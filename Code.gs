/** Code.gs - Dynamic Campaign-aware Routing System
 *
 * this host to the ORIGINAL_SCRIPT_URL Apps Script exec endpoint.
 */
// ───────────────────────────────────────────────────────────────────────────────
// GLOBAL CONSTANTS AND CONFIGURATION
// ───────────────────────────────────────────────────────────────────────────────

// Main script configuration
const SCRIPT_URL = 'https://script.google.com/a/macros/vlbpo.com/s/AKfycbxeQ0AnupBHM71M6co3LVc5NPrxTblRXLd6AuTOpxMs2rMehF9dBSkGykIcLGHROywQ/exec';
const FAVICON_URL = 'https://res.cloudinary.com/dr8qd3xfc/image/upload/v1754763514/vlbpo/lumina/3_dgitcx.png';


/** Toggle to print deep traces into console for every access decision. */
const ACCESS_DEBUG = true;

/** Canonical page sets (lower-case keys) */
const ACCESS = {
  ADMIN_ONLY_PAGES: new Set(['admin.users', 'admin.roles', 'admin.campaigns']),
  PUBLIC_PAGES: new Set([
    'login', 'setpassword', 'resetpassword', 'forgotpassword', 'forgot-password',
    'resendverification', 'resend-verification', 'emailconfirmed', 'email-confirmed'
  ]),
  DEFAULT_PAGE: 'dashboard',
  PRIVS: { SYSTEM_ADMIN: 'SYSTEM_ADMIN', MANAGE_USERS: 'MANAGE_USERS', MANAGE_PAGES: 'MANAGE_PAGES' }
};

// ---- small utilities ---------------------------------------------------------
function _truthy(v) {
  return v === true || String(v).toUpperCase() === 'TRUE';
}

function _normalizePageKey(k) {
  return (k || '').toString().trim().toLowerCase();
}

function _normalizeId(x) {
  return (x == null ? '' : String(x)).trim();
}

function _now() {
  return new Date();
}

function _safeDate(d) {
  try {
    return new Date(d);
  } catch (e) {
    return null;
  }
}

function _isFuture(d) {
  try {
    return d && d.getTime && d.getTime() > Date.now();
  } catch (_) {
    return false;
  }
}

function _csvToSet(csv) {
  const set = new Set();
  (String(csv || '').split(',') || []).forEach(s => {
    const k = _normalizePageKey(s);
    if (k) set.add(k);
  });
  return set;
}

// ---- caching helpers (CacheService with fallback to script properties) -------
function _cacheGet(key) {
  try {
    const c = CacheService.getScriptCache().get(key);
    if (c) return JSON.parse(c);
  } catch (_) {
  }
  try {
    const p = PropertiesService.getScriptProperties().getProperty(key);
    if (p) return JSON.parse(p);
  } catch (_) {
  }
  return null;
}

function _cachePut(key, obj, ttlSec) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(obj), Math.min(21600, Math.max(5, ttlSec || 60)));
  } catch (_) {
  }
  try {
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(obj));
  } catch (_) {
  }
}

// ---- load metadata about pages ----------------------------------------------
function _loadAllPagesMeta() {
  const CK = 'ACCESS_PAGES_META_V1';
  const cached = _cacheGet(CK);
  if (cached) return cached;
  let rows = [];
  try {
    rows = readSheet(PAGES_SHEET) || [];
  } catch (_) {
  }
  // Build meta map: key -> meta
  const map = {};
  rows.forEach(r => {
    const key = _normalizePageKey(r.PageKey || r.key);
    if (!key) return;
    map[key] = {
      key,
      title: r.PageTitle || r.Name || '',
      requiresAdmin: _truthy(r.RequiresAdmin),
      campaignSpecific: _truthy(r.CampaignSpecific) || _truthy(r.IsCampaignSpecific) || false,
      active: (r.Active === undefined) ? true : _truthy(r.Active)
    };
  });
  _cachePut(CK, map, 180);
  return map;
}

function _getPageMeta(pageKey) {
  const key = _normalizePageKey(pageKey);
  const meta = _loadAllPagesMeta()[key];
  // Safe default if not registered: treat as non-admin, non-campaign, active
  return meta || { key, title: pageKey, requiresAdmin: false, campaignSpecific: false, active: true };
}

// ---- user normalization + grant loading -------------------------------------
function _normalizeUser(user) {
  const out = Object.assign({}, user || {});
  // hard booleans
  out.IsAdmin = _truthy(out.IsAdmin) || String((out.roleNames || [])).toLowerCase().includes('system admin');
  out.CanLogin = _truthy(out.CanLogin !== undefined ? out.CanLogin : true);
  out.EmailConfirmed = _truthy(out.EmailConfirmed !== undefined ? out.EmailConfirmed : true);
  out.ResetRequired = _truthy(out.ResetRequired);
  out.LockoutEnd = out.LockoutEnd ? _safeDate(out.LockoutEnd) : null;
  out.ID = _normalizeId(out.ID || out.id);
  out.CampaignID = _normalizeId(out.CampaignID || out.campaignId);
  out.PagesCsv = String(out.Pages || out.pages || '');
  return out;
}

/** Load effective grants for user: assigned pages + per-campaign privileges + explicit per-page grants */
function _loadUserGrants(userId) {
  const id = _normalizeId(userId);
  const CK = 'ACCESS_USER_GRANTS_' + id;
  const cached = _cacheGet(CK);
  if (cached) return cached;

  // Users.Pages (CSV) -> Set of canonical keys
  let users = [];
  try { users = readSheet(USERS_SHEET) || []; } catch (_) { }
  const me = users.find(u => _normalizeId(u.ID) === id) || {};
  const rawAssigned = _csvToSet(me.Pages || me.pages);
  const assignedPagesCanon = new Set();
  rawAssigned.forEach(p => assignedPagesCanon.add(canonicalizePageKey(p)));

  // Per-campaign privileges
  let perms = [];
  try {
    perms = readSheet((typeof CAMPAIGN_USER_PERMISSIONS_SHEET !== 'undefined'
      ? CAMPAIGN_USER_PERMISSIONS_SHEET : 'CAMPAIGN_USER_PERMISSIONS')) || [];
  } catch (_) { }
  const campaignPrivs = {};
  perms.filter(p => _normalizeId(p.UserID) === id).forEach(p => {
    const cid = _normalizeId(p.CampaignID);
    if (!cid) return;
    campaignPrivs[cid] = {
      level: String(p.PermissionLevel || '').toUpperCase(), // VIEWER/MANAGER/ADMIN
      canManageUsers: _truthy(p.CanManageUsers),
      canManagePages: _truthy(p.CanManagePages)
    };
  });

  // Explicit page grants (UserPagePermissions) — canonicalize keys here too
  let upp = [];
  try { upp = readSheet('UserPagePermissions') || []; } catch (_) { }
  const pageGrants = {}; // campaignId -> Set(pageKey)
  upp.filter(p => _normalizeId(p.UserID || p.userId) === id).forEach(p => {
    const cid = _normalizeId(p.CampaignID || p.campaignId) || '*';
    const page = canonicalizePageKey(p.PageKey || p.pageKey);
    const allowed = (p.CanView || p.View || p.Access || p.access);
    if (!page || !allowed) return;
    if (!pageGrants[cid]) pageGrants[cid] = new Set();
    pageGrants[cid].add(page);
  });

  const grants = { assignedPagesCanon, campaignPrivs, pageGrants };
  _cachePut(CK, grants, 120);
  return grants;
}


function handleLogoutRequest(e) {
  try {
    const token = e.parameter.token || '';

    // Attempt to logout via AuthenticationService
    let logoutResult = false;
    if (typeof AuthenticationService !== 'undefined' && AuthenticationService.logout) {
      try {
        logoutResult = AuthenticationService.logout(token);
      } catch (logoutError) {
        console.warn('AuthenticationService logout failed:', logoutError);
      }
    }

    // Create logout confirmation page
    return HtmlService.createHtmlOutput(`
            <!DOCTYPE html>
            <html>
            <head>
                <title>Logged Out</title>
                <meta name="viewport" content="width=device-width,initial-scale=1">
                <style>
                    body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
                    .success { color: #27ae60; }
                    .message { margin: 20px 0; }
                    .login-link { color: #3498db; text-decoration: none; padding: 10px 20px; background: #ecf0f1; border-radius: 5px; }
                </style>
            </head>
            <body>
                <h1 class="success">Successfully Logged Out</h1>
                <p class="message">You have been logged out of the system.</p>
                <a href="${SCRIPT_URL}" class="login-link">Return to Login</a>
                <script>
                    setTimeout(function() {
                        window.location.href = "${SCRIPT_URL}";
                    }, 3000);
                </script>
            </body>
            </html>
        `).setTitle('Logged Out')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    console.error('Error handling logout request:', error);
    writeError('handleLogoutRequest', error);

    // Fallback logout page
    return HtmlService.createHtmlOutput(`
            <!DOCTYPE html>
            <html>
            <head>
                <title>Logout</title>
                <meta name="viewport" content="width=device-width,initial-scale=1">
            </head>
            <body>
                <h1>Logout</h1>
                <p>Redirecting to login...</p>
                <script>
                    window.location.href = "${SCRIPT_URL}";
                </script>
            </body>
            </html>
        `).setTitle('Logout')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// CLIENT-ACCESSIBLE FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

function getUserRoles(token) {
  try {
    if (typeof AuthenticationService !== 'undefined' && AuthenticationService.getUserRoles) {
      return AuthenticationService.getUserRoles(token);
    }
    return { success: false, error: 'AuthenticationService not available' };
  } catch (error) {
    writeError('getUserRoles', error);
    return { success: false, error: error.message };
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// CURRENT USER LOADER (universal; available to all pages)
// ───────────────────────────────────────────────────────────────────────────────

/** internal: safe JSON stringify for HTML templates */
function _stringifyForTemplate_(obj) {
  try { return JSON.stringify(obj || {}).replace(/<\/script>/g, '<\\/script>'); }
  catch (_) { return '{}'; }
}

/** internal: find a Users row by email (case-insensitive) */
function _findUserByEmail_(email) {
  if (!email) return null;
  try {
    const CK = 'USR_BY_EMAIL_' + email.toLowerCase();
    const cached = _cacheGet(CK);
    if (cached) return cached;

    const rows = (typeof readSheet === 'function') ? (readSheet(USERS_SHEET) || []) : [];
    const hit = rows.find(r => String(r.Email || '').trim().toLowerCase() === email.toLowerCase()) || null;
    if (hit) _cachePut(CK, hit, 120);
    return hit;
  } catch (e) {
    writeError && writeError('_findUserByEmail_', e);
    return null;
  }
}

/** internal: convert a Users row to a safe client user object */
function _toClientUser_(row, fallbackEmail) {
  const rolesMap = (typeof getRolesMapping === 'function') ? getRolesMapping() : {};
  const roleIds = String(row && row.Roles || '')
    .split(',').map(s => s.trim()).filter(Boolean);
  const roleNames = roleIds.map(id => rolesMap[id]).filter(Boolean);

  const isAdminFlag =
    (String(row && row.IsAdmin).toLowerCase() === 'true') ||
    roleNames.some(n => /system\s*admin|administrator|super\s*admin/i.test(String(n || '')));

  const client = {
    ID: String(row && row.ID || ''),
    Email: String(row && row.Email || fallbackEmail || '').trim(),
    FullName: String(row && (row.FullName || row.UserName || '') || '').trim() ||
      String(fallbackEmail || '').split('@')[0],
    CampaignID: String(row && (row.CampaignID || row.CampaignId || '') || ''),
    roleNames: roleNames,
    IsAdmin: !!isAdminFlag,
    CanLogin: (row && row.CanLogin !== undefined) ? _truthy(row.CanLogin) : true,
    EmailConfirmed: (row && row.EmailConfirmed !== undefined) ? _truthy(row.EmailConfirmed) : true,
    ResetRequired: !!(row && _truthy(row.ResetRequired)),
    Pages: String(row && (row.Pages || row.pages || '') || '')
  };
  return client;
}

/** getCurrentUser(): resolves the current active user from the account email */
function getCurrentUser() {
  try {
    // Prefer the real account; fall back to effective user if needed
    const email = String(
      (Session.getActiveUser() && Session.getActiveUser().getEmail()) ||
      (Session.getEffectiveUser() && Session.getEffectiveUser().getEmail()) ||
      ''
    ).trim().toLowerCase();

    const row = _findUserByEmail_(email);
    // If not in Users sheet, still return a minimal shape so UI can work
    const client = _toClientUser_(row, email);

    // hydrate some campaign context if available (non-blocking)
    try {
      if (client.CampaignID) {
        if (typeof getCampaignById === 'function') {
          const c = getCampaignById(client.CampaignID);
          client.campaignName = c ? (c.Name || c.name || '') : '';
        }
        if (typeof getUserCampaignPermissions === 'function') {
          client.campaignPermissions = getUserCampaignPermissions(client.ID);
        }
        if (typeof getCampaignNavigation === 'function') {
          client.campaignNavigation = getCampaignNavigation(client.CampaignID);
        }
      }
    } catch (ctxErr) {
      console.warn('getCurrentUser: campaign context hydrate failed:', ctxErr);
    }

    return client;
  } catch (e) {
    writeError && writeError('getCurrentUser', e);
    return _toClientUser_(null, '');
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// STRICT users: self + explicitly assigned only
// ───────────────────────────────────────────────────────────────────────────────

// Optional sheet name helpers
function getManagerUsersSheetName_() { return 'MANAGER_USERS'; }
function getCampaignsSheetName_() { return 'CAMPAIGNS'; }

function _normStr_(s) { return String(s || '').trim(); }
function _normEmail_(s) { return _normStr_(s).toLowerCase(); }
function _toBool_(v) { return v === true || String(v || '').trim().toUpperCase() === 'TRUE'; }

function _tryGetRoleNames_(userId) {
  try {
    if (typeof getUserRoleNamesSafe === 'function') {
      var list = getUserRoleNamesSafe(userId) || [];
      return Array.isArray(list) ? list : [];
    }
  } catch (_) { }
  return [];
}

function _normEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function _normStr_(str) {
  return String(str || '').trim();
}

function _campaignNameMap_() {
  try {
    const rows = (typeof readSheet === 'function') ? (readSheet('CAMPAIGNS') || []) : [];
    const map = {};
    rows.forEach(function (r) {
      const id = _normStr_(r.ID || r.Id || r.id);
      const nm = _normStr_(r.Name || r.name);
      if (id) map[id] = nm;
    });
    return map;
  } catch (_) {
    return {};
  }
}

function _uiUserShape_(u, cmap) {
  const cid = _normStr_(u.CampaignID || u.campaignId);
  return {
    ID: u.ID,
    UserName: u.UserName,
    FullName: u.FullName,
    Email: u.Email,
    CampaignID: u.CampaignID,
    campaignName: cmap[cid] || _normStr_(u.CampaignName || ''),
    CanLogin: u.CanLogin,
    IsAdmin: u.IsAdmin,
    canLoginBool: _toBool_(u.CanLogin),
    isAdminBool: _toBool_(u.IsAdmin),
    roleNames: _tryGetRoleNames_(u.ID),
    pages: []
  };
}

/**
 * Enhanced getUsers function with better error handling
 * This should replace the existing function in Code.gs
 */
function getUsers() {
  try {
    console.log('getUsers() called');

    // Get current user email
    const meEmail = _normEmail_((Session.getActiveUser() && Session.getActiveUser().getEmail()) || '');
    console.log('Current user email:', meEmail);

    if (!meEmail) {
      console.warn('No current user email found');
      return [];
    }

    // Read all users
    const allUsers = (typeof readSheet === 'function') ? (readSheet(USERS_SHEET) || []) : [];
    console.log('All users from sheet:', allUsers.length);

    if (allUsers.length === 0) {
      console.warn('No users found in Users sheet');
      return [];
    }

    // Find current user
    const me = allUsers.find(function (u) {
      return _normEmail_(u.Email || u.email) === meEmail;
    });

    if (!me) {
      console.warn('Current user not found in Users sheet');
      // Fallback: return all users if current user is not found
      return allUsers.map(u => _uiUserShape_(u, _campaignNameMap_()));
    }

    console.log('Found current user:', me.FullName || me.UserName);

    // Try to get manager-user relationships
    let muRows = [];
    try {
      muRows = (typeof readSheet === 'function') ? (readSheet(getManagerUsersSheetName_()) || []) : [];
      console.log('Manager-user relationships found:', muRows.length);
    } catch (error) {
      console.warn('Could not read manager-user relationships, will return all users:', error);
      // Fallback: return all users if manager relationships can't be read
      return allUsers.map(u => _uiUserShape_(u, _campaignNameMap_()));
    }

    // Get assigned user IDs
    const assignedIds = new Set(
      muRows.filter(function (a) { return String(a.ManagerUserID) === String(me.ID); })
        .map(function (a) { return String(a.UserID); })
        .filter(Boolean)
    );

    console.log('Assigned user IDs:', Array.from(assignedIds));

    // If no assigned users, return current user + users from same campaign
    if (assignedIds.size === 0) {
      console.log('No assigned users found, using campaign-based filtering');
      const sameCampaignUsers = allUsers.filter(u =>
        (u.CampaignID || u.campaignId) === (me.CampaignID || me.campaignId)
      );
      return sameCampaignUsers.map(u => _uiUserShape_(u, _campaignNameMap_()));
    }

    // Build user map
    const byId = new Map(allUsers.filter(function (u) { return u && u.ID; }).map(function (u) { return [String(u.ID), u]; }));
    const cmap = _campaignNameMap_();

    // Build output with current user + assigned users
    const out = [_uiUserShape_(me, cmap)];
    assignedIds.forEach(function (id) {
      const u = byId.get(id);
      if (u) out.push(_uiUserShape_(u, cmap));
    });

    // Remove duplicates
    const seen = new Set();
    const dedup = out.filter(function (u) {
      const k = String(u.ID || '');
      if (!k || seen.has(k)) return false;
      seen.add(k);
      return true;
    });

    // Sort by name
    dedup.sort(function (a, b) {
      return String(a.FullName || '').localeCompare(String(b.FullName || '')) ||
        String(a.UserName || '').localeCompare(String(b.UserName || ''));
    });

    console.log('Final user list:', dedup.length, 'users');
    return dedup;

  } catch (e) {
    console.error('Error in getUsers:', e);
    writeError && writeError('getUsers (enhanced)', e);

    // Final fallback: try to return at least the current user
    try {
      const currentUser = getCurrentUser();
      if (currentUser) {
        return [currentUser];
      }
    } catch (fallbackError) {
      console.error('Fallback getCurrentUser also failed:', fallbackError);
    }

    return [];
  }
}

/**
 * Alternative user list function that doesn't depend on manager relationships
 */
function getSimpleUserList(campaignId) {
  try {
    console.log('getSimpleUserList called with campaignId:', campaignId);

    // Read users directly from sheet
    const users = readSheet(USERS_SHEET) || [];
    console.log('Read', users.length, 'users from sheet');

    let filteredUsers = users;

    // Filter by campaign if specified
    if (campaignId && campaignId.trim() !== '') {
      filteredUsers = users.filter(u =>
        String(u.CampaignID || u.campaignId || '').trim() === String(campaignId).trim()
      );
      console.log('Filtered to', filteredUsers.length, 'users for campaign:', campaignId);
    }

    // Extract names
    const names = filteredUsers
      .map(u => u.FullName || u.UserName || u.Email || u.name || '')
      .filter(name => name.trim() !== '')
      .filter((name, index, arr) => arr.indexOf(name) === index) // Remove duplicates
      .sort();

    console.log('Returning', names.length, 'user names');
    return names;

  } catch (error) {
    console.error('Error in getSimpleUserList:', error);
    return [];
  }
}

function clientGetAssignedAgentNames(campaignId) {
  try {
    console.log('clientGetAssignedAgentNames called with campaignId:', campaignId);

    // Strategy 1: Try the original assigned-only approach
    let users = [];
    try {
      users = getUsers();
      console.log('getUsers() returned:', users.length, 'users');
    } catch (error) {
      console.warn('getUsers() failed, trying fallback approaches:', error);
    }

    // Strategy 2: If no users from assigned approach, try reading Users sheet directly
    if (!users || users.length === 0) {
      try {
        console.log('Trying direct Users sheet read...');
        const allUsers = readSheet(USERS_SHEET) || [];
        console.log('Direct Users sheet read returned:', allUsers.length, 'users');

        // Filter by campaign if specified
        if (campaignId && campaignId.trim() !== '') {
          users = allUsers.filter(u =>
            String(u.CampaignID || u.campaignId || '').trim() === String(campaignId).trim()
          );
          console.log('Filtered by campaignId:', users.length, 'users');
        } else {
          users = allUsers;
        }

        // Convert to expected format
        users = users.map(u => ({
          ID: u.ID || u.id,
          FullName: u.FullName || u.UserName || u.name,
          UserName: u.UserName || u.name,
          Email: u.Email || u.email,
          CampaignID: u.CampaignID || u.campaignId
        }));
      } catch (error) {
        console.error('Direct Users sheet read failed:', error);
      }
    }

    // Strategy 3: If still no users, try getting current user at least
    if (!users || users.length === 0) {
      try {
        console.log('No users found, trying to get current user...');
        const currentUser = getCurrentUser();
        if (currentUser && currentUser.FullName) {
          users = [currentUser];
          console.log('Using current user as fallback:', currentUser.FullName);
        }
      } catch (error) {
        console.error('getCurrentUser failed:', error);
      }
    }

    // Strategy 4: Final fallback - hardcoded test users for development
    if (!users || users.length === 0) {
      console.warn('All user retrieval strategies failed, using empty array');
      users = [];
    }

    // Extract display names
    const displayNames = users
      .map(u => {
        const name = u.FullName || u.UserName || u.Email || u.name || '';
        return name.trim();
      })
      .filter(name => name !== '')
      .filter((name, index, arr) => arr.indexOf(name) === index) // Remove duplicates
      .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));

    console.log('Final userList:', displayNames);
    return displayNames;

  } catch (error) {
    console.error('Error in clientGetAssignedAgentNames:', error);
    writeError('clientGetAssignedAgentNames', error);
    return [];
  }
}

/**
 * Client-accessible wrapper with multiple fallback strategies
 */
function clientGetUserList(campaignId) {
  try {
    // Try the enhanced assigned approach first
    let userList = clientGetAssignedAgentNames(campaignId);

    // If that fails, try simple approach
    if (!userList || userList.length === 0) {
      console.log('Assigned approach failed, trying simple approach');
      userList = getSimpleUserList(campaignId);
    }

    // If still empty, try without campaign filter
    if (!userList || userList.length === 0) {
      console.log('Campaign-filtered approach failed, trying all users');
      userList = getSimpleUserList('');
    }

    return userList;

  } catch (error) {
    console.error('Error in clientGetUserList:', error);
    return [];
  }
}

/** client-callable wrapper for pages that want to refresh on demand */
function clientGetCurrentUser() {
  return getCurrentUser();
}

/** inject user + JSON into any HtmlTemplate (use in all page renderers) */
function __injectCurrentUser_(tpl, explicitUser) {
  try {
    const u = explicitUser && (explicitUser.Email || explicitUser.email || explicitUser.ID)
      ? explicitUser
      : getCurrentUser();
    tpl.user = u;
    tpl.currentUserJson = _stringifyForTemplate_(u);
  } catch (e) {
    writeError && writeError('__injectCurrentUser_', e);
    tpl.user = null;
    tpl.currentUserJson = '{}';
  }
}

function login(email, password) {
  try {
    if (typeof AuthenticationService !== 'undefined' && AuthenticationService.login) {
      return AuthenticationService.login(email, password);
    }
    return { success: false, error: 'AuthenticationService not available' };
  } catch (error) {
    writeError('login', error);
    return { success: false, error: error.message };
  }
}

/**
 * Enhanced logout function that handles session cleanup and redirect
 */
function logout(token) {
  try {
    console.log('logout function called with token:', token ? token.substring(0, 8) + '...' : 'no token');

    // Clean up the session using AuthenticationService
    let sessionRemoved = false;

    try {
      if (typeof AuthenticationService !== 'undefined' && AuthenticationService.logout) {
        sessionRemoved = AuthenticationService.logout(token);
        console.log('AuthenticationService.logout result:', sessionRemoved);
      }
    } catch (authError) {
      console.warn('AuthenticationService.logout failed:', authError);
      // Continue with logout even if session removal fails
    }

    // Get the base URL for redirect (without token = login page)
    const baseUrl = SCRIPT_URL; // No token parameter = login page

    const result = {
      success: true,
      sessionRemoved: sessionRemoved,
      redirectUrl: baseUrl,
      message: 'Logged out successfully'
    };

    console.log('logout function returning:', result);
    return result;

  } catch (error) {
    console.error('Error during logout:', error);
    writeError('logout', error);

    // Even if there's an error, still redirect to login
    return {
      success: false,
      error: error.message,
      redirectUrl: SCRIPT_URL,
      message: 'Logout completed with errors'
    };
  }
}

/**
 * Enhanced keepAlive function that handles session expiration
 */
function keepAlive(token) {
  try {
    console.log('keepAlive called with token:', token ? token.substring(0, 8) + '...' : 'no token');

    if (typeof AuthenticationService !== 'undefined' && AuthenticationService.keepAlive) {
      const isValid = AuthenticationService.keepAlive(token);

      if (!isValid) {
        // Session expired, return logout info
        return {
          success: false,
          expired: true,
          redirectUrl: SCRIPT_URL,
          message: 'Session expired'
        };
      }

      return {
        success: true,
        message: 'Session active'
      };
    }

    return {
      success: false,
      error: 'AuthenticationService not available',
      redirectUrl: SCRIPT_URL
    };
  } catch (error) {
    console.error('Error in keepAlive:', error);
    writeError('keepAlive', error);

    return {
      success: false,
      error: error.message,
      redirectUrl: SCRIPT_URL,
      message: 'Session check failed'
    };
  }
}

/**
 * Client-accessible logout function for google.script.run
 */
function clientLogout(token) {
  try {
    console.log('clientLogout called with token:', token ? token.substring(0, 8) + '...' : 'no token');

    // Call the main logout function
    const result = logout(token);

    console.log('clientLogout result:', result);
    return result;

  } catch (error) {
    console.error('Error in clientLogout:', error);
    writeError('clientLogout', error);

    // Return a safe result even on error
    return {
      success: false,
      error: error.message,
      redirectUrl: SCRIPT_URL,
      message: 'Logout failed but redirecting to login'
    };
  }
}

function resendVerificationEmail(email) {
  try {
    if (typeof clientResendVerificationEmail === 'function') {
      return clientResendVerificationEmail(email);
    } else {
      return { success: false, error: 'Email service not available' };
    }
  } catch (error) {
    writeError('resendVerificationEmail', error);
    return { success: false, error: error.message };
  }
}

function sendTestEmail(email) {
  try {
    if (typeof clientSendTestEmail === 'function') {
      return clientSendTestEmail(email);
    } else {
      return { success: false, error: 'Email service not available' };
    }
  } catch (error) {
    writeError('sendTestEmail', error);
    return { success: false, error: error.message };
  }
}

function requestPasswordReset(email) {
  try {
    if (typeof AuthenticationService !== 'undefined' && AuthenticationService.requestPasswordReset) {
      return AuthenticationService.requestPasswordReset(email);
    }
    return { success: false, error: 'AuthenticationService not available' };
  } catch (error) {
    writeError('requestPasswordReset', error);
    return { success: false, error: error.message };
  }
}

function setPasswordWithToken(token, newPassword) {
  try {
    if (typeof AuthenticationService !== 'undefined' && AuthenticationService.setPasswordWithToken) {
      return AuthenticationService.setPasswordWithToken(token, newPassword);
    }
    return { success: false, error: 'AuthenticationService not available' };
  } catch (error) {
    writeError('setPasswordWithToken', error);
    return { success: false, error: error.message };
  }
}

function changePassword(token, oldPassword, newPassword) {
  try {
    if (typeof AuthenticationService !== 'undefined' && AuthenticationService.changePassword) {
      return AuthenticationService.changePassword(token, oldPassword, newPassword);
    }
    return { success: false, error: 'AuthenticationService not available' };
  } catch (error) {
    writeError('changePassword', error);
    return { success: false, error: error.message };
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// CACHE UTILITIES
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Cache utilities
 */
if (typeof scriptCache === 'undefined') {
  const scriptCache = {
    get: function (key) {
      try {
        return PropertiesService.getScriptProperties().getProperty(key);
      } catch (e) {
        return null;
      }
    },

    put: function (key, value, ttl) {
      try {
        PropertiesService.getScriptProperties().setProperty(key, value);
        return true;
      } catch (e) {
        return false;
      }
    },

    remove: function (key) {
      try {
        PropertiesService.getScriptProperties().deleteProperty(key);
        return true;
      } catch (e) {
        return false;
      }
    }
  };

  this.scriptCache = scriptCache;
}

// ───────────────────────────────────────────────────────────────────────────────
// QA ANALYTICS UTILITY FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Utility function for empty QA analytics
 */
function getEmptyQAAnalytics() {
  return {
    avgScore: 0,
    passRate: 0,
    totalEvaluations: 0,
    agentsEvaluated: 0,
    avgScoreChange: 0,
    passRateChange: 0,
    evaluationsChange: 0,
    agentsChange: 0,
    categories: { labels: [], values: [] },
    trends: { labels: [], values: [] },
    agents: { labels: [], values: [] }
  };
}

// ───────────────────────────────────────────────────────────────────────────────
// SHEET CONSTANT DEFINITIONS (if not defined elsewhere)
// ───────────────────────────────────────────────────────────────────────────────

if (typeof USERS_SHEET === 'undefined') {
  const USERS_SHEET = 'Users';
}
if (typeof ROLES_SHEET === 'undefined') {
  const ROLES_SHEET = 'Roles';
}
if (typeof PAGES_SHEET === 'undefined') {
  const PAGES_SHEET = 'Pages';
}
if (typeof CAMPAIGNS_SHEET === 'undefined') {
  const CAMPAIGNS_SHEET = 'Campaigns';
}

// ───────────────────────────────────────────────────────────────────────────────
// SYSTEM PAGES INITIALIZATION
// ───────────────────────────────────────────────────────────────────────────────

function initializeSystemPages() {
  try {
    const systemPages = [
      {
        PageKey: 'dashboard',
        Name: 'Dashboard',
        Description: 'Main dashboard page',
        RequiredRole: '',
        CampaignSpecific: false,
        Active: true
      },
      {
        PageKey: 'users',
        Name: 'User Management',
        Description: 'Manage system users',
        RequiredRole: 'admin',
        CampaignSpecific: false,
        Active: true
      },
      {
        PageKey: 'roles',
        Name: 'Role Management',
        Description: 'Manage user roles',
        RequiredRole: 'admin',
        CampaignSpecific: false,
        Active: true
      },
      {
        PageKey: 'campaigns',
        Name: 'Campaign Management',
        Description: 'Manage campaigns',
        RequiredRole: 'admin',
        CampaignSpecific: false,
        Active: true
      },
      {
        PageKey: 'tasks',
        Name: 'Task Management',
        Description: 'Task board and management',
        RequiredRole: '',
        CampaignSpecific: true,
        Active: true
      },
      {
        PageKey: 'search',
        Name: 'Search',
        Description: 'Global search functionality',
        RequiredRole: '',
        CampaignSpecific: false,
        Active: true
      },
      {
        PageKey: 'chat',
        Name: 'Chat',
        Description: 'Team communication',
        RequiredRole: '',
        CampaignSpecific: false,
        Active: true
      },
      {
        PageKey: 'notifications',
        Name: 'Notifications',
        Description: 'System notifications',
        RequiredRole: '',
        CampaignSpecific: false,
        Active: true
      }
    ];

    if (typeof writeSheet === 'function') {
      writeSheet(PAGES_SHEET, systemPages);
      console.log('System pages initialized');
    }

  } catch (error) {
    console.error('Error initializing system pages:', error);
    writeError('initializeSystemPages', error);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// INITIALIZATION ON LOAD
// ───────────────────────────────────────────────────────────────────────────────

console.log('Enhanced Code.gs loaded successfully');
console.log('Features: Campaign-aware routing, Enhanced authentication, Improved error handling');
console.log('Script URL:', SCRIPT_URL);

// Auto-initialize system if needed (with error handling)
try {
  // Run basic validation
  if (typeof SpreadsheetApp !== 'undefined') {
    console.log('SpreadsheetApp available - system ready');
  }
} catch (error) {
  console.error('Error during auto-initialization:', error);
  writeError('autoInitialization', error);
}

// ───────────────────────────────────────────────────────────────────────────────
// GLOBAL FAVICON INJECTOR (idempotent)
// ───────────────────────────────────────────────────────────────────────────────

(function () {
  // If this patch already ran (another file or re-deploy), don't wrap again
  if (HtmlService.__faviconPatched === true) return;
  HtmlService.__faviconPatched = true;

  const _origCreate = HtmlService.createTemplateFromFile;

  HtmlService.createTemplateFromFile = function (file) {
    const tpl = _origCreate.call(HtmlService, file);

    // Wrap evaluate only once per template instance
    const _origEval = tpl.evaluate;
    tpl.evaluate = function () {
      const out = _origEval.apply(tpl, arguments);
      return (typeof out.setFaviconUrl === 'function') ? out.setFaviconUrl(FAVICON_URL) : out;
    };

    return tpl;
  };
})();

// ───────────────────────────────────────────────────────────────────────────────
// TEMPLATE INCLUDE HELPER (cycle-safe)
// ───────────────────────────────────────────────────────────────────────────────

const __INCLUDE_STACK = [];
const __INCLUDED_ONCE = new Set();

function include(file, params) {
  if (__INCLUDE_STACK.indexOf(file) !== -1) {
    throw new RangeError('Cyclic include detected: ' + __INCLUDE_STACK.concat(file).join(' → '));
  }
  if (__INCLUDE_STACK.length > 20) {
    throw new RangeError('Include depth > 20; aborting to avoid stack overflow.');
  }

  __INCLUDE_STACK.push(file);
  try {
    const tpl = HtmlService.createTemplateFromFile(file);
    if (params) Object.keys(params).forEach(k => (tpl[k] = params[k]));

    try {
      if (!tpl.user) tpl.user = getCurrentUserByEmail_();
      tpl.currentUserJson = _stringifyForTemplate_(tpl.user);
    } catch (_) {
      tpl.user = null;
      tpl.currentUserJson = '{}';
    }

    return tpl.evaluate().getContent();
  } finally {
    __INCLUDE_STACK.pop();
  }
}

function includeOnce(file, params) {
  if (__INCLUDED_ONCE.has(file)) return '';
  __INCLUDED_ONCE.add(file);
  return include(file, params);
}

function getCurrentUserByEmail_() {
  try {
    const email = String(
      (Session.getActiveUser() && Session.getActiveUser().getEmail()) ||
      (Session.getEffectiveUser() && Session.getEffectiveUser().getEmail()) || ''
    ).trim().toLowerCase();
    if (!email) return null;

    const rows = (typeof readSheet === 'function') ? (readSheet(USERS_SHEET) || []) : [];
    const row = rows.find(r => String(r.Email || '').trim().toLowerCase() === email);
    return row ? _normalizeUser(row) : null;  // you already have _normalizeUser in your code
  } catch (e) {
    return null;
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// SYSTEM INITIALIZATION
// ───────────────────────────────────────────────────────────────────────────────

function initializeSystem() {
  try {
    console.log('Initializing system...');

    // Ensure all required sheets exist
    if (typeof AuthenticationService !== 'undefined' && AuthenticationService.ensureSheets) {
      AuthenticationService.ensureSheets();
    }

    // Initialize main sheets
    initializeMainSheets();

    // Initialize campaign-specific systems
    initializeCampaignSystems();

    console.log('System initialization completed successfully');
    return { success: true, message: 'System initialized' };

  } catch (error) {
    console.error('System initialization failed:', error);
    writeError('initializeSystem', error);
    return { success: false, error: error.message };
  }
}

function initializeMainSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Initialize required sheets if they don't exist
    const requiredSheets = [USERS_SHEET, ROLES_SHEET, PAGES_SHEET, CAMPAIGNS_SHEET];

    requiredSheets.forEach(sheetName => {
      try {
        let sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
          console.log(`Creating missing sheet: ${sheetName}`);
          sheet = ss.insertSheet(sheetName);
          initializeSheetHeaders(sheet, sheetName);
        }
      } catch (sheetError) {
        console.warn(`Could not initialize sheet ${sheetName}:`, sheetError);
      }
    });

    console.log('Main sheets initialized');

  } catch (error) {
    console.error('Error initializing main sheets:', error);
    writeError('initializeMainSheets', error);
  }
}

function initializeSheetHeaders(sheet, sheetName) {
  try {
    let headers = [];

    switch (sheetName) {
      case USERS_SHEET:
        headers = ['ID', 'Email', 'FullName', 'Password', 'Roles', 'CampaignID', 'EmailConfirmed', 'EmailConfirmation', 'CreatedAt', 'UpdatedAt'];
        break;
      case ROLES_SHEET:
        headers = ['ID', 'Name', 'Description', 'Permissions', 'CreatedAt'];
        break;
      case PAGES_SHEET:
        headers = ['PageKey', 'Name', 'Description', 'RequiredRole', 'CampaignSpecific', 'Active'];
        break;
      case CAMPAIGNS_SHEET:
        headers = ['ID', 'Name', 'Description', 'Active', 'Settings', 'CreatedAt'];
        break;
      default:
        console.warn(`No headers defined for sheet: ${sheetName}`);
        return;
    }

    if (headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      console.log(`Headers set for ${sheetName}:`, headers);
    }

  } catch (error) {
    console.error(`Error setting headers for ${sheetName}:`, error);
  }
}

function initializeCampaignSystems() {
  try {
    // Initialize Independence Insurance QA system if available
    if (typeof initializeIndependenceQASystem === 'function') {
      try {
        initializeIndependenceQASystem();
        console.log('Independence QA system initialized');
      } catch (error) {
        console.warn('Independence QA system initialization failed:', error);
      }
    }

    // Initialize Credit Suite QA system if available
    if (typeof initializeCreditSuiteQASystem === 'function') {
      try {
        initializeCreditSuiteQASystem();
        console.log('Credit Suite QA system initialized');
      } catch (error) {
        console.warn('Credit Suite QA system initialization failed:', error);
      }
    }

    // Initialize system pages if Pages sheet is empty
    if (typeof readSheet === 'function') {
      const pages = readSheet(PAGES_SHEET);
      if (pages.length === 0) {
        initializeSystemPages();
      }
    }

    console.log('Campaign systems initialization completed');

  } catch (error) {
    console.error('Error initializing campaign systems:', error);
    writeError('initializeCampaignSystems', error);
  }
}

function initializeCampaignService() {
  try {
    // Initialize system components
    initializeSystem();

    writeDebug('Campaign service initialized successfully');
    return { success: true, message: 'Campaign service initialized' };
  } catch (e) {
    writeError('initializeCampaignService', e);
    return { success: false, error: 'Failed to initialize campaign service: ' + e.message };
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// ENHANCED AUTHENTICATION WITH CAMPAIGN AWARENESS
// ───────────────────────────────────────────────────────────────────────────────

// Centralized "Access Denied" renderer
function renderAccessDenied(baseUrl, message) {
  const tpl = HtmlService.createTemplateFromFile('AccessDenied');
  tpl.baseUrl = baseUrl;
  tpl.message = message || 'You do not have permission to view this page.';
  return tpl.evaluate()
    .setTitle('Access Denied')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---- campaign membership / access -------------------------------------------
function hasCampaignAccess(user, campaignId) {
  try {
    const u = _normalizeUser(user);
    const cid = _normalizeId(campaignId);
    if (!cid) return true;               // no campaign = nothing to gate
    if (isSystemAdmin(u)) return true;   // system admin can see all
    // direct membership
    if (_normalizeId(u.CampaignID) === cid) return true;
    // explicit permission rows give view-level access automatically
    const grants = _loadUserGrants(u.ID);
    if (grants.campaignPrivs[cid]) return true;
    return false;
  } catch (e) {
    writeError && writeError('hasCampaignAccess', e);
    return false;
  }
}

// ---- the core evaluator ------------------------------------------------------
/**
 * Evaluate whether a user may view a page (route guard).
 * @returns { allow:boolean, reason:string, trace:string[] }
 */
function evaluatePageAccess(user, pageKey, campaignId) {
  const trace = [];
  try {
    const u = _normalizeUser(user);
    const page = _normalizePageKey(pageKey || '');
    const cid = _normalizeId(campaignId || '');

    // 0) hygiene & account state
    if (!u || !u.ID) return { allow: false, reason: 'No session', trace };
    if (!_truthy(u.CanLogin)) return { allow: false, reason: 'Account disabled', trace };
    if (u.LockoutEnd && _isFuture(u.LockoutEnd)) return { allow: false, reason: 'Account locked', trace };
    if (!ACCESS.PUBLIC_PAGES.has(page)) {
      if (!_truthy(u.EmailConfirmed)) return { allow: false, reason: 'Email not confirmed', trace };
      if (_truthy(u.ResetRequired)) return { allow: false, reason: 'Password reset required', trace };
    }

    // 1) public pages: always OK
    if (ACCESS.PUBLIC_PAGES.has(page)) {
      trace.push('PUBLIC page');
      return { allow: true, reason: 'public', trace };
    }

    // 2) page metadata sanity
    const meta = _getPageMeta(page);
    if (!meta.active) return { allow: false, reason: 'Page disabled', trace };

    // 3) admin-only pages
    if (ACCESS.ADMIN_ONLY_PAGES.has(page) || meta.requiresAdmin === true) {
      if (isSystemAdmin(u)) {
        trace.push('admin-only: allowed');
        return { allow: true, reason: 'admin', trace };
      }
      return { allow: false, reason: 'System Admin required', trace };
    }

    // 4) campaign gating (only if page is campaign-specific)
    if (meta.campaignSpecific) {
      if (!cid) return { allow: false, reason: 'Campaign required', trace };
      if (!hasCampaignAccess(u, cid)) return { allow: false, reason: 'No campaign access', trace };
    }

    // 5) page grants for NON-admins
    if (!isSystemAdmin(u)) {
      const grants = _loadUserGrants(u.ID);

      // 5a) user-assigned pages (global)
      if (grants.assignedPagesCanon.has(page)) {
        trace.push('assigned via Users.Pages (canonical)');
        return { allow: true, reason: 'assigned', trace };
      }

      // 5b) explicit per-campaign page grant
      if (cid && grants.pageGrants[cid] && grants.pageGrants[cid].has(page)) {
        trace.push('explicit campaign page grant');
        return { allow: true, reason: 'campaign-grant', trace };
      }

      // 5c) any explicit *global* page grant (campaignId='*')
      if (grants.pageGrants['*'] && grants.pageGrants['*'].has(page)) {
        trace.push('explicit global page grant');
        return { allow: true, reason: 'global-grant', trace };
      }

      return { allow: false, reason: 'Page not assigned', trace };
    }

    // 6) System Admin falls through to allow
    trace.push('System Admin: allow');
    return { allow: true, reason: 'admin', trace };

  } catch (e) {
    writeError && writeError('evaluatePageAccess', e);
    trace.push('exception:' + e.message);
    return { allow: false, reason: 'Evaluator error', trace };
  }
}

/**
 * Check if user has access to a specific campaign
 */
function checkUserCampaignAccess(userId, campaignId) {
  try {
    if (!userId || !campaignId) return false;

    // Get user's campaign permissions
    const permissions = getUserCampaignPermissions(userId);

    // Check if user has access to this campaign
    return permissions.some(perm =>
      String(perm.CampaignID || perm.campaignId) === String(campaignId) &&
      (perm.CanView || perm.View || perm.Access || perm.access)
    );
  } catch (error) {
    console.error('Error checking campaign access:', error);
    writeError('checkUserCampaignAccess', error);
    return false;
  }
}

function requireCampaignAwareAuth(e) {
  try {
    // 0) Auth service availability
    if (typeof AuthenticationService === 'undefined' || typeof AuthenticationService.requireAuth !== 'function') {
      return renderAccessDenied(getBaseUrl(e?.parameter?.token || ''), 'Authentication service unavailable');
    }

    // 1) Authentication handshake
    const auth = AuthenticationService.requireAuth(e);
    const rawToken = e?.parameter?.token || '';
    const baseUrl = getBaseUrl(rawToken);

    if (auth && typeof auth.getContent === 'function') {
      return rawToken
        ? renderAccessDenied(baseUrl, 'Your session is invalid or expired. Please sign in again.')
        : auth;
    }

    const user = auth;
    const pageParam = String(e?.parameter?.page || '').toLowerCase();
    const page = canonicalizePageKey(pageParam);   // ← see helper below
    const campaignId = String(e?.parameter?.campaign || user.CampaignID || '');

    const decision = evaluatePageAccess(user, page, campaignId);
    _debugAccess && _debugAccess('route', decision, user, page, campaignId); // OK if it doesn’t log PII

    if (!decision || decision.allow !== true) {
      return renderAccessDenied(baseUrl, (decision && decision.reason) || 'You do not have permission to view this page.');
    }

    if (campaignId) {
      try {
        if (typeof getCampaignNavigation === 'function') user.campaignNavigation = getCampaignNavigation(campaignId);
        if (typeof getUserCampaignPermissions === 'function') user.campaignPermissions = getUserCampaignPermissions(user.ID);
        if (typeof getCampaignById === 'function') {
          const c = getCampaignById(campaignId);
          user.campaignName = c ? (c.Name || c.name || '') : '';
        }
      } catch (ctxErr) {
        // Don’t block the page on context failures
        console.warn('Campaign context hydrate failed:', ctxErr);
      }
    }

    return user;

  } catch (error) {
    writeError('requireCampaignAwareAuth', error);
    return renderAccessDenied(getBaseUrl(e?.parameter?.token || ''), 'Authentication error occurred');
  }
}

function canonicalizePageKey(k) {
  const key = String(k || '').trim().toLowerCase();
  if (!key) return key;

  // Map legacy slugs/aliases → canonical keys used by the Access Engine
  switch (key) {
    case 'manageuser':
    case 'users':
      return 'admin.users';
    case 'manageroles':
    case 'roles':
      return 'admin.roles';
    case 'managecampaign':
    case 'campaigns':
      return 'admin.campaigns';
    case 'qa-dashboard':
    case 'ibtrqualityreports':
      return 'qa.dashboard';
    case 'qualityform':
    case 'independencequality':
    case 'creditsuiteqa':
      return 'qa.form';
    case 'qualityview':
      return 'qa.view';
    case 'qualitylist':
      return 'qa.list';
    case 'callreports':
      return 'reports.calls';
    case 'attendancereports':
      return 'reports.attendance';
    case 'tasklist':
      return 'tasks.list';
    case 'taskboard':
      return 'tasks.board';
    case 'taskform':
    case 'newtask':
    case 'edittask':
      return 'tasks.form';
    case 'eodreport':
      return 'tasks.eod';
    case 'coachingdashboard':
      return 'coaching.dashboard';
    case 'coachinglist':
    case 'coachings':
      return 'coaching.list';
    case 'coachingview':
      return 'coaching.view';
    case 'coachingsheet':
    case 'coaching':
      return 'coaching.form';
    case 'attendancecalendar':
      return 'calendar.attendance';
    case 'schedulemanagement':
      return 'schedule.management';
    case 'slotmanagement':
      return 'schedule.slots';
    case 'search':
      return 'global.search';
    case 'chat':
      return 'global.chat';
    case 'bookmarks':
      return 'global.bookmarks';
    case 'notifications':
      return 'global.notifications';
    case 'dashboard':
      return 'dashboard';
    default:
      return key; // unknowns fall through; let the engine decide
  }
}

/** Guard for *actions inside* a page (not navigation). */
function hasActionPrivilege(user, campaignId, privilegeKey) {
  try {
    const u = _normalizeUser(user);
    if (!u || !u.ID) return false;
    if (privilegeKey === ACCESS.PRIVS.SYSTEM_ADMIN) return isSystemAdmin(u);
    if (isSystemAdmin(u)) return true; // admin can do any action

    const cid = _normalizeId(campaignId);
    const grants = _loadUserGrants(u.ID);
    const p = grants.campaignPrivs[cid] || {};
    if (privilegeKey === ACCESS.PRIVS.MANAGE_USERS) return !!p.canManageUsers;
    if (privilegeKey === ACCESS.PRIVS.MANAGE_PAGES) return !!p.canManagePages;
    return false;
  } catch (e) {
    writeError && writeError('hasActionPrivilege', e);
    return false;
  }
}

/** Convenience: throws on failure for in-page actions. */
function requireActionPrivilegeOrThrow(user, campaignId, privilegeKey) {
  if (!hasActionPrivilege(user, campaignId, privilegeKey)) {
    throw new Error('Insufficient privileges: ' + privilegeKey + (campaignId ? (' for campaign ' + campaignId) : ''));
  }
}

/**
 * Get campaign navigation structure
 */
function getCampaignNavigation(campaignId) {
  try {
    if (typeof readSheet === 'function') {
      const navigation = readSheet('CampaignNavigation') || [];
      const campaignNav = navigation.filter(nav =>
        String(nav.CampaignID || nav.campaignId) === String(campaignId)
      );

      // Group by categories
      const categories = {};
      const uncategorizedPages = [];

      campaignNav.forEach(nav => {
        const category = nav.Category || nav.category || 'Uncategorized';
        if (category === 'Uncategorized') {
          uncategorizedPages.push(nav);
        } else {
          if (!categories[category]) {
            categories[category] = [];
          }
          categories[category].push(nav);
        }
      });

      return {
        categories: Object.entries(categories).map(([name, pages]) => ({
          name,
          pages
        })),
        uncategorizedPages
      };
    }

    // Default navigation structure
    return {
      categories: [
        {
          name: 'Main',
          pages: [
            { PageKey: 'dashboard', Name: 'Dashboard', Icon: 'dashboard' },
            { PageKey: 'reports', Name: 'Reports', Icon: 'bar_chart' }
          ]
        }
      ],
      uncategorizedPages: []
    };
  } catch (error) {
    console.error('Error getting campaign navigation:', error);
    writeError('getCampaignNavigation', error);
    return { categories: [], uncategorizedPages: [] };
  }
}

/**
 * Get campaign by ID
 */
function getCampaignById(campaignId) {
  try {
    const campaigns = getAllCampaigns();
    return campaigns.find(campaign =>
      String(campaign.ID || campaign.id) === String(campaignId)
    );
  } catch (error) {
    console.error('Error getting campaign by ID:', error);
    writeError('getCampaignById', error);
    return null;
  }
}

function checkUserPageAccess(userId, campaignId, pageKey) {
  try {
    const users = getAllUsers();
    const user = users.find(u => String(u.ID) === String(userId));
    if (!user) return false;
    return !!evaluatePageAccess(user, pageKey, campaignId).allow;
  } catch (e) {
    writeError('checkUserPageAccess', e);
    return false;
  }
}

/**
 * Get user's page permissions for a specific campaign
 */
function getUserPagePermissions(userId, campaignId) {
  try {
    if (typeof readSheet === 'function') {
      const permissions = readSheet('UserPagePermissions') || [];
      return permissions.filter(perm =>
        String(perm.UserID || perm.userId) === String(userId) &&
        String(perm.CampaignID || perm.campaignId) === String(campaignId)
      );
    }

    return [];
  } catch (error) {
    console.error('Error getting user page permissions:', error);
    writeError('getUserPagePermissions', error);
    return [];
  }
}

/**
 * Get base URL with token
 */
function getBaseUrl(token) {
  return `${SCRIPT_URL}?token=${encodeURIComponent(token || '')}`;
}

// ───────────────────────────────────────────────────────────────────────────────
// ENHANCED doGet WITH DYNAMIC ROUTING
// ───────────────────────────────────────────────────────────────────────────────

function doGet(e) {
  try {
    const rawToken = e.parameter.token || "";
    const baseUrl = getBaseUrl(rawToken);

    // Initialize system components
    initializeSystem();

    // ───────────────────────────────────────────────────────────────────────────
    // ENHANCED PROXY HANDLER
    // ───────────────────────────────────────────────────────────────────────────
    if (e.parameter.page === 'proxy') {
      console.log('doGet: Handling enhanced proxy request');
      return serveEnhancedProxy(e);
    }

    // ───────────────────────────────────────────────────────────────────────────
    // LOGOUT HANDLER
    // ───────────────────────────────────────────────────────────────────────────
    if (e.parameter.action === 'logout') {
      console.log('doGet: Handling logout action');
      try {
        return handleLogoutRequest(e);
      } catch (error) {
        console.error('doGet: Logout handler error:', error);
        return HtmlService.createHtmlOutput(`
                    <!DOCTYPE html>
                    <html>
                    <head>
                        <title>Logged Out</title>
                        <meta name="viewport" content="width=device-width,initial-scale=1">
                    </head>
                    <body>
                        <h1>Logged Out</h1>
                        <p>You have been logged out. Redirecting...</p>
                        <script>
                            setTimeout(function() {
                                window.location.href = "${SCRIPT_URL}";
                            }, 2000);
                        </script>
                    </body>
                    </html>
                `).setTitle('Logged Out')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }
    }

    // ───────────────────────────────────────────────────────────────────────────
    // EMAIL CONFIRMATION HANDLER
    // ───────────────────────────────────────────────────────────────────────────
    if (e.parameter.action === 'confirmEmail' && e.parameter.token) {
      const confirmationToken = e.parameter.token;
      const success = confirmEmail(confirmationToken);

      if (success) {
        const redirectUrl = `${baseUrl}&page=setpassword&token=${encodeURIComponent(confirmationToken)}`;
        return HtmlService
          .createHtmlOutput(`<script>window.location.href = "${redirectUrl}";</script>`)
          .setTitle('Redirecting...');
      } else {
        const tpl = HtmlService.createTemplateFromFile('EmailConfirmed');
        tpl.baseUrl = baseUrl;
        tpl.success = false;
        tpl.token = confirmationToken;
        return tpl.evaluate()
          .setTitle('Email Confirmation')
          .addMetaTag('viewport', 'width=device-width,initial-scale=1');
      }
    }

    // ───────────────────────────────────────────────────────────────────────────
    // PUBLIC PAGES (No Authentication Required)
    // ───────────────────────────────────────────────────────────────────────────
    const page = (e.parameter.page || "").toLowerCase();

    if (page === 'login' || (!rawToken && !page)) {
      const tpl = HtmlService.createTemplateFromFile('Login');
      tpl.baseUrl = baseUrl;
      tpl.rawToken = rawToken;
      tpl.scriptUrl = SCRIPT_URL;

      return tpl.evaluate()
        .setTitle('Login - VLBPO LuminaHQ')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // Handle other public pages
    const publicPages = ['setpassword', 'resetpassword', 'resend-verification', 'resendverification',
      'forgotpassword', 'forgot-password', 'emailconfirmed', 'email-confirmed'];

    if (publicPages.includes(page)) {
      return handlePublicPage(page, e, baseUrl, rawToken);
    }

    // ───────────────────────────────────────────────────────────────────────────
    // PROTECTED PAGES (Authentication Required)
    // ───────────────────────────────────────────────────────────────────────────

    // Use enhanced authentication
    const auth = requireCampaignAwareAuth(e);
    if (auth.getContent) {
      return auth; // HtmlOutput: Login or Forbidden
    }

    const user = auth;

    // Add role names to user object
    const rolesMap = getRolesMapping();
    user.roleNames = (user.roles || []).map(id => rolesMap[id]);

    // Handle password reset requirement
    if (_truthy(user.ResetRequired)) {
      const tpl = HtmlService.createTemplateFromFile('ChangePassword');
      tpl.baseUrl = baseUrl;
      tpl.rawToken = rawToken;
      tpl.scriptUrl = SCRIPT_URL;

      return tpl.evaluate()
        .setTitle('Change Password')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // DEFAULT ROUTE: Redirect to user's campaign dashboard
    // ─────────────────────────────────────────────────────────────────────────
    if (rawToken && !page) {
      const userCampaignId = user.CampaignID || '';

      const redirectUrl = userCampaignId
        ? `${baseUrl}&page=dashboard&campaign=${userCampaignId}`
        : `${baseUrl}&page=dashboard`;

      return HtmlService
        .createHtmlOutput(`<script>window.location.href = "${redirectUrl}";</script>`)
        .setTitle('Redirecting to Dashboard...');
    }

    // ───────────────────────────────────────────────────────────────────────────
    // CAMPAIGN-SPECIFIC PAGE ROUTING
    // ───────────────────────────────────────────────────────────────────────────

    const campaignId = e.parameter.campaign || user.CampaignID || '';

    // Handle CSV exports
    if (page === "callreports" && e.parameter.action === "exportCallCsv") {
      const gran = e.parameter.granularity || "Week";
      const period = e.parameter.period || weekStringFromDate(new Date());
      const agent = e.parameter.agent || "";
      const csv = exportCallAnalyticsCsv(gran, period, agent);
      return ContentService.createTextOutput(csv)
        .setMimeType(ContentService.MimeType.CSV)
        .downloadAsFile(`callAnalytics_${period}.csv`);
    }

    if (page === 'attendancereports' && e.parameter.action === 'exportCsv') {
      const gran = e.parameter.granularity || 'Week';
      const period = e.parameter.period || weekStringFromDate(new Date());
      const agent = e.parameter.agent || '';
      const csv = exportAttendanceCsv(gran, period, agent);
      return ContentService.createTextOutput(csv)
        .setMimeType(ContentService.MimeType.CSV)
        .downloadAsFile(`attendance_${period}.csv`);
    }

    if (page === 'attendancereports' && e.parameter.action === 'exportCsvEnhanced') {
      try {
        const gran = e.parameter.granularity || 'Week';
        const period = e.parameter.period || weekStringFromDate(new Date());
        const agents = e.parameter.agents || '';
        const format = e.parameter.format || 'combined';
        const includeBreaks = e.parameter.includeBreaks === 'true';
        const includeLunch = e.parameter.includeLunch === 'true';
        const includeWeekends = e.parameter.includeWeekends === 'true';

        const exportOptions = {
          format: format,
          includeBreaks: includeBreaks,
          includeLunch: includeLunch,
          includeWeekends: includeWeekends
        };

        // Generate the CSV using the enhanced export function
        const csv = exportAttendanceCsvEnhanced(gran, period, agents, exportOptions);

        // Determine filename based on selection
        let filename = `attendance_${gran.toLowerCase()}_${period}`;

        if (agents && agents !== '') {
          const agentList = agents.split(',');
          if (agentList.length === 1) {
            filename += `_${agentList[0].replace(/[^a-zA-Z0-9]/g, '_')}`;
          } else {
            filename += `_${agentList.length}_agents`;
          }
        } else {
          filename += '_all_agents';
        }

        filename += `_${format}.csv`;

        return ContentService.createTextOutput(csv)
          .setMimeType(ContentService.MimeType.CSV)
          .downloadAsFile(filename);

      } catch (error) {
        console.error('Enhanced CSV export error:', error);
        writeError('enhancedCsvExport', error);

        // Return error as plain text
        return ContentService.createTextOutput(
          `Error generating export: ${error.message}\n\nPlease try again or contact support.`
        ).setMimeType(ContentService.MimeType.TEXT);
      }
    }
    // Enhanced routing with campaign awareness
    return routeToPage(page, e, baseUrl, user, campaignId);

  } catch (error) {
    console.error('Error in doGet:', error);
    writeError('doGet', error);
    return createErrorPage('System Error', `An error occurred: ${error.message}`);
  }
}



// ───────────────────────────────────────────────────────────────────────────────
// PAGE ROUTING LOGIC
// ───────────────────────────────────────────────────────────────────────────────

function routeToPage(page, e, baseUrl, user, campaignId) {
  try {
    switch (page) {
      case "dashboard":
        return serveCampaignPage('Dashboard', e, baseUrl, user, campaignId);

      // QA System Routes
      case "qualityform":
        return routeQAForm(e, baseUrl, user, campaignId);

      case "ibtrqualityreports":
        return routeQADashboard(e, baseUrl, user, campaignId);

      case 'qualityview':
        return routeQAView(e, baseUrl, user, campaignId);

      case 'qualitylist':
        return routeQAList(e, baseUrl, user, campaignId);

      // Independence QA specific routes
      case "indpcoachingform":
        return serveCampaignPage('IdependenceCoachingForm', e, baseUrl, user, campaignId);

      case "indpcoachingform":
        return serveCampaignPage('IdependenceCoachingForm', e, baseUrl, user, campaignId);

      case "independencequality":
        return serveCampaignPage('IndependenceQAForm', e, baseUrl, user, campaignId);

      case "independenceqadashboard":
        return serveCampaignPage('UnifiedQADashboard', e, baseUrl, user, campaignId);

      // Credit Suite QA specific routes
      case "creditsuiteqa":
        return serveCampaignPage('CreditSuiteQAForm', e, baseUrl, user, campaignId);

      case "qa-dashboard":
        return serveCampaignPage('CreditSuiteQADashboard', e, baseUrl, user, campaignId);

      // Call and Attendance Reports
      case "callreports":
        return serveCampaignPage('CallReports', e, baseUrl, user, campaignId);

      case "attendancereports":
        return serveCampaignPage('AttendanceReports', e, baseUrl, user, campaignId);

      // Coaching System
      case "coachingdashboard":
        return serveCampaignPage('CoachingDashboard', e, baseUrl, user, campaignId);

      case "coachingview":
        return serveCampaignPage('CoachingView', e, baseUrl, user, campaignId);

      case "coachinglist":
      case "coachings":
        return serveCampaignPage('CoachingList', e, baseUrl, user, campaignId);

      case "coachingsheet":
      case "coaching":
        return serveCampaignPage('CoachingForm', e, baseUrl, user, campaignId);

      // ── Task Management Pages ───────────────────────────────────────────────
      case "tasklist":
        return serveTaskPage('TaskList', e, baseUrl, user, campaignId);

      case "taskboard":
        return serveTaskPage('TaskBoard', e, baseUrl, user, campaignId);

      case "taskform":
      case "newtask":
      case "edittask":
        return serveTaskPage('TaskForm', e, baseUrl, user, campaignId);

      case "eodreport":
        return serveTaskPage('EODReport', e, baseUrl, user, campaignId);

      // Schedule and Calendar
      case "schedulemanagement":
        return serveCampaignPage('ScheduleManagement', e, baseUrl, user, campaignId);

      case "attendancecalendar":
        return serveCampaignPage('Calendar', e, baseUrl, user, campaignId);

      // Other Campaign Pages
      case "escalations":
        return serveCampaignPage('Escalations', e, baseUrl, user, campaignId);

      case "settings":
        return serveCampaignPage('Settings', e, baseUrl, user, campaignId);

      case "incentives":
        return serveCampaignPage('Incentives', e, baseUrl, user, campaignId);

      // Global Pages
      case 'search':
        return serveGlobalPage('Search', e, baseUrl, user);

      case 'chat':
        return serveGlobalPage('Chat', e, baseUrl, user);

      case 'bookmarks':
        return serveGlobalPage('BookmarkManager', e, baseUrl, user);

      case 'notifications':
        return serveGlobalPage('Notifications', e, baseUrl, user);

      // Admin Pages
      case 'manageuser':
        return serveAdminPage('Users', e, baseUrl, user);

      case 'manageroles':
        return serveAdminPage('RoleManagement', e, baseUrl, user);

      case 'managecampaign':
        return serveAdminPage('CampaignManagement', e, baseUrl, user);

      case 'import':
        return serveAdminPage('ImportCsv', e, baseUrl, user);

      case 'importattendance':
        return serveAdminPage('ImportAttendance', e, baseUrl, user);

      case 'slotmanagement':
        return serveShiftSlotManagement(e, baseUrl, user, campaignId);

      // Special Pages
      case 'ackform':
        return serveAckForm(e, baseUrl, user);

      case 'proxy':
        return serveProxy(e);

      default:
        // For unknown pages, redirect to dashboard
        const defaultCampaignId = user.CampaignID || '';
        const redirectUrl = defaultCampaignId
          ? `${baseUrl}&page=dashboard&campaign=${defaultCampaignId}`
          : `${baseUrl}&page=dashboard`;

        return HtmlService
          .createHtmlOutput(`<script>window.location.href = "${redirectUrl}";</script>`)
          .setTitle('Redirecting to Dashboard...');
    }
  } catch (error) {
    console.error(`Error routing to page ${page}:`, error);
    writeError('routeToPage', error);
    return createErrorPage('Page Error', `Error loading page: ${error.message}`);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// QA SYSTEM ROUTING
// ───────────────────────────────────────────────────────────────────────────────

function routeQAForm(e, baseUrl, user, campaignId) {
  // Determine which QA form to use based on campaign
  const campaign = getCampaignById(campaignId);
  const campaignName = campaign?.Name?.toLowerCase() || '';

  if (campaignName.includes('independence') || campaignName.includes('insurance')) {
    return serveCampaignPage('IndependenceQAForm', e, baseUrl, user, campaignId);
  } else if (campaignName.includes('credit') || campaignName.includes('suite')) {
    return serveCampaignPage('CreditSuiteQAForm', e, baseUrl, user, campaignId);
  } else {
    // Default QA form
    return serveCampaignPage('QualityForm', e, baseUrl, user, campaignId);
  }
}

function routeQADashboard(e, baseUrl, user, campaignId) {
  // Determine which QA dashboard to use based on campaign
  const campaign = getCampaignById(campaignId);
  const campaignName = campaign?.Name?.toLowerCase() || '';

  if (campaignName.includes('independence') || campaignName.includes('insurance')) {
    return serveCampaignPage('UnifiedQADashboard', e, baseUrl, user, campaignId);
  } else if (campaignName.includes('credit') || campaignName.includes('suite')) {
    return serveCampaignPage('CreditSuiteQADashboard', e, baseUrl, user, campaignId);
  } else {
    // Default QA dashboard
    return serveCampaignPage('QADashboard', e, baseUrl, user, campaignId);
  }
}

function routeQAView(e, baseUrl, user, campaignId) {
  // Determine which QA view to use based on campaign
  const campaign = getCampaignById(campaignId);
  const campaignName = campaign?.Name?.toLowerCase() || '';

  if (campaignName.includes('independence') || campaignName.includes('insurance')) {
    return serveCampaignPage('IndependenceQAView', e, baseUrl, user, campaignId);
  } else if (campaignName.includes('credit') || campaignName.includes('suite')) {
    return serveCampaignPage('CreditSuiteQAView', e, baseUrl, user, campaignId);
  } else {
    // Default QA view
    return serveCampaignPage('QualityView', e, baseUrl, user, campaignId);
  }
}

function routeQAList(e, baseUrl, user, campaignId) {
  // Determine which QA list to use based on campaign
  const campaign = getCampaignById(campaignId);
  const campaignName = campaign?.Name?.toLowerCase() || '';

  if (campaignName.includes('independence') || campaignName.includes('insurance')) {
    return serveCampaignPage('IndependenceQAList', e, baseUrl, user, campaignId);
  } else if (campaignName.includes('credit') || campaignName.includes('suite')) {
    return serveCampaignPage('CreditSuiteQAList', e, baseUrl, user, campaignId);
  } else {
    // Default QA list
    return serveCampaignPage('QAList', e, baseUrl, user, campaignId);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// PAGE SERVING FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Serve task-related pages with proper data and permissions
 */
function serveTaskPage(templateName, e, baseUrl, user, campaignId) {
  try {
    // Check if user has access to tasks (you can customize this logic)
    if (!user) {
      return redirectToLogin(baseUrl);
    }

    const tpl = HtmlService.createTemplateFromFile(templateName);
    tpl.baseUrl = baseUrl;
    tpl.rawToken = e.parameter.token || '';
    tpl.scriptUrl = baseUrl.split('?')[0];
    tpl.currentPage = getPageTitle(templateName);
    tpl.user = user;
    tpl.campaignId = campaignId;

    // Add template-specific data
    addTaskTemplateData(tpl, templateName, e, user, campaignId);

    return tpl.evaluate()
      .setTitle(tpl.currentPage + ' - VLBPO DataLog')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    console.error(`Error serving task page ${templateName}:`, error);
    writeError(`serveTaskPage(${templateName})`, error);
    return createErrorPage('Failed to load page', error.message);
  }
}

/**
 * Add template-specific data for task pages
 * (Aligned with task logic; uses getEODTasksByDate for EOD.)
 */
function addTaskTemplateData(tpl, templateName, e, user, campaignId) {
  try {
    // Get all tasks for the user or campaign
    const allTasks = getAllTasks();
    let userTasks = [];

    // Filter tasks based on user permissions
    if (isUserAdmin(user)) {
      userTasks = allTasks; // Admins see all tasks
    } else {
      userTasks = getTasksFor(user.Email || user.UserName);
    }

    // Get list of users for delegation
    const users = getAllUsers();
    const userEmails = users.map(u => u.Email || u.UserName).filter(Boolean);

    switch (templateName) {
      case 'TaskList':
        tpl.tasks = JSON.stringify(userTasks).replace(/<\/script>/g, '<\\/script>');
        tpl.tasksJson = tpl.tasks; // Alias for backward compatibility
        tpl.users = userEmails;
        tpl.totalTasks = userTasks.length;
        tpl.completedTasks = userTasks.filter(t => t.Status === 'Completed' || t.Status === 'Done').length;
        break;

      case 'TaskBoard':
        tpl.tasks = JSON.stringify(userTasks).replace(/<\/script>/g, '<\\/script>');
        tpl.tasksJson = tpl.tasks;
        tpl.owners = JSON.stringify([...new Set(userTasks.map(t => t.Owner))]).replace(/<\/script>/g, '<\\/script>');
        tpl.ownersJson = tpl.owners;
        tpl.selectedOwner = e.parameter.owner || '';
        break;

      case 'TaskForm':
        const recordId = e.parameter.id || '';
        let record = {};

        if (recordId) {
          try {
            record = getTaskById(recordId);
          } catch (err) {
            console.warn('Task not found:', recordId);
          }
        }

        tpl.recordId = recordId;
        tpl.record = JSON.stringify(record).replace(/<\/script>/g, '<\\/script>');
        tpl.tasks = JSON.stringify(userTasks).replace(/<\/script>/g, '<\\/script>');
        tpl.tasksJson = tpl.tasks;
        tpl.users = userEmails;
        break;

      case 'EODReport':
        const reportDate = e.parameter.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        const eodTasks = (typeof getEODTasksByDate === 'function') ? getEODTasksByDate(reportDate) : [];

        tpl.reportDate = reportDate;
        tpl.eodTasks = JSON.stringify(eodTasks).replace(/<\/script>/g, '<\\/script>');
        tpl.users = userEmails;
        break;

      default:
        // Default data for any task page
        tpl.tasks = JSON.stringify(userTasks).replace(/<\/script>/g, '<\\/script>');
        tpl.users = userEmails;
        break;
    }

  } catch (error) {
    console.error(`Error adding task template data for ${templateName}:`, error);
    writeError(`addTaskTemplateData(${templateName})`, error);

    // Provide safe defaults
    tpl.tasks = '[]';
    tpl.tasksJson = '[]';
    tpl.users = [];
    tpl.recordId = '';
    tpl.record = '{}';
  }
}

/**
 * Get page title from template name
 */
function getPageTitle(templateName) {
  const titles = {
    'TaskList': 'Task List',
    'TaskBoard': 'Task Board',
    'TaskForm': 'Task Form',
    'EODReport': 'EOD Report'
  };

  return titles[templateName] || templateName;
}

/**
 * Serve a campaign-specific page
 */
function serveCampaignPage(templateName, e, baseUrl, user, campaignId) {
  try {
    const tpl = HtmlService.createTemplateFromFile(templateName);
    tpl.baseUrl = baseUrl;
    tpl.rawToken = e.parameter.token || '';
    tpl.scriptUrl = SCRIPT_URL;
    tpl.currentPage = templateName.replace(/([A-Z])/g, ' $1').trim();
    tpl.user = user;
    tpl.campaignId = campaignId;

    __injectCurrentUser_(tpl, user); // ← add

    if (campaignId) {
      tpl.campaignName = (tpl.user && tpl.user.campaignName) || '';
      tpl.campaignNavigation = (tpl.user && tpl.user.campaignNavigation) || { categories: [], uncategorizedPages: [] };
    }

    addTemplateSpecificData(tpl, templateName, e, tpl.user, campaignId);
    return tpl.evaluate()
      .setTitle(tpl.currentPage + ' - VLBPO LuminaHQ')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    writeError(`serveCampaignPage(${templateName})`, error);
    return createErrorPage('Failed to load page', error.message);
  }
}


/**
 * Serve a global page (accessible from any campaign)
 */
function serveGlobalPage(templateName, e, baseUrl, user) {
  try {
    const tpl = HtmlService.createTemplateFromFile(templateName);
    tpl.baseUrl = baseUrl;
    tpl.rawToken = e.parameter.token || '';
    tpl.scriptUrl = SCRIPT_URL;
    tpl.currentPage = templateName.replace(/([A-Z])/g, ' $1').trim();
    tpl.user = user;

    __injectCurrentUser_(tpl, user); // ← add

    addTemplateSpecificData(tpl, templateName, e, tpl.user, null);
    return tpl.evaluate()
      .setTitle(tpl.currentPage + ' - VLBPO LuminaHQ')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    writeError(`serveGlobalPage(${templateName})`, error);
    return createErrorPage('Failed to load page', error.message);
  }
}

function serveShiftSlotManagement(e, baseUrl, user, campaignId) {
  try {
    // Check if user has permission to manage shift slots
    // You might want to add specific permission checking here
    if (!isUserAdmin(user) && !hasPermission(user, 'manage_shifts')) {
      return renderAccessDenied(baseUrl, 'You do not have permission to manage shift slots.');
    }

    const tpl = HtmlService.createTemplateFromFile('SlotManagementInterface');
    tpl.baseUrl = baseUrl;
    tpl.rawToken = e.parameter.token || '';
    tpl.scriptUrl = SCRIPT_URL;
    tpl.currentPage = 'Shift Slot Management';
    tpl.user = user;
    tpl.campaignId = campaignId;

    // Add any additional template data
    if (campaignId) {
      tpl.campaignName = user.campaignName || '';
    }

    return tpl.evaluate()
      .setTitle('Shift Slot Management - VLBPO LuminaHQ')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    console.error('Error serving shift slot management page:', error);
    writeError('serveShiftSlotManagement', error);
    return createErrorPage('Failed to load page', error.message);
  }
}

function hasPermission(user, permission) {
  try {
    // Check if user has specific permission
    // This is a placeholder - implement based on your permission system
    const userRoles = user.Roles || user.roles || '';

    // Admins have all permissions
    if (isUserAdmin(user)) return true;

    // Check specific permissions based on your role system
    switch (permission) {
      case 'manage_shifts':
        return userRoles.toLowerCase().includes('manager') ||
          userRoles.toLowerCase().includes('supervisor') ||
          userRoles.toLowerCase().includes('scheduler');
      default:
        return false;
    }
  } catch (error) {
    console.error('Error checking permission:', error);
    return false;
  }
}

/**
 * Serve an admin page (requires admin permissions)
 */
function serveAdminPage(templateName, e, baseUrl, user) {
  if (!isSystemAdmin(user)) return renderAccessDenied(baseUrl, 'This page requires System Admin privileges.');
  try {
    // Check admin permissions
    if (!isUserAdmin(user)) {
      const tpl = HtmlService.createTemplateFromFile('AccessDenied');
      tpl.baseUrl = baseUrl;
      tpl.message = 'This page requires administrator privileges.';
      return tpl.evaluate()
        .setTitle('Access Denied')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    const tpl = HtmlService.createTemplateFromFile(templateName);
    tpl.baseUrl = baseUrl;
    tpl.rawToken = e.parameter.token || '';
    tpl.scriptUrl = SCRIPT_URL;
    tpl.currentPage = templateName.replace(/([A-Z])/g, ' $1').trim();
    tpl.user = user;

    // Add admin page specific data
    addTemplateSpecificData(tpl, templateName, e, user, null);

    return tpl.evaluate()
      .setTitle(tpl.currentPage + ' - VLBPO LuminaHQ')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    console.error(`Error serving admin page ${templateName}:`, error);
    writeError(`serveAdminPage(${templateName})`, error);
    return createErrorPage('Failed to load page', error.message);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// TEMPLATE DATA HANDLER WITH IMPROVED ERROR HANDLING
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Add template-specific data based on the template name
 * FIXED: Now properly passes campaignId to all handlers
 */
function addTemplateSpecificData(tpl, templateName, e, user, campaignId) {
  try {
    const timeZone = Session.getScriptTimeZone();

    // Consistent variable definitions with proper defaults
    const granularity = (e.parameter.granularity || "Week").toString();
    const periodValue = (e.parameter.period || weekStringFromDate(new Date())).toString();
    const selectedAgent = (e.parameter.agent || "").toString();
    const selectedCampaign = (e.parameter.campaign || campaignId || "").toString();
    const selectedDepartment = (e.parameter.department || "").toString();

    // Set common template properties consistently
    tpl.granularity = granularity;
    tpl.period = periodValue;
    tpl.periodValue = periodValue;
    tpl.selectedPeriod = periodValue;
    tpl.selectedAgent = selectedAgent;
    tpl.selectedCampaign = selectedCampaign;
    tpl.selectedDepartment = selectedDepartment;
    tpl.timeZone = timeZone;

    // Template-specific data handling - FIXED: now passes campaignId
    handleTemplateSpecificData(tpl, templateName, e, user, campaignId);

  } catch (error) {
    console.error(`Error adding template data for ${templateName}:`, error);
    writeError(`addTemplateSpecificData(${templateName})`, error);
  }
}

function handleTemplateSpecificData(tpl, templateName, e, user, campaignId) {
  try {
    switch (templateName) {
      case 'Dashboard':
        handleDashboardData(tpl, e, user, campaignId);
        break;

      case 'AttendanceReports':
        handleAttendanceReportsData(tpl, e, user, campaignId);
        break;

      case 'CallReports':
        handleCallReportsData(tpl, e, user, campaignId);
        break;

      case 'QualityForm':
      case 'IndependenceQAForm':
      case 'CreditSuiteQAForm':
        handleQAFormData(tpl, e, user, templateName, campaignId);
        break;

      case 'QualityView':
      case 'IndependenceQAView':
      case 'CreditSuiteQAView':
        handleQAViewData(tpl, e, user, templateName);
        break;

      case 'QADashboard':
      case 'UnifiedQADashboard':
      case 'CreditSuiteQADashboard':
        handleQADashboardData(tpl, e, user, templateName, campaignId);
        break;

      case 'QAList':
      case 'IndependenceQAList':
      case 'CreditSuiteQAList':
        handleQAListData(tpl, e, user, templateName);
        break;

      case 'Users':
        handleUsersData(tpl, e, user);
        break;

      case 'TaskForm':
      case 'TaskList':
      case 'TaskBoard':
        handleTaskData(tpl, e, user, templateName);
        break;

      case 'CoachingForm':
      case 'CoachingList':
      case 'CoachingView':
      case 'CoachingDashboard':
        handleCoachingData(tpl, e, user, templateName, campaignId);
        break;

      case 'EODReport':
        handleEODReportData(tpl, e, user);
        break;

      case 'Calendar':
        handleCalendarData(tpl, e, user, campaignId);
        break;

      case 'Search':
        handleSearchData(tpl, e);
        break;

      case 'Chat':
        handleChatData(tpl, e, user);
        break;

      case 'Notifications':
        handleNotificationsData(tpl, e, user);
        break;

      case 'Escalations':
        handleEscalationsData(tpl, e, user, campaignId);
        break;

      case 'CoachingAckForm':
        handleCoachingAckData(tpl, e, user);
        break;

      default:
        console.log(`No specific data handler for template: ${templateName}`);
        break;
    }
  } catch (error) {
    console.error(`Error handling data for ${templateName}:`, error);
    writeError(`handleTemplateSpecificData(${templateName})`, error);
  }
}

// Individual template data handlers
function handleDashboardData(tpl, e, user, campaignId) {
  try {
    const granularity = e.parameter.granularity || "Week";
    const periodValue = e.parameter.period || weekStringFromDate(new Date());
    const selectedAgent = e.parameter.agent || "";

    if (typeof getDashboardOkrs === 'function') {
      const okrData = getDashboardOkrs(granularity, periodValue, selectedAgent);
      tpl.okrData = JSON.stringify(okrData).replace(/<\/script>/g, '<\\/script>');
    } else {
      tpl.okrData = JSON.stringify({});
    }

    // FIXED: Pass campaignId properly
    tpl.userList = clientGetAssignedAgentNames(campaignId || '');
  } catch (error) {
    console.error('Error handling dashboard data:', error);
    tpl.okrData = JSON.stringify({});
    tpl.userList = [];
  }
}

function handleAttendanceReportsData(tpl, e, user, campaignId) {
  try {
    const granularity = e.parameter.granularity || "Week";
    const periodValue = e.parameter.period || weekStringFromDate(new Date());
    const selectedAgent = e.parameter.agent || "";

    if (typeof getAttendanceAnalyticsByPeriod === 'function') {
      const attendanceAnalytics = getAttendanceAnalyticsByPeriod(granularity, periodValue, selectedAgent);
      const allRows = attendanceAnalytics.filteredRows || [];
      const rowPage = parseInt(e.parameter.rowPage, 10) || 1;
      const PAGE_SIZE = 50;
      const startRowIdx = (rowPage - 1) * PAGE_SIZE;
      attendanceAnalytics.filteredRows = allRows.slice(startRowIdx, startRowIdx + PAGE_SIZE);

      tpl.attendanceData = JSON.stringify(attendanceAnalytics).replace(/<\/script>/g, '<\\/script>');
      tpl.currentRowPage = rowPage;
      tpl.totalRows = allRows.length;
      tpl.PAGE_SIZE = PAGE_SIZE;
    } else {
      tpl.attendanceData = JSON.stringify({ filteredRows: [], summary: {} });
      tpl.currentRowPage = 1;
      tpl.totalRows = 0;
      tpl.PAGE_SIZE = 50;
    }

    // FIXED: Pass campaignId properly and add fallback
    tpl.userList = clientGetAssignedAgentNames(campaignId || user.CampaignID || '');
    
    // ADDITIONAL: Add debug logging
    console.log('AttendanceReports userList length:', tpl.userList.length);
    console.log('CampaignId used:', campaignId || user.CampaignID || '');
    
  } catch (error) {
    console.error('Error handling attendance reports data:', error);
    writeError('handleAttendanceReportsData', error);
    tpl.attendanceData = JSON.stringify({ filteredRows: [], summary: {} });
    tpl.userList = [];
  }
}

function handleQAFormData(tpl, e, user, templateName) {
  try {
    const qaId = (e.parameter.id || "").toString();
    tpl.recordId = qaId;

    // Determine which users function to call based on template
    if (templateName === 'IndependenceQAForm') {
      if (typeof clientGetIndependenceUsers === 'function') {
        tpl.users = clientGetIndependenceUsers();
      } else {
        tpl.users = getIndependenceInsuranceUsers();
      }
      tpl.campaignName = 'Independence Insurance';
    } else if (templateName === 'CreditSuiteQAForm') {
      if (typeof clientGetCreditSuiteUsers === 'function') {
        tpl.users = clientGetCreditSuiteUsers();
      } else {
        tpl.users = getUsers();
      }
      tpl.campaignName = 'Credit Suite';
    } else {
      tpl.users = getUsers();
    }

    // Get existing record if editing
    if (qaId) {
      if (templateName === 'IndependenceQAForm' && typeof clientGetIndependenceQAById === 'function') {
        const record = clientGetIndependenceQAById(qaId);
        tpl.record = record ? JSON.stringify(record).replace(/<\/script>/g, '<\\/script>') : "{}";
      } else if (templateName === 'CreditSuiteQAForm' && typeof clientGetCreditSuiteQAById === 'function') {
        const record = clientGetCreditSuiteQAById(qaId);
        tpl.record = record ? JSON.stringify(record).replace(/<\/script>/g, '<\\/script>') : "{}";
      } else if (typeof getQARecordById === 'function') {
        const record = getQARecordById(qaId);
        tpl.record = record ? JSON.stringify(record).replace(/<\/script>/g, '<\\/script>') : "{}";
      } else {
        tpl.record = "{}";
      }
    } else {
      tpl.record = "{}";
    }
  } catch (error) {
    console.error('Error handling QA form data:', error);
    tpl.users = [];
    tpl.recordId = "";
    tpl.record = "{}";
  }
}

/**
 * Client-accessible function for Credit Suite users
 */
function clientGetCreditSuiteUsers(campaignId = null) {
  try {
    return getCreditSuiteUsers(campaignId);
  } catch (error) {
    console.error('Error in clientGetCreditSuiteUsers:', error);
    writeError('clientGetCreditSuiteUsers', error);
    return [];
  }
}

/**
 * Get Credit Suite users - campaign-aware version
 */
function getCreditSuiteUsers(campaignId = null) {
  try {
    let users;

    if (campaignId) {
      users = getUsersByCampaign(campaignId);
    } else {
      const allUsers = getAllUsers();
      users = allUsers.filter(user => {
        const campaignName = user.CampaignName || user.campaignName || '';
        const userCampaignId = user.CampaignID || user.campaignId || '';

        return campaignName.toLowerCase().includes('credit') ||
          campaignName.toLowerCase().includes('suite') ||
          userCampaignId === 'credit-suite' ||
          userCampaignId.toLowerCase().includes('credit');
      });
    }

    return users
      .map(user => ({
        name: user.FullName || user.UserName || user.name,
        email: user.Email || user.email,
        id: user.ID || user.id
      }))
      .filter(user => user.name && user.email)
      .sort((a, b) => a.name.localeCompare(b.name));

  } catch (error) {
    console.error('Error getting Credit Suite users:', error);
    writeError('getCreditSuiteUsers', error);
    return [];
  }
}

function handleQAViewData(tpl, e, user, templateName) {
  try {
    const qaViewId = (e.parameter.id || '').toString();
    tpl.recordId = qaViewId;

    if (!qaViewId) {
      tpl.record = "{}";
      return;
    }

    if (templateName === 'IndependenceQAView' && typeof clientGetIndependenceQAById === 'function') {
      const qaRec = clientGetIndependenceQAById(qaViewId);
      tpl.record = JSON.stringify(qaRec || {}).replace(/<\/script>/g, '<\\/script>');
      tpl.campaignName = 'Independence Insurance';
    } else if (templateName === 'CreditSuiteQAView' && typeof clientGetCreditSuiteQAById === 'function') {
      const qaRec = clientGetCreditSuiteQAById(qaViewId);
      tpl.record = JSON.stringify(qaRec || {}).replace(/<\/script>/g, '<\\/script>');
      tpl.campaignName = 'Credit Suite';
    } else if (typeof getQARecordById === 'function') {
      const qaRec = getQARecordById(qaViewId);
      tpl.record = JSON.stringify(qaRec || {}).replace(/<\/script>/g, '<\\/script>');
    } else {
      tpl.record = "{}";
    }
  } catch (error) {
    console.error('Error handling QA view data:', error);
    tpl.record = "{}";
    tpl.recordId = "";
  }
}

function handleQADashboardData(tpl, e, user, templateName) {
  try {
    const granularity = e.parameter.granularity || "Week";
    const periodValue = e.parameter.period || weekStringFromDate(new Date());
    const selectedAgent = e.parameter.agent || "";

    // Handle different QA dashboard types
    if (templateName === 'UnifiedQADashboard') {
      // For unified dashboard, we need both campaign data
      tpl.userList = clientGetAssignedAgentNames(campaignId || '');

      // Get Independence analytics
      try {
        if (typeof clientGetIndependenceQAAnalytics === 'function') {
          const independenceAnalytics = clientGetIndependenceQAAnalytics(granularity, periodValue, selectedAgent, "");
          tpl.independenceAnalytics = JSON.stringify(independenceAnalytics).replace(/<\/script>/g, '<\\/script>');
        } else {
          tpl.independenceAnalytics = JSON.stringify(getEmptyQAAnalytics());
        }
      } catch (error) {
        console.error('Error getting Independence analytics:', error);
        tpl.independenceAnalytics = JSON.stringify(getEmptyQAAnalytics());
      }

      // Get Credit Suite analytics
      try {
        if (typeof clientGetCreditSuiteQAAnalytics === 'function') {
          const creditSuiteAnalytics = clientGetCreditSuiteQAAnalytics(granularity, periodValue, selectedAgent, "");
          tpl.creditSuiteAnalytics = JSON.stringify(creditSuiteAnalytics).replace(/<\/script>/g, '<\\/script>');
        } else {
          tpl.creditSuiteAnalytics = JSON.stringify(getEmptyQAAnalytics());
        }
      } catch (error) {
        console.error('Error getting Credit Suite analytics:', error);
        tpl.creditSuiteAnalytics = JSON.stringify(getEmptyQAAnalytics());
      }

      // Set initial analytics (Independence by default)
      tpl.qaAnalytics = tpl.independenceAnalytics;

    } else if (templateName === 'CreditSuiteQADashboard') {
      // Credit Suite specific dashboard
      tpl.userList = clientGetAssignedAgentNames(campaignId || '');
      tpl.campaignName = 'Credit Suite';

      try {
        if (typeof clientGetCreditSuiteQAAnalytics === 'function') {
          const analytics = clientGetCreditSuiteQAAnalytics(granularity, periodValue, selectedAgent);
          tpl.qaAnalytics = JSON.stringify(analytics).replace(/<\/script>/g, '<\\/script>');
        } else {
          tpl.qaAnalytics = JSON.stringify(getEmptyQAAnalytics());
        }
      } catch (error) {
        console.error('Error getting Credit Suite QA analytics:', error);
        tpl.qaAnalytics = JSON.stringify(getEmptyQAAnalytics());
      }

    } else {
      // Default QA dashboard
      try {
        if (typeof getAllQA === 'function') {
          const qaRecords = getAllQA();
          const qaUserList = getUsers().map(u => u.name || u.FullName || u.UserName).filter(n => n).sort();

          tpl.userList = qaUserList;
          tpl.qaRecords = JSON.stringify(qaRecords).replace(/<\/script>/g, '<\\/script>');
        } else {
          tpl.userList = [];
          tpl.qaRecords = JSON.stringify([]);
        }
      } catch (error) {
        console.error('Error getting QA dashboard data:', error);
        tpl.userList = [];
        tpl.qaRecords = JSON.stringify([]);
      }
    }

  } catch (error) {
    console.error('Error handling QA dashboard data:', error);
    tpl.userList = [];
    tpl.qaAnalytics = JSON.stringify(getEmptyQAAnalytics());
  }
}

function handleQAListData(tpl, e, user, templateName) {
  try {
    if (templateName === 'IndependenceQAList') {
      let qaRecords = [];
      if (typeof clientGetIndependenceQARecords === 'function') {
        qaRecords = clientGetIndependenceQARecords();
      }
      tpl.qaRecords = JSON.stringify(qaRecords).replace(/<\/script>/g, '<\\/script>');
      tpl.campaignName = 'Independence Insurance';

    } else if (templateName === 'CreditSuiteQAList') {
      let qaRecords = [];
      if (typeof clientGetCreditSuiteQARecords === 'function') {
        qaRecords = clientGetCreditSuiteQARecords();
      }
      tpl.qaRecords = JSON.stringify(qaRecords).replace(/<\/script>/g, '<\\/script>');
      tpl.campaignName = 'Credit Suite';

    } else {
      // Default QA list
      if (typeof getAllQA === 'function') {
        tpl.qaRecords = JSON.stringify(getAllQA()).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.qaRecords = JSON.stringify([]);
      }
    }
  } catch (error) {
    console.error('Error handling QA list data:', error);
    tpl.qaRecords = JSON.stringify([]);
  }
}

function handleUsersData(tpl, e, user) {
  try {
    tpl.campaignList = getAllCampaigns();

    const roleList = getAllRoles() || [];
    tpl.roleList = roleList.filter(r => {
      const n = String(r.name || r.Name || '').toLowerCase();
      return n !== 'super admin' && n !== 'administrator';
    });

    const knownPages = Object.keys(_loadAllPagesMeta());
    tpl.pagesList = knownPages; // the multiselect source
  } catch (error) {
    writeError('handleUsersData', error);
    tpl.campaignList = [];
    tpl.roleList = [];
    tpl.pagesList = [];
  }
}


function handleTaskData(tpl, e, user, templateName) {
  try {
    const currentUserEmail = user.email || user.Email || '';

    if (templateName === 'TaskForm') {
      tpl.currentUser = currentUserEmail;
      tpl.users = getUsers().map(u => u.email || u.Email).filter(e => e);

      if (typeof getTasksFor === 'function') {
        const visible = getTasksFor(currentUserEmail);
        tpl.tasksJson = JSON.stringify(visible).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.tasksJson = JSON.stringify([]);
      }

      const taskId = (e.parameter.id || '').toString();
      tpl.recordId = taskId;

      if (taskId && typeof getTaskById === 'function') {
        tpl.record = JSON.stringify(getTaskById(taskId)).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.record = '{}';
      }

    } else if (templateName === 'TaskList') {
      tpl.currentUser = currentUserEmail;

      if (typeof getTasksFor === 'function') {
        tpl.tasks = JSON.stringify(getTasksFor(currentUserEmail)).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.tasks = JSON.stringify([]);
      }

      tpl.users = JSON.stringify(getUsers().map(u => u.email || u.Email).filter(e => e));
      tpl.recordId = '';
      tpl.record = '{}';

    } else if (templateName === 'TaskBoard') {
      tpl.currentUser = currentUserEmail;

      if (typeof getTasksFor === 'function') {
        const visibleTasks = getTasksFor(currentUserEmail);
        tpl.tasksJson = JSON.stringify(visibleTasks).replace(/<\/script>/g, '<\\/script>');

        const ownersArr = [...new Set(visibleTasks.map(t => t.Owner).filter(x => x))].sort();
        tpl.ownersJson = JSON.stringify(ownersArr);
        tpl.selectedOwner = e.parameter.owner || '';
      } else {
        tpl.tasksJson = JSON.stringify([]);
        tpl.ownersJson = JSON.stringify([]);
        tpl.selectedOwner = '';
      }
    }

  } catch (error) {
    console.error('Error handling task data:', error);
    tpl.currentUser = user.email || user.Email || '';
    tpl.users = [];
    tpl.tasksJson = JSON.stringify([]);
    tpl.recordId = '';
    tpl.record = '{}';
  }
}

function handleCoachingData(tpl, e, user, templateName) {
  try {
    const granularity = e.parameter.granularity || "Week";
    const periodValue = e.parameter.period || weekStringFromDate(new Date());
    const selectedAgent = e.parameter.agent || "";

    if (templateName === 'CoachingForm') {
      // No additional data needed beyond base template data

    } else if (templateName === 'CoachingList') {
      if (typeof getAllCoaching === 'function') {
        tpl.coachingRecords = JSON.stringify(getAllCoaching()).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.coachingRecords = JSON.stringify([]);
      }

    } else if (templateName === 'CoachingView') {
      const coachingId = (e.parameter.id || '').toString();
      tpl.recordId = coachingId;

      if (coachingId && typeof getCoachingRecordById === 'function') {
        const coachingRec = getCoachingRecordById(coachingId);
        tpl.record = JSON.stringify(coachingRec || {}).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.record = '{}';
      }

    } else if (templateName === 'CoachingDashboard') {
      if (typeof getDashboardCoaching === 'function') {
        const coachingMetrics = getDashboardCoaching(granularity, periodValue, selectedAgent);
        const coachingUserList = getUsers().map(u => u.name || u.FullName || u.UserName).filter(n => n).sort();

        tpl.userList = coachingUserList;
        tpl.sessionsTrend = JSON.stringify(coachingMetrics.sessionsTrend || []);
        tpl.topicsDist = JSON.stringify(coachingMetrics.topicsDist || []);
        tpl.upcomingFollow = JSON.stringify(coachingMetrics.upcoming || []);
      } else {
        tpl.userList = [];
        tpl.sessionsTrend = JSON.stringify([]);
        tpl.topicsDist = JSON.stringify([]);
        tpl.upcomingFollow = JSON.stringify([]);
      }
    }

  } catch (error) {
    console.error('Error handling coaching data:', error);
    tpl.userList = [];
    tpl.sessionsTrend = JSON.stringify([]);
    tpl.topicsDist = JSON.stringify([]);
    tpl.upcomingFollow = JSON.stringify([]);
  }
}

function handleEODReportData(tpl, e, user) {
  try {
    const today = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'
    );

    if (typeof getEODTasksByDate === 'function') {
      const allComplete = getEODTasksByDate(today);
      const me = (user.email || user.Email || '').toLowerCase();
      const visibleEOD = allComplete.filter(t =>
        String(t.Owner || '').toLowerCase() === me ||
        String(t.Delegations || '')
          .split(',')
          .map(x => x.trim().toLowerCase())
          .includes(me)
      );

      tpl.reportDate = today;
      tpl.tasks = JSON.stringify(visibleEOD).replace(/<\/script>/g, '<\\/script>');
    } else {
      tpl.reportDate = today;
      tpl.tasks = JSON.stringify([]);
    }
  } catch (error) {
    console.error('Error handling EOD report data:', error);
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    tpl.reportDate = today;
    tpl.tasks = JSON.stringify([]);
  }
}

function handleCalendarData(tpl, e, user) {
  try {
    const selectedAgent = e.parameter.agent || "";

    if (e.parameter.page === 'attendancecalendar') {
      // Attendance Calendar
      if (typeof getAttendanceAnalyticsByPeriod === 'function') {
        const today = new Date();
        const isoWeek = weekStringFromDate(today);
        const attendanceAnalytics = getAttendanceAnalyticsByPeriod("Week", isoWeek, selectedAgent);

        tpl.userList = clientGetAssignedAgentNames(campaignId || '');
        tpl.attendanceData = JSON.stringify(attendanceAnalytics).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.userList = clientGetAssignedAgentNames(campaignId || '');
        tpl.attendanceData = JSON.stringify({});
      }
    } else {
      tpl.userList = clientGetAssignedAgentNames(campaignId || '');
    }
  } catch (error) {
    console.error('Error handling calendar data:', error);
    tpl.userList = [];
    tpl.attendanceData = JSON.stringify({});
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// ENHANCED SEARCH DATA HANDLER
// ───────────────────────────────────────────────────────────────────────────────

function handleSearchData(tpl, e) {
  try {
    const query = (e.parameter.query || '').trim();
    const pageIndex = parseInt(e.parameter.pageIndex || '1', 10);
    const startIdx = (pageIndex - 1) * 10 + 1;
    let items = [], totalResults = 0, error = null;

    if (query) {
      try {
        if (typeof searchWeb === 'function') {
          const resp = searchWeb(query, startIdx);
          items = resp.items || [];
          totalResults = parseInt(resp.searchInformation.totalResults || '0', 10);

          // Enhance search results with proxy information
          items = items.map(item => ({
            ...item,
            proxyUrl: `${SCRIPT_URL}?page=proxy&url=${encodeURIComponent(item.link)}`,
            displayTitle: item.htmlTitle || item.title || 'Untitled',
            displaySnippet: item.htmlSnippet || item.snippet || 'No description available'
          }));
        }
      } catch (err) {
        error = err.message;
      }
    }

    tpl.query = query;
    tpl.results = items;
    tpl.totalResults = totalResults;
    tpl.pageIndex = pageIndex;
    tpl.error = error;
    tpl.scriptUrl = SCRIPT_URL;

  } catch (error) {
    console.error('Error handling search data:', error);
    tpl.query = '';
    tpl.results = [];
    tpl.totalResults = 0;
    tpl.pageIndex = 1;
    tpl.error = error.message;
    tpl.scriptUrl = SCRIPT_URL;
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// ENHANCED BOOKMARK FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Client-accessible function to get user bookmarks
 */
function clientGetUserBookmarks() {
  try {
    // This function should get bookmarks for the current user
    // Implement based on your bookmark storage system
    if (typeof readSheet === 'function') {
      const bookmarks = readSheet('UserBookmarks') || [];
      return bookmarks.map(bookmark => ({
        ...bookmark,
        proxyUrl: `${SCRIPT_URL}?page=proxy&url=${encodeURIComponent(bookmark.URL)}`
      }));
    }
    return [];
  } catch (error) {
    console.error('Error getting user bookmarks:', error);
    writeError('clientGetUserBookmarks', error);
    return [];
  }
}

/**
 * Client-accessible function to add bookmark
 */
function clientAddBookmark(title, url, description, tags, folder) {
  try {
    const bookmark = {
      ID: Utilities.getUuid(),
      Title: title,
      URL: url,
      Description: description || '',
      Tags: tags || '',
      Folder: folder || 'General',
      Created: new Date(),
      LastAccessed: new Date(),
      AccessCount: 0
    };

    if (typeof writeSheet === 'function') {
      writeSheet('UserBookmarks', [bookmark]);
      return { success: true, bookmark: bookmark };
    }

    return { success: false, error: 'Bookmark service not available' };
  } catch (error) {
    console.error('Error adding bookmark:', error);
    writeError('clientAddBookmark', error);
    return { success: false, error: error.message };
  }
}

/**
 * Client-accessible function to delete bookmark
 */
function clientDeleteBookmark(bookmarkId) {
  try {
    if (typeof readSheet === 'function' && typeof writeSheet === 'function') {
      const bookmarks = readSheet('UserBookmarks') || [];
      const filteredBookmarks = bookmarks.filter(b => b.ID !== bookmarkId);

      // Rewrite the sheet with filtered data
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName('UserBookmarks');
      if (sheet) {
        sheet.clear();
        if (filteredBookmarks.length > 0) {
          const headers = Object.keys(filteredBookmarks[0]);
          sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
          const data = filteredBookmarks.map(b => headers.map(h => b[h]));
          sheet.getRange(2, 1, data.length, headers.length).setValues(data);
        }
      }
      return true;
    }
    return false;
  } catch (error) {
    console.error('Error deleting bookmark:', error);
    writeError('clientDeleteBookmark', error);
    return false;
  }
}

function getUsersByCampaign(campaignId) {
  try {
    var rows = (typeof readSheet === 'function') ? (readSheet(USERS_SHEET) || []) : [];
    return rows.filter(function (r) { return String(r.CampaignID || r.campaignId) === String(campaignId); });
  } catch (_) { return []; }
}

/**
 * Client-accessible function to update bookmark access
 */
function clientUpdateBookmarkAccess(bookmarkId) {
  try {
    if (typeof readSheet === 'function' && typeof writeSheet === 'function') {
      const bookmarks = readSheet('UserBookmarks') || [];
      const bookmark = bookmarks.find(b => b.ID === bookmarkId);

      if (bookmark) {
        bookmark.LastAccessed = new Date();
        bookmark.AccessCount = (bookmark.AccessCount || 0) + 1;

        // Update the bookmark in the sheet
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('UserBookmarks');
        if (sheet) {
          const data = sheet.getDataRange().getValues();
          const headers = data[0];
          const idIndex = headers.indexOf('ID');
          const accessIndex = headers.indexOf('LastAccessed');
          const countIndex = headers.indexOf('AccessCount');

          for (let i = 1; i < data.length; i++) {
            if (data[i][idIndex] === bookmarkId) {
              if (accessIndex >= 0) sheet.getRange(i + 1, accessIndex + 1).setValue(bookmark.LastAccessed);
              if (countIndex >= 0) sheet.getRange(i + 1, countIndex + 1).setValue(bookmark.AccessCount);
              break;
            }
          }
        }
      }
      return true;
    }
    return false;
  } catch (error) {
    console.error('Error updating bookmark access:', error);
    writeError('clientUpdateBookmarkAccess', error);
    return false;
  }
}

/**
 * Client-accessible function to get bookmark folders
 */
function clientGetUserBookmarkFolders() {
  try {
    if (typeof readSheet === 'function') {
      const bookmarks = readSheet('UserBookmarks') || [];
      const folders = [...new Set(bookmarks.map(b => b.Folder).filter(f => f))];
      return folders.length > 0 ? folders : ['General', 'Work', 'Reference', 'Tools'];
    }
    return ['General', 'Work', 'Reference', 'Tools'];
  } catch (error) {
    console.error('Error getting bookmark folders:', error);
    return ['General', 'Work', 'Reference', 'Tools'];
  }
}

function handleChatData(tpl, e, user) {
  try {
    const groupId = e.parameter.groupId || '';
    const channelId = e.parameter.channelId || '';
    tpl.groupId = groupId;
    tpl.channelId = channelId;

    // Safely get chat data with error handling
    let userGroups = [];
    let groupChannels = [];
    let messages = [];

    try {
      if (typeof getUserChatGroups === 'function') {
        userGroups = getUserChatGroups(user.ID) || [];
      }
    } catch (e) {
      console.warn('Error getting user groups:', e);
    }

    try {
      if (groupId && typeof getChatChannels === 'function') {
        groupChannels = getChatChannels(groupId);
      }
    } catch (e) {
      console.warn('Error getting group channels:', e);
    }

    try {
      if (channelId && typeof getChatMessages === 'function') {
        messages = getChatMessages(channelId);
      }
    } catch (e) {
      console.warn('Error getting messages:', e);
    }

    tpl.userGroups = JSON.stringify(userGroups)
      .replace(/\\/g, '\\\\')
      .replace(/'/g, '\\\'')
      .replace(/<\/script>/g, '<\\/script>');
    tpl.groupChannels = JSON.stringify(groupChannels)
      .replace(/\\/g, '\\\\')
      .replace(/'/g, '\\\'')
      .replace(/<\/script>/g, '<\\/script>');
    tpl.messages = JSON.stringify(messages)
      .replace(/\\/g, '\\\\')
      .replace(/'/g, '\\\'')
      .replace(/<\/script>/g, '<\\/script>');
  } catch (error) {
    console.error('Error handling chat data:', error);
    tpl.groupId = '';
    tpl.channelId = '';
    tpl.userGroups = JSON.stringify([]);
    tpl.groupChannels = JSON.stringify([]);
    tpl.messages = JSON.stringify([]);
  }
}

function handleNotificationsData(tpl, e, user) {
  try {
    if (typeof getAllTasks === 'function') {
      const notifToday = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const dueTasks = getAllTasks()
        .filter(t => (t.Status !== 'Done' && t.Status !== 'Completed') && t.DueDate === notifToday);
      tpl.tasksJson = JSON.stringify(dueTasks).replace(/<\/script>/g, '<\\/script>');
    } else {
      tpl.tasksJson = JSON.stringify([]);
    }
  } catch (error) {
    console.error('Error handling notifications data:', error);
    tpl.tasksJson = JSON.stringify([]);
  }
}

function handleEscalationsData(tpl, e, user) {
  try {
    tpl.userList = clientGetAssignedAgentNames(campaignId || '');
  } catch (error) {
    console.error('Error handling escalations data:', error);
    tpl.userList = [];
  }
}

function handleCoachingAckData(tpl, e, user) {
  try {
    const ackId = (e.parameter.id || "").toString();
    tpl.recordId = ackId;
    tpl.record = ackId && typeof getCoachingRecordById === 'function'
      ? JSON.stringify(getCoachingRecordById(ackId)).replace(/<\/script>/g, '<\\/script>')
      : '{}';
  } catch (error) {
    console.error('Error handling coaching ack data:', error);
    tpl.recordId = '';
    tpl.record = '{}';
  }
}

function handleCallReportsData(tpl, e, user, campaignId) {
  try {
    const granularity = e.parameter.granularity || "Week";
    const periodValue = e.parameter.period || weekStringFromDate(new Date());
    const selectedAgent = e.parameter.agent || "";

    if (typeof getAnalyticsByPeriod === 'function') {
      const attendanceAnalytics = getAttendanceAnalyticsByPeriod(granularity, periodValue, selectedAgent);
      const analytics = getAnalyticsByPeriod("Week", periodValue, "");
      const rawReps = analytics.repMetrics || [];
      const pageNum = parseInt(e.parameter.page, 10) || 1;
      const PAGE_SIZE = 50;
      const startPageIdx = (pageNum - 1) * PAGE_SIZE;
      const pageSlice = rawReps.slice(startPageIdx, startPageIdx + PAGE_SIZE);
      attendanceAnalytics.filteredRows = allRows.slice(startRowIdx, startRowIdx + PAGE_SIZE);

      const formattedReps = pageSlice.map((r) => {
        const totalCalls = r.totalCalls || 0;
        const totalTalkDecimal = parseFloat(r.totalTalk) || 0;
        const totalTalkFormatted = formatDuration ? formatDuration(totalTalkDecimal) : totalTalkDecimal;
        const avgTalkDecimal = totalCalls > 0 ? totalTalkDecimal / totalCalls : 0;
        const avgTalkFormatted = formatDuration ? formatDuration(avgTalkDecimal) : avgTalkDecimal;
        return {
          agent: r.agent,
          totalCalls: totalCalls,
          totalTalkFormatted: totalTalkFormatted,
          avgTalkFormatted: avgTalkFormatted,
        };
      });

      tpl.PAGE_SIZE = PAGE_SIZE;
      tpl.data = formattedReps;
      
      // FIXED: Use the improved user list function
      tpl.userList = clientGetAssignedAgentNames(campaignId || user.CampaignID || '');

      // Chart data
      tpl.callVolumeLast7 = JSON.stringify(analytics.callTrend || []);
      tpl.hourlyHeatmapLast7 = JSON.stringify([]);
      tpl.avgIntervalByAgentLast7 = JSON.stringify([]);
      tpl.talkTimeByAgentLast7 = JSON.stringify(analytics.talkTrend || []);
      tpl.wrapupCountsLast7 = JSON.stringify(analytics.wrapDist || []);
      tpl.csatDistLast7 = JSON.stringify(analytics.csatDist || []);
      tpl.policyCountsLast7 = JSON.stringify(analytics.policyDist || []);
      tpl.agentLeaderboardLast7 = JSON.stringify(analytics.repMetrics || []);
    } else {
      tpl.PAGE_SIZE = 50;
      tpl.data = [];
      tpl.userList = clientGetAssignedAgentNames(campaignId || user.CampaignID || '');
      tpl.callVolumeLast7 = JSON.stringify([]);
      tpl.hourlyHeatmapLast7 = JSON.stringify([]);
      tpl.avgIntervalByAgentLast7 = JSON.stringify([]);
      tpl.talkTimeByAgentLast7 = JSON.stringify([]);
      tpl.wrapupCountsLast7 = JSON.stringify([]);
      tpl.csatDistLast7 = JSON.stringify([]);
      tpl.policyCountsLast7 = JSON.stringify([]);
      tpl.agentLeaderboardLast7 = JSON.stringify([]);
    }
  } catch (error) {
    console.error('Error handling call reports data:', error);
    writeError('handleCallReportsData', error);
    tpl.PAGE_SIZE = 50;
    tpl.data = [];
    tpl.userList = [];
    // Set empty arrays for all chart data...
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// UTILITY FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Redirect to login page
 */
function redirectToLogin(baseUrl) {
  const loginUrl = baseUrl.split('&')[0] + '&page=login';
  return HtmlService
    .createHtmlOutput(`<script>window.location.href = "${loginUrl}";</script>`)
    .setTitle('Redirecting to Login...');
}

/**
 * Check if tasks module is properly initialized
 * (Uses configured spreadsheet if available; falls back to active.)
 */
function isTasksModuleInitialized() {
  try {
    const ss = (typeof getIBTRSpreadsheet === 'function')
      ? getIBTRSpreadsheet()
      : SpreadsheetApp.getActiveSpreadsheet();
    const tasksSheet = ss.getSheetByName(TASKS_SHEET);
    const commentsSheet = ss.getSheetByName(COMMENTS_SHEET);
    return tasksSheet !== null && commentsSheet !== null;
  } catch (error) {
    return false;
  }
}

/**
 * Auto-initialize tasks module if needed
 */
function ensureTasksModuleInitialized() {
  if (!isTasksModuleInitialized()) {
    console.log('Tasks module not initialized, initializing now...');
    initializeTasks();
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// BACKWARD COMPATIBILITY FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

// These functions provide backward compatibility with existing code

function getTasks() {
  return getAllTasks();
}

function getTasksForUser(userEmail) {
  return getTasksFor(userEmail);
}

function createTask(taskData) {
  return addOrUpdateTask(taskData);
}

function updateTask(taskData) {
  return addOrUpdateTask(taskData);
}

function removeTask(taskId) {
  return deleteTask(taskId);
}

// ───────────────────────────────────────────────────────────────────────────────
// TRIGGER FUNCTIONS (for automatic execution)
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Called by time-based trigger to send EOD emails
 */
function triggerSendEODEmail() {
  try {
    sendEODEmail();
  } catch (error) {
    console.error('Error in triggerSendEODEmail:', error);
    writeError('triggerSendEODEmail', error);
  }
}

/**
 * Called by time-based trigger to generate recurring tasks
 */
function triggerGenerateRecurringTasks() {
  try {
    generateRecurringTasks();
  } catch (error) {
    console.error('Error in triggerGenerateRecurringTasks:', error);
    writeError('triggerGenerateRecurringTasks', error);
  }
}

/**
 * Set up all necessary triggers for the tasks module
 */
function setupTaskTriggers() {
  try {
    // Delete existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'triggerSendEODEmail' ||
        trigger.getHandlerFunction() === 'triggerGenerateRecurringTasks') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    // Create new triggers
    ScriptApp.newTrigger('triggerSendEODEmail')
      .timeBased()
      .everyDays(1)
      .atHour(18) // 6 PM
      .create();

    ScriptApp.newTrigger('triggerGenerateRecurringTasks')
      .timeBased()
      .everyDays(1)
      .atHour(0) // Midnight
      .create();

    console.log('Task triggers set up successfully');

  } catch (error) {
    console.error('Error setting up task triggers:', error);
    writeError('setupTaskTriggers', error);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// ADMIN FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Admin function to reset all tasks (use with caution)
 */
function adminResetAllTasks() {
  try {
    if (!isUserAdmin(Session.getActiveUser().getEmail())) {
      throw new Error('Admin privileges required');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tasksSheet = ss.getSheetByName(TASKS_SHEET);
    const commentsSheet = ss.getSheetByName(COMMENTS_SHEET);

    if (tasksSheet) {
      tasksSheet.clear();
      tasksSheet.getRange(1, 1, 1, TASKS_HEADERS.length).setValues([TASKS_HEADERS]);
    }

    if (commentsSheet) {
      commentsSheet.clear();
      commentsSheet.getRange(1, 1, 1, 4).setValues([['TaskID', 'Timestamp', 'Author', 'Text']]);
    }

    console.log('All tasks reset successfully');
    return { success: true };

  } catch (error) {
    console.error('Error resetting tasks:', error);
    writeError('adminResetAllTasks', error);
    return { success: false, error: error.message };
  }
}

/**
 * Admin function to export all tasks data
 */
function adminExportAllTasksData() {
  try {
    if (!isUserAdmin(Session.getActiveUser().getEmail())) {
      throw new Error('Admin privileges required');
    }

    const tasks = getAllTasks();
    const comments = readSheet(COMMENTS_SHEET);

    return {
      success: true,
      data: {
        tasks: tasks,
        comments: comments,
        exportDate: new Date().toISOString()
      }
    };

  } catch (error) {
    console.error('Error exporting tasks data:', error);
    writeError('adminExportAllTasksData', error);
    return { success: false, error: error.message };
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// CLIENT-ACCESSIBLE TASK FUNCTIONS (aligned to task logic)
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Client-accessible function to get all tasks
 */
function clientGetAllTasks() {
  try {
    return getAllTasks();
  } catch (error) {
    console.error('Error in clientGetAllTasks:', error);
    writeError('clientGetAllTasks', error);
    return [];
  }
}

/**
 * Client-accessible function to get tasks for a specific user
 */
function clientGetTasksFor(userEmail) {
  try {
    return getTasksFor(userEmail);
  } catch (error) {
    console.error('Error in clientGetTasksFor:', error);
    writeError('clientGetTasksFor', error);
    return [];
  }
}

/**
 * Client-accessible function to add or update a task
 */
function clientAddOrUpdateTask(formData) {
  try {
    return addOrUpdateTask(formData);
  } catch (error) {
    console.error('Error in clientAddOrUpdateTask:', error);
    writeError('clientAddOrUpdateTask', error);
    return { success: false, error: error.message };
  }
}

/**
 * Client-accessible function to delete a task
 */
function clientDeleteTask(taskId) {
  try {
    return deleteTask(taskId);
  } catch (error) {
    console.error('Error in clientDeleteTask:', error);
    writeError('clientDeleteTask', error);
    return { success: false, error: error.message };
  }
}

/**
 * Client-accessible function to get a task by ID
 */
function clientGetTaskById(taskId) {
  try {
    return getTaskById(taskId);
  } catch (error) {
    console.error('Error in clientGetTaskById:', error);
    writeError('clientGetTaskById', error);
    return null;
  }
}

/**
 * Client-accessible function to snooze a task
 */
function clientSnoozeTask(taskId, newDueDate) {
  try {
    return snoozeTask(taskId, newDueDate);
  } catch (error) {
    console.error('Error in clientSnoozeTask:', error);
    writeError('clientSnoozeTask', error);
    return { success: false, error: error.message };
  }
}

/**
 * Client-accessible function to add a comment to a task
 */
function clientAddTaskComment(taskId, author, text) {
  try {
    return addComment(taskId, author, text);
  } catch (error) {
    console.error('Error in clientAddTaskComment:', error);
    writeError('clientAddTaskComment', error);
    return { success: false, error: error.message };
  }
}

/**
 * Client-accessible function to get comments for a task
 */
function clientGetTaskComments(taskId) {
  try {
    return getComments(taskId);
  } catch (error) {
    console.error('Error in clientGetTaskComments:', error);
    writeError('clientGetTaskComments', error);
    return [];
  }
}

/**
 * Client-accessible function to export tasks
 * (Graceful fallback if exportTasks is not defined in task module)
 */
function clientExportTasks(format = 'CSV') {
  try {
    if (typeof exportTasks === 'function') {
      return exportTasks(format);
    }
    // Fallback export (CSV) using existing data:
    const tasks = getAllTasks();
    const headers = Object.keys(tasks[0] || {}).length ? Object.keys(tasks[0]) : TASKS_HEADERS;
    const rows = [headers].concat(tasks.map(t => headers.map(h => t[h] ?? '')));
    const csv = rows.map(r => r.map(v => `"${String(v instanceof Date ? v.toISOString() : v).replace(/"/g, '""')}"`).join(',')).join('\n');
    const blob = Utilities.newBlob(csv, 'text/csv', 'tasks.csv');
    let folder;
    try {
      folder = DriveApp.getFolderById(TASKS_FOLDER_ID);
    } catch (e) {
      folder = DriveApp.createFolder('Task Attachments');
    }
    const name = `tasks_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}.csv`;
    const file = folder.createFile(blob).setName(name);
    return { success: true, fileId: file.getId(), url: file.getUrl() };
  } catch (error) {
    console.error('Error in clientExportTasks:', error);
    writeError('clientExportTasks', error);
    return { success: false, error: error.message };
  }
}

/**
 * Client-accessible function to get task notifications count
 * (Compute inline so it doesn't depend on an external helper.)
 */
function clientGetTaskNotifications() {
  try {
    const me = (Session.getActiveUser().getEmail() || '').toLowerCase();
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const today = new Date(todayStr);

    const tasks = (typeof getTasksFor === 'function' && me)
      ? getTasksFor(me)
      : getAllTasks();

    const count = tasks.filter(t => {
      const status = String(t.Status || '').toLowerCase();
      if (status === 'completed' || status === 'done') return false;
      if (!t.DueDate) return false;
      const due = new Date(t.DueDate);
      return !isNaN(due) && due <= today;
    }).length;

    return count;
  } catch (error) {
    console.error('Error in clientGetTaskNotifications:', error);
    writeError('clientGetTaskNotifications', error);
    return 0;
  }
}

/**
 * Client-accessible function to generate EOD report (unchanged)
 */
function clientGenerateEODReport(date, teamMember = '') {
  try {
    const tasks = getAllTasks();

    // Filter tasks by date and team member
    const filteredTasks = tasks.filter(task => {
      const matchesDate = !date ||
        task.CompletedDate === date ||
        task.StartDate === date ||
        task.DueDate === date;
      const matchesMember = !teamMember || task.Owner === teamMember;
      return matchesDate && matchesMember;
    });

    // Categorize tasks
    const completedTasks = filteredTasks.filter(t =>
      t.Status === 'Completed' || t.Status === 'Done'
    );
    const inProgressTasks = filteredTasks.filter(t =>
      t.Status === 'In Progress'
    );
    const overdueTasks = filteredTasks.filter(t => {
      if (t.Status === 'Completed' || t.Status === 'Done') return false;
      return t.DueDate && new Date(t.DueDate) < new Date(date);
    });

    // Create team summaries
    const teamSummaries = {};
    filteredTasks.forEach(task => {
      if (!teamSummaries[task.Owner]) {
        teamSummaries[task.Owner] = {
          total: 0,
          completed: 0,
          inProgress: 0,
          overdue: 0
        };
      }

      teamSummaries[task.Owner].total++;

      if (task.Status === 'Completed' || task.Status === 'Done') {
        teamSummaries[task.Owner].completed++;
      } else if (task.Status === 'In Progress') {
        teamSummaries[task.Owner].inProgress++;
      } else if (task.DueDate && new Date(task.DueDate) < new Date(date)) {
        teamSummaries[task.Owner].overdue++;
      }
    });

    return {
      success: true,
      data: {
        date: date,
        teamMember: teamMember,
        totalTasks: filteredTasks.length,
        completedTasks: completedTasks,
        inProgressTasks: inProgressTasks,
        overdueTasks: overdueTasks,
        teamSummaries: teamSummaries
      }
    };

  } catch (error) {
    console.error('Error in clientGenerateEODReport:', error);
    writeError('clientGenerateEODReport', error);
    return { success: false, error: error.message };
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// INITIALIZATION AND SETUP FUNCTIONS (tasks)
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Initialize the tasks module (call this once to set up)
 */
function initializeTasks() {
  try {
    console.log('Initializing Tasks Module...');

    // Initialize tasks module
    if (typeof initializeTasksModule === 'function') {
      initializeTasksModule();
    } else if (typeof setupTasksSheets === 'function') {
      setupTasksSheets();
    }

    console.log('Tasks Module initialized successfully');
    return { success: true };

  } catch (error) {
    console.error('Error initializing Tasks Module:', error);
    writeError('initializeTasks', error);
    return { success: false, error: error.message };
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// PUBLIC PAGES / PROXY / BOOKMARK SHEETS INIT
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Handle public pages
 */
function handlePublicPage(page, e, baseUrl, rawToken) {
  const scriptUrl = SCRIPT_URL;

  switch (page) {
    case 'setpassword':
    case 'resetpassword':
      const resetToken = e.parameter.token || '';
      const tpl = HtmlService.createTemplateFromFile('ChangePassword');
      tpl.baseUrl = baseUrl;
      tpl.rawToken = rawToken;
      tpl.scriptUrl = scriptUrl;
      tpl.token = resetToken;
      tpl.isReset = page === 'resetpassword';

      return tpl.evaluate()
        .setTitle((page === 'resetpassword' ? 'Reset' : 'Set') + ' Your Password - VLBPO LuminaHQ')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    case 'resend-verification':
    case 'resendverification':
      const verifyTpl = HtmlService.createTemplateFromFile('ResendVerification');
      verifyTpl.baseUrl = baseUrl;
      verifyTpl.rawToken = rawToken;
      verifyTpl.scriptUrl = scriptUrl;

      return verifyTpl.evaluate()
        .setTitle('Resend Email Verification - VLBPO LuminaHQ')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    case 'forgotpassword':
    case 'forgot-password':
      const forgotTpl = HtmlService.createTemplateFromFile('ForgotPassword');
      forgotTpl.baseUrl = baseUrl;
      forgotTpl.rawToken = rawToken;
      forgotTpl.scriptUrl = scriptUrl;

      return forgotTpl.evaluate()
        .setTitle('Forgot Password - VLBPO LuminaHQ')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    case 'emailconfirmed':
    case 'email-confirmed':
      const confirmTpl = HtmlService.createTemplateFromFile('EmailConfirmed');
      confirmTpl.baseUrl = baseUrl;
      confirmTpl.success = e.parameter.success === 'true';
      confirmTpl.token = e.parameter.token || '';

      return confirmTpl.evaluate()
        .setTitle('Email Confirmation - VLBPO LuminaHQ')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    default:
      return createErrorPage('Page Not Found', `The page "${page}" was not found.`);
  }
}

/**
 * Serve acknowledgment form
 */
function serveAckForm(e, baseUrl, user) {
  try {
    const id = (e.parameter.id || "").toString();
    const tpl = HtmlService.createTemplateFromFile('CoachingAckForm');

    tpl.baseUrl = baseUrl;
    tpl.rawToken = e.parameter.token || '';
    tpl.scriptUrl = SCRIPT_URL;
    tpl.currentPage = 'Acknowledge Coaching';
    tpl.recordId = id;

    if (id && typeof getCoachingRecordById === 'function') {
      tpl.record = JSON.stringify(getCoachingRecordById(id)).replace(/<\/script>/g, '<\\/script>');
    } else {
      tpl.record = '{}';
    }

    tpl.user = user;

    return tpl.evaluate()
      .setTitle('Acknowledge Coaching')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('Error serving ack form:', error);
    return createErrorPage('Form Error', error.message);
  }
}

/**
 * Enhanced proxy service that replaces the original serveProxy function
 * This should replace the existing serveProxy function in your Code.gs file
 */
function serveProxy(e) {
  try {
    return serveEnhancedProxy(e);
  } catch (error) {
    console.error('Proxy service error:', error);
    writeError('serveProxy', error);

    // Fallback to basic proxy if enhanced fails
    return serveBasicProxy(e);
  }
}

function serveEnhancedProxy(e) {
  try {
    var target = e.parameter.url || '';
    if (!target) {
      return HtmlService.createHtmlOutput('<h3>Proxy Error</h3><p>Missing url</p>')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    if (!/^https?:\/\//i.test(target)) target = 'https://' + target;

    var resp = UrlFetchApp.fetch(target, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
    });
    var ct = String(resp.getHeaders()['Content-Type'] || '').toLowerCase();
    var body = resp.getContentText();

    if (ct.indexOf('text/html') === -1) {
      // Non-HTML: serve raw
      return ContentService.createTextOutput(body);
    }

    // Strip common frame-busters and CSP meta tags
    body = body
      .replace(/<meta[^>]+http-equiv=["']?content-security-policy["']?[^>]*>/gi, '')
      .replace(/<script[^>]*>[\s\S]*?top\.location(?:.|\s)*?<\/script>/gi, '')
      .replace(/<script[^>]*>[\s\S]*?if\s*\(\s*top\s*[!=]=?\s*self\s*\)[\s\S]*?<\/script>/gi, '');

    // Add <base> to fix relative URLs
    var base = target.replace(/([?#].*)$/, '');
    if (body.indexOf('<base ') === -1) {
      body = body.replace(/<head([^>]*)>/i, function (m, attrs) {
        return '<head' + attrs + '><base href="' + base + '">';
      });
    }

    // Minimal CSS guard
    body = body.replace('</head>', '<style>*,*:before,*:after{box-sizing:border-box}</style></head>');

    return HtmlService.createHtmlOutput(body)
      .setTitle('LuminaHQ Browser')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    writeError && writeError('serveEnhancedProxy', err);
    return HtmlService.createHtmlOutput('<h3>Proxy Error</h3><pre>' + String(err) + '</pre>')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * Fallback basic proxy service (your original implementation with improvements)
 */
function serveBasicProxy(e) {
  try {
    const target = e.parameter.url;
    if (!target) {
      return HtmlService.createHtmlOutput(`
                <!DOCTYPE html>
                <html>
                <head>
                    <title>Proxy Error</title>
                    <style>
                        body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
                        .error { color: #e74c3c; }
                    </style>
                </head>
                <body>
                    <h1 class="error">Proxy Error</h1>
                    <p>Missing URL parameter</p>
                    <button onclick="history.back()">Go Back</button>
                </body>
                </html>
            `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // Normalize URL
    let normalizedUrl = target;
    if (!normalizedUrl.startsWith('http')) {
      normalizedUrl = 'https://' + normalizedUrl;
    }

    const resp = UrlFetchApp.fetch(normalizedUrl, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });

    const content = resp.getContentText();
    const contentType = resp.getHeaders()['Content-Type'] || '';

    if (contentType.toLowerCase().includes('text/html')) {
      // Basic HTML processing to remove frame-busting
      let processedContent = content;

      // Remove common frame-busting scripts
      processedContent = processedContent.replace(
        /<script[^>]*>[\s\S]*?if\s*\(\s*top\s*[!=]=?\s*self\s*\)[\s\S]*?<\/script>/gi,
        ''
      );
      processedContent = processedContent.replace(
        /<script[^>]*>[\s\S]*?top\.location[\s\S]*?<\/script>/gi,
        ''
      );

      // Add basic styles to prevent overlap issues
      const enhancedContent = processedContent.replace(
        '</head>',
        `<style>
                    body { margin-top: 0 !important; }
                    * { box-sizing: border-box; }
                </style></head>`
      );

      return HtmlService.createHtmlOutput(enhancedContent)
        .setTitle('LuminaHQ Browser')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } else {
      // Return non-HTML content as-is
      return ContentService.createTextOutput(content);
    }

  } catch (error) {
    console.error('Basic proxy error:', error);
    return HtmlService.createHtmlOutput(`
            <!DOCTYPE html>
            <html>
            <head>
                <title>Proxy Error</title>
                <style>
                    body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
                    .error { color: #e74c3c; }
                </style>
            </head>
            <body>
                <h1 class="error">Unable to Load Page</h1>
                <p>Error: ${error.message}</p>
                <button onclick="history.back()">Go Back</button>
                <button onclick="location.reload()">Retry</button>
            </body>
            </html>
        `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// INITIALIZATION FUNCTION TO ENSURE BOOKMARK SHEETS EXIST
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Initialize bookmark-related sheets
 * Call this from your main initialization function
 */
function initializeBookmarkSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Create UserBookmarks sheet if it doesn't exist
    let bookmarkSheet = ss.getSheetByName('UserBookmarks');
    if (!bookmarkSheet) {
      bookmarkSheet = ss.insertSheet('UserBookmarks');
      const headers = ['ID', 'UserID', 'Title', 'URL', 'Description', 'Tags', 'Folder', 'Created', 'LastAccessed', 'AccessCount'];
      bookmarkSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      bookmarkSheet.setFrozenRows(1);
      console.log('UserBookmarks sheet created');
    }

    return true;
  } catch (error) {
    console.error('Error initializing bookmark sheets:', error);
    writeError('initializeBookmarkSheets', error);
    return false;
  }
}


/**
 * Get user ID from session token (implement based on your auth system)
 */
function getUserIdFromToken(token) {
  try {
    // This should integrate with your existing authentication system
    if (typeof AuthenticationService !== 'undefined' && AuthenticationService.validateToken) {
      const user = AuthenticationService.validateToken(token);
      return user ? (user.ID || user.id) : null;
    }
    return null;
  } catch (error) {
    console.error('Error getting user ID from token:', error);
    return null;
  }
}

/**
 * Enhanced bookmark retrieval with user filtering
 */
function clientGetUserBookmarksForUser(userId) {
  try {
    if (typeof readSheet === 'function') {
      const allBookmarks = readSheet('UserBookmarks') || [];
      const userBookmarks = allBookmarks.filter(bookmark =>
        String(bookmark.UserID || bookmark.userId) === String(userId)
      );

      return userBookmarks.map(bookmark => ({
        ...bookmark,
        proxyUrl: `${SCRIPT_URL}?page=proxy&url=${encodeURIComponent(bookmark.URL)}`
      }));
    }
    return [];
  } catch (error) {
    console.error('Error getting user bookmarks:', error);
    writeError('clientGetUserBookmarksForUser', error);
    return [];
  }
}

/**
 * Enhanced bookmark addition with user association
 */
function clientAddBookmarkForUser(userId, title, url, description, tags, folder) {
  try {
    const bookmark = {
      ID: Utilities.getUuid(),
      UserID: userId,
      Title: title,
      URL: url,
      Description: description || '',
      Tags: tags || '',
      Folder: folder || 'General',
      Created: new Date(),
      LastAccessed: new Date(),
      AccessCount: 0
    };

    if (typeof writeSheet === 'function') {
      // Append to existing bookmarks
      const existingBookmarks = readSheet('UserBookmarks') || [];
      existingBookmarks.push(bookmark);

      // Write back to sheet
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName('UserBookmarks');
      if (sheet) {
        const headers = Object.keys(bookmark);
        const data = existingBookmarks.map(b => headers.map(h => b[h] || ''));

        sheet.clear();
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        if (data.length > 0) {
          sheet.getRange(2, 1, data.length, headers.length).setValues(data);
        }
      }

      return { success: true, bookmark: bookmark };
    }

    return { success: false, error: 'Bookmark service not available' };
  } catch (error) {
    console.error('Error adding bookmark for user:', error);
    writeError('clientAddBookmarkForUser', error);
    return { success: false, error: error.message };
  }
}

/**
 * Create error page
 */
function createErrorPage(title, message) {
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <title>${title}</title>
      <meta name="viewport" content="width=device-width,initial-scale=1">
      <style>
        body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
        .error { color: #e74c3c; }
        .message { margin: 20px 0; }
        .back-link { color: #3498db; text-decoration: none; }
      </style>
    </head>
    <body>
      <h1 class="error">${title}</h1>
      <p class="message">${message}</p>
      <a href="#" onclick="history.back()" class="back-link">← Go Back</a>
    </body>
    </html>
  `;

  return HtmlService.createHtmlOutput(html)
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ───────────────────────────────────────────────────────────────────────────────
// UTILITY FUNCTIONS - Enhanced with better error handling
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Safe error writing function
 */
function writeError(functionName, error) {
  try {
    const errorMessage = error && error.message ? error.message : error.toString();
    console.error(`[${functionName}] ${errorMessage}`);

    // Try to log to a sheet if possible
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let errorSheet = ss.getSheetByName('ErrorLog');

      if (!errorSheet) {
        errorSheet = ss.insertSheet('ErrorLog');
        errorSheet.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Function', 'Error', 'Stack']]);
      }

      const timestamp = new Date();
      const stack = error && error.stack ? error.stack : 'No stack trace';

      errorSheet.appendRow([timestamp, functionName, errorMessage, stack]);
    } catch (logError) {
      console.error('Failed to log error to sheet:', logError);
    }
  } catch (e) {
    console.error('Error in writeError function:', e);
  }
}

/**
 * Safe alias for writeError
 */
function writeDebug(message) {
  console.log(`[DEBUG] ${message}`);
}

/**
 * Enhanced email confirmation function
 */
function confirmEmail(token) {
  try {
    if (!token) {
      console.warn('confirmEmail called with empty token');
      return false;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(USERS_SHEET);
    if (!sh) {
      console.error('Users sheet not found');
      return false;
    }

    const data = sh.getDataRange().getValues();
    if (data.length < 2) {
      console.warn('No users found in Users sheet');
      return false;
    }

    const headers = data.shift();
    const colTok = headers.indexOf('EmailConfirmation');
    const colConf = headers.indexOf('EmailConfirmed');
    const colUpd = headers.indexOf('UpdatedAt');
    const colEmail = headers.indexOf('Email');

    if (colTok < 0 || colConf < 0) {
      console.error('Required columns not found in Users sheet');
      return false;
    }

    let found = false;
    let userEmail = null;

    data.forEach((row, i) => {
      if (String(row[colTok]) === String(token)) {
        found = true;
        userEmail = row[colEmail];
        const rowNum = i + 2;

        sh.getRange(rowNum, colConf + 1).setValue(true);

        if (colUpd >= 0) {
          sh.getRange(rowNum, colUpd + 1).setValue(new Date());
        }

        console.log(`Email confirmed for user: ${userEmail}`);
      }
    });

    if (typeof invalidateCache === 'function') {
      invalidateCache(USERS_SHEET);
    }

    return found;
  } catch (error) {
    console.error('Error confirming email:', error);
    writeError('confirmEmail', error);
    return false;
  }
}

// ---- global admin helper (the only global admin) -----------------------------
function isSystemAdmin(user) {
  try {
    const u = _normalizeUser(user);
    return !!u.IsAdmin; // driven by Users.IsAdmin or explicit "System Admin" name you already attach in roleNames.
  } catch (e) {
    writeError && writeError('isSystemAdmin', e);
    return false;
  }
}

/**
 * Get roles mapping function
 */
function getRolesMapping() {
  try {
    let roles = [];
    if (typeof getAllRoles === 'function') {
      roles = getAllRoles();
    } else if (typeof readSheet === 'function') {
      roles = readSheet(ROLES_SHEET) || [];
    }

    const mapping = {};
    roles.forEach(role => {
      mapping[role.ID || role.id] = role.Name || role.name;
    });
    return mapping;
  } catch (error) {
    console.error('Error getting roles mapping:', error);
    return {};
  }
}

function getAvailableAgentsForExport(granularity, periodId) {
  try {
    // Use the enhanced authentication to verify user access
    const user = getCurrentUser();
    if (!user || !user.ID) {
      console.warn('getAvailableAgentsForExport: No valid user session');
      return [];
    }

    // Get analytics for the period to find agents with data
    const analytics = getAttendanceAnalyticsByPeriod(granularity, periodId, '');

    if (!analytics || !analytics.userCompliance) {
      console.warn('getAvailableAgentsForExport: No analytics data available');
      return [];
    }

    // Extract unique agent names from user compliance data
    const agents = analytics.userCompliance
      .map(uc => uc.user)
      .filter(user => user && user.trim() !== '')
      .sort();

    // Remove duplicates and return
    return [...new Set(agents)];

  } catch (error) {
    console.error('Error in getAvailableAgentsForExport:', error);
    writeError('getAvailableAgentsForExport', error);
    return [];
  }
}

/**
 * Search function (unchanged but with error handling)
 */
function searchWeb(query, startIndex) {
  try {
    if (!query || typeof query !== 'string') {
      throw new Error('Invalid search query.');
    }
    const CSE_ID = '130aba31c8a2d439c';
    const API_KEY = 'AIzaSyAg-puM5l9iQpjz_NplMJaKbUNRH7ld7sY';
    const baseUrl = 'https://www.googleapis.com/customsearch/v1';
    const params = [
      `key=${API_KEY}`,
      `cx=${CSE_ID}`,
      `q=${encodeURIComponent(query)}`,
      startIndex ? `start=${startIndex}` : ''
    ].filter(Boolean).join('&');
    const url = `${baseUrl}?${params}`;

    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = resp.getResponseCode();
    const text = resp.getContentText();
    if (code !== 200) {
      throw new Error(`Search API error [${code}]: ${text}`);
    }
    return JSON.parse(text);
  } catch (error) {
    console.error('Error in searchWeb:', error);
    writeError('searchWeb', error);
    throw error;
  }
}

function getAllPageKeys() {
  try {
    if (typeof readSheet === 'function') {
      const pages = readSheet(PAGES_SHEET) || [];
      return pages.map(p => p.PageKey || p.key).filter(k => k);
    } else {
      return [];
    }
  } catch (error) {
    console.error('Error getting page keys:', error);
    return [];
  }
}

function validateExportRequest(granularity, periodId, agents, options) {
  const validation = {
    isValid: true,
    errors: []
  };

  // Validate granularity
  const validGranularities = ['Week', 'Month', 'Quarter', 'Year'];
  if (!validGranularities.includes(granularity)) {
    validation.isValid = false;
    validation.errors.push(`Invalid granularity: ${granularity}`);
  }

  // Validate period format
  if (!periodId || periodId.trim() === '') {
    validation.isValid = false;
    validation.errors.push('Period ID is required');
  }

  // Validate agent selection
  if (agents && agents.trim() !== '') {
    const agentList = agents.split(',').map(a => a.trim()).filter(a => a !== '');

    // Check for reasonable limits
    if (agentList.length > 50) {
      validation.isValid = false;
      validation.errors.push('Too many agents selected (maximum 50)');
    }
  }

  // Validate export options
  if (options && typeof options === 'object') {
    const validFormats = ['combined', 'separate', 'summary'];
    if (options.format && !validFormats.includes(options.format)) {
      validation.isValid = false;
      validation.errors.push(`Invalid export format: ${options.format}`);
    }
  }

  return validation;
}

function exportMultiplePeriods(granularity, periods, agents, options) {
  try {
    const exports = [];
    const errors = [];

    periods.forEach(period => {
      try {
        const csv = exportAttendanceCsvEnhanced(granularity, period, agents, options);
        exports.push({
          period: period,
          csv: csv,
          filename: `attendance_${granularity.toLowerCase()}_${period}.csv`
        });
      } catch (error) {
        errors.push(`Failed to export ${period}: ${error.message}`);
      }
    });

    return {
      success: true,
      exports: exports,
      errors: errors
    };

  } catch (error) {
    console.error('Error in exportMultiplePeriods:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Export summary statistics across multiple periods
 */
function exportPeriodComparison(granularity, periods, agents) {
  try {
    const comparisonData = [];

    periods.forEach(period => {
      try {
        const analytics = getAttendanceAnalyticsByPeriod(granularity, period, agents);

        // Extract key metrics for comparison
        const summary = {
          period: period,
          totalProductiveHours: analytics.totalProductiveHours || 0,
          totalNonProductiveHours: analytics.totalNonProductiveHours || 0,
          agentCount: analytics.userCompliance ? analytics.userCompliance.length : 0,
          avgAttendanceRate: 0
        };

        // Calculate average attendance rate
        if (analytics.top5Attendance && analytics.top5Attendance.length > 0) {
          const totalRate = analytics.top5Attendance.reduce((sum, agent) => sum + (agent.percentage || 0), 0);
          summary.avgAttendanceRate = Math.round(totalRate / analytics.top5Attendance.length);
        }

        comparisonData.push(summary);

      } catch (error) {
        console.warn(`Failed to get analytics for period ${period}:`, error);
      }
    });

    // Generate comparison CSV
    if (comparisonData.length === 0) {
      return 'No data available for comparison';
    }

    const headers = ['Period', 'Productive Hours', 'Non-Productive Hours', 'Agent Count', 'Avg Attendance Rate %'];
    const rows = [headers];

    comparisonData.forEach(data => {
      rows.push([
        data.period,
        data.totalProductiveHours.toFixed(2),
        data.totalNonProductiveHours.toFixed(2),
        data.agentCount,
        data.avgAttendanceRate
      ]);
    });

    return rows.map(row => row.map(cell => `"${cell}"`).join(',')).join('\r\n');

  } catch (error) {
    console.error('Error in exportPeriodComparison:', error);
    writeError('exportPeriodComparison', error);
    return 'Error generating comparison export: ' + error.message;
  }
}

/**
 * Client-accessible function for period comparison
 */
function clientExportPeriodComparison(granularity, periods, agents) {
  try {
    return exportPeriodComparison(granularity, periods, agents);
  } catch (error) {
    console.error('Error in clientExportPeriodComparison:', error);
    writeError('clientExportPeriodComparison', error);
    return 'Error generating comparison: ' + error.message;
  }
}

/**
 * Generate quick export templates for common use cases
 */
function getExportTemplates() {
  return [
    {
      name: 'Weekly Summary (All Agents)',
      granularity: 'Week',
      agents: '',
      options: {
        format: 'summary',
        includeBreaks: false,
        includeLunch: false,
        includeWeekends: false
      }
    },
    {
      name: 'Monthly Detailed (All Agents)',
      granularity: 'Month',
      agents: '',
      options: {
        format: 'combined',
        includeBreaks: true,
        includeLunch: true,
        includeWeekends: false
      }
    },
    {
      name: 'Quarterly Overview (All Agents)',
      granularity: 'Quarter',
      agents: '',
      options: {
        format: 'separate',
        includeBreaks: true,
        includeLunch: true,
        includeWeekends: true
      }
    },
    {
      name: 'Weekly Individual Agent Report',
      granularity: 'Week',
      agents: 'SINGLE_AGENT', // Placeholder - will be replaced
      options: {
        format: 'combined',
        includeBreaks: true,
        includeLunch: true,
        includeWeekends: false
      }
    }
  ];
}

/**
 * Client-accessible function to get export templates
 */
function clientGetExportTemplates() {
  try {
    return getExportTemplates();
  } catch (error) {
    console.error('Error getting export templates:', error);
    return [];
  }
}

/**
 * Apply export template with custom parameters
 */
function applyExportTemplate(templateName, customPeriod, customAgents) {
  try {
    const templates = getExportTemplates();
    const template = templates.find(t => t.name === templateName);

    if (!template) {
      throw new Error(`Template "${templateName}" not found`);
    }

    // Apply customizations
    const exportConfig = {
      granularity: template.granularity,
      period: customPeriod || weekStringFromDate(new Date()),
      agents: customAgents || template.agents,
      options: { ...template.options }
    };

    // Replace placeholder for single agent templates
    if (exportConfig.agents === 'SINGLE_AGENT' && customAgents) {
      exportConfig.agents = customAgents;
    }

    return exportConfig;

  } catch (error) {
    console.error('Error applying export template:', error);
    writeError('applyExportTemplate', error);
    return null;
  }
}

/**
 * Generate export with template
 */
function exportWithTemplate(templateName, customPeriod, customAgents) {
  try {
    const config = applyExportTemplate(templateName, customPeriod, customAgents);

    if (!config) {
      throw new Error('Failed to apply export template');
    }

    return exportAttendanceCsvEnhanced(
      config.granularity,
      config.period,
      config.agents,
      config.options
    );

  } catch (error) {
    console.error('Error in exportWithTemplate:', error);
    writeError('exportWithTemplate', error);
    return 'Error generating template export: ' + error.message;
  }
}

/**
 * Week string utility with error handling
 */
function weekStringFromDate(date) {
  try {
    if (!date || !(date instanceof Date)) {
      date = new Date();
    }

    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    const weekNum = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    return `${d.getUTCFullYear()}-W${weekNum.toString().padStart(2, '0')}`;
  } catch (error) {
    console.error('Error generating week string:', error);
    return new Date().getFullYear() + '-W01';
  }
}

/** Handy debug hook: dump decision to console. */
function _debugAccess(label, result, user, pageKey, campaignId) {
  try {
    if (!ACCESS_DEBUG) return;
    console.log('[ACCESS] ' + label + ' → allow=' + result.allow + ' reason=' + result.reason +
      ' page=' + pageKey + ' campaign=' + (campaignId || '-') +
      ' user=' + (user && (user.Email || user.UserName || user.ID)));
    (result.trace || []).forEach((t, i) => console.log('   [' + i + '] ' + t));
  } catch (_) {
  }
}

function isUserAdmin(user) {
  // Keep simple and consistent: our global admin is the only real admin
  return isSystemAdmin(user);
}

function getUserList() {
  try {
    // limit to assigned scope by default
    return clientGetAssignedAgentNames('');
  } catch (error) {
    console.error('Error getting user list:', error);
    return [];
  }
}

function getAttendanceUserList() {
  try {
    return clientGetAssignedAgentNames('');
  } catch (error) {
    console.error('Error getting attendance user list:', error);
    return [];
  }
}

function getAllCampaigns() {
  try {
    if (typeof clientGetAllCampaigns === 'function') {
      return clientGetAllCampaigns();
    } else if (typeof readSheet === 'function') {
      return readSheet(CAMPAIGNS_SHEET) || [];
    } else {
      return [];
    }
  } catch (error) {
    console.error('Error getting all campaigns:', error);
    return [];
  }
}

function getAllRoles() {
  try {
    if (typeof readSheet === 'function') {
      return readSheet(ROLES_SHEET) || [];
    } else {
      return [];
    }
  } catch (error) {
    console.error('Error getting all roles:', error);
    return [];
  }
}