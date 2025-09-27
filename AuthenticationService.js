/**
 * AuthenticationService.gs - Simplified Token-Based Authentication Service
 * Updated to use standard token-based session management
 * 
 * Features:
 * - Token-based session management
 * - Enhanced security with session tokens
 * - Email integration for password resets and confirmations
 * - Simple URL structure
 * - Improved error handling and logging
 */

// ───────────────────────────────────────────────────────────────────────────────
// AUTHENTICATION CONFIGURATION
// ───────────────────────────────────────────────────────────────────────────────

// Session TTL: 1 hour for regular sessions, 24 hours for remember me
const SESSION_TTL_MS = 60 * 60 * 1000; // 1 hour
const REMEMBER_ME_TTL_MS = 24 * 60 * 60 * 1000; // 24 hours

// ───────────────────────────────────────────────────────────────────────────────
// GUARDED GLOBAL DEFAULTS (supports standalone deployments)
// ───────────────────────────────────────────────────────────────────────────────
if (typeof USERS_SHEET === 'undefined') var USERS_SHEET = 'Users';
if (typeof ROLES_SHEET === 'undefined') var ROLES_SHEET = 'Roles';
if (typeof USER_ROLES_SHEET === 'undefined') var USER_ROLES_SHEET = 'UserRoles';
if (typeof USER_CLAIMS_SHEET === 'undefined') var USER_CLAIMS_SHEET = 'UserClaims';
if (typeof SESSIONS_SHEET === 'undefined') var SESSIONS_SHEET = 'Sessions';

if (typeof USERS_HEADERS === 'undefined') var USERS_HEADERS = [
  'ID', 'UserName', 'FullName', 'Email', 'CampaignID', 'PasswordHash', 'ResetRequired',
  'EmailConfirmation', 'EmailConfirmed', 'PhoneNumber', 'EmploymentStatus', 'HireDate', 'Country',
  'LockoutEnd', 'TwoFactorEnabled', 'CanLogin', 'Roles', 'Pages', 'CreatedAt', 'UpdatedAt', 'IsAdmin'
];
if (typeof ROLES_HEADER === 'undefined') var ROLES_HEADER = ['ID', 'Name', 'NormalizedName', 'CreatedAt', 'UpdatedAt'];
if (typeof USER_ROLES_HEADER === 'undefined') var USER_ROLES_HEADER = ['UserId', 'RoleId', 'CreatedAt', 'UpdatedAt'];
if (typeof CLAIMS_HEADERS === 'undefined') var CLAIMS_HEADERS = ['ID', 'UserId', 'ClaimType', 'CreatedAt', 'UpdatedAt'];
if (typeof SESSIONS_HEADERS === 'undefined') var SESSIONS_HEADERS = [
  'Token',
  'UserId',
  'CreatedAt',
  'ExpiresAt',
  'RememberMe',
  'CampaignScope',
  'UserAgent',
  'IpAddress'
];

function resolveScriptUrl() {
  if (typeof SCRIPT_URL !== 'undefined' && SCRIPT_URL) {
    return SCRIPT_URL;
  }

  try {
    return ScriptApp.getService().getUrl();
  } catch (error) {
    console.warn('resolveScriptUrl: unable to determine script URL', error);
    return '';
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// AUTHENTICATION SERVICE IMPLEMENTATION
// ───────────────────────────────────────────────────────────────────────────────

var AuthenticationService = (function () {

  const passwordUtils = (function resolvePasswordUtilities() {
    if (typeof ensurePasswordUtilities === 'function') {
      return ensurePasswordUtilities();
    }
    
    if (typeof PasswordUtilities !== 'undefined' && PasswordUtilities) {
      return PasswordUtilities;
    }

    if (typeof __createPasswordUtilitiesModule === 'function') {
      const utils = __createPasswordUtilitiesModule();

      try {
        if (typeof PasswordUtilities === 'undefined' || !PasswordUtilities) {
          PasswordUtilities = utils;
        }
      } catch (assignErr) {
        // Ignore assignment issues (e.g., strict mode) and just return the instance.
      }

      if (typeof ensurePasswordUtilities !== 'function') {
        try {
          ensurePasswordUtilities = function ensurePasswordUtilities() { return utils; };
        } catch (ensureErr) {
          // Ignore if the global cannot be reassigned.
        }
      }

      return utils;
    }

    throw new Error('PasswordUtilities module is not available.');
    
  })();

  function normalizeHashValue(hash) {
    if (passwordUtils && typeof passwordUtils.normalizeHash === 'function') {
      return passwordUtils.normalizeHash(hash);
    }
    if (passwordUtils && typeof passwordUtils.decodePasswordHash === 'function') {
      return passwordUtils.decodePasswordHash(hash);
    }
    return toStr(hash).toLowerCase();
  }

  // ─── Internal helpers ─────────────────────────────────────────────────────

  function getSS() {
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  function getOrCreateSheet(name, headers) {
    let sh = getSS().getSheetByName(name);
    if (!sh) {
      sh = getSS().insertSheet(name);
      sh.clear();
      if (headers && headers.length > 0) {
        sh.appendRow(headers);
      }
    }
    return sh;
  }

  function hasDatabaseManager() {
    return typeof DatabaseManager !== 'undefined'
      && DatabaseManager
      && typeof DatabaseManager.table === 'function';
  }

  function getDbTable(sheetName) {
    if (!hasDatabaseManager()) return null;
    try {
      return DatabaseManager.table(sheetName);
    } catch (err) {
      console.warn('DatabaseManager.table failed for ' + sheetName + ':', err);
      return null;
    }
  }

  function toStr(value) {
    if (value === null || typeof value === 'undefined') return '';
    if (value instanceof Date) return value.toISOString();
    return String(value).trim();
  }

  function toBool(value) {
    if (value === true) return true;
    if (value === false) return false;
    const str = toStr(value).toLowerCase();
    return str === 'true' || str === '1' || str === 'yes' || str === 'y';
  }

  function dedupeStrings(values) {
    const set = {};
    const out = [];
    (values || []).forEach(val => {
      const key = toStr(val);
      if (key && !set[key]) {
        set[key] = true;
        out.push(key);
      }
    });
    return out;
  }

  function normalizePrincipal(value) {
    const raw = toStr(value);
    return {
      raw,
      lower: raw.toLowerCase()
    };
  }

  function cloneScope(scope) {
    if (!scope) {
      return {
        defaultCampaignId: '',
        allowedCampaignIds: [],
        managedCampaignIds: [],
        adminCampaignIds: [],
        isGlobalAdmin: false
      };
    }

    return {
      defaultCampaignId: toStr(scope.defaultCampaignId || scope.DefaultCampaignId),
      allowedCampaignIds: dedupeStrings(scope.allowedCampaignIds || scope.AllowedCampaignIds),
      managedCampaignIds: dedupeStrings(scope.managedCampaignIds || scope.ManagedCampaignIds),
      adminCampaignIds: dedupeStrings(scope.adminCampaignIds || scope.AdminCampaignIds),
      isGlobalAdmin: !!scope.isGlobalAdmin
    };
  }

  function sanitizeRole(role) {
    if (!role) return null;
    const id = toStr(role.ID || role.Id || role.id);
    if (!id) return null;

    return {
      ID: id,
      Name: toStr(role.Name || role.name),
      NormalizedName: toStr(
        role.NormalizedName
        || role.normalizedName
        || (role.Name || role.name || '').toUpperCase()
      ),
      CreatedAt: role.CreatedAt || role.createdAt || null,
      UpdatedAt: role.UpdatedAt || role.updatedAt || null
    };
  }

  function sanitizeClaim(claim) {
    if (!claim) return null;
    const id = toStr(claim.ID || claim.Id || claim.id);
    const userId = toStr(claim.UserId || claim.UserID || claim.userId);
    if (!userId) return null;

    return {
      ID: id,
      UserId: userId,
      ClaimType: toStr(claim.ClaimType || claim.claimType),
      CreatedAt: claim.CreatedAt || claim.createdAt || null,
      UpdatedAt: claim.UpdatedAt || claim.updatedAt || null
    };
  }

  function sanitizeUserForTransport(user) {
    if (!user || typeof user !== 'object') return null;

    const scope = cloneScope(user.TenantScope || user.tenantScope);

    return {
      ID: toStr(user.ID || user.Id || user.id),
      UserName: toStr(user.UserName || user.username),
      FullName: toStr(user.FullName || user.fullName || user.UserName || user.username),
      Email: toStr(user.Email || user.email).toLowerCase(),
      IsAdmin: toBool(user.IsAdmin),
      IsGlobalAdmin: !!user.IsGlobalAdmin,
      CampaignID: toStr(user.CampaignID || user.CampaignId || user.campaignId),
      DefaultCampaignId: toStr(user.DefaultCampaignId || user.defaultCampaignId),
      AllowedCampaignIds: scope.allowedCampaignIds.slice(),
      ManagedCampaignIds: scope.managedCampaignIds.slice(),
      AdminCampaignIds: scope.adminCampaignIds.slice(),
      TenantScope: scope,
      roles: (Array.isArray(user.roles) ? user.roles : []).map(sanitizeRole).filter(Boolean),
      claims: (Array.isArray(user.claims) ? user.claims : []).map(sanitizeClaim).filter(Boolean),
      pages: (Array.isArray(user.pages) ? user.pages : []).map(toStr).filter(Boolean)
    };
  }

  function normalizePrincipal(value) {
    const raw = toStr(value);
    return {
      raw,
      lower: raw.toLowerCase()
    };
  }

  function cloneScope(scope) {
    if (!scope) {
      return {
        defaultCampaignId: '',
        allowedCampaignIds: [],
        managedCampaignIds: [],
        adminCampaignIds: [],
        isGlobalAdmin: false
      };
    }

    return {
      defaultCampaignId: toStr(scope.defaultCampaignId || scope.DefaultCampaignId),
      allowedCampaignIds: dedupeStrings(scope.allowedCampaignIds || scope.AllowedCampaignIds),
      managedCampaignIds: dedupeStrings(scope.managedCampaignIds || scope.ManagedCampaignIds),
      adminCampaignIds: dedupeStrings(scope.adminCampaignIds || scope.AdminCampaignIds),
      isGlobalAdmin: !!scope.isGlobalAdmin
    };
  }

  function sanitizeRole(role) {
    if (!role) return null;
    const id = toStr(role.ID || role.Id || role.id);
    if (!id) return null;

    return {
      ID: id,
      Name: toStr(role.Name || role.name),
      NormalizedName: toStr(
        role.NormalizedName
        || role.normalizedName
        || (role.Name || role.name || '').toUpperCase()
      ),
      CreatedAt: role.CreatedAt || role.createdAt || null,
      UpdatedAt: role.UpdatedAt || role.updatedAt || null
    };
  }

  function sanitizeClaim(claim) {
    if (!claim) return null;
    const id = toStr(claim.ID || claim.Id || claim.id);
    const userId = toStr(claim.UserId || claim.UserID || claim.userId);
    if (!userId) return null;

    return {
      ID: id,
      UserId: userId,
      ClaimType: toStr(claim.ClaimType || claim.claimType),
      CreatedAt: claim.CreatedAt || claim.createdAt || null,
      UpdatedAt: claim.UpdatedAt || claim.updatedAt || null
    };
  }

  function sanitizeUserForTransport(user) {
    if (!user || typeof user !== 'object') return null;

    const scope = cloneScope(user.TenantScope || user.tenantScope);

    return {
      ID: toStr(user.ID || user.Id || user.id),
      UserName: toStr(user.UserName || user.username),
      FullName: toStr(user.FullName || user.fullName || user.UserName || user.username),
      Email: toStr(user.Email || user.email).toLowerCase(),
      IsAdmin: toBool(user.IsAdmin),
      IsGlobalAdmin: !!user.IsGlobalAdmin,
      CampaignID: toStr(user.CampaignID || user.CampaignId || user.campaignId),
      DefaultCampaignId: toStr(user.DefaultCampaignId || user.defaultCampaignId),
      AllowedCampaignIds: scope.allowedCampaignIds.slice(),
      ManagedCampaignIds: scope.managedCampaignIds.slice(),
      AdminCampaignIds: scope.adminCampaignIds.slice(),
      TenantScope: scope,
      roles: (Array.isArray(user.roles) ? user.roles : []).map(sanitizeRole).filter(Boolean),
      claims: (Array.isArray(user.claims) ? user.claims : []).map(sanitizeClaim).filter(Boolean),
      pages: (Array.isArray(user.pages) ? user.pages : []).map(toStr).filter(Boolean)
    };
  }

  function readTable(sheetName, options = {}) {
    const table = getDbTable(sheetName);
    if (table) {
      try {
        const results = table.find(options) || [];
        if (Array.isArray(results)) {
          return results;
        }
      } catch (err) {
        console.warn('DatabaseManager read failed for ' + sheetName + ':', err);
      }
    }

    const sh = getOrCreateSheet(sheetName, options.headers || []);
    const vals = sh.getDataRange().getValues();
    if (vals.length < 2) return [];
    const hdrs = vals.shift();
    let rows = vals.map(row => {
      const obj = {};
      hdrs.forEach((h, i) => {
        if (h) obj[h] = row[i];
      });
      return obj;
    });

    if (options.where && typeof options.where === 'object') {
      rows = rows.filter(r => Object.keys(options.where).every(key => toStr(r[key]) === toStr(options.where[key])));
    }

    if (typeof options.limit === 'number') {
      rows = rows.slice(0, Math.max(0, options.limit));
    }

    return rows;
  }

  function getUserById(userId) {
    if (!userId && userId !== 0) return null;
    try {
      const table = getDbTable(USERS_SHEET);
      if (table && typeof table.findById === 'function') {
        const found = table.findById(userId);
        if (found) return found;
      }
    } catch (err) {
      console.warn('getUserById DatabaseManager lookup failed:', err);
    }

    return readTable(USERS_SHEET)
      .find(u => String(u.ID) === String(userId)) || null;
  }

  function buildSessionScope(userId, userRecord) {
    let profile = null;
    try {
      if (typeof TenantSecurity !== 'undefined'
        && TenantSecurity
        && typeof TenantSecurity.getAccessProfile === 'function') {
        profile = TenantSecurity.getAccessProfile(userId);
      }
    } catch (err) {
      console.warn('buildSessionScope: tenant profile lookup failed', err);
    }

    const user = userRecord || getUserById(userId) || {};
    const defaultCampaignId = profile && profile.defaultCampaignId
      ? toStr(profile.defaultCampaignId)
      : toStr(user.CampaignID || user.CampaignId);

    const allowedCampaignIds = dedupeStrings(profile ? profile.allowedCampaignIds : [defaultCampaignId]);
    if (!allowedCampaignIds.length && defaultCampaignId) {
      allowedCampaignIds.push(defaultCampaignId);
    }

    const scope = {
      defaultCampaignId,
      allowedCampaignIds,
      managedCampaignIds: dedupeStrings(profile ? profile.managedCampaignIds : []),
      adminCampaignIds: dedupeStrings(profile ? profile.adminCampaignIds : []),
      isGlobalAdmin: profile ? !!profile.isGlobalAdmin : toBool(user.IsAdmin)
    };

    return cloneScope(scope);
  }

  function updateSessionRecord(sessionToken, updates) {
    if (!sessionToken || !updates || typeof updates !== 'object') return false;
    const table = getDbTable(SESSIONS_SHEET);
    if (table) {
      try {
        table.update(sessionToken, updates);
        return true;
      } catch (err) {
        console.warn('updateSessionRecord: DatabaseManager update failed', err);
      }
    }

    const sh = getOrCreateSheet(SESSIONS_SHEET, SESSIONS_HEADERS);
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return false;
    let headers = data[0].map(h => toStr(h));
    let mutated = false;

    Object.keys(updates).forEach(key => {
      if (headers.indexOf(key) === -1) {
        headers.push(key);
        mutated = true;
      }
    });

    if (mutated) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    }

    const tokenIdx = headers.indexOf('Token') !== -1 ? headers.indexOf('Token') : headers.indexOf('SessionToken');
    if (tokenIdx === -1) return false;

    for (let i = 1; i < data.length; i++) {
      if (toStr(data[i][tokenIdx]) === toStr(sessionToken)) {
        Object.keys(updates).forEach(key => {
          const colIndex = headers.indexOf(key);
          if (colIndex !== -1) {
            sh.getRange(i + 1, colIndex + 1).setValue(updates[key]);
          }
        });
        return true;
      }
    }

    return false;
  }

  function removeSessionByToken(sessionToken) {
    if (!sessionToken) return false;
    const table = getDbTable(SESSIONS_SHEET);
    if (table) {
      try {
        return !!table.delete(sessionToken);
      } catch (err) {
        console.warn('removeSessionByToken: DatabaseManager delete failed', err);
      }
    }

    const sh = getOrCreateSheet(SESSIONS_SHEET, SESSIONS_HEADERS);
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return false;
    const headers = data[0].map(h => toStr(h));
    const tokenIdx = headers.indexOf('Token') !== -1 ? headers.indexOf('Token') : headers.indexOf('SessionToken');
    if (tokenIdx === -1) return false;

    for (let i = 1; i < data.length; i++) {
      if (toStr(data[i][tokenIdx]) === toStr(sessionToken)) {
        sh.deleteRow(i + 1);
        return true;
      }
    }
    return false;
  }

  function cleanExpiredSessions() {
    const now = Date.now();
    const sessions = readTable(SESSIONS_SHEET);
    sessions.forEach((s) => {
      try {
        const expiry = new Date(s.ExpiresAt).getTime();
        if (!isNaN(expiry) && expiry < now) {
          removeSessionByToken(s.Token || s.SessionToken);
        }
      } catch (e) {
        console.warn('Error cleaning expired session:', e);
        removeSessionByToken(s.Token || s.SessionToken);
      }
    });
  }

  function hashPwd(raw) {
    return passwordUtils.hashPassword(raw);
  }

  function generateSecureToken() {
    return Utilities.getUuid() + '_' + Date.now();
  }

  // ─── Public API ────────────────────────────────────────────────────────────

  /** Ensure all identity & session sheets exist with proper headers */
  function ensureSheets() {
    try {
      getOrCreateSheet(USERS_SHEET, USERS_HEADERS);
      getOrCreateSheet(ROLES_SHEET, ROLES_HEADER);
      getOrCreateSheet(USER_ROLES_SHEET, USER_ROLES_HEADER);
      getOrCreateSheet(USER_CLAIMS_SHEET, CLAIMS_HEADERS);
      const sessionsSheet = getOrCreateSheet(SESSIONS_SHEET, SESSIONS_HEADERS);
      if (sessionsSheet) {
        const lastColumn = Math.max(sessionsSheet.getLastColumn(), SESSIONS_HEADERS.length, 1);
        const headerRange = sessionsSheet.getRange(1, 1, 1, lastColumn);
        let existing = headerRange.getValues()[0].map(h => toStr(h));
        let mutated = false;

        if (existing.indexOf('Token') === -1) {
          const legacyIdx = existing.indexOf('SessionToken');
          if (legacyIdx !== -1) {
            existing[legacyIdx] = 'Token';
            mutated = true;
          }
        }

        SESSIONS_HEADERS.forEach(header => {
          if (existing.indexOf(header) === -1) {
            existing.push(header);
            mutated = true;
          }
        });

        const beforeFilterLength = existing.length;
        existing = existing.filter(Boolean);
        if (existing.length !== beforeFilterLength) {
          mutated = true;
        }
        if (!existing.length) {
          existing = SESSIONS_HEADERS.slice();
          mutated = true;
        }

        if (mutated) {
          sessionsSheet.getRange(1, 1, 1, existing.length).setValues([existing]);
          sessionsSheet.setFrozenRows(1);
        }
      }
      console.log('Authentication sheets initialized successfully');
    } catch (error) {
      console.error('Error ensuring authentication sheets:', error);
      throw error;
    }
  }

  function findUserByPrincipal(principal) {
    const norm = normalizePrincipal(principal);
    if (!norm.lower) return null;

    try {
      const match = readTable(USERS_SHEET)
        .find(u => {
          const email = toStr(u.Email).toLowerCase();
          const username = toStr(u.UserName).toLowerCase();
          return email === norm.lower || username === norm.lower;
        });
      if (match) return match;
    } catch (error) {
      console.warn('Primary user lookup failed:', error);
    }

    try {
      if (typeof readSheet === 'function') {
        const rows = readSheet(USERS_SHEET) || [];
        return rows.find(u => {
          const email = toStr(u.Email).toLowerCase();
          const username = toStr(u.UserName).toLowerCase();
          return email === norm.lower || username === norm.lower;
        }) || null;
      }
    } catch (fallbackError) {
      console.warn('Fallback user lookup failed:', fallbackError);
    }

    return null;
  }

  /** Find a user row by email (case-insensitive) */
  function getUserByEmail(email) {
    try {
      return findUserByPrincipal(email);
    } catch (error) {
      console.error('Error getting user by email:', error);
      return null;
    }
  }

  /** Create a new session for a user ID, returning the session token */
  function createSessionFor(userId, existingToken = null, rememberMe = false, options) {
    try {
      const token = existingToken || generateSecureToken();
      const now = new Date();
      const ttl = rememberMe ? REMEMBER_ME_TTL_MS : SESSION_TTL_MS;
      const expiresAt = new Date(now.getTime() + ttl);

      const scope = cloneScope((options && options.scope) || buildSessionScope(userId, options && options.user));

      const sessionRecord = {
        Token: token,
        UserId: userId,
        CreatedAt: now.toISOString(),
        ExpiresAt: expiresAt.toISOString(),
        RememberMe: rememberMe ? 'TRUE' : 'FALSE',
        CampaignScope: scope ? JSON.stringify(scope) : '',
        UserAgent: (options && options.userAgent) || 'Google Apps Script',
        IpAddress: (options && options.ipAddress) || 'N/A'
      };

      const table = getDbTable(SESSIONS_SHEET);
      if (table) {
        if (existingToken && typeof table.findById === 'function' && table.findById(token)) {
          table.update(token, sessionRecord);
        } else {
          table.insert(sessionRecord);
        }
      } else {
        if (existingToken) {
          updateSessionRecord(token, sessionRecord);
        } else {
          const headers = (typeof SESSIONS_HEADERS !== 'undefined' && Array.isArray(SESSIONS_HEADERS))
            ? SESSIONS_HEADERS
            : ['Token', 'UserId', 'CreatedAt', 'ExpiresAt', 'RememberMe', 'CampaignScope', 'UserAgent', 'IpAddress'];
          ensureSheets();
          const sheet = getOrCreateSheet(SESSIONS_SHEET, headers);
          const lastColumn = Math.max(sheet.getLastColumn(), headers.length, 1);
          let headerRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(h => toStr(h));
          if (headerRow.indexOf('Token') === -1) {
            const legacyIdx = headerRow.indexOf('SessionToken');
            if (legacyIdx !== -1) headerRow[legacyIdx] = 'Token';
          }
          headers.forEach(h => {
            if (headerRow.indexOf(h) === -1) headerRow.push(h);
          });
          headerRow = headerRow.filter(Boolean);
          sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
          sheet.setFrozenRows(1);
          const rowValues = headerRow.map(h => typeof sessionRecord[h] !== 'undefined' ? sessionRecord[h] : '');
          sheet.appendRow(rowValues);
        }
      }

      let sessionMeta = null;
      if (options && typeof options === 'object') {
        if (options.sessionMeta && typeof options.sessionMeta === 'object') {
          sessionMeta = options.sessionMeta;
        } else {
          sessionMeta = {};
          options.sessionMeta = sessionMeta;
        }
      }

      if (sessionMeta) {
        sessionMeta.token = token;
        sessionMeta.issuedAt = now.toISOString();
        sessionMeta.expiresAt = expiresAt.toISOString();
        sessionMeta.ttlMs = ttl;
        sessionMeta.rememberMe = !!rememberMe;
      }

      console.log(`Session created for user ${userId}, expires: ${expiresAt.toISOString()}`);
      return token;
    } catch (error) {
      console.error('Error creating session:', error);
      return null;
    }
  }

  /**
   * Simplified login function 
   * @param {string} email - User email
   * @param {string} rawPwd - Plain text password
   * @param {boolean} rememberMe - Whether to create long-term session
   * @return {Object} Login result with detailed status information
   */
  function login(email, rawPwd, rememberMe = false) {
    try {
      ensureSheets();

      const normalizedEmail = toStr(email).toLowerCase();
      const passwordInput = rawPwd == null ? '' : String(rawPwd);

      if (!normalizedEmail || !passwordInput.trim()) {
        return {
          success: false,
          error: 'Email and password are required',
          errorCode: 'MISSING_CREDENTIALS'
        };
      }

      cleanExpiredSessions();
      const user = getUserByEmail(normalizedEmail);

      if (!user) {
        return {
          success: false,
          error: 'Invalid email or password',
          errorCode: 'INVALID_CREDENTIALS'
        };
      }

      // Check if user can login
      const canLogin = toBool(user.CanLogin);
      if (!canLogin) {
        return {
          success: false,
          error: 'Your account has been disabled. Please contact support.',
          errorCode: 'ACCOUNT_DISABLED'
        };
      }

      // Check email confirmation
      const emailConfirmed = toBool(user.EmailConfirmed);
      if (!emailConfirmed) {
        return {
          success: false,
          error: 'Please confirm your email address before logging in.',
          errorCode: 'EMAIL_NOT_CONFIRMED',
          needsEmailConfirmation: true
        };
      }

      // Check if password has been set
      const storedHash = normalizeHashValue(user.PasswordHash);
      const hasPassword = storedHash.length > 0;
      if (!hasPassword) {
        return {
          success: false,
          error: 'Please set up your password using the link from your welcome email.',
          errorCode: 'PASSWORD_NOT_SET',
          needsPasswordSetup: true
        };
      }

      // Verify password
      if (!passwordUtils.verifyPassword(passwordInput, user.PasswordHash)) {
        return {
          success: false,
          error: 'Invalid email or password',
          errorCode: 'INVALID_CREDENTIALS'
        };
      }

      const tenantScope = buildSessionScope(user.ID, user);
      if (!tenantScope.isGlobalAdmin && (!tenantScope.allowedCampaignIds || tenantScope.allowedCampaignIds.length === 0)) {
        return {
          success: false,
          error: 'No campaign assignments were found for your account. Please contact your administrator.',
          errorCode: 'NO_CAMPAIGN_ACCESS'
        };
      }

      const sessionScope = cloneScope(tenantScope);
      const sessionMeta = {};
      const sessionOptions = { scope: sessionScope, user, sessionMeta };

      // Check if password reset is required
      const resetRequired = toBool(user.ResetRequired);
      if (resetRequired) {
        // Create a temporary session for password reset
        const resetToken = createSessionFor(user.ID, null, false, sessionOptions);
        return {
          success: false,
          error: 'You must change your password before continuing.',
          errorCode: 'PASSWORD_RESET_REQUIRED',
          resetToken: resetToken,
          needsPasswordReset: true
        };
      }

      // Create session
      const token = createSessionFor(user.ID, null, rememberMe, sessionOptions);
      if (!token) {
        return {
          success: false,
          error: 'Failed to create session. Please try again.',
          errorCode: 'SESSION_CREATION_FAILED'
        };
      }

      const sessionExpiresAt = sessionMeta.expiresAt || new Date(Date.now() + (rememberMe ? REMEMBER_ME_TTL_MS : SESSION_TTL_MS)).toISOString();
      const sessionIssuedAt = sessionMeta.issuedAt || new Date().toISOString();
      const sessionTtlMs = sessionMeta.ttlMs || (rememberMe ? REMEMBER_ME_TTL_MS : SESSION_TTL_MS);
      const sessionRememberMe = sessionMeta.rememberMe !== undefined ? !!sessionMeta.rememberMe : !!rememberMe;

      // Update last login
      updateLastLogin(user.ID);

      const userPayload = sanitizeUserForTransport(
        buildClientUserPayload(user, sessionScope)
      );

      return {
        success: true,
        sessionToken: token,
        user: userPayload,
        message: 'Login successful',
        sessionExpiresAt,
        sessionIssuedAt,
        sessionTtlMs,
        rememberMe: sessionRememberMe
      };

    } catch (error) {
      console.error('Login error:', error);
      writeError('AuthenticationService.login', error);
      return {
        success: false,
        error: 'An error occurred during login. Please try again.',
        errorCode: 'SYSTEM_ERROR'
      };
    }
  }

  function getUserCampaignPages(userId, campaignId) {
    try {
      if (!campaignId) return [];

      // Get campaign pages from campaign configuration
      const campaignPages = getCampaignPages(campaignId);
      return campaignPages
        .filter(cp => cp && cp.IsActive !== false)
        .map(cp => toStr(cp.PageKey))
        .filter(Boolean);
    } catch (e) {
      console.warn('Error getting user campaign pages:', e);
      return [];
    }
  }

  function getCampaignPages(campaignId) {
    try {
      // This would typically come from a CAMPAIGN_PAGES sheet
      // For now, return default pages
      return [
        { PageKey: 'dashboard', IsActive: true },
        { PageKey: 'reports', IsActive: true },
        { PageKey: 'qa', IsActive: true }
      ];
    } catch (error) {
      console.warn('Error getting campaign pages:', error);
      return [];
    }
  }

  function buildClientUserPayload(userRow, scope) {
    if (!userRow) return null;

    const safeScope = cloneScope(scope || buildSessionScope(userRow.ID, userRow));
    const allowedCampaignIds = safeScope.allowedCampaignIds.slice();
    const effectiveCampaignId = safeScope.defaultCampaignId
      || (allowedCampaignIds.length ? allowedCampaignIds[0] : toStr(userRow.CampaignID));

    return {
      ID: toStr(userRow.ID),
      UserName: toStr(userRow.UserName),
      FullName: toStr(userRow.FullName) || toStr(userRow.UserName),
      Email: toStr(userRow.Email).toLowerCase(),
      IsAdmin: toBool(userRow.IsAdmin),
      IsGlobalAdmin: !!safeScope.isGlobalAdmin,
      CampaignID: effectiveCampaignId,
      DefaultCampaignId: safeScope.defaultCampaignId,
      AllowedCampaignIds: allowedCampaignIds,
      ManagedCampaignIds: safeScope.managedCampaignIds.slice(),
      AdminCampaignIds: safeScope.adminCampaignIds.slice(),
      TenantScope: safeScope,
      roles: getUserRoles(userRow.ID),
      claims: getUserClaims(userRow.ID),
      pages: getUserCampaignPages(userRow.ID, effectiveCampaignId)
    };
  }

  function buildClientUserPayload(userRow, scope) {
    if (!userRow) return null;

    const safeScope = cloneScope(scope || buildSessionScope(userRow.ID, userRow));
    const allowedCampaignIds = safeScope.allowedCampaignIds.slice();
    const effectiveCampaignId = safeScope.defaultCampaignId
      || (allowedCampaignIds.length ? allowedCampaignIds[0] : toStr(userRow.CampaignID));

    return {
      ID: toStr(userRow.ID),
      UserName: toStr(userRow.UserName),
      FullName: toStr(userRow.FullName) || toStr(userRow.UserName),
      Email: toStr(userRow.Email).toLowerCase(),
      IsAdmin: toBool(userRow.IsAdmin),
      IsGlobalAdmin: !!safeScope.isGlobalAdmin,
      CampaignID: effectiveCampaignId,
      DefaultCampaignId: safeScope.defaultCampaignId,
      AllowedCampaignIds: allowedCampaignIds,
      ManagedCampaignIds: safeScope.managedCampaignIds.slice(),
      AdminCampaignIds: safeScope.adminCampaignIds.slice(),
      TenantScope: safeScope,
      roles: getUserRoles(userRow.ID),
      claims: getUserClaims(userRow.ID),
      pages: getUserCampaignPages(userRow.ID, effectiveCampaignId)
    };
  }

  /**
   * Simplified logout function that handles session cleanup
   * @param {string} sessionToken - Session token to invalidate
   * @return {Object} Logout result
   */
  function logout(sessionToken) {
    try {
      console.log('Logout initiated for token:', sessionToken ? sessionToken.substring(0, 8) + '...' : 'no token');
      
      let sessionRemoved = false;
      
      if (sessionToken) {
        sessionRemoved = _invalidateSessionByToken(sessionToken);
        console.log('Session invalidation result:', sessionRemoved);
      }

      return {
        success: true,
        sessionRemoved: sessionRemoved,
        message: 'Logged out successfully'
      };
    } catch (error) {
      console.error('Error during logout:', error);
      writeError('logout', error);
      return {
        success: false,
        error: error.message,
        message: 'Logout completed with errors'
      };
    }
  }

  /**
   * Validate & extend a session token (sliding expiration).
   * Returns the full user object (with roles/pages arrays) or null.
   */
  function getSessionUser(sessionToken) {
    try {
      if (!sessionToken) {
        return null;
      }

      cleanExpiredSessions();

      let sessionRec = null;
      const sessionTable = getDbTable(SESSIONS_SHEET);
      if (sessionTable && typeof sessionTable.findById === 'function') {
        try {
          sessionRec = sessionTable.findById(sessionToken);
        } catch (err) {
          console.warn('Session lookup via DatabaseManager failed:', err);
        }
      }

      if (!sessionRec) {
        sessionRec = readTable(SESSIONS_SHEET)
          .find(s => toStr(s.Token || s.SessionToken) === toStr(sessionToken));
      }

      if (!sessionRec) {
        console.log('Session not found for token:', sessionToken.substring(0, 8) + '...');
        return null;
      }

      // Check if session is expired
      const expiryTime = new Date(sessionRec.ExpiresAt).getTime();
      const now = Date.now();

      if (!expiryTime || isNaN(expiryTime) || expiryTime < now) {
        console.log('Session expired for token:', sessionToken.substring(0, 8) + '...');
        _invalidateSessionByToken(sessionToken);
        return null;
      }

      // Extend session expiry (sliding expiration)
      const isRememberMe = toBool(sessionRec.RememberMe);
      const ttl = isRememberMe ? REMEMBER_ME_TTL_MS : SESSION_TTL_MS;
      const newExpiry = new Date(now + ttl);

      const sessionUserId = sessionRec.UserId || sessionRec.UserID || sessionRec.userid || sessionRec.userId;
      const userRow = getUserById(sessionUserId);
      if (!userRow) {
        console.warn('User not found for session:', sessionRec.UserId);
        _invalidateSessionByToken(sessionToken);
        return null;
      }

      const tenantScope = buildSessionScope(userRow.ID, userRow);
      const sessionScope = cloneScope(tenantScope);

      try {
        updateSessionRecord(sessionToken, {
          ExpiresAt: newExpiry.toISOString(),
          CampaignScope: JSON.stringify(sessionScope)
        });
      } catch (updateErr) {
        console.warn('Failed to update session metadata:', updateErr);
      }

      const userPayload = sanitizeUserForTransport(
        buildClientUserPayload(userRow, sessionScope)
      );

      if (!userPayload) {
        console.warn('Unable to build user payload for session user:', userRow && userRow.ID);
        return null;
      }

      const sessionUser = Object.assign({}, userPayload, {
        sessionToken,
        sessionExpiry: newExpiry.toISOString(),
        rememberMe: isRememberMe,
        needsReset: toBool(userRow.ResetRequired)
      });

      console.log('Session validated and extended for user:', userPayload.Email);
      return sessionUser;

    } catch (error) {
      console.error('Error validating session:', error);
      writeError('getSessionUser', error);
      return null;
    }
  }

  function updateLastLogin(userId) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(USERS_SHEET);
      const data = sh.getDataRange().getValues();
      const headers = data[0];

      // Add LastLogin column if it doesn't exist
      let lastLoginIndex = headers.indexOf('LastLogin');
      if (lastLoginIndex === -1) {
        // Add the column
        sh.getRange(1, headers.length + 1).setValue('LastLogin');
        lastLoginIndex = headers.length;
      }

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === userId) {
          const row = i + 1;
          const col = lastLoginIndex + 1;

          const now = new Date();
          sh.getRange(row, col).setValue(now);
          sh.getRange(row, col).setNumberFormat('yyyy-mm-dd hh:mm:ss');
          break;
        }
      }
    } catch (error) {
      console.warn('Failed to update last login:', error);
    }
  }

  /**
   * Simplified keepAlive function 
   * @param {string} sessionToken - Session token to validate
   */
  function keepAlive(sessionToken) {
    try {
      if (!sessionToken) {
        return {
          success: false,
          expired: true,
          message: 'No session token provided'
        };
      }

      const user = getSessionUser(sessionToken);
      if (!user) {
        return {
          success: false,
          expired: true,
          message: 'Session expired or invalid'
        };
      }

      const expiresAt = user.sessionExpiry || null;
      let sessionTtlSeconds = null;
      if (expiresAt) {
        const expiryTime = Date.parse(expiresAt);
        if (!isNaN(expiryTime)) {
          sessionTtlSeconds = Math.max(0, Math.floor((expiryTime - Date.now()) / 1000));
        }
      }

      return {
        success: true,
        message: 'Session active',
        user: {
          ID: user.ID,
          FullName: user.FullName,
          Email: user.Email,
          CampaignID: user.CampaignID,
          IsAdmin: user.IsAdmin,
          rememberMe: !!user.rememberMe
        },
        sessionToken: user.sessionToken,
        sessionExpiresAt: expiresAt,
        sessionTtlSeconds,
        rememberMe: !!user.rememberMe
      };
    } catch (error) {
      console.error('Error in keepAlive:', error);
      writeError('keepAlive', error);
      return {
        success: false,
        error: error.message,
        message: 'Session check failed'
      };
    }
  }

  /**
   * Get user claims by user ID
   */
  function getUserClaims(userId) {
    try {
      const target = toStr(userId);
      if (!target) return [];

      const claims = readTable(USER_CLAIMS_SHEET);
      return claims
        .filter(claim => toStr(claim.UserId || claim.UserID) === target)
        .map(sanitizeClaim)
        .filter(Boolean);
    } catch (error) {
      console.error('Error getting user claims:', error);
      return [];
    }
  }

  /**
   * Get user roles by user ID
   */
  function getUserRoles(userId) {
    try {
      const target = toStr(userId);
      if (!target) return [];

      const userRoles = readTable(USER_ROLES_SHEET);
      const allRoles = readTable(ROLES_SHEET);

      const userRoleIds = userRoles
        .filter(ur => toStr(ur.UserId || ur.UserID) === target)
        .map(ur => toStr(ur.RoleId || ur.RoleID));

      return allRoles
        .filter(role => userRoleIds.indexOf(toStr(role.ID || role.Id)) !== -1)
        .map(sanitizeRole)
        .filter(Boolean);
    } catch (error) {
      console.error('Error getting user roles:', error);
      return [];
    }
  }

  /**
   * Simplified authentication requirement function
   */
  function requireAuth(e) {
    try {
      // Check for token parameter
      const token = e.parameter.token;
      if (token) {
        const user = getSessionUser(token);
        if (user) {
          return user;
        }
      }

      // Fall back to current user via Google session
      const user = getCurrentUserProfile_();
      if (!user || !user.ID) {
        return HtmlService
          .createTemplateFromFile('Login')
          .evaluate()
          .setTitle('Please Log In')
          .addMetaTag('viewport', 'width=device-width,initial-scale=1')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }

      const pageParam = (e.parameter.page || 'dashboard').toLowerCase();
      const allowed = (user.pages || []).map(p => p.toLowerCase());

      if (e.parameter.page
        && allowed.length > 0
        && allowed.indexOf(pageParam) < 0) {
        const tpl = HtmlService.createTemplateFromFile('AccessDenied');
        tpl.baseUrl = ScriptApp.getService().getUrl();
        return tpl
          .evaluate()
          .setTitle('Access Denied')
          .addMetaTag('viewport', 'width=device-width,initial-scale=1')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }

      return user;
    } catch (error) {
      console.error('Error in requireAuth:', error);
      writeError('requireAuth', error);
      throw error;
    }
  }

  /**
   * Enforce both auth & campaign match.
   * Returns user or throws.
   */
  function requireCampaignAuth(e) {
    const maybe = requireAuth(e);
    if (maybe.getContent) return maybe;
    const user = maybe;
    const cid = toStr(e.parameter.campaignId || e.parameter.campaignID || e.parameter.campaign || '');
    if (!cid) {
      return user;
    }

    try {
      if (typeof TenantSecurity !== 'undefined'
        && TenantSecurity
        && typeof TenantSecurity.assertCampaignAccess === 'function') {
        TenantSecurity.assertCampaignAccess(user.ID, cid);
      } else {
        const allowed = Array.isArray(user.AllowedCampaignIds)
          ? user.AllowedCampaignIds.map(toStr)
          : [];
        if (!user.IsGlobalAdmin && allowed.indexOf(cid) === -1 && toStr(user.CampaignID) !== cid) {
          throw new Error('Not authorized for campaign: ' + cid);
        }
      }
      return user;
    } catch (err) {
      console.warn('Campaign authorization failed:', err);
      const tpl = HtmlService.createTemplateFromFile('AccessDenied');
      tpl.baseUrl = ScriptApp.getService().getUrl();
      return tpl
        .evaluate()
        .setTitle('Access Denied')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }

  /**
   * Change password function
   * @param {string} sessionToken - Session token
   * @param {string} oldPassword - Current password
   * @param {string} newPassword - New password
   * @return {Object} Change result
   */
  function changePassword(sessionToken, oldPassword, newPassword) {
    try {
      const user = getSessionUser(sessionToken);
      if (!user) {
        return { success: false, message: 'Not authenticated.' };
      }

      // Validate new password strength
      if (!newPassword || newPassword.length < 8) {
        return { success: false, message: 'Password must be at least 8 characters long.' };
      }

      // Check password complexity
      const hasUpper = /[A-Z]/.test(newPassword);
      const hasLower = /[a-z]/.test(newPassword);
      const hasNumber = /[0-9]/.test(newPassword);
      const hasSpecial = /[^A-Za-z0-9]/.test(newPassword);

      if (!hasUpper || !hasLower || !hasNumber || !hasSpecial) {
        return {
          success: false,
          message: 'Password must contain at least one uppercase letter, one lowercase letter, one number, and one special character.'
        };
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(USERS_SHEET);
      const data = sh.getDataRange().getValues();
      const headers = data[0];

      const now = new Date();

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === user.ID) {
          const pwdColIndex = headers.indexOf('PasswordHash');
          const resetColIndex = headers.indexOf('ResetRequired');
          const updatedAtIndex = headers.indexOf('UpdatedAt');

          if (!passwordUtils.verifyPassword(oldPassword, data[i][pwdColIndex])) {
            return { success: false, message: 'Current password is incorrect.' };
          }

          const rowNum = i + 1;
          const newHash = passwordUtils.hashPassword(newPassword);
          sh.getRange(rowNum, pwdColIndex + 1).setValue(newHash);
          sh.getRange(rowNum, resetColIndex + 1).setValue(false);
          if (updatedAtIndex >= 0) {
            sh.getRange(rowNum, updatedAtIndex + 1).setValue(now);
          }

          // Send password change confirmation email
          try {
            if (typeof sendPasswordChangeConfirmation === 'function') {
              sendPasswordChangeConfirmation(user.Email, { timestamp: now });
            }
          } catch (emailError) {
            console.warn('Failed to send password change email:', emailError);
          }

          return { success: true, message: 'Password changed successfully.' };
        }
      }

      return { success: false, message: 'User record not found.' };
    } catch (error) {
      console.error('Error changing password:', error);
      writeError('changePassword', error);
      return { success: false, message: 'An error occurred while changing password.' };
    }
  }

  /**
   * Set password with token function (for new users and password resets)
   * @param {string} token - Email confirmation or reset token
   * @param {string} newPassword - New password
   * @return {Object} Set password result
   */
  function setPasswordWithToken(token, newPassword) {
    try {
      // Validate password strength
      if (!newPassword || newPassword.length < 8) {
        return { success: false, message: 'Password must be at least 8 characters long.' };
      }

      const hasUpper = /[A-Z]/.test(newPassword);
      const hasLower = /[a-z]/.test(newPassword);
      const hasNumber = /[0-9]/.test(newPassword);
      const hasSpecial = /[^A-Za-z0-9]/.test(newPassword);

      if (!hasUpper || !hasLower || !hasNumber || !hasSpecial) {
        return {
          success: false,
          message: 'Password must contain uppercase, lowercase, number, and special character.'
        };
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(USERS_SHEET);
      const data = sh.getDataRange().getValues();
      const headers = data[0];

      const now = new Date();

      for (let i = 1; i < data.length; i++) {
        const tokenColIndex = headers.indexOf('EmailConfirmation'); // Password setup token

        if (String(data[i][tokenColIndex]) === token) {
          const rowNum = i + 1;
          const email = data[i][headers.indexOf('Email')];

          // Set password and update status
          const newHash = passwordUtils.hashPassword(newPassword);
          sh.getRange(rowNum, headers.indexOf('PasswordHash') + 1).setValue(newHash);
          sh.getRange(rowNum, headers.indexOf('ResetRequired') + 1).setValue('FALSE');
          sh.getRange(rowNum, headers.indexOf('EmailConfirmation') + 1).setValue(''); // Clear token
          sh.getRange(rowNum, headers.indexOf('UpdatedAt') + 1).setValue(now);

          // Send confirmation email
          try {
            if (typeof sendPasswordChangeConfirmation === 'function') {
              sendPasswordChangeConfirmation(email, { timestamp: now });
            }
          } catch (emailError) {
            console.warn('Failed to send password confirmation email:', emailError);
          }

          return {
            success: true,
            message: 'Password set successfully. You can now log in.'
          };
        }
      }

      return { success: false, message: 'Invalid or expired setup link.' };
    } catch (error) {
      writeError('setPasswordWithToken', error);
      return { success: false, message: 'An error occurred while setting password.' };
    }
  }

  /**
   * Generate and send password reset token
   * @param {string} email - User email
   * @return {Object} Reset result
   */
  function requestPasswordReset(email) {
    try {
      const user = getUserByEmail(email);
      if (!user) {
        return {
          success: true,
          message: 'If an account with this email exists, a password reset link has been sent.'
        };
      }

      // Check if user has a password set
      if (!normalizeHashValue(user.PasswordHash)) {
        return {
          success: false,
          error: 'This account needs initial password setup. Check your welcome email.'
        };
      }

      // Generate reset token and send email
      const resetToken = Utilities.getUuid();

      // Update user record
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(USERS_SHEET);
      const data = sh.getDataRange().getValues();
      const headers = data[0];

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][headers.indexOf('Email')]).toLowerCase() === email.toLowerCase()) {
          const rowNum = i + 1;
          sh.getRange(rowNum, headers.indexOf('EmailConfirmation') + 1).setValue(resetToken);
          sh.getRange(rowNum, headers.indexOf('UpdatedAt') + 1).setValue(new Date());
          break;
        }
      }

      // Send reset email
      let emailSent = false;
      try {
        if (typeof sendPasswordResetEmail === 'function') {
          emailSent = sendPasswordResetEmail(email, resetToken);
        }
      } catch (emailError) {
        console.error('Error sending password reset email:', emailError);
      }

      return emailSent ?
        { success: true, message: 'Password reset email sent.' } :
        { success: false, error: 'Failed to send email. Please try again.' };

    } catch (error) {
      writeError('requestPasswordReset', error);
      return { success: false, error: 'An error occurred. Please try again later.' };
    }
  }

  /**
   * Resend password setup email for new users
   */
  function resendPasswordSetupEmail(email) {
    try {
      const user = getUserByEmail(email);
      if (!user) {
        return {
          success: true,
          message: 'If an account with this email exists, a password setup link has been sent.'
        };
      }

      // Check if user needs password setup
      const needsSetup = !normalizeHashValue(user.PasswordHash)
        || String(user.ResetRequired).toUpperCase() === 'TRUE';
      if (!needsSetup) {
        return {
          success: false,
          error: 'This account already has a password set. Use "Forgot Password" instead.'
        };
      }

      // Generate new setup token
      const setupToken = Utilities.getUuid();

      // Update user record with new token
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(USERS_SHEET);
      const data = sh.getDataRange().getValues();
      const headers = data[0];

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][headers.indexOf('Email')]).toLowerCase() === email.toLowerCase()) {
          const rowNum = i + 1;
          sh.getRange(rowNum, headers.indexOf('EmailConfirmation') + 1).setValue(setupToken);
          sh.getRange(rowNum, headers.indexOf('UpdatedAt') + 1).setValue(new Date());
          break;
        }
      }

      // Send email
      let emailSent = false;
      try {
        if (typeof sendPasswordSetupEmail === 'function') {
          emailSent = sendPasswordSetupEmail(email, {
            userName: user.UserName,
            fullName: user.FullName || user.UserName,
            passwordSetupToken: setupToken
          });
        }
      } catch (emailError) {
        console.error('Error sending password setup email:', emailError);
      }

      return emailSent ?
        { success: true, message: 'Password setup email sent successfully.' } :
        { success: false, error: 'Failed to send email. Please try again.' };

    } catch (error) {
      writeError('resendPasswordSetupEmail', error);
      return { success: false, error: 'An error occurred. Please try again later.' };
    }
  }

  /**
   * Log user activity for audit purposes
   * @param {string} sessionToken - Session token
   * @param {Object} activity - Activity data to log
   */
  function logUserActivity(sessionToken, activity) {
    try {
      const user = getSessionUser(sessionToken);
      if (!user) {
        throw new Error('User not authenticated');
      }

      // Ensure "UserActivityLog" sheet exists
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sh = ss.getSheetByName('UserActivityLog');
      if (!sh) {
        sh = ss.insertSheet('UserActivityLog');
        sh.appendRow(['Timestamp', 'UserEmail', 'ActivityPayload']);
      }

      // Append the audit entry
      sh.appendRow([
        new Date(),
        user.Email,
        JSON.stringify(activity, null, 2)
      ]);

      return { success: true };
    } catch (error) {
      console.error('Error logging user activity:', error);
      writeError('logUserActivity', error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Private helper to invalidate a session by token
   */
  function _invalidateSessionByToken(sessionToken) {
    try {
      if (!sessionToken) return false;
      const removed = removeSessionByToken(sessionToken);
      if (!removed) {
        console.log('Session not found for invalidation:', sessionToken.substring(0, 8) + '...');
        return false;
      }

      console.log('Session invalidated:', sessionToken.substring(0, 8) + '...');
      return true;
    } catch (error) {
      console.error('Error invalidating session:', error);
      return false;
    }
  }

  // Return public API
  return {
    ensureSheets,
    login,
    logout,
    getSessionUser,
    keepAlive,
    requireAuth,
    requireCampaignAuth,
    getUserByEmail,
    getUserClaims,
    getUserRoles,
    changePassword,
    setPasswordWithToken,
    requestPasswordReset,
    resendPasswordSetupEmail,
    logUserActivity,
    createSessionFor,
    _invalidateSessionByToken
  };
})();

// ───────────────────────────────────────────────────────────────────────────────
// CLIENT-ACCESSIBLE FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Simplified login function for standard authentication
 * Called by the client-side code via google.script.run
 */
function loginUser(email, password, rememberMe = false) {
  try {
    const normalizedEmail = typeof email === 'string' ? email.trim() : '';
    const rememberFlag = rememberMe === true
      || rememberMe === 'true'
      || rememberMe === 1
      || rememberMe === '1';
    console.log('Server-side login for:', normalizedEmail || '[empty]');

    const result = AuthenticationService.login(normalizedEmail, password, rememberFlag);

    if (!result || typeof result !== 'object') {
      return {
        success: false,
        error: 'The authentication service did not return a response. Please try again.',
        errorCode: 'NO_RESPONSE'
      };
    }

    if (!result.success) {
      return {
        success: false,
        error: result.error || 'Login failed. Please try again.',
        errorCode: result.errorCode || 'AUTHENTICATION_FAILED',
        needsEmailConfirmation: !!result.needsEmailConfirmation,
        needsPasswordReset: !!result.needsPasswordReset,
        needsPasswordSetup: !!result.needsPasswordSetup,
        resetToken: result.resetToken || null
      };
    }

    const sessionToken = toStr(result.sessionToken);
    if (!sessionToken) {
      return {
        success: false,
        error: 'Login succeeded but no session token was generated. Please try again.',
        errorCode: 'SESSION_TOKEN_MISSING'
      };
    }

    const sanitizedUser = sanitizeUserForTransport(result.user);

    const issuedAt = result.sessionIssuedAt ? new Date(result.sessionIssuedAt) : new Date();
    const ttlMs = (typeof result.sessionTtlMs === 'number' && !isNaN(result.sessionTtlMs))
      ? result.sessionTtlMs
      : (rememberFlag ? REMEMBER_ME_TTL_MS : SESSION_TTL_MS);
    let expiresAtIso = result.sessionExpiresAt || null;
    if (!expiresAtIso) {
      const computedExpiry = new Date(issuedAt.getTime() + ttlMs);
      expiresAtIso = computedExpiry.toISOString();
    }

    const ttlSeconds = Math.max(0, Math.floor(ttlMs / 1000));
    const rememberResponse = result.rememberMe !== undefined ? !!result.rememberMe : rememberFlag;

    const baseScriptUrl = resolveScriptUrl();
    const redirectUrl = baseScriptUrl
      ? (baseScriptUrl + '?page=dashboard&token=' + encodeURIComponent(sessionToken))
      : ('?page=dashboard&token=' + encodeURIComponent(sessionToken));

    return {
      success: true,
      message: result.message || 'Login successful',
      redirectUrl,
      sessionToken,
      user: sanitizedUser,
      rememberMe: rememberResponse,
      sessionExpiresAt: expiresAtIso,
      sessionIssuedAt: issuedAt.toISOString(),
      sessionTtlSeconds: ttlSeconds
    };
  } catch (error) {
    console.error('Server login error:', error);
    writeError('loginUser', error);
    return {
      success: false,

      error: 'Login failed. Please try again.',
      errorCode: 'SYSTEM_ERROR'
    };
  }
}

/**
 * Simplified logout function
 */
function logoutUser(sessionToken) {
  try {
    console.log('Server-side logout');
    
    const result = AuthenticationService.logout(sessionToken);
    
    return {
      success: true,
      message: 'Logout successful',
      redirectUrl: resolveScriptUrl()
    };
  } catch (error) {
    console.error('Server logout error:', error);
    writeError('logoutUser', error);
    return {
      success: false,
      error: error.message,
      redirectUrl: resolveScriptUrl()
    };
  }
}

/**
 * Keep alive function for sessions
 */
function keepAliveSession(sessionToken) {
  try {
    return AuthenticationService.keepAlive(sessionToken);
  } catch (error) {
    console.error('Keep alive error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

// Expose individual functions for google.script.run
function getUserClaims(userId) {
  return AuthenticationService.getUserClaims(userId);
}

function getUserRoles(userId) {
  return AuthenticationService.getUserRoles(userId);
}

function setPasswordWithToken(token, newPassword) {
  return AuthenticationService.setPasswordWithToken(token, newPassword);
}

function changePassword(sessionToken, oldPassword, newPassword) {
  return AuthenticationService.changePassword(sessionToken, oldPassword, newPassword);
}

function requestPasswordReset(email) {
  return AuthenticationService.requestPasswordReset(email);
}

function resendPasswordSetupEmail(email) {
  return AuthenticationService.resendPasswordSetupEmail(email);
}

function login(email, password) {
  return AuthenticationService.login(email, password);
}

function logout(sessionToken) {
  return AuthenticationService.logout(sessionToken);
}

function keepAlive(sessionToken) {
  return AuthenticationService.keepAlive(sessionToken);
}

function clientLogUserActivity(sessionToken, activity) {
  return AuthenticationService.logUserActivity(sessionToken, activity);
}

function ensureAuthenticationSheets() {
  return AuthenticationService.ensureSheets();
}

// ───────────────────────────────────────────────────────────────────────────────
// UTILITY FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

function writeError(functionName, error) {
  try {
    const errorMessage = error && error.message ? error.message : error.toString();
    console.error(`[${functionName}] ${errorMessage}`);

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
 * Get current user without requiring token
 * Uses Google Apps Script's built-in session management
 */
function getCurrentUserProfile_() {
  try {
    // Use Google's built-in user session
    const email = String(
      (Session.getActiveUser() && Session.getActiveUser().getEmail()) ||
      (Session.getEffectiveUser() && Session.getEffectiveUser().getEmail()) ||
      ''
    ).trim().toLowerCase();

    if (!email) {
      return null;
    }

    // Get user from your existing user management system
    const user = AuthenticationService.getUserByEmail(email);
    if (!user) {
      return null;
    }

    var tenantProfile = null;
    try {
      if (typeof TenantSecurity !== 'undefined' && TenantSecurity && typeof TenantSecurity.getAccessProfile === 'function') {
        tenantProfile = TenantSecurity.getAccessProfile(user.ID);
      }
    } catch (ctxErr) {
      console.warn('getCurrentUserProfile_: tenant profile load failed', ctxErr);
    }

    var allowedCampaigns = tenantProfile ? tenantProfile.allowedCampaignIds.slice() : [];
    if (!allowedCampaigns.length && user.CampaignID) {
      allowedCampaigns.push(String(user.CampaignID));
    }

    // Return user in expected format
    return {
      ID: user.ID,
      Email: user.Email,
      FullName: user.FullName,
      UserName: user.UserName,
      CampaignID: user.CampaignID,
      IsAdmin: String(user.IsAdmin).toUpperCase() === 'TRUE',
      IsGlobalAdmin: tenantProfile ? !!tenantProfile.isGlobalAdmin : String(user.IsAdmin).toUpperCase() === 'TRUE',

      CanLogin: String(user.CanLogin).toUpperCase() === 'TRUE',
      EmailConfirmed: String(user.EmailConfirmed).toUpperCase() === 'TRUE',
      ResetRequired: String(user.ResetRequired).toUpperCase() === 'TRUE',
      AllowedCampaignIds: allowedCampaigns,
      ManagedCampaignIds: tenantProfile ? tenantProfile.managedCampaignIds.slice() : [],
      AdminCampaignIds: tenantProfile ? tenantProfile.adminCampaignIds.slice() : [],
      DefaultCampaignId: tenantProfile ? tenantProfile.defaultCampaignId : toStr(user.CampaignID),

      TenantAccess: tenantProfile
    };
  } catch (error) {
    console.error('Error getting current user:', error);
    return null;
  }
}

// Provide a fallback global getCurrentUser implementation when another module
// (for example Code.js) has not already registered one. This ensures
// authentication-dependent modules can resolve the active user consistently.
(function ensureGlobalGetCurrentUser() {
  const root = (typeof globalThis !== 'undefined')
    ? globalThis
    : (typeof self !== 'undefined')
      ? self
      : this;

  if (root && typeof root.getCurrentUser !== 'function') {
    root.getCurrentUser = function () {
      return getCurrentUserProfile_();
    };
  }
}).call(this);

// ───────────────────────────────────────────────────────────────────────────────
// INITIALIZATION LOG
// ───────────────────────────────────────────────────────────────────────────────

console.log('Simplified AuthenticationService.gs loaded successfully');
console.log('Features: Token-based sessions, Enhanced security, Simple URLs, Comprehensive error handling');
