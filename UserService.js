/******************************************************************************* 
 * UserService.gs — Complete Campaign-aware User Management + HR/Benefits
 * - FIXED globals: safe constant pattern (no block-scoped const in if)
 * - Adds HR/Benefits columns & logic:
 *     TerminationDate, ProbationMonths, ProbationEndDate,
 *     InsuranceQualifiedDate, InsuranceEligible, InsuranceSignedUp,
 *     InsuranceCardReceivedDate
 *   Insurance eligibility = 3 months AFTER probation ends (configurable)
 * - Safe fallbacks for readSheet/ensureSheetWithHeaders/writeError/etc.
 * - User CRUD, Page assignment, Roles, Campaign permissions, Manager mapping
 *******************************************************************************/

// ───────────────────────────────────────────────────────────────────────────────
// Global guard (Apps Script V8-safe)
// ───────────────────────────────────────────────────────────────────────────────
const G = (typeof globalThis !== 'undefined') ? globalThis : (typeof this !== 'undefined' ? this : {});

// Sheets
if (typeof G.USERS_SHEET === 'undefined') G.USERS_SHEET = 'Users';
if (typeof G.ROLES_SHEET === 'undefined') G.ROLES_SHEET = 'Roles';
if (typeof G.PAGES_SHEET === 'undefined') G.PAGES_SHEET = 'Pages';
if (typeof G.CAMPAIGNS_SHEET === 'undefined') G.CAMPAIGNS_SHEET = 'Campaigns';
if (typeof G.USER_ROLES_SHEET === 'undefined') G.USER_ROLES_SHEET = 'UserRoles';
if (typeof G.CAMPAIGN_USER_PERMISSIONS_SHEET === 'undefined') G.CAMPAIGN_USER_PERMISSIONS_SHEET = 'CampaignUserPermissions';
if (typeof G.MANAGER_USERS_SHEET === 'undefined') G.MANAGER_USERS_SHEET = 'MANAGER_USERS';
if (typeof G.MANAGER_USERS_HEADER === 'undefined') G.MANAGER_USERS_HEADER = ['ID', 'ManagerUserID', 'UserID', 'CreatedAt', 'UpdatedAt'];

// Campaign user permissions headers
if (typeof G.CAMPAIGN_USER_PERMISSIONS_HEADERS === 'undefined') {
  G.CAMPAIGN_USER_PERMISSIONS_HEADERS = [
    'ID', 'CampaignID', 'UserID', 'PermissionLevel', 'CanManageUsers', 'CanManagePages', 'CreatedAt', 'UpdatedAt'
  ];
}

// HR/Benefits config
if (typeof G.INSURANCE_MONTHS_AFTER_PROBATION === 'undefined') G.INSURANCE_MONTHS_AFTER_PROBATION = 3;

// Optional extra columns we’ll ensure on Users sheet if missing
const OPTIONAL_USER_COLUMNS = [
  'TerminationDate',
  'ProbationMonths',
  'ProbationEndDate',
  'InsuranceQualifiedDate',
  'InsuranceEligible',
  'InsuranceSignedUp',
  'InsuranceCardReceivedDate'
];

// ───────────────────────────────────────────────────────────────────────────────
// Safe fallbacks for common helpers (no-ops if you already defined them)
// ───────────────────────────────────────────────────────────────────────────────
if (typeof writeError !== 'function') {
  function writeError(where, err) { try { console.error('[ERROR]', where, err && (err.stack || err)); } catch (_) { Logger.log(where + ': ' + err); } }
}
if (typeof invalidateCache !== 'function') {
  function invalidateCache() { /* no-op fallback */ }
}
if (typeof ensureSheetWithHeaders !== 'function') {
  function ensureSheetWithHeaders(name, headers) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    } else {
      const lastCol = sh.getLastColumn();
      const row = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
      const hdrs = row.map(String);
      // append any missing headers
      let changed = false;
      headers.forEach(h => {
        if (hdrs.indexOf(h) === -1) { hdrs.push(h); changed = true; }
      });
      if (changed) sh.getRange(1, 1, 1, hdrs.length).setValues([hdrs]);
    }
    return sh;
  }
}
if (typeof readSheet !== 'function') {
  function readSheet(name) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(name);
    if (!sh) return [];
    const vals = sh.getDataRange().getValues();
    if (vals.length < 2) return [];
    const headers = vals[0].map(String);
    return vals.slice(1).map(r => {
      const o = {};
      headers.forEach((h, i) => o[h] = (typeof r[i] !== 'undefined') ? r[i] : '');
      return o;
    });
  }
}

// Validators (only if you don’t already have them)
if (typeof getValidEmploymentStatuses !== 'function') {
  function getValidEmploymentStatuses() {
    return ['Active', 'Leave', 'Terminated', 'Pending', 'Probation', 'Contractor', 'Intern'];
  }
}
if (typeof validateEmploymentStatus !== 'function') {
  function validateEmploymentStatus(s) { return !s || getValidEmploymentStatuses().indexOf(String(s).trim()) !== -1; }
}
if (typeof validateHireDate !== 'function') {
  function validateHireDate(d) {
    if (!d) return true;
    const dt = new Date(d);
    if (isNaN(dt)) return false;
    const now = new Date();
    return dt <= now;
  }
}
if (typeof validateCountry !== 'function') {
  function validateCountry(c) {
    const s = String(c || '').trim();
    return s.length === 0 || (s.length >= 2 && s.length <= 100);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// Date helpers + Benefits calculators
// ───────────────────────────────────────────────────────────────────────────────
function _toIsoDateOnly_(d) {
  if (!d) return '';
  const dt = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dt)) return '';
  return new Date(Date.UTC(dt.getFullYear(), dt.getMonth(), dt.getDate())).toISOString().slice(0, 10);
}
function _addMonths_(dateStr, months) {
  if (!dateStr) return '';
  const dt = new Date(dateStr);
  if (isNaN(dt)) return '';
  const d = new Date(dt);
  const targetMonth = d.getMonth() + Number(months || 0);
  d.setMonth(targetMonth);
  // normalize end-of-month issues
  if (d.getDate() !== dt.getDate()) d.setDate(0);
  return _toIsoDateOnly_(d);
}
function calcProbationEndDate_(hireDateStr, probationMonths) {
  if (!hireDateStr || !probationMonths) return '';
  return _addMonths_(hireDateStr, Number(probationMonths || 0));
}
function calcInsuranceQualifiedDate_(probationEndStr, monthsAfter) {
  if (!probationEndStr) return '';
  const m = (typeof monthsAfter === 'number') ? monthsAfter : G.INSURANCE_MONTHS_AFTER_PROBATION;
  return _addMonths_(probationEndStr, m);
}
function isInsuranceEligibleNow_(qualifiedDateStr, terminationDateStr) {
  if (!qualifiedDateStr) return false;
  const q = new Date(qualifiedDateStr);
  if (isNaN(q)) return false;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  if (q > today) return false;
  if (terminationDateStr) {
    const t = new Date(terminationDateStr);
    if (!isNaN(t) && t <= today) return false;
  }
  return true;
}
function _boolToStr_(v) { return (v === true || String(v).trim().toUpperCase() === 'TRUE' || String(v).trim().toUpperCase() === 'YES' || String(v).trim() === '1') ? 'TRUE' : 'FALSE'; }
function _strToBool_(v) { return (v === true || String(v).trim().toUpperCase() === 'TRUE'); }

// ───────────────────────────────────────────────────────────────────────────────
// Sheet utils (local)
// ───────────────────────────────────────────────────────────────────────────────
function _getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet not found: ' + name);
  return sh;
}
function _scanSheet_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Sheet has no header row: ' + sheet.getName());
  const headers = values[0].map(String);
  const idx = {}; headers.forEach((h, i) => idx[h] = i);
  return { headers, values, idx };
}
function ensureOptionalUserColumns_(sh, headers, idx) {
  let changed = false;
  const hdrs = headers.slice();
  OPTIONAL_USER_COLUMNS.forEach(h => {
    if (hdrs.indexOf(h) === -1) { hdrs.push(h); changed = true; }
  });
  if (changed) {
    sh.getRange(1, 1, 1, hdrs.length).setValues([hdrs]);
    const lastCol = sh.getLastColumn();
    const newHeaders = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    const newIdx = {}; newHeaders.forEach((h, i) => newIdx[h] = i);
    return { headers: newHeaders, idx: newIdx, changed: true };
  }
  return { headers, idx, changed: false };
}

// Base required columns we always expect on Users sheet
const REQUIRED_USER_COLUMNS = [
  'ID', 'UserName', 'FullName', 'Email', 'CampaignID', 'PasswordHash', 'ResetRequired',
  'EmailConfirmation', 'EmailConfirmed', 'PhoneNumber', 'EmploymentStatus', 'HireDate', 'Country',
  'LockoutEnd', 'TwoFactorEnabled', 'CanLogin', 'Roles', 'Pages', 'CreatedAt', 'UpdatedAt', 'IsAdmin'
];

function _ensureUserHeaders_(idx) {
  // Throw only for the base minimal set; optional benefits columns are handled separately
  for (let i = 0; i < REQUIRED_USER_COLUMNS.length; i++) {
    const h = REQUIRED_USER_COLUMNS[i];
    if (typeof idx[h] !== 'number' || idx[h] < 0) {
      throw new Error('USERS sheet missing column: ' + h);
    }
  }
}

function checkEmailExists(email) {
  try {
    if (!email) return { exists: false };

    const users = readSheet('Users') || [];
    const normalizedEmail = String(email).trim().toLowerCase();

    const existingUser = users.find(user =>
      user.Email && String(user.Email).trim().toLowerCase() === normalizedEmail
    );

    if (existingUser) {
      return {
        exists: true,
        user: {
          ID: existingUser.ID,
          FullName: existingUser.FullName || existingUser.UserName,
          UserName: existingUser.UserName,
          Email: existingUser.Email,
          CampaignID: existingUser.CampaignID,
          campaignName: getCampaignNameSafe(existingUser.CampaignID),
          canLoginBool: (existingUser.CanLogin === true || String(existingUser.CanLogin).toUpperCase() === 'TRUE')
        }
      };
    }

    return { exists: false };
  } catch (error) {
    writeError('checkEmailExists', error);
    return { exists: false, error: error.message };
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// Admin detection (global + role + campaign permissions)
// ───────────────────────────────────────────────────────────────────────────────
function isUserAdmin(uOrId) {
  if (!uOrId) return false;
  const user = (typeof uOrId === 'string')
    ? (readSheet(G.USERS_SHEET) || []).find(x => String(x.ID) === String(uOrId))
    : uOrId;

  if (!user) return false;
  const flag = (user.IsAdmin === true || String(user.IsAdmin).toUpperCase() === 'TRUE');
  if (flag) return true;

  try {
    const roleNames = getUserRolesSafe(user.ID).map(r => String(r.name || '').toLowerCase());
    if (roleNames.some(n => /\b(system\s*admin|super\s*admin|administrator|admin)\b/.test(n))) return true;
  } catch (_) { }

  try {
    const perms = readCampaignPermsSafely_();
    if (perms.some(p => String(p.UserID) === String(user.ID) &&
      String(p.PermissionLevel || '').toUpperCase() === 'ADMIN')) return true;
  } catch (_) { }

  return false;
}

// ───────────────────────────────────────────────────────────────────────────────
// Campaign permissions CRUD
// ───────────────────────────────────────────────────────────────────────────────
function _hasSheet_(name) { try { return !!SpreadsheetApp.getActive().getSheetByName(name); } catch (_) { return false; } }

function _permsSheetReady_() {
  if (!_hasSheet_(G.CAMPAIGN_USER_PERMISSIONS_SHEET)) return false;
  const sh = SpreadsheetApp.getActive().getSheetByName(G.CAMPAIGN_USER_PERMISSIONS_SHEET);
  const lastCol = sh.getLastColumn(); if (!lastCol) return false;
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  return G.CAMPAIGN_USER_PERMISSIONS_HEADERS.every(h => headers.indexOf(h) !== -1);
}
function readCampaignPermsSafely_() {
  try {
    if (!_permsSheetReady_()) return [];
    const rows = readSheet(G.CAMPAIGN_USER_PERMISSIONS_SHEET) || [];
    if (!Array.isArray(rows) || !rows.length) return [];
    return rows.map(p => ({
      ID: p.ID,
      CampaignID: String(p.CampaignID || ''),
      UserID: String(p.UserID || ''),
      PermissionLevel: String(p.PermissionLevel || '').toUpperCase(),
      CanManageUsers: (p.CanManageUsers === true || String(p.CanManageUsers).toUpperCase() === 'TRUE'),
      CanManagePages: (p.CanManagePages === true || String(p.CanManagePages).toUpperCase() === 'TRUE'),
      CreatedAt: p.CreatedAt || null,
      UpdatedAt: p.UpdatedAt || null
    }));
  } catch (e) { writeError && writeError('readCampaignPermsSafely_', e); return []; }
}
function debugCampaignPerms() {
  const ready = _permsSheetReady_();
  return {
    sheetName: G.CAMPAIGN_USER_PERMISSIONS_SHEET,
    ready,
    note: ready ? 'Permissions sheet looks OK.' :
      'Sheet missing or headers not complete. Expected headers: ' + G.CAMPAIGN_USER_PERMISSIONS_HEADERS.join(', ')
  };
}

function getCampaignUserPermissions(campaignId, userId) {
  try {
    const rows = readSheet(G.CAMPAIGN_USER_PERMISSIONS_SHEET) || [];
    const row = rows.find(r => String(r.CampaignID) === String(campaignId) && String(r.UserID) === String(userId));
    const toBool = v => v === true || String(v).toUpperCase() === 'TRUE';
    if (!row) return { permissionLevel: 'USER', canManageUsers: false, canManagePages: false };
    return {
      permissionLevel: String(row.PermissionLevel || 'USER').toUpperCase(),
      canManageUsers: toBool(row.CanManageUsers),
      canManagePages: toBool(row.CanManagePages)
    };
  } catch (e) { writeError && writeError('getCampaignUserPermissions', e); return { permissionLevel: 'USER', canManageUsers: false, canManagePages: false }; }
}

function getCampaignPermsHeaders_() { return G.CAMPAIGN_USER_PERMISSIONS_HEADERS; }
function getOrCreateCampaignPermsSheet_() {
  if (typeof ensureSheetWithHeaders === 'function')
    return ensureSheetWithHeaders(G.CAMPAIGN_USER_PERMISSIONS_SHEET, getCampaignPermsHeaders_());
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(G.CAMPAIGN_USER_PERMISSIONS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(G.CAMPAIGN_USER_PERMISSIONS_SHEET);
    sh.getRange(1, 1, 1, getCampaignPermsHeaders_().length).setValues([getCampaignPermsHeaders_()]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function setCampaignUserPermissions(campaignId, userId, permissionLevel, canManageUsers, canManagePages) {
  try {
    if (!campaignId || !userId) return { success: false, error: 'campaignId and userId are required' };
    const lvl = String(permissionLevel || 'VIEWER').toUpperCase();
    const cmu = (canManageUsers === true || String(canManageUsers).toLowerCase() === 'true');
    const cmp = (canManagePages === true || String(canManagePages).toLowerCase() === 'true');

    const sh = getOrCreateCampaignPermsSheet_();
    const data = sh.getDataRange().getValues();
    const headers = data[0] || G.CAMPAIGN_USER_PERMISSIONS_HEADERS;
    const idx = {}; headers.forEach((h, i) => idx[String(h)] = i);

    ['ID', 'CampaignID', 'UserID', 'PermissionLevel', 'CanManageUsers', 'CanManagePages', 'CreatedAt', 'UpdatedAt']
      .forEach(h => { if (!(h in idx)) throw new Error('Campaign permissions sheet is missing header: ' + h); });

    let rowIndex = -1;
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][idx.CampaignID]) === String(campaignId) && String(data[r][idx.UserID]) === String(userId)) {
        rowIndex = r; break;
      }
    }
    const nowIso = new Date().toISOString();
    if (rowIndex === -1) {
      const row = [];
      row[idx.ID] = Utilities.getUuid();
      row[idx.CampaignID] = campaignId;
      row[idx.UserID] = userId;
      row[idx.PermissionLevel] = lvl;
      row[idx.CanManageUsers] = cmu ? 'TRUE' : 'FALSE';
      row[idx.CanManagePages] = cmp ? 'TRUE' : 'FALSE';
      row[idx.CreatedAt] = nowIso;
      row[idx.UpdatedAt] = nowIso;
      for (let c = 0; c < headers.length; c++) if (typeof row[c] === 'undefined') row[c] = '';
      sh.appendRow(row);
      if (typeof invalidateCache === 'function') invalidateCache(G.CAMPAIGN_USER_PERMISSIONS_SHEET);
      return { success: true, created: true, message: 'Permissions saved', campaignId, userId, permissionLevel: lvl, canManageUsers: cmu, canManagePages: cmp };
    } else {
      const r = rowIndex + 1;
      if (idx.PermissionLevel >= 0) sh.getRange(r, idx.PermissionLevel + 1).setValue(lvl);
      if (idx.CanManageUsers >= 0) sh.getRange(r, idx.CanManageUsers + 1).setValue(cmu ? 'TRUE' : 'FALSE');
      if (idx.CanManagePages >= 0) sh.getRange(r, idx.CanManagePages + 1).setValue(cmp ? 'TRUE' : 'FALSE');
      if (idx.UpdatedAt >= 0) sh.getRange(r, idx.UpdatedAt + 1).setValue(nowIso);
      if (typeof invalidateCache === 'function') invalidateCache(G.CAMPAIGN_USER_PERMISSIONS_SHEET);
      return { success: true, updated: true, message: 'Permissions saved', campaignId, userId, permissionLevel: lvl, canManageUsers: cmu, canManagePages: cmp };
    }
  } catch (e) { writeError && writeError('setCampaignUserPermissions', e); return { success: false, error: e.message }; }
}

// ───────────────────────────────────────────────────────────────────────────────
// Users: get all (campaign-aware) + safe mappers
// ───────────────────────────────────────────────────────────────────────────────
function clientGetAllUsers(requestingUserId) {
  try {
    let users = [];
    try { users = readSheet(G.USERS_SHEET); } catch (e) { writeError('clientGetAllUsers - readSheet', e); return []; }
    if (!Array.isArray(users) || users.length === 0) return [];

    const enhancedUsers = [];
    for (let i = 0; i < users.length; i++) {
      try {
        const u = users[i];
        if (!u || !u.ID) continue;
        enhancedUsers.push(createSafeUserObject(u));
      } catch (userErr) {
        enhancedUsers.push(createMinimalUserObject(users[i]));
      }
    }

    let filteredUsers = enhancedUsers;
    if (requestingUserId) {
      try {
        const requestingUser = users.find(u => String(u.ID) === String(requestingUserId));
        if (requestingUser) {
          if (isUserAdmin(requestingUser)) {
            filteredUsers = enhancedUsers;
          } else {
            const managedCampaigns = getUserManagedCampaigns(requestingUserId) || [];
            const managedSet = new Set(managedCampaigns.map(c => String(c.ID)));
            if (managedSet.size > 0) {
              const managedIds = new Set(Array.from(managedSet).map(String));
              filteredUsers = enhancedUsers.filter(u => {
                if (String(u.ID) === String(requestingUserId)) return true;
                const uCamps = (typeof getUserCampaignsSafe === 'function')
                  ? (getUserCampaignsSafe(u.ID) || []).map(x => String(x.campaignId))
                  : (u.CampaignID ? [String(u.CampaignID)] : []);
                return uCamps.some(cid => managedIds.has(cid));
              });
            } else {
              filteredUsers = enhancedUsers.filter(user => String(user.ID) === String(requestingUserId));
            }
          }
        } else {
          filteredUsers = [];
        }
      } catch (permissionError) {
        filteredUsers = enhancedUsers;
      }
    }
    return filteredUsers;
  } catch (globalError) { writeError('clientGetAllUsers', globalError); return []; }
}

function createSafeUserObject(user) {
  const safe = {
    ID: user.ID || '',
    UserName: user.UserName || user.userName || user.username || '', // Fix: Check all variations
    FullName: user.FullName || user.fullName || '',
    Email: user.Email || user.email || '',
    PhoneNumber: user.PhoneNumber || user.phoneNumber || '',
    CampaignID: user.CampaignID || user.campaignID || '',
    EmploymentStatus: user.EmploymentStatus || '',
    HireDate: user.HireDate || '',
    Country: user.Country || '',
    // HR/Benefits raw
    TerminationDate: user.TerminationDate || '',
    ProbationMonths: (user.ProbationMonths !== '' && user.ProbationMonths != null) ? Number(user.ProbationMonths) : '',
    ProbationEndDate: user.ProbationEndDate || '',
    InsuranceQualifiedDate: user.InsuranceQualifiedDate || '',
    InsuranceEligible: _strToBool_(user.InsuranceEligible || false),
    InsuranceSignedUp: _strToBool_(user.InsuranceSignedUp || false),
    InsuranceCardReceivedDate: user.InsuranceCardReceivedDate || ''
  };

  // Add convenience aliases
  safe.userName = safe.UserName; // Add camelCase alias for frontend
  safe.fullName = safe.FullName;
  safe.email = safe.Email;

  safe.InsuranceEligibilityDate = safe.InsuranceQualifiedDate;  // alias for UI
  safe.InsuranceEnrolled = safe.InsuranceSignedUp;              // alias for UI
  safe.InsuranceEnrolledBool = !!safe.InsuranceSignedUp;        // convenience
  safe.InsuranceQualifiedBool = !!safe.InsuranceEligible;       // convenience

  try {
    safe.canLoginBool = _strToBool_(user.CanLogin);
    safe.isAdminBool = _strToBool_(user.IsAdmin);
    safe.emailConfirmedBool = _strToBool_(user.EmailConfirmed);
    safe.needsPasswordSetup = !user.PasswordHash || _strToBool_(user.ResetRequired);
  } catch (e) {
    safe.canLoginBool = false; safe.isAdminBool = false; safe.emailConfirmedBool = false; safe.needsPasswordSetup = true;
  }

  try {
    safe.CreatedAt = user.CreatedAt ? new Date(user.CreatedAt).toISOString() : new Date().toISOString();
    safe.UpdatedAt = user.UpdatedAt ? new Date(user.UpdatedAt).toISOString() : new Date().toISOString();
    safe.LastLogin = user.LastLogin ? new Date(user.LastLogin).toISOString() : null;

    // nice formats
    safe.hireDateFormatted = user.HireDate ? new Date(user.HireDate).toLocaleDateString() : '';
    safe.terminationDateFormatted = user.TerminationDate ? new Date(user.TerminationDate).toLocaleDateString() : '';
    safe.probationEndDateFormatted = user.ProbationEndDate ? new Date(user.ProbationEndDate).toLocaleDateString() : '';
    safe.insuranceQualifiedDateFormatted = user.InsuranceQualifiedDate ? new Date(user.InsuranceQualifiedDate).toLocaleDateString() : '';
    safe.insuranceCardReceivedDateFormatted = user.InsuranceCardReceivedDate ? new Date(user.InsuranceCardReceivedDate).toLocaleDateString() : '';
  } catch (_) { }

  try { safe.campaignName = user.CampaignID ? getCampaignNameSafe(user.CampaignID) : ''; } catch (_) { safe.campaignName = ''; }

  try {
    const userRoleIds = getUserRoleIdsSafe(user.ID);
    const userRoles = getUserRolesSafe(user.ID);
    safe.roleIds = userRoleIds;
    safe.roles = userRoles.map(r => r.id);
    safe.roleNames = userRoles.map(r => r.name);
    if (safe.roleNames.length === 0 && user.Roles) {
      const csvRoles = String(user.Roles).split(',').filter(Boolean);
      safe.csvRoles = csvRoles;
      safe.roleNames = csvRoles.map(roleId => getRoleNameSafe(roleId));
    } else {
      safe.csvRoles = user.Roles ? String(user.Roles).split(',').filter(Boolean) : [];
    }
  } catch (_) {
    safe.roleIds = []; safe.roles = []; safe.roleNames = []; safe.csvRoles = [];
  }

  try { safe.pages = getUserPagesSafe(user.ID); } catch (_) { safe.pages = []; }
  try { safe.campaignPermissions = getUserCampaignPermissionsSafe(user.ID); } catch (_) { safe.campaignPermissions = []; }

  // Derived: InsuranceEligible recompute if needed
  try {
    const eligible = isInsuranceEligibleNow_(safe.InsuranceQualifiedDate, safe.TerminationDate);
    safe.InsuranceEligible = !!eligible;
  } catch (_) { }

  return safe;
}

function createMinimalUserObject(user) {
  return {
    ID: user.ID || '',
    UserName: user.UserName || '',
    FullName: user.FullName || '',
    Email: user.Email || '',
    PhoneNumber: user.PhoneNumber || '',
    EmploymentStatus: user.EmploymentStatus || '',
    HireDate: user.HireDate || null,
    Country: user.Country || '',
    canLoginBool: _strToBool_(user.CanLogin),
    isAdminBool: _strToBool_(user.IsAdmin),
    campaignName: '',
    roleNames: [],
    campaignPermissions: [],
    pages: [],
    CreatedAt: new Date().toISOString(),
    UpdatedAt: new Date().toISOString()
  };
}

// ───────────────────────────────────────────────────────────────────────────────
// Safe helpers
// ───────────────────────────────────────────────────────────────────────────────
function getCampaignNameSafe(campaignId) {
  try {
    if (!campaignId) return '';
    const campaigns = readSheet(G.CAMPAIGNS_SHEET) || [];
    const c = campaigns.find(x => String(x.ID) === String(campaignId));
    return c ? c.Name : '';
  } catch (e) { return ''; }
}
function getUserRoleIdsSafe(userId) {
  try {
    if (!userId) return [];
    const userRoles = readSheet(G.USER_ROLES_SHEET) || [];
    return userRoles
      .filter(ur => String(ur.UserId || ur.UserID) === String(userId))
      .map(ur => ur.RoleId || ur.RoleID)
      .filter(Boolean);
  } catch (e) { return []; }
}
function getUserRolesSafe(userId) {
  try {
    if (!userId) return [];
    const roleIds = getUserRoleIdsSafe(userId);
    if (roleIds.length === 0) return [];
    const allRoles = readSheet(G.ROLES_SHEET) || [];
    return roleIds.map(roleId => {
      const role = allRoles.find(r => r.ID === roleId);
      return role ? { id: role.ID, name: role.Name } : null;
    }).filter(Boolean);
  } catch (e) { return []; }
}
function getRoleNameSafe(roleId) {
  try {
    if (!roleId) return '';
    const roles = readSheet(G.ROLES_SHEET) || [];
    const role = roles.find(r => r.ID === roleId);
    return role ? role.Name : roleId;
  } catch (e) { return roleId; }
}
function getUserPagesSafe(userId) {
  try {
    if (!userId) return [];
    const users = readSheet(G.USERS_SHEET) || [];
    const user = users.find(u => u.ID === userId);
    if (!user || !user.Pages) return [];
    return String(user.Pages).split(',').map(p => p.trim()).filter(p => p);
  } catch (e) { return []; }
}
function getUserCampaignPermissionsSafe(userId) {
  try {
    if (!userId) return [];
    return readCampaignPermsSafely_().filter(p => String(p.UserID) === String(userId));
  } catch (e) { return []; }
}

// ───────────────────────────────────────────────────────────────────────────────
// Register / Update user (enhanced with benefits)
// ───────────────────────────────────────────────────────────────────────────────
if (typeof _buildUniqIndexes_ !== 'function') {
  function _buildUniqIndexes_(users) {
    const byEmail = new Map();
    const byUser = new Map();
    (users || []).forEach(u => {
      const e = _normEmail_(u && u.Email);
      const n = _normUser_(u && u.UserName);
      // Prefer first occurrence per key
      if (e && !byEmail.has(e)) byEmail.set(e, u);
      if (n && !byUser.has(n)) byUser.set(n, u);
    });
    return { byEmail, byUser };
  }
}


if (typeof _findRowIndexById_ !== 'function') {
  function _findRowIndexById_(values, idx, userId) {
    if (!values || !values.length) return -1;
    // Fallback to column 0 if idx.ID missing
    const col = (idx && typeof idx.ID === 'number') ? idx.ID : 0;
    for (var r = 1; r < values.length; r++) {
      if (String(values[r][col]) === String(userId)) return r;
    }
    return -1;
  }
}

function clientRegisterUser(userData) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { success: false, error: 'System is busy. Please try again.' }; }

  try {
    const v = _validateUserInput_(userData);
    if (!v.ok) return { success: false, error: v.errors.join('; ') };
    const data = _normalizeIncoming_(userData);

    const sh = _getSheet_(G.USERS_SHEET);
    let { headers, values, idx } = _scanSheet_(sh);
    _ensureUserHeaders_(idx);
    const ensured = ensureOptionalUserColumns_(sh, headers, idx);
    headers = ensured.headers; idx = ensured.idx;

    const users = _readUsersAsObjects_();
    const uniq = _buildUniqIndexes_(users);

    const emailKey = _normEmail_(data.email);
    const existing = emailKey ? uniq.byEmail.get(emailKey) : null;

    if (existing) {
      if (data.mergeIfExists === true || String(data.mergeIfExists).toLowerCase() === 'true') {
        const rowIndex = _findRowIndexById_(values, idx, existing.ID);
        if (rowIndex === -1) return { success: false, error: 'Existing user not found in sheet for merge' };

        if (data.userName && _normUser_(data.userName) !== _normUser_(existing.UserName)) {
          const taken = uniq.byUser.has(_normUser_(data.userName));
          if (taken) return { success: false, error: 'Username already exists' };
        }

        // recompute benefits if needed
        const hire = data.hireDate || existing.HireDate || '';
        const probMonths = (data.probationMonths != null) ? data.probationMonths : (existing.ProbationMonths || '');
        const probEnd = data.probationEndDate || calcProbationEndDate_(hire, probMonths);
        const insQual = data.insuranceQualifiedDate || calcInsuranceQualifiedDate_(probEnd, G.INSURANCE_MONTHS_AFTER_PROBATION);
        const insEligible = isInsuranceEligibleNow_(insQual, data.terminationDate || existing.TerminationDate);

        const merged = {
          ID: existing.ID,
          UserName: data.userName || existing.UserName,
          FullName: data.fullName || existing.FullName,
          Email: existing.Email,
          CampaignID: data.campaignId || existing.CampaignID,
          PasswordHash: values[rowIndex][idx['PasswordHash']],
          ResetRequired: values[rowIndex][idx['ResetRequired']],
          EmailConfirmation: values[rowIndex][idx['EmailConfirmation']],
          EmailConfirmed: 'TRUE',
          PhoneNumber: data.phoneNumber || existing.PhoneNumber,
          EmploymentStatus: data.employmentStatus || existing.EmploymentStatus,
          HireDate: hire || '',
          Country: data.country || existing.Country,
          LockoutEnd: values[rowIndex][idx['LockoutEnd']],
          TwoFactorEnabled: values[rowIndex][idx['TwoFactorEnabled']],
          CanLogin: _boolToStr_((data.canLogin != null) ? data.canLogin : _strToBool_(existing.CanLogin)),
          Roles: _csvUnion_(_splitCsv_(existing.Roles), data.roles).join(','),
          Pages: _csvUnion_(_splitCsv_(existing.Pages), data.pages).join(','),
          CreatedAt: values[rowIndex][idx['CreatedAt']],
          UpdatedAt: _now_(),
          IsAdmin: _boolToStr_((data.isAdmin != null) ? data.isAdmin : _strToBool_(existing.IsAdmin)),
          // Benefits
          TerminationDate: _toIsoDateOnly_(data.terminationDate || existing.TerminationDate || ''),
          ProbationMonths: (probMonths === '' ? '' : Number(probMonths)),
          ProbationEndDate: _toIsoDateOnly_(probEnd),
          InsuranceQualifiedDate: _toIsoDateOnly_(insQual),
          InsuranceEligible: _boolToStr_(insEligible),
          InsuranceSignedUp: _boolToStr_(data.insuranceSignedUp != null ? data.insuranceSignedUp : existing.InsuranceSignedUp),
          InsuranceCardReceivedDate: _toIsoDateOnly_(data.insuranceCardReceivedDate || existing.InsuranceCardReceivedDate || '')
        };

        const row = [];
        Object.keys(idx).forEach(header => {
          row[idx[header]] = (typeof merged[header] !== 'undefined') ? merged[header] : values[rowIndex][idx[header]];
        });
        sh.getRange(rowIndex + 1, 1, 1, headers.length).setValues([row]);

        try { invalidateCache && invalidateCache(G.USERS_SHEET); } catch (_) { }
        return { success: true, alreadyExisted: true, merged: true, userId: existing.ID, message: 'Existing user merged safely (roles/pages/fields updated).' };
      } else {
        return { success: true, alreadyExisted: true, userId: existing.ID, message: 'User already exists; no changes made.' };
      }
    }

    const userKey = _normUser_(data.userName);
    if (userKey && uniq.byUser.has(userKey)) return { success: false, error: 'Username already exists' };

    const id = Utilities.getUuid();
    const createdAt = _now_();
    const setupToken = Utilities.getUuid();

    // Benefits compute
    const probEnd = calcProbationEndDate_(data.hireDate || '', data.probationMonths || '');
    const insQual = calcInsuranceQualifiedDate_(probEnd, G.INSURANCE_MONTHS_AFTER_PROBATION);
    const insEligible = isInsuranceEligibleNow_(insQual, data.terminationDate);

    const newUser = {
      ID: id,
      UserName: data.userName,
      FullName: data.fullName || '',
      Email: data.email,
      CampaignID: data.campaignId || '',
      PasswordHash: '',
      ResetRequired: _boolToStr_(data.canLogin),
      EmailConfirmation: setupToken,
      EmailConfirmed: 'TRUE',
      PhoneNumber: data.phoneNumber || '',
      EmploymentStatus: data.employmentStatus || '',
      HireDate: data.hireDate || '',
      Country: data.country || '',
      LockoutEnd: '',
      TwoFactorEnabled: 'FALSE',
      CanLogin: _boolToStr_(data.canLogin),
      Roles: (data.roles || []).join(','),
      Pages: (data.pages || []).join(','),
      CreatedAt: createdAt,
      UpdatedAt: createdAt,
      IsAdmin: _boolToStr_(data.isAdmin),
      // Benefits fields
      TerminationDate: _toIsoDateOnly_(data.terminationDate || ''),
      ProbationMonths: (data.probationMonths === '' || data.probationMonths == null) ? '' : Number(data.probationMonths),
      ProbationEndDate: _toIsoDateOnly_(data.probationEndDate || probEnd),
      InsuranceQualifiedDate: _toIsoDateOnly_(data.insuranceQualifiedDate || insQual),
      InsuranceEligible: _boolToStr_(data.insuranceEligible != null ? data.insuranceEligible : insEligible),
      InsuranceSignedUp: _boolToStr_(data.insuranceSignedUp),
      InsuranceCardReceivedDate: _toIsoDateOnly_(data.insuranceCardReceivedDate || '')
    };

    const row = [];
    Object.keys(idx).forEach(header => { row[idx[header]] = (typeof newUser[header] !== 'undefined') ? newUser[header] : ''; });
    sh.appendRow(row);

    try {
      if (data.campaignId && data.permissionLevel && typeof setCampaignUserPermissions === 'function') {
        setCampaignUserPermissions(data.campaignId, id, String(data.permissionLevel).toUpperCase(),
          _strToBool_(data.canManageUsers), _strToBool_(data.canManagePages));
      }
      if (Array.isArray(data.roles) && typeof addUserRole === 'function') {
        data.roles.forEach(rid => addUserRole(id, rid));
      }
    } catch (pe) { writeError && writeError('clientRegisterUser:perms/roles', pe); }

    try { invalidateCache && invalidateCache(G.USERS_SHEET); } catch (_) { }

    if (data.canLogin && typeof sendPasswordSetupEmail === 'function') {
      try {
        sendPasswordSetupEmail(data.email, {
          userName: data.userName,
          fullName: data.fullName || data.userName,
          passwordSetupToken: setupToken
        });
      } catch (mailErr) { writeError && writeError('clientRegisterUser:sendPasswordSetupEmail', mailErr); }
    }

    try {
      notifyOnUserRegistered_ && notifyOnUserRegistered_({
        id, userName: data.userName, fullName: data.fullName || '', email: data.email,
        phoneNumber: data.phoneNumber || '', campaignId: data.campaignId || '',
        employmentStatus: data.employmentStatus || '', hireDate: data.hireDate || '',
        country: data.country || '', pages: data.pages || [], roles: data.roles || [],
        isAdmin: !!data.isAdmin, canLogin: !!data.canLogin, createdAt: createdAt
      });
    } catch (nerr) { writeError && writeError('clientRegisterUser:notify', nerr); }

    return {
      success: true,
      userId: id,
      message: (data.canLogin
        ? `User created. ${(data.pages || []).length} page(s) assigned. Password setup email sent.`
        : `User created. ${(data.pages || []).length} page(s) assigned. Login disabled.`)
    };

  } catch (e) {
    writeError && writeError('clientRegisterUser', e);
    return { success: false, error: e.message || String(e) };
  } finally {
    try { lock.releaseLock(); } catch (_) { }
  }
}

function clientUpdateUser(userId, userData) {
  console.log('Updating user', userId, 'with data:', userData); // Debug log
  
  if (!userId) return { success: false, error: 'User ID is required' };
  
  const lock = LockService.getScriptLock();
  try { 
    lock.waitLock(20000); 
  } catch (e) { 
    return { success: false, error: 'System is busy. Please try again.' }; 
  }

  try {
    // Fix: Ensure userName is properly mapped from all possible field names
    if (!userData.userName && userData.UserName) {
      userData.userName = userData.UserName;
    }
    if (!userData.userName && userData.username) {
      userData.userName = userData.username;
    }
    
    console.log('userName after mapping:', userData.userName); // Debug log
    
    const v = _validateUserInput_(userData);
    if (!v.ok) {
      console.error('Update validation failed:', v.errors);
      return { success: false, error: v.errors.join('; ') };
    }
    
    const data = _normalizeIncoming_(userData);
    
    // Additional check for userName after normalization
    if (!data.userName) {
      return { success: false, error: 'Username is required for update' };
    }

    const sh = _getSheet_(G.USERS_SHEET);
    let { headers, values, idx } = _scanSheet_(sh);
    _ensureUserHeaders_(idx);
    const ensured = ensureOptionalUserColumns_(sh, headers, idx);
    headers = ensured.headers; 
    idx = ensured.idx;

    const users = _readUsersAsObjects_();
    const uniq = _buildUniqIndexes_(users);

    const rowIndex = _findRowIndexById_(values, idx, userId);
    if (rowIndex === -1) {
      return { success: false, error: 'User not found in sheet' };
    }

    const current = {}; 
    headers.forEach((h, i) => current[h] = values[rowIndex][i]);
    
    console.log('Current user data:', current); // Debug log

    const desiredEmailKey = _normEmail_(data.email);
    const desiredUserKey = _normUser_(data.userName);

    // Check for email conflicts
    if (desiredEmailKey) {
      const hit = uniq.byEmail.get(desiredEmailKey);
      if (hit && String(hit.ID) !== String(userId)) {
        return { success: false, error: 'Email already exists for another user' };
      }
    }
    
    // Check for username conflicts
    if (desiredUserKey) {
      const hitU = uniq.byUser.get(desiredUserKey);
      if (hitU && String(hitU.ID) !== String(userId)) {
        return { success: false, error: 'Username already exists for another user' };
      }
    }

    // Benefits recomputation
    const hire = data.hireDate || current['HireDate'] || '';
    const probMonths = (data.probationMonths != null) ? data.probationMonths : (current['ProbationMonths'] || '');
    const probEnd = data.probationEndDate || calcProbationEndDate_(hire, probMonths);
    const term = data.terminationDate || current['TerminationDate'] || '';
    const insQual = data.insuranceQualifiedDate || calcInsuranceQualifiedDate_(probEnd, G.INSURANCE_MONTHS_AFTER_PROBATION);
    const insEligible = (data.insuranceEligible != null) ? data.insuranceEligible : isInsuranceEligibleNow_(insQual, term);

    const updated = {
      ID: userId,
      UserName: data.userName, // This should now be properly set
      FullName: data.fullName || '',
      Email: String(data.email).trim(),
      CampaignID: data.campaignId || '',
      PasswordHash: current['PasswordHash'],
      ResetRequired: current['ResetRequired'],
      EmailConfirmation: current['EmailConfirmation'],
      EmailConfirmed: 'TRUE',
      PhoneNumber: data.phoneNumber || '',
      EmploymentStatus: data.employmentStatus || '',
      HireDate: hire || '',
      Country: data.country || '',
      LockoutEnd: current['LockoutEnd'],
      TwoFactorEnabled: current['TwoFactorEnabled'],
      CanLogin: _boolToStr_(data.canLogin),
      Roles: (data.roles || []).join(','),
      Pages: (data.pages || []).join(','),
      CreatedAt: current['CreatedAt'],
      UpdatedAt: _now_(),
      IsAdmin: _boolToStr_(data.isAdmin),
      // Benefits
      TerminationDate: _toIsoDateOnly_(term),
      ProbationMonths: (probMonths === '' ? '' : Number(probMonths)),
      ProbationEndDate: _toIsoDateOnly_(probEnd),
      InsuranceQualifiedDate: _toIsoDateOnly_(insQual),
      InsuranceEligible: _boolToStr_(insEligible),
      InsuranceSignedUp: _boolToStr_(data.insuranceSignedUp != null ? data.insuranceSignedUp : current['InsuranceSignedUp']),
      InsuranceCardReceivedDate: _toIsoDateOnly_(data.insuranceCardReceivedDate || current['InsuranceCardReceivedDate'] || '')
    };

    console.log('Updated user object:', updated); // Debug log

    const row = [];
    Object.keys(idx).forEach(header => {
      row[idx[header]] = (typeof updated[header] !== 'undefined') ? updated[header] : current[header];
    });
    
    sh.getRange(rowIndex + 1, 1, 1, headers.length).setValues([row]);

    console.log('User updated successfully with userName:', data.userName);

    try {
      if (data.campaignId && data.permissionLevel && typeof setCampaignUserPermissions === 'function') {
        setCampaignUserPermissions(data.campaignId, userId, String(data.permissionLevel).toUpperCase(),
          _strToBool_(data.canManageUsers), _strToBool_(data.canManagePages));
      }
    } catch (pe) { 
      writeError && writeError('clientUpdateUser:perms', pe); 
    }

    try {
      invalidateCache && invalidateCache(G.USERS_SHEET);
      invalidateCache && invalidateCache(G.CAMPAIGN_USER_PERMISSIONS_SHEET);
    } catch (_) { }

    return { 
      success: true, 
      message: `User updated successfully with username "${data.userName}". ${(data.pages || []).length} page(s) assigned.` 
    };

  } catch (e) {
    console.error('clientUpdateUser error:', e);
    writeError && writeError('clientUpdateUser', e);
    return { success: false, error: e.message || String(e) };
  } finally {
    try { 
      lock.releaseLock(); 
    } catch (_) { }
  }
}

function debugUsernameIssues(userId = null) {
  try {
    const results = {
      sheetStatus: null,
      userSample: null,
      specificUser: null,
      headerMapping: null,
      recommendations: []
    };

    // Check sheet structure
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(G.USERS_SHEET);
    
    if (!sheet) {
      results.sheetStatus = { error: 'Users sheet not found' };
      results.recommendations.push('Create Users sheet with proper headers');
      return results;
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow < 2) {
      results.sheetStatus = { error: 'No user data found', dimensions: { rows: lastRow, cols: lastCol } };
      results.recommendations.push('Add at least one user to the sheet');
      return results;
    }

    // Get headers and check UserName column
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const userNameIndex = headers.indexOf('UserName');
    
    results.headerMapping = {
      allHeaders: headers,
      userNameIndex: userNameIndex,
      userNameExists: userNameIndex !== -1
    };

    if (userNameIndex === -1) {
      results.recommendations.push('Add UserName column to Users sheet headers');
      return results;
    }

    // Get sample data
    const sampleRows = Math.min(5, lastRow - 1);
    const sampleData = sheet.getRange(2, 1, sampleRows, lastCol).getValues();
    
    results.userSample = {
      count: sampleRows,
      data: sampleData.map((row, index) => ({
        rowNumber: index + 2,
        ID: row[0],
        UserName: row[userNameIndex],
        FullName: row[headers.indexOf('FullName')],
        Email: row[headers.indexOf('Email')]
      }))
    };

    // Check for empty usernames in sample
    const emptyUsernames = results.userSample.data.filter(user => !user.UserName || user.UserName === '');
    if (emptyUsernames.length > 0) {
      results.recommendations.push(`Found ${emptyUsernames.length} users with empty usernames in sample`);
    }

    // Check specific user if provided
    if (userId) {
      const userData = sheet.getDataRange().getValues();
      const userRow = userData.find((row, index) => index > 0 && String(row[0]) === String(userId));
      
      if (userRow) {
        results.specificUser = {
          found: true,
          ID: userRow[0],
          UserName: userRow[userNameIndex],
          FullName: userRow[headers.indexOf('FullName')],
          Email: userRow[headers.indexOf('Email')],
          rawRow: userRow
        };
        
        if (!results.specificUser.UserName || results.specificUser.UserName === '') {
          results.recommendations.push(`Specific user ${userId} has empty UserName field`);
        }
      } else {
        results.specificUser = { found: false, searchedId: userId };
        results.recommendations.push(`User with ID ${userId} not found in sheet`);
      }
    }

    // Test createSafeUserObject function
    if (results.userSample.data.length > 0) {
      try {
        const firstUser = results.userSample.data[0];
        const rawUserData = {};
        headers.forEach((header, index) => {
          rawUserData[header] = sampleData[0][index];
        });
        
        const safeUser = createSafeUserObject(rawUserData);
        results.createSafeUserTest = {
          success: true,
          input: rawUserData,
          output: safeUser,
          userNameMapped: !!safeUser.UserName
        };
        
        if (!safeUser.UserName) {
          results.recommendations.push('createSafeUserObject is not properly mapping UserName field');
        }
      } catch (error) {
        results.createSafeUserTest = {
          success: false,
          error: error.message
        };
        results.recommendations.push('createSafeUserObject function has errors');
      }
    }

    results.sheetStatus = { 
      success: true, 
      dimensions: { rows: lastRow, cols: lastCol },
      userNameColumn: userNameIndex + 1 // 1-based for sheet reference
    };

    if (results.recommendations.length === 0) {
      results.recommendations.push('No issues detected in username handling');
    }

    return results;

  } catch (error) {
    return {
      error: error.message,
      stack: error.stack,
      recommendations: ['Contact system administrator - unexpected error occurred']
    };
  }
}

function repairUsernamesInSheet() {
  try {
    const results = {
      processed: 0,
      repaired: 0,
      errors: [],
      details: []
    };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(G.USERS_SHEET);
    
    if (!sheet) {
      return { success: false, error: 'Users sheet not found' };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { success: false, error: 'No user data to process' };
    }

    const headers = data[0];
    const userNameIndex = headers.indexOf('UserName');
    const fullNameIndex = headers.indexOf('FullName');
    const emailIndex = headers.indexOf('Email');
    const updatedAtIndex = headers.indexOf('UpdatedAt');

    if (userNameIndex === -1) {
      return { success: false, error: 'UserName column not found' };
    }

    // Process each user row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowNumber = i + 1;
      
      results.processed++;
      
      try {
        let userName = row[userNameIndex];
        let needsRepair = false;
        
        // Check if username is empty or invalid
        if (!userName || userName === '' || typeof userName !== 'string') {
          needsRepair = true;
          
          // Try to generate username from email or fullname
          const email = row[emailIndex];
          const fullName = row[fullNameIndex];
          
          if (email && typeof email === 'string') {
            userName = email.split('@')[0]; // Use part before @ as username
          } else if (fullName && typeof fullName === 'string') {
            userName = fullName.toLowerCase().replace(/\s+/g, ''); // Remove spaces from full name
          } else {
            userName = `user${Date.now()}_${Math.floor(Math.random() * 1000)}`; // Generate random username
          }
          
          // Ensure username is valid
          userName = userName.substring(0, 50); // Limit length
          
          results.details.push({
            rowNumber: rowNumber,
            oldUserName: row[userNameIndex],
            newUserName: userName,
            method: email ? 'email' : (fullName ? 'fullname' : 'generated')
          });
        }
        
        if (needsRepair) {
          // Update the username in the sheet
          sheet.getRange(rowNumber, userNameIndex + 1).setValue(userName);
          
          // Update the UpdatedAt timestamp if column exists
          if (updatedAtIndex !== -1) {
            sheet.getRange(rowNumber, updatedAtIndex + 1).setValue(new Date());
          }
          
          results.repaired++;
        }
        
      } catch (error) {
        results.errors.push({
          rowNumber: rowNumber,
          error: error.message
        });
      }
    }

    // Clear cache after repairs
    try {
      invalidateCache && invalidateCache(G.USERS_SHEET);
    } catch (_) {}

    return {
      success: true,
      ...results,
      message: `Processed ${results.processed} users, repaired ${results.repaired} usernames`
    };

  } catch (error) {
    return {
      success: false,
      error: error.message,
      stack: error.stack
    };
  }
}

function clientDebugUsernameIssues(userId = null) {
  try {
    return debugUsernameIssues(userId);
  } catch (error) {
    writeError && writeError('clientDebugUsernameIssues', error);
    return { success: false, error: error.message };
  }
}

function clientRepairUsernames() {
  try {
    return repairUsernamesInSheet();
  } catch (error) {
    writeError && writeError('clientRepairUsernames', error);
    return { success: false, error: error.message };
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// Password / email utilities (admin paths)
// ───────────────────────────────────────────────────────────────────────────────
function clientAdminResetPassword(userId, requestingUserId) {
  try {
    if (!userId) return { success: false, error: 'User ID is required' };
    if (requestingUserId) {
      const allUsers = readSheet(G.USERS_SHEET) || [];
      const requester = allUsers.find(u => String(u.ID) === String(requestingUserId));
      if (!requester || !_strToBool_(requester.IsAdmin)) return { success: false, error: 'Only administrators can perform this action' };
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(G.USERS_SHEET);
    const data = sheet.getDataRange().getValues();
    const headers = data[0]; const idx = {}; headers.forEach((h, i) => idx[String(h)] = i);

    let rowIndex = -1, row;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(userId)) { rowIndex = i; row = data[i]; break; }
    }
    if (rowIndex === -1) return { success: false, error: 'User not found' };

    const email = row[idx['Email']] || '';
    if (!email) return { success: false, error: 'User has no email on file' };

    const canLogin = _strToBool_(row[idx['CanLogin']]);
    if (!canLogin) return { success: false, error: 'User cannot login (CanLogin is FALSE)' };

    const token = Utilities.getUuid();
    if (idx['EmailConfirmation'] >= 0) sheet.getRange(rowIndex + 1, idx['EmailConfirmation'] + 1).setValue(token);
    if (idx['ResetRequired'] >= 0) sheet.getRange(rowIndex + 1, idx['ResetRequired'] + 1).setValue('TRUE');
    if (idx['UpdatedAt'] >= 0) sheet.getRange(rowIndex + 1, idx['UpdatedAt'] + 1).setValue(new Date());

    invalidateCache && invalidateCache(G.USERS_SHEET);

    try {
      if (typeof sendAdminPasswordResetEmail === 'function') {
        sendAdminPasswordResetEmail(email, { resetToken: token });
      } else if (typeof sendPasswordResetEmail === 'function') {
        sendPasswordResetEmail(email, token);
      } else {
        throw new Error('No email template available (sendAdminPasswordResetEmail/sendPasswordResetEmail)');
      }
    } catch (mailErr) {
      writeError && writeError('clientAdminResetPassword:send', mailErr);
      return { success: false, error: 'Token saved, but failed to send email: ' + mailErr };
    }
    return { success: true, message: 'Password reset email sent' };
  } catch (e) { writeError && writeError('clientAdminResetPassword', e); return { success: false, error: e.message }; }
}
function clientAdminResetPasswordById(userId, requestingUserId) { return clientAdminResetPassword(userId, requestingUserId); }

function clientResendFirstLoginEmail(userId, requestingUserId) {
  try {
    if (!userId) return { success: false, error: 'User ID is required' };
    if (requestingUserId) {
      const allUsers = readSheet(G.USERS_SHEET) || [];
      const requester = allUsers.find(u => String(u.ID) === String(requestingUserId));
      if (!requester || !_strToBool_(requester.IsAdmin)) return { success: false, error: 'Only administrators can perform this action' };
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(G.USERS_SHEET);
    const data = sheet.getDataRange().getValues();
    const headers = data[0]; const idx = {}; headers.forEach((h, i) => idx[String(h)] = i);

    let rowIndex = -1, row;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(userId)) { rowIndex = i; row = data[i]; break; }
    }
    if (rowIndex === -1) return { success: false, error: 'User not found' };

    const email = row[idx['Email']] || '';
    const userName = row[idx['UserName']] || '';
    const fullName = row[idx['FullName']] || '';
    if (!email) return { success: false, error: 'User has no email on file' };

    const canLogin = _strToBool_(row[idx['CanLogin']]);
    if (!canLogin) return { success: false, error: 'User cannot login (CanLogin is FALSE)' };

    const token = Utilities.getUuid();
    if (idx['EmailConfirmation'] >= 0) sheet.getRange(rowIndex + 1, idx['EmailConfirmation'] + 1).setValue(token);
    if (idx['ResetRequired'] >= 0) sheet.getRange(rowIndex + 1, idx['ResetRequired'] + 1).setValue('TRUE');
    if (idx['UpdatedAt'] >= 0) sheet.getRange(rowIndex + 1, idx['UpdatedAt'] + 1).setValue(new Date());

    invalidateCache && invalidateCache(G.USERS_SHEET);

    try {
      if (typeof sendFirstLoginResendEmail === 'function') {
        sendFirstLoginResendEmail(email, { userName, fullName, passwordSetupToken: token });
      } else if (typeof sendPasswordSetupEmail === 'function') {
        sendPasswordSetupEmail(email, { userName, fullName, passwordSetupToken: token });
      } else {
        throw new Error('No email template available (sendFirstLoginResendEmail/sendPasswordSetupEmail)');
      }
    } catch (mailErr) {
      writeError && writeError('clientResendFirstLoginEmail:send', mailErr);
      return { success: false, error: 'Token saved, but failed to send email: ' + mailErr };
    }
    return { success: true, message: 'First-login/setup email resent' };
  } catch (e) { writeError && writeError('clientResendFirstLoginEmail', e); return { success: false, error: e.message }; }
}

// ───────────────────────────────────────────────────────────────────────────────
// Delete user (also clears campaign perms, roles, manager links)
// ───────────────────────────────────────────────────────────────────────────────
function clientDeleteUser(userId) {
  try {
    if (!userId) return { success: false, error: 'User ID is required' };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(G.USERS_SHEET);
    if (!sheet) return { success: false, error: 'Users sheet not found' };

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1; let userName = '';
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(userId)) { rowIndex = i; userName = data[i][1] || data[i][2] || 'Unknown User'; break; }
    }
    if (rowIndex === -1) return { success: false, error: 'User not found in database' };

    const permissionsSheet = ss.getSheetByName(G.CAMPAIGN_USER_PERMISSIONS_SHEET);
    if (permissionsSheet) {
      const permissionsData = permissionsSheet.getDataRange().getValues();
      for (let i = permissionsData.length - 1; i >= 1; i--) {
        if (String(permissionsData[i][2]) === String(userId)) permissionsSheet.deleteRow(i + 1);
      }
    }

    deleteUserRoles(userId);

    try {
      const muSheet = getOrCreateManagerUsersSheet_();
      const muData = muSheet.getDataRange().getValues();
      const headers = muData[0] || [];
      const midx = { ManagerUserID: headers.indexOf('ManagerUserID'), UserID: headers.indexOf('UserID') };
      for (let i = muData.length - 1; i >= 1; i--) {
        if (String(muData[i][midx.ManagerUserID]) === String(userId) || String(muData[i][midx.UserID]) === String(userId)) {
          muSheet.deleteRow(i + 1);
        }
      }
      invalidateCache && invalidateCache(getManagerUsersSheetName_());
    } catch (e) { }

    sheet.deleteRow(rowIndex + 1);

    try {
      invalidateCache && invalidateCache(G.USERS_SHEET);
      invalidateCache && invalidateCache(G.CAMPAIGN_USER_PERMISSIONS_SHEET);
      invalidateCache && invalidateCache(G.USER_ROLES_SHEET);
    } catch (_) { }

    return { success: true, message: `User "${userName}" has been deleted successfully`, deletedUserId: userId };
  } catch (e) {
    writeError && writeError('clientDeleteUser', e);
    return { success: false, error: 'Failed to delete user: ' + e.message };
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// Available pages / discovery / assign pages
// ───────────────────────────────────────────────────────────────────────────────
function clientGetAvailablePages() {
  try {
    if (typeof enhancedAutoDiscoverAndSavePages === 'function') enhancedAutoDiscoverAndSavePages({ minIntervalSec: 300 });
    const pages = readSheet(G.PAGES_SHEET) || [];
    return pages.map(p => ({
      key: p.PageKey,
      title: p.PageTitle || p.Name,
      icon: p.PageIcon || p.Icon,
      description: p.Description,
      isSystem: (p.IsSystemPage === true || String(p.IsSystemPage).toUpperCase() === 'TRUE'),
      requiresAdmin: (p.RequiresAdmin === true || String(p.RequiresAdmin).toUpperCase() === 'TRUE')
    })).sort((a, b) => (a.title || '').localeCompare(b.title || ''));
  } catch (e) { writeError && writeError('clientGetAvailablePages', e); return []; }
}
function clientRunEnhancedDiscovery() {
  try {
    let res = {};
    if (typeof enhancedAutoDiscoverAndSavePages === 'function') {
      res = enhancedAutoDiscoverAndSavePages({ force: true }) || {};
    }
    const pages = readSheet(G.PAGES_SHEET) || [];
    const out = { success: true, total: pages.length };
    if (res && typeof res === 'object') {
      out.added = res.added || 0; out.updated = res.updated || 0;
      out.newPages = res.newPages || []; out.categories = res.categories || {};
      out.message = res.message || 'Discovery completed.';
    } else {
      out.added = 0; out.updated = 0; out.newPages = []; out.categories = {}; out.message = 'Discovery completed.';
    }
    return out;
  } catch (e) { writeError && writeError('clientRunEnhancedDiscovery', e); return { success: false, error: e.message }; }
}
function clientGetUserPages(userId) {
  try {
    if (!userId) return [];
    const users = readSheet(G.USERS_SHEET) || [];
    const user = users.find(u => u.ID === userId);
    if (!user || !user.Pages) return [];
    return String(user.Pages).split(',').map(p => p.trim()).filter(Boolean);
  } catch (e) { writeError && writeError('clientGetUserPages', e); return []; }
}
function clientAssignPagesToUser(userId, pageKeys) {
  try {
    if (!userId || !Array.isArray(pageKeys)) return { success: false, error: 'Invalid parameters' };
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(G.USERS_SHEET);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const pagesIndex = headers.indexOf('Pages');
    const updatedAtIndex = headers.indexOf('UpdatedAt');
    if (pagesIndex === -1) return { success: false, error: 'Pages column not found' };

    let userRowIndex = -1;
    for (let i = 1; i < data.length; i++) if (String(data[i][0]) === String(userId)) { userRowIndex = i; break; }
    if (userRowIndex === -1) return { success: false, error: 'User not found' };

    const pagesString = pageKeys.join(',');
    sheet.getRange(userRowIndex + 1, pagesIndex + 1).setValue(pagesString);
    if (updatedAtIndex !== -1) sheet.getRange(userRowIndex + 1, updatedAtIndex + 1).setValue(new Date());
    try { invalidateCache && invalidateCache(G.USERS_SHEET); } catch (_) { }
    return { success: true, message: `Assigned ${pageKeys.length} pages to user` };
  } catch (e) { writeError && writeError('clientAssignPagesToUser', e); return { success: false, error: 'Failed to assign pages: ' + e.message }; }
}

// ───────────────────────────────────────────────────────────────────────────────
// Manager assignments
// ───────────────────────────────────────────────────────────────────────────────
function clientGetUserPermissions(userId, campaignId) { return getCampaignUserPermissions(campaignId, userId); }
function clientSetUserPermissions(userId, campaignId, permissionLevel, canManageUsers, canManagePages) {
  return setCampaignUserPermissions(campaignId, userId, permissionLevel, canManageUsers, canManagePages);
}
function canUserManageOthers(userId) {
  try {
    const users = readSheet(G.USERS_SHEET) || [];
    const u = users.find(x => x.ID === userId);
    if (u && _strToBool_(u.IsAdmin)) return true;
    const perms = readSheet(G.CAMPAIGN_USER_PERMISSIONS_SHEET) || [];
    const userPerms = perms.filter(p => String(p.UserID) === String(userId));
    const toBool = v => v === true || String(v).toUpperCase() === 'TRUE';
    return userPerms.some(p =>
      ['MANAGER', 'ADMIN'].includes(String(p.PermissionLevel).toUpperCase()) || toBool(p.CanManageUsers)
    );
  } catch (e) { writeError && writeError('canUserManageOthers', e); return false; }
}
function clientGetAvailableUsersForManager(managerUserId) {
  try {
    if (!managerUserId) return { success: false, error: 'Manager ID is required' };
    const users = readSheet(G.USERS_SHEET) || [];
    const manager = users.find(u => String(u.ID) === String(managerUserId));
    if (!manager) return { success: false, error: 'Manager not found' };

    const admin = _strToBool_(manager.IsAdmin);
    let allowedCampaignIds = [];
    if (admin) {
      allowedCampaignIds = (readSheet(G.CAMPAIGNS_SHEET) || []).map(c => c.ID);
    } else {
      allowedCampaignIds = getManagedCampaignIdsForUser_(managerUserId);
    }
    const list = users
      .filter(u => String(u.ID) !== String(managerUserId))
      .filter(u => {
        const allowed = new Set((allowedCampaignIds || []).map(String));
        if (!allowed.size) return true;
        const uCamps = (typeof getUserCampaignsSafe === 'function')
          ? (getUserCampaignsSafe(u.ID) || []).map(x => String(x.campaignId))
          : (u.CampaignID ? [String(u.CampaignID)] : []);
        return uCamps.some(cid => allowed.has(cid));
      })
      .map(u => ({
        ID: u.ID, UserName: u.UserName, FullName: u.FullName, Email: u.Email,
        CampaignID: u.CampaignID, campaignName: getCampaignNameSafe(u.CampaignID),
        roleNames: getUserRolesSafe(u.ID).map(r => r.name)
      }));

    return { success: true, users: list };
  } catch (e) { writeError && writeError('clientGetAvailableUsersForManager', e); return { success: false, error: e.message }; }
}
function clientGetManagedUsers(managerUserId) {
  try {
    if (!managerUserId) return { success: false, error: 'Manager ID is required' };
    const sh = getOrCreateManagerUsersSheet_();
    const data = sh.getDataRange().getValues();
    const headers = data[0] || [];
    const midx = { ManagerUserID: headers.indexOf('ManagerUserID'), UserID: headers.indexOf('UserID') };
    const users = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][midx.ManagerUserID]) === String(managerUserId)) users.push({ ID: data[i][midx.UserID] });
    }
    return { success: true, users };
  } catch (e) { writeError && writeError('clientGetManagedUsers', e); return { success: false, error: e.message }; }
}
function clientAssignUsersToManager(managerUserId, userIds) {
  try {
    if (!managerUserId || !Array.isArray(userIds)) return { success: false, error: 'Invalid parameters' };
    const users = readSheet(G.USERS_SHEET) || [];
    const manager = users.find(u => String(u.ID) === String(managerUserId));
    if (!manager) return { success: false, error: 'Manager not found' };

    const admin = _strToBool_(manager.IsAdmin);
    const allowedCampaignIds = admin ? (readSheet(G.CAMPAIGNS_SHEET) || []).map(c => c.ID)
      : getManagedCampaignIdsForUser_(managerUserId);
    const allowedSet = new Set((allowedCampaignIds || []).map(String));
    const eligible = userIds.filter(id => {
      const u = users.find(x => String(x.ID) === String(id));
      if (!u) return false;
      if (admin || !allowedSet.size) return true;
      const uCamps = (typeof getUserCampaignsSafe === 'function')
        ? (getUserCampaignsSafe(u.ID) || []).map(x => String(x.campaignId))
        : (u.CampaignID ? [String(u.CampaignID)] : []);
      return uCamps.some(cid => allowedSet.has(cid));
    });

    const sh = getOrCreateManagerUsersSheet_();
    const data = sh.getDataRange().getValues();
    const headers = data[0] || [];
    const idx = { ID: headers.indexOf('ID'), ManagerUserID: headers.indexOf('ManagerUserID'), UserID: headers.indexOf('UserID') };

    const current = new Set();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idx.ManagerUserID]) === String(managerUserId)) current.add(String(data[i][idx.UserID]));
    }

    const target = new Set(eligible.map(String));
    const toAdd = [...target].filter(id => !current.has(id));
    const toRemove = [...current].filter(id => !target.has(id));

    for (let r = data.length - 1; r >= 1; r--) {
      const row = data[r];
      if (String(row[idx.ManagerUserID]) === String(managerUserId) && toRemove.includes(String(row[idx.UserID]))) {
        sh.deleteRow(r + 1);
      }
    }
    const now = new Date();
    toAdd.forEach(uid => { sh.appendRow([Utilities.getUuid(), managerUserId, uid, now, now]); });

    invalidateCache && invalidateCache(getManagerUsersSheetName_());
    try { notifyOnManagerAssignment_(managerUserId, toAdd); } catch (_) { }

    return { success: true, message: `Assigned ${target.size} user(s) to manager`, added: toAdd.length, removed: toRemove.length };
  } catch (e) { writeError && writeError('clientAssignUsersToManager', e); return { success: false, error: e.message }; }
}
function getManagedCampaignIdsForUser_(userId) {
  const perms = readCampaignPermsSafely_();
  return perms
    .filter(p => String(p.UserID) === String(userId) && ['MANAGER', 'ADMIN'].includes(String(p.PermissionLevel || '').toUpperCase()))
    .map(p => p.CampaignID);
}
function notifyOnManagerAssignment_(managerUserId, userIds) {
  if (!userIds || !userIds.length) return;
  try {
    const users = readSheet(G.USERS_SHEET) || [];
    const mgr = users.find(u => String(u.ID) === String(managerUserId));
    if (!mgr) return;
    const managerName = mgr.FullName || mgr.UserName || 'Your manager';
    const html = buildHtmlEmail_('Manager Assigned', 'You have been assigned a manager.', '<p>You have been assigned to <strong>' + managerName + '</strong>.</p>');
    userIds.forEach(uid => {
      const u = users.find(x => String(x.ID) === String(uid));
      if (u && u.Email) safeSendHtmlEmail_(u.Email, 'You have a new manager', html);
    });
  } catch (e) { writeError && writeError('notifyOnManagerAssignment_', e); }
}
function getManagerUsersSheetName_() { return G.MANAGER_USERS_SHEET; }
function getManagerUsersHeaders_() { return G.MANAGER_USERS_HEADER; }
function getOrCreateManagerUsersSheet_() {
  if (typeof ensureSheetWithHeaders === 'function') return ensureSheetWithHeaders(getManagerUsersSheetName_(), getManagerUsersHeaders_());
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = getManagerUsersSheetName_();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(getManagerUsersHeaders_()); }
  return sh;
}

// ───────────────────────────────────────────────────────────────────────────────
/** Roles **/
// ───────────────────────────────────────────────────────────────────────────────
function getAllRoles() {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(G.ROLES_SHEET);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    const rows = data.slice(1);
    return rows.map(r => ({ ID: r[0], Name: r[1], normalizedName: r[2] }));
  } catch (e) { writeError && writeError('getAllRoles', e); return []; }
}
function addUserRole(userId, roleId) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(G.USER_ROLES_SHEET);
    if (!sheet) return;
    const id = Utilities.getUuid(); const now = new Date();
    sheet.appendRow([id, userId, roleId, now, now]);
    invalidateCache && invalidateCache(G.USER_ROLES_SHEET);
  } catch (e) { writeError && writeError('addUserRole', e); }
}
function deleteUserRoles(userId) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(G.USER_ROLES_SHEET);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][1]) === String(userId)) sheet.deleteRow(i + 1);
    }
    invalidateCache && invalidateCache(G.USER_ROLES_SHEET);
  } catch (e) { writeError && writeError('deleteUserRoles', e); }
}
function getRolesMapping() {
  try {
    const roles = getAllRoles();
    const mapping = {}; roles.forEach(r => mapping[r.ID] = r.Name);
    return mapping;
  } catch (error) { writeError && writeError('getRolesMapping', error); return {}; }
}

// ───────────────────────────────────────────────────────────────────────────────
/** Utilities, validation, emails **/
// ───────────────────────────────────────────────────────────────────────────────
function getAllCampaigns() { try { return readSheet(G.CAMPAIGNS_SHEET) || []; } catch (e) { writeError && writeError('getAllCampaigns', e); return []; } }
function getCampaignName(campaignId) {
  try {
    const campaigns = readSheet(G.CAMPAIGNS_SHEET) || []; const campaign = campaigns.find(c => c.ID === campaignId);
    return campaign ? campaign.Name : '';
  } catch (e) { writeError && writeError('getCampaignName', e); return ''; }
}

function _normEmail_(e) { return String(e || '').trim().toLowerCase(); }
function _normUser_(u) { return String(u || '').trim().toLowerCase(); }
function _splitCsv_(s) { return String(s || '').split(',').map(x => x.trim()).filter(Boolean); }
function _csvUnion_(a1, a2) { const set = new Set([...(a1 || []), ...(a2 || [])].map(x => String(x).trim()).filter(Boolean)); return Array.from(set); }
function _now_() { return new Date(); }

function _readUsersAsObjects_() {
  try { if (typeof readSheet === 'function') return readSheet(G.USERS_SHEET) || []; } catch (e) { writeError && writeError('_readUsersAsObjects_', e); }
  const sh = _getSheet_(G.USERS_SHEET);
  const { headers, values } = _scanSheet_(sh);
  const out = []; for (let r = 1; r < values.length; r++) { const row = values[r]; const o = {}; headers.forEach((h, i) => o[h] = row[i]); out.push(o); }
  return out;
}
function _logErr_(where, err) { try { writeError && writeError(where, err); } catch (_) { } }

function _validateUserInput_(userData) {
  const errors = [];
  if (!userData || !String(userData.userName).trim()) errors.push('Username is required');
  if (!userData || !String(userData.email).trim()) errors.push('Email is required');

  const email = String(userData && userData.email || '').trim();
  if (email && !/^[^@]+@[^@]+\.[^@]+$/.test(email)) errors.push('Email appears invalid');

  if (userData.employmentStatus && !validateEmploymentStatus(userData.employmentStatus)) {
    errors.push('Invalid employment status. Valid options: ' + getValidEmploymentStatuses().join(', '));
  }
  if (userData.hireDate && !validateHireDate(userData.hireDate)) {
    errors.push('Invalid hire date. Must be a valid date not in the future');
  }
  if (userData.country && !validateCountry(userData.country)) {
    errors.push('Country must be between 2 and 100 characters');
  }
  // dates sanity (optional)
  ['terminationDate', 'probationEndDate', 'insuranceQualifiedDate', 'insuranceCardReceivedDate'].forEach(k => {
    if (userData[k]) {
      const d = new Date(userData[k]);
      if (isNaN(d)) errors.push(`${k} is not a valid date`);
    }
  });
  if (userData.probationMonths != null && userData.probationMonths !== '') {
    const n = Number(userData.probationMonths);
    if (isNaN(n) || n < 0 || n > 24) errors.push('probationMonths must be 0..24');
  }
  return { ok: errors.length === 0, errors };
}

function _normalizeIncoming_(userData) {
  const out = Object.assign({}, userData || {});
  out.userName = String(out.userName || '').trim();
  out.fullName = String(out.fullName || '').trim();
  out.email = String(out.email || '').trim();
  out.phoneNumber = String(out.phoneNumber || '').trim();
  out.campaignId = String(out.campaignId || '').trim();

  out.employmentStatus = String(out.employmentStatus || '').trim();
  out.country = String(out.country || '').trim();

  if (out.insuranceEnrolled != null && out.insuranceSignedUp == null) {
    out.insuranceSignedUp = out.insuranceEnrolled;
  }
  if (out.insuranceEligibilityDate && !out.insuranceQualifiedDate) {
    out.insuranceQualifiedDate = out.insuranceEligibilityDate;
  }
  if (out.insuranceQualified != null && out.insuranceEligible == null) {
    out.insuranceEligible = out.insuranceQualified;
  }

  // Normalize date-only strings (keep after alias mapping)
  out.hireDate = out.hireDate ? _toIsoDateOnly_(out.hireDate) : '';
  out.terminationDate = out.terminationDate ? _toIsoDateOnly_(out.terminationDate) : '';
  out.probationEndDate = out.probationEndDate ? _toIsoDateOnly_(out.probationEndDate) : '';
  out.insuranceQualifiedDate = out.insuranceQualifiedDate ? _toIsoDateOnly_(out.insuranceQualifiedDate) : '';
  out.insuranceCardReceivedDate = out.insuranceCardReceivedDate ? _toIsoDateOnly_(out.insuranceCardReceivedDate) : '';

  // Normalize booleans
  out.isAdmin = _strToBool_(out.isAdmin);
  out.canLogin = _strToBool_(out.canLogin);
  out.insuranceEligible = (out.insuranceEligible == null) ? null : _strToBool_(out.insuranceEligible);
  out.insuranceSignedUp = _strToBool_(out.insuranceSignedUp);

  // Normalize numbers
  if (out.probationMonths === '' || out.probationMonths == null) {
    // leave as '' to avoid forcing a value
  } else {
    out.probationMonths = Number(out.probationMonths);
    if (isNaN(out.probationMonths)) out.probationMonths = '';
  }

  const pages = Array.isArray(out.pages) ? out.pages : _splitCsv_(out.pages);
  const roles = Array.isArray(out.roles) ? out.roles : _splitCsv_(out.roles);
  out.pages = pages; out.roles = roles;

  return out;
}

// Email helpers
function buildHtmlEmail_(title, preheader, innerHtml) {
  try {
    if (typeof renderEmail_ === 'function') {
      return renderEmail_({
        headerTitle: title,
        headerGradient: 'linear-gradient(135deg, #0ea5e9 0%, #0891b2 100%)',
        logoUrl: (typeof EMAIL_CONFIG !== 'undefined' && EMAIL_CONFIG.logoUrl) ? EMAIL_CONFIG.logoUrl : '',
        preheader: preheader,
        contentHtml: innerHtml
      });
    }
  } catch (_) { }
  return '<html><body style="font-family:Arial,sans-serif;"><h2>' + title + '</h2>' + innerHtml + '<hr><small>Automated notification</small></body></html>';
}

function safeSendHtmlEmail_(to, subject, html) {
  try {
    const prefix = (typeof EMAIL_CONFIG !== 'undefined' && EMAIL_CONFIG.subjectPrefix) ? EMAIL_CONFIG.subjectPrefix : '';
    const fromName = (typeof EMAIL_CONFIG !== 'undefined' && EMAIL_CONFIG.fromName) ? EMAIL_CONFIG.fromName : 'System';
    if (typeof sendEmail_ === 'function') {
      sendEmail_({ to, subject: prefix + subject, htmlBody: html });
      return;
    }
    MailApp.sendEmail({ to, subject: prefix + subject, htmlBody: html, body: 'This message contains HTML content.', name: fromName });
  } catch (e) { writeError && writeError('safeSendHtmlEmail_', e); }
}

/** Return the user row (object from Users sheet) for the direct manager of a user, if any. */
function getDirectManagerUser_(userId) {
  try {
    if (!userId) return null;
    const muSheet = getOrCreateManagerUsersSheet_();
    const data = muSheet.getDataRange().getValues();
    if (!data || data.length < 2) return null;
    const headers = data[0] || [];
    const midx = { ManagerUserID: headers.indexOf('ManagerUserID'), UserID: headers.indexOf('UserID') };
    if (midx.ManagerUserID === -1 || midx.UserID === -1) return null;
    let managerUserId = null;
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][midx.UserID]) === String(userId)) {
        managerUserId = data[r][midx.ManagerUserID];
        break;
      }
    }
    if (!managerUserId) return null;
    const users = readSheet(G.USERS_SHEET) || [];
    return users.find(u => String(u.ID) === String(managerUserId)) || null;
  } catch (e) { writeError && writeError('getDirectManagerUser_', e); return null; }
}

/** Pick ONE primary manager for a campaign from CampaignUserPermissions (PermissionLevel MANAGER). */
function getPrimaryCampaignManagerUser_(campaignId) {
  try {
    if (!campaignId) return null;
    const perms = readCampaignPermsSafely_() || [];
    // Only MANAGER level for the campaign
    const managers = perms.filter(p => String(p.CampaignID) === String(campaignId) && String(p.PermissionLevel).toUpperCase() === 'MANAGER');
    if (!managers.length) return null;
    // Choose the most recently updated (fallback to created)
    const toTime = v => (v ? new Date(v).getTime() : 0);
    managers.sort((a, b) => (toTime(b.UpdatedAt) || toTime(b.CreatedAt)) - (toTime(a.UpdatedAt) || toTime(a.CreatedAt)));
    const chosen = managers[0];
    const users = readSheet(G.USERS_SHEET) || [];
    return users.find(u => String(u.ID) === String(chosen.UserID)) || null;
  } catch (e) { writeError && writeError('getPrimaryCampaignManagerUser_', e); return null; }
}

/** Extract a sendable email from a user row, only if CanLogin is TRUE. */
function _emailIfActive_(u) {
  if (!u) return '';
  const canLogin = _strToBool_(u.CanLogin);
  const email = (u.Email || '').trim();
  return (canLogin && email) ? email : '';
}

function notifyOnUserRegistered_(newUser) {
  try {
    // Resolve a single recipient in this order:
    // 1) Direct manager assigned to this user (MANAGER_USERS mapping)
    // 2) Primary campaign manager for the user's campaign (CampaignUserPermissions MANAGER)
    // No more emailing every campaign manager or system admins.
    let recipient = '';

    // 1) Direct manager (if mapping already exists)
    const directMgr = getDirectManagerUser_(newUser && newUser.id);
    recipient = _emailIfActive_(directMgr);

    // 2) Fallback to primary campaign manager
    if (!recipient && newUser && newUser.campaignId) {
      const campMgr = getPrimaryCampaignManagerUser_(newUser.campaignId);
      recipient = _emailIfActive_(campMgr);
    }

    if (!recipient) return; // nothing to notify

    const rolesList = (newUser.roles || []).join(', ') || '—';
    const pagesList = (newUser.pages || []).join(', ') || '—';

    let campaignName = '';
    try { campaignName = getCampaignNameSafe ? getCampaignNameSafe(newUser.campaignId) : getCampaignName(newUser.campaignId); } catch (_) { }

    const html = buildHtmlEmail_(
      'New User Registered', 'A new user account was created.',
      '<div style="border:1px solid #e2e8f0;border-radius:12px;padding:16px;">'
      + '<p><strong>User Name:</strong> ' + (newUser.userName || '') + '</p>'
      + '<p><strong>Full Name:</strong> ' + (newUser.fullName || '') + '</p>'
      + '<p><strong>Email:</strong> ' + (newUser.email || '') + '</p>'
      + '<p><strong>Phone:</strong> ' + (newUser.phoneNumber || '') + '</p>'
      + '<p><strong>Campaign:</strong> ' + (campaignName || newUser.campaignId || '—') + '</p>'
      + '<p><strong>Roles:</strong> ' + rolesList + '</p>'
      + '<p><strong>Pages:</strong> ' + pagesList + '</p>'
      + '<p><strong>Can Login:</strong> ' + (newUser.canLogin ? 'Yes' : 'No') + '</p>'
      + '<p><strong>Is Admin:</strong> ' + (newUser.isAdmin ? 'Yes' : 'No') + '</p>'
      + '<p><strong>Created:</strong> ' + new Date(newUser.createdAt).toLocaleString() + '</p>'
      + '</div>'
    );


    safeSendHtmlEmail_(recipient, 'New User Registered: ' + (newUser.userName || newUser.email || newUser.id), html);
  } catch (e) { writeError && writeError('notifyOnUserRegistered_', e); }
}

function getEmailsForCampaignManagers_(campaignId) {
  try {
    if (!campaignId) return [];
    const perms = readCampaignPermsSafely_() || [];
    const managers = perms.filter(p => String(p.CampaignID) === String(campaignId) && ['MANAGER', 'ADMIN'].includes(String(p.PermissionLevel || '').toUpperCase()));
    if (!managers.length) return [];
    const users = readSheet(G.USERS_SHEET) || [];
    const emails = managers.map(m => { const u = users.find(x => String(x.ID) === String(m.UserID)); return u && u.Email; }).filter(Boolean);
    return Array.from(new Set(emails));
  } catch (e) { writeError && writeError('getEmailsForCampaignManagers_', e); return []; }
}
function getSystemAdminsEmails_() {
  try {
    const users = readSheet(G.USERS_SHEET) || [];
    const emails = users.filter(u => _strToBool_(u.IsAdmin) && _strToBool_(u.CanLogin)).map(u => u.Email).filter(Boolean);
    return Array.from(new Set(emails));
  } catch (e) { writeError && writeError('getSystemAdminsEmails_', e); return []; }
}

// ───────────────────────────────────────────────────────────────────────────────
// Public: Employment status helpers
// ───────────────────────────────────────────────────────────────────────────────
function clientGetValidEmploymentStatuses() {
  try { return { success: true, statuses: getValidEmploymentStatuses(), message: 'Valid employment statuses retrieved' }; }
  catch (e) { writeError && writeError('clientGetValidEmploymentStatuses', e); return { success: false, error: e.message, statuses: [] }; }
}
function clientGetEmploymentStatusReport(campaignId) {
  try {
    const users = readSheet(G.USERS_SHEET) || [];
    let filtered = users;
    if (campaignId) filtered = users.filter(u => u.CampaignID === campaignId);

    const statusCounts = {}; const valid = getValidEmploymentStatuses();
    valid.forEach(s => statusCounts[s] = 0); statusCounts['Unspecified'] = 0;

    filtered.forEach(u => {
      const s = String(u.EmploymentStatus || '').trim();
      if (s && valid.includes(s)) statusCounts[s]++; else statusCounts['Unspecified']++;
    });
    return { success: true, campaignId, totalUsers: filtered.length, statusCounts, message: `Employment status report for ${filtered.length} users` };
  } catch (e) { writeError && writeError('clientGetEmploymentStatusReport', e); return { success: false, error: e.message }; }
}

// ───────────────────────────────────────────────────────────────────────────────
// Benefits: snapshots + batch normalize/write
// ───────────────────────────────────────────────────────────────────────────────
function clientGetBenefitsSnapshot(userId) {
  try {
    const users = readSheet(G.USERS_SHEET) || [];
    const u = users.find(x => String(x.ID) === String(userId));
    if (!u) return { success: false, error: 'User not found' };
    const hire = u.HireDate || '';
    const probMonths = (u.ProbationMonths !== '' && u.ProbationMonths != null) ? Number(u.ProbationMonths) : '';
    const probEnd = u.ProbationEndDate || calcProbationEndDate_(hire, probMonths);
    const insQual = u.InsuranceQualifiedDate || calcInsuranceQualifiedDate_(probEnd, G.INSURANCE_MONTHS_AFTER_PROBATION);
    const eligible = isInsuranceEligibleNow_(insQual, u.TerminationDate || '');
    return {
      success: true,
      userId,
      hireDate: hire || '',
      probationMonths: probMonths === '' ? '' : Number(probMonths),
      probationEndDate: _toIsoDateOnly_(probEnd),
      insuranceQualifiedDate: _toIsoDateOnly_(insQual),
      insuranceEligible: !!eligible,
      insuranceSignedUp: _strToBool_(u.InsuranceSignedUp),
      insuranceCardReceivedDate: _toIsoDateOnly_(u.InsuranceCardReceivedDate || '')
    };
  } catch (e) { writeError && writeError('clientGetBenefitsSnapshot', e); return { success: false, error: e.message }; }
}

function clientBatchNormalizeBenefits() {
  try {
    const sh = _getSheet_(G.USERS_SHEET);
    let { headers, values, idx } = _scanSheet_(sh);
    _ensureUserHeaders_(idx);
    const ensured = ensureOptionalUserColumns_(sh, headers, idx);
    headers = ensured.headers; idx = ensured.idx;

    let updated = 0;
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const hire = row[idx['HireDate']] || '';
      const term = row[idx['TerminationDate']] || '';
      let probMonths = row[idx['ProbationMonths']];
      probMonths = (probMonths === '' || probMonths == null) ? '' : Number(probMonths);
      const curProbEnd = row[idx['ProbationEndDate']] || '';
      const curInsQual = row[idx['InsuranceQualifiedDate']] || '';
      const curEligible = row[idx['InsuranceEligible']];
      const curCard = row[idx['InsuranceCardReceivedDate']] || '';

      const probEnd = curProbEnd || calcProbationEndDate_(hire, probMonths);
      const insQual = curInsQual || calcInsuranceQualifiedDate_(probEnd, G.INSURANCE_MONTHS_AFTER_PROBATION);
      const eligible = isInsuranceEligibleNow_(insQual, term);

      const newProbEnd = _toIsoDateOnly_(probEnd);
      const newInsQual = _toIsoDateOnly_(insQual);
      const newEligibleStr = _boolToStr_(eligible);

      let changed = false;
      function setIfDiff(colName, val) {
        if (typeof idx[colName] !== 'number') return;
        const col = idx[colName] + 1;
        const cur = row[idx[colName]];
        if (String(cur) !== String(val)) { sh.getRange(r + 1, col).setValue(val); row[idx[colName]] = val; changed = true; }
      }
      setIfDiff('ProbationEndDate', newProbEnd);
      setIfDiff('InsuranceQualifiedDate', newInsQual);
      setIfDiff('InsuranceEligible', newEligibleStr);

      if (changed) {
        if (typeof idx['UpdatedAt'] === 'number') sh.getRange(r + 1, idx['UpdatedAt'] + 1).setValue(new Date());
        updated++;
      }
    }
    try { invalidateCache && invalidateCache(G.USERS_SHEET); } catch (_) { }
    return { success: true, updated, message: `Benefits normalized for ${updated} user(s)` };
  } catch (e) { writeError && writeError('clientBatchNormalizeBenefits', e); return { success: false, error: e.message }; }
}

// ───────────────────────────────────────────────────────────────────────────────
// Debug
// ───────────────────────────────────────────────────────────────────────────────
function testConnection() { return { success: true, timestamp: new Date().toISOString(), message: 'Connection successful' }; }

function debugUsersSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(G.USERS_SHEET);
    if (!sheet) {
      return { error: 'USERS_SHEET not found', sheetName: G.USERS_SHEET, availableSheets: ss.getSheets().map(s => s.getName()) };
    }
    const lastRow = sheet.getLastRow(); const lastCol = sheet.getLastColumn();
    if (lastRow < 2) {
      return {
        error: 'No data rows in sheet',
        dimensions: { rows: lastRow, cols: lastCol },
        sheetName: G.USERS_SHEET,
        headers: lastRow > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : []
      };
    }
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const sampleData = sheet.getRange(2, 1, Math.min(2, lastRow - 1), lastCol).getValues();
    return { success: true, sheetName: G.USERS_SHEET, dimensions: { rows: lastRow, cols: lastCol }, headers, sampleDataCount: sampleData.length, sampleData };
  } catch (e) { return { error: 'Error reading sheet: ' + e.message }; }
}

function _buildUniqIndexes_(users) {
  console.log('Building unique indexes for', users.length, 'users'); // Debug log
  
  const byEmail = new Map();
  const byUser = new Map();
  
  (users || []).forEach((u, index) => {
    try {
      const e = _normEmail_(u && (u.Email || u.email));
      // Fix: Check all possible username field variations
      const n = _normUser_(u && (u.UserName || u.userName || u.username));
      
      // Log problematic entries
      if (!e && u && (u.Email || u.email)) {
        console.warn('Failed to normalize email for user at index', index, ':', u.Email || u.email);
      }
      if (!n && u && (u.UserName || u.userName || u.username)) {
        console.warn('Failed to normalize username for user at index', index, ':', u.UserName || u.userName || u.username);
      }
      
      // Prefer first occurrence per key
      if (e && !byEmail.has(e)) byEmail.set(e, u);
      if (n && !byUser.has(n)) byUser.set(n, u);
      
    } catch (error) {
      console.error('Error processing user at index', index, ':', error, u);
    }
  });
  
  console.log('Built indexes:', { emailEntries: byEmail.size, userEntries: byUser.size }); // Debug log
  
  return { byEmail, byUser };
}

// ───────────────────────────────────────────────────────────────────────────────
// Optional: quick dry-run logger (does not send mail)
// ───────────────────────────────────────────────────────────────────────────────
function debugNotifyResolution_(userIdOrNewUser) {
  try {
    const users = readSheet(G.USERS_SHEET) || [];
    const newUser = (userIdOrNewUser && typeof userIdOrNewUser === 'object')
      ? userIdOrNewUser
      : (users.find(u => String(u.ID) === String(userIdOrNewUser)) || null);
    if (!newUser) return { success: false, error: 'User not found' };


    const directMgr = getDirectManagerUser_(newUser.id || newUser.ID);
    const directEmail = _emailIfActive_(directMgr);


    const campMgr = getPrimaryCampaignManagerUser_(newUser.campaignId || newUser.CampaignID);
    const campEmail = _emailIfActive_(campMgr);


    return { success: true, directMgr: directMgr && directMgr.Email, directEmail, campMgr: campMgr && campMgr.Email, campEmail };
  } catch (e) { return { success: false, error: String(e) }; }
}

console.log('✅ UserService.gs loaded');
console.log('📦 Features: User CRUD, Roles, Pages, Campaign perms, Manager mapping, HR/Benefits (probation + insurance)');
