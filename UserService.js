/******************************************************************************* 
 * UserService.gs — Complete Campaign-aware User Management + HR/Benefits
 * - FIXED globals: safe constant pattern (no block-scoped const in if)
 * - Adds HR/Benefits columns & logic:
 *     TerminationDate, ProbationMonths, ProbationEnd,
 *     InsuranceEligibleDate, InsuranceQualified, InsuranceEnrolled,
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
if (typeof G.USER_CAMPAIGNS_SHEET === 'undefined') {
  if (typeof USER_CAMPAIGNS_SHEET !== 'undefined') {
    G.USER_CAMPAIGNS_SHEET = USER_CAMPAIGNS_SHEET;
  } else {
    G.USER_CAMPAIGNS_SHEET = 'UserCampaigns';
  }
}
if (typeof G.MANAGER_USERS_HEADER === 'undefined') G.MANAGER_USERS_HEADER = ['ID', 'ManagerUserID', 'UserID', 'CreatedAt', 'UpdatedAt'];
if (typeof G.USER_EQUIPMENT_SHEET === 'undefined') G.USER_EQUIPMENT_SHEET = 'UserEquipment';
if (typeof G.USER_EQUIPMENT_HEADERS === 'undefined') {
  G.USER_EQUIPMENT_HEADERS = [
    'ID', 'UserID', 'ItemName', 'ItemType', 'SerialNumber', 'Condition', 'IssuedDate', 'ReturnedDate',
    'Notes', 'PhotoIds', 'PhotoUrls', 'PhotoNames', 'CreatedAt', 'UpdatedAt'
  ];
}
if (typeof G.USER_EQUIPMENT_FOLDER_NAME === 'undefined') G.USER_EQUIPMENT_FOLDER_NAME = 'LuminaHQ User Equipment';
if (typeof G.USER_EQUIPMENT_FOLDER_PROPERTY === 'undefined') G.USER_EQUIPMENT_FOLDER_PROPERTY = 'USER_EQUIPMENT_FOLDER_ID';

// Campaign user permissions headers
if (typeof G.CAMPAIGN_USER_PERMISSIONS_HEADERS === 'undefined') {
  G.CAMPAIGN_USER_PERMISSIONS_HEADERS = [
    'ID', 'CampaignID', 'UserID', 'PermissionLevel', 'CanManageUsers', 'CanManagePages', 'CreatedAt', 'UpdatedAt'
  ];
}

// HR/Benefits config
if (typeof G.INSURANCE_MONTHS_AFTER_PROBATION === 'undefined') G.INSURANCE_MONTHS_AFTER_PROBATION = 3;

// Canonical Users sheet column order (must remain in sync with sheet format)
var USER_SHEET_COLUMN_ORDER = (function resolveUserSheetColumnOrder_() {
  if (Array.isArray(G.USER_SHEET_COLUMN_ORDER) && G.USER_SHEET_COLUMN_ORDER.length) {
    return G.USER_SHEET_COLUMN_ORDER.slice();
  }
  const canonical = [
    'ID',
    'UserName',
    'FullName',
    'Email',
    'CampaignID',
    'PasswordHash',
    'ResetRequired',
    'EmailConfirmation',
    'EmailConfirmed',
    'PhoneNumber',
    'EmploymentStatus',
    'HireDate',
    'Country',
    'LockoutEnd',
    'TwoFactorEnabled',
    'CanLogin',
    'Roles',
    'Pages',
    'CreatedAt',
    'UpdatedAt',
    'IsAdmin',
    'DeletedAt',
    'IsAdmin',
    'ProbationMonths',
    'ProbationEnd',
    'InsuranceEligibleDate',
    'InsuranceQualified',
    'InsuranceEnrolled',
    'InsuranceCardReceivedDate',
    'ProbationEndDate',
    'InsuranceQualifiedDate',
    'InsuranceEligible',
    'InsuranceSignedUp',
    'TerminationDate',
    'MFASecret',
    'MFABackupCodes',
    'MFADeliveryPreference',
    'MFAEnabled',
    'NormalizedUserName',
    'NormalizedEmail',
    'PhoneNumberConfirmed',
    'LockoutEnabled',
    'AccessFailedCount',
    'TwoFactorDelivery',
    'TwoFactorSecret',
    'TwoFactorRecoveryCodes',
    'SecurityStamp',
    'ConcurrencyStamp',
    'EmailConfirmationTokenHash',
    'EmailConfirmationSentAt',
    'EmailConfirmationExpiresAt',
    'ResetPasswordToken',
    'ResetPasswordTokenHash',
    'ResetPasswordSentAt',
    'ResetPasswordExpiresAt',
    'LastLoginAt',
    'LastLoginIp',
    'LastLoginUserAgent'
  ];
  G.USER_SHEET_COLUMN_ORDER = canonical.slice();
  return canonical;
})();
if (!Array.isArray(G.USER_SHEET_COLUMN_ORDER) || G.USER_SHEET_COLUMN_ORDER.length !== USER_SHEET_COLUMN_ORDER.length ||
  G.USER_SHEET_COLUMN_ORDER.some((val, idx) => val !== USER_SHEET_COLUMN_ORDER[idx])) {
  G.USER_SHEET_COLUMN_ORDER = USER_SHEET_COLUMN_ORDER.slice();
}
if (!Array.isArray(G.USERS_HEADERS) || G.USERS_HEADERS.length !== USER_SHEET_COLUMN_ORDER.length ||
  G.USERS_HEADERS.some((val, idx) => val !== USER_SHEET_COLUMN_ORDER[idx])) {
  G.USERS_HEADERS = USER_SHEET_COLUMN_ORDER.slice();
}

// Optional extra columns we’ll ensure on Users sheet if missing
const OPTIONAL_USER_COLUMNS = USER_SHEET_COLUMN_ORDER.slice(21);

const USER_LOG_MAX_DEPTH = 4;
const USER_LOG_MAX_KEYS = 40;

function _userSanitizeForLog_(value, depth = 0, seen) {
  if (value === null || value === undefined) return value;
  const type = typeof value;
  if (type === 'string' || type === 'number' || type === 'boolean') return value;
  if (value instanceof Date) return value.toISOString();
  if (type === 'function') return '[Function]';

  const tracker = seen || new WeakSet();
  if (depth >= USER_LOG_MAX_DEPTH) {
    return Array.isArray(value) ? '[Array]' : '[Object]';
  }

  if (value && typeof value === 'object') {
    if (tracker.has(value)) return '[Circular]';
    tracker.add(value);
  }

  if (Array.isArray(value)) {
    return value.slice(0, USER_LOG_MAX_KEYS).map(item => _userSanitizeForLog_(item, depth + 1, tracker));
  }

  if (value && typeof value === 'object') {
    const out = {};
    Object.keys(value).slice(0, USER_LOG_MAX_KEYS).forEach(key => {
      try {
        out[key] = _userSanitizeForLog_(value[key], depth + 1, tracker);
      } catch (err) {
        out[key] = '[Unserializable]';
      }
    });
    return out;
  }

  try {
    return String(value);
  } catch (_) {
    return '[Unserializable primitive]';
  }
}

function _userLog_(label, payload, level) {
  try {
    const sanitized = _userSanitizeForLog_(payload);
    if (level === 'error') {
      console.error(label, sanitized);
    } else if (level === 'warn') {
      console.warn(label, sanitized);
    } else {
      console.log(label, sanitized);
    }
  } catch (err) {
    try {
      console.warn('UserService logging failed', label, err && err.message ? err.message : err);
    } catch (_) { }
  }
}

function _userGetPasswordUtilities_() {
  try {
    if (typeof ensurePasswordUtilities === 'function') {
      return ensurePasswordUtilities();
    }
  } catch (ensureError) {
    _userLog_('UserService.passwordUtilities.ensure', ensureError, 'warn');
  }

  try {
    if (typeof PasswordUtilities !== 'undefined' && PasswordUtilities) {
      return PasswordUtilities;
    }
  } catch (globalError) {
    _userLog_('UserService.passwordUtilities.global', globalError, 'warn');
  }

  return null;
}

function _userDigestToHex_(digest) {
  if (!digest) {
    return '';
  }

  try {
    const utils = _userGetPasswordUtilities_();
    if (!utils || typeof utils.digestToHex !== 'function') {
      throw new Error('Password utilities digestToHex unavailable');
    }

    return utils.digestToHex(digest);
  } catch (error) {
    _userLog_('UserService.digestToHex', error, 'error');
    throw error;
  }
}

function clientGetUserSummaries(context) {
  try {
    return getEntitySummaries('users', context);
  } catch (error) {
    console.error('clientGetUserSummaries failed:', error);
    throw error;
  }
}

function clientGetUserDetail(id, context) {
  try {
    const detail = getEntityDetail('users', id, context);
    try {
      if (detail && (detail.record || detail.id)) {
        const identifier = detail.record && detail.record.ID ? detail.record.ID : id;
        detail.hasActiveSession = userHasActiveSession(identifier);
      }
    } catch (sessionError) {
      console.warn('clientGetUserDetail: unable to resolve active session status', sessionError);
    }
    return detail;
  } catch (error) {
    console.error('clientGetUserDetail failed:', error);
    throw error;
  }
}

function userHasActiveSession(userIdentifier) {
  try {
    if (typeof AuthenticationService !== 'undefined'
      && AuthenticationService
      && typeof AuthenticationService.userHasActiveSession === 'function') {
      return AuthenticationService.userHasActiveSession(userIdentifier);
    }

    if (typeof IdentityService !== 'undefined'
      && IdentityService
      && typeof IdentityService.hasActiveSession === 'function') {
      return IdentityService.hasActiveSession(userIdentifier);
    }
  } catch (error) {
    console.warn('userHasActiveSession helper failed', error);
  }
  return false;
}

function _normalizeFieldKey_(key) {
  return String(key || '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function _findUserFieldValue_(user, targetKey) {
  if (!user || typeof user !== 'object') return '';
  const normalizedTarget = _normalizeFieldKey_(targetKey);
  if (!normalizedTarget) return '';

  try {
    if (user.sheetFieldMap && typeof user.sheetFieldMap === 'object') {
      for (const key in user.sheetFieldMap) {
        if (!Object.prototype.hasOwnProperty.call(user.sheetFieldMap, key)) continue;
        if (_normalizeFieldKey_(key) !== normalizedTarget) continue;
        const value = user.sheetFieldMap[key];
        if (value !== null && value !== undefined && value !== '') return value;
      }
    }
  } catch (_) { /* ignore lookup issues */ }

  if (Array.isArray(user.sheetFields)) {
    for (let i = 0; i < user.sheetFields.length; i++) {
      const field = user.sheetFields[i];
      if (!field || typeof field.key === 'undefined') continue;
      if (_normalizeFieldKey_(field.key) !== normalizedTarget) continue;
      const value = field.value;
      if (value !== null && value !== undefined && value !== '') return value;
    }
  }

  const keys = Object.keys(user);
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (_normalizeFieldKey_(key) !== normalizedTarget) continue;
    const value = user[key];
    if (value !== null && value !== undefined && value !== '') return value;
  }

  return '';
}

function _getUserName_(user) {
  if (!user) return '';
  const direct = user.UserName || user.userName || user.username || user.Username;
  if (direct !== null && direct !== undefined && String(direct).trim() !== '') return direct;
  return _findUserFieldValue_(user, 'username');
}

function _equipmentParseArray_(value) {
  if (value === null || typeof value === 'undefined') return [];
  if (Array.isArray(value)) return value.filter(v => v !== null && typeof v !== 'undefined').map(v => String(v));
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return [];
    try {
      const parsed = JSON.parse(trimmed);
      if (Array.isArray(parsed)) return parsed.filter(v => v !== null && typeof v !== 'undefined').map(v => String(v));
    } catch (_) {
      // ignore
    }
    return trimmed.split(/[\n,;]/).map(part => part.trim()).filter(Boolean);
  }
  return [];
}

function _equipmentStringifyArray_(arr) {
  return (Array.isArray(arr) && arr.length) ? JSON.stringify(arr) : '';
}

function _equipmentEnsureSheet_() {
  return ensureSheetWithHeaders(G.USER_EQUIPMENT_SHEET, G.USER_EQUIPMENT_HEADERS);
}

function _equipmentSanitizeSegment_(value, fallback) {
  const raw = String(value || '').trim();
  const sanitized = raw.replace(/[\\/:*?"<>|#]+/g, '_').replace(/\s+/g, ' ').trim();
  return sanitized || fallback || 'untitled';
}

function _equipmentEnsureRootFolder_() {
  if (typeof DriveApp === 'undefined') throw new Error('Drive service unavailable');
  let props = null;
  try { props = PropertiesService.getScriptProperties(); } catch (_) { props = null; }
  const propKey = G.USER_EQUIPMENT_FOLDER_PROPERTY || 'USER_EQUIPMENT_FOLDER_ID';
  let folderId = props ? props.getProperty(propKey) : null;
  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (err) {
      if (props) props.deleteProperty(propKey);
    }
  }
  const root = DriveApp.getRootFolder();
  const desiredName = G.USER_EQUIPMENT_FOLDER_NAME || 'User Equipment';
  const iter = root.getFoldersByName(desiredName);
  const folder = iter.hasNext() ? iter.next() : root.createFolder(desiredName);
  if (props) props.setProperty(propKey, folder.getId());
  return folder;
}

function _equipmentEnsureSubfolder_(userId, equipmentId) {
  const root = _equipmentEnsureRootFolder_();
  const userFolderName = _equipmentSanitizeSegment_('user-' + userId, 'user');
  let userFolderIter = root.getFoldersByName(userFolderName);
  let userFolder = userFolderIter.hasNext() ? userFolderIter.next() : root.createFolder(userFolderName);
  const equipmentFolderName = _equipmentSanitizeSegment_('equipment-' + equipmentId, 'equipment');
  let equipmentIter = userFolder.getFoldersByName(equipmentFolderName);
  let equipmentFolder = equipmentIter.hasNext() ? equipmentIter.next() : userFolder.createFolder(equipmentFolderName);
  return equipmentFolder;
}

function _equipmentToIsoString_(value) {
  if (!value && value !== 0) return '';
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? '' : value.toISOString();
  }
  if (typeof value === 'number') {
    const dt = new Date(value);
    return isNaN(dt.getTime()) ? '' : dt.toISOString();
  }
  const str = String(value || '').trim();
  if (!str) return '';
  const dateObj = new Date(str);
  if (!isNaN(dateObj.getTime())) return dateObj.toISOString();
  return str;
}

function _equipmentSavePhotos_(userId, equipmentId, photos) {
  const result = { ids: [], urls: [], names: [] };
  if (!Array.isArray(photos) || !photos.length) return result;
  if (typeof Utilities === 'undefined' || typeof Utilities.base64Decode !== 'function') {
    throw new Error('Utilities service unavailable for file upload');
  }
  const folder = _equipmentEnsureSubfolder_(userId, equipmentId);
  photos.forEach(photo => {
    try {
      if (!photo || !photo.dataUrl) return;
      const rawName = photo.name || ('Equipment Photo ' + (result.ids.length + 1));
      const safeName = _equipmentSanitizeSegment_(rawName, 'photo') + (rawName.includes('.') ? '' : '.png');
      let mimeType = photo.mimeType || '';
      let base64 = photo.dataUrl;
      if (typeof base64 === 'string' && base64.indexOf(',') > -1) {
        const parts = base64.split(',');
        const meta = parts.shift();
        base64 = parts.join(',');
        const match = meta && meta.match(/^data:(.*?);base64$/i);
        if (match && match[1]) mimeType = mimeType || match[1];
      }
      if (!mimeType) mimeType = 'image/png';
      const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, safeName);
      const file = folder.createFile(blob);
      result.ids.push(file.getId());
      result.urls.push(file.getUrl());
      result.names.push(safeName);
    } catch (err) {
      writeError && writeError('_equipmentSavePhotos_', err);
    }
  });
  return result;
}

function _equipmentRemovePhotos_(photoIds) {
  if (!Array.isArray(photoIds) || !photoIds.length) return;
  if (typeof DriveApp === 'undefined') return;
  photoIds.forEach(id => {
    const safeId = String(id || '').trim();
    if (!safeId) return;
    try {
      DriveApp.getFileById(safeId).setTrashed(true);
    } catch (err) {
      writeError && writeError('_equipmentRemovePhotos_', err);
    }
  });
}

function _equipmentMapRow_(headers, row) {
  const data = {};
  for (let i = 0; i < headers.length; i++) {
    data[headers[i]] = typeof row[i] !== 'undefined' ? row[i] : '';
  }
  const ids = _equipmentParseArray_(data.PhotoIds);
  const urls = _equipmentParseArray_(data.PhotoUrls);
  const names = _equipmentParseArray_(data.PhotoNames);
  const photos = [];
  const count = Math.max(ids.length, urls.length, names.length);
  for (let i = 0; i < count; i++) {
    photos.push({
      id: ids[i] || '',
      url: urls[i] || '',
      name: names[i] || ''
    });
  }
  return {
    id: data.ID || data.id || '',
    userId: data.UserID || data.userId || '',
    itemName: data.ItemName || data.itemName || '',
    itemType: data.ItemType || data.itemType || '',
    serialNumber: data.SerialNumber || data.serialNumber || '',
    condition: data.Condition || data.condition || '',
    issuedDate: _toIsoDateOnly_(data.IssuedDate || data.issuedDate || ''),
    returnedDate: _toIsoDateOnly_(data.ReturnedDate || data.returnedDate || ''),
    notes: data.Notes || data.notes || '',
    photoIds: ids,
    photoUrls: urls,
    photoNames: names,
    photos: photos,
    createdAt: _equipmentToIsoString_(data.CreatedAt || data.createdAt || ''),
    updatedAt: _equipmentToIsoString_(data.UpdatedAt || data.updatedAt || '')
  };
}

const EMPLOYMENT_STATUS_CANONICAL = [
  'Active',
  'Inactive',
  'Terminated',
  'On Leave',
  'Pending',
  'Probation',
  'Contract',
  'Contractor',
  'Full Time',
  'Part Time',
  'Seasonal',
  'Temporary',
  'Suspended',
  'Retired',
  'Intern',
  'Consultant'
];

const EMPLOYMENT_STATUS_ALIAS_MAP = {
  'active': 'Active',
  'inactive': 'Inactive',
  'terminated': 'Terminated',
  'leave': 'On Leave',
  'on leave': 'On Leave',
  'leave of absence': 'On Leave',
  'pending': 'Pending',
  'probation': 'Probation',
  'probationary': 'Probation',
  'probationary period': 'Probation',
  'contract': 'Contract',
  'contract employee': 'Contract',
  'contractor': 'Contractor',
  'consultant': 'Consultant',
  'full time': 'Full Time',
  'full-time': 'Full Time',
  'fulltime': 'Full Time',
  'part time': 'Part Time',
  'part-time': 'Part Time',
  'parttime': 'Part Time',
  'seasonal': 'Seasonal',
  'temporary': 'Temporary',
  'temp': 'Temporary',
  'suspended': 'Suspended',
  'retired': 'Retired',
  'intern': 'Intern',
  'consultant/contractor': 'Consultant'
};


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
    return EMPLOYMENT_STATUS_CANONICAL.slice();
  }
}
if (typeof normalizeEmploymentStatus !== 'function') {
  function normalizeEmploymentStatus(status) {
    const raw = String(status || '').trim();
    if (!raw) return '';
    const lower = raw.toLowerCase();
    if (Object.prototype.hasOwnProperty.call(EMPLOYMENT_STATUS_ALIAS_MAP, lower)) {
      return EMPLOYMENT_STATUS_ALIAS_MAP[lower];
    }
    const canonical = EMPLOYMENT_STATUS_CANONICAL.find(s => s.toLowerCase() === lower);
    return canonical || '';
  }
}
if (typeof validateEmploymentStatus !== 'function') {
  function validateEmploymentStatus(s) { return !s || normalizeEmploymentStatus(s) !== ''; }
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
function calcInsuranceEligibleDate_(probationEndStr, monthsAfter) {
  if (!probationEndStr) return '';
  const m = (typeof monthsAfter === 'number') ? monthsAfter : G.INSURANCE_MONTHS_AFTER_PROBATION;
  return _addMonths_(probationEndStr, m);
}
const calcInsuranceQualifiedDate_ = calcInsuranceEligibleDate_;
function isInsuranceQualifiedNow_(eligibleDateStr, terminationDateStr) {
  if (!eligibleDateStr) return false;
  const q = new Date(eligibleDateStr);
  if (isNaN(q)) return false;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  if (q > today) return false;
  if (terminationDateStr) {
    const t = new Date(terminationDateStr);
    if (!isNaN(t) && t <= today) return false;
  }
  return true;
}
const isInsuranceEligibleNow_ = isInsuranceQualifiedNow_;
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

function _normalizeUserIdHeaderKey_(header) {
  return String(header || '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function _isLikelyUserIdHeader_(normalizedKey, extraKeys) {
  if (!normalizedKey) return false;
  if (normalizedKey === 'id' || normalizedKey === 'userid') return true;
  if (normalizedKey === 'useridentifier' || normalizedKey === 'useruniqueid') return true;
  if (normalizedKey === 'useruuid' || normalizedKey === 'userguid' || normalizedKey === 'useridlegacy') return true;
  if (normalizedKey.indexOf('user') !== -1 && /id$/.test(normalizedKey)) return true;
  if (extraKeys && extraKeys[normalizedKey]) return true;
  return false;
}

function _normalizePotentialIdValue_(value) {
  if (value === null || typeof value === 'undefined') return '';
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return '';
    if (trimmed.toLowerCase() === 'null' || trimmed.toLowerCase() === 'undefined') return '';
    return trimmed;
  }
  if (typeof value === 'number') {
    if (!isFinite(value)) return '';
    return String(value);
  }
  if (value instanceof Date) {
    return String(value.getTime());
  }
  try {
    return String(value);
  } catch (_) {
    return '';
  }
}

function _generateUniqueUserId_(usedIds) {
  let attempts = 0;
  while (attempts < 20) {
    let candidate = '';
    if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.getUuid === 'function') {
      candidate = Utilities.getUuid();
    } else {
      const rand = Math.random().toString(36).slice(2, 10);
      candidate = 'USR-' + Date.now().toString(36) + '-' + rand;
    }
    candidate = _normalizePotentialIdValue_(candidate);
    if (candidate && !usedIds[candidate]) {
      return candidate;
    }
    attempts += 1;
  }
  const fallback = 'USR-' + (new Date().getTime()) + '-' + Math.floor(Math.random() * 1000000);
  return _normalizePotentialIdValue_(fallback) || fallback;
}

function _collectUserIdCandidatesFromStructure_(value, addCandidate, seenObjects, extraKeys) {
  if (value === null || typeof value === 'undefined') return;
  const valueType = typeof value;
  if (valueType === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return;
    if (trimmed[0] === '{' || trimmed[0] === '[') {
      try {
        const parsed = JSON.parse(trimmed);
        _collectUserIdCandidatesFromStructure_(parsed, addCandidate, seenObjects, extraKeys);
      } catch (_) { /* ignore malformed JSON */ }
    }
    return;
  }
  if (valueType === 'number' || valueType === 'boolean') return;
  if (valueType !== 'object') return;

  if (typeof WeakSet !== 'undefined') {
    if (!seenObjects) {
      seenObjects = new WeakSet();
    } else if (seenObjects.has(value)) {
      return;
    }
    seenObjects.add(value);
  }

  if (Array.isArray(value)) {
    for (let i = 0; i < value.length; i++) {
      _collectUserIdCandidatesFromStructure_(value[i], addCandidate, seenObjects, extraKeys);
    }
    return;
  }

  for (const key in value) {
    if (!Object.prototype.hasOwnProperty.call(value, key)) continue;
    const normalizedKey = _normalizeUserIdHeaderKey_(key);
    if (_isLikelyUserIdHeader_(normalizedKey, extraKeys)) {
      const candidate = _normalizePotentialIdValue_(value[key]);
      if (candidate) addCandidate(candidate);
    }
    _collectUserIdCandidatesFromStructure_(value[key], addCandidate, seenObjects, extraKeys);
  }
}

function _collectRowUserIdCandidates_(row, headerMeta, options) {
  const candidates = [];
  const seen = {};
  const addCandidate = function (candidate) {
    if (!candidate || seen[candidate]) return;
    seen[candidate] = true;
    candidates.push(candidate);
  };

  const candidateIndexes = headerMeta && headerMeta.candidateIndexes ? headerMeta.candidateIndexes : [];
  for (let i = 0; i < candidateIndexes.length; i++) {
    const idx = candidateIndexes[i];
    if (idx < 0 || idx >= row.length) continue;
    const candidate = _normalizePotentialIdValue_(row[idx]);
    if (candidate) addCandidate(candidate);
  }

  const allowNestedScan = !options || options.scanNested !== false;
  if (allowNestedScan) {
    for (let i = 0; i < row.length; i++) {
      const cell = row[i];
      if (cell === null || typeof cell === 'undefined') continue;
      if (typeof cell === 'string') {
        const trimmed = cell.trim();
        if (!trimmed) continue;
        if (trimmed[0] !== '{' && trimmed[0] !== '[') continue;
      } else if (typeof cell !== 'object') {
        continue;
      }
      try {
        _collectUserIdCandidatesFromStructure_(cell, addCandidate, null, headerMeta.extraCandidateKeys);
      } catch (_) { /* ignore traversal issues */ }
    }
  }

  return candidates;
}

function _canonicalUserId_(user) {
  if (!user || typeof user !== 'object') return '';
  const candidates = [user.ID, user.Id, user.id, user.UserID, user.userId];
  for (let i = 0; i < candidates.length; i++) {
    const normalized = _normalizePotentialIdValue_(candidates[i]);
    if (normalized) return normalized;
  }
  return '';
}

function _buildUserIdentifierLookup_() {
  const users = _readUsersAsObjects_();
  const lookup = {
    users,
    byId: {},
    byUser: {},
    byEmail: {}
  };

  for (let i = 0; i < users.length; i++) {
    const user = users[i];
    if (!user || typeof user !== 'object') continue;

    const canonicalId = _canonicalUserId_(user);
    if (canonicalId && !lookup.byId[canonicalId]) {
      lookup.byId[canonicalId] = user;
    }

    const userKey = _normUser_(_getUserName_(user));
    if (userKey && !lookup.byUser[userKey]) {
      lookup.byUser[userKey] = user;
    }

    const emailKey = _normEmail_(user.Email || user.email);
    if (emailKey && !lookup.byEmail[emailKey]) {
      lookup.byEmail[emailKey] = user;
    }
  }

  return lookup;
}

function ensureUsersHaveIds(options) {
  const summary = {
    success: true,
    total: 0,
    updated: 0,
    generated: 0,
    reused: 0,
    duplicatesResolved: 0,
    details: []
  };

  if (typeof SpreadsheetApp === 'undefined') {
    summary.success = false;
    summary.error = 'SpreadsheetApp is not available in this context.';
    return summary;
  }

  let sheet;
  try {
    sheet = _getSheet_(G.USERS_SHEET);
  } catch (sheetError) {
    summary.success = false;
    summary.error = sheetError && sheetError.message ? sheetError.message : String(sheetError);
    return summary;
  }

  let scan;
  try {
    scan = _scanSheet_(sheet);
  } catch (scanError) {
    summary.success = false;
    summary.error = scanError && scanError.message ? scanError.message : String(scanError);
    return summary;
  }

  const headers = scan.headers || [];
  const values = scan.values || [];
  if (values.length <= 1) {
    return summary;
  }

  const normalizedHeaderKeys = headers.map(_normalizeUserIdHeaderKey_);
  const headerMeta = {
    normalizedHeaderKeys,
    candidateIndexes: [],
    extraCandidateKeys: {}
  };

  if (options && Array.isArray(options.extraCandidateHeaders)) {
    options.extraCandidateHeaders.forEach(header => {
      const key = _normalizeUserIdHeaderKey_(header);
      if (key) headerMeta.extraCandidateKeys[key] = true;
    });
  }

  let idColumnIndex = -1;
  for (let i = 0; i < normalizedHeaderKeys.length; i++) {
    const key = normalizedHeaderKeys[i];
    if (key === 'id') {
      idColumnIndex = i;
      break;
    }
  }
  if (idColumnIndex === -1 && typeof scan.idx.ID === 'number') {
    idColumnIndex = scan.idx.ID;
  }
  if (idColumnIndex === -1) {
    summary.success = false;
    summary.error = 'Users sheet is missing an ID column.';
    return summary;
  }

  for (let i = 0; i < normalizedHeaderKeys.length; i++) {
    if (i === idColumnIndex) continue;
    const key = normalizedHeaderKeys[i];
    if (_isLikelyUserIdHeader_(key, headerMeta.extraCandidateKeys)) {
      headerMeta.candidateIndexes.push(i);
    }
  }

  const rows = values.slice(1);
  summary.total = rows.length;

  const idUsageCounts = {};
  const originalIds = new Array(rows.length);
  for (let r = 0; r < rows.length; r++) {
    const original = _normalizePotentialIdValue_(rows[r][idColumnIndex]);
    originalIds[r] = original;
    if (!original) continue;
    idUsageCounts[original] = (idUsageCounts[original] || 0) + 1;
  }

  const duplicateSeenCounts = {};
  const usedIds = {};
  const idColumnValues = new Array(rows.length);
  const columnUpdates = {};
  const columnDirtyFlags = {};
  headerMeta.candidateIndexes.forEach(index => {
    columnUpdates[index] = rows.map(row => row[index]);
    columnDirtyFlags[index] = false;
  });

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    const originalId = originalIds[r];
    let finalId = originalId;
    let action = 'kept';

    if (finalId) {
      if (idUsageCounts[finalId] > 1) {
        duplicateSeenCounts[finalId] = (duplicateSeenCounts[finalId] || 0) + 1;
        if (duplicateSeenCounts[finalId] === 1 && !usedIds[finalId]) {
          usedIds[finalId] = true;
          action = 'kept';
        } else {
          finalId = '';
        }
      } else if (usedIds[finalId]) {
        finalId = '';
      } else {
        usedIds[finalId] = true;
      }
    }

    if (!finalId) {
      const candidates = _collectRowUserIdCandidates_(row, headerMeta, options || {});
      for (let i = 0; i < candidates.length; i++) {
        const candidate = candidates[i];
        if (!usedIds[candidate]) {
          finalId = candidate;
          usedIds[finalId] = true;
          summary.reused += 1;
          action = 'reused';
          break;
        }
      }
    }

    if (!finalId) {
      finalId = _generateUniqueUserId_(usedIds);
      usedIds[finalId] = true;
      summary.generated += 1;
      action = 'generated';
    }

    if (originalId !== finalId) {
      summary.updated += 1;
      if (originalId && idUsageCounts[originalId] > 1) {
        summary.duplicatesResolved += 1;
      }
      row[idColumnIndex] = finalId;
    }

    idColumnValues[r] = [finalId];
    summary.details.push({
      rowNumber: r + 2,
      originalId: originalId,
      finalId: finalId,
      action: action
    });

    for (let c = 0; c < headerMeta.candidateIndexes.length; c++) {
      const idx = headerMeta.candidateIndexes[c];
      const current = _normalizePotentialIdValue_(row[idx]);
      if (!current || current === originalId || current === finalId) {
        if (_normalizePotentialIdValue_(columnUpdates[idx][r]) !== finalId) {
          columnUpdates[idx][r] = finalId;
          row[idx] = finalId;
          columnDirtyFlags[idx] = true;
        }
      }
    }
  }

  try {
    if (summary.updated > 0) {
      sheet.getRange(2, idColumnIndex + 1, rows.length, 1).setValues(idColumnValues);
    }
    Object.keys(columnUpdates).forEach(indexKey => {
      if (!columnDirtyFlags[indexKey]) return;
      const columnIndex = Number(indexKey);
      const columnValues = columnUpdates[columnIndex].map(value => [value]);
      sheet.getRange(2, columnIndex + 1, rows.length, 1).setValues(columnValues);
    });
  } catch (writeError) {
    summary.success = false;
    summary.error = writeError && writeError.message ? writeError.message : String(writeError);
  }

  return summary;
}

function reconcileUserIdReferencesAcrossSheets(options) {
  const summary = {
    success: true,
    metrics: {
      totalSheets: 0,
      processedSheets: 0,
      sheetsWithUpdates: 0,
      totalRows: 0,
      totalBlankUserIdRows: 0,
      totalUpdatedRows: 0,
      totalResolved: 0,
      totalUnresolved: 0
    },
    sheets: [],
    warnings: [],
    errors: []
  };

  const sheetExistenceCache = Object.create(null);
  function sheetExists(name) {
    const trimmed = String(name || '').trim();
    if (!trimmed) return false;
    const cacheKey = trimmed.toLowerCase();
    if (Object.prototype.hasOwnProperty.call(sheetExistenceCache, cacheKey)) {
      return sheetExistenceCache[cacheKey];
    }

    let environmentChecked = false;

    function markResult(result) {
      sheetExistenceCache[cacheKey] = result;
      return result;
    }

    // Check the main spreadsheet first (legacy behaviour)
    if (typeof SpreadsheetApp === 'undefined' || !SpreadsheetApp || typeof SpreadsheetApp.getActiveSpreadsheet !== 'function') {
      return markResult(true);
    }

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss && typeof ss.getSheetByName === 'function') {
        environmentChecked = true;
        const existsOnMain = !!ss.getSheetByName(trimmed);
        if (existsOnMain) {
          return markResult(true);
        }
      }
    } catch (existsError) {
      const message = existsError && existsError.message ? existsError.message : String(existsError);
      summary.warnings.push({ sheet: trimmed, stage: 'existsCheck', error: message });
      _userLog_('UserService.reconcileUserIdReferences.sheetExistsCheckFailed', { sheet: trimmed, error: message }, 'warn');
      return markResult(true);
    }

    // If not found on the main spreadsheet, check the dedicated schedule spreadsheet
    if (typeof getScheduleSpreadsheet === 'function') {
      try {
        const scheduleSs = getScheduleSpreadsheet();
        if (scheduleSs && typeof scheduleSs.getSheetByName === 'function') {
          environmentChecked = true;
          const existsOnSchedule = !!scheduleSs.getSheetByName(trimmed);
          if (existsOnSchedule) {
            return markResult(true);
          }
        }
      } catch (scheduleError) {
        const message = scheduleError && scheduleError.message ? scheduleError.message : String(scheduleError);
        summary.warnings.push({ sheet: trimmed, stage: 'scheduleExistsCheck', error: message });
        _userLog_('UserService.reconcileUserIdReferences.scheduleExistsCheckFailed', { sheet: trimmed, error: message }, 'warn');
        return markResult(true);
      }
    }

    // If no environment check succeeded, default to true to avoid false negatives during dry runs
    if (!environmentChecked) {
      return markResult(true);
    }

    return markResult(false);
  }

  let lookup;
  try {
    lookup = _buildUserIdentifierLookup_();
  } catch (lookupError) {
    summary.success = false;
    summary.error = lookupError && lookupError.message ? lookupError.message : String(lookupError);
    _userLog_('UserService.reconcileUserIdReferences.lookupFailed', { error: summary.error }, 'error');
    return summary;
  }

  const requestedSheets = Array.isArray(options && options.sheets)
    ? options.sheets.map(name => String(name || '').trim().toLowerCase()).filter(Boolean)
    : null;

  const baseAltIdHeaders = [
    'UserID', 'UserId', 'User UUID', 'User GUID', 'EmployeeID', 'EmployeeId',
    'AgentID', 'AgentId', 'AgentUserID', 'AgentUserId', 'ProfileID', 'ProfileId',
    'UserProfileID', 'UserProfileId', 'RequestorUserID', 'RequestorUserId',
    'TargetUserID', 'TargetUserId', 'AssigneeID', 'AssigneeId',
    'ManagedUserID', 'ManagedUserId'
  ];
  const baseUsernameHeaders = [
    'UserName', 'User', 'FullName', 'DisplayName', 'UserFullName',
    'AgentName', 'AgentUserName', 'ManagedUserName', 'RequestorUserName',
    'TargetUserName', 'AssignedUser', 'AssignedTo', 'EmployeeName',
    'OwnerName', 'ContactName'
  ];
  const baseEmailHeaders = [
    'Email', 'UserEmail', 'LoginEmail', 'AgentEmail', 'EmployeeEmail',
    'ContactEmail', 'OwnerEmail', 'ManagedUserEmail'
  ];

  function toNormalizedSet(base, extras) {
    const set = Object.create(null);
    (base || []).forEach(value => {
      const key = _normalizeUserIdHeaderKey_(value);
      if (key) set[key] = true;
    });
    (extras || []).forEach(value => {
      const key = _normalizeUserIdHeaderKey_(value);
      if (key) set[key] = true;
    });
    return set;
  }

  function resolveSheetConfigs() {
    const attendanceSheetName = (typeof G.ATTENDANCE_LOG_SHEET === 'string' && G.ATTENDANCE_LOG_SHEET) ||
      (typeof G.ATTENDANCE_SHEET === 'string' && G.ATTENDANCE_SHEET) ||
      'AttendanceLog';
    const seen = Object.create(null);
    const out = [];
    function pushConfig(sheetName, cfg) {
      const normalizedName = String(sheetName || '').trim();
      if (!normalizedName) return;
      const key = normalizedName.toLowerCase();
      if (!key || seen[key]) return;
      if (requestedSheets && requestedSheets.length && requestedSheets.indexOf(key) === -1) return;
      if (!sheetExists(normalizedName)) return;
      seen[key] = true;
      out.push({
        name: normalizedName,
        label: cfg.label || normalizedName,
        idColumnKeys: cfg.idColumnKeys || ['UserID', 'UserId'],
        altIdKeys: cfg.altIdKeys || [],
        usernameKeys: cfg.usernameKeys || [],
        emailKeys: cfg.emailKeys || []
      });
    }

    function addSheetCandidates(candidates, cfg) {
      const list = Array.isArray(candidates) ? candidates : [candidates];
      list.forEach(candidate => {
        let value = candidate;
        if (typeof value === 'function') {
          try {
            value = value();
          } catch (candidateError) {
            value = null;
          }
        }
        if (value === null || typeof value === 'undefined') return;
        pushConfig(value, cfg);
      });
    }

    addSheetCandidates([G.USER_ROLES_SHEET, (typeof USER_ROLES_SHEET !== 'undefined') ? USER_ROLES_SHEET : null, 'UserRoles'], {
      label: 'User roles',
      idColumnKeys: ['UserID', 'UserId'],
      usernameKeys: ['UserName'],
      emailKeys: ['UserEmail', 'Email']
    });

    addSheetCandidates([
      G.CAMPAIGN_USER_PERMISSIONS_SHEET,
      (typeof CAMPAIGN_USER_PERMISSIONS_SHEET !== 'undefined') ? CAMPAIGN_USER_PERMISSIONS_SHEET : null,
      'CampaignUserPermissions'
    ], {
      label: 'Campaign user permissions',
      idColumnKeys: ['UserID', 'UserId'],
      usernameKeys: ['UserName', 'User'],
      emailKeys: ['UserEmail', 'Email']
    });

    addSheetCandidates([
      G.USER_EQUIPMENT_SHEET,
      (typeof USER_EQUIPMENT_SHEET !== 'undefined') ? USER_EQUIPMENT_SHEET : null,
      'UserEquipment'
    ], {
      label: 'User equipment',
      idColumnKeys: ['UserID', 'UserId'],
      usernameKeys: ['UserName', 'AssignedTo', 'AssignedUser', 'EmployeeName', 'OwnerName'],
      emailKeys: ['UserEmail', 'Email', 'OwnerEmail', 'EmployeeEmail']
    });

    addSheetCandidates([
      G.MANAGER_USERS_SHEET,
      (typeof USER_MANAGERS_SHEET !== 'undefined') ? USER_MANAGERS_SHEET : null,
      'MANAGER_USERS',
      'ManagerUsers',
      'UserManagers'
    ], {
      label: 'Manager user mappings',
      idColumnKeys: ['UserID', 'UserId'],
      altIdKeys: ['ManagedUserID', 'ManagedUserId'],
      usernameKeys: ['ManagedUserName', 'UserName', 'User'],
      emailKeys: ['ManagedUserEmail', 'UserEmail', 'Email']
    });

    addSheetCandidates([
      attendanceSheetName,
      (typeof ATTENDANCE_LOG_SHEET !== 'undefined') ? ATTENDANCE_LOG_SHEET : null,
      (typeof ATTENDANCE_SHEET !== 'undefined') ? ATTENDANCE_SHEET : null
    ], {
      label: 'Attendance log',
      idColumnKeys: ['UserID', 'UserId'],
      usernameKeys: ['User', 'UserName', 'Agent', 'AgentName'],
      emailKeys: ['UserEmail', 'Email', 'AgentEmail']
    });

    addSheetCandidates([
      G.USER_CAMPAIGNS_SHEET,
      (typeof USER_CAMPAIGNS_SHEET !== 'undefined') ? USER_CAMPAIGNS_SHEET : null,
      'UserCampaigns'
    ], {
      label: 'User campaign assignments',
      idColumnKeys: ['UserID', 'UserId'],
      altIdKeys: [],
      usernameKeys: ['UserName', 'AgentName', 'FullName'],
      emailKeys: ['UserEmail', 'Email']
    });

    addSheetCandidates([
      (typeof SCHEDULE_GENERATION_SHEET !== 'undefined') ? SCHEDULE_GENERATION_SHEET : null,
      'GeneratedSchedules'
    ], {
      label: 'Generated schedules',
      idColumnKeys: ['UserID', 'UserId'],
      usernameKeys: ['UserName', 'AgentName', 'AssignedUser'],
      emailKeys: ['UserEmail', 'Email', 'AgentEmail']
    });

    addSheetCandidates([
      (typeof SCHEDULE_NOTIFICATIONS_SHEET !== 'undefined') ? SCHEDULE_NOTIFICATIONS_SHEET : null,
      'ScheduleNotifications'
    ], {
      label: 'Schedule notifications',
      idColumnKeys: ['UserID', 'UserId'],
      usernameKeys: ['UserName', 'AssignedUser', 'RecipientName'],
      emailKeys: ['UserEmail', 'Email', 'RecipientEmail']
    });

    addSheetCandidates([
      (typeof SCHEDULE_ADHERENCE_SHEET !== 'undefined') ? SCHEDULE_ADHERENCE_SHEET : null,
      'ScheduleAdherence'
    ], {
      label: 'Schedule adherence',
      idColumnKeys: ['UserID', 'UserId'],
      usernameKeys: ['UserName', 'AgentName'],
      emailKeys: ['UserEmail', 'Email']
    });

    addSheetCandidates([
      (typeof RECURRING_SCHEDULES_SHEET !== 'undefined') ? RECURRING_SCHEDULES_SHEET : null,
      'RecurringSchedules'
    ], {
      label: 'Recurring schedules',
      idColumnKeys: ['UserID', 'UserId'],
      usernameKeys: ['UserName', 'AgentName'],
      emailKeys: ['UserEmail', 'Email']
    });

    addSheetCandidates([
      (typeof ATTENDANCE_STATUS_SHEET !== 'undefined') ? ATTENDANCE_STATUS_SHEET : null,
      'AttendanceStatus'
    ], {
      label: 'Attendance status',
      idColumnKeys: ['UserID', 'UserId'],
      usernameKeys: ['UserName', 'AgentName'],
      emailKeys: ['UserEmail', 'Email']
    });

    addSheetCandidates([
      (typeof USER_HOLIDAY_PAY_STATUS_SHEET !== 'undefined') ? USER_HOLIDAY_PAY_STATUS_SHEET : null,
      'UserHolidayPayStatus'
    ], {
      label: 'Holiday pay status',
      idColumnKeys: ['UserID', 'UserId'],
      usernameKeys: ['UserName', 'EmployeeName'],
      emailKeys: ['UserEmail', 'Email', 'EmployeeEmail']
    });

    addSheetCandidates('UserBookmarks', {
      label: 'User bookmarks',
      idColumnKeys: ['UserID', 'UserId'],
      emailKeys: ['UserEmail', 'Email']
    });

    addSheetCandidates('BrowsingAnalytics', {
      label: 'Browsing analytics',
      idColumnKeys: ['UserID', 'UserId'],
      emailKeys: ['UserEmail', 'Email']
    });

    addSheetCandidates('SecurityIncidents', {
      label: 'Security incidents',
      idColumnKeys: ['UserID', 'UserId'],
      emailKeys: ['UserEmail', 'Email']
    });

    addSheetCandidates('ComplianceAuditTrail', {
      label: 'Compliance audit trail',
      idColumnKeys: ['UserID', 'UserId'],
      emailKeys: ['UserEmail', 'Email']
    });

    return out;
  }

  const configs = resolveSheetConfigs();
  summary.metrics.totalSheets = configs.length;

  if (!configs.length) {
    summary.warnings.push({ message: 'No sheets resolved for reconciliation', requestedSheets: requestedSheets || [] });
    _userLog_('UserService.reconcileUserIdReferences.noSheets', { requestedSheets: requestedSheets || [] }, 'warn');
    return summary;
  }

  configs.forEach(config => {
    const sheetName = config.name;
    summary.metrics.processedSheets += 1;

    let sheet;
    try {
      sheet = _getSheet_(sheetName);
    } catch (sheetError) {
      const message = sheetError && sheetError.message ? sheetError.message : String(sheetError);
      summary.success = false;
      summary.errors.push({ sheet: sheetName, stage: 'getSheet', error: message });
      _userLog_('UserService.reconcileUserIdReferences.sheetMissing', { sheet: sheetName, error: message }, 'warn');
      return;
    }

    let scan;
    try {
      scan = _scanSheet_(sheet);
    } catch (scanError) {
      const message = scanError && scanError.message ? scanError.message : String(scanError);
      summary.success = false;
      summary.errors.push({ sheet: sheetName, stage: 'scanSheet', error: message });
      _userLog_('UserService.reconcileUserIdReferences.scanFailed', { sheet: sheetName, error: message }, 'warn');
      return;
    }

    const headers = scan.headers || [];
    const values = scan.values || [];
    if (values.length <= 1) {
      const sheetSummary = {
        sheetName,
        label: config.label,
        totalRows: 0,
        blankUserIdRows: 0,
        resolvedRows: 0,
        unresolvedRows: 0,
        identifierUsage: { id: 0, username: 0, email: 0 },
        skipped: 'noDataRows'
      };
      summary.sheets.push(sheetSummary);
      _userLog_('UserService.reconcileUserIdReferences.sheetSummary', sheetSummary);
      return;
    }

    const normalizedHeaders = headers.map(_normalizeUserIdHeaderKey_);
    const idColumnKeySet = toNormalizedSet(config.idColumnKeys || ['UserID', 'UserId']);
    const idColumnIndexes = [];
    const idColumnIndexSet = Object.create(null);
    for (let i = 0; i < normalizedHeaders.length; i++) {
      if (idColumnKeySet[normalizedHeaders[i]]) {
        idColumnIndexes.push(i);
        idColumnIndexSet[i] = true;
      }
    }

    if (!idColumnIndexes.length) {
      const sheetSummary = {
        sheetName,
        label: config.label,
        totalRows: values.length - 1,
        blankUserIdRows: 0,
        resolvedRows: 0,
        unresolvedRows: 0,
        identifierUsage: { id: 0, username: 0, email: 0 },
        skipped: 'missingUserIdColumn'
      };
      summary.warnings.push({ sheet: sheetName, message: 'UserID column not found' });
      summary.sheets.push(sheetSummary);
      _userLog_('UserService.reconcileUserIdReferences.sheetSummary', sheetSummary, 'warn');
      return;
    }

    const altIdKeys = toNormalizedSet(baseAltIdHeaders, config.altIdKeys);
    const usernameKeys = toNormalizedSet(baseUsernameHeaders, config.usernameKeys);
    const emailKeys = toNormalizedSet(baseEmailHeaders, config.emailKeys);

    const altIndexes = { id: [], username: [], email: [] };
    const seenIndexes = Object.create(null);
    for (let i = 0; i < normalizedHeaders.length; i++) {
      if (idColumnIndexSet[i]) continue;
      const key = normalizedHeaders[i];
      if (!key) continue;
      if (altIdKeys[key]) {
        if (!seenIndexes['id:' + i]) {
          altIndexes.id.push(i);
          seenIndexes['id:' + i] = true;
        }
      }
      if (usernameKeys[key]) {
        if (!seenIndexes['username:' + i]) {
          altIndexes.username.push(i);
          seenIndexes['username:' + i] = true;
        }
      }
      if (emailKeys[key]) {
        if (!seenIndexes['email:' + i]) {
          altIndexes.email.push(i);
          seenIndexes['email:' + i] = true;
        }
      }
    }

    const rows = values.slice(1);
    const sheetSummary = {
      sheetName,
      label: config.label,
      totalRows: rows.length,
      blankUserIdRows: 0,
      resolvedRows: 0,
      unresolvedRows: 0,
      identifierUsage: { id: 0, username: 0, email: 0 },
      sampleResolved: [],
      columnSummaries: []
    };
    const columnWriteQueue = [];

    for (let c = 0; c < idColumnIndexes.length; c++) {
      const columnIndex = idColumnIndexes[c];
      const columnHeader = headers[columnIndex] || '';
      const columnValues = new Array(rows.length);
      const columnSummary = {
        columnIndex,
        header: columnHeader,
        blankUserIdRows: 0,
        resolvedRows: 0,
        unresolvedRows: 0,
        identifierUsage: { id: 0, username: 0, email: 0 }
      };
      let columnHadChanges = false;

      for (let r = 0; r < rows.length; r++) {
        const row = rows[r] || [];
        const currentValue = row[columnIndex];
        const normalizedCurrent = _normalizePotentialIdValue_(currentValue);
        let finalValue = (currentValue === null || typeof currentValue === 'undefined') ? '' : currentValue;
        let resolvedThisCell = false;

        if (!normalizedCurrent) {
          columnSummary.blankUserIdRows += 1;

          let matchedUser = null;
          let matchedVia = '';

          for (let i = 0; i < altIndexes.id.length && !matchedUser; i++) {
            const candidate = _normalizePotentialIdValue_(row[altIndexes.id[i]]);
            if (!candidate) continue;
            matchedUser = lookup.byId[candidate];
            if (matchedUser) {
              matchedVia = 'id';
              break;
            }
          }

          if (!matchedUser) {
            for (let i = 0; i < altIndexes.username.length && !matchedUser; i++) {
              const candidate = _normUser_(row[altIndexes.username[i]]);
              if (!candidate) continue;
              matchedUser = lookup.byUser[candidate];
              if (matchedUser) {
                matchedVia = 'username';
                break;
              }
            }
          }

          if (!matchedUser) {
            for (let i = 0; i < altIndexes.email.length && !matchedUser; i++) {
              const candidate = _normEmail_(row[altIndexes.email[i]]);
              if (!candidate) continue;
              matchedUser = lookup.byEmail[candidate];
              if (matchedUser) {
                matchedVia = 'email';
                break;
              }
            }
          }

          if (matchedUser) {
            const canonicalId = _canonicalUserId_(matchedUser);
            if (canonicalId) {
              finalValue = canonicalId;
              resolvedThisCell = true;
              columnSummary.resolvedRows += 1;
              const usageKey = matchedVia || 'id';
              if (columnSummary.identifierUsage[usageKey] !== undefined) {
                columnSummary.identifierUsage[usageKey] += 1;
              }
              if (sheetSummary.sampleResolved.length < 5) {
                sheetSummary.sampleResolved.push({
                  row: r + 2,
                  column: columnHeader,
                  userId: canonicalId,
                  via: usageKey
                });
              }
            } else {
              columnSummary.unresolvedRows += 1;
            }
          } else {
            columnSummary.unresolvedRows += 1;
          }
        }

        if (resolvedThisCell && normalizedCurrent !== finalValue) {
          columnHadChanges = true;
        }

        row[columnIndex] = finalValue;
        columnValues[r] = [finalValue];
      }

      sheetSummary.blankUserIdRows += columnSummary.blankUserIdRows;
      sheetSummary.resolvedRows += columnSummary.resolvedRows;
      sheetSummary.unresolvedRows += columnSummary.unresolvedRows;
      ['id', 'username', 'email'].forEach(key => {
        sheetSummary.identifierUsage[key] += columnSummary.identifierUsage[key];
      });

      if (columnHadChanges) {
        columnWriteQueue.push({ index: columnIndex, values: columnValues });
      }

      sheetSummary.columnSummaries.push(Object.assign({ hadChanges: columnHadChanges }, columnSummary));
    }

    summary.metrics.totalRows += sheetSummary.totalRows;
    summary.metrics.totalBlankUserIdRows += sheetSummary.blankUserIdRows;
    summary.metrics.totalResolved += sheetSummary.resolvedRows;
    summary.metrics.totalUnresolved += sheetSummary.unresolvedRows;

    if (columnWriteQueue.length && rows.length > 0) {
      try {
        columnWriteQueue.forEach(column => {
          sheet.getRange(2, column.index + 1, rows.length, 1).setValues(column.values);
        });
        summary.metrics.sheetsWithUpdates += 1;
        summary.metrics.totalUpdatedRows += sheetSummary.resolvedRows;
        try { invalidateCache && invalidateCache(sheetName); } catch (_) { }
      } catch (writeError) {
        const message = writeError && writeError.message ? writeError.message : String(writeError);
        summary.success = false;
        summary.errors.push({ sheet: sheetName, stage: 'write', error: message });
        sheetSummary.writeError = message;
        _userLog_('UserService.reconcileUserIdReferences.writeFailed', { sheet: sheetName, error: message }, 'error');
      }
    }

    summary.sheets.push(sheetSummary);
    _userLog_('UserService.reconcileUserIdReferences.sheetSummary', sheetSummary);
  });

  _userLog_('UserService.reconcileUserIdReferences.summary', summary);
  return summary;
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
const REQUIRED_USER_COLUMNS = USER_SHEET_COLUMN_ORDER.slice(0, 21);

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
          FullName: existingUser.FullName || _getUserName_(existingUser),
          UserName: _getUserName_(existingUser),
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

function clientCheckUserConflicts(payload) {
  try {
    const email = payload && (payload.email || payload.Email) ? String(payload.email || payload.Email).trim() : '';
    const userName = payload && (payload.userName || payload.UserName) ? String(payload.userName || payload.UserName).trim() : '';
    const excludeId = payload && (payload.excludeUserId || payload.excludeId || payload.userId || payload.ID || payload.id)
      ? String(payload.excludeUserId || payload.excludeId || payload.userId || payload.ID || payload.id)
      : '';

    const emailKey = _normEmail_(email);
    const userKey = _normUser_(userName);

    const conflicts = {
      emailConflict: null,
      userNameConflict: null
    };

    if (!emailKey && !userKey) {
      return { success: true, hasConflicts: false, conflicts };
    }

    const users = _readUsersAsObjects_();
    users.forEach(u => {
      if (!u || !u.ID) return;
      if (excludeId && String(u.ID) === excludeId) return;

      if (emailKey && !conflicts.emailConflict) {
        if (_normEmail_(u.Email || u.email) === emailKey) {
          conflicts.emailConflict = {
            id: u.ID,
            email: u.Email || u.email || '',
            userName: _getUserName_(u),
            campaignId: u.CampaignID || u.campaignId || ''
          };
        }
      }

      if (userKey && !conflicts.userNameConflict) {
        if (_normUser_(_getUserName_(u)) === userKey) {
          conflicts.userNameConflict = {
            id: u.ID,
            email: u.Email || u.email || '',
            userName: _getUserName_(u),
            campaignId: u.CampaignID || u.campaignId || ''
          };
        }
      }
    });

    return {
      success: true,
      hasConflicts: !!(conflicts.emailConflict || conflicts.userNameConflict),
      conflicts
    };
  } catch (error) {
    writeError && writeError('clientCheckUserConflicts', error);
    return { success: false, error: error.message || String(error) };
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
    try {
      ensureUsersHaveIds();
    } catch (ensureError) {
      try {
        writeError && writeError('clientGetAllUsers.ensureUsersHaveIds', ensureError);
      } catch (_) { }
    }

    let users = [];
    try {
      if (typeof getAllUsersRaw === 'function') {
        users = getAllUsersRaw();
      } else {
        users = readSheet(G.USERS_SHEET);
      }
    } catch (e) { writeError('clientGetAllUsers - readSheet', e); return []; }
    if (!Array.isArray(users) || users.length === 0) return [];

    const enhancedUsers = [];
    for (let i = 0; i < users.length; i++) {
      try {
        const u = users[i];
        if (!u || typeof u !== 'object') {
          _userLog_('clientGetAllUsers.invalidRow', { index: i, row: _userSanitizeForLog_(u) }, 'warn');
          enhancedUsers.push(createMinimalUserObject(u || {}));
          continue;
        }

        if (!u.ID) {
          const coercedId = u.Id || u.id;
          if (coercedId) {
            u.ID = coercedId;
            users[i] = u;
          }
        }

        if (!u.ID) {
          try {
            ensureUsersHaveIds();
          } catch (ensureLoopError) {
            try {
              writeError && writeError('clientGetAllUsers.ensureUsersHaveIds.row', ensureLoopError);
            } catch (_) { }
          }

          if (!u.ID) {
            const refreshedId = u.Id || u.id;
            if (refreshedId) {
              u.ID = refreshedId;
              users[i] = u;
            }
          }
        }

        if (!u.ID) {
          _userLog_('clientGetAllUsers.rowMissingId', {
            index: i,
            row: _userSanitizeForLog_(u)
          }, 'warn');
          enhancedUsers.push(createMinimalUserObject(u));
          continue;
        }

        enhancedUsers.push(createSafeUserObject(u));
      } catch (userErr) {
        enhancedUsers.push(createMinimalUserObject(users[i] || {}));
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
  const raw = (user && typeof user === 'object') ? user : {};
  const safe = {};
  const sheetFieldMap = {};
  const sheetFieldOrder = [];

  const assignField = (key, value) => {
    if (!key && key !== 0) return;
    const strKey = String(key);
    if (!Object.prototype.hasOwnProperty.call(sheetFieldMap, strKey)) sheetFieldOrder.push(strKey);
    sheetFieldMap[strKey] = value;
  };

  Object.keys(raw).forEach(key => {
    const value = raw[key];
    safe[key] = value;
    assignField(key, value);
    if (typeof key === 'string' && key.length) {
      const camel = key.charAt(0).toLowerCase() + key.slice(1);
      if (!Object.prototype.hasOwnProperty.call(safe, camel)) {
        safe[camel] = value;
      }
    }
  });

  safe.ID = safe.ID || safe.id || '';
  assignField('ID', safe.ID);

  const resolvedUserName = _getUserName_(safe) || safe.Email || safe.email || '';
  safe.UserName = resolvedUserName;
  safe.userName = resolvedUserName;
  safe.username = resolvedUserName;
  assignField('UserName', safe.UserName);

  const firstName = safe.FirstName || safe.firstName || '';
  const lastName = safe.LastName || safe.lastName || '';
  if (!safe.FullName && (firstName || lastName)) {
    safe.FullName = [firstName, lastName].filter(Boolean).join(' ');
  }
  safe.FullName = safe.FullName || safe.fullName || '';
  safe.fullName = safe.FullName;
  assignField('FullName', safe.FullName);

  safe.Email = safe.Email || safe.email || safe.EmailAddress || safe.emailAddress || safe.PrimaryEmail || safe.primaryEmail || '';
  safe.email = safe.Email;
  assignField('Email', safe.Email);

  safe.PhoneNumber = safe.PhoneNumber || safe.phoneNumber || safe.Phone || safe.phone || safe.Mobile || safe.mobile || safe.ContactNumber || safe.contactNumber || '';
  safe.phoneNumber = safe.PhoneNumber;
  assignField('PhoneNumber', safe.PhoneNumber);

  safe.CampaignID = safe.CampaignID || safe.CampaignId || safe.campaignID || safe.campaignId || '';
  safe.campaignId = safe.CampaignID;
  assignField('CampaignID', safe.CampaignID);

  const normalizedStatus = normalizeEmploymentStatus(safe.EmploymentStatus || safe.employmentStatus || safe.Status || safe.EmployeeStatus);
  safe.EmploymentStatus = normalizedStatus || safe.EmploymentStatus || safe.employmentStatus || '';
  safe.employmentStatus = safe.EmploymentStatus;
  assignField('EmploymentStatus', safe.EmploymentStatus);

  const hireDateValue = safe.HireDate || safe.hireDate || safe.DateOfHire || safe.dateOfHire || safe.Hire_Date || '';
  safe.HireDate = hireDateValue ? _toIsoDateOnly_(hireDateValue) : '';
  safe.hireDate = safe.HireDate;
  assignField('HireDate', safe.HireDate);

  const countryValue = safe.Country || safe.country || safe.Location || safe.location || safe.CountryOfResidence || '';
  safe.Country = countryValue;
  safe.country = safe.Country;
  assignField('Country', safe.Country);

  const terminationValue = safe.TerminationDate || safe.terminationDate || safe.DateOfTermination || safe.dateOfTermination || '';
  safe.TerminationDate = terminationValue ? _toIsoDateOnly_(terminationValue) : '';
  safe.terminationDate = safe.TerminationDate;
  assignField('TerminationDate', safe.TerminationDate);

  let probationMonths = '';
  if (safe.ProbationMonths !== '' && safe.ProbationMonths != null) probationMonths = Number(safe.ProbationMonths);
  if ((probationMonths === '' || isNaN(probationMonths)) && safe.probationMonths !== '' && safe.probationMonths != null) {
    probationMonths = Number(safe.probationMonths);
  }
  probationMonths = (probationMonths === '' || isNaN(probationMonths)) ? '' : probationMonths;
  safe.ProbationMonths = probationMonths;
  safe.probationMonths = probationMonths;
  assignField('ProbationMonths', safe.ProbationMonths);

  const probationEndValue = safe.ProbationEnd || safe.probationEnd || safe.ProbationEndDate || safe.probationEndDate || '';
  safe.ProbationEnd = probationEndValue ? _toIsoDateOnly_(probationEndValue) : '';
  safe.ProbationEndDate = safe.ProbationEnd;
  safe.probationEnd = safe.ProbationEnd;
  safe.probationEndDate = safe.ProbationEnd;
  assignField('ProbationEnd', safe.ProbationEnd);
  assignField('ProbationEndDate', safe.ProbationEndDate);

  const insuranceEligibleValue = safe.InsuranceEligibleDate || safe.insuranceEligibleDate || safe.InsuranceQualifiedDate || safe.insuranceQualifiedDate || safe.InsuranceEligibilityDate || '';
  safe.InsuranceEligibleDate = insuranceEligibleValue ? _toIsoDateOnly_(insuranceEligibleValue) : '';
  safe.InsuranceQualifiedDate = safe.InsuranceEligibleDate;
  safe.InsuranceEligibilityDate = safe.InsuranceEligibleDate;
  assignField('InsuranceEligibleDate', safe.InsuranceEligibleDate);
  assignField('InsuranceQualifiedDate', safe.InsuranceQualifiedDate);

  let insuranceQualified = null;
  if (safe.InsuranceQualified != null) {
    insuranceQualified = _strToBool_(safe.InsuranceQualified);
  } else if (safe.InsuranceEligible != null) {
    insuranceQualified = _strToBool_(safe.InsuranceEligible);
  } else if (safe.InsuranceEligibleDate) {
    insuranceQualified = isInsuranceEligibleNow_(safe.InsuranceEligibleDate, safe.TerminationDate);
  }
  if (insuranceQualified == null) insuranceQualified = false;
  safe.InsuranceQualified = insuranceQualified;
  safe.InsuranceEligible = insuranceQualified;
  safe.InsuranceQualifiedBool = !!insuranceQualified;
  assignField('InsuranceQualified', _boolToStr_(insuranceQualified));
  assignField('InsuranceEligible', _boolToStr_(insuranceQualified));

  let insuranceEnrolled = null;
  if (safe.InsuranceEnrolled != null) {
    insuranceEnrolled = _strToBool_(safe.InsuranceEnrolled);
  } else if (safe.InsuranceSignedUp != null) {
    insuranceEnrolled = _strToBool_(safe.InsuranceSignedUp);
  }
  if (insuranceEnrolled == null) insuranceEnrolled = false;
  safe.InsuranceEnrolled = insuranceEnrolled;
  safe.InsuranceSignedUp = insuranceEnrolled;
  safe.InsuranceEnrolledBool = !!insuranceEnrolled;
  assignField('InsuranceEnrolled', _boolToStr_(insuranceEnrolled));
  assignField('InsuranceSignedUp', _boolToStr_(insuranceEnrolled));

  const insuranceCardValue = safe.InsuranceCardReceivedDate || safe.insuranceCardReceivedDate || safe.InsuranceCardDate || '';
  safe.InsuranceCardReceivedDate = insuranceCardValue ? _toIsoDateOnly_(insuranceCardValue) : '';
  assignField('InsuranceCardReceivedDate', safe.InsuranceCardReceivedDate);

  try {
    const canLogin = _strToBool_(safe.CanLogin != null ? safe.CanLogin : safe.canLogin);
    safe.canLoginBool = canLogin;
    safe.CanLogin = _boolToStr_(canLogin);
    assignField('CanLogin', safe.CanLogin);
  } catch (_) {
    safe.canLoginBool = false;
  }

  try {
    const isAdmin = _strToBool_(safe.IsAdmin != null ? safe.IsAdmin : safe.isAdmin);
    safe.isAdminBool = isAdmin;
    safe.IsAdmin = _boolToStr_(isAdmin);
    assignField('IsAdmin', safe.IsAdmin);
  } catch (_) {
    safe.isAdminBool = false;
  }

  try {
    const confirmed = _strToBool_(safe.EmailConfirmed != null ? safe.EmailConfirmed : safe.emailConfirmed);
    safe.emailConfirmedBool = confirmed;
    safe.EmailConfirmed = _boolToStr_(confirmed);
    assignField('EmailConfirmed', safe.EmailConfirmed);
  } catch (_) {
    safe.emailConfirmedBool = false;
  }

  safe.needsPasswordSetup = !safe.PasswordHash || _strToBool_(safe.ResetRequired);

  const pagesRaw = Array.isArray(safe.Pages) ? safe.Pages : Array.isArray(safe.pages) ? safe.pages : String(safe.Pages || safe.pages || '').split(',');
  safe.pages = (pagesRaw || []).map(p => String(p || '').trim()).filter(Boolean);
  safe.Pages = safe.pages.join(',');
  assignField('Pages', safe.Pages);

  try {
    safe.campaignName = safe.CampaignID ? getCampaignNameSafe(safe.CampaignID) : '';
  } catch (_) {
    safe.campaignName = '';
  }

  try {
    const userRoleIds = getUserRoleIdsSafe(safe.ID);
    const userRoles = getUserRolesSafe(safe.ID);
    safe.roleIds = userRoleIds;
    safe.roles = userRoles.map(r => r.id);
    safe.roleNames = userRoles.map(r => r.name);
    if (!safe.roleNames.length) {
      const csvRoles = safe.Roles ? String(safe.Roles).split(',').map(r => r.trim()).filter(Boolean) : [];
      safe.csvRoles = csvRoles;
      if (!safe.roleNames.length && csvRoles.length) {
        safe.roleNames = csvRoles.map(roleId => getRoleNameSafe(roleId));
      }
    } else {
      safe.csvRoles = safe.roleNames.length ? userRoles.map(r => r.id) : [];
    }
    safe.Roles = (safe.csvRoles || []).join(',');
    assignField('Roles', safe.Roles);
  } catch (_) {
    safe.roleIds = []; safe.roles = []; safe.roleNames = []; safe.csvRoles = [];
    assignField('Roles', safe.Roles || '');
  }

  try { safe.campaignPermissions = getUserCampaignPermissionsSafe(safe.ID); } catch (_) { safe.campaignPermissions = []; }

  try {
    safe.CreatedAt = safe.CreatedAt ? new Date(safe.CreatedAt).toISOString() : new Date().toISOString();
    safe.UpdatedAt = safe.UpdatedAt ? new Date(safe.UpdatedAt).toISOString() : new Date().toISOString();
    safe.LastLogin = safe.LastLogin ? new Date(safe.LastLogin).toISOString() : (safe.lastLogin ? new Date(safe.lastLogin).toISOString() : null);
    assignField('CreatedAt', safe.CreatedAt);
    assignField('UpdatedAt', safe.UpdatedAt);
    if (safe.LastLogin) assignField('LastLogin', safe.LastLogin);

    safe.hireDateFormatted = safe.HireDate ? new Date(safe.HireDate).toLocaleDateString() : '';
    safe.terminationDateFormatted = safe.TerminationDate ? new Date(safe.TerminationDate).toLocaleDateString() : '';
    safe.probationEndDateFormatted = safe.ProbationEnd ? new Date(safe.ProbationEnd).toLocaleDateString() : '';
    safe.insuranceQualifiedDateFormatted = safe.InsuranceQualifiedDate ? new Date(safe.InsuranceQualifiedDate).toLocaleDateString() : '';
    safe.insuranceCardReceivedDateFormatted = safe.InsuranceCardReceivedDate ? new Date(safe.InsuranceCardReceivedDate).toLocaleDateString() : '';
  } catch (_) { }

  // Ensure canonical sheet columns exist in the export snapshot
  const ensureColumns = (Array.isArray(REQUIRED_USER_COLUMNS) ? REQUIRED_USER_COLUMNS : [])
    .concat(Array.isArray(OPTIONAL_USER_COLUMNS) ? OPTIONAL_USER_COLUMNS : []);
  ensureColumns.forEach(col => {
    if (!Object.prototype.hasOwnProperty.call(sheetFieldMap, col)) {
      assignField(col, typeof safe[col] !== 'undefined' ? safe[col] : '');
    }
  });

  const canonicalOrder = Array.isArray(USER_SHEET_COLUMN_ORDER) ? USER_SHEET_COLUMN_ORDER : [];
  if (canonicalOrder.length) {
    const seen = new Set();
    const ordered = [];
    canonicalOrder.forEach(col => {
      if (!Object.prototype.hasOwnProperty.call(sheetFieldMap, col)) return;
      if (seen.has(col)) return;
      ordered.push(col);
      seen.add(col);
    });
    for (let i = 0; i < sheetFieldOrder.length; i++) {
      const key = sheetFieldOrder[i];
      if (seen.has(key)) continue;
      ordered.push(key);
      seen.add(key);
    }
    sheetFieldOrder.length = 0;
    Array.prototype.push.apply(sheetFieldOrder, ordered);
  }

  safe.sheetFieldMap = {};
  for (let i = 0; i < sheetFieldOrder.length; i++) {
    const key = sheetFieldOrder[i];
    safe.sheetFieldMap[key] = sheetFieldMap[key];
  }
  safe.sheetFieldOrder = sheetFieldOrder.slice();
  safe.sheetFields = sheetFieldOrder.map(key => ({ key, value: sheetFieldMap[key] }));
  safe.sheetFieldCount = safe.sheetFields.length;

  return safe;
}

function createMinimalUserObject(user) {
  return {
    ID: user.ID || '',
    UserName: _getUserName_(user),
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
    ProbationEnd: user.ProbationEnd || user.ProbationEndDate || '',
    InsuranceEligibleDate: user.InsuranceEligibleDate || user.InsuranceQualifiedDate || '',
    InsuranceQualified: (user.InsuranceQualified != null)
      ? _strToBool_(user.InsuranceQualified)
      : _strToBool_(user.InsuranceEligible || false),
    InsuranceEnrolled: (user.InsuranceEnrolled != null)
      ? _strToBool_(user.InsuranceEnrolled)
      : _strToBool_(user.InsuranceSignedUp || false),
    ProbationEndDate: user.ProbationEnd || user.ProbationEndDate || '',
    InsuranceQualifiedDate: user.InsuranceEligibleDate || user.InsuranceQualifiedDate || '',
    InsuranceEligible: (user.InsuranceQualified != null)
      ? _strToBool_(user.InsuranceQualified)
      : _strToBool_(user.InsuranceEligible || false),
    InsuranceSignedUp: (user.InsuranceEnrolled != null)
      ? _strToBool_(user.InsuranceEnrolled)
      : _strToBool_(user.InsuranceSignedUp || false),
    CreatedAt: new Date().toISOString(),
    UpdatedAt: new Date().toISOString(),
    sheetFieldMap: {},
    sheetFieldOrder: [],
    sheetFields: [],
    sheetFieldCount: 0
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

function normalizeRoleIds_(roleIds) {
  if (!Array.isArray(roleIds)) return [];
  var normalized = [];
  var seen = {};
  roleIds.forEach(function (rid) {
    var key = String(rid || '').trim();
    if (!key) return;
    if (!seen[key]) {
      seen[key] = true;
      normalized.push(key);
    }
  });
  return normalized;
}

function syncUserRoles_(userId, desiredRoleIds, options) {
  var summary = { added: 0, removed: 0 };
  try {
    if (!userId) return summary;

    var normalizedDesired = normalizeRoleIds_(desiredRoleIds);

    var existing = getUserRoleIdsSafe(userId).map(function (rid) { return String(rid || '').trim(); }).filter(Boolean);

    var existingUnique = normalizeRoleIds_(existing);
    var hasDuplicates = existing.length !== existingUnique.length;

    var differs = false;
    if (existingUnique.length !== normalizedDesired.length) {
      differs = true;
    } else {
      var desiredLookup = {};
      normalizedDesired.forEach(function (rid) { desiredLookup[rid] = true; });
      for (var i = 0; i < normalizedDesired.length; i++) {
        if (!desiredLookup[existingUnique[i]]) {
          differs = true;
          break;
        }
      }
    }

    if (!differs && !hasDuplicates) {
      return summary;
    }

    if (existing.length && typeof deleteUserRoles === 'function') {
      deleteUserRoles(userId);
      summary.removed = existing.length;
    }

    if (normalizedDesired.length && typeof addUserRole === 'function') {
      var scope = options && options.scope ? options.scope : (options && options.defaultScope ? options.defaultScope : '');
      var assignedBy = options && options.assignedBy ? options.assignedBy : resolveAssignedBy_();
      normalizedDesired.forEach(function (rid) {
        addUserRole(userId, rid, { scope: scope, assignedBy: assignedBy });
      });
      summary.added = normalizedDesired.length;
    }

    return summary;
  } catch (err) {
    throw err;
  }
}

function resolveAssignedBy_() {
  try {
    if (typeof Session === 'undefined' || !Session) return 'System';

    if (typeof Session.getEffectiveUser === 'function') {
      var effective = Session.getEffectiveUser();
      if (effective && typeof effective.getEmail === 'function') {
        var effEmail = effective.getEmail();
        if (effEmail) return effEmail;
      }
    }

    if (typeof Session.getActiveUser === 'function') {
      var active = Session.getActiveUser();
      if (active && typeof active.getEmail === 'function') {
        var actEmail = active.getEmail();
        if (actEmail) return actEmail;
      }
    }
  } catch (_) { }
  return 'System';
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
      const n = _normUser_(_getUserName_(u));
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
    _userLog_('[clientRegisterUser] invoked', { userData });
    const v = _validateUserInput_(userData);
    if (!v.ok) {
      _userLog_('[clientRegisterUser] validation failed', { errors: v.errors }, 'warn');
      return { success: false, error: v.errors.join('; ') };
    }
    const data = _normalizeIncoming_(userData);
    _userLog_('[clientRegisterUser] normalized data', { data });

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
      _userLog_('[clientRegisterUser] existing record located by email', {
        existingId: existing.ID,
        existingUserName: _getUserName_(existing),
        incomingUserName: data.userName,
        mergeRequested: data.mergeIfExists
      });
    } else {
      _userLog_('[clientRegisterUser] no existing user for email', { emailKey });
    }

    if (existing) {
      if (data.mergeIfExists === true || String(data.mergeIfExists).toLowerCase() === 'true') {
        const rowIndex = _findRowIndexById_(values, idx, existing.ID);
        if (rowIndex === -1) return { success: false, error: 'Existing user not found in sheet for merge' };

        if (data.userName && _normUser_(data.userName) !== _normUser_(_getUserName_(existing))) {
          const taken = uniq.byUser.has(_normUser_(data.userName));
          if (taken) {
            _userLog_('[clientRegisterUser] merge aborted due to username conflict', {
              requestedUserName: data.userName
            }, 'warn');
            return { success: false, error: 'Username already exists' };
          }
        }

        // recompute benefits if needed
        const hire = data.hireDate || existing.HireDate || '';
        const probMonths = (data.probationMonths != null) ? data.probationMonths : (existing.ProbationMonths || '');
        const existingProbEnd = existing.ProbationEnd || existing.ProbationEndDate || '';
        const existingEligibleDate = existing.InsuranceEligibleDate || existing.InsuranceQualifiedDate || '';
        const existingQualified = (existing.InsuranceQualified != null)
          ? _strToBool_(existing.InsuranceQualified)
          : _strToBool_(existing.InsuranceEligible || false);
        const existingEnrolled = (existing.InsuranceEnrolled != null)
          ? _strToBool_(existing.InsuranceEnrolled)
          : _strToBool_(existing.InsuranceSignedUp || false);

        const probEnd = data.probationEnd || data.probationEndDate || existingProbEnd || calcProbationEndDate_(hire, probMonths);
        const eligibleDate = data.insuranceEligibleDate || data.insuranceQualifiedDate || existingEligibleDate
          || calcInsuranceEligibleDate_(probEnd, G.INSURANCE_MONTHS_AFTER_PROBATION);
        const qualified = (data.insuranceQualified != null)
          ? _strToBool_(data.insuranceQualified)
          : ((existing.InsuranceQualified != null || existing.InsuranceEligible != null)
            ? existingQualified
            : isInsuranceQualifiedNow_(eligibleDate, data.terminationDate || existing.TerminationDate));
        const enrolled = (data.insuranceEnrolled != null)
          ? _strToBool_(data.insuranceEnrolled)
          : existingEnrolled;

        const merged = {
          ID: existing.ID,
          UserName: data.userName || _getUserName_(existing),
          FullName: data.fullName || existing.FullName,
          Email: existing.Email,
          CampaignID: data.campaignId || existing.CampaignID,
          PasswordHash: values[rowIndex][idx['PasswordHash']],
          ResetRequired: values[rowIndex][idx['ResetRequired']],
          EmailConfirmation: values[rowIndex][idx['EmailConfirmation']],
          EmailConfirmed: 'TRUE',
          PhoneNumber: data.phoneNumber || existing.PhoneNumber,
          EmploymentStatus: data.employmentStatus || normalizeEmploymentStatus(existing.EmploymentStatus) || existing.EmploymentStatus,
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
          ProbationEnd: _toIsoDateOnly_(probEnd),
          ProbationEndDate: _toIsoDateOnly_(probEnd),
          InsuranceEligibleDate: _toIsoDateOnly_(eligibleDate),
          InsuranceQualifiedDate: _toIsoDateOnly_(eligibleDate),
          InsuranceQualified: qualified,
          InsuranceEligible: _boolToStr_(qualified),
          InsuranceEnrolled: enrolled,
          InsuranceSignedUp: _boolToStr_(enrolled),
          InsuranceCardReceivedDate: _toIsoDateOnly_(data.insuranceCardReceivedDate || existing.InsuranceCardReceivedDate || '')
        };

        _userLog_('[clientRegisterUser] merging user payload', { merged });

        const row = [];
        Object.keys(idx).forEach(header => {
          row[idx[header]] = (typeof merged[header] !== 'undefined') ? merged[header] : values[rowIndex][idx[header]];
        });
        sh.getRange(rowIndex + 1, 1, 1, headers.length).setValues([row]);

        try {
          if (typeof invalidateCache === 'function') {
            _userLog_('[clientRegisterUser] cache invalidation start (merge)', { sheets: [G.USERS_SHEET] });
            invalidateCache(G.USERS_SHEET);
            _userLog_('[clientRegisterUser] cache invalidation complete (merge)', {});
          }
        } catch (cacheErr) {
          _userLog_('[clientRegisterUser] cache invalidation failed (merge)', { error: cacheErr && cacheErr.message }, 'warn');
        }
        const mergeResponse = { success: true, alreadyExisted: true, merged: true, userId: existing.ID, message: 'Existing user merged safely (roles/pages/fields updated).' };
        _userLog_('[clientRegisterUser] merge completed', mergeResponse);
        return mergeResponse;
      } else {
        const conflictResponse = {
          success: false,
          alreadyExisted: true,
          userId: existing.ID,
          error: 'A user with this email already exists.',
          conflict: {
            email: existing.Email || '',
            userId: existing.ID,
            userName: _getUserName_(existing) || '',
            campaignId: existing.CampaignID || ''
          }
        };
        _userLog_('[clientRegisterUser] merge denied (email conflict)', conflictResponse, 'warn');
        return conflictResponse;
      }
    }

    const userKey = _normUser_(data.userName);
    if (userKey && uniq.byUser.has(userKey)) {
      _userLog_('[clientRegisterUser] username conflict detected', { requestedUserName: data.userName }, 'warn');
      return { success: false, error: 'Username already exists' };
    }

    const id = Utilities.getUuid();
    const createdAt = _now_();
    const setupToken = Utilities.getUuid();

    // Benefits compute
    const probEnd = data.probationEnd || calcProbationEndDate_(data.hireDate || '', data.probationMonths || '');
    const eligibleDate = data.insuranceEligibleDate || calcInsuranceEligibleDate_(probEnd, G.INSURANCE_MONTHS_AFTER_PROBATION);
    const qualified = (data.insuranceQualified != null)
      ? _strToBool_(data.insuranceQualified)
      : isInsuranceQualifiedNow_(eligibleDate, data.terminationDate);
    const enrolled = (data.insuranceEnrolled != null)
      ? _strToBool_(data.insuranceEnrolled)
      : _strToBool_(data.insuranceSignedUp);

    const normalizedRoleIds = normalizeRoleIds_(data.roles);

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
      Roles: normalizedRoleIds.join(','),
      Pages: (data.pages || []).join(','),
      CreatedAt: createdAt,
      UpdatedAt: createdAt,
      IsAdmin: _boolToStr_(data.isAdmin),
      // Benefits fields
      TerminationDate: _toIsoDateOnly_(data.terminationDate || ''),
      ProbationMonths: (data.probationMonths === '' || data.probationMonths == null) ? '' : Number(data.probationMonths),
      ProbationEnd: _toIsoDateOnly_(data.probationEnd || probEnd),
      ProbationEndDate: _toIsoDateOnly_(data.probationEndDate || probEnd),
      InsuranceEligibleDate: _toIsoDateOnly_(data.insuranceEligibleDate || eligibleDate),
      InsuranceQualifiedDate: _toIsoDateOnly_(data.insuranceQualifiedDate || eligibleDate),
      InsuranceQualified: (data.insuranceQualified != null) ? _strToBool_(data.insuranceQualified) : qualified,
      InsuranceEligible: _boolToStr_((data.insuranceQualified != null) ? data.insuranceQualified : qualified),
      InsuranceEnrolled: enrolled,
      InsuranceSignedUp: _boolToStr_(enrolled),
      InsuranceCardReceivedDate: _toIsoDateOnly_(data.insuranceCardReceivedDate || '')
    };

    _userLog_('[clientRegisterUser] new user row prepared', { newUser });

    const row = [];
    Object.keys(idx).forEach(header => { row[idx[header]] = (typeof newUser[header] !== 'undefined') ? newUser[header] : ''; });
    sh.appendRow(row);
    _userLog_('[clientRegisterUser] row appended', { userId: id, rowValues: row });

    try {
      if (data.campaignId && data.permissionLevel && typeof setCampaignUserPermissions === 'function') {
        _userLog_('[clientRegisterUser] setCampaignUserPermissions invoked', {
          campaignId: data.campaignId,
          userId: id,
          permissionLevel: data.permissionLevel,
          canManageUsers: data.canManageUsers,
          canManagePages: data.canManagePages
        });
        setCampaignUserPermissions(data.campaignId, id, String(data.permissionLevel).toUpperCase(),
          _strToBool_(data.canManageUsers), _strToBool_(data.canManagePages));
      }
      if (Array.isArray(data.roles)) {
        try {
          syncUserRoles_(id, normalizedRoleIds, { assignedBy: resolveAssignedBy_() });
          _userLog_('[clientRegisterUser] syncUserRoles invoked', { userId: id, roles: normalizedRoleIds });
        } catch (syncErr) {
          writeError && writeError('clientRegisterUser:syncUserRoles', syncErr);
          _userLog_('[clientRegisterUser] syncUserRoles failed', { error: syncErr && syncErr.message }, 'warn');
        }
      }
    } catch (pe) { writeError && writeError('clientRegisterUser:perms/roles', pe); }

    try {
      if (typeof invalidateCache === 'function') {
        _userLog_('[clientRegisterUser] cache invalidation start', { sheets: [G.USERS_SHEET, G.CAMPAIGN_USER_PERMISSIONS_SHEET] });
        invalidateCache(G.USERS_SHEET);
        if (G.CAMPAIGN_USER_PERMISSIONS_SHEET) {
          invalidateCache(G.CAMPAIGN_USER_PERMISSIONS_SHEET);
        }
        _userLog_('[clientRegisterUser] cache invalidation complete', {});
      }
    } catch (cacheErr) {
      _userLog_('[clientRegisterUser] cache invalidation failed', { error: cacheErr && cacheErr.message }, 'warn');
    }

    if (data.canLogin && typeof sendPasswordSetupEmail === 'function') {
      try {
        _userLog_('[clientRegisterUser] sending password setup email', { email: data.email, userId: id });
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

    const result = {
      success: true,
      userId: id,
      message: (data.canLogin
        ? `User created. ${(data.pages || []).length} page(s) assigned. Password setup email sent.`
        : `User created. ${(data.pages || []).length} page(s) assigned. Login disabled.`)
    };
    _userLog_('[clientRegisterUser] success response', result);
    return result;

  } catch (e) {
    _userLog_('[clientRegisterUser] error', { message: e && e.message, stack: e && e.stack }, 'error');
    writeError && writeError('clientRegisterUser', e);
    return { success: false, error: e.message || String(e) };
  } finally {
    try { lock.releaseLock(); } catch (_) { }
  }
}

function clientUpdateUser(userId, userData) {
  _userLog_('[clientUpdateUser] invoked', { userId, userData });
  
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
    
    _userLog_('[clientUpdateUser] username after mapping', { userName: userData.userName });
    
    const v = _validateUserInput_(userData);
    if (!v.ok) {
      _userLog_('[clientUpdateUser] validation failed', { errors: v.errors }, 'warn');
      return { success: false, error: v.errors.join('; ') };
    }

    const data = _normalizeIncoming_(userData);
    _userLog_('[clientUpdateUser] normalized data', { data });

    // Additional check for userName after normalization
    if (!data.userName) {
      _userLog_('[clientUpdateUser] username missing after normalization', { userId }, 'warn');
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
      _userLog_('[clientUpdateUser] user not found', { userId }, 'warn');
      return { success: false, error: 'User not found in sheet' };
    }

    const current = {};
    headers.forEach((h, i) => current[h] = values[rowIndex][i]);
    _userLog_('[clientUpdateUser] current row snapshot', { current });

    const desiredEmailKey = _normEmail_(data.email);
    const desiredUserKey = _normUser_(data.userName);

    // Check for email conflicts
    if (desiredEmailKey) {
      const hit = uniq.byEmail.get(desiredEmailKey);
      if (hit && String(hit.ID) !== String(userId)) {
        _userLog_('[clientUpdateUser] email conflict detected', { conflictingUserId: hit.ID, desiredEmailKey }, 'warn');
        return { success: false, error: 'Email already exists for another user' };
      }
    }
    
    // Check for username conflicts
    if (desiredUserKey) {
      const hitU = uniq.byUser.get(desiredUserKey);
      if (hitU && String(hitU.ID) !== String(userId)) {
        _userLog_('[clientUpdateUser] username conflict detected', { conflictingUserId: hitU.ID, desiredUserKey }, 'warn');
        return { success: false, error: 'Username already exists for another user' };
      }
    }

    // Benefits recomputation
    const hire = data.hireDate || current['HireDate'] || '';
    const probMonths = (data.probationMonths != null) ? data.probationMonths : (current['ProbationMonths'] || '');
    const currentProbEnd = current['ProbationEnd'] || current['ProbationEndDate'] || '';
    const term = data.terminationDate || current['TerminationDate'] || '';
    const currentEligibleDate = current['InsuranceEligibleDate'] || current['InsuranceQualifiedDate'] || '';
    const currentQualified = (Object.prototype.hasOwnProperty.call(current, 'InsuranceQualified'))
      ? _strToBool_(current['InsuranceQualified'])
      : _strToBool_(current['InsuranceEligible']);
    const currentEnrolled = (Object.prototype.hasOwnProperty.call(current, 'InsuranceEnrolled'))
      ? _strToBool_(current['InsuranceEnrolled'])
      : _strToBool_(current['InsuranceSignedUp']);

    const probEnd = data.probationEnd || data.probationEndDate || currentProbEnd || calcProbationEndDate_(hire, probMonths);
    const eligibleDate = data.insuranceEligibleDate || data.insuranceQualifiedDate || currentEligibleDate
      || calcInsuranceEligibleDate_(probEnd, G.INSURANCE_MONTHS_AFTER_PROBATION);
    const qualified = (data.insuranceQualified != null)
      ? _strToBool_(data.insuranceQualified)
      : (Object.prototype.hasOwnProperty.call(current, 'InsuranceQualified')
        || Object.prototype.hasOwnProperty.call(current, 'InsuranceEligible'))
        ? currentQualified
        : isInsuranceQualifiedNow_(eligibleDate, term);
    const enrolled = (data.insuranceEnrolled != null)
      ? _strToBool_(data.insuranceEnrolled)
      : currentEnrolled;

    const normalizedRoleIds = normalizeRoleIds_(data.roles);

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
      Roles: normalizedRoleIds.join(','),
      Pages: (data.pages || []).join(','),
      CreatedAt: current['CreatedAt'],
      UpdatedAt: _now_(),
      IsAdmin: _boolToStr_(data.isAdmin),
      // Benefits
      TerminationDate: _toIsoDateOnly_(term),
      ProbationMonths: (probMonths === '' ? '' : Number(probMonths)),
      ProbationEnd: _toIsoDateOnly_(probEnd),
      ProbationEndDate: _toIsoDateOnly_(probEnd),
      InsuranceEligibleDate: _toIsoDateOnly_(eligibleDate),
      InsuranceQualifiedDate: _toIsoDateOnly_(eligibleDate),
      InsuranceQualified: qualified,
      InsuranceEligible: _boolToStr_(qualified),
      InsuranceEnrolled: enrolled,
      InsuranceSignedUp: _boolToStr_(enrolled),
      InsuranceCardReceivedDate: _toIsoDateOnly_(data.insuranceCardReceivedDate || current['InsuranceCardReceivedDate'] || '')
    };
    _userLog_('[clientUpdateUser] updated row payload', { updated });

    const row = [];
    Object.keys(idx).forEach(header => {
      row[idx[header]] = (typeof updated[header] !== 'undefined') ? updated[header] : current[header];
    });
    
    sh.getRange(rowIndex + 1, 1, 1, headers.length).setValues([row]);
    _userLog_('[clientUpdateUser] row updated', { userId, userName: data.userName, rowValues: row });

    try {
      if (data.campaignId && data.permissionLevel && typeof setCampaignUserPermissions === 'function') {
        _userLog_('[clientUpdateUser] setCampaignUserPermissions invoked', {
          campaignId: data.campaignId,
          userId,
          permissionLevel: data.permissionLevel,
          canManageUsers: data.canManageUsers,
          canManagePages: data.canManagePages
        });
        setCampaignUserPermissions(data.campaignId, userId, String(data.permissionLevel).toUpperCase(),
          _strToBool_(data.canManageUsers), _strToBool_(data.canManagePages));
      }
    } catch (pe) {
      writeError && writeError('clientUpdateUser:perms', pe);
      _userLog_('[clientUpdateUser] setCampaignUserPermissions failed', { error: pe && pe.message }, 'warn');
    }

    try {
      if (Array.isArray(data.roles)) {
        syncUserRoles_(userId, normalizedRoleIds, { assignedBy: resolveAssignedBy_() });
        _userLog_('[clientUpdateUser] syncUserRoles invoked', { userId, roles: normalizedRoleIds });
      }
    } catch (roleSyncErr) {
      writeError && writeError('clientUpdateUser:syncUserRoles', roleSyncErr);
      _userLog_('[clientUpdateUser] syncUserRoles failed', { error: roleSyncErr && roleSyncErr.message }, 'warn');
    }

    try {
      if (typeof invalidateCache === 'function') {
        _userLog_('[clientUpdateUser] cache invalidation start', { sheets: [G.USERS_SHEET, G.CAMPAIGN_USER_PERMISSIONS_SHEET] });
        invalidateCache(G.USERS_SHEET);
        if (G.CAMPAIGN_USER_PERMISSIONS_SHEET) {
          invalidateCache(G.CAMPAIGN_USER_PERMISSIONS_SHEET);
        }
        _userLog_('[clientUpdateUser] cache invalidation complete', {});
      }
    } catch (cacheErr) {
      _userLog_('[clientUpdateUser] cache invalidation failed', { error: cacheErr && cacheErr.message }, 'warn');
    }

    const response = {
      success: true,
      message: `User updated successfully with username "${data.userName}". ${(data.pages || []).length} page(s) assigned.`
    };
    _userLog_('[clientUpdateUser] success response', { userId, response });
    return response;

  } catch (e) {
    _userLog_('[clientUpdateUser] error', { message: e && e.message, stack: e && e.stack }, 'error');
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

    let token = null;
    if (typeof IdentityService !== 'undefined'
      && IdentityService
      && typeof IdentityService.beginPasswordReset === 'function') {
      try {
        const resetResult = IdentityService.beginPasswordReset(email, { sendEmail: false, returnTokens: true });
        if (resetResult && resetResult.success !== false && resetResult.resetToken) {
          token = resetResult.resetToken;
        } else {
          console.warn('clientAdminResetPassword: IdentityService returned unexpected result', resetResult);
        }
      } catch (identityErr) {
        writeError && writeError('clientAdminResetPassword:IdentityService', identityErr);
      }
    }

    if (!token) {
      token = Utilities.getUuid();
      const sentAt = new Date();
      const expiresAtDate = new Date(sentAt.getTime() + 60 * 60000);
      const sentAtIso = sentAt.toISOString();
      const expiresAtIso = expiresAtDate.toISOString();
      const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token), Utilities.Charset.UTF_8);
      const tokenHash = _userDigestToHex_(digest);

      if (idx['EmailConfirmation'] >= 0) sheet.getRange(rowIndex + 1, idx['EmailConfirmation'] + 1).setValue(token);
      if (idx['ResetPasswordToken'] >= 0) sheet.getRange(rowIndex + 1, idx['ResetPasswordToken'] + 1).setValue(token);
      if (idx['ResetPasswordTokenHash'] >= 0) sheet.getRange(rowIndex + 1, idx['ResetPasswordTokenHash'] + 1).setValue(tokenHash);
      if (idx['ResetPasswordSentAt'] >= 0) sheet.getRange(rowIndex + 1, idx['ResetPasswordSentAt'] + 1).setValue(sentAtIso);
      if (idx['ResetPasswordExpiresAt'] >= 0) sheet.getRange(rowIndex + 1, idx['ResetPasswordExpiresAt'] + 1).setValue(expiresAtIso);
      if (idx['ResetRequired'] >= 0) sheet.getRange(rowIndex + 1, idx['ResetRequired'] + 1).setValue('TRUE');
      if (idx['UpdatedAt'] >= 0) sheet.getRange(rowIndex + 1, idx['UpdatedAt'] + 1).setValue(sentAtIso);
    }

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

function listUserEquipment_(userId) {
  if (!userId) return [];
  const sheet = _equipmentEnsureSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rowCount = Math.max(0, lastRow - 1);
  if (!rowCount) return [];
  const values = sheet.getRange(2, 1, rowCount, lastCol).getValues();
  const idx = {};
  headers.forEach((h, i) => { idx[h] = i; });
  if (typeof idx.UserID !== 'number') return [];
  const userKey = String(userId);
  const items = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowUser = String(row[idx.UserID] || row[idx.userId] || '');
    if (rowUser !== userKey) continue;
    items.push(_equipmentMapRow_(headers, row));
  }
  items.sort((a, b) => {
    const aTime = Date.parse(a.updatedAt || a.createdAt || 0) || 0;
    const bTime = Date.parse(b.updatedAt || b.createdAt || 0) || 0;
    return bTime - aTime;
  });
  return items;
}

function clientGetUserEquipment(userId) {
  try {
    if (!userId) return { success: false, error: 'User ID required' };
    const items = listUserEquipment_(userId);
    return { success: true, items: items };
  } catch (e) {
    writeError && writeError('clientGetUserEquipment', e);
    return { success: false, error: e.message };
  }
}

function clientSaveUserEquipment(userId, equipmentData) {
  try {
    if (!userId) return { success: false, error: 'User ID required' };
    const data = equipmentData || {};
    const sheet = _equipmentEnsureSheet_();
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const idx = {};
    headers.forEach((h, i) => { idx[h] = i; });
    ['ID', 'UserID', 'ItemName', 'ItemType', 'SerialNumber', 'Condition', 'IssuedDate', 'ReturnedDate',
      'Notes', 'PhotoIds', 'PhotoUrls', 'PhotoNames', 'CreatedAt', 'UpdatedAt']
      .forEach(header => {
        if (typeof idx[header] !== 'number') {
          throw new Error('Equipment sheet missing header: ' + header);
        }
      });

    const lastRow = sheet.getLastRow();
    const rowCount = Math.max(0, lastRow - 1);
    const rangeValues = rowCount ? sheet.getRange(2, 1, rowCount, lastCol).getValues() : [];

    const targetId = data.id ? String(data.id) : '';
    let rowNumber = -1;
    let existingRow = null;
    if (targetId) {
      for (let r = 0; r < rangeValues.length; r++) {
        const row = rangeValues[r];
        if (String(row[idx.ID]) === targetId) {
          rowNumber = r + 2;
          existingRow = row;
          break;
        }
      }
    }

    const equipmentId = targetId || (typeof Utilities !== 'undefined' && Utilities.getUuid ? Utilities.getUuid() : 'EQ-' + Date.now());
    const now = new Date();

    let existingPhotos = [];
    let createdAt = now;
    if (existingRow) {
      const mapped = _equipmentMapRow_(headers, existingRow);
      existingPhotos = Array.isArray(mapped.photos) ? mapped.photos.slice() : [];
      const createdRaw = (typeof idx.CreatedAt === 'number') ? existingRow[idx.CreatedAt] : null;
      if (createdRaw instanceof Date) {
        createdAt = createdRaw;
      } else if (createdRaw) {
        const parsed = new Date(createdRaw);
        createdAt = isNaN(parsed) ? now : parsed;
      } else if (mapped.createdAt) {
        const parsed = new Date(mapped.createdAt);
        if (!isNaN(parsed)) createdAt = parsed;
      }
    }

    const removeSet = new Set((Array.isArray(data.removePhotoIds) ? data.removePhotoIds : []).map(id => String(id)));
    if (removeSet.size) {
      const removedIds = existingPhotos.filter(photo => removeSet.has(String(photo.id))).map(photo => photo.id);
      existingPhotos = existingPhotos.filter(photo => !removeSet.has(String(photo.id)));
      _equipmentRemovePhotos_(removedIds);
    }

    if (Array.isArray(data.newPhotos) && data.newPhotos.length) {
      const uploaded = _equipmentSavePhotos_(userId, equipmentId, data.newPhotos);
      for (let i = 0; i < uploaded.ids.length; i++) {
        existingPhotos.push({
          id: uploaded.ids[i],
          url: uploaded.urls[i] || '',
          name: uploaded.names[i] || ''
        });
      }
    }

    const photoIds = existingPhotos.map(photo => photo.id).filter(Boolean);
    const photoUrls = existingPhotos.map(photo => photo.url || '');
    const photoNames = existingPhotos.map(photo => photo.name || '');

    const rowValues = new Array(headers.length).fill('');
    rowValues[idx.ID] = equipmentId;
    rowValues[idx.UserID] = String(userId);
    rowValues[idx.ItemName] = data.itemName ? String(data.itemName) : '';
    rowValues[idx.ItemType] = data.itemType ? String(data.itemType) : '';
    rowValues[idx.SerialNumber] = data.serialNumber ? String(data.serialNumber) : '';
    rowValues[idx.Condition] = data.condition ? String(data.condition) : '';
    rowValues[idx.IssuedDate] = data.issuedDate ? _toIsoDateOnly_(data.issuedDate) : '';
    rowValues[idx.ReturnedDate] = data.returnedDate ? _toIsoDateOnly_(data.returnedDate) : '';
    rowValues[idx.Notes] = data.notes ? String(data.notes) : '';
    rowValues[idx.PhotoIds] = _equipmentStringifyArray_(photoIds);
    rowValues[idx.PhotoUrls] = _equipmentStringifyArray_(photoUrls);
    rowValues[idx.PhotoNames] = _equipmentStringifyArray_(photoNames);
    rowValues[idx.CreatedAt] = createdAt;
    rowValues[idx.UpdatedAt] = now;

    if (rowNumber > 0) {
      sheet.getRange(rowNumber, 1, 1, headers.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
      rowNumber = sheet.getLastRow();
    }

    const mappedRow = _equipmentMapRow_(headers, rowValues);
    return {
      success: true,
      id: equipmentId,
      item: mappedRow,
      message: existingRow ? 'Equipment updated.' : 'Equipment added.'
    };
  } catch (e) {
    writeError && writeError('clientSaveUserEquipment', e);
    return { success: false, error: e.message };
  }
}

function clientDeleteUserEquipment(equipmentId) {
  try {
    if (!equipmentId) return { success: false, error: 'Equipment ID required' };
    const sheet = _equipmentEnsureSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'Equipment not found' };
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const rowCount = Math.max(0, lastRow - 1);
    const values = rowCount ? sheet.getRange(2, 1, rowCount, lastCol).getValues() : [];
    const idx = {};
    headers.forEach((h, i) => { idx[h] = i; });
    if (typeof idx.ID !== 'number') throw new Error('Equipment sheet missing ID header');
    if (typeof idx.PhotoIds !== 'number') throw new Error('Equipment sheet missing PhotoIds header');

    for (let r = 0; r < values.length; r++) {
      const row = values[r];
      if (String(row[idx.ID]) !== String(equipmentId)) continue;
      const photoIds = _equipmentParseArray_(row[idx.PhotoIds]);
      sheet.deleteRow(r + 2);
      _equipmentRemovePhotos_(photoIds);
      return { success: true, message: 'Equipment entry removed.' };
    }
    return { success: false, error: 'Equipment not found' };
  } catch (e) {
    writeError && writeError('clientDeleteUserEquipment', e);
    return { success: false, error: e.message };
  }
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
        ID: u.ID,
        UserName: _getUserName_(u),
        FullName: u.FullName || _getUserName_(u),
        Email: u.Email,
        CampaignID: u.CampaignID,
        campaignName: getCampaignNameSafe(u.CampaignID),
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
    if (midx.ManagerUserID < 0 || midx.UserID < 0) {
      return { success: true, users: [] };
    }

    const managedIds = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][midx.ManagerUserID]) === String(managerUserId)) {
        managedIds.push(String(data[i][midx.UserID]));
      }
    }
    if (!managedIds.length) return { success: true, users: [] };

    const allUsers = readSheet(G.USERS_SHEET) || [];
    const managedUsers = managedIds.map(id => {
      const match = allUsers.find(u => String(u.ID) === id);
      if (match) {
        const basic = createSafeUserObject(match);
        return {
          ID: basic.ID,
          UserName: basic.UserName,
          FullName: basic.FullName,
          Email: basic.Email,
          CampaignID: basic.CampaignID,
          campaignName: basic.campaignName,
          roleNames: basic.roleNames || []
        };
      }
      return { ID: id };
    });

    return { success: true, users: managedUsers };
  } catch (e) { writeError && writeError('clientGetManagedUsers', e); return { success: false, error: e.message }; }
}

function clientGetManagerTeamSummary(managerUserId) {
  try {
    if (!managerUserId) return { success: false, error: 'Manager ID is required' };

    const managedResponse = clientGetManagedUsers(managerUserId);
    if (!managedResponse || managedResponse.success === false) {
      return managedResponse || { success: false, error: 'Unable to load managed users' };
    }

    const managedUsers = Array.isArray(managedResponse.users) ? managedResponse.users : [];
    const qualityRecords = _loadQualityRecordsForManagerSummary_();
    const summary = _buildManagerTeamSummary_(managedUsers, qualityRecords);
    if (summary && typeof summary === 'object') {
      summary.generatedAt = new Date().toISOString();
    }
    return summary;
  } catch (e) {
    if (typeof writeError === 'function') writeError('clientGetManagerTeamSummary', e);
    return { success: false, error: e && e.message ? e.message : 'Failed to build manager team summary' };
  }
}

function _loadQualityRecordsForManagerSummary_() {
  const sheetName = (typeof QA_RECORDS === 'string' && QA_RECORDS) ? QA_RECORDS : 'Quality';
  try {
    if (typeof readSheet === 'function') {
      const rows = readSheet(sheetName);
      return Array.isArray(rows) ? rows : [];
    }
  } catch (err) {
    if (typeof writeError === 'function') writeError('_loadQualityRecordsForManagerSummary_', err);
  }
  return [];
}

function _buildManagerTeamSummary_(managedUsers, qualityRows) {
  const metrics = [];
  const idMap = Object.create(null);
  const emailMap = Object.create(null);
  const usernameMap = Object.create(null);
  const nameMap = Object.create(null);

  (managedUsers || []).forEach(user => {
    const safeUser = (typeof createSafeUserObject === 'function') ? createSafeUserObject(user) : (user || {});
    const metric = {
      id: _managerSummaryNormalize_(safeUser.ID || safeUser.id),
      name: _managerSummaryNormalize_(safeUser.FullName || safeUser.fullName || safeUser.DisplayName || safeUser.displayName || safeUser.UserName || safeUser.userName),
      email: _managerSummaryNormalize_(safeUser.Email || safeUser.email),
      campaignName: _managerSummaryNormalize_(safeUser.campaignName || safeUser.CampaignName),
      roles: _managerSummaryParseList_(safeUser.roleNames || safeUser.roles),
      evaluations: 0,
      scoreSum: 0,
      scoreCount: 0,
      lastEvaluation: null
    };
    metric.roles = metric.roles.filter(Boolean);
    metrics.push(metric);

    if (metric.id) {
      idMap[metric.id] = metric;
    }
    const emailKey = _managerSummaryNormalizeLower_(metric.email);
    if (emailKey) _managerSummaryMapPush_(emailMap, emailKey, metric);
    const usernameKey = _managerSummaryNormalizeLower_(safeUser.UserName || safeUser.userName || safeUser.Username);
    if (usernameKey) _managerSummaryMapPush_(usernameMap, usernameKey, metric);
    const nameKey = _managerSummaryNormalizeLower_(metric.name);
    if (nameKey) _managerSummaryMapPush_(nameMap, nameKey, metric);
  });

  (Array.isArray(qualityRows) ? qualityRows : []).forEach(record => {
    if (!record || typeof record !== 'object') return;
    const matches = [];

    const idCandidates = [
      record.AgentUserID, record.AgentUserId, record.AgentId,
      record.UserID, record.UserId, record.EmployeeID, record.EmployeeId
    ];
    idCandidates.forEach(candidate => {
      const key = _managerSummaryNormalize_(candidate);
      if (key && idMap[key]) matches.push(idMap[key]);
    });

    const emailCandidates = [
      record.AgentEmail, record.Email, record.AgentWorkEmail,
      record.UserEmail, record.ContactEmail
    ];
    emailCandidates.forEach(candidate => {
      const key = _managerSummaryNormalizeLower_(candidate);
      if (key && emailMap[key]) Array.prototype.push.apply(matches, emailMap[key]);
    });

    const usernameCandidates = [record.AgentUsername, record.UserName, record.Username, record.AgentUserName];
    usernameCandidates.forEach(candidate => {
      const key = _managerSummaryNormalizeLower_(candidate);
      if (key && usernameMap[key]) Array.prototype.push.apply(matches, usernameMap[key]);
    });

    const nameCandidates = [record.AgentName, record.AgentFullName, record.FullName, record.Agent];
    nameCandidates.forEach(candidate => {
      const key = _managerSummaryNormalizeLower_(candidate);
      if (key && nameMap[key]) Array.prototype.push.apply(matches, nameMap[key]);
    });

    if (!matches.length) return;
    const uniqueMatches = [];
    const seen = new Set();
    matches.forEach(metric => {
      if (metric && !seen.has(metric)) {
        seen.add(metric);
        uniqueMatches.push(metric);
      }
    });

    if (!uniqueMatches.length) return;

    const score = _managerSummaryNumber_(
      record.Percentage || record.PercentageScore || record.AgentScore ||
      record.Score || record.TotalScore || record.OverallScore || record.QualityScore || record.AgentPercentage
    );
    const fallbackScore = _managerSummaryNumber_(record.AverageScore || record.AgentAverage || record.QAScore);
    const resolvedScore = score !== null ? score : fallbackScore;
    const evaluatedAt = _managerSummaryDate_(
      record.Timestamp || record.AssessmentDate || record.EvaluatedAt ||
      record.EvaluationDate || record.CreatedAt || record.SubmittedAt || record.UpdatedAt
    );

    uniqueMatches.forEach(metric => {
      metric.evaluations += 1;
      if (resolvedScore !== null) {
        metric.scoreSum += resolvedScore;
        metric.scoreCount += 1;
      }
      if (evaluatedAt && (!metric.lastEvaluation || evaluatedAt > metric.lastEvaluation)) {
        metric.lastEvaluation = evaluatedAt;
      }
    });
  });

  let totalEvaluations = 0;
  let totalScoreSum = 0;
  let totalScoreCount = 0;
  let latestEvaluation = null;

  const users = metrics.map(metric => {
    const average = metric.scoreCount ? metric.scoreSum / metric.scoreCount : null;
    totalEvaluations += metric.evaluations;
    if (metric.scoreCount) {
      totalScoreSum += metric.scoreSum;
      totalScoreCount += metric.scoreCount;
    }
    if (metric.lastEvaluation && (!latestEvaluation || metric.lastEvaluation > latestEvaluation)) {
      latestEvaluation = metric.lastEvaluation;
    }
    return {
      id: metric.id,
      name: metric.name || '',
      email: metric.email || '',
      campaignName: metric.campaignName || '',
      roles: metric.roles,
      evaluations: metric.evaluations,
      averageScore: average !== null ? (Math.round(average * 10) / 10) : null,
      lastEvaluation: metric.lastEvaluation ? metric.lastEvaluation.toISOString() : null
    };
  });

  const teamAverage = totalScoreCount ? (totalScoreSum / totalScoreCount) : null;

  return {
    success: true,
    users: users,
    totals: {
      managedCount: metrics.length,
      evaluatedUsers: metrics.filter(metric => metric.evaluations > 0).length,
      totalEvaluations: totalEvaluations,
      scoreAverage: teamAverage !== null ? (Math.round(teamAverage * 10) / 10) : null,
      lastEvaluation: latestEvaluation ? latestEvaluation.toISOString() : null
    }
  };
}

function _managerSummaryMapPush_(map, key, metric) {
  if (!key || !metric) return;
  if (!map[key]) {
    map[key] = [metric];
    return;
  }
  if (map[key].indexOf(metric) === -1) {
    map[key].push(metric);
  }
}

function _managerSummaryNormalize_(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) {
    return String(value.getTime());
  }
  const text = String(value);
  return text ? text.trim() : '';
}

function _managerSummaryNormalizeLower_(value) {
  const text = _managerSummaryNormalize_(value);
  return text ? text.toLowerCase() : '';
}

function _managerSummaryParseList_(value) {
  if (!value) return [];
  if (Array.isArray(value)) {
    return value.filter(Boolean).map(v => _managerSummaryNormalize_(v)).filter(Boolean);
  }
  if (typeof value === 'string') {
    return value.split(/[,;|]/).map(part => _managerSummaryNormalize_(part)).filter(Boolean);
  }
  return [];
}

function _managerSummaryNumber_(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number' && !isNaN(value)) return value;
  const parsed = parseFloat(String(value).replace(/[^0-9.+-]/g, ''));
  return Number.isFinite(parsed) ? parsed : null;
}

function _managerSummaryDate_(value) {
  if (!value) return null;
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return null;
    return new Date(value.getTime());
  }
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
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
    const managerName = mgr.FullName || _getUserName_(mgr) || 'Your manager';
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
    const rows = (typeof _readRolesSheet_ === 'function')
      ? _readRolesSheet_()
      : __readSheetFallback__(G.ROLES_SHEET);
    const normalizer = (typeof normalizeRoleRecord_ === 'function') ? normalizeRoleRecord_ : __normalizeRoleRecordFallback__;
    return rows.map(normalizer).filter(Boolean);
  } catch (e) {
    writeError && writeError('getAllRoles', e);
    return [];
  }
}
function addUserRole(userId, roleId, scopeOrOptions, assignedBy) {
  try {
    if (!userId || !roleId) throw new Error('User ID and Role ID are required');
    const sheet = SpreadsheetApp.getActive().getSheetByName(G.USER_ROLES_SHEET);
    if (!sheet) throw new Error(`Sheet ${G.USER_ROLES_SHEET} not found`);

    const headers = __getHeaders__(sheet, G.USER_ROLES_HEADER);
    const now = new Date();
    const id = Utilities.getUuid();

    let scope = '';
    let assigned = '';
    if (scopeOrOptions && typeof scopeOrOptions === 'object' && !Array.isArray(scopeOrOptions)) {
      scope = scopeOrOptions.scope || scopeOrOptions.Scope || '';
      assigned = scopeOrOptions.assignedBy || scopeOrOptions.AssignedBy || scopeOrOptions.assigned_by || '';
    } else {
      scope = scopeOrOptions || '';
      assigned = assignedBy || '';
    }

    const rowMap = {
      ID: id,
      UserId: userId,
      RoleId: roleId,
      Scope: scope,
      AssignedBy: assigned,
      CreatedAt: now,
      UpdatedAt: now,
      DeletedAt: ''
    };

    const mapper = (typeof _mapRowFromHeaders_ === 'function') ? _mapRowFromHeaders_ : __mapRowFromHeaders__;
    sheet.appendRow(mapper(headers, rowMap));
    invalidateCache && invalidateCache(G.USER_ROLES_SHEET);
    return id;
  } catch (e) {
    writeError && writeError('addUserRole', e);
  }
}
function deleteUserRoles(userId, roleId) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(G.USER_ROLES_SHEET);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (!data.length) return;
    const headers = data[0].map(h => String(h || '').trim());
    const userIdx = headers.indexOf('UserId');
    const roleIdx = headers.indexOf('RoleId');
    if (userIdx < 0) return;

    for (let i = data.length - 1; i >= 1; i--) {
      const matchesUser = String(data[i][userIdx]) === String(userId);
      const matchesRole = roleId ? String(data[i][roleIdx]) === String(roleId) : true;
      if (matchesUser && matchesRole) {
        sheet.deleteRow(i + 1);
      }
    }
    invalidateCache && invalidateCache(G.USER_ROLES_SHEET);
  } catch (e) {
    writeError && writeError('deleteUserRoles', e);
  }
}
function getRolesMapping() {
  try {
    const roles = getAllRoles();
    const mapping = {};
    roles.forEach(r => {
      const id = r.id || r.ID;
      const name = r.name || r.Name;
      if (id) mapping[id] = name;
    });
    return mapping;
  } catch (error) {
    writeError && writeError('getRolesMapping', error);
    return {};
  }
}

function __readSheetFallback__(sheetName) {
  try {
    const ss = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!ss) return [];
    const range = ss.getDataRange();
    if (!range) return [];
    const values = range.getValues();
    if (!values.length) return [];
    const headers = values.shift().map(h => String(h || '').trim());
    return values
      .map(row => {
        const obj = {};
        let hasData = false;
        headers.forEach((header, idx) => {
          if (!header) return;
          const value = row[idx];
          if (value !== '' && value != null) hasData = true;
          obj[header] = value;
        });
        return hasData ? obj : null;
      })
      .filter(Boolean);
  } catch (error) {
    writeError && writeError('readSheetFallback', error);
    return [];
  }
}

function __normalizeRoleRecordFallback__(record) {
  if (!record || typeof record !== 'object') return null;
  const base = Object.assign({}, record);
  let id = base.ID || base.Id || base.id || '';
  if (id != null) id = String(id).trim();
  let name = base.Name || base.name || '';
  if (name != null) name = String(name).trim();
  const normalizedName = (base.NormalizedName || base.normalizedName || (name ? name.toUpperCase() : '') || '').toString();
  const scope = base.Scope || base.scope || '';
  const description = base.Description || base.description || '';
  const createdValue = base.CreatedAt || base.createdAt || null;
  const updatedValue = base.UpdatedAt || base.updatedAt || null;
  const deletedValue = base.DeletedAt || base.deletedAt || null;

  return Object.assign({}, base, {
    ID: id || base.ID,
    id: id,
    Name: name || base.Name,
    name: name,
    NormalizedName: normalizedName,
    normalizedName: normalizedName,
    Scope: scope,
    scope: scope,
    Description: description,
    description: description,
    CreatedAt: createdValue,
    createdAt: __toClientDateFallback__(createdValue),
    UpdatedAt: updatedValue,
    updatedAt: __toClientDateFallback__(updatedValue),
    DeletedAt: deletedValue,
    deletedAt: __toClientDateFallback__(deletedValue)
  });
}

function __toClientDateFallback__(value) {
  if (!value && value !== 0) return null;
  if (value instanceof Date) return value.toISOString();
  return value;
}

function __getHeaders__(sheet, fallback) {
  if (typeof _getSheetHeaders_ === 'function') return _getSheetHeaders_(sheet, fallback);
  if (!sheet) return Array.isArray(fallback) ? fallback.slice() : [];
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return Array.isArray(fallback) ? fallback.slice() : [];
  const values = sheet.getRange(1, 1, 1, lastColumn).getValues();
  if (!values.length) return Array.isArray(fallback) ? fallback.slice() : [];
  return values[0].map(h => String(h || '').trim());
}

function __mapRowFromHeaders__(headers, valueMap) {
  return headers.map(header => {
    if (!header) return '';
    return Object.prototype.hasOwnProperty.call(valueMap, header) ? valueMap[header] : '';
  });
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
  try {
    ensureUsersHaveIds();
  } catch (ensureError) {
    try { writeError && writeError('_readUsersAsObjects_.ensureUsersHaveIds', ensureError); } catch (_) { }
  }
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

  if (userData.employmentStatus) {
    const normalizedStatus = normalizeEmploymentStatus(userData.employmentStatus);
    if (!normalizedStatus) {
      errors.push('Invalid employment status. Valid options: ' + getValidEmploymentStatuses().join(', '));
    }
  }
  if (userData.hireDate && !validateHireDate(userData.hireDate)) {
    errors.push('Invalid hire date. Must be a valid date not in the future');
  }
  if (userData.country && !validateCountry(userData.country)) {
    errors.push('Country must be between 2 and 100 characters');
  }
  // dates sanity (optional)
  ['terminationDate', 'probationEnd', 'probationEndDate', 'insuranceEligibleDate', 'insuranceQualifiedDate', 'insuranceCardReceivedDate'].forEach(k => {
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

  out.employmentStatus = normalizeEmploymentStatus(out.employmentStatus);
  out.country = String(out.country || '').trim();

  if (out.probationEnd == null && out.probationEndDate != null) {
    out.probationEnd = out.probationEndDate;
  }
  if (out.insuranceEligibleDate == null && out.insuranceQualifiedDate != null) {
    out.insuranceEligibleDate = out.insuranceQualifiedDate;
  }
  if (out.insuranceQualified == null && out.insuranceEligible != null) {
    out.insuranceQualified = out.insuranceEligible;
  }
  if (out.insuranceEnrolled == null && out.insuranceSignedUp != null) {
    out.insuranceEnrolled = out.insuranceSignedUp;
  }

  // Maintain legacy mirrors for downstream compatibility
  if (out.probationEnd != null && out.probationEndDate == null) {
    out.probationEndDate = out.probationEnd;
  }
  if (out.insuranceEligibleDate != null && out.insuranceQualifiedDate == null) {
    out.insuranceQualifiedDate = out.insuranceEligibleDate;
  }
  if (out.insuranceQualified != null && out.insuranceEligible == null) {
    out.insuranceEligible = out.insuranceQualified;
  }
  if (out.insuranceEnrolled != null && out.insuranceSignedUp == null) {
    out.insuranceSignedUp = out.insuranceEnrolled;
  }

  // Normalize date-only strings (keep after alias mapping)
  out.hireDate = out.hireDate ? _toIsoDateOnly_(out.hireDate) : '';
  out.terminationDate = out.terminationDate ? _toIsoDateOnly_(out.terminationDate) : '';
  out.probationEnd = out.probationEnd ? _toIsoDateOnly_(out.probationEnd) : '';
  out.probationEndDate = out.probationEnd;
  out.insuranceEligibleDate = out.insuranceEligibleDate ? _toIsoDateOnly_(out.insuranceEligibleDate) : '';
  out.insuranceQualifiedDate = out.insuranceEligibleDate;
  out.insuranceCardReceivedDate = out.insuranceCardReceivedDate ? _toIsoDateOnly_(out.insuranceCardReceivedDate) : '';

  // Normalize booleans
  out.isAdmin = _strToBool_(out.isAdmin);
  out.canLogin = _strToBool_(out.canLogin);
  out.insuranceQualified = (out.insuranceQualified == null || out.insuranceQualified === '')
    ? null
    : _strToBool_(out.insuranceQualified);
  out.insuranceEligible = out.insuranceQualified;
  out.insuranceEnrolled = _strToBool_(out.insuranceEnrolled);
  out.insuranceSignedUp = out.insuranceEnrolled;

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

    const escapeHtml = function (value) {
      if (value === null || value === undefined) return '';
      return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
    };

    const normalizeList = function (value) {
      if (Array.isArray(value)) return value;
      if (value === null || value === undefined) return [];
      if (typeof value === 'string') {
        return value.split(',').map(function (part) { return part.trim(); }).filter(Boolean);
      }
      return [value];
    };

    const brandName = (typeof EMAIL_CONFIG !== 'undefined' && EMAIL_CONFIG.brandName) ? EMAIL_CONFIG.brandName : 'Lumina HQ';
    const supportEmail = (typeof EMAIL_CONFIG !== 'undefined' && EMAIL_CONFIG.supportEmail) ? EMAIL_CONFIG.supportEmail : 'it@vlbpo.com';

    const roleSource = normalizeList(newUser.roles);
    const resolvedRoles = roleSource.map(function (role) {
      let raw = role;
      if (role && typeof role === 'object') {
        raw = role.id || role.ID || role.roleId || role.RoleId || role.name || role.Name || '';
      }
      raw = raw === undefined || raw === null ? '' : raw;
      let display = raw;
      if (typeof getRoleNameSafe === 'function' && raw) {
        try {
          const lookup = getRoleNameSafe(raw);
          display = lookup || display;
        } catch (roleErr) {
          display = display || raw;
        }
      }
      if ((!display || !String(display).trim()) && role && typeof role === 'object') {
        display = role.name || role.Name || raw;
      }
      return String(display || '').trim();
    }).filter(Boolean);

    const roleBadges = resolvedRoles.length
      ? resolvedRoles.map(function (name) { return '<span class="role-chip">' + escapeHtml(name) + '</span>'; }).join('')
      : '<span class="role-chip role-chip--ghost">No roles assigned yet</span>';

    const pageSource = normalizeList(newUser.pages);
    const resolvedPages = pageSource.map(function (page) {
      if (page && typeof page === 'object') {
        const candidate = page.name || page.title || page.displayName || page.PageName || page.Page || page.key || page.Key || page.slug || page;
        return String(candidate || '').trim();
      }
      return String(page || '').trim();
    }).filter(Boolean);

    const pageBadges = resolvedPages.length
      ? resolvedPages.map(function (name) { return '<span class="page-chip">' + escapeHtml(name) + '</span>'; }).join('')
      : '<span class="page-chip page-chip--ghost">No pages assigned</span>';

    let campaignName = '';
    try { campaignName = getCampaignNameSafe ? getCampaignNameSafe(newUser.campaignId) : getCampaignName(newUser.campaignId); } catch (_) { }
    const campaignDisplay = campaignName || newUser.campaignId || '—';

    let hireDateDisplay = '—';
    if (newUser.hireDate) {
      try {
        const hireDate = new Date(newUser.hireDate);
        hireDateDisplay = isNaN(hireDate.getTime()) ? String(newUser.hireDate) : hireDate.toLocaleDateString();
      } catch (hireErr) {
        hireDateDisplay = String(newUser.hireDate);
      }
    }

    let createdAtDisplay = '—';
    if (newUser.createdAt) {
      try {
        const created = new Date(newUser.createdAt);
        createdAtDisplay = isNaN(created.getTime()) ? String(newUser.createdAt) : created.toLocaleString();
      } catch (createdErr) {
        createdAtDisplay = String(newUser.createdAt);
      }
    }

    const loginIndicator = newUser.canLogin
      ? '<span class="status-indicator"><span class="status-dot"></span>Login Enabled</span>'
      : '<span class="status-indicator status-indicator--warning"><span class="status-dot"></span>Login Disabled</span>';

    const adminIndicator = newUser.isAdmin
      ? '<span class="status-indicator status-indicator--info"><span class="status-dot"></span>Administrator Privileges</span>'
      : '<span class="status-indicator status-indicator--muted"><span class="status-dot"></span>Standard Access</span>';

    const brandNameEscaped = escapeHtml(brandName);
    const supportEmailEscaped = escapeHtml(supportEmail);

    const phoneMarkup = newUser.phoneNumber
      ? '            <p style="margin:4px 0 0;color:#64748b;font-size:13px;">' + escapeHtml(newUser.phoneNumber) + '</p>'
      : '';

    const employmentStatusDisplay = escapeHtml(newUser.employmentStatus || '—');
    const countryDisplay = escapeHtml(newUser.country || '—');

    const content = [
      '<div class="welcome-badge" style="background:linear-gradient(135deg,#00AEEF 0%,#003F87 100%);border:1px solid rgba(255,255,255,.35);box-shadow:0 15px 35px rgba(0,63,135,.35);">🚀 ' + brandNameEscaped + ' Account</div>',
      '<p class="subtitle">A new team member has been activated in ' + brandNameEscaped + '. Here\'s their launch profile.</p>',
      '<div class="futuristic-card">',
      '    <div class="detail-grid">',
      '        <div class="detail-item">',
      '            <span class="muted-label">Employee</span>',
      '            <strong>' + escapeHtml(newUser.fullName || newUser.userName || newUser.email || '—') + '</strong>',
      '            <p style="margin:6px 0 0;color:#475569;">' + escapeHtml(newUser.email || '—') + '</p>',
      phoneMarkup,
      '        </div>',
      '        <div class="detail-item">',
      '            <span class="muted-label">Account Created</span>',
      '            <strong>' + escapeHtml(createdAtDisplay) + '</strong>',
      '            <div style="margin-top:12px;display:flex;flex-wrap:wrap;gap:10px;">',
      '                ' + loginIndicator,
      '                ' + adminIndicator,
      '            </div>',
      '        </div>',
      '    </div>',
      '    <div style="margin-top:24px;">',
      '        <span class="muted-label">Roles &amp; Access Levels</span>',
      '        ' + roleBadges,
      '    </div>',
      '    <div style="margin-top:20px;">',
      '        <span class="muted-label">Pages Enabled</span>',
      '        ' + pageBadges,
      '    </div>',
      '</div>',
      '<div class="info-matrix">',
      '    <div class="info-matrix-row">',
      '        <div class="info-matrix-cell">',
      '            <span class="muted-label">Campaign</span>',
      '            <strong>' + escapeHtml(campaignDisplay) + '</strong>',
      '        </div>',
      '        <div class="info-matrix-cell">',
      '            <span class="muted-label">Employment Status</span>',
      '            <strong>' + employmentStatusDisplay + '</strong>',
      '        </div>',
      '    </div>',
      '    <div class="info-matrix-row">',
      '        <div class="info-matrix-cell">',
      '            <span class="muted-label">Hire Date</span>',
      '            <strong>' + escapeHtml(hireDateDisplay) + '</strong>',
      '        </div>',
      '        <div class="info-matrix-cell">',
      '            <span class="muted-label">Country</span>',
      '            <strong>' + countryDisplay + '</strong>',
      '        </div>',
      '    </div>',
      '</div>',
      '<div class="info-card" style="background:linear-gradient(135deg,rgba(0,174,239,.08),rgba(0,63,135,.08));border-left-color:#00AEEF;">',
      '    <h3 style="margin:0 0 10px;color:#003F87;">Next Actions</h3>',
      '    <ul class="feature-list">',
      '        <li style="border-left-color:#00AEEF;">Confirm tool provisioning within 24 hours</li>',
      '        <li style="border-left-color:#00AEEF;">Share onboarding checklist with the employee</li>',
      '        <li style="border-left-color:#00AEEF;">Welcome them in the Lumina HQ collaboration channels</li>',
      '    </ul>',
      '    <p style="margin-top:18px;color:#475569;">Need adjustments? Contact <a href="mailto:' + supportEmailEscaped + '" class="support-link">' + supportEmailEscaped + '</a>.</p>',
      '</div>'
    ].filter(Boolean).join('\n');

    const html = buildHtmlEmail_(
      'New Team Member Activated', 'Access summary for ' + brandName + '.',
      content
    );

    safeSendHtmlEmail_(recipient, 'New Team Member: ' + (newUser.userName || newUser.email || newUser.id), html);
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
  try {
    return {
      success: true,
      statuses: getValidEmploymentStatuses(),
      aliases: Object.assign({}, EMPLOYMENT_STATUS_ALIAS_MAP),
      message: 'Valid employment statuses retrieved'
    };
  }
  catch (e) {
    writeError && writeError('clientGetValidEmploymentStatuses', e);
    return { success: false, error: e.message, statuses: [] };
  }
}
function clientGetEmploymentStatusReport(campaignId) {
  try {
    const users = readSheet(G.USERS_SHEET) || [];
    let filtered = users;
    if (campaignId) filtered = users.filter(u => u.CampaignID === campaignId);

    const statusCounts = {};
    const valid = getValidEmploymentStatuses();
    valid.forEach(s => statusCounts[s] = 0);
    statusCounts['Unspecified'] = 0;

    const LOOKBACK_DAYS = 90;
    const toStartOfDay = (value) => {
      const d = (value instanceof Date) ? new Date(value.getTime()) : new Date(value);
      if (isNaN(d)) return null;
      d.setHours(0, 0, 0, 0);
      return d;
    };
    const fromSpreadsheetDate = (value) => {
      if (!value && value !== 0) return null;
      if (value instanceof Date) return isNaN(value) ? null : new Date(value.getTime());
      if (typeof value === 'number') {
        const ms = Math.round((Number(value) - 25569) * 86400 * 1000);
        const d = new Date(ms);
        return isNaN(d) ? null : d;
      }
      if (typeof value === 'string') {
        const trimmed = value.trim();
        if (!trimmed) return null;
        const numeric = Number(trimmed);
        if (!Number.isNaN(numeric) && trimmed.replace(/[0-9.]/g, '') === '') {
          const numericDate = new Date(Math.round((numeric - 25569) * 86400 * 1000));
          if (!isNaN(numericDate)) return numericDate;
        }
        const parsed = new Date(trimmed);
        if (!isNaN(parsed)) return parsed;
        const match = trimmed.match(/^([0-9]{1,2})[\/\-]([0-9]{1,2})[\/\-]([0-9]{2,4})$/);
        if (match) {
          const month = Number(match[1]);
          const day = Number(match[2]);
          let year = Number(match[3]);
          if (year < 100) year += year >= 70 ? 1900 : 2000;
          const fallback = new Date(year, month - 1, day);
          if (!isNaN(fallback)) return fallback;
        }
      }
      return null;
    };

    const today = toStartOfDay(new Date());
    const windowStart = today ? new Date(today.getTime()) : null;
    if (windowStart) windowStart.setDate(windowStart.getDate() - Math.max(0, LOOKBACK_DAYS - 1));

    let totalHires = 0;
    let totalTerminations = 0;
    let hiresWithinWindow = 0;
    let terminationsWithinWindow = 0;
    let activeHeadcount = 0;

    filtered.forEach(u => {
      const normalized = normalizeEmploymentStatus(u && (u.EmploymentStatus || u.employmentStatus));
      if (normalized) {
        if (typeof statusCounts[normalized] !== 'number') statusCounts[normalized] = 0;
        statusCounts[normalized]++;
      } else {
        statusCounts['Unspecified']++;
      }

      const hireDateRaw = fromSpreadsheetDate(u && (u.HireDate || u.hireDate || u.DateHired || u.dateHired));
      const terminationDateRaw = fromSpreadsheetDate(u && (u.TerminationDate || u.terminationDate || u.DateOfTermination || u.dateOfTermination));

      const hireDate = hireDateRaw ? toStartOfDay(hireDateRaw) : null;
      const terminationDate = terminationDateRaw ? toStartOfDay(terminationDateRaw) : null;

      if (hireDate && today && hireDate <= today) {
        totalHires++;
        if (windowStart && hireDate >= windowStart) {
          hiresWithinWindow++;
        }
      }

      const terminationEffective = terminationDate && today && terminationDate <= today;
      if (terminationEffective) {
        totalTerminations++;
        if (windowStart && terminationDate >= windowStart) {
          terminationsWithinWindow++;
        }
      }

      if (!terminationDate || (today && terminationDate > today)) {
        activeHeadcount++;
      }
    });

    const hiringRate = activeHeadcount > 0
      ? Number(((hiresWithinWindow / activeHeadcount) * 100).toFixed(1))
      : 0;
    const turnoverRate = activeHeadcount > 0
      ? Number(((terminationsWithinWindow / activeHeadcount) * 100).toFixed(1))
      : 0;

    return {
      success: true,
      campaignId,
      totalUsers: filtered.length,
      statusCounts,
      validStatuses: valid,
      turnoverMetrics: {
        lookbackDays: LOOKBACK_DAYS,
        windowStart: windowStart ? _toIsoDateOnly_(windowStart) : '',
        windowEnd: today ? _toIsoDateOnly_(today) : '',
        activeHeadcount,
        totalHires,
        totalTerminations,
        hiresWithinWindow,
        terminationsWithinWindow,
        hiringRate,
        turnoverRate
      },
      message: `Employment status report for ${filtered.length} users`
    };
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
    const probEnd = u.ProbationEnd || u.ProbationEndDate || calcProbationEndDate_(hire, probMonths);
    const eligibleDate = u.InsuranceEligibleDate || u.InsuranceQualifiedDate
      || calcInsuranceEligibleDate_(probEnd, G.INSURANCE_MONTHS_AFTER_PROBATION);
    const qualified = (u.InsuranceQualified != null)
      ? _strToBool_(u.InsuranceQualified)
      : isInsuranceQualifiedNow_(eligibleDate, u.TerminationDate || '');
    const enrolled = (u.InsuranceEnrolled != null)
      ? _strToBool_(u.InsuranceEnrolled)
      : _strToBool_(u.InsuranceSignedUp);
    return {
      success: true,
      userId,
      hireDate: hire || '',
      probationMonths: probMonths === '' ? '' : Number(probMonths),
      probationEnd: _toIsoDateOnly_(probEnd),
      probationEndDate: _toIsoDateOnly_(probEnd),
      insuranceEligibleDate: _toIsoDateOnly_(eligibleDate),
      insuranceQualifiedDate: _toIsoDateOnly_(eligibleDate),
      insuranceQualified: !!qualified,
      insuranceEligible: !!qualified,
      insuranceEnrolled: !!enrolled,
      insuranceSignedUp: !!enrolled,
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

      const curProbEnd = (typeof idx['ProbationEnd'] === 'number') ? row[idx['ProbationEnd']] : '';
      const curProbEndLegacy = (typeof idx['ProbationEndDate'] === 'number') ? row[idx['ProbationEndDate']] : '';
      const curEligibleDate = (typeof idx['InsuranceEligibleDate'] === 'number') ? row[idx['InsuranceEligibleDate']] : '';
      const curEligibleDateLegacy = (typeof idx['InsuranceQualifiedDate'] === 'number') ? row[idx['InsuranceQualifiedDate']] : '';

      const probEnd = curProbEnd || curProbEndLegacy || calcProbationEndDate_(hire, probMonths);
      const eligibleDate = curEligibleDate || curEligibleDateLegacy || calcInsuranceEligibleDate_(probEnd, G.INSURANCE_MONTHS_AFTER_PROBATION);
      const qualified = isInsuranceQualifiedNow_(eligibleDate, term);

      const newProbEnd = _toIsoDateOnly_(probEnd);
      const newEligibleIso = _toIsoDateOnly_(eligibleDate);
      const newQualifiedBool = !!qualified;
      const newQualifiedStr = _boolToStr_(newQualifiedBool);

      let changed = false;
      function setIfDiff(colName, val) {
        if (typeof idx[colName] !== 'number') return;
        const col = idx[colName] + 1;
        const cur = row[idx[colName]];
        if (String(cur) !== String(val)) { sh.getRange(r + 1, col).setValue(val); row[idx[colName]] = val; changed = true; }
      }
      setIfDiff('ProbationEnd', newProbEnd);
      setIfDiff('ProbationEndDate', newProbEnd);
      setIfDiff('InsuranceEligibleDate', newEligibleIso);
      setIfDiff('InsuranceQualifiedDate', newEligibleIso);
      setIfDiff('InsuranceQualified', newQualifiedBool);
      setIfDiff('InsuranceEligible', newQualifiedStr);

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
      const n = _normUser_(_getUserName_(u));
      
      // Log problematic entries
      if (!e && u && (u.Email || u.email)) {
        console.warn('Failed to normalize email for user at index', index, ':', u.Email || u.email);
      }
      if (!n && _getUserName_(u)) {
        console.warn('Failed to normalize username for user at index', index, ':', _getUserName_(u));
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

var UserService = typeof UserService !== 'undefined' ? UserService : {};
UserService.ensureUsersHaveIds = ensureUsersHaveIds;
UserService.reconcileUserIdReferencesAcrossSheets = reconcileUserIdReferencesAcrossSheets;
UserService.buildUserIdentifierLookup = _buildUserIdentifierLookup_;

console.log('✅ UserService.gs loaded');
console.log('📦 Features: User CRUD, Roles, Pages, Campaign perms, Manager mapping, HR/Benefits (probation + insurance)');
