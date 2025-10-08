/**
 * MainUtilities.gs
 * Core system utilities for identity management, campaigns, pages, chat, and shared infrastructure.
 * - Collision-safe globals (all sheet names & headers are guarded)
 * - Enhanced page discovery & categorization
 * - Multi-campaign navigation
 * - Manager–Users canonical helpers
 */

// ────────────────────────────────────────────────────────────────────────────
// Cache & constants (guarded)
// ────────────────────────────────────────────────────────────────────────────
if (typeof CACHE_TTL_SEC === 'undefined') var CACHE_TTL_SEC = 600;
if (typeof PAGE_SIZE === 'undefined') var PAGE_SIZE = 10;
if (typeof MAX_BATCH_SIZE === 'undefined') var MAX_BATCH_SIZE = 200;
if (typeof scriptCache === 'undefined') var scriptCache = CacheService.getScriptCache();

// ────────────────────────────────────────────────────────────────────────────
// Identity & Core System Sheet names (guarded)
// ────────────────────────────────────────────────────────────────────────────
if (typeof USERS_SHEET === 'undefined') var USERS_SHEET = "Users";
if (typeof ROLES_SHEET === 'undefined') var ROLES_SHEET = "Roles";
if (typeof USER_ROLES_SHEET === 'undefined') var USER_ROLES_SHEET = "UserRoles";
if (typeof USER_CLAIMS_SHEET === 'undefined') var USER_CLAIMS_SHEET = "UserClaims";
if (typeof SESSIONS_SHEET === 'undefined') var SESSIONS_SHEET = "Sessions";

if (typeof CAMPAIGNS_SHEET === 'undefined') var CAMPAIGNS_SHEET = "Campaigns";
if (typeof PAGES_SHEET === 'undefined') var PAGES_SHEET = "Pages";
if (typeof CAMPAIGN_PAGES_SHEET === 'undefined') var CAMPAIGN_PAGES_SHEET = "CampaignPages";
if (typeof PAGE_CATEGORIES_SHEET === 'undefined') var PAGE_CATEGORIES_SHEET = "PageCategories";
if (typeof CAMPAIGN_USER_PERMISSIONS_SHEET === 'undefined') var CAMPAIGN_USER_PERMISSIONS_SHEET = "CampaignUserPermissions";
if (typeof USER_MANAGERS_SHEET === 'undefined') var USER_MANAGERS_SHEET = "UserManagers";
if (typeof USER_CAMPAIGNS_SHEET === 'undefined') var USER_CAMPAIGNS_SHEET = "UserCampaigns";

if (typeof DEBUG_LOGS_SHEET === 'undefined') var DEBUG_LOGS_SHEET = "DebugLogs";
if (typeof ERROR_LOGS_SHEET === 'undefined') var ERROR_LOGS_SHEET = "ErrorLogs";
if (typeof NOTIFICATIONS_SHEET === 'undefined') var NOTIFICATIONS_SHEET = "Notifications";

// Multi-campaign banner settings
if (typeof MULTI_CAMPAIGN_NAME === 'undefined') var MULTI_CAMPAIGN_NAME = 'MultiCampaign(System Admin)';
if (typeof MULTI_CAMPAIGN_ICON === 'undefined') var MULTI_CAMPAIGN_ICON = 'fas fa-building';

// Chat system sheets (guarded)
if (typeof CHAT_GROUPS_SHEET === 'undefined') var CHAT_GROUPS_SHEET = 'ChatGroups';
if (typeof CHAT_CHANNELS_SHEET === 'undefined') var CHAT_CHANNELS_SHEET = 'ChatChannels';
if (typeof CHAT_MESSAGES_SHEET === 'undefined') var CHAT_MESSAGES_SHEET = 'ChatMessages';
if (typeof CHAT_GROUP_MEMBERS_SHEET === 'undefined') var CHAT_GROUP_MEMBERS_SHEET = 'ChatGroupMembers';
if (typeof CHAT_MESSAGE_REACTIONS_SHEET === 'undefined') var CHAT_MESSAGE_REACTIONS_SHEET = 'ChatMessageReactions';
if (typeof CHAT_USER_PREFERENCES_SHEET === 'undefined') var CHAT_USER_PREFERENCES_SHEET = 'ChatUserPreferences';
if (typeof CHAT_ANALYTICS_SHEET === 'undefined') var CHAT_ANALYTICS_SHEET = 'ChatAnalytics';
if (typeof CHAT_CHANNEL_MEMBERS_SHEET === 'undefined') var CHAT_CHANNEL_MEMBERS_SHEET = 'ChatChannelMembers';

// ────────────────────────────────────────────────────────────────────────────
// Headers (guarded)
// ────────────────────────────────────────────────────────────────────────────
if (typeof USERS_HEADERS === 'undefined') var USERS_HEADERS = [
  "ID",
  "UserName",
  "FullName",
  "Email",
  "CampaignID",
  "PasswordHash",
  "ResetRequired",
  "EmailConfirmation",
  "EmailConfirmed",
  "PhoneNumber",
  "EmploymentStatus",
  "HireDate",
  "Country",
  "LockoutEnd",
  "TwoFactorEnabled",
  "CanLogin",
  "Roles",
  "Pages",
  "CreatedAt",
  "UpdatedAt",
  "IsAdmin",
  "NormalizedUserName",
  "NormalizedEmail",
  "PhoneNumberConfirmed",
  "LockoutEnabled",
  "AccessFailedCount",
  "TwoFactorDelivery",
  "TwoFactorSecret",
  "TwoFactorRecoveryCodes",
  "SecurityStamp",
  "ConcurrencyStamp",
  "EmailConfirmationTokenHash",
  "EmailConfirmationSentAt",
  "EmailConfirmationExpiresAt",
  "ResetPasswordToken",
  "ResetPasswordTokenHash",
  "ResetPasswordSentAt",
  "ResetPasswordExpiresAt",
  "LastLogin",
  "LastLoginAt",
  "LastLoginIp",
  "LastLoginUserAgent",
  "DeletedAt",
  "TerminationDate",
  "ProbationMonths",
  "ProbationEnd",
  "ProbationEndDate",
  "InsuranceEligibleDate",
  "InsuranceQualifiedDate",
  "InsuranceEligible",
  "InsuranceQualified",
  "InsuranceEnrolled",
  "InsuranceSignedUp",
  "InsuranceCardReceivedDate",
  "MFASecret",
  "MFABackupCodes",
  "MFADeliveryPreference",
  "MFAEnabled"
];

if (typeof ROLES_HEADER === 'undefined') var ROLES_HEADER = [
  "ID",
  "Name",
  "NormalizedName",
  "Scope",
  "Description",
  "CreatedAt",
  "UpdatedAt",
  "DeletedAt"
];
if (typeof USER_ROLES_HEADER === 'undefined') var USER_ROLES_HEADER = [
  "ID",
  "UserId",
  "RoleId",
  "Scope",
  "AssignedBy",
  "CreatedAt",
  "UpdatedAt",
  "DeletedAt"
];
if (typeof CLAIMS_HEADERS === 'undefined') var CLAIMS_HEADERS = ["ID", "UserId", "ClaimType", "CreatedAt", "UpdatedAt"];
if (typeof SESSIONS_HEADERS === 'undefined') var SESSIONS_HEADERS = [
  "Token",
  "TokenHash",
  "TokenSalt",
  "UserId",
  "CreatedAt",
  "LastActivityAt",
  "ExpiresAt",
  "IdleTimeoutMinutes",
  "RememberMe",
  "CampaignScope",
  "UserAgent",
  "IpAddress",
  "ServerIp"
];

if (typeof CHAT_GROUPS_HEADERS === 'undefined') var CHAT_GROUPS_HEADERS = ['ID', 'Name', 'Description', 'CreatedBy', 'CreatedAt', 'UpdatedAt'];
if (typeof CHAT_CHANNELS_HEADERS === 'undefined') var CHAT_CHANNELS_HEADERS = ['ID', 'GroupId', 'Name', 'Description', 'IsPrivate', 'CreatedBy', 'CreatedAt', 'UpdatedAt'];
if (typeof CHAT_MESSAGES_HEADERS === 'undefined') var CHAT_MESSAGES_HEADERS = ['ID', 'ChannelId', 'UserId', 'Message', 'Timestamp', 'EditedAt', 'ParentMessageId', 'IsDeleted'];
if (typeof CHAT_GROUP_MEMBERS_HEADERS === 'undefined') var CHAT_GROUP_MEMBERS_HEADERS = ['ID', 'GroupId', 'UserId', 'JoinedAt', 'Role', 'IsActive'];
if (typeof CHAT_MESSAGE_REACTIONS_HEADERS === 'undefined') var CHAT_MESSAGE_REACTIONS_HEADERS = ['ID', 'MessageId', 'UserId', 'Reaction', 'Timestamp'];
if (typeof CHAT_USER_PREFERENCES_HEADERS === 'undefined') var CHAT_USER_PREFERENCES_HEADERS = ['UserId', 'NotificationSettings', 'Theme', 'LastSeen', 'Status'];
if (typeof CHAT_ANALYTICS_HEADERS === 'undefined') var CHAT_ANALYTICS_HEADERS = ['Timestamp', 'UserId', 'Action', 'Details', 'SessionId'];
if (typeof CHAT_CHANNEL_MEMBERS_HEADERS === 'undefined') var CHAT_CHANNEL_MEMBERS_HEADERS = ["ID", "ChannelId", "UserId", "JoinedAt", "Role", "IsActive"];

if (typeof CAMPAIGNS_HEADERS === 'undefined') var CAMPAIGNS_HEADERS = [
  "ID",
  "Name",
  "Description",
  "ClientName",
  "Status",
  "Channel",
  "Timezone",
  "SlaTier",
  "CreatedAt",
  "UpdatedAt",
  "DeletedAt"
];
if (typeof PAGES_HEADERS === 'undefined') var PAGES_HEADERS = ["PageKey", "PageTitle", "PageIcon", "Description", "IsSystemPage", "RequiresAdmin", "CreatedAt", "UpdatedAt"];
if (typeof CAMPAIGN_PAGES_HEADERS === 'undefined') var CAMPAIGN_PAGES_HEADERS = ["ID", "CampaignID", "PageKey", "PageTitle", "PageIcon", "CategoryID", "SortOrder", "IsActive", "CreatedAt", "UpdatedAt"];
if (typeof PAGE_CATEGORIES_HEADERS === 'undefined') var PAGE_CATEGORIES_HEADERS = ["ID", "CampaignID", "CategoryName", "CategoryIcon", "SortOrder", "IsActive", "CreatedAt", "UpdatedAt"];
if (typeof CAMPAIGN_USER_PERMISSIONS_HEADERS === 'undefined') var CAMPAIGN_USER_PERMISSIONS_HEADERS = [
  "ID",
  "CampaignID",
  "UserID",
  "PermissionLevel",
  "Role",
  "CanManageUsers",
  "CanManagePages",
  "Notes",
  "CreatedAt",
  "UpdatedAt",
  "DeletedAt"
];
if (typeof USER_MANAGERS_HEADERS === 'undefined') var USER_MANAGERS_HEADERS = ["ID", "ManagerUserID", "ManagedUserID", "CampaignID", "CreatedAt", "UpdatedAt"];
if (typeof ATTENDANCE_LOG_HEADERS === 'undefined') var ATTENDANCE_LOG_HEADERS = ["ID", "Timestamp", "User", "DurationMin", "State", "Date", "UserID", "CreatedAt", "UpdatedAt"];
if (typeof USER_CAMPAIGNS_HEADERS === 'undefined') var USER_CAMPAIGNS_HEADERS = [
  "ID",
  "UserId",
  "CampaignId",
  "Role",
  "IsPrimary",
  "CreatedAt",
  "UpdatedAt",
  "DeletedAt"
];

if (typeof DEBUG_LOGS_HEADERS === 'undefined') var DEBUG_LOGS_HEADERS = ["Timestamp", "Message"];
if (typeof ERROR_LOGS_HEADERS === 'undefined') var ERROR_LOGS_HEADERS = ["Timestamp", "Error", "Stack"];
if (typeof NOTIFICATIONS_HEADERS === 'undefined') var NOTIFICATIONS_HEADERS = ["ID", "UserId", "Type", "Severity", "Title", "Message", "Data", "Read", "ActionTaken", "CreatedAt", "ReadAt", "ExpiresAt"];

// ────────────────────────────────────────────────────────────────────────────
// Canonical identity/authentication sheet definitions
// ────────────────────────────────────────────────────────────────────────────

function buildCanonicalIdentitySheetMap_() {
  const definitions = {};

  const addDefinition = (name, headers) => {
    if (!name) return;
    const trimmedName = String(name).trim();
    if (!trimmedName || definitions[trimmedName]) return;

    if (!Array.isArray(headers)) return;
    const normalized = headers
      .map(header => (header === null || typeof header === 'undefined') ? '' : String(header).trim())
      .filter(Boolean);
    if (!normalized.length) return;

    definitions[trimmedName] = normalized;
  };

  addDefinition(USERS_SHEET, USERS_HEADERS);
  addDefinition(ROLES_SHEET, ROLES_HEADER);
  addDefinition(USER_ROLES_SHEET, USER_ROLES_HEADER);
  addDefinition(USER_CLAIMS_SHEET, CLAIMS_HEADERS);
  addDefinition(SESSIONS_SHEET, SESSIONS_HEADERS);
  addDefinition(CAMPAIGNS_SHEET, CAMPAIGNS_HEADERS);
  addDefinition(PAGES_SHEET, PAGES_HEADERS);
  addDefinition(CAMPAIGN_PAGES_SHEET, CAMPAIGN_PAGES_HEADERS);
  addDefinition(PAGE_CATEGORIES_SHEET, PAGE_CATEGORIES_HEADERS);
  addDefinition(CAMPAIGN_USER_PERMISSIONS_SHEET, CAMPAIGN_USER_PERMISSIONS_HEADERS);
  addDefinition(USER_MANAGERS_SHEET, USER_MANAGERS_HEADERS);
  addDefinition(USER_CAMPAIGNS_SHEET, USER_CAMPAIGNS_HEADERS);

  return definitions;
}

if (typeof listCanonicalIdentitySheets !== 'function') {
  function listCanonicalIdentitySheets() {
    const map = buildCanonicalIdentitySheetMap_();
    return Object.keys(map).map(name => ({ name, headers: map[name].slice() }));
  }
}

if (typeof getCanonicalSheetHeaders !== 'function') {
  function getCanonicalSheetHeaders(sheetName) {
    if (!sheetName && sheetName !== 0) return null;
    const target = String(sheetName).trim();
    if (!target) return null;
    const map = buildCanonicalIdentitySheetMap_();
    const headers = map[target];
    return Array.isArray(headers) ? headers.slice() : null;
  }
}

function ensureSheetStructureFromDefinition_(definition) {
  const result = {
    sheet: definition && definition.name ? definition.name : '',
    ensured: false,
    method: null,
    error: null
  };

  if (!definition || !definition.name || !Array.isArray(definition.headers) || !definition.headers.length) {
    result.error = 'INVALID_DEFINITION';
    return result;
  }

  try {
    let sheet = null;

    if (typeof ensureSheetWithHeaders === 'function') {
      sheet = ensureSheetWithHeaders(definition.name, definition.headers);
      result.method = 'ensureSheetWithHeaders';
    } else if (typeof synchronizeSheetHeaders === 'function') {
      sheet = synchronizeSheetHeaders(definition.name, definition.headers);
      result.method = 'synchronizeSheetHeaders';
    }

    if (!sheet) {
      if (typeof SpreadsheetApp === 'undefined') {
        throw new Error('SpreadsheetApp not available');
      }
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        throw new Error('Active spreadsheet not available');
      }
      sheet = ss.getSheetByName(definition.name);
      if (!sheet) {
        sheet = ss.insertSheet(definition.name);
        result.method = 'manual-create';
      } else if (!result.method) {
        result.method = 'manual-sync';
      }

      sheet.getRange(1, 1, 1, definition.headers.length).setValues([definition.headers]);
      sheet.setFrozenRows(1);
    }

    result.ensured = true;
    if (sheet && typeof sheet.getSheetId === 'function') {
      result.sheetId = sheet.getSheetId();
    }
    return result;
  } catch (error) {
    result.error = error && error.message ? error.message : String(error);
    console.warn(`ensureSheetStructureFromDefinition_: failed for ${result.sheet}`, error);
    return result;
  }
}

if (typeof ensureIdentitySheetStructures !== 'function') {
  function ensureIdentitySheetStructures(options = {}) {
    const { sheetNames } = options || {};
    const allowList = Array.isArray(sheetNames) && sheetNames.length
      ? new Set(sheetNames.map(name => String(name || '').trim()).filter(Boolean))
      : null;

    const definitions = listCanonicalIdentitySheets();
    const results = [];

    definitions.forEach(definition => {
      if (allowList && !allowList.has(definition.name)) {
        return;
      }
      results.push(ensureSheetStructureFromDefinition_(definition));
    });

    return results;
  }
}

// ────────────────────────────────────────────────────────────────────────────
// HR / Benefits – Users sheet upgrade + calculators
// ────────────────────────────────────────────────────────────────────────────

const USER_BENEFIT_FIELDS = {
  TerminationDate: 'TerminationDate',
  ProbationMonths: 'ProbationMonths',
  ProbationEnd: 'ProbationEnd',
  InsuranceEligibleDate: 'InsuranceEligibleDate', // = ProbationEnd + 3 months
  InsuranceQualified: 'InsuranceQualified',       // boolean: today >= InsuranceEligibleDate and not Terminated/Inactive
  InsuranceEnrolled: 'InsuranceEnrolled',         // boolean: user signed up
  InsuranceCardReceivedDate: 'InsuranceCardReceivedDate'
};

/** Ensure benefit columns exist (appended) without disturbing existing Users headers */
function ensureUsersBenefitsColumns_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(USERS_SHEET) || ensureSheetWithHeaders(USERS_SHEET, USERS_HEADERS);
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return;

  const headerRange = sh.getRange(1, 1, 1, sh.getLastColumn());
  const headers = headerRange.getValues()[0].map(String);

  const missing = Object.values(USER_BENEFIT_FIELDS).filter(h => !headers.includes(h));
  if (missing.length === 0) return;

  // Append missing columns at the end
  const newHeaders = headers.concat(missing);
  sh.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]).setFontWeight('bold');
  sh.setFrozenRows(1);

  const rows = Math.max(0, sh.getLastRow() - 1);
  if (rows > 0) {
    // Set sensible defaults for booleans / numbers (without overwriting existing data)
    const colIndex = name => newHeaders.indexOf(name) + 1;
    const defaults = {
      [USER_BENEFIT_FIELDS.ProbationMonths]: 3,
      [USER_BENEFIT_FIELDS.InsuranceQualified]: false,
      [USER_BENEFIT_FIELDS.InsuranceEnrolled]: false
    };
    Object.entries(defaults).forEach(([key, defVal]) => {
      if (missing.includes(key)) {
        const rng = sh.getRange(2, colIndex(key), rows, 1);
        const arr = Array.from({ length: rows }, () => [defVal]);
        rng.setValues(arr);
      }
    });
  }
}

/** Date helpers */
function parseAsDate_(v) {
  if (!v && v !== 0) return null;
  if (v instanceof Date) return new Date(v);
  const s = String(v).trim();
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
function daysInMonth_(y, m) { return new Date(y, m + 1, 0).getDate(); }
function addMonths_(d, n) {
  if (!d) return null;
  const dt = new Date(d.getTime());
  const y = dt.getFullYear();
  const m = dt.getMonth();
  const targetM = m + n;
  const targetY = y + Math.floor(targetM / 12);
  const cleanM = (targetM % 12 + 12) % 12;
  const day = Math.min(dt.getDate(), daysInMonth_(targetY, cleanM));
  return new Date(targetY, cleanM, day);
}

/** Compute whether a user is qualified for insurance: today >= eligible && status not Inactive/Terminated */
function isInsuranceQualified_(employmentStatus, eligibleDate) {
  const bad = new Set(['Inactive', 'Terminated']);
  const okStatus = !bad.has(String(employmentStatus || '').trim());
  const ed = parseAsDate_(eligibleDate);
  return !!(okStatus && ed && (new Date()).getTime() >= ed.getTime());
}

/**
 * Recalculate benefits for all users.
 * opts: { force?: boolean } -> if true, overwrite ProbationEnd / InsuranceEligibleDate even if already set
 */
function recalcBenefitsForAllUsers(opts = {}) {
  try {
    ensureUsersBenefitsColumns_();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(USERS_SHEET);
    const data = sh.getDataRange().getValues();
    if (data.length <= 1) return { updated: 0 };

    const headers = data[0].map(String);
    const col = name => headers.indexOf(name);

    const iHire = col('HireDate');
    const iEmpStatus = col('EmploymentStatus');

    const iProbMonths = col(USER_BENEFIT_FIELDS.ProbationMonths);
    const iProbEnd = col(USER_BENEFIT_FIELDS.ProbationEnd);
    const iEligible = col(USER_BENEFIT_FIELDS.InsuranceEligibleDate);
    const iQualified = col(USER_BENEFIT_FIELDS.InsuranceQualified);

    if ([iHire, iEmpStatus, iProbMonths, iProbEnd, iEligible, iQualified].some(ix => ix < 0)) {
      return { updated: 0, warning: 'One or more benefit columns are missing.' };
    }

    let updated = 0;
    const force = !!opts.force;

    for (let r = 1; r < data.length; r++) {
      const row = data[r];

      // Source fields
      const hire = parseAsDate_(row[iHire]);
      let probMonths = Number(row[iProbMonths]);
      if (!Number.isFinite(probMonths) || probMonths <= 0) probMonths = 3;

      let probEnd = parseAsDate_(row[iProbEnd]);
      if (!probEnd || force) {
        probEnd = hire ? addMonths_(hire, probMonths) : null;
        if (probEnd) { row[iProbEnd] = probEnd; updated++; }
      }

      let eligible = parseAsDate_(row[iEligible]);
      const recomputedEligible = probEnd ? addMonths_(probEnd, 3) : null;
      if (!eligible || force) {
        eligible = recomputedEligible;
        if (eligible) { row[iEligible] = eligible; updated++; }
      }

      // Qualified boolean
      const qualified = isInsuranceQualified_(row[iEmpStatus], eligible);
      if (row[iQualified] !== qualified) {
        row[iQualified] = qualified;
        updated++;
      }

      data[r] = row;
    }

    // Write back only changed columns for safety
    const rows = data.length - 1;
    function writeCol(ix) {
      if (ix < 0) return;
      const rng = sh.getRange(2, ix + 1, rows, 1);
      const colVals = [];
      for (let r = 1; r < data.length; r++) colVals.push([data[r][ix]]);
      rng.setValues(colVals);
    }
    writeCol(iProbEnd);
    writeCol(iEligible);
    writeCol(iQualified);

    invalidateCache(USERS_SHEET);
    return { updated };
  } catch (e) {
    safeWriteError('recalcBenefitsForAllUsers', e);
    return { updated: 0, error: e.message };
  }
}

/** Recalc for a single user by ID (returns summary) */
function recalcBenefitsForUser(userId, opts = {}) {
  try {
    ensureUsersBenefitsColumns_();
    const users = readSheet(USERS_SHEET) || [];
    const user = users.find(u => u.ID === userId);
    if (!user) return { success: false, message: 'User not found' };

    // Pull live sheet to update in place
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(USERS_SHEET);
    const data = sh.getDataRange().getValues();
    const headers = data[0].map(String);
    const col = name => headers.indexOf(name);

    const iId = col('ID');
    const iHire = col('HireDate');
    const iEmpStatus = col('EmploymentStatus');
    const iProbMonths = col(USER_BENEFIT_FIELDS.ProbationMonths);
    const iProbEnd = col(USER_BENEFIT_FIELDS.ProbationEnd);
    const iEligible = col(USER_BENEFIT_FIELDS.InsuranceEligibleDate);
    const iQualified = col(USER_BENEFIT_FIELDS.InsuranceQualified);

    const rowIdx = data.findIndex((r, idx) => idx > 0 && String(r[iId]) === userId);
    if (rowIdx < 1) return { success: false, message: 'User row not found' };

    const row = data[rowIdx];
    const force = !!opts.force;

    const hire = parseAsDate_(row[iHire]);
    let probMonths = Number(row[iProbMonths]);
    if (!Number.isFinite(probMonths) || probMonths <= 0) probMonths = 3;

    let probEnd = parseAsDate_(row[iProbEnd]);
    if (!probEnd || force) probEnd = hire ? addMonths_(hire, probMonths) : null;

    let eligible = parseAsDate_(row[iEligible]);
    const recomputedEligible = probEnd ? addMonths_(probEnd, 3) : null;
    if (!eligible || force) eligible = recomputedEligible;

    const qualified = isInsuranceQualified_(row[iEmpStatus], eligible);

    // Persist back
    if (iProbEnd >= 0) sh.getRange(rowIdx + 1, iProbEnd + 1).setValue(probEnd || '');
    if (iEligible >= 0) sh.getRange(rowIdx + 1, iEligible + 1).setValue(eligible || '');
    if (iQualified >= 0) sh.getRange(rowIdx + 1, iQualified + 1).setValue(qualified);

    invalidateCache(USERS_SHEET);

    return {
      success: true,
      userId,
      probationMonths: probMonths,
      probationEnd: probEnd,
      insuranceEligibleDate: eligible,
      insuranceQualified: qualified
    };
  } catch (e) {
    safeWriteError('recalcBenefitsForUser', e);
    return { success: false, message: e.message };
  }
}

/** Mutators */
function setUserInsuranceEnrollment(userId, enrolled, cardReceivedDate = null) {
  try {
    ensureUsersBenefitsColumns_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(USERS_SHEET);
    const data = sh.getDataRange().getValues();
    const headers = data[0].map(String);
    const col = name => headers.indexOf(name);

    const iId = col('ID');
    const iEnrolled = col(USER_BENEFIT_FIELDS.InsuranceEnrolled);
    const iCard = col(USER_BENEFIT_FIELDS.InsuranceCardReceivedDate);

    const rowIdx = data.findIndex((r, idx) => idx > 0 && String(r[iId]) === userId);
    if (rowIdx < 1) return { success: false, message: 'User not found' };

    sh.getRange(rowIdx + 1, iEnrolled + 1).setValue(!!enrolled);
    if (iCard >= 0) {
      const dt = parseAsDate_(cardReceivedDate);
      sh.getRange(rowIdx + 1, iCard + 1).setValue(dt || '');
    }
    invalidateCache(USERS_SHEET);
    return { success: true };
  } catch (e) {
    safeWriteError('setUserInsuranceEnrollment', e);
    return { success: false, message: e.message };
  }
}

function setUserTermination(userId, terminationDate) {
  try {
    ensureUsersBenefitsColumns_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(USERS_SHEET);
    const data = sh.getDataRange().getValues();
    const headers = data[0].map(String);
    const col = name => headers.indexOf(name);

    const iId = col('ID');
    const iTerm = col(USER_BENEFIT_FIELDS.TerminationDate);
    const iStatus = col('EmploymentStatus');

    const rowIdx = data.findIndex((r, idx) => idx > 0 && String(r[iId]) === userId);
    if (rowIdx < 1) return { success: false, message: 'User not found' };

    sh.getRange(rowIdx + 1, iTerm + 1).setValue(parseAsDate_(terminationDate) || '');
    if (iStatus >= 0) sh.getRange(rowIdx + 1, iStatus + 1).setValue('Terminated');
    invalidateCache(USERS_SHEET);
    return { success: true };
  } catch (e) {
    safeWriteError('setUserTermination', e);
    return { success: false, message: e.message };
  }
}

function setUserProbation(userId, probationMonths, probationEndOverride = null) {
  try {
    ensureUsersBenefitsColumns_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(USERS_SHEET);
    const data = sh.getDataRange().getValues();
    const headers = data[0].map(String);
    const col = name => headers.indexOf(name);

    const iId = col('ID');
    const iProbMonths = col(USER_BENEFIT_FIELDS.ProbationMonths);
    const iProbEnd = col(USER_BENEFIT_FIELDS.ProbationEnd);

    const rowIdx = data.findIndex((r, idx) => idx > 0 && String(r[iId]) === userId);
    if (rowIdx < 1) return { success: false, message: 'User not found' };

    const pm = Number(probationMonths);
    if (Number.isFinite(pm) && pm > 0) sh.getRange(rowIdx + 1, iProbMonths + 1).setValue(pm);
    if (probationEndOverride) {
      sh.getRange(rowIdx + 1, iProbEnd + 1).setValue(parseAsDate_(probationEndOverride) || '');
    }
    invalidateCache(USERS_SHEET);
    // Recompute eligibility after change
    recalcBenefitsForUser(userId, { force: true });
    return { success: true };
  } catch (e) {
    safeWriteError('setUserProbation', e);
    return { success: false, message: e.message };
  }
}

/** Client wrappers */
function clientRecalcAllBenefits(force = false) {
  return recalcBenefitsForAllUsers({ force });
}
function clientRecalcUserBenefits(userId, force = false) {
  return recalcBenefitsForUser(userId, { force });
}
function clientSetInsuranceEnrollment(userId, enrolled, cardReceivedDate) {
  return setUserInsuranceEnrollment(userId, enrolled, cardReceivedDate);
}
function clientSetTermination(userId, terminationDate) {
  return setUserTermination(userId, terminationDate);
}
function clientSetProbation(userId, probationMonths, probationEndOverride) {
  return setUserProbation(userId, probationMonths, probationEndOverride);
}

// ────────────────────────────────────────────────────────────────────────────
// Canonical Manager–Users helpers (single source of truth; guarded)
// ────────────────────────────────────────────────────────────────────────────
if (typeof MANAGER_USERS_CANON_HEADERS === 'undefined')
  var MANAGER_USERS_CANON_HEADERS = ['ID', 'ManagerUserID', 'UserID', 'CampaignID', 'CreatedAt', 'UpdatedAt'];

if (typeof getManagerUsersSheetName_ !== 'function') {
  function getManagerUsersSheetName_() {
    return (typeof USER_MANAGERS_SHEET !== 'undefined') ? USER_MANAGERS_SHEET : 'UserManagers';
  }
}
if (typeof getManagerUsersHeaders_ !== 'function') {
  function getManagerUsersHeaders_() {
    return (typeof USER_MANAGERS_HEADERS !== 'undefined') ? USER_MANAGERS_HEADERS : MANAGER_USERS_CANON_HEADERS;
  }
}
if (typeof getOrCreateManagerUsersSheet_ !== 'function') {
  function getOrCreateManagerUsersSheet_() {
    const sh = ensureSheetWithHeaders(getManagerUsersSheetName_(), getManagerUsersHeaders_());

    // Auto-upgrade legacy headers: ManagedUserID -> UserID, add CampaignID if missing
    const hdrRange = sh.getRange(1, 1, 1, sh.getLastColumn());
    const hdrs = hdrRange.getValues()[0].map(String);
    let mutated = false;

    const idxManaged = hdrs.indexOf('ManagedUserID');
    const idxUser = hdrs.indexOf('UserID');
    if (idxManaged !== -1 && idxUser === -1) { hdrs[idxManaged] = 'UserID'; mutated = true; }
    if (!hdrs.includes('CampaignID')) { hdrs.splice(Math.min(hdrs.length, 3), 0, 'CampaignID'); mutated = true; }

    if (mutated) {
      sh.getRange(1, 1, 1, hdrs.length).setValues([hdrs]);
      sh.setFrozenRows(1);
    }
    return sh;
  }
}

function normalizeHeaderName_(value) {
  if (value === null || typeof value === 'undefined') return '';
  return String(value).trim();
}

function areHeadersAligned_(sheet, headers) {
  if (!sheet) return false;
  if (!headers || headers.length === 0) return false;
  if (sheet.getLastColumn() < headers.length) return false;
  const range = sheet.getRange(1, 1, 1, headers.length);
  const existing = range.getValues()[0].map(normalizeHeaderName_);
  return existing.length === headers.length && existing.every((h, i) => h === headers[i]);
}

function syncSheetColumnsAndHeaders_(sheet, headers) {
  if (!sheet) throw new Error('Sheet is required');
  if (!headers || headers.length === 0) throw new Error('Headers array is required');

  const headerCount = headers.length;
  const maxCols = sheet.getMaxColumns();
  if (maxCols < headerCount) {
    sheet.insertColumnsAfter(maxCols, headerCount - maxCols);
  }

  let lastCol = sheet.getLastColumn();
  if (lastCol < headerCount) {
    const missing = headerCount - lastCol;
    const anchor = Math.max(lastCol, 1);
    sheet.insertColumnsAfter(anchor, missing);
    lastCol = sheet.getLastColumn();
  }

  const rangeWidth = Math.max(sheet.getLastColumn(), headerCount);
  const initialHeaderRange = sheet.getRange(1, 1, 1, rangeWidth);
  const currentHeaders = initialHeaderRange.getValues()[0].map(normalizeHeaderName_);
  const maxRows = sheet.getMaxRows();
  let structureMutated = false;

  for (let i = 0; i < headerCount; i++) {
    const expected = headers[i];
    const current = currentHeaders[i] || '';
    if (current === expected) continue;

    const existingIdx = currentHeaders.indexOf(expected, i + 1);
    if (existingIdx !== -1) {
      const colRange = sheet.getRange(1, existingIdx + 1, maxRows, 1);
      sheet.moveColumns(colRange, i + 1);
      const [moved] = currentHeaders.splice(existingIdx, 1);
      currentHeaders.splice(i, 0, moved);
      structureMutated = true;
    } else if (current) {
      const targetCol = i + 1;
      const currentLast = Math.max(sheet.getLastColumn(), 1);
      if (targetCol > currentLast) {
        sheet.insertColumnsAfter(currentLast, 1);
      } else {
        sheet.insertColumnsBefore(targetCol, 1);
      }
      currentHeaders.splice(i, 0, '');
      structureMutated = true;
    }

    if (currentHeaders[i] !== expected) {
      currentHeaders[i] = expected;
    }
  }

  let finalLastCol = sheet.getLastColumn();
  let finalRange = sheet.getRange(1, 1, 1, finalLastCol);
  let finalRaw = finalRange.getValues()[0];
  let finalNormalized = finalRaw.map(normalizeHeaderName_);
  const allowedHeaders = new Set(headers);
  const seenHeaders = new Set();
  const columnsToDelete = [];

  for (let idx = 0; idx < finalNormalized.length; idx++) {
    const headerName = finalNormalized[idx];
    if (headerName) {
      if (!allowedHeaders.has(headerName)) {
        columnsToDelete.push(idx + 1);
        continue;
      }
      if (seenHeaders.has(headerName)) {
        columnsToDelete.push(idx + 1);
        continue;
      }
      seenHeaders.add(headerName);
    } else if (idx >= headerCount) {
      columnsToDelete.push(idx + 1);
    }
  }

  if (columnsToDelete.length) {
    for (let i = columnsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteColumn(columnsToDelete[i]);
    }
    structureMutated = true;
    finalLastCol = sheet.getLastColumn();
    finalRange = sheet.getRange(1, 1, 1, finalLastCol);
    finalRaw = finalRange.getValues()[0];
    finalNormalized = finalRaw.map(normalizeHeaderName_);
  }

  let headerMutated = false;

  for (let i = 0; i < headerCount; i++) {
    if (finalNormalized[i] !== headers[i]) {
      headerMutated = true;
      break;
    }
  }

  const finalHeaderRange = sheet.getRange(1, 1, 1, headerCount);
  if (headerMutated) {
    finalHeaderRange.setValues([headers.slice()]);
  }
  finalHeaderRange.setFontWeight('bold');
  sheet.setFrozenRows(1);

  return structureMutated || headerMutated;
}

if (typeof synchronizeSheetHeaders !== 'function') {
  function synchronizeSheetHeaders(sheetName, headers) {
    if (!sheetName) throw new Error('sheetName is required');
    if (!headers || !Array.isArray(headers) || headers.length === 0) {
      throw new Error('headers array is required');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return ensureSheetWithHeaders(sheetName, headers);
    }

    const mutated = syncSheetColumnsAndHeaders_(sheet, headers);
    if (mutated) {
      console.log(`Synchronized sheet structure for ${sheetName}`);
    }
    return sheet;
  }
}

(function (global) {
  if (global && typeof global.synchronizeSheetHeaders !== 'function' && typeof synchronizeSheetHeaders === 'function') {
    global.synchronizeSheetHeaders = synchronizeSheetHeaders;
  }
})(typeof globalThis !== 'undefined' ? globalThis : this);

function normalizeUserNameValue_(value) {
  if (value === null || typeof value === 'undefined') return '';
  return String(value).trim().toUpperCase();
}

function normalizeEmailValue_(value) {
  if (value === null || typeof value === 'undefined') return '';
  return String(value).trim().toLowerCase();
}

function isBlankCell_(value) {
  if (value === null || typeof value === 'undefined') return true;
  if (value instanceof Date) return isNaN(value.getTime());
  if (typeof value === 'string') return value.trim() === '';
  return false;
}

function rowHasMeaningfulData_(row) {
  if (!Array.isArray(row)) return false;
  for (let i = 0; i < row.length; i++) {
    if (!isBlankCell_(row[i])) return true;
  }
  return false;
}

function toSheetBooleanString_(value) {
  return value ? 'TRUE' : 'FALSE';
}

function generateUuid_() {
  try {
    if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.getUuid === 'function') {
      return Utilities.getUuid();
    }
  } catch (_) { }
  const rand = Math.floor(Math.random() * 1e12);
  return `uuid-${Date.now()}-${rand}`;
}

function backfillSheetDataIfPossible_(sheet) {
  try {
    if (!sheet) return false;
    const sheetName = normalizeHeaderName_(sheet.getName());
    const usersSheetName = normalizeHeaderName_(USERS_SHEET);
    if (sheetName === usersSheetName) {
      return backfillUsersSheetDerivedColumns_(sheet);
    }
  } catch (err) {
    console.warn('backfillSheetDataIfPossible_ failed', err);
  }
  return false;
}

function backfillUsersSheetDerivedColumns_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) {
    return false;
  }

  const headerValues = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const normalizedHeaders = headerValues.map(normalizeHeaderName_);

  const indexByHeader = {};
  for (let i = 0; i < normalizedHeaders.length; i++) {
    const header = normalizedHeaders[i];
    if (!header) continue;
    if (!Object.prototype.hasOwnProperty.call(indexByHeader, header)) {
      indexByHeader[header] = i;
    }
  }

  const idxId = Object.prototype.hasOwnProperty.call(indexByHeader, 'ID') ? indexByHeader.ID : -1;
  const idxUserName = Object.prototype.hasOwnProperty.call(indexByHeader, 'UserName') ? indexByHeader.UserName : -1;
  const idxFullName = Object.prototype.hasOwnProperty.call(indexByHeader, 'FullName') ? indexByHeader.FullName : -1;
  const idxEmail = Object.prototype.hasOwnProperty.call(indexByHeader, 'Email') ? indexByHeader.Email : -1;
  const idxNormalizedUserName = Object.prototype.hasOwnProperty.call(indexByHeader, 'NormalizedUserName') ? indexByHeader.NormalizedUserName : -1;
  const idxNormalizedEmail = Object.prototype.hasOwnProperty.call(indexByHeader, 'NormalizedEmail') ? indexByHeader.NormalizedEmail : -1;
  const idxCampaignId = Object.prototype.hasOwnProperty.call(indexByHeader, 'CampaignID') ? indexByHeader.CampaignID : -1;
  const idxSecurityStamp = Object.prototype.hasOwnProperty.call(indexByHeader, 'SecurityStamp') ? indexByHeader.SecurityStamp : -1;
  const idxConcurrencyStamp = Object.prototype.hasOwnProperty.call(indexByHeader, 'ConcurrencyStamp') ? indexByHeader.ConcurrencyStamp : -1;
  const idxCreatedAt = Object.prototype.hasOwnProperty.call(indexByHeader, 'CreatedAt') ? indexByHeader.CreatedAt : -1;
  const idxUpdatedAt = Object.prototype.hasOwnProperty.call(indexByHeader, 'UpdatedAt') ? indexByHeader.UpdatedAt : -1;
  const idxResetRequired = Object.prototype.hasOwnProperty.call(indexByHeader, 'ResetRequired') ? indexByHeader.ResetRequired : -1;
  const idxEmailConfirmed = Object.prototype.hasOwnProperty.call(indexByHeader, 'EmailConfirmed') ? indexByHeader.EmailConfirmed : -1;
  const idxPhoneConfirmed = Object.prototype.hasOwnProperty.call(indexByHeader, 'PhoneNumberConfirmed') ? indexByHeader.PhoneNumberConfirmed : -1;
  const idxLockoutEnabled = Object.prototype.hasOwnProperty.call(indexByHeader, 'LockoutEnabled') ? indexByHeader.LockoutEnabled : -1;
  const idxAccessFailedCount = Object.prototype.hasOwnProperty.call(indexByHeader, 'AccessFailedCount') ? indexByHeader.AccessFailedCount : -1;
  const idxTwoFactorEnabled = Object.prototype.hasOwnProperty.call(indexByHeader, 'TwoFactorEnabled') ? indexByHeader.TwoFactorEnabled : -1;
  const idxCanLogin = Object.prototype.hasOwnProperty.call(indexByHeader, 'CanLogin') ? indexByHeader.CanLogin : -1;
  const idxIsAdmin = Object.prototype.hasOwnProperty.call(indexByHeader, 'IsAdmin') ? indexByHeader.IsAdmin : -1;

  const hasAnyManagedColumn = [
    idxNormalizedUserName,
    idxNormalizedEmail,
    idxId,
    idxSecurityStamp,
    idxConcurrencyStamp,
    idxResetRequired,
    idxEmailConfirmed,
    idxPhoneConfirmed,
    idxLockoutEnabled,
    idxAccessFailedCount,
    idxTwoFactorEnabled,
    idxCanLogin,
    idxIsAdmin,
    idxCreatedAt,
    idxUpdatedAt
  ].some(idx => idx >= 0);

  if (!hasAnyManagedColumn) {
    return false;
  }

  const dataRange = sheet.getRange(2, 1, lastRow - 1, lastColumn);
  const values = dataRange.getValues();
  let mutated = false;

  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    const row = values[rowIndex];
    if (!rowHasMeaningfulData_(row)) {
      continue;
    }

    const rawUserName = idxUserName >= 0 ? String(row[idxUserName] || '').trim() : '';
    const rawEmail = idxEmail >= 0 ? String(row[idxEmail] || '').trim() : '';
    const rawFullName = idxFullName >= 0 ? String(row[idxFullName] || '').trim() : '';
    const rawCampaignId = idxCampaignId >= 0 ? String(row[idxCampaignId] || '').trim() : '';

    if (idxNormalizedUserName >= 0) {
      const candidate = normalizeUserNameValue_(rawUserName || rawEmail);
      const existingRaw = row[idxNormalizedUserName];
      const existingNormalized = normalizeUserNameValue_(existingRaw);
      if (candidate) {
        if (existingNormalized !== candidate || String(existingRaw || '').trim() !== candidate) {
          row[idxNormalizedUserName] = candidate;
          mutated = true;
        }
      } else if (existingNormalized && !rawUserName && !rawEmail) {
        row[idxNormalizedUserName] = '';
        mutated = true;
      }
    }

    if (idxNormalizedEmail >= 0) {
      const candidate = normalizeEmailValue_(rawEmail);
      const existingRaw = row[idxNormalizedEmail];
      const existingNormalized = normalizeEmailValue_(existingRaw);
      if (candidate) {
        if (existingNormalized !== candidate || String(existingRaw || '').trim() !== candidate) {
          row[idxNormalizedEmail] = candidate;
          mutated = true;
        }
      } else if (existingNormalized && !rawEmail) {
        row[idxNormalizedEmail] = '';
        mutated = true;
      }
    }

    const hasIdentityData = Boolean(
      (idxId >= 0 && !isBlankCell_(row[idxId])) ||
      rawUserName ||
      rawEmail ||
      rawFullName ||
      rawCampaignId
    );

    if (idxId >= 0 && isBlankCell_(row[idxId]) && hasIdentityData) {
      row[idxId] = generateUuid_();
      mutated = true;
    }

    if (idxSecurityStamp >= 0 && isBlankCell_(row[idxSecurityStamp]) && hasIdentityData) {
      row[idxSecurityStamp] = generateUuid_();
      mutated = true;
    }

    if (idxConcurrencyStamp >= 0 && isBlankCell_(row[idxConcurrencyStamp]) && hasIdentityData) {
      row[idxConcurrencyStamp] = generateUuid_();
      mutated = true;
    }

    const ensureBooleanDefault = (idx, defaultValue) => {
      if (idx < 0) return;
      if (!isBlankCell_(row[idx])) return;
      row[idx] = toSheetBooleanString_(defaultValue);
      mutated = true;
    };

    ensureBooleanDefault(idxResetRequired, false);
    ensureBooleanDefault(idxEmailConfirmed, false);
    ensureBooleanDefault(idxPhoneConfirmed, false);
    ensureBooleanDefault(idxLockoutEnabled, false);
    ensureBooleanDefault(idxTwoFactorEnabled, false);
    ensureBooleanDefault(idxCanLogin, true);
    ensureBooleanDefault(idxIsAdmin, false);

    if (idxAccessFailedCount >= 0 && isBlankCell_(row[idxAccessFailedCount])) {
      row[idxAccessFailedCount] = 0;
      mutated = true;
    }

    let rowTimestamp = null;
    const ensureTimestamp = idx => {
      if (idx < 0) return;
      if (!isBlankCell_(row[idx])) return;
      if (!rowTimestamp) rowTimestamp = new Date();
      row[idx] = rowTimestamp;
      mutated = true;
    };

    if (idxCreatedAt >= 0 && isBlankCell_(row[idxCreatedAt])) {
      if (idxUpdatedAt >= 0 && !isBlankCell_(row[idxUpdatedAt])) {
        row[idxCreatedAt] = row[idxUpdatedAt];
        mutated = true;
      } else {
        ensureTimestamp(idxCreatedAt);
      }
    }

    if (idxUpdatedAt >= 0 && isBlankCell_(row[idxUpdatedAt])) {
      if (idxCreatedAt >= 0 && !isBlankCell_(row[idxCreatedAt])) {
        row[idxUpdatedAt] = row[idxCreatedAt];
        mutated = true;
      } else {
        ensureTimestamp(idxUpdatedAt);
      }
    }
  }

  if (mutated) {
    dataRange.setValues(values);
  }

  return mutated;
}
if (typeof readManagerAssignments_ !== 'function') {
  function readManagerAssignments_() {
    const name = getManagerUsersSheetName_();
    const rows = readSheet(name) || [];
    return rows.map(r => ({
      ManagerUserID: String(r.ManagerUserID || r.ManagerID || r.ManagerId || r.ManagerUserId || '').trim(),
      UserID: String(r.UserID || r.ManagedUserID || r.UserId || r.ManagedUserId || '').trim(),
      CampaignID: String(r.CampaignID || '').trim()
    })).filter(x => x.ManagerUserID && x.UserID);
  }
}

// ────────────────────────────────────────────────────────────────────────────
// Setup: Create/Repair sheet with headers (idempotent w/ caching & retry)
// ────────────────────────────────────────────────────────────────────────────
const __ensureSheetWithHeaders = (function () {
  function ensureSheetWithHeadersImpl(name, headers) {
    const MAX_RETRIES = 3;
    const BASE_DELAY_MS = 1000;
    const recursionGuardKey = `RECURSION_GUARD_${name}`;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    function sleep(ms) { Utilities.sleep(ms); }

    try {
      const recursionGuard = PropertiesService.getScriptProperties().getProperty(recursionGuardKey);
      if (recursionGuard === 'active') {
        console.log(`Recursion detected for sheet ${name}, skipping ensureSheetWithHeaders`);
        return ss.getSheetByName(name);
      }
      PropertiesService.getScriptProperties().setProperty(recursionGuardKey, 'active');

      if (!headers || !Array.isArray(headers) || headers.some(h => !h || typeof h !== 'string')) {
        throw new Error('Invalid or empty headers provided');
      }
      const uniqueHeaders = new Set(headers);
      if (uniqueHeaders.size !== headers.length) {
        throw new Error('Duplicate headers detected');
      }

      const cacheKey = `SHEET_EXISTS_${name}`;
      const cached = scriptCache.get(cacheKey);
      let sh = ss.getSheetByName(name);

      if (sh && areHeadersAligned_(sh, headers)) {
        if (typeof registerTableSchema === 'function') {
          try { registerTableSchema(name, { headers }); } catch (regErr) { console.warn(`registerTableSchema(${name}) failed`, regErr); }
        }
        return sh;
      }

      if (cached === 'true' && sh) {
        console.log(`Cache hit: Sheet ${name} exists`);
      }

      let lastError = null;
      for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
        try {
          sh = ss.getSheetByName(name);
          if (!sh) {
            sh = ss.insertSheet(name);
            console.log(`Created sheet ${name}`);
          }

          const mutated = syncSheetColumnsAndHeaders_(sh, headers);
          if (mutated) {
            console.log(`Synchronized headers for ${name}`);
          }

          const backfilled = backfillSheetDataIfPossible_(sh);
          if (backfilled) {
            console.log(`Backfilled derived data for ${name}`);
          }

          scriptCache.put(cacheKey, 'true', CACHE_TTL_SEC);

          if (typeof registerTableSchema === 'function') {
            try { registerTableSchema(name, { headers }); } catch (regErr) { console.warn(`registerTableSchema(${name}) failed`, regErr); }
          }

          return sh;
        } catch (e) {
          lastError = e;
          if (e.message.includes('timed out') && attempt < MAX_RETRIES) {
            const delay = BASE_DELAY_MS * Math.pow(2, attempt - 1);
            console.log(`Attempt ${attempt} failed for ${name}: ${e.message}. Retrying after ${delay}ms`);
            sleep(delay);
            continue;
          }
          throw e;
        }
      }
      throw lastError || new Error(`Failed to ensure sheet ${name} after ${MAX_RETRIES} attempts`);
    } catch (e) {
      console.error(`ensureSheetWithHeaders(${name}) failed: ${e.message}, Document ID: ${ss.getId()}`);
      throw e;
    } finally {
      PropertiesService.getScriptProperties().deleteProperty(recursionGuardKey);
    }
  }

  return ensureSheetWithHeadersImpl;
})();

(function (global) {
  const previous = (typeof global.ensureSheetWithHeaders === 'function') ? global.ensureSheetWithHeaders.bind(global) : null;
  global.ensureSheetWithHeaders = function (name, headers) {
    try {
      return __ensureSheetWithHeaders(name, headers);
    } catch (err) {
      if (previous) {
        try { return previous(name, headers); } catch (fallbackErr) { safeWriteError && safeWriteError('ensureSheetWithHeadersFallback', fallbackErr); }
      }
      throw err;
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);

// ────────────────────────────────────────────────────────────────────────────
// Error / Debug logging
// ────────────────────────────────────────────────────────────────────────────
if (typeof writeError !== 'function') {
  function writeError(context, error) {
    try {
      const errorMsg = error && error.message ? error.message : String(error);
      console.error(`${context}: ${errorMsg}`);
      if (error && error.stack) console.error(error.stack);
      const sh = ensureSheetWithHeaders(ERROR_LOGS_SHEET, ERROR_LOGS_HEADERS);
      sh.appendRow([new Date(), errorMsg, error && error.stack ? error.stack : '']);
      invalidateCache(ERROR_LOGS_SHEET);
    } catch (_) { }
  }
}
function writeDebug(msg) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(DEBUG_LOGS_SHEET);
    if (!sh) {
      sh = ss.insertSheet(DEBUG_LOGS_SHEET);
      sh.getRange(1, 1, 1, DEBUG_LOGS_HEADERS.length).setValues([DEBUG_LOGS_HEADERS]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), String(msg)]);
    invalidateCache(DEBUG_LOGS_SHEET);
  } catch (e) { console.log(`writeDebug failed: ${e.message}`); }
}
function safeWriteError(context, error) {
  try {
    const msg = error && error.message ? error.message : String(error);
    const stack = error && error.stack ? error.stack : '';
    writeError(context, msg);
    if (stack) writeError(`${context} Stack`, stack);
  } catch (writeErr) {
    console.error('safeWriteError failed:', { context, writeErr: writeErr && writeErr.message });
  }
}

// ────────────────────────────────────────────────────────────────────────────
// Sheet read w/ caching
// ────────────────────────────────────────────────────────────────────────────
const __readSheet = (function () {
  function readSheetImpl(sheetName, optionsOrCache) {
    const { useCache, allowScriptCache, queryOptions, tenantContext: explicitTenantContext, campaignId: requestedCampaignId, userId: requestedUserId, suppressTenantContext } = _normalizeReadSheetOptions_(optionsOrCache);
    const cacheKey = allowScriptCache ? `DATA_${sheetName}` : null;

    if (allowScriptCache && cacheKey) {
      try {
        const cached = scriptCache.get(cacheKey);
        if (cached) return JSON.parse(cached);
      } catch (cacheErr) {
        console.warn(`Failed to read cache for ${sheetName}: ${cacheErr.message}`);
      }
    }

    let data = null;
    let usedDatabaseManager = false;

    var schemaRegistry;
    try {
      schemaRegistry = typeof __DB_SCHEMA_REGISTRY__ !== 'undefined'
        ? __DB_SCHEMA_REGISTRY__
        : (typeof globalThis !== 'undefined' ? globalThis.__DB_SCHEMA_REGISTRY__ : null);
    } catch (_) { schemaRegistry = null; }
    var schema = schemaRegistry ? schemaRegistry[sheetName] : null;

    const guardActive = (function () {
      try {
        const globalObj = (typeof globalThis !== 'undefined') ? globalThis : this;
        return Boolean(globalObj && Number(globalObj.__READ_SHEET_SUPPRESS_TENANT_CONTEXT__ || 0) > 0);
      } catch (_) {
        return false;
      }
    })();

    var tenantContext = explicitTenantContext || null;
    var currentUser = null;
    const skipTenantContext = suppressTenantContext || guardActive;
    if (!tenantContext && schema && schema.tenantColumn && !skipTenantContext) {
      var targetCampaignId = requestedCampaignId || null;
      var targetUserId = requestedUserId || null;
      if (!targetCampaignId && queryOptions && queryOptions.where) {
        targetCampaignId = queryOptions.where.CampaignID || queryOptions.where.CampaignId || queryOptions.where.campaignId || null;
      }
      if (typeof getCurrentUser === 'function') {
        try { currentUser = getCurrentUser(); } catch (ctxErr) { console.warn('readSheet: getCurrentUser failed', ctxErr); }
      }
      if (!targetUserId && currentUser && currentUser.ID) {
        targetUserId = currentUser.ID;
      }
      if (typeof TenantSecurity !== 'undefined' && TenantSecurity && targetUserId) {
        try {
          tenantContext = TenantSecurity.getTenantContext(targetUserId, targetCampaignId).context;
        } catch (tenantErr) {
          console.warn(`readSheet(${sheetName}): tenant context error`, tenantErr);
        }
      }
    }

    if (typeof dbSelect === 'function') {
      try {
        const query = Object.assign({}, queryOptions);
        if (!useCache) query.cache = false;
        data = dbSelect(sheetName, query, tenantContext);
        if (Array.isArray(data)) usedDatabaseManager = true;
      } catch (dbErr) {
        safeWriteError && safeWriteError(`readSheet(${sheetName})`, dbErr);
        data = null;
      }
    }

    if (!Array.isArray(data)) {
      data = _legacyReadSheet_(sheetName);
    }

    if (allowScriptCache && cacheKey && Array.isArray(data)) {
      try {
        scriptCache.put(cacheKey, JSON.stringify(data), CACHE_TTL_SEC);
      } catch (cachePutErr) {
        console.warn(`Failed to cache ${sheetName}: ${cachePutErr.message}`);
      }
    }

    if (!usedDatabaseManager && typeof queryOptions === 'object' && Object.keys(queryOptions).length) {
      data = _applyQueryOptions_(data, queryOptions);
    }

    return Array.isArray(data) ? data : [];
  }

  return readSheetImpl;
})();

(function (global) {
  const previousRead = (typeof global.readSheet === 'function') ? global.readSheet.bind(global) : null;
  global.readSheet = function (sheetName, optionsOrCache) {
    try {
      return __readSheet(sheetName, optionsOrCache);
    } catch (err) {
      if (previousRead) {
        try { return previousRead(sheetName, optionsOrCache); } catch (fallbackErr) { safeWriteError && safeWriteError('readSheetFallback', fallbackErr); }
      }
      throw err;
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);

const __invalidateCache = (function () {
  function invalidateCacheImpl(sheetName) {
    try { scriptCache.remove(`DATA_${sheetName}`); } catch (e) { console.error('invalidateCache failed:', e); }
    if (typeof DatabaseManager !== 'undefined' && DatabaseManager && typeof DatabaseManager.dropTableCache === 'function') {
      try { DatabaseManager.dropTableCache(sheetName); } catch (err) { console.error('DatabaseManager cache drop failed:', err); }
    }
  }

  return invalidateCacheImpl;
})();

(function (global) {
  const previousInvalidate = (typeof global.invalidateCache === 'function') ? global.invalidateCache.bind(global) : null;
  global.invalidateCache = function (sheetName) {
    __invalidateCache(sheetName);
    if (previousInvalidate && previousInvalidate !== global.invalidateCache) {
      try { previousInvalidate(sheetName); } catch (fallbackErr) { safeWriteError && safeWriteError('invalidateCacheFallback', fallbackErr); }
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);

function _normalizeReadSheetOptions_(optionsOrCache) {
  let options = {};
  let useCache = true;
  if (typeof optionsOrCache === 'boolean') {
    useCache = optionsOrCache;
  } else if (optionsOrCache && typeof optionsOrCache === 'object') {
    options = optionsOrCache;
    if (Object.prototype.hasOwnProperty.call(options, 'useCache')) {
      useCache = options.useCache !== false;
    } else if (Object.prototype.hasOwnProperty.call(options, 'cache')) {
      useCache = options.cache !== false;
    }
  }

  const queryKeys = ['where', 'filter', 'map', 'sortBy', 'sortDesc', 'offset', 'limit', 'columns'];
  const queryOptions = {};
  queryKeys.forEach(key => {
    if (typeof options[key] !== 'undefined') queryOptions[key] = options[key];
  });

  const hasFunctions = typeof queryOptions.filter === 'function' || typeof queryOptions.map === 'function';
  const allowScriptCache = useCache && !hasFunctions && Object.keys(queryOptions).length === 0;

  let tenantContext = null;
  if (options && typeof options === 'object') {
    if (typeof options.tenantContext !== 'undefined') {
      if (typeof options.tenantContext === 'string' || typeof options.tenantContext === 'number') {
        tenantContext = { tenantId: String(options.tenantContext) };
      } else if (options.tenantContext && typeof options.tenantContext === 'object') {
        tenantContext = Object.assign({}, options.tenantContext);
      }
    } else if (typeof options.tenantId !== 'undefined') {
      tenantContext = { tenantId: String(options.tenantId) };
    } else if (typeof options.campaignId !== 'undefined') {
      tenantContext = { tenantId: String(options.campaignId) };
    }
  }

  const campaignId = options && typeof options === 'object' && typeof options.campaignId !== 'undefined'
    ? String(options.campaignId)
    : null;
  const userId = options && typeof options === 'object' && typeof options.userId !== 'undefined'
    ? String(options.userId)
    : null;

  const suppressTenantContext = Boolean(options && (options.suppressTenantContext || options.skipTenantContext));

  return { useCache, allowScriptCache, queryOptions, tenantContext, campaignId, userId, suppressTenantContext };
}

function _legacyReadSheet_(sheetName) {
  try {
    const isCampaignSheet = sheetName === CAMPAIGNS_SHEET;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(sheetName);
    if (!sh) { safeWriteError && safeWriteError(`readSheet(${sheetName})`, `Sheet ${sheetName} not found`); return []; }

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return [];

    const vals = sh.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = vals.shift().map(h => String(h).trim() || null);
    if (headers.some(h => !h)) return [];
    const uniqueHeaders = new Set(headers);
    if (uniqueHeaders.size !== headers.length) return [];

    return vals.map(row => {
      const obj = {};
      let hasData = row.some(v => v !== '' && v != null);
      if (isCampaignSheet) {
        const idIndex = headers.indexOf('ID');
        const nameIndex = headers.indexOf('Name');
        hasData = row[idIndex] && row[nameIndex];
      }
      if (!hasData) return null;

      row.forEach((v, i) => {
        if (!headers[i]) return;
        obj[headers[i]] = (isCampaignSheet && headers[i].includes('At') && v && !isNaN(new Date(v))) ? new Date(v).toISOString() : v;
      });
      return obj;
    }).filter(Boolean);
  } catch (e) {
    safeWriteError && safeWriteError(`readSheet(${sheetName})`, e);
    return [];
  }
}

function _applyQueryOptions_(rows, options) {
  if (!Array.isArray(rows) || !options) return Array.isArray(rows) ? rows : [];
  let filtered = rows.slice();

  if (options.where && typeof options.where === 'object') {
    filtered = filtered.filter(row => _matchesWhere_(row, options.where));
  }

  if (options.filter && typeof options.filter === 'function') {
    filtered = filtered.filter(options.filter);
  }

  if (options.map && typeof options.map === 'function') {
    filtered = filtered.map(options.map);
  }

  if (options.sortBy) {
    const key = options.sortBy;
    const desc = !!options.sortDesc;
    filtered.sort((a, b) => {
      const av = a[key];
      const bv = b[key];
      if (av === bv) return 0;
      if (av === undefined || av === null || av === '') return desc ? 1 : -1;
      if (bv === undefined || bv === null || bv === '') return desc ? -1 : 1;
      if (av > bv) return desc ? -1 : 1;
      if (av < bv) return desc ? 1 : -1;
      return 0;
    });
  }

  const offset = options.offset || 0;
  const limit = typeof options.limit === 'number' ? options.limit : null;
  if (offset || limit !== null) {
    const start = offset;
    const end = limit !== null ? offset + limit : filtered.length;
    filtered = filtered.slice(start, end);
  }

  if (options.columns && Array.isArray(options.columns) && options.columns.length) {
    filtered = filtered.map(row => {
      const projected = {};
      options.columns.forEach(col => { projected[col] = row[col]; });
      return projected;
    });
  }

  return filtered;
}

function _matchesWhere_(row, where) {
  const keys = Object.keys(where);
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    const expected = where[key];
    const actual = row[key];
    if (expected instanceof RegExp) {
      if (!expected.test(String(actual || ''))) return false;
    } else if (expected && typeof expected === 'object' && !Array.isArray(expected)) {
      if (!_evaluateWhereOperator_(actual, expected)) return false;
    } else if (String(actual) !== String(expected)) {
      return false;
    }
  }
  return true;
}

function _evaluateWhereOperator_(actual, expression) {
  const ops = Object.keys(expression);
  for (let i = 0; i < ops.length; i++) {
    const op = ops[i];
    const value = expression[op];
    switch (op) {
      case '$gt':
        if (!(actual > value)) return false;
        break;
      case '$gte':
        if (!(actual >= value)) return false;
        break;
      case '$lt':
        if (!(actual < value)) return false;
        break;
      case '$lte':
        if (!(actual <= value)) return false;
        break;
      case '$ne':
        if (actual === value) return false;
        break;
      case '$in':
        if (!Array.isArray(value) || value.indexOf(actual) === -1) return false;
        break;
      case '$nin':
        if (Array.isArray(value) && value.indexOf(actual) !== -1) return false;
        break;
      default:
        if (String(actual) !== String(expression[op])) return false;
    }
  }
  return true;
}

// ────────────────────────────────────────────────────────────────────────────
// Employment helpers (basic validation)
// ────────────────────────────────────────────────────────────────────────────
function getValidEmploymentStatuses() {
  return ['Active', 'Inactive', 'Terminated', 'On Leave', 'Probation', 'Contract', 'Part Time', 'Full Time', 'Suspended'];
}
function validateEmploymentStatus(status) {
  if (!status) return true;
  return getValidEmploymentStatuses().includes(String(status).trim());
}
function validateHireDate(dateStr) {
  if (!dateStr) return true;
  try { const d = new Date(dateStr); return !isNaN(d.getTime()) && d <= new Date(); } catch (_) { return false; }
}
function validateCountry(country) {
  if (!country) return true;
  const s = String(country).trim(); return s.length >= 2 && s.length <= 100;
}

// ────────────────────────────────────────────────────────────────────────────
// Multi-campaign: special campaign row
// ────────────────────────────────────────────────────────────────────────────
function getOrCreateMultiCampaignId() {
  try {
    const sheet = ensureSheetWithHeaders(CAMPAIGNS_SHEET, CAMPAIGNS_HEADERS);
    const rows = readSheet(CAMPAIGNS_SHEET) || [];
    const row = rows.find(c => String(c.Name || '').trim().toLowerCase() === MULTI_CAMPAIGN_NAME.toLowerCase());
    if (row) return row.ID;

    const id = Utilities.getUuid(), now = new Date();
    sheet.appendRow([id, MULTI_CAMPAIGN_NAME, 'System-wide navigation container', now, now]);
    invalidateCache(CAMPAIGNS_SHEET);
    return id;
  } catch (e) {
    safeWriteError('getOrCreateMultiCampaignId', e);
    return null;
  }
}
function invalidateNavigationCache(campaignId) {
  if (!campaignId) return;
  ['NAVIGATION_', 'CAMPAIGN_PAGES_', 'PAGE_CATEGORIES_'].forEach(prefix => {
    try { scriptCache.remove(`${prefix}${campaignId}`); } catch (_) { }
  });
}

// ────────────────────────────────────────────────────────────────────────────
// Enhanced Page Discovery: Define all pages reflected by routing
// ────────────────────────────────────────────────────────────────────────────
function getAllPagesFromActualRouting() {
  const discoveredPages = [
    // DASHBOARD & ANALYTICS
    { key: 'dashboard', title: 'Dashboard', icon: 'fas fa-tachometer-alt', description: 'Main dashboard with OKR metrics and analytics', isSystem: true, requiresAdmin: false, category: 'Dashboard & Analytics', isMainPage: true },
    { key: 'agent-experience', title: 'Agent Experience', icon: 'fas fa-user-headset', description: 'Personalized workspace for frontline specialists with schedules, tasks, and communications', isSystem: true, requiresAdmin: false, category: 'Dashboard & Analytics', isMainPage: true },
    { key: 'goalsetting', title: 'Goal Setting', icon: 'fas fa-bullseye', description: 'Create, track, and update OKRs and performance targets', isSystem: true, requiresAdmin: false, category: 'Dashboard & Analytics' },

    // LEADERSHIP & EXECUTIVE HUBS
    { key: 'manager-executive-experience', title: 'Manager & Executive Experience', icon: 'fas fa-briefcase', description: 'Leadership analytics, coaching metrics, and cross-campaign oversight', isSystem: true, requiresAdmin: true, category: 'Dashboard & Analytics' },

    // QUALITY ASSURANCE
    { key: 'qualityform', title: 'Quality Form', icon: 'fas fa-clipboard-check', description: 'Quality assurance evaluation form (campaign-aware routing)', isSystem: true, requiresAdmin: false, category: 'Quality Assurance' },
    { key: 'ibtrqualityreports', title: 'QA Dashboard', icon: 'fas fa-chart-pie', description: 'Quality assurance dashboard and reports (campaign-aware routing)', isSystem: true, requiresAdmin: false, category: 'Quality Assurance' },
    { key: 'qualityview', title: 'Quality View', icon: 'fas fa-eye', description: 'View individual quality assurance records', isSystem: true, requiresAdmin: false, category: 'Quality Assurance' },
    { key: 'qualitylist', title: 'Quality List', icon: 'fas fa-list-check', description: 'List all quality assurance records', isSystem: true, requiresAdmin: false, category: 'Quality Assurance' },
    { key: 'unifiedqadashboard', title: 'Unified QA Dashboard', icon: 'fas fa-layer-group', description: 'Cross-campaign QA dashboard for consolidated quality analytics', isSystem: true, requiresAdmin: false, category: 'Quality Assurance' },
    { key: 'qacollablist', title: 'QA Collaboration List', icon: 'fas fa-users-gear', description: 'Collaborative queue of QA audits and action items', isSystem: true, requiresAdmin: false, category: 'Quality Assurance' },
    { key: 'qacollabview', title: 'QA Collaboration View', icon: 'fas fa-users-viewfinder', description: 'Detailed view of collaborative QA audit items', isSystem: true, requiresAdmin: false, category: 'Quality Assurance' },
    { key: 'qualitycollabform', title: 'QA Collaboration Form', icon: 'fas fa-clipboard-list', description: 'Submit or update collaborative QA action plans', isSystem: true, requiresAdmin: false, category: 'Quality Assurance' },

    // CAMPAIGN-SPECIFIC QA
    { key: 'independencequality', title: 'Independence Insurance QA', icon: 'fas fa-shield-alt', description: 'Quality assurance form for Independence Insurance campaign', isSystem: true, requiresAdmin: false, category: 'Campaign QA' },
    { key: 'independenceqadashboard', title: 'Independence QA Dashboard', icon: 'fas fa-chart-line', description: 'Unified QA dashboard for Independence Insurance', isSystem: true, requiresAdmin: false, category: 'Campaign QA' },
    { key: 'creditsuiteqa', title: 'Credit Suite QA', icon: 'fas fa-credit-card', description: 'Quality assurance form for Credit Suite campaign', isSystem: true, requiresAdmin: false, category: 'Campaign QA' },
    { key: 'qa-dashboard', title: 'Credit Suite QA Dashboard', icon: 'fas fa-chart-area', description: 'QA dashboard for Credit Suite campaign', isSystem: true, requiresAdmin: false, category: 'Campaign QA' },
    { key: 'groundingqaform', title: 'Grounding QA Form', icon: 'fas fa-seedling', description: 'Campaign-specific QA form for Grounding support teams', isSystem: true, requiresAdmin: false, category: 'Campaign QA' },

    // REPORTING
    { key: 'callreports', title: 'Call Reports', icon: 'fas fa-phone-volume', description: 'Call analytics and reporting dashboard with CSV export', isSystem: true, requiresAdmin: false, category: 'Reporting & Analytics' },
    { key: 'attendancereports', title: 'Attendance Reports', icon: 'fas fa-chart-bar', description: 'Attendance analytics and reports with CSV export', isSystem: true, requiresAdmin: false, category: 'Reporting & Analytics' },
    { key: 'collaboration-reporting', title: 'Collaboration Reporting', icon: 'fas fa-people-arrows', description: 'Collaboration analytics, sentiment tracking, and engagement insights', isSystem: true, requiresAdmin: false, category: 'Reporting & Analytics' },

    // COACHING
    { key: 'coachingdashboard', title: 'Coaching Dashboard', icon: 'fas fa-chalkboard-teacher', description: 'Coaching metrics and management dashboard', isSystem: true, requiresAdmin: false, category: 'Coaching & Development' },
    { key: 'coachingview', title: 'Coaching View', icon: 'fas fa-search-plus', description: 'View individual coaching session details', isSystem: true, requiresAdmin: false, category: 'Coaching & Development' },
    { key: 'coachinglist', title: 'Coaching List', icon: 'fas fa-list-ul', description: 'List of all coaching sessions', isSystem: true, requiresAdmin: false, category: 'Coaching & Development' },
    { key: 'coachings', title: 'Coachings', icon: 'fas fa-users', description: 'Alternative route for coaching sessions management', isSystem: true, requiresAdmin: false, category: 'Coaching & Development' },
    { key: 'coachingsheet', title: 'Coaching Form', icon: 'fas fa-file-signature', description: 'Create and edit coaching session forms', isSystem: true, requiresAdmin: false, category: 'Coaching & Development' },
    { key: 'coaching', title: 'Coaching Interface', icon: 'fas fa-graduation-cap', description: 'Alternative coaching interface', isSystem: true, requiresAdmin: false, category: 'Coaching & Development' },

    // TASKS
    { key: 'tasksmanager', title: 'Task Board', icon: 'fas fa-columns', description: 'Kanban-style task board management', isSystem: true, requiresAdmin: false, category: 'Task Management' },
    { key: 'taskform', title: 'Task Form', icon: 'fas fa-plus-square', description: 'Create and edit individual tasks', isSystem: true, requiresAdmin: false, category: 'Task Management' },

    // SCHEDULING
    { key: 'schedulemanagement', title: 'Schedule Management', icon: 'fas fa-calendar-week', description: 'Manage work schedules and shifts', isSystem: true, requiresAdmin: false, category: 'Scheduling & Time' },
    { key: 'attendancecalendar', title: 'Attendance Calendar', icon: 'fas fa-calendar-check', description: 'Calendar view for attendance tracking', isSystem: true, requiresAdmin: false, category: 'Scheduling & Time' },
    { key: 'slotmanagement', title: 'Shift Slot Management', icon: 'fas fa-clock', description: 'Manage shift slots and time allocations', isSystem: true, requiresAdmin: false, category: 'Scheduling & Time' },
    { key: 'agent-schedule', title: 'Agent Schedule', icon: 'fas fa-calendar-day', description: 'Personal schedule and shift details for agents', isSystem: true, requiresAdmin: false, category: 'Scheduling & Time' },

    // WORKFLOW
    { key: 'escalations', title: 'Escalations', icon: 'fas fa-exclamation-triangle', description: 'Issue escalation management and tracking', isSystem: true, requiresAdmin: false, category: 'Workflow & Operations' },
    { key: 'eodreport', title: 'EOD Report', icon: 'fas fa-clipboard-check', description: 'End of day reporting and task completion', isSystem: true, requiresAdmin: false, category: 'Workflow & Operations' },
    { key: 'incentives', title: 'Incentives', icon: 'fas fa-trophy', description: 'Employee incentives and rewards program', isSystem: true, requiresAdmin: false, category: 'Workflow & Operations' },

    // COMMUNICATION
    { key: 'search', title: 'Web Search', icon: 'fas fa-search', description: 'Global web search functionality', isSystem: true, requiresAdmin: false, category: 'Communication' },
    { key: 'chat', title: 'Team Chat', icon: 'fas fa-comments', description: 'Team communication and messaging system', isSystem: true, requiresAdmin: false, category: 'Communication' },
    { key: 'bookmarks', title: 'Bookmarks', icon: 'fas fa-bookmark', description: 'Personal bookmark manager', isSystem: true, requiresAdmin: false, category: 'Communication' },
    { key: 'notifications', title: 'Notifications', icon: 'fas fa-bell', description: 'System notifications and alerts', isSystem: true, requiresAdmin: false, category: 'Communication' },

    // ADMINISTRATION
    { key: 'manageuser', title: 'User Management', icon: 'fas fa-users-cog', description: 'Manage system users and permissions', isSystem: true, requiresAdmin: true, category: 'Administration' },
    { key: 'manageroles', title: 'Role Management', icon: 'fas fa-user-shield', description: 'Manage user roles and permissions', isSystem: true, requiresAdmin: true, category: 'Administration' },
    { key: 'managecampaign', title: 'Campaign Management', icon: 'fas fa-bullhorn', description: 'Manage campaigns and their configurations', isSystem: true, requiresAdmin: true, category: 'Administration' },
    { key: 'settings', title: 'Settings', icon: 'fas fa-cogs', description: 'System configuration and settings', isSystem: true, requiresAdmin: false, category: 'Administration' },

    // DATA MGMT
    { key: 'import', title: 'Data Import', icon: 'fas fa-file-import', description: 'Import data from CSV and other sources', isSystem: true, requiresAdmin: true, category: 'Data Management' },
    { key: 'importattendance', title: 'Import Attendance', icon: 'fas fa-file-upload', description: 'Import attendance data from external sources', isSystem: true, requiresAdmin: true, category: 'Data Management' },

    // UTILITIES
    { key: 'ackform', title: 'Acknowledgment Form', icon: 'fas fa-signature', description: 'Employee acknowledgment and signature forms', isSystem: true, requiresAdmin: false, category: 'Forms & Utilities' },
    { key: 'proxy', title: 'Proxy Service', icon: 'fas fa-exchange-alt', description: 'Proxy service for external content access', isSystem: true, requiresAdmin: false, category: 'Forms & Utilities' },

    // AUTH
    { key: 'setpassword', title: 'Set Password', icon: 'fas fa-key', description: 'Set new password for user account', isSystem: true, requiresAdmin: false, category: 'Authentication', isPublic: true },
    { key: 'resetpassword', title: 'Reset Password', icon: 'fas fa-unlock-alt', description: 'Reset forgotten password', isSystem: true, requiresAdmin: false, category: 'Authentication', isPublic: true },
    { key: 'resend-verification', title: 'Resend Verification', icon: 'fas fa-envelope-circle-check', description: 'Resend email verification link', isSystem: true, requiresAdmin: false, category: 'Authentication', isPublic: true },
    { key: 'resendverification', title: 'Resend Verification', icon: 'fas fa-envelope-circle-check', description: 'Alternative route for resending verification', isSystem: true, requiresAdmin: false, category: 'Authentication', isPublic: true },
    { key: 'forgotpassword', title: 'Forgot Password', icon: 'fas fa-question-circle', description: 'Initiate password reset process', isSystem: true, requiresAdmin: false, category: 'Authentication', isPublic: true },
    { key: 'forgot-password', title: 'Forgot Password', icon: 'fas fa-question-circle', description: 'Alternative route for password reset', isSystem: true, requiresAdmin: false, category: 'Authentication', isPublic: true },
    { key: 'emailconfirmed', title: 'Email Confirmed', icon: 'fas fa-check-circle', description: 'Email confirmation success page', isSystem: true, requiresAdmin: false, category: 'Authentication', isPublic: true },
    { key: 'email-confirmed', title: 'Email Confirmed', icon: 'fas fa-check-circle', description: 'Alternative route for email confirmation', isSystem: true, requiresAdmin: false, category: 'Authentication', isPublic: true }
  ];
  return discoveredPages;
}

// Category definitions
function getEnhancedPageCategories() {
  return {
    'Dashboard & Analytics': { icon: 'fas fa-chart-line', description: 'Main dashboards and analytics views', sortOrder: 1, color: '#3498db' },
    'Quality Assurance': { icon: 'fas fa-clipboard-check', description: 'Quality assurance forms and reporting', sortOrder: 2, color: '#e74c3c' },
    'Campaign QA': { icon: 'fas fa-shield-alt', description: 'Campaign-specific quality assurance', sortOrder: 3, color: '#9b59b6' },
    'Reporting & Analytics': { icon: 'fas fa-chart-bar', description: 'Reports and data analytics', sortOrder: 4, color: '#f39c12' },
    'Coaching & Development': { icon: 'fas fa-chalkboard-teacher', description: 'Employee coaching and development', sortOrder: 5, color: '#27ae60' },
    'Task Management': { icon: 'fas fa-tasks', description: 'Task creation and management', sortOrder: 6, color: '#2ecc71' },
    'Scheduling & Time': { icon: 'fas fa-calendar-alt', description: 'Schedule and time management', sortOrder: 7, color: '#16a085' },
    'Workflow & Operations': { icon: 'fas fa-cogs', description: 'Daily operations and workflow', sortOrder: 8, color: '#34495e' },
    'Communication': { icon: 'fas fa-comments', description: 'Communication and collaboration tools', sortOrder: 9, color: '#8e44ad' },
    'Administration': { icon: 'fas fa-users-cog', description: 'System administration and management', sortOrder: 10, color: '#c0392b' },
    'Data Management': { icon: 'fas fa-database', description: 'Data import and management', sortOrder: 11, color: '#d35400' },
    'Forms & Utilities': { icon: 'fas fa-tools', description: 'Special forms and utility functions', sortOrder: 12, color: '#7f8c8d' },
    'Authentication': { icon: 'fas fa-lock', description: 'Authentication and security pages', sortOrder: 13, color: '#95a5a6' }
  };
}

// Initialize system pages
function initializeEnhancedSystemPages() {
  try {
    ensureSheetWithHeaders(PAGES_SHEET, PAGES_HEADERS);
    const pagesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PAGES_SHEET);
    const existingPages = readSheet(PAGES_SHEET);
    if ((existingPages || []).length > 0) {
      enhancedAutoDiscoverAndSavePages({ force: true });
      return;
    }

    const systemPages = getAllPagesFromActualRouting();
    const now = new Date().toISOString();
    systemPages.forEach(page => {
      try {
        pagesSheet.appendRow([
          page.key, page.title, page.icon, page.description,
          page.isSystem === true, page.requiresAdmin === true, now, now
        ]);
      } catch (err) { console.error(`Failed to add page ${page.key}:`, err); }
    });

    invalidateCache(PAGES_SHEET);
  } catch (e) {
    safeWriteError('initializeEnhancedSystemPages', e);
  }
}

// Discovery + save/update Pages rows
function enhancedAutoDiscoverAndSavePages(opts = {}) {
  try {
    const { force = false, minIntervalSec = 300 } = opts;

    ensureSheetWithHeaders(PAGES_SHEET, PAGES_HEADERS);

    const prop = PropertiesService.getScriptProperties();
    const k = 'ENHANCED_PAGE_DISCOVERY_LAST_RUN';
    const last = Number(prop.getProperty(k) || '0');
    const now = Date.now();
    if (!force && last && (now - last) / 1000 < minIntervalSec) {
      return { skipped: true, reason: 'throttled' };
    }

    const discovered = getAllPagesFromActualRouting();
    const existing = readSheet(PAGES_SHEET) || [];
    const existingKeys = existing.map(p => (p.PageKey || '').toLowerCase());

    const newPages = discovered.filter(p => !existingKeys.includes(p.key.toLowerCase()));
    const pagesToUpdate = discovered.filter(p => existingKeys.includes(p.key.toLowerCase()));

    const sheet = ensureSheetWithHeaders(PAGES_SHEET, PAGES_HEADERS);
    const nowIso = new Date().toISOString();
    let addedCount = 0, updatedCount = 0;

    newPages.forEach(page => {
      try {
        sheet.appendRow([
          page.key, page.title, page.icon, page.description,
          !!page.isSystem, !!page.requiresAdmin, nowIso, nowIso
        ]);
        addedCount++;
      } catch (e) { console.error(`Failed to add page ${page.key}:`, e); }
    });

    if (pagesToUpdate.length > 0) {
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const pageKeyIndex = headers.indexOf('PageKey');
      const titleIndex = headers.indexOf('PageTitle');
      const iconIndex = headers.indexOf('PageIcon');
      const descIndex = headers.indexOf('Description');
      const updatedAtIndex = headers.indexOf('UpdatedAt');

      pagesToUpdate.forEach(page => {
        try {
          const rowIndex = data.findIndex((row, idx) =>
            idx > 0 && row[pageKeyIndex] && String(row[pageKeyIndex]).toLowerCase() === page.key.toLowerCase()
          );
          if (rowIndex > 0) {
            const r = rowIndex + 1;
            if (data[rowIndex][titleIndex] !== page.title) sheet.getRange(r, titleIndex + 1).setValue(page.title);
            if (data[rowIndex][iconIndex] !== page.icon) sheet.getRange(r, iconIndex + 1).setValue(page.icon);
            if (data[rowIndex][descIndex] !== page.description) sheet.getRange(r, descIndex + 1).setValue(page.description);
            sheet.getRange(r, updatedAtIndex + 1).setValue(nowIso);
            updatedCount++;
          }
        } catch (e) { console.error(`Failed to update page ${page.key}:`, e); }
      });
    }

    invalidateCache(PAGES_SHEET);
    prop.setProperty(k, String(now));

    return { skipped: false, success: true, added: addedCount, updated: updatedCount, total: (existing.length + addedCount), newPages: newPages.map(p => ({ key: p.key, title: p.title, category: p.category })) };
  } catch (e) {
    safeWriteError('enhancedAutoDiscoverAndSavePages', e);
    return { skipped: false, success: false, error: e.message };
  }
}

// Icons
function suggestIconForPageKey(key) {
  try {
    const k = String(key || '').toLowerCase();
    const iconMap = {
      dashboard: 'fa-tachometer-alt',
      'agent-experience': 'fa-user-headset',
      goalsetting: 'fa-bullseye',
      'manager-executive-experience': 'fa-briefcase',
      qualityform: 'fa-clipboard-check',
      ibtrqualityreports: 'fa-chart-pie',
      qualityview: 'fa-eye',
      qualitylist: 'fa-list-check',
      unifiedqadashboard: 'fa-layer-group',
      qacollablist: 'fa-users-gear',
      qacollabview: 'fa-users-viewfinder',
      qualitycollabform: 'fa-clipboard-list',
      qacollabform: 'fa-clipboard-list',
      independencequality: 'fa-shield-alt',
      independenceqadashboard: 'fa-chart-line',
      creditsuiteqa: 'fa-credit-card',
      'qa-dashboard': 'fa-chart-area',
      groundingqaform: 'fa-seedling',
      callreports: 'fa-phone-volume',
      attendancereports: 'fa-chart-bar',
      'collaboration-reporting': 'fa-people-arrows',
      collaborationreporting: 'fa-people-arrows',
      coachingdashboard: 'fa-chalkboard-teacher',
      coachingview: 'fa-search-plus',
      coachinglist: 'fa-list-ul',
      coachings: 'fa-users',
      coachingsheet: 'fa-file-signature',
      coaching: 'fa-graduation-cap',
      tasksmanager: 'fa-columns',
      taskform: 'fa-plus-square',
      schedulemanagement: 'fa-calendar-week',
      attendancecalendar: 'fa-calendar-check',
      calendar: 'fa-calendar-days',
      slotmanagement: 'fa-clock',
      'agent-schedule': 'fa-calendar-day',
      escalations: 'fa-exclamation-triangle',
      eodreport: 'fa-clipboard-check',
      incentives: 'fa-trophy',
      search: 'fa-search',
      chat: 'fa-comments',
      bookmarks: 'fa-bookmark',
      notifications: 'fa-bell',
      manageuser: 'fa-users-cog',
      manageroles: 'fa-user-shield',
      managecampaign: 'fa-bullhorn',
      settings: 'fa-cogs',
      import: 'fa-file-import',
      importattendance: 'fa-file-upload',
      ackform: 'fa-signature',
      proxy: 'fa-exchange-alt',
      setpassword: 'fa-key',
      resetpassword: 'fa-unlock-alt',
      'resend-verification': 'fa-envelope-circle-check',
      resendverification: 'fa-envelope-circle-check',
      forgotpassword: 'fa-question-circle',
      'forgot-password': 'fa-question-circle',
      emailconfirmed: 'fa-check-circle',
      'email-confirmed': 'fa-check-circle'
    };
    if (iconMap[k]) return 'fas ' + iconMap[k];

    if (k.includes('dashboard')) return 'fas fa-chart-line';
    if (k.includes('quality') || k.includes('qa')) return 'fas fa-clipboard-check';
    if (k.includes('coach')) return 'fas fa-chalkboard-teacher';
    if (k.includes('task')) return 'fas fa-tasks';
    if (k.includes('schedule') || k.includes('calendar')) return 'fas fa-calendar-alt';
    if (k.includes('report')) return 'fas fa-chart-bar';
    if (k.includes('chat') || k.includes('message')) return 'fas fa-comments';
    if (k.includes('user') || k.includes('manage')) return 'fas fa-users-cog';
    if (k.includes('admin')) return 'fas fa-user-shield';
    if (k.includes('import')) return 'fas fa-file-import';
    if (k.includes('password') || k.includes('auth')) return 'fas fa-key';
    return 'fas fa-file';
  } catch (e) {
    safeWriteError('suggestIconForPageKey', e);
    return 'fas fa-file';
  }
}
function refreshAllPageIcons() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(PAGES_SHEET);
    const rows = readSheet(PAGES_SHEET);
    if (!rows.length) return { updated: 0 };

    const headers = PAGES_HEADERS, col = name => headers.indexOf(name) + 1;
    let updated = 0;

    rows.forEach((r, i) => {
      const key = String(r.PageKey || '').toLowerCase();
      if (!key) return;
      const suggestedIcon = suggestIconForPageKey(key);
      const rowNum = i + 2;
      if (r.PageIcon !== suggestedIcon) {
        sh.getRange(rowNum, col('PageIcon')).setValue(suggestedIcon);
        sh.getRange(rowNum, col('UpdatedAt')).setValue(new Date().toISOString());
        updated++;
      }
    });

    if (updated) invalidateCache(PAGES_SHEET);
    return { updated, total: rows.length };
  } catch (e) {
    safeWriteError('refreshAllPageIcons', e);
    return { updated: 0, error: e.message };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// Category setup & assignment
// ────────────────────────────────────────────────────────────────────────────
function createEnhancedCategoriesForCampaign(campaignId, categoryDefinitions) {
  try {
    const existing = readSheet(PAGE_CATEGORIES_SHEET).filter(x => x.CampaignID === campaignId);
    if (existing.length > 0) return { success: true, categoriesCreated: 0, skipped: true };

    const sheet = ensureSheetWithHeaders(PAGE_CATEGORIES_SHEET, PAGE_CATEGORIES_HEADERS);
    const now = new Date().toISOString();
    let created = 0;

    Object.entries(categoryDefinitions).forEach(([name, data]) => {
      const id = generateUniqueId();
      sheet.appendRow([id, campaignId, name, data.icon, data.sortOrder, true, now, now]);
      created++;
    });

    invalidateCache(PAGE_CATEGORIES_SHEET);
    invalidateCache(`PAGE_CATEGORIES_${campaignId}`);
    return { success: true, categoriesCreated: created };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
function assignPagesToEnhancedCategories(campaignId) {
  try {
    const pages = getAllPagesFromActualRouting();
    const categories = getCampaignPageCategories(campaignId);
    const categoryMap = {};
    categories.forEach(c => categoryMap[c.CategoryName] = c.ID);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CAMPAIGN_PAGES_SHEET);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const campaignIdIndex = headers.indexOf('CampaignID');
    const pageKeyIndex = headers.indexOf('PageKey');
    const categoryIdIndex = headers.indexOf('CategoryID');

    let updated = 0;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[campaignIdIndex] === campaignId && row[pageKeyIndex]) {
        const pk = String(row[pageKeyIndex]).toLowerCase();
        const def = pages.find(p => p.key.toLowerCase() === pk);
        if (def && def.category) {
          const catId = categoryMap[def.category];
          if (catId && row[categoryIdIndex] !== catId) {
            sheet.getRange(i + 1, categoryIdIndex + 1).setValue(catId);
            updated++;
          }
        }
      }
    }

    invalidateCache(CAMPAIGN_PAGES_SHEET);
    invalidateCache(`CAMPAIGN_PAGES_${campaignId}`);
    invalidateCache(`NAVIGATION_${campaignId}`);
    return { success: true, pagesAssigned: updated };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
function setupEnhancedPageCategoriesForAllCampaigns() {
  try {
    const campaigns = readSheet(CAMPAIGNS_SHEET);
    const results = { success: 0, failed: 0, errors: [], details: [] };
    if (campaigns.length === 0) return results;

    const categories = getEnhancedPageCategories();

    campaigns.forEach(c => {
      try {
        const catRes = createEnhancedCategoriesForCampaign(c.ID, categories);
        const pagesRes = createCampaignPagesFromSystem(c.ID);
        const assignRes = assignPagesToEnhancedCategories(c.ID);
        if (catRes.success && pagesRes && assignRes.success) {
          results.success++;
          results.details.push({
            campaignId: c.ID,
            campaignName: c.Name,
            status: 'success',
            categoriesCreated: catRes.categoriesCreated || 0,
            pagesAssigned: assignRes.pagesAssigned || 0,
            skipped: !!catRes.skipped
          });
        } else {
          results.failed++;
          results.errors.push(`Failed for ${c.Name}: ${(catRes.error || '') || (!pagesRes ? 'create pages failed' : '') || (assignRes.error || '')}`);
        }
      } catch (e) {
        results.failed++;
        results.errors.push(`Campaign ${c.Name} (${c.ID}): ${e.message}`);
        safeWriteError('setupEnhancedPageCategoriesForAllCampaigns', e);
      }
    });

    return results;
  } catch (e) {
    safeWriteError('setupEnhancedPageCategoriesForAllCampaigns', e);
    return { success: 0, failed: 1, errors: [e.message], details: [] };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// Legacy compatibility helpers
// ────────────────────────────────────────────────────────────────────────────
function getAllPageKeys() {
  try {
    enhancedAutoDiscoverAndSavePages({ minIntervalSec: 300 });
    const cacheKey = 'PAGE_KEYS';
    const cached = scriptCache.get(cacheKey);
    if (cached) return JSON.parse(cached);

    const pagesFromSheet = readSheet(PAGES_SHEET).map(p => p.PageKey);
    const pagesFromCode = getAllPagesFromActualRouting().map(p => p.key);
    const keys = [...new Set([...pagesFromSheet, ...pagesFromCode])].sort();
    scriptCache.put(cacheKey, JSON.stringify(keys), CACHE_TTL_SEC);
    return keys;
  } catch (e) { safeWriteError('getAllPageKeys', e); return []; }
}
function getAllPages() {
  try {
    const cacheKey = `DATA_${PAGES_SHEET}`;
    const cached = scriptCache.get(cacheKey);
    if (cached) return JSON.parse(cached);
    const pages = readSheet(PAGES_SHEET);
    scriptCache.put(cacheKey, JSON.stringify(pages), CACHE_TTL_SEC);
    return pages;
  } catch (e) { safeWriteError('getAllPages', e); return []; }
}

// ────────────────────────────────────────────────────────────────────────────
// Campaign pages / categories / navigation
// ────────────────────────────────────────────────────────────────────────────
function getCampaignPages(campaignId) {
  try {
    const cacheKey = `CAMPAIGN_PAGES_${campaignId}`;
    const cached = scriptCache.get(cacheKey);
    if (cached) return JSON.parse(cached);

    const campaignPages = readSheet(CAMPAIGN_PAGES_SHEET)
      .filter(cp => cp && cp.CampaignID === campaignId && (cp.IsActive === true || cp.IsActive === 'TRUE'));

    const allPages = getAllPages();
    const pageMap = {}; allPages.forEach(p => pageMap[p.PageKey] = p);

    const result = campaignPages.map(cp => {
      const sys = pageMap[cp.PageKey];
      return {
        ID: cp.ID, CampaignID: cp.CampaignID, PageKey: cp.PageKey,
        PageTitle: cp.PageTitle || (sys ? sys.PageTitle : cp.PageKey),
        PageIcon: cp.PageIcon || (sys ? sys.PageIcon : 'fas fa-file'),
        CategoryID: cp.CategoryID || null, SortOrder: cp.SortOrder || 999,
        IsActive: cp.IsActive, PageDescription: sys ? sys.Description : ''
      };
    }).sort((a, b) => (a.SortOrder || 999) - (b.SortOrder || 999));

    scriptCache.put(cacheKey, JSON.stringify(result), CACHE_TTL_SEC);
    return result;
  } catch (e) { safeWriteError('getCampaignPages', e); return []; }
}
function getCampaignPageCategories(campaignId) {
  try {
    const cacheKey = `PAGE_CATEGORIES_${campaignId}`;
    const cached = scriptCache.get(cacheKey);
    if (cached) return JSON.parse(cached);

    const categories = readSheet(PAGE_CATEGORIES_SHEET)
      .filter(pc => pc && pc.CampaignID === campaignId && (pc.IsActive === true || pc.IsActive === 'TRUE'))
      .map(pc => ({
        ID: pc.ID, CampaignID: pc.CampaignID, CategoryName: pc.CategoryName,
        CategoryIcon: pc.CategoryIcon || 'fas fa-folder', SortOrder: pc.SortOrder || 999,
        IsActive: pc.IsActive, CreatedAt: pc.CreatedAt, UpdatedAt: pc.UpdatedAt
      }))
      .sort((a, b) => (a.SortOrder || 999) - (b.SortOrder || 999));

    scriptCache.put(cacheKey, JSON.stringify(categories), CACHE_TTL_SEC);
    return categories;
  } catch (e) { safeWriteError('getCampaignPageCategories', e); return []; }
}
function getCampaignNavigation(campaignId) {
  try {
    const cacheKey = `NAVIGATION_${campaignId}`;
    const cached = scriptCache.get(cacheKey);
    if (cached) return JSON.parse(cached);

    const pages = getCampaignPages(campaignId);
    const categories = getCampaignPageCategories(campaignId);
    const navigation = { categories: [], uncategorizedPages: [] };

    if (categories.length === 0) {
      navigation.uncategorizedPages = pages.map(p => ({ ...p, PageIcon: p.PageIcon || 'fas fa-file' }));
      scriptCache.put(cacheKey, JSON.stringify(navigation), CACHE_TTL_SEC);
      return navigation;
    }

    const categoryMap = {};
    categories.forEach(cat => {
      categoryMap[cat.ID] = {
        ID: cat.ID, CampaignID: cat.CampaignID, CategoryName: cat.CategoryName,
        CategoryIcon: cat.CategoryIcon || 'fas fa-folder', SortOrder: cat.SortOrder || 999,
        IsActive: cat.IsActive, pages: []
      };
    });

    pages.forEach(page => {
      page.PageIcon = page.PageIcon || 'fas fa-file';
      if (page.CategoryID && categoryMap[page.CategoryID]) categoryMap[page.CategoryID].pages.push(page);
      else navigation.uncategorizedPages.push(page);
    });

    navigation.categories = Object.values(categoryMap).filter(cat => cat.pages.length > 0);
    scriptCache.put(cacheKey, JSON.stringify(navigation), CACHE_TTL_SEC);
    return navigation;
  } catch (e) {
    safeWriteError('getCampaignNavigation', e);
    return { categories: [], uncategorizedPages: [] };
  }
}
function clientGetCampaignNavigation(campaignId) {
  try {
    if (!campaignId) return { categories: [], uncategorizedPages: [] };
    return getCampaignNavigation(campaignId);
  } catch (e) {
    safeWriteError('clientGetCampaignNavigation', e);
    return { categories: [], uncategorizedPages: [] };
  }
}

// Create campaign pages from system (if none)
function createCampaignPagesFromSystem(campaignId) {
  try {
    if (!campaignId) { safeWriteError('createCampaignPagesFromSystem', 'Campaign ID is required'); return false; }
    const existing = readSheet(CAMPAIGN_PAGES_SHEET).filter(cp => cp.CampaignID === campaignId);
    if (existing.length > 0) return true;

    const cpSheet = ensureSheetWithHeaders(CAMPAIGN_PAGES_SHEET, CAMPAIGN_PAGES_HEADERS);
    const systemPages = readSheet(PAGES_SHEET);
    const now = new Date().toISOString();
    let created = 0;

    systemPages.forEach((page, idx) => {
      if (page.RequiresAdmin === true || page.RequiresAdmin === 'TRUE') return;
      cpSheet.appendRow([generateUniqueId(), campaignId, page.PageKey, page.PageTitle, page.PageIcon, '', idx + 1, true, now, now]);
      created++;
    });

    invalidateCache(CAMPAIGN_PAGES_SHEET);
    invalidateCache(`CAMPAIGN_PAGES_${campaignId}`);
    invalidateCache(`NAVIGATION_${campaignId}`);
    return true;
  } catch (e) {
    safeWriteError('createCampaignPagesFromSystem', e);
    return false;
  }
}

// ────────────────────────────────────────────────────────────────────────────
// Identity / permissions helpers (guarded where likely to collide)
// ────────────────────────────────────────────────────────────────────────────
if (typeof getUserManagedCampaigns !== 'function') {
  function getUserManagedCampaigns(userId) {
    try {
      if (!userId) return [];
      const users = readSheet(USERS_SHEET);
      const u = users.find(x => x.ID === userId);
      if (u && (u.IsAdmin === 'TRUE' || u.IsAdmin === true)) return readSheet(CAMPAIGNS_SHEET);

      const perms = readSheet(CAMPAIGN_USER_PERMISSIONS_SHEET);
      const managedIds = perms.filter(p => p.UserID === userId && (p.PermissionLevel === 'MANAGER' || p.PermissionLevel === 'ADMIN')).map(p => p.CampaignID);
      if (managedIds.length === 0) return [];
      const all = readSheet(CAMPAIGNS_SHEET);
      return all.filter(c => managedIds.includes(c.ID));
    } catch (e) { writeError('getUserManagedCampaigns', e); return []; }
  }
}
function getUsersByCampaign(campaignId) {
  try {
    if (!campaignId) return [];
    const ucs = readSheet(USER_CAMPAIGNS_SHEET);
    const ids = ucs.filter(uc => uc.CampaignId === campaignId).map(uc => uc.UserId);
    const all = readSheet(USERS_SHEET);
    if (ids.length === 0) return all.filter(u => u.CampaignID === campaignId); // legacy
    return all.filter(u => ids.includes(u.ID));
  } catch (e) { writeError('getUsersByCampaign', e); return []; }
}
function userCanManageCampaign(userId, campaignId) {
  try {
    if (!userId || !campaignId) return false;
    const users = readSheet(USERS_SHEET);
    const u = users.find(x => x.ID === userId);
    if (u && (u.IsAdmin === 'TRUE' || u.IsAdmin === true)) return true;

    const perms = readSheet(CAMPAIGN_USER_PERMISSIONS_SHEET);
    const p = perms.find(x => x.UserID === userId && x.CampaignID === campaignId && (x.PermissionLevel === 'MANAGER' || x.PermissionLevel === 'ADMIN'));
    return !!p;
  } catch (e) { safeWriteError('userCanManageCampaign', e); return false; }
}
function getCampaignsWithUserCounts() {
  try {
    const campaigns = readSheet(CAMPAIGNS_SHEET);
    const ucs = readSheet(USER_CAMPAIGNS_SHEET);
    const users = readSheet(USERS_SHEET);
    return campaigns.map(c => {
      const joined = ucs.filter(uc => uc.CampaignId === c.ID).length;
      const legacy = users.filter(u => u.CampaignID === c.ID).length;
      return { ...c, userCount: joined + legacy, newUserCount: joined, legacyUserCount: legacy };
    });
  } catch (e) { writeError('getCampaignsWithUserCounts', e); return []; }
}
function addUserToCampaign(userId, campaignId) {
  try {
    if (!userId || !campaignId) return false;
    const sh = ensureSheetWithHeaders(USER_CAMPAIGNS_SHEET, USER_CAMPAIGNS_HEADERS);
    const rows = readSheet(USER_CAMPAIGNS_SHEET);
    const exists = rows.some(r => r.UserId === userId && r.CampaignId === campaignId);
    if (exists) return true;
    const now = new Date().toISOString();
    sh.appendRow([Utilities.getUuid(), userId, campaignId, now, now]);
    invalidateCache(USER_CAMPAIGNS_SHEET);
    return true;
  } catch (e) { safeWriteError('addUserToCampaign', e); return false; }
}
function getUserCampaignsSafe(userId) {
  try {
    const ucs = readSheet(USER_CAMPAIGNS_SHEET) || [];
    const joined = ucs.filter(r => r.UserId === userId).map(r => ({ campaignId: r.CampaignId, source: 'multi' }));
    if (joined.length) return joined;

    const users = readSheet(USERS_SHEET) || [];
    const u = users.find(x => x.ID === userId);
    if (u && u.CampaignID) return [{ campaignId: u.CampaignID, source: 'legacy' }];
    return [];
  } catch (e) { safeWriteError('getUserCampaignsSafe', e); return []; }
}
function clientGetAvailableCampaigns(requestingUserId = null) {
  try {
    const all = getAllCampaigns();
    if (!requestingUserId) return all.map(c => ({ id: c.ID, name: c.Name, description: c.Description || '' }));

    const users = readSheet(USERS_SHEET);
    const u = users.find(x => x.ID === requestingUserId);
    if (u && (u.IsAdmin === 'TRUE' || u.IsAdmin === true)) return all.map(c => ({ id: c.ID, name: c.Name, description: c.Description || '' }));

    const managed = getUserManagedCampaigns(requestingUserId);
    return managed.map(c => ({ id: c.ID, name: c.Name, description: c.Description || '' }));
  } catch (e) { writeError('clientGetAvailableCampaigns', e); return []; }
}
if (typeof getAllCampaigns !== 'function') {
  function getAllCampaigns() {
    try { return readSheet(CAMPAIGNS_SHEET); } catch (e) { writeError('getAllCampaigns', e); return []; }
  }
}
function clientGetCampaignStats(requestingUserId = null) {
  try {
    const stats = getCampaignsWithUserCounts();
    if (!requestingUserId) return stats;
    const managedIds = getUserManagedCampaigns(requestingUserId).map(c => c.ID);
    return stats.filter(c => managedIds.includes(c.ID));
  } catch (e) { writeError('clientGetCampaignStats', e); return []; }
}
function clientCanAccessUser(requestingUserId, targetUserId) {
  try {
    if (!requestingUserId || !targetUserId) return false;
    if (requestingUserId === targetUserId) return true;

    const users = readSheet(USERS_SHEET);
    const r = users.find(u => u.ID === requestingUserId);
    if (r && (r.IsAdmin === 'TRUE' || r.IsAdmin === true)) return true;

    const managedCampaignIds = getUserManagedCampaigns(requestingUserId).map(c => c.ID);
    const targetCampaigns = getUserCampaignsSafe(targetUserId).map(uc => uc.campaignId);
    return targetCampaigns.some(cId => managedCampaignIds.includes(cId));
  } catch (e) { writeError('clientCanAccessUser', e); return false; }
}

// ────────────────────────────────────────────────────────────────────────────
// Navigation for a user (multi-campaign aware)
// ────────────────────────────────────────────────────────────────────────────
function clientGetNavigationForUser(userId) {
  try {
    if (!userId) return { categories: [], uncategorizedPages: [] };

    const users = readSheet(USERS_SHEET) || [];
    const user = users.find(u => u.ID === userId);
    if (!user) return { categories: [], uncategorizedPages: [] };

    // Clear nav cache for all campaigns this user belongs to
    const userCampaigns = getUserCampaignsSafe(userId);
    userCampaigns.forEach(uc => invalidateNavigationCache(uc.campaignId));

    // Single-campaign user → primary campaign nav
    if (!isMultiCampaignUser(user)) {
      const primaryCampaignId = userCampaigns.length ? userCampaigns[0].campaignId : user.CampaignID;
      if (primaryCampaignId) return getCampaignNavigation(primaryCampaignId);
    }

    // Multi-campaign user → categories per accessible campaign
    const cacheKey = `NAVIGATION_MULTI_${userId}`;
    const cached = scriptCache.get(cacheKey);
    if (cached) return JSON.parse(cached);

    const isSysAdmin = String(user.IsAdmin).toUpperCase() === 'TRUE';
    const multiId = getOrCreateMultiCampaignId();

    let accessibleCampaigns = [];
    if (isSysAdmin) {
      accessibleCampaigns = (readSheet(CAMPAIGNS_SHEET) || []).filter(c => c && c.ID && c.Name && c.ID !== multiId);
    } else {
      const campaignIds = userCampaigns.map(uc => uc.campaignId);
      const all = readSheet(CAMPAIGNS_SHEET) || [];
      accessibleCampaigns = all.filter(c => campaignIds.includes(c.ID));
    }

    const categories = accessibleCampaigns.sort((a, b) => (a.Name || '').localeCompare(b.Name || ''))
      .map(c => ({
        ID: `cat_multi_${c.ID}`, CampaignID: c.ID, CategoryName: c.Name,
        CategoryIcon: MULTI_CAMPAIGN_ICON, SortOrder: 999, IsActive: true, pages: getCampaignPages(c.ID)
      }));

    const nav = { categories, uncategorizedPages: [] };
    scriptCache.put(cacheKey, JSON.stringify(nav), CACHE_TTL_SEC);
    return nav;
  } catch (e) {
    safeWriteError('clientGetNavigationForUser', e);
    return { categories: [], uncategorizedPages: [] };
  }
}
function isMultiCampaignUser(user) {
  try {
    if (!user) return false;
    const multiId = getOrCreateMultiCampaignId();
    return multiId && String(user.CampaignID || '') === String(multiId);
  } catch (e) { safeWriteError('isMultiCampaignUser', e); return false; }
}

// ────────────────────────────────────────────────────────────────────────────
// Utilities
// ────────────────────────────────────────────────────────────────────────────
function generateUniqueId() {
  return 'cat_' + Utilities.getUuid().replace(/-/g, '').substring(0, 12);
}
function ok(msg = 'OK', extra = {}) {
  return Object.assign({ success: true, message: msg, _refresh: true, refreshScopes: ['all'] }, extra);
}
function fail(msg, extra = {}) {
  return Object.assign({ success: false, message: msg, _refresh: false }, extra);
}
function commitWrites() { try { SpreadsheetApp.flush(); } catch (_) { } }

if (typeof getCampaignById !== 'function') {
  function getCampaignById(campaignId) {
    try { const c = readSheet(CAMPAIGNS_SHEET).find(x => x.ID === campaignId); return c || null; }
    catch (e) { safeWriteError('getCampaignById', e); return null; }
  }
}
if (typeof getAllRoles !== 'function') {
  function getAllRoles() {
    try { return readSheet(ROLES_SHEET); } catch (e) { safeWriteError('getAllRoles', e); return []; }
  }
}

// ────────────────────────────────────────────────────────────────────────────
/** Client-accessible setup & discovery wrappers */
// ────────────────────────────────────────────────────────────────────────────
if (typeof clientRunEnhancedDiscovery !== 'function') {
  function clientRunEnhancedDiscovery() {
    try { return enhancedAutoDiscoverAndSavePages({ force: true }); }
    catch (e) { safeWriteError('clientRunEnhancedDiscovery', e); return { success: false, error: e.message }; }
  }
}
function clientSetupEnhancedCategories() {
  try { return setupEnhancedPageCategoriesForAllCampaigns(); }
  catch (e) { safeWriteError('clientSetupEnhancedCategories', e); return { success: false, error: e.message }; }
}
function runEnhancedMainSetup() {
  try {
    const discovery = enhancedAutoDiscoverAndSavePages({ force: true });
    if (!discovery.success) throw new Error(`Page discovery failed: ${discovery.error}`);

    const categoriesResult = setupEnhancedPageCategoriesForAllCampaigns();

    const campaigns = readSheet(CAMPAIGNS_SHEET) || [];
    campaigns.forEach(c => invalidateNavigationCache(c.ID));
    [PAGES_SHEET, CAMPAIGN_PAGES_SHEET, PAGE_CATEGORIES_SHEET, CAMPAIGNS_SHEET].forEach(invalidateCache);

    return {
      success: true,
      discovery, categories: categoriesResult,
      summary: {
        totalPages: discovery.total, newPages: discovery.added, updatedPages: discovery.updated,
        campaignsConfigured: categoriesResult.success,
        categoriesAvailable: Object.keys(getEnhancedPageCategories()).length
      }
    };
  } catch (e) {
    safeWriteError('runEnhancedMainSetup', e);
    return { success: false, error: e.message };
  }
}
function clientRunEnhancedSetup() {
  try { return runEnhancedMainSetup(); }
  catch (e) { safeWriteError('clientRunEnhancedSetup', e); return { success: false, error: e.message }; }
}
function debugMultiCampaignSetup() {
  try {
    const multiId = getOrCreateMultiCampaignId();
    const counts = {
      campaigns: (readSheet(CAMPAIGNS_SHEET) || []).length,
      pages: (readSheet(PAGES_SHEET) || []).length,
      categories: (readSheet(PAGE_CATEGORIES_SHEET) || []).length,
      campaignPages: (readSheet(CAMPAIGN_PAGES_SHEET) || []).length,
      userCampaigns: (readSheet(USER_CAMPAIGNS_SHEET) || []).length,
      multiCampaignId: multiId
    };
    return { success: true, counts };
  } catch (e) { safeWriteError('debugMultiCampaignSetup', e); return { success: false, error: e.message }; }
}
function clientDebugMultiCampaignSetup() { return debugMultiCampaignSetup(); }

// ────────────────────────────────────────────────────────────────────────────
// Health & bootstrap
// ────────────────────────────────────────────────────────────────────────────
function setupMainSheets() {
  try {
    ensureSheetWithHeaders(USERS_SHEET, USERS_HEADERS);
    ensureSheetWithHeaders(ROLES_SHEET, ROLES_HEADER);
    ensureSheetWithHeaders(USER_ROLES_SHEET, USER_ROLES_HEADER);
    ensureSheetWithHeaders(USER_CLAIMS_SHEET, CLAIMS_HEADERS);
    ensureSheetWithHeaders(SESSIONS_SHEET, SESSIONS_HEADERS);

    // Multi-campaign support
    ensureSheetWithHeaders(USER_CAMPAIGNS_SHEET, USER_CAMPAIGNS_HEADERS);

    // Chat system
    ensureSheetWithHeaders(CHAT_GROUPS_SHEET, CHAT_GROUPS_HEADERS);
    ensureSheetWithHeaders(CHAT_GROUP_MEMBERS_SHEET, CHAT_GROUP_MEMBERS_HEADERS);
    ensureSheetWithHeaders(CHAT_CHANNELS_SHEET, CHAT_CHANNELS_HEADERS);
    ensureSheetWithHeaders(CHAT_CHANNEL_MEMBERS_SHEET, CHAT_CHANNEL_MEMBERS_HEADERS);
    ensureSheetWithHeaders(CHAT_MESSAGES_SHEET, CHAT_MESSAGES_HEADERS);
    ensureSheetWithHeaders(CHAT_MESSAGE_REACTIONS_SHEET, CHAT_MESSAGE_REACTIONS_HEADERS);
    ensureSheetWithHeaders(CHAT_USER_PREFERENCES_SHEET, CHAT_USER_PREFERENCES_HEADERS);
    ensureSheetWithHeaders(CHAT_ANALYTICS_SHEET, CHAT_ANALYTICS_HEADERS);

    // Campaign and page management
    ensureSheetWithHeaders(CAMPAIGNS_SHEET, CAMPAIGNS_HEADERS);
    ensureSheetWithHeaders(PAGES_SHEET, PAGES_HEADERS);
    ensureSheetWithHeaders(CAMPAIGN_PAGES_SHEET, CAMPAIGN_PAGES_HEADERS);
    ensureSheetWithHeaders(PAGE_CATEGORIES_SHEET, PAGE_CATEGORIES_HEADERS);
    ensureSheetWithHeaders(CAMPAIGN_USER_PERMISSIONS_SHEET, CAMPAIGN_USER_PERMISSIONS_HEADERS);
    ensureSheetWithHeaders(USER_MANAGERS_SHEET, getManagerUsersHeaders_());

    // Logs and notifications
    ensureSheetWithHeaders(DEBUG_LOGS_SHEET, DEBUG_LOGS_HEADERS);
    ensureSheetWithHeaders(ERROR_LOGS_SHEET, ERROR_LOGS_HEADERS);
    ensureSheetWithHeaders(NOTIFICATIONS_SHEET, NOTIFICATIONS_HEADERS);

    // ✅ NEW: upgrade Users with HR/Benefits columns
    ensureUsersBenefitsColumns_();

    // Initialize enhanced system pages and icons
    initializeEnhancedSystemPages();
    enhancedAutoDiscoverAndSavePages({ minIntervalSec: 300 });
    refreshAllPageIcons();
    getOrCreateMultiCampaignId();

    console.log('setupMainSheets completed with multi-campaign + benefits support');
  } catch (e) {
    console.error('setupMainSheets failed:', e);
  }
}

function healthCheckMain() {
  const results = {};
  const specs = [
    { name: USERS_SHEET, headers: USERS_HEADERS },
    { name: ROLES_SHEET, headers: ROLES_HEADER },
    { name: USER_ROLES_SHEET, headers: USER_ROLES_HEADER },
    { name: USER_CLAIMS_SHEET, headers: CLAIMS_HEADERS },
    { name: SESSIONS_SHEET, headers: SESSIONS_HEADERS },
    { name: CHAT_GROUPS_SHEET, headers: CHAT_GROUPS_HEADERS },
    { name: CHAT_GROUP_MEMBERS_SHEET, headers: CHAT_GROUP_MEMBERS_HEADERS },
    { name: CHAT_CHANNELS_SHEET, headers: CHAT_CHANNELS_HEADERS },
    { name: CHAT_CHANNEL_MEMBERS_SHEET, headers: CHAT_CHANNEL_MEMBERS_HEADERS },
    { name: CHAT_MESSAGES_SHEET, headers: CHAT_MESSAGES_HEADERS },
    { name: CHAT_MESSAGE_REACTIONS_SHEET, headers: CHAT_MESSAGE_REACTIONS_HEADERS },
    { name: CHAT_USER_PREFERENCES_SHEET, headers: CHAT_USER_PREFERENCES_HEADERS },
    { name: CHAT_ANALYTICS_SHEET, headers: CHAT_ANALYTICS_HEADERS },
    { name: CAMPAIGNS_SHEET, headers: CAMPAIGNS_HEADERS },
    { name: PAGES_SHEET, headers: PAGES_HEADERS },
    { name: CAMPAIGN_PAGES_SHEET, headers: CAMPAIGN_PAGES_HEADERS },
    { name: PAGE_CATEGORIES_SHEET, headers: PAGE_CATEGORIES_HEADERS },
    { name: CAMPAIGN_USER_PERMISSIONS_SHEET, headers: CAMPAIGN_USER_PERMISSIONS_HEADERS },
    { name: DEBUG_LOGS_SHEET, headers: DEBUG_LOGS_HEADERS },
    { name: ERROR_LOGS_SHEET, headers: ERROR_LOGS_HEADERS },
    { name: NOTIFICATIONS_SHEET, headers: NOTIFICATIONS_HEADERS },
    { name: USER_MANAGERS_SHEET, headers: getManagerUsersHeaders_() },
    { name: USER_CAMPAIGNS_SHEET, headers: USER_CAMPAIGNS_HEADERS }
  ];
  specs.forEach(({ name, headers }) => {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(name);
      if (!sh) { results[name] = { ok: false, message: 'Missing sheet' }; }
      else {
        const existing = sh.getRange(1, 1, 1, headers.length).getValues()[0];
        const match = headers.every((h, i) => existing[i] === h);
        results[name] = match ? { ok: true, message: 'OK' } : { ok: false, message: 'Header mismatch' };
      }
    } catch (e) { results[name] = { ok: false, message: `Error: ${e.message}` }; }
  });
  return results;
}

// ────────────────────────────────────────────────────────────────────────────
// Banner
// ────────────────────────────────────────────────────────────────────────────
console.log('📦 MainUtilities.gs loaded (enhanced, collision-safe, multi-campaign ready)');
console.log('🔧 Available: navigation, discovery, categories, manager-users, and setup utilities');
console.log('🔧 Multi-campaign functions available:');
console.log('   - getUserManagedCampaigns()');
console.log('   - getUsersByCampaign()');
console.log('   - userCanManageCampaign()');
console.log('   - clientGetAvailableCampaigns()');
console.log('   - clientMigrateLegacyUsers()');
console.log('   - clientGetNavigationForUser()');
console.log('   - clientGetCampaignNavigation()');
console.log('   - clientRunEnhancedDiscovery() / clientSetupEnhancedCategories() / clientRunEnhancedSetup()');
console.log('   - clientDebugMultiCampaignSetup()');
