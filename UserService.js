/**
 * UserService.js
 * -----------------------------------------------------------------------------
 * Centralised call-centre user directory backed by the DatabaseManager/SheetsDB
 * tables introduced in the authentication + campaign refactors.  The goal is to
 * expose predictable CRUD helpers that work against the `Users`, `UserRoles`,
 * `CampaignUserPermissions`, and manager mapping sheets while remaining friendly
 * to legacy Google Apps Script clients.
 *
 * The module exposes a small `UserDirectory` facade internally and then
 * re-exports the most commonly used global functions (clientGetAllUsers,
 * clientRegisterUser, etc.) so existing HTML service front-ends continue to
 * operate without modification.
 */
(function (global) {
  'use strict';

  if (global.UserDirectory) {
    return;
  }

  // ---------------------------------------------------------------------------
  // Configuration & constants
  // ---------------------------------------------------------------------------

  var USERS_TABLE_NAME = (typeof global.USERS_SHEET === 'string' && global.USERS_SHEET)
    ? global.USERS_SHEET
    : 'Users';
  var USER_ROLES_TABLE_NAME = (typeof global.USER_ROLES_SHEET === 'string' && global.USER_ROLES_SHEET)
    ? global.USER_ROLES_SHEET
    : 'UserRoles';
  var MANAGER_USERS_TABLE_NAME = (typeof global.MANAGER_USERS_SHEET === 'string' && global.MANAGER_USERS_SHEET)
    ? global.MANAGER_USERS_SHEET
    : 'ManagerUsers';
  var CAMPAIGN_PERMISSIONS_TABLE_NAME = (typeof global.CAMPAIGN_USER_PERMISSIONS_SHEET === 'string'
    && global.CAMPAIGN_USER_PERMISSIONS_SHEET)
    ? global.CAMPAIGN_USER_PERMISSIONS_SHEET
    : 'CampaignUserPermissions';
  var PAGES_TABLE_NAME = (typeof global.PAGES_SHEET === 'string' && global.PAGES_SHEET)
    ? global.PAGES_SHEET
    : 'Pages';

  var DEFAULT_USER_HEADERS = ['ID', 'UserName', 'FullName', 'Email', 'CampaignID', 'PasswordHash', 'ResetRequired',
    'EmailConfirmation', 'EmailConfirmed', 'PhoneNumber', 'EmploymentStatus', 'HireDate', 'Country',
    'LockoutEnd', 'TwoFactorEnabled', 'CanLogin', 'Roles', 'Pages', 'CreatedAt', 'UpdatedAt', 'LastLogin',
    'DeletedAt', 'IsAdmin'];

  var OPTIONAL_USER_HEADERS = ['TerminationDate', 'ProbationMonths', 'ProbationEndDate', 'InsuranceQualifiedDate',
    'InsuranceEligible', 'InsuranceSignedUp', 'InsuranceCardReceivedDate'];

  var EMPLOYMENT_STATUS = ['Active', 'Inactive', 'Terminated', 'On Leave', 'Pending', 'Probation', 'Contract',
    'Contractor', 'Full Time', 'Part Time', 'Seasonal', 'Temporary', 'Suspended', 'Retired', 'Intern', 'Consultant'];

  // ---------------------------------------------------------------------------
  // Guards & helpers
  // ---------------------------------------------------------------------------

  function hasDatabaseManager() {
    return typeof global.DatabaseManager !== 'undefined'
      && global.DatabaseManager
      && typeof global.DatabaseManager.defineTable === 'function';
  }

  function normalizeString(value) {
    if (value === null || typeof value === 'undefined') {
      return '';
    }
    return String(value).trim();
  }

  function normalizeEmail(value) {
    return normalizeString(value).toLowerCase();
  }

  function normalizeUserName(value) {
    return normalizeString(value).toLowerCase();
  }

  function toBoolean(value) {
    if (value === true || value === false) {
      return value;
    }
    var normalized = normalizeString(value).toLowerCase();
    if (!normalized) return false;
    return normalized === 'true' || normalized === '1' || normalized === 'yes' || normalized === 'y';
  }

  function parseList(value) {
    if (!value && value !== 0) {
      return [];
    }
    if (Array.isArray(value)) {
      return value.filter(function (entry) { return normalizeString(entry); }).map(function (entry) {
        return normalizeString(entry);
      });
    }
    return String(value).split(',').map(function (part) { return normalizeString(part); }).filter(function (part) {
      return part.length > 0;
    });
  }

  function joinList(values) {
    if (!Array.isArray(values) || !values.length) {
      return '';
    }
    var trimmed = values.map(function (entry) { return normalizeString(entry); }).filter(function (entry) {
      return entry.length > 0;
    });
    return trimmed.join(',');
  }

  function isSoftDeleted(row) {
    if (!row) {
      return false;
    }
    var marker = row.DeletedAt || row.deletedAt || row.deleted_at;
    if (marker === null || typeof marker === 'undefined') {
      return false;
    }
    return normalizeString(marker).length > 0;
  }

  function isValidEmploymentStatus(value) {
    if (!value && value !== 0) {
      return true;
    }
    var normalized = normalizeString(value).toLowerCase();
    if (!normalized) {
      return true;
    }
    return EMPLOYMENT_STATUS.some(function (status) {
      return status.toLowerCase() === normalized;
    });
  }

  function ensureArray(value) {
    if (!value && value !== 0) return [];
    if (Array.isArray(value)) return value.slice();
    return [value];
  }

  function ensureDateValue(value) {
    if (!value && value !== 0) {
      return '';
    }
    if (value instanceof Date) {
      return value;
    }
    var asNumber = Number(value);
    if (!isNaN(asNumber)) {
      return new Date(asNumber);
    }
    var parsed = new Date(String(value));
    if (isNaN(parsed)) {
      return '';
    }
    return parsed;
  }

  function getUuid() {
    return (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.getUuid === 'function')
      ? Utilities.getUuid()
      : String(Math.random()).slice(2) + String(new Date().getTime());
  }

  function getNow() {
    return new Date();
  }

  // ---------------------------------------------------------------------------
  // Database table lookups
  // ---------------------------------------------------------------------------

  function getUsersTable() {
    if (!hasDatabaseManager()) {
      throw new Error('DatabaseManager is required for UserDirectory operations');
    }
    return global.DatabaseManager.defineTable(USERS_TABLE_NAME, {
      headers: DEFAULT_USER_HEADERS.concat(OPTIONAL_USER_HEADERS),
      idColumn: 'ID',
      timestamps: { created: 'CreatedAt', updated: 'UpdatedAt' },
      defaults: {
        ResetRequired: false,
        EmailConfirmed: false,
        CanLogin: true,
        TwoFactorEnabled: false,
        Roles: '',
        Pages: '',
        IsAdmin: false,
        DeletedAt: ''
      },
      validators: {
        Email: function (value) { return normalizeEmail(value).length > 0; },
        EmploymentStatus: function (value) { return isValidEmploymentStatus(value); }
      }
    });
  }

  }

  function isValidEmploymentStatus(value) {
    if (!value && value !== 0) {
      return true;
    }
    var normalized = normalizeString(value).toLowerCase();
    if (!normalized) {
      return true;
    }
    return EMPLOYMENT_STATUS.some(function (status) {
      return status.toLowerCase() === normalized;
    });
  }

  function ensureArray(value) {
    if (!value && value !== 0) return [];
    if (Array.isArray(value)) return value.slice();
    return [value];
  }

  function ensureDateValue(value) {
    if (!value && value !== 0) {
      return '';
    }
    if (value instanceof Date) {
      return value;
    }
    var asNumber = Number(value);
    if (!isNaN(asNumber)) {
      return new Date(asNumber);
    }
    var parsed = new Date(String(value));
    if (isNaN(parsed)) {
      return '';
    }
    return parsed;
  }

  function getUuid() {
    return (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.getUuid === 'function')
      ? Utilities.getUuid()
      : String(Math.random()).slice(2) + String(new Date().getTime());
  }

  function getNow() {
    return new Date();
  }

  // ---------------------------------------------------------------------------
  // Database table lookups
  // ---------------------------------------------------------------------------

  function getUsersTable() {
    if (!hasDatabaseManager()) {
      throw new Error('DatabaseManager is required for UserDirectory operations');
    }
    return global.DatabaseManager.defineTable(USERS_TABLE_NAME, {
      headers: DEFAULT_USER_HEADERS.concat(OPTIONAL_USER_HEADERS),
      idColumn: 'ID',
      timestamps: { created: 'CreatedAt', updated: 'UpdatedAt' },
      defaults: {
        ResetRequired: false,
        EmailConfirmed: false,
        CanLogin: true,
        TwoFactorEnabled: false,
        Roles: '',
        Pages: '',
        IsAdmin: false,
        DeletedAt: ''
      },
      validators: {
        Email: function (value) { return normalizeEmail(value).length > 0; },
        EmploymentStatus: function (value) { return isValidEmploymentStatus(value); }
      }
    });
  }

  function getUserRolesTable() {
    if (!hasDatabaseManager()) {
      throw new Error('DatabaseManager is required for UserDirectory operations');
    }
    return global.DatabaseManager.defineTable(USER_ROLES_TABLE_NAME, {
      headers: ['ID', 'UserId', 'RoleId', 'Scope', 'AssignedBy', 'CreatedAt', 'UpdatedAt', 'DeletedAt'],
      idColumn: 'ID',
      timestamps: { created: 'CreatedAt', updated: 'UpdatedAt' },
      defaults: { Scope: 'global', AssignedBy: '', DeletedAt: '' }
    });
  }

  function getCampaignPermissionsTable() {
    if (!hasDatabaseManager()) {
      throw new Error('DatabaseManager is required for UserDirectory operations');
    }
    return global.DatabaseManager.defineTable(CAMPAIGN_PERMISSIONS_TABLE_NAME, {
      headers: ['ID', 'CampaignID', 'UserID', 'PermissionLevel', 'Role', 'CanManageUsers', 'CanManagePages', 'Notes',
        'CreatedAt', 'UpdatedAt', 'DeletedAt'],
      idColumn: 'ID',
      timestamps: { created: 'CreatedAt', updated: 'UpdatedAt' },
      defaults: { PermissionLevel: 'USER', Role: '', CanManageUsers: false, CanManagePages: false, Notes: '', DeletedAt: '' }
    });
  }

  function getManagerAssignmentsTable() {
    if (!hasDatabaseManager()) {
      throw new Error('DatabaseManager is required for UserDirectory operations');
    }
    return global.DatabaseManager.defineTable(MANAGER_USERS_TABLE_NAME, {
      headers: ['ID', 'ManagerUserID', 'UserID', 'CreatedAt', 'UpdatedAt', 'DeletedAt'],
      idColumn: 'ID',
      timestamps: { created: 'CreatedAt', updated: 'UpdatedAt' },
      defaults: { DeletedAt: '' }
    });
  }

  // Page definitions are frequently customised, so fall back to a readSheet helper
  function readPagesSheet() {
    if (typeof global.readSheet === 'function') {
      return global.readSheet(PAGES_TABLE_NAME) || [];
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PAGES_TABLE_NAME);
    if (!sheet) {
      return [];
    }
    var values = sheet.getDataRange().getValues();
    if (!values.length) {
      return [];
    }
    var headers = values[0];
    return values.slice(1).map(function (row) {
      var record = {};
      for (var i = 0; i < headers.length; i++) {
        record[String(headers[i])] = row[i];
      }
      return record;
    });
  }

  // ---------------------------------------------------------------------------
  // Role helpers
  // ---------------------------------------------------------------------------

  function resolveRoleIds(userId, fallbackColumn) {
    var ids = [];
    if (typeof global.getUserRoleIds === 'function') {
      try {
        ids = global.getUserRoleIds(userId) || [];
      } catch (err) {
        logError('resolveRoleIds:getUserRoleIds', err);
      }
    }
    if ((!ids || !ids.length) && fallbackColumn) {
      ids = parseList(fallbackColumn);
    }
    return ids;
  }

  function resolveRoleNames(roleIds) {
    if (!roleIds || !roleIds.length) {
      return [];
    }
    if (typeof global.getUserRoleNames === 'function') {
      try {
        return global.getUserRoleNames(roleIds) || [];
      } catch (err) {
        logError('resolveRoleNames:getUserRoleNames', err);
      }
    }
    return roleIds.slice();
  }

  function syncRoles(userId, desiredRoleIds) {
    var targetRoles = ensureArray(desiredRoleIds).map(function (roleId) {
      return normalizeString(roleId);
    }).filter(function (roleId) { return roleId.length > 0; });

    if (!targetRoles.length) {
      if (typeof global.deleteUserRoles === 'function') {
        global.deleteUserRoles(userId);
      }
      updateRolesColumn(userId, '');
      return;
    }

    if (typeof global.deleteUserRoles === 'function') {
      try {
        global.deleteUserRoles(userId);
      } catch (err) {
        logError('syncRoles:deleteUserRoles', err);
      }
    }

    if (typeof global.addUserRole === 'function') {
      targetRoles.forEach(function (roleId) {
        try {
          global.addUserRole(userId, roleId);
        } catch (err) {
          logError('syncRoles:addUserRole:' + roleId, err);
        }
      });
    }

    updateRolesColumn(userId, joinList(targetRoles));
  }

  function updateRolesColumn(userId, joinedRoles) {
    try {
      var table = getUsersTable();
      table.update(userId, { Roles: joinedRoles || '' });
    } catch (err) {
      logError('updateRolesColumn', err);
    }
  }

  // ---------------------------------------------------------------------------
  // Logging helper
  // ---------------------------------------------------------------------------

  function logError(where, err) {
    if (typeof global.safeWriteError === 'function') {
      try { global.safeWriteError(where, err); } catch (_) { }
      return;
    }
    if (typeof global.writeError === 'function') {
      try { global.writeError(where, err); } catch (_) { }
      return;
    }
    try {
      console.error('[UserService:' + where + ']', err && (err.stack || err));
    } catch (_) {
      if (typeof Logger !== 'undefined' && Logger && typeof Logger.log === 'function') {
        Logger.log('UserService.' + where + ': ' + (err && err.message ? err.message : err));
      }
    }
  }

  // ---------------------------------------------------------------------------
  // Mapping helpers
  // ---------------------------------------------------------------------------

  function createSafeUserObject(record) {
    if (!record) {
      return null;
    }
    var obj = {
      id: record.ID || record.Id || record.id || '',
      userName: record.UserName || record.userName || record.Username || '',
      fullName: record.FullName || record.fullName || '',
      email: record.Email || record.email || '',
      campaignId: record.CampaignID || record.CampaignId || record.campaignId || '',
      phoneNumber: record.PhoneNumber || record.phoneNumber || '',
      employmentStatus: record.EmploymentStatus || record.employmentStatus || '',
      hireDate: record.HireDate || record.hireDate || '',
      country: record.Country || record.country || '',
      canLogin: toBoolean(record.CanLogin),
      isAdmin: toBoolean(record.IsAdmin),
      twoFactorEnabled: toBoolean(record.TwoFactorEnabled),
      lockoutEnd: record.LockoutEnd || '',
      lastLogin: record.LastLogin || '',
      createdAt: record.CreatedAt || record.createdAt || '',
      updatedAt: record.UpdatedAt || record.updatedAt || '',
      deletedAt: record.DeletedAt || record.deletedAt || '',
      rolesCsv: record.Roles || '',
      pages: parseList(record.Pages || record.pages || ''),
      metadata: {}
    };

    obj.roleIds = resolveRoleIds(obj.id, record.Roles || record.roles || '');
    obj.roleNames = resolveRoleNames(obj.roleIds);

    return obj;
  }

  function mapUsers(records) {
    return (records || []).map(function (record) { return createSafeUserObject(record); }).filter(function (entry) {
      return !!entry;
    });
  }

  // ---------------------------------------------------------------------------
  // Query helpers
  // ---------------------------------------------------------------------------

  function listUsers(options) {
    options = options || {};
    var includeDeleted = options.includeDeleted === true;
    var campaignFilter = normalizeString(options.campaignId || options.campaignID || '');
    var search = normalizeString(options.search || options.query || '');

    var table = getUsersTable();
    var rows = table.find({
      filter: function (row) {
        if (!includeDeleted && isSoftDeleted(row)) {
          return false;
        }
        if (campaignFilter) {
          var rowCampaign = normalizeString(row.CampaignID || row.CampaignId || row.campaignId);
          if (rowCampaign !== campaignFilter) {
            return false;
          }
        }
        if (search) {
          var haystack = [row.UserName, row.FullName, row.Email, row.PhoneNumber]
            .map(function (value) { return normalizeString(value).toLowerCase(); })
            .join(' ');
          if (haystack.indexOf(search.toLowerCase()) === -1) {
            return false;
          }
        }
        return true;
      },
      sortBy: 'FullName'
    });
    return mapUsers(rows);
  }

  function findUserById(userId) {
    if (!userId && userId !== 0) {
      return null;
    }
    var table = getUsersTable();
    var record = table.findById(String(userId));
    return record ? createSafeUserObject(record) : null;
  }

  function findRawUserById(userId) {
    if (!userId && userId !== 0) {
      return null;
    }
    var table = getUsersTable();
    return table.findById(String(userId));
  }

  function findUserByEmail(email) {
    var normalized = normalizeEmail(email);
    if (!normalized) {
      return null;
    }
    var table = getUsersTable();
    var results = table.find({
      filter: function (row) {
        if (isSoftDeleted(row)) {
          return false;
        }
        return normalizeEmail(row.Email || row.email) === normalized;
      },
      limit: 1
    });
    return results && results.length ? createSafeUserObject(results[0]) : null;
  }

  function checkUserConflicts(email, userName, excludeId) {
    var conflicts = { emailConflict: null, userNameConflict: null };
    var normalizedEmail = normalizeEmail(email);
    var normalizedUser = normalizeUserName(userName);
    var table = getUsersTable();

    if (normalizedEmail) {
      var emailMatches = table.find({
        filter: function (row) {
          if (isSoftDeleted(row)) return false;
          if (excludeId && String(row.ID) === String(excludeId)) return false;
          return normalizeEmail(row.Email || row.email) === normalizedEmail;
        },
        limit: 1
      });
      if (emailMatches && emailMatches.length) {
        conflicts.emailConflict = createSafeUserObject(emailMatches[0]);
      }
    }

    if (normalizedUser) {
      var userMatches = table.find({
        filter: function (row) {
          if (isSoftDeleted(row)) return false;
          if (excludeId && String(row.ID) === String(excludeId)) return false;
          return normalizeUserName(row.UserName || row.userName || row.Username) === normalizedUser;
        },
        limit: 1
      });
      if (userMatches && userMatches.length) {
        conflicts.userNameConflict = createSafeUserObject(userMatches[0]);
      }
    }

    conflicts.hasConflicts = !!(conflicts.emailConflict || conflicts.userNameConflict);
    return conflicts;
  }

  // ---------------------------------------------------------------------------
  // Mutations
  // ---------------------------------------------------------------------------

  function buildUserInsertPayload(input) {
    var userName = normalizeString(input.userName || input.UserName || input.username || '');
    var email = normalizeString(input.email || input.Email || '');
    if (!userName) {
      throw new Error('User name is required');
    }
    if (!email) {
      throw new Error('Email is required');
    }
    var employmentStatus = normalizeString(input.employmentStatus || input.EmploymentStatus || '');
    if (!isValidEmploymentStatus(employmentStatus)) {
      throw new Error('Invalid employment status value');
    }

    var payload = {
      UserName: userName,
      FullName: normalizeString(input.fullName || input.FullName || userName),
      Email: email,
      CampaignID: normalizeString(input.campaignId || input.CampaignID || input.campaignID || ''),
      PhoneNumber: normalizeString(input.phoneNumber || input.PhoneNumber || ''),
      EmploymentStatus: employmentStatus,
      HireDate: ensureDateValue(input.hireDate || input.HireDate || ''),
      Country: normalizeString(input.country || input.Country || ''),
      LockoutEnd: ensureDateValue(input.lockoutEnd || input.LockoutEnd || ''),
      TwoFactorEnabled: toBoolean(input.twoFactorEnabled || input.TwoFactorEnabled),
      CanLogin: toBoolean(input.canLogin != null ? input.canLogin : true),
      IsAdmin: toBoolean(input.isAdmin || input.IsAdmin),
      Roles: joinList(parseList(input.roles || input.Roles || [])),
      Pages: joinList(parseList(input.pages || input.Pages || [])),
      ResetRequired: toBoolean(input.resetRequired != null ? input.resetRequired : true),
      EmailConfirmation: normalizeString(input.emailConfirmation || input.EmailConfirmation || getUuid()),
      EmailConfirmed: toBoolean(input.emailConfirmed || input.EmailConfirmed),
      PasswordHash: normalizeString(input.passwordHash || input.PasswordHash || ''),
      DeletedAt: ''
    };

    return payload;
  }

  function createUser(input) {
    var table = getUsersTable();
    var payload = buildUserInsertPayload(input);
    var conflicts = checkUserConflicts(payload.Email, payload.UserName);
    if (conflicts.hasConflicts) {
      var parts = [];
      if (conflicts.emailConflict) parts.push('email already in use');
      if (conflicts.userNameConflict) parts.push('username already in use');
      throw new Error('Cannot create user: ' + parts.join(', '));
    }

    var inserted = table.insert(payload);

    if (input.roles) {
      try { syncRoles(inserted.ID, parseList(input.roles)); } catch (err) { logError('createUser:syncRoles', err); }
    }

    if (input.permissionLevel && input.campaignId) {
      try {
        setCampaignUserPermissions(input.campaignId, inserted.ID, input.permissionLevel, input.canManageUsers, input.canManagePages);
      } catch (err) {
        logError('createUser:setCampaignUserPermissions', err);
      }
    }

    if (payload.CanLogin && typeof global.sendPasswordSetupEmail === 'function') {
      try {
        global.sendPasswordSetupEmail(payload.Email, {
          userName: payload.UserName,
          fullName: payload.FullName,
          passwordSetupToken: payload.EmailConfirmation
        });
      } catch (err) {
        logError('createUser:sendPasswordSetupEmail', err);
      }
    }

    return createSafeUserObject(inserted);
  }

  function updateUser(userId, updates) {
    if (!userId && userId !== 0) {
      throw new Error('User ID is required');
    }
    var table = getUsersTable();
    var existing = table.findById(String(userId));
    if (!existing) {
      throw new Error('User not found');
    }

    var payload = {};
    if (Object.prototype.hasOwnProperty.call(updates, 'userName') || Object.prototype.hasOwnProperty.call(updates, 'UserName')) {
      payload.UserName = normalizeString(updates.userName || updates.UserName);
    }
    if (Object.prototype.hasOwnProperty.call(updates, 'fullName') || Object.prototype.hasOwnProperty.call(updates, 'FullName')) {
      payload.FullName = normalizeString(updates.fullName || updates.FullName);
    }
    if (Object.prototype.hasOwnProperty.call(updates, 'email') || Object.prototype.hasOwnProperty.call(updates, 'Email')) {
      payload.Email = normalizeString(updates.email || updates.Email);
    }
    if (Object.prototype.hasOwnProperty.call(updates, 'campaignId') || Object.prototype.hasOwnProperty.call(updates, 'CampaignID')) {
      payload.CampaignID = normalizeString(updates.campaignId || updates.CampaignID || updates.campaignID || '');
    }
    if (Object.prototype.hasOwnProperty.call(updates, 'phoneNumber') || Object.prototype.hasOwnProperty.call(updates, 'PhoneNumber')) {
      payload.PhoneNumber = normalizeString(updates.phoneNumber || updates.PhoneNumber || '');
    }
    if (Object.prototype.hasOwnProperty.call(updates, 'employmentStatus') || Object.prototype.hasOwnProperty.call(updates, 'EmploymentStatus')) {
      var status = normalizeString(updates.employmentStatus || updates.EmploymentStatus || '');
      if (!isValidEmploymentStatus(status)) {
        throw new Error('Invalid employment status value');
      }
      payload.EmploymentStatus = status;
    }
    if (Object.prototype.hasOwnProperty.call(updates, 'hireDate') || Object.prototype.hasOwnProperty.call(updates, 'HireDate')) {
      payload.HireDate = ensureDateValue(updates.hireDate || updates.HireDate || '');
    }
    if (Object.prototype.hasOwnProperty.call(updates, 'country') || Object.prototype.hasOwnProperty.call(updates, 'Country')) {
      payload.Country = normalizeString(updates.country || updates.Country || '');
    }
    if (Object.prototype.hasOwnProperty.call(updates, 'canLogin') || Object.prototype.hasOwnProperty.call(updates, 'CanLogin')) {
      payload.CanLogin = toBoolean(updates.canLogin != null ? updates.canLogin : updates.CanLogin);
    }
    if (Object.prototype.hasOwnProperty.call(updates, 'isAdmin') || Object.prototype.hasOwnProperty.call(updates, 'IsAdmin')) {
      payload.IsAdmin = toBoolean(updates.isAdmin != null ? updates.isAdmin : updates.IsAdmin);
    }
    if (Object.prototype.hasOwnProperty.call(updates, 'pages') || Object.prototype.hasOwnProperty.call(updates, 'Pages')) {
      payload.Pages = joinList(parseList(updates.pages || updates.Pages || []));
    }
    if (Object.prototype.hasOwnProperty.call(updates, 'roles') || Object.prototype.hasOwnProperty.call(updates, 'Roles')) {
      payload.Roles = joinList(parseList(updates.roles || updates.Roles || []));
    }

    if (payload.Email || payload.UserName) {
      var conflicts = checkUserConflicts(payload.Email || existing.Email, payload.UserName || existing.UserName, userId);
      if (conflicts.emailConflict || conflicts.userNameConflict) {
        throw new Error('Cannot update user: conflicts detected');
      }
    }

    if (Object.keys(payload).length) {
      table.update(String(userId), payload);
    }

    if (Object.prototype.hasOwnProperty.call(updates, 'roles') || Object.prototype.hasOwnProperty.call(updates, 'Roles')) {
      syncRoles(String(userId), parseList(updates.roles || updates.Roles || []));
    }

    if (Object.prototype.hasOwnProperty.call(updates, 'permissionLevel') || Object.prototype.hasOwnProperty.call(updates, 'canManageUsers')
      || Object.prototype.hasOwnProperty.call(updates, 'canManagePages')) {
      setCampaignUserPermissions(updates.campaignId || updates.CampaignID || existing.CampaignID,
        String(userId),
        updates.permissionLevel || (existing.PermissionLevel || 'USER'),
        updates.canManageUsers,
        updates.canManagePages);
    }

    var refreshed = table.findById(String(userId));
    return createSafeUserObject(refreshed);
  }

  function deleteUser(userId) {
    if (!userId && userId !== 0) {
      throw new Error('User ID is required');
    }
    var table = getUsersTable();
    var existing = table.findById(String(userId));
    if (!existing) {
      throw new Error('User not found');
    }
    if (isSoftDeleted(existing)) {
      return createSafeUserObject(existing);
    }
    table.update(String(userId), { DeletedAt: getNow() });
    return createSafeUserObject(table.findById(String(userId)));
  }

  function restoreUser(userId) {
    if (!userId && userId !== 0) {
      throw new Error('User ID is required');
    }
    var table = getUsersTable();
    var existing = table.findById(String(userId));
    if (!existing) {
      throw new Error('User not found');
    }
    table.update(String(userId), { DeletedAt: '' });
    return createSafeUserObject(table.findById(String(userId)));
  }

  function assignPagesToUser(userId, pages) {
    var table = getUsersTable();
    var existing = table.findById(String(userId));
    if (!existing) {
      throw new Error('User not found');
    }
    var joined = joinList(parseList(pages));
    table.update(String(userId), { Pages: joined });
    return createSafeUserObject(table.findById(String(userId)));
  }

  function getUserPages(userId) {
    var user = findRawUserById(userId);
    if (!user) {
      return [];
    }
    return parseList(user.Pages || user.pages || '');
  }

  // ---------------------------------------------------------------------------
  // Campaign permission helpers
  // ---------------------------------------------------------------------------

  function getCampaignUserPermissions(campaignId, userId) {
    var normalizedCampaign = normalizeString(campaignId);
    var normalizedUser = normalizeString(userId);
    if (!normalizedCampaign || !normalizedUser) {
      return { permissionLevel: 'USER', canManageUsers: false, canManagePages: false };
    }
    var table = getCampaignPermissionsTable();
    var match = table.find({
      filter: function (row) {
        if (isSoftDeleted(row)) return false;
        return normalizeString(row.CampaignID || row.CampaignId) === normalizedCampaign
          && normalizeString(row.UserID || row.UserId) === normalizedUser;
      },
      limit: 1
    });
    if (!match || !match.length) {
      return { permissionLevel: 'USER', canManageUsers: false, canManagePages: false };
    }
    var row = match[0];
    return {
      permissionLevel: String(row.PermissionLevel || 'USER').toUpperCase(),
      canManageUsers: toBoolean(row.CanManageUsers),
      canManagePages: toBoolean(row.CanManagePages),
      notes: row.Notes || ''
    };
  }

  function setCampaignUserPermissions(campaignId, userId, permissionLevel, canManageUsers, canManagePages) {
    var normalizedCampaign = normalizeString(campaignId);
    var normalizedUser = normalizeString(userId);
    if (!normalizedCampaign || !normalizedUser) {
      return { success: false, error: 'Campaign ID and User ID are required' };
    }
    var level = normalizeString(permissionLevel || 'USER').toUpperCase();
    var table = getCampaignPermissionsTable();
    var existing = table.find({
      filter: function (row) {
        return normalizeString(row.CampaignID || row.CampaignId) === normalizedCampaign
          && normalizeString(row.UserID || row.UserId) === normalizedUser;
      },
      limit: 1
    });

    var payload = {
      CampaignID: normalizedCampaign,
      UserID: normalizedUser,
      PermissionLevel: level,
      CanManageUsers: toBoolean(canManageUsers),
      CanManagePages: toBoolean(canManagePages),
      DeletedAt: ''
    };

    if (existing && existing.length) {
      table.update(existing[0].ID, payload);
    } else {
      table.insert(payload);
    }
    return { success: true };
  }

  // ---------------------------------------------------------------------------
  // Manager assignments
  // ---------------------------------------------------------------------------

  function getManagedUsers(managerUserId) {
    var managerId = normalizeString(managerUserId);
    if (!managerId) {
      return { success: true, managerUserId: managerUserId, users: [] };
    }
    var assignmentsTable = getManagerAssignmentsTable();
    var assignments = assignmentsTable.find({
      filter: function (row) {
        if (isSoftDeleted(row)) return false;
        return normalizeString(row.ManagerUserID || row.ManagerId) === managerId;
      }
    });
    var userIds = assignments.map(function (row) { return String(row.UserID || row.UserId); });
    var users = listUsers({ includeDeleted: false }).filter(function (user) {
      return userIds.indexOf(String(user.id)) !== -1;
    });
    return { success: true, managerUserId: managerUserId, users: users };
  }

  function getAvailableUsersForManager(managerUserId) {
    var assigned = getManagedUsers(managerUserId).users || [];
    var assignedIds = assigned.map(function (user) { return String(user.id); });
    var allUsers = listUsers({ includeDeleted: false });
    var available = allUsers.filter(function (user) {
      return assignedIds.indexOf(String(user.id)) === -1;
    });
    return { success: true, managerUserId: managerUserId, users: available };
  }

  function assignUsersToManager(managerUserId, userIds) {
    var managerId = normalizeString(managerUserId);
    if (!managerId) {
      throw new Error('Manager ID is required');
    }
    var assignmentsTable = getManagerAssignmentsTable();
    var existing = assignmentsTable.find({
      filter: function (row) {
        return normalizeString(row.ManagerUserID || row.ManagerId) === managerId;
      }
    });
    (existing || []).forEach(function (row) {
      if (row && row.ID) {
        assignmentsTable.delete(row.ID);
      }
    });

    var targetIds = ensureArray(userIds).map(function (id) { return normalizeString(id); }).filter(function (id) { return id; });
    targetIds.forEach(function (userId) {
      assignmentsTable.insert({ ManagerUserID: managerId, UserID: userId });
    });

    return getManagedUsers(managerId);
  }

  // ---------------------------------------------------------------------------
  // Password administration
  // ---------------------------------------------------------------------------

  function adminResetPassword(userId, requestingUserId) {
    var target = findRawUserById(userId);
    if (!target) {
      return { success: false, error: 'User not found' };
    }
    if (!toBoolean(target.CanLogin)) {
      return { success: false, error: 'User cannot login (CanLogin is FALSE)' };
    }
    if (requestingUserId) {
      var requester = findRawUserById(requestingUserId);
      if (!requester || !toBoolean(requester.IsAdmin)) {
        return { success: false, error: 'Only administrators can perform this action' };
      }
    }
    var token = getUuid();
    var table = getUsersTable();
    table.update(String(userId), { EmailConfirmation: token, ResetRequired: true });

    try {
      if (typeof global.sendAdminPasswordResetEmail === 'function') {
        global.sendAdminPasswordResetEmail(target.Email, { resetToken: token });
      } else if (typeof global.sendPasswordResetEmail === 'function') {
        global.sendPasswordResetEmail(target.Email, token);
      }
    } catch (err) {
      logError('adminResetPassword:sendEmail', err);
      return { success: false, error: 'Token saved, but failed to send email: ' + err };
    }

    return { success: true, message: 'Password reset email sent' };
  }

  function resendFirstLoginEmail(userId, requestingUserId) {
    var target = findRawUserById(userId);
    if (!target) {
      return { success: false, error: 'User not found' };
    }
    if (!toBoolean(target.CanLogin)) {
      return { success: false, error: 'User cannot login (CanLogin is FALSE)' };
    }
    if (requestingUserId) {
      var requester = findRawUserById(requestingUserId);
      if (!requester || !toBoolean(requester.IsAdmin)) {
        return { success: false, error: 'Only administrators can perform this action' };
      }
    }
    var token = getUuid();
    var table = getUsersTable();
    table.update(String(userId), { EmailConfirmation: token, ResetRequired: true });

    try {
      if (typeof global.sendFirstLoginResendEmail === 'function') {
        global.sendFirstLoginResendEmail(target.Email, {
          userName: target.UserName,
          fullName: target.FullName || target.UserName,
          passwordSetupToken: token
        });
      } else if (typeof global.sendPasswordSetupEmail === 'function') {
        global.sendPasswordSetupEmail(target.Email, {
          userName: target.UserName,
          fullName: target.FullName || target.UserName,
          passwordSetupToken: token
        });
      }
    } catch (err) {
      logError('resendFirstLoginEmail:sendEmail', err);
      return { success: false, error: 'Token saved, but failed to send email: ' + err };
    }

    return { success: true, message: 'First login email sent' };
  }

  // ---------------------------------------------------------------------------
  // Diagnostics & utilities
  // ---------------------------------------------------------------------------

  function runEnhancedDiscovery() {
    try {
      var users = listUsers({ includeDeleted: true });
      var duplicateEmails = {};
      var duplicateUserNames = {};
      var emailMap = {};
      var userMap = {};
      users.forEach(function (user) {
        var emailKey = normalizeEmail(user.email);
        if (emailKey) {
          if (!emailMap[emailKey]) {
            emailMap[emailKey] = [];
          }
          emailMap[emailKey].push(user.id);
        }
        var userKey = normalizeUserName(user.userName);
        if (userKey) {
          if (!userMap[userKey]) {
            userMap[userKey] = [];
          }
          userMap[userKey].push(user.id);
        }
      });
      Object.keys(emailMap).forEach(function (key) {
        if (emailMap[key].length > 1) {
          duplicateEmails[key] = emailMap[key];
        }
      });
      Object.keys(userMap).forEach(function (key) {
        if (userMap[key].length > 1) {
          duplicateUserNames[key] = userMap[key];
        }
      });
      return {
        success: true,
        totalUsers: users.length,
        duplicateEmails: duplicateEmails,
        duplicateUserNames: duplicateUserNames
      };
    } catch (err) {
      logError('runEnhancedDiscovery', err);
      return { success: false, error: err.message };
    }
  }

  function getAvailablePages() {
    try {
      var rows = readPagesSheet();
      if (!rows.length) {
        return [];
      }
      return rows.map(function (row) {
        return {
          key: row.Key || row.key || row.PageKey || row.pageKey || row.ID || row.Id || row.id || '',
          label: row.Label || row.label || row.Name || row.name || '',
          category: row.Category || row.category || row.Section || row.section || '',
          description: row.Description || row.description || ''
        };
      }).filter(function (page) { return page.key; });
    } catch (err) {
      logError('getAvailablePages', err);
      return [];
    }
  }

  // ---------------------------------------------------------------------------
  // Public facade
  // ---------------------------------------------------------------------------

  var UserDirectory = {
    list: listUsers,
    findById: findUserById,
    findByEmail: findUserByEmail,
    create: createUser,
    update: updateUser,
    remove: deleteUser,
    restore: restoreUser,
    assignPages: assignPagesToUser,
    getUserPages: getUserPages,
    getCampaignUserPermissions: getCampaignUserPermissions,
    setCampaignUserPermissions: setCampaignUserPermissions,
    getManagedUsers: getManagedUsers,
    getAvailableUsersForManager: getAvailableUsersForManager,
    assignUsersToManager: assignUsersToManager,
    adminResetPassword: adminResetPassword,
    resendFirstLoginEmail: resendFirstLoginEmail,
    runEnhancedDiscovery: runEnhancedDiscovery,
    getAvailablePages: getAvailablePages,
    createSafeUserObject: createSafeUserObject,
    checkConflicts: checkUserConflicts
  };

  global.UserDirectory = UserDirectory;

  // ---------------------------------------------------------------------------
  // Legacy global wrappers (HTML Service expects these names)
  // ---------------------------------------------------------------------------

  function clientGetAllUsers() {
    try {
      return UserDirectory.list({ includeDeleted: false });
    } catch (err) {
      logError('clientGetAllUsers', err);
      return [];
    }
  }

  function clientRegisterUser(userData) {
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(20000);
    } catch (err) {
      return { success: false, error: 'System is busy. Please try again.' };
    }
    try {
      var user = UserDirectory.create(userData || {});
      return { success: true, userId: user.id, user: user, message: 'User created successfully' };
    } catch (err) {
      logError('clientRegisterUser', err);
      return { success: false, error: err.message };
    } finally {
      try { lock.releaseLock(); } catch (_) { }
    }
  }

  function clientUpdateUser(userId, updates) {
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(20000);
    } catch (err) {
      return { success: false, error: 'System is busy. Please try again.' };
    }
    try {
      var user = UserDirectory.update(userId, updates || {});
      return { success: true, userId: userId, user: user, message: 'User updated successfully' };
    } catch (err) {
      logError('clientUpdateUser', err);
      return { success: false, error: err.message };
    } finally {
      try { lock.releaseLock(); } catch (_) { }
    }
  }

  function clientDeleteUser(userId) {
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(20000);
    } catch (err) {
      return { success: false, error: 'System is busy. Please try again.' };
    }
    try {
      var result = UserDirectory.remove(userId);
      return { success: true, userId: userId, user: result, message: 'User deactivated' };
    } catch (err) {
      logError('clientDeleteUser', err);
      return { success: false, error: err.message };
    } finally {
      try { lock.releaseLock(); } catch (_) { }
    }
  }

  function clientCheckUserConflicts(payload) {
    try {
      payload = payload || {};
      var conflicts = UserDirectory.checkConflicts(payload.email || payload.Email, payload.userName || payload.UserName, payload.excludeId || payload.excludeUserId || payload.userId || payload.ID);
      return { success: true, hasConflicts: conflicts.hasConflicts, conflicts: conflicts };
    } catch (err) {
      logError('clientCheckUserConflicts', err);
      return { success: false, error: err.message, conflicts: { emailConflict: null, userNameConflict: null } };
    }
  }

  function clientGetAvailablePages() {
    return UserDirectory.getAvailablePages();
  }

  function clientGetUserPages(userId) {
    try {
      return UserDirectory.getUserPages(userId);
    } catch (err) {
      logError('clientGetUserPages', err);
      return [];
    }
  }

  function clientAssignPagesToUser(userId, pageKeys) {
    try {
      var user = UserDirectory.assignPages(userId, pageKeys);
      return { success: true, user: user };
    } catch (err) {
      logError('clientAssignPagesToUser', err);
      return { success: false, error: err.message };
    }
  }

  function clientGetUserPermissions(userId, campaignId) {
    try {
      return UserDirectory.getCampaignUserPermissions(campaignId, userId);
    } catch (err) {
      logError('clientGetUserPermissions', err);
      return { permissionLevel: 'USER', canManageUsers: false, canManagePages: false };
    }
  }

  function clientSetUserPermissions(userId, campaignId, permissionLevel, canManageUsers, canManagePages) {
    try {
      return UserDirectory.setCampaignUserPermissions(campaignId, userId, permissionLevel, canManageUsers, canManagePages);
    } catch (err) {
      logError('clientSetUserPermissions', err);
      return { success: false, error: err.message };
    }
  }

  function clientGetAvailableUsersForManager(managerUserId) {
    try {
      return UserDirectory.getAvailableUsersForManager(managerUserId);
    } catch (err) {
      logError('clientGetAvailableUsersForManager', err);
      return { success: false, error: err.message, users: [] };
    }
  }

  function clientGetManagedUsers(managerUserId) {
    try {
      return UserDirectory.getManagedUsers(managerUserId);
    } catch (err) {
      logError('clientGetManagedUsers', err);
      return { success: false, error: err.message, users: [] };
    }
  }

  function clientAssignUsersToManager(managerUserId, userIds) {
    try {
      return UserDirectory.assignUsersToManager(managerUserId, userIds);
    } catch (err) {
      logError('clientAssignUsersToManager', err);
      return { success: false, error: err.message };
    }
  }

  function clientAdminResetPassword(userId, requestingUserId) {
    return UserDirectory.adminResetPassword(userId, requestingUserId);
  }

  function clientAdminResetPasswordById(userId, requestingUserId) {
    return clientAdminResetPassword(userId, requestingUserId);
  }

  function clientResendFirstLoginEmail(userId, requestingUserId) {
    return UserDirectory.resendFirstLoginEmail(userId, requestingUserId);
  }

  function clientRunEnhancedDiscovery() {
    return UserDirectory.runEnhancedDiscovery();
  }

  function clientGetValidEmploymentStatuses() {
    return { success: true, statuses: EMPLOYMENT_STATUS.slice() };
  }

  function clientGetEmploymentStatusReport(campaignId) {
    try {
      var users = UserDirectory.list({ includeDeleted: false, campaignId: campaignId });
      var counts = {};
      EMPLOYMENT_STATUS.forEach(function (status) { counts[status] = 0; });
      counts.Unspecified = 0;
      users.forEach(function (user) {
        var status = normalizeString(user.employmentStatus);
        if (!status) {
          counts.Unspecified++;
        } else if (counts.hasOwnProperty(status)) {
          counts[status]++;
        } else {
          counts.Unspecified++;
        }
      });
      return { success: true, campaignId: campaignId || '', totalUsers: users.length, statusCounts: counts };
    } catch (err) {
      logError('clientGetEmploymentStatusReport', err);
      return { success: false, error: err.message };
    }
  }

  // Re-export key helpers
  global.createSafeUserObject = createSafeUserObject;
  global.clientGetAllUsers = clientGetAllUsers;
  global.clientRegisterUser = clientRegisterUser;
  global.clientUpdateUser = clientUpdateUser;
  global.clientDeleteUser = clientDeleteUser;
  global.clientCheckUserConflicts = clientCheckUserConflicts;
  global.clientGetAvailablePages = clientGetAvailablePages;
  global.clientGetUserPages = clientGetUserPages;
  global.clientAssignPagesToUser = clientAssignPagesToUser;
  global.clientGetUserPermissions = clientGetUserPermissions;
  global.clientSetUserPermissions = clientSetUserPermissions;
  global.clientGetAvailableUsersForManager = clientGetAvailableUsersForManager;
  global.clientGetManagedUsers = clientGetManagedUsers;
  global.clientAssignUsersToManager = clientAssignUsersToManager;
  global.clientAdminResetPassword = clientAdminResetPassword;
  global.clientAdminResetPasswordById = clientAdminResetPasswordById;
  global.clientResendFirstLoginEmail = clientResendFirstLoginEmail;
  global.clientRunEnhancedDiscovery = clientRunEnhancedDiscovery;
  global.clientGetValidEmploymentStatuses = clientGetValidEmploymentStatuses;
  global.clientGetEmploymentStatusReport = clientGetEmploymentStatusReport;

})(this);

