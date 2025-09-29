/**
 * AuthenticationService.js
 * -----------------------------------------------------------------------------
 * Minimal authentication facade backed by the shared DatabaseManager / SheetsDB
 * adapter. This service is responsible for locating user records, validating
 * passwords and issuing session tokens that are stored in the Sessions table.
 *
 * The implementation favours clarity over cleverness so that it is easy to keep
 * in sync with Spreadsheet based tables as well as the typed SheetsDB tables
 * defined in `SheetsDatabaseBootstrap.js`.
 */

(function (global) {
  if (global.AuthenticationService) {
    return;
  }

  var SESSION_TTL_MS = 60 * 60 * 1000; // 1 hour
  var REMEMBER_ME_TTL_MS = 24 * 60 * 60 * 1000; // 24 hours
  var TIMEZONE = 'Etc/UTC';

  var USERS_TABLE_NAME = (typeof global.USERS_SHEET === 'string' && global.USERS_SHEET)
    ? global.USERS_SHEET
    : 'Users';
  var SESSIONS_TABLE_NAME = (typeof global.SESSIONS_SHEET === 'string' && global.SESSIONS_SHEET)
    ? global.SESSIONS_SHEET
    : 'Sessions';

  var AUTH_TABLE_SCHEMAS = [
    {
      name: USERS_TABLE_NAME,
      version: 1,
      primaryKey: 'ID',
      idPrefix: 'USR_',
      columns: [
        { name: 'ID', type: 'string', primaryKey: true },
        { name: 'UserName', type: 'string', nullable: true },
        { name: 'FullName', type: 'string', nullable: true },
        { name: 'Email', type: 'string', required: true, unique: true, maxLength: 320 },
        { name: 'CampaignID', type: 'string', nullable: true },
        { name: 'PasswordHash', type: 'string', required: true },
        { name: 'ResetRequired', type: 'boolean', defaultValue: false },
        { name: 'EmailConfirmation', type: 'string', nullable: true },
        { name: 'EmailConfirmed', type: 'boolean', defaultValue: false },
        { name: 'PhoneNumber', type: 'string', nullable: true },
        { name: 'EmploymentStatus', type: 'string', nullable: true },
        { name: 'HireDate', type: 'date', nullable: true },
        { name: 'Country', type: 'string', nullable: true },
        { name: 'LockoutEnd', type: 'timestamp', nullable: true },
        { name: 'TwoFactorEnabled', type: 'boolean', defaultValue: false },
        { name: 'CanLogin', type: 'boolean', defaultValue: true },
        { name: 'Roles', type: 'string', nullable: true },
        { name: 'Pages', type: 'string', nullable: true },
        { name: 'CreatedAt', type: 'timestamp', required: true },
        { name: 'UpdatedAt', type: 'timestamp', required: true },
        { name: 'LastLogin', type: 'timestamp', nullable: true },
        { name: 'DeletedAt', type: 'timestamp', nullable: true },
        { name: 'IsAdmin', type: 'boolean', defaultValue: false }
      ],
      indexes: [
        { name: USERS_TABLE_NAME + '_Email_idx', field: 'Email', unique: true },
        { name: USERS_TABLE_NAME + '_Campaign_idx', field: 'CampaignID' }
      ]
    },
    {
      name: SESSIONS_TABLE_NAME,
      version: 1,
      primaryKey: 'Token',
      idPrefix: 'SES_',
      columns: [
        { name: 'Token', type: 'string', primaryKey: true },
        { name: 'UserId', type: 'string', required: true },
        { name: 'CreatedAt', type: 'timestamp', required: true },
        { name: 'UpdatedAt', type: 'timestamp', required: true },
        { name: 'ExpiresAt', type: 'timestamp', required: true },
        { name: 'RememberMe', type: 'boolean', defaultValue: false },
        { name: 'CampaignScope', type: 'json', nullable: true },
        { name: 'ClientContext', type: 'json', nullable: true },
        { name: 'UserAgent', type: 'string', nullable: true },
        { name: 'IpAddress', type: 'string', nullable: true },
        { name: 'DeletedAt', type: 'timestamp', nullable: true }
      ],
      indexes: [
        { name: SESSIONS_TABLE_NAME + '_User_idx', field: 'UserId' },
        { name: SESSIONS_TABLE_NAME + '_Expiry_idx', field: 'ExpiresAt' }
      ],
      retentionDays: 45
    }
  ];

  var passwordUtils = null;
  function getPasswordUtils() {
    if (!passwordUtils) {
      if (typeof global.ensurePasswordUtilities === 'function') {
        passwordUtils = global.ensurePasswordUtilities();
      } else if (typeof global.PasswordUtilities !== 'undefined') {
        passwordUtils = global.PasswordUtilities;
      }
    }
    if (!passwordUtils) {
      throw new Error('Password utilities are not available');
    }
    return passwordUtils;
  }

  var tablesInitialised = false;
  function ensureTables() {
    if (tablesInitialised) {
      return;
    }
    if (typeof global.DatabaseManager === 'undefined' || !global.DatabaseManager) {
      throw new Error('DatabaseManager is not available');
    }

    var userHeaders = AUTH_TABLE_SCHEMAS[0].columns.map(function (column) { return column.name; });
    var sessionHeaders = AUTH_TABLE_SCHEMAS[1].columns.map(function (column) { return column.name; });

    global.DatabaseManager.defineTable(USERS_TABLE_NAME, {
      headers: userHeaders,
      idColumn: 'ID',
      timestamps: { created: 'CreatedAt', updated: 'UpdatedAt' }
    });

    global.DatabaseManager.defineTable(SESSIONS_TABLE_NAME, {
      headers: sessionHeaders,
      idColumn: 'Token',
      timestamps: { created: 'CreatedAt', updated: 'UpdatedAt' }
    });

    tablesInitialised = true;
  }

  function getUsersTable() {
    ensureTables();
    return global.DatabaseManager.table(USERS_TABLE_NAME);
  }

  function getSessionsTable() {
    ensureTables();
    return global.DatabaseManager.table(SESSIONS_TABLE_NAME);
  }

  function normalizeEmail(email) {
    if (email === null || typeof email === 'undefined') {
      return '';
    }
    return String(email).trim().toLowerCase();
  }

  function normalizeString(value) {
    if (value === null || typeof value === 'undefined') {
      return '';
    }
    return String(value).trim();
  }

  function toBoolean(value) {
    if (value === true || value === false) {
      return value;
    }
    var normalized = normalizeString(value).toLowerCase();
    if (!normalized) {
      return false;
    }
    return normalized === 'true' || normalized === '1' || normalized === 'yes' || normalized === 'y';
  }

  function nowIso() {
    return Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss'Z'");
  }

  function futureIso(offsetMs) {
    return Utilities.formatDate(new Date(Date.now() + offsetMs), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss'Z'");
  }

  function cloneRecord(record) {
    if (!record || typeof record !== 'object') {
      return record;
    }
    var copy = {};
    Object.keys(record).forEach(function (key) {
      copy[key] = record[key];
    });
    return copy;
  }

  function sanitizeUser(user) {
    if (!user) {
      return null;
    }
    var sanitized = cloneRecord(user);
    delete sanitized.PasswordHash;
    return sanitized;
  }

  function isEmailConfirmed(user) {
    if (!user) {
      return false;
    }
    if (!Object.prototype.hasOwnProperty.call(user, 'EmailConfirmed')) {
      return true;
    }
    return toBoolean(user.EmailConfirmed);
  }

  function invalidateUsersCache() {
    try {
      if (typeof global.invalidateCache === 'function') {
        global.invalidateCache(USERS_TABLE_NAME);
      }
    } catch (err) {
      console.warn('invalidateUsersCache: failed to invalidate cache', err);
    }
  }

  function findUserByResetToken(token) {
    var normalized = normalizeString(token);
    if (!normalized) {
      return null;
    }

    var table = getUsersTable();
    var user = null;

    if (typeof table.findOne === 'function') {
      user = table.findOne({ EmailConfirmation: normalized });
    }

    if (!user && typeof table.find === 'function') {
      var results = table.find({ where: { EmailConfirmation: normalized }, limit: 5 }) || [];
      user = results.length ? results[0] : null;
    }

    if (!user && typeof table.read === 'function') {
      var rows = table.read({ limit: 2000 }) || [];
      for (var i = 0; i < rows.length; i++) {
        if (normalizeString(rows[i].EmailConfirmation) === normalized) {
          user = rows[i];
          break;
        }
      }
    }

    if (user && user.Email) {
      user.Email = normalizeEmail(user.Email);
    }

    return user;
  }

  function ensureResetTokenForUser(user) {
    if (!user || !user.ID) {
      throw new Error('User record is missing an ID');
    }

    var table = getUsersTable();
    var token = normalizeString(user.EmailConfirmation);
    var needsUpdate = false;
    var updates = { UpdatedAt: nowIso() };

    if (!token) {
      token = Utilities.getUuid();
      updates.EmailConfirmation = token;
      needsUpdate = true;
    }

    if (!toBoolean(user.ResetRequired)) {
      updates.ResetRequired = true;
      needsUpdate = true;
    }

    if (needsUpdate) {
      if (typeof table.update !== 'function') {
        throw new Error('Users table does not support updates');
      }
      table.update(String(user.ID), updates);
      invalidateUsersCache();
      user.ResetRequired = true;
    }

    user.EmailConfirmation = token;

    return token;
  }

  function sendPasswordResetNotification(user, token) {
    if (!user || !user.Email || !token) {
      return false;
    }

    if (typeof global.sendPasswordResetEmail === 'function') {
      try {
        return global.sendPasswordResetEmail(user.Email, token) !== false;
      } catch (err) {
        throw err;
      }
    }

    if (typeof global.sendAdminPasswordResetEmail === 'function') {
      try {
        global.sendAdminPasswordResetEmail(user.Email, { resetToken: token });
        return true;
      } catch (err) {
        throw err;
      }
    }

    return false;
  }

  function sendPasswordSetupNotification(user, token) {
    if (!user || !user.Email || !token) {
      return false;
    }

    var payload = {
      userName: user.UserName || user.Email,
      fullName: user.FullName || user.UserName || user.Email,
      passwordSetupToken: token
    };

    if (typeof global.sendFirstLoginResendEmail === 'function') {
      try {
        return global.sendFirstLoginResendEmail(user.Email, payload) !== false;
      } catch (err) {
        throw err;
      }
    }

    if (typeof global.sendPasswordSetupEmail === 'function') {
      try {
        return global.sendPasswordSetupEmail(user.Email, payload) !== false;
      } catch (err) {
        throw err;
      }
    }

    return sendPasswordResetNotification(user, token);
  }

  function sendPasswordChangeConfirmation(user) {
    if (!user || !user.Email) {
      return;
    }

    if (typeof global.sendPasswordChangeConfirmation === 'function') {
      try {
        global.sendPasswordChangeConfirmation(user.Email, { timestamp: nowIso() });
      } catch (err) {
        console.warn('sendPasswordChangeConfirmation: failed to send confirmation', err);
      }
    }
  }

  function isSoftDeleted(record) {
    if (!record) {
      return false;
    }
    return !!normalizeString(record.DeletedAt);
  }

  function canLogin(user) {
    if (!user) {
      return false;
    }
    if (isSoftDeleted(user)) {
      return false;
    }
    if (Object.prototype.hasOwnProperty.call(user, 'CanLogin')) {
      return toBoolean(user.CanLogin);
    }
    return true;
  }

  function findUserByEmail(email) {
    var normalized = normalizeEmail(email);
    if (!normalized) {
      return null;
    }

    var table = getUsersTable();
    var user = null;
    if (typeof table.findOne === 'function') {
      user = table.findOne({ Email: normalized });
    }

    if (!user && typeof table.find === 'function') {
      var results = table.find({ where: { Email: normalized }, limit: 5 }) || [];
      user = results.length ? results[0] : null;
    }

    if (!user && typeof table.read === 'function') {
      var rows = table.read({ limit: 2000 }) || [];
      for (var i = 0; i < rows.length; i++) {
        if (normalizeEmail(rows[i].Email) === normalized) {
          user = rows[i];
          break;
        }
      }
    }

    if (user && user.Email) {
      user.Email = normalizeEmail(user.Email);
    }

    return user;
  }

  function findUserById(id) {
    if (!id) {
      return null;
    }
    var table = getUsersTable();
    if (typeof table.findById === 'function') {
      return table.findById(id);
    }
    if (typeof table.findOne === 'function') {
      return table.findOne({ ID: id });
    }
    var rows = (typeof table.read === 'function') ? table.read({ limit: 2000 }) : [];
    for (var i = 0; i < rows.length; i++) {
      if (normalizeString(rows[i].ID) === normalizeString(id)) {
        return rows[i];
      }
    }
    return null;
  }

  function verifyUserPassword(user, rawPassword) {
    if (!user) {
      return false;
    }
    var hash = user.PasswordHash;
    if (!hash) {
      return false;
    }
    var utils = getPasswordUtils();
    return utils.verifyPassword(rawPassword, hash);
  }

  function computeSessionTtl(rememberMe) {
    return rememberMe ? REMEMBER_ME_TTL_MS : SESSION_TTL_MS;
  }

  function buildSessionRecord(userId, rememberMe, campaignScope, metadata) {
    var token = Utilities.getUuid();
    var createdAt = nowIso();
    var ttlMs = computeSessionTtl(rememberMe);
    var expiresAt = futureIso(ttlMs);

    var record = {
      Token: token,
      UserId: userId,
      CreatedAt: createdAt,
      UpdatedAt: createdAt,
      ExpiresAt: expiresAt,
      RememberMe: !!rememberMe,
      DeletedAt: ''
    };

    if (campaignScope) {
      record.CampaignScope = campaignScope;
    }

    if (metadata && typeof metadata === 'object') {
      record.ClientContext = metadata;
      if (metadata.userAgent) {
        record.UserAgent = metadata.userAgent;
      }
      if (metadata.ipAddress) {
        record.IpAddress = metadata.ipAddress;
      }
    }

    return { record: record, ttlMs: ttlMs };
  }

  function persistSession(record) {
    var table = getSessionsTable();
    if (typeof table.insert === 'function') {
      return table.insert(record);
    }
    if (typeof table.upsert === 'function') {
      return table.upsert({ Token: record.Token }, record);
    }
    throw new Error('Sessions table does not support inserts');
  }

  function attachSessionMeta(user, session) {
    if (!user || !session) {
      return null;
    }
    var sanitized = sanitizeUser(user);
    sanitized.sessionToken = session.Token;
    sanitized.sessionExpiresAt = session.ExpiresAt;
    sanitized.sessionTtlSeconds = Math.max(0, Math.floor((Date.parse(session.ExpiresAt) - Date.now()) / 1000));
    return sanitized;
  }

  function loadActiveSession(token) {
    var normalized = normalizeString(token);
    if (!normalized) {
      return null;
    }

    var table = getSessionsTable();
    var session = null;
    if (typeof table.findById === 'function') {
      session = table.findById(normalized);
    } else if (typeof table.findOne === 'function') {
      session = table.findOne({ Token: normalized });
    }

    if (!session) {
      return null;
    }

    if (isSoftDeleted(session)) {
      return null;
    }

    if (session.ExpiresAt) {
      var expires = Date.parse(session.ExpiresAt);
      if (!isNaN(expires) && expires < Date.now()) {
        expireSession(session.Token);
        return null;
      }
    }

    var user = findUserById(session.UserId);
    if (!user || !canLogin(user)) {
      return null;
    }

    return { session: session, user: user };
  }

  function expireSession(token) {
    if (!token) {
      return;
    }
    try {
      var table = getSessionsTable();
      if (typeof table.update === 'function') {
        table.update(token, { DeletedAt: nowIso() });
      } else if (typeof table.delete === 'function') {
        table.delete(token);
      }
    } catch (err) {
      console.warn('expireSession: failed to expire session', err);
    }
  }

  function refreshSession(session) {
    if (!session) {
      return null;
    }
    var ttlMs = computeSessionTtl(toBoolean(session.RememberMe));
    var updates = {
      UpdatedAt: nowIso(),
      ExpiresAt: futureIso(ttlMs)
    };

    var table = getSessionsTable();
    if (typeof table.update === 'function') {
      table.update(session.Token, updates);
      session.UpdatedAt = updates.UpdatedAt;
      session.ExpiresAt = updates.ExpiresAt;
    }

    return session;
  }

  function updateLastLogin(userId) {
    if (!userId) {
      return;
    }
    try {
      var table = getUsersTable();
      if (typeof table.update === 'function') {
        table.update(userId, { LastLogin: nowIso(), UpdatedAt: nowIso() });
      }
    } catch (err) {
      console.warn('updateLastLogin: unable to update last login', err);
    }
  }

  function login(email, password, rememberMe, clientMetadata) {
    try {
      ensureTables();
    } catch (err) {
      return {
        success: false,
        error: err.message || 'Authentication system unavailable',
        errorCode: 'AUTH_UNAVAILABLE'
      };
    }

    var normalizedEmail = normalizeEmail(email);
    if (!normalizedEmail) {
      return { success: false, error: 'Email is required', errorCode: 'EMAIL_REQUIRED' };
    }

    if (!password && password !== 0) {
      return { success: false, error: 'Password is required', errorCode: 'PASSWORD_REQUIRED' };
    }

    var user = findUserByEmail(normalizedEmail);
    if (!user) {
      return { success: false, error: 'Invalid email or password', errorCode: 'INVALID_CREDENTIALS' };
    }

    if (!canLogin(user)) {
      return { success: false, error: 'Account is disabled', errorCode: 'ACCOUNT_DISABLED' };
    }

    var existingHash = normalizeString(user.PasswordHash);
    if (!existingHash) {
      var setupToken;
      try {
        setupToken = ensureResetTokenForUser(user);
      } catch (tokenErr) {
        console.error('login: failed to ensure setup token', tokenErr);
        if (typeof global.writeError === 'function') {
          global.writeError('login.ensureResetToken', tokenErr);
        }
      }
      return {
        success: false,
        error: 'You need to set a password before signing in. Check your email for the activation link.',
        errorCode: 'PASSWORD_NOT_SET',
        resetToken: setupToken || normalizeString(user.EmailConfirmation) || ''
      };
    }

    if (!verifyUserPassword(user, password)) {
      return { success: false, error: 'Invalid email or password', errorCode: 'INVALID_CREDENTIALS' };
    }

    if (!isEmailConfirmed(user)) {
      return { success: false, error: 'Please confirm your email address before signing in.', errorCode: 'EMAIL_NOT_CONFIRMED' };
    }

    if (toBoolean(user.ResetRequired)) {
      var resetToken;
      try {
        resetToken = ensureResetTokenForUser(user);
      } catch (resetErr) {
        console.error('login: failed to ensure reset token', resetErr);
        if (typeof global.writeError === 'function') {
          global.writeError('login.ensureResetToken', resetErr);
        }
      }
      return {
        success: false,
        error: 'Password reset required before signing in.',
        errorCode: 'PASSWORD_RESET_REQUIRED',
        resetToken: resetToken || normalizeString(user.EmailConfirmation) || ''
      };
    }

    var sessionEnvelope = buildSessionRecord(user.ID, !!rememberMe, null, clientMetadata);
    var persisted;
    try {
      persisted = persistSession(sessionEnvelope.record);
    } catch (err) {
      console.error('login: failed to persist session', err);
      return { success: false, error: 'Unable to create session', errorCode: 'SESSION_ERROR' };
    }

    updateLastLogin(user.ID);

    var sessionRecord = persisted && persisted.Token ? persisted : sessionEnvelope.record;
    var sanitizedUser = attachSessionMeta(user, sessionRecord);

    return {
      success: true,
      message: 'Login successful',
      sessionToken: sessionRecord.Token,
      sessionExpiresAt: sessionRecord.ExpiresAt,
      sessionTtlSeconds: Math.floor(sessionEnvelope.ttlMs / 1000),
      rememberMe: !!rememberMe,
      user: sanitizedUser,
      redirectUrl: null
    };
  }

  function createSessionFor(userId, campaignScope, rememberMe, metadata) {
    ensureTables();
    var user = typeof userId === 'object' ? userId : findUserById(userId);
    if (!user) {
      return { success: false, error: 'User not found', errorCode: 'USER_NOT_FOUND' };
    }
    if (!canLogin(user)) {
      return { success: false, error: 'Account is disabled', errorCode: 'ACCOUNT_DISABLED' };
    }

    var sessionEnvelope = buildSessionRecord(user.ID, !!rememberMe, campaignScope || null, metadata);
    var persisted;
    try {
      persisted = persistSession(sessionEnvelope.record);
    } catch (err) {
      console.error('createSessionFor: failed to persist session', err);
      return { success: false, error: 'Unable to create session', errorCode: 'SESSION_ERROR' };
    }

    var sessionRecord = persisted && persisted.Token ? persisted : sessionEnvelope.record;
    var sanitizedUser = attachSessionMeta(user, sessionRecord);

    return {
      success: true,
      sessionToken: sessionRecord.Token,
      sessionExpiresAt: sessionRecord.ExpiresAt,
      sessionTtlSeconds: Math.floor(sessionEnvelope.ttlMs / 1000),
      rememberMe: !!rememberMe,
      user: sanitizedUser,
      campaignScope: campaignScope || null
    };
  }

  function getSessionUser(token) {
    var payload = loadActiveSession(token);
    if (!payload) {
      return null;
    }
    var refreshed = refreshSession(payload.session);
    return attachSessionMeta(payload.user, refreshed || payload.session);
  }

  function validateToken(token) {
    return getSessionUser(token);
  }

  function logout(token) {
    if (!token) {
      return { success: true, message: 'No active session' };
    }
    expireSession(token);
    return { success: true, message: 'Logged out' };
  }

  function keepAlive(token) {
    var payload = loadActiveSession(token);
    if (!payload) {
      return { success: false, expired: true, message: 'Session expired or not found' };
    }
    var refreshed = refreshSession(payload.session);
    var sanitized = attachSessionMeta(payload.user, refreshed || payload.session);
    return {
      success: true,
      message: 'Session active',
      sessionToken: sanitized.sessionToken,
      sessionExpiresAt: sanitized.sessionExpiresAt,
      sessionTtlSeconds: sanitized.sessionTtlSeconds,
      user: sanitized
    };
  }

  function getTableSchemas() {
    return AUTH_TABLE_SCHEMAS.map(function (schema) {
      return JSON.parse(JSON.stringify(schema));
    });
  }

  function requestPasswordReset(email) {
    try {
      ensureTables();
    } catch (err) {
      return {
        success: false,
        error: 'Authentication system unavailable',
        errorCode: 'AUTH_UNAVAILABLE'
      };
    }

    var normalizedEmail = normalizeEmail(email);
    if (!normalizedEmail) {
      return { success: false, error: 'Email is required', errorCode: 'EMAIL_REQUIRED' };
    }

    var user = findUserByEmail(normalizedEmail);
    var genericMessage = 'If an account with that email exists, a password reset link has been sent.';

    if (!user) {
      return { success: true, message: genericMessage };
    }

    if (!canLogin(user)) {
      return { success: true, message: genericMessage };
    }

    var lock = null;
    if (typeof LockService !== 'undefined' && LockService) {
      try {
        lock = LockService.getScriptLock();
        lock.waitLock(20000);
      } catch (err) {
        return { success: false, error: 'System is busy. Please try again shortly.', errorCode: 'LOCK_TIMEOUT' };
      }
    }

    try {
      var token = ensureResetTokenForUser(user);
      if (!sendPasswordResetNotification(user, token)) {
        return {
          success: false,
          error: 'Unable to send password reset email. Please contact support.',
          errorCode: 'EMAIL_FAILURE'
        };
      }
      return { success: true, message: genericMessage };
    } catch (err) {
      console.error('requestPasswordReset failed', err);
      if (typeof global.writeError === 'function') {
        global.writeError('requestPasswordReset', err);
      }
      return {
        success: false,
        error: 'Unable to process the password reset request. Please try again later.',
        errorCode: 'RESET_FAILED'
      };
    } finally {
      if (lock) {
        try { lock.releaseLock(); } catch (_) { }
      }
    }
  }

  function resendPasswordSetupEmail(email) {
    try {
      ensureTables();
    } catch (err) {
      return {
        success: false,
        error: 'Authentication system unavailable',
        errorCode: 'AUTH_UNAVAILABLE'
      };
    }

    var normalizedEmail = normalizeEmail(email);
    if (!normalizedEmail) {
      return { success: false, error: 'Email is required', errorCode: 'EMAIL_REQUIRED' };
    }

    var user = findUserByEmail(normalizedEmail);
    var genericMessage = 'If an account with that email exists, setup instructions have been sent.';

    if (!user) {
      return { success: true, message: genericMessage };
    }

    if (!canLogin(user)) {
      return { success: true, message: genericMessage };
    }

    // If a password already exists, fall back to the reset flow.
    if (normalizeString(user.PasswordHash)) {
      return requestPasswordReset(normalizedEmail);
    }

    var lock = null;
    if (typeof LockService !== 'undefined' && LockService) {
      try {
        lock = LockService.getScriptLock();
        lock.waitLock(20000);
      } catch (err) {
        return { success: false, error: 'System is busy. Please try again shortly.', errorCode: 'LOCK_TIMEOUT' };
      }
    }

    try {
      var token = ensureResetTokenForUser(user);
      if (!sendPasswordSetupNotification(user, token)) {
        return {
          success: false,
          error: 'Unable to send the setup email. Please contact support.',
          errorCode: 'EMAIL_FAILURE'
        };
      }
      return { success: true, message: genericMessage };
    } catch (err) {
      console.error('resendPasswordSetupEmail failed', err);
      if (typeof global.writeError === 'function') {
        global.writeError('resendPasswordSetupEmail', err);
      }
      return {
        success: false,
        error: 'Unable to resend the setup email. Please try again later.',
        errorCode: 'SETUP_FAILED'
      };
    } finally {
      if (lock) {
        try { lock.releaseLock(); } catch (_) { }
      }
    }
  }

  function setPasswordWithToken(token, newPassword) {
    try {
      ensureTables();
    } catch (err) {
      return {
        success: false,
        error: 'Authentication system unavailable',
        errorCode: 'AUTH_UNAVAILABLE'
      };
    }

    var normalizedToken = normalizeString(token);
    if (!normalizedToken) {
      return { success: false, error: 'Invalid or missing token.', errorCode: 'TOKEN_MISSING' };
    }

    var normalizedPassword = newPassword == null ? '' : String(newPassword);
    if (normalizedPassword.length < 8) {
      return { success: false, error: 'Password must be at least 8 characters long.', errorCode: 'PASSWORD_TOO_SHORT' };
    }

    var user = findUserByResetToken(normalizedToken);
    if (!user) {
      return { success: false, error: 'Reset link is invalid or has expired.', errorCode: 'TOKEN_INVALID' };
    }

    var lock = null;
    if (typeof LockService !== 'undefined' && LockService) {
      try {
        lock = LockService.getScriptLock();
        lock.waitLock(20000);
      } catch (err) {
        return { success: false, error: 'System is busy. Please try again shortly.', errorCode: 'LOCK_TIMEOUT' };
      }
    }

    try {
      var utils = getPasswordUtils();
      var hash = utils.createPasswordHash(normalizedPassword);
      var table = getUsersTable();
      if (typeof table.update !== 'function') {
        throw new Error('Users table does not support updates');
      }

      table.update(String(user.ID), {
        PasswordHash: hash,
        ResetRequired: false,
        EmailConfirmation: '',
        EmailConfirmed: true,
        UpdatedAt: nowIso()
      });
      invalidateUsersCache();
      sendPasswordChangeConfirmation(user);
      return {
        success: true,
        message: 'Your password has been updated. You may now sign in.'
      };
    } catch (err) {
      console.error('setPasswordWithToken failed', err);
      if (typeof global.writeError === 'function') {
        global.writeError('setPasswordWithToken', err);
      }
      return {
        success: false,
        error: 'Unable to update your password. Please try again later.',
        errorCode: 'RESET_FAILED'
      };
    } finally {
      if (lock) {
        try { lock.releaseLock(); } catch (_) { }
      }
    }
  }

  function ensureSheets() {
    ensureTables();
  }

  global.AuthenticationService = {
    login: login,
    logout: logout,
    keepAlive: keepAlive,
    createSessionFor: createSessionFor,
    getSessionUser: getSessionUser,
    validateToken: validateToken,
    ensureSheets: ensureSheets,
    findUserByEmail: findUserByEmail,
    findUserById: findUserById,
    getUserByEmail: findUserByEmail,
    verifyUserPassword: verifyUserPassword,
    getTableSchemas: getTableSchemas,
    requestPasswordReset: requestPasswordReset,
    resendPasswordSetupEmail: resendPasswordSetupEmail,
    setPasswordWithToken: setPasswordWithToken
  };

})(typeof globalThis !== 'undefined' ? globalThis : this);

function loginUser(email, password, rememberMe, clientMetadata) {
  return AuthenticationService.login(email, password, rememberMe, clientMetadata);
}

function logoutUser(sessionToken) {
  return AuthenticationService.logout(sessionToken);
}

function keepAliveSession(sessionToken) {
  return AuthenticationService.keepAlive(sessionToken);
}

function login(email, password, rememberMe, clientMetadata) {
  return AuthenticationService.login(email, password, rememberMe, clientMetadata);
}

function requestPasswordReset(email) {
  return AuthenticationService.requestPasswordReset(email);
}

function resendPasswordSetupEmail(email) {
  return AuthenticationService.resendPasswordSetupEmail(email);
}

function setPasswordWithToken(token, newPassword) {
  return AuthenticationService.setPasswordWithToken(token, newPassword);
}

function logout(sessionToken) {
  return AuthenticationService.logout(sessionToken);
}

function keepAlive(sessionToken) {
  return AuthenticationService.keepAlive(sessionToken);
}
