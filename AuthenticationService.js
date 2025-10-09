/**
 * AuthenticationService.js
 * -----------------------------------------------------------------------------
 * Rebuilt authentication, credential and session management service.  The
 * previous implementation grew into a 5,000+ line monolith with tightly coupled
 * helpers, ad-hoc token handling and multiple redundant code paths.  This file
 * provides a clean room implementation whose sole responsibility is to manage
 * user credentials, login flows and session state.  Every exported method is
 * designed to be composable so UserService, IdentityService and UI surfaces can
 * collaborate without sharing internal details.
 */

(function (global) {
  var SESSION_TTL_MS = 30 * 60 * 1000;            // 30 minutes
  var REMEMBER_ME_TTL_MS = 48 * 60 * 60 * 1000;   // 48 hours
  var DEFAULT_IDLE_TIMEOUT_MINUTES = 30;
  var PASSWORD_RESET_TTL_MINUTES = 120;

  var USERS_TABLE = 'Users';
  var CREDENTIALS_TABLE = 'UserCredentials';
  var SESSIONS_TABLE = 'Sessions';
  var PASSWORD_RESET_TABLE = 'PasswordResets';

  var LOGIN_CONTEXT_CACHE_KEY = 'AUTH:LOGIN_CONTEXT';
  var LOGIN_CONTEXT_CACHE_TTL_SECONDS = 300; // 5 minutes

  var PASSWORD_PURPOSE_SETUP = 'setup';
  var PASSWORD_PURPOSE_RESET = 'reset';

  function requirePasswordUtils() {
    if (typeof ensurePasswordUtilities === 'function') {
      return ensurePasswordUtilities();
    }
    if (global.PasswordUtilities) {
      return global.PasswordUtilities;
    }
    throw new Error('Password utilities are unavailable. Ensure PasswordUtilities.js is loaded.');
  }

  function normalizeString(value) {
    if (value === null || typeof value === 'undefined') return '';
    return String(value).trim();
  }

  function normalizeEmail(email) {
    return normalizeString(email).toLowerCase();
  }

  function sheetBoolean(value) {
    return value ? 'TRUE' : 'FALSE';
  }

  function parseBooleanFlag(value) {
    if (value === true || value === false) return !!value;
    if (value === null || typeof value === 'undefined') return false;
    var str = String(value).trim().toLowerCase();
    return str === 'true' || str === '1' || str === 'yes';
  }

  function now() {
    return new Date();
  }

  function iso(date) {
    return (date instanceof Date) ? date.toISOString() : new Date(date).toISOString();
  }

  function toNumber(value, fallback) {
    var num = Number(value);
    return isFinite(num) ? num : (typeof fallback === 'number' ? fallback : 0);
  }

  function clone(obj) {
    if (!obj || typeof obj !== 'object') return obj;
    var out = {};
    Object.keys(obj).forEach(function (key) { out[key] = obj[key]; });
    return out;
  }

  function createMetadataBlob(metadata) {
    if (!metadata || typeof metadata !== 'object') return '';
    try {
      return JSON.stringify(metadata);
    } catch (err) {
      return '';
    }
  }

  function parseMetadataBlob(blob) {
    if (!blob) return null;
    try {
      return JSON.parse(blob);
    } catch (_) {
      return null;
    }
  }

  function requireDatabaseManager() {
    if (!global.DatabaseManager || typeof global.DatabaseManager.defineTable !== 'function') {
      throw new Error('DatabaseManager is required for AuthenticationService');
    }
    return global.DatabaseManager;
  }

  var AuthenticationService = (function () {
    var tables = {
      users: null,
      credentials: null,
      sessions: null,
      resets: null
    };

    // -------------------------------------------------------------------------
    // Table helpers
    // -------------------------------------------------------------------------

    function ensureUsersTable() {
      if (tables.users) return tables.users;
      var db = requireDatabaseManager();
      tables.users = db.defineTable(USERS_TABLE, {
        idColumn: 'ID',
        headers: [
          'ID', 'UserName', 'FullName', 'Email', 'NormalizedEmail', 'CanLogin', 'IsAdmin',
          'Roles', 'Pages', 'SecurityStamp', 'ConcurrencyStamp', 'LastLoginAt',
          'ResetRequired', 'ResetPasswordToken', 'ResetPasswordTokenHash',
          'ResetPasswordSentAt', 'ResetPasswordExpiresAt', 'UpdatedAt', 'CreatedAt'
        ],
        timestamps: { created: 'CreatedAt', updated: 'UpdatedAt' }
      });
      return tables.users;
    }

    function ensureCredentialTable() {
      if (tables.credentials) return tables.credentials;
      var db = requireDatabaseManager();
      tables.credentials = db.defineTable(CREDENTIALS_TABLE, {
        idColumn: 'ID',
        headers: [
          'ID', 'UserId', 'PasswordHash', 'PasswordSalt', 'PasswordIterations',
          'PasswordAlgorithm', 'PasswordVersion', 'PasswordUpdatedAt',
          'MustChangePassword', 'CreatedAt', 'UpdatedAt'
        ],
        timestamps: { created: 'CreatedAt', updated: 'UpdatedAt' }
      });
      return tables.credentials;
    }

    function ensureSessionTable() {
      if (tables.sessions) return tables.sessions;
      var db = requireDatabaseManager();
      tables.sessions = db.defineTable(SESSIONS_TABLE, {
        idColumn: 'ID',
        headers: [
          'ID', 'TokenHash', 'UserId', 'RememberMe', 'IdleTimeoutMinutes',
          'MetadataJson', 'CreatedAt', 'UpdatedAt', 'LastActivityAt',
          'ExpiresAt', 'RevokedAt'
        ],
        timestamps: { created: 'CreatedAt', updated: 'UpdatedAt' }
      });
      return tables.sessions;
    }

    function ensureResetTable() {
      if (tables.resets) return tables.resets;
      var db = requireDatabaseManager();
      tables.resets = db.defineTable(PASSWORD_RESET_TABLE, {
        idColumn: 'ID',
        headers: [
          'ID', 'UserId', 'TokenHash', 'Purpose', 'IssuedAt', 'ExpiresAt',
          'ConsumedAt', 'MetadataJson', 'CreatedAt', 'UpdatedAt'
        ],
        timestamps: { created: 'CreatedAt', updated: 'UpdatedAt' }
      });
      return tables.resets;
    }

    function ensureInfrastructure() {
      var summary = { success: true, tables: {} };
      try {
        ensureUsersTable();
        summary.tables.users = true;
      } catch (err) {
        summary.success = false;
        summary.tables.users = { error: err.message };
      }
      try {
        ensureCredentialTable();
        summary.tables.credentials = true;
      } catch (err2) {
        summary.success = false;
        summary.tables.credentials = { error: err2.message };
      }
      try {
        ensureSessionTable();
        summary.tables.sessions = true;
      } catch (err3) {
        summary.success = false;
        summary.tables.sessions = { error: err3.message };
      }
      try {
        ensureResetTable();
        summary.tables.resets = true;
      } catch (err4) {
        summary.success = false;
        summary.tables.resets = { error: err4.message };
      }
      return summary;
    }

    // -------------------------------------------------------------------------
    // Lookup helpers
    // -------------------------------------------------------------------------

    function getUsersIndex() {
      var table = ensureUsersTable();
      var rows = table.project([
        'ID', 'UserName', 'FullName', 'Email', 'NormalizedEmail', 'CanLogin',
        'IsAdmin', 'Roles', 'Pages', 'SecurityStamp', 'ResetRequired',
        'LastLoginAt', 'UpdatedAt'
      ], { cache: true });
      var byId = {};
      var byEmail = {};
      for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        byId[String(row.ID)] = row;
        var emailKey = normalizeEmail(row.NormalizedEmail || row.Email);
        if (emailKey) {
          byEmail[emailKey] = row;
        }
      }
      return { byId: byId, byEmail: byEmail };
    }

    function findUserByEmail(email) {
      var normalized = normalizeEmail(email);
      if (!normalized) return null;
      var index = getUsersIndex();
      return index.byEmail[normalized] || null;
    }

    function findUserById(userId) {
      if (!userId && userId !== 0) return null;
      var index = getUsersIndex();
      return index.byId[String(userId)] || null;
    }

    function getCredentialRecord(userId) {
      var table = ensureCredentialTable();
      return table.findOne({ UserId: String(userId) });
    }

    function saveCredentialRecord(userId, passwordRecord, options) {
      var table = ensureCredentialTable();
      var normalizedUserId = String(userId);
      var existing = table.findOne({ UserId: normalizedUserId });
      var payload = {
        UserId: normalizedUserId,
        PasswordHash: passwordRecord.hash,
        PasswordSalt: passwordRecord.salt,
        PasswordIterations: passwordRecord.iterations,
        PasswordAlgorithm: passwordRecord.algorithm || 'SHA-256',
        PasswordVersion: passwordRecord.version || 1,
        PasswordUpdatedAt: passwordRecord.createdAt || iso(now()),
        MustChangePassword: sheetBoolean(options && options.mustChange === true)
      };

      if (existing) {
        table.update(existing.ID, payload);
        return clone(existing);
      }
      table.insert(payload);
      return payload;
    }

    function markUserForPasswordChange(userId, reason) {
      var table = ensureUsersTable();
      var updates = {
        ResetRequired: sheetBoolean(true)
      };
      if (reason && reason.token && reason.expiresAt) {
        updates.ResetPasswordToken = reason.token;
        updates.ResetPasswordTokenHash = reason.tokenHash || '';
        updates.ResetPasswordSentAt = reason.issuedAt || iso(now());
        updates.ResetPasswordExpiresAt = reason.expiresAt;
      }
      table.update(String(userId), updates);
    }

    function clearPasswordResetColumns(userId) {
      var table = ensureUsersTable();
      table.update(String(userId), {
        ResetPasswordToken: '',
        ResetPasswordTokenHash: '',
        ResetPasswordSentAt: '',
        ResetPasswordExpiresAt: '',
        ResetRequired: sheetBoolean(false)
      });
    }

    // -------------------------------------------------------------------------
    // Password token helpers
    // -------------------------------------------------------------------------

    function persistPasswordToken(userId, purpose, tokenPayload, options) {
      var table = ensureResetTable();
      var record = {
        UserId: String(userId),
        TokenHash: tokenPayload.tokenHash,
        Purpose: purpose,
        IssuedAt: tokenPayload.issuedAt,
        ExpiresAt: tokenPayload.expiresAt,
        ConsumedAt: '',
        MetadataJson: createMetadataBlob(options && options.metadata)
      };
      table.insert(record);
      markUserForPasswordChange(userId, tokenPayload);
      return tokenPayload;
    }

    function findResetRecordByToken(token) {
      var utils = requirePasswordUtils();
      var hash = utils.hashToken(token);
      if (!hash) return null;
      var table = ensureResetTable();
      var matches = table.find({ where: { TokenHash: hash } });
      if (!matches.length) return null;
      // Latest entry wins
      matches.sort(function (a, b) {
        return (new Date(b.IssuedAt).getTime() || 0) - (new Date(a.IssuedAt).getTime() || 0);
      });
      return matches[0];
    }

    function consumeResetRecord(record) {
      if (!record || !record.ID) return;
      var table = ensureResetTable();
      table.update(record.ID, {
        ConsumedAt: iso(now())
      });
    }

    // -------------------------------------------------------------------------
    // Session helpers
    // -------------------------------------------------------------------------

    function createSessionFor(userId, metadata, rememberMe) {
      var utils = requirePasswordUtils();
      var table = ensureSessionTable();
      var nowDate = now();
      var ttl = rememberMe ? REMEMBER_ME_TTL_MS : SESSION_TTL_MS;
      var token = utils.generateToken({ length: 48 });
      var tokenHash = utils.hashToken(token);
      var expiresAt = new Date(nowDate.getTime() + ttl);
      var record = {
        TokenHash: tokenHash,
        UserId: String(userId),
        RememberMe: sheetBoolean(!!rememberMe),
        IdleTimeoutMinutes: String(DEFAULT_IDLE_TIMEOUT_MINUTES),
        MetadataJson: createMetadataBlob(metadata),
        CreatedAt: iso(nowDate),
        UpdatedAt: iso(nowDate),
        LastActivityAt: iso(nowDate),
        ExpiresAt: iso(expiresAt),
        RevokedAt: ''
      };
      table.insert(record);
      return {
        token: token,
        expiresAt: record.ExpiresAt,
        ttlMilliseconds: ttl,
        idleTimeoutMinutes: DEFAULT_IDLE_TIMEOUT_MINUTES,
        metadata: metadata || null
      };
    }

    function findSessionByToken(token) {
      var utils = requirePasswordUtils();
      var hash = utils.hashToken(token);
      if (!hash) return null;
      var table = ensureSessionTable();
      var session = table.findOne({ TokenHash: hash });
      if (!session) return null;
      return session;
    }

    function validateToken(token) {
      var session = findSessionByToken(token);
      if (!session) {
        return { valid: false, reason: 'NOT_FOUND' };
      }
      if (session.RevokedAt) {
        return { valid: false, reason: 'REVOKED' };
      }
      var expires = new Date(session.ExpiresAt || 0).getTime();
      if (!expires || expires <= now().getTime()) {
        return { valid: false, reason: 'EXPIRED', session: session };
      }
      var user = findUserById(session.UserId);
      if (!user) {
        return { valid: false, reason: 'USER_MISSING', session: session };
      }
      return { valid: true, session: session, user: user };
    }

    function getSessionUser(token) {
      var result = validateToken(token);
      if (!result.valid) return null;
      return sanitizeUserForClient(result.user);
    }

    function keepAlive(token) {
      var session = findSessionByToken(token);
      if (!session) return { success: false, error: 'Session not found' };
      if (session.RevokedAt) return { success: false, error: 'Session revoked' };
      var table = ensureSessionTable();
      table.update(session.ID, {
        LastActivityAt: iso(now()),
        UpdatedAt: iso(now())
      });
      return { success: true };
    }

    function logout(token) {
      var session = findSessionByToken(token);
      if (!session) {
        return { success: true, message: 'Session already closed' };
      }
      var table = ensureSessionTable();
      table.update(session.ID, {
        RevokedAt: iso(now()),
        UpdatedAt: iso(now())
      });
      return { success: true };
    }

    function cleanupExpiredSessions() {
      var table = ensureSessionTable();
      var sessions = table.read({ cache: false });
      var nowMs = now().getTime();
      var removed = 0;
      for (var i = 0; i < sessions.length; i++) {
        var session = sessions[i];
        var expiresAt = new Date(session.ExpiresAt || 0).getTime();
        if (!expiresAt || expiresAt <= nowMs || session.RevokedAt) {
          table.delete(session.ID);
          removed++;
        }
      }
      return { success: true, evaluated: sessions.length, removed: removed };
    }

    function userHasActiveSession(userIdentifier) {
      if (!userIdentifier && userIdentifier !== 0) return false;
      var user = typeof userIdentifier === 'string' && userIdentifier.indexOf('@') !== -1
        ? findUserByEmail(userIdentifier)
        : findUserById(userIdentifier);
      if (!user) return false;
      var table = ensureSessionTable();
      var sessions = table.find({ where: { UserId: String(user.ID) } });
      var nowMs = now().getTime();
      for (var i = 0; i < sessions.length; i++) {
        var record = sessions[i];
        if (record.RevokedAt) continue;
        var expires = new Date(record.ExpiresAt || 0).getTime();
        if (expires && expires > nowMs) {
          return true;
        }
      }
      return false;
    }

    // -------------------------------------------------------------------------
    // User helpers
    // -------------------------------------------------------------------------

    function sanitizeUserForClient(user) {
      if (!user) return null;
      return {
        ID: user.ID,
        UserName: user.UserName || '',
        FullName: user.FullName || '',
        Email: user.Email || '',
        Roles: user.Roles || '',
        Pages: user.Pages || '',
        IsAdmin: parseBooleanFlag(user.IsAdmin),
        CanLogin: parseBooleanFlag(user.CanLogin)
      };
    }

    function updateLastLogin(userId) {
      var table = ensureUsersTable();
      table.update(String(userId), {
        LastLoginAt: iso(now())
      });
    }

    // -------------------------------------------------------------------------
    // Public workflows
    // -------------------------------------------------------------------------

    function login(email, password, rememberMe, metadata) {
      ensureInfrastructure();
      var utils = requirePasswordUtils();
      var normalizedEmail = normalizeEmail(email);
      if (!normalizedEmail) {
        return { success: false, error: 'Email is required', errorCode: 'MISSING_EMAIL' };
      }
      var passwordInput = utils.normalizePasswordInput(password);
      if (!passwordInput) {
        return { success: false, error: 'Password is required', errorCode: 'MISSING_PASSWORD' };
      }

      var user = findUserByEmail(normalizedEmail);
      if (!user) {
        return { success: false, error: 'Invalid email or password', errorCode: 'INVALID_CREDENTIALS' };
      }

      if (!parseBooleanFlag(user.CanLogin)) {
        return { success: false, error: 'This account is disabled.', errorCode: 'ACCOUNT_DISABLED' };
      }

      var credential = getCredentialRecord(user.ID);
      if (!credential || !credential.PasswordHash) {
        return { success: false, error: 'Account has no password. Request a setup email.', errorCode: 'PASSWORD_NOT_SET' };
      }

      var passwordRecord = {
        hash: credential.PasswordHash,
        salt: credential.PasswordSalt,
        iterations: toNumber(credential.PasswordIterations, 150000)
      };

      if (!utils.verifyPassword(passwordInput, passwordRecord)) {
        return { success: false, error: 'Invalid email or password', errorCode: 'INVALID_CREDENTIALS' };
      }

      var session = createSessionFor(user.ID, metadata || {}, rememberMe === true);
      updateLastLogin(user.ID);

      var sanitizedUser = sanitizeUserForClient(user);
      var redirectSlug = getLandingSlug(user, { metadata: metadata, user: sanitizedUser });
      var redirectUrl = buildLandingRedirectUrlFromSlug(redirectSlug);

      return {
        success: true,
        message: 'Login successful',
        user: sanitizedUser,
        sessionToken: session.token,
        rememberMe: rememberMe === true,
        sessionExpiresAt: session.expiresAt,
        sessionTtlSeconds: Math.max(60, Math.floor(session.ttlMilliseconds / 1000)),
        sessionIdleTimeoutMinutes: session.idleTimeoutMinutes,
        redirectSlug: redirectSlug,
        redirectUrl: redirectUrl || '',
        requestedReturnUrl: (metadata && metadata.requestedReturnUrl) || '',
        identityWarnings: [],
        needsCampaignAssignment: false
      };
    }

    function initializeCredentialsForUser(userId, options) {
      ensureInfrastructure();
      var user = findUserById(userId);
      if (!user) {
        return { success: false, error: 'User not found' };
      }
      var utils = requirePasswordUtils();
      var passwordRecord = utils.createPasswordRecord(utils.generateRandomPassword({ length: 14 }), { skipValidation: true });
      saveCredentialRecord(userId, passwordRecord, { mustChange: true });
      var tokenPayload = utils.createResetToken({ ttlMinutes: PASSWORD_RESET_TTL_MINUTES });
      persistPasswordToken(userId, PASSWORD_PURPOSE_SETUP, tokenPayload, options || {});
      return {
        success: true,
        token: tokenPayload.token,
        expiresAt: tokenPayload.expiresAt
      };
    }

    function issuePasswordSetupToken(userId, options) {
      ensureInfrastructure();
      var user = findUserById(userId);
      if (!user) {
        return { success: false, error: 'User not found' };
      }
      var utils = requirePasswordUtils();
      var tokenPayload = utils.createResetToken({ ttlMinutes: PASSWORD_RESET_TTL_MINUTES });
      persistPasswordToken(userId, PASSWORD_PURPOSE_SETUP, tokenPayload, options || {});
      return {
        success: true,
        token: tokenPayload.token,
        expiresAt: tokenPayload.expiresAt
      };
    }

    function issuePasswordResetToken(userId, options) {
      ensureInfrastructure();
      var user = findUserById(userId);
      if (!user) {
        return { success: false, error: 'User not found' };
      }
      var utils = requirePasswordUtils();
      var tokenPayload = utils.createResetToken({ ttlMinutes: PASSWORD_RESET_TTL_MINUTES });
      persistPasswordToken(userId, PASSWORD_PURPOSE_RESET, tokenPayload, options || {});
      return {
        success: true,
        token: tokenPayload.token,
        expiresAt: tokenPayload.expiresAt
      };
    }

    function changePassword(userId, currentPassword, newPassword, options) {
      ensureInfrastructure();
      var user = findUserById(userId);
      if (!user) {
        return { success: false, error: 'User not found' };
      }
      var utils = requirePasswordUtils();
      var credential = getCredentialRecord(userId);
      if (!credential || !credential.PasswordHash) {
        return { success: false, error: 'Password has not been set yet.' };
      }
      var record = {
        hash: credential.PasswordHash,
        salt: credential.PasswordSalt,
        iterations: toNumber(credential.PasswordIterations, 150000)
      };
      if (!utils.verifyPassword(currentPassword, record)) {
        return { success: false, error: 'Current password is incorrect.' };
      }
      var newRecord = utils.createPasswordRecord(newPassword, options || {});
      saveCredentialRecord(userId, newRecord, { mustChange: false });
      clearPasswordResetColumns(userId);
      return { success: true };
    }

    function resetPasswordWithToken(token, newPassword, options) {
      ensureInfrastructure();
      var record = findResetRecordByToken(token);
      if (!record) {
        return { success: false, error: 'Invalid or expired token.' };
      }
      if (record.ConsumedAt) {
        return { success: false, error: 'Token has already been used.' };
      }
      var expires = new Date(record.ExpiresAt || 0).getTime();
      if (!expires || expires <= now().getTime()) {
        return { success: false, error: 'Token has expired.' };
      }
      var utils = requirePasswordUtils();
      var user = findUserById(record.UserId);
      if (!user) {
        return { success: false, error: 'User not found.' };
      }
      var passwordRecord = utils.createPasswordRecord(newPassword, options || {});
      saveCredentialRecord(user.ID || user.Id || record.UserId, passwordRecord, { mustChange: false });
      consumeResetRecord(record);
      clearPasswordResetColumns(record.UserId);
      return { success: true };
    }

    // -------------------------------------------------------------------------
    // Device & MFA placeholders (not implemented yet)
    // -------------------------------------------------------------------------

    function beginMfaChallenge() {
      return { success: false, error: 'Multi-factor authentication is not configured.' };
    }

    function verifyMfaCode() {
      return { success: false, error: 'Multi-factor authentication is not configured.' };
    }

    function confirmDeviceVerification() {
      return { success: false, error: 'Device verification is not enabled.' };
    }

    function denyDeviceVerification() {
      return { success: false, error: 'Device verification is not enabled.' };
    }

    // -------------------------------------------------------------------------
    // Landing helpers
    // -------------------------------------------------------------------------

    function getLandingSlug(user, context) {
      if (context && context.requestedSlug) {
        return normalizeString(context.requestedSlug) || 'dashboard';
      }
      return 'dashboard';
    }

    function resolveLandingDestination(user, context) {
      var slug = getLandingSlug(user, context || {});
      return {
        slug: slug,
        redirectUrl: buildLandingRedirectUrlFromSlug(slug)
      };
    }

    function buildLandingRedirectUrlFromSlug(slug) {
      var normalized = normalizeString(slug);
      if (!normalized) return '';
      return '?page=' + encodeURIComponent(normalized);
    }

    // -------------------------------------------------------------------------
    // Request context helpers
    // -------------------------------------------------------------------------

    function captureLoginRequestContext(e) {
      if (!e) return null;
      var context = {
        serverObservedAt: iso(now()),
        host: (e.headers && e.headers.host) || '',
        serverUserAgent: (e.headers && e.headers['user-agent']) || '',
        forwardedFor: (e.headers && (e.headers['x-forwarded-for'] || e.headers['x-appengine-user-ip'])) || '',
        requestedReturnUrl: (e.parameter && e.parameter.returnUrl) || ''
      };
      try {
        CacheService.getScriptCache().put(LOGIN_CONTEXT_CACHE_KEY, JSON.stringify(context), LOGIN_CONTEXT_CACHE_TTL_SECONDS);
      } catch (_) { }
      return context;
    }

    function consumeLoginRequestContext() {
      try {
        var cache = CacheService.getScriptCache();
        var raw = cache.get(LOGIN_CONTEXT_CACHE_KEY);
        if (!raw) return null;
        cache.remove(LOGIN_CONTEXT_CACHE_KEY);
        return JSON.parse(raw);
      } catch (err) {
        return null;
      }
    }

    function deriveLoginReturnUrlFromEvent(e) {
      if (!e || !e.parameter) return '';
      var preferred = normalizeString(e.parameter.returnUrl || e.parameter.redirect || '');
      if (!preferred) return '';
      if (/^https?:\/\//i.test(preferred)) {
        return preferred;
      }
      return preferred.charAt(0) === '/' ? preferred : '/' + preferred;
    }

    // -------------------------------------------------------------------------
    // Public API
    // -------------------------------------------------------------------------

    return {
      ensureSheets: ensureInfrastructure,
      login: login,
      logout: logout,
      keepAlive: keepAlive,
      createSessionFor: createSessionFor,
      getSessionUser: getSessionUser,
      validateToken: validateToken,
      findUserByEmail: findUserByEmail,
      findUserById: findUserById,
      getUserByEmail: findUserByEmail,
      findUserByPrincipal: findUserByEmail,
      verifyUserPassword: function (userId, password) {
        var user = typeof userId === 'object' ? userId : findUserById(userId);
        if (!user) return false;
        var credential = getCredentialRecord(user.ID);
        if (!credential) return false;
        var utils = requirePasswordUtils();
        var record = {
          hash: credential.PasswordHash,
          salt: credential.PasswordSalt,
          iterations: toNumber(credential.PasswordIterations, 150000)
        };
        return utils.verifyPassword(password, record);
      },
      initializeCredentialsForUser: initializeCredentialsForUser,
      issuePasswordSetupToken: issuePasswordSetupToken,
      issuePasswordResetToken: issuePasswordResetToken,
      changePassword: changePassword,
      resetPasswordWithToken: resetPasswordWithToken,
      userHasActiveSession: userHasActiveSession,
      cleanupExpiredSessions: cleanupExpiredSessions,
      resolveLandingDestination: resolveLandingDestination,
      getLandingSlug: getLandingSlug,
      buildLandingRedirectUrl: buildLandingRedirectUrlFromSlug,
      captureLoginRequestContext: captureLoginRequestContext,
      consumeLoginRequestContext: consumeLoginRequestContext,
      deriveLoginReturnUrlFromEvent: deriveLoginReturnUrlFromEvent,
      beginMfaChallenge: beginMfaChallenge,
      verifyMfaCode: verifyMfaCode,
      confirmDeviceVerification: confirmDeviceVerification,
      denyDeviceVerification: denyDeviceVerification
    };
  })();

  global.AuthenticationService = AuthenticationService;

  // ---------------------------------------------------------------------------
  // Client wrappers for google.script.run exposure
  // ---------------------------------------------------------------------------

  function loginUser(email, password, rememberMe, clientMetadata) {
    try {
      var metadata = clientMetadata && typeof clientMetadata === 'object'
        ? clone(clientMetadata)
        : {};
      var serverContext = null;
      if (AuthenticationService && typeof AuthenticationService.consumeLoginRequestContext === 'function') {
        serverContext = AuthenticationService.consumeLoginRequestContext();
      }
      if (serverContext) {
        metadata = metadata || {};
        metadata.serverContext = serverContext;
        if (serverContext.requestedReturnUrl && !metadata.requestedReturnUrl) {
          metadata.requestedReturnUrl = serverContext.requestedReturnUrl;
        }
      }
      return AuthenticationService.login(email, password, rememberMe === true, metadata);
    } catch (err) {
      return {
        success: false,
        error: err && err.message ? err.message : String(err),
        errorCode: 'AUTH_FAILURE'
      };
    }
  }

  function beginMfaChallengeClient(challengeId, options) {
    return AuthenticationService.beginMfaChallenge(challengeId, options);
  }

  function verifyMfaCodeClient(challengeId, code, metadata) {
    return AuthenticationService.verifyMfaCode(challengeId, code, metadata);
  }

  function confirmDeviceVerificationClient(verificationId, code, metadata) {
    return AuthenticationService.confirmDeviceVerification(verificationId, code, metadata);
  }

  function denyDeviceVerificationClient(verificationId, metadata) {
    return AuthenticationService.denyDeviceVerification(verificationId, metadata);
  }

  function logoutUser(sessionToken) {
    return AuthenticationService.logout(sessionToken);
  }

  function keepAlive(sessionToken) {
    return AuthenticationService.keepAlive(sessionToken);
  }

  function cleanupExpiredSessionsJob() {
    return AuthenticationService.cleanupExpiredSessions();
  }

  global.loginUser = loginUser;
  global.beginMfaChallenge = beginMfaChallengeClient;
  global.verifyMfaCode = verifyMfaCodeClient;
  global.confirmDeviceVerification = confirmDeviceVerificationClient;
  global.denyDeviceVerification = denyDeviceVerificationClient;
  global.logoutUser = logoutUser;
  global.keepAlive = keepAlive;
  global.cleanupExpiredSessionsJob = cleanupExpiredSessionsJob;

  console.log('AuthenticationService rebuilt and loaded.');
})(typeof globalThis !== 'undefined' ? globalThis : this);

