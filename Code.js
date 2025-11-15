const AUTH_CONFIG = Object.freeze({
  COOKIE_NAME: 'lumina_auth_token',
  TOKEN_TTL_MINUTES: 60,
  REMEMBER_ME_TTL_MINUTES: 24 * 60,
  SECRET_PROPERTY_KEY: 'AUTH_SIGNING_SECRET',
  TOKEN_PROPERTY_PREFIX: 'AUTH_TOKEN_',
  SECRET_LENGTH_BYTES: 48
});

function doGet(e) {
  purgeExpiredTokens();
  var page = '';
  if (e && e.parameter && typeof e.parameter.page === 'string') {
    page = e.parameter.page.toLowerCase();
  }

  var templateData = {
    baseUrl: getWebAppBaseUrl(),
    cookieName: AUTH_CONFIG.COOKIE_NAME,
    tokenTtlMinutes: AUTH_CONFIG.TOKEN_TTL_MINUTES
  };

  if (page === 'dashboard') {
    return renderTemplate('Dashboard', templateData);
  }

  return renderTemplate('Login', templateData);
}

function renderTemplate(name, data) {
  var template = HtmlService.createTemplateFromFile(name);
  if (data && typeof data === 'object') {
    Object.keys(data).forEach(function (key) {
      template[key] = data[key];
    });
  }

  if (TEMPLATE_INCLUDE_STATE && typeof TEMPLATE_INCLUDE_STATE === 'object') {
    TEMPLATE_INCLUDE_STATE.once = Object.create(null);
  }

  return template.evaluate()
    .setTitle('LuminaHQ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

var TEMPLATE_INCLUDE_STATE = (typeof TEMPLATE_INCLUDE_STATE !== 'undefined' && TEMPLATE_INCLUDE_STATE)
  ? TEMPLATE_INCLUDE_STATE
  : { once: Object.create(null) };

function include(filename, data) {
  if (!filename) {
    return '';
  }
  var tpl = HtmlService.createTemplateFromFile(filename);
  if (data && typeof data === 'object') {
    Object.keys(data).forEach(function (key) {
      tpl[key] = data[key];
    });
  }
  return tpl.evaluate().getContent();
}

function includeOnce(filename, data) {
  if (!filename) {
    return '';
  }
  if (TEMPLATE_INCLUDE_STATE.once[filename]) {
    return '';
  }
  TEMPLATE_INCLUDE_STATE.once[filename] = true;
  return include(filename, data);
}

function validateToken(token) {
  try {
    var validation = verifyToken(token);
    return validation;
  } catch (error) {
    return {
      valid: false,
      reason: 'ERROR',
      message: error && error.message ? error.message : 'Validation failed.'
    };
  }
}

function applyAppTokenToLoginResult(loginResult, rememberMeOverride) {
  if (!loginResult || !loginResult.success) {
    return loginResult;
  }

  var identity = deriveLoginIdentity(loginResult);
  if (!identity) {
    return loginResult;
  }

  var rememberMe = (typeof rememberMeOverride === 'boolean')
    ? rememberMeOverride
    : !!loginResult.rememberMe;
  var sessionToken = extractSessionTokenFromLoginResult(loginResult);
  var issuedToken = issueAuthToken(identity, { rememberMe: rememberMe }, sessionToken);

  loginResult.token = issuedToken.token;
  loginResult.expiresAt = issuedToken.expiresAt;
  loginResult.ttlMinutes = issuedToken.ttlMinutes;
  loginResult.rememberMe = issuedToken.rememberMe;
  loginResult.cookieName = AUTH_CONFIG.COOKIE_NAME;

  return loginResult;
}

function deriveLoginIdentity(loginResult) {
  if (!loginResult) {
    return null;
  }

  var user = loginResult.user
    || loginResult.profile
    || loginResult.account
    || (loginResult.session && loginResult.session.user)
    || null;

  var email = normalizeEmail(
    (user && (user.Email || user.email || user.UserName || user.username))
    || loginResult.email
    || loginResult.userEmail
    || ''
  );

  var identifier = '';
  if (user) {
    identifier = extractUserId(user);
  }

  if (!identifier) {
    identifier = normalizeString(
      loginResult.userId
      || loginResult.accountId
      || loginResult.userIdentifier
      || ''
    );
  }

  if (!identifier && email) {
    identifier = email;
  }

  if (!identifier) {
    return null;
  }

  var name = '';
  if (user) {
    name = normalizeString(user.FullName || user.fullName || user.Name || user.name || user.UserName || user.username || '');
  }

  return {
    id: identifier,
    email: email,
    name: name
  };
}

function extractSessionTokenFromLoginResult(loginResult) {
  if (!loginResult) {
    return '';
  }
  if (loginResult.sessionToken) {
    return loginResult.sessionToken;
  }
  if (loginResult.session && loginResult.session.token) {
    return loginResult.session.token;
  }
  if (loginResult.session && loginResult.session.sessionToken) {
    return loginResult.session.sessionToken;
  }
  return '';
}

function isAppTokenFormat(token) {
  if (!token || typeof token !== 'string') {
    return false;
  }
  if (token.indexOf('.') === -1) {
    return false;
  }
  var parts = token.split('.');
  return parts.length === 2 && parts[0].length > 8 && parts[1].length > 8;
}

function resolveSessionTokenFromAppToken(token) {
  if (!token || typeof token !== 'string') {
    return { sessionToken: null, jti: null, record: null };
  }

  var parsed = parseToken(token);
  if (!parsed || !parsed.valid || !parsed.payload || !parsed.payload.jti) {
    return { sessionToken: null, jti: null, record: null };
  }

  var record = loadTokenRecord(parsed.payload.jti);
  return {
    sessionToken: record ? (record.sessionToken || null) : null,
    jti: parsed.payload.jti,
    record: record || null
  };
}

function invalidateAppToken(token) {
  var resolution = resolveSessionTokenFromAppToken(token);
  if (resolution && resolution.jti) {
    deleteTokenRecord(resolution.jti);
  }
  return resolution;
}

function invalidateTokensForSessionToken(sessionToken) {
  if (!sessionToken) {
    return 0;
  }
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(buildTokenPropertyKey(jti));
}

function purgeExpiredTokens() {
  var props = PropertiesService.getScriptProperties();
  var keys = props.getKeys();
  if (!keys || !keys.length) {
    return 0;
  }

  props.setProperty(REALTIME_JOB_LAST_RUN_PROP, String(now));
  props.setProperty(REALTIME_JOB_STATUS_PROP, 'running');

  try {
    runRealtimeJob(props, now, config);
  } catch (error) {
    var message = (error && error.message) ? error.message : String(error);
    props.setProperty(REALTIME_JOB_STATUS_PROP, 'error:' + message);
    if (typeof logError === 'function') {
      logError('checkRealtimeUpdatesJob', error);
    } else {
      console.error('[checkRealtimeUpdatesJob] ' + message, error);
    }
  }

  return removed;
}

function runRealtimeJob(props, now, config) {
  var handlers = getRealtimeUpdateHandlers();
  if (!handlers.length) {
    console.log('[checkRealtimeUpdatesJob] No realtime handlers registered; exiting early.');
    props.setProperty(REALTIME_JOB_STATUS_PROP, 'idle');
    props.setProperty(REALTIME_JOB_LAST_SUCCESS_PROP, String(Date.now()));
    return;
  }

  var start = now;
  var iteration = 0;
  var hasMoreWork = true;
  var workPerformed = false;

  while (hasMoreWork && Date.now() - start < config.maxRuntimeMs) {
    hasMoreWork = false;
    for (var i = 0; i < handlers.length; i++) {
      var handler = handlers[i];
      var handlerHasMore = false;
      try {
        handlerHasMore = runRealtimeUpdateHandler(handler, iteration, config);
      } catch (handlerError) {
        if (typeof logError === 'function') {
          logError('checkRealtimeUpdatesJob.handler', handlerError);
        } else {
          console.error('[checkRealtimeUpdatesJob] Handler error', handlerError);
        }
      }
      if (handlerHasMore) {
        hasMoreWork = true;
        workPerformed = true;
      }
    }
    iteration++;
    if (hasMoreWork && config.sleepMs > 0) {
      Utilities.sleep(config.sleepMs);
    }
  }

  if (!workPerformed) {
    console.log('[checkRealtimeUpdatesJob] No realtime updates were processed during this window.');
  } else if (hasMoreWork) {
    console.log('[checkRealtimeUpdatesJob] Max runtime reached; remaining work will continue on the next trigger.');
  }

  props.setProperty(REALTIME_JOB_STATUS_PROP, 'idle');
  props.setProperty(REALTIME_JOB_LAST_SUCCESS_PROP, String(Date.now()));
  props.setProperty('REALTIME_JOB_LAST_ITERATIONS', String(iteration));
}

/**
 * Reads realtime job configuration from Script Properties, falling back to the
 * defaults defined above.
 */
function getRealtimeJobConfig(props) {
  if (!props) {
    props = PropertiesService.getScriptProperties();
  }
  var expires = new Date(now.getTime() + ttlMinutes * 60 * 1000);
  var payload = {
    uid: user.id,
    email: user.email,
    name: user.name || '',
    jti: Utilities.getUuid(),
    iat: now.getTime(),
    exp: expires.getTime()
  };

  var payloadEncoded = Utilities.base64EncodeWebSafe(JSON.stringify(payload));
  var signature = signPayload(payloadEncoded);
  var token = payloadEncoded + '.' + signature;

  storeTokenRecord({
    jti: payload.jti,
    uid: user.id,
    expiresAt: payload.exp,
    rememberMe: !!(options && options.rememberMe),
    sessionToken: sessionToken || null
  });

  return {
    token: token,
    expiresAt: new Date(payload.exp).toISOString(),
    ttlMinutes: ttlMinutes,
    rememberMe: !!(options && options.rememberMe)
  };
}

function verifyToken(token) {
  if (!token || typeof token !== 'string') {
    return { valid: false, reason: 'MISSING' };
  }

  var parsed = parseToken(token);
  if (!parsed.valid) {
    return parsed;
  }

  var payload = parsed.payload;
  var now = Date.now();
  if (!payload.exp || payload.exp <= now) {
    deleteTokenRecord(payload.jti);
    return { valid: false, reason: 'EXPIRED' };
  }

  var record = loadTokenRecord(payload.jti);
  if (!record) {
    return { valid: false, reason: 'REVOKED' };
  }

  if (record.uid !== payload.uid) {
    deleteTokenRecord(payload.jti);
    return { valid: false, reason: 'INVALID' };
  }

  var sessionState = resolveSessionStateForRecord(record);
  if (sessionState && sessionState.status === 'invalid') {
    deleteTokenRecord(payload.jti);
    return {
      valid: false,
      reason: sessionState.reason || 'SESSION_INVALID'
    };
  }
  if (sessionState && sessionState.status === 'expired') {
    deleteTokenRecord(payload.jti);
    return {
      valid: false,
      reason: sessionState.reason || 'SESSION_EXPIRED'
    };
  }
  var email = resolution.user ? resolution.user.email : '';

  var user = (sessionState && sessionState.user) ? sessionState.user : null;
  if (!user) {
    user = getUserById(payload.uid);
  }
  if (!user && payload.email) {
    user = getUserByEmail(payload.email);
  }
  if (!user) {
    deleteTokenRecord(payload.jti);
    return { valid: false, reason: 'UNKNOWN_USER' };
  }

  if (sessionState && sessionState.user) {
    var sessionUserId = normalizeString(sessionState.user.id || '');
    var payloadUserId = normalizeString(payload.uid || '');
    if (sessionUserId && payloadUserId && sessionUserId !== payloadUserId) {
      deleteTokenRecord(payload.jti);
      return { valid: false, reason: 'SESSION_MISMATCH' };
    }

    var sessionEmail = normalizeEmail(sessionState.user.email || '');
    var payloadEmail = normalizeEmail(payload.email || '');
    if (sessionEmail && payloadEmail && sessionEmail !== payloadEmail) {
      deleteTokenRecord(payload.jti);
      return { valid: false, reason: 'SESSION_MISMATCH' };
    }
  } catch (error) {
    console.warn('verifyPassword: SHA-256 fallback failed', error);
  }

  var expiresAtMs = payload.exp;
  if (record && record.expiresAt) {
    expiresAtMs = Math.min(expiresAtMs, record.expiresAt);
  }
  if (sessionState && sessionState.expiresAt) {
    expiresAtMs = Math.min(expiresAtMs, sessionState.expiresAt);
  }

  return {
    valid: true,
    token: token,
    expiresAt: new Date(expiresAtMs).toISOString(),
    rememberMe: !!(record && record.rememberMe),
    user: {
      id: user.id,
      email: user.email,
      name: user.name || ''
    },
    session: sessionState
  };
}

function resolveSessionStateForRecord(record) {
  if (!record || !record.sessionToken) {
    return { status: 'unlinked' };
  }

  if (typeof AuthenticationService === 'undefined' || !AuthenticationService) {
    return { status: 'unavailable' };
  }

  if (typeof AuthenticationService.getSessionStatus === 'function') {
    try {
      var status = AuthenticationService.getSessionStatus(record.sessionToken, { touch: true });
      if (!status || status.valid === false) {
        return {
          status: status && status.status ? status.status : 'invalid',
          reason: (status && status.reason) || 'SESSION_INVALID'
        };
      }

      return {
        status: 'active',
        sessionToken: record.sessionToken,
        expiresAt: toMillis(status.expiresAt || status.sessionExpiresAt || null),
        rememberMe: !!status.rememberMe,
        user: normalizeSessionUserForToken(status.user || status.sessionUser || null),
        raw: status
      };
    } catch (error) {
      console.warn('resolveSessionStateForRecord: getSessionStatus failed', error);
      return { status: 'error', reason: 'SESSION_ERROR' };
    }
  }

  if (typeof AuthenticationService.getSessionUser === 'function') {
    try {
      var sessionUser = AuthenticationService.getSessionUser(record.sessionToken);
      if (!sessionUser) {
        return { status: 'invalid', reason: 'SESSION_INVALID' };
      }
      return {
        status: 'active',
        sessionToken: record.sessionToken,
        expiresAt: null,
        user: normalizeSessionUserForToken(sessionUser),
        raw: sessionUser
      };
    } catch (fallbackError) {
      console.warn('resolveSessionStateForRecord: getSessionUser failed', fallbackError);
      return { status: 'error', reason: 'SESSION_ERROR' };
    }
  }

  return { status: 'unsupported' };
}

function normalizeSessionUserForToken(sessionUser) {
  if (!sessionUser) {
    return null;
  }
  var email = normalizeEmail(
    sessionUser.Email
    || sessionUser.email
    || sessionUser.UserName
    || sessionUser.username
    || ''
  );
  var identifier = extractUserId(sessionUser);
  if (!identifier) {
    identifier = normalizeString(
      sessionUser.id
      || sessionUser.userId
      || sessionUser.UserId
      || ''
    );
  }
  if (!identifier && email) {
    identifier = email;
  }
  return {
    id: identifier || '',
    email: email,
    name: normalizeString(
      sessionUser.FullName
      || sessionUser.fullName
      || sessionUser.Name
      || sessionUser.name
      || sessionUser.DisplayName
      || sessionUser.displayName
      || ''
    )
  };
}

function toMillis(value) {
  if (!value) {
    return null;
  }
  if (typeof value === 'number') {
    return value;
  }
  var parsed = Date.parse(value);
  return isNaN(parsed) ? null : parsed;
}

function parseToken(token) {
  try {
    var parts = token.split('.');
    if (parts.length !== 2) {
      return { valid: false, reason: 'FORMAT' };
    }

    var payloadSegment = parts[0];
    var signatureSegment = parts[1];
    var expectedSignature = signPayload(payloadSegment);

    if (!constantTimeEquals(signatureSegment, expectedSignature)) {
      return { valid: false, reason: 'SIGNATURE' };
    }

    var payloadJson = Utilities.newBlob(Utilities.base64DecodeWebSafe(payloadSegment)).getDataAsString();
    var payload = JSON.parse(payloadJson);

    return {
      valid: true,
      payload: payload,
      signature: signatureSegment
    };
  } catch (error) {
    return {
      valid: false,
      reason: 'ERROR',
      message: error && error.message ? error.message : 'Unable to parse token.'
    };
  }
}

function signPayload(payloadSegment) {
  var keyBytes = getSigningKey();
  var signatureBytes = Utilities.computeHmacSha256Signature(payloadSegment, keyBytes);
  return Utilities.base64EncodeWebSafe(signatureBytes);
}

function getSigningKey() {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty(AUTH_CONFIG.SECRET_PROPERTY_KEY);
  if (existing) {
    return Utilities.base64Decode(existing);
  }

  var randomBytes = generateRandomBytes(AUTH_CONFIG.SECRET_LENGTH_BYTES);
  var encoded = Utilities.base64Encode(randomBytes);
  props.setProperty(AUTH_CONFIG.SECRET_PROPERTY_KEY, encoded);
  return randomBytes;
}

function generateRandomBytes(length) {
  var bytes = [];
  while (bytes.length < length) {
    var seed = Utilities.getUuid() + ':' + Date.now() + ':' + Math.random();
    var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed);
    for (var i = 0; i < digest.length && bytes.length < length; i++) {
      bytes.push(digest[i]);
    }
  }
  return bytes.slice(0, length);
}

function storeTokenRecord(record) {
  if (!record || !record.jti) {
    return;
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    var props = PropertiesService.getScriptProperties();
    props.setProperty(buildTokenPropertyKey(record.jti), JSON.stringify(record));
  } finally {
    lock.releaseLock();
  }
}

function loadTokenRecord(jti) {
  if (!jti) {
    return null;
  }
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(buildTokenPropertyKey(jti));
  if (!raw) {
    return null;
  }
  try {
    var parsed = JSON.parse(raw);
    if (parsed && parsed.expiresAt && parsed.expiresAt <= Date.now()) {
      deleteTokenRecord(jti);
      return null;
    }
    return parsed;
  } catch (error) {
    deleteTokenRecord(jti);
    return null;
  }
}

function deleteTokenRecord(jti) {
  if (!jti) {
    return;
  }
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(buildTokenPropertyKey(jti));
}

function purgeExpiredTokens() {
  var props = PropertiesService.getScriptProperties();
  var keys = props.getKeys();
  if (!keys || !keys.length) {
    return;
  }

  var now = Date.now();
  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    keys.forEach(function (key) {
      if (key.indexOf(AUTH_CONFIG.TOKEN_PROPERTY_PREFIX) !== 0) {
        return;
      }
      var raw = props.getProperty(key);
      if (!raw) {
        props.deleteProperty(key);
        return;
      }
      try {
        var parsed = JSON.parse(raw);
        if (!parsed.expiresAt || parsed.expiresAt <= now) {
          props.deleteProperty(key);
        }
      } catch (error) {
        props.deleteProperty(key);
      }
    });
  } finally {
    lock.releaseLock();
  }
}

function buildTokenPropertyKey(jti) {
  return AUTH_CONFIG.TOKEN_PROPERTY_PREFIX + jti;
}

function constantTimeEquals(a, b) {
  if (!a || !b || a.length !== b.length) {
    return false;
  }
  var diff = 0;
  for (var i = 0; i < a.length; i++) {
    diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return diff === 0;
}

function getWebAppBaseUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (error) {
    return '';
  }
}

function resolveUserByEmail(email) {
  if (!email) {
    return null;
  }
  var raw = fetchRawUserByEmail(email);
  if (!raw && typeof email === 'string') {
    raw = fetchRawUserByEmail(normalizeEmail(email));
  }
  return raw ? normalizeResolvedUser(raw, email) : null;
}

function resolveUserById(id) {
  if (!id) {
    return null;
  }
  var raw = fetchRawUserById(id);
  if (!raw && typeof id === 'string' && id.indexOf('@') > -1) {
    raw = fetchRawUserByEmail(id);
    if (raw) {
      return normalizeResolvedUser(raw, id);
    }
  }
  return raw ? normalizeResolvedUser(raw) : null;
}

function getUserByEmail(email) {
  var resolved = resolveUserByEmail(email);
  return resolved ? resolved.user : null;
}

function getUserById(id) {
  var resolved = resolveUserById(id);
  return resolved ? resolved.user : null;
}

function fetchRawUserByEmail(email) {
  if (!email) {
    return null;
  }
  var normalized = normalizeEmail(email);
  if (!normalized) {
    return null;
  }
  try {
    if (typeof AuthenticationService !== 'undefined'
      && AuthenticationService
      && typeof AuthenticationService.getUserByEmail === 'function') {
      var user = AuthenticationService.getUserByEmail(normalized);
      if (user) {
        return user;
      }
    }
  } catch (error) {
    console.warn('fetchRawUserByEmail: AuthenticationService lookup failed', error);
  }

  try {
    if (typeof readSheet === 'function') {
      var rows = readSheet('Users');
      if (Array.isArray(rows)) {
        for (var j = 0; j < rows.length; j++) {
          var row = rows[j];
          if (row && normalizeEmail(row.Email || row.email) === normalized) {
            return row;
          }
        }
      }
    }
  } catch (sheetError) {
    console.warn('fetchRawUserByEmail: Sheet lookup failed', sheetError);
  }

  return null;
}

function fetchRawUserById(id) {
  var normalized = normalizeString(id);
  if (!normalized) {
    return null;
  }

  try {
    if (typeof AuthenticationService !== 'undefined'
      && AuthenticationService
      && typeof AuthenticationService.findUserById === 'function') {
      var user = AuthenticationService.findUserById(normalized);
      if (user) {
        return user;
      }
    }
  } catch (error) {
    console.warn('fetchRawUserById: AuthenticationService lookup failed', error);
  }

  try {
    if (typeof readSheet === 'function') {
      var rows = readSheet('Users');
      if (Array.isArray(rows)) {
        for (var j = 0; j < rows.length; j++) {
          var row = rows[j];
          var rowId = extractUserId(row);
          if (row && rowId && rowId === normalized) {
            return row;
          }
        }
      }
    }
  } catch (sheetError) {
    console.warn('fetchRawUserById: Sheet lookup failed', sheetError);
  }

  return null;
}

function normalizeResolvedUser(rawUser, fallbackEmail) {
  if (!rawUser) {
    return null;
  }

  var email = normalizeEmail(rawUser.Email || rawUser.email || rawUser.NormalizedEmail || fallbackEmail || '');
  var userId = extractUserId(rawUser);
  if (!userId && email) {
    userId = email;
  }

  if (!userId) {
    return null;
  }

  var name = normalizeString(
    rawUser.FullName || rawUser.fullName || rawUser.Name || rawUser.name || rawUser.UserName || rawUser.username || ''
  );

  var passwordHash = extractPasswordHash(rawUser);

  var status = {
    canLogin: interpretBoolean(rawUser.CanLogin || rawUser.canLogin),
    emailConfirmed: interpretBoolean(rawUser.EmailConfirmed || rawUser.emailConfirmed),
    requiresReset: interpretBoolean(rawUser.ResetRequired || rawUser.resetRequired)
  };

  return {
    user: {
      id: userId,
      email: email,
      name: name
    },
    passwordHash: passwordHash,
    status: status,
    raw: rawUser
  };
}

function verifyPassword(password, resolution) {
  if (!resolution || !resolution.passwordHash) {
    return false;
  }

  var hash = normalizeString(resolution.passwordHash);
  if (!hash) {
    return false;
  }
  var email = resolution.user ? resolution.user.email : '';

  if (typeof AuthenticationService !== 'undefined'
    && AuthenticationService
    && typeof AuthenticationService.verifyUserPassword === 'function') {
    try {
      var verification = AuthenticationService.verifyUserPassword(password, hash, { email: email });
      if (verification && verification.success) {
        return true;
      }
    } catch (error) {
      console.warn('verifyPassword: AuthenticationService verification failed', error);
    }
  }

  var utils = getPasswordUtilitiesSafe();
  if (utils && typeof utils.verifyPassword === 'function') {
    try {
      if (utils.verifyPassword(password, hash)) {
        return true;
      }
    } catch (error) {
      console.warn('verifyPassword: PasswordUtilities verifyPassword failed', error);
    }
  }

  if (utils && typeof utils.normalizeHash === 'function' && typeof utils.hashPassword === 'function') {
    try {
      var normalizedHash = utils.normalizeHash(hash);
      var computed = utils.hashPassword(password);
      if (typeof utils.safeCompare === 'function') {
        if (utils.safeCompare(computed, normalizedHash)) {
          return true;
        }
      } else if (typeof utils.constantTimeEquals === 'function') {
        if (utils.constantTimeEquals(computed, normalizedHash)) {
          return true;
        }
      } else if (constantTimeEquals(computed, normalizedHash)) {
        return true;
      }
    } catch (error) {
      console.warn('verifyPassword: PasswordUtilities fallback comparison failed', error);
    }
  }

  try {
    var shaDigest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
    var shaHex = bytesToHex(shaDigest);
    if (shaHex && constantTimeEquals(shaHex, hash)) {
      return true;
    }
  } catch (error) {
    console.warn('verifyPassword: SHA-256 fallback failed', error);
  }

  return false;
}

function extractUserId(rawUser) {
  if (!rawUser) {
    return '';
  }
  var candidates = [
    rawUser.ID,
    rawUser.Id,
    rawUser.id,
    rawUser.UserId,
    rawUser.UserID,
    rawUser.userId,
    rawUser.userID,
    rawUser.EmployeeId,
    rawUser.EmployeeID,
    rawUser.employeeId
  ];

  for (var i = 0; i < candidates.length; i++) {
    var value = candidates[i];
    if (value || value === 0) {
      var normalized = normalizeString(value);
      if (normalized) {
        return normalized;
      }
    }
  }

  return '';
}

function extractPasswordHash(rawUser) {
  if (!rawUser) {
    return '';
  }
  var keys = [
    'PasswordHash',
    'passwordHash',
    'Password_Hash',
    'Password',
    'password',
    'PasswordDigest',
    'passwordDigest',
    'HashedPassword',
    'hashedPassword'
  ];

  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    if (Object.prototype.hasOwnProperty.call(rawUser, key)) {
      var value = rawUser[key];
      if (value || value === 0) {
        return String(value);
      }
    }
  }

  return '';
}

function interpretBoolean(value) {
  if (value === null || typeof value === 'undefined') {
    return undefined;
  }
  if (typeof value === 'boolean') {
    return value;
  }
  if (typeof value === 'number') {
    if (value === 1) {
      return true;
    }
    if (value === 0) {
      return false;
    }
  }
  var normalized = normalizeString(value);
  if (!normalized) {
    return undefined;
  }
  var truthy = ['true', 'yes', 'y', '1', 'enabled', 'active'];
  var falsy = ['false', 'no', 'n', '0', 'disabled', 'inactive'];
  if (truthy.indexOf(normalized.toLowerCase()) !== -1) {
    return true;
  }
  if (falsy.indexOf(normalized.toLowerCase()) !== -1) {
    return false;
  }
  return undefined;
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

function getPasswordUtilitiesSafe() {
  try {
    if (typeof ensurePasswordUtilities === 'function') {
      return ensurePasswordUtilities();
    }
  } catch (error) {
    console.warn('getPasswordUtilitiesSafe: ensurePasswordUtilities failed', error);
  }

  if (typeof PasswordUtilities !== 'undefined' && PasswordUtilities) {
    return PasswordUtilities;
  }

  return null;
}

function bytesToHex(bytes) {
  if (!bytes || typeof bytes.length === 'undefined') {
    return '';
  }
  var out = [];
  for (var i = 0; i < bytes.length; i++) {
    var hex = (bytes[i] & 0xff).toString(16);
    out.push(hex.length === 1 ? '0' + hex : hex);
  }
  return out.join('');
}
