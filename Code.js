const AUTH_CONFIG = Object.freeze({
  COOKIE_NAME: 'lumina_auth_token',
  TOKEN_TTL_MINUTES: 60,
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

  return template.evaluate()
    .setTitle('LuminaHQ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .addMetaTag('referrer', 'no-referrer')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function loginUser(email, password) {
  if (!email || !password) {
    return { success: false, message: 'Enter your email and password.' };
  }

  var normalizedEmail = normalizeEmail(email);
  var resolution = resolveUserByEmail(normalizedEmail);
  if (!resolution || !resolution.user) {
    return { success: false, message: 'Invalid username or password.' };
  }

  if (resolution.status && resolution.status.canLogin === false) {
    return {
      success: false,
      message: 'Your account has been disabled. Please contact your administrator.'
    };
  }

  if (resolution.status && resolution.status.emailConfirmed === false) {
    return {
      success: false,
      message: 'Please confirm your email address before logging in.'
    };
  }

  if (resolution.status && resolution.status.requiresReset === true) {
    return {
      success: false,
      message: 'You must change your password before continuing.'
    };
  }

  if (!verifyPassword(password, resolution)) {
    return { success: false, message: 'Invalid username or password.' };
  }

  var tokenInfo = issueAuthToken(resolution.user);
  return {
    success: true,
    token: tokenInfo.token,
    expiresAt: tokenInfo.expiresAt,
    user: {
      id: resolution.user.id,
      email: resolution.user.email,
      name: resolution.user.name || ''
    }
  };
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

function logoutUser(token) {
  if (!token) {
    return { success: true };
  }

  var parsed = parseToken(token);
  if (parsed && parsed.payload && parsed.payload.jti) {
    deleteTokenRecord(parsed.payload.jti);
  }

  return { success: true };
}

function issueAuthToken(user) {
  var now = new Date();
  var expires = new Date(now.getTime() + AUTH_CONFIG.TOKEN_TTL_MINUTES * 60 * 1000);
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
    expiresAt: payload.exp
  });

  return {
    token: token,
    expiresAt: new Date(payload.exp).toISOString()
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

  var user = getUserById(payload.uid);
  if (!user && payload.email) {
    user = getUserByEmail(payload.email);
  }
  if (!user) {
    deleteTokenRecord(payload.jti);
    return { valid: false, reason: 'UNKNOWN_USER' };
  }

  return {
    valid: true,
    token: token,
    expiresAt: new Date(payload.exp).toISOString(),
    user: {
      id: user.id,
      email: user.email,
      name: user.name || ''
    }
  };
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

  var randomBytes = Utilities.getRandomBytes(AUTH_CONFIG.SECRET_LENGTH_BYTES);
  var encoded = Utilities.base64Encode(randomBytes);
  props.setProperty(AUTH_CONFIG.SECRET_PROPERTY_KEY, encoded);
  return randomBytes;
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
      if (typeof utils.constantTimeEquals === 'function') {
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
