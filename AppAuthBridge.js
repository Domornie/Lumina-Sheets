const APP_AUTH_CONFIG = Object.freeze({
  COOKIE_NAME: 'lumina_auth_token',
  TOKEN_TTL_MINUTES: 60,
  REMEMBER_ME_TTL_MINUTES: 24 * 60,
  SECRET_PROPERTY_KEY: 'APP_AUTH_SIGNING_SECRET',
  TOKEN_PROPERTY_PREFIX: 'APP_AUTH_TOKEN_',
  SECRET_LENGTH_BYTES: 48
});

var AppAuthBridge = (function () {
  function applyTokenToLoginResult(loginResult, rememberMeOverride) {
    if (!loginResult || !loginResult.success) {
      return loginResult;
    }

    purgeExpiredTokens();

    var identity = deriveIdentityFromLogin(loginResult);
    if (!identity) {
      return loginResult;
    }

    var rememberMe = (typeof rememberMeOverride === 'boolean')
      ? rememberMeOverride
      : !!loginResult.rememberMe;

    var sessionToken = extractSessionToken(loginResult);
    if (!sessionToken) {
      return loginResult;
    }

    var issued = issueAuthToken(identity, { rememberMe: rememberMe }, sessionToken);
    if (!issued) {
      return loginResult;
    }

    attachTokenMetadata(loginResult, issued);
    return loginResult;
  }

  function ensureSessionToken(sessionToken, rememberMe) {
    if (!sessionToken) {
      return null;
    }

    purgeExpiredTokens();

    var identity = resolveIdentityForSession(sessionToken);
    if (!identity) {
      return null;
    }

    invalidateTokensForSession(sessionToken);
    var issued = issueAuthToken(identity, { rememberMe: !!rememberMe }, sessionToken);
    return issued ? enrichTokenPayload(issued) : null;
  }

  function attachTokenMetadata(target, issued) {
    if (!target || !issued) {
      return;
    }
    target.appToken = issued.token;
    target.appTokenExpiresAt = issued.expiresAt;
    target.appTokenTtlMinutes = issued.ttlMinutes;
    target.appTokenCookieName = APP_AUTH_CONFIG.COOKIE_NAME;
    target.appTokenRememberMe = issued.rememberMe;
  }

  function enrichTokenPayload(issued) {
    if (!issued) {
      return null;
    }
    return {
      token: issued.token,
      expiresAt: issued.expiresAt,
      ttlMinutes: issued.ttlMinutes,
      rememberMe: issued.rememberMe,
      cookieName: APP_AUTH_CONFIG.COOKIE_NAME
    };
  }

  function validateToken(token) {
    if (!token || typeof token !== 'string') {
      return { valid: false, reason: 'MISSING' };
    }

    purgeExpiredTokens();

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
      return { valid: false, reason: 'MISMATCH' };
    }

    var sessionState = resolveSessionState(record);
    if (!sessionState || sessionState.status !== 'active') {
      deleteTokenRecord(payload.jti);
      return {
        valid: false,
        reason: (sessionState && sessionState.reason) || 'SESSION_INVALID'
      };
    }

    var user = sessionState.user || resolveUserById(payload.uid) || resolveUserByEmail(payload.email);
    if (!user) {
      deleteTokenRecord(payload.jti);
      return { valid: false, reason: 'UNKNOWN_USER' };
    }

    var expiresAt = Math.min(payload.exp, record.expiresAt || payload.exp);
    if (sessionState.expiresAt) {
      expiresAt = Math.min(expiresAt, sessionState.expiresAt);
    }

    return {
      valid: true,
      token: token,
      expiresAt: new Date(expiresAt).toISOString(),
      rememberMe: !!record.rememberMe,
      user: user,
      sessionToken: record.sessionToken,
      session: sessionState
    };
  }

  function invalidateToken(token) {
    if (!token) {
      return { success: true };
    }
    var parsed = parseToken(token);
    if (!parsed.valid) {
      return { success: true };
    }
    deleteTokenRecord(parsed.payload.jti);
    return { success: true };
  }

  function invalidateTokensForSession(sessionToken) {
    if (!sessionToken) {
      return 0;
    }

    var props = PropertiesService.getScriptProperties();
    var keys = props.getKeys();
    if (!keys || !keys.length) {
      return 0;
    }

    var removed = 0;
    for (var i = 0; i < keys.length; i++) {
      var key = keys[i];
      if (key.indexOf(APP_AUTH_CONFIG.TOKEN_PROPERTY_PREFIX) !== 0) {
        continue;
      }
      var raw = props.getProperty(key);
      if (!raw) {
        continue;
      }
      try {
        var parsed = JSON.parse(raw);
        if (parsed && parsed.sessionToken === sessionToken) {
          props.deleteProperty(key);
          removed++;
        }
      } catch (error) {
        props.deleteProperty(key);
      }
    }
    return removed;
  }

  function purgeExpiredTokens() {
    var props = PropertiesService.getScriptProperties();
    var keys = props.getKeys();
    if (!keys || !keys.length) {
      return;
    }

    var now = Date.now();
    var lock = LockService.getScriptLock();
    lock.waitLock(2000);
    try {
      keys.forEach(function (key) {
        if (key.indexOf(APP_AUTH_CONFIG.TOKEN_PROPERTY_PREFIX) !== 0) {
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

  function deriveIdentityFromLogin(loginResult) {
    if (!loginResult) {
      return null;
    }
    var user = loginResult.user || loginResult.profile || loginResult.account;
    if (!user) {
      return null;
    }

    var identifier = extractUserId(user);
    if (!identifier) {
      identifier = normalizeString(loginResult.userId || loginResult.accountId || '');
    }

    var email = normalizeEmail(
      (user.Email || user.email || user.UserName || user.username || '')
    );
    if (!identifier && email) {
      identifier = email;
    }

    if (!identifier) {
      return null;
    }

    return {
      id: identifier,
      email: email,
      name: normalizeString(user.FullName || user.fullName || user.Name || user.name || '')
    };
  }

  function extractSessionToken(loginResult) {
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

  function issueAuthToken(user, options, sessionToken) {
    if (!user || !user.id) {
      return null;
    }

    var now = new Date();
    var ttlMinutes = APP_AUTH_CONFIG.TOKEN_TTL_MINUTES;
    if (options && options.rememberMe) {
      ttlMinutes = APP_AUTH_CONFIG.REMEMBER_ME_TTL_MINUTES;
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
      email: user.email,
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

  function resolveIdentityForSession(sessionToken) {
    if (!sessionToken || typeof AuthenticationService === 'undefined' || !AuthenticationService) {
      return null;
    }

    if (typeof AuthenticationService.getSessionStatus === 'function') {
      try {
        var status = AuthenticationService.getSessionStatus(sessionToken, { touch: false });
        if (status && status.user) {
          return normalizeSessionUser(status.user || status.sessionUser || {});
        }
      } catch (statusError) {
        console.warn('resolveIdentityForSession: getSessionStatus failed', statusError);
      }
    }

    if (typeof AuthenticationService.getSessionUser === 'function') {
      try {
        var sessionUser = AuthenticationService.getSessionUser(sessionToken);
        if (sessionUser) {
          return normalizeSessionUser(sessionUser);
        }
      } catch (sessionError) {
        console.warn('resolveIdentityForSession: getSessionUser failed', sessionError);
      }
    }

    return null;
  }

  function resolveSessionState(record) {
    if (!record || !record.sessionToken || typeof AuthenticationService === 'undefined' || !AuthenticationService) {
      return { status: 'missing', reason: 'SESSION_MISSING' };
    }

    if (typeof AuthenticationService.getSessionStatus === 'function') {
      try {
        var status = AuthenticationService.getSessionStatus(record.sessionToken, { touch: true });
        if (!status || status.valid === false) {
          return {
            status: 'invalid',
            reason: (status && status.reason) || 'SESSION_INVALID'
          };
        }
        return {
          status: 'active',
          sessionToken: record.sessionToken,
          expiresAt: toMillis(status.expiresAt || status.sessionExpiresAt || null),
          rememberMe: !!status.rememberMe,
          user: normalizeSessionUser(status.user || status.sessionUser || {})
        };
      } catch (error) {
        console.warn('resolveSessionState: getSessionStatus failed', error);
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
          rememberMe: record.rememberMe,
          user: normalizeSessionUser(sessionUser)
        };
      } catch (fallbackError) {
        console.warn('resolveSessionState: getSessionUser failed', fallbackError);
        return { status: 'error', reason: 'SESSION_ERROR' };
      }
    }

    return { status: 'unknown', reason: 'SESSION_UNSUPPORTED' };
  }

  function normalizeSessionUser(user) {
    if (!user) {
      return null;
    }
    var email = normalizeEmail(user.Email || user.email || user.UserName || user.username || '');
    var identifier = extractUserId(user);
    if (!identifier && email) {
      identifier = email;
    }
    return {
      id: identifier || '',
      email: email,
      name: normalizeString(user.FullName || user.fullName || user.Name || user.name || '')
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
      return { valid: true, payload: payload, signature: signatureSegment };
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
    var existing = props.getProperty(APP_AUTH_CONFIG.SECRET_PROPERTY_KEY);
    if (existing) {
      return Utilities.base64Decode(existing);
    }
    var randomBytes = generateRandomBytes(APP_AUTH_CONFIG.SECRET_LENGTH_BYTES);
    props.setProperty(APP_AUTH_CONFIG.SECRET_PROPERTY_KEY, Utilities.base64Encode(randomBytes));
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
    lock.waitLock(2000);
    try {
      var props = PropertiesService.getScriptProperties();
      props.setProperty(buildTokenKey(record.jti), JSON.stringify(record));
    } finally {
      lock.releaseLock();
    }
  }

  function loadTokenRecord(jti) {
    if (!jti) {
      return null;
    }
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(buildTokenKey(jti));
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
    props.deleteProperty(buildTokenKey(jti));
  }

  function buildTokenKey(jti) {
    return APP_AUTH_CONFIG.TOKEN_PROPERTY_PREFIX + jti;
  }

  function resolveUserByEmail(email) {
    if (!email || typeof AuthenticationService === 'undefined' || !AuthenticationService) {
      return null;
    }
    if (typeof AuthenticationService.findUserByEmail === 'function') {
      try {
        return AuthenticationService.findUserByEmail(email);
      } catch (error) {
        console.warn('resolveUserByEmail: findUserByEmail failed', error);
      }
    }
    if (typeof AuthenticationService.getUserByEmail === 'function') {
      try {
        return AuthenticationService.getUserByEmail(email);
      } catch (fallbackError) {
        console.warn('resolveUserByEmail: getUserByEmail failed', fallbackError);
      }
    }
    return null;
  }

  function resolveUserById(id) {
    if (!id || typeof AuthenticationService === 'undefined' || !AuthenticationService) {
      return null;
    }
    if (typeof AuthenticationService.findUserById === 'function') {
      try {
        return AuthenticationService.findUserById(id);
      } catch (error) {
        console.warn('resolveUserById: findUserById failed', error);
      }
    }
    return null;
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

  function extractUserId(user) {
    if (!user) {
      return '';
    }
    var candidates = [
      user.ID,
      user.Id,
      user.id,
      user.UserId,
      user.UserID,
      user.userId,
      user.userID,
      user.EmployeeId,
      user.employeeId
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

  return {
    applyTokenToLoginResult: applyTokenToLoginResult,
    ensureSessionToken: ensureSessionToken,
    validateToken: validateToken,
    invalidateToken: invalidateToken,
    invalidateTokensForSession: invalidateTokensForSession,
    purgeExpiredTokens: purgeExpiredTokens,
    attachTokenMetadata: attachTokenMetadata,
    config: APP_AUTH_CONFIG
  };
})();

function applyAppTokenToLoginResult(loginResult, rememberMeOverride) {
  return AppAuthBridge.applyTokenToLoginResult(loginResult, rememberMeOverride);
}

function ensureAppTokenForSession(sessionToken, rememberMe) {
  return AppAuthBridge.ensureSessionToken(sessionToken, rememberMe);
}

function validateAppToken(token) {
  return AppAuthBridge.validateToken(token);
}

function invalidateAppToken(token) {
  return AppAuthBridge.invalidateToken(token);
}

function invalidateTokensForSessionToken(sessionToken) {
  return AppAuthBridge.invalidateTokensForSession(sessionToken);
}

function purgeExpiredAppTokens() {
  return AppAuthBridge.purgeExpiredTokens();
}
