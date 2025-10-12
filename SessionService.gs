/**
 * SessionService.gs
 * -----------------------------------------------------------------------------
 * Encapsulates session handling for the Lumina Identity platform. Sessions are
 * issued after successful authentication and persisted via CacheService with a
 * fallback to PropertiesService to satisfy the 10–30 minute sliding expiration
 * requirement.
 */
(function bootstrapSessionService(global) {
  if (!global) return;
  if (global.SessionService && typeof global.SessionService === 'object') {
    return;
  }

  var CacheService = global.CacheService;
  var PropertiesService = global.PropertiesService;
  var Utilities = global.Utilities;
  function getIdentityRepository() {
    var repo = global.IdentityRepository;
    if (!repo || typeof repo.upsert !== 'function') {
      throw new Error('IdentityRepository not initialized');
    }
    return repo;
  }
  var DEFAULT_TTL_SECONDS = 20 * 60; // 20 minutes
  var MIN_TTL_SECONDS = 10 * 60;
  var MAX_TTL_SECONDS = 30 * 60;
  var SESSION_PREFIX = 'session::';
  var CSRF_PREFIX = 'csrf::';

  var cache = CacheService ? CacheService.getUserCache() : null;
  var scriptProperties = PropertiesService ? PropertiesService.getScriptProperties() : null;

  function now() {
    return new Date().getTime();
  }

  function clampTtl(ttlSeconds) {
    if (!ttlSeconds) {
      return DEFAULT_TTL_SECONDS;
    }
    return Math.max(MIN_TTL_SECONDS, Math.min(MAX_TTL_SECONDS, ttlSeconds));
  }

  function generateId(prefix) {
    var uuid = Utilities.getUuid();
    return prefix ? prefix + uuid : uuid;
  }

  function persistSession(session) {
    getIdentityRepository().upsert('Sessions', 'SessionId', session);
  }

  function toCachePayload(session) {
    return JSON.stringify(session);
  }

  function putCache(key, payload, ttlSeconds) {
    if (!cache) {
      return;
    }
    cache.put(key, payload, ttlSeconds);
  }

  function readCache(key) {
    if (!cache) {
      return null;
    }
    var cached = cache.get(key);
    return cached || null;
  }

  function writeFallback(key, payload) {
    if (!scriptProperties) {
      return;
    }
    scriptProperties.setProperty(key, payload);
  }

  function readFallback(key) {
    if (!scriptProperties) {
      return null;
    }
    return scriptProperties.getProperty(key);
  }

  function deleteFallback(key) {
    if (!scriptProperties) {
      return;
    }
    scriptProperties.deleteProperty(key);
  }

  function issueSession(user, campaignId, ip, ua, ttlSeconds) {
    var issuedAt = now();
    var expiresAt = issuedAt + clampTtl(ttlSeconds) * 1000;
    var session = {
      SessionId: generateId('SID-'),
      UserId: user.UserId,
      CampaignId: campaignId,
      IssuedAt: issuedAt,
      ExpiresAt: expiresAt,
      CSRF: generateCsrfToken(),
      IP: ip || '',
      UA: ua || ''
    };
    var payload = toCachePayload(session);
    putCache(SESSION_PREFIX + session.SessionId, payload, clampTtl(ttlSeconds));
    writeFallback(SESSION_PREFIX + session.SessionId, payload);
    persistSession(session);
    return session;
  }

  function generateCsrfToken() {
    var randomBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid());
    return Utilities.base64Encode(randomBytes).replace(/[^a-zA-Z0-9]/g, '').slice(0, 40);
  }

  function issueCsrf(sessionId) {
    var token = generateCsrfToken();
    putCache(CSRF_PREFIX + sessionId, token, clampTtl());
    writeFallback(CSRF_PREFIX + sessionId, token);
    return token;
  }

  function readSession(sessionId) {
    if (!sessionId) {
      return null;
    }
    var payload = readCache(SESSION_PREFIX + sessionId) || readFallback(SESSION_PREFIX + sessionId);
    if (!payload) {
      return null;
    }
    try {
      var session = JSON.parse(payload);
      if (session.ExpiresAt && session.ExpiresAt < now()) {
        invalidateSession(sessionId);
        return null;
      }
      return session;
    } catch (err) {
      invalidateSession(sessionId);
      return null;
    }
  }

  function validateCsrf(sessionId, token) {
    if (!sessionId || !token) {
      return false;
    }
    var expected = readCache(CSRF_PREFIX + sessionId) || readFallback(CSRF_PREFIX + sessionId);
    return expected && expected === token;
  }

  function renewSession(sessionId, ttlSeconds) {
    var session = readSession(sessionId);
    if (!session) {
      return null;
    }
    var issuedAt = now();
    session.IssuedAt = issuedAt;
    session.ExpiresAt = issuedAt + clampTtl(ttlSeconds) * 1000;
    var payload = toCachePayload(session);
    putCache(SESSION_PREFIX + sessionId, payload, clampTtl(ttlSeconds));
    writeFallback(SESSION_PREFIX + sessionId, payload);
    getIdentityRepository().upsert('Sessions', 'SessionId', session);
    return session;
  }

  function invalidateSession(sessionId) {
    if (!sessionId) {
      return;
    }
    if (cache) {
      cache.remove(SESSION_PREFIX + sessionId);
      cache.remove(CSRF_PREFIX + sessionId);
    }
    deleteFallback(SESSION_PREFIX + sessionId);
    deleteFallback(CSRF_PREFIX + sessionId);
    getIdentityRepository().remove('Sessions', 'SessionId', sessionId);
  }

  function invalidateUserSessions(userId) {
    var sessions = getIdentityRepository().list('Sessions').filter(function(row) {
      return row.UserId === userId;
    });
    sessions.forEach(function(session) {
      invalidateSession(session.SessionId);
    });
  }

  global.SessionService = {
    issueSession: issueSession,
    renewSession: renewSession,
    invalidateSession: invalidateSession,
    invalidateUserSessions: invalidateUserSessions,
    readSession: readSession,
    issueCsrf: issueCsrf,
    validateCsrf: validateCsrf,
    DEFAULT_TTL_SECONDS: DEFAULT_TTL_SECONDS
  };
})(GLOBAL_SCOPE);
