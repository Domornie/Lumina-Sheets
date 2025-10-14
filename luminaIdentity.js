var LuminaIdentity = (function () {
  var AUTH_COOKIE_NAME = 'authToken';
  var CACHE_PREFIX = 'LUMINA_IDENTITY:';
  var SESSION_CACHE_PREFIX = CACHE_PREFIX + 'SESSION:';
  var USER_CACHE_PREFIX = CACHE_PREFIX + 'USER:';
  var USER_ROW_CACHE_PREFIX = CACHE_PREFIX + 'USER_ROW:';
  var CLAIM_CACHE_PREFIX = CACHE_PREFIX + 'CLAIMS:';
  var ACTIVE_SESSION_CACHE_PREFIX = CACHE_PREFIX + 'ACTIVE_SESSION:';
  var CACHE_TTL_SECONDS = 300;
  var ACTIVE_SESSION_MIN_TTL_SECONDS = 60;
  var ACTIVE_SESSION_MAX_TTL_SECONDS = 21600; // 6 hours (CacheService limit)
  var ACTIVE_SESSION_IDLE_TIMEOUT_SECONDS = 20 * 60; // 20 minutes inactivity window

  function logWarning(label, error) {
    try {
      if (typeof console !== 'undefined' && console && typeof console.warn === 'function') {
        console.warn(label, error);
      }
      if (typeof writeError === 'function') {
        writeError(label, error);
      }
    } catch (_) {
      // Ignore logging failures
    }
  }

  function safeString(value) {
    if (value === null || typeof value === 'undefined') {
      return '';
    }
    var text = String(value).trim();
    if (!text || text.toLowerCase() === 'undefined' || text.toLowerCase() === 'null') {
      return '';
    }
    return text;
  }

  function safeLower(value) {
    var str = safeString(value);
    return str ? str.toLowerCase() : '';
  }

  function parseCookies(e) {
    var header = '';
    try {
      if (e && e.headers) {
        header = e.headers.Cookie || e.headers.cookie || '';
      }
    } catch (err) {
      logWarning('LuminaIdentity.parseCookies', err);
    }

    var cookies = {};
    if (!header) {
      return cookies;
    }

    header.split(';').forEach(function (pair) {
      if (!pair) {
        return;
      }
      var idx = pair.indexOf('=');
      if (idx === -1) {
        return;
      }
      var key = safeString(pair.slice(0, idx));
      var value = safeString(pair.slice(idx + 1));
      if (key) {
        try {
          cookies[key] = decodeURIComponent(value);
        } catch (_) {
          cookies[key] = value;
        }
      }
    });

    return cookies;
  }

  function readQueryValue(e, key) {
    var query = safeString(e && e.queryString);
    if (!query) {
      return '';
    }

    var pairs = query.split('&');
    for (var i = 0; i < pairs.length; i++) {
      var part = pairs[i];
      if (!part) {
        continue;
      }
      var idx = part.indexOf('=');
      var rawKey = idx === -1 ? part : part.slice(0, idx);
      if (safeLower(rawKey) !== safeLower(key)) {
        continue;
      }
      var rawValue = idx === -1 ? '' : part.slice(idx + 1);
      try {
        return decodeURIComponent(rawValue.replace(/\+/g, ' '));
      } catch (_) {
        return rawValue;
      }
    }
    return '';
  }

  function readParameter(e, key) {
    if (!e) {
      return '';
    }
    if (e.parameter && e.parameter[key]) {
      return safeString(e.parameter[key]);
    }
    if (e.parameters && e.parameters[key]) {
      var picked = e.parameters[key];
      if (Array.isArray(picked)) {
        return picked.length ? safeString(picked[0]) : '';
      }
      return safeString(picked);
    }
    return '';
  }

  function resolveSessionToken(e, options) {
    var opts = options || {};
    var candidates = [];

    if (opts.sessionToken) {
      candidates.push(opts.sessionToken);
    }
    if (opts.token) {
      candidates.push(opts.token);
    }
    if (opts.authToken) {
      candidates.push(opts.authToken);
    }

    var storedToken = readActiveUserSessionToken();
    if (storedToken) {
      candidates.push(storedToken);
    }

    ['sessionToken', 'token', 'authToken'].forEach(function (name) {
      candidates.push(readParameter(e, name));
      candidates.push(readQueryValue(e, name));
    });

    var cookies = parseCookies(e);
    if (cookies[AUTH_COOKIE_NAME]) {
      candidates.push(cookies[AUTH_COOKIE_NAME]);
    }

    for (var i = 0; i < candidates.length; i++) {
      var value = safeString(candidates[i]);
      if (value) {
        return value;
      }
    }

    return '';
  }

  function getCache() {
    try {
      if (typeof CacheService !== 'undefined' && CacheService) {
        return CacheService.getScriptCache();
      }
    } catch (err) {
      logWarning('LuminaIdentity.getCache', err);
    }
    return null;
  }

  function readCache(key) {
    if (!key) {
      return null;
    }
    try {
      var cache = getCache();
      if (cache) {
        var value = cache.get(key);
        if (value) {
          return JSON.parse(value);
        }
      }
    } catch (err) {
      logWarning('LuminaIdentity.readCache', err);
    }
    return null;
  }

  function writeCache(key, value, ttl) {
    if (!key) {
      return;
    }
    try {
      var cache = getCache();
      if (cache) {
        cache.put(key, JSON.stringify(value), Math.min(21600, Math.max(5, ttl || CACHE_TTL_SECONDS)));
      }
    } catch (err) {
      logWarning('LuminaIdentity.writeCache', err);
    }
  }

  function removeCache(key) {
    if (!key) {
      return;
    }
    try {
      var cache = getCache();
      if (cache) {
        cache.remove(key);
      }
    } catch (err) {
      logWarning('LuminaIdentity.removeCache', err);
    }
  }

  function getActiveUserStoreKey() {
    try {
      if (typeof Session !== 'undefined'
        && Session
        && typeof Session.getTemporaryActiveUserKey === 'function') {
        var key = Session.getTemporaryActiveUserKey();
        if (key) {
          return ACTIVE_SESSION_CACHE_PREFIX + safeString(key);
        }
      }
    } catch (err) {
      logWarning('LuminaIdentity.getActiveUserStoreKey', err);
    }
    return null;
  }

  function normalizeIdleTimeoutSeconds(metadata) {
    var idleSeconds = ACTIVE_SESSION_IDLE_TIMEOUT_SECONDS;

    if (metadata && typeof metadata === 'object') {
      if (typeof metadata.idleTimeoutSeconds === 'number' && metadata.idleTimeoutSeconds > 0) {
        idleSeconds = metadata.idleTimeoutSeconds;
      } else if (metadata.idleTimeoutSeconds) {
        var parsedIdleSeconds = Number(metadata.idleTimeoutSeconds);
        if (!isNaN(parsedIdleSeconds) && parsedIdleSeconds > 0) {
          idleSeconds = parsedIdleSeconds;
        }
      } else if (typeof metadata.idleTimeoutMinutes === 'number' && metadata.idleTimeoutMinutes > 0) {
        idleSeconds = Math.floor(metadata.idleTimeoutMinutes * 60);
      } else if (metadata.idleTimeoutMinutes) {
        var parsedIdleMinutes = Number(metadata.idleTimeoutMinutes);
        if (!isNaN(parsedIdleMinutes) && parsedIdleMinutes > 0) {
          idleSeconds = Math.floor(parsedIdleMinutes * 60);
        }
      } else if (typeof metadata.sessionIdleTimeoutMinutes === 'number' && metadata.sessionIdleTimeoutMinutes > 0) {
        idleSeconds = Math.floor(metadata.sessionIdleTimeoutMinutes * 60);
      } else if (metadata.sessionIdleTimeoutMinutes) {
        var parsedSessionIdle = Number(metadata.sessionIdleTimeoutMinutes);
        if (!isNaN(parsedSessionIdle) && parsedSessionIdle > 0) {
          idleSeconds = Math.floor(parsedSessionIdle * 60);
        }
      }
    }

    idleSeconds = Math.min(idleSeconds, ACTIVE_SESSION_IDLE_TIMEOUT_SECONDS);
    idleSeconds = Math.max(ACTIVE_SESSION_MIN_TTL_SECONDS, idleSeconds);
    idleSeconds = Math.min(ACTIVE_SESSION_MAX_TTL_SECONDS, idleSeconds);

    return idleSeconds;
  }

  function computeActiveSessionTtlSeconds(metadata) {
    var idleTimeoutSeconds = normalizeIdleTimeoutSeconds(metadata);
    var ttl = idleTimeoutSeconds;

    if (metadata && typeof metadata === 'object') {
      if (typeof metadata.ttlSeconds === 'number' && metadata.ttlSeconds > 0) {
        ttl = metadata.ttlSeconds;
      } else if (metadata.ttlSeconds) {
        var parsedTtl = Number(metadata.ttlSeconds);
        if (!isNaN(parsedTtl) && parsedTtl > 0) {
          ttl = parsedTtl;
        }
      } else if (typeof metadata.sessionTtlSeconds === 'number' && metadata.sessionTtlSeconds > 0) {
        ttl = metadata.sessionTtlSeconds;
      } else if (metadata.sessionTtlSeconds) {
        var parsedSessionTtl = Number(metadata.sessionTtlSeconds);
        if (!isNaN(parsedSessionTtl) && parsedSessionTtl > 0) {
          ttl = parsedSessionTtl;
        }
      } else if (metadata.expiresAt) {
        var expiry = Date.parse(metadata.expiresAt);
        if (!isNaN(expiry)) {
          var delta = Math.floor((expiry - Date.now()) / 1000);
          if (delta > 0) {
            ttl = delta;
          }
        }
      } else if (metadata.sessionExpiresAt) {
        var altExpiry = Date.parse(metadata.sessionExpiresAt);
        if (!isNaN(altExpiry)) {
          var altDelta = Math.floor((altExpiry - Date.now()) / 1000);
          if (altDelta > 0) {
            ttl = altDelta;
          }
        }
      }

      if (metadata.rememberMe === true && ttl < 3600) {
        ttl = 3600;
      }
    }

    ttl = Math.max(ACTIVE_SESSION_MIN_TTL_SECONDS, ttl);
    ttl = Math.min(ACTIVE_SESSION_MAX_TTL_SECONDS, ttl);
    ttl = Math.min(ttl, idleTimeoutSeconds);

    return ttl;
  }

  function persistActiveUserSessionToken(token, metadata) {
    var normalized = safeString(token);
    if (!normalized) {
      return;
    }

    var storeKey = getActiveUserStoreKey();
    if (!storeKey) {
      return;
    }

    var idleTimeoutSeconds = normalizeIdleTimeoutSeconds(metadata);
    var nowIso = new Date().toISOString();

    var payload = {
      token: normalized,
      updatedAt: nowIso,
      lastActivityAt: nowIso,
      idleTimeoutSeconds: idleTimeoutSeconds
    };

    if (metadata && typeof metadata === 'object') {
      if (metadata.expiresAt || metadata.sessionExpiresAt) {
        var expiryIso = metadata.expiresAt || metadata.sessionExpiresAt;
        payload.expiresAt = safeString(expiryIso);
      }
      if (metadata.ttlSeconds || metadata.sessionTtlSeconds) {
        var ttlCandidate = metadata.ttlSeconds || metadata.sessionTtlSeconds;
        var numericTtl = typeof ttlCandidate === 'number' ? ttlCandidate : Number(ttlCandidate);
        if (!isNaN(numericTtl) && numericTtl > 0) {
          payload.ttlSeconds = numericTtl;
        }
      }
      if (metadata.rememberMe !== undefined) {
        payload.rememberMe = !!metadata.rememberMe;
      }
      if (metadata.idleTimeoutMinutes || metadata.sessionIdleTimeoutMinutes) {
        payload.idleTimeoutMinutes = metadata.idleTimeoutMinutes || metadata.sessionIdleTimeoutMinutes;
      }
    }

    var serialized = null;
    try {
      serialized = JSON.stringify(payload);
    } catch (err) {
      logWarning('LuminaIdentity.persistActiveUserSessionToken.serialize', err);
      return;
    }

    var ttlSeconds = computeActiveSessionTtlSeconds(metadata || {});

    try {
      if (typeof CacheService !== 'undefined' && CacheService && typeof CacheService.getUserCache === 'function') {
        CacheService.getUserCache().put(storeKey, serialized, ttlSeconds);
      }
    } catch (cacheErr) {
      logWarning('LuminaIdentity.persistActiveUserSessionToken.cache', cacheErr);
    }

    try {
      if (typeof PropertiesService !== 'undefined' && PropertiesService && typeof PropertiesService.getUserProperties === 'function') {
        PropertiesService.getUserProperties().setProperty(storeKey, serialized);
      }
    } catch (propErr) {
      logWarning('LuminaIdentity.persistActiveUserSessionToken.props', propErr);
    }
  }

  function readActiveUserSessionToken() {
    var storeKey = getActiveUserStoreKey();
    if (!storeKey) {
      return '';
    }

    var raw = null;

    try {
      if (typeof CacheService !== 'undefined' && CacheService && typeof CacheService.getUserCache === 'function') {
        raw = CacheService.getUserCache().get(storeKey);
      }
    } catch (cacheErr) {
      logWarning('LuminaIdentity.readActiveUserSessionToken.cache', cacheErr);
    }

    if (!raw) {
      try {
        if (typeof PropertiesService !== 'undefined' && PropertiesService && typeof PropertiesService.getUserProperties === 'function') {
          raw = PropertiesService.getUserProperties().getProperty(storeKey);
        }
      } catch (propErr) {
        logWarning('LuminaIdentity.readActiveUserSessionToken.props', propErr);
      }
    }

    if (!raw) {
      return '';
    }

    try {
      var payload = JSON.parse(raw);
      if (!payload) {
        return '';
      }

      var idleTimeoutSeconds = normalizeIdleTimeoutSeconds(payload);
      var lastActivity = safeString(payload.lastActivityAt || payload.updatedAt);
      if (lastActivity) {
        var lastActivityDate = Date.parse(lastActivity);
        if (!isNaN(lastActivityDate)) {
          var inactiveSeconds = Math.floor((Date.now() - lastActivityDate) / 1000);
          if (inactiveSeconds > idleTimeoutSeconds) {
            try {
              clearActiveUserSessionToken();
            } catch (clearErr) {
              logWarning('LuminaIdentity.readActiveUserSessionToken.expire', clearErr);
            }
            return '';
          }
        }
      }

      return safeString(payload.token);
    } catch (err) {
      logWarning('LuminaIdentity.readActiveUserSessionToken.parse', err);
    }

    return '';
  }

  function clearActiveUserSessionToken() {
    var storeKey = getActiveUserStoreKey();
    if (!storeKey) {
      return;
    }

    try {
      if (typeof CacheService !== 'undefined' && CacheService && typeof CacheService.getUserCache === 'function') {
        CacheService.getUserCache().remove(storeKey);
      }
    } catch (cacheErr) {
      logWarning('LuminaIdentity.clearActiveUserSessionToken.cache', cacheErr);
    }

    try {
      if (typeof PropertiesService !== 'undefined' && PropertiesService && typeof PropertiesService.getUserProperties === 'function') {
        PropertiesService.getUserProperties().deleteProperty(storeKey);
      }
    } catch (propErr) {
      logWarning('LuminaIdentity.clearActiveUserSessionToken.props', propErr);
    }
  }

  function clone(value) {
    try {
      return JSON.parse(JSON.stringify(value));
    } catch (_) {
      return value;
    }
  }

  function ensurePasswordToolkit() {
    try {
      if (typeof ensurePasswordUtilities === 'function') {
        return ensurePasswordUtilities();
      }
      if (typeof PasswordUtilities !== 'undefined' && PasswordUtilities) {
        return PasswordUtilities;
      }
    } catch (err) {
      logWarning('LuminaIdentity.ensurePasswordToolkit', err);
    }
    throw new Error('Password utilities unavailable');
  }

  function readUsersDataset() {
    try {
      if (typeof getAllUsersRaw === 'function') {
        var fromHelper = getAllUsersRaw();
        if (Array.isArray(fromHelper) && fromHelper.length) {
          return fromHelper;
        }
      }
    } catch (err) {
      logWarning('LuminaIdentity.readUsersDataset(getAllUsersRaw)', err);
    }

    try {
      if (typeof readSheet === 'function') {
        var sheetData = readSheet('Users') || [];
        if (Array.isArray(sheetData) && sheetData.length) {
          return sheetData;
        }
      }
    } catch (err2) {
      logWarning('LuminaIdentity.readUsersDataset(readSheet)', err2);
    }

    return [];
  }

  function lookupUserRowById(userId) {
    var normalized = safeString(userId);
    if (!normalized) {
      return null;
    }
    var cacheKey = USER_ROW_CACHE_PREFIX + normalized;
    var cached = readCache(cacheKey);
    if (cached) {
      return cached;
    }

    var dataset = readUsersDataset();
    for (var i = 0; i < dataset.length; i++) {
      var row = dataset[i];
      if (!row) {
        continue;
      }
      if (safeString(row.ID) === normalized) {
        writeCache(cacheKey, row, CACHE_TTL_SECONDS);
        return row;
      }
    }
    return null;
  }

  function lookupUserRowByEmail(email) {
    var normalized = safeLower(email);
    if (!normalized) {
      return null;
    }
    var cacheKey = USER_ROW_CACHE_PREFIX + 'EMAIL:' + normalized;
    var cached = readCache(cacheKey);
    if (cached) {
      return cached;
    }

    var dataset = readUsersDataset();
    for (var i = 0; i < dataset.length; i++) {
      var row = dataset[i];
      if (!row) {
        continue;
      }
      if (safeLower(row.Email || row.email) === normalized) {
        writeCache(cacheKey, row, CACHE_TTL_SECONDS);
        return row;
      }
    }
    return null;
  }

  function mergeRecords(primary, fallback) {
    var output = {};
    [fallback || {}, primary || {}].forEach(function (source, index) {
      Object.keys(source).forEach(function (key) {
        var value = source[key];
        if (index === 0) {
          if (!Object.prototype.hasOwnProperty.call(output, key)) {
            output[key] = value;
          }
        } else {
          var hasKey = Object.prototype.hasOwnProperty.call(output, key);
          if (!hasKey || (value !== null && typeof value !== 'undefined' && value !== '')) {
            output[key] = value;
          }
        }
      });
    });
    return output;
  }

  function fetchSessionUser(sessionToken) {
    if (!sessionToken) {
      return null;
    }

    var cached = readCache(SESSION_CACHE_PREFIX + 'RAW:' + sessionToken);
    if (cached && cached.sessionToken === sessionToken) {
      return cached;
    }

    var sessionUser = null;
    try {
      if (typeof AuthenticationService !== 'undefined'
        && AuthenticationService
        && typeof AuthenticationService.getSessionUser === 'function') {
        sessionUser = AuthenticationService.getSessionUser(sessionToken);
      }
    } catch (err) {
      logWarning('LuminaIdentity.fetchSessionUser', err);
    }

    if (sessionUser) {
      sessionUser.sessionToken = sessionUser.sessionToken || sessionToken;
      writeCache(SESSION_CACHE_PREFIX + 'RAW:' + sessionToken, clone(sessionUser), CACHE_TTL_SECONDS);
    }

    return sessionUser;
  }

  function hydrateUserRecord(sessionUser, explicitUser) {
    var base = mergeRecords(explicitUser || {}, sessionUser || {});

    var normalizedId = safeString(base.ID || base.Id || base.id || base.UserId || base.userId);
    var normalizedEmail = safeLower(base.Email || base.email || base.EmailAddress || base.emailAddress);

    if (normalizedId) {
      var byId = lookupUserRowById(normalizedId);
      if (byId) {
        base = mergeRecords(base, byId);
      }
    } else if (normalizedEmail) {
      var byEmail = lookupUserRowByEmail(normalizedEmail);
      if (byEmail) {
        base = mergeRecords(base, byEmail);
      }
    }

    return base;
  }

  function readRoleAssignments(userId) {
    var roles = [];
    if (!userId) {
      return { ids: [], names: [], records: [] };
    }

    try {
      if (typeof getUserRolesSafe === 'function') {
        roles = getUserRolesSafe(userId) || [];
      } else if (typeof getUserRoles === 'function') {
        roles = getUserRoles(userId) || [];
      }
    } catch (err) {
      logWarning('LuminaIdentity.readRoleAssignments', err);
      roles = [];
    }

    var ids = [];
    var names = [];
    var seenNames = {};
    roles.forEach(function (role) {
      if (!role) {
        return;
      }
      var id = safeString(role.id || role.ID);
      var name = safeString(role.name || role.Name || role.displayName || role.DisplayName);
      if (id) {
        ids.push(id);
      }
      if (name) {
        var lower = name.toLowerCase();
        if (!seenNames[lower]) {
          seenNames[lower] = true;
          names.push(name);
        }
      }
    });

    return { ids: ids, names: names, records: roles };
  }

  function readCampaignAssignments(userId) {
    if (!userId) {
      return [];
    }
    try {
      if (typeof csGetUserCampaigns === 'function') {
        return csGetUserCampaigns(userId) || [];
      }
    } catch (err) {
      logWarning('LuminaIdentity.readCampaignAssignments', err);
    }
    return [];
  }

  function readCampaignPermissions(userId) {
    if (!userId) {
      return [];
    }
    try {
      if (typeof getUserCampaignPermissionsSafe === 'function') {
        return getUserCampaignPermissionsSafe(userId) || [];
      }
    } catch (err) {
      logWarning('LuminaIdentity.readCampaignPermissions', err);
    }
    return [];
  }

  function readPageAssignments(userId) {
    if (!userId) {
      return [];
    }
    try {
      if (typeof getUserPagesSafe === 'function') {
        return getUserPagesSafe(userId) || [];
      }
    } catch (err) {
      logWarning('LuminaIdentity.readPageAssignments', err);
    }
    return [];
  }

  function buildClaims(user) {
    var userId = safeString(user && (user.ID || user.Id || user.id || user.UserId || user.userId));
    var baseClaims = {
      userId: userId,
      roles: [],
      roleIds: [],
      roleRecords: [],
      campaigns: [],
      permissions: [],
      pages: [],
      permissionFlags: {},
      isAdmin: false,
      isSupervisor: false,
      isTrainer: false,
      isAgent: false,
      isQa: false,
      claimsVersion: '2024.11.0',
      generatedAt: new Date().toISOString()
    };

    if (!userId) {
      return baseClaims;
    }

    var cacheKey = CLAIM_CACHE_PREFIX + userId;
    var cached = readCache(cacheKey);
    if (cached) {
      return cached;
    }

    var roles = readRoleAssignments(userId);
    var campaigns = readCampaignAssignments(userId);
    var permissions = readCampaignPermissions(userId);
    var pages = readPageAssignments(userId);

    var permissionFlags = {};
    permissions.forEach(function (perm) {
      if (!perm) {
        return;
      }
      var level = safeString(perm.PermissionLevel || perm.permissionLevel);
      if (level) {
        permissionFlags[level.toLowerCase()] = true;
      }
      if (perm.CanManageUsers || perm.canManageUsers) {
        permissionFlags.manageusers = true;
      }
      if (perm.CanManagePages || perm.canManagePages) {
        permissionFlags.managepages = true;
      }
    });

    var lowerRoles = {};
    roles.names.forEach(function (name) {
      lowerRoles[name.toLowerCase()] = true;
    });

    var claims = {
      userId: userId,
      roles: roles.names.slice(),
      roleIds: roles.ids.slice(),
      roleRecords: clone(roles.records || []),
      campaigns: clone(campaigns || []),
      permissions: clone(permissions || []),
      pages: clone(pages || []),
      permissionFlags: permissionFlags,
      isAdmin: !!lowerRoles.admin,
      isSupervisor: !!(lowerRoles.supervisor || lowerRoles.lead),
      isTrainer: !!lowerRoles.trainer,
      isAgent: !!(lowerRoles.agent || lowerRoles.specialist),
      isQa: !!(lowerRoles.qa || lowerRoles.quality),
      claimsVersion: '2024.11.0',
      generatedAt: new Date().toISOString()
    };

    writeCache(cacheKey, clone(claims), CACHE_TTL_SECONDS);
    return claims;
  }

  function buildIdentity(sessionUser, explicitUser, metaOptions) {
    var hydrated = hydrateUserRecord(sessionUser, explicitUser);
    var claims = buildClaims(hydrated);

    var sessionToken = safeString((sessionUser && sessionUser.sessionToken) || (hydrated && hydrated.sessionToken));
    var sessionExpiresAt = sessionUser && sessionUser.sessionExpiresAt ? sessionUser.sessionExpiresAt : (hydrated && hydrated.sessionExpiresAt);
    var sessionIdleTimeout = sessionUser && sessionUser.sessionIdleTimeoutMinutes ? sessionUser.sessionIdleTimeoutMinutes : (hydrated && hydrated.sessionIdleTimeoutMinutes);

    var identityMeta = mergeRecords((hydrated && hydrated.identityMeta) || {}, {
      source: metaOptions && metaOptions.source ? metaOptions.source : (sessionUser ? 'session' : 'anonymous'),
      resolvedAt: new Date().toISOString(),
      sessionToken: sessionToken || '',
      cacheHit: metaOptions && metaOptions.cacheHit ? true : false
    });

    var identity = mergeRecords(hydrated, {
      id: safeString(hydrated && (hydrated.ID || hydrated.Id || hydrated.UserId || hydrated.userId || hydrated.id)),
      email: safeString(hydrated && (hydrated.Email || hydrated.email || hydrated.EmailAddress || hydrated.emailAddress)),
      displayName: safeString(hydrated && (hydrated.FullName || hydrated.fullName || hydrated.Name || hydrated.name || hydrated.UserName || hydrated.username)),
      sessionToken: sessionToken,
      sessionExpiresAt: sessionExpiresAt || '',
      sessionIdleTimeoutMinutes: sessionIdleTimeout || '',
      claims: claims,
      roles: claims.roles.slice(),
      roleNames: claims.roles.slice(),
      roleRecords: claims.roleRecords.slice ? claims.roleRecords.slice() : clone(claims.roleRecords),
      campaigns: claims.campaigns.slice(),
      permissions: claims.permissions.slice(),
      permissionFlags: clone(claims.permissionFlags),
      pages: claims.pages.slice(),
      identityMeta: identityMeta
    });

    identity.session = {
      token: sessionToken,
      expiresAt: sessionExpiresAt || '',
      idleTimeoutMinutes: sessionIdleTimeout || '',
      rememberMe: !!(sessionUser && sessionUser.sessionRememberMe)
    };

    identity.campaignRoles = identity.campaigns.map(function (campaign) {
      return {
        id: safeString(campaign && (campaign.id || campaign.ID || campaign.CampaignId || campaign.CampaignID)),
        role: safeString(campaign && (campaign.role || campaign.Role))
      };
    });

    return identity;
  }

  function resolveIdentity(e, options) {
    var opts = options || {};

    if (opts.identity || opts.explicitIdentity) {
      var explicit = opts.identity || opts.explicitIdentity;
      return buildIdentity(explicit, explicit, { source: 'explicit', cacheHit: false });
    }

    var sessionToken = resolveSessionToken(e, opts);
    if (!sessionToken && opts.sessionToken) {
      sessionToken = safeString(opts.sessionToken);
    }

    var explicitUser = opts.explicitUser || null;
    var useCache = opts.useCache !== false;

    if (!sessionToken) {
      return buildIdentity({}, explicitUser, { source: 'anonymous', cacheHit: false });
    }

    var cachedIdentity = useCache ? readCache(SESSION_CACHE_PREFIX + sessionToken) : null;
    if (cachedIdentity && cachedIdentity.sessionToken === sessionToken) {
      try {
        persistActiveUserSessionToken(sessionToken, {
          sessionExpiresAt: (cachedIdentity.session && cachedIdentity.session.expiresAt) || cachedIdentity.sessionExpiresAt,
          sessionTtlSeconds: (cachedIdentity.session && (cachedIdentity.session.ttlSeconds || cachedIdentity.session.sessionTtlSeconds))
            || cachedIdentity.sessionTtlSeconds,
          sessionIdleTimeoutMinutes: (cachedIdentity.session && (cachedIdentity.session.idleTimeoutMinutes || cachedIdentity.session.sessionIdleTimeoutMinutes))
            || cachedIdentity.sessionIdleTimeoutMinutes,
          rememberMe: (cachedIdentity.session && cachedIdentity.session.rememberMe !== undefined)
            ? cachedIdentity.session.rememberMe
            : cachedIdentity.sessionRememberMe
        });
      } catch (cachePersistErr) {
        logWarning('LuminaIdentity.resolveIdentity.persistCacheActive', cachePersistErr);
      }
      return buildIdentity(cachedIdentity, explicitUser, { source: 'cache', cacheHit: true });
    }

    var sessionUser = fetchSessionUser(sessionToken);
    if (!sessionUser) {
      if (useCache) {
        removeCache(SESSION_CACHE_PREFIX + sessionToken);
      }
      try {
        clearActiveUserSessionToken();
      } catch (clearErr) {
        logWarning('LuminaIdentity.resolveIdentity.clearActive', clearErr);
      }
      return buildIdentity({ sessionToken: sessionToken }, explicitUser, { source: 'anonymous', cacheHit: false });
    }

    var identity = buildIdentity(sessionUser, explicitUser, { source: 'session', cacheHit: false });

    if (sessionToken) {
      try {
        persistActiveUserSessionToken(sessionToken, {
          sessionExpiresAt: identity.sessionExpiresAt || (sessionUser && (sessionUser.sessionExpiresAt || sessionUser.expiresAt)),
          sessionTtlSeconds: sessionUser && (sessionUser.sessionTtlSeconds || sessionUser.ttlSeconds),
          sessionIdleTimeoutMinutes: sessionUser && (sessionUser.sessionIdleTimeoutMinutes || sessionUser.idleTimeoutMinutes),
          rememberMe: sessionUser && (sessionUser.sessionRememberMe || sessionUser.rememberMe)
        });
      } catch (persistErr) {
        logWarning('LuminaIdentity.resolveIdentity.persistActive', persistErr);
      }
    }

    if (useCache) {
      writeCache(SESSION_CACHE_PREFIX + sessionToken, clone(identity), CACHE_TTL_SECONDS);
      var userId = safeString(identity.id);
      if (userId) {
        writeCache(USER_CACHE_PREFIX + userId, clone(identity), CACHE_TTL_SECONDS);
      }
    }

    return identity;
  }

  function ensureAuthenticated(e, options) {
    var identity = resolveIdentity(e, options);
    if (!identity || !identity.sessionToken) {
      var error = new Error('Authentication required');
      error.code = 'AUTH_REQUIRED';
      error.identity = identity;
      throw error;
    }
    return identity;
  }

  function login(email, password, rememberMe, clientMetadata) {
    var normalizedEmail = safeLower(email);
    var passwordValue = safeString(password);
    if (!normalizedEmail || !passwordValue) {
      return {
        success: false,
        error: 'Email and password are required.'
      };
    }

    var result = null;
    try {
      if (typeof AuthenticationService !== 'undefined'
        && AuthenticationService
        && typeof AuthenticationService.login === 'function') {
        result = AuthenticationService.login(normalizedEmail, passwordValue, !!rememberMe, clientMetadata);
      } else {
        throw new Error('AuthenticationService.login unavailable');
      }
    } catch (err) {
      logWarning('LuminaIdentity.login', err);
      return { success: false, error: 'Authentication service unavailable' };
    }

    if (!result || !result.success || !result.sessionToken) {
      return result || { success: false, error: 'Unable to authenticate' };
    }

    var identity = resolveIdentity(null, { sessionToken: result.sessionToken });
    try {
      persistActiveUserSessionToken(result.sessionToken, {
        sessionExpiresAt: result.sessionExpiresAt,
        sessionTtlSeconds: result.sessionTtlSeconds,
        sessionIdleTimeoutMinutes: result.sessionIdleTimeoutMinutes,
        rememberMe: result.rememberMe
      });
    } catch (persistErr) {
      logWarning('LuminaIdentity.login.persistResultActive', persistErr);
    }
    result.identity = identity;
    result.user = identity;
    result.roles = identity.roleNames;
    result.claims = identity.claims;
    result.campaigns = identity.campaigns;
    return result;
  }

  function logout(sessionToken) {
    var token = safeString(sessionToken);
    if (!token) {
      return { success: false, error: 'Session token required' };
    }

    try {
      if (typeof AuthenticationService !== 'undefined'
        && AuthenticationService
        && typeof AuthenticationService.logout === 'function') {
        var response = AuthenticationService.logout(token);
        removeCache(SESSION_CACHE_PREFIX + token);
        removeCache(SESSION_CACHE_PREFIX + 'RAW:' + token);
        try {
          if (typeof clearActiveUserSessionToken === 'function') {
            clearActiveUserSessionToken();
          }
        } catch (clearErr) {
          logWarning('LuminaIdentity.logout.clearActive', clearErr);
        }
        return response;
      }
    } catch (err) {
      logWarning('LuminaIdentity.logout', err);
    }

    removeCache(SESSION_CACHE_PREFIX + token);
    removeCache(SESSION_CACHE_PREFIX + 'RAW:' + token);
    try {
      if (typeof clearActiveUserSessionToken === 'function') {
        clearActiveUserSessionToken();
      }
    } catch (clearError) {
      logWarning('LuminaIdentity.logout.clearActiveFallback', clearError);
    }
    return { success: false, error: 'Authentication service unavailable' };
  }

  function keepAlive(sessionToken) {
    var token = safeString(sessionToken);
    if (!token) {
      return { success: false, error: 'Session token required' };
    }

    try {
      if (typeof AuthenticationService !== 'undefined'
        && AuthenticationService
        && typeof AuthenticationService.keepAlive === 'function') {
        return AuthenticationService.keepAlive(token);
      }
    } catch (err) {
      logWarning('LuminaIdentity.keepAlive', err);
    }

    return { success: false, error: 'Authentication service unavailable' };
  }

  function refreshIdentity(sessionToken) {
    var token = safeString(sessionToken);
    if (!token) {
      return null;
    }
    removeCache(SESSION_CACHE_PREFIX + token);
    removeCache(SESSION_CACHE_PREFIX + 'RAW:' + token);
    return resolveIdentity(null, { sessionToken: token, useCache: false });
  }

  function clearUserCache(userId) {
    var id = safeString(userId);
    if (!id) {
      return;
    }
    removeCache(USER_CACHE_PREFIX + id);
    removeCache(USER_ROW_CACHE_PREFIX + id);
    removeCache(CLAIM_CACHE_PREFIX + id);
  }

  function injectIntoTemplate(template, identity) {
    var tpl = template || {};
    var resolvedIdentity = identity || resolveIdentity(null, {});

    try {
      tpl.user = resolvedIdentity;
      tpl.identity = resolvedIdentity;
      tpl.safeUser = resolvedIdentity;
    } catch (err) {
      logWarning('LuminaIdentity.injectTemplate.assign', err);
    }

    try {
      var json = JSON.stringify(resolvedIdentity || {}).replace(/<\/script>/gi, '<\\/script>');
      tpl.identityJson = json;
      tpl.safeUserJson = json;
      tpl.currentUserJson = json;
    } catch (errJson) {
      logWarning('LuminaIdentity.injectTemplate.json', errJson);
    }

    try {
      tpl.identityMetaJson = JSON.stringify((resolvedIdentity && resolvedIdentity.identityMeta) || {});
    } catch (errMeta) {
      logWarning('LuminaIdentity.injectTemplate.meta', errMeta);
    }

    return resolvedIdentity;
  }

  function passwordApi() {
    var toolkit = ensurePasswordToolkit();
    return {
      hash: function (plain) { return toolkit.hashPassword(plain); },
      verify: function (plain, hash) { return toolkit.verifyPassword(plain, hash); },
      create: function (plain) { return toolkit.createPasswordHash(plain); },
      normalize: function (value) { return toolkit.normalizePasswordInput(value); },
      constantTimeEquals: function (a, b) { return toolkit.constantTimeEquals(a, b); }
    };
  }

  return {
    resolve: resolveIdentity,
    getIdentity: resolveIdentity,
    ensureAuthenticated: ensureAuthenticated,
    login: login,
    logout: logout,
    keepAlive: keepAlive,
    refreshIdentity: refreshIdentity,
    clearUserCache: clearUserCache,
    resolveSessionToken: resolveSessionToken,
    injectTemplate: injectIntoTemplate,
    buildClaims: buildClaims,
    persistActiveSessionToken: persistActiveUserSessionToken,
    readActiveSessionToken: readActiveUserSessionToken,
    clearActiveSessionToken: clearActiveUserSessionToken,
    password: passwordApi,
    getPasswordToolkit: ensurePasswordToolkit
  };
})();

var luminaIdentity = LuminaIdentity;
