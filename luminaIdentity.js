var luminaIdentity = (function () {
  var AUTH_COOKIE_NAME = 'authToken';
  var CACHE_PREFIX = 'ILLUMINA_IDENTITY:';
  var SESSION_CACHE_PREFIX = CACHE_PREFIX + 'SESSION:';
  var USER_CACHE_PREFIX = CACHE_PREFIX + 'USER:';
  var EMAIL_CACHE_PREFIX = CACHE_PREFIX + 'EMAIL:';
  var CACHE_SECONDS = 300;

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

  function pickFirst(value) {
    if (Array.isArray(value)) {
      return value.length ? value[0] : '';
    }
    return value;
  }

  function parseCookies(e) {
    var header = '';
    try {
      if (e && e.headers) {
        header = e.headers.Cookie || e.headers.cookie || '';
      }
    } catch (err) {
      logWarning('luminaIdentity.parseCookies', err);
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
      if (safeString(rawKey).toLowerCase() !== safeString(key).toLowerCase()) {
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
      return safeString(pickFirst(e.parameters[key]));
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

  function scriptCache() {
    try {
      if (typeof CacheService !== 'undefined' && CacheService) {
        return CacheService.getScriptCache();
      }
    } catch (err) {
      logWarning('luminaIdentity.scriptCache', err);
    }
    return null;
  }

  function readCache(key) {
    if (!key) {
      return null;
    }
    try {
      var cache = scriptCache();
      if (cache) {
        var value = cache.get(key);
        if (value) {
          return JSON.parse(value);
        }
      }
    } catch (err) {
      logWarning('luminaIdentity.readCache', err);
    }
    return null;
  }

  function writeCache(key, value, ttl) {
    if (!key) {
      return;
    }
    try {
      var cache = scriptCache();
      if (cache) {
        cache.put(key, JSON.stringify(value), Math.min(21600, Math.max(5, ttl || CACHE_SECONDS)));
      }
    } catch (err) {
      logWarning('luminaIdentity.writeCache', err);
    }
  }

  function normalizeUserId(user) {
    if (!user) {
      return '';
    }
    return safeString(user.ID || user.Id || user.id || user.UserId || user.userId);
  }

  function normalizeEmail(user) {
    if (!user) {
      return '';
    }
    return safeString(user.Email || user.email || user.EmailAddress || user.emailAddress);
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
      logWarning('luminaIdentity.readUsersDataset(getAllUsersRaw)', err);
    }

    try {
      if (typeof readSheet === 'function') {
        var sheetData = readSheet('Users') || [];
        if (Array.isArray(sheetData) && sheetData.length) {
          return sheetData;
        }
      }
    } catch (err2) {
      logWarning('luminaIdentity.readUsersDataset(readSheet)', err2);
    }

    return [];
  }

  function lookupUserRowById(userId) {
    var normalized = safeString(userId);
    if (!normalized) {
      return null;
    }
    var cacheKey = USER_CACHE_PREFIX + normalized;
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
        writeCache(cacheKey, row, CACHE_SECONDS);
        return row;
      }
    }
    return null;
  }

  function lookupUserRowByEmail(email) {
    var normalized = safeString(email).toLowerCase();
    if (!normalized) {
      return null;
    }
    var cacheKey = EMAIL_CACHE_PREFIX + normalized;
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
      if (safeString(row.Email || row.email).toLowerCase() === normalized) {
        writeCache(cacheKey, row, CACHE_SECONDS);
        return row;
      }
    }
    return null;
  }

  function mergeUserRecords(primary, fallback) {
    var output = {};
    [fallback || {}, primary || {}].forEach(function (source) {
      Object.keys(source).forEach(function (key) {
        if (!Object.prototype.hasOwnProperty.call(output, key) || safeString(output[key]) === '') {
          output[key] = source[key];
        }
      });
    });
    return output;
  }

  function resolveCampaignNavigation(user) {
    try {
      var campaignId = safeString(user && (user.CampaignID || user.campaignId || user.CampaignId));
      if (!campaignId) {
        return;
      }
      if (user.campaignNavigation && typeof user.campaignNavigation === 'object') {
        return;
      }
      if (typeof getCampaignNavigation === 'function') {
        var navigation = getCampaignNavigation(campaignId);
        if (navigation && typeof navigation === 'object') {
          user.campaignNavigation = navigation;
        }
      }
    } catch (err) {
      logWarning('luminaIdentity.resolveCampaignNavigation', err);
    }
  }

  function resolveCampaignName(user) {
    try {
      var campaignName = safeString(user && (user.CampaignName || user.campaignName));
      if (campaignName) {
        return;
      }
      var campaignId = safeString(user && (user.CampaignID || user.campaignId));
      if (!campaignId) {
        return;
      }
      if (typeof getCampaignById === 'function') {
        var campaign = getCampaignById(campaignId);
        if (campaign) {
          var resolved = safeString(campaign.Name || campaign.name || campaign.DisplayName || campaign.displayName);
          if (resolved) {
            user.CampaignName = resolved;
            user.campaignName = resolved;
          }
        }
      }
    } catch (err) {
      logWarning('luminaIdentity.resolveCampaignName', err);
    }
  }

  function ensureRoleInformation(user) {
    if (!user) {
      return;
    }
    var rolesPresent = Array.isArray(user.roleNames) && user.roleNames.length;
    if (!rolesPresent) {
      var userId = normalizeUserId(user);
      if (userId && typeof getUserRolesSafe === 'function') {
        try {
          var userRoles = getUserRolesSafe(userId) || [];
          var roleNames = userRoles
            .map(function (role) {
              return safeString(role && (role.name || role.Name || role.displayName || role.DisplayName));
            })
            .filter(Boolean);
          if (roleNames.length) {
            user.roleNames = roleNames;
          }
        } catch (err) {
          logWarning('luminaIdentity.ensureRoleInformation', err);
        }
      }
    }

    if ((!Array.isArray(user.roleNames) || !user.roleNames.length) && safeString(user.Roles)) {
      user.roleNames = safeString(user.Roles).split(',').map(function (part) {
        return safeString(part);
      }).filter(Boolean);
    }

    if (!user.RoleName && Array.isArray(user.roleNames) && user.roleNames.length) {
      user.RoleName = user.roleNames[0];
    }
  }

  function ensureSafeWrapper(user, meta) {
    var hydrated = mergeUserRecords(user, {});
    try {
      if (typeof createSafeUserObject === 'function') {
        hydrated = createSafeUserObject(hydrated);
      }
    } catch (err) {
      logWarning('luminaIdentity.ensureSafeWrapper', err);
    }

    ensureRoleInformation(hydrated);
    resolveCampaignName(hydrated);
    resolveCampaignNavigation(hydrated);

    try {
      var metaTarget = hydrated.identityMeta && typeof hydrated.identityMeta === 'object'
        ? hydrated.identityMeta
        : {};
      metaTarget = Object.assign({}, metaTarget, meta || {});
      hydrated.identityMeta = metaTarget;
    } catch (errMeta) {
      logWarning('luminaIdentity.ensureSafeWrapper.meta', errMeta);
      hydrated.identityMeta = meta || {};
    }

    return hydrated;
  }

  function cloneForCache(value) {
    try {
      return JSON.parse(JSON.stringify(value));
    } catch (_) {
      return value;
    }
  }

  function fetchSessionUser(sessionToken) {
    if (!sessionToken) {
      return null;
    }
    try {
      if (typeof AuthenticationService !== 'undefined'
        && AuthenticationService
        && typeof AuthenticationService.getSessionUser === 'function') {
        var sessionUser = AuthenticationService.getSessionUser(sessionToken);
        if (sessionUser) {
          sessionUser.sessionToken = sessionUser.sessionToken || sessionToken;
          return sessionUser;
        }
      }
    } catch (err) {
      logWarning('luminaIdentity.fetchSessionUser', err);
    }
    return null;
  }

  function resolveBaseUser(e, options) {
    var opts = options || {};
    var meta = { source: 'unknown', cacheHit: false };
    var sessionToken = resolveSessionToken(e, opts);

    var user = null;
    if (opts.explicitUser && typeof opts.explicitUser === 'object') {
      user = opts.explicitUser;
      meta.source = 'explicit';
    }

    if ((!user || !normalizeUserId(user)) && sessionToken) {
      var sessionCacheKey = SESSION_CACHE_PREFIX + sessionToken;
      if (opts.useCache !== false) {
        var cachedIdentity = readCache(sessionCacheKey);
        if (cachedIdentity) {
          cachedIdentity.sessionToken = cachedIdentity.sessionToken || sessionToken;
          meta.source = (cachedIdentity.identityMeta && cachedIdentity.identityMeta.source) || 'cache';
          meta.cacheHit = true;
          user = cachedIdentity;
        }
      }

      if (!user) {
        user = fetchSessionUser(sessionToken);
        if (user) {
          meta.source = meta.source === 'explicit' ? 'explicit+session' : 'session';
        }
      }
    }

    if ((!user || !normalizeUserId(user)) && opts.allowCurrentUser !== false) {
      try {
        if (typeof getCurrentUser === 'function') {
          var current = getCurrentUser();
          if (current && normalizeUserId(current)) {
            if (user) {
              user = mergeUserRecords(user, current);
              meta.source = meta.source + '+current';
            } else {
              user = current;
              meta.source = 'current';
            }
          }
        }
      } catch (err) {
        logWarning('luminaIdentity.resolveBaseUser.current', err);
      }
    }

    if (!user) {
      user = {};
    }

    if (!user.sessionToken && sessionToken) {
      user.sessionToken = sessionToken;
    }

    meta.sessionToken = user.sessionToken || sessionToken || '';
    return { user: user, meta: meta };
  }

  function hydrateUserRecord(baseUser) {
    var user = baseUser || {};
    var normalizedId = normalizeUserId(user);
    var normalizedEmail = normalizeEmail(user);

    if (normalizedId) {
      var row = lookupUserRowById(normalizedId);
      if (row) {
        user = mergeUserRecords(user, row);
      }
    } else if (normalizedEmail) {
      var byEmail = lookupUserRowByEmail(normalizedEmail);
      if (byEmail) {
        user = mergeUserRecords(user, byEmail);
      }
    }

    return user;
  }

  function resolveIdentity(e, options) {
    var context = resolveBaseUser(e, options);
    var user = hydrateUserRecord(context.user);
    var meta = Object.assign({}, context.meta, { resolvedAt: new Date().toISOString() });

    var identity = ensureSafeWrapper(user, meta);

    if (identity && identity.sessionToken && options && options.useCache !== false) {
      writeCache(SESSION_CACHE_PREFIX + identity.sessionToken, cloneForCache(identity), CACHE_SECONDS);
    }

    var cacheId = normalizeUserId(identity);
    if (cacheId && options && options.useCache !== false) {
      writeCache(USER_CACHE_PREFIX + cacheId, cloneForCache(identity), CACHE_SECONDS);
    }

    return identity;
  }

  function stringifyForTemplate(value) {
    if (typeof _stringifyForTemplate_ === 'function') {
      return _stringifyForTemplate_(value);
    }
    try {
      return JSON.stringify(value || {}).replace(/<\/script>/gi, '<\\/script>');
    } catch (err) {
      logWarning('luminaIdentity.stringifyForTemplate', err);
      return '{}';
    }
  }

  function injectIntoTemplate(tpl, identity) {
    if (!tpl) {
      return identity || {};
    }

    var user = identity || {};
    try {
      tpl.user = user;
      tpl.safeUser = user;
    } catch (err) {
      logWarning('luminaIdentity.injectIntoTemplate.assign', err);
    }

    var json = stringifyForTemplate(user);
    try {
      tpl.currentUserJson = json;
    } catch (errJson) {
      logWarning('luminaIdentity.injectIntoTemplate.currentUserJson', errJson);
    }
    try {
      tpl.identityJson = json;
    } catch (errIdentityJson) {
      logWarning('luminaIdentity.injectIntoTemplate.identityJson', errIdentityJson);
    }
    try {
      tpl.safeUserJson = json;
    } catch (errSafeJson) {
      logWarning('luminaIdentity.injectIntoTemplate.safeUserJson', errSafeJson);
    }

    try {
      tpl.identityMetaJson = stringifyForTemplate(user && user.identityMeta ? user.identityMeta : {});
    } catch (errMeta) {
      logWarning('luminaIdentity.injectIntoTemplate.identityMetaJson', errMeta);
    }

    return user;
  }

  return {
    resolve: resolveIdentity,
    injectTemplate: injectIntoTemplate,
    resolveSessionToken: resolveSessionToken
  };
})();
