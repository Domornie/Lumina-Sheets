/**
 * AuthorizationRegistry.js
 * -----------------------------------------------------------------------------
 * Central repository for authorization metadata so the entire system can resolve
 * role hierarchy rules, capability flags, and user authorization snapshots.
 *
 * The registry keeps a normalized copy of the role hierarchy, exposes helpers to
 * compare levels, and persists authorization snapshots keyed by both user ID and
 * session token.  Other modules can query these snapshots to enforce hierarchy
 * rules without having to rebuild the authorization profile on every request.
 */

var AuthorizationRegistry = (function () {
  var PROFILE_CACHE_TTL_SECONDS = 10 * 60; // 10 minutes
  var PROFILE_CACHE_PREFIX = 'AUTHZ_PROFILE:';
  var SESSION_CACHE_PREFIX = 'AUTHZ_SESSION:';
  var ROLE_RULES_PROPERTY_KEY = 'AUTHZ_ROLE_RULES';
  var PROFILE_PROPERTY_PREFIX = 'AUTHZ_PROFILE:';
  var SESSION_PROPERTY_PREFIX = 'AUTHZ_SESSION:';
  var ROLE_INDEX_CACHE_KEY = 'AUTHZ_ROLE_INDEX';

  function nowIsoString() {
    try {
      return new Date().toISOString();
    } catch (err) {
      return '';
    }
  }

  function getScriptCache() {
    try {
      if (typeof CacheService !== 'undefined' && CacheService.getScriptCache) {
        return CacheService.getScriptCache();
      }
    } catch (err) {
      console.warn('AuthorizationRegistry: unable to access CacheService', err);
    }
    return null;
  }

  function getScriptProperties() {
    try {
      if (typeof PropertiesService !== 'undefined' && PropertiesService.getScriptProperties) {
        return PropertiesService.getScriptProperties();
      }
    } catch (err) {
      console.warn('AuthorizationRegistry: unable to access PropertiesService', err);
    }
    return null;
  }

  function safeCachePut(cache, key, value, ttl) {
    if (!cache || !key) return;
    try {
      cache.put(key, value, ttl || PROFILE_CACHE_TTL_SECONDS);
    } catch (err) {
      console.warn('AuthorizationRegistry: cache.put failed for key', key, err);
    }
  }

  function safeCacheRemove(cache, key) {
    if (!cache || !key) return;
    try {
      cache.remove(key);
    } catch (err) {
      console.warn('AuthorizationRegistry: cache.remove failed for key', key, err);
    }
  }

  function cloneCapabilities(capabilities) {
    var clone = {
      isSystemAdmin: false,
      isExecutive: false,
      isManager: false,
      canManageUsers: false,
      canManagePages: false
    };

    if (!capabilities || typeof capabilities !== 'object') {
      return clone;
    }

    Object.keys(clone).forEach(function (key) {
      if (typeof capabilities[key] !== 'undefined') {
        clone[key] = !!capabilities[key];
      }
    });

    return clone;
  }

  function normalizeAliases(rawAliases) {
    if (!Array.isArray(rawAliases)) return [];
    var aliases = [];
    rawAliases.forEach(function (alias) {
      var normalized = '';
      if (alias || alias === 0) {
        normalized = String(alias).trim().toLowerCase();
      }
      if (normalized && aliases.indexOf(normalized) === -1) {
        aliases.push(normalized);
      }
    });
    return aliases;
  }

  function cloneRule(rule) {
    if (!rule || typeof rule !== 'object') {
      return null;
    }
    return {
      key: String(rule.key || rule.level || '').trim().toUpperCase(),
      label: String(rule.label || rule.title || rule.key || '').trim() || '',
      weight: Number(rule.weight || 0),
      aliases: normalizeAliases(rule.aliases),
      capabilities: cloneCapabilities(rule.capabilities)
    };
  }

  function sanitizeRules(rules) {
    var list = Array.isArray(rules) ? rules : [];
    var sanitized = [];

    list.forEach(function (rule, index) {
      var cloned = cloneRule(rule) || {};
      if (!cloned.key) {
        var fallback = (cloned.label || 'CUSTOM_' + index).toUpperCase().replace(/[^A-Z0-9_]/g, '_');
        cloned.key = fallback || ('CUSTOM_' + index);
      }
      if (!cloned.label) {
        cloned.label = cloned.key;
      }
      if (typeof cloned.weight !== 'number' || isNaN(cloned.weight)) {
        cloned.weight = 0;
      }
      sanitized.push(cloned);
    });

    return sanitized;
  }

  var DEFAULT_ROLE_RULES = sanitizeRules([
    {
      key: 'SYSTEM_ADMIN',
      label: 'System Administrator',
      weight: 2200,
      aliases: ['system administrator', 'system admin', 'administrator', 'admin'],
      capabilities: { isSystemAdmin: true, isExecutive: true, isManager: true, canManageUsers: true, canManagePages: true }
    },
    {
      key: 'CEO',
      label: 'Chief Executive Officer',
      weight: 2100,
      aliases: ['ceo', 'chief executive officer'],
      capabilities: { isExecutive: true, isManager: true, canManageUsers: true, canManagePages: true }
    },
    {
      key: 'COO',
      label: 'Chief Operating Officer',
      weight: 2050,
      aliases: ['coo', 'chief operating officer'],
      capabilities: { isExecutive: true, isManager: true, canManageUsers: true, canManagePages: true }
    },
    {
      key: 'CFO',
      label: 'Chief Financial Officer',
      weight: 2000,
      aliases: ['cfo', 'chief financial officer'],
      capabilities: { isExecutive: true, canManagePages: true }
    },
    {
      key: 'CTO',
      label: 'Chief Technology Officer',
      weight: 1950,
      aliases: ['cto', 'chief technology officer'],
      capabilities: { isExecutive: true, isManager: true, canManageUsers: true, canManagePages: true }
    },
    {
      key: 'DIRECTOR',
      label: 'Director',
      weight: 1800,
      aliases: ['director'],
      capabilities: { isExecutive: true, isManager: true, canManageUsers: true, canManagePages: true }
    },
    {
      key: 'OPERATIONS_MANAGER',
      label: 'Operations Manager',
      weight: 1700,
      aliases: ['operations manager', 'ops manager'],
      capabilities: { isManager: true, canManageUsers: true, canManagePages: true }
    },
    {
      key: 'ACCOUNT_MANAGER',
      label: 'Account Manager',
      weight: 1650,
      aliases: ['account manager'],
      capabilities: { isManager: true, canManageUsers: true }
    },
    {
      key: 'WORKFORCE_MANAGER',
      label: 'Workforce Manager',
      weight: 1600,
      aliases: ['workforce manager'],
      capabilities: { isManager: true, canManageUsers: true }
    },
    {
      key: 'QUALITY_ASSURANCE_MANAGER',
      label: 'Quality Assurance Manager',
      weight: 1550,
      aliases: ['quality assurance manager', 'qa manager'],
      capabilities: { isManager: true, canManagePages: true }
    },
    {
      key: 'TRAINING_MANAGER',
      label: 'Training Manager',
      weight: 1500,
      aliases: ['training manager'],
      capabilities: { isManager: true, canManageUsers: true }
    },
    {
      key: 'TEAM_SUPERVISOR',
      label: 'Team Supervisor',
      weight: 1400,
      aliases: ['team supervisor', 'team lead'],
      capabilities: { isManager: true, canManageUsers: true }
    },
    {
      key: 'FLOOR_SUPERVISOR',
      label: 'Floor Supervisor',
      weight: 1350,
      aliases: ['floor supervisor'],
      capabilities: { isManager: true, canManageUsers: true }
    },
    {
      key: 'ESCALATIONS_MANAGER',
      label: 'Escalations Manager',
      weight: 1300,
      aliases: ['escalations manager'],
      capabilities: { isManager: true, canManagePages: true }
    },
    {
      key: 'CLIENT_SUCCESS_MANAGER',
      label: 'Client Success Manager',
      weight: 1250,
      aliases: ['client success manager', 'customer success manager'],
      capabilities: { isManager: true, canManageUsers: true }
    },
    {
      key: 'COMPLIANCE_MANAGER',
      label: 'Compliance Manager',
      weight: 1200,
      aliases: ['compliance manager'],
      capabilities: { isManager: true, canManagePages: true }
    },
    {
      key: 'IT_SUPPORT_MANAGER',
      label: 'IT Support Manager',
      weight: 1150,
      aliases: ['it support manager', 'it manager', 'technology manager'],
      capabilities: { isManager: true, canManagePages: true }
    },
    {
      key: 'REPORTING_ANALYST',
      label: 'Reporting Analyst',
      weight: 900,
      aliases: ['reporting analyst', 'analyst'],
      capabilities: { canManagePages: true }
    },
    {
      key: 'QUALITY',
      label: 'Quality',
      weight: 850,
      aliases: ['quality', 'qa'],
      capabilities: { canManagePages: true }
    },
    {
      key: 'MANAGER',
      label: 'Manager',
      weight: 1450,
      aliases: ['manager', 'supervisor'],
      capabilities: { isManager: true, canManageUsers: true }
    },
    {
      key: 'AGENT',
      label: 'Agent',
      weight: 400,
      aliases: ['agent', 'associate'],
      capabilities: {}
    },
    {
      key: 'GUEST',
      label: 'Guest',
      weight: 100,
      aliases: ['guest', 'viewer'],
      capabilities: {}
    }
  ]);

  function getStoredRoleRules() {
    var props = getScriptProperties();
    if (!props) {
      return DEFAULT_ROLE_RULES.slice();
    }

    try {
      var stored = props.getProperty(ROLE_RULES_PROPERTY_KEY);
      if (!stored) {
        props.setProperty(ROLE_RULES_PROPERTY_KEY, JSON.stringify(DEFAULT_ROLE_RULES));
        return DEFAULT_ROLE_RULES.slice();
      }

      var parsed = JSON.parse(stored);
      var sanitized = sanitizeRules(parsed);
      return sanitized.length ? sanitized : DEFAULT_ROLE_RULES.slice();
    } catch (err) {
      console.warn('AuthorizationRegistry: unable to parse stored role rules', err);
      try {
        props.setProperty(ROLE_RULES_PROPERTY_KEY, JSON.stringify(DEFAULT_ROLE_RULES));
      } catch (persistError) {
        console.warn('AuthorizationRegistry: failed to persist default rules', persistError);
      }
      return DEFAULT_ROLE_RULES.slice();
    }
  }

  function persistRoleRules(rules) {
    var props = getScriptProperties();
    if (!props) {
      return;
    }
    try {
      props.setProperty(ROLE_RULES_PROPERTY_KEY, JSON.stringify(rules));
    } catch (err) {
      console.warn('AuthorizationRegistry: failed to persist role rules', err);
    }
  }

  function buildRoleIndex(rules) {
    var index = {};
    if (!Array.isArray(rules)) {
      return index;
    }

    rules.forEach(function (rule) {
      if (!rule || !rule.key) return;
      var key = String(rule.key).trim().toUpperCase();
      if (!key) return;
      index[key] = {
        key: key,
        label: rule.label || key,
        weight: Number(rule.weight || 0),
        capabilities: cloneCapabilities(rule.capabilities)
      };
    });

    return index;
  }

  function getRoleIndex() {
    var cache = getScriptCache();
    if (cache) {
      try {
        var cached = cache.get(ROLE_INDEX_CACHE_KEY);
        if (cached) {
          return JSON.parse(cached);
        }
      } catch (err) {
        console.warn('AuthorizationRegistry: failed to read role index cache', err);
      }
    }

    var rules = getStoredRoleRules();
    var index = buildRoleIndex(rules);

    if (cache) {
      try {
        cache.put(ROLE_INDEX_CACHE_KEY, JSON.stringify(index), PROFILE_CACHE_TTL_SECONDS);
      } catch (err) {
        console.warn('AuthorizationRegistry: unable to cache role index', err);
      }
    }

    return index;
  }

  function compareRoleLevels(levelA, levelB) {
    var index = getRoleIndex();
    var normalizedA = String(levelA || '').trim().toUpperCase();
    var normalizedB = String(levelB || '').trim().toUpperCase();
    var weightA = normalizedA && index[normalizedA] ? Number(index[normalizedA].weight || 0) : 0;
    var weightB = normalizedB && index[normalizedB] ? Number(index[normalizedB].weight || 0) : 0;
    return weightA - weightB;
  }

  function roleIsAtLeast(level, requiredLevel) {
    if (!requiredLevel && requiredLevel !== 0) {
      return true;
    }
    return compareRoleLevels(level, requiredLevel) >= 0;
  }

  function uniqStrings(list) {
    if (!Array.isArray(list)) return [];
    var seen = {};
    var result = [];
    list.forEach(function (item) {
      if (!item && item !== 0) return;
      var normalized = String(item).trim();
      if (!normalized) return;
      if (seen[normalized]) return;
      seen[normalized] = true;
      result.push(normalized);
    });
    return result;
  }

  function sanitizeUserReference(user) {
    if (!user || typeof user !== 'object') return null;
    var ref = {};
    if (user.ID || user.Id || user.id) {
      ref.id = String(user.ID || user.Id || user.id).trim();
    }
    if (user.UserId || user.userId) {
      ref.userId = String(user.UserId || user.userId).trim();
      if (!ref.id) ref.id = ref.userId;
    }
    if (user.UserName || user.username) {
      ref.userName = String(user.UserName || user.username).trim();
    }
    if (user.FullName || user.fullName || user.name) {
      ref.fullName = String(user.FullName || user.fullName || user.name).trim();
    }
    if (user.Email || user.email) {
      ref.email = String(user.Email || user.email).trim();
    }
    return Object.keys(ref).length ? ref : null;
  }

  function buildAuthorizationSnapshot(userPayload, options) {
    if (!userPayload || typeof userPayload !== 'object') {
      return null;
    }

    var userId = '';
    if (userPayload.ID || userPayload.Id || userPayload.id) {
      userId = String(userPayload.ID || userPayload.Id || userPayload.id).trim();
    }
    if (!userId && userPayload.UserId) {
      userId = String(userPayload.UserId).trim();
    }
    if (!userId) {
      return null;
    }

    var sessionToken = '';
    if (options && options.sessionToken) {
      sessionToken = String(options.sessionToken).trim();
    } else if (userPayload.sessionToken) {
      sessionToken = String(userPayload.sessionToken).trim();
    }

    var claims = Array.isArray(userPayload.AuthorizationClaims) ? userPayload.AuthorizationClaims.slice()
      : Array.isArray(userPayload.Claims) ? userPayload.Claims.slice()
      : [];

    var uniqueClaims = uniqStrings(claims);

    var roleHierarchy = Array.isArray(userPayload.RoleHierarchy) ? userPayload.RoleHierarchy : [];
    var normalizedRoles = roleHierarchy.map(function (role) {
      var normalized = cloneRule(role) || {};
      normalized.name = String(role && (role.name || role.label || '') || '').trim();
      return normalized;
    });

    var highestRole = null;
    if (userPayload.HighestRole && typeof userPayload.HighestRole === 'object') {
      highestRole = cloneRule(userPayload.HighestRole);
      if (highestRole) {
        highestRole.name = String(userPayload.HighestRole.name || userPayload.HighestRole.label || '').trim();
      }
    }

    var directReports = uniqStrings(userPayload.ManagedUserIds || []);
    var managedUsers = Array.isArray(userPayload.ManagedUsers)
      ? userPayload.ManagedUsers.map(sanitizeUserReference).filter(Boolean)
      : [];

    var snapshot = {
      userId: userId,
      sessionToken: sessionToken || null,
      userName: String(userPayload.UserName || '').trim(),
      fullName: String(userPayload.FullName || '').trim(),
      email: String(userPayload.Email || '').trim().toLowerCase(),
      highestRole: highestRole,
      roles: normalizedRoles,
      claims: uniqueClaims,
      roleCapabilities: cloneCapabilities(userPayload.RoleCapabilities || (userPayload.Authorization && userPayload.Authorization.roleCapabilities)),
      flags: {
        isSystemAdmin: !!userPayload.IsSystemAdmin,
        isExecutive: !!userPayload.IsExecutive,
        isManager: !!userPayload.IsManager,
        canManageUsers: !!userPayload.CanManageUsers,
        canManagePages: !!userPayload.CanManagePages
      },
      permissionLevels: uniqStrings(userPayload.PermissionLevels || []),
      campaign: {
        activeCampaignId: String(userPayload.ActiveCampaignId || userPayload.CurrentCampaignId || '').trim(),
        defaultCampaignId: String(userPayload.DefaultCampaignId || '').trim(),
        allowedCampaignIds: uniqStrings(userPayload.AllowedCampaignIds || []),
        managedCampaignIds: uniqStrings(userPayload.ManagedCampaignIds || []),
        adminCampaignIds: uniqStrings(userPayload.AdminCampaignIds || []),
        activePermission: userPayload.ActiveCampaignPermission ? JSON.parse(JSON.stringify(userPayload.ActiveCampaignPermission)) : null
      },
      directReports: directReports,
      managedUsers: managedUsers,
      directManagerId: userPayload.DirectManagerId ? String(userPayload.DirectManagerId).trim() : null,
      directManager: sanitizeUserReference(userPayload.DirectManager),
      timestamp: nowIsoString()
    };

    if (options && options.tenantPayload) {
      try {
        snapshot.tenant = JSON.parse(JSON.stringify(options.tenantPayload));
      } catch (err) {
        snapshot.tenant = null;
      }
    }

    if (options && options.rawScope) {
      try {
        snapshot.rawScope = JSON.parse(JSON.stringify(options.rawScope));
      } catch (err) {
        snapshot.rawScope = null;
      }
    }

    return snapshot;
  }

  function storeSnapshot(snapshot) {
    if (!snapshot || !snapshot.userId) {
      return null;
    }

    var json;
    try {
      json = JSON.stringify(snapshot);
    } catch (err) {
      console.warn('AuthorizationRegistry: unable to serialize snapshot for user', snapshot.userId, err);
      return null;
    }

    var cache = getScriptCache();
    if (cache) {
      safeCachePut(cache, PROFILE_CACHE_PREFIX + snapshot.userId, json, PROFILE_CACHE_TTL_SECONDS);
      if (snapshot.sessionToken) {
        safeCachePut(cache, SESSION_CACHE_PREFIX + snapshot.sessionToken, json, PROFILE_CACHE_TTL_SECONDS);
      }
    }

    var props = getScriptProperties();
    if (props) {
      try {
        props.setProperty(PROFILE_PROPERTY_PREFIX + snapshot.userId, json);
        if (snapshot.sessionToken) {
          props.setProperty(SESSION_PROPERTY_PREFIX + snapshot.sessionToken, json);
        }
      } catch (err) {
        console.warn('AuthorizationRegistry: failed to persist snapshot for user', snapshot.userId, err);
      }
    }

    return snapshot;
  }

  function parseSnapshot(serialized) {
    if (!serialized) return null;
    try {
      return JSON.parse(serialized);
    } catch (err) {
      console.warn('AuthorizationRegistry: unable to parse stored snapshot', err);
      return null;
    }
  }

  function getSnapshotFromCache(key) {
    var cache = getScriptCache();
    if (!cache) return null;
    try {
      var serialized = cache.get(key);
      return parseSnapshot(serialized);
    } catch (err) {
      console.warn('AuthorizationRegistry: cache lookup failed for key', key, err);
      return null;
    }
  }

  function getSnapshotFromProperties(key) {
    var props = getScriptProperties();
    if (!props) return null;
    try {
      var serialized = props.getProperty(key);
      return parseSnapshot(serialized);
    } catch (err) {
      console.warn('AuthorizationRegistry: properties lookup failed for key', key, err);
      return null;
    }
  }

  function getAuthorizationSnapshotForUser(userId) {
    if (!userId && userId !== 0) {
      return null;
    }
    var key = PROFILE_CACHE_PREFIX + String(userId).trim();
    var snapshot = getSnapshotFromCache(key);
    if (snapshot) {
      return snapshot;
    }

    var propertyKey = PROFILE_PROPERTY_PREFIX + String(userId).trim();
    snapshot = getSnapshotFromProperties(propertyKey);
    if (snapshot) {
      storeSnapshot(snapshot);
    }
    return snapshot;
  }

  function getAuthorizationSnapshotForSession(sessionToken) {
    if (!sessionToken && sessionToken !== 0) {
      return null;
    }
    var key = SESSION_CACHE_PREFIX + String(sessionToken).trim();
    var snapshot = getSnapshotFromCache(key);
    if (snapshot) {
      return snapshot;
    }

    var propertyKey = SESSION_PROPERTY_PREFIX + String(sessionToken).trim();
    snapshot = getSnapshotFromProperties(propertyKey);
    if (snapshot) {
      storeSnapshot(snapshot);
    }
    return snapshot;
  }

  function clearAuthorizationSnapshot(criteria) {
    var userId = criteria && criteria.userId ? String(criteria.userId).trim() : '';
    var sessionToken = criteria && criteria.sessionToken ? String(criteria.sessionToken).trim() : '';

    var cache = getScriptCache();
    if (userId) {
      safeCacheRemove(cache, PROFILE_CACHE_PREFIX + userId);
    }
    if (sessionToken) {
      safeCacheRemove(cache, SESSION_CACHE_PREFIX + sessionToken);
    }

    var props = getScriptProperties();
    if (props) {
      try {
        if (userId) props.deleteProperty(PROFILE_PROPERTY_PREFIX + userId);
        if (sessionToken) props.deleteProperty(SESSION_PROPERTY_PREFIX + sessionToken);
      } catch (err) {
        console.warn('AuthorizationRegistry: failed to clear stored snapshot', err);
      }
    }
  }

  function registerAuthorizationSnapshot(userPayload, options) {
    var snapshot = buildAuthorizationSnapshot(userPayload, options || {});
    if (!snapshot) {
      return null;
    }
    return storeSnapshot(snapshot);
  }

  function userHasCapability(subject, capability) {
    if (!capability && capability !== 0) {
      return false;
    }
    var snapshot = (subject && typeof subject === 'object' && subject.userId)
      ? subject
      : getAuthorizationSnapshotForUser(subject);
    if (!snapshot) {
      return false;
    }

    var key = String(capability).trim();
    if (!key) {
      return false;
    }

    if (snapshot.roleCapabilities && typeof snapshot.roleCapabilities === 'object' && key in snapshot.roleCapabilities) {
      return !!snapshot.roleCapabilities[key];
    }

    if (snapshot.flags && typeof snapshot.flags === 'object' && key in snapshot.flags) {
      return !!snapshot.flags[key];
    }

    return false;
  }

  function userManagesUser(managerSubject, targetUserId) {
    if (!targetUserId && targetUserId !== 0) {
      return false;
    }

    var snapshot = (managerSubject && typeof managerSubject === 'object' && managerSubject.userId)
      ? managerSubject
      : getAuthorizationSnapshotForUser(managerSubject);

    if (!snapshot) {
      return false;
    }

    if (snapshot.flags && snapshot.flags.isSystemAdmin) {
      return true;
    }

    var target = String(targetUserId).trim().toLowerCase();
    if (!target) {
      return false;
    }

    var directReports = Array.isArray(snapshot.directReports) ? snapshot.directReports : [];
    for (var i = 0; i < directReports.length; i++) {
      if (String(directReports[i]).trim().toLowerCase() === target) {
        return true;
      }
    }

    var managedUsers = Array.isArray(snapshot.managedUsers) ? snapshot.managedUsers : [];
    for (var j = 0; j < managedUsers.length; j++) {
      var managed = managedUsers[j];
      var managedId = managed && (managed.id || managed.userId);
      if (managedId && String(managedId).trim().toLowerCase() === target) {
        return true;
      }
    }

    return false;
  }

  function ensureRoleHierarchyRules(defaultRules) {
    var props = getScriptProperties();
    if (!props) {
      return DEFAULT_ROLE_RULES.slice();
    }

    try {
      var stored = props.getProperty(ROLE_RULES_PROPERTY_KEY);
      if (!stored) {
        var rulesToPersist = sanitizeRules(defaultRules && defaultRules.length ? defaultRules : DEFAULT_ROLE_RULES);
        props.setProperty(ROLE_RULES_PROPERTY_KEY, JSON.stringify(rulesToPersist));
        var cache = getScriptCache();
        if (cache) {
          cache.remove(ROLE_INDEX_CACHE_KEY);
        }
        return rulesToPersist.slice();
      }
    } catch (err) {
      console.warn('AuthorizationRegistry: unable to ensure role hierarchy rules', err);
      try {
        props.setProperty(ROLE_RULES_PROPERTY_KEY, JSON.stringify(DEFAULT_ROLE_RULES));
      } catch (persistError) {
        console.warn('AuthorizationRegistry: failed to reset role rules to defaults', persistError);
      }
    }

    return getStoredRoleRules();
  }

  function setRoleHierarchyRules(rules) {
    var sanitized = sanitizeRules(rules);
    if (!sanitized.length) {
      sanitized = DEFAULT_ROLE_RULES.slice();
    }

    persistRoleRules(sanitized);

    var cache = getScriptCache();
    if (cache) {
      try {
        cache.remove(ROLE_INDEX_CACHE_KEY);
      } catch (err) {
        console.warn('AuthorizationRegistry: unable to clear role index cache', err);
      }
    }

    return sanitized.slice();
  }

  function getRoleHierarchyRules(defaults) {
    if (defaults && defaults.length) {
      ensureRoleHierarchyRules(defaults);
    }
    var stored = getStoredRoleRules();
    return stored.slice();
  }

  function getDefaultRoleHierarchyRules() {
    return DEFAULT_ROLE_RULES.slice();
  }

  return {
    getRoleHierarchyRules: getRoleHierarchyRules,
    getDefaultRoleHierarchyRules: getDefaultRoleHierarchyRules,
    setRoleHierarchyRules: setRoleHierarchyRules,
    ensureRoleHierarchyRules: ensureRoleHierarchyRules,
    registerAuthorizationSnapshot: registerAuthorizationSnapshot,
    getAuthorizationSnapshotForUser: getAuthorizationSnapshotForUser,
    getAuthorizationSnapshotForSession: getAuthorizationSnapshotForSession,
    clearAuthorizationSnapshot: clearAuthorizationSnapshot,
    compareRoleLevels: compareRoleLevels,
    roleIsAtLeast: roleIsAtLeast,
    userHasCapability: userHasCapability,
    userManagesUser: userManagesUser
  };
})();

if (typeof module !== 'undefined' && module.exports) {
  module.exports = AuthorizationRegistry;
}
