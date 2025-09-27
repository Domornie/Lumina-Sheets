/**
 * TenantSecurityService.gs
 * Centralizes multi-tenant access controls for campaign-aware data.
 * - Computes per-user access profiles and allowed campaign lists
 * - Produces DatabaseManager tenant contexts for scoped CRUD
 * - Provides helper utilities to guard read/write operations against tenant leaks
 */

(function (global) {
  if (global.TenantSecurity) return;

  function toStr(value) {
    if (value === null || typeof value === 'undefined') return '';
    return String(value).trim();
  }

  function toBool(value) {
    if (value === true) return true;
    if (value === false) return false;
    var str = toStr(value).toLowerCase();
    return str === 'true' || str === '1' || str === 'yes' || str === 'y';
  }

  function logWarning(scope, err) {
    if (typeof console !== 'undefined' && console && console.warn) {
      console.warn('[TenantSecurity:' + scope + ']', err);
    }
    if (global && typeof global.safeWriteError === 'function') {
      try { global.safeWriteError('TenantSecurity.' + scope, err); } catch (_) { }
    }
  }

  function safeDbSelect(name, options, allowAllTenants) {
    var ctx = allowAllTenants ? { allowAllTenants: true } : null;
    try {
      if (typeof global.dbSelect === 'function') {
        return global.dbSelect(name, options || {}, ctx) || [];
      }
    } catch (err) {
      logWarning('safeDbSelect(' + name + ')', err);
    }

    if (typeof global.readSheet === 'function') {
      try {
        var rows = global.readSheet(name) || [];
        if (!options || !options.where) return rows;
        return (global.__dbApplyQueryOptions || global.applyQueryOptions || function (rowsInput, opts) {
          return rowsInput.filter(function (row) {
            return Object.keys(opts.where).every(function (key) {
              return toStr(row[key]) === toStr(opts.where[key]);
            });
          });
        })(rows, options);
      } catch (sheetErr) {
        logWarning('readSheet(' + name + ')', sheetErr);
      }
    }
    return [];
  }

  function loadUser(userId) {
    if (!userId) return null;
    var tableName = global.USERS_SHEET || 'Users';
    try {
      if (global.DatabaseManager) {
        return global.DatabaseManager.table(tableName).findById(userId);
      }
    } catch (err) {
      logWarning('loadUser.table', err);
    }
    var rows = safeDbSelect(tableName, { where: { ID: userId } }, true);
    return rows && rows.length ? rows[0] : null;
  }

  function loadUserCampaignAssignments(userId) {
    if (!userId) return [];
    var tableName = global.USER_CAMPAIGNS_SHEET || 'UserCampaigns';
    return safeDbSelect(tableName, { where: { UserId: userId } }, true);
  }

  function loadCampaignPermissions(userId) {
    if (!userId) return [];
    var tableName = global.CAMPAIGN_USER_PERMISSIONS_SHEET || 'CampaignUserPermissions';
    return safeDbSelect(tableName, { where: { UserID: userId } }, true);
  }

  function dedupe(values) {
    var set = {};
    var out = [];
    for (var i = 0; i < values.length; i++) {
      var key = toStr(values[i]);
      if (!key) continue;
      if (!set[key]) {
        set[key] = true;
        out.push(key);
      }
    }
    return out;
  }

  function buildAccessProfile(userId) {
    var user = loadUser(userId);
    if (!user) {
      throw new Error('User not found for tenant access profile: ' + userId);
    }

    var defaultCampaignId = toStr(user.CampaignID || user.campaignId || user.CampaignId);
    var isGlobalAdmin = toBool(user.IsAdmin);
    var assignments = loadUserCampaignAssignments(userId);
    var permissions = loadCampaignPermissions(userId);

    var allowed = [];
    if (defaultCampaignId) allowed.push(defaultCampaignId);
    assignments.forEach(function (row) {
      var cid = toStr(row.CampaignId || row.CampaignID);
      if (cid) allowed.push(cid);
    });
    permissions.forEach(function (row) {
      var cid = toStr(row.CampaignID || row.CampaignId);
      if (cid) allowed.push(cid);
    });

    var managed = [];
    var adminCampaigns = [];
    permissions.forEach(function (row) {
      var cid = toStr(row.CampaignID || row.CampaignId);
      if (!cid) return;
      var level = toStr(row.PermissionLevel).toUpperCase();
      if (level === 'MANAGER' || level === 'ADMIN') {
        managed.push(cid);
      }
      if (level === 'ADMIN') {
        adminCampaigns.push(cid);
      }
    });

    var profile = {
      userId: userId,
      user: user,
      isGlobalAdmin: isGlobalAdmin,
      defaultCampaignId: defaultCampaignId || '',
      allowedCampaignIds: dedupe(allowed),
      managedCampaignIds: dedupe(managed),
      adminCampaignIds: dedupe(adminCampaigns),
      assignments: assignments,
      permissions: permissions
    };

    profile.hasAccessTo = function (campaignId) {
      if (!campaignId) return false;
      if (profile.isGlobalAdmin) return true;
      var key = toStr(campaignId);
      return profile.allowedCampaignIds.indexOf(key) !== -1;
    };

    return profile;
  }

  function assertCampaignAccess(userId, campaignId, options) {
    options = options || {};
    var profile = buildAccessProfile(userId);
    if (profile.isGlobalAdmin) {
      return { profile: profile, campaignId: toStr(campaignId) };
    }
    var cid = toStr(campaignId);
    if (!cid) {
      throw new Error('Campaign ID is required to verify access');
    }
    if (profile.allowedCampaignIds.indexOf(cid) === -1) {
      throw new Error('User ' + userId + ' does not have access to campaign ' + cid);
    }
    if (options.requireManager && profile.managedCampaignIds.indexOf(cid) === -1) {
      throw new Error('User ' + userId + ' must be a manager for campaign ' + cid);
    }
    if (options.requireAdmin && profile.adminCampaignIds.indexOf(cid) === -1) {
      throw new Error('User ' + userId + ' must be a campaign admin for ' + cid);
    }
    return { profile: profile, campaignId: cid };
  }

  function getTenantContext(userId, campaignId, options) {
    options = options || {};
    var profile = buildAccessProfile(userId);
    var cid = toStr(campaignId);
    var context = { userId: userId };

    if (profile.isGlobalAdmin) {
      if (cid) {
        context.tenantId = cid;
      } else {
        context.allowAllTenants = true;
      }
      return { profile: profile, context: context };
    }

    if (cid) {
      assertCampaignAccess(userId, cid, options);
      context.tenantId = cid;
    } else {
      context.tenantIds = profile.allowedCampaignIds.slice();
    }

    return { profile: profile, context: context };
  }

  function getScopedTable(userId, campaignId, tableName, options) {
    if (!global.DatabaseManager) {
      throw new Error('DatabaseManager is required for scoped table access');
    }
    var ctxInfo = getTenantContext(userId, campaignId, options);
    return global.DatabaseManager.table(tableName, ctxInfo.context);
  }

  function scopedCrud(userId, campaignId, options) {
    var ctxInfo = getTenantContext(userId, campaignId, options);
    if (global.DatabaseManager) {
      return global.DatabaseManager.tenant(ctxInfo.context);
    }
    return global.dbWithContext ? global.dbWithContext(ctxInfo.context) : { context: ctxInfo.context };
  }

  function filterCampaignsForUser(userId, campaigns) {
    var list = Array.isArray(campaigns) ? campaigns.slice() : [];
    try {
      var profile = buildAccessProfile(userId);
      if (profile.isGlobalAdmin) return list;
      var allowed = {};
      profile.allowedCampaignIds.forEach(function (cid) { allowed[cid] = true; });
      return list.filter(function (campaign) {
        var cid = toStr(campaign.ID || campaign.Id || campaign.id);
        return allowed[cid];
      });
    } catch (err) {
      logWarning('filterCampaignsForUser', err);
      return list;
    }
  }

  var TenantSecurity = {
    getAccessProfile: buildAccessProfile,
    assertCampaignAccess: assertCampaignAccess,
    getTenantContext: getTenantContext,
    getScopedTable: getScopedTable,
    scopedCrud: scopedCrud,
    filterCampaignList: filterCampaignsForUser
  };

  global.TenantSecurity = TenantSecurity;
  global.getTenantAccessProfile = function (userId) { return TenantSecurity.getAccessProfile(userId); };
  global.assertTenantCampaignAccess = function (userId, campaignId, options) { return TenantSecurity.assertCampaignAccess(userId, campaignId, options); };
  global.getTenantContextForUser = function (userId, campaignId, options) { return TenantSecurity.getTenantContext(userId, campaignId, options); };
  global.getTenantScopedTable = function (userId, campaignId, tableName, options) { return TenantSecurity.getScopedTable(userId, campaignId, tableName, options); };
  global.getTenantScopedDb = function (userId, campaignId, options) { return TenantSecurity.scopedCrud(userId, campaignId, options); };
  global.filterCampaignsForUser = function (userId, campaigns) { return TenantSecurity.filterCampaignList(userId, campaigns); };
})(typeof globalThis !== 'undefined' ? globalThis : this);
