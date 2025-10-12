/**
 * RBACService.gs
 * -----------------------------------------------------------------------------
 * Centralized role-based access control provider for Lumina Identity. Roles and
 * capabilities are loaded from the RolePermissions sheet. The service exposes
 * helpers to determine whether an actor is authorised to perform a capability
 * within a given scope.
 */
(function bootstrapRBACService(global) {
  if (!global) return;
  if (global.RBACService && typeof global.RBACService === 'object') {
    return;
  }

  function getIdentityRepository() {
    var repo = global.IdentityRepository;
    if (!repo || typeof repo.list !== 'function') {
      throw new Error('IdentityRepository not initialized');
    }
    return repo;
  }

  var CAPABILITIES = {
    VIEW_USERS: 'VIEW_USERS',
    MANAGE_USERS: 'MANAGE_USERS',
    ASSIGN_ROLES: 'ASSIGN_ROLES',
    TRANSFER_USERS: 'TRANSFER_USERS',
    TERMINATE_USERS: 'TERMINATE_USERS',
    MANAGE_EQUIPMENT: 'MANAGE_EQUIPMENT',
    VIEW_AUDIT: 'VIEW_AUDIT',
    MANAGE_POLICIES: 'MANAGE_POLICIES'
  };

  function getRolePermissions(role) {
    return getIdentityRepository().list('RolePermissions').filter(function(row) {
      return row.Role === role;
    });
  }

  function isAllowed(role, capability, campaignContext) {
    var permissions = getRolePermissions(role);
    var allowed = permissions.some(function(permission) {
      if (permission.Capability !== capability) {
        return false;
      }
      return permission.Allowed === 'Y' || permission.Allowed === true;
    });
    if (!allowed) {
      return false;
    }
    if (!campaignContext) {
      return true;
    }
    if (role === 'System Admin' || role === 'CEO' || role === 'CTO') {
      return true;
    }
    if (permissionScopeAllowsCampaign(permissions, capability, campaignContext.scope)) {
      return true;
    }
    return false;
  }

  function permissionScopeAllowsCampaign(permissions, capability, scope) {
    var relevant = permissions.filter(function(permission) {
      return permission.Capability === capability;
    });
    if (!relevant.length) {
      return false;
    }
    return relevant.some(function(permission) {
      if (!permission.Scope) {
        return false;
      }
      var normalized = String(permission.Scope).toLowerCase();
      if (normalized === 'global') {
        return true;
      }
      if (!scope) {
        return false;
      }
      if (normalized === 'campaign') {
        return scope === 'campaign';
      }
      if (normalized === 'team') {
        return scope === 'team';
      }
      return false;
    });
  }

  function campaignScopeForUser(userId, campaignId) {
    var assignments = getIdentityRepository().list('UserCampaigns').filter(function(row) {
      return row.UserId === userId && row.CampaignId === campaignId;
    });
    if (!assignments.length) {
      return null;
    }
    return assignments.map(function(assignment) {
      return assignment.Role;
    });
  }

  function assertPermission(userId, campaignId, capability, actorRoles) {
    var roles = actorRoles || campaignScopeForUser(userId, campaignId);
    if (!roles || !roles.length) {
      throw new Error('User lacks campaign assignment');
    }
    var granted = roles.some(function(role) {
      return isAllowed(role, capability, { scope: 'campaign', campaignId: campaignId });
    });
    if (!granted) {
      throw new Error('Permission denied: ' + capability);
    }
  }

  global.RBACService = {
    CAPABILITIES: CAPABILITIES,
    isAllowed: isAllowed,
    assertPermission: assertPermission,
    campaignScopeForUser: campaignScopeForUser
  };
})(GLOBAL_SCOPE);
