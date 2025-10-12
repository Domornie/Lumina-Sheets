/**
 * UserService.gs
 * -----------------------------------------------------------------------------
 * User administration, lifecycle, and transfer management.
 */
(function bootstrapUserService(global) {
  if (!global) return;
  if (global.UserService && typeof global.UserService === 'object') {
    return;
  }

  var Utilities = global.Utilities;
  var dependencies = null;

  function ensureDependencies() {
    if (dependencies) {
      return dependencies;
    }
    var repo = global.IdentityRepository;
    var rbac = global.RBACService;
    var auth = global.AuthService;
    var audit = global.AuditService;
    if (!repo || !rbac || !auth || !audit) {
      throw new Error('UserService dependencies not initialized');
    }
    dependencies = {
      repository: repo,
      rbac: rbac,
      auth: auth,
      audit: audit,
      equipment: global.EquipmentService || null,
      eligibility: global.EligibilityService || null
    };
    return dependencies;
  }

  function getRepository() { return ensureDependencies().repository; }
  function getRBAC() { return ensureDependencies().rbac; }
  function getAuthService() { return ensureDependencies().auth; }
  function getAuditService() { return ensureDependencies().audit; }
  function getEquipmentService() { return ensureDependencies().equipment; }
  function getEligibilityService() { return ensureDependencies().eligibility; }

  function listUsers(actor, campaignId) {
    var rbac = getRBAC();
    rbac.assertPermission(actor.UserId, campaignId, rbac.CAPABILITIES.VIEW_USERS, actor.Roles);
    var repo = getRepository();
    var eligibilityService = getEligibilityService();
    return repo.list('UserCampaigns')
      .filter(function(row) { return row.CampaignId === campaignId; })
      .map(function(row) {
        var user = repo.find('Users', function(item) { return item.UserId === row.UserId; });
        if (!user) {
          return null;
        }
        var eligibility = eligibilityService ? eligibilityService.evaluateEligibility(row.UserId, campaignId) : [];
        return {
          AssignmentId: row.AssignmentId,
          UserId: user.UserId,
          Email: user.Email,
          Username: user.Username,
          Role: row.Role,
          Status: user.Status,
          Watchlist: row.Watchlist,
          Eligibility: eligibility
        };
      })
      .filter(Boolean);
  }

  function createUser(actor, payload) {
    var rbac = getRBAC();
    rbac.assertPermission(actor.UserId, payload.CampaignId, rbac.CAPABILITIES.MANAGE_USERS, actor.Roles);
    var auth = getAuthService();
    auth.validatePassword(payload.password);
    var now = new Date().toISOString();
    var userId = payload.UserId || Utilities.getUuid();
    var record = {
      UserId: userId,
      Email: payload.Email,
      Username: payload.Username || payload.Email,
      PasswordHash: auth.hashPassword(payload.password),
      EmailVerified: payload.EmailVerified ? 'Y' : 'N',
      TOTPEnabled: 'N',
      TOTPSecretHash: '',
      Status: payload.Status || 'Active',
      LastLoginAt: '',
      CreatedAt: now
    };
    getRepository().upsert('Users', 'UserId', record);
    addUserCampaignAssignment(actor, userId, payload.CampaignId, payload.Role, payload.IsPrimary);
    getAuditService().log({
      ActorUserId: actor.UserId,
      ActorRole: actor.PrimaryRole,
      CampaignId: payload.CampaignId,
      Target: userId,
      Action: 'USER_CREATED',
      After: record
    });
    return record;
  }

  function addUserCampaignAssignment(actor, userId, campaignId, role, isPrimary) {
    var rbac = getRBAC();
    rbac.assertPermission(actor.UserId, campaignId, rbac.CAPABILITIES.ASSIGN_ROLES, actor.Roles);
    var record = {
      AssignmentId: Utilities.getUuid(),
      UserId: userId,
      CampaignId: campaignId,
      Role: role,
      IsPrimary: isPrimary ? 'Y' : 'N',
      AddedBy: actor.UserId,
      AddedAt: new Date().toISOString(),
      Watchlist: 'N'
    };
    getRepository().upsert('UserCampaigns', 'AssignmentId', record);
    getAuditService().log({
      ActorUserId: actor.UserId,
      ActorRole: actor.PrimaryRole,
      CampaignId: campaignId,
      Target: userId,
      Action: 'ROLE_ASSIGNED',
      After: record
    });
  }

  function updateUser(actor, userId, updates) {
    var repo = getRepository();
    var user = repo.find('Users', function(row) { return row.UserId === userId; });
    if (!user) {
      throw new Error('User not found');
    }
    var rbac = getRBAC();
    rbac.assertPermission(actor.UserId, updates.CampaignId, rbac.CAPABILITIES.MANAGE_USERS, actor.Roles);
    var updated = Object.assign({}, user, updates);
    if (updates.password) {
      var auth = getAuthService();
      auth.validatePassword(updates.password);
      updated.PasswordHash = auth.hashPassword(updates.password);
    }
    repo.upsert('Users', 'UserId', updated);
    if (typeof updates.Watchlist !== 'undefined') {
      updateWatchlist(userId, updates.CampaignId, updates.Watchlist, actor);
    }
    getAuditService().log({
      ActorUserId: actor.UserId,
      ActorRole: actor.PrimaryRole,
      CampaignId: updates.CampaignId,
      Target: userId,
      Action: 'USER_UPDATED',
      Before: user,
      After: updated
    });
    return updated;
  }

  function updateWatchlist(userId, campaignId, value, actor) {
    var repo = getRepository();
    var assignment = repo.find('UserCampaigns', function(row) {
      return row.UserId === userId && row.CampaignId === campaignId;
    });
    if (!assignment) {
      throw new Error('User is not assigned to campaign');
    }
    var updated = Object.assign({}, assignment, {
      Watchlist: value ? 'Y' : 'N'
    });
    repo.upsert('UserCampaigns', 'AssignmentId', updated);
    getAuditService().log({
      ActorUserId: actor.UserId,
      ActorRole: actor.PrimaryRole,
      CampaignId: campaignId,
      Target: userId,
      Action: value ? 'WATCHLIST_ADDED' : 'WATCHLIST_REMOVED'
    });
  }

  function transferUser(actor, userId, toCampaignId) {
    var repo = getRepository();
    var existing = repo.list('UserCampaigns').filter(function(row) { return row.UserId === userId; });
    if (!existing.length) {
      throw new Error('User has no assignments');
    }
    var rbac = getRBAC();
    rbac.assertPermission(actor.UserId, toCampaignId, rbac.CAPABILITIES.TRANSFER_USERS, actor.Roles);
    var roleToCarry = existing[0].Role;
    existing.forEach(function(assignment) {
      repo.remove('UserCampaigns', 'AssignmentId', assignment.AssignmentId);
    });
    addUserCampaignAssignment(actor, userId, toCampaignId, roleToCarry || 'Team Member', true);
    getAuditService().log({
      ActorUserId: actor.UserId,
      ActorRole: actor.PrimaryRole,
      CampaignId: toCampaignId,
      Target: userId,
      Action: 'USER_TRANSFERRED',
      Before: existing,
      After: { campaignId: toCampaignId }
    });
  }

  function updateLifecycle(actor, userId, payload) {
    var campaignId = payload.CampaignId;
    var rbac = getRBAC();
    rbac.assertPermission(actor.UserId, campaignId, rbac.CAPABILITIES.MANAGE_USERS, actor.Roles);
    var equipmentService = getEquipmentService();
    if (['Terminated', 'Resigned'].indexOf(payload.State) >= 0 && !payload.Override && equipmentService && equipmentService.hasOutstandingEquipment(userId, campaignId)) {
      throw new Error('Outstanding equipment must be returned before termination.');
    }
    var repo = getRepository();
    var existingUser = repo.find('Users', function(row) { return row.UserId === userId; });
    var record = {
      UserId: userId,
      CampaignId: campaignId,
      State: payload.State,
      EffectiveDate: payload.EffectiveDate || new Date().toISOString(),
      Reason: payload.Reason || '',
      Notes: payload.Notes || ''
    };
    if (existingUser) {
      var nextStatus = existingUser.Status;
      if (payload.State === 'Terminated' || payload.State === 'Resigned') {
        nextStatus = 'Locked';
      }
      repo.upsert('Users', 'UserId', Object.assign({}, existingUser, { Status: nextStatus }));
    }
    repo.append('EmploymentStatus', record);
    getAuditService().log({
      ActorUserId: actor.UserId,
      ActorRole: actor.PrimaryRole,
      CampaignId: campaignId,
      Target: userId,
      Action: 'LIFECYCLE_UPDATE',
      After: record
    });
    return record;
  }

  function getUserProfile(actor, userId, campaignId) {
    var rbac = getRBAC();
    rbac.assertPermission(actor.UserId, campaignId, rbac.CAPABILITIES.VIEW_USERS, actor.Roles);
    var repo = getRepository();
    var user = repo.find('Users', function(row) { return row.UserId === userId; });
    if (!user) {
      throw new Error('User not found');
    }
    var assignments = repo.list('UserCampaigns').filter(function(row) {
      return row.UserId === userId;
    });
    var equipmentService = getEquipmentService();
    var equipment = equipmentService ? equipmentService.listEquipment(actor, { userId: userId, campaignId: campaignId }) : [];
    var eligibilityService = getEligibilityService();
    var eligibility = eligibilityService ? eligibilityService.evaluateEligibility(userId, campaignId) : [];
    return {
      user: user,
      assignments: assignments,
      equipment: equipment,
      eligibility: eligibility
    };
  }

  global.UserService = {
    listUsers: listUsers,
    createUser: createUser,
    updateUser: updateUser,
    transferUser: transferUser,
    updateLifecycle: updateLifecycle,
    getUserProfile: getUserProfile
  };
})(GLOBAL_SCOPE);
