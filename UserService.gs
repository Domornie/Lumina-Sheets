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

  var IdentityRepository = global.IdentityRepository;
  var RBACService = global.RBACService;
  var AuthService = global.AuthService;
  var EquipmentService = global.EquipmentService;
  var AuditService = global.AuditService;
  var EligibilityService = global.EligibilityService;
  var Utilities = global.Utilities;

  if (!IdentityRepository || !RBACService || !AuthService || !AuditService) {
    throw new Error('UserService dependencies missing');
  }

  function listUsers(actor, campaignId) {
    RBACService.assertPermission(actor.UserId, campaignId, RBACService.CAPABILITIES.VIEW_USERS, actor.Roles);
    return IdentityRepository.list('UserCampaigns')
      .filter(function(row) { return row.CampaignId === campaignId; })
      .map(function(row) {
        var user = IdentityRepository.find('Users', function(item) { return item.UserId === row.UserId; });
        if (!user) {
          return null;
        }
        var eligibility = EligibilityService ? EligibilityService.evaluateEligibility(row.UserId, campaignId) : [];
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
    RBACService.assertPermission(actor.UserId, payload.CampaignId, RBACService.CAPABILITIES.MANAGE_USERS, actor.Roles);
    AuthService.validatePassword(payload.password);
    var now = new Date().toISOString();
    var userId = payload.UserId || Utilities.getUuid();
    var record = {
      UserId: userId,
      Email: payload.Email,
      Username: payload.Username || payload.Email,
      PasswordHash: AuthService.hashPassword(payload.password),
      EmailVerified: payload.EmailVerified ? 'Y' : 'N',
      TOTPEnabled: 'N',
      TOTPSecretHash: '',
      Status: payload.Status || 'Active',
      LastLoginAt: '',
      CreatedAt: now
    };
    IdentityRepository.upsert('Users', 'UserId', record);
    addUserCampaignAssignment(actor, userId, payload.CampaignId, payload.Role, payload.IsPrimary);
    AuditService.log({
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
    RBACService.assertPermission(actor.UserId, campaignId, RBACService.CAPABILITIES.ASSIGN_ROLES, actor.Roles);
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
    IdentityRepository.upsert('UserCampaigns', 'AssignmentId', record);
    AuditService.log({
      ActorUserId: actor.UserId,
      ActorRole: actor.PrimaryRole,
      CampaignId: campaignId,
      Target: userId,
      Action: 'ROLE_ASSIGNED',
      After: record
    });
  }

  function updateUser(actor, userId, updates) {
    var user = IdentityRepository.find('Users', function(row) { return row.UserId === userId; });
    if (!user) {
      throw new Error('User not found');
    }
    RBACService.assertPermission(actor.UserId, updates.CampaignId, RBACService.CAPABILITIES.MANAGE_USERS, actor.Roles);
    var updated = Object.assign({}, user, updates);
    if (updates.password) {
      AuthService.validatePassword(updates.password);
      updated.PasswordHash = AuthService.hashPassword(updates.password);
    }
    IdentityRepository.upsert('Users', 'UserId', updated);
    if (typeof updates.Watchlist !== 'undefined') {
      updateWatchlist(userId, updates.CampaignId, updates.Watchlist, actor);
    }
    AuditService.log({
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
    var assignment = IdentityRepository.find('UserCampaigns', function(row) {
      return row.UserId === userId && row.CampaignId === campaignId;
    });
    if (!assignment) {
      throw new Error('User is not assigned to campaign');
    }
    var updated = Object.assign({}, assignment, {
      Watchlist: value ? 'Y' : 'N'
    });
    IdentityRepository.upsert('UserCampaigns', 'AssignmentId', updated);
    AuditService.log({
      ActorUserId: actor.UserId,
      ActorRole: actor.PrimaryRole,
      CampaignId: campaignId,
      Target: userId,
      Action: value ? 'WATCHLIST_ADDED' : 'WATCHLIST_REMOVED'
    });
  }

  function transferUser(actor, userId, toCampaignId) {
    var existing = IdentityRepository.list('UserCampaigns').filter(function(row) { return row.UserId === userId; });
    if (!existing.length) {
      throw new Error('User has no assignments');
    }
    RBACService.assertPermission(actor.UserId, toCampaignId, RBACService.CAPABILITIES.TRANSFER_USERS, actor.Roles);
    var roleToCarry = existing[0].Role;
    existing.forEach(function(assignment) {
      IdentityRepository.remove('UserCampaigns', 'AssignmentId', assignment.AssignmentId);
    });
    addUserCampaignAssignment(actor, userId, toCampaignId, roleToCarry || 'Team Member', true);
    AuditService.log({
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
    RBACService.assertPermission(actor.UserId, campaignId, RBACService.CAPABILITIES.MANAGE_USERS, actor.Roles);
    if (['Terminated', 'Resigned'].indexOf(payload.State) >= 0 && !payload.Override && EquipmentService.hasOutstandingEquipment(userId, campaignId)) {
      throw new Error('Outstanding equipment must be returned before termination.');
    }
    var existingUser = IdentityRepository.find('Users', function(row) { return row.UserId === userId; });
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
      IdentityRepository.upsert('Users', 'UserId', Object.assign({}, existingUser, { Status: nextStatus }));
    }
    IdentityRepository.append('EmploymentStatus', record);
    AuditService.log({
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
    RBACService.assertPermission(actor.UserId, campaignId, RBACService.CAPABILITIES.VIEW_USERS, actor.Roles);
    var user = IdentityRepository.find('Users', function(row) { return row.UserId === userId; });
    if (!user) {
      throw new Error('User not found');
    }
    var assignments = IdentityRepository.list('UserCampaigns').filter(function(row) {
      return row.UserId === userId;
    });
    var equipment = EquipmentService.listEquipment(actor, { userId: userId, campaignId: campaignId });
    var eligibility = EligibilityService ? EligibilityService.evaluateEligibility(userId, campaignId) : [];
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
