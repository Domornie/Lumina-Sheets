/**
 * PolicyService.gs
 * -----------------------------------------------------------------------------
 * Provides read/write helpers for Policies, FeatureFlags, and EligibilityRules.
 */
(function bootstrapPolicyService(global) {
  if (!global) return;
  if (global.PolicyService && typeof global.PolicyService === 'object') {
    return;
  }

  var Utilities = global.Utilities;

  function getIdentityRepository() {
    var repo = global.IdentityRepository;
    if (!repo || typeof repo.list !== 'function') {
      throw new Error('IdentityRepository not initialized');
    }
    return repo;
  }

  function getRBACService() {
    var service = global.RBACService;
    if (!service || typeof service.assertPermission !== 'function') {
      throw new Error('RBACService not initialized');
    }
    return service;
  }

  function getAuditService() {
    var service = global.AuditService;
    if (!service || typeof service.log !== 'function') {
      throw new Error('AuditService not initialized');
    }
    return service;
  }

  function listPolicies(scope) {
    var rows = getIdentityRepository().list('Policies');
    if (!scope) {
      return rows;
    }
    return rows.filter(function(row) {
      return row.Scope === scope || row.Scope === 'Global';
    });
  }

  function upsertPolicy(actor, payload) {
    var rbac = getRBACService();
    rbac.assertPermission(actor.UserId, payload.CampaignId || '', rbac.CAPABILITIES.MANAGE_POLICIES, actor.Roles);
    var record = {
      PolicyId: payload.PolicyId || Utilities.getUuid(),
      Name: payload.Name,
      Scope: payload.Scope || (payload.CampaignId ? payload.CampaignId : 'Global'),
      Key: payload.Key,
      Value: payload.Value,
      UpdatedAt: new Date().toISOString()
    };
    getIdentityRepository().upsert('Policies', 'PolicyId', record);
    getAuditService().log({
      ActorUserId: actor.UserId,
      ActorRole: actor.PrimaryRole,
      CampaignId: payload.CampaignId || '',
      Target: 'Policy:' + record.Key,
      Action: 'POLICY_UPSERT',
      After: record
    });
    return record;
  }

  function listFeatureFlags() {
    return getIdentityRepository().list('FeatureFlags');
  }

  function setFeatureFlag(actor, flag, value) {
    var rbac = getRBACService();
    rbac.assertPermission(actor.UserId, actor.CampaignId || '', rbac.CAPABILITIES.MANAGE_POLICIES, actor.Roles);
    var record = {
      Flag: flag,
      Value: value,
      Notes: '',
      UpdatedAt: new Date().toISOString()
    };
    getIdentityRepository().upsert('FeatureFlags', 'Flag', record);
    getAuditService().log({
      ActorUserId: actor.UserId,
      ActorRole: actor.PrimaryRole,
      CampaignId: actor.CampaignId || '',
      Target: 'FeatureFlag:' + flag,
      Action: 'FLAG_UPDATE',
      After: record
    });
    return record;
  }

  function listEligibilityRules(campaignId) {
    var rows = getIdentityRepository().list('EligibilityRules');
    return rows.filter(function(row) {
      return row.Scope === 'Global' || row.Scope === campaignId;
    });
  }

  global.PolicyService = {
    listPolicies: listPolicies,
    upsertPolicy: upsertPolicy,
    listFeatureFlags: listFeatureFlags,
    setFeatureFlag: setFeatureFlag,
    listEligibilityRules: listEligibilityRules
  };
})(GLOBAL_SCOPE);
