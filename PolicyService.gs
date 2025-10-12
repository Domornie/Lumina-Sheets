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

  var IdentityRepository = global.IdentityRepository;
  var RBACService = global.RBACService;
  var Utilities = global.Utilities;

  if (!IdentityRepository || !RBACService) {
    throw new Error('PolicyService requires IdentityRepository and RBACService');
  }

  function listPolicies(scope) {
    var rows = IdentityRepository.list('Policies');
    if (!scope) {
      return rows;
    }
    return rows.filter(function(row) {
      return row.Scope === scope || row.Scope === 'Global';
    });
  }

  function upsertPolicy(actor, payload) {
    RBACService.assertPermission(actor.UserId, payload.CampaignId || '', RBACService.CAPABILITIES.MANAGE_POLICIES, actor.Roles);
    var record = {
      PolicyId: payload.PolicyId || Utilities.getUuid(),
      Name: payload.Name,
      Scope: payload.Scope || (payload.CampaignId ? payload.CampaignId : 'Global'),
      Key: payload.Key,
      Value: payload.Value,
      UpdatedAt: new Date().toISOString()
    };
    IdentityRepository.upsert('Policies', 'PolicyId', record);
    AuditService.log({
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
    return IdentityRepository.list('FeatureFlags');
  }

  function setFeatureFlag(actor, flag, value) {
    RBACService.assertPermission(actor.UserId, actor.CampaignId || '', RBACService.CAPABILITIES.MANAGE_POLICIES, actor.Roles);
    var record = {
      Flag: flag,
      Value: value,
      Notes: '',
      UpdatedAt: new Date().toISOString()
    };
    IdentityRepository.upsert('FeatureFlags', 'Flag', record);
    AuditService.log({
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
    var rows = IdentityRepository.list('EligibilityRules');
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
