/**
 * EquipmentService.gs
 * -----------------------------------------------------------------------------
 * Handles assignment and reclamation of company assets.
 */
(function bootstrapEquipmentService(global) {
  if (!global) return;
  if (global.EquipmentService && typeof global.EquipmentService === 'object') {
    return;
  }

  var IdentityRepository = global.IdentityRepository;
  var RBACService = global.RBACService;
  var AuditService = global.AuditService;
  var Utilities = global.Utilities;

  if (!IdentityRepository || !RBACService || !AuditService) {
    throw new Error('EquipmentService dependencies missing');
  }

  function assignEquipment(actor, payload) {
    RBACService.assertPermission(actor.UserId, payload.CampaignId, RBACService.CAPABILITIES.MANAGE_EQUIPMENT, actor.Roles);
    var record = {
      EquipmentId: payload.EquipmentId || Utilities.getUuid(),
      UserId: payload.UserId,
      CampaignId: payload.CampaignId,
      Type: payload.Type,
      Serial: payload.Serial,
      Condition: payload.Condition || 'Good',
      AssignedAt: payload.AssignedAt || new Date().toISOString(),
      ReturnedAt: payload.ReturnedAt || '',
      Notes: payload.Notes || '',
      Status: payload.Status || 'Assigned'
    };
    IdentityRepository.upsert('Equipment', 'EquipmentId', record);
    AuditService.log({
      ActorUserId: actor.UserId,
      ActorRole: actor.PrimaryRole,
      CampaignId: payload.CampaignId,
      Target: 'Equipment:' + record.EquipmentId,
      Action: 'EQUIPMENT_ASSIGN',
      After: record
    });
    return record;
  }

  function updateEquipment(actor, equipmentId, updates) {
    var record = IdentityRepository.find('Equipment', function(row) { return row.EquipmentId === equipmentId; });
    if (!record) {
      throw new Error('Equipment not found');
    }
    RBACService.assertPermission(actor.UserId, record.CampaignId, RBACService.CAPABILITIES.MANAGE_EQUIPMENT, actor.Roles);
    var updated = Object.assign({}, record, updates, {
      ReturnedAt: updates.ReturnedAt || record.ReturnedAt,
      Status: updates.Status || record.Status
    });
    IdentityRepository.upsert('Equipment', 'EquipmentId', updated);
    AuditService.log({
      ActorUserId: actor.UserId,
      ActorRole: actor.PrimaryRole,
      CampaignId: record.CampaignId,
      Target: 'Equipment:' + equipmentId,
      Action: 'EQUIPMENT_UPDATE',
      Before: record,
      After: updated
    });
    return updated;
  }

  function listEquipment(actor, filters) {
    var rows = IdentityRepository.list('Equipment');
    if (filters && filters.campaignId) {
      rows = rows.filter(function(row) { return row.CampaignId === filters.campaignId; });
    }
    if (filters && filters.userId) {
      rows = rows.filter(function(row) { return row.UserId === filters.userId; });
    }
    return rows;
  }

  function hasOutstandingEquipment(userId, campaignId) {
    return IdentityRepository.list('Equipment').some(function(row) {
      return row.UserId === userId && row.CampaignId === campaignId && (!row.ReturnedAt || row.Status === 'Assigned');
    });
  }

  global.EquipmentService = {
    assignEquipment: assignEquipment,
    updateEquipment: updateEquipment,
    listEquipment: listEquipment,
    hasOutstandingEquipment: hasOutstandingEquipment
  };
})(GLOBAL_SCOPE);
