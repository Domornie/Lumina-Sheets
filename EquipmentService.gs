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

  function assignEquipment(actor, payload) {
    var rbac = getRBACService();
    rbac.assertPermission(actor.UserId, payload.CampaignId, rbac.CAPABILITIES.MANAGE_EQUIPMENT, actor.Roles);
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
    getIdentityRepository().upsert('Equipment', 'EquipmentId', record);
    getAuditService().log({
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
    var record = getIdentityRepository().find('Equipment', function(row) { return row.EquipmentId === equipmentId; });
    if (!record) {
      throw new Error('Equipment not found');
    }
    var rbac = getRBACService();
    rbac.assertPermission(actor.UserId, record.CampaignId, rbac.CAPABILITIES.MANAGE_EQUIPMENT, actor.Roles);
    var updated = Object.assign({}, record, updates, {
      ReturnedAt: updates.ReturnedAt || record.ReturnedAt,
      Status: updates.Status || record.Status
    });
    getIdentityRepository().upsert('Equipment', 'EquipmentId', updated);
    getAuditService().log({
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
    var rows = getIdentityRepository().list('Equipment');
    if (filters && filters.campaignId) {
      rows = rows.filter(function(row) { return row.CampaignId === filters.campaignId; });
    }
    if (filters && filters.userId) {
      rows = rows.filter(function(row) { return row.UserId === filters.userId; });
    }
    return rows;
  }

  function hasOutstandingEquipment(userId, campaignId) {
    return getIdentityRepository().list('Equipment').some(function(row) {
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
