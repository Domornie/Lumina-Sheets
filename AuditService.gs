/**
 * AuditService.gs
 * -----------------------------------------------------------------------------
 * Immutable audit logging helper. Every administrative or security-sensitive
 * event must be written to the AuditLog sheet using the schema from the Lumina
 * Identity specification.
 */
(function bootstrapAuditService(global) {
  if (!global) return;
  if (global.AuditService && typeof global.AuditService === 'object') {
    return;
  }

  var Utilities = global.Utilities;

  function getRepository() {
    var repo = global.IdentityRepository;
    if (!repo || typeof repo.append !== 'function') {
      throw new Error('AuditService requires IdentityRepository bootstrap');
    }
    return repo;
  }

  function toJson(value) {
    if (value === null || typeof value === 'undefined') return '';
    if (typeof value === 'string') return value;
    try {
      return JSON.stringify(value);
    } catch (err) {
      return String(value);
    }
  }

  function log(event) {
    if (!event) {
      throw new Error('Audit event payload is required');
    }
    var payload = {
      EventId: event.EventId || Utilities.getUuid(),
      Timestamp: event.Timestamp || new Date().toISOString(),
      ActorUserId: event.ActorUserId || '',
      ActorRole: event.ActorRole || '',
      CampaignId: event.CampaignId || '',
      Target: event.Target || '',
      Action: event.Action || '',
      BeforeJSON: toJson(event.BeforeJSON || event.Before || ''),
      AfterJSON: toJson(event.AfterJSON || event.After || ''),
      IP: event.IP || '',
      UA: event.UA || ''
    };
    getRepository().append('AuditLog', payload);
    return payload;
  }

  function list(filters) {
    var rows = getRepository().list('AuditLog');
    if (!filters) {
      return rows;
    }
    return rows.filter(function(row) {
      if (filters.campaignId && row.CampaignId !== filters.campaignId) return false;
      if (filters.target && row.Target !== filters.target) return false;
      if (filters.action && row.Action !== filters.action) return false;
      if (filters.from && new Date(row.Timestamp).getTime() < new Date(filters.from).getTime()) return false;
      if (filters.to && new Date(row.Timestamp).getTime() > new Date(filters.to).getTime()) return false;
      return true;
    });
  }

  global.AuditService = {
    log: log,
    list: list
  };
})(GLOBAL_SCOPE);
