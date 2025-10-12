/**
 * LuminaAdminClient.gs
 * -----------------------------------------------------------------------------
 * Lightweight wrappers exposed to HtmlService frontends. Delegates to
 * LuminaAdmin service for data access and mutations.
 */
(function(global) {
  if (!global) return;

  function ensureService() {
    if (!global.LuminaAdmin || typeof global.LuminaAdmin.ensureSeeded !== 'function') {
      throw new Error('LuminaAdmin service not initialized');
    }
    global.LuminaAdmin.ensureSeeded();
    return global.LuminaAdmin;
  }

  function getMessages() {
    var service = ensureService();
    return service.getMessages();
  }

  function updateMessageStatus(messageId, status) {
    if (!messageId) {
      throw new Error('messageId required');
    }
    if (!status) {
      throw new Error('status required');
    }
    var service = ensureService();
    return service.updateMessageStatus(messageId, status);
  }

  function runJob(jobName) {
    if (!jobName) {
      throw new Error('jobName required');
    }
    var service = ensureService();
    return service.runJob(jobName, { mode: 'interactive' });
  }

  global.LuminaAdminClient_getMessages = getMessages;
  global.LuminaAdminClient_updateMessageStatus = updateMessageStatus;
  global.LuminaAdminClient_runJob = runJob;
})(GLOBAL_SCOPE);
