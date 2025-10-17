/** Enhanced Multi-Campaign Google Apps Script - Code.gs
 * Simplified Token-Based Authentication System
 * 
 * Features:
 * - Token-based authentication
 * - Campaign-specific routing (e.g., CreditSuite.QAForm, IBTR.QACollabList)
 * - Enhanced authentication and access control
 * - Clean URL structure
 * - Secure session management via tokens
 */

// ───────────────────────────────────────────────────────────────────────────────
// GLOBAL CONSTANTS AND CONFIGURATION
// ───────────────────────────────────────────────────────────────────────────────

var GLOBAL_SCOPE = (typeof GLOBAL_SCOPE !== 'undefined') ? GLOBAL_SCOPE
  : (typeof globalThis === 'object' && globalThis)
    ? globalThis
    : (typeof this === 'object' && this)
      ? this
      : {};

const SCRIPT_URL = 'https://script.google.com/a/macros/vlbpo.com/s/AKfycbxeQ0AnupBHM71M6co3LVc5NPrxTblRXLd6AuTOpxMs2rMehF9dBSkGykIcLGHROywQ/exec';
const FAVICON_URL = 'https://res.cloudinary.com/dr8qd3xfc/image/upload/v1754763514/vlbpo/lumina/3_dgitcx.png';

/** Toggle for debug traces */
const ACCESS_DEBUG = true;

/** Canonical page access definitions */
const ACCESS = {
  ADMIN_ONLY_PAGES: new Set(['admin.users', 'admin.roles', 'admin.campaigns']),
  PUBLIC_PAGES: new Set([
    'login',
    'landing',
    'landing-about',
    'landing-capabilities',
    'setpassword',
    'resetpassword',
    'forgotpassword',
    'forgot-password',
    'resendverification',
    'resend-verification',
    'emailconfirmed',
    'email-confirmed',
    'terms-of-service',
    'privacy-policy',
    'lumina-user-guide'
  ]),
  DEFAULT_PAGE: 'dashboard',
  PRIVS: { SYSTEM_ADMIN: 'SYSTEM_ADMIN', MANAGE_USERS: 'MANAGE_USERS', MANAGE_PAGES: 'MANAGE_PAGES' }
};

// ───────────────────────────────────────────────────────────────────────────────
// LUMINA ENTITY REGISTRY
// ───────────────────────────────────────────────────────────────────────────────

const LUMINA_ENTITY_REGISTRY = (function buildEntityRegistry() {
  var registry = Object.create(null);

  function sliceHeaders(headers) {
    return Array.isArray(headers) ? headers.slice() : null;
  }

  function coerceString(value) {
    return value === null || typeof value === 'undefined' ? '' : String(value);
  }

  var qualityTableName = (typeof QA_RECORDS === 'string' && QA_RECORDS) ? QA_RECORDS : 'Quality';
  var qualitySchema = (typeof QA_HEADERS !== 'undefined' && Array.isArray(QA_HEADERS))
    ? { headers: sliceHeaders(QA_HEADERS), idColumn: 'ID' }
    : { idColumn: 'ID' };
  qualitySchema.cacheTTL = 2700;
  registry.quality = {
    name: 'quality',
    tableName: qualityTableName,
    idColumn: 'ID',
    summaryColumns: ['ID', 'Timestamp', 'AgentName', 'TotalScore', 'Percentage', 'FeedbackShared'],
    summaryOptions: { sortBy: 'Timestamp', sortDesc: true },
    schema: qualitySchema,
    normalizeSummary: function (row) {
      return {
        id: row.ID,
        agentName: coerceString(row.AgentName || row.agentName),
        timestamp: row.Timestamp || row.timestamp || '',
        totalScore: row.TotalScore || row.totalScore || '',
        percentage: row.Percentage || row.percentage || '',
        feedbackShared: row.FeedbackShared || row.feedbackShared || ''
      };
    },
    normalizeDetail: function (row) { return row; }
  };
  registry.qa = registry.quality;
  registry.qualityrecords = registry.quality;

  var usersTableName = (typeof USERS_SHEET === 'string' && USERS_SHEET) ? USERS_SHEET : 'Users';
  var usersSchema = (typeof USERS_HEADERS !== 'undefined' && Array.isArray(USERS_HEADERS))
    ? { headers: sliceHeaders(USERS_HEADERS), idColumn: 'ID' }
    : { idColumn: 'ID' };
  usersSchema.cacheTTL = 1800;
  registry.users = {
    name: 'users',
    tableName: usersTableName,
    idColumn: 'ID',
    summaryColumns: ['ID', 'FullName', 'UserName', 'CampaignID'],
    schema: usersSchema,
    normalizeSummary: function (row) {
      var fullName = coerceString(row.FullName || row.fullName);
      var userName = coerceString(row.UserName || row.userName || row.Username);
      return {
        id: row.ID,
        displayName: fullName || userName || row.ID,
        fullName: fullName,
        userName: userName,
        campaignId: coerceString(row.CampaignID || row.CampaignId)
      };
    },
    normalizeDetail: function (row) {
      return {
        id: row.ID,
        fullName: coerceString(row.FullName || row.fullName),
        userName: coerceString(row.UserName || row.userName),
        email: coerceString(row.Email || row.email || row.EmailAddress),
        record: row
      };
    }
  };
  registry.user = registry.users;

  return registry;
})();

function resolveLuminaEntityDefinition(entityName) {
  var key = String(entityName || '').toLowerCase();
  if (!key) {
    throw new Error('Entity name is required.');
  }
  var def = LUMINA_ENTITY_REGISTRY[key];
  if (!def) {
    throw new Error('Unknown entity: ' + entityName);
  }
  return def;
}

function ensureEntitySchema(definition) {
  if (!definition || !definition.tableName) {
    throw new Error('Invalid entity definition.');
  }

  try {
    if (definition.schema && typeof registerTableSchema === 'function') {
      registerTableSchema(definition.tableName, definition.schema);
    } else if (typeof DatabaseManager !== 'undefined' && DatabaseManager && typeof DatabaseManager.defineTable === 'function') {
      DatabaseManager.defineTable(definition.tableName, { idColumn: definition.idColumn || 'ID' });
    }
  } catch (schemaError) {
    console.warn('ensureEntitySchema: unable to register schema for ' + definition.tableName + ':', schemaError);
  }
}

function projectEntityRows(definition, context, options, columns) {
  var manager = (typeof DatabaseManager !== 'undefined') ? DatabaseManager : null;
  if (!manager || typeof manager.table !== 'function') {
    throw new Error('DatabaseManager.table is not available.');
  }

  ensureEntitySchema(definition);

  var table = manager.table(definition.tableName, context);
  var cols = Array.isArray(columns) ? columns.slice() : null;
  if (cols && definition.idColumn && cols.indexOf(definition.idColumn) === -1) {
    cols.push(definition.idColumn);
  }

  if (cols && typeof table.project === 'function') {
    return table.project(cols, options || {});
  }

  var opts = Object.assign({}, options || {});
  if (cols) {
    opts.columns = cols;
  }
  return table.read(opts);
}

function getEntitySummaries(entityName, context) {
  try {
    var def = resolveLuminaEntityDefinition(entityName);
    var options = def.summaryOptions ? Object.assign({}, def.summaryOptions) : {};
    var rows = projectEntityRows(def, context, options, def.summaryColumns);
    if (typeof def.normalizeSummary === 'function') {
      return rows.map(function (row) { return def.normalizeSummary(row); });
    }
    return rows;
  } catch (error) {
    console.error('getEntitySummaries failed for "' + entityName + '":', error);
    throw error;
  }
}

function getEntityDetail(entityName, id, context) {
  try {
    if (!id && id !== 0) {
      throw new Error('Record id is required.');
    }

    var def = resolveLuminaEntityDefinition(entityName);
    var manager = (typeof DatabaseManager !== 'undefined') ? DatabaseManager : null;
    if (!manager || typeof manager.table !== 'function') {
      throw new Error('DatabaseManager.table is not available.');
    }

    ensureEntitySchema(def);
    var table = manager.table(def.tableName, context);
    var record = (typeof table.findById === 'function') ? table.findById(id) : null;
    if (!record && typeof table.findOne === 'function' && def.idColumn) {
      var where = {};
      where[def.idColumn] = id;
      record = table.findOne(where);
    }
    if (!record) {
      return null;
    }
    return (typeof def.normalizeDetail === 'function') ? def.normalizeDetail(record) : record;
  } catch (error) {
    console.error('getEntityDetail failed for "' + entityName + '":', error);
    throw error;
  }
}

/**
 * Utility helpers for managing time-driven triggers that may have become
 * orphaned after refactors. These are meant to be run manually from the Apps
 * Script editor when cleaning up legacy jobs such as the removed
 * `checkRealtimeUpdatesJob` trigger.
 */

/**
 * Returns a summary of all project triggers and logs it to Stackdriver.
 * Useful for debugging lingering time-driven triggers.
 */
function listProjectTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var summary = triggers.map(function (t) {
    var details = null;
    try {
      details = t.getTriggerSourceId ? t.getTriggerSourceId() : null;
    } catch (e) {
      details = null;
    }
    return {
      handler: t.getHandlerFunction(),
      type: t.getEventType(),
      details: details
    };
  });
  console.log('[listProjectTriggers] ' + JSON.stringify(summary));
  return summary;
}

/**
 * Removes the legacy `checkRealtimeUpdatesJob` trigger if it still exists.
 * Invoke this once from the Script Editor to stop Apps Script from trying to
 * run the deleted function.
 */
function removeLegacyRealtimeTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var trigger = triggers[i];
    if (trigger.getHandlerFunction && trigger.getHandlerFunction() === 'checkRealtimeUpdatesJob') {
      ScriptApp.deleteTrigger(trigger);
      console.log('[removeLegacyRealtimeTrigger] Deleted trigger for checkRealtimeUpdatesJob');
    }
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// REALTIME UPDATE JOB CONFIGURATION
// ───────────────────────────────────────────────────────────────────────────────

const REALTIME_JOB_DEFAULT_MAX_RUNTIME_MS = 60 * 1000; // 1 minute safety window
const REALTIME_JOB_DEFAULT_MIN_INTERVAL_MS = 5 * 60 * 1000; // Do not run more than every 5 minutes
const REALTIME_JOB_DEFAULT_SLEEP_MS = 250; // Pause between handler batches
const REALTIME_JOB_LOCK_WAIT_MS = 5000; // Wait up to 5 seconds to acquire the script lock
const REALTIME_JOB_LAST_RUN_PROP = 'REALTIME_JOB_LAST_RUN_AT';
const REALTIME_JOB_LAST_SUCCESS_PROP = 'REALTIME_JOB_LAST_SUCCESS_AT';
const REALTIME_JOB_STATUS_PROP = 'REALTIME_JOB_STATUS';

/**
 * Utility helpers for managing time-driven triggers that may have become
 * orphaned after refactors. These are meant to be run manually from the Apps
 * Script editor when cleaning up or rescheduling jobs such as
 * `checkRealtimeUpdatesJob`.
 */

/**
 * Returns a summary of all project triggers and logs it to Stackdriver.
 * Useful for debugging lingering time-driven triggers.
 */
function listProjectTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var summary = triggers.map(function (t) {
    var details = null;
    try {
      details = t.getTriggerSourceId ? t.getTriggerSourceId() : null;
    } catch (e) {
      details = null;
    }
    return {
      handler: t.getHandlerFunction(),
      type: t.getEventType(),
      details: details
    };
  });
  console.log('[listProjectTriggers] ' + JSON.stringify(summary));
  return summary;
}

/**
 * Removes the legacy `checkRealtimeUpdatesJob` trigger if it still exists.
 * Invoke this once from the Script Editor to stop Apps Script from trying to
 * run the deleted function.
 */
function removeLegacyRealtimeTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var trigger = triggers[i];
    if (trigger.getHandlerFunction && trigger.getHandlerFunction() === 'checkRealtimeUpdatesJob') {
      ScriptApp.deleteTrigger(trigger);
      console.log('[removeLegacyRealtimeTrigger] Deleted trigger for checkRealtimeUpdatesJob');
    }
  }
}

/**
 * Time-driven job that checks for realtime updates without exceeding the
 * configured execution window. The job self-throttles by tracking its own
 * runtime in Script Properties so repeated triggers cannot overlap or hog the
 * Apps Script runtime.
 */
function checkRealtimeUpdatesJob() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(REALTIME_JOB_LOCK_WAIT_MS)) {
    console.log('[checkRealtimeUpdatesJob] Another run is already in progress; skipping.');
    return;
  }

  var props = PropertiesService.getScriptProperties();
  var config = getRealtimeJobConfig(props);
  var now = Date.now();
  var lastRun = Number(props.getProperty(REALTIME_JOB_LAST_RUN_PROP)) || 0;
  if (lastRun && now - lastRun < config.minIntervalMs) {
    console.log('[checkRealtimeUpdatesJob] Last run was ' + Math.round((now - lastRun) / 1000) + 's ago; waiting ' + Math.round(config.minIntervalMs / 1000) + 's between executions.');
    lock.releaseLock();
    return;
  }

  props.setProperty(REALTIME_JOB_LAST_RUN_PROP, String(now));
  props.setProperty(REALTIME_JOB_STATUS_PROP, 'running');

  try {
    var handlers = getRealtimeUpdateHandlers();
    if (!handlers.length) {
      console.log('[checkRealtimeUpdatesJob] No realtime handlers registered; exiting early.');
      props.setProperty(REALTIME_JOB_STATUS_PROP, 'idle');
      props.setProperty(REALTIME_JOB_LAST_SUCCESS_PROP, String(Date.now()));
      return;
    }

    var start = now;
    var iteration = 0;
    var hasMoreWork = true;
    var workPerformed = false;
    while (hasMoreWork && Date.now() - start < config.maxRuntimeMs) {
      hasMoreWork = false;
      for (var i = 0; i < handlers.length; i++) {
        var handler = handlers[i];
        var handlerHasMore = false;
        try {
          handlerHasMore = runRealtimeUpdateHandler(handler, iteration, config);
        } catch (handlerError) {
          if (typeof logError === 'function') {
            logError('checkRealtimeUpdatesJob.handler', handlerError);
          } else {
            console.error('[checkRealtimeUpdatesJob] Handler error', handlerError);
          }
        }
        if (handlerHasMore) {
          hasMoreWork = true;
          workPerformed = true;
        }
      }
      iteration++;
      if (hasMoreWork && config.sleepMs > 0) {
        Utilities.sleep(config.sleepMs);
      }
    }

    if (!workPerformed) {
      console.log('[checkRealtimeUpdatesJob] No realtime updates were processed during this window.');
    } else if (hasMoreWork) {
      console.log('[checkRealtimeUpdatesJob] Max runtime reached; remaining work will continue on the next trigger.');
    }

    props.setProperty(REALTIME_JOB_STATUS_PROP, 'idle');
    props.setProperty(REALTIME_JOB_LAST_SUCCESS_PROP, String(Date.now()));
    props.setProperty('REALTIME_JOB_LAST_ITERATIONS', String(iteration));
  } catch (error) {
    var message = (error && error.message) ? error.message : String(error);
    props.setProperty(REALTIME_JOB_STATUS_PROP, 'error:' + message);
    if (typeof logError === 'function') {
      logError('checkRealtimeUpdatesJob', error);
    } else {
      console.error('[checkRealtimeUpdatesJob] ' + message, error);
    }
    throw error;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Reads realtime job configuration from Script Properties, falling back to the
 * defaults defined above.
 */
function getRealtimeJobConfig(props) {
  if (!props) {
    props = PropertiesService.getScriptProperties();
  }

  var config = {
    maxRuntimeMs: REALTIME_JOB_DEFAULT_MAX_RUNTIME_MS,
    minIntervalMs: REALTIME_JOB_DEFAULT_MIN_INTERVAL_MS,
    sleepMs: REALTIME_JOB_DEFAULT_SLEEP_MS
  };

  try {
    var customRuntime = Number(props.getProperty('REALTIME_JOB_MAX_RUNTIME_MS'));
    if (customRuntime && customRuntime > 0) {
      config.maxRuntimeMs = customRuntime;
    }
  } catch (e) {}

  try {
    var customInterval = Number(props.getProperty('REALTIME_JOB_MIN_INTERVAL_MS'));
    if (customInterval && customInterval > 0) {
      config.minIntervalMs = customInterval;
    }
  } catch (e2) {}

  try {
    var customSleep = Number(props.getProperty('REALTIME_JOB_SLEEP_MS'));
    if (customSleep || customSleep === 0) {
      if (customSleep >= 0) {
        config.sleepMs = customSleep;
      }
    }
  } catch (e3) {}

  return config;
}

/**
 * Collects realtime update handlers registered globally within the Apps Script
 * project. Handlers should return a truthy value when additional work remains.
 */
function getRealtimeUpdateHandlers() {
  var handlers = [];

  try {
    if (typeof globalThis !== 'undefined' && globalThis.REALTIME_UPDATE_HANDLERS && globalThis.REALTIME_UPDATE_HANDLERS.length) {
      for (var i = 0; i < globalThis.REALTIME_UPDATE_HANDLERS.length; i++) {
        if (typeof globalThis.REALTIME_UPDATE_HANDLERS[i] === 'function') {
          handlers.push(globalThis.REALTIME_UPDATE_HANDLERS[i]);
        }
      }
    }
  } catch (e) {}

  if (typeof processRealtimeUpdateQueue === 'function') {
    handlers.push(processRealtimeUpdateQueue);
  }

  if (typeof synchronizeRealtimeUpdates === 'function') {
    handlers.push(synchronizeRealtimeUpdates);
  }

  if (typeof pullRealtimeUpdates === 'function') {
    handlers.push(pullRealtimeUpdates);
  }

  if (typeof flushCampaignDirtyRows === 'function') {
    handlers.push(function () {
      flushCampaignDirtyRows();
      return false;
    });
  }

  return handlers;
}

/**
 * Executes a realtime handler and normalizes the response to a boolean
 * indicating whether additional work remains.
 */
function runRealtimeUpdateHandler(handler, iteration, config) {
  var state = {
    iteration: iteration,
    maxRuntimeMs: config.maxRuntimeMs,
    minIntervalMs: config.minIntervalMs
  };

  var result = handler(state);
  if (!result) {
    return false;
  }

  if (result === true) {
    return true;
  }

  if (typeof result === 'number') {
    return result > 0;
  }

  if (typeof result === 'object') {
    if (typeof result.hasMore === 'boolean') {
      return result.hasMore;
    }
    if (typeof result.more === 'boolean') {
      return result.more;
    }
    if (typeof result.pending === 'number') {
      return result.pending > 0;
    }
    if (typeof result.remaining === 'number') {
      return result.remaining > 0;
    }
    if (typeof result.length === 'number') {
      return result.length > 0;
    }
  }

  return false;
}

// ───────────────────────────────────────────────────────────────────────────────
// CAMPAIGN-SCOPED ROUTING SYSTEM
// ───────────────────────────────────────────────────────────────────────────────

function __case_slug__(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/[^\w\s-]/g, '')
    .trim()
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-');
}

function __case_norm__(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

/** Campaign definitions with their specific templates */
const CASE_DEFS = [
  // Credit Suite
  {
    case: 'CreditSuite',
    aliases: ['CS', 'Credit Suite', 'creditsuite', 'credit-suite'],
    idHint: 'credit-suite',
    pages: {
      QAForm: 'CreditSuiteQAForm',
      QADashboard: 'CreditSuiteQADashboard',
      QAList: 'CreditSuiteQAList',
      QAView: 'CreditSuiteQAView',
      AttendanceReports: 'CreditSuiteAttendanceReports',
      CallReports: 'CreditSuiteCallReports',
      Dashboard: 'CreditSuiteDashboard'
    }
  },

  // HiyaCar
  {
    case: 'HiyaCar',
    aliases: ['HYC', 'hiya car', 'hiya-car', 'hiyacar'],
    idHint: 'hiya-car',
    pages: {
      QAForm: 'HiyaCarQAForm',
      QADashboard: 'HiyaCarQADashboard',
      QAList: 'HiyaCarQAList',
      QAView: 'HiyaCarQAView',
      AttendanceReports: 'HiyaCarAttendanceReports',
      CallReports: 'HiyaCarCallReports',
      Dashboard: 'HiyaCarDashboard'
    }
  },

  // Benefits Resource Center (iBTR)
  {
    case: 'IBTR',
    aliases: ['Benefits Resource Center (iBTR)', 'Benefits Resource Center', 'iBTR', 'IBTR', 'benefits-resource-center-ibtr', 'BRC'],
    idHint: 'ibtr',
    pages: {
      QAForm: 'IBTRQAForm',
      QADashboard: 'IBTRQADashboard',
      QAList: 'IBTRQAList',
      QAView: 'IBTRQAView',
      QACollabList: 'IBTRQACollabList', // special page
      AttendanceReports: 'IBTRAttendanceReports',
      CallReports: 'IBTRCallReports',
      Dashboard: 'IBTRDashboard'
    }
  },

  // Independence Insurance Agency
  {
    case: 'IndependenceInsuranceAgency',
    aliases: ['Independence Insurance Agency', 'Independence', 'IIA', 'independence-insurance-agency'],
    idHint: 'independence-insurance-agency',
    pages: {
      QAForm: 'IndependenceQAForm',
      QADashboard: 'IndependenceQADashboard',
      QAList: 'IndependenceQAList',
      QAView: 'IndependenceQAView',
      AttendanceReports: 'IndependenceAttendanceReports',
      CallReports: 'IndependenceCallReports',
      Dashboard: 'IndependenceDashboard',
      CoachingForm: 'IndependenceCoachingForm'
    }
  },

  // JSC
  {
    case: 'JSC',
    aliases: ['JSC'],
    idHint: 'jsc',
    pages: {
      QAForm: 'JSCQAForm',
      QADashboard: 'JSCQADashboard',
      QAList: 'JSCQAList',
      QAView: 'JSCQAView',
      AttendanceReports: 'JSCAttendanceReports',
      CallReports: 'JSCCallReports',
      Dashboard: 'JSCDashboard'
    }
  },

  // Kids in the Game
  {
    case: 'KidsInTheGame',
    aliases: ['Kids in the Game', 'KITG', 'kids-in-the-game'],
    idHint: 'kids-in-the-game',
    pages: {
      QAForm: 'KidsInTheGameQAForm',
      QADashboard: 'KidsInTheGameQADashboard',
      QAList: 'KidsInTheGameQAList',
      QAView: 'KidsInTheGameQAView',
      AttendanceReports: 'KidsInTheGameAttendanceReports',
      CallReports: 'KidsInTheGameCallReports',
      Dashboard: 'KidsInTheGameDashboard'
    }
  },

  // Kofi Group
  {
    case: 'KofiGroup',
    aliases: ['Kofi Group', 'KOFI', 'kofi-group'],
    idHint: 'kofi-group',
    pages: {
      QAForm: 'KofiGroupQAForm',
      QADashboard: 'KofiGroupQADashboard',
      QAList: 'KofiGroupQAList',
      QAView: 'KofiGroupQAView',
      AttendanceReports: 'KofiGroupAttendanceReports',
      CallReports: 'KofiGroupCallReports',
      Dashboard: 'KofiGroupDashboard'
    }
  },

  // PAW LAW FIRM
  {
    case: 'PAWLawFirm',
    aliases: ['PAW LAW FIRM', 'PAW', 'paw-law-firm'],
    idHint: 'paw-law-firm',
    pages: {
      QAForm: 'PAWLawFirmQAForm',
      QADashboard: 'PAWLawFirmQADashboard',
      QAList: 'PAWLawFirmQAList',
      QAView: 'PAWLawFirmQAView',
      AttendanceReports: 'PAWLawFirmAttendanceReports',
      CallReports: 'PAWLawFirmCallReports',
      Dashboard: 'PAWLawFirmDashboard'
    }
  },

  // Pro House Photos
  {
    case: 'ProHousePhotos',
    aliases: ['Pro House Photos', 'PHP', 'pro-house-photos'],
    idHint: 'pro-house-photos',
    pages: {
      QAForm: 'ProHousePhotosQAForm',
      QADashboard: 'ProHousePhotosQADashboard',
      QAList: 'ProHousePhotosQAList',
      QAView: 'ProHousePhotosQAView',
      AttendanceReports: 'ProHousePhotosAttendanceReports',
      CallReports: 'ProHousePhotosCallReports',
      Dashboard: 'ProHousePhotosDashboard'
    }
  },

  // Independence Agency & Credit Suite
  {
    case: 'IndependenceAgencyCreditSuite',
    aliases: ['Independence Agency & Credit Suite', 'IACS', 'independence-agency-credit-suite'],
    idHint: 'independence-agency-credit-suite',
    pages: {
      QAForm: 'IACSQAForm',
      QADashboard: 'IACSQADashboard',
      QAList: 'IACSQAList',
      QAView: 'IACSQAView',
      AttendanceReports: 'IACSAttendanceReports',
      CallReports: 'IACSCallReports',
      Dashboard: 'IACSQADashboard'
    }
  },

  // Proozy
  {
    case: 'Proozy',
    aliases: ['Proozy', 'PRZ', 'proozy'],
    idHint: 'proozy',
    pages: {
      QAForm: 'ProozyQAForm',
      QADashboard: 'ProozyQADashboard',
      QAList: 'ProozyQAList',
      QAView: 'ProozyQAView',
      AttendanceReports: 'ProozyAttendanceReports',
      CallReports: 'ProozyCallReports',
      Dashboard: 'ProozyDashboard'
    }
  },

  // The Grounding (TGC)
  {
    case: 'TGC',
    aliases: ['The Grounding', 'TG', 'TGC', 'the-grounding'],
    idHint: 'the-grounding',
    pages: {
      QAForm: 'TheGroundingQAForm',
      QADashboard: 'TheGroundingQADashboard',
      QAList: 'TheGroundingQAList',
      QAView: 'TheGroundingQAView',
      AttendanceReports: 'TGCAttendanceReports',
      ChatReport: 'TGCChatReport', // Special: chat instead of calls
      Dashboard: 'TheGroundingDashboard'
    }
  },

  // CO
  {
    case: 'CO',
    aliases: ['CO', 'CO ()', 'co'],
    idHint: 'co',
    pages: {
      QAForm: 'COQAForm',
      QADashboard: 'COQADashboard',
      QAList: 'COQAList',
      QAView: 'COQAView',
      AttendanceReports: 'COAttendanceReports',
      CallReports: 'COCallReports',
      Dashboard: 'CODashboard'
    }
  }
];

/** Generic fallback templates */
const GENERIC_FALLBACKS = {
  QAForm: ['QualityForm'],
  QADashboard: ['QADashboard', 'UnifiedQADashboard'],
  QAList: ['QAList'],
  QAView: ['QualityView'],
  AttendanceReports: ['AttendanceReports'],
  CallReports: ['CallReports'],
  ChatReport: ['ChatReports', 'Chat'],
  Dashboard: ['Dashboard'],
  CoachingForm: ['CoachingForm'],
  TaskForm: ['TaskForm'],
  TaskList: ['TaskList'],
  TaskBoard: ['TaskBoard']
};

// ───────────────────────────────────────────────────────────────────────────────
// UTILITY FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

function _truthy(v) {
  return v === true || String(v).toUpperCase() === 'TRUE';
}

function _normalizePageKey(k) {
  return (k || '').toString().trim().toLowerCase();
}

function _normalizeId(x) {
  return (x == null ? '' : String(x)).trim();
}

function _now() {
  return new Date();
}

function _safeDate(d) {
  try {
    return new Date(d);
  } catch (e) {
    return null;
  }
}

function _isFuture(d) {
  try {
    return d && d.getTime && d.getTime() > Date.now();
  } catch (_) {
    return false;
  }
}

function _csvToSet(csv) {
  const set = new Set();
  (String(csv || '').split(',') || []).forEach(s => {
    const k = _normalizePageKey(s);
    if (k) set.add(k);
  });
  return set;
}

function _stringifyForTemplate_(obj) {
  try {
    return JSON.stringify(obj || {}).replace(/<\/script>/g, '<\\/script>');
  } catch (_) {
    return '{}';
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// CACHING HELPERS
// ───────────────────────────────────────────────────────────────────────────────

function _cacheGet(key) {
  try {
    const c = CacheService.getScriptCache().get(key);
    if (c) return JSON.parse(c);
  } catch (_) { }
  try {
    const p = PropertiesService.getScriptProperties().getProperty(key);
    if (p) return JSON.parse(p);
  } catch (_) { }
  return null;
}

function _cachePut(key, obj, ttlSec) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(obj), Math.min(21600, Math.max(5, ttlSec || 60)));
  } catch (_) { }
  try {
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(obj));
  } catch (_) { }
}


// ───────────────────────────────────────────────────────────────────────────────
// SIMPLIFIED AUTHENTICATION FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Get current user from session using Google Apps Script built-in authentication
 */
function getCurrentUser() {
  try {
    const email = String(
      (Session.getActiveUser() && Session.getActiveUser().getEmail()) ||
      (Session.getEffectiveUser() && Session.getEffectiveUser().getEmail()) ||
      ''
    ).trim().toLowerCase();

    const row = _findUserByEmail_(email);
    const client = _toClientUser_(row, email);

    // Hydrate campaign context
    try {
      const globalObj = (typeof globalThis !== 'undefined') ? globalThis : this;
      const guardKey = '__READ_SHEET_SUPPRESS_TENANT_CONTEXT__';
      const previousGuard = globalObj && globalObj[guardKey] ? Number(globalObj[guardKey]) : 0;
      if (globalObj) {
        globalObj[guardKey] = previousGuard + 1;
      }

      try {
        if (client.CampaignID) {
          if (typeof getCampaignById === 'function') {
            const c = getCampaignById(client.CampaignID);
            client.campaignName = c ? (c.Name || c.name || '') : '';
          }
          if (typeof getUserCampaignPermissions === 'function') {
            client.campaignPermissions = getUserCampaignPermissions(client.ID);
          }
          if (typeof getCampaignNavigation === 'function') {
            client.campaignNavigation = getCampaignNavigation(client.CampaignID);
          }
        }
      } finally {
        if (globalObj) {
          if (previousGuard) {
            globalObj[guardKey] = previousGuard;
          } else {
            try {
              delete globalObj[guardKey];
            } catch (_) {
              globalObj[guardKey] = 0;
            }
          }
        }
      }
    } catch (ctxErr) {
      console.warn('getCurrentUser: campaign context hydrate failed:', ctxErr);
    }

    return client;
  } catch (e) {
    writeError && writeError('getCurrentUser', e);
    return _toClientUser_(null, '');
  }
}

/**
 * Simple authentication using token parameter
 */
function authenticateUser(e) {
  try {
    // First resolve the session token from any available source
    const rawToken = e && e.parameter ? (e.parameter.token || e.parameter.sessionToken || '') : '';
    let token = rawToken;

    if (typeof LuminaIdentity !== 'undefined'
      && LuminaIdentity
      && typeof LuminaIdentity.resolveSessionToken === 'function') {
      try {
        token = LuminaIdentity.resolveSessionToken(e, { sessionToken: rawToken }) || rawToken;
      } catch (resolveError) {
        console.warn('authenticateUser: unable to resolve session token via LuminaIdentity', resolveError);
        token = rawToken;
      }
    }

    if (!token && rawToken) {
      token = rawToken;
    }

    let identity = null;
    if (token && typeof LuminaIdentity !== 'undefined' && LuminaIdentity && typeof LuminaIdentity.resolve === 'function') {
      try {
        identity = LuminaIdentity.resolve(e, { sessionToken: token, useCache: true });
      } catch (identityError) {
        console.warn('authenticateUser: unable to resolve identity via LuminaIdentity', identityError);
        identity = null;
      }
    }

    if (identity && (identity.ID || identity.id)) {
      identity.sessionToken = identity.sessionToken || token;
      return identity;
    }

    if (token && typeof AuthenticationService !== 'undefined' && AuthenticationService.getSessionUser) {
      const user = AuthenticationService.getSessionUser(token);
      if (user) {
        user.sessionToken = user.sessionToken || token;

        try {
          if (typeof LuminaIdentity !== 'undefined' && LuminaIdentity) {
            try {
              if (typeof LuminaIdentity.resolve === 'function') {
                LuminaIdentity.resolve(null, { sessionToken: token, explicitUser: user, useCache: false });
              }
            } catch (resolvePersistError) {
              console.warn('authenticateUser: unable to hydrate identity cache', resolvePersistError);
            }

            if (typeof LuminaIdentity.persistActiveSessionToken === 'function') {
              const rememberValue = (function () {
                if (Object.prototype.hasOwnProperty.call(user, 'sessionRememberMe')) {
                  return !!user.sessionRememberMe;
                }
                if (Object.prototype.hasOwnProperty.call(user, 'rememberMe')) {
                  return !!user.rememberMe;
                }
                if (Object.prototype.hasOwnProperty.call(user, 'RememberMe')) {
                  return _truthy(user.RememberMe);
                }
                return undefined;
              })();

              const metadata = {
                sessionExpiresAt: user.sessionExpiresAt || user.expiresAt || user.ExpiresAt || '',
                sessionTtlSeconds: user.sessionTtlSeconds || user.ttlSeconds || null,
                sessionIdleTimeoutMinutes: user.sessionIdleTimeoutMinutes || user.idleTimeoutMinutes || user.IdleTimeoutMinutes || null,
                lastActivityAt: new Date().toISOString()
              };

              if (rememberValue !== undefined) {
                metadata.rememberMe = rememberValue;
              }

              LuminaIdentity.persistActiveSessionToken(token, metadata);
            }
          }
        } catch (persistError) {
          console.warn('authenticateUser: unable to persist active session token', persistError);
        }

        return user;
      }
    }

    // Fall back to current user via Google session
    const user = getCurrentUser();
    if (!user || !user.ID) {
      return null;
    }

    if (typeof AuthenticationService !== 'undefined'
      && AuthenticationService
      && typeof AuthenticationService.userHasActiveSession === 'function') {
      try {
        const hasActiveSession = AuthenticationService.userHasActiveSession(user);
        if (!hasActiveSession) {
          return null;
        }
      } catch (sessionCheckError) {
        console.warn('authenticateUser: active session verification failed', sessionCheckError);
        return null;
      }
    }

    // Check if user can login
    if (!_truthy(user.CanLogin)) {
      return null;
    }

    return user;
  } catch (error) {
    console.error('Authentication error:', error);
    writeError('authenticateUser', error);
    return null;
  }
}

function getAuthenticatedUrl(page, campaignId, additionalParams = {}) {
  const baseUrl = getBaseUrl();
  const url = baseUrl || SCRIPT_URL || '';
  const queryParts = [];

  if (page) {
    queryParts.push('page=' + encodeURIComponent(page));
  }

  if (campaignId) {
    queryParts.push('campaign=' + encodeURIComponent(campaignId));
  }

  Object.keys(additionalParams || {}).forEach(function(key) {
    var value = additionalParams[key];
    if (value !== null && value !== undefined) {
      queryParts.push(encodeURIComponent(key) + '=' + encodeURIComponent(value));
    }
  });

  if (!url) {
    return queryParts.length ? ('?' + queryParts.join('&')) : '';
  }

  if (queryParts.length === 0) {
    return url;
  }

  const separator = url.indexOf('?') === -1 ? '?' : '&';
  return url + separator + queryParts.join('&');
}

function getBaseUrl() {
  function normalizeWebAppUrl(candidate) {
    if (!candidate) {
      return '';
    }

    try {
      const raw = String(candidate).trim();
      if (!raw) {
        return '';
      }

      // Ensure we always point to an /exec deployment and not the editor-facing /dev URL
      const normalized = raw.replace(/\/dev(\?|$)/i, '/exec$1');

      if (/usercodeapppanel/i.test(normalized)) {
        return normalized.replace(/\/userCodeAppPanel.*$/i, '/exec');
      }

      return normalized;
    } catch (error) {
      console.warn('getBaseUrl: unable to normalize candidate URL', error);
      return '';
    }
  }

  const candidates = [];
  const seen = new Set();

  function pushCandidate(value, label) {
    try {
      const normalized = normalizeWebAppUrl(value);
      if (normalized && !seen.has(normalized)) {
        seen.add(normalized);
        candidates.push(normalized);
      }
    } catch (err) {
      if (label) {
        console.warn('getBaseUrl: failed to add candidate from ' + label, err);
      } else {
        console.warn('getBaseUrl: failed to add candidate', err);
      }
    }
  }

  try {
    if (typeof resolveScriptUrl === 'function') {
      pushCandidate(resolveScriptUrl(), 'resolveScriptUrl');
    }
  } catch (err) {
    console.warn('getBaseUrl: resolveScriptUrl failed', err);
  }

  if (typeof SCRIPT_URL !== 'undefined' && SCRIPT_URL) {
    pushCandidate(SCRIPT_URL, 'SCRIPT_URL');
  }

  try {
    if (typeof ScriptApp !== 'undefined' && ScriptApp && ScriptApp.getService) {
      const serviceUrl = ScriptApp.getService().getUrl();
      if (serviceUrl) {
        pushCandidate(serviceUrl, 'ScriptApp');
      }
    }
  } catch (error) {
    console.warn('getBaseUrl: unable to determine script URL via ScriptApp', error);
  }

  if (!candidates.length) {
    return '';
  }

  // Prefer an /exec deployment if available.
  const execCandidate = candidates.find(function (candidate) {
    return /\/exec(\?|$)/i.test(candidate);
  });

  return execCandidate || candidates[0] || '';
}

// ───────────────────────────────────────────────────────────────────────────────
// USER MANAGEMENT FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

function _findUserByEmail_(email) {
  if (!email) return null;
  try {
    const CK = 'USR_BY_EMAIL_' + email.toLowerCase();
    const cached = _cacheGet(CK);
    if (cached) return cached;

    const rows = (typeof readSheet === 'function') ? (readSheet('Users') || []) : [];
    const hit = rows.find(r => String(r.Email || '').trim().toLowerCase() === email.toLowerCase()) || null;
    if (hit) _cachePut(CK, hit, 120);
    return hit;
  } catch (e) {
    writeError && writeError('_findUserByEmail_', e);
    return null;
  }
}

function _findUserById_(id) {
  try {
    const normalized = _normalizeId(id);
    if (!normalized) return null;

    const cacheKey = 'USR_BY_ID_' + normalized;
    try {
      const cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) return JSON.parse(cached);
    } catch (_) { }

    const rows = (typeof readSheet === 'function')
      ? (readSheet('Users', {
        suppressTenantContext: true,
        useCache: true,
        where: { ID: normalized },
        limit: 1
      }) || [])
      : [];

    const hit = Array.isArray(rows)
      ? rows.find(function (r) {
        return _normalizeId(r && (r.ID || r.Id || r.id)) === normalized;
      })
      : null;

    if (hit) {
      try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(hit), 120); } catch (_) { }
    }

    return hit || null;
  } catch (e) {
    writeError && writeError('_findUserById_', e);
    return null;
  }
}

function _toClientUser_(row, fallbackEmail) {
  const rolesMap = (typeof getRolesMapping === 'function') ? getRolesMapping() : {};
  const roleIds = String(row && row.Roles || '')
    .split(',').map(s => s.trim()).filter(Boolean);
  const roleNames = roleIds.map(id => rolesMap[id]).filter(Boolean);

  const isAdminFlag =
    (String(row && row.IsAdmin).toLowerCase() === 'true') ||
    roleNames.some(n => /system\s*admin|administrator|super\s*admin/i.test(String(n || '')));

  const client = {
    ID: String(row && row.ID || ''),
    Email: String(row && row.Email || fallbackEmail || '').trim(),
    FullName: String(row && (row.FullName || row.UserName || '') || '').trim() ||
      String(fallbackEmail || '').split('@')[0],
    CampaignID: String(row && (row.CampaignID || row.CampaignId || '') || ''),
    roleNames: roleNames,
    IsAdmin: !!isAdminFlag,
    CanLogin: (row && row.CanLogin !== undefined) ? _truthy(row.CanLogin) : true,
    EmailConfirmed: (row && row.EmailConfirmed !== undefined) ? _truthy(row.EmailConfirmed) : true,
    ResetRequired: !!(row && _truthy(row.ResetRequired)),
    Pages: String(row && (row.Pages || row.pages || '') || '')
  };
  return client;
}

function clientGetCurrentUser() {
  return getCurrentUser();
}

function __injectCurrentUser_(tpl, explicitUser) {
  try {
    const u = explicitUser && (explicitUser.Email || explicitUser.email || explicitUser.ID)
      ? explicitUser
      : getCurrentUser();
    tpl.user = u;
    tpl.currentUserJson = _stringifyForTemplate_(u);
  } catch (e) {
    writeError && writeError('__injectCurrentUser_', e);
    tpl.user = null;
    tpl.currentUserJson = '{}';
  }
}


// ───────────────────────────────────────────────────────────────────────────────
// ACCESS CONTROL SYSTEM
// ───────────────────────────────────────────────────────────────────────────────

function _normalizeUser(user) {
  const out = Object.assign({}, user || {});
  out.IsAdmin = _truthy(out.IsAdmin) || String((out.roleNames || [])).toLowerCase().includes('system admin');
  out.CanLogin = _truthy(out.CanLogin !== undefined ? out.CanLogin : true);
  out.EmailConfirmed = _truthy(out.EmailConfirmed !== undefined ? out.EmailConfirmed : true);
  out.ResetRequired = _truthy(out.ResetRequired);
  out.LockoutEnd = out.LockoutEnd ? _safeDate(out.LockoutEnd) : null;
  out.ID = _normalizeId(out.ID || out.id);
  out.CampaignID = _normalizeId(out.CampaignID || out.campaignId);
  out.PagesCsv = String(out.Pages || out.pages || '');
  return out;
}

function isSystemAdmin(user) {
  try {
    const u = _normalizeUser(user);
    return !!u.IsAdmin;
  } catch (e) {
    writeError && writeError('isSystemAdmin', e);
    return false;
  }
}

function evaluatePageAccess(user, pageKey, campaignId) {
  const trace = [];
  try {
    const u = _normalizeUser(user);
    const page = _normalizePageKey(pageKey || '');
    const cid = _normalizeId(campaignId || '');

    // Basic account checks
    if (!u || !u.ID) return { allow: false, reason: 'No session', trace };
    if (!_truthy(u.CanLogin)) return { allow: false, reason: 'Account disabled', trace };
    if (u.LockoutEnd && _isFuture(u.LockoutEnd)) return { allow: false, reason: 'Account locked', trace };
    if (!ACCESS.PUBLIC_PAGES.has(page)) {
      if (!_truthy(u.EmailConfirmed)) return { allow: false, reason: 'Email not confirmed', trace };
      if (_truthy(u.ResetRequired)) return { allow: false, reason: 'Password reset required', trace };
    }

    // Public pages
    if (ACCESS.PUBLIC_PAGES.has(page)) {
      trace.push('PUBLIC page');
      return { allow: true, reason: 'public', trace };
    }

    // Admin-only pages
    if (ACCESS.ADMIN_ONLY_PAGES.has(page)) {
      if (isSystemAdmin(u)) {
        trace.push('admin-only: allowed');
        return { allow: true, reason: 'admin', trace };
      }
      return { allow: false, reason: 'System Admin required', trace };
    }

    // System admin has access to everything
    if (isSystemAdmin(u)) {
      trace.push('System Admin: allow');
      return { allow: true, reason: 'admin', trace };
    }

    // For regular users, check if they have campaign access
    if (cid && !hasCampaignAccess(u, cid)) {
      return { allow: false, reason: 'No campaign access', trace };
    }

    // Default allow for authenticated users
    trace.push('Authenticated user access');
    return { allow: true, reason: 'authenticated', trace };

  } catch (e) {
    writeError && writeError('evaluatePageAccess', e);
    trace.push('exception:' + e.message);
    return { allow: false, reason: 'Evaluator error', trace };
  }
}

function hasCampaignAccess(user, campaignId) {
  try {
    const u = _normalizeUser(user);
    const cid = _normalizeId(campaignId);
    if (!cid) return true;
    if (isSystemAdmin(u)) return true;
    if (_normalizeId(u.CampaignID) === cid) return true;
    return false;
  } catch (e) {
    writeError && writeError('hasCampaignAccess', e);
    return false;
  }
}

function renderAccessDenied(message) {
  const baseUrl = getBaseUrl();
  const tpl = HtmlService.createTemplateFromFile('AccessDenied');
  tpl.baseUrl = baseUrl;
  tpl.message = message || 'You do not have permission to view this page.';
  return tpl.evaluate()
    .setTitle('Access Denied')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Simplified authentication requirement function
 */
function requireAuth(e) {
  try {
    const user = authenticateUser(e);

    if (!user) {
      return renderLoginPage(e);
    }

    const pageParam = String(e?.parameter?.page || '').toLowerCase();
    const page = canonicalizePageKey(pageParam);
    const campaignId = String(e?.parameter?.campaign || user.CampaignID || '');

    const decision = evaluatePageAccess(user, page, campaignId);
    _debugAccess && _debugAccess('route', decision, user, page, campaignId);

    if (!decision || decision.allow !== true) {
      return renderAccessDenied((decision && decision.reason) || 'You do not have permission to view this page.');
    }

    // Hydrate campaign context
    if (campaignId) {
      try {
        if (typeof getCampaignNavigation === 'function') user.campaignNavigation = getCampaignNavigation(campaignId);
        if (typeof getUserCampaignPermissions === 'function') user.campaignPermissions = getUserCampaignPermissions(user.ID);
        if (typeof getCampaignById === 'function') {
          const c = getCampaignById(campaignId);
          user.campaignName = c ? (c.Name || c.name || '') : '';
        }
      } catch (ctxErr) {
        console.warn('Campaign context hydrate failed:', ctxErr);
      }
    }

    return user;

  } catch (error) {
    writeError('requireAuth', error);
    return renderAccessDenied('Authentication error occurred');
  }
}

function renderLoginPage(e) {
  let serverMetadata = null;
  let initialReturnUrl = '';

  if (typeof AuthenticationService !== 'undefined'
    && AuthenticationService
    && typeof AuthenticationService.captureLoginRequestContext === 'function') {
    try {
      serverMetadata = AuthenticationService.captureLoginRequestContext(e) || null;
    } catch (contextError) {
      console.warn('renderLoginPage: unable to capture server metadata', contextError);
    }
  }

  if (typeof AuthenticationService !== 'undefined'
    && AuthenticationService
    && typeof AuthenticationService.deriveLoginReturnUrlFromEvent === 'function') {
    try {
      initialReturnUrl = AuthenticationService.deriveLoginReturnUrlFromEvent(e) || '';
    } catch (returnError) {
      console.warn('renderLoginPage: deriveLoginReturnUrlFromEvent failed', returnError);
    }
  }

  if (!initialReturnUrl && e && e.parameter) {
    try {
      const directReturn = sanitizeLoginReturnUrl(e.parameter.returnUrl || e.parameter.ReturnUrl || '');
      if (directReturn) {
        initialReturnUrl = directReturn;
      }
    } catch (directError) {
      console.warn('renderLoginPage: unable to sanitize direct returnUrl', directError);
    }
  }

  if (!initialReturnUrl && e && e.parameter) {
    const requestedPage = String(e.parameter.page || e.parameter.Page || '').trim();
    if (requestedPage && requestedPage.toLowerCase() !== 'login') {
      const additionalParams = {};
      let campaignId = '';

      Object.keys(e.parameter).forEach(function (key) {
        if (!key) return;
        if (/^page$/i.test(key)) return;
        if (/^token$/i.test(key)) return;
        if (/^returnurl$/i.test(key)) return;

        const value = e.parameter[key];
        if (value === null || typeof value === 'undefined' || value === '') {
          return;
        }

        if (!campaignId && /^campaign$/i.test(key)) {
          campaignId = value;
          return;
        }

        additionalParams[key] = value;
      });

      try {
        const builtUrl = getAuthenticatedUrl(requestedPage, campaignId, additionalParams);
        const sanitizedBuilt = sanitizeLoginReturnUrl(builtUrl);
        if (sanitizedBuilt) {
          initialReturnUrl = sanitizedBuilt;
        }
      } catch (buildReturnError) {
        console.warn('renderLoginPage: unable to build fallback return URL', buildReturnError);
      }
    }
  }

  const tpl = HtmlService.createTemplateFromFile('Login');
  tpl.baseUrl = getBaseUrl();
  tpl.scriptUrl = SCRIPT_URL;
  tpl.serverMetadataJson = JSON.stringify(serverMetadata || null);
  tpl.initialReturnUrl = initialReturnUrl || '';
  return tpl.evaluate()
    .setTitle('Login - VLBPO LuminaHQ')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function sanitizeLoginReturnUrl(raw) {
  try {
    if (typeof IdentityService !== 'undefined'
      && IdentityService
      && typeof IdentityService.sanitizeLoginReturnUrl === 'function') {
      return IdentityService.sanitizeLoginReturnUrl(raw);
    }
  } catch (identityError) {
    console.warn('sanitizeLoginReturnUrl: IdentityService helper failed', identityError);
  }

  if (!raw && raw !== 0) {
    return '';
  }

  try {
    var value = String(raw).trim();
    if (!value) {
      return '';
    }

    if (/^javascript:/i.test(value)) {
      return '';
    }

    if (/^https?:/i.test(value)) {
      try {
        var base = getBaseUrl() || '';
        if (!base && typeof SCRIPT_URL === 'string') {
          base = SCRIPT_URL;
        }

        if (base) {
          var baseMatch = /^https?:\/\/[^/]+/i.exec(base);
          var targetMatch = /^https?:\/\/[^/]+/i.exec(value);
          if (baseMatch && targetMatch && baseMatch[0].toLowerCase() !== targetMatch[0].toLowerCase()) {
            return '';
          }
        }
      } catch (originError) {
        console.warn('sanitizeLoginReturnUrl fallback origin check failed', originError);
      }
    }

    if (value.length > 500) {
      value = value.slice(0, 500);
    }

    return value;
  } catch (error) {
    console.warn('sanitizeLoginReturnUrl fallback failed', error);
    return '';
  }
}

function buildLoginPageUrl(options) {
  var base = getBaseUrl() || SCRIPT_URL || '';
  var parts = ['page=login'];

  if (options && options.returnUrl) {
    var sanitized = sanitizeLoginReturnUrl(options.returnUrl);
    if (sanitized) {
      parts.push('returnUrl=' + encodeURIComponent(sanitized));
    }
  }

  if (options && options.message) {
    parts.push('message=' + encodeURIComponent(String(options.message)));
  }

  if (options && options.error) {
    parts.push('error=' + encodeURIComponent(String(options.error)));
  }

  var query = parts.join('&');
  if (!base) {
    return '?' + query;
  }

  return base + (base.indexOf('?') === -1 ? '?' : '&') + query;
}

function handleLogoutRequest(e) {
  try {
    var token = (e && e.parameter && e.parameter.token) ? String(e.parameter.token) : '';

    try {
      if (typeof AuthenticationService !== 'undefined'
        && AuthenticationService
        && typeof AuthenticationService.logout === 'function') {
        AuthenticationService.logout(token || '');
      } else if (typeof identitySignOut === 'function') {
        identitySignOut(token || '');
      }
    } catch (logoutError) {
      console.warn('handleLogoutRequest: server logout failed', logoutError);
    }

    var returnCandidate = (e && e.parameter && e.parameter.returnUrl) ? e.parameter.returnUrl : '';
    var loginUrl = buildLoginPageUrl({ returnUrl: returnCandidate });

    return HtmlService
      .createHtmlOutput('<script>window.top.location.href = ' + JSON.stringify(loginUrl) + ';</script>')
      .setTitle('Logging out…')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('handleLogoutRequest: unexpected failure', error);
    writeError('handleLogoutRequest', error);
    return renderLoginPage({ parameter: { page: 'login' } });
  }
}

function canonicalizePageKey(k) {
  const original = String(k || '').trim();
  const key = original.toLowerCase()
    .replace(/%2f/g, '/')
    .replace(/%5c/g, '\\')
    .replace(/%3a/g, ':')
    .replace(/%7c/g, '|')
    .replace(/%40/g, '@');
  if (!key) return key;

  if (/^(userprofile|user-profile|profile)(?:[:\/@|\\-]+.+)?$/.test(key)) {
    return 'userprofile';
  }

  if (/^(agent-experience|workspace.agent)(?:[:\/@|\\-]+.+)?$/.test(key)) {
    return 'agent-experience';
  }

  // Map legacy slugs/aliases → canonical keys used by the Access Engine
  switch (key) {
    // Landing pages
    case 'landing':
    case 'landing-page':
      return 'landing';
    case 'landing-about':
    case 'landingabout':
    case 'about':
    case 'about-luminahq':
      return 'landing-about';
    case 'landing-capabilities':
    case 'landingcapabilities':
    case 'capabilities':
    case 'explore-capabilities':
      return 'landing-capabilities';

    // Legal & public resources
    case 'terms-of-service':
    case 'terms-and-conditions':
    case 'termsofservice':
    case 'terms':
      return 'terms-of-service';
    case 'privacy-policy':
    case 'privacypolicy':
    case 'privacy':
    case 'privacy-notice':
      return 'privacy-policy';
    case 'lumina-user-guide':
    case 'lumina-hq-user-guide':
    case 'luminauserguide':
    case 'user-guide':
      return 'lumina-user-guide';

    // Admin Pages
    case 'manageuser':
    case 'users':
      return 'admin.users';
    case 'manageroles':
    case 'roles':
      return 'admin.roles';
    case 'managecampaign':
    case 'campaigns':
      return 'admin.campaigns';

    // Experience hubs
    case 'agent-experience':
    case 'workspace.agent':
      return 'agent-experience';
    case 'userprofile':
    case 'user-profile':
    case 'profile':
      return 'userprofile';
    case 'manager-executive-experience':
      return 'workspace.executive';
    case 'goalsetting':
      return 'performance.goals';

    // Task Management (Default Pages)
    case 'tasklist':
    case 'task-list':
      return 'tasks.list';
    case 'taskboard':
    case 'task-board':
      return 'tasks.board';
    case 'taskform':
    case 'task-form':
    case 'newtask':
    case 'edittask':
      return 'tasks.form';
    case 'taskview':
    case 'task-view':
      return 'tasks.view';

    // Communication & Collaboration (Default Pages)
    case 'chat':
      return 'global.chat';
    case 'search':
      return 'global.search';
    case 'bookmarks':
      return 'global.bookmarks';
    case 'collaboration-reporting':
    case 'collaborationreporting':
      return 'global.collaboration-reporting';

    // Schedule Management (Default Page)
    case 'schedule':
    case 'schedulemanagement':
      return 'global.schedule';
    case 'agent-schedule':
      return 'schedule.agent';

    // Other Global Pages
    case 'notifications':
      return 'global.notifications';
    case 'settings':
      return 'global.settings';

    // QA System Pages
    case 'unifiedqadashboard':
    case 'qa-dashboard':
    case 'ibtrqualityreports':
      return 'qa.dashboard';
    case 'qualityform':
    case 'independencequality':
    case 'creditsuiteqa':
    case 'groundingqaform':
      return 'qa.form';
    case 'qualitycollabform':
    case 'qacollabform':
      return 'qa.collaboration.form';
    case 'qualityview':
      return 'qa.view';
    case 'qualitylist':
      return 'qa.list';
    case 'qacollablist':
      return 'qa.collaboration.list';
    case 'qacollabview':
      return 'qa.collaboration.view';

    // Report Pages
    case 'callreports':
      return 'reports.calls';
    case 'attendancereports':
      return 'reports.attendance';

    // Coaching Pages
    case 'coachingdashboard':
      return 'coaching.dashboard';
    case 'coachinglist':
    case 'coachings':
      return 'coaching.list';
    case 'coachingview':
      return 'coaching.view';
    case 'coachingsheet':
    case 'coaching':
      return 'coaching.form';

    // Calendar and Schedule
    case 'calendar':
    case 'attendancecalendar':
      return 'calendar.attendance';
    case 'slotmanagement':
      return 'schedule.slots';

    case 'importcsv':
    case 'import-csv':
      return 'import';
    case 'importattendance':
    case 'import-attendance':
      return 'importattendance';

    // Dashboard
    case 'dashboard':
      return 'dashboard';

    // Proxy and Special Pages
    case 'proxy':
      return 'global.proxy';

    default:
      return key; // unknowns fall through; let the engine decide
  }
}

function __normalizeProfileIdentifierValue(value) {
  if (value === null || typeof value === 'undefined') {
    return '';
  }

  const text = String(value).trim();
  if (!text) {
    return '';
  }

  try {
    const decoded = decodeURIComponent(text);
    return decoded && decoded.trim() ? decoded.trim() : text;
  } catch (_) {
    return text;
  }
}

function __extractProfileIdentifierFromPageKey(rawPage) {
  if (rawPage === null || typeof rawPage === 'undefined') {
    return '';
  }

  const text = String(rawPage).trim();
  if (!text) {
    return '';
  }

  const sanitized = text.replace(/\+/g, ' ');
  const lowered = sanitized.toLowerCase();
  const prefixes = ['userprofile', 'user-profile', 'profile', 'agent-experience', 'workspace.agent'];

  for (let idx = 0; idx < prefixes.length; idx++) {
    const prefix = prefixes[idx];
    if (!lowered.startsWith(prefix)) {
      continue;
    }

    const remainder = sanitized.slice(prefix.length);
    if (!remainder) {
      continue;
    }

    const cleaned = remainder
      .replace(/^[\s:\/@|\\-]+/, '')
      .split(/[?#]/)[0];

    const normalized = __normalizeProfileIdentifierValue(cleaned);
    if (normalized) {
      return normalized;
    }
  }

  return '';
}

function resolveProfileIdentifierFromRequest(e, rawPage) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const candidateKeys = [
      'profileId', 'profileID', 'profileid', 'ProfileID', 'ProfileId',
      'profile', 'Profile',
      'profileSlug', 'ProfileSlug', 'profileslug',
      'slug', 'Slug',
      'handle', 'Handle',
      'userId', 'userID', 'userid', 'UserID', 'UserId',
      'ID', 'Id', 'id'
    ];

    for (let idx = 0; idx < candidateKeys.length; idx++) {
      const key = candidateKeys[idx];
      if (Object.prototype.hasOwnProperty.call(params, key)) {
        const normalized = __normalizeProfileIdentifierValue(params[key]);
        if (normalized) {
          return normalized;
        }
      }
    }

    const pathCandidates = [
      rawPage,
      params && params.page,
      params && params.path,
      params && params.route,
      params && params.profilePath,
      params && params.profileSlug,
      params && params.slug,
      params && params.handle
    ];

    for (let i = 0; i < pathCandidates.length; i++) {
      const extracted = __extractProfileIdentifierFromPageKey(pathCandidates[i]);
      if (extracted) {
        return extracted;
      }
    }
  } catch (err) {
    console.warn('resolveProfileIdentifierFromRequest: failed to resolve identifier', err);
  }

    return '';
}

// ───────────────────────────────────────────────────────────────────────────────
// ENHANCED doGet WITH SIMPLIFIED ROUTING
// ───────────────────────────────────────────────────────────────────────────────

function doGet(e) {
  try {
    const baseUrl = getBaseUrl();

    function redirectToLanding(user) {
      const userCampaignId = user.CampaignID || '';
      let landingSlug = '';

      try {
        if (typeof AuthenticationService !== 'undefined' && AuthenticationService) {
          if (typeof AuthenticationService.resolveLandingDestination === 'function') {
            const landingInfo = AuthenticationService.resolveLandingDestination(user, {
              user: user,
              userPayload: user,
              rawUser: user
            });
            if (landingInfo && landingInfo.slug) {
              landingSlug = landingInfo.slug;
            }
          } else if (typeof AuthenticationService.getLandingSlug === 'function') {
            landingSlug = AuthenticationService.getLandingSlug(user, { user: user, userPayload: user });
          }
        }
      } catch (landingError) {
        console.warn('doGet: failed to compute landing slug', landingError);
      }

      const redirectPage = landingSlug || 'dashboard';
      const redirectUrl = getAuthenticatedUrl(redirectPage, userCampaignId);
      return HtmlService
        .createHtmlOutput(`<script>window.location.href = "${redirectUrl}";</script>`)
        .setTitle('Redirecting...');
    }

    // Initialize system
    // initializeSystem();

    // Handle special actions
    if (e.parameter.page === 'proxy') {
      console.log('doGet: Handling proxy request');
      return serveEnhancedProxy(e);
    }

    if (e.parameter.action === 'logout') {
      console.log('doGet: Handling logout action');
      return handleLogoutRequest(e);
    }

    if (e.parameter.action === 'confirmEmail' && e.parameter.token) {
      const confirmationToken = e.parameter.token;
      const success = confirmEmail(confirmationToken);
      if (success) {
        const redirectUrl = `${baseUrl}?page=setpassword&token=${encodeURIComponent(confirmationToken)}`;
        return HtmlService
          .createHtmlOutput(`<script>window.location.href = "${redirectUrl}";</script>`)
          .setTitle('Redirecting...');
      } else {
        const tpl = HtmlService.createTemplateFromFile('EmailConfirmed');
        tpl.baseUrl = baseUrl;
        tpl.success = false;
        tpl.token = confirmationToken;
        return tpl.evaluate()
          .setTitle('Email Confirmation')
          .addMetaTag('viewport', 'width=device-width,initial-scale=1');
      }
    }

    // Handle public pages
    const rawPageParam = (typeof e.parameter.page === 'string') ? e.parameter.page : '';
    const page = rawPageParam.toLowerCase();

    if (!page) {
      return handlePublicPage('landing', e, baseUrl);
    }

    if (page === 'login') {
      try {
        const existingSession = authenticateUser(e);
        if (existingSession && existingSession.ID) {
          return redirectToLanding(existingSession);
        }
      } catch (sessionError) {
        console.warn('doGet: session probe failed for login/default route', sessionError);
      }

      return renderLoginPage(e);
    }

    // Handle other public pages
    const publicPages = [
      'landing',
      'landing-about',
      'about',
      'landing-capabilities',
      'capabilities',
      'setpassword',
      'resetpassword',
      'resend-verification',
      'resendverification',
      'forgotpassword',
      'forgot-password',
      'emailconfirmed',
      'email-confirmed',
      'terms-of-service',
      'termsofservice',
      'terms',
      'privacy-policy',
      'privacypolicy',
      'privacy',
      'lumina-user-guide',
      'lumina-hq-user-guide',
      'user-guide'
    ];

    if (publicPages.includes(page)) {
      return handlePublicPage(page, e, baseUrl);
    }

    // Protected pages - require authentication
    const auth = requireAuth(e);
    if (auth.getContent) {
      return auth; // Login or access denied page
    }

    const user = auth;

    // Handle password reset requirement
    if (_truthy(user.ResetRequired)) {
      const tpl = HtmlService.createTemplateFromFile('ChangePassword');
      tpl.baseUrl = baseUrl;
      tpl.scriptUrl = SCRIPT_URL;
      return tpl.evaluate()
        .setTitle('Change Password')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    const campaignId = e.parameter.campaign || user.CampaignID || '';

    // Handle CSV exports
    if (page === "callreports" && e.parameter.action === "exportCallCsv") {
      const gran = e.parameter.granularity || "Week";
      const period = e.parameter.period || weekStringFromDate(new Date());
      const agent = e.parameter.agent || "";
      const csv = exportCallAnalyticsCsv(gran, period, agent);
      return ContentService.createTextOutput(csv)
        .setMimeType(ContentService.MimeType.CSV)
        .downloadAsFile(`callAnalytics_${period}.csv`);
    }

    if (page === 'attendancereports' && e.parameter.action === 'exportCsv') {
      const gran = e.parameter.granularity || 'Week';
      const period = e.parameter.period || weekStringFromDate(new Date());
      const agent = e.parameter.agent || '';
      const csv = exportAttendanceCsv(gran, period, agent);
      return ContentService.createTextOutput(csv)
        .setMimeType(ContentService.MimeType.CSV)
        .downloadAsFile(`attendance_${period}.csv`);
    }

    // Route to appropriate page
    return routeToPage(page, e, baseUrl, user, campaignId);

  } catch (error) {
    console.error('Error in doGet:', error);
    writeError('doGet', error);
    return createErrorPage('System Error', `An error occurred: ${error.message}`);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// ENHANCED ROUTING WITH CAMPAIGN SUPPORT (Updated for Clean URLs)
// ───────────────────────────────────────────────────────────────────────────────

function routeToPage(page, e, baseUrl, user, campaignIdFromCaller) {
  try {
    const raw = String(page || '').trim();
    const canonicalPage = canonicalizePageKey(raw);
    const resolvedProfileId = resolveProfileIdentifierFromRequest(e, raw);
    const hasProfileParameter = resolvedProfileId !== '';

    if (hasProfileParameter && e) {
      if (!e.parameter) {
        e.parameter = {};
      }

      const targetParam = e.parameter;
      const existingProfileId = targetParam.profileId || targetParam.profileID || targetParam.ProfileID || targetParam.ProfileId;
      if (!existingProfileId) {
        targetParam.profileId = resolvedProfileId;
      }
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // DEFAULT/GLOBAL PAGES (Always available, campaign-independent)
    // ═══════════════════════════════════════════════════════════════════════════

    if (canonicalPage === 'userprofile' || (canonicalPage === 'agent-experience' && hasProfileParameter)) {
      return serveGlobalPage('UserProfile', e, baseUrl, user);
    }

    if (canonicalPage === 'agent-experience') {
      return serveGlobalPage('AgentExperience', e, baseUrl, user);
    }

    if (page === 'goalsetting') {
      return serveGlobalPage('GoalSetting', e, baseUrl, user);
    }

    // Task Management (Default Pages)
    if (page === "tasklist" || page === "task-list") {
      return serveGlobalPage('TaskList', e, baseUrl, user);
    }

    if (page === "taskboard" || page === "task-board") {
      return serveGlobalPage('TaskBoard', e, baseUrl, user);
    }

    if (page === "taskform" || page === "task-form" || page === "newtask" || page === "edittask") {
      return serveGlobalPage('TaskForm', e, baseUrl, user);
    }

    if (page === "taskview" || page === "task-view") {
      return serveGlobalPage('TaskView', e, baseUrl, user);
    }

    // Communication & Collaboration (Default Pages)
    if (page === 'chat') {
      return serveGlobalPage('Chat', e, baseUrl, user);
    }

    if (page === 'search') {
      return serveGlobalPage('Search', e, baseUrl, user);
    }

    if (page === 'bookmarks') {
      return serveGlobalPage('BookmarkManager', e, baseUrl, user);
    }

    // Administration (Default Pages)
    if (page === 'manager-executive-experience') {
      return serveAdminPage('ManagerExecutiveExperience', e, baseUrl, user, {
        allowManagers: true,
        allowSupervisors: true,
        accessDeniedMessage: 'Manager or executive access required.'
      });
    }

    if (page === 'users' || page === 'manageuser') {
      return serveAdminPage('Users', e, baseUrl, user);
    }

    if (page === 'roles' || page === 'manageroles') {
      return serveAdminPage('RoleManagement', e, baseUrl, user);
    }

    if (page === 'campaigns' || page === 'managecampaign') {
      return serveAdminPage('CampaignManagement', e, baseUrl, user);
    }

    // Schedule Management (Default Page)
    if (page === 'schedule' || page === 'schedulemanagement') {
      return serveGlobalPage('ScheduleManagement', e, baseUrl, user);
    }

    // Other Default Global Pages
    if (page === 'notifications') {
      return serveGlobalPage('Notifications', e, baseUrl, user);
    }

    if (page === 'settings') {
      return serveGlobalPage('Settings', e, baseUrl, user);
    }

    if (page === 'proxy') {
      return serveProxy(e);
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // CAMPAIGN-SPECIFIC ROUTING
    // ═══════════════════════════════════════════════════════════════════════════

    // Campaign-specific routing pattern: Case.PageKind
    const PATTERN = /^([^._:\-][^._:\-]*(?:[ _:\-][^._:\-]+)*)[._:\-](QAForm|QADashboard|QAList|QAView|QACollabList|AttendanceReports|CallReports|ChatReport|Dashboard|CoachingForm|TaskForm|TaskList|TaskBoard)$/i;
    const match = PATTERN.exec(raw);

    if (match) {
      const caseToken = match[1].trim();
      const pageKind = match[2];
      const def = __case_resolve__(caseToken);

      if (!def) {
        return createErrorPage('Unknown Campaign', 'No case mapping found for: ' + caseToken);
      }

      // Determine campaign ID
      const explicitCid = String(e?.parameter?.campaign || '').trim();
      const sheetCid = __case_findCampaignId__(def);
      const cid = explicitCid || campaignIdFromCaller || sheetCid || (user && user.CampaignID) || '';

      const candidates = __case_templateCandidates__(def, pageKind);
      return __case_serve__(candidates, e, baseUrl, user, cid);
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // LEGACY/BACKWARD COMPATIBILITY ROUTES
    // ═══════════════════════════════════════════════════════════════════════════

    switch (page) {
      case "dashboard":
        return serveCampaignPage('Dashboard', e, baseUrl, user, campaignIdFromCaller);

      case "qualityform":
        return routeQAForm(e, baseUrl, user, campaignIdFromCaller);

      case "ibtrqualityreports":
        return routeQADashboard(e, baseUrl, user, campaignIdFromCaller);

      case 'qualityview':
        return routeQAView(e, baseUrl, user, campaignIdFromCaller);

      case 'qualitylist':
        return routeQAList(e, baseUrl, user, campaignIdFromCaller);

      case 'qacollablist':
        return serveCampaignPage('QACollabList', e, baseUrl, user, campaignIdFromCaller);

      case 'qacollabview':
        return serveCampaignPage('QualityCollabView', e, baseUrl, user, campaignIdFromCaller);

      case 'qualitycollabform':
      case 'qacollabform':
        return serveCampaignPage('QualityCollabForm', e, baseUrl, user, campaignIdFromCaller);

      case "independencequality":
        return serveCampaignPage('IndependenceQAForm', e, baseUrl, user, campaignIdFromCaller);

      case "independenceqadashboard":
        return serveCampaignPage('UnifiedQADashboard', e, baseUrl, user, campaignIdFromCaller);

      case "creditsuiteqa":
        return serveCampaignPage('CreditSuiteQAForm', e, baseUrl, user, campaignIdFromCaller);

      case "groundingqaform":
        return serveCampaignPage('GroundingQAForm', e, baseUrl, user, campaignIdFromCaller);

      case "qa-dashboard":
        return serveCampaignPage('CreditSuiteQADashboard', e, baseUrl, user, campaignIdFromCaller);

      case "unifiedqadashboard":
        return serveCampaignPage('UnifiedQADashboard', e, baseUrl, user, campaignIdFromCaller);

      case "callreports":
        return serveCampaignPage('CallReports', e, baseUrl, user, campaignIdFromCaller);

      case "attendancereports":
        return serveCampaignPage('AttendanceReports', e, baseUrl, user, campaignIdFromCaller);

      case "coachingdashboard":
        return serveCampaignPage('CoachingDashboard', e, baseUrl, user, campaignIdFromCaller);

      case "coachingview":
        return serveCampaignPage('CoachingView', e, baseUrl, user, campaignIdFromCaller);

      case "coachinglist":
      case "coachings":
        return serveCampaignPage('CoachingList', e, baseUrl, user, campaignIdFromCaller);

      case "coachingsheet":
      case "coaching":
        return serveCampaignPage('CoachingForm', e, baseUrl, user, campaignIdFromCaller);

      case "eodreport":
        return serveGlobalPage('EODReport', e, baseUrl, user);

      case "collaboration-reporting":
      case "collaborationreporting":
        return serveGlobalPage('CollaborationReporting', e, baseUrl, user);

      case "calendar":
      case "attendancecalendar":
        return serveCampaignPage('Calendar', e, baseUrl, user, campaignIdFromCaller);

      case "escalations":
        return serveCampaignPage('Escalations', e, baseUrl, user, campaignIdFromCaller);

      case "incentives":
        return serveCampaignPage('Incentives', e, baseUrl, user, campaignIdFromCaller);

      case 'slotmanagement':
        return serveShiftSlotManagement(e, baseUrl, user, campaignIdFromCaller);

      case 'ackform':
        return serveAckForm(e, baseUrl, user);

      case "agent-schedule":
        return serveAgentSchedulePage(e, baseUrl, e.parameter.token);

      case 'importcsv':
      case 'import-csv':
      case 'import':
        // Mirror Import Attendance authentication so managers and supervisors can import call reports
        if (isSystemAdmin(user) || hasManagerRole(user) || hasSupervisorRole(user)) {
          return serveAdminPage('ImportCsv', e, baseUrl, user, {
            allowManagers: true,
            allowSupervisors: true,
            accessDeniedMessage: 'You need manager or supervisor privileges to import call reports.'
          });
        } else {
          return renderAccessDenied('You need manager or supervisor privileges to import call reports.');
        }

      case 'importattendance':
      case 'import-attendance':
        // Allow managers and supervisors for attendance imports
        if (isSystemAdmin(user) || hasManagerRole(user) || hasSupervisorRole(user)) {
          return serveAdminPage('ImportAttendance', e, baseUrl, user, {
            allowManagers: true,
            allowSupervisors: true,
            accessDeniedMessage: 'You need manager or supervisor privileges to import attendance data.'
          });
        } else {
          return renderAccessDenied('You need manager or supervisor privileges to import attendance data.');
        }

      default:
        // Unknown page - redirect to dashboard
        const defaultCampaignId = user.CampaignID || '';
        const redirectUrl = getAuthenticatedUrl('dashboard', defaultCampaignId);

        return HtmlService
          .createHtmlOutput(`<script>window.location.href = "${redirectUrl}";</script>`)
          .setTitle('Redirecting to Dashboard...');
    }
  } catch (error) {
    console.error(`Error routing to page ${page}:`, error);
    writeError('routeToPage', error);
    return createErrorPage('Page Error', `Error loading page: ${error.message}`);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// PAGE SERVING FUNCTIONS (Updated for Clean URLs)
// ───────────────────────────────────────────────────────────────────────────────

function serveCampaignPage(templateName, e, baseUrl, user, campaignId) {
  try {
    const tpl = HtmlService.createTemplateFromFile(templateName);
    tpl.baseUrl = baseUrl;
    tpl.scriptUrl = SCRIPT_URL;
    tpl.currentPage = templateName.replace(/([A-Z])/g, ' $1').trim();
    tpl.user = user;
    tpl.campaignId = campaignId;

    __injectCurrentUser_(tpl, user);

    if (campaignId) {
      tpl.campaignName = (tpl.user && tpl.user.campaignName) || '';
      tpl.campaignNavigation = (tpl.user && tpl.user.campaignNavigation) || { categories: [], uncategorizedPages: [] };
    }

    addTemplateSpecificData(tpl, templateName, e, tpl.user, campaignId);

    return tpl.evaluate()
      .setTitle(tpl.currentPage + ' - VLBPO LuminaHQ')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    writeError(`serveCampaignPage(${templateName})`, error);
    return createErrorPage('Failed to load page', error.message);
  }
}

function serveGlobalPage(templateName, e, baseUrl, user) {
  try {
    const tpl = HtmlService.createTemplateFromFile(templateName);
    tpl.baseUrl = baseUrl;
    tpl.scriptUrl = SCRIPT_URL;
    tpl.currentPage = templateName.replace(/([A-Z])/g, ' $1').trim();
    tpl.user = user;

    __injectCurrentUser_(tpl, user);

    addTemplateSpecificData(tpl, templateName, e, tpl.user, null);

    return tpl.evaluate()
      .setTitle(tpl.currentPage + ' - VLBPO LuminaHQ')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    writeError(`serveGlobalPage(${templateName})`, error);
    return createErrorPage('Failed to load page', error.message);
  }
}

function __collectUserRoleNames(user) {
  const roles = [];

  const append = (value) => {
    if (value === null || typeof value === 'undefined') {
      return;
    }

    if (Array.isArray(value)) {
      value.forEach(append);
      return;
    }

    const raw = String(value);
    if (!raw) {
      return;
    }

    raw
      .split(/[,;|]+/)
      .map(part => part.trim().toLowerCase())
      .filter(Boolean)
      .forEach(part => roles.push(part));
  };

  try {
    if (user) {
      append(user.roleNames);
      append(user.RoleNames);
      append(user.roles);
      append(user.Roles);
      append(user.Role);
      append(user.role);
      append(user.PrimaryRole);
      append(user.primaryRole);
      append(user.Position);
      append(user.position);
      append(user.JobTitle);
      append(user.jobTitle);
      append(user.Title);
      append(user.title);
    }
  } catch (roleError) {
    writeError && writeError('__collectUserRoleNames', roleError);
  }

  return Array.from(new Set(roles));
}

function __userHasRoleKeyword(user, keywords) {
  const list = Array.isArray(keywords) ? keywords : [keywords];
  const normalizedKeywords = list
    .map(keyword => String(keyword || '').trim().toLowerCase())
    .filter(Boolean);

  if (!normalizedKeywords.length) {
    return false;
  }

  try {
    const roles = __collectUserRoleNames(user);
    return roles.some(role => normalizedKeywords.some(keyword => role.includes(keyword)));
  } catch (err) {
    writeError && writeError('__userHasRoleKeyword', err);
    return false;
  }
}

function hasManagerRole(user) {
  if (isSystemAdmin(user)) {
    return true;
  }
  return __userHasRoleKeyword(user, ['manager']);
}

function hasSupervisorRole(user) {
  if (isSystemAdmin(user)) {
    return true;
  }
  return __userHasRoleKeyword(user, ['supervisor', 'team lead', 'teamlead']);
}

function serveAdminPage(templateName, e, baseUrl, user, options) {
  const opts = Object.assign({
    allowManagers: false,
    allowSupervisors: false,
    allowedRoles: [],
    accessDeniedMessage: 'This page requires System Admin privileges.'
  }, options || {});

  const hasAdditionalAccess =
    (opts.allowManagers && hasManagerRole(user)) ||
    (opts.allowSupervisors && hasSupervisorRole(user)) ||
    (Array.isArray(opts.allowedRoles) && opts.allowedRoles.length
      ? __userHasRoleKeyword(user, opts.allowedRoles)
      : false);

  if (!isSystemAdmin(user) && !hasAdditionalAccess) {
    return renderAccessDenied(opts.accessDeniedMessage);
  }

  try {
    const tpl = HtmlService.createTemplateFromFile(templateName);
    tpl.baseUrl = baseUrl;
    tpl.scriptUrl = SCRIPT_URL;
    tpl.currentPage = templateName.replace(/([A-Z])/g, ' $1').trim();
    tpl.user = user;

    // Inject current user data consistently with other serve functions
    __injectCurrentUser_(tpl, user);

    // Add admin page specific data
    addTemplateSpecificData(tpl, templateName, e, user, null);

    return tpl.evaluate()
      .setTitle(tpl.currentPage + ' - VLBPO LuminaHQ')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    console.error(`Error serving admin page ${templateName}:`, error);
    writeError(`serveAdminPage(${templateName})`, error);
    return createErrorPage('Failed to load page', error.message);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// PUBLIC PAGE HANDLERS (Updated for Clean URLs)
// ───────────────────────────────────────────────────────────────────────────────

function handlePublicPage(page, e, baseUrl) {
  const scriptUrl = SCRIPT_URL;

  switch (page) {
    case 'landing':
      const landingTpl = HtmlService.createTemplateFromFile('Landing');
      landingTpl.baseUrl = baseUrl;
      landingTpl.scriptUrl = scriptUrl;

      return landingTpl.evaluate()
        .setTitle('LuminaHQ – Intelligent Workforce Command Center')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    case 'landing-about':
    case 'about':
    case 'about-luminahq':
      const aboutTpl = HtmlService.createTemplateFromFile('LandingAbout');
      aboutTpl.baseUrl = baseUrl;
      aboutTpl.scriptUrl = scriptUrl;

      return aboutTpl.evaluate()
        .setTitle('About LuminaHQ')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    case 'landing-capabilities':
    case 'capabilities':
    case 'explore-capabilities':
      const capabilitiesTpl = HtmlService.createTemplateFromFile('LandingCapabilities');
      capabilitiesTpl.baseUrl = baseUrl;
      capabilitiesTpl.scriptUrl = scriptUrl;

      return capabilitiesTpl.evaluate()
        .setTitle('Explore LuminaHQ Capabilities')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    case 'terms-of-service':
    case 'termsofservice':
    case 'terms':
    case 'terms-and-conditions':
      const termsTpl = HtmlService.createTemplateFromFile('TermsOfService');
      termsTpl.baseUrl = baseUrl;
      termsTpl.scriptUrl = scriptUrl;
      termsTpl.user = {};
      termsTpl.currentPage = 'terms-of-service';

      return termsTpl.evaluate()
        .setTitle('Terms of Service - VLBPO LuminaHQ')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    case 'privacy-policy':
    case 'privacypolicy':
    case 'privacy':
    case 'privacy-notice':
      const privacyTpl = HtmlService.createTemplateFromFile('PrivacyPolicy');
      privacyTpl.baseUrl = baseUrl;
      privacyTpl.scriptUrl = scriptUrl;
      privacyTpl.user = {};
      privacyTpl.currentPage = 'privacy-policy';

      return privacyTpl.evaluate()
        .setTitle('Privacy Policy - VLBPO LuminaHQ')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    case 'lumina-user-guide':
    case 'lumina-hq-user-guide':
    case 'luminauserguide':
    case 'user-guide':
      const guideTpl = HtmlService.createTemplateFromFile('LuminaHQUserGuide');
      guideTpl.baseUrl = baseUrl;
      guideTpl.scriptUrl = scriptUrl;
      guideTpl.user = {};
      guideTpl.currentPage = 'lumina-user-guide';

      return guideTpl.evaluate()
        .setTitle('LuminaHQ User Guide')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    case 'setpassword':
    case 'resetpassword':
      const resetToken = e.parameter.token || '';
      const tpl = HtmlService.createTemplateFromFile('ChangePassword');
      tpl.baseUrl = baseUrl;
      tpl.scriptUrl = scriptUrl;
      tpl.token = resetToken;
      tpl.isReset = page === 'resetpassword';

      return tpl.evaluate()
        .setTitle((page === 'resetpassword' ? 'Reset' : 'Set') + ' Your Password - VLBPO LuminaHQ')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    case 'resend-verification':
    case 'resendverification':
      const verifyTpl = HtmlService.createTemplateFromFile('ResendVerification');
      verifyTpl.baseUrl = baseUrl;
      verifyTpl.scriptUrl = scriptUrl;

      return verifyTpl.evaluate()
        .setTitle('Resend Email Verification - VLBPO LuminaHQ')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    case 'forgotpassword':
    case 'forgot-password':
      const forgotTpl = HtmlService.createTemplateFromFile('ForgotPassword');
      forgotTpl.baseUrl = baseUrl;
      forgotTpl.scriptUrl = scriptUrl;

      return forgotTpl.evaluate()
        .setTitle('Forgot Password - VLBPO LuminaHQ')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    case 'emailconfirmed':
    case 'email-confirmed':
      const confirmTpl = HtmlService.createTemplateFromFile('EmailConfirmed');
      confirmTpl.baseUrl = baseUrl;
      confirmTpl.success = e.parameter.success === 'true';
      confirmTpl.token = e.parameter.token || '';

      return confirmTpl.evaluate()
        .setTitle('Email Confirmation - VLBPO LuminaHQ')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    default:
      return createErrorPage('Page Not Found', `The page "${page}" was not found.`);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// REST OF THE CODE (INCLUDES ALL EXISTING FUNCTIONS)
// These functions remain largely unchanged but with clean URL updates where needed
// ───────────────────────────────────────────────────────────────────────────────

// [Include all existing functions from the original Code.gs with URL updates]
// Campaign resolution helpers, template handling, user management, etc.
// The implementation continues with all existing functionality...

// Campaign resolution helpers
function __case_allCampaignsRaw__() {
  try {
    return (typeof readSheet === 'function') ? (readSheet('CAMPAIGNS') || []) : [];
  } catch (_) {
    return [];
  }
}

function __case_findCampaignId__(def) {
  try {
    const rows = __case_allCampaignsRaw__();
    const aliasSet = new Set([def.case].concat(def.aliases || []));
    const aliasSlugs = new Set(Array.from(aliasSet).map(__case_slug__));

    const hit = rows.find(r => aliasSet.has(String(r.Name || r.name || '').trim())) ||
      rows.find(r => aliasSlugs.has(__case_slug__(String(r.Name || r.name || '').trim())));
    if (hit) return String(hit.ID || hit.Id || hit.id || '').trim() || def.idHint || '';
  } catch (_) { }
  return def.idHint || '';
}

function __case_registry__() {
  const byCase = new Map();
  const byAlias = new Map();
  CASE_DEFS.forEach(def => {
    byCase.set(def.case, def);
    (def.aliases || []).forEach(a => byAlias.set(__case_norm__(a), def));
    byAlias.set(__case_norm__(def.case), def);
  });
  return { byCase, byAlias };
}

function __case_resolve__(token) {
  const reg = __case_registry__();
  const key = __case_norm__(token);
  return reg.byAlias.get(key) || null;
}

function __case_templateCandidates__(def, pageKind) {
  const chosen = (def.pages || {})[pageKind];
  const candidates = [];
  if (chosen) candidates.push(chosen);

  // Special handling for Independence dashboard
  if (pageKind === 'QADashboard' && def.case === 'IndependenceInsuranceAgency') {
    candidates.push('UnifiedQADashboard');
  }

  // Add generic fallbacks
  (GENERIC_FALLBACKS[pageKind] || []).forEach(t => candidates.push(t));

  // Remove duplicates
  return candidates.filter((v, i, a) => a.indexOf(v) === i);
}

function __case_serve__(candidates, e, baseUrl, user, campaignId) {
  for (var i = 0; i < candidates.length; i++) {
    try {
      HtmlService.createTemplateFromFile(candidates[i]); // Test existence
      return serveCampaignPage(candidates[i], e, baseUrl, user, campaignId);
    } catch (_) {
      // Template doesn't exist, try next
    }
  }
  return createErrorPage('Missing Page', `Template not found for this campaign/page. Tried: ${candidates.join(', ')}`);
}

// QA System routing
function routeQAForm(e, baseUrl, user, campaignId) {
  const campaign = getCampaignById(campaignId);
  const campaignName = campaign?.Name?.toLowerCase() || '';

  if (campaignName.includes('independence') || campaignName.includes('insurance')) {
    return serveCampaignPage('IndependenceQAForm', e, baseUrl, user, campaignId);
  } else if (campaignName.includes('credit') || campaignName.includes('suite')) {
    return serveCampaignPage('CreditSuiteQAForm', e, baseUrl, user, campaignId);
  } else {
    return serveCampaignPage('QualityForm', e, baseUrl, user, campaignId);
  }
}

function routeQADashboard(e, baseUrl, user, campaignId) {
  const campaign = getCampaignById(campaignId);
  const campaignName = campaign?.Name?.toLowerCase() || '';

  if (campaignName.includes('independence') || campaignName.includes('insurance')) {
    return serveCampaignPage('UnifiedQADashboard', e, baseUrl, user, campaignId);
  } else if (campaignName.includes('credit') || campaignName.includes('suite')) {
    return serveCampaignPage('CreditSuiteQADashboard', e, baseUrl, user, campaignId);
  } else {
    return serveCampaignPage('QADashboard', e, baseUrl, user, campaignId);
  }
}

function routeQAView(e, baseUrl, user, campaignId) {
  const campaign = getCampaignById(campaignId);
  const campaignName = campaign?.Name?.toLowerCase() || '';

  if (campaignName.includes('independence') || campaignName.includes('insurance')) {
    return serveCampaignPage('IndependenceQAView', e, baseUrl, user, campaignId);
  } else if (campaignName.includes('credit') || campaignName.includes('suite')) {
    return serveCampaignPage('CreditSuiteQAView', e, baseUrl, user, campaignId);
  } else {
    return serveCampaignPage('QualityView', e, baseUrl, user, campaignId);
  }
}

function routeQAList(e, baseUrl, user, campaignId) {
  const campaign = getCampaignById(campaignId);
  const campaignName = campaign?.Name?.toLowerCase() || '';

  if (campaignName.includes('independence') || campaignName.includes('insurance')) {
    return serveCampaignPage('IndependenceQAList', e, baseUrl, user, campaignId);
  } else if (campaignName.includes('credit') || campaignName.includes('suite')) {
    return serveCampaignPage('CreditSuiteQAList', e, baseUrl, user, campaignId);
  } else {
    return serveCampaignPage('QAList', e, baseUrl, user, campaignId);
  }
}

// [Continue with all remaining functions from the original Code.gs...]
// This includes all the template data handling, user management, utility functions, etc.
// The functions remain the same but with updated URL generation

// Template data handling functions
function addTemplateSpecificData(tpl, templateName, e, user, campaignId) {
  try {
    const timeZone = Session.getScriptTimeZone();
    const granularity = (e.parameter.granularity || "Week").toString();
    const periodValue = (e.parameter.period || weekStringFromDate(new Date())).toString();
    const selectedAgent = (e.parameter.agent || "").toString();
    const selectedCampaign = (e.parameter.campaign || campaignId || "").toString();

    // Set common template properties
    tpl.granularity = granularity;
    tpl.period = periodValue;
    tpl.periodValue = periodValue;
    tpl.selectedPeriod = periodValue;
    tpl.selectedAgent = selectedAgent;
    tpl.selectedCampaign = selectedCampaign;
    tpl.timeZone = timeZone;

    handleTemplateSpecificData(tpl, templateName, e, user, campaignId);
  } catch (error) {
    console.error(`Error adding template data for ${templateName}:`, error);
    writeError(`addTemplateSpecificData(${templateName})`, error);
  }
}

function handleTemplateSpecificData(tpl, templateName, e, user, campaignId) {
  try {
    switch (templateName) {
      case 'Dashboard':
        handleDashboardData(tpl, e, user, campaignId);
        break;

      case 'AttendanceReports':
        handleAttendanceReportsData(tpl, e, user, campaignId);
        break;

      case 'CallReports':
        handleCallReportsData(tpl, e, user, campaignId);
        break;

      // Campaign-specific QA forms
      case 'IndependenceQAForm':
      case 'CreditSuiteQAForm':
      case 'HiyaCarQAForm':
      case 'IBTRQAForm':
      case 'JSCQAForm':
      case 'KidsInTheGameQAForm':
      case 'KofiGroupQAForm':
      case 'PAWLawFirmQAForm':
      case 'ProHousePhotosQAForm':
      case 'IACSQAForm':
      case 'ProozyQAForm':
      case 'TheGroundingQAForm':
      case 'COQAForm':
        handleQAFormData(tpl, e, user, templateName, campaignId);
        break;

      case 'QualityForm':
        handleQAFormData(tpl, e, user, templateName, campaignId);
        break;

      // Campaign-specific QA views
      case 'IndependenceQAView':
      case 'CreditSuiteQAView':
      case 'HiyaCarQAView':
      case 'IBTRQAView':
      case 'JSCQAView':
      case 'KidsInTheGameQAView':
      case 'KofiGroupQAView':
      case 'PAWLawFirmQAView':
      case 'ProHousePhotosQAView':
      case 'IACSQAView':
      case 'ProozyQAView':
      case 'TheGroundingQAView':
      case 'COQAView':
        handleQAViewData(tpl, e, user, templateName);
        break;

      case 'QualityView':
        handleQAViewData(tpl, e, user, templateName);
        break;

      // Campaign-specific QA dashboards
      case 'IndependenceQADashboard':
      case 'UnifiedQADashboard':
      case 'CreditSuiteQADashboard':
      case 'HiyaCarQADashboard':
      case 'IBTRQADashboard':
      case 'JSCQADashboard':
      case 'KidsInTheGameQADashboard':
      case 'KofiGroupQADashboard':
      case 'PAWLawFirmQADashboard':
      case 'ProHousePhotosQADashboard':
      case 'IACSQADashboard':
      case 'ProozyQADashboard':
      case 'TheGroundingQADashboard':
      case 'COQADashboard':
        handleQADashboardData(tpl, e, user, templateName, campaignId);
        break;

      case 'QADashboard':
        handleQADashboardData(tpl, e, user, templateName, campaignId);
        break;

      // Campaign-specific QA lists
      case 'IndependenceQAList':
      case 'CreditSuiteQAList':
      case 'HiyaCarQAList':
      case 'IBTRQAList':
      case 'JSCQAList':
      case 'KidsInTheGameQAList':
      case 'KofiGroupQAList':
      case 'PAWLawFirmQAList':
      case 'ProHousePhotosQAList':
      case 'IACSQAList':
      case 'ProozyQAList':
      case 'TheGroundingQAList':
      case 'COQAList':
      case 'IBTRQACollabList':
        handleQAListData(tpl, e, user, templateName);
        break;

      case 'QAList':
        handleQAListData(tpl, e, user, templateName);
        break;

      case 'Users':
        handleUsersData(tpl, e, user);
        break;

      case 'TaskForm':
      case 'TaskList':
      case 'TaskBoard':
      case 'TaskView':
        handleTaskData(tpl, e, user, templateName);
        break;

      case 'CoachingForm':
      case 'CoachingList':
      case 'CoachingView':
      case 'CoachingDashboard':
        handleCoachingData(tpl, e, user, templateName, campaignId);
        break;

      case 'EODReport':
        handleEODReportData(tpl, e, user);
        break;

      case 'Calendar':
        handleCalendarData(tpl, e, user, campaignId);
        break;

      case 'Search':
        handleSearchData(tpl, e);
        break;

      case 'UserProfile':
        handleUserProfileData(tpl, e, user);
        break;

      case 'Chat':
        handleChatData(tpl, e, user);
        break;

      case 'Notifications':
        handleNotificationsData(tpl, e, user);
        break;

      default:
        console.log(`No specific data handler for template: ${templateName}`);
        break;
    }
  } catch (error) {
    console.error(`Error handling data for ${templateName}:`, error);
    writeError(`handleTemplateSpecificData(${templateName})`, error);
  }
}

// Individual template data handlers (these remain largely the same)
function handleDashboardData(tpl, e, user, campaignId) {
  try {
    const granularity = e.parameter.granularity || "Week";
    const periodValue = e.parameter.period || weekStringFromDate(new Date());
    const selectedAgent = e.parameter.agent || "";

    if (typeof getDashboardOkrs === 'function') {
      const okrData = getDashboardOkrs(granularity, periodValue, selectedAgent);
      tpl.okrData = JSON.stringify(okrData).replace(/<\/script>/g, '<\\/script>');
    } else {
      tpl.okrData = JSON.stringify({});
    }

    // Use manager-filtered user list
    const requestingUserId = user && user.ID ? user.ID : null;
    tpl.userList = clientGetAssignedAgentNames(campaignId || '', requestingUserId);
  } catch (error) {
    console.error('Error handling dashboard data:', error);
    tpl.okrData = JSON.stringify({});
    tpl.userList = [];
  }
}

function handleAttendanceReportsData(tpl, e, user, campaignId) {
  try {
    const granularity = e.parameter.granularity || "Week";
    const periodValue = e.parameter.period || weekStringFromDate(new Date());
    const selectedAgent = e.parameter.agent || "";

    if (typeof clientGetEnhancedAttendanceAnalytics === 'function') {
      const enhancedAnalytics = clientGetEnhancedAttendanceAnalytics(granularity, periodValue, selectedAgent);

      const allRows = enhancedAnalytics.filteredRows || [];
      const rowPage = parseInt(e.parameter.rowPage, 10) || 1;
      const PAGE_SIZE = 50;
      const startRowIdx = (rowPage - 1) * PAGE_SIZE;
      enhancedAnalytics.filteredRows = allRows.slice(startRowIdx, startRowIdx + PAGE_SIZE);

      tpl.attendanceData = JSON.stringify(enhancedAnalytics).replace(/<\/script>/g, '<\\/script>');
      tpl.currentRowPage = rowPage;
      tpl.totalRows = allRows.length;
      tpl.PAGE_SIZE = PAGE_SIZE;
      tpl.executiveMetrics = JSON.stringify(enhancedAnalytics.executiveMetrics || {}).replace(/<\/script>/g, '<\\/script>');
    } else {
      const attendanceAnalytics = getAttendanceAnalyticsByPeriod(granularity, periodValue, selectedAgent);
      const allRows = attendanceAnalytics.filteredRows || [];
      const rowPage = parseInt(e.parameter.rowPage, 10) || 1;
      const PAGE_SIZE = 50;
      const startRowIdx = (rowPage - 1) * PAGE_SIZE;
      attendanceAnalytics.filteredRows = allRows.slice(startRowIdx, startRowIdx + PAGE_SIZE);

      tpl.attendanceData = JSON.stringify(attendanceAnalytics).replace(/<\/script>/g, '<\\/script>');
      tpl.currentRowPage = rowPage;
      tpl.totalRows = allRows.length;
      tpl.PAGE_SIZE = PAGE_SIZE;
      tpl.executiveMetrics = JSON.stringify({});
    }

    // Use manager-filtered user list
    const requestingUserId = user && user.ID ? user.ID : null;
    tpl.userList = clientGetAssignedAgentNames(campaignId || user.CampaignID || '', requestingUserId);

    const resolvedTimezone = (typeof GLOBAL_SCOPE.ATTENDANCE_TIMEZONE === 'string' && GLOBAL_SCOPE.ATTENDANCE_TIMEZONE)
      ? GLOBAL_SCOPE.ATTENDANCE_TIMEZONE
      : (typeof Session !== 'undefined' && Session.getScriptTimeZone ? Session.getScriptTimeZone() : 'America/Jamaica');
    const resolvedTimezoneLabel = (typeof GLOBAL_SCOPE.ATTENDANCE_TIMEZONE_LABEL === 'string' && GLOBAL_SCOPE.ATTENDANCE_TIMEZONE_LABEL)
      ? GLOBAL_SCOPE.ATTENDANCE_TIMEZONE_LABEL
      : 'Company Time';

    tpl.managerUserId = user && user.ID ? user.ID : '';
    tpl.attendanceTimezone = resolvedTimezone;
    tpl.attendanceTimezoneLabel = resolvedTimezoneLabel;
    tpl.attendanceDataJSON = tpl.attendanceData;
    tpl.userListJSON = JSON.stringify(tpl.userList || []).replace(/<\/script>/g, '<\\/script>');
    tpl.currentUserJSON = JSON.stringify(user || {}).replace(/<\/script>/g, '<\\/script>');

  } catch (error) {
    console.error('Error handling attendance reports data:', error);
    writeError('handleAttendanceReportsData', error);
    tpl.attendanceData = JSON.stringify({ filteredRows: [], summary: {} });
    tpl.executiveMetrics = JSON.stringify({});
    tpl.userList = [];
    tpl.managerUserId = user && user.ID ? user.ID : '';
    const fallbackTimezone = (typeof GLOBAL_SCOPE.ATTENDANCE_TIMEZONE === 'string' && GLOBAL_SCOPE.ATTENDANCE_TIMEZONE)
      ? GLOBAL_SCOPE.ATTENDANCE_TIMEZONE
      : (typeof Session !== 'undefined' && Session.getScriptTimeZone ? Session.getScriptTimeZone() : 'America/Jamaica');
    const fallbackTimezoneLabel = (typeof GLOBAL_SCOPE.ATTENDANCE_TIMEZONE_LABEL === 'string' && GLOBAL_SCOPE.ATTENDANCE_TIMEZONE_LABEL)
      ? GLOBAL_SCOPE.ATTENDANCE_TIMEZONE_LABEL
      : 'Company Time';
    tpl.attendanceTimezone = fallbackTimezone;
    tpl.attendanceTimezoneLabel = fallbackTimezoneLabel;
    tpl.attendanceDataJSON = tpl.attendanceData;
    tpl.userListJSON = JSON.stringify([]);
    tpl.currentUserJSON = JSON.stringify(user || {}).replace(/<\/script>/g, '<\\/script>');
  }
}

function handleCallReportsData(tpl, e, user, campaignId) {
  try {
    const granularity = e.parameter.granularity || "Week";
    const periodValue = e.parameter.period || weekStringFromDate(new Date());
    const selectedAgent = e.parameter.agent || "";

    if (typeof getAnalyticsByPeriod === 'function') {
      const analytics = getAnalyticsByPeriod("Week", periodValue, "");
      const rawReps = analytics.repMetrics || [];
      const pageNum = parseInt(e.parameter.page, 10) || 1;
      const PAGE_SIZE = 50;
      const startPageIdx = (pageNum - 1) * PAGE_SIZE;
      const pageSlice = rawReps.slice(startPageIdx, startPageIdx + PAGE_SIZE);

      const formattedReps = pageSlice.map((r) => {
        const totalCalls = r.totalCalls || 0;
        const totalTalkDecimal = parseFloat(r.totalTalk) || 0;
        const totalTalkFormatted = formatDuration ? formatDuration(totalTalkDecimal) : totalTalkDecimal;
        const avgTalkDecimal = totalCalls > 0 ? totalTalkDecimal / totalCalls : 0;
        const avgTalkFormatted = formatDuration ? formatDuration(avgTalkDecimal) : avgTalkDecimal;
        return {
          agent: r.agent,
          totalCalls: totalCalls,
          totalTalkFormatted: totalTalkFormatted,
          avgTalkFormatted: avgTalkFormatted,
        };
      });

      tpl.PAGE_SIZE = PAGE_SIZE;
      tpl.data = formattedReps;

      // Use manager-filtered user list
      const requestingUserId = user && user.ID ? user.ID : null;
      tpl.userList = clientGetAssignedAgentNames(campaignId || user.CampaignID || '', requestingUserId);

      // Chart data
      tpl.callVolumeLast7 = JSON.stringify(analytics.callTrend || []);
      tpl.hourlyHeatmapLast7 = JSON.stringify(analytics.hourlyHeatmap || []);
      tpl.avgIntervalByAgentLast7 = JSON.stringify(analytics.avgInterval || []);
      tpl.talkTimeByAgentLast7 = JSON.stringify(analytics.talkTrend || []);
      tpl.wrapupCountsLast7 = JSON.stringify(analytics.wrapDist || []);
      tpl.csatDistLast7 = JSON.stringify(analytics.csatDist || []);
      tpl.policyCountsLast7 = JSON.stringify(analytics.policyDist || []);
      tpl.agentLeaderboardLast7 = JSON.stringify(analytics.repMetrics || []);
    } else {
      tpl.PAGE_SIZE = 50;
      tpl.data = [];

      // Use manager-filtered user list
      const requestingUserId = user && user.ID ? user.ID : null;
      tpl.userList = clientGetAssignedAgentNames(campaignId || user.CampaignID || '', requestingUserId);

      tpl.callVolumeLast7 = JSON.stringify([]);
      tpl.talkTimeByAgentLast7 = JSON.stringify([]);
      tpl.wrapupCountsLast7 = JSON.stringify([]);
      tpl.csatDistLast7 = JSON.stringify([]);
      tpl.policyCountsLast7 = JSON.stringify([]);
      tpl.agentLeaderboardLast7 = JSON.stringify([]);
    }
  } catch (error) {
    console.error('Error handling call reports data:', error);
    writeError('handleCallReportsData', error);
    tpl.PAGE_SIZE = 50;
    tpl.data = [];
    tpl.userList = [];
  }
}

// Updated QA Form Data Handlers for Cookie-Based Auth
function handleQAFormData(tpl, e, user, templateName, campaignId) {
  try {
    const qaId = (e.parameter.id || "").toString();
    tpl.recordId = qaId;

    // Get manager-filtered user list based on template and user
    let userList = [];
    const requestingUserId = user && user.ID ? user.ID : null;

    console.log('handleQAFormData - Getting users for template:', templateName, 'campaignId:', campaignId, 'requestingUserId:', requestingUserId);

    if (templateName.includes('Independence')) {
      userList = clientGetIndependenceUsers(requestingUserId);
      tpl.campaignName = 'Independence Insurance';
    } else if (templateName.includes('CreditSuite')) {
      userList = clientGetCreditSuiteUsers(campaignId, requestingUserId);
      tpl.campaignName = 'Credit Suite';
    } else if (templateName.includes('IBTR')) {
      userList = clientGetIBTRUsers(requestingUserId);
      tpl.campaignName = 'Benefits Resource Center (iBTR)';
    } else {
      // For generic QA forms, use the new unified function
      userList = getQAFormUsers(campaignId, requestingUserId);
    }

    console.log('handleQAFormData - Final user list count:', userList.length);

    // Set both users and userList for backward compatibility
    tpl.users = userList;
    tpl.userList = userList;

    // Get existing record if editing
    if (qaId) {
      if (templateName.includes('Independence') && typeof clientGetIndependenceQAById === 'function') {
        const record = clientGetIndependenceQAById(qaId);
        tpl.record = record ? JSON.stringify(record).replace(/<\/script>/g, '<\\/script>') : "{}";
      } else if (templateName.includes('CreditSuite') && typeof clientGetCreditSuiteQAById === 'function') {
        const record = clientGetCreditSuiteQAById(qaId);
        tpl.record = record ? JSON.stringify(record).replace(/<\/script>/g, '<\\/script>') : "{}";
      } else if (typeof getQARecordById === 'function') {
        const record = getQARecordById(qaId);
        tpl.record = record ? JSON.stringify(record).replace(/<\/script>/g, '<\\/script>') : "{}";
      } else {
        tpl.record = "{}";
      }
    } else {
      tpl.record = "{}";
    }
  } catch (error) {
    console.error('Error handling QA form data:', error);
    tpl.users = [];
    tpl.userList = [];
    tpl.recordId = "";
    tpl.record = "{}";
  }
}

function handleQAViewData(tpl, e, user, templateName) {
  try {
    const qaViewId = (e.parameter.id || '').toString();
    tpl.recordId = qaViewId;

    if (!qaViewId) {
      tpl.record = "{}";
      return;
    }

    if (templateName.includes('Independence') && typeof clientGetIndependenceQAById === 'function') {
      const qaRec = clientGetIndependenceQAById(qaViewId);
      tpl.record = JSON.stringify(qaRec || {}).replace(/<\/script>/g, '<\\/script>');
      tpl.campaignName = 'Independence Insurance';
    } else if (templateName.includes('CreditSuite') && typeof clientGetCreditSuiteQAById === 'function') {
      const qaRec = clientGetCreditSuiteQAById(qaViewId);
      tpl.record = JSON.stringify(qaRec || {}).replace(/<\/script>/g, '<\\/script>');
      tpl.campaignName = 'Credit Suite';
    } else if (typeof getQARecordById === 'function') {
      const qaRec = getQARecordById(qaViewId);
      tpl.record = JSON.stringify(qaRec || {}).replace(/<\/script>/g, '<\\/script>');
    } else {
      tpl.record = "{}";
    }
  } catch (error) {
    console.error('Error handling QA view data:', error);
    tpl.record = "{}";
    tpl.recordId = "";
  }
}

function handleQADashboardData(tpl, e, user, templateName, campaignId) {
  try {
    const granularity = e.parameter.granularity || "Week";
    const periodValue = e.parameter.period || weekStringFromDate(new Date());
    const selectedAgent = e.parameter.agent || "";

    tpl.userList = clientGetAssignedAgentNames(campaignId || '');

    if (templateName === 'UnifiedQADashboard') {
      // For unified dashboard, get both campaign analytics
      try {
        if (typeof clientGetIndependenceQAAnalytics === 'function') {
          const independenceAnalytics = clientGetIndependenceQAAnalytics(granularity, periodValue, selectedAgent, "");
          tpl.independenceAnalytics = JSON.stringify(independenceAnalytics).replace(/<\/script>/g, '<\\/script>');
        } else {
          tpl.independenceAnalytics = JSON.stringify(getEmptyQAAnalytics());
        }
      } catch (error) {
        console.error('Error getting Independence analytics:', error);
        tpl.independenceAnalytics = JSON.stringify(getEmptyQAAnalytics());
      }

      try {
        if (typeof clientGetCreditSuiteQAAnalytics === 'function') {
          const creditSuiteAnalytics = clientGetCreditSuiteQAAnalytics(granularity, periodValue, selectedAgent, "");
          tpl.creditSuiteAnalytics = JSON.stringify(creditSuiteAnalytics).replace(/<\/script>/g, '<\\/script>');
        } else {
          tpl.creditSuiteAnalytics = JSON.stringify(getEmptyQAAnalytics());
        }
      } catch (error) {
        console.error('Error getting Credit Suite analytics:', error);
        tpl.creditSuiteAnalytics = JSON.stringify(getEmptyQAAnalytics());
      }

      tpl.qaAnalytics = tpl.independenceAnalytics;

    } else if (templateName.includes('CreditSuite')) {
      tpl.campaignName = 'Credit Suite';
      try {
        if (typeof clientGetCreditSuiteQAAnalytics === 'function') {
          const analytics = clientGetCreditSuiteQAAnalytics(granularity, periodValue, selectedAgent);
          tpl.qaAnalytics = JSON.stringify(analytics).replace(/<\/script>/g, '<\\/script>');
        } else {
          tpl.qaAnalytics = JSON.stringify(getEmptyQAAnalytics());
        }
      } catch (error) {
        console.error('Error getting Credit Suite QA analytics:', error);
        tpl.qaAnalytics = JSON.stringify(getEmptyQAAnalytics());
      }

    } else if (templateName.includes('Independence')) {
      tpl.campaignName = 'Independence Insurance';
      try {
        if (typeof clientGetIndependenceQAAnalytics === 'function') {
          const analytics = clientGetIndependenceQAAnalytics(granularity, periodValue, selectedAgent);
          tpl.qaAnalytics = JSON.stringify(analytics).replace(/<\/script>/g, '<\\/script>');
        } else {
          tpl.qaAnalytics = JSON.stringify(getEmptyQAAnalytics());
        }
      } catch (error) {
        console.error('Error getting Independence QA analytics:', error);
        tpl.qaAnalytics = JSON.stringify(getEmptyQAAnalytics());
      }

    } else {
      // Default QA dashboard
      try {
        if (typeof getAllQA === 'function') {
          const qaRecords = getAllQA();
          tpl.qaRecords = JSON.stringify(qaRecords).replace(/<\/script>/g, '<\\/script>');
        } else {
          tpl.qaRecords = JSON.stringify([]);
        }
      } catch (error) {
        console.error('Error getting QA dashboard data:', error);
        tpl.qaRecords = JSON.stringify([]);
      }
    }

  } catch (error) {
    console.error('Error handling QA dashboard data:', error);
    tpl.userList = [];
    tpl.qaAnalytics = JSON.stringify(getEmptyQAAnalytics());
  }
}

function handleQAListData(tpl, e, user, templateName) {
  try {
    if (templateName.includes('Independence')) {
      let qaRecords = [];
      if (typeof clientGetIndependenceQARecords === 'function') {
        qaRecords = clientGetIndependenceQARecords();
      }
      tpl.qaRecords = JSON.stringify(qaRecords).replace(/<\/script>/g, '<\\/script>');
      tpl.campaignName = 'Independence Insurance';

    } else if (templateName.includes('CreditSuite')) {
      let qaRecords = [];
      if (typeof clientGetCreditSuiteQARecords === 'function') {
        qaRecords = clientGetCreditSuiteQARecords();
      }
      tpl.qaRecords = JSON.stringify(qaRecords).replace(/<\/script>/g, '<\\/script>');
      tpl.campaignName = 'Credit Suite';

    } else {
      // Default or other campaign QA list
      if (typeof getAllQA === 'function') {
        tpl.qaRecords = JSON.stringify(getAllQA()).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.qaRecords = JSON.stringify([]);
      }
    }
  } catch (error) {
    console.error('Error handling QA list data:', error);
    tpl.qaRecords = JSON.stringify([]);
  }
}

function handleUsersData(tpl, e, user) {
  try {
    tpl.campaignList = getAllCampaigns();
    const roleList = getAllRoles() || [];
    tpl.roleList = roleList.filter(r => {
      const n = String(r.name || r.Name || '').toLowerCase();
      return n !== 'super admin' && n !== 'administrator';
    });
    const knownPages = Object.keys(_loadAllPagesMeta());
    tpl.pagesList = knownPages;
  } catch (error) {
    writeError('handleUsersData', error);
    tpl.campaignList = [];
    tpl.roleList = [];
    tpl.pagesList = [];
  }
}

function handleTaskData(tpl, e, user, templateName) {
  try {
    const currentUserEmail = user.email || user.Email || '';

    if (templateName === 'TaskForm') {
      tpl.currentUser = currentUserEmail;
      tpl.users = getUsers().map(u => u.email || u.Email).filter(e => e);

      if (typeof getTasksFor === 'function') {
        const visible = getTasksFor(currentUserEmail);
        tpl.tasksJson = JSON.stringify(visible).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.tasksJson = JSON.stringify([]);
      }

      const taskId = (e.parameter.id || '').toString();
      tpl.recordId = taskId;

      if (taskId && typeof getTaskById === 'function') {
        tpl.record = JSON.stringify(getTaskById(taskId)).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.record = '{}';
      }

    } else if (templateName === 'TaskList') {
      tpl.currentUser = currentUserEmail;

      if (typeof getTasksFor === 'function') {
        tpl.tasks = JSON.stringify(getTasksFor(currentUserEmail)).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.tasks = JSON.stringify([]);
      }

      tpl.users = JSON.stringify(getUsers().map(u => u.email || u.Email).filter(e => e));
      tpl.recordId = '';
      tpl.record = '{}';

    } else if (templateName === 'TaskBoard') {
      tpl.currentUser = currentUserEmail;

      if (typeof getTasksFor === 'function') {
        const visibleTasks = getTasksFor(currentUserEmail);
        tpl.tasksJson = JSON.stringify(visibleTasks).replace(/<\/script>/g, '<\\/script>');

        const ownersArr = [...new Set(visibleTasks.map(t => t.Owner).filter(x => x))].sort();
        tpl.ownersJson = JSON.stringify(ownersArr);
        tpl.selectedOwner = e.parameter.owner || '';
      } else {
        tpl.tasksJson = JSON.stringify([]);
        tpl.ownersJson = JSON.stringify([]);
        tpl.selectedOwner = '';
      }
    }

  } catch (error) {
    console.error('Error handling task data:', error);
    tpl.currentUser = user.email || user.Email || '';
    tpl.users = [];
    tpl.tasksJson = JSON.stringify([]);
    tpl.recordId = '';
    tpl.record = '{}';
  }
}

function handleCoachingData(tpl, e, user, templateName) {
  try {
    const granularity = e.parameter.granularity || "Week";
    const periodValue = e.parameter.period || weekStringFromDate(new Date());
    const selectedAgent = e.parameter.agent || "";

    if (templateName === 'CoachingForm') {
      // No additional data needed beyond base template data

    } else if (templateName === 'CoachingList') {
      if (typeof getAllCoaching === 'function') {
        tpl.coachingRecords = JSON.stringify(getAllCoaching()).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.coachingRecords = JSON.stringify([]);
      }

    } else if (templateName === 'CoachingView') {
      const coachingId = (e.parameter.id || '').toString();
      tpl.recordId = coachingId;

      if (coachingId && typeof getCoachingRecordById === 'function') {
        const coachingRec = getCoachingRecordById(coachingId);
        tpl.record = JSON.stringify(coachingRec || {}).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.record = '{}';
      }

    } else if (templateName === 'CoachingDashboard') {
      if (typeof getDashboardCoaching === 'function') {
        const coachingMetrics = getDashboardCoaching(granularity, periodValue, selectedAgent);
        const coachingUserList = getUsers().map(u => u.name || u.FullName || u.UserName).filter(n => n).sort();

        tpl.userList = coachingUserList;
        tpl.sessionsTrend = JSON.stringify(coachingMetrics.sessionsTrend || []);
        tpl.topicsDist = JSON.stringify(coachingMetrics.topicsDist || []);
        tpl.upcomingFollow = JSON.stringify(coachingMetrics.upcoming || []);
      } else {
        tpl.userList = [];
        tpl.sessionsTrend = JSON.stringify([]);
        tpl.topicsDist = JSON.stringify([]);
        tpl.upcomingFollow = JSON.stringify([]);
      }
    }

  } catch (error) {
    console.error('Error handling coaching data:', error);
    tpl.userList = [];
    tpl.sessionsTrend = JSON.stringify([]);
    tpl.topicsDist = JSON.stringify([]);
    tpl.upcomingFollow = JSON.stringify([]);
  }
}

function handleEODReportData(tpl, e, user) {
  try {
    const today = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'
    );

    if (typeof getEODTasksByDate === 'function') {
      const allComplete = getEODTasksByDate(today);
      const me = (user.email || user.Email || '').toLowerCase();
      const visibleEOD = allComplete.filter(t =>
        String(t.Owner || '').toLowerCase() === me ||
        String(t.Delegations || '')
          .split(',')
          .map(x => x.trim().toLowerCase())
          .includes(me)
      );

      tpl.reportDate = today;
      tpl.tasks = JSON.stringify(visibleEOD).replace(/<\/script>/g, '<\\/script>');
    } else {
      tpl.reportDate = today;
      tpl.tasks = JSON.stringify([]);
    }
  } catch (error) {
    console.error('Error handling EOD report data:', error);
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    tpl.reportDate = today;
    tpl.tasks = JSON.stringify([]);
  }
}

function handleCalendarData(tpl, e, user, campaignId) {
  try {
    const selectedAgent = e.parameter.agent || "";

    if (e.parameter.page === 'attendancecalendar') {
      if (typeof getAttendanceAnalyticsByPeriod === 'function') {
        const today = new Date();
        const isoWeek = weekStringFromDate(today);
        const attendanceAnalytics = getAttendanceAnalyticsByPeriod("Week", isoWeek, selectedAgent);

        tpl.userList = clientGetAssignedAgentNames(campaignId || '');
        tpl.attendanceData = JSON.stringify(attendanceAnalytics).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.userList = clientGetAssignedAgentNames(campaignId || '');
        tpl.attendanceData = JSON.stringify({});
      }
    } else {
      tpl.userList = clientGetAssignedAgentNames(campaignId || '');
    }
  } catch (error) {
    console.error('Error handling calendar data:', error);
    tpl.userList = [];
    tpl.attendanceData = JSON.stringify({});
  }
}

function handleSearchData(tpl, e) {
  try {
    const query = (e.parameter.query || '').trim();
    const pageIndex = parseInt(e.parameter.pageIndex || '1', 10);
    const startIdx = (pageIndex - 1) * 10 + 1;
    let items = [], totalResults = 0, error = null;

    if (query) {
      try {
        if (typeof searchWeb === 'function') {
          const resp = searchWeb(query, startIdx);
          items = resp.items || [];
          totalResults = parseInt(resp.searchInformation.totalResults || '0', 10);

          items = items.map(item => ({
            ...item,
            proxyUrl: `${SCRIPT_URL}?page=proxy&url=${encodeURIComponent(item.link)}`,
            displayTitle: item.htmlTitle || item.title || 'Untitled',
            displaySnippet: item.htmlSnippet || item.snippet || 'No description available'
          }));
        }
      } catch (err) {
        error = err.message;
      }
    }

    tpl.query = query;
    tpl.results = items;
    tpl.totalResults = totalResults;
    tpl.pageIndex = pageIndex;
    tpl.error = error;
    tpl.scriptUrl = SCRIPT_URL;

  } catch (error) {
    console.error('Error handling search data:', error);
    tpl.query = '';
    tpl.results = [];
    tpl.totalResults = 0;
    tpl.pageIndex = 1;
    tpl.error = error.message;
    tpl.scriptUrl = SCRIPT_URL;
  }
}

function computeUserProfileSlug(userRecord, detailRecord) {
  try {
    const record = detailRecord && detailRecord.record ? detailRecord.record : detailRecord || {};
    const normalize = (value) => {
      if (value === null || typeof value === 'undefined') {
        return '';
      }
      const text = String(value).trim();
      return text ? text : '';
    };

    const identifierCandidates = [
      userRecord && (userRecord.ID || userRecord.Id || userRecord.id || userRecord.EmployeeID || userRecord.employeeId || userRecord.ProfileID || userRecord.profileId),
      record && (record.ID || record.Id || record.id || record.EmployeeID || record.employeeId || record.ProfileID || record.profileId)
    ];

    for (let idx = 0; idx < identifierCandidates.length; idx++) {
      const normalizedId = normalize(identifierCandidates[idx]);
      if (normalizedId) {
        return normalizedId;
      }
    }

    const candidates = [];

    const pushCandidate = (value) => {
      if (value === null || typeof value === 'undefined') {
        return;
      }
      const text = String(value).trim();
      if (text) {
        candidates.push(text);
      }
    };

    pushCandidate(userRecord && (userRecord.UserName || userRecord.userName || userRecord.username));
    pushCandidate(record && (record.UserName || record.userName || record.username));
    pushCandidate(userRecord && (userRecord.DisplayName || userRecord.displayName));
    pushCandidate(record && (record.DisplayName || record.displayName));

    const firstName = userRecord && (userRecord.FirstName || userRecord.firstName);
    const lastName = userRecord && (userRecord.LastName || userRecord.lastName);
    if (firstName || lastName) {
      pushCandidate([firstName, lastName].filter(Boolean).join(' '));
    }

    const recordFirstName = record && (record.FirstName || record.firstName);
    const recordLastName = record && (record.LastName || record.lastName);
    if (recordFirstName || recordLastName) {
      pushCandidate([recordFirstName, recordLastName].filter(Boolean).join(' '));
    }

    pushCandidate(userRecord && (userRecord.Email || userRecord.email));
    pushCandidate(record && (record.Email || record.email));

    for (let i = 0; i < candidates.length; i++) {
      const slug = candidates[i]
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, '-')
        .replace(/^-+|-+$/g, '');
      if (slug) {
        return slug;
      }
    }
  } catch (error) {
    console.warn('computeUserProfileSlug: failed to compute slug', error);
  }

    return 'user';
}

function handleUserProfileData(tpl, e, user) {
  try {
    const bootstrap = {
      user: user || null,
      viewer: user || null,
      detail: null,
      pages: [],
      equipment: [],
      permissions: null,
      managerSummary: null,
      campaignId: '',
      generatedAt: new Date().toISOString()
    };

    const viewerId = user && user.ID ? String(user.ID) : '';
    const rawPageParam = (e && e.parameter && e.parameter.page) || '';
    const requestedProfileId = resolveProfileIdentifierFromRequest(e, rawPageParam);
    const normalizedRequested = requestedProfileId ? String(requestedProfileId).trim() : '';
    const profileId = normalizedRequested || viewerId;
    const normalizedProfileId = profileId ? profileId.toString().trim().toLowerCase() : '';
    const requestingUserId = viewerId || profileId;

    const extractProfileId = (record) => {
      if (!record || typeof record !== 'object') {
        return '';
      }
      const keys = ['ID', 'Id', 'id', 'EmployeeID', 'employeeId', 'ProfileID', 'profileId', 'UserID', 'userId'];
      for (let i = 0; i < keys.length; i++) {
        const value = record[keys[i]];
        if (value !== undefined && value !== null) {
          const text = String(value).trim();
          if (text) {
            return text;
          }
        }
      }
      return '';
    };

    let profileRecord = null;

    if (profileId && typeof getUsers === 'function') {
      try {
        const roster = getUsers();
        if (Array.isArray(roster) && roster.length) {
          for (let idx = 0; idx < roster.length; idx++) {
            const candidate = extractProfileId(roster[idx]);
            if (candidate && candidate.toLowerCase() === normalizedProfileId) {
              profileRecord = roster[idx];
              break;
            }
          }
        }
      } catch (rosterError) {
        console.warn('handleUserProfileData: unable to resolve profile user from roster', rosterError);
      }
    }

    if (profileId && typeof clientGetUserDetail === 'function') {
      try {
        const detailOptions = requestingUserId ? { requestingUserId: requestingUserId } : {};
        bootstrap.detail = clientGetUserDetail(profileId, detailOptions) || null;
        if (!profileRecord && bootstrap.detail && bootstrap.detail.record) {
          profileRecord = bootstrap.detail.record;
        }
      } catch (detailError) {
        console.warn('handleUserProfileData: unable to load user detail', detailError);
      }
    }

    if (!profileRecord && viewerId && profileId && viewerId.toLowerCase() === normalizedProfileId) {
      profileRecord = user || null;
    }

    if (!profileRecord && profileId && typeof clientGetUserProfile === 'function') {
      try {
        const fallbackRecord = clientGetUserProfile(profileId);
        if (fallbackRecord) {
          profileRecord = fallbackRecord;
        }
      } catch (profileError) {
        console.warn('handleUserProfileData: fallback profile lookup failed', profileError);
      }
    }

    if (profileId && typeof clientGetUserPages === 'function') {
      try {
        bootstrap.pages = clientGetUserPages(profileId) || [];
      } catch (pagesError) {
        console.warn('handleUserProfileData: unable to load user pages', pagesError);
      }
    }

    if (profileId && typeof clientGetUserEquipment === 'function') {
      try {
        const equipmentResponse = clientGetUserEquipment(profileId);
        if (equipmentResponse && equipmentResponse.success && Array.isArray(equipmentResponse.items)) {
          bootstrap.equipment = equipmentResponse.items;
        }
      } catch (equipmentError) {
        console.warn('handleUserProfileData: unable to load user equipment', equipmentError);
      }
    }

    let campaignId = '';
    if (profileRecord && (profileRecord.CampaignID || profileRecord.campaignId)) {
      campaignId = String(profileRecord.CampaignID || profileRecord.campaignId);
    }

    if (!campaignId && bootstrap.detail && bootstrap.detail.record) {
      const record = bootstrap.detail.record;
      campaignId = String(record.CampaignID || record.campaignId || record.CampaignId || '');
    }

    bootstrap.campaignId = campaignId;

    if (profileId && campaignId && typeof getCampaignUserPermissions === 'function') {
      try {
        bootstrap.permissions = getCampaignUserPermissions(campaignId, profileId) || null;
      } catch (permissionsError) {
        console.warn('handleUserProfileData: unable to load campaign permissions', permissionsError);
      }
    }

    if (profileId && typeof clientGetManagerTeamSummary === 'function') {
      try {
        const summary = clientGetManagerTeamSummary(profileId);
        if (summary && summary.success) {
          bootstrap.managerSummary = summary;
        }
      } catch (managerSummaryError) {
        console.warn('handleUserProfileData: unable to load manager summary', managerSummaryError);
      }
    }

    bootstrap.user = profileRecord || user || null;
    bootstrap.profileId = profileId;
    bootstrap.requestedProfileId = normalizedRequested;
    const computedSlug = computeUserProfileSlug(bootstrap.user, bootstrap.detail);
    bootstrap.profileSlug = computedSlug && computedSlug !== 'user' ? computedSlug : (profileId || computedSlug);

    tpl.profileBootstrap = _stringifyForTemplate_(bootstrap);
  } catch (error) {
    console.error('Error handling user profile data:', error);
    try {
      writeError && writeError('handleUserProfileData', error);
    } catch (_) { }
    tpl.profileBootstrap = '{}';
  }
}

function handleChatData(tpl, e, user) {
  try {
    const groupId = e.parameter.groupId || '';
    const channelId = e.parameter.channelId || '';
    tpl.groupId = groupId;
    tpl.channelId = channelId;

    let userGroups = [];
    let groupChannels = [];
    let messages = [];

    try {
      if (typeof getUserChatGroups === 'function') {
        userGroups = getUserChatGroups(user.ID) || [];
      }
    } catch (e) {
      console.warn('Error getting user groups:', e);
    }

    try {
      if (groupId && typeof getChatChannels === 'function') {
        groupChannels = getChatChannels(groupId);
      }
    } catch (e) {
      console.warn('Error getting group channels:', e);
    }

    try {
      if (channelId && typeof getChatMessages === 'function') {
        messages = getChatMessages(channelId);
      }
    } catch (e) {
      console.warn('Error getting messages:', e);
    }

    tpl.userGroups = JSON.stringify(userGroups).replace(/<\/script>/g, '<\\/script>');
    tpl.groupChannels = JSON.stringify(groupChannels).replace(/<\/script>/g, '<\\/script>');
    tpl.messages = JSON.stringify(messages).replace(/<\/script>/g, '<\\/script>');
  } catch (error) {
    console.error('Error handling chat data:', error);
    tpl.groupId = '';
    tpl.channelId = '';
    tpl.userGroups = JSON.stringify([]);
    tpl.groupChannels = JSON.stringify([]);
    tpl.messages = JSON.stringify([]);
  }
}

function handleNotificationsData(tpl, e, user) {
  try {
    if (typeof getAllTasks === 'function') {
      const notifToday = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const dueTasks = getAllTasks()
        .filter(t => (t.Status !== 'Done' && t.Status !== 'Completed') && t.DueDate === notifToday);
      tpl.tasksJson = JSON.stringify(dueTasks).replace(/<\/script>/g, '<\\/script>');
    } else {
      tpl.tasksJson = JSON.stringify([]);
    }
  } catch (error) {
    console.error('Error handling notifications data:', error);
    tpl.tasksJson = JSON.stringify([]);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// USER MANAGEMENT FUNCTIONS (Updated for Cookie Auth)
// ───────────────────────────────────────────────────────────────────────────────

function _normStr_(s) { return String(s || '').trim(); }
function _normEmail_(s) { return _normStr_(s).toLowerCase(); }
function _toBool_(v) { return v === true || String(v || '').trim().toUpperCase() === 'TRUE'; }

function _tryGetRoleNames_(userId) {
  try {
    if (typeof getUserRoleNamesSafe === 'function') {
      var list = getUserRoleNamesSafe(userId) || [];
      return Array.isArray(list) ? list : [];
    }
  } catch (_) { }
  return [];
}

const _GUEST_USER_KEYWORDS_ = ['guest', 'client', 'partner', 'collab'];

function _normLower_(value) { return _normStr_(value).toLowerCase(); }

function _extractCampaignIdsFromUser_(user) {
  const ids = new Set();
  const primary = _normStr_(user && (user.CampaignID || user.CampaignId || user.campaignId));
  if (primary) ids.add(primary);

  const rawList = user && (user.CampaignIds || user.campaignIds);
  if (Array.isArray(rawList)) {
    rawList.forEach(function (val) {
      const normalized = _normStr_(val);
      if (normalized) ids.add(normalized);
    });
  } else if (typeof rawList !== 'undefined' && rawList !== null) {
    String(rawList || '')
      .split(/[,\s]+/)
      .map(function (part) { return _normStr_(part); })
      .filter(Boolean)
      .forEach(function (val) { ids.add(val); });
  }

  return Array.from(ids);
}

function _collectUserRoleTokens_(user) {
  const tokens = [];
  if (!user) return tokens;

  const directFields = [
    user.Role, user.role, user.PrimaryRole, user.primaryRole,
    user.UserType, user.userType,
    user.Persona, user.persona, user.PersonaType, user.personaType,
    user.AccessLevel, user.accessLevel,
    user.AccountType, user.accountType,
    user.Classification, user.classification,
    user.EmploymentStatus, user.employmentStatus,
    user.Department, user.department,
    user.Title, user.title,
    user.JobTitle, user.jobTitle
  ];

  directFields.forEach(function (value) {
    const normalized = _normLower_(value);
    if (normalized) tokens.push(normalized);
  });

  const rolesCsv = _normLower_(user && (user.Roles || user.roles || ''));
  if (rolesCsv) {
    rolesCsv.split(/[,\s]+/).forEach(function (part) {
      const normalized = String(part || '').trim();
      if (normalized) tokens.push(normalized);
    });
  }

  if (Array.isArray(user && user.roleNames)) {
    user.roleNames.forEach(function (name) {
      const normalized = _normLower_(name);
      if (normalized) tokens.push(normalized);
    });
  } else if (user && user.ID) {
    try {
      const safeRoles = _tryGetRoleNames_(user.ID) || [];
      safeRoles.forEach(function (name) {
        const normalized = _normLower_(name);
        if (normalized) tokens.push(normalized);
      });
    } catch (_) { }
  }

  return tokens;
}

function _isGuestUser_(user) {
  const tokens = _collectUserRoleTokens_(user);
  if (!tokens.length) return false;
  return tokens.some(function (token) {
    return _GUEST_USER_KEYWORDS_.some(function (keyword) {
      return token.indexOf(keyword) !== -1;
    });
  });
}

function _filterUsersForCampaign_(rows, campaignId, options) {
  const cid = _normStr_(campaignId);
  const excludeGuests = !options || options.excludeGuests !== false;
  const list = Array.isArray(rows) ? rows : [];

  return list.filter(function (user) {
    if (!user) return false;
    if (cid) {
      const campaigns = _extractCampaignIdsFromUser_(user);
      if (campaigns.indexOf(cid) === -1) {
        return false;
      }
    }
    if (excludeGuests && _isGuestUser_(user)) {
      return false;
    }
    return true;
  });
}

function _resolveCampaignIdFromInput_(input) {
  if (typeof input === 'string' || typeof input === 'number') {
    const candidate = _normStr_(input);
    if (!candidate) return '';
    if (/^\d+$/.test(candidate)) {
      const userRow = _findUserById_(candidate);
      if (userRow) {
        const fromRow = _normStr_(userRow.CampaignID || userRow.campaignId);
        if (fromRow) return fromRow;
      }
    }
    return candidate;
  }

  if (Array.isArray(input)) {
    for (let i = 0; i < input.length; i++) {
      const candidate = _resolveCampaignIdFromInput_(input[i]);
      if (candidate) return candidate;
    }
    return '';
  }

  if (input && typeof input === 'object') {
    const props = ['campaignId', 'CampaignID', 'campaignid', 'CampaignId'];
    for (let i = 0; i < props.length; i++) {
      if (Object.prototype.hasOwnProperty.call(input, props[i])) {
        const candidate = _normStr_(input[props[i]]);
        if (candidate) return candidate;
      }
    }

    const userProps = ['managerUserId', 'userId', 'UserID', 'managerId'];
    for (let j = 0; j < userProps.length; j++) {
      if (Object.prototype.hasOwnProperty.call(input, userProps[j])) {
        const candidate = _resolveCampaignIdFromInput_(input[userProps[j]]);
        if (candidate) return candidate;
      }
    }
  }

  const currentUser = (typeof getCurrentUser === 'function') ? getCurrentUser() : null;
  if (currentUser) {
    const fallback = _normStr_(
      currentUser.activeCampaignId ||
      currentUser.activeCampaignID ||
      currentUser.campaignId ||
      currentUser.CampaignId ||
      currentUser.CampaignID
    );
    if (fallback) return fallback;
  }

  return '';
}

function _campaignNameMap_() {
  try {
    const rows = (typeof readSheet === 'function') ? (readSheet('CAMPAIGNS') || []) : [];
    const map = {};
    rows.forEach(function (r) {
      const id = _normStr_(r.ID || r.Id || r.id);
      const nm = _normStr_(r.Name || r.name);
      if (id) map[id] = nm;
    });
    return map;
  } catch (_) {
    return {};
  }
}

function _uiUserShape_(u, cmap) {
  const primaryCampaignId = _normStr_(u.CampaignID || u.campaignId);
  const campaignIdsRaw = u.CampaignIds || u.campaignIds || '';
  const campaignIds = Array.isArray(campaignIdsRaw)
    ? campaignIdsRaw.filter(Boolean).map(function (val) { return _normStr_(val); })
    : String(campaignIdsRaw || '')
        .split(/[,\s]+/)
        .map(function (part) { return _normStr_(part); })
        .filter(Boolean);

  const employmentStatus = _normStr_(u.EmploymentStatus || u.employmentStatus || u.Status || '');
  const department = _normStr_(u.Department || u.department || '');
  const role = _normStr_(u.Role || u.role || u.PrimaryRole || '');
  const managerEmail = _normStr_(u.ManagerEmail || u.managerEmail || '');
  const activeFlag = (typeof u.Active === 'undefined') ? u.active : u.Active;
  const activeBool = (function (val) {
    if (typeof val === 'boolean') return val;
    const str = String(val || '').trim().toLowerCase();
    if (!str) return false;
    return str === 'true' || str === 'yes' || str === '1' || str === 'active';
  })(activeFlag);

  return {
    ID: u.ID,
    id: u.ID,
    UserName: u.UserName,
    FullName: u.FullName,
    name: _normStr_(u.FullName || u.UserName || ''),
    Email: u.Email,
    email: _normStr_(u.Email || ''),
    CampaignID: u.CampaignID,
    campaignName: cmap[primaryCampaignId] || _normStr_(u.CampaignName || ''),
    CampaignIds: campaignIds,
    campaignIds: campaignIds,
    CanLogin: u.CanLogin,
    IsAdmin: u.IsAdmin,
    canLoginBool: _toBool_(u.CanLogin),
    isAdminBool: _toBool_(u.IsAdmin),
    Active: typeof u.Active === 'undefined' ? u.active : u.Active,
    active: activeBool,
    activeBool: activeBool,
    EmploymentStatus: employmentStatus,
    employmentStatus: employmentStatus,
    Department: department,
    department: department,
    Role: role,
    role: role,
    ManagerEmail: managerEmail,
    managerEmail: managerEmail,
    roleNames: _tryGetRoleNames_(u.ID),
    pages: []
  };
}

function _readUsersSheetSafe_(options) {
  try {
    let opts = {};
    if (typeof options === 'string' || typeof options === 'number') {
      opts.campaignId = options;
    } else if (options && typeof options === 'object') {
      opts = Object.assign({}, options);
    }

    const campaignId = _resolveCampaignIdFromInput_(opts.campaignId);
    const excludeGuests = opts.excludeGuests !== false;
    const allowAllCampaigns = opts.allowAllCampaigns === true;

    const readOptions = { useCache: opts.useCache !== false };
    if (campaignId && !allowAllCampaigns) {
      readOptions.campaignId = campaignId;
      readOptions.where = Object.assign({}, opts.where || {}, { CampaignID: campaignId });
    }
    if (opts.suppressTenantContext) {
      readOptions.suppressTenantContext = true;
    }

    let rows = (typeof readSheet === 'function') ? (readSheet('Users', readOptions) || []) : [];
    if (!Array.isArray(rows)) rows = [];

    if (allowAllCampaigns && campaignId) {
      rows = rows.filter(function (row) {
        const campaigns = _extractCampaignIdsFromUser_(row);
        return campaigns.indexOf(campaignId) !== -1;
      });
    }

    if (campaignId || excludeGuests) {
      rows = _filterUsersForCampaign_(rows, allowAllCampaigns ? (campaignId || '') : campaignId, { excludeGuests: excludeGuests });
    }

    return rows;
  } catch (err) {
    console.warn('Unable to read Users sheet:', err);
    return [];
  }
}

function _readManagerUsersSheetSafe_() {
  try {
    return (typeof readSheet === 'function') ? (readSheet('MANAGER_USERS') || []) : [];
  } catch (err) {
    console.warn('Unable to read MANAGER_USERS sheet:', err);
    return [];
  }
}

function _dedupeAndSortUsers_(list) {
  const seenIds = new Set();
  const seenEmails = new Set();
  const out = [];
  for (let i = 0; i < list.length; i++) {
    const user = list[i];
    if (!user) continue;
    const idKey = String(user.ID || '');
    const emailKey = String(user.email || user.Email || '').trim().toLowerCase();
    if (idKey && seenIds.has(idKey)) continue;
    if (emailKey && seenEmails.has(emailKey)) continue;
    if (idKey) seenIds.add(idKey);
    if (emailKey) seenEmails.add(emailKey);
    out.push(user);
  }
  out.sort(function (a, b) {
    return String(a.name || a.FullName || '').localeCompare(String(b.name || b.FullName || '')) ||
      String(a.UserName || '').localeCompare(String(b.UserName || ''));
  });
  return out;
}

function getUsersByManager(managerUserId, options) {
  try {
    const opts = Object.assign({ excludeGuests: true }, options || {});
    const readOpts = {
      excludeGuests: opts.excludeGuests !== false,
      allowAllCampaigns: true
    };

    const cmap = _campaignNameMap_();
    const rows = _readUsersSheetSafe_(readOpts);
    const shaped = rows.map(function (row) { return _uiUserShape_(row, cmap); });

    return _dedupeAndSortUsers_(shaped);

  } catch (error) {
    console.error('Error in getUsersByManager:', error);
    writeError && writeError('getUsersByManager', error);
    return [];
  }
}

function getUser(managerUserId, options) {
  try {
    const finalOpts = (typeof options === 'object' && options !== null)
      ? Object.assign({}, options)
      : {};
    return getUsersByManager(managerUserId, finalOpts);

  } catch (error) {
    console.error('Error in getUser:', error);
    writeError && writeError('getUser', error);
    return [];
  }
}

function getUsers(options) {
  try {
    const opts = (typeof options === 'string' || typeof options === 'number')
      ? { campaignId: options }
      : (options && typeof options === 'object')
        ? Object.assign({}, options)
        : {};

    const excludeGuests = opts.excludeGuests !== false;
    const cmap = _campaignNameMap_();

    let rows = _readUsersSheetSafe_({
      allowAllCampaigns: true,
      excludeGuests: excludeGuests
    });

    const campaignCandidates = [opts.campaignId, opts.activeCampaignId, opts.managerCampaignId];
    const resolvedCampaignId = _resolveCampaignIdFromInput_(campaignCandidates);
    if (resolvedCampaignId) {
      rows = rows.filter(function (row) {
        const campaigns = _extractCampaignIdsFromUser_(row);
        return campaigns.indexOf(resolvedCampaignId) !== -1;
      });
    }

    const shaped = rows.map(function (row) { return _uiUserShape_(row, cmap); });
    return _dedupeAndSortUsers_(shaped);

  } catch (e) {
    console.error('Error in getUsers:', e);
    writeError && writeError('getUsers (enhanced)', e);
    return [];
  }
}

function clientGetAssignedAgentNames(campaignInput) {
  try {
    const campaignId = _resolveCampaignIdFromInput_(campaignInput);
    console.log('clientGetAssignedAgentNames resolved campaignId:', campaignId, 'from input:', campaignInput);

    const cmap = _campaignNameMap_();
    let users = [];

    try {
      users = getUsers({ campaignId: campaignId, excludeGuests: true });
      if (Array.isArray(users)) {
        console.log('getUsers() returned:', users.length, 'users');
      } else {
        console.warn('getUsers() returned non-array data');
        users = [];
      }
    } catch (error) {
      console.warn('getUsers() failed, trying fallback approaches:', error);
      users = [];
    }

    if (!users.length) {
      try {
        const rows = _readUsersSheetSafe_({ campaignId: campaignId, excludeGuests: true });
        if (rows.length) {
          users = rows.map(function (row) { return _uiUserShape_(row, cmap); });
          console.log('Hydrated users from campaign-scoped sheet read:', users.length);
        }
      } catch (error) {
        console.error('Campaign-scoped user read failed:', error);
      }
    }

    if (!users.length) {
      try {
        const currentUser = getCurrentUser();
        if (currentUser && !_isGuestUser_(currentUser)) {
          users = [_uiUserShape_(currentUser, cmap)];
          console.log('Using current user as fallback:', currentUser.FullName || currentUser.UserName || currentUser.Email);
        }
      } catch (error) {
        console.error('getCurrentUser fallback failed:', error);
      }
    }

    if (!users.length) {
      console.warn('All user retrieval strategies failed, using empty array');
      users = [];
    }

    const displayNames = users
      .map(function (user) {
        const name = user && (user.FullName || user.UserName || user.name || user.Email || '');
        return String(name || '').trim();
      })
      .filter(function (name) { return name !== ''; })
      .filter(function (name, index, arr) { return arr.indexOf(name) === index; })
      .sort(function (a, b) { return a.localeCompare(b, undefined, { sensitivity: 'base' }); });

    console.log('Final userList:', displayNames);
    return displayNames;

  } catch (error) {
    console.error('Error in clientGetAssignedAgentNames:', error);
    writeError('clientGetAssignedAgentNames', error);
    return [];
  }
}

function clientGetUserList(campaignId) {
  try {
    let userList = clientGetAssignedAgentNames(campaignId);

    if (!userList || userList.length === 0) {
      console.log('Assigned approach failed, trying simple approach');
      userList = getSimpleUserList(campaignId);
    }

    if (!userList || userList.length === 0) {
      console.log('Campaign-filtered approach failed, trying all users');
      userList = getSimpleUserList('');
    }

    return userList;

  } catch (error) {
    console.error('Error in clientGetUserList:', error);
    return [];
  }
}

function getSimpleUserList(campaignInput) {
  try {
    const campaignId = _resolveCampaignIdFromInput_(campaignInput);
    console.log('getSimpleUserList called with campaignId:', campaignId);

    const users = _readUsersSheetSafe_({ campaignId: campaignId, excludeGuests: true });
    console.log('Read', users.length, 'users from scoped reader');

    const names = users
      .map(function (u) { return u.FullName || u.UserName || u.Email || u.name || ''; })
      .map(function (name) { return String(name || '').trim(); })
      .filter(function (name) { return name !== ''; })
      .filter(function (name, index, arr) { return arr.indexOf(name) === index; })
      .sort();

    console.log('Returning', names.length, 'user names');
    return names;

  } catch (error) {
    console.error('Error in getSimpleUserList:', error);
    return [];
  }
}

function getUsersByCampaign(campaignInput) {
  try {
    const campaignId = _resolveCampaignIdFromInput_(campaignInput);
    return _readUsersSheetSafe_({ campaignId: campaignId, excludeGuests: true });
  } catch (error) {
    console.warn('getUsersByCampaign failed:', error);
    return [];
  }
}

function getAllUsersRaw(campaignInput, options) {
  try {
    let opts = {};
    if (campaignInput && typeof campaignInput === 'object' && !Array.isArray(campaignInput)) {
      opts = Object.assign({}, campaignInput);
    } else {
      opts = Object.assign({}, options || {});
      if (typeof campaignInput !== 'undefined') {
        opts.campaignId = campaignInput;
      }
    }

    const campaignId = _resolveCampaignIdFromInput_(opts.campaignId);
    return _readUsersSheetSafe_({
      campaignId: campaignId,
      excludeGuests: opts.excludeGuests !== false,
      allowAllCampaigns: opts.allowAllCampaigns === true
    });
  } catch (error) {
    console.error('Error getting raw user list:', error);
    return [];
  }
}

function getAllUsers() {
  try {
    if (typeof getUsers === 'function') {
      const scoped = getUsers();
      if (Array.isArray(scoped)) {
        return scoped;
      }
    }

    console.warn('getAllUsers: getUsers unavailable or returned invalid data; defaulting to empty list');
    return [];
  } catch (error) {
    console.error('Error getting manager-scoped users:', error);
    return [];
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// CAMPAIGN-SPECIFIC USER FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

function clientGetIndependenceUsers() {
  try {
    return getIndependenceInsuranceUsers();
  } catch (error) {
    console.error('Error in clientGetIndependenceUsers:', error);
    writeError('clientGetIndependenceUsers', error);
    return [];
  }
}

function getIndependenceInsuranceUsers() {
  try {
    const allUsers = getAllUsers();
    const users = allUsers.filter(user => {
      const campaignName = user.CampaignName || user.campaignName || '';
      const userCampaignId = user.CampaignID || user.campaignId || '';

      return campaignName.toLowerCase().includes('independence') ||
        campaignName.toLowerCase().includes('insurance') ||
        userCampaignId === 'independence-insurance-agency' ||
        userCampaignId.toLowerCase().includes('independence');
    });

    return users
      .map(user => ({
        name: user.FullName || user.UserName || user.name,
        email: user.Email || user.email,
        id: user.ID || user.id
      }))
      .filter(user => user.name && user.email)
      .sort((a, b) => a.name.localeCompare(b.name));

  } catch (error) {
    console.error('Error getting Independence users:', error);
    writeError('getIndependenceInsuranceUsers', error);
    return [];
  }
}

function clientGetCreditSuiteUsers(campaignId = null) {
  try {
    return getCreditSuiteUsers(campaignId);
  } catch (error) {
    console.error('Error in clientGetCreditSuiteUsers:', error);
    writeError('clientGetCreditSuiteUsers', error);
    return [];
  }
}

function getCreditSuiteUsers(campaignId = null) {
  try {
    let users;

    if (campaignId) {
      users = getUsersByCampaign(campaignId);
    } else {
      const allUsers = getAllUsers();
      users = allUsers.filter(user => {
        const campaignName = user.CampaignName || user.campaignName || '';
        const userCampaignId = user.CampaignID || user.campaignId || '';

        return campaignName.toLowerCase().includes('credit') ||
          campaignName.toLowerCase().includes('suite') ||
          userCampaignId === 'credit-suite' ||
          userCampaignId.toLowerCase().includes('credit');
      });
    }

    return users
      .map(user => ({
        name: user.FullName || user.UserName || user.name,
        email: user.Email || user.email,
        id: user.ID || user.id
      }))
      .filter(user => user.name && user.email)
      .sort((a, b) => a.name.localeCompare(b.name));

  } catch (error) {
    console.error('Error getting Credit Suite users:', error);
    writeError('getCreditSuiteUsers', error);
    return [];
  }
}

function clientGetIBTRUsers() {
  try {
    const allUsers = getAllUsers();
    const users = allUsers.filter(user => {
      const campaignName = user.CampaignName || user.campaignName || '';
      const userCampaignId = user.CampaignID || user.campaignId || '';

      return campaignName.toLowerCase().includes('benefits resource center') ||
        campaignName.toLowerCase().includes('ibtr') ||
        userCampaignId === 'ibtr' ||
        userCampaignId.toLowerCase().includes('benefits');
    });

    return users
      .map(user => ({
        name: user.FullName || user.UserName || user.name,
        email: user.Email || user.email,
        id: user.ID || user.id
      }))
      .filter(user => user.name && user.email)
      .sort((a, b) => a.name.localeCompare(b.name));

  } catch (error) {
    console.error('Error getting IBTR users:', error);
    writeError('clientGetIBTRUsers', error);
    return [];
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// CAMPAIGN AND NAVIGATION FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

function getAllCampaigns() {
  try {
    var campaigns;
    if (typeof clientGetAllCampaigns === 'function') {
      campaigns = clientGetAllCampaigns();
    } else if (typeof readSheet === 'function') {
      campaigns = readSheet('CAMPAIGNS') || [];
    } else {
      campaigns = [];
    }

    var currentUser = null;
    try {
      if (typeof getCurrentUser === 'function') {
        currentUser = getCurrentUser();
      }
    } catch (err) {
      console.warn('getAllCampaigns: failed to hydrate current user', err);
    }

    if (currentUser && currentUser.ID && typeof TenantSecurity !== 'undefined' && TenantSecurity) {
      try {
        return TenantSecurity.filterCampaignList(currentUser.ID, campaigns);
      } catch (filterErr) {
        console.warn('getAllCampaigns: tenant filter error', filterErr);
      }
    }

    return campaigns;
  } catch (error) {
    console.error('Error getting all campaigns:', error);
    return [];
  }
}

function getCampaignById(campaignId) {
  try {
    const campaigns = getAllCampaigns();
    return campaigns.find(campaign =>
      String(campaign.ID || campaign.id) === String(campaignId)
    );
  } catch (error) {
    console.error('Error getting campaign by ID:', error);
    writeError('getCampaignById', error);
    return null;
  }
}

function getCampaignNavigation(campaignId) {
  try {
    if (!campaignId) {
      return { categories: [], uncategorizedPages: [] };
    }

    const cache = (typeof CacheService !== 'undefined' && CacheService)
      ? CacheService.getScriptCache()
      : null;
    const cacheTtl = (typeof CACHE_TTL_SEC !== 'undefined') ? CACHE_TTL_SEC : 300;
    const cacheKey = `NAVIGATION_${campaignId}`;

    if (cache) {
      try {
        const cached = cache.get(cacheKey);
        if (cached) {
          return JSON.parse(cached);
        }
      } catch (cacheErr) {
        console.warn('getCampaignNavigation: cache read failed', cacheErr);
      }
    }

    const pages = (typeof getCampaignPages === 'function')
      ? getCampaignPages(campaignId)
      : (function () {
          if (typeof readSheet !== 'function') return [];
          const sheetName = (typeof CAMPAIGN_PAGES_SHEET !== 'undefined') ? CAMPAIGN_PAGES_SHEET : 'CAMPAIGN_PAGES';
          const allPages = readSheet(sheetName) || [];
          return allPages.filter(p => String(p.CampaignID || p.campaignId) === String(campaignId));
        })();

    const categories = (typeof getCampaignPageCategories === 'function')
      ? getCampaignPageCategories(campaignId)
      : (function () {
          if (typeof readSheet !== 'function') return [];
          const sheetName = (typeof PAGE_CATEGORIES_SHEET !== 'undefined') ? PAGE_CATEGORIES_SHEET : 'PAGE_CATEGORIES';
          const allCategories = readSheet(sheetName) || [];
          return allCategories
            .filter(cat => String(cat.CampaignID || cat.campaignId) === String(campaignId))
            .map(cat => ({
              ID: cat.ID || cat.id,
              CampaignID: cat.CampaignID || cat.campaignId,
              CategoryName: cat.CategoryName || cat.categoryName || cat.Category || cat.category,
              CategoryIcon: cat.CategoryIcon || cat.categoryIcon || 'fas fa-folder',
              SortOrder: cat.SortOrder || cat.sortOrder || 999
            }))
            .sort((a, b) => (a.SortOrder || 999) - (b.SortOrder || 999));
        })();

    const navigation = { categories: [], uncategorizedPages: [] };

    if (!Array.isArray(pages) || pages.length === 0) {
      if (cache) {
        try { cache.put(cacheKey, JSON.stringify(navigation), cacheTtl); } catch (_) {}
      }
      return navigation;
    }

    const categoryMap = {};
    categories.forEach(cat => {
      if (!cat || !cat.ID) return;
      categoryMap[cat.ID] = Object.assign({}, cat, { pages: [] });
    });

    pages.forEach(page => {
      if (!page) return;
      const normalizedPage = Object.assign({}, page, {
        PageIcon: page.PageIcon || page.pageIcon || 'fas fa-file'
      });
      const categoryId = page.CategoryID || page.categoryId;
      if (categoryId && categoryMap[categoryId]) {
        categoryMap[categoryId].pages.push(normalizedPage);
      } else {
        navigation.uncategorizedPages.push(normalizedPage);
      }
    });

    navigation.categories = Object.values(categoryMap).filter(cat => cat.pages && cat.pages.length > 0);

    if (cache) {
      try {
        cache.put(cacheKey, JSON.stringify(navigation), cacheTtl);
      } catch (cachePutErr) {
        console.warn('getCampaignNavigation: cache write failed', cachePutErr);
      }
    }

    return navigation;
  } catch (error) {
    console.error('Error getting campaign navigation:', error);
    writeError('getCampaignNavigation', error);
    return { categories: [], uncategorizedPages: [] };
  }
}

function getUserCampaignPermissions(userId) {
  try {
    if (typeof readSheet === 'function') {
      const permissions = readSheet('CAMPAIGN_USER_PERMISSIONS') || [];
      return permissions.filter(perm =>
        String(perm.UserID || perm.userId) === String(userId)
      );
    }
    return [];
  } catch (error) {
    console.error('Error getting user campaign permissions:', error);
    writeError('getUserCampaignPermissions', error);
    return [];
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// ROLES AND PERMISSIONS
// ───────────────────────────────────────────────────────────────────────────────

function getAllRoles() {
  try {
    if (typeof readSheet === 'function') {
      return readSheet('Roles') || [];
    } else {
      return [];
    }
  } catch (error) {
    console.error('Error getting all roles:', error);
    return [];
  }
}

function getRolesMapping() {
  try {
    let roles = [];
    if (typeof getAllRoles === 'function') {
      roles = getAllRoles();
    } else if (typeof readSheet === 'function') {
      roles = readSheet('Roles') || [];
    }

    const mapping = {};
    roles.forEach(role => {
      mapping[role.ID || role.id] = role.Name || role.name;
    });
    return mapping;
  } catch (error) {
    console.error('Error getting roles mapping:', error);
    return {};
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// PAGE ACCESS AND METADATA
// ───────────────────────────────────────────────────────────────────────────────

function _loadAllPagesMeta() {
  const CK = 'ACCESS_PAGES_META_V1';
  const cached = _cacheGet(CK);
  if (cached) return cached;

  let rows = [];
  try {
    rows = readSheet('Pages') || [];
  } catch (_) {
  }

  const map = {};
  rows.forEach(r => {
    const key = _normalizePageKey(r.PageKey || r.key);
    if (!key) return;
    map[key] = {
      key,
      title: r.PageTitle || r.Name || '',
      requiresAdmin: _truthy(r.RequiresAdmin),
      campaignSpecific: _truthy(r.CampaignSpecific) || _truthy(r.IsCampaignSpecific) || false,
      active: (r.Active === undefined) ? true : _truthy(r.Active)
    };
  });
  _cachePut(CK, map, 180);
  return map;
}

function _getPageMeta(pageKey) {
  const key = _normalizePageKey(pageKey);
  const meta = _loadAllPagesMeta()[key];
  return meta || { key, title: pageKey, requiresAdmin: false, campaignSpecific: false, active: true };
}

// ───────────────────────────────────────────────────────────────────────────────
// SPECIAL PAGE HANDLERS (Updated for Clean URLs)
// ───────────────────────────────────────────────────────────────────────────────

function serveAckForm(e, baseUrl, user) {
  try {
    const id = (e.parameter.id || "").toString();
    const tpl = HtmlService.createTemplateFromFile('CoachingAckForm');

    tpl.baseUrl = baseUrl;
    tpl.scriptUrl = SCRIPT_URL;
    tpl.currentPage = 'Acknowledge Coaching';
    tpl.recordId = id;

    if (id && typeof getCoachingRecordById === 'function') {
      tpl.record = JSON.stringify(getCoachingRecordById(id)).replace(/<\/script>/g, '<\\/script>');
    } else {
      tpl.record = '{}';
    }

    tpl.user = user;
    __injectCurrentUser_(tpl, user);

    return tpl.evaluate()
      .setTitle('Acknowledge Coaching')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('Error serving ack form:', error);
    return createErrorPage('Form Error', error.message);
  }
}

function serveAgentSchedulePage(e, baseUrl, token) {
  try {
    const agentData = validateAgentToken(token);

    if (!agentData.success) {
      return createErrorPage('Invalid Token', 'Your schedule link has expired or is invalid.');
    }

    const tpl = HtmlService.createTemplateFromFile('AgentSchedule');
    tpl.baseUrl = baseUrl;
    tpl.token = token;
    tpl.agentData = JSON.stringify(agentData).replace(/<\/script>/g, '<\\/script>');

    return tpl.evaluate()
      .setTitle('My Schedule - VLBPO')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('Error serving agent schedule page:', error);
    return createErrorPage('Schedule Error', error.message);
  }
}

function serveShiftSlotManagement(e, baseUrl, user, campaignId) {
  try {
    if (!isUserAdmin(user) && !hasPermission(user, 'manage_shifts')) {
      return renderAccessDenied('You do not have permission to manage shift slots.');
    }

    const tpl = HtmlService.createTemplateFromFile('SlotManagementInterface');
    tpl.baseUrl = baseUrl;
    tpl.scriptUrl = SCRIPT_URL;
    tpl.currentPage = 'Shift Slot Management';
    tpl.user = user;
    tpl.campaignId = campaignId;

    __injectCurrentUser_(tpl, user);

    if (campaignId) {
      tpl.campaignName = user.campaignName || '';
    }

    return tpl.evaluate()
      .setTitle('Shift Slot Management - VLBPO LuminaHQ')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    console.error('Error serving shift slot management page:', error);
    writeError('serveShiftSlotManagement', error);
    return createErrorPage('Failed to load page', error.message);
  }
}

function hasPermission(user, permission) {
  try {
    const userRoles = user.Roles || user.roles || '';

    if (isUserAdmin(user)) return true;

    switch (permission) {
      case 'manage_shifts':
        return userRoles.toLowerCase().includes('manager') ||
          userRoles.toLowerCase().includes('supervisor') ||
          userRoles.toLowerCase().includes('scheduler');
      default:
        return false;
    }
  } catch (error) {
    console.error('Error checking permission:', error);
    return false;
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// ENHANCED PROXY SERVICES
// ───────────────────────────────────────────────────────────────────────────────

function serveProxy(e) {
  try {
    return serveEnhancedProxy(e);
  } catch (error) {
    console.error('Proxy service error:', error);
    writeError('serveProxy', error);
    return serveBasicProxy(e);
  }
}

function serveEnhancedProxy(e) {
  try {
    var target = e.parameter.url || '';
    if (!target) {
      return HtmlService.createHtmlOutput('<h3>Proxy Error</h3><p>Missing url</p>')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    if (!/^https?:\/\//i.test(target)) target = 'https://' + target;

    var resp = UrlFetchApp.fetch(target, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
    });
    var ct = String(resp.getHeaders()['Content-Type'] || '').toLowerCase();
    var body = resp.getContentText();

    if (ct.indexOf('text/html') === -1) {
      return ContentService.createTextOutput(body);
    }

    body = body
      .replace(/<meta[^>]+http-equiv=["']?content-security-policy["']?[^>]*>/gi, '')
      .replace(/<script[^>]*>[\s\S]*?top\.location(?:.|\s)*?<\/script>/gi, '')
      .replace(/<script[^>]*>[\s\S]*?if\s*\(\s*top\s*[!=]=?\s*self\s*\)[\s\S]*?<\/script>/gi, '');

    var base = target.replace(/([?#].*)$/, '');
    if (body.indexOf('<base ') === -1) {
      body = body.replace(/<head([^>]*)>/i, function (m, attrs) {
        return '<head' + attrs + '><base href="' + base + '">';
      });
    }

    body = body.replace('</head>', '<style>*,*:before,*:after{box-sizing:border-box}</style></head>');

    return HtmlService.createHtmlOutput(body)
      .setTitle('LuminaHQ Browser')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    writeError && writeError('serveEnhancedProxy', err);
    return HtmlService.createHtmlOutput('<h3>Proxy Error</h3><pre>' + String(err) + '</pre>')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function serveBasicProxy(e) {
  try {
    const target = e.parameter.url;
    if (!target) {
      return HtmlService.createHtmlOutput(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>Proxy Error</title>
            <style>
                body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
                .error { color: #e74c3c; }
            </style>
        </head>
        <body>
            <h1 class="error">Proxy Error</h1>
            <p>Missing URL parameter</p>
            <button onclick="history.back()">Go Back</button>
        </body>
        </html>
      `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    let normalizedUrl = target;
    if (!normalizedUrl.startsWith('http')) {
      normalizedUrl = 'https://' + normalizedUrl;
    }

    const resp = UrlFetchApp.fetch(normalizedUrl, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });

    const content = resp.getContentText();
    const contentType = resp.getHeaders()['Content-Type'] || '';

    if (contentType.toLowerCase().includes('text/html')) {
      let processedContent = content;

      processedContent = processedContent.replace(
        /<script[^>]*>[\s\S]*?if\s*\(\s*top\s*[!=]=?\s*self\s*\)[\s\S]*?<\/script>/gi,
        ''
      );
      processedContent = processedContent.replace(
        /<script[^>]*>[\s\S]*?top\.location[\s\S]*?<\/script>/gi,
        ''
      );

      const enhancedContent = processedContent.replace(
        '</head>',
        `<style>
          body { margin-top: 0 !important; }
          * { box-sizing: border-box; }
        </style></head>`
      );

      return HtmlService.createHtmlOutput(enhancedContent)
        .setTitle('LuminaHQ Browser')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } else {
      return ContentService.createTextOutput(content);
    }

  } catch (error) {
    console.error('Basic proxy error:', error);
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
      <head>
          <title>Proxy Error</title>
          <style>
              body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
              .error { color: #e74c3c; }
          </style>
      </head>
      <body>
          <h1 class="error">Unable to Load Page</h1>
          <p>Error: ${error.message}</p>
          <button onclick="history.back()">Go Back</button>
          <button onclick="location.reload()">Retry</button>
      </body>
      </html>
    `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// UTILITY FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

function createErrorPage(title, message) {
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <title>${title}</title>
      <meta name="viewport" content="width=device-width,initial-scale=1">
      <style>
        body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
        .error { color: #e74c3c; }
        .message { margin: 20px 0; }
        .back-link { color: #3498db; text-decoration: none; }
      </style>
    </head>
    <body>
      <h1 class="error">${title}</h1>
      <p class="message">${message}</p>
      <a href="#" onclick="history.back()" class="back-link">← Go Back</a>
    </body>
    </html>
  `;

  return HtmlService.createHtmlOutput(html)
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function writeError(functionName, error) {
  try {
    const errorMessage = error && error.message ? error.message : error.toString();
    console.error(`[${functionName}] ${errorMessage}`);

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let errorSheet = ss.getSheetByName('ErrorLog');

      if (!errorSheet) {
        errorSheet = ss.insertSheet('ErrorLog');
        errorSheet.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Function', 'Error', 'Stack']]);
      }

      const timestamp = new Date();
      const stack = error && error.stack ? error.stack : 'No stack trace';

      errorSheet.appendRow([timestamp, functionName, errorMessage, stack]);
    } catch (logError) {
      console.error('Failed to log error to sheet:', logError);
    }
  } catch (e) {
    console.error('Error in writeError function:', e);
  }
}

function writeDebug(message) {
  console.log(`[DEBUG] ${message}`);
}

function confirmEmail(token) {
  try {
    if (!token) {
      console.warn('confirmEmail called with empty token');
      return false;
    }

    if (typeof IdentityService !== 'undefined'
      && IdentityService
      && typeof IdentityService.confirmEmail === 'function') {
      const result = IdentityService.confirmEmail(token);
      if (result && result.success) {
        return true;
      }
      console.warn('confirmEmail: IdentityService returned failure', result);
      return false;
    }

    console.warn('confirmEmail: IdentityService unavailable; skipping confirmation');
    return false;
  } catch (error) {
    console.error('Error confirming email:', error);
    writeError('confirmEmail', error);
    return false;
  }
}

function weekStringFromDate(date) {
  try {
    if (!date || !(date instanceof Date)) {
      date = new Date();
    }

    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    const weekNum = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    return `${d.getUTCFullYear()}-W${weekNum.toString().padStart(2, '0')}`;
  } catch (error) {
    console.error('Error generating week string:', error);
    return new Date().getFullYear() + '-W01';
  }
}

function searchWeb(query, startIndex) {
  try {
    if (!query || typeof query !== 'string') {
      throw new Error('Invalid search query.');
    }
    const CSE_ID = '130aba31c8a2d439c';
    const API_KEY = 'AIzaSyAg-puM5l9iQpjz_NplMJaKbUNRH7ld7sY';
    const baseUrl = 'https://www.googleapis.com/customsearch/v1';
    const params = [
      `key=${API_KEY}`,
      `cx=${CSE_ID}`,
      `q=${encodeURIComponent(query)}`,
      startIndex ? `start=${startIndex}` : ''
    ].filter(Boolean).join('&');
    const url = `${baseUrl}?${params}`;

    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = resp.getResponseCode();
    const text = resp.getContentText();
    if (code !== 200) {
      throw new Error(`Search API error [${code}]: ${text}`);
    }
    return JSON.parse(text);
  } catch (error) {
    console.error('Error in searchWeb:', error);
    writeError('searchWeb', error);
    throw error;
  }
}

function getEmptyQAAnalytics() {
  return {
    avgScore: 0,
    passRate: 0,
    totalEvaluations: 0,
    agentsEvaluated: 0,
    avgScoreChange: 0,
    passRateChange: 0,
    evaluationsChange: 0,
    agentsChange: 0,
    categories: { labels: [], values: [] },
    trends: { labels: [], values: [] },
    agents: { labels: [], values: [] }
  };
}

// ───────────────────────────────────────────────────────────────────────────────
// SYSTEM INITIALIZATION
// ───────────────────────────────────────────────────────────────────────────────

function initializeSystem() {
  try {
    console.log('Initializing system...');

    if (typeof AuthenticationService !== 'undefined' && AuthenticationService.ensureSheets) {
      AuthenticationService.ensureSheets();
    }

    try {
      ensureSessionCleanupTrigger();
    } catch (triggerError) {
      console.warn('initializeSystem: unable to ensure session cleanup trigger', triggerError);
    }

    initializeMainSheets();
    initializeCampaignSystems();

    if (typeof CallCenterWorkflowService !== 'undefined' && CallCenterWorkflowService.initialize) {
      try {
        CallCenterWorkflowService.initialize();
      } catch (err) {
        console.warn('CallCenterWorkflowService initialization failed', err);
      }
    }

    console.log('System initialization completed successfully');
    return { success: true, message: 'System initialized' };

  } catch (error) {
    console.error('System initialization failed:', error);
    writeError('initializeSystem', error);
    return { success: false, error: error.message };
  }
}

function ensureSessionCleanupTrigger() {
  if (typeof ScriptApp === 'undefined' || !ScriptApp || typeof ScriptApp.getProjectTriggers !== 'function') {
    console.warn('ensureSessionCleanupTrigger: ScriptApp not available; skipping trigger registration.');
    return false;
  }

  var handlerName = 'cleanupExpiredSessionsJob';

  try {
    var triggers = ScriptApp.getProjectTriggers();
    var hasTrigger = triggers.some(function (trigger) {
      return trigger && typeof trigger.getHandlerFunction === 'function'
        ? trigger.getHandlerFunction() === handlerName
        : false;
    });

    if (!hasTrigger) {
      ScriptApp.newTrigger(handlerName)
        .timeBased()
        .everyHours(1)
        .create();
      console.log('ensureSessionCleanupTrigger: created hourly trigger for cleanupExpiredSessionsJob');
    }

    return true;
  } catch (error) {
    console.error('ensureSessionCleanupTrigger: failed to ensure trigger', error);
    if (typeof writeError === 'function') {
      try { writeError('ensureSessionCleanupTrigger', error); } catch (loggingError) { console.warn('ensureSessionCleanupTrigger: unable to log error', loggingError); }
    }
    return false;
  }
}

function initializeMainSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requiredSheets = ['Users', 'Roles', 'Pages', 'CAMPAIGNS'];

    requiredSheets.forEach(sheetName => {
      try {
        let sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
          console.log(`Creating missing sheet: ${sheetName}`);
          sheet = ss.insertSheet(sheetName);
          initializeSheetHeaders(sheet, sheetName);
        }
      } catch (sheetError) {
        console.warn(`Could not initialize sheet ${sheetName}:`, sheetError);
      }
    });

    console.log('Main sheets initialized');

  } catch (error) {
    console.error('Error initializing main sheets:', error);
    writeError('initializeMainSheets', error);
  }
}

function initializeSheetHeaders(sheet, sheetName) {
  try {
    let headers = [];

    switch (sheetName) {
      case 'Users':
        headers = ['ID', 'Email', 'FullName', 'Password', 'Roles', 'CampaignID', 'EmailConfirmed', 'EmailConfirmation', 'CreatedAt', 'UpdatedAt'];
        break;
      case 'Roles':
        headers = ['ID', 'Name', 'Description', 'Permissions', 'CreatedAt'];
        break;
      case 'Pages':
        headers = ['PageKey', 'Name', 'Description', 'RequiredRole', 'CampaignSpecific', 'Active'];
        break;
      case 'CAMPAIGNS':
        headers = ['ID', 'Name', 'Description', 'Active', 'Settings', 'CreatedAt'];
        break;
      default:
        console.warn(`No headers defined for sheet: ${sheetName}`);
        return;
    }

    if (headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      console.log(`Headers set for ${sheetName}:`, headers);
    }

  } catch (error) {
    console.error(`Error setting headers for ${sheetName}:`, error);
  }
}

function initializeCampaignSystems() {
  try {
    if (typeof initializeIndependenceQASystem === 'function') {
      try {
        initializeIndependenceQASystem();
        console.log('Independence QA system initialized');
      } catch (error) {
        console.warn('Independence QA system initialization failed:', error);
      }
    }

    if (typeof initializeCreditSuiteQASystem === 'function') {
      try {
        initializeCreditSuiteQASystem();
        console.log('Credit Suite QA system initialized');
      } catch (error) {
        console.warn('Credit Suite QA system initialization failed:', error);
      }
    }

    if (typeof readSheet === 'function') {
      const pages = readSheet('Pages');
      if (pages.length === 0) {
        initializeSystemPages();
      }
    }

    console.log('Campaign systems initialization completed');

  } catch (error) {
    console.error('Error initializing campaign systems:', error);
    writeError('initializeCampaignSystems', error);
  }
}

function initializeSystemPages() {
  try {
    const systemPages = [
      {
        PageKey: 'dashboard',
        Name: 'Dashboard',
        Description: 'Main dashboard page',
        RequiredRole: '',
        CampaignSpecific: false,
        Active: true
      },
      {
        PageKey: 'users',
        Name: 'User Management',
        Description: 'Manage system users',
        RequiredRole: 'admin',
        CampaignSpecific: false,
        Active: true
      },
      {
        PageKey: 'roles',
        Name: 'Role Management',
        Description: 'Manage user roles',
        RequiredRole: 'admin',
        CampaignSpecific: false,
        Active: true
      },
      {
        PageKey: 'campaigns',
        Name: 'Campaign Management',
        Description: 'Manage campaigns',
        RequiredRole: 'admin',
        CampaignSpecific: false,
        Active: true
      },
      {
        PageKey: 'tasks',
        Name: 'Task Management',
        Description: 'Task board and management',
        RequiredRole: '',
        CampaignSpecific: true,
        Active: true
      },
      {
        PageKey: 'search',
        Name: 'Search',
        Description: 'Global search functionality',
        RequiredRole: '',
        CampaignSpecific: false,
        Active: true
      },
      {
        PageKey: 'chat',
        Name: 'Chat',
        Description: 'Team communication',
        RequiredRole: '',
        CampaignSpecific: false,
        Active: true
      },
      {
        PageKey: 'notifications',
        Name: 'Notifications',
        Description: 'System notifications',
        RequiredRole: '',
        CampaignSpecific: false,
        Active: true
      }
    ];

    if (typeof writeSheet === 'function') {
      writeSheet('Pages', systemPages);
      console.log('System pages initialized');
    }

  } catch (error) {
    console.error('Error initializing system pages:', error);
    writeError('initializeSystemPages', error);
  }
}

function queueBackgroundInitialization(options) {
  var safeConsole = (typeof console !== 'undefined' && console) ? console : {
    error: function () { },
    warn: function () { },
    log: function () { }
  };

  function logTaskError(label, error) {
    if (safeConsole && typeof safeConsole.error === 'function') {
      safeConsole.error('queueBackgroundInitialization task failed [' + label + ']:', error);
    }
    if (typeof writeError === 'function') {
      try {
        writeError('queueBackgroundInitialization::' + label, error);
      } catch (loggingError) {
        if (safeConsole && typeof safeConsole.error === 'function') {
          safeConsole.error('Failed to log queueBackgroundInitialization error for [' + label + ']:', loggingError);
        }
      }
    }
  }

  function extractContext(opts) {
    if (!opts || typeof opts !== 'object') {
      return null;
    }
    if (opts.context && typeof opts.context === 'object') {
      return opts.context;
    }
    var contextKeys = ['tenantId', 'tenantIds', 'campaignId', 'campaignIds', 'allowAllTenants', 'globalTenantAccess'];
    var context = {};
    var hasContext = false;
    contextKeys.forEach(function (key) {
      if (Object.prototype.hasOwnProperty.call(opts, key)) {
        context[key] = opts[key];
        hasContext = true;
      }
    });
    return hasContext ? context : null;
  }

  try {
    var requestOptions = (options && typeof options === 'object') ? options : {};
    var context = extractContext(requestOptions);
    var manager = (typeof DatabaseManager !== 'undefined' && DatabaseManager && typeof DatabaseManager.table === 'function')
      ? DatabaseManager
      : null;

    var seenEntities = Object.create(null);
    var entityNames = [];
    function addEntity(name) {
      if (!name) return;
      var normalized = String(name).trim().toLowerCase();
      if (!normalized || seenEntities[normalized]) return;
      seenEntities[normalized] = true;
      entityNames.push(normalized);
    }

    if (Array.isArray(requestOptions.entities)) {
      requestOptions.entities.forEach(addEntity);
    }

    if (!entityNames.length) {
      addEntity('quality');
      addEntity('users');
    }

    var tasks = [];

    if (manager && typeof resolveLuminaEntityDefinition === 'function') {
      var warmedTables = Object.create(null);
      entityNames.forEach(function (entityName) {
        try {
          var definition = resolveLuminaEntityDefinition(entityName);
          if (!definition || !definition.tableName || warmedTables[definition.tableName]) {
            return;
          }
          warmedTables[definition.tableName] = true;
          tasks.push({
            label: 'warmEntity:' + (definition.name || entityName),
            run: function () {
              var readOptions = { cache: true, limit: 50 };
              if (Array.isArray(definition.summaryColumns) && definition.summaryColumns.length) {
                readOptions.columns = definition.summaryColumns.slice();
              }
              manager.table(definition.tableName, context).read(readOptions);
            }
          });
        } catch (resolveError) {
          logTaskError('resolveEntity(' + entityName + ')', resolveError);
        }
      });
    }

    if (manager && typeof manager.backfillAllMissingIds === 'function') {
      tasks.push({
        label: 'DatabaseManager.backfillAllMissingIds',
        run: function () {
          var maintenanceContext = Object.assign({ allowAllTenants: true }, context || {});
          var summaries = manager.backfillAllMissingIds(maintenanceContext);
          if (safeConsole && typeof safeConsole.log === 'function') {
            safeConsole.log('DatabaseManager.backfillAllMissingIds summaries:', summaries);
          }
          if (Array.isArray(summaries)) {
            for (var i = 0; i < summaries.length; i++) {
              var summary = summaries[i];
              if (summary && summary.error && safeConsole && typeof safeConsole.error === 'function') {
                safeConsole.error('ID backfill error for table ' + summary.table + ': ' + summary.error);
              }
            }
          }
        }
      });
    }

    if (typeof QualityService !== 'undefined' && QualityService && typeof QualityService.queueBackgroundInitialization === 'function') {
      tasks.push({
        label: 'QualityService.queueBackgroundInitialization',
        run: function () {
          QualityService.queueBackgroundInitialization(context, requestOptions);
        }
      });
    }

    tasks.forEach(function (task) {
      try {
        task.run();
      } catch (taskError) {
        logTaskError(task.label, taskError);
      }
    });

    return true;
  } catch (error) {
    logTaskError('root', error);
    return false;
  }
}

function scheduledWarmup() {
  var safeConsole = (typeof console !== 'undefined' && console) ? console : {
    error: function () { },
    warn: function () { },
    log: function () { }
  };

  try {
    return queueBackgroundInitialization({});
  } catch (error) {
    if (safeConsole && typeof safeConsole.error === 'function') {
      safeConsole.error('scheduledWarmup failed:', error);
    }
    if (typeof writeError === 'function') {
      try {
        writeError('scheduledWarmup', error);
      } catch (loggingError) {
        if (safeConsole && typeof safeConsole.error === 'function') {
          safeConsole.error('Failed to log scheduledWarmup error:', loggingError);
        }
      }
    }
    return false;
  }
}

function runDatabaseIdBackfill() {
  var safeConsole = (typeof console !== 'undefined' && console) ? console : {
    error: function () { },
    warn: function () { },
    log: function () { }
  };

  try {
    if (typeof DatabaseManager === 'undefined' || !DatabaseManager || typeof DatabaseManager.backfillAllMissingIds !== 'function') {
      throw new Error('DatabaseManager.backfillAllMissingIds is not available');
    }
    var summaries = DatabaseManager.backfillAllMissingIds({ allowAllTenants: true });
    if (safeConsole && typeof safeConsole.log === 'function') {
      safeConsole.log('runDatabaseIdBackfill summaries:', summaries);
    }
    return summaries;
  } catch (error) {
    if (safeConsole && typeof safeConsole.error === 'function') {
      safeConsole.error('runDatabaseIdBackfill failed:', error);
    }
    if (typeof writeError === 'function') {
      try {
        writeError('runDatabaseIdBackfill', error);
      } catch (loggingError) {
        if (safeConsole && typeof safeConsole.error === 'function') {
          safeConsole.error('Failed to log runDatabaseIdBackfill error:', loggingError);
        }
      }
    }
    throw error;
  }
}

function ensureScheduledWarmupTrigger() {
  var triggers = [];
  try {
    triggers = ScriptApp.getProjectTriggers();
  } catch (error) {
    var safeConsole = (typeof console !== 'undefined' && console) ? console : {
      error: function () { },
      warn: function () { },
      log: function () { }
    };
    if (safeConsole && typeof safeConsole.error === 'function') {
      safeConsole.error('ensureScheduledWarmupTrigger: unable to list triggers:', error);
    }
    throw error;
  }

  var hasTrigger = triggers.some(function (trigger) {
    if (trigger && typeof trigger.getHandlerFunction === 'function') {
      return trigger.getHandlerFunction() === 'scheduledWarmup';
    }
    return false;
  });

  if (!hasTrigger) {
    ScriptApp.newTrigger('scheduledWarmup')
      .timeBased()
      .everyMinutes(5)
      .create();
  }

  return hasTrigger;
}

// ───────────────────────────────────────────────────────────────────────────────
// GLOBAL FAVICON INJECTOR AND TEMPLATE HELPERS
// ───────────────────────────────────────────────────────────────────────────────

(function () {
  if (HtmlService.__faviconPatched === true) return;
  HtmlService.__faviconPatched = true;

  const _origCreate = HtmlService.createTemplateFromFile;

  HtmlService.createTemplateFromFile = function (file) {
    const tpl = _origCreate.call(HtmlService, file);

    const _origEval = tpl.evaluate;
    tpl.evaluate = function () {
      const out = _origEval.apply(tpl, arguments);
      return (typeof out.setFaviconUrl === 'function') ? out.setFaviconUrl(FAVICON_URL) : out;
    };

    return tpl;
  };
})();

const __INCLUDE_STACK = [];
const __INCLUDED_ONCE = new Set();

function include(file, params) {
  if (__INCLUDE_STACK.indexOf(file) !== -1) {
    throw new RangeError('Cyclic include detected: ' + __INCLUDE_STACK.concat(file).join(' → '));
  }
  if (__INCLUDE_STACK.length > 20) {
    throw new RangeError('Include depth > 20; aborting to avoid stack overflow.');
  }

  __INCLUDE_STACK.push(file);
  try {
    const tpl = HtmlService.createTemplateFromFile(file);
    if (params) Object.keys(params).forEach(k => (tpl[k] = params[k]));

    try {
      if (!tpl.user) tpl.user = getCurrentUser();
      tpl.currentUserJson = _stringifyForTemplate_(tpl.user);
    } catch (_) {
      tpl.user = null;
      tpl.currentUserJson = '{}';
    }

    return tpl.evaluate().getContent();
  } finally {
    __INCLUDE_STACK.pop();
  }
}

function includeOnce(file, params) {
  if (__INCLUDED_ONCE.has(file)) return '';
  __INCLUDED_ONCE.add(file);
  return include(file, params);
}

// ───────────────────────────────────────────────────────────────────────────────
// CAMPAIGN HELPERS FOR CLIENT ACCESS
// ───────────────────────────────────────────────────────────────────────────────

function clientListCasesAndPages() {
  return CASE_DEFS.map(def => ({
    case: def.case,
    aliases: def.aliases,
    idHint: def.idHint,
    pages: Object.keys(def.pages || {})
  }));
}

function clientBuildCaseHref(caseNameOrAlias, pageKind) {
  const base = getBaseUrl();
  const page = String(caseNameOrAlias || '').trim() + '.' + String(pageKind || 'QAForm').trim();
  return base + '?page=' + encodeURIComponent(page);
}

function __routeCase__(caseName, pageKind, e, baseUrl, user, cid) {
  const def = __case_resolve__(caseName);
  if (!def) return createErrorPage('Unknown Campaign', 'No case mapping for: ' + caseName);
  const targetCid = cid || __case_findCampaignId__(def) || (user && user.CampaignID) || '';
  const candidates = __case_templateCandidates__(def, pageKind);
  return __case_serve__(candidates, e, baseUrl, user, targetCid);
}

// Named route wrappers for specific campaigns
function routeCreditSuiteQAForm(e, baseUrl, user, cid) {
  return __routeCase__('CreditSuite', 'QAForm', e, baseUrl, user, cid);
}

function routeIBTRQACollabList(e, baseUrl, user, cid) {
  return __routeCase__('IBTR', 'QACollabList', e, baseUrl, user, cid);
}

function routeTGCChatReport(e, baseUrl, user, cid) {
  return __routeCase__('TGC', 'ChatReport', e, baseUrl, user, cid);
}

function routeIBTRAttendanceReports(e, baseUrl, user, cid) {
  return __routeCase__('IBTR', 'AttendanceReports', e, baseUrl, user, cid);
}

function routeIndependenceQAForm(e, baseUrl, user, cid) {
  return __routeCase__('IndependenceInsuranceAgency', 'QAForm', e, baseUrl, user, cid);
}

function _debugAccess(label, result, user, pageKey, campaignId) {
  try {
    if (!ACCESS_DEBUG) return;
    console.log('[ACCESS] ' + label + ' → allow=' + result.allow + ' reason=' + result.reason +
      ' page=' + pageKey + ' campaign=' + (campaignId || '-') +
      ' user=' + (user && (user.Email || user.UserName || user.ID)));
    (result.trace || []).forEach((t, i) => console.log('   [' + i + '] ' + t));
  } catch (_) {
  }
}

function isUserAdmin(user) {
  return isSystemAdmin(user);
}

function validateAgentToken(token) {
  // Placeholder function - implement based on your token validation system
  try {
    // Add your token validation logic here
    return {
      success: true,
      agentId: 'sample-agent',
      agentName: 'Sample Agent'
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

// Helper function to get QA form users (unified approach)
function getQAFormUsers(campaignId, requestingUserId) {
  try {
    console.log('getQAFormUsers called with campaignId:', campaignId, 'requestingUserId:', requestingUserId);

    // Use the existing user management system
    const users = getUsers();

    return users
      .filter(user => {
        // Filter by campaign if specified
        if (campaignId && user.CampaignID && user.CampaignID !== campaignId) {
          return false;
        }
        return true;
      })
      .map(user => ({
        name: user.FullName || user.UserName,
        email: user.Email,
        id: user.ID
      }))
      .filter(user => user.name && user.email)
      .sort((a, b) => a.name.localeCompare(b.name));
  } catch (error) {
    console.error('Error in getQAFormUsers:', error);
    return [];
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// FINAL INITIALIZATION LOG
// ───────────────────────────────────────────────────────────────────────────────

console.log('Enhanced Multi-Campaign Code.gs with Simplified Authentication loaded successfully');
console.log('Features: Token-based authentication, Campaign-aware routing, Enhanced access control');
console.log('Base URL:', SCRIPT_URL);
console.log('Supported Campaigns:', CASE_DEFS.map(def => def.case).join(', '));
