/**
 * CampaignResourceRegistry
 * -------------------------------------------------------------
 * Shared lookup helpers that map each campaign to its configured
 * Utilities and Schedule spreadsheets. Configuration is read from
 * the main Campaigns sheet so adding a campaign is a configuration
 * change instead of a code change.
 *
 * Expected Campaigns columns (extend the existing sheet rather than
 * replacing it):
 * - UtilitiesFileId / UtilitiesSpreadsheetId
 * - ScheduleFileId / ScheduleSpreadsheetId
 * - CampaignType (call/chat/email/mixed) or Channel fallback
 * - Optional derived sheet naming pattern: "<Campaign Name> - Escalations", etc.
 *
 * The helpers stay backward-compatible with the legacy iBTR setup by
 * falling back to CAMPAIGN_SPREADSHEET_ID and SCHEDULE_SPREADSHEET_ID
 * when campaign-specific IDs are missing.
 *
 * Adding a new campaign requires only updating the Campaigns sheet with
 * the Utilities and Schedule spreadsheet IDs (no code migrations). The
 * registry will attempt campaign-prefixed sheet names before using the
 * legacy defaults, so utilities can live side-by-side without mixing data.
 */
(function () {
  const root = typeof globalThis === 'object' ? globalThis : (function () { return this; })();

  if (root.CampaignResourceRegistry) {
    return;
  }

  // Static catalog of campaigns to keep utilities/schedule resolution predictable
  // without performing any migrations or sheet creation. The main Campaigns sheet
  // remains the source of truth; these entries only supply defaults for names and
  // channel types when configuration rows are missing.
  var DEFAULT_CAMPAIGN_CATALOG = [
    { id: 'Credit Suite', name: 'Credit Suite', channel: 'call', campaignType: 'call' },
    { id: 'HiyaCar', name: 'HiyaCar', channel: 'call', campaignType: 'call' },
    { id: 'iBTR', name: 'Benefits Resource Center (iBTR)', channel: 'call', campaignType: 'call' },
    { id: 'Independence Insurance Agency', name: 'Independence Insurance Agency', channel: 'call', campaignType: 'call' },
    { id: 'JSC', name: 'JSC', channel: 'call', campaignType: 'call' },
    { id: 'Kids in the Game', name: 'Kids in the Game', channel: 'call', campaignType: 'call' },
    { id: 'Kofi Group', name: 'Kofi Group', channel: 'call', campaignType: 'call' },
    { id: 'PAW LAW FIRM', name: 'PAW LAW FIRM', channel: 'call', campaignType: 'call' },
    { id: 'Pro House Photos', name: 'Pro House Photos', channel: 'call', campaignType: 'call' },
    { id: 'Proozy', name: 'Proozy', channel: 'call', campaignType: 'call' },
    { id: 'The Grounding CO', name: 'The Grounding CO', channel: 'email', campaignType: 'email' },
    { id: 'Lumina HQ', name: 'Lumina HQ', channel: 'mixed', campaignType: 'mixed' }
  ];

  function normalize(value) {
    if (value === null || value === undefined) return '';
    return String(value).trim();
  }

  function normalizeId(value) {
    return normalize(value).toLowerCase();
  }

  function toCampaignKey(record) {
    return normalizeId(record && (record.ID || record.Id || record.id));
  }

  function safeLogError(where, error) {
    try { console.error('[CampaignResourceRegistry] ' + where, error); } catch (_) { }
    if (typeof safeWriteError === 'function') {
      try { safeWriteError(where, error); } catch (_) { }
    }
  }

  function safeWarn(where, message) {
    try { console.warn('[CampaignResourceRegistry] ' + where + ': ' + message); } catch (_) { }
  }

  function readCampaignRows() {
    try {
      if (typeof dbSelect === 'function') {
        const rows = dbSelect(CAMPAIGNS_SHEET || 'Campaigns');
        if (Array.isArray(rows)) return rows;
      }
    } catch (error) {
      safeLogError('dbSelect(Campaigns)', error);
    }

    try {
      if (typeof readSheet === 'function') {
        const rows = readSheet(CAMPAIGNS_SHEET || 'Campaigns');
        if (Array.isArray(rows)) return rows;
      }
    } catch (error) {
      safeLogError('readSheet(Campaigns)', error);
    }

    return [];
  }

  function parseCampaign(row) {
    const id = normalize(row.ID || row.Id || row.id);
    const key = normalizeId(id);
    if (!key) return null;

    const campaign = {
      id,
      key,
      name: normalize(row.Name || row.name),
      description: normalize(row.Description || row.description),
      client: normalize(row.ClientName || row.clientName || row.Client || row.client),
      status: normalize(row.Status || row.status),
      timezone: normalize(row.Timezone || row.timezone || row.TimeZone),
      channel: normalize(row.Channel || row.channel || row.CampaignType),
      utilitiesFileId: normalize(row.UtilitiesFileId || row.UtilitiesSpreadsheetId || row.UtilitiesFileID),
      scheduleFileId: normalize(row.ScheduleFileId || row.ScheduleSpreadsheetId || row.ScheduleFileID),
      campaignType: normalize(row.CampaignType || row.Channel || row.channel),
      isCatalogOnly: false,
      source: 'sheet'
    };

    return campaign;
  }

  function normalizeCampaignRecord(record) {
    if (!record) return null;
    var normalized = {
      id: normalize(record.id || record.ID || record.Id || record.key || record.name),
      name: normalize(record.name || record.Name),
      description: normalize(record.description || record.Description),
      client: normalize(record.client || record.Client || record.ClientName),
      status: normalize(record.status || record.Status),
      timezone: normalize(record.timezone || record.Timezone || record.TimeZone),
      channel: normalize(record.channel || record.Channel || record.campaignType || record.CampaignType),
      utilitiesFileId: normalize(record.utilitiesFileId || record.UtilitiesFileId || record.UtilitiesSpreadsheetId || record.UtilitiesFileID),
      scheduleFileId: normalize(record.scheduleFileId || record.ScheduleFileId || record.ScheduleSpreadsheetId || record.ScheduleFileID),
      campaignType: normalize(record.campaignType || record.CampaignType || record.channel || record.Channel),
      isCatalogOnly: record.isCatalogOnly === true,
      source: normalize(record.source || 'catalog')
    };

    normalized.key = normalizeId(normalized.id || normalized.name);
    if (!normalized.id && normalized.name) {
      normalized.id = normalized.name;
    }

    return normalized.key ? normalized : null;
  }

  function mergeCampaignConfigs(base, override) {
    var merged = {};
    var fields = ['id', 'key', 'name', 'description', 'client', 'status', 'timezone', 'channel', 'utilitiesFileId', 'scheduleFileId', 'campaignType'];
    fields.forEach(function (field) {
      var value = (override && override[field]) ? override[field] : (base && base[field]);
      merged[field] = normalize(value);
    });
    merged.key = normalizeId(merged.id || merged.key || merged.name);
    if (!merged.id && merged.name) merged.id = merged.name;
    if (!merged.campaignType && merged.channel) merged.campaignType = merged.channel;
    merged.isCatalogOnly = (typeof (override && override.isCatalogOnly) === 'boolean')
      ? override.isCatalogOnly
      : !!(base && base.isCatalogOnly);
    merged.source = normalize((override && override.source) || (base && base.source) || '');
    return merged.key ? merged : null;
  }

  function buildDefaultCampaignIndex() {
    var index = new Map();
    DEFAULT_CAMPAIGN_CATALOG.forEach(function (entry) {
      var normalized = normalizeCampaignRecord(Object.assign({ isCatalogOnly: true, source: 'catalog' }, entry));
      if (normalized && normalized.key) {
        index.set(normalized.key, normalized);
      }
    });
    return index;
  }

  function buildCampaignIndex() {
    const rows = readCampaignRows();
    const index = buildDefaultCampaignIndex();

    rows.forEach(function (row) {
      const parsed = parseCampaign(row || {});
      if (!parsed) return;
      const existing = index.get(parsed.key);
      const merged = mergeCampaignConfigs(existing, parsed);
      if (merged) {
        index.set(merged.key, merged);
      }
    });

    return index;
  }

  function getCampaignConfig(identifier) {
    const index = buildCampaignIndex();
    const key = normalizeId(identifier);
    if (key && index.has(key)) {
      return index.get(key);
    }

    // Fallback by name for convenience
    const nameKey = normalizeId(identifier && identifier.Name ? identifier.Name : identifier);
    if (nameKey) {
      for (const value of index.values()) {
        if (normalizeId(value.name) === nameKey) return value;
      }
    }

    // Provide backwards-compatible IBTR defaults
    if (key === 'ibtr' || nameKey === 'benefits resource center' || nameKey === 'ibtr') {
      return {
        id: identifier,
        key: key || nameKey || 'ibtr',
        name: 'Benefits Resource Center (iBTR)',
        utilitiesFileId: (typeof CAMPAIGN_SPREADSHEET_ID !== 'undefined') ? normalize(CAMPAIGN_SPREADSHEET_ID) : '',
        scheduleFileId: (typeof SCHEDULE_SPREADSHEET_ID !== 'undefined') ? normalize(SCHEDULE_SPREADSHEET_ID) : '',
        campaignType: 'call',
        isCatalogOnly: false,
        source: 'legacy'
      };
    }

    return null;
  }

  function getCampaignConfigs() {
    const index = buildCampaignIndex();
    return Array.from(index.values());
  }

  function resolveSpreadsheetById(spreadsheetId) {
    if (!spreadsheetId) return null;
    try {
      return SpreadsheetApp.openById(spreadsheetId);
    } catch (error) {
      safeLogError('openById(' + spreadsheetId + ')', error);
      return null;
    }
  }

  function getCampaignUtilitiesSpreadsheet(identifier) {
    const config = getCampaignConfig(identifier);
    if (config && config.utilitiesFileId) {
      const ss = resolveSpreadsheetById(config.utilitiesFileId);
      if (ss) return ss;
    }

    // Legacy fallback: reuse IBTR helper
    if (typeof getIBTRSpreadsheet === 'function') {
      try { return getIBTRSpreadsheet(); } catch (error) { safeLogError('getIBTRSpreadsheet()', error); }
    }

    // Ultimate fallback: active spreadsheet
    try { return SpreadsheetApp.getActiveSpreadsheet(); } catch (error) { safeLogError('activeSpreadsheet', error); }
    return null;
  }

  function getCampaignScheduleSpreadsheet(identifier) {
    const config = getCampaignConfig(identifier);
    if (config && config.scheduleFileId) {
      const ss = resolveSpreadsheetById(config.scheduleFileId);
      if (ss) return ss;
    }

    // Reuse ScheduleUtilities fallback when available
    if (typeof getScheduleSpreadsheet === 'function') {
      try { return getScheduleSpreadsheet(); } catch (error) { safeLogError('getScheduleSpreadsheet()', error); }
    }

    // Legacy IBTR schedule fallback
    if (typeof getIBTRSpreadsheet === 'function') {
      try { return getIBTRSpreadsheet(); } catch (error) { safeLogError('getIBTRSpreadsheet()', error); }
    }

    try { return SpreadsheetApp.getActiveSpreadsheet(); } catch (error) { safeLogError('activeSpreadsheet', error); }
    return null;
  }

  function ensureSheet(ss, sheetName) {
    if (!ss || !sheetName) return null;
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) return sheet;
    safeWarn('ensureSheet', 'Missing sheet "' + sheetName + '" in ' + ss.getName());
    return null;
  }

  function getCampaignName(config) {
    if (!config) return '';
    return normalize(config.name || config.Name || config.id || config.Id || config.ID);
  }

  function buildSheetNameCandidates(config) {
    const baseName = getCampaignName(config);
    const candidates = function (label, globalFallback) {
      const names = [];
      if (baseName) {
        names.push(baseName + ' - ' + label);
      }
      if (globalFallback) {
        names.push(globalFallback);
      }
      return names.filter(Boolean);
    };

    const channel = normalize(config && (config.campaignType || config.channel));
    const reportPrefersChat = channel === 'chat' || channel === 'email';

    return {
      escalations: candidates('Escalations', (typeof ESCALATIONS_SHEET !== 'undefined' && ESCALATIONS_SHEET) || 'Escalations'),
      tasks: candidates('Tasks', (typeof TASKS_SHEET !== 'undefined' && TASKS_SHEET) || 'Tasks'),
      quality: candidates('Quality', (typeof G !== 'undefined' && G.QA_RECORDS) || 'Quality'),
      callReports: candidates(reportPrefersChat ? 'Call Reports (Voice Optional)' : 'Call Reports', (typeof CALL_REPORT !== 'undefined' && CALL_REPORT) || 'CallReport'),
      chatReports: candidates('Chat Reports', 'ChatReport'),
      attendanceLog: candidates('Attendance Log', (typeof ATTENDANCE_SHEET_NAME !== 'undefined' && ATTENDANCE_SHEET_NAME) || (typeof ATTENDANCE !== 'undefined' && ATTENDANCE) || 'AttendanceLog')
    };
  }

  function getCampaignUtilitySheetNames(campaign) {
    return buildSheetNameCandidates(campaign || getCampaignConfig(campaign));
  }

  function getCampaignUtilitySheets(identifier) {
    const config = getCampaignConfig(identifier);
    const ss = getCampaignUtilitiesSpreadsheet(identifier);
    const names = buildSheetNameCandidates(config || identifier || {});
    const sheets = {};

    Object.keys(names).forEach(function (key) {
      const options = names[key];
      for (let i = 0; i < options.length; i++) {
        const sheet = ensureSheet(ss, options[i]);
        if (sheet) {
          sheets[key] = sheet;
          break;
        }
      }
    });

    return { config, spreadsheet: ss, sheets };
  }

  root.CampaignResourceRegistry = {
    getCampaignConfig,
    getCampaignConfigs,
    getCampaignUtilitiesSpreadsheet,
    getCampaignScheduleSpreadsheet,
    ensureSheet,
    getCampaignUtilitySheetNames,
    getCampaignUtilitySheets
  };

  // Convenience aliases for consumers
  if (typeof root.getCampaignConfig !== 'function') root.getCampaignConfig = getCampaignConfig;
  if (typeof root.getCampaignConfigs !== 'function') root.getCampaignConfigs = getCampaignConfigs;
  if (typeof root.getCampaignUtilitiesSpreadsheet !== 'function') root.getCampaignUtilitiesSpreadsheet = getCampaignUtilitiesSpreadsheet;
  if (typeof root.getCampaignScheduleSpreadsheet !== 'function') root.getCampaignScheduleSpreadsheet = getCampaignScheduleSpreadsheet;
})();
