/**
 * IdentityRepository.gs
 * -----------------------------------------------------------------------------
 * Shared data-access helpers for the Lumina Identity platform. Provides typed
 * wrappers around the Sheets-based data model defined in the Lumina Identity
 * specification. Every table is represented by a sheet whose header row matches
 * the contract from the design document. The repository automatically ensures
 * header integrity, provides optimistic locking semantics, and exposes helpers
 * to read/write strongly typed objects.
 */
(function bootstrapIdentityRepository(global) {
  if (!global) return;
  if (global.IdentityRepository && typeof global.IdentityRepository === 'object') {
    return;
  }

  var SpreadsheetApp = global.SpreadsheetApp;
  var PropertiesService = global.PropertiesService;
  var CacheService = global.CacheService;
  var LockService = global.LockService;
  var Utilities = global.Utilities;

  var SPREADSHEET_ID_PROPERTY = 'IDENTITY_SPREADSHEET_ID';
  var CACHE_TTL_SECONDS = 60;
  var HEADER_ROW = 1;

  var LEGACY_FALLBACK_TABLE_HEADERS = {
    Campaigns: ['CampaignId', 'Name', 'Status', 'ClientOwnerEmail', 'CreatedAt', 'SettingsJSON'],
    Users: ['UserId', 'Email', 'Username', 'PasswordHash', 'EmailVerified', 'TOTPEnabled', 'TOTPSecretHash', 'Status', 'LastLoginAt', 'CreatedAt', 'Role', 'CampaignId', 'UpdatedAt', 'FlagsJson', 'Watchlist'],
    UserCampaigns: ['AssignmentId', 'UserId', 'CampaignId', 'Role', 'IsPrimary', 'AddedBy', 'AddedAt', 'Watchlist'],
    Roles: ['RoleId', 'Role', 'Name', 'Description', 'PermissionsJson', 'DefaultForCampaignManager', 'IsGlobal'],
    RolePermissions: ['PermissionId', 'Role', 'Capability', 'Scope', 'Allowed'],
    OTP: ['Key', 'Email', 'Code', 'Purpose', 'ExpiresAt', 'Attempts', 'LastSentAt', 'ResendCount'],
    Sessions: ['SessionId', 'UserId', 'CampaignId', 'IssuedAt', 'ExpiresAt', 'CSRF', 'IP', 'UA'],
    LoginAttempts: ['EmailOrUsername', 'Count1m', 'Count15m', 'LastAttemptAt', 'LockedUntil'],
    Equipment: ['EquipmentId', 'UserId', 'CampaignId', 'Type', 'Serial', 'Condition', 'AssignedAt', 'ReturnedAt', 'Notes', 'Status'],
    EmploymentStatus: ['UserId', 'CampaignId', 'State', 'EffectiveDate', 'Reason', 'Notes'],
    EligibilityRules: ['RuleId', 'Name', 'Scope', 'RuleType', 'ParamsJSON', 'Active'],
    AuditLog: ['EventId', 'Timestamp', 'ActorUserId', 'ActorRole', 'CampaignId', 'Target', 'Action', 'BeforeJSON', 'AfterJSON', 'Mode', 'IP', 'UA'],
    FeatureFlags: ['Key', 'Value', 'Env', 'Description', 'UpdatedAt', 'Flag', 'Notes'],
    Policies: ['PolicyId', 'Name', 'Scope', 'Key', 'Value', 'UpdatedAt'],
    QualityScores: ['RecordId', 'UserId', 'CampaignId', 'Score', 'Date'],
    Attendance: ['RecordId', 'UserId', 'CampaignId', 'Date', 'State', 'Start', 'End', 'Productive', 'Minutes'],
    Performance: ['RecordId', 'UserId', 'CampaignId', 'Metric', 'Score', 'Date'],
    Shifts: ['ShiftId', 'CampaignId', 'UserId', 'Date', 'StartTime', 'EndTime', 'Status'],
    QAAudits: ['AuditId', 'UserId', 'CampaignId', 'Score', 'Band', 'AutoFail', 'CreatedAt', 'DetailsUrl'],
    Coaching: ['CoachId', 'UserId', 'CampaignId', 'Plan', 'DueDate', 'Status'],
    Benefits: ['UserId', 'Eligible', 'Reason', 'EffectiveDate'],
    PayrollSync: ['CampaignId', 'RunId', 'Status', 'ErrorsJson', 'StartedAt', 'EndedAt'],
    SystemMessages: ['MessageId', 'Severity', 'Title', 'Body', 'TargetRole', 'TargetCampaignId', 'Status', 'CreatedAt', 'ResolvedAt', 'CreatedBy', 'MetadataJson'],
    Jobs: ['JobId', 'Name', 'Schedule', 'LastRunAt', 'LastStatus', 'ConfigJson', 'Enabled', 'RunHash']
  };

  function cloneHeaderMap(map) {
    var clone = {};
    Object.keys(map).forEach(function(name) {
      clone[name] = Array.isArray(map[name]) ? map[name].slice() : [];
    });
    return clone;
  }

  function coerceHeaderMap(source, fallback) {
    var resolved = {};
    if (source && typeof source === 'object') {
      Object.keys(source).forEach(function(name) {
        if (Array.isArray(source[name]) && source[name].length) {
          resolved[name] = source[name].slice();
        }
      });
    }
    if (fallback && typeof fallback === 'object') {
      Object.keys(fallback).forEach(function(name) {
        if (!Array.isArray(resolved[name]) || !resolved[name].length) {
          var headers = fallback[name];
          if (Array.isArray(headers) && headers.length) {
            resolved[name] = headers.slice();
          }
        }
      });
    }
    return resolved;
  }

  var defaultsFromMain = coerceHeaderMap(
    global && global.LUMINA_IDENTITY_TABLE_HEADER_DEFAULTS,
    LEGACY_FALLBACK_TABLE_HEADERS
  );

  if (global && !global.LUMINA_IDENTITY_TABLE_HEADER_DEFAULTS) {
    global.LUMINA_IDENTITY_TABLE_HEADER_DEFAULTS = cloneHeaderMap(defaultsFromMain);
  }

  var canonicalHeaderSource = coerceHeaderMap(
    global && global.LUMINA_IDENTITY_CANONICAL_TABLE_HEADERS,
    defaultsFromMain
  );

  if (!global || !global.LUMINA_IDENTITY_CANONICAL_TABLE_HEADERS) {
    if (global) {
      global.LUMINA_IDENTITY_CANONICAL_TABLE_HEADERS = cloneHeaderMap(canonicalHeaderSource);
    }
  } else {
    Object.keys(canonicalHeaderSource).forEach(function(name) {
      if (!Array.isArray(global.LUMINA_IDENTITY_CANONICAL_TABLE_HEADERS[name])
        || !global.LUMINA_IDENTITY_CANONICAL_TABLE_HEADERS[name].length) {
        global.LUMINA_IDENTITY_CANONICAL_TABLE_HEADERS[name] = canonicalHeaderSource[name].slice();
      }
    });
  }

  var TABLE_HEADERS = cloneHeaderMap(
    (global && global.LUMINA_IDENTITY_CANONICAL_TABLE_HEADERS)
      ? global.LUMINA_IDENTITY_CANONICAL_TABLE_HEADERS
      : canonicalHeaderSource
  );

  var repositoryCache = CacheService ? CacheService.getScriptCache() : null;
  var spreadsheetCache = null;

  function getSpreadsheetId() {
    var scriptProperties = PropertiesService && PropertiesService.getScriptProperties();
    if (!scriptProperties) {
      throw new Error('Script properties unavailable – configure IDENTITY_SPREADSHEET_ID.');
    }
    var id = scriptProperties.getProperty(SPREADSHEET_ID_PROPERTY);
    if (!id) {
      throw new Error('Missing script property: IDENTITY_SPREADSHEET_ID');
    }
    return id;
  }

  function getSpreadsheet() {
    if (spreadsheetCache) {
      return spreadsheetCache;
    }
    if (!SpreadsheetApp) {
      throw new Error('SpreadsheetApp unavailable');
    }
    spreadsheetCache = SpreadsheetApp.openById(getSpreadsheetId());
    return spreadsheetCache;
  }

  function normalizeHeaders(values) {
    return (values || []).map(function(value) {
      return String(value || '').trim();
    });
  }

  function ensureSheet(name) {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    var headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn() || TABLE_HEADERS[name].length).getValues()[0];
    var normalized = normalizeHeaders(headers);
    var expected = TABLE_HEADERS[name];
    var needsWrite = expected.length !== normalized.length || expected.some(function(header, idx) {
      return normalized[idx] !== header;
    });

    if (needsWrite) {
      sheet.clear();
      sheet.getRange(HEADER_ROW, 1, 1, expected.length).setValues([expected]);
    }
    return sheet;
  }

  function withLock(name, callback) {
    var lock = LockService ? LockService.getScriptLock() : null;
    if (!lock) {
      return callback();
    }
    lock.waitLock(30000);
    try {
      return callback();
    } finally {
      lock.releaseLock();
    }
  }

  function list(name) {
    var cacheKey = 'identity-table-' + name;
    if (repositoryCache) {
      var cached = repositoryCache.get(cacheKey);
      if (cached) {
        try {
          return JSON.parse(cached);
        } catch (err) {
          repositoryCache.remove(cacheKey);
        }
      }
    }

    var sheet = ensureSheet(name);
    var range = sheet.getDataRange();
    var values = range.getValues();
    if (values.length <= 1) {
      return [];
    }
    var headers = values[0];
    var rows = [];
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      if (row.join('').trim() === '') {
        continue;
      }
      var obj = {};
      headers.forEach(function(header, idx) {
        obj[String(header)] = row[idx];
      });
      rows.push(obj);
    }

    if (repositoryCache) {
      repositoryCache.put(cacheKey, JSON.stringify(rows), CACHE_TTL_SECONDS);
    }
    return rows;
  }

  function write(name, rows) {
    var sheet = ensureSheet(name);
    var headers = TABLE_HEADERS[name];
    var values = [headers];
    rows.forEach(function(row) {
      var arr = headers.map(function(header) {
        return (row && Object.prototype.hasOwnProperty.call(row, header)) ? row[header] : '';
      });
      values.push(arr);
    });
    sheet.clearContents();
    sheet.getRange(1, 1, values.length, headers.length).setValues(values);
    if (repositoryCache) {
      repositoryCache.remove('identity-table-' + name);
    }
  }

  function upsert(name, key, payload) {
    if (!payload || !key) {
      throw new Error('Invalid upsert payload for ' + name);
    }
    return withLock('identity-upsert-' + name, function() {
      var rows = list(name);
      var found = false;
      for (var i = 0; i < rows.length; i++) {
        if (rows[i][key] === payload[key]) {
          rows[i] = Object.assign({}, rows[i], payload);
          found = true;
          break;
        }
      }
      if (!found) {
        rows.push(payload);
      }
      write(name, rows);
      return payload;
    });
  }

  function remove(name, key, value) {
    return withLock('identity-remove-' + name, function() {
      var rows = list(name).filter(function(row) {
        return row[key] !== value;
      });
      write(name, rows);
    });
  }

  function append(name, payload) {
    var sheet = ensureSheet(name);
    var headers = TABLE_HEADERS[name];
    var row = headers.map(function(header) {
      return (payload && Object.prototype.hasOwnProperty.call(payload, header)) ? payload[header] : '';
    });
    sheet.appendRow(row);
    if (repositoryCache) {
      repositoryCache.remove('identity-table-' + name);
    }
    return payload;
  }

  function find(name, predicate) {
    var rows = list(name);
    for (var i = 0; i < rows.length; i++) {
      if (predicate(rows[i])) {
        return rows[i];
      }
    }
    return null;
  }

  global.IdentityRepository = {
    list: list,
    append: append,
    upsert: upsert,
    remove: remove,
    find: find,
    ensureSheet: ensureSheet,
    TABLE_HEADERS: TABLE_HEADERS
  };
})(GLOBAL_SCOPE);
