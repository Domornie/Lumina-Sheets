/**
 * DatabaseBindings.gs
 * Bridges legacy sheet helpers with the centralized DatabaseManager CRUD abstraction.
 * - Automatically registers known sheet schemas with DatabaseManager
 * - Provides global CRUD helpers (dbSelect/dbCreate/dbUpdate/dbDelete/dbUpsert)
 * - Falls back to direct Spreadsheet operations if DatabaseManager is unavailable
 */

(function (global) {
  if (!global) return;

  var schemaRegistry = global.__DB_SCHEMA_REGISTRY__ || {};
  global.__DB_SCHEMA_REGISTRY__ = schemaRegistry;

  function getManager() {
    return typeof global.DatabaseManager !== 'undefined' ? global.DatabaseManager : null;
  }

  function inferIdFromHeaders(headers) {
    if (!Array.isArray(headers) || headers.length === 0) return null;
    var normalized = headers.map(function (h) { return String(h || '').trim(); });
    var preferred = ['ID', 'Id', 'Uuid', 'UUID'];
    for (var i = 0; i < preferred.length; i++) {
      var idx = normalized.indexOf(preferred[i]);
      if (idx !== -1) return headers[idx];
    }
    for (var j = 0; j < normalized.length; j++) {
      if (/Id$/i.test(normalized[j])) return headers[j];
      if (/Key$/i.test(normalized[j])) return headers[j];
    }
    return null;
  }

  function cloneHeaders(headers) {
    return Array.isArray(headers) ? headers.slice() : undefined;
  }

  function registerIfDefined(sheetName, headers, idColumn, extra) {
    if (!sheetName || !Array.isArray(headers)) return;
    if (schemaRegistry[sheetName] && schemaRegistry[sheetName].headers && schemaRegistry[sheetName].headers.length) return;
    var config = { headers: cloneHeaders(headers) };
    if (typeof idColumn !== 'undefined') config.idColumn = idColumn;
    if (extra) {
      for (var key in extra) {
        if (Object.prototype.hasOwnProperty.call(extra, key)) {
          if (key.toLowerCase().indexOf('tenant') !== -1) continue;
          config[key] = extra[key];
        }
      }
    }
    registerTableSchema(sheetName, config);
  }

  function attemptRegisterKnownSchemas() {
    registerIfDefined(global.USERS_SHEET || 'Users', global.USERS_HEADERS, 'ID');
    registerIfDefined(global.ROLES_SHEET || 'Roles', global.ROLES_HEADER, 'ID');
    registerIfDefined(global.USER_ROLES_SHEET || 'UserRoles', global.USER_ROLES_HEADER, 'UserId');
    registerIfDefined(global.USER_CLAIMS_SHEET || 'UserClaims', global.CLAIMS_HEADERS, 'ID');
    registerIfDefined(global.SESSIONS_SHEET || 'Sessions', global.SESSIONS_HEADERS, 'Token');
    registerIfDefined(global.CAMPAIGNS_SHEET || 'Campaigns', global.CAMPAIGNS_HEADERS, 'ID');
    registerIfDefined(global.PAGES_SHEET || 'Pages', global.PAGES_HEADERS, 'PageKey');
    registerIfDefined(global.CAMPAIGN_PAGES_SHEET || 'CampaignPages', global.CAMPAIGN_PAGES_HEADERS, 'ID');
    registerIfDefined(global.PAGE_CATEGORIES_SHEET || 'PageCategories', global.PAGE_CATEGORIES_HEADERS, 'ID');
    registerIfDefined(global.CAMPAIGN_USER_PERMISSIONS_SHEET || 'CampaignUserPermissions', global.CAMPAIGN_USER_PERMISSIONS_HEADERS, 'ID');
    registerIfDefined(global.USER_MANAGERS_SHEET || 'UserManagers', global.USER_MANAGERS_HEADERS, 'ID');
    registerIfDefined(global.USER_CAMPAIGNS_SHEET || 'UserCampaigns', global.USER_CAMPAIGNS_HEADERS, 'ID');
    registerIfDefined(global.NOTIFICATIONS_SHEET || 'Notifications', global.NOTIFICATIONS_HEADERS, 'ID');
    registerIfDefined(global.DEBUG_LOGS_SHEET || 'DebugLogs', global.DEBUG_LOGS_HEADERS, 'Timestamp', { timestamps: false, idColumn: 'Timestamp' });
    registerIfDefined(global.ERROR_LOGS_SHEET || 'ErrorLogs', global.ERROR_LOGS_HEADERS, 'Timestamp', { timestamps: false, idColumn: 'Timestamp' });

    registerIfDefined(global.CHAT_GROUPS_SHEET || 'ChatGroups', global.CHAT_GROUPS_HEADERS, 'ID');
    registerIfDefined(global.CHAT_CHANNELS_SHEET || 'ChatChannels', global.CHAT_CHANNELS_HEADERS, 'ID');
    registerIfDefined(global.CHAT_MESSAGES_SHEET || 'ChatMessages', global.CHAT_MESSAGES_HEADERS, 'ID');
    registerIfDefined(global.CHAT_GROUP_MEMBERS_SHEET || 'ChatGroupMembers', global.CHAT_GROUP_MEMBERS_HEADERS, 'ID');
    registerIfDefined(global.CHAT_MESSAGE_REACTIONS_SHEET || 'ChatMessageReactions', global.CHAT_MESSAGE_REACTIONS_HEADERS, 'ID');
    registerIfDefined(global.CHAT_USER_PREFERENCES_SHEET || 'ChatUserPreferences', global.CHAT_USER_PREFERENCES_HEADERS, 'UserId');
    registerIfDefined(global.CHAT_ANALYTICS_SHEET || 'ChatAnalytics', global.CHAT_ANALYTICS_HEADERS, 'Timestamp', { timestamps: false, idColumn: 'Timestamp' });
    registerIfDefined(global.CHAT_CHANNEL_MEMBERS_SHEET || 'ChatChannelMembers', global.CHAT_CHANNEL_MEMBERS_HEADERS, 'ID');

    registerIfDefined(global.ATTENDANCE_LOG_SHEET || global.ATTENDANCE_SHEET || 'AttendanceLog', global.ATTENDANCE_LOG_HEADERS, 'ID');
  }

  function applySchemaToManager(name) {
    var manager = getManager();
    if (!manager) return;
    var schema = schemaRegistry[name];
    if (!schema) return;
    manager.defineTable(name, schema);
  }

  function inferSchemaFromSheet(name) {
    var schema = { headers: [], idColumn: null };
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) return schema;
      var sh = ss.getSheetByName(name);
      if (!sh) return schema;
      var lastCol = sh.getLastColumn();
      if (lastCol < 1) return schema;
      var headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
      schema.headers = headerRow.map(function (h) { return String(h || '').trim(); });
      schema.idColumn = inferIdFromHeaders(schema.headers);
      return schema;
    } catch (err) {
      if (global.safeWriteError) {
        try { global.safeWriteError('inferSchemaFromSheet', err); } catch (_) { }
      }
      return schema;
    }
  }

  function ensureSchema(name) {
    attemptRegisterKnownSchemas();
    if (!schemaRegistry[name]) {
      schemaRegistry[name] = inferSchemaFromSheet(name);
    }
    applySchemaToManager(name);
    return schemaRegistry[name];
  }

  function registerTableSchema(name, options) {
    if (!name) return null;
    var schema = options ? Object.assign({}, options) : {};
    if (schema.headers) schema.headers = cloneHeaders(schema.headers);
    if (schema.tenantColumn) delete schema.tenantColumn;
    if (schema.requireTenant) delete schema.requireTenant;
    if (schema.allowGlobalTenantBypass) delete schema.allowGlobalTenantBypass;
    if (!Object.prototype.hasOwnProperty.call(schema, 'idColumn')) {
      schema.idColumn = inferIdFromHeaders(schema.headers);
    }
    schemaRegistry[name] = schema;
    applySchemaToManager(name);
    return schema;
  }

  function dbTable(name, context) {
    if (!name) throw new Error('Sheet name is required');
    var manager = getManager();
    ensureSchema(name);
    if (!manager) return null;
    if (typeof context !== 'undefined' && context !== null) {
      return manager.table(name, context);
    }
    return manager.table(name);
  }

  function dbSelect(name, options, context) {
    var manager = getManager();
    ensureSchema(name);
    var query = options ? Object.assign({}, options) : {};
    if (manager) {
      try {
        return manager.table(name, context).find(query, context);
      } catch (err) {
        if (global.safeWriteError) {
          try { global.safeWriteError('dbSelect', err); } catch (_) { }
        }
      }
    }
    return applyQueryOptions(legacyReadSheetData(name), query);
  }

  function dbCreate(name, record, context) {
    if (!record || typeof record !== 'object') return null;
    var table = dbTable(name, context);
    if (table) {
      return table.insert(record, context);
    }
    return legacyInsert(name, record);
  }

  function buildWhereFromIdentifier(table, identifier) {
    if (identifier && typeof identifier === 'object') return identifier;
    if (!table || !table.idColumn) return null;
    if (identifier === null || typeof identifier === 'undefined') return null;
    var where = {};
    where[table.idColumn] = identifier;
    return where;
  }

  function dbUpdate(name, identifier, updates, context) {
    if (!updates || typeof updates !== 'object') return null;
    var table = dbTable(name, context);
    if (table) {
      if (table.idColumn && typeof identifier !== 'object') {
        return table.update(identifier, updates, context);
      }
      var whereClause = buildWhereFromIdentifier(table, identifier) || (identifier && typeof identifier === 'object' ? identifier : null);
      if (!whereClause) throw new Error('Update requires an ID or where clause');
      var existing = table.findOne(whereClause, context);
      if (existing && table.idColumn && existing[table.idColumn]) {
        return table.update(existing[table.idColumn], updates, context);
      }
      if (existing) {
        return legacyUpdate(name, whereClause, updates);
      }
      return null;
    }
    var where = identifier && typeof identifier === 'object' ? identifier : null;
    if (!where) throw new Error('Update requires an ID or where clause');
    return legacyUpdate(name, where, updates);
  }

  function dbUpsert(name, where, updates, context) {
    var table = dbTable(name, context);
    if (table) {
      return table.upsert(where || {}, updates || {}, context);
    }
    var existing = applyQueryOptions(legacyReadSheetData(name), { where: where, limit: 1 });
    if (existing.length) {
      return legacyUpdate(name, where, updates || {});
    }
    var payload = Object.assign({}, where || {}, updates || {});
    return dbCreate(name, payload, context);
  }

  function dbDelete(name, identifier, context) {
    var table = dbTable(name, context);
    if (table) {
      if (table.idColumn && typeof identifier !== 'object') {
        return table.delete(identifier, context);
      }
      var whereClause = buildWhereFromIdentifier(table, identifier) || (identifier && typeof identifier === 'object' ? identifier : null);
      if (!whereClause) throw new Error('Delete requires an ID or where clause');
      if (table.idColumn) {
        var existing = table.findOne(whereClause, context);
        if (existing && existing[table.idColumn]) {
          return table.delete(existing[table.idColumn], context);
        }
      }
      return legacyDelete(name, whereClause);
    }
    var where = identifier && typeof identifier === 'object' ? identifier : null;
    if (!where) throw new Error('Delete requires an ID or where clause');
    return legacyDelete(name, where);
  }

  function legacyInsert(name, record) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    var sh = ss.getSheetByName(name);
    var headers;
    if (sh) {
      var lastCol = sh.getLastColumn();
      headers = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').trim(); }) : [];
    } else {
      var headerKeys = Object.keys(record || {});
      if (typeof ensureSheetWithHeaders === 'function' && headerKeys.length) {
        sh = ensureSheetWithHeaders(name, headerKeys);
        headers = headerKeys;
      } else {
        sh = ss.insertSheet(name);
        headers = headerKeys;
        if (headers.length) {
          sh.getRange(1, 1, 1, headers.length).setValues([headers]);
          sh.setFrozenRows(1);
        }
      }
    }
    if (!headers || headers.length === 0) {
      headers = Object.keys(record || {});
      if (headers.length) {
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
        sh.setFrozenRows(1);
      }
    }
    var row = headers.map(function (header) { return typeof record[header] === 'undefined' ? '' : record[header]; });
    sh.appendRow(row);
    return record;
  }

  function legacyUpdate(name, where, updates) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    var sh = ss.getSheetByName(name);
    if (!sh || !where) return null;
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return null;
    var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').trim(); });
    var range = sh.getRange(2, 1, lastRow - 1, lastCol);
    var values = range.getValues();
    for (var i = 0; i < values.length; i++) {
      var rowObj = {};
      for (var j = 0; j < headers.length; j++) {
        rowObj[headers[j]] = values[i][j];
      }
      if (!matchesWhere(rowObj, where)) continue;

      Object.keys(updates || {}).forEach(function (key) {
        rowObj[key] = updates[key];
      });
      var serialized = headers.map(function (header) { return typeof rowObj[header] === 'undefined' ? '' : rowObj[header]; });
      range.getCell(i + 1, 1).offset(0, 0, 1, lastCol).setValues([serialized]);
      return rowObj;
    }
    return null;
  }

  function legacyDelete(name, where) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return false;
    var sh = ss.getSheetByName(name);
    if (!sh || !where) return false;
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return false;
    var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').trim(); });
    var range = sh.getRange(2, 1, lastRow - 1, lastCol);
    var values = range.getValues();
    for (var i = 0; i < values.length; i++) {
      var rowObj = {};
      for (var j = 0; j < headers.length; j++) {
        rowObj[headers[j]] = values[i][j];
      }
      if (!matchesWhere(rowObj, where)) continue;
      sh.deleteRow(i + 2);
      return true;
    }
    return false;
  }

  function legacyReadSheetData(name) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) return [];
      var sh = ss.getSheetByName(name);
      if (!sh) return [];
      var lastRow = sh.getLastRow();
      var lastCol = sh.getLastColumn();
      if (lastRow < 2 || lastCol < 1) return [];
      var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
      var headers = values.shift().map(function (h) { return String(h || '').trim(); });
      if (headers.some(function (h) { return !h; })) return [];
      var unique = {};
      for (var i = 0; i < headers.length; i++) {
        if (unique[headers[i]]) return [];
        unique[headers[i]] = true;
      }
      return values.map(function (row) {
        var obj = {};
        for (var j = 0; j < headers.length; j++) {
          obj[headers[j]] = typeof row[j] === 'undefined' ? '' : row[j];
        }
        return obj;
      });
    } catch (err) {
      if (global.safeWriteError) {
        try { global.safeWriteError('legacyReadSheetData', err); } catch (_) { }
      }
      return [];
    }
  }

  function applyQueryOptions(rows, options) {
    if (!Array.isArray(rows)) return [];
    var result = rows.slice();
    if (options && options.where) {
      result = result.filter(function (row) { return matchesWhere(row, options.where); });
    }
    if (options && typeof options.filter === 'function') {
      result = result.filter(options.filter);
    }
    if (options && typeof options.map === 'function') {
      result = result.map(options.map);
    }
    if (options && options.sortBy) {
      var key = options.sortBy;
      var desc = !!options.sortDesc;
      result.sort(function (a, b) {
        var av = a[key];
        var bv = b[key];
        if (av === bv) return 0;
        if (av === undefined || av === null || av === '') return desc ? 1 : -1;
        if (bv === undefined || bv === null || bv === '') return desc ? -1 : 1;
        if (av > bv) return desc ? -1 : 1;
        if (av < bv) return desc ? 1 : -1;
        return 0;
      });
    }
    if (options && (options.offset || typeof options.limit === 'number')) {
      var start = options.offset || 0;
      var end = typeof options.limit === 'number' ? start + options.limit : result.length;
      result = result.slice(start, end);
    }
    if (options && Array.isArray(options.columns) && options.columns.length) {
      result = result.map(function (row) {
        var projected = {};
        options.columns.forEach(function (col) { projected[col] = row[col]; });
        return projected;
      });
    }
    return result;
  }

  function matchesWhere(row, where) {
    if (!where || typeof where !== 'object') return true;
    var keys = Object.keys(where);
    for (var i = 0; i < keys.length; i++) {
      var key = keys[i];
      var expected = where[key];
      var actual = row[key];
      if (expected instanceof RegExp) {
        if (!expected.test(String(actual || ''))) return false;
      } else if (expected && typeof expected === 'object' && !Array.isArray(expected)) {
        if (!evaluateWhereOperator(actual, expected)) return false;
      } else if (String(actual) !== String(expected)) {
        return false;
      }
    }
    return true;
  }

  function evaluateWhereOperator(actual, expression) {
    var ops = Object.keys(expression);
    for (var i = 0; i < ops.length; i++) {
      var op = ops[i];
      var value = expression[op];
      switch (op) {
        case '$gt':
          if (!(actual > value)) return false;
          break;
        case '$gte':
          if (!(actual >= value)) return false;
          break;
        case '$lt':
          if (!(actual < value)) return false;
          break;
        case '$lte':
          if (!(actual <= value)) return false;
          break;
        case '$ne':
          if (actual === value) return false;
          break;
        case '$in':
          if (!Array.isArray(value) || value.indexOf(actual) === -1) return false;
          break;
        case '$nin':
          if (Array.isArray(value) && value.indexOf(actual) !== -1) return false;
          break;
        default:
          if (String(actual) !== String(expression[op])) return false;
      }
    }
    return true;
  }

  function dbTenantSelect(name, context, options) {
    return dbSelect(name, options || {}, context);
  }

  function dbTenantCreate(name, context, record) {
    return dbCreate(name, record, context);
  }

  function dbTenantUpdate(name, context, identifier, updates) {
    return dbUpdate(name, identifier, updates, context);
  }

  function dbTenantUpsert(name, context, where, updates) {
    return dbUpsert(name, where, updates, context);
  }

  function dbTenantDelete(name, context, identifier) {
    return dbDelete(name, identifier, context);
  }

  function expose(name, fn) {
    var previous = (typeof global[name] === 'function') ? global[name].bind(global) : null;
    var wrapped = function () {
      try {
        return fn.apply(this, arguments);
      } catch (err) {
        if (previous) {
          try { return previous.apply(this, arguments); } catch (fallbackErr) { if (global.safeWriteError) { try { global.safeWriteError(name + 'Fallback', fallbackErr); } catch (_) { } } }
        }
        throw err;
      }
    };
    wrapped.previous = previous;
    global[name] = wrapped;
  }

  expose('registerTableSchema', registerTableSchema);
  expose('dbTable', dbTable);
  expose('dbSelect', dbSelect);
  expose('dbCreate', dbCreate);
  expose('dbUpdate', dbUpdate);
  expose('dbUpsert', dbUpsert);
  expose('dbDelete', dbDelete);
  expose('dbTenantSelect', dbTenantSelect);
  expose('dbTenantCreate', dbTenantCreate);
  expose('dbTenantUpdate', dbTenantUpdate);
  expose('dbTenantUpsert', dbTenantUpsert);
  expose('dbTenantDelete', dbTenantDelete);
  expose('dbWithContext', function (context) {
    var manager = getManager();
    if (manager) {
      if (typeof manager.withContext === 'function') {
        return manager.withContext(context);
      }
      if (typeof manager.tenant === 'function') {
        return manager.tenant(context);
      }
    }
    return {
      select: function (name, options) { return dbSelect(name, options, context); },
      create: function (name, record) { return dbCreate(name, record, context); },
      update: function (name, identifier, updates) { return dbUpdate(name, identifier, updates, context); },
      upsert: function (name, where, updates) { return dbUpsert(name, where, updates, context); },
      delete: function (name, identifier) { return dbDelete(name, identifier, context); }
    };
  });
  expose('__dbApplyQueryOptions', applyQueryOptions);

})(typeof globalThis !== 'undefined' ? globalThis : this);
