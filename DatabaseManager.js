/**
 * DatabaseManager.gs
 * Unified Google Sheets database abstraction for CRUD operations across all sheets.
 *
 * Usage:
 *   const usersTable = DatabaseManager.defineTable('Users', {
 *     headers: ['ID', 'UserName', 'Email'],
 *     defaults: { CanLogin: true },
 *   });
 *   const user = usersTable.insert({ UserName: 'alice', Email: 'alice@example.com' });
 *
 *   const schedules = DatabaseManager.table('Schedules').find({ where: { UserId: user.ID } });
 *
 * The manager automatically:
 *   - Ensures sheet + headers exist (appends missing headers without overwriting data)
 *   - Provides strongly-typed CRUD helpers with optional caching and timestamps
 *   - Treats Google Sheets like database tables that can be registered at runtime
 */

(function (global) {
  if (global.DatabaseManager) return;

  var DEFAULT_CACHE_TTL = 300; // seconds
  var DEFAULT_ID_COLUMN = 'ID';
  var DEFAULT_CREATED_AT = 'CreatedAt';
  var DEFAULT_UPDATED_AT = 'UpdatedAt';
  var tables = {};
  var logger = (typeof console !== 'undefined') ? console : {
    log: function () { },
    info: function () { },
    warn: function () { },
    error: function () { }
  };

  function isObject(value) {
    return value && typeof value === 'object' && !Array.isArray(value);
  }

  function clone(value) {
    if (!isObject(value)) return value;
    var copy = {};
    Object.keys(value).forEach(function (key) {
      copy[key] = value[key];
    });
    return copy;
  }

  function SpreadsheetHandle(table) {
    this.table = table;
  }

  SpreadsheetHandle.prototype.getSheet = function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(this.table.name);
    if (!sh) {
      sh = ss.insertSheet(this.table.name);
    }
    this.table.ensureHeaders(sh);
    return sh;
  };

  SpreadsheetHandle.prototype.readAllRows = function () {
    var sheet = this.getSheet();
    var lastRow = sheet.getLastRow();
    var headerCount = this.table.headers.length;
    if (lastRow < 2 || headerCount === 0) {
      return [];
    }
    var range = sheet.getRange(2, 1, lastRow - 1, headerCount);
    return range.getValues();
  };

  SpreadsheetHandle.prototype.writeRow = function (rowIndex, rowValues) {
    var sheet = this.getSheet();
    var headerCount = this.table.headers.length;
    var range = sheet.getRange(rowIndex, 1, 1, headerCount);
    range.setValues([rowValues]);
  };

  SpreadsheetHandle.prototype.appendRow = function (rowValues) {
    var sheet = this.getSheet();
    sheet.appendRow(rowValues);
  };

  SpreadsheetHandle.prototype.deleteRow = function (rowIndex) {
    var sheet = this.getSheet();
    sheet.deleteRow(rowIndex);
  };

  function toStringValue(value) {
    if (value === null || typeof value === 'undefined') return '';
    return String(value);
  }

  function dedupeValues(values) {
    var out = [];
    var seen = {};
    for (var i = 0; i < values.length; i++) {
      var v = toStringValue(values[i]);
      if (!v) continue;
      if (!seen[v]) {
        seen[v] = true;
        out.push(v);
      }
    }
    return out;
  }

  function normalizeTenantContext(context) {
    if (!context && context !== 0) return null;
    if (typeof context === 'string' || typeof context === 'number') {
      return { tenantId: toStringValue(context) };
    }
    if (!isObject(context)) return null;
    var normalized = {};
    Object.keys(context).forEach(function (key) {
      normalized[key] = context[key];
    });
    if (normalized.campaignId && !normalized.tenantId) {
      normalized.tenantId = normalized.campaignId;
    }
    if (normalized.campaignIds && !normalized.tenantIds) {
      normalized.tenantIds = Array.isArray(normalized.campaignIds)
        ? normalized.campaignIds.slice()
        : normalized.campaignIds;
    }
    return normalized;
  }

  function Table(name, config) {
    this.name = name;
    if (config && Object.prototype.hasOwnProperty.call(config, 'idColumn')) {
      this.idColumn = config.idColumn;
    } else {
      this.idColumn = DEFAULT_ID_COLUMN;
    }
    this.cacheTTL = (config && config.cacheTTL) || DEFAULT_CACHE_TTL;
    if (config && config.timestamps === false) {
      this.timestamps = null;
    } else if (config && config.timestamps) {
      this.timestamps = {
        created: config.timestamps.created || DEFAULT_CREATED_AT,
        updated: config.timestamps.updated || DEFAULT_UPDATED_AT
      };
    } else {
      this.timestamps = {
        created: DEFAULT_CREATED_AT,
        updated: DEFAULT_UPDATED_AT
      };
    }
    this.defaults = (config && config.defaults) || {};
    this.validators = (config && config.validators) || {};
    this.tenantColumn = (config && config.tenantColumn) || null;
    if (this.tenantColumn) {
      if (config && Object.prototype.hasOwnProperty.call(config, 'requireTenant')) {
        this.requireTenant = !!config.requireTenant;
      } else {
        this.requireTenant = true;
      }
    } else {
      this.requireTenant = false;
    }
    this.allowGlobalTenantBypass = !!(config && config.allowGlobalTenantBypass);
    this.cacheKey = 'DB_TABLE_' + name;

    var providedHeaders = [];
    if (config && Array.isArray(config.headers)) {
      providedHeaders = config.headers.slice();
    } else if (config && Array.isArray(config.columns)) {
      providedHeaders = config.columns.slice();
    }

    this.headers = normalizeHeaders(this, providedHeaders);
    this.sheetHandle = new SpreadsheetHandle(this);
  }

  Table.prototype.withContext = function (context) {
    return new ScopedTable(this, context || null);
  };

  Table.prototype.normalizeContext = function (context) {
    return normalizeTenantContext(context);
  };

  Table.prototype.getTenantAccess = function (context, allowGlobal) {
    if (!this.tenantColumn) {
      return { enforce: false, allowed: null };
    }

    var ctx = this.normalizeContext(context) || {};
    if (ctx.allowAllTenants || ctx.globalTenantAccess || this.allowGlobalTenantBypass) {
      return { enforce: false, allowed: null };
    }

    var all = [];
    if (Object.prototype.hasOwnProperty.call(ctx, 'tenantId')) {
      all.push(ctx.tenantId);
    }
    if (Array.isArray(ctx.tenantIds)) {
      all = all.concat(ctx.tenantIds);
    }
    if (Array.isArray(ctx.allowedTenants)) {
      all = all.concat(ctx.allowedTenants);
    }
    if (Array.isArray(ctx.allowedTenantIds)) {
      all = all.concat(ctx.allowedTenantIds);
    }
    if (ctx.campaignId) {
      all.push(ctx.campaignId);
    }
    if (Array.isArray(ctx.campaignIds)) {
      all = all.concat(ctx.campaignIds);
    }

    var allowed = dedupeValues(all);
    if (!allowed.length) {
      if (this.requireTenant && !allowGlobal) {
        throw new Error('Tenant context required for table ' + this.name);
      }
      return { enforce: false, allowed: null };
    }

    return { enforce: true, allowed: allowed, context: ctx };
  };

  Table.prototype.ensureTenantConditionAllowed = function (condition, allowedSet) {
    if (!condition) return;
    if (condition instanceof RegExp) {
      throw new Error('Regex filters are not permitted on tenant column ' + this.tenantColumn + ' for table ' + this.name);
    }

    function hasAllowed(value) {
      return allowedSet[toStringValue(value)] === true;
    }

    if (isObject(condition)) {
      var keys = Object.keys(condition);
      if (!keys.length) return;
      for (var i = 0; i < keys.length; i++) {
        var key = keys[i];
        var value = condition[key];
        if (Array.isArray(value)) {
          for (var j = 0; j < value.length; j++) {
            if (!hasAllowed(value[j])) {
              throw new Error('Tenant filter includes unauthorized campaign for table ' + this.name);
            }
          }
        } else if (!hasAllowed(value)) {
          throw new Error('Tenant filter includes unauthorized campaign for table ' + this.name);
        }
      }
      return;
    }

    if (!hasAllowed(condition)) {
      throw new Error('Tenant filter includes unauthorized campaign for table ' + this.name);
    }
  };

  Table.prototype.prepareTenantOptions = function (options, context, allowGlobal) {
    var cfg = this.getTenantAccess(context, allowGlobal);
    if (!cfg.enforce) {
      return { options: options || {}, tenantAccess: cfg };
    }

    var finalOptions = options ? Object.assign({}, options) : {};
    var allowedSet = {};
    for (var i = 0; i < cfg.allowed.length; i++) {
      allowedSet[toStringValue(cfg.allowed[i])] = true;
    }

    if (finalOptions.where && Object.prototype.hasOwnProperty.call(finalOptions.where, this.tenantColumn)) {
      this.ensureTenantConditionAllowed(finalOptions.where[this.tenantColumn], allowedSet);
    }

    var baseFilter = finalOptions.filter;
    var column = this.tenantColumn;
    var tenantFilter = function (row) {
      return allowedSet[toStringValue(row[column])];
    };
    if (typeof baseFilter === 'function') {
      finalOptions.filter = function (row) {
        return baseFilter(row) && tenantFilter(row);
      };
    } else {
      finalOptions.filter = tenantFilter;
    }

    return { options: finalOptions, tenantAccess: cfg };
  };

  Table.prototype.enforceTenantOnRecord = function (record, tenantAccess) {
    if (!this.tenantColumn) return record;
    if (!tenantAccess.enforce) {
      if (this.requireTenant && (record[this.tenantColumn] === null || typeof record[this.tenantColumn] === 'undefined' || record[this.tenantColumn] === '')) {
        throw new Error('Tenant column "' + this.tenantColumn + '" must be provided for table ' + this.name);
      }
      return record;
    }

    var allowed = {};
    for (var i = 0; i < tenantAccess.allowed.length; i++) {
      allowed[toStringValue(tenantAccess.allowed[i])] = true;
    }

    var value = toStringValue(record[this.tenantColumn]);
    if (!value) {
      if (tenantAccess.allowed.length === 1) {
        record[this.tenantColumn] = tenantAccess.allowed[0];
      } else {
        throw new Error('Tenant column "' + this.tenantColumn + '" must be specified when multiple campaigns are available for table ' + this.name);
      }
    } else if (!allowed[value]) {
      throw new Error('Tenant access denied for campaign ' + value + ' on table ' + this.name);
    }

    return record;
  };

  Table.prototype.ensureExistingTenantAllowed = function (existingRecord, tenantAccess) {
    if (!this.tenantColumn || !existingRecord) return;
    if (!tenantAccess.enforce) return;
    var currentTenant = toStringValue(existingRecord[this.tenantColumn]);
    if (!currentTenant) {
      throw new Error('Existing record missing tenant column "' + this.tenantColumn + '" in table ' + this.name);
    }
    var allowed = tenantAccess.allowed;
    for (var i = 0; i < allowed.length; i++) {
      if (toStringValue(allowed[i]) === currentTenant) {
        return;
      }
    }
    throw new Error('Tenant access denied for existing record in table ' + this.name);
  };

  Table.prototype.ensureHeaders = function (sheet) {
    var desiredHeaders = this.headers.slice();
    var lastColumn = sheet.getLastColumn();
    var existingHeaders = [];
    if (lastColumn > 0) {
      existingHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(String);
    }

    var finalHeaders;
    var changed = false;

    if (existingHeaders.length === 0) {
      finalHeaders = desiredHeaders.slice();
      if (finalHeaders.length) {
        sheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
        sheet.setFrozenRows(1);
      }
    } else {
      finalHeaders = existingHeaders.slice();
      desiredHeaders.forEach(function (header) {
        if (finalHeaders.indexOf(header) === -1) {
          finalHeaders.push(header);
          changed = true;
        }
      });
      if (changed) {
        sheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
      }
    }

    this.headers = finalHeaders;
  };

  Table.prototype.toObjects = function (rows) {
    var headers = this.headers;
    return rows.map(function (row) {
      var obj = {};
      headers.forEach(function (header, index) {
        obj[header] = typeof row[index] === 'undefined' ? '' : row[index];
      });
      return obj;
    });
  };

  Table.prototype.serialize = function (record) {
    var row = [];
    var headers = this.headers;
    for (var i = 0; i < headers.length; i++) {
      var header = headers[i];
      row.push(typeof record[header] === 'undefined' ? '' : record[header]);
    }
    return row;
  };

  Table.prototype.validateRecord = function (record) {
    var validators = this.validators;
    var keys = Object.keys(validators);
    for (var i = 0; i < keys.length; i++) {
      var key = keys[i];
      var validator = validators[key];
      if (typeof validator === 'function') {
        var result = validator(record[key], record);
        if (result === false) {
          throw new Error('Validation failed for column "' + key + '"');
        }
      }
    }
  };

  Table.prototype.applyDefaults = function (record, isInsert) {
    var defaults = this.defaults;
    var keys = Object.keys(defaults);
    for (var i = 0; i < keys.length; i++) {
      var key = keys[i];
      if (typeof record[key] === 'undefined' || record[key] === null || record[key] === '') {
        if (typeof defaults[key] === 'function') {
          record[key] = defaults[key](record, isInsert);
        } else {
          record[key] = defaults[key];
        }
      }
    }
    return record;
  };

  Table.prototype.ensureId = function (record) {
    if (!this.idColumn) return record;
    if (!record[this.idColumn]) {
      record[this.idColumn] = Utilities.getUuid();
    }
    return record;
  };

  Table.prototype.touchTimestamps = function (record, isInsert) {
    if (!this.timestamps) return record;
    var now = new Date();
    if (isInsert && this.timestamps.created && !record[this.timestamps.created]) {
      record[this.timestamps.created] = now;
    }
    if (this.timestamps.updated) {
      record[this.timestamps.updated] = now;
    }
    return record;
  };

  Table.prototype.invalidateCache = function () {
    try {
      CacheService.getScriptCache().remove(this.cacheKey);
    } catch (err) {
      logger.error('Failed to invalidate cache for table ' + this.name + ': ' + err);
    }
  };

  Table.prototype.read = function (options, context) {
    options = options || {};
    var prepared = this.prepareTenantOptions(options, context, true);
    var finalOptions = prepared.options;
    var useCache = finalOptions.cache !== false;
    var cache = CacheService.getScriptCache();
    var headers = this.headers;

    if (useCache) {
      var cached = cache.get(this.cacheKey);
      if (cached) {
        try {
          var parsed = JSON.parse(cached);
          return applyQueryOptions(parsed, headers, finalOptions);
        } catch (err) {
          logger.warn('Cache parse failed for table ' + this.name + ': ' + err);
        }
      }
    }

    var rows = this.sheetHandle.readAllRows();
    var objects = this.toObjects(rows);

    if (useCache) {
      try {
        cache.put(this.cacheKey, JSON.stringify(objects), this.cacheTTL);
      } catch (err) {
        logger.warn('Cache put failed for table ' + this.name + ': ' + err);
      }
    }

    return applyQueryOptions(objects, headers, finalOptions);
  };

  Table.prototype.project = function (columns, options, context) {
    if (!Array.isArray(columns) || !columns.length) {
      return this.read(options || {}, context);
    }

    var opts = options ? clone(options) : {};
    opts.columns = columns.slice();
    return this.read(opts, context);
  };

  Table.prototype.find = function (options, context) {
    return this.read(options || {}, context);
  };

  Table.prototype.findOne = function (where, context) {
    var options = { where: where, limit: 1 };
    var results = this.read(options, context);
    return results.length ? results[0] : null;
  };

  Table.prototype.findById = function (id, context) {
    if (!this.idColumn) return null;
    return this.findOne(createWhereClause(this.idColumn, id), context);
  };

  Table.prototype.insert = function (record, context) {

    if (!record || typeof record !== 'object') {
      throw new Error('Record must be an object for insert');
    }
    var copy = clone(record);
    var tenantAccess = this.getTenantAccess(context, false);
    this.enforceTenantOnRecord(copy, tenantAccess);
    this.ensureId(copy);
    this.applyDefaults(copy, true);
    this.touchTimestamps(copy, true);
    this.validateRecord(copy);

    var rowValues = this.serialize(copy);
    this.sheetHandle.appendRow(rowValues);
    this.invalidateCache();
    return copy;
  };

  Table.prototype.batchInsert = function (records, context) {
    if (!Array.isArray(records) || records.length === 0) {
      return [];
    }
    var sheet = this.sheetHandle.getSheet();
    var processed = [];
    var rows = [];
    var tenantAccess = this.getTenantAccess(context, false);

    for (var i = 0; i < records.length; i++) {
      var copy = clone(records[i]);
      this.enforceTenantOnRecord(copy, tenantAccess);
      this.ensureId(copy);
      this.applyDefaults(copy, true);
      this.touchTimestamps(copy, true);
      this.validateRecord(copy);
      processed.push(copy);
      rows.push(this.serialize(copy));
    }

    if (rows.length) {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rows.length, this.headers.length).setValues(rows);
      this.invalidateCache();
    }

    return processed;
  };

  Table.prototype.update = function (id, updates) {
    if (!this.idColumn) {
      throw new Error('Cannot update without idColumn configuration');
    }
    var sheet = this.sheetHandle.getSheet();
    var headers = this.headers;
    var idIndex = headers.indexOf(this.idColumn);
    if (idIndex === -1) {
      throw new Error('ID column ' + this.idColumn + ' not found in headers');
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return null;
    }

    var range = sheet.getRange(2, 1, lastRow - 1, headers.length);
    var values = range.getValues();

    var updatedRecord = null;
    var tenantAccess = this.getTenantAccess(context, false);

    for (var i = 0; i < values.length; i++) {
      if (String(values[i][idIndex]) === String(id)) {
        var record = {};
        for (var j = 0; j < headers.length; j++) {
          record[headers[j]] = values[i][j];
        }

        var existingRecord = clone(record);
        this.ensureExistingTenantAllowed(existingRecord, tenantAccess);

        Object.keys(updates || {}).forEach(function (key) {
          record[key] = updates[key];
        });

        this.enforceTenantOnRecord(record, tenantAccess);
        this.touchTimestamps(record, false);
        this.validateRecord(record);
        var serialized = this.serialize(record);
        range.getCell(i + 1, 1).offset(0, 0, 1, headers.length).setValues([serialized]);
        updatedRecord = record;
        break;
      }
    }

    if (updatedRecord) {
      this.invalidateCache();
    }
    return updatedRecord;
  };

  Table.prototype.upsert = function (where, updates, context) {
    var existing = this.findOne(where, context);
    if (existing) {
      var id = this.idColumn ? existing[this.idColumn] : null;
      if (id) {
        return this.update(id, updates, context);
      }
      var merged = clone(existing);
      Object.keys(updates || {}).forEach(function (key) {
        merged[key] = updates[key];
      });
      return this.insert(merged, context);
    }
    var insertRecord = clone(where || {});
    Object.keys(updates || {}).forEach(function (key) {
      insertRecord[key] = updates[key];
    });
    return this.insert(insertRecord, context);
  };

  Table.prototype.delete = function (id, context) {
    if (!this.idColumn) {
      throw new Error('Cannot delete without idColumn configuration');
    }
    var sheet = this.sheetHandle.getSheet();
    var headers = this.headers;
    var idIndex = headers.indexOf(this.idColumn);
    if (idIndex === -1) {
      throw new Error('ID column ' + this.idColumn + ' not found in headers');
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return false;
    }

    var range = sheet.getRange(2, 1, lastRow - 1, headers.length);
    var values = range.getValues();

    var tenantAccess = this.getTenantAccess(context, false);

    for (var i = 0; i < values.length; i++) {
      if (String(values[i][idIndex]) === String(id)) {
        var record = {};
        for (var j = 0; j < headers.length; j++) {
          record[headers[j]] = values[i][j];
        }
        this.ensureExistingTenantAllowed(record, tenantAccess);
        sheet.deleteRow(i + 2);
        this.invalidateCache();
        return true;
      }
    }
    return false;
  };

  Table.prototype.clear = function () {
    var sheet = this.sheetHandle.getSheet();
    var headers = this.headers;
    sheet.clearContents();
    if (headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
    this.invalidateCache();
  };

  Table.prototype.count = function (where, context) {
    var options = where ? { where: where } : {};
    return this.read(options, context).length;
  };

  Table.prototype.listColumns = function () {
    return this.headers.slice();
  };

  function normalizeHeaders(table, providedHeaders) {
    var headers = providedHeaders.slice();
    if (table.idColumn && headers.indexOf(table.idColumn) === -1) {
      headers.unshift(table.idColumn);
    }
    if (table.timestamps) {
      if (table.timestamps.created && headers.indexOf(table.timestamps.created) === -1) {
        headers.push(table.timestamps.created);
      }
      if (table.timestamps.updated && headers.indexOf(table.timestamps.updated) === -1) {
        headers.push(table.timestamps.updated);
      }
    }
    Object.keys(table.defaults).forEach(function (key) {
      if (headers.indexOf(key) === -1) {
        headers.push(key);
      }
    });
    Object.keys(table.validators).forEach(function (key) {
      if (headers.indexOf(key) === -1) {
        headers.push(key);
      }
    });
    return headers;
  }

  function applyQueryOptions(rows, headers, options) {
    var filtered = rows.slice();

    if (options.where && isObject(options.where)) {
      filtered = filtered.filter(function (row) {
        return matchesWhere(row, options.where);
      });
    }

    if (options.filter && typeof options.filter === 'function') {
      filtered = filtered.filter(options.filter);
    }

    if (options.map && typeof options.map === 'function') {
      filtered = filtered.map(options.map);
    }

    if (options.sortBy) {
      var sortKey = options.sortBy;
      var descending = !!options.sortDesc;
      filtered.sort(function (a, b) {
        var av = a[sortKey];
        var bv = b[sortKey];
        if (av === bv) return 0;
        if (av === undefined || av === null || av === '') return descending ? 1 : -1;
        if (bv === undefined || bv === null || bv === '') return descending ? -1 : 1;
        if (av > bv) return descending ? -1 : 1;
        if (av < bv) return descending ? 1 : -1;
        return 0;
      });
    }

    var offset = options.offset || 0;
    var limit = typeof options.limit === 'number' ? options.limit : null;
    if (offset || limit !== null) {
      var start = offset;
      var end = limit !== null ? offset + limit : filtered.length;
      filtered = filtered.slice(start, end);
    }

    if (options.columns && Array.isArray(options.columns) && options.columns.length) {
      filtered = filtered.map(function (row) {
        var projected = {};
        options.columns.forEach(function (col) {
          projected[col] = row[col];
        });
        return projected;
      });
    }

    return filtered;
  }

  function matchesWhere(row, where) {
    var keys = Object.keys(where);
    for (var i = 0; i < keys.length; i++) {
      var key = keys[i];
      var expected = where[key];
      var actual = row[key];
      if (expected instanceof RegExp) {
        if (!expected.test(String(actual || ''))) {
          return false;
        }
      } else if (isObject(expected)) {
        if (!evaluateWhereOperator(actual, expected)) {
          return false;
        }
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

  function createWhereClause(column, value) {
    var where = {};
    where[column] = value;
    return where;
  }

  function ensureTable(name, config) {
    if (!tables[name]) {
      tables[name] = new Table(name, config || {});
    } else if (config) {
      tables[name] = new Table(name, mergeConfigs(tables[name], config));
    }
    return tables[name];
  }

  function mergeConfigs(existingTable, config) {
    var merged = {
      idColumn: existingTable.idColumn,
      cacheTTL: typeof config.cacheTTL === 'number' ? config.cacheTTL : existingTable.cacheTTL,
      timestamps: existingTable.timestamps,
      defaults: Object.assign({}, existingTable.defaults || {}, config.defaults || {}),
      validators: Object.assign({}, existingTable.validators || {}, config.validators || {}),
      headers: config.headers || existingTable.headers,
      columns: config.columns || existingTable.headers,
      tenantColumn: Object.prototype.hasOwnProperty.call(config, 'tenantColumn')
        ? config.tenantColumn
        : existingTable.tenantColumn,
      requireTenant: Object.prototype.hasOwnProperty.call(config, 'requireTenant')
        ? config.requireTenant
        : existingTable.requireTenant,
      allowGlobalTenantBypass: Object.prototype.hasOwnProperty.call(config, 'allowGlobalTenantBypass')
        ? config.allowGlobalTenantBypass
        : existingTable.allowGlobalTenantBypass
    };

    if (Object.prototype.hasOwnProperty.call(config, 'idColumn')) {
      merged.idColumn = config.idColumn;
    }

    if (config.timestamps === false) {
      merged.timestamps = false;
    } else if (config.timestamps) {
      merged.timestamps = config.timestamps;
    }

    return merged;
  }

  function mergeContextObjects(base, extra) {
    var result = {};
    Object.keys(base || {}).forEach(function (key) {
      result[key] = base[key];
    });
    Object.keys(extra || {}).forEach(function (key) {
      if (Array.isArray(extra[key])) {
        result[key] = extra[key].slice();
      } else {
        result[key] = extra[key];
      }
    });
    return result;
  }

  function ScopedTable(table, context) {
    this.table = table;
    this.context = table && typeof table.normalizeContext === 'function'
      ? table.normalizeContext(context)
      : normalizeTenantContext(context);
  }

  ScopedTable.prototype.getContext = function () {
    return this.context || null;
  };

  ScopedTable.prototype.withContext = function (extraContext) {
    var normalized = normalizeTenantContext(extraContext) || {};
    var merged = mergeContextObjects(this.context || {}, normalized);
    return new ScopedTable(this.table, merged);
  };

  var PROXIED_METHODS = ['read', 'project', 'find', 'findOne', 'findById', 'insert', 'batchInsert', 'update', 'upsert', 'delete', 'count'];
  PROXIED_METHODS.forEach(function (method) {
    if (typeof Table.prototype[method] !== 'function') return;
    ScopedTable.prototype[method] = function () {
      var args = Array.prototype.slice.call(arguments);
      args.push(this.context);
      return this.table[method].apply(this.table, args);
    };
  });

  ScopedTable.prototype.clear = function () {
    return this.table.clear.apply(this.table, arguments);
  };

  ScopedTable.prototype.listColumns = function () {
    return this.table.listColumns();
  };

  var DatabaseManager = {
    defineTable: function (name, config) {
      if (!name) {
        throw new Error('Table name is required');
      }
      return ensureTable(name, config || {});
    },
    table: function (name, context) {
      if (!name) {
        throw new Error('Table name is required');
      }
      var table = ensureTable(name);
      if (typeof context !== 'undefined' && context !== null) {
        return table.withContext(context);
      }
      return table;
    },
    listTables: function () {
      return Object.keys(tables);
    },
    dropTableCache: function (name) {
      if (tables[name]) {
        tables[name].invalidateCache();
      }
    },
    tenant: function (context) {
      var normalized = normalizeTenantContext(context) || context || null;
      return {
        context: normalized,
        table: function (name) {
          return DatabaseManager.table(name, normalized);
        },
        withContext: function (extra) {
          var merged = mergeContextObjects(normalized || {}, normalizeTenantContext(extra) || extra || {});
          return DatabaseManager.tenant(merged);
        }
      };
    },
    withContext: function (context) {
      return this.tenant(context);
    },
    normalizeContext: normalizeTenantContext
  };

  global.DatabaseManager = DatabaseManager;

  function exposeHelper(name, fn) {
    var previous = (typeof global[name] === 'function') ? global[name].bind(global) : null;
    var wrapped = function () {
      try {
        return fn.apply(this, arguments);
      } catch (err) {
        if (previous) {
          try { return previous.apply(this, arguments); } catch (fallbackErr) {
            if (typeof safeWriteError === 'function') {
              try { safeWriteError(name + 'Fallback', fallbackErr); } catch (_) { }
            }
          }
        }
        throw err;
      }
    };
    wrapped.previous = previous;
    global[name] = wrapped;
  }

  exposeHelper('defineTable', DatabaseManager.defineTable);
  exposeHelper('getTable', DatabaseManager.table);

})(typeof globalThis !== 'undefined' ? globalThis : this);
