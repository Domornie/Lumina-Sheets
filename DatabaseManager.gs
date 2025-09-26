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

  Table.prototype.read = function (options) {
    options = options || {};
    var useCache = options.cache !== false;
    var cache = CacheService.getScriptCache();
    var headers = this.headers;

    if (useCache) {
      var cached = cache.get(this.cacheKey);
      if (cached) {
        try {
          var parsed = JSON.parse(cached);
          return applyQueryOptions(parsed, headers, options);
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

    return applyQueryOptions(objects, headers, options);
  };

  Table.prototype.find = function (options) {
    return this.read(options || {});
  };

  Table.prototype.findOne = function (where) {
    var options = { where: where, limit: 1 };
    var results = this.read(options);
    return results.length ? results[0] : null;
  };

  Table.prototype.findById = function (id) {
    if (!this.idColumn) return null;
    return this.findOne(createWhereClause(this.idColumn, id));
  };

  Table.prototype.insert = function (record) {
    if (!record || typeof record !== 'object') {
      throw new Error('Record must be an object for insert');
    }
    var copy = clone(record);
    this.ensureId(copy);
    this.applyDefaults(copy, true);
    this.touchTimestamps(copy, true);
    this.validateRecord(copy);

    var rowValues = this.serialize(copy);
    this.sheetHandle.appendRow(rowValues);
    this.invalidateCache();
    return copy;
  };

  Table.prototype.batchInsert = function (records) {
    if (!Array.isArray(records) || records.length === 0) {
      return [];
    }
    var sheet = this.sheetHandle.getSheet();
    var processed = [];
    var rows = [];

    for (var i = 0; i < records.length; i++) {
      var copy = clone(records[i]);
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
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][idIndex]) === String(id)) {
        var record = {};
        for (var j = 0; j < headers.length; j++) {
          record[headers[j]] = values[i][j];
        }

        Object.keys(updates || {}).forEach(function (key) {
          record[key] = updates[key];
        });

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

  Table.prototype.upsert = function (where, updates) {
    var existing = this.findOne(where);
    if (existing) {
      var id = this.idColumn ? existing[this.idColumn] : null;
      if (id) {
        return this.update(id, updates);
      }
      var merged = clone(existing);
      Object.keys(updates || {}).forEach(function (key) {
        merged[key] = updates[key];
      });
      return this.insert(merged);
    }
    var insertRecord = clone(where || {});
    Object.keys(updates || {}).forEach(function (key) {
      insertRecord[key] = updates[key];
    });
    return this.insert(insertRecord);
  };

  Table.prototype.delete = function (id) {
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

    for (var i = 0; i < values.length; i++) {
      if (String(values[i][idIndex]) === String(id)) {
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

  Table.prototype.count = function (where) {
    var options = where ? { where: where } : {};
    return this.read(options).length;
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
      columns: config.columns || existingTable.headers
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

  var DatabaseManager = {
    defineTable: function (name, config) {
      if (!name) {
        throw new Error('Table name is required');
      }
      return ensureTable(name, config || {});
    },
    table: function (name) {
      if (!name) {
        throw new Error('Table name is required');
      }
      return ensureTable(name);
    },
    listTables: function () {
      return Object.keys(tables);
    },
    dropTableCache: function (name) {
      if (tables[name]) {
        tables[name].invalidateCache();
      }
    }
  };

  global.DatabaseManager = DatabaseManager;
  global.defineTable = DatabaseManager.defineTable;
  global.getTable = DatabaseManager.table;

})(typeof globalThis !== 'undefined' ? globalThis : this);
