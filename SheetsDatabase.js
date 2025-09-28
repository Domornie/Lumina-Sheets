(function (global) {
  if (global.SheetsDB) {
    return;
  }

  var CONFIG = {
    auditSheetName: 'AuditLogs',
    migrationSheetName: '__SchemaMigrations',
    indexPrefix: '__idx',
    archivePrefix: '__Archive',
    registryKeyPrefix: 'SHEETS_DB_SCHEMA_',
    registrySheetName: '__SchemaRegistry',
    idempotencySheetName: '__IdempotencyKeys',
    schemaVersionProperty: 'schemaVersion',
    timezone: 'Etc/UTC',
    defaultPrimaryKey: 'id',
    createdAtColumn: 'createdAt',
    updatedAtColumn: 'updatedAt',
    deletedAtColumn: 'deletedAt'
  };

  function nowIsoString() {
    return Utilities.formatDate(new Date(), CONFIG.timezone, "yyyy-MM-dd'T'HH:mm:ss'Z'");
  }

  function toIso(value) {
    if (!value && value !== 0) {
      return '';
    }
    if (Object.prototype.toString.call(value) === '[object Date]') {
      return Utilities.formatDate(value, CONFIG.timezone, "yyyy-MM-dd'T'HH:mm:ss'Z'");
    }
    if (typeof value === 'number') {
      return Utilities.formatDate(new Date(value), CONFIG.timezone, "yyyy-MM-dd'T'HH:mm:ss'Z'");
    }
    var parsed = new Date(value);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, CONFIG.timezone, "yyyy-MM-dd'T'HH:mm:ss'Z'");
    }
    throw new Error('Unable to coerce value to ISO timestamp: ' + value);
  }

  function toDate(value) {
    if (!value && value !== 0) return null;
    if (Object.prototype.toString.call(value) === '[object Date]') {
      return value;
    }
    if (typeof value === 'number') {
      return new Date(value);
    }
    var parsed = new Date(value);
    if (!isNaN(parsed.getTime())) {
      return parsed;
    }
    return null;
  }

  function deepClone(value) {
    if (value === null || typeof value !== 'object') {
      return value;
    }
    if (Array.isArray(value)) {
      return value.map(deepClone);
    }
    var copy = {};
    Object.keys(value).forEach(function (key) {
      copy[key] = deepClone(value[key]);
    });
    return copy;
  }

  function valuesEqual(a, b) {
    if (a === b) return true;
    if (a === null || b === null || typeof a !== typeof b) return false;
    if (Array.isArray(a) && Array.isArray(b)) {
      if (a.length !== b.length) return false;
      for (var i = 0; i < a.length; i++) {
        if (!valuesEqual(a[i], b[i])) return false;
      }
      return true;
    }
    if (typeof a === 'object') {
      var keysA = Object.keys(a);
      var keysB = Object.keys(b);
      if (keysA.length !== keysB.length) return false;
      for (var j = 0; j < keysA.length; j++) {
        var key = keysA[j];
        if (!valuesEqual(a[key], b[key])) return false;
      }
      return true;
    }
    return String(a) === String(b);
  }

  function hashId(prefix, counter) {
    var padded = String(counter);
    while (padded.length < 6) {
      padded = '0' + padded;
    }
    return (prefix || '').toUpperCase() + padded;
  }

  function ensureSupportSheet(name, headers) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Active spreadsheet not available');
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    var lastCol = sheet.getLastColumn();
    if (lastCol === 0 && headers && headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    return sheet;
  }

  function getDocumentProperties() {
    return PropertiesService.getDocumentProperties();
  }

  function readSchemaVersion(tableName) {
    var props = getDocumentProperties();
    var raw = props.getProperty(CONFIG.registryKeyPrefix + tableName);
    if (!raw) return null;
    try {
      return JSON.parse(raw);
    } catch (err) {
      return null;
    }
  }

  function writeSchemaVersion(tableName, data) {
    var props = getDocumentProperties();
    props.setProperty(CONFIG.registryKeyPrefix + tableName, JSON.stringify(data));
  }

  function appendMigrationLog(entry) {
    var sheet = ensureSupportSheet(CONFIG.migrationSheetName, [
      'timestamp',
      'table',
      'fromVersion',
      'toVersion',
      'description'
    ]);
    sheet.appendRow([
      nowIsoString(),
      entry.table || '',
      entry.fromVersion || '',
      entry.toVersion || '',
      entry.description || ''
    ]);
  }

  function auditLog(action, table, id, actor, before, after, metadata) {
    var sheet = ensureSupportSheet(CONFIG.auditSheetName, [
      'timestamp',
      'action',
      'table',
      'recordId',
      'actor',
      'before',
      'after',
      'metadata'
    ]);
    sheet.appendRow([
      nowIsoString(),
      action,
      table,
      id || '',
      actor || '',
      before ? JSON.stringify(before) : '',
      after ? JSON.stringify(after) : '',
      metadata ? JSON.stringify(metadata) : ''
    ]);
  }

  function ensureIdempotencyKey(key, payload) {
    if (!key) return null;
    var sheet = ensureSupportSheet(CONFIG.idempotencySheetName, [
      'key',
      'table',
      'action',
      'recordId',
      'response',
      'createdAt'
    ]);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        var existing = data[i][4];
        return existing ? JSON.parse(existing) : null;
      }
    }
    sheet.appendRow([
      key,
      payload.table || '',
      payload.action || '',
      payload.recordId || '',
      payload.response ? JSON.stringify(payload.response) : '',
      nowIsoString()
    ]);
    return null;
  }

  function recordIdempotencyResult(key, payload) {
    if (!key) return;
    var sheet = ensureSupportSheet(CONFIG.idempotencySheetName, [
      'key',
      'table',
      'action',
      'recordId',
      'response',
      'createdAt'
    ]);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 5).setValue(payload.response ? JSON.stringify(payload.response) : '');
        sheet.getRange(i + 1, 3).setValue(payload.action || '');
        sheet.getRange(i + 1, 4).setValue(payload.recordId || '');
        return;
      }
    }
    sheet.appendRow([
      key,
      payload.table || '',
      payload.action || '',
      payload.recordId || '',
      payload.response ? JSON.stringify(payload.response) : '',
      nowIsoString()
    ]);
  }

  function coerceValue(value, column) {
    if (!column) return value;
    var type = column.type || 'string';
    if (value === null || typeof value === 'undefined' || value === '') {
      if (column.required && !column.nullable && !column.primaryKey) {
        throw new Error('Missing required field "' + column.name + '"');
      }
      return column.nullable ? null : '';
    }

    switch (type) {
      case 'number':
        var num = Number(value);
        if (isNaN(num)) {
          throw new Error('Invalid number for field "' + column.name + '"');
        }
        if (column.min !== undefined && num < column.min) {
          throw new Error('Value for "' + column.name + '" below minimum ' + column.min);
        }
        if (column.max !== undefined && num > column.max) {
          throw new Error('Value for "' + column.name + '" above maximum ' + column.max);
        }
        return num;
      case 'boolean':
        if (typeof value === 'boolean') return value;
        if (value === 'true' || value === '1' || value === 1) return true;
        if (value === 'false' || value === '0' || value === 0) return false;
        throw new Error('Invalid boolean for field "' + column.name + '"');
      case 'date':
      case 'datetime':
      case 'timestamp':
        return toIso(value);
      case 'enum':
        if (!column.allowedValues || column.allowedValues.indexOf(value) === -1) {
          throw new Error('Value for "' + column.name + '" must be one of: ' + (column.allowedValues || []).join(', '));
        }
        return value;
      case 'json':
        if (typeof value === 'object') {
          return JSON.stringify(value);
        }
        try {
          JSON.parse(value);
          return value;
        } catch (err) {
          throw new Error('Invalid JSON for field "' + column.name + '"');
        }
      case 'string':
      default:
        var str = String(value);
        if (column.minLength && str.length < column.minLength) {
          throw new Error('Value for "' + column.name + '" shorter than minimum length ' + column.minLength);
        }
        if (column.maxLength && str.length > column.maxLength) {
          throw new Error('Value for "' + column.name + '" exceeds maximum length ' + column.maxLength);
        }
        if (column.pattern && !(new RegExp(column.pattern).test(str))) {
          throw new Error('Value for "' + column.name + '" does not match pattern ' + column.pattern);
        }
        return str;
    }
  }

  function normalizeColumn(column) {
    var copy = Object.assign({}, column);
    copy.name = String(copy.name);
    if (!copy.type) copy.type = 'string';
    if (copy.required === undefined) copy.required = false;
    if (copy.nullable === undefined) copy.nullable = !copy.required;
    if (copy.primaryKey) {
      copy.required = true;
      copy.nullable = false;
    }
    return copy;
  }

  function ensureTimestamps(columns) {
    var columnNames = columns.map(function (c) { return c.name; });
    if (columnNames.indexOf(CONFIG.createdAtColumn) === -1) {
      columns.push(normalizeColumn({
        name: CONFIG.createdAtColumn,
        type: 'timestamp',
        required: true
      }));
    }
    if (columnNames.indexOf(CONFIG.updatedAtColumn) === -1) {
      columns.push(normalizeColumn({
        name: CONFIG.updatedAtColumn,
        type: 'timestamp',
        required: true
      }));
    }
    if (columnNames.indexOf(CONFIG.deletedAtColumn) === -1) {
      columns.push(normalizeColumn({
        name: CONFIG.deletedAtColumn,
        type: 'timestamp',
        required: false,
        nullable: true
      }));
    }
    return columns;
  }

  function Table(schema) {
    this.name = schema.name;
    this.primaryKey = schema.primaryKey || CONFIG.defaultPrimaryKey;
    this.idPrefix = schema.idPrefix || (this.primaryKey.toUpperCase().substring(0, 3) + '_');
    this.version = schema.version || 1;
    this.columns = (schema.columns || []).map(normalizeColumn);
    if (this.columns.filter(function (c) { return c.primaryKey; }).length === 0) {
      this.columns.unshift(normalizeColumn({
        name: this.primaryKey,
        type: 'string',
        required: true,
        primaryKey: true
      }));
    }
    this.columns = ensureTimestamps(this.columns);
    this.indexes = schema.indexes || [];
    this.uniques = this.columns.filter(function (c) { return !!c.unique; }).map(function (c) { return c.name; });
    this.references = {};
    var self = this;
    (schema.columns || []).forEach(function (col) {
      if (col.references) {
        self.references[col.name] = col.references;
      }
    });
    this.sheet = null;
    this.columnIndex = {};
    this.headers = this.columns.map(function (c, idx) {
      self.columnIndex[c.name] = idx;
      return c.name;
    });
    this.archiveAfterDays = schema.archiveAfterDays || null;
    this.retentionDays = schema.retentionDays || null;
    this.registered = false;
  }

  Table.prototype.getSheet = function () {
    if (this.sheet) return this.sheet;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Active spreadsheet not available');
    var sheet = ss.getSheetByName(this.name);
    if (!sheet) {
      sheet = ss.insertSheet(this.name);
    }
    var lastCol = sheet.getLastColumn();
    if (lastCol === 0) {
      sheet.getRange(1, 1, 1, this.headers.length).setValues([this.headers]);
    } else {
      var existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      var updates = [];
      for (var i = 0; i < this.headers.length; i++) {
        if (existingHeaders[i] !== this.headers[i]) {
          updates.push(this.headers[i]);
        }
      }
      if (lastCol < this.headers.length) {
        sheet.getRange(1, 1, 1, this.headers.length).setValues([this.headers]);
      }
      if (updates.length) {
        sheet.getRange(1, 1, 1, this.headers.length).setValues([this.headers]);
      }
    }
    this.sheet = sheet;
    return this.sheet;
  };

  Table.prototype.ensureRegistered = function () {
    if (this.registered) return;
    this.getSheet();
    var schemaMeta = readSchemaVersion(this.name) || {};
    if (schemaMeta.version !== this.version) {
      appendMigrationLog({
        table: this.name,
        fromVersion: schemaMeta.version || '',
        toVersion: this.version,
        description: 'Auto-aligned headers for schema registration'
      });
      writeSchemaVersion(this.name, {
        version: this.version,
        headers: this.headers
      });
    }
    this.registered = true;
  };

  Table.prototype.getAllRows = function () {
    this.ensureRegistered();
    var sheet = this.getSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var range = sheet.getRange(2, 1, lastRow - 1, this.headers.length);
    var values = range.getValues();
    return values.map(function (row) {
      return row.slice();
    });
  };

  Table.prototype.deserialize = function (row) {
    var obj = {};
    for (var i = 0; i < this.headers.length; i++) {
      var header = this.headers[i];
      var column = this.columns[i];
      var value = row[i];
      if (column.type === 'number' && value !== '') {
        obj[header] = Number(value);
      } else if ((column.type === 'timestamp' || column.type === 'datetime' || column.type === 'date') && value !== '') {
        obj[header] = typeof value === 'string' ? value : toIso(value);
      } else if (column.type === 'boolean' && value !== '') {
        obj[header] = value === true || value === 'true';
      } else if (column.type === 'json' && value) {
        try {
          obj[header] = JSON.parse(value);
        } catch (_) {
          obj[header] = value;
        }
      } else {
        obj[header] = value;
      }
    }
    return obj;
  };

  Table.prototype.serialize = function (record) {
    var row = new Array(this.headers.length);
    for (var i = 0; i < this.headers.length; i++) {
      var column = this.columns[i];
      var value = record[this.headers[i]];
      if (value === undefined || value === null) {
        row[i] = '';
        continue;
      }
      if (column.type === 'json' && typeof value === 'object') {
        row[i] = JSON.stringify(value);
      } else {
        row[i] = value;
      }
    }
    return row;
  };

  Table.prototype.generateId = function () {
    var props = getDocumentProperties();
    var key = 'SHEETS_DB_SEQ_' + this.name;
    var counter = Number(props.getProperty(key) || '0') + 1;
    props.setProperty(key, String(counter));
    return hashId(this.idPrefix, counter);
  };

  Table.prototype.validateRecord = function (record, options) {
    options = options || {};
    var sanitized = {};
    var keys = Object.keys(record || {});
    var self = this;
    keys.forEach(function (key) {
      if (self.columnIndex[key] === undefined) {
        throw new Error('Unknown field "' + key + '" for table ' + self.name);
      }
    });
    this.columns.forEach(function (column) {
      var value = record[column.name];
      if ((value === undefined || value === null || value === '') && column.defaultValue !== undefined && !options.forUpdate) {
        value = typeof column.defaultValue === 'function' ? column.defaultValue() : column.defaultValue;
      }
      if (column.primaryKey && !value && !options.forUpdate) {
        value = self.generateId();
      }
      if (column.name === CONFIG.createdAtColumn && !value) {
        value = nowIsoString();
      }
      if (column.name === CONFIG.updatedAtColumn) {
        value = nowIsoString();
      }
      if (column.name === CONFIG.deletedAtColumn && !value) {
        value = '';
      }
      var coerced = coerceValue(value, column);
      sanitized[column.name] = coerced;
    });

    var primaryId = sanitized[this.primaryKey];
    if (!primaryId) {
      throw new Error('Record missing primary key ' + this.primaryKey);
    }

    if (!options.skipUniques) {
      this.ensureUniqueConstraints(sanitized, options.currentRowIndex, options.existingRecord);
    }
    this.ensureForeignKeys(sanitized);
    return sanitized;
  };

  Table.prototype.ensureUniqueConstraints = function (record, currentRowIndex, existingRecord) {
    var rows = this.getAllRows();
    var self = this;
    this.uniques.forEach(function (columnName) {
      var value = record[columnName];
      if (!value && value !== 0) return;
      for (var i = 0; i < rows.length; i++) {
        if (currentRowIndex !== undefined && i === currentRowIndex) continue;
        var other = self.deserialize(rows[i]);
        if (other[columnName] === value && other[self.primaryKey] !== record[self.primaryKey] && !other[CONFIG.deletedAtColumn]) {
          throw new Error('Field "' + columnName + '" must be unique. Duplicate found for value ' + value);
        }
      }
    });
  };

  Table.prototype.ensureForeignKeys = function (record) {
    var self = this;
    Object.keys(this.references).forEach(function (field) {
      var reference = self.references[field];
      var value = record[field];
      if (!value && reference.allowNull !== false) {
        return;
      }
      if (!value) {
        throw new Error('Field "' + field + '" requires a reference value');
      }
      var table = SheetsDB.getTable(reference.table);
      if (!table) {
        throw new Error('Referenced table "' + reference.table + '" not registered');
      }
      var foreign = table.get(value, { includeDeleted: false });
      if (!foreign) {
        throw new Error('Referenced record not found for "' + field + '" => ' + value);
      }
    });
  };

  Table.prototype.findRowIndexById = function (id) {
    var rows = this.getAllRows();
    for (var i = 0; i < rows.length; i++) {
      var obj = this.deserialize(rows[i]);
      if (obj[this.primaryKey] === id) {
        return i;
      }
    }
    return -1;
  };

  Table.prototype.get = function (id, options) {
    options = options || {};
    this.ensureRegistered();
    var sheet = this.getSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    var range = sheet.getRange(2, 1, lastRow - 1, this.headers.length);
    var values = range.getValues();
    for (var i = 0; i < values.length; i++) {
      var record = this.deserialize(values[i]);
      if (record[this.primaryKey] === id) {
        if (!options.includeDeleted && record[CONFIG.deletedAtColumn]) {
          return null;
        }
        return record;
      }
    }
    return null;
  };

  Table.prototype.list = function (options) {
    options = options || {};
    var rows = this.getAllRows();
    var includeDeleted = !!options.includeDeleted;
    var filters = options.filters || [];
    var cursor = options.cursor ? JSON.parse(Utilities.newBlob(Utilities.base64Decode(options.cursor)).getDataAsString()) : null;
    var limit = Math.min(options.limit || 50, 500);
    var offset = options.offset || 0;

    var self = this;
    function normalizeFilterValue(filter) {
      var idx = self.columnIndex[filter.field];
      if (idx === undefined) {
        return filter.value;
      }
      var column = self.columns[idx];
      if (!column) {
        return filter.value;
      }
      try {
        return coerceValue(filter.value, column);
      } catch (err) {
        return filter.value;
      }
    }
    var records = rows.map(function (row) { return self.deserialize(row); });
    records = records.filter(function (record) {
      if (!includeDeleted && record[CONFIG.deletedAtColumn]) {
        return false;
      }
      return filters.every(function (filter) {
        var value = record[filter.field];
        if (value === undefined) return false;
        var target = normalizeFilterValue(filter);
        switch (filter.operator) {
          case '=':
            return value === target;
          case 'contains':
            return String(value).toLowerCase().indexOf(String(target).toLowerCase()) !== -1;
          case '>':
            return value > target;
          case '<':
            return value < target;
          case '>=':
            return value >= target;
          case '<=':
            return value <= target;
          default:
            return false;
        }
      });
    });

    records.sort(function (a, b) {
      var updatedA = a[CONFIG.updatedAtColumn] || '';
      var updatedB = b[CONFIG.updatedAtColumn] || '';
      if (updatedA === updatedB) {
        if (a[self.primaryKey] < b[self.primaryKey]) return -1;
        if (a[self.primaryKey] > b[self.primaryKey]) return 1;
        return 0;
      }
      return updatedA < updatedB ? -1 : 1;
    });

    if (cursor) {
      records = records.filter(function (record) {
        if (record[CONFIG.updatedAtColumn] < cursor.updatedAt) return false;
        if (record[CONFIG.updatedAtColumn] === cursor.updatedAt) {
          return record[self.primaryKey] > cursor.id;
        }
        return true;
      });
    }

    var totalCount = records.length;
    if (offset) {
      records = records.slice(offset);
    }

    var limited = records.slice(0, limit);
    var nextCursor = null;
    if (records.length > limit) {
      var last = limited[limited.length - 1];
      if (last) {
        nextCursor = Utilities.base64Encode(JSON.stringify({
          updatedAt: last[CONFIG.updatedAtColumn] || '',
          id: last[self.primaryKey]
        }));
      }
    }

    return {
      records: limited,
      nextCursor: nextCursor,
      total: totalCount
    };
  };

  Table.prototype.updateIndexes = function () {
    if (!this.indexes || !this.indexes.length) return;
    var self = this;
    var rows = this.getAllRows();
    this.indexes.forEach(function (index) {
      var name = index.name || (self.name + '_' + index.field);
      var sheetName = CONFIG.indexPrefix + '_' + name;
      var sheet = ensureSupportSheet(sheetName, ['value', 'rowNumbers']);
      var map = {};
      for (var i = 0; i < rows.length; i++) {
        var record = self.deserialize(rows[i]);
        var value = record[index.field];
        if (!value && value !== 0) continue;
        if (!map[value]) {
          map[value] = [];
        }
        map[value].push(i + 2);
      }
      var entries = Object.keys(map);
      sheet.clearContents();
      if (!entries.length) {
        sheet.getRange(1, 1, 1, 2).setValues([['value', 'rowNumbers']]);
      } else {
        var values = [['value', 'rowNumbers']];
        entries.forEach(function (key) {
          values.push([key, map[key].join(',')]);
        });
        sheet.getRange(1, 1, values.length, 2).setValues(values);
      }
    });
  };

  Table.prototype.archiveOlderThan = function (isoDate) {
    var cutoff = toDate(isoDate);
    if (!cutoff) return 0;
    var rows = this.getAllRows();
    if (!rows.length) return 0;
    var archiveName = CONFIG.archivePrefix + '_' + this.name + '_' + Utilities.formatDate(cutoff, CONFIG.timezone, 'yyyyMM');
    var archiveSheet = ensureSupportSheet(archiveName, this.headers);
    var sheet = this.getSheet();
    var removed = 0;
    for (var i = rows.length - 1; i >= 0; i--) {
      var record = this.deserialize(rows[i]);
      var deletedAt = toDate(record[CONFIG.deletedAtColumn]);
      var updatedAt = toDate(record[CONFIG.updatedAtColumn]);
      if ((deletedAt && deletedAt <= cutoff) || (updatedAt && updatedAt <= cutoff)) {
        archiveSheet.appendRow(rows[i]);
        sheet.deleteRow(i + 2);
        removed++;
      }
    }
    if (removed) {
      this.updateIndexes();
    }
    return removed;
  };

  Table.prototype.purgeSoftDeleted = function (days) {
    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - (days || 30));
    return this.archiveOlderThan(cutoff);
  };

  Table.prototype.create = function (record, options) {
    options = options || {};
    var lock = LockService.getDocumentLock();
    if (!lock.tryLock(5000)) {
      throw new Error('Unable to obtain lock for create on ' + this.name);
    }
    try {
      this.ensureRegistered();
      var sanitized = this.validateRecord(record, { forUpdate: false });
      var id = sanitized[this.primaryKey];
      var idempotencyKey = options.idempotencyKey;
      if (idempotencyKey) {
        var existing = ensureIdempotencyKey(idempotencyKey, {
          table: this.name,
          action: 'create',
          recordId: id
        });
        if (existing) {
          return existing;
        }
      }
      var sheet = this.getSheet();
      sheet.appendRow(this.serialize(sanitized));
      this.updateIndexes();
      auditLog('create', this.name, id, options.actor || '', null, sanitized, options.metadata || {});
      var response = { record: sanitized };
      if (idempotencyKey) {
        recordIdempotencyResult(idempotencyKey, {
          table: this.name,
          action: 'create',
          recordId: id,
          response: response
        });
      }
      return response;
    } finally {
      lock.releaseLock();
    }
  };

  Table.prototype.update = function (id, updates, options) {
    options = options || {};
    var lock = LockService.getDocumentLock();
    if (!lock.tryLock(5000)) {
      throw new Error('Unable to obtain lock for update on ' + this.name);
    }
    try {
      this.ensureRegistered();
      var rowIndex = this.findRowIndexById(id);
      if (rowIndex === -1) {
        throw new Error('Record not found: ' + id);
      }
      var sheet = this.getSheet();
      var range = sheet.getRange(rowIndex + 2, 1, 1, this.headers.length);
      var existingRow = range.getValues()[0];
      var existingRecord = this.deserialize(existingRow);
      if (existingRecord[CONFIG.deletedAtColumn] && !options.allowDeleted) {
        throw new Error('Cannot update a deleted record. Restore first.');
      }
      if (options.expectedUpdatedAt && existingRecord[CONFIG.updatedAtColumn] !== options.expectedUpdatedAt) {
        throw new Error('Record has been modified since last read. Expected updatedAt ' + options.expectedUpdatedAt + ' but found ' + existingRecord[CONFIG.updatedAtColumn]);
      }
      var merged = Object.assign({}, existingRecord, updates);
      merged[this.primaryKey] = existingRecord[this.primaryKey];
      var sanitized = this.validateRecord(merged, {
        forUpdate: true,
        currentRowIndex: rowIndex,
        existingRecord: existingRecord
      });
      var serialized = this.serialize(sanitized);
      range.setValues([serialized]);
      this.updateIndexes();
      auditLog('update', this.name, id, options.actor || '', existingRecord, sanitized, options.metadata || {});
      return { record: sanitized };
    } finally {
      lock.releaseLock();
    }
  };

  Table.prototype.softDelete = function (id, options) {
    options = options || {};
    return this.update(id, {
      deletedAt: nowIsoString()
    }, Object.assign({}, options, {
      allowDeleted: true
    }));
  };

  Table.prototype.restore = function (id, options) {
    options = options || {};
    return this.update(id, {
      deletedAt: ''
    }, Object.assign({}, options, {
      allowDeleted: true
    }));
  };

  Table.prototype.hardDelete = function (id, options) {
    options = options || {};
    var lock = LockService.getDocumentLock();
    if (!lock.tryLock(5000)) {
      throw new Error('Unable to obtain lock for delete on ' + this.name);
    }
    try {
      this.ensureRegistered();
      var rowIndex = this.findRowIndexById(id);
      if (rowIndex === -1) {
        throw new Error('Record not found: ' + id);
      }
      var sheet = this.getSheet();
      var range = sheet.getRange(rowIndex + 2, 1, 1, this.headers.length);
      var existingRecord = this.deserialize(range.getValues()[0]);
      sheet.deleteRow(rowIndex + 2);
      this.updateIndexes();
      auditLog('delete', this.name, id, options.actor || '', existingRecord, null, options.metadata || {});
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  };

  Table.prototype.backup = function () {
    this.ensureRegistered();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = this.getSheet();
    var backupName = this.name + '_Backup_' + Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyyMMdd_HHmmss');
    sheet.copyTo(ss).setName(backupName);
    return backupName;
  };

  Table.prototype.refresh = function () {
    this.sheet = null;
    this.ensureRegistered();
  };

  var SheetsDB = {
    tables: {},
    defineTable: function (schema) {
      if (!schema || !schema.name) {
        throw new Error('Schema definition requires a name');
      }
      var table = new Table(schema);
      table.ensureRegistered();
      table.updateIndexes();
      this.tables[schema.name] = table;
      return table;
    },
    getTable: function (name) {
      return this.tables[name];
    },
    listTables: function () {
      return Object.keys(this.tables);
    },
    ensureSupportStructures: function () {
      ensureSupportSheet(CONFIG.auditSheetName, ['timestamp', 'action', 'table', 'recordId', 'actor', 'before', 'after', 'metadata']);
      ensureSupportSheet(CONFIG.migrationSheetName, ['timestamp', 'table', 'fromVersion', 'toVersion', 'description']);
      ensureSupportSheet(CONFIG.idempotencySheetName, ['key', 'table', 'action', 'recordId', 'response', 'createdAt']);
    },
    backupTables: function (names) {
      var results = [];
      var list = names || this.listTables();
      for (var i = 0; i < list.length; i++) {
        var table = this.getTable(list[i]);
        if (!table) continue;
        results.push({ table: list[i], backupSheet: table.backup() });
      }
      return results;
    },
    enforceRetention: function () {
      var results = [];
      var names = this.listTables();
      for (var i = 0; i < names.length; i++) {
        var table = this.getTable(names[i]);
        if (!table || !table.retentionDays) continue;
        results.push({
          table: names[i],
          archived: table.purgeSoftDeleted(table.retentionDays)
        });
      }
      return results;
    },
    runMaintenance: function () {
      this.ensureSupportStructures();
      var backups = this.backupTables();
      var retention = this.enforceRetention();
      return { backups: backups, retention: retention };
    }
  };

  SheetsDB.ensureSupportStructures();

  global.SheetsDB = SheetsDB;
})(this);
