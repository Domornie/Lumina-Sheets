(function (global) {
  if (global.SheetsDbApi) {
    return;
  }

  function jsonResponse(payload, status) {
    var output = ContentService.createTextOutput(JSON.stringify(payload));
    output.setMimeType(ContentService.MimeType.JSON);
    if (status) {
      try {
        output.setResponseCode(status);
      } catch (err) {
        // Apps Script setResponseCode available in advanced services only; ignore for compatibility
      }
    }
    return output;
  }

  function parseFilters(parameter) {
    if (!parameter) return [];
    try {
      var decoded = JSON.parse(parameter);
      if (Array.isArray(decoded)) {
        return decoded.filter(function (item) {
          return item && item.field && item.operator;
        });
      }
    } catch (err) {
      return [];
    }
    return [];
  }

  function parseCursor(parameter) {
    if (!parameter) return null;
    try {
      return parameter;
    } catch (err) {
      return null;
    }
  }

  function readBody(e) {
    if (!e || !e.postData || !e.postData.contents) {
      return {};
    }
    try {
      return JSON.parse(e.postData.contents);
    } catch (err) {
      throw new Error('Invalid JSON body');
    }
  }

  function loadApiKeyMap() {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty('SHEETS_DB_API_KEYS');
    if (!raw) return {};
    try {
      var parsed = JSON.parse(raw);
      return parsed && typeof parsed === 'object' ? parsed : {};
    } catch (err) {
      return {};
    }
  }

  var ROLE_PERMISSIONS = {
    admin: ['read', 'write', 'delete', 'manage'],
    writer: ['read', 'write'],
    reader: ['read']
  };

  function resolveRoleFromKey(apiKey) {
    if (!apiKey) return null;
    var map = loadApiKeyMap();
    return map[apiKey] || null;
  }

  function requirePermission(role, permission) {
    if (!permission) return true;
    if (!role) {
      throw new Error('Unauthorized: missing API key');
    }
    var allowed = ROLE_PERMISSIONS[role];
    if (!allowed || allowed.indexOf(permission) === -1) {
      throw new Error('Forbidden: insufficient permissions');
    }
    return true;
  }

  function normalizeFilters(params) {
    var filters = [];
    if (params.filter) {
      filters = parseFilters(params.filter);
    }
    if (params.filterField && params.filterValue !== undefined) {
      filters.push({
        field: params.filterField,
        operator: params.filterOperator || '=',
        value: params.filterValue
      });
    }
    return filters;
  }

  function safeListTables() {
    if (!global.SheetsDB) {
      throw new Error('SheetsDB is not initialized');
    }
    return global.SheetsDB.listTables();
  }

  function handleGet(e) {
    try {
      var params = e && e.parameter ? e.parameter : {};
      var apiKey = params.apiKey || '';
      var role = resolveRoleFromKey(apiKey);
      var action = params.action || '';
      if (action === 'tables') {
        requirePermission(role, 'read');
        return jsonResponse({ tables: safeListTables() });
      }
      var tableName = params.table;
      if (!tableName) {
        throw new Error('Missing table parameter');
      }
      var table = SheetsDB.getTable(tableName);
      if (!table) {
        throw new Error('Unknown table: ' + tableName);
      }
      if (params.id) {
        requirePermission(role, 'read');
        var record = table.get(params.id, { includeDeleted: params.includeDeleted === 'true' });
        if (!record) {
          return jsonResponse({ error: 'Not found' }, 404);
        }
        return jsonResponse({ record: record });
      }
      if (action === 'schema') {
        requirePermission(role, 'read');
        return jsonResponse({
          table: tableName,
          version: table.version,
          columns: table.columns
        });
      }
      requirePermission(role, 'read');
      var filters = normalizeFilters(params);
      var list = table.list({
        includeDeleted: params.includeDeleted === 'true',
        filters: filters,
        limit: params.limit ? Number(params.limit) : undefined,
        offset: params.offset ? Number(params.offset) : undefined,
        cursor: parseCursor(params.cursor)
      });
      return jsonResponse(list);
    } catch (err) {
      return jsonResponse({ error: err.message || String(err) }, 400);
    }
  }

  function handlePost(e) {
    try {
      var body = readBody(e);
      var params = e && e.parameter ? e.parameter : {};
      var apiKey = body.apiKey || params.apiKey || '';
      var role = resolveRoleFromKey(apiKey);
      var action = (body.action || params.action || '').toLowerCase();
      var tableName = body.table || params.table;
      if (!tableName) {
        throw new Error('Missing table parameter');
      }
      var table = SheetsDB.getTable(tableName);
      if (!table) {
        throw new Error('Unknown table: ' + tableName);
      }
      var actor = body.actor || params.actor || '';
      if (action === 'create') {
        requirePermission(role, 'write');
        var createResult = table.create(body.record || {}, {
          actor: actor,
          metadata: body.metadata,
          idempotencyKey: body.idempotencyKey
        });
        return jsonResponse(createResult, 201);
      }
      if (action === 'update') {
        requirePermission(role, 'write');
        if (!body.id) {
          throw new Error('Missing id for update');
        }
        var updateResult = table.update(body.id, body.record || {}, {
          actor: actor,
          metadata: body.metadata,
          expectedUpdatedAt: body.expectedUpdatedAt,
          allowDeleted: body.allowDeleted === true
        });
        return jsonResponse(updateResult);
      }
      if (action === 'delete') {
        requirePermission(role, 'delete');
        if (!body.id) {
          throw new Error('Missing id for delete');
        }
        if (body.hard === true) {
          return jsonResponse(table.hardDelete(body.id, {
            actor: actor,
            metadata: body.metadata
          }));
        }
        return jsonResponse(table.softDelete(body.id, {
          actor: actor,
          metadata: body.metadata
        }));
      }
      if (action === 'restore') {
        requirePermission(role, 'write');
        if (!body.id) {
          throw new Error('Missing id for restore');
        }
        return jsonResponse(table.restore(body.id, {
          actor: actor,
          metadata: body.metadata
        }));
      }
      if (action === 'archive') {
        requirePermission(role, 'manage');
        if (!body.cutoff) {
          throw new Error('Missing cutoff for archive');
        }
        return jsonResponse({ archived: table.archiveOlderThan(body.cutoff) });
      }
      if (action === 'backup') {
        requirePermission(role, 'manage');
        return jsonResponse({ backupSheet: table.backup() });
      }
      if (action === 'maintenance') {
        requirePermission(role, 'manage');
        return jsonResponse(SheetsDB.runMaintenance());
      }
      throw new Error('Unknown or unsupported action: ' + action);
    } catch (err) {
      return jsonResponse({ error: err.message || String(err) }, 400);
    }
  }

  global.SheetsDbApi = {
    handleGet: handleGet,
    handlePost: handlePost
  };
})(this);
