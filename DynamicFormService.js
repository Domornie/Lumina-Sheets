/**
 * DynamicFormService.js
 *
 * Provides a lightweight, tenant-aware form builder that lets managers compose
 * Google Form-style questionnaires and map specific prompts directly to user
 * profile columns. Form responses always persist, even when no bindings are
 * supplied, so teams can review every answer submitted by a user alongside the
 * automatic profile updates.
 */

var DynamicFormService = (function (global) {
  var service = {};
  var USERS_TABLE = (typeof USERS_SHEET === 'string' && USERS_SHEET) ? USERS_SHEET : 'Users';
  var FORMS_TABLE = 'DynamicForms';
  var RESPONSES_TABLE = 'DynamicFormResponses';
  var RESPONSE_SHEET_BASE_HEADERS = (typeof DynamicFormUtilities !== 'undefined' && DynamicFormUtilities && Array.isArray(DynamicFormUtilities.RESPONSE_BASE_HEADERS))
    ? DynamicFormUtilities.RESPONSE_BASE_HEADERS.slice()
    : ['SubmissionID', 'FormID', 'FormName', 'SubmittedAt', 'UserID', 'UserName', 'SubmittedBy'];
  var safeConsole = (typeof console !== 'undefined' && console) ? console : {
    log: function () { },
    warn: function () { },
    error: function () { }
  };

  var activeSpreadsheetCache = {
    spreadsheet: null,
    id: '',
    timestamp: 0
  };

  var bindingAliases = {
    email: { type: 'userColumn', column: 'Email' },
    'user.email': { type: 'userColumn', column: 'Email' },
    phone: { type: 'userColumn', column: 'PhoneNumber' },
    phonenumber: { type: 'userColumn', column: 'PhoneNumber' },
    'user.phone': { type: 'userColumn', column: 'PhoneNumber' },
    fullname: { type: 'userColumn', column: 'FullName' },
    'user.fullname': { type: 'userColumn', column: 'FullName' },
    'user.full-name': { type: 'userColumn', column: 'FullName' },
    firstname: { type: 'namePart', part: 'first' },
    lastname: { type: 'namePart', part: 'last' },
    'user.firstname': { type: 'namePart', part: 'first' },
    'user.lastname': { type: 'namePart', part: 'last' },
    'insurance.provider': { type: 'insurance', key: 'provider' },
    'insurance.policy': { type: 'insurance', key: 'policyNumber' },
    'insurance.policyNumber': { type: 'insurance', key: 'policyNumber' },
    'insurance.groupNumber': { type: 'insurance', key: 'groupNumber' },
    'insurance.memberId': { type: 'insurance', key: 'memberId' },
    'insurance.notes': { type: 'insurance', key: 'notes' }
  };

  function ensureTables() {
    if (typeof DatabaseManager === 'undefined' || !DatabaseManager || typeof DatabaseManager.defineTable !== 'function') {
      throw new Error('DatabaseManager is required for DynamicFormService.');
    }

    if (!ensureTables.initialized) {
      DatabaseManager.defineTable(FORMS_TABLE, {
        headers: ['ID', 'CampaignID', 'Name', 'Description', 'Fields', 'ResponseSheetId', 'ResponseSheetName', 'CreatedBy', 'CreatedAt', 'UpdatedAt'],
        idColumn: 'ID',
        tenantColumn: 'CampaignID',
        requireTenant: false
      });

      DatabaseManager.defineTable(RESPONSES_TABLE, {
        headers: ['ID', 'FormID', 'UserID', 'CampaignID', 'Responses', 'UserUpdates', 'SubmittedBy', 'CreatedAt', 'UpdatedAt'],
        idColumn: 'ID',
        tenantColumn: 'CampaignID',
        requireTenant: false
      });

      ensureTables.initialized = true;
    }
  }

  function copyObject(source, allowedKeys) {
    var copy = {};
    for (var i = 0; i < allowedKeys.length; i++) {
      var key = allowedKeys[i];
      if (Object.prototype.hasOwnProperty.call(source, key)) {
        copy[key] = source[key];
      }
    }
    return copy;
  }

  function toStringValue(value) {
    if (value === null || typeof value === 'undefined') return '';
    return String(value);
  }

  function toDateValue(value) {
    if (!value && value !== 0) return null;
    if (value instanceof Date) return value;
    var asDate = new Date(value);
    if (isNaN(asDate.getTime())) return null;
    return asDate;
  }

  function toIsoString(value) {
    var date = toDateValue(value);
    return date ? date.toISOString() : null;
  }

  function safeJsonParse(value, fallback) {
    if (!value && value !== 0) return fallback;
    try {
      return JSON.parse(value);
    } catch (err) {
      if (safeConsole && typeof safeConsole.warn === 'function') {
        safeConsole.warn('DynamicFormService: JSON parse failed', err);
      }
      return fallback;
    }
  }

  function resolveCampaignId(context, formRecord) {
    if (formRecord) {
      if (formRecord.CampaignID) {
        return formRecord.CampaignID;
      }
      if (formRecord.campaignId) {
        return String(formRecord.campaignId);
      }
      if (formRecord.tenantId) {
        return String(formRecord.tenantId);
      }
    }
    if (!context) return '';
    if (context.campaignId) return String(context.campaignId);
    if (context.tenantId) return String(context.tenantId);
    if (Array.isArray(context.campaignIds) && context.campaignIds.length === 1) {
      return String(context.campaignIds[0]);
    }
    if (Array.isArray(context.tenantIds) && context.tenantIds.length === 1) {
      return String(context.tenantIds[0]);
    }
    return '';
  }

  function invalidateSpreadsheetCache() {
    activeSpreadsheetCache.spreadsheet = null;
    activeSpreadsheetCache.timestamp = 0;
  }

  function shouldRetrySpreadsheetError(err) {
    if (!err || !err.message) return false;
    var message = String(err.message);
    if (message.indexOf('Service Spreadsheets') !== -1) return true;
    if (message.indexOf('Document is locked') !== -1) return true;
    if (message.indexOf('locked') !== -1 && message.indexOf('Spreadsheet') !== -1) return true;
    return false;
  }

  function sleepForRetry(attempt) {
    if (typeof Utilities === 'undefined' || !Utilities || typeof Utilities.sleep !== 'function') {
      return;
    }
    var delay = 200 * Math.max(1, attempt + 1);
    if (delay > 2000) {
      delay = 2000;
    }
    try {
      Utilities.sleep(delay);
    } catch (sleepErr) {
      if (safeConsole && typeof safeConsole.warn === 'function') {
        safeConsole.warn('DynamicFormService: sleep failed during retry', sleepErr);
      }
    }
  }

  function acquireDocumentLock() {
    if (typeof LockService === 'undefined' || !LockService || typeof LockService.getDocumentLock !== 'function') {
      return null;
    }
    var attempts = 0;
    var lock = null;
    while (attempts < 3) {
      try {
        lock = LockService.getDocumentLock();
        if (!lock) {
          return null;
        }
        if (typeof lock.tryLock === 'function') {
          if (lock.tryLock(5000)) {
            return lock;
          }
        } else if (typeof lock.waitLock === 'function') {
          lock.waitLock(5000);
          return lock;
        }
      } catch (lockErr) {
        if (safeConsole && typeof safeConsole.warn === 'function') {
          safeConsole.warn('DynamicFormService: Failed to acquire document lock', lockErr);
        }
      }
      attempts += 1;
      sleepForRetry(attempts);
    }
    return null;
  }

  function releaseDocumentLock(lock) {
    if (!lock) return;
    try {
      if (typeof lock.releaseLock === 'function') {
        lock.releaseLock();
      }
    } catch (err) {
      if (safeConsole && typeof safeConsole.warn === 'function') {
        safeConsole.warn('DynamicFormService: Failed to release document lock', err);
      }
    }
  }

  function getActiveSpreadsheet() {
    if (typeof SpreadsheetApp === 'undefined' || !SpreadsheetApp || typeof SpreadsheetApp.getActiveSpreadsheet !== 'function') {
      return null;
    }
    var now = (typeof Date !== 'undefined' && Date.now) ? Date.now() : new Date().getTime();
    if (activeSpreadsheetCache.spreadsheet && (now - activeSpreadsheetCache.timestamp) < 60000) {
      return activeSpreadsheetCache.spreadsheet;
    }
    var attempts = 0;
    var lastError = null;
    while (attempts < 3) {
      try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        if (ss) {
          activeSpreadsheetCache.spreadsheet = ss;
          activeSpreadsheetCache.timestamp = now;
          if (typeof ss.getId === 'function') {
            activeSpreadsheetCache.id = ss.getId();
          }
          return ss;
        }
      } catch (err) {
        lastError = err;
        invalidateSpreadsheetCache();
        if (!shouldRetrySpreadsheetError(err)) {
          break;
        }
      }
      sleepForRetry(attempts);
      attempts += 1;
    }
    if (activeSpreadsheetCache.id && typeof SpreadsheetApp.openById === 'function') {
      try {
        var reopened = SpreadsheetApp.openById(activeSpreadsheetCache.id);
        if (reopened) {
          activeSpreadsheetCache.spreadsheet = reopened;
          activeSpreadsheetCache.timestamp = now;
          return reopened;
        }
      } catch (reopenErr) {
        lastError = reopenErr;
      }
    }
    if (lastError && safeConsole && typeof safeConsole.warn === 'function') {
      safeConsole.warn('DynamicFormService: Failed to access active spreadsheet', lastError);
    }
    return activeSpreadsheetCache.spreadsheet;
  }

  function sanitizeSheetName(name) {
    var sanitized = toStringValue(name || '');
    sanitized = sanitized.replace(/[\[\]\*\/\\\?:]/g, ' ');
    sanitized = sanitized.replace(/\s+/g, ' ').trim();
    return sanitized;
  }

  function ensureUniqueSheetName(ss, desiredName) {
    if (!ss) return desiredName || 'Dynamic Form Responses';
    var base = desiredName && desiredName.trim() ? desiredName.trim() : 'Dynamic Form Responses';
    if (base.length > 99) {
      base = base.substring(0, 99);
    }
    var candidate = base;
    var attempt = 2;
    while (ss.getSheetByName(candidate)) {
      var suffix = ' (' + attempt + ')';
      var maxLength = 99 - suffix.length;
      var shortened = base;
      if (shortened.length > maxLength) {
        shortened = shortened.substring(0, maxLength);
      }
      candidate = shortened + suffix;
      attempt += 1;
    }
    return candidate;
  }

  function buildResponseSheetHeaders(fields) {
    if (typeof DynamicFormUtilities !== 'undefined' && DynamicFormUtilities && typeof DynamicFormUtilities.buildResponseHeaders === 'function') {
      return DynamicFormUtilities.buildResponseHeaders(fields);
    }
    var headers = RESPONSE_SHEET_BASE_HEADERS.slice();
    var seen = {};
    for (var i = 0; i < headers.length; i++) {
      seen[headers[i]] = true;
    }
    if (Array.isArray(fields)) {
      for (var j = 0; j < fields.length; j++) {
        var field = fields[j];
        if (!field) continue;
        var label = toStringValue(field.label || field.title || field.name || ('Field ' + (j + 1)));
        if (!label) {
          label = 'Field ' + (j + 1);
        }
        var candidate = label;
        var counter = 2;
        while (seen[candidate]) {
          candidate = label + ' (' + counter + ')';
          counter += 1;
        }
        seen[candidate] = true;
        headers.push(candidate);
      }
    }
    return headers;
  }

  function ensureResponseSheetHeaders(sheet, fields) {
    if (typeof DynamicFormUtilities !== 'undefined' && DynamicFormUtilities && typeof DynamicFormUtilities.ensureResponseHeaders === 'function') {
      DynamicFormUtilities.ensureResponseHeaders(sheet, fields);
      return;
    }
    if (!sheet) return;
    try {
      var headers = buildResponseSheetHeaders(fields);
      var range = sheet.getRange(1, 1, 1, headers.length);
      var current = range.getValues();
      var needsUpdate = true;
      if (current && current.length) {
        var row = current[0];
        needsUpdate = row.length !== headers.length;
        if (!needsUpdate) {
          needsUpdate = false;
          for (var i = 0; i < headers.length; i++) {
            if (toStringValue(row[i]) !== toStringValue(headers[i])) {
              needsUpdate = true;
              break;
            }
          }
        }
      }
      if (needsUpdate) {
        range.setValues([headers]);
      }
      if (typeof sheet.getFrozenRows === 'function' && sheet.getFrozenRows() < 1 && typeof sheet.setFrozenRows === 'function') {
        sheet.setFrozenRows(1);
      }
    } catch (err) {
      if (safeConsole && typeof safeConsole.warn === 'function') {
        safeConsole.warn('DynamicFormService: Failed to ensure response sheet headers', err);
      }
    }
  }

  function createResponseSheet(formRecord, fields, ssOverride) {
    if (typeof DynamicFormUtilities !== 'undefined' && DynamicFormUtilities && typeof DynamicFormUtilities.createResponseSheet === 'function') {
      return DynamicFormUtilities.createResponseSheet(formRecord, fields, ssOverride || getActiveSpreadsheet());
    }
    var ss = ssOverride || getActiveSpreadsheet();
    if (!ss || !formRecord) return null;
    try {
      var baseName = sanitizeSheetName((formRecord && (formRecord.Name || formRecord.name)) || 'Dynamic Form');
      if (!baseName) {
        baseName = 'Dynamic Form';
      }
      var suffix = '';
      if (formRecord.ID || formRecord.id) {
        var idFragment = String(formRecord.ID || formRecord.id).replace(/[^a-zA-Z0-9]/g, '').substring(0, 6).toUpperCase();
        if (idFragment) {
          suffix = ' [' + idFragment + ']';
        }
      }
      var desiredName = sanitizeSheetName(baseName + ' Responses' + suffix);
      var sheetName = ensureUniqueSheetName(ss, desiredName);
      var sheet = ss.insertSheet(sheetName);
      ensureResponseSheetHeaders(sheet, fields);
      return {
        id: sheet.getSheetId(),
        name: sheetName,
        sheet: sheet
      };
    } catch (err) {
      if (safeConsole && typeof safeConsole.error === 'function') {
        safeConsole.error('DynamicFormService: Failed to create response sheet for form ' + (formRecord.ID || ''), err);
      }
      return null;
    }
  }

  function normalizeSheetValue(value) {
    if (value === null || typeof value === 'undefined') return '';
    if (value instanceof Date) return value;
    if (Array.isArray(value)) {
      var joined = [];
      for (var i = 0; i < value.length; i++) {
        if (value[i] === null || typeof value[i] === 'undefined') continue;
        joined.push(String(value[i]));
      }
      return joined.length ? joined.join(', ') : '';
    }
    if (typeof value === 'object') {
      try {
        return JSON.stringify(value);
      } catch (err) {
        return String(value);
      }
    }
    return value;
  }

  function resolveResponseSheetUrl(formRecord) {
    if (!formRecord || typeof formRecord !== 'object') return '';
    var ss = getActiveSpreadsheet();
    if (!ss || typeof ss.getUrl !== 'function') return '';
    var baseUrl = ss.getUrl();
    if (!baseUrl) return '';
    var sheetIdRaw = formRecord.ResponseSheetId || formRecord.responseSheetId;
    var sheetName = formRecord.ResponseSheetName || formRecord.responseSheetName;
    var sheetId = sheetIdRaw;
    if (!sheetIdRaw && sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        sheetId = sheet.getSheetId();
        formRecord.ResponseSheetId = String(sheetId);
      }
    }
    if (sheetId && typeof sheetId === 'number') {
      sheetId = String(sheetId);
    }
    if (typeof DynamicFormUtilities !== 'undefined' && DynamicFormUtilities && typeof DynamicFormUtilities.getSheetUrl === 'function' && sheet) {
      return DynamicFormUtilities.getSheetUrl(sheet);
    }
    if (sheetId) {
      return baseUrl + '#gid=' + sheetId;
    }
    return '';
  }

  function persistResponseSheetMetadata(formsTable, formRecord, updates) {
    if (!formsTable || !formRecord || !formRecord.ID) return;
    try {
      formsTable.update(formRecord.ID, updates);
    } catch (err) {
      if (safeConsole && typeof safeConsole.warn === 'function') {
        safeConsole.warn('DynamicFormService: Failed to persist response sheet metadata for form ' + formRecord.ID, err);
      }
    }
  }

  function ensureResponseSheetForForm(formRecord, fields, formsTable) {
    if (!formRecord) return null;
    var attempts = 0;
    var lastError = null;
    while (attempts < 3) {
      var ss = getActiveSpreadsheet();
      if (!ss) {
        return null;
      }
      var lock = acquireDocumentLock();
      try {
        var sheetName = formRecord.ResponseSheetName || formRecord.responseSheetName || '';
        var sheetIdRaw = formRecord.ResponseSheetId || formRecord.responseSheetId || '';
        var sheet = sheetName ? ss.getSheetByName(sheetName) : null;
        var numericId = sheetIdRaw ? Number(sheetIdRaw) : NaN;

        if (!sheet && sheetIdRaw && !isNaN(numericId)) {
          var sheets = ss.getSheets();
          for (var i = 0; i < sheets.length; i++) {
            if (sheets[i] && typeof sheets[i].getSheetId === 'function' && sheets[i].getSheetId() === numericId) {
              sheet = sheets[i];
              break;
            }
          }
          if (sheet && !sheetName) {
            sheetName = sheet.getName();
            formRecord.ResponseSheetName = sheetName;
            persistResponseSheetMetadata(formsTable, formRecord, { ResponseSheetName: sheetName });
          }
        }

        if (!sheet) {
          var created = createResponseSheet(formRecord, fields, ss);
          if (!created || !created.sheet) {
            throw new Error('DynamicFormService: Failed to create response sheet.');
          }
          sheet = created.sheet;
          sheetName = created.name;
          numericId = created.id;
          formRecord.ResponseSheetName = sheetName;
          formRecord.ResponseSheetId = String(numericId);
          persistResponseSheetMetadata(formsTable, formRecord, {
            ResponseSheetName: sheetName,
            ResponseSheetId: String(numericId)
          });
        } else {
          ensureResponseSheetHeaders(sheet, fields);
          if (!sheetName) {
            sheetName = sheet.getName();
            formRecord.ResponseSheetName = sheetName;
            persistResponseSheetMetadata(formsTable, formRecord, { ResponseSheetName: sheetName });
          }
          if (!sheetIdRaw || isNaN(numericId)) {
            numericId = sheet.getSheetId();
            formRecord.ResponseSheetId = String(numericId);
            persistResponseSheetMetadata(formsTable, formRecord, { ResponseSheetId: String(numericId) });
          }
        }

        ensureResponseSheetHeaders(sheet, fields);
        formRecord.ResponseSheetUrl = resolveResponseSheetUrl(formRecord);
        if (typeof DynamicFormUtilities !== 'undefined' && DynamicFormUtilities && typeof DynamicFormUtilities.upsertFormMetadata === 'function') {
          try {
            DynamicFormUtilities.upsertFormMetadata(formRecord, fields);
          } catch (metaErr) {
            if (safeConsole && typeof safeConsole.warn === 'function') {
              safeConsole.warn('DynamicFormService: Failed to sync form metadata sheet for form ' + (formRecord.ID || ''), metaErr);
            }
          }
        }

        return {
          sheet: sheet,
          name: sheetName,
          id: sheet.getSheetId()
        };
      } catch (err) {
        lastError = err;
        invalidateSpreadsheetCache();
        if (!shouldRetrySpreadsheetError(err)) {
          if (safeConsole && typeof safeConsole.error === 'function') {
            safeConsole.error('DynamicFormService: Unable to ensure response sheet for form ' + (formRecord.ID || ''), err);
          }
          throw err;
        }
      } finally {
        releaseDocumentLock(lock);
      }
      attempts += 1;
      sleepForRetry(attempts);
    }
    if (lastError) {
      throw lastError;
    }
    return null;
  }

  function appendResponseToSheet(formRecord, responseRecord, normalizedAnswers, existingUser, formsTable, fields) {
    if (!formRecord || !responseRecord) return;
    var sheetMeta = ensureResponseSheetForForm(formRecord, fields, formsTable);
    if (!sheetMeta || !sheetMeta.sheet) return;
    try {
      var sheet = sheetMeta.sheet;
      ensureResponseSheetHeaders(sheet, fields);
      var submittedAt = toDateValue(responseRecord.CreatedAt || responseRecord.createdAt || responseRecord.UpdatedAt || responseRecord.updatedAt || new Date());
      if (!submittedAt) {
        submittedAt = new Date();
      }
      var existingName = existingUser && (existingUser.FullName || existingUser.UserName || existingUser.Email || existingUser.fullName || existingUser.userName || existingUser.email);
      var answersList = normalizedAnswers && Array.isArray(normalizedAnswers.list) ? normalizedAnswers.list : [];
      var row = [
        responseRecord.ID || responseRecord.Id || responseRecord.id || '',
        formRecord.ID || formRecord.Id || formRecord.id || '',
        formRecord.Name || formRecord.name || 'Untitled form',
        submittedAt,
        responseRecord.UserID || responseRecord.userId || responseRecord.userID || '',
        existingName ? existingName : '',
        responseRecord.SubmittedBy || responseRecord.submittedBy || ''
      ];
      for (var i = 0; i < answersList.length; i++) {
        var answer = answersList[i];
        row.push(normalizeSheetValue(answer ? answer.value : ''));
      }
      sheet.appendRow(row);
    } catch (err) {
      if (safeConsole && typeof safeConsole.error === 'function') {
        safeConsole.error('DynamicFormService: Failed to append response to sheet for form ' + (formRecord.ID || ''), err);
      }
    }
  }

  function applyResponseSheetMetadata(formRecord) {
    if (!formRecord || typeof formRecord !== 'object') return formRecord;
    if (formRecord.ResponseSheetId && typeof formRecord.ResponseSheetId === 'number') {
      formRecord.ResponseSheetId = String(formRecord.ResponseSheetId);
    }
    if (formRecord.responseSheetId && typeof formRecord.responseSheetId === 'number' && !formRecord.ResponseSheetId) {
      formRecord.ResponseSheetId = String(formRecord.responseSheetId);
    }
    if (formRecord.responseSheetName && !formRecord.ResponseSheetName) {
      formRecord.ResponseSheetName = formRecord.responseSheetName;
    }
    var url = resolveResponseSheetUrl(formRecord);
    if (url) {
      formRecord.ResponseSheetUrl = url;
    }
    return formRecord;
  }

  function resolveUserDisplay(usersTable, cache, userId) {
    var key = userId ? String(userId) : '';
    if (!key) {
      return { name: '', email: '' };
    }
    if (cache && Object.prototype.hasOwnProperty.call(cache, key)) {
      return cache[key];
    }
    var summary = { name: '', email: '' };
    if (!usersTable || typeof usersTable.findById !== 'function') {
      if (cache) cache[key] = summary;
      return summary;
    }
    try {
      var user = usersTable.findById(key);
      if (user) {
        summary.name = toStringValue(user.FullName || user.UserName || user.DisplayName || '');
        summary.email = toStringValue(user.Email || '');
      }
    } catch (err) {
      if (safeConsole && typeof safeConsole.warn === 'function') {
        safeConsole.warn('DynamicFormService: Unable to resolve user ' + key, err);
      }
    }
    if (cache) {
      cache[key] = summary;
    }
    return summary;
  }

  function normalizeField(field, index) {
    if (!field || typeof field !== 'object') {
      throw new Error('Field definition at index ' + index + ' must be an object.');
    }

    var normalized = {};
    var identifier = field.id || field.name || (field.label ? String(field.label).replace(/\s+/g, '-').toLowerCase() : 'field-' + (index + 1));
    normalized.id = String(identifier);
    normalized.label = toStringValue(field.label || field.title || ('Field ' + (index + 1)));
    normalized.type = field.type ? String(field.type) : 'text';
    if (field.required) {
      normalized.required = true;
    }
    if (field.helpText) {
      normalized.helpText = toStringValue(field.helpText);
    }
    if (field.options && Array.isArray(field.options)) {
      normalized.options = field.options.slice();
    }

    var binding = normalizeBinding(field.binding || field.mapTo);
    if (binding) {
      normalized.binding = binding;
    }

    return normalized;
  }

  function normalizeBinding(binding) {
    if (!binding) return null;
    var bindingObject;
    if (typeof binding === 'string') {
      var key = binding.replace(/\s+/g, '').toLowerCase();
      bindingObject = bindingAliases[key];
      if (!bindingObject && key.indexOf('insurance:') === 0) {
        bindingObject = { type: 'insurance', key: key.split(':')[1] };
      }
    } else if (binding && typeof binding === 'object') {
      bindingObject = copyObject(binding, ['type', 'column', 'part', 'key']);
    }

    if (!bindingObject) {
      return null;
    }

    if (bindingObject.type === 'userColumn' && bindingObject.column) {
      bindingObject.column = String(bindingObject.column);
      return bindingObject;
    }

    if (bindingObject.type === 'namePart' && bindingObject.part) {
      bindingObject.part = bindingObject.part === 'last' ? 'last' : 'first';
      return bindingObject;
    }

    if (bindingObject.type === 'insurance' && bindingObject.key) {
      bindingObject.key = String(bindingObject.key);
      return bindingObject;
    }

    return null;
  }

  function serializeFields(fields) {
    return JSON.stringify(fields);
  }

  function parseFields(value) {
    return safeJsonParse(value, []);
  }

  function getUsersTable(context) {
    ensureTables();
    return DatabaseManager.table(USERS_TABLE, context || null);
  }

  function getFormsTable(context) {
    ensureTables();
    return DatabaseManager.table(FORMS_TABLE, context || null);
  }

  function getResponsesTable(context) {
    ensureTables();
    return DatabaseManager.table(RESPONSES_TABLE, context || null);
  }

  function normalizeAnswers(fields, answers) {
    var answerMap = {};
    if (!answers) {
      return { map: answerMap, list: [] };
    }

    if (Array.isArray(answers)) {
      for (var i = 0; i < answers.length; i++) {
        var entry = answers[i];
        if (!entry) continue;
        if (entry.id || entry.fieldId) {
          answerMap[entry.id || entry.fieldId] = entry.value;
        } else if (entry.name) {
          answerMap[entry.name] = entry.value;
        }
      }
    } else if (typeof answers === 'object') {
      var keys = Object.keys(answers);
      for (var j = 0; j < keys.length; j++) {
        answerMap[keys[j]] = answers[keys[j]];
      }
    }

    var ordered = [];
    for (var k = 0; k < fields.length; k++) {
      var field = fields[k];
      var value = answerMap[field.id];
      ordered.push({ id: field.id, label: field.label, value: value });
    }

    return { map: answerMap, list: ordered };
  }

  function getExistingUser(usersTable, userId) {
    if (!userId) return null;
    try {
      return usersTable.findById(userId);
    } catch (err) {
      if (safeConsole && typeof safeConsole.warn === 'function') {
        safeConsole.warn('DynamicFormService: failed to fetch user ' + userId, err);
      }
      return null;
    }
  }

  function buildUserUpdates(fields, normalizedAnswers, existingUser) {
    var map = normalizedAnswers.map;
    var updates = {};
    var nameParts = { first: null, last: null };
    var insurance = {};
    var appliedBindings = [];

    for (var i = 0; i < fields.length; i++) {
      var field = fields[i];
      if (!field.binding) continue;
      var value = map[field.id];
      if (value === null || typeof value === 'undefined' || value === '') {
        continue;
      }

      if (field.binding.type === 'userColumn' && field.binding.column) {
        updates[field.binding.column] = value;
        appliedBindings.push({ fieldId: field.id, column: field.binding.column, value: value });
        continue;
      }

      if (field.binding.type === 'namePart') {
        if (field.binding.part === 'first') {
          nameParts.first = value;
        } else if (field.binding.part === 'last') {
          nameParts.last = value;
        }
        appliedBindings.push({ fieldId: field.id, column: field.binding.part === 'last' ? 'FullName.last' : 'FullName.first', value: value });
        continue;
      }

      if (field.binding.type === 'insurance' && field.binding.key) {
        insurance[field.binding.key] = value;
        appliedBindings.push({ fieldId: field.id, column: 'InsuranceInformation.' + field.binding.key, value: value });
      }
    }

    if (nameParts.first || nameParts.last) {
      var existingFullName = existingUser && existingUser.FullName ? String(existingUser.FullName) : '';
      var existingTokens = existingFullName ? existingFullName.trim().split(/\s+/) : [];
      var existingFirst = existingTokens.length ? existingTokens[0] : '';
      var existingLast = existingTokens.length > 1 ? existingTokens.slice(1).join(' ') : '';
      var finalFirst = (nameParts.first !== null && nameParts.first !== undefined && nameParts.first !== '') ? nameParts.first : existingFirst;
      var finalLast = (nameParts.last !== null && nameParts.last !== undefined && nameParts.last !== '') ? nameParts.last : existingLast;
      var combined = (finalFirst + ' ' + finalLast).trim();
      if (!combined) {
        combined = finalFirst || finalLast;
      }
      if (combined) {
        updates.FullName = combined;
      }
    }

    if (Object.keys(insurance).length) {
      var existingInsurance = existingUser && existingUser.InsuranceInformation
        ? safeJsonParse(existingUser.InsuranceInformation, {})
        : {};
      var insuranceKeys = Object.keys(insurance);
      for (var j = 0; j < insuranceKeys.length; j++) {
        var key = insuranceKeys[j];
        existingInsurance[key] = insurance[key];
      }
      updates.InsuranceInformation = JSON.stringify(existingInsurance);
    }

    return { updates: updates, bindings: appliedBindings };
  }

  service.createForm = function (context, config) {
    ensureTables();
    if (!config || typeof config !== 'object') {
      throw new Error('Form configuration is required.');
    }
    if (!config.name && !config.title) {
      throw new Error('Form name is required.');
    }
    if (!Array.isArray(config.fields) || !config.fields.length) {
      throw new Error('At least one field definition is required.');
    }

    var fields = [];
    for (var i = 0; i < config.fields.length; i++) {
      fields.push(normalizeField(config.fields[i], i));
    }

    var record = {
      ID: (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.getUuid === 'function') ? Utilities.getUuid() : String(new Date().getTime()),
      CampaignID: resolveCampaignId(context, config),
      Name: toStringValue(config.name || config.title),
      Description: toStringValue(config.description || ''),
      Fields: serializeFields(fields),
      ResponseSheetId: '',
      ResponseSheetName: '',
      CreatedBy: config.createdBy || ''
    };

    var table = getFormsTable(context);
    var inserted = table.insert(record);
    inserted.Fields = fields;
    ensureResponseSheetForForm(inserted, fields, table);
    applyResponseSheetMetadata(inserted);
    return inserted;
  };

  service.getForm = function (context, formId) {
    ensureTables();
    if (!formId) {
      throw new Error('Form ID is required.');
    }
    var table = getFormsTable(context);
    var form = table.findById(formId);
    if (!form) return null;
    form.Fields = parseFields(form.Fields);
    ensureResponseSheetForForm(form, form.Fields, table);
    applyResponseSheetMetadata(form);
    return form;
  };

  service.listForms = function (context, options) {
    ensureTables();
    var table = getFormsTable(context);
    var rows = table.read(options || {});
    for (var i = 0; i < rows.length; i++) {
      rows[i].Fields = parseFields(rows[i].Fields);
      ensureResponseSheetForForm(rows[i], rows[i].Fields, table);
      applyResponseSheetMetadata(rows[i]);
    }
    return rows;
  };

  service.submitFormResponse = function (context, formId, userId, answers, metadata) {
    ensureTables();
    if (!formId) {
      throw new Error('Form ID is required to submit a response.');
    }
    if (!userId) {
      throw new Error('User ID is required to submit a response.');
    }

    var formsTable = getFormsTable(context);
    var form = formsTable.findById(formId);
    if (!form) {
      throw new Error('Form not found: ' + formId);
    }

    var fields = parseFields(form.Fields);
    var normalizedAnswers = normalizeAnswers(fields, answers);
    var usersTable = getUsersTable(context);
    var existingUser = getExistingUser(usersTable, userId);
    var userUpdatePayload = buildUserUpdates(fields, normalizedAnswers, existingUser);

    if (existingUser && Object.keys(userUpdatePayload.updates).length) {
      try {
        usersTable.update(userId, userUpdatePayload.updates);
      } catch (err) {
        if (safeConsole && typeof safeConsole.error === 'function') {
          safeConsole.error('DynamicFormService: failed to update user ' + userId, err);
        }
        throw err;
      }
    }

    var responseRecord = {
      ID: (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.getUuid === 'function') ? Utilities.getUuid() : String(new Date().getTime()),
      FormID: formId,
      UserID: userId,
      CampaignID: resolveCampaignId(context, form),
      Responses: JSON.stringify(normalizedAnswers.list),
      UserUpdates: JSON.stringify(userUpdatePayload.bindings),
      SubmittedBy: metadata && metadata.submittedBy ? metadata.submittedBy : ''
    };

    var responsesTable = getResponsesTable(context);
    var stored = responsesTable.insert(responseRecord);
    stored.Responses = normalizedAnswers.list;
    stored.UserUpdates = userUpdatePayload.bindings;

    appendResponseToSheet(form, stored, normalizedAnswers, existingUser, formsTable, fields);
    return stored;
  };

  service.listResponsesForUser = function (context, userId, options) {
    ensureTables();
    if (!userId) {
      throw new Error('User ID is required.');
    }
    var table = getResponsesTable(context);
    var query = options ? Object.assign({}, options) : {};
    query.where = query.where || {};
    query.where.UserID = userId;
    var rows = table.read(query);
    for (var i = 0; i < rows.length; i++) {
      rows[i].Responses = safeJsonParse(rows[i].Responses, []);
      rows[i].UserUpdates = safeJsonParse(rows[i].UserUpdates, []);
    }
    return rows;
  };

  service.listResponsesForForm = function (context, formId, options) {
    ensureTables();
    if (!formId) {
      throw new Error('Form ID is required.');
    }
    var table = getResponsesTable(context);
    var query = options ? Object.assign({}, options) : {};
    query.where = query.where || {};
    query.where.FormID = formId;
    var rows = table.read(query);
    for (var i = 0; i < rows.length; i++) {
      rows[i].Responses = safeJsonParse(rows[i].Responses, []);
      rows[i].UserUpdates = safeJsonParse(rows[i].UserUpdates, []);
    }
    return rows;
  };

  service.getDashboardSummary = function (context, options) {
    ensureTables();
    var formsTable = getFormsTable(context);
    var responsesTable = getResponsesTable(context);
    var usersTable = getUsersTable(context);

    var forms = formsTable.read(options && options.formsQuery ? options.formsQuery : {});
    var formsById = {};
    var formsWithSheets = 0;
    var formsWithoutSheets = 0;
    for (var i = 0; i < forms.length; i++) {
      var form = forms[i];
      form.Fields = parseFields(form.Fields);
      ensureResponseSheetForForm(form, form.Fields, formsTable);
      applyResponseSheetMetadata(form);
      formsById[String(form.ID)] = form;
      if (form.ResponseSheetName) {
        formsWithSheets += 1;
      } else {
        formsWithoutSheets += 1;
      }
    }

    var responses = responsesTable.read(options && options.responsesQuery ? options.responsesQuery : {});
    var responsesByFormMap = {};
    var parsedResponses = [];
    for (var j = 0; j < responses.length; j++) {
      var response = responses[j];
      response.Responses = safeJsonParse(response.Responses, []);
      response.UserUpdates = safeJsonParse(response.UserUpdates, []);
      parsedResponses.push(response);
      var formKey = String(response.FormID || '');
      if (!responsesByFormMap[formKey]) {
        responsesByFormMap[formKey] = {
          formId: formKey,
          totalResponses: 0,
          lastSubmissionAt: null
        };
      }
      responsesByFormMap[formKey].totalResponses += 1;
      var submissionIso = toIsoString(response.CreatedAt || response.createdAt || response.UpdatedAt || response.updatedAt);
      if (submissionIso && (!responsesByFormMap[formKey].lastSubmissionAt || responsesByFormMap[formKey].lastSubmissionAt < submissionIso)) {
        responsesByFormMap[formKey].lastSubmissionAt = submissionIso;
      }
    }

    var responsesByForm = [];
    var responseFormKeys = Object.keys(responsesByFormMap);
    for (var k = 0; k < responseFormKeys.length; k++) {
      var formId = responseFormKeys[k];
      var aggregate = responsesByFormMap[formId];
      var relatedForm = formsById[formId];
      responsesByForm.push({
        formId: formId,
        formName: relatedForm ? (relatedForm.Name || relatedForm.name || 'Untitled form') : 'Untitled form',
        totalResponses: aggregate.totalResponses,
        lastSubmissionAt: aggregate.lastSubmissionAt,
        responseSheetUrl: relatedForm && relatedForm.ResponseSheetUrl ? relatedForm.ResponseSheetUrl : '',
        responseSheetName: relatedForm && relatedForm.ResponseSheetName ? relatedForm.ResponseSheetName : ''
      });
    }
    responsesByForm.sort(function (a, b) {
      if (b.totalResponses === a.totalResponses) {
        var aDate = a.lastSubmissionAt || '';
        var bDate = b.lastSubmissionAt || '';
        return bDate < aDate ? -1 : bDate > aDate ? 1 : 0;
      }
      return b.totalResponses - a.totalResponses;
    });

    var recentResponses = [];
    var sortedResponses = parsedResponses.slice();
    sortedResponses.sort(function (a, b) {
      var aDate = toDateValue(a.CreatedAt || a.createdAt || a.UpdatedAt || a.updatedAt || 0);
      var bDate = toDateValue(b.CreatedAt || b.createdAt || b.UpdatedAt || b.updatedAt || 0);
      var aTime = aDate ? aDate.getTime() : 0;
      var bTime = bDate ? bDate.getTime() : 0;
      return bTime - aTime;
    });

    var userCache = {};
    var recentLimit = Math.min(sortedResponses.length, 10);
    for (var r = 0; r < recentLimit; r++) {
      var entry = sortedResponses[r];
      var formEntry = formsById[String(entry.FormID)] || null;
      var userSummary = resolveUserDisplay(usersTable, userCache, entry.UserID || entry.userId || entry.userID || '');
      recentResponses.push({
        submissionId: entry.ID,
        formId: entry.FormID,
        formName: formEntry ? (formEntry.Name || formEntry.name || 'Untitled form') : 'Untitled form',
        userId: entry.UserID || entry.userId || entry.userID || '',
        userName: userSummary.name,
        userEmail: userSummary.email,
        submittedBy: entry.SubmittedBy || entry.submittedBy || '',
        submittedAt: toIsoString(entry.CreatedAt || entry.createdAt || entry.UpdatedAt || entry.updatedAt),
        answerCount: Array.isArray(entry.Responses) ? entry.Responses.length : 0
      });
    }

    return {
      generatedAt: new Date().toISOString(),
      totalForms: forms.length,
      totalResponses: parsedResponses.length,
      formsWithSheets: formsWithSheets,
      formsWithoutSheets: formsWithoutSheets,
      averageResponsesPerForm: forms.length ? (parsedResponses.length / forms.length) : 0,
      forms: forms,
      responsesByForm: responsesByForm,
      recentResponses: recentResponses
    };
  };

  return service;
})(typeof globalThis !== 'undefined' ? globalThis : this);
