/**
 * DynamicFormUtilities.js
 *
 * Helper utilities for working with Dynamic Form spreadsheets. Responsible for
 * provisioning the metadata sheet that tracks each form and the backing
 * response sheet where submissions land.
 */
var DynamicFormUtilities = (function (global) {
  var utils = {};
  var INDEX_SHEET_NAME = 'DynamicFormsIndex';
  var INDEX_HEADERS = [
    'FormID',
    'FormName',
    'CampaignID',
    'Description',
    'ResponseSheetName',
    'ResponseSheetId',
    'ResponseSheetUrl',
    'FieldsJson',
    'CreatedAt',
    'UpdatedAt'
  ];
  var RESPONSE_BASE_HEADERS = [
    'SubmissionID',
    'FormID',
    'FormName',
    'SubmittedAt',
    'UserID',
    'UserName',
    'SubmittedBy'
  ];
  var safeConsole = (typeof console !== 'undefined' && console) ? console : {
    log: function () { },
    warn: function () { },
    error: function () { }
  };

  var spreadsheetCache = {
    id: '',
    timestamp: 0,
    spreadsheet: null
  };

  function toStringValue(value) {
    if (value === null || typeof value === 'undefined') return '';
    return String(value);
  }

  function toDate(value) {
    if (value instanceof Date) return value;
    if (!value && value !== 0) return new Date();
    var asDate = new Date(value);
    if (isNaN(asDate.getTime())) return new Date();
    return asDate;
  }

  function sanitizeSheetName(name) {
    var sanitized = toStringValue(name);
    sanitized = sanitized.replace(/[\[\]\*\/\\\?:]/g, ' ');
    sanitized = sanitized.replace(/\s+/g, ' ').trim();
    return sanitized || 'Dynamic Form';
  }

  function ensureUniqueSheetName(ss, desiredName) {
    if (!ss) return desiredName || 'Dynamic Form';
    var base = desiredName && desiredName.trim() ? desiredName.trim() : 'Dynamic Form';
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

  function getSpreadsheet() {
    if (typeof SpreadsheetApp === 'undefined' || !SpreadsheetApp || typeof SpreadsheetApp.getActiveSpreadsheet !== 'function') {
      return null;
    }
    var now = Date.now ? Date.now() : new Date().getTime();
    if (spreadsheetCache.spreadsheet && (now - spreadsheetCache.timestamp) < 60000) {
      return spreadsheetCache.spreadsheet;
    }
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss) {
        spreadsheetCache.spreadsheet = ss;
        spreadsheetCache.timestamp = now;
        if (typeof ss.getId === 'function') {
          spreadsheetCache.id = ss.getId();
        }
        return ss;
      }
    } catch (err) {
      if (safeConsole && typeof safeConsole.warn === 'function') {
        safeConsole.warn('DynamicFormUtilities: Failed to fetch active spreadsheet', err);
      }
    }
    if (spreadsheetCache.id && typeof SpreadsheetApp.openById === 'function') {
      try {
        var reopened = SpreadsheetApp.openById(spreadsheetCache.id);
        if (reopened) {
          spreadsheetCache.spreadsheet = reopened;
          spreadsheetCache.timestamp = now;
          return reopened;
        }
      } catch (openErr) {
        if (safeConsole && typeof safeConsole.warn === 'function') {
          safeConsole.warn('DynamicFormUtilities: Failed to reopen spreadsheet by id', openErr);
        }
      }
    }
    return spreadsheetCache.spreadsheet;
  }

  function ensureIndexSheet(ss) {
    if (!ss) return null;
    var sheet = ss.getSheetByName(INDEX_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(INDEX_SHEET_NAME);
      sheet.getRange(1, 1, 1, INDEX_HEADERS.length).setValues([INDEX_HEADERS]);
      if (typeof sheet.setFrozenRows === 'function') {
        sheet.setFrozenRows(1);
      }
    } else {
      try {
        var range = sheet.getRange(1, 1, 1, INDEX_HEADERS.length);
        var existing = range.getValues();
        var needsUpdate = !existing || !existing.length;
        if (!needsUpdate) {
          var row = existing[0];
          needsUpdate = row.length !== INDEX_HEADERS.length;
          if (!needsUpdate) {
            for (var i = 0; i < INDEX_HEADERS.length; i++) {
              if (toStringValue(row[i]) !== toStringValue(INDEX_HEADERS[i])) {
                needsUpdate = true;
                break;
              }
            }
          }
        }
        if (needsUpdate) {
          range.setValues([INDEX_HEADERS]);
        }
        if (typeof sheet.getFrozenRows === 'function' && sheet.getFrozenRows() < 1 && typeof sheet.setFrozenRows === 'function') {
          sheet.setFrozenRows(1);
        }
      } catch (err) {
        if (safeConsole && typeof safeConsole.warn === 'function') {
          safeConsole.warn('DynamicFormUtilities: Failed to ensure index headers', err);
        }
      }
    }
    return sheet;
  }

  function buildResponseHeaders(fields) {
    var headers = RESPONSE_BASE_HEADERS.slice();
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

  function ensureResponseHeaders(sheet, fields) {
    if (!sheet) return;
    try {
      var headers = buildResponseHeaders(fields);
      var range = sheet.getRange(1, 1, 1, headers.length);
      var existing = range.getValues();
      var needsUpdate = !existing || !existing.length;
      if (!needsUpdate) {
        var row = existing[0];
        needsUpdate = row.length !== headers.length;
        if (!needsUpdate) {
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
        safeConsole.warn('DynamicFormUtilities: Failed to ensure response headers', err);
      }
    }
  }

  function createResponseSheet(formRecord, fields, ss) {
    if (!formRecord) return null;
    var spreadsheet = ss || getSpreadsheet();
    if (!spreadsheet) return null;
    try {
      var baseName = sanitizeSheetName((formRecord && (formRecord.Name || formRecord.name)) || 'Dynamic Form');
      var suffix = '';
      if (formRecord.ID || formRecord.id) {
        var idFragment = String(formRecord.ID || formRecord.id).replace(/[^a-zA-Z0-9]/g, '').substring(0, 6).toUpperCase();
        if (idFragment) {
          suffix = ' [' + idFragment + ']';
        }
      }
      var desiredName = sanitizeSheetName(baseName + ' Responses' + suffix);
      var sheetName = ensureUniqueSheetName(spreadsheet, desiredName);
      var sheet = spreadsheet.insertSheet(sheetName);
      ensureResponseHeaders(sheet, fields);
      return {
        sheet: sheet,
        id: sheet.getSheetId(),
        name: sheetName
      };
    } catch (err) {
      if (safeConsole && typeof safeConsole.error === 'function') {
        safeConsole.error('DynamicFormUtilities: Failed to create response sheet', err);
      }
      return null;
    }
  }

  function getSheetUrl(sheet) {
    if (!sheet) return '';
    try {
      var ss = sheet.getParent();
      if (ss && typeof ss.getUrl === 'function') {
        var baseUrl = ss.getUrl();
        if (!baseUrl) return '';
        if (typeof sheet.getSheetId === 'function') {
          var id = sheet.getSheetId();
          return baseUrl + '#gid=' + id;
        }
        return baseUrl;
      }
    } catch (err) {
      if (safeConsole && typeof safeConsole.warn === 'function') {
        safeConsole.warn('DynamicFormUtilities: Failed to resolve sheet url', err);
      }
    }
    return '';
  }

  function upsertFormMetadata(formRecord, fields) {
    if (!formRecord) return;
    var ss = getSpreadsheet();
    if (!ss) return;
    var sheet = ensureIndexSheet(ss);
    if (!sheet) return;

    var lock = null;
    if (typeof LockService !== 'undefined' && LockService && typeof LockService.getDocumentLock === 'function') {
      try {
        lock = LockService.getDocumentLock();
        if (lock && typeof lock.waitLock === 'function') {
          lock.waitLock(5000);
        }
      } catch (err) {
        if (safeConsole && typeof safeConsole.warn === 'function') {
          safeConsole.warn('DynamicFormUtilities: Failed to acquire document lock for metadata', err);
        }
        lock = null;
      }
    }

    try {
      ensureIndexSheet(ss);
      var formId = toStringValue(formRecord.ID || formRecord.Id || formRecord.id || '');
      if (!formId) return;
      var lastRow = sheet.getLastRow();
      var targetRow = -1;
      if (lastRow >= 2) {
        var idRange = sheet.getRange(2, 1, lastRow - 1, 1);
        var ids = idRange.getValues();
        for (var i = 0; i < ids.length; i++) {
          if (toStringValue(ids[i][0]) === formId) {
            targetRow = i + 2;
            break;
          }
        }
      }

      var now = toDate(new Date());
      var createdAt = now;
      if (targetRow > -1) {
        var existingCreated = sheet.getRange(targetRow, 9).getValue();
        if (existingCreated) {
          createdAt = existingCreated;
        }
      }

      var responseSheetName = toStringValue(formRecord.ResponseSheetName || formRecord.responseSheetName || '');
      var responseSheetId = toStringValue(formRecord.ResponseSheetId || formRecord.responseSheetId || '');
      var responseUrl = toStringValue(formRecord.ResponseSheetUrl || formRecord.responseSheetUrl || '');
      if (!responseUrl && responseSheetName) {
        var sheetByName = ss.getSheetByName(responseSheetName);
        if (sheetByName) {
          responseUrl = getSheetUrl(sheetByName);
        }
      }

      var serializedFields = fields;
      if (!serializedFields && formRecord && formRecord.Fields) {
        serializedFields = formRecord.Fields;
      }
      var fieldsJson;
      if (typeof serializedFields === 'string') {
        fieldsJson = serializedFields;
      } else {
        try {
          fieldsJson = JSON.stringify(serializedFields || []);
        } catch (stringifyErr) {
          fieldsJson = '[]';
          if (safeConsole && typeof safeConsole.warn === 'function') {
            safeConsole.warn('DynamicFormUtilities: Failed to stringify fields for metadata', stringifyErr);
          }
        }
      }

      var payload = [
        formId,
        toStringValue(formRecord.Name || formRecord.name || ''),
        toStringValue(formRecord.CampaignID || formRecord.campaignId || ''),
        toStringValue(formRecord.Description || formRecord.description || ''),
        responseSheetName,
        responseSheetId,
        responseUrl,
        fieldsJson || '[]',
        createdAt,
        now
      ];

      if (targetRow > -1) {
        sheet.getRange(targetRow, 1, 1, payload.length).setValues([payload]);
      } else {
        sheet.appendRow(payload);
      }
    } catch (err) {
      if (safeConsole && typeof safeConsole.error === 'function') {
        safeConsole.error('DynamicFormUtilities: Failed to upsert form metadata', err);
      }
    } finally {
      if (lock && typeof lock.releaseLock === 'function') {
        try {
          lock.releaseLock();
        } catch (releaseErr) {
          if (safeConsole && typeof safeConsole.warn === 'function') {
            safeConsole.warn('DynamicFormUtilities: Failed to release lock', releaseErr);
          }
        }
      }
    }
  }

  utils.getSpreadsheet = getSpreadsheet;
  utils.ensureIndexSheet = ensureIndexSheet;
  utils.buildResponseHeaders = buildResponseHeaders;
  utils.ensureResponseHeaders = ensureResponseHeaders;
  utils.createResponseSheet = createResponseSheet;
  utils.getSheetUrl = getSheetUrl;
  utils.upsertFormMetadata = upsertFormMetadata;
  utils.RESPONSE_BASE_HEADERS = RESPONSE_BASE_HEADERS.slice();
  utils.INDEX_HEADERS = INDEX_HEADERS.slice();
  utils.INDEX_SHEET_NAME = INDEX_SHEET_NAME;

  return utils;
})(this);
