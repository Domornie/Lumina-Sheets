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
  var safeConsole = (typeof console !== 'undefined' && console) ? console : {
    log: function () { },
    warn: function () { },
    error: function () { }
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
        headers: ['ID', 'CampaignID', 'Name', 'Description', 'Fields', 'CreatedBy', 'CreatedAt', 'UpdatedAt'],
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
      CreatedBy: config.createdBy || ''
    };

    var table = getFormsTable(context);
    var inserted = table.insert(record);
    inserted.Fields = fields;
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
    return form;
  };

  service.listForms = function (context, options) {
    ensureTables();
    var table = getFormsTable(context);
    var rows = table.read(options || {});
    for (var i = 0; i < rows.length; i++) {
      rows[i].Fields = parseFields(rows[i].Fields);
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

  return service;
})(typeof globalThis !== 'undefined' ? globalThis : this);
