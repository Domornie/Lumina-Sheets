/**
 * EnterpriseSecurityService.js
 *
 * Provides enterprise-grade security controls for LuminaHQ data stores:
 *   - Deterministic, per-column encryption using an HMAC-based stream cipher
 *   - Tamper-evident record signatures bound to tenant and user context
 *   - Centralized security audit logging with automatic redaction of secrets
 *
 * The service is intentionally stateless – all key material is derived from a
 * single master key stored in Apps Script properties (or an in-memory fallback
 * during local execution). Every security primitive takes a metadata payload so
 * that tenant and actor context are incorporated into the cryptography.
 */
(function bootstrapEnterpriseSecurity(global) {
  if (!global) return;
  if (global.EnterpriseSecurity && typeof global.EnterpriseSecurity === 'object') {
    return;
  }

  var Utilities = global.Utilities;
  if (!Utilities) {
    throw new Error('Utilities service is required for EnterpriseSecurityService');
  }

  var Charset = Utilities.Charset || { UTF_8: 'UTF-8' };
  var MASTER_KEY_PROPERTY = 'ENTERPRISE_SECURITY_MASTER_KEY_V1';
  var MASTER_KEY_CACHE = null;
  var VERSION = 'v1';

  function getScriptProperties() {
    try {
      return (global.PropertiesService && global.PropertiesService.getScriptProperties()) || null;
    } catch (err) {
      return null;
    }
  }

  function ensureMasterKey() {
    if (MASTER_KEY_CACHE) {
      return MASTER_KEY_CACHE;
    }
    var scriptProperties = getScriptProperties();
    var stored = null;
    if (scriptProperties) {
      stored = scriptProperties.getProperty(MASTER_KEY_PROPERTY);
    } else if (global.__ENTERPRISE_SECURITY_MASTER_KEY__) {
      stored = global.__ENTERPRISE_SECURITY_MASTER_KEY__;
    }

    if (!stored) {
      var entropy = [];
      for (var i = 0; i < 8; i++) {
        entropy.push(Utilities.getUuid());
      }
      entropy.push(String(new Date().getTime()));
      entropy.push(String(Math.random()));
      var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_512, entropy.join('|'), Charset.UTF_8);
      stored = Utilities.base64Encode(digest);
      if (scriptProperties) {
        scriptProperties.setProperty(MASTER_KEY_PROPERTY, stored);
      } else {
        global.__ENTERPRISE_SECURITY_MASTER_KEY__ = stored;
      }
    }

    MASTER_KEY_CACHE = stored;
    return stored;
  }

  function getMasterKeyBytes() {
    return Utilities.base64Decode(ensureMasterKey());
  }

  function constantTimeEquals(a, b) {
    if (a === b) return true;
    if (!a || !b) return false;
    var strA = String(a);
    var strB = String(b);
    if (strA.length !== strB.length) return false;
    var diff = 0;
    for (var i = 0; i < strA.length; i++) {
      diff |= strA.charCodeAt(i) ^ strB.charCodeAt(i);
    }
    return diff === 0;
  }

  function cloneObject(value) {
    if (!value || typeof value !== 'object') return value;
    var copy = Array.isArray(value) ? [] : {};
    Object.keys(value).forEach(function (key) {
      copy[key] = cloneObject(value[key]);
    });
    return copy;
  }

  function toJson(value) {
    if (value === null || typeof value === 'undefined') return '';
    if (typeof value === 'string') return value;
    try {
      return JSON.stringify(value);
    } catch (err) {
      return String(value);
    }
  }

  function fromJson(value) {
    if (value === null || typeof value === 'undefined') return value;
    if (typeof value !== 'string') return value;
    if (!value) return '';
    if (value.charAt(0) !== '{' && value.charAt(0) !== '[') {
      return value;
    }
    try {
      return JSON.parse(value);
    } catch (err) {
      return value;
    }
  }

  function deriveKeyBytes(purpose, saltParts) {
    var masterBytes = getMasterKeyBytes();
    var keyString = Utilities.base64Encode(masterBytes);
    var messageParts = [String(purpose || '')];
    if (Array.isArray(saltParts)) {
      for (var i = 0; i < saltParts.length; i++) {
        messageParts.push(String(saltParts[i] || ''));
      }
    } else if (saltParts) {
      messageParts.push(String(saltParts));
    }
    var message = messageParts.join('|');
    var digest = Utilities.computeHmacSha256Signature(message, keyString, Charset.UTF_8);
    return digest;
  }

  function generateIvBytes() {
    var entropy = Utilities.getUuid() + '|' + new Date().getTime() + '|' + Math.random();
    var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, entropy, Charset.UTF_8);
    return digest.slice(0, 16);
  }

  function concatByteArrays(arrays) {
    var total = 0;
    for (var i = 0; i < arrays.length; i++) {
      total += arrays[i].length;
    }
    var result = new Array(total);
    var offset = 0;
    for (var j = 0; j < arrays.length; j++) {
      var arr = arrays[j];
      for (var k = 0; k < arr.length; k++) {
        result[offset + k] = arr[k];
      }
      offset += arr.length;
    }
    return result;
  }

  function deriveContextSalt(meta) {
    if (!meta) return '';
    var parts = [];
    if (meta.table) parts.push('table:' + meta.table);
    if (meta.column) parts.push('column:' + meta.column);
    if (meta.operation) parts.push('op:' + meta.operation);
    if (meta.recordId) parts.push('record:' + meta.recordId);
    var context = meta.context || {};
    if (context.tenantId) parts.push('tenant:' + context.tenantId);
    if (context.campaignId) parts.push('campaign:' + context.campaignId);
    if (context.campaign) parts.push('campaign:' + context.campaign);
    if (Array.isArray(context.allowedTenants)) {
      parts.push('allowed:' + context.allowedTenants.join(','));
    }
    if (context.userId) parts.push('user:' + context.userId);
    if (context.actorId) parts.push('actor:' + context.actorId);
    if (context.requesterId) parts.push('request:' + context.requesterId);
    if (context.sessionId) parts.push('session:' + context.sessionId);
    if (meta.classification) parts.push('class:' + meta.classification);
    return parts.join('|');
  }

  function deriveKeystream(ivBytes, length, meta) {
    var stream = [];
    var block = 0;
    var salt = deriveContextSalt(meta);
    var keyBytes = deriveKeyBytes('enc:' + VERSION, [salt]);
    var keyString = Utilities.base64Encode(keyBytes);
    while (stream.length < length) {
      var seed = Utilities.base64Encode(ivBytes) + '|' + block + '|' + salt;
      var digest = Utilities.computeHmacSha256Signature(seed, keyString, Charset.UTF_8);
      for (var i = 0; i < digest.length && stream.length < length; i++) {
        stream.push(digest[i] & 0xFF);
      }
      block += 1;
    }
    return stream;
  }

  function encryptValue(rawValue, meta) {
    if (rawValue === null || typeof rawValue === 'undefined' || rawValue === '') {
      return rawValue;
    }
    if (typeof rawValue === 'string' && rawValue.indexOf('ENC:' + VERSION + ':') === 0) {
      return rawValue;
    }
    var plainString = toJson(rawValue);
    var plainBytes = Utilities.newBlob(plainString, 'application/octet-stream').getBytes();
    var ivBytes = generateIvBytes();
    var keystream = deriveKeystream(ivBytes, plainBytes.length, meta);
    var cipherBytes = new Array(plainBytes.length);
    for (var i = 0; i < plainBytes.length; i++) {
      cipherBytes[i] = (plainBytes[i] ^ keystream[i]) & 0xFF;
    }
    var macSalt = deriveContextSalt(meta);
    var macKeyBytes = deriveKeyBytes('mac:' + VERSION, [macSalt]);
    var macKeyString = Utilities.base64Encode(macKeyBytes);
    var macPayload = VERSION + '|' + Utilities.base64Encode(ivBytes) + '|' + Utilities.base64Encode(cipherBytes) + '|' + macSalt;
    var macBytes = Utilities.computeHmacSha256Signature(macPayload, macKeyString, Charset.UTF_8);
    return [
      'ENC',
      VERSION,
      Utilities.base64Encode(ivBytes),
      Utilities.base64Encode(cipherBytes),
      Utilities.base64Encode(macBytes)
    ].join(':');
  }

  function verifyMac(ivBase64, cipherBase64, macBase64, meta) {
    var macSalt = deriveContextSalt(meta);
    var macKeyBytes = deriveKeyBytes('mac:' + VERSION, [macSalt]);
    var macKeyString = Utilities.base64Encode(macKeyBytes);
    var macPayload = VERSION + '|' + ivBase64 + '|' + cipherBase64 + '|' + macSalt;
    var expectedBytes = Utilities.computeHmacSha256Signature(macPayload, macKeyString, Charset.UTF_8);
    var expectedBase64 = Utilities.base64Encode(expectedBytes);
    if (!constantTimeEquals(expectedBase64, macBase64)) {
      throw new Error('Security envelope verification failed for ' + (meta && meta.table ? meta.table : 'unknown table'));
    }
  }

  function decryptValue(storedValue, meta) {
    if (storedValue === null || typeof storedValue === 'undefined' || storedValue === '') {
      return storedValue;
    }
    if (typeof storedValue !== 'string') {
      return storedValue;
    }
    if (storedValue.indexOf('ENC:' + VERSION + ':') !== 0) {
      return storedValue;
    }
    var parts = storedValue.split(':');
    if (parts.length < 5) {
      throw new Error('Invalid encrypted payload encountered.');
    }
    var ivBase64 = parts[2];
    var cipherBase64 = parts[3];
    var macBase64 = parts[4];
    verifyMac(ivBase64, cipherBase64, macBase64, meta);
    var ivBytes = Utilities.base64Decode(ivBase64);
    var cipherBytes = Utilities.base64Decode(cipherBase64);
    var keystream = deriveKeystream(ivBytes, cipherBytes.length, meta);
    var plainBytes = new Array(cipherBytes.length);
    for (var i = 0; i < cipherBytes.length; i++) {
      plainBytes[i] = (cipherBytes[i] ^ keystream[i]) & 0xFF;
    }
    var plainString = Utilities.newBlob(plainBytes).getDataAsString();
    return fromJson(plainString);
  }

  function canonicalizeRecord(record, options) {
    var ignore = {};
    if (options && options.signatureColumn) {
      ignore[options.signatureColumn] = true;
    }
    var keys = Object.keys(record || {});
    keys.sort();
    var parts = [];
    for (var i = 0; i < keys.length; i++) {
      var key = keys[i];
      if (ignore[key]) continue;
      var value = record[key];
      if (typeof value === 'object') {
        value = JSON.stringify(value);
      }
      parts.push(key + '=' + String(value || ''));
    }
    return parts.join('|');
  }

  function createRecordSignature(record, options) {
    var payload = canonicalizeRecord(record, options);
    var salt = deriveContextSalt(options);
    var keyBytes = deriveKeyBytes('sig:' + VERSION, [salt]);
    var keyString = Utilities.base64Encode(keyBytes);
    var digest = Utilities.computeHmacSha256Signature(VERSION + '|' + payload, keyString, Charset.UTF_8);
    return ['SIG', VERSION, Utilities.base64Encode(digest)].join(':');
  }

  function verifyRecordSignature(record, signature, options) {
    if (!signature || typeof signature !== 'string') return true;
    if (signature.indexOf('SIG:' + VERSION + ':') !== 0) return true;
    var expected = createRecordSignature(record, options);
    if (!constantTimeEquals(expected, signature)) {
      throw new Error('Record signature mismatch detected for ' + (options && options.table ? options.table : 'table'));
    }
    return true;
  }

  function normalizeSecurityConfig(config) {
    if (!config) return null;
    var normalized = cloneObject(config);
    normalized.enabled = config.enabled !== false;
    normalized.sensitiveFields = Array.isArray(config.sensitiveFields) ? config.sensitiveFields.slice() : [];
    normalized.redactedFields = Array.isArray(config.redactedFields) ? config.redactedFields.slice() : normalized.sensitiveFields.slice();
    if (!normalized.auditSheet) {
      normalized.auditSheet = 'SecurityAuditTrail';
    }
    normalized.signatureColumn = config.signatureColumn || null;
    normalized.classification = config.classification || 'restricted';
    return normalized;
  }

  function shouldEncryptField(field, config) {
    if (!config || config.enabled === false) return false;
    if (config.encryptAll === true) return true;
    if (!field) return false;
    if (Array.isArray(config.sensitiveFields) && config.sensitiveFields.indexOf(field) !== -1) {
      return true;
    }
    if (config.classifications && config.classifications[field] === 'secret') {
      return true;
    }
    return false;
  }

  function shouldRedactField(field, config) {
    if (!config) return false;
    if (Array.isArray(config.redactedFields)) {
      return config.redactedFields.indexOf(field) !== -1;
    }
    return false;
  }

  function redactRecord(record, options) {
    var config = options && options.config ? options.config : null;
    var copy = {};
    Object.keys(record || {}).forEach(function (key) {
      var value = record[key];
      if (shouldRedactField(key, config)) {
        copy[key] = value ? '***' : value;
      } else {
        copy[key] = value;
      }
    });
    if (config && config.signatureColumn && record && Object.prototype.hasOwnProperty.call(record, config.signatureColumn)) {
      copy[config.signatureColumn] = record[config.signatureColumn];
    }
    return copy;
  }

  function protectRecord(record, meta) {
    var config = meta && meta.config ? meta.config : null;
    if (!config || config.enabled === false) {
      return cloneObject(record);
    }
    var copy = cloneObject(record || {});
    var signatureColumn = config.signatureColumn;
    var signatureValue = null;
    if (signatureColumn && Object.prototype.hasOwnProperty.call(copy, signatureColumn)) {
      signatureValue = copy[signatureColumn];
      delete copy[signatureColumn];
    }
    Object.keys(copy).forEach(function (key) {
      if (shouldEncryptField(key, config)) {
        copy[key] = encryptValue(copy[key], {
          table: meta.table,
          column: key,
          context: meta.context,
          operation: meta.operation,
          recordId: meta.recordId,
          classification: config.classification
        });
      }
    });
    if (signatureColumn) {
      var signature = createRecordSignature(copy, {
        table: meta.table,
        signatureColumn: signatureColumn,
        context: meta.context,
        operation: meta.operation,
        recordId: meta.recordId,
        classification: config.classification
      });
      copy[signatureColumn] = signature;
    } else if (signatureValue) {
      copy[signatureColumn] = signatureValue;
    }
    return copy;
  }

  function revealRecord(record, meta) {
    var config = meta && meta.config ? meta.config : null;
    if (!config || config.enabled === false) {
      return cloneObject(record);
    }
    var stored = cloneObject(record || {});
    var signatureColumn = config.signatureColumn;
    var signatureValue = signatureColumn ? stored[signatureColumn] : null;
    if (signatureColumn && Object.prototype.hasOwnProperty.call(stored, signatureColumn)) {
      delete stored[signatureColumn];
    }
    if (signatureColumn && signatureValue) {
      verifyRecordSignature(stored, signatureValue, {
        table: meta.table,
        signatureColumn: signatureColumn,
        context: meta.context,
        operation: meta.operation,
        recordId: meta.recordId,
        classification: config.classification
      });
    }
    Object.keys(stored).forEach(function (key) {
      if (shouldEncryptField(key, config)) {
        stored[key] = decryptValue(stored[key], {
          table: meta.table,
          column: key,
          context: meta.context,
          operation: meta.operation,
          recordId: meta.recordId,
          classification: config.classification
        });
      }
    });
    if (signatureColumn && signatureValue) {
      stored[signatureColumn] = signatureValue;
    }
    return stored;
  }

  function normalizeAuditPayload(event) {
    var payload = cloneObject(event || {});
    if (!payload.timestamp) payload.timestamp = new Date();
    if (!payload.table) payload.table = 'Unknown';
    if (!payload.action) payload.action = 'unknown';
    if (!payload.classification) payload.classification = 'restricted';
    if (!payload.status) payload.status = 'OK';
    return payload;
  }

  function ensureAuditSheet(sheetName) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) return null;
      var name = sheetName || 'SecurityAuditTrail';
      var sheet = ss.getSheetByName(name);
      if (!sheet) {
        sheet = ss.insertSheet(name);
        var headers = [
          'Timestamp',
          'Table',
          'Action',
          'RecordId',
          'Actor',
          'Tenant',
          'Classification',
          'Status',
          'Metadata',
          'Before',
          'After'
        ];
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.setFrozenRows(1);
      }
      return sheet;
    } catch (err) {
      return null;
    }
  }

  function recordAuditEvent(event) {
    var payload = normalizeAuditPayload(event);
    var sheet = ensureAuditSheet(payload.sheetName || payload.auditSheet);
    if (!sheet) return false;
    var row = [
      payload.timestamp,
      payload.table,
      payload.action,
      payload.recordId || '',
      payload.actor || '',
      payload.tenant || '',
      payload.classification,
      payload.status,
      payload.metadata ? JSON.stringify(payload.metadata) : '',
      payload.before ? JSON.stringify(payload.before) : '',
      payload.after ? JSON.stringify(payload.after) : ''
    ];
    sheet.appendRow(row);
    return true;
  }

  function extractActor(context) {
    if (!context) return '';
    if (context.actorId) return context.actorId;
    if (context.userId) return context.userId;
    if (context.requesterId) return context.requesterId;
    if (context.auth && context.auth.userId) return context.auth.userId;
    return '';
  }

  function extractTenant(context) {
    if (!context) return '';
    if (context.tenantId) return context.tenantId;
    if (context.campaignId) return context.campaignId;
    if (context.campaign) return context.campaign;
    if (Array.isArray(context.allowedTenants) && context.allowedTenants.length === 1) {
      return context.allowedTenants[0];
    }
    return '';
  }

  var EnterpriseSecurity = {
    version: VERSION,
    constantTimeEquals: constantTimeEquals,
    encryptValue: function (value, meta) {
      return encryptValue(value, meta || {});
    },
    decryptValue: function (value, meta) {
      return decryptValue(value, meta || {});
    },
    isEncryptedValue: function (value) {
      return typeof value === 'string' && value.indexOf('ENC:' + VERSION + ':') === 0;
    },
    createRecordSignature: function (record, options) {
      return createRecordSignature(record, options || {});
    },
    verifyRecordSignature: function (record, signature, options) {
      return verifyRecordSignature(record, signature, options || {});
    },
    protectRecord: function (record, meta) {
      var normalized = cloneObject(meta || {});
      normalized.config = normalizeSecurityConfig(meta && meta.config ? meta.config : {});
      return protectRecord(record, normalized);
    },
    revealRecord: function (record, meta) {
      var normalized = cloneObject(meta || {});
      normalized.config = normalizeSecurityConfig(meta && meta.config ? meta.config : {});
      return revealRecord(record, normalized);
    },
    redactRecord: function (record, meta) {
      var normalized = cloneObject(meta || {});
      normalized.config = normalizeSecurityConfig(meta && meta.config ? meta.config : {});
      return redactRecord(record, normalized);
    },
    recordAuditEvent: function (event) {
      var payload = cloneObject(event || {});
      if (payload.before && payload.meta && payload.meta.config) {
        payload.before = redactRecord(payload.before, payload.meta);
      }
      if (payload.after && payload.meta && payload.meta.config) {
        payload.after = redactRecord(payload.after, payload.meta);
      }
      payload.actor = payload.actor || extractActor(payload.context);
      payload.tenant = payload.tenant || extractTenant(payload.context);
      payload.classification = payload.classification || (payload.meta && payload.meta.config ? payload.meta.config.classification : 'restricted');
      payload.sheetName = (payload.meta && payload.meta.config && payload.meta.config.auditSheet) || payload.sheetName;
      return recordAuditEvent(payload);
    },
    normalizeConfig: normalizeSecurityConfig,
    deriveContextSalt: deriveContextSalt,
    extractActor: extractActor,
    extractTenant: extractTenant
  };

  global.EnterpriseSecurity = EnterpriseSecurity;
})(this);
