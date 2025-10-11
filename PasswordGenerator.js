/**
 * PasswordGenerator.js
 * -----------------------------------------------------------------------------
 * Centralized helper for generating passwords and managing encrypted payloads.
 * This module provides a consistent API for generating secure passwords, as well
 * as encoding and decoding password payloads for use during account creation
 * and seed data provisioning. It builds on top of PasswordUtilities when
 * available so hashed records remain consistent across the codebase.
 */

function __createPasswordGeneratorModule() {
  var DEFAULT_LENGTH = 16;
  var MIN_LENGTH = 8;
  var MAX_LENGTH = 128;

  var CHARSET_UPPER = 'ABCDEFGHJKLMNPQRSTUVWXYZ';
  var CHARSET_LOWER = 'abcdefghijkmnopqrstuvwxyz';
  var CHARSET_NUM = '23456789';
  var CHARSET_SYMBOL = '!@#$%^&*()-_=+[]{};:,.?/';

  var BASE64_REGEX = /^[A-Za-z0-9+/]+={0,2}$/;
  var BASE64_WEBSAFE_REGEX = /^[A-Za-z0-9_-]+={0,2}$/;

  function normalizeBoolean(value, defaultValue) {
    if (value === null || typeof value === 'undefined') {
      return defaultValue;
    }

    if (typeof value === 'boolean') {
      return value;
    }

    var str = String(value).trim().toLowerCase();
    if (!str) {
      return defaultValue;
    }

    if (str === 'true' || str === '1' || str === 'yes' || str === 'y' || str === 'on') {
      return true;
    }

    if (str === 'false' || str === '0' || str === 'no' || str === 'n' || str === 'off') {
      return false;
    }

    return defaultValue;
  }

  function normalizeLength(length) {
    var parsed = parseInt(length, 10);
    if (isNaN(parsed) || parsed < MIN_LENGTH) {
      parsed = DEFAULT_LENGTH;
    }
    if (parsed > MAX_LENGTH) {
      parsed = MAX_LENGTH;
    }
    return parsed;
  }

  function getPasswordUtilities() {
    if (typeof ensurePasswordUtilities === 'function') {
      return ensurePasswordUtilities();
    }
    if (typeof PasswordUtilities !== 'undefined' && PasswordUtilities) {
      return PasswordUtilities;
    }
    if (typeof __createPasswordUtilitiesModule === 'function') {
      var utils = __createPasswordUtilitiesModule();
      if (typeof PasswordUtilities === 'undefined' || !PasswordUtilities) {
        PasswordUtilities = utils;
      }
      if (typeof ensurePasswordUtilities !== 'function') {
        ensurePasswordUtilities = function ensurePasswordUtilities() { return utils; };
      }
      return utils;
    }
    return null;
  }

  function randomInt(maxExclusive) {
    if (maxExclusive <= 0) {
      return 0;
    }

    if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.getRandomInteger === 'function') {
      return Utilities.getRandomInteger(0, maxExclusive - 1);
    }

    return Math.floor(Math.random() * maxExclusive);
  }

  function buildCharacterPool(options) {
    var includeUppercase = normalizeBoolean(options && options.includeUppercase, true);
    var includeLowercase = normalizeBoolean(options && options.includeLowercase, true);
    var includeNumbers = normalizeBoolean(options && options.includeNumbers, true);
    var includeSymbols = normalizeBoolean(options && options.includeSymbols, false);

    var pools = [];
    if (includeUppercase) { pools.push(CHARSET_UPPER); }
    if (includeLowercase) { pools.push(CHARSET_LOWER); }
    if (includeNumbers) { pools.push(CHARSET_NUM); }
    if (includeSymbols) { pools.push(CHARSET_SYMBOL); }

    if (!pools.length) {
      pools.push(CHARSET_UPPER, CHARSET_LOWER, CHARSET_NUM);
    }

    return pools;
  }

  function pickRandomChar(pool) {
    if (!pool || !pool.length) {
      return '';
    }
    var index = randomInt(pool.length);
    return pool.charAt(index);
  }

  function shuffleArray(arr) {
    for (var i = arr.length - 1; i > 0; i--) {
      var j = randomInt(i + 1);
      var tmp = arr[i];
      arr[i] = arr[j];
      arr[j] = tmp;
    }
    return arr;
  }

  function ensureCharFromEachPool(pools) {
    var chars = [];
    pools.forEach(function (pool) {
      if (pool && pool.length) {
        chars.push(pickRandomChar(pool));
      }
    });
    return chars;
  }

  function generatePassword(options) {
    var length = normalizeLength(options && options.length);
    var pools = buildCharacterPool(options);
    var requiredChars = ensureCharFromEachPool(pools);
    var combinedPool = pools.join('');
    var passwordChars = requiredChars.slice();

    while (passwordChars.length < length) {
      passwordChars.push(pickRandomChar(combinedPool));
    }

    shuffleArray(passwordChars);

    var password = passwordChars.join('').slice(0, length);
    var encrypted = encryptPassword(password, { format: options && options.encryptedFormat });
    var passwordUtils = getPasswordUtilities();
    var hashRecord = passwordUtils && typeof passwordUtils.createPasswordRecord === 'function'
      ? passwordUtils.createPasswordRecord(password, { format: options && options.hashFormat })
      : null;

    return {
      password: password,
      encrypted: encrypted ? encrypted.value : '',
      encryptedFormat: encrypted ? encrypted.format : '',
      hashRecord: hashRecord
    };
  }

  function normalizeEncryptedFormat(format) {
    var normalized = (format || '').toString().trim().toLowerCase();
    if (normalized === 'base64' || normalized === 'b64') {
      return 'base64';
    }
    if (normalized === 'base64websafe' || normalized === 'base64-websafe' || normalized === 'base64_url'
      || normalized === 'base64url' || normalized === 'websafe') {
      return 'base64-websafe';
    }
    return 'base64-websafe';
  }

  function bytesToString(bytes) {
    if (!bytes) {
      return '';
    }

    if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.newBlob === 'function') {
      try {
        return Utilities.newBlob(bytes).getDataAsString(Utilities.Charset.UTF_8);
      } catch (blobErr) {
        try {
          return Utilities.newBlob(bytes).getDataAsString();
        } catch (_) {
          return '';
        }
      }
    }

    if (typeof Buffer !== 'undefined') {
      try {
        return Buffer.from(bytes).toString('utf8');
      } catch (_) {
        return '';
      }
    }

    if (typeof bytes === 'string') {
      return bytes;
    }

    try {
      return String.fromCharCode.apply(null, bytes);
    } catch (_) {
      var result = '';
      for (var i = 0; i < bytes.length; i++) {
        result += String.fromCharCode(bytes[i]);
      }
      return result;
    }
  }

  function stringToBytes(value) {
    if (value === null || typeof value === 'undefined') {
      return [];
    }
    if (typeof value === 'string') {
      if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.newBlob === 'function') {
        return Utilities.newBlob(value).getBytes();
      }
      if (typeof Buffer !== 'undefined') {
        return Buffer.from(value, 'utf8');
      }
      var bytes = [];
      for (var i = 0; i < value.length; i++) {
        bytes.push(value.charCodeAt(i));
      }
      return bytes;
    }
    return value;
  }

  function encodeBase64(value) {
    if (!value && value !== '') {
      return '';
    }

    if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.base64Encode === 'function') {
      try {
        return Utilities.base64Encode(value, Utilities.Charset.UTF_8);
      } catch (_) {
        try {
          return Utilities.base64Encode(value);
        } catch (err) {
          var bytes = stringToBytes(value);
          return Utilities.base64Encode(bytes);
        }
      }
    }

    if (typeof Buffer !== 'undefined') {
      return Buffer.from(value, 'utf8').toString('base64');
    }

    if (typeof btoa === 'function') {
      return btoa(value);
    }

    return '';
  }

  function encodeBase64WebSafe(value) {
    var encoded = encodeBase64(value);
    if (!encoded) {
      return encoded;
    }
    return encoded.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }

  function addBase64Padding(value) {
    if (!value) {
      return value;
    }
    var remainder = value.length % 4;
    if (remainder === 0) {
      return value;
    }
    var padding = '===='.slice(remainder);
    return value + padding;
  }

  function decodeBase64(value) {
    if (!value) {
      return '';
    }

    if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.base64Decode === 'function') {
      try {
        var bytes = Utilities.base64Decode(value);
        return bytesToString(bytes);
      } catch (_) {
        return '';
      }
    }

    if (typeof Buffer !== 'undefined') {
      try {
        return Buffer.from(value, 'base64').toString('utf8');
      } catch (_) {
        return '';
      }
    }

    if (typeof atob === 'function') {
      try {
        return atob(value);
      } catch (_) {
        return '';
      }
    }

    return '';
  }

  function decodeBase64WebSafe(value) {
    if (!value) {
      return '';
    }
    var normalized = value.replace(/-/g, '+').replace(/_/g, '/');
    normalized = addBase64Padding(normalized);
    return decodeBase64(normalized);
  }

  function encryptPassword(password, options) {
    var normalizedPassword = (password === null || typeof password === 'undefined') ? '' : String(password);
    if (!normalizedPassword) {
      return { value: '', format: normalizeEncryptedFormat(options && options.format) };
    }

    var format = normalizeEncryptedFormat(options && options.format);
    if (format === 'base64') {
      return { value: encodeBase64(normalizedPassword), format: 'base64' };
    }

    return { value: encodeBase64WebSafe(normalizedPassword), format: 'base64-websafe' };
  }

  function decryptPassword(encrypted, options) {
    if (!encrypted) {
      return '';
    }

    var format = options && options.format ? normalizeEncryptedFormat(options.format) : null;
    if (!format || format === 'base64-websafe') {
      var decodedWebSafe = decodeBase64WebSafe(encrypted);
      if (decodedWebSafe) {
        return decodedWebSafe;
      }
    }

    if (!format || format === 'base64') {
      return decodeBase64(encrypted);
    }

    return '';
  }

  function detectEncryptedFormat(value) {
    if (!value) {
      return 'unknown';
    }
    if (BASE64_WEBSAFE_REGEX.test(value)) {
      return 'base64-websafe';
    }
    if (BASE64_REGEX.test(value)) {
      return 'base64';
    }
    return 'unknown';
  }

  function resolvePlaintextPassword(password, encrypted, options) {
    if (password) {
      return String(password);
    }
    if (encrypted) {
      return decryptPassword(encrypted, options);
    }
    return '';
  }

  function createPasswordRecord(password, options) {
    var utils = getPasswordUtilities();
    if (!utils || typeof utils.createPasswordRecord !== 'function') {
      return null;
    }
    return utils.createPasswordRecord(password, options);
  }

  function createPasswordUpdate(password, options) {
    var utils = getPasswordUtilities();
    if (!utils || typeof utils.createPasswordUpdate !== 'function') {
      return null;
    }
    return utils.createPasswordUpdate(password, options);
  }

  return {
    generatePassword: generatePassword,
    encryptPassword: encryptPassword,
    decryptPassword: decryptPassword,
    detectEncryptedFormat: detectEncryptedFormat,
    resolvePlaintextPassword: resolvePlaintextPassword,
    createPasswordRecord: createPasswordRecord,
    createPasswordUpdate: createPasswordUpdate,
    normalizeEncryptedFormat: normalizeEncryptedFormat
  };
}

if (typeof PasswordGenerator === 'undefined' || !PasswordGenerator) {
  var PasswordGenerator = __createPasswordGeneratorModule();
}

var ensurePasswordGenerator = (typeof ensurePasswordGenerator === 'function')
  ? ensurePasswordGenerator
  : function ensurePasswordGenerator() {
      if (typeof PasswordGenerator === 'undefined' || !PasswordGenerator) {
        PasswordGenerator = __createPasswordGeneratorModule();
      }
      return PasswordGenerator;
    };
