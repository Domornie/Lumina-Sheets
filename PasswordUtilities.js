/**
 * PasswordUtilities.js
 * -----------------------------------------------------------------------------
 * Centralized helpers for password hashing and verification across the Lumina
 * Sheets codebase. These utilities wrap the Google Apps Script `Utilities`
 * cryptographic helpers and provide a consistent API for creating, storing, and
 * validating password hashes.
 */

function __createPasswordUtilitiesModule() {
  var HASH_VERSION = 'v1';
  var HASH_DELIMITER = '$';

  function normalizePasswordInput(raw) {
    return raw == null ? '' : String(raw);
  }

  function digestToHex(digest) {
    if (!digest || typeof digest.map !== 'function') return '';
    return digest
      .map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); })
      .join('');
  }

  function generateSalt(length) {
    var targetLength = Math.max(16, length || 32);
    try {
      var seed = Utilities.getUuid() + ':' + Utilities.getUuid() + ':' + Date.now();
      var digest = Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        seed,
        Utilities.Charset.UTF_8
      );
      var encoded = Utilities.base64EncodeWebSafe(digest).replace(/=+$/, '');
      if (encoded.length >= targetLength) {
        return encoded.slice(0, targetLength);
      }
      var uuidSalt = Utilities.getUuid().replace(/[^A-Za-z0-9]/g, '');
      return (encoded + uuidSalt).slice(0, targetLength);
    } catch (error) {
      var fallback = String(Math.random()).replace(/[^A-Za-z0-9]/g, '');
      try {
        fallback = (Utilities.getUuid() || '').replace(/[^A-Za-z0-9]/g, '') + fallback;
      } catch (_) {}
      if (fallback.length < targetLength) {
        fallback = (fallback + Math.random().toString(36).slice(2)).slice(0, targetLength);
      }
      return fallback.slice(0, targetLength);
    }
  }

  function parseHashStructure(hash) {
    var raw = (hash === null || typeof hash === 'undefined') ? '' : String(hash).trim();
    if (!raw) {
      return { kind: 'empty', raw: '' };
    }

    var parts = raw.split(HASH_DELIMITER);
    if (parts.length === 3 && /^v\d+$/i.test(parts[0]) && parts[1] && parts[2]) {
      return {
        kind: 'versioned',
        version: parts[0].toLowerCase(),
        salt: parts[1],
        hash: parts[2],
        raw: raw
      };
    }

    return {
      kind: 'legacy',
      version: 'legacy',
      salt: '',
      hash: raw.toLowerCase(),
      raw: raw
    };
  }

  function normalizeHash(hash) {
    var parsed = parseHashStructure(hash);
    if (parsed.kind === 'empty') {
      return '';
    }
    if (parsed.kind === 'versioned') {
      return [parsed.version, parsed.salt, parsed.hash].join(HASH_DELIMITER);
    }
    return parsed.hash;
  }

  function hashPassword(raw) {
    var normalized = normalizePasswordInput(raw);
    var digest = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      normalized,
      Utilities.Charset.UTF_8
    );
    return digestToHex(digest);
  }

  function computeVersionedHash(password, salt) {
    var normalized = normalizePasswordInput(password);
    var digest = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      salt + '|' + normalized,
      Utilities.Charset.UTF_8
    );
    return digestToHex(digest);
  }

  function constantTimeEquals(a, b) {
    if (a == null || b == null) return false;
    var strA = String(a);
    var strB = String(b);
    if (strA.length !== strB.length) return false;
    var diff = 0;
    for (var i = 0; i < strA.length; i++) {
      diff |= strA.charCodeAt(i) ^ strB.charCodeAt(i);
    }
    return diff === 0;
  }

  function verifyPassword(raw, expectedHash) {
    var parsed = parseHashStructure(expectedHash);
    if (parsed.kind === 'empty') {
      return false;
    }

    if (parsed.kind === 'versioned') {
      var recomputed = computeVersionedHash(raw, parsed.salt);
      return constantTimeEquals(recomputed, parsed.hash);
    }

    var hashed = hashPassword(raw);
    return constantTimeEquals(hashed, parsed.hash);
  }

  function createPasswordHash(raw) {
    var salt = generateSalt(32);
    var hashed = computeVersionedHash(raw, salt);
    return normalizeHash([HASH_VERSION, salt, hashed].join(HASH_DELIMITER));
  }

  function decodePasswordHash(hash) {
    return normalizeHash(hash);
  }

  return {
    normalizePasswordInput: normalizePasswordInput,
    normalizeHash: normalizeHash,
    decodePasswordHash: decodePasswordHash,
    digestToHex: digestToHex,
    hashPassword: hashPassword,
    createPasswordHash: createPasswordHash,
    verifyPassword: verifyPassword,
    comparePassword: verifyPassword,
    constantTimeEquals: constantTimeEquals
  };
}

if (typeof PasswordUtilities === 'undefined' || !PasswordUtilities) {
  var PasswordUtilities = __createPasswordUtilitiesModule();
}

var ensurePasswordUtilities = (typeof ensurePasswordUtilities === 'function')
  ? ensurePasswordUtilities
  : function ensurePasswordUtilities() {
    if (typeof PasswordUtilities === 'undefined' || !PasswordUtilities) {
      PasswordUtilities = __createPasswordUtilitiesModule();
    }
    return PasswordUtilities;
  };
