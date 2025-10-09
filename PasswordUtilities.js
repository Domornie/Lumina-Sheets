/**
 * PasswordUtilities.js
 * -----------------------------------------------------------------------------
 * Centralized helpers for password hashing and verification across the Lumina
 * Sheets codebase. These utilities wrap the Google Apps Script `Utilities`
 * cryptographic helpers and provide a consistent API for creating, storing, and
 * validating password hashes.
 */

function __createPasswordUtilitiesModule() {
  function normalizePasswordInput(raw) {
    return raw == null ? '' : String(raw);
  }

  var HEX_HASH_REGEX = /^[0-9a-fA-F]+$/;
  var BASE64_HASH_REGEX = /^[A-Za-z0-9+/]+={0,2}$/;
  var BASE64_WEBSAFE_REGEX = /^[A-Za-z0-9_-]+={0,2}$/;

  function isHexHash(value) {
    return !!value && HEX_HASH_REGEX.test(value);
  }

  function isBase64Hash(value) {
    return !!value && BASE64_HASH_REGEX.test(value);
  }

  function isBase64WebSafeHash(value) {
    return !!value && BASE64_WEBSAFE_REGEX.test(value);
  }

  function stripBase64Padding(value) {
    if (value === null || typeof value === 'undefined') return '';
    return String(value).replace(/=+$/, '');
  }

  function normalizeHash(hash) {
    if (hash === null || typeof hash === 'undefined') return '';
    if (hash instanceof Date) return hash.toISOString();
    var str = String(hash).trim();
    if (!str) return '';
    if (isHexHash(str)) {
      return str.toLowerCase();
    }
    return str;
  }

  function digestToHex(digest) {
    if (!digest || typeof digest.map !== 'function') return '';
    return digest
      .map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); })
      .join('');
  }

  function computeHashVariants(raw) {
    var normalized = normalizePasswordInput(raw);
    var digest = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      normalized,
      Utilities.Charset.UTF_8
    );

    return {
      hex: digestToHex(digest),
      base64: Utilities.base64Encode(digest),
      base64WebSafe: Utilities.base64EncodeWebSafe(digest)
    };
  }

  function hashPassword(raw) {
    return computeHashVariants(raw).hex;
  }

  function hashPasswordBase64(raw) {
    return computeHashVariants(raw).base64;
  }

  function hashPasswordWebSafe(raw) {
    return computeHashVariants(raw).base64WebSafe;
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
    var normalizedExpected = normalizeHash(expectedHash);
    if (!normalizedExpected) return false;

    var variants = computeHashVariants(raw);

    if (constantTimeEquals(variants.hex, normalizedExpected)) {
      return true;
    }

    var looksBase64 = isBase64Hash(normalizedExpected);
    var looksWebSafe = isBase64WebSafeHash(normalizedExpected);

    if (looksBase64 || looksWebSafe) {
      if (variants.base64 && constantTimeEquals(variants.base64, normalizedExpected)) {
        return true;
      }

      if (variants.base64WebSafe && constantTimeEquals(variants.base64WebSafe, normalizedExpected)) {
        return true;
      }

      var storedNoPad = stripBase64Padding(normalizedExpected);
      if (storedNoPad && storedNoPad !== normalizedExpected) {
        var base64NoPad = stripBase64Padding(variants.base64);
        var webSafeNoPad = stripBase64Padding(variants.base64WebSafe);

        if (base64NoPad && constantTimeEquals(base64NoPad, storedNoPad)) {
          return true;
        }

        if (webSafeNoPad && constantTimeEquals(webSafeNoPad, storedNoPad)) {
          return true;
        }
      }

      try {
        var decodedHex = digestToHex(Utilities.base64Decode(normalizedExpected));
        if (decodedHex && constantTimeEquals(decodedHex, variants.hex)) {
          return true;
        }
      } catch (err1) {}

      try {
        var decodedWebSafeHex = digestToHex(Utilities.base64DecodeWebSafe(normalizedExpected));
        if (decodedWebSafeHex && constantTimeEquals(decodedWebSafeHex, variants.hex)) {
          return true;
        }
      } catch (err2) {}
    }

    return false;
  }

  function createPasswordHash(raw) {
    return hashPassword(raw);
  }

  function decodePasswordHash(hash) {
    return normalizeHash(hash);
  }

  function detectHashFormat(hash) {
    if (hash === null || typeof hash === 'undefined') {
      return 'empty';
    }

    var trimmed = String(hash).trim();
    if (!trimmed) {
      return 'empty';
    }

    if (isHexHash(trimmed)) {
      return 'hex';
    }

    if (isBase64WebSafeHash(trimmed)) {
      return 'base64-websafe';
    }

    if (isBase64Hash(trimmed)) {
      return 'base64';
    }

    return 'unknown';
  }

  return {
    normalizePasswordInput: normalizePasswordInput,
    normalizeHash: normalizeHash,
    decodePasswordHash: decodePasswordHash,
    digestToHex: digestToHex,
    hashPassword: hashPassword,
    hashPasswordBase64: hashPasswordBase64,
    hashPasswordWebSafe: hashPasswordWebSafe,
    createPasswordHash: createPasswordHash,
    verifyPassword: verifyPassword,
    comparePassword: verifyPassword,
    constantTimeEquals: constantTimeEquals,
    detectHashFormat: detectHashFormat,
    getPasswordHashVariants: computeHashVariants,
    isHexHash: isHexHash,
    isBase64Hash: isBase64Hash,
    isBase64WebSafeHash: isBase64WebSafeHash
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
