/**
 * PasswordUtilities.js
 * -----------------------------------------------------------------------------
 * Centralized helpers for password hashing and verification across the Lumina
 * Sheets codebase. These utilities wrap the Google Apps Script `Utilities`
 * cryptographic helpers and provide a consistent API for creating, storing, and
 * validating password hashes.
 */

if (typeof PasswordUtilities === 'undefined') {
  var PasswordUtilities = (function () {
    function normalizePasswordInput(raw) {
      return raw == null ? '' : String(raw);
    }

    function normalizeHash(hash) {
      if (hash === null || typeof hash === 'undefined') return '';
      if (hash instanceof Date) return hash.toISOString();
      return String(hash).trim().toLowerCase();
    }

    function digestToHex(digest) {
      if (!digest || typeof digest.map !== 'function') return '';
      return digest
        .map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); })
        .join('');
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
      var hashed = hashPassword(raw);
      return constantTimeEquals(hashed, normalizedExpected);
    }

    function createPasswordHash(raw) {
      return hashPassword(raw);
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
  })();
}
