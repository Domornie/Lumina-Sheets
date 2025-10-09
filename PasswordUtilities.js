/**
 * PasswordUtilities.js
 * -----------------------------------------------------------------------------
 * Fresh password toolkit that coordinates credential hashing, verification and
 * token generation for the rebuilt authentication stack.  The previous
 * implementation only generated raw SHA-256 digests which provided no salting,
 * no version metadata and very little room for evolution.  This version adds a
 * higher level API that can be composed by the AuthenticationService and the
 * UserService without leaking implementation details.
 *
 * All helpers are pure and intentionally framework-agnostic so they can be
 * reused inside Apps Script custom functions, triggers or services.  Every
 * operation returns explicit metadata describing the hashing parameters to
 * support future migrations.
 */

(function (global) {
  if (global.PasswordUtilities && global.PasswordUtilities.__version__ === 2) {
    return;
  }

  // ---------------------------------------------------------------------------
  // Utility helpers
  // ---------------------------------------------------------------------------

  function toStringValue(value) {
    if (value === null || typeof value === 'undefined') return '';
    return String(value);
  }

  function normalizePasswordInput(raw) {
    return toStringValue(raw);
  }

  function now() {
    return new Date();
  }

  function isoTimestamp(date) {
    return (date instanceof Date) ? date.toISOString() : new Date(date).toISOString();
  }

  function randomBytes(length) {
    var size = Math.max(8, length || 16);
    return Utilities.getRandomBytes(size);
  }

  function toBase64(bytes) {
    return Utilities.base64Encode(bytes);
  }

  function digestToHex(bytes) {
    if (!bytes || typeof bytes.map !== 'function') return '';
    return bytes
      .map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); })
      .join('');
  }

  function generateSalt(length) {
    return toBase64(randomBytes(length || 16));
  }

  function constantTimeEquals(a, b) {
    if (!a || !b) return false;
    var strA = toStringValue(a);
    var strB = toStringValue(b);
    if (strA.length !== strB.length) return false;
    var diff = 0;
    for (var i = 0; i < strA.length; i++) {
      diff |= strA.charCodeAt(i) ^ strB.charCodeAt(i);
    }
    return diff === 0;
  }

  function deriveHash(password, salt, iterations) {
    var normalizedSalt = toStringValue(salt);
    var normalizedPassword = normalizePasswordInput(password);
    var rounds = Math.max(1, iterations || 150000);

    var seed = normalizedSalt + '\u0000' + normalizedPassword;
    var digest = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      seed,
      Utilities.Charset.UTF_8
    );

    for (var i = 1; i < rounds; i++) {
      var input = Utilities.base64Encode(digest) + '\u0000' + seed + '\u0000' + i;
      digest = Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        input,
        Utilities.Charset.UTF_8
      );
    }

    return Utilities.base64Encode(digest);
  }

  function hashToken(token) {
    var normalized = toStringValue(token).trim();
    if (!normalized) return '';
    var digest = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      normalized,
      Utilities.Charset.UTF_8
    );
    return digestToHex(digest);
  }

  function validatePasswordStrength(password) {
    var value = normalizePasswordInput(password);
    var errors = [];

    if (value.length < 10) {
      errors.push('Password must be at least 10 characters long.');
    }
    if (!/[A-Z]/.test(value)) {
      errors.push('Include at least one uppercase letter.');
    }
    if (!/[a-z]/.test(value)) {
      errors.push('Include at least one lowercase letter.');
    }
    if (!/[0-9]/.test(value)) {
      errors.push('Include at least one number.');
    }
    if (!/[^A-Za-z0-9]/.test(value)) {
      errors.push('Include at least one special character.');
    }

    return {
      valid: errors.length === 0,
      errors: errors
    };
  }

  function createPasswordRecord(rawPassword, options) {
    var password = normalizePasswordInput(rawPassword);
    var strength = validatePasswordStrength(password);
    if (options && options.skipValidation !== true && !strength.valid) {
      var error = new Error('Password does not meet minimum complexity requirements.');
      error.validationErrors = strength.errors;
      throw error;
    }

    var iterations = (options && options.iterations) ? Math.max(1, options.iterations | 0) : 150000;
    var saltLength = (options && options.saltLength) ? Math.max(8, options.saltLength | 0) : 16;
    var salt = generateSalt(saltLength);
    var hash = deriveHash(password, salt, iterations);

    return {
      hash: hash,
      salt: salt,
      iterations: iterations,
      algorithm: 'SHA-256',
      version: 1,
      createdAt: isoTimestamp(now())
    };
  }

  function verifyPassword(rawPassword, record) {
    if (!record || !record.hash || !record.salt) {
      return false;
    }
    var hash = deriveHash(rawPassword, record.salt, record.iterations || 150000);
    return constantTimeEquals(hash, record.hash);
  }

  function generateRandomPassword(options) {
    var length = (options && options.length) ? Math.max(8, options.length | 0) : 14;
    var alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%^&*()-_=+[]{}';
    var bytes = randomBytes(length);
    var chars = [];
    for (var i = 0; i < length; i++) {
      var index = bytes[i] % alphabet.length;
      chars.push(alphabet.charAt(index));
    }
    return chars.join('');
  }

  function generateToken(options) {
    var length = (options && options.length) ? Math.max(16, options.length | 0) : 32;
    return toBase64(randomBytes(length));
  }

  function createResetToken(options) {
    var ttlMinutes = (options && options.ttlMinutes) ? Math.max(5, options.ttlMinutes | 0) : 60;
    var token = generateToken({ length: 24 });
    var tokenHash = hashToken(token);
    var issuedAt = now();
    var expiresAt = new Date(issuedAt.getTime() + ttlMinutes * 60 * 1000);

    return {
      token: token,
      tokenHash: tokenHash,
      issuedAt: isoTimestamp(issuedAt),
      expiresAt: isoTimestamp(expiresAt)
    };
  }

  var api = {
    __version__: 2,
    normalizePasswordInput: normalizePasswordInput,
    validatePasswordStrength: validatePasswordStrength,
    createPasswordRecord: createPasswordRecord,
    verifyPassword: verifyPassword,
    generateRandomPassword: generateRandomPassword,
    generateToken: generateToken,
    createResetToken: createResetToken,
    hashToken: hashToken,
    now: now
  };

  function ensurePasswordUtilities() {
    return api;
  }

  global.PasswordUtilities = api;
  global.ensurePasswordUtilities = ensurePasswordUtilities;
})(typeof globalThis !== 'undefined' ? globalThis : this);

