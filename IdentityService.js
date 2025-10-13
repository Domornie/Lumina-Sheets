/**
 * IdentityService.js
 * -----------------------------------------------------------------------------
 * Simplified identity management helpers that build on top of the existing
 * AuthenticationService and UserService modules.  The goal of this implementation
 * is to keep all authentication state inside the legacy user schema without
 * introducing any new headers or auxiliary sheets.  Only the columns that already
 * exist in the Users sheet are touched, and every operation gracefully skips
 * fields that are not present.
 */

var IdentityService = (function () {
  const USERS_SHEET_NAME = (typeof USERS_SHEET === 'string' && USERS_SHEET)
    ? USERS_SHEET
    : 'Users';

  const EMAIL_CONFIRMATION_TTL_MINUTES = 60;
  const PASSWORD_RESET_TTL_MINUTES = 60;

  function now() {
    return new Date();
  }

  function toSheetBoolean(value) {
    return value ? 'TRUE' : 'FALSE';
  }

  function normalizeEmail(email) {
    if (email === null || typeof email === 'undefined') {
      return '';
    }
    return String(email).trim().toLowerCase();
  }

  function hashToken(token) {
    try {
      const digest = Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        String(token || ''),
        Utilities.Charset.UTF_8
      );

      try {
        const utils = getPasswordUtils();
        if (utils && typeof utils.digestToHex === 'function') {
          return utils.digestToHex(digest);
        }
      } catch (utilsError) {
        console.warn('hashToken: falling back to manual hex conversion', utilsError);
      }

      if (digest && typeof digest.map === 'function') {
        return digest
          .map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); })
          .join('');
      }
    } catch (digestError) {
      console.warn('hashToken: Failed to compute digest', digestError);
    }

    return '';
  }

  function generateToken() {
    return Utilities.getUuid();
  }

  function sanitizeLoginReturnUrl(url) {
    if (!url && url !== 0) {
      return '';
    }

    try {
      var raw = String(url).trim();
      if (!raw) {
        return '';
      }

      if (/^javascript:/i.test(raw)) {
        return '';
      }

      if (/^https?:/i.test(raw)) {
        try {
          var base = '';
          if (typeof getBaseUrl === 'function') {
            base = String(getBaseUrl() || '');
          }
          if (!base && typeof SCRIPT_URL === 'string') {
            base = SCRIPT_URL;
          }

          if (base) {
            var baseMatch = /^https?:\/\/[^/]+/i.exec(base);
            var targetMatch = /^https?:\/\/[^/]+/i.exec(raw);
            if (baseMatch && targetMatch && baseMatch[0].toLowerCase() !== targetMatch[0].toLowerCase()) {
              return '';
            }
          }
        } catch (originError) {
          console.warn('sanitizeLoginReturnUrl: origin comparison failed', originError);
        }
      }

      if (raw.length > 500) {
        raw = raw.slice(0, 500);
      }

      return raw;
    } catch (err) {
      console.warn('sanitizeLoginReturnUrl: unable to sanitize return URL', err);
      return '';
    }
  }

  function getPasswordUtils() {
    if (typeof ensurePasswordUtilities === 'function') {
      return ensurePasswordUtilities();
    }
    if (typeof PasswordUtilities !== 'undefined' && PasswordUtilities) {
      return PasswordUtilities;
    }
    throw new Error('Password utilities unavailable');
  }

  function invalidateUsersCache() {
    try {
      if (typeof invalidateCache === 'function') {
        invalidateCache(USERS_SHEET_NAME);
      }
    } catch (err) {
      console.warn('IdentityService: unable to invalidate cache', err);
    }
  }

  function loadUserContext() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(USERS_SHEET_NAME);
    if (!sheet) {
      throw new Error('Users sheet not found');
    }

    const lastColumn = sheet.getLastColumn();
    if (lastColumn < 1) {
      return { sheet: sheet, headers: [], index: {} };
    }

    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
      .map(function (header) { return String(header || '').trim(); });

    const index = {};
    headers.forEach(function (header, idx) {
      if (header) {
        index[header] = idx;
      }
    });

    return { sheet: sheet, headers: headers, index: index };
  }

  function buildRecord(headers, row) {
    const record = {};
    for (let i = 0; i < headers.length; i += 1) {
      record[headers[i]] = row[i];
    }
    return record;
  }

  function findUserRow(predicate) {
    const context = loadUserContext();
    const sheet = context.sheet;
    const headers = context.headers;
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return null;
    }

    const range = sheet.getRange(2, 1, lastRow - 1, headers.length);
    const values = range.getValues();

    for (let i = 0; i < values.length; i += 1) {
      const record = buildRecord(headers, values[i]);
      if (predicate(record)) {
        return {
          context: context,
          rowIndex: i + 2,
          rowValues: values[i].slice(),
          user: record
        };
      }
    }

    return null;
  }

  function findUserRowByEmail(email) {
    const normalized = normalizeEmail(email);
    if (!normalized) {
      return null;
    }
    return findUserRow(function (row) {
      return normalizeEmail(row.Email) === normalized;
    });
  }

  function hasColumn(rowContext, column) {
    return rowContext && rowContext.context && Object.prototype.hasOwnProperty.call(rowContext.context.index, column);
  }

  function setColumnValue(rowContext, column, value) {
    if (!hasColumn(rowContext, column)) {
      return false;
    }
    const columnIndex = rowContext.context.index[column];
    rowContext.rowValues[columnIndex] = value;
    rowContext.context.sheet.getRange(rowContext.rowIndex, columnIndex + 1).setValue(value);
    rowContext.user[column] = value;
    return true;
  }

  function clearColumns(rowContext, columns) {
    columns.forEach(function (column) {
      setColumnValue(rowContext, column, '');
    });
  }

  function applyTimestamp(rowContext) {
    if (hasColumn(rowContext, 'UpdatedAt')) {
      setColumnValue(rowContext, 'UpdatedAt', now());
    }
  }

  function createEmailConfirmation(rowContext, options) {
    const token = generateToken();
    const issuedAt = now();
    const ttlMinutes = (options && options.ttlMinutes) || EMAIL_CONFIRMATION_TTL_MINUTES;
    const expiresAt = new Date(issuedAt.getTime() + ttlMinutes * 60000);

    setColumnValue(rowContext, 'EmailConfirmation', token);
    if (hasColumn(rowContext, 'EmailConfirmationTokenHash')) {
      setColumnValue(rowContext, 'EmailConfirmationTokenHash', hashToken(token));
    }
    if (hasColumn(rowContext, 'EmailConfirmationSentAt')) {
      setColumnValue(rowContext, 'EmailConfirmationSentAt', issuedAt);
    }
    if (hasColumn(rowContext, 'EmailConfirmationExpiresAt')) {
      setColumnValue(rowContext, 'EmailConfirmationExpiresAt', expiresAt);
    }
    if (hasColumn(rowContext, 'EmailConfirmed')) {
      setColumnValue(rowContext, 'EmailConfirmed', toSheetBoolean(false));
    }
    if (hasColumn(rowContext, 'SecurityStamp')) {
      setColumnValue(rowContext, 'SecurityStamp', Utilities.getUuid());
    }
    applyTimestamp(rowContext);

    return { token: token, issuedAt: issuedAt, expiresAt: expiresAt };
  }

  function createPasswordReset(rowContext, options) {
    const token = generateToken();
    const issuedAt = now();
    const ttlMinutes = (options && options.ttlMinutes) || PASSWORD_RESET_TTL_MINUTES;
    const expiresAt = new Date(issuedAt.getTime() + ttlMinutes * 60000);

    if (!setColumnValue(rowContext, 'ResetPasswordToken', token)) {
      setColumnValue(rowContext, 'EmailConfirmation', token);
    }
    if (hasColumn(rowContext, 'ResetPasswordTokenHash')) {
      setColumnValue(rowContext, 'ResetPasswordTokenHash', hashToken(token));
    }
    if (hasColumn(rowContext, 'ResetPasswordSentAt')) {
      setColumnValue(rowContext, 'ResetPasswordSentAt', issuedAt);
    }
    if (hasColumn(rowContext, 'ResetPasswordExpiresAt')) {
      setColumnValue(rowContext, 'ResetPasswordExpiresAt', expiresAt);
    }
    if (hasColumn(rowContext, 'ResetRequired')) {
      setColumnValue(rowContext, 'ResetRequired', toSheetBoolean(true));
    }
    if (hasColumn(rowContext, 'SecurityStamp')) {
      setColumnValue(rowContext, 'SecurityStamp', Utilities.getUuid());
    }
    applyTimestamp(rowContext);

    return { token: token, issuedAt: issuedAt, expiresAt: expiresAt };
  }

  function parseDate(value) {
    if (!value) {
      return null;
    }
    if (value instanceof Date) {
      return value;
    }
    const parsed = new Date(value);
    return isNaN(parsed.getTime()) ? null : parsed;
  }

  function confirmEmail(token) {
    if (!token) {
      return { success: false, error: 'Token is required', errorCode: 'TOKEN_REQUIRED' };
    }

    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(20000);
    } catch (err) {
      return { success: false, error: 'System busy. Try again shortly.' };
    }

    try {
      const tokenHash = hashToken(token);
      const match = findUserRow(function (row) {
        if (row.EmailConfirmationTokenHash && String(row.EmailConfirmationTokenHash).trim() === tokenHash) {
          return true;
        }
        if (row.EmailConfirmation && String(row.EmailConfirmation).trim() === String(token)) {
          return true;
        }
        return false;
      });

      if (!match) {
        return { success: false, error: 'Invalid or expired token', errorCode: 'TOKEN_INVALID' };
      }

      if (hasColumn(match, 'EmailConfirmationExpiresAt')) {
        const expiry = parseDate(match.user.EmailConfirmationExpiresAt);
        if (expiry && expiry.getTime() < now().getTime()) {
          return { success: false, error: 'Confirmation link has expired', errorCode: 'TOKEN_EXPIRED' };
        }
      }

      setColumnValue(match, 'EmailConfirmed', toSheetBoolean(true));
      clearColumns(match, [
        'EmailConfirmation',
        'EmailConfirmationTokenHash',
        'EmailConfirmationSentAt',
        'EmailConfirmationExpiresAt'
      ]);
      if (hasColumn(match, 'ResetRequired')) {
        setColumnValue(match, 'ResetRequired', toSheetBoolean(false));
      }
      if (hasColumn(match, 'SecurityStamp')) {
        setColumnValue(match, 'SecurityStamp', Utilities.getUuid());
      }
      applyTimestamp(match);
      invalidateUsersCache();

      return { success: true };
    } catch (error) {
      console.error('IdentityService.confirmEmail error', error);
      return { success: false, error: error.message || String(error) };
    } finally {
      try { lock.releaseLock(); } catch (releaseErr) { console.warn('confirmEmail: releaseLock failed', releaseErr); }
    }
  }

  function resendEmailConfirmation(email, options) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(20000);
    } catch (err) {
      return { success: false, error: 'System busy. Try again shortly.' };
    }

    try {
      const match = findUserRowByEmail(email);
      if (!match) {
        return { success: true, message: 'If an account exists, a confirmation email has been sent.' };
      }

      if (hasColumn(match, 'EmailConfirmed')) {
        const confirmed = String(match.user.EmailConfirmed || '').toUpperCase() === 'TRUE';
        if (confirmed) {
          return { success: false, error: 'Email already confirmed', errorCode: 'EMAIL_CONFIRMED' };
        }
      }

      const confirmation = createEmailConfirmation(match, options || {});
      invalidateUsersCache();

      if (!options || options.sendEmail !== false) {
        try {
          if (typeof sendPasswordSetupEmail === 'function') {
            sendPasswordSetupEmail(match.user.Email, {
              userName: match.user.UserName || '',
              fullName: match.user.FullName || '',
              passwordSetupToken: confirmation.token
            });
          } else if (typeof sendPasswordResetEmail === 'function') {
            sendPasswordResetEmail(match.user.Email, confirmation.token);
          }
        } catch (mailErr) {
          console.warn('resendEmailConfirmation: unable to send email', mailErr);
        }
      }

      return {
        success: true,
        emailConfirmationToken: (options && options.returnTokens) ? confirmation.token : null,
        expiresAt: confirmation.expiresAt
      };
    } catch (error) {
      console.error('IdentityService.resendEmailConfirmation error', error);
      return { success: false, error: error.message || String(error) };
    } finally {
      try { lock.releaseLock(); } catch (releaseErr) { console.warn('resendEmailConfirmation: releaseLock failed', releaseErr); }
    }
  }

  function beginPasswordReset(email, options) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(20000);
    } catch (err) {
      return { success: false, error: 'System busy. Try again shortly.' };
    }

    try {
      const match = findUserRowByEmail(email);
      if (!match) {
        return { success: true, message: 'If an account exists, password reset instructions were sent.' };
      }

      const reset = createPasswordReset(match, options || {});
      invalidateUsersCache();

      if (!options || options.sendEmail !== false) {
        try {
          if (typeof sendPasswordResetEmail === 'function') {
            sendPasswordResetEmail(match.user.Email, reset.token);
          }
        } catch (mailErr) {
          console.warn('beginPasswordReset: unable to send email', mailErr);
        }
      }

      return {
        success: true,
        resetToken: (options && options.returnTokens) ? reset.token : null,
        expiresAt: reset.expiresAt
      };
    } catch (error) {
      console.error('IdentityService.beginPasswordReset error', error);
      return { success: false, error: error.message || String(error) };
    } finally {
      try { lock.releaseLock(); } catch (releaseErr) { console.warn('beginPasswordReset: releaseLock failed', releaseErr); }
    }
  }

  function findResetMatch(token) {
    const tokenHash = hashToken(token);
    return findUserRow(function (row) {
      if (row.ResetPasswordTokenHash && String(row.ResetPasswordTokenHash).trim() === tokenHash) {
        return true;
      }
      if (row.ResetPasswordToken && String(row.ResetPasswordToken).trim() === String(token)) {
        return true;
      }
      if (!row.ResetPasswordToken && row.EmailConfirmation && String(row.EmailConfirmation).trim() === String(token)) {
        return true;
      }
      return false;
    });
  }

  function resetPassword(token, newPassword) {
    if (!token) {
      return { success: false, error: 'Token is required', errorCode: 'TOKEN_REQUIRED' };
    }
    if (!newPassword) {
      return { success: false, error: 'Password is required', errorCode: 'PASSWORD_REQUIRED' };
    }

    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(20000);
    } catch (err) {
      return { success: false, error: 'System busy. Try again shortly.' };
    }

    try {
      const match = findResetMatch(token);
      if (!match) {
        return { success: false, error: 'Invalid or expired token', errorCode: 'TOKEN_INVALID' };
      }

      if (hasColumn(match, 'ResetPasswordExpiresAt')) {
        const expiry = parseDate(match.user.ResetPasswordExpiresAt);
        if (expiry && expiry.getTime() < now().getTime()) {
          return { success: false, error: 'Reset token has expired', errorCode: 'TOKEN_EXPIRED' };
        }
      }

      const utils = getPasswordUtils();
      const passwordHash = utils.createPasswordHash(newPassword);
      setColumnValue(match, 'PasswordHash', passwordHash);

      clearColumns(match, [
        'ResetPasswordToken',
        'ResetPasswordTokenHash',
        'ResetPasswordSentAt',
        'ResetPasswordExpiresAt'
      ]);
      if (hasColumn(match, 'EmailConfirmation')) {
        setColumnValue(match, 'EmailConfirmation', '');
      }
      if (hasColumn(match, 'ResetRequired')) {
        setColumnValue(match, 'ResetRequired', toSheetBoolean(false));
      }
      if (hasColumn(match, 'SecurityStamp')) {
        setColumnValue(match, 'SecurityStamp', Utilities.getUuid());
      }
      applyTimestamp(match);
      invalidateUsersCache();

      return { success: true };
    } catch (error) {
      console.error('IdentityService.resetPassword error', error);
      return { success: false, error: error.message || String(error) };
    } finally {
      try { lock.releaseLock(); } catch (releaseErr) { console.warn('resetPassword: releaseLock failed', releaseErr); }
    }
  }

  function signIn(email, password, options) {
    options = options || {};
    const normalizedEmail = normalizeEmail(email);
    if (!normalizedEmail) {
      return { success: false, error: 'Email is required', errorCode: 'EMAIL_REQUIRED' };
    }
    if (!password) {
      return { success: false, error: 'Password is required', errorCode: 'PASSWORD_REQUIRED' };
    }

    try {
      const auth = (typeof AuthenticationService !== 'undefined' && AuthenticationService)
        ? AuthenticationService
        : null;
      if (!auth || typeof auth.login !== 'function') {
        throw new Error('AuthenticationService.login unavailable');
      }
      const rememberMe = !!options.rememberMe;
      const metadata = options.metadata ? Object.assign({}, options.metadata) : {};
      if (options.ipAddress && !metadata.ipAddress) metadata.ipAddress = options.ipAddress;
      if (options.userAgent && !metadata.userAgent) metadata.userAgent = options.userAgent;
      if (options.campaignId && !metadata.requestedCampaignId) {
        metadata.requestedCampaignId = options.campaignId;
      }
      if (options.returnUrl) {
        const sanitizedReturn = sanitizeLoginReturnUrl(options.returnUrl);
        if (sanitizedReturn) {
          metadata.requestedReturnUrl = sanitizedReturn;
        }
      }
      const result = auth.login(normalizedEmail, password, rememberMe, metadata);
      if (result && result.success && options && options.returnUrl) {
        const requestedReturn = sanitizeLoginReturnUrl(options.returnUrl);
        if (requestedReturn) {
          result.requestedReturnUrl = requestedReturn;
        }
      }
      return result;
    } catch (error) {
      console.error('IdentityService.signIn error', error);
      return { success: false, error: error.message || String(error) };
    }
  }

  function verifyTwoFactorCode(challengeId, code, options) {
    try {
      if (typeof AuthenticationService !== 'undefined'
        && AuthenticationService
        && typeof AuthenticationService.verifyMfaCode === 'function') {
        return AuthenticationService.verifyMfaCode(challengeId, code, options || {});
      }
      return { success: false, error: 'Two-factor verification unavailable' };
    } catch (error) {
      console.error('IdentityService.verifyTwoFactorCode error', error);
      return { success: false, error: error.message || String(error) };
    }
  }

  function signOut(sessionToken) {
    try {
      if (typeof AuthenticationService !== 'undefined'
        && AuthenticationService
        && typeof AuthenticationService.logout === 'function') {
        return AuthenticationService.logout(sessionToken || '');
      }
      return { success: true };
    } catch (error) {
      console.error('IdentityService.signOut error', error);
      return { success: false, error: error.message || String(error) };
    }
  }

  function hasActiveSession(userIdentifier) {
    try {
      if (typeof AuthenticationService !== 'undefined'
        && AuthenticationService
        && typeof AuthenticationService.userHasActiveSession === 'function') {
        return AuthenticationService.userHasActiveSession(userIdentifier);
      }
    } catch (error) {
      console.error('IdentityService.hasActiveSession error', error);
    }
    return false;
  }

  return {
    confirmEmail: confirmEmail,
    resendEmailConfirmation: resendEmailConfirmation,
    beginPasswordReset: beginPasswordReset,
    resetPassword: resetPassword,
    signIn: signIn,
    verifyTwoFactorCode: verifyTwoFactorCode,
    signOut: signOut,
    hasActiveSession: hasActiveSession,
    sanitizeLoginReturnUrl: sanitizeLoginReturnUrl
  };
})();

function identityConfirmEmail(token) {
  try {
    return IdentityService.confirmEmail(token);
  } catch (error) {
    return { success: false, error: error.message || String(error) };
  }
}

function identityResendConfirmation(email, options) {
  try {
    return IdentityService.resendEmailConfirmation(email, options);
  } catch (error) {
    return { success: false, error: error.message || String(error) };
  }
}

function identityBeginPasswordReset(email, options) {
  try {
    return IdentityService.beginPasswordReset(email, options);
  } catch (error) {
    return { success: false, error: error.message || String(error) };
  }
}

function identityResetPassword(token, newPassword) {
  try {
    return IdentityService.resetPassword(token, newPassword);
  } catch (error) {
    return { success: false, error: error.message || String(error) };
  }
}

function identitySignIn(email, password, options) {
  try {
    return IdentityService.signIn(email, password, options);
  } catch (error) {
    return { success: false, error: error.message || String(error) };
  }
}

function identityVerifyTwoFactorCode(challengeId, code, options) {
  try {
    return IdentityService.verifyTwoFactorCode(challengeId, code, options);
  } catch (error) {
    return { success: false, error: error.message || String(error) };
  }
}

function identitySignOut(sessionToken) {
  try {
    return IdentityService.signOut(sessionToken);
  } catch (error) {
    return { success: false, error: error.message || String(error) };
  }
}

function identityHasActiveSession(userIdentifier) {
  try {
    return IdentityService.hasActiveSession(userIdentifier);
  } catch (error) {
    return false;
  }
}

function resendPasswordSetupEmail(email) {
  try {
    const result = IdentityService.resendEmailConfirmation(email, { sendEmail: true });
    if (result && result.success) {
      return {
        success: true,
        message: result.message || 'If an account exists, a password setup email has been sent.'
      };
    }
    return {
      success: false,
      error: (result && result.error) || 'Unable to send the password setup email.'
    };
  } catch (error) {
    console.error('resendPasswordSetupEmail error', error);
    return { success: false, error: error.message || String(error) };
  }
}

function requestPasswordReset(email) {
  try {
    const result = IdentityService.beginPasswordReset(email, { sendEmail: true });
    if (!result || result.success === false) {
      return {
        success: false,
        error: (result && result.error) || 'Unable to send the password reset email.'
      };
    }
    return {
      success: true,
      message: result.message || 'If an account exists, password reset instructions were sent.'
    };
  } catch (error) {
    console.error('requestPasswordReset error', error);
    return { success: false, error: error.message || String(error) };
  }
}

function setPasswordWithToken(token, newPassword) {
  try {
    const result = IdentityService.resetPassword(token, newPassword);
    if (result && result.success) {
      return {
        success: true,
        message: result.message || 'Password updated successfully.'
      };
    }
    return {
      success: false,
      error: (result && result.error) || 'Failed to set the password. Please try again.'
    };
  } catch (error) {
    console.error('setPasswordWithToken error', error);
    return { success: false, error: error.message || String(error) };
  }
}
