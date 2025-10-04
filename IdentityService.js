/**
 * IdentityService.js
 * -----------------------------------------------------------------------------
 * High level identity management service that mirrors the core workflow of
 * ASP.NET Core Identity (registration, login, logout, email verification,
 * password reset, and two-factor authentication) using Google Apps Script
 * primitives.
 *
 * The service builds on top of the existing AuthenticationService module while
 * ensuring users follow a centralized set of authentication rules across every
 * entry point in the Lumina Sheets solution.
 */

var IdentityService = (function () {
  const USERS_SHEET_NAME = (typeof USERS_SHEET === 'string' && USERS_SHEET)
    ? USERS_SHEET
    : 'Users';

  const LEGACY_USER_HEADERS = [
    'ID', 'UserName', 'FullName', 'Email', 'CampaignID', 'PasswordHash', 'ResetRequired',
    'EmailConfirmation', 'EmailConfirmed', 'PhoneNumber', 'EmploymentStatus', 'HireDate', 'Country',
    'LockoutEnd', 'TwoFactorEnabled', 'CanLogin', 'Roles', 'Pages', 'CreatedAt', 'UpdatedAt', 'IsAdmin'
  ];

  const IDENTITY_REQUIRED_HEADERS = [
    'NormalizedUserName',
    'NormalizedEmail',
    'PhoneNumberConfirmed',
    'LockoutEnabled',
    'AccessFailedCount',
    'TwoFactorDelivery',
    'TwoFactorSecret',
    'TwoFactorRecoveryCodes',
    'SecurityStamp',
    'ConcurrencyStamp',
    'EmailConfirmationTokenHash',
    'EmailConfirmationSentAt',
    'EmailConfirmationExpiresAt',
    'ResetPasswordToken',
    'ResetPasswordTokenHash',
    'ResetPasswordSentAt',
    'ResetPasswordExpiresAt',
    'LastLoginAt',
    'LastLoginIp',
    'LastLoginUserAgent'
  ];

  const USER_HEADERS = (function buildUserHeaders() {
    const base = (typeof USERS_HEADERS !== 'undefined'
      && Array.isArray(USERS_HEADERS)
      && USERS_HEADERS.length)
      ? USERS_HEADERS.slice()
      : LEGACY_USER_HEADERS.slice();

    const seen = {};
    base.forEach(function (header) {
      if (header) {
        seen[header] = true;
      }
    });

    IDENTITY_REQUIRED_HEADERS.forEach(function (header) {
      if (header && !seen[header]) {
        base.push(header);
        seen[header] = true;
      }
    });

    return base;
  })();

  const EMAIL_CONFIRMATION_TTL_MINUTES = 60;
  const PASSWORD_RESET_TTL_MINUTES = 60;

  function now() {
    return new Date();
  }

  function toIsoString(date) {
    if (!date) return '';
    if (date instanceof Date) {
      return date.toISOString();
    }
    try {
      const parsed = new Date(date);
      return isNaN(parsed.getTime()) ? '' : parsed.toISOString();
    } catch (err) {
      return '';
    }
  }

  function toSheetBoolean(value) {
    return value ? 'TRUE' : 'FALSE';
  }

  function fromSheetBoolean(value) {
    if (value === true) return true;
    if (value === false) return false;
    if (value == null) return false;
    const normalized = String(value).trim().toLowerCase();
    return normalized === 'true'
      || normalized === '1'
      || normalized === 'y'
      || normalized === 'yes';
  }

  function normalizeEmail(email) {
    if (!email && email !== 0) return '';
    return String(email).trim().toLowerCase();
  }

  function normalizeUserName(userName) {
    if (!userName && userName !== 0) return '';
    return String(userName).trim();
  }

  function normalizedForLookup(value) {
    return normalizeUserName(value).toUpperCase();
  }

  function hashToken(token) {
    const digest = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      String(token || ''),
      Utilities.Charset.UTF_8
    );
    return digest
      .map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); })
      .join('');
  }

  function generateSecurityStamp() {
    return Utilities.getUuid();
  }

  function generateToken() {
    const raw = Utilities.getUuid() + Utilities.getUuid();
    return Utilities.base64EncodeWebSafe(raw, Utilities.Charset.UTF_8)
      .replace(/=+$/g, '');
  }

  function getPasswordUtils() {
    if (typeof ensurePasswordUtilities === 'function') {
      return ensurePasswordUtilities();
    }
    if (typeof PasswordUtilities !== 'undefined' && PasswordUtilities) {
      return PasswordUtilities;
    }
    throw new Error('Password utilities are not available');
  }

  function ensureAuthenticationService() {
    if (typeof AuthenticationService === 'undefined' || !AuthenticationService) {
      throw new Error('AuthenticationService unavailable');
    }
    return AuthenticationService;
  }

  function getUserSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(USERS_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(USERS_SHEET_NAME);
    }
    return sheet;
  }

  function ensureHeaders(sheet, requiredHeaders) {
    const required = (requiredHeaders && requiredHeaders.length)
      ? requiredHeaders.slice()
      : USER_HEADERS.slice();
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (lastRow === 0 || lastColumn === 0) {
      sheet.clear();
      if (sheet.getMaxColumns() < required.length) {
        sheet.insertColumns(1, required.length - sheet.getMaxColumns());
      }
      sheet.getRange(1, 1, 1, required.length).setValues([required]);
      sheet.setFrozenRows(1);
      return required;
    }

    const headerRange = sheet.getRange(1, 1, 1, lastColumn);
    let headers = headerRange.getValues()[0].map(function (value) {
      return value ? String(value) : '';
    });

    if (!headers.length || (headers.length === 1 && !headers[0])) {
      sheet.clear();
      if (sheet.getMaxColumns() < required.length) {
        sheet.insertColumns(1, required.length - sheet.getMaxColumns());
      }
      sheet.getRange(1, 1, 1, required.length).setValues([required]);
      sheet.setFrozenRows(1);
      return required;
    }

    const existing = headers.slice();
    required.forEach(function (header) {
      if (existing.indexOf(header) !== -1) {
        return;
      }

      const blankIndex = existing.indexOf('');
      if (blankIndex !== -1) {
        existing[blankIndex] = header;
        sheet.getRange(1, blankIndex + 1).setValue(header);
        headers[blankIndex] = header;
        return;
      }

      const lastExistingColumn = headers.length || 1;
      sheet.insertColumnsAfter(lastExistingColumn, 1);
      const targetColumn = lastExistingColumn + 1;
      sheet.getRange(1, targetColumn).setValue(header);
      headers.push(header);
      existing.push(header);
    });

    return headers;
  }

  function buildIndexMap(headers) {
    const map = {};
    headers.forEach(function (header, index) {
      if (header) {
        map[header] = index;
      }
    });
    return map;
  }

  function ensureSheetColumns(sheet, columnCount) {
    const maxColumns = sheet.getMaxColumns();
    if (maxColumns < columnCount) {
      sheet.insertColumnsAfter(maxColumns, columnCount - maxColumns);
    }
  }

  function loadUserContext() {
    let sheet = getUserSheet();
    if (typeof ensureSheetWithHeaders === 'function') {
      try {
        sheet = ensureSheetWithHeaders(USERS_SHEET_NAME, USER_HEADERS.slice());
      } catch (ensureError) {
        console.warn('loadUserContext: ensureSheetWithHeaders failed', ensureError);
      }
    }

    const headers = ensureHeaders(sheet, USER_HEADERS);
    const index = buildIndexMap(headers);
    if (!index.ID) {
      throw new Error('Users sheet is missing the ID column');
    }
    return { sheet: sheet, headers: headers, index: index };
  }

  function readRow(sheet, headers, rowIndex) {
    const range = sheet.getRange(rowIndex, 1, 1, headers.length);
    const values = range.getValues()[0];
    const record = {};
    for (let i = 0; i < headers.length; i++) {
      record[headers[i]] = values[i];
    }
    return record;
  }

  function writeRow(sheet, headers, rowIndex, record) {
    const row = [];
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      row[i] = (typeof record[header] !== 'undefined' && record[header] !== null)
        ? record[header]
        : '';
    }
    ensureSheetColumns(sheet, headers.length);
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
  }

  function appendRow(sheet, headers, record) {
    const row = [];
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      row[i] = (typeof record[header] !== 'undefined' && record[header] !== null)
        ? record[header]
        : '';
    }
    const nextRow = sheet.getLastRow() + 1;
    ensureSheetColumns(sheet, headers.length);
    sheet.getRange(nextRow, 1, 1, headers.length).setValues([row]);
    return nextRow;
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

    for (let i = 0; i < values.length; i++) {
      const record = {};
      for (let j = 0; j < headers.length; j++) {
        record[headers[j]] = values[i][j];
      }
      if (predicate(record)) {
        return {
          context: context,
          rowIndex: i + 2,
          user: record
        };
      }
    }

    return null;
  }

  function findUserRowByEmail(email) {
    const normalized = normalizedForLookup(email);
    if (!normalized) return null;
    return findUserRow(function (row) {
      const stored = row.NormalizedEmail || normalizedForLookup(row.Email || '');
      return stored === normalized;
    });
  }

  function findUserRowById(userId) {
    if (!userId) return null;
    return findUserRow(function (row) {
      return String(row.ID || '') === String(userId);
    });
  }

  function sanitizeUser(row) {
    if (!row) return null;
    const safe = {
      id: row.ID,
      userName: row.UserName,
      normalizedUserName: row.NormalizedUserName,
      email: row.Email,
      normalizedEmail: row.NormalizedEmail,
      fullName: row.FullName,
      phoneNumber: row.PhoneNumber,
      twoFactorEnabled: fromSheetBoolean(row.TwoFactorEnabled),
      twoFactorDelivery: row.TwoFactorDelivery || 'email',
      accessFailedCount: parseInt(row.AccessFailedCount, 10) || 0,
      lockoutEnabled: fromSheetBoolean(row.LockoutEnabled),
      lockoutEnd: row.LockoutEnd || '',
      emailConfirmed: fromSheetBoolean(row.EmailConfirmed),
      lastLoginAt: row.LastLoginAt || '',
      createdAt: row.CreatedAt || '',
      updatedAt: row.UpdatedAt || ''
    };
    return safe;
  }

  function createEmailConfirmation(row, context, options) {
    const token = generateToken();
    const tokenHash = hashToken(token);
    const sentAt = toIsoString(now());
    const expiresAt = toIsoString(new Date(now().getTime() + EMAIL_CONFIRMATION_TTL_MINUTES * 60000));

    row.EmailConfirmationTokenHash = tokenHash;
    row.EmailConfirmationSentAt = sentAt;
    row.EmailConfirmationExpiresAt = expiresAt;
    row.EmailConfirmation = token;
    row.ResetRequired = toSheetBoolean(true);
    row.ResetPasswordToken = token;
    row.ResetPasswordTokenHash = tokenHash;
    row.ResetPasswordSentAt = sentAt;
    row.ResetPasswordExpiresAt = expiresAt;
    if (!fromSheetBoolean(row.EmailConfirmed)) {
      row.EmailConfirmed = toSheetBoolean(false);
    }
    row.SecurityStamp = row.SecurityStamp || generateSecurityStamp();
    row.UpdatedAt = sentAt;

    writeRow(context.sheet, context.headers, context.rowIndex, row);

    if (!options || options.sendEmail !== false) {
      try {
        if (typeof sendPasswordSetupEmail === 'function') {
          sendPasswordSetupEmail(row.Email, {
            userName: row.UserName,
            fullName: row.FullName,
            passwordSetupToken: token
          });
        }
      } catch (emailError) {
        console.warn('createEmailConfirmation: failed to send email', emailError);
      }
    }

    return {
      token: token,
      expiresAt: expiresAt
    };
  }

  function createPasswordReset(row, context, options) {
    const token = generateToken();
    const tokenHash = hashToken(token);
    const sentAt = toIsoString(now());
    const expiresAt = toIsoString(new Date(now().getTime() + PASSWORD_RESET_TTL_MINUTES * 60000));

    row.ResetPasswordToken = token;
    row.ResetPasswordTokenHash = tokenHash;
    row.ResetPasswordSentAt = sentAt;
    row.ResetPasswordExpiresAt = expiresAt;
    row.ResetRequired = toSheetBoolean(true);
    row.UpdatedAt = sentAt;

    writeRow(context.sheet, context.headers, context.rowIndex, row);

    if (!options || options.sendEmail !== false) {
      try {
        if (options && options.useAdminTemplate && typeof sendAdminPasswordResetEmail === 'function') {
          sendAdminPasswordResetEmail(row.Email, { resetToken: token });
        } else if (typeof sendPasswordResetEmail === 'function') {
          sendPasswordResetEmail(row.Email, token);
        } else if (typeof sendAdminPasswordResetEmail === 'function') {
          sendAdminPasswordResetEmail(row.Email, { resetToken: token });
        }
      } catch (emailError) {
        console.warn('createPasswordReset: failed to send email', emailError);
      }
    }

    return {
      token: token,
      expiresAt: expiresAt
    };
  }

  function registerUser(payload) {
    const data = payload || {};
    const email = normalizeEmail(data.email);
    const userName = normalizeUserName(data.userName || data.email);

    if (!email) {
      return { success: false, error: 'Email is required', errorCode: 'EMAIL_REQUIRED' };
    }
    if (!userName) {
      return { success: false, error: 'Username is required', errorCode: 'USERNAME_REQUIRED' };
    }
    if (!data.password) {
      return { success: false, error: 'Password is required', errorCode: 'PASSWORD_REQUIRED' };
    }

    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(20000);
    } catch (err) {
      return { success: false, error: 'System busy. Try again shortly.' };
    }

    try {
      const existing = findUserRowByEmail(email);
      if (existing) {
        return {
          success: false,
          error: 'A user with this email already exists.',
          errorCode: 'DUPLICATE_EMAIL',
          userId: existing.user.ID
        };
      }

      const context = loadUserContext();
      const utils = getPasswordUtils();
      const passwordHash = utils.createPasswordHash(data.password);
      const createdAt = toIsoString(now());
      const id = Utilities.getUuid();
      const securityStamp = generateSecurityStamp();
      const concurrencyStamp = generateSecurityStamp();

      const record = {
        ID: id,
        UserName: userName,
        NormalizedUserName: normalizedForLookup(userName),
        Email: data.email,
        NormalizedEmail: normalizedForLookup(email),
        FullName: data.fullName || '',
        CampaignID: data.campaignId || '',
        PasswordHash: passwordHash,
        ResetRequired: toSheetBoolean(
          typeof data.resetRequired === 'boolean'
            ? data.resetRequired
            : data.canLogin !== false
        ),
        EmailConfirmation: '',
        EmailConfirmed: toSheetBoolean(!!data.emailConfirmed),
        PhoneNumber: data.phoneNumber || '',
        PhoneNumberConfirmed: toSheetBoolean(!!data.phoneNumberConfirmed),
        EmploymentStatus: data.employmentStatus || '',
        HireDate: data.hireDate || '',
        Country: data.country || '',
        TwoFactorEnabled: toSheetBoolean(!!data.twoFactorEnabled),
        TwoFactorDelivery: data.twoFactorDelivery || 'email',
        TwoFactorSecret: data.twoFactorSecret || '',
        TwoFactorRecoveryCodes: Array.isArray(data.twoFactorRecoveryCodes)
          ? data.twoFactorRecoveryCodes.join(',')
          : (data.twoFactorRecoveryCodes || ''),
        AccessFailedCount: 0,
        LockoutEnabled: toSheetBoolean(data.lockoutEnabled !== false),
        LockoutEnd: '',
        CanLogin: toSheetBoolean(data.canLogin !== false),
        Roles: Array.isArray(data.roles) ? data.roles.join(',') : (data.roles || ''),
        Pages: Array.isArray(data.pages) ? data.pages.join(',') : (data.pages || ''),
        SecurityStamp: securityStamp,
        ConcurrencyStamp: concurrencyStamp,
        EmailConfirmationTokenHash: '',
        EmailConfirmationSentAt: '',
        EmailConfirmationExpiresAt: '',
        ResetPasswordToken: '',
        ResetPasswordTokenHash: '',
        ResetPasswordSentAt: '',
        ResetPasswordExpiresAt: '',
        LastLoginAt: '',
        LastLoginIp: '',
        LastLoginUserAgent: '',
        CreatedAt: createdAt,
        UpdatedAt: createdAt,
        IsAdmin: toSheetBoolean(!!data.isAdmin)
      };

      const rowIndex = appendRow(context.sheet, context.headers, record);
      const rowContext = {
        sheet: context.sheet,
        headers: context.headers,
        index: context.index,
        rowIndex: rowIndex,
        user: record
      };

      let confirmationToken = null;
      if (!record.EmailConfirmed || !fromSheetBoolean(record.EmailConfirmed)) {
        const confirmation = createEmailConfirmation(record, rowContext, {
          sendEmail: data.sendEmail !== false
        });
        confirmationToken = confirmation.token;
      }

      return {
        success: true,
        userId: id,
        email: data.email,
        requiresEmailConfirmation: !fromSheetBoolean(record.EmailConfirmed),
        emailConfirmationToken: data.returnTokens ? confirmationToken : null
      };
    } catch (error) {
      console.error('IdentityService.registerUser error', error);
      return { success: false, error: error.message || String(error) };
    } finally {
      try { lock.releaseLock(); } catch (releaseErr) { console.warn('registerUser: releaseLock failed', releaseErr); }
    }
  }

  function confirmEmail(token) {
    if (!token) {
      return { success: false, error: 'Token is required', errorCode: 'TOKEN_REQUIRED' };
    }
    const tokenHash = hashToken(token);
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(20000);
    } catch (err) {
      return { success: false, error: 'System busy. Try again shortly.' };
    }

    try {
      let match = findUserRow(function (row) {
        return row.EmailConfirmationTokenHash && row.EmailConfirmationTokenHash === tokenHash;
      });
      if (!match) {
        match = findUserRow(function (row) {
          return row.EmailConfirmation && String(row.EmailConfirmation).trim() === String(token);
        });
      }
      if (!match) {
        return { success: false, error: 'Invalid or expired token', errorCode: 'TOKEN_INVALID' };
      }

      const expiresAt = match.user.EmailConfirmationExpiresAt;
      if (expiresAt) {
        const expiry = new Date(expiresAt);
        if (!isNaN(expiry.getTime()) && expiry.getTime() < now().getTime()) {
          return { success: false, error: 'Confirmation token has expired', errorCode: 'TOKEN_EXPIRED' };
        }
      }

      match.user.EmailConfirmed = toSheetBoolean(true);
      match.user.EmailConfirmationTokenHash = '';
      match.user.EmailConfirmationSentAt = '';
      match.user.EmailConfirmationExpiresAt = '';
      match.user.EmailConfirmation = '';
      match.user.SecurityStamp = generateSecurityStamp();
      match.user.UpdatedAt = toIsoString(now());

      writeRow(match.context.sheet, match.context.headers, match.rowIndex, match.user);

      return { success: true };
    } catch (error) {
      console.error('IdentityService.confirmEmail error', error);
      return { success: false, error: error.message || String(error) };
    } finally {
      try { lock.releaseLock(); } catch (releaseErr) { console.warn('confirmEmail: releaseLock failed', releaseErr); }
    }
  }

  function resendEmailConfirmation(email, options) {
    const userRow = findUserRowByEmail(email);
    if (!userRow) {
      return { success: true, message: 'If an account exists, a confirmation email has been sent.' };
    }
    if (fromSheetBoolean(userRow.user.EmailConfirmed)) {
      return { success: false, error: 'Email already confirmed', errorCode: 'EMAIL_CONFIRMED' };
    }

    const confirmation = createEmailConfirmation(userRow.user, userRow, options || {});
    return {
      success: true,
      emailConfirmationToken: (options && options.returnTokens) ? confirmation.token : null,
      expiresAt: confirmation.expiresAt
    };
  }

  function beginPasswordReset(email, options) {
    const userRow = findUserRowByEmail(email);
    if (!userRow) {
      return { success: true, message: 'If an account exists, password reset instructions were sent.' };
    }

    const reset = createPasswordReset(userRow.user, userRow, options || {});
    return {
      success: true,
      resetToken: (options && options.returnTokens) ? reset.token : null,
      expiresAt: reset.expiresAt
    };
  }

  function resetPassword(token, newPassword) {
    if (!token) {
      return { success: false, error: 'Token is required', errorCode: 'TOKEN_REQUIRED' };
    }
    if (!newPassword) {
      return { success: false, error: 'Password is required', errorCode: 'PASSWORD_REQUIRED' };
    }

    const tokenHash = hashToken(token);
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(20000);
    } catch (err) {
      return { success: false, error: 'System busy. Try again shortly.' };
    }

    try {
      let match = findUserRow(function (row) {
        return row.ResetPasswordTokenHash && row.ResetPasswordTokenHash === tokenHash;
      });
      if (!match) {
        match = findUserRow(function (row) {
          return row.ResetPasswordToken && String(row.ResetPasswordToken).trim() === String(token);
        });
      }
      if (!match) {
        match = findUserRow(function (row) {
          return row.EmailConfirmation && String(row.EmailConfirmation).trim() === String(token);
        });
      }
      if (!match) {
        return { success: false, error: 'Invalid or expired token', errorCode: 'TOKEN_INVALID' };
      }

      const expiresAt = match.user.ResetPasswordExpiresAt;
      if (expiresAt) {
        const expiry = new Date(expiresAt);
        if (!isNaN(expiry.getTime()) && expiry.getTime() < now().getTime()) {
          return { success: false, error: 'Reset token has expired', errorCode: 'TOKEN_EXPIRED' };
        }
      }

      const utils = getPasswordUtils();
      const passwordHash = utils.createPasswordHash(newPassword);

      match.user.PasswordHash = passwordHash;
      match.user.ResetPasswordToken = '';
      match.user.ResetPasswordTokenHash = '';
      match.user.ResetPasswordSentAt = '';
      match.user.ResetPasswordExpiresAt = '';
      match.user.EmailConfirmation = '';
      match.user.ResetRequired = toSheetBoolean(false);
      match.user.SecurityStamp = generateSecurityStamp();
      match.user.UpdatedAt = toIsoString(now());

      writeRow(match.context.sheet, match.context.headers, match.rowIndex, match.user);

      return { success: true };
    } catch (error) {
      console.error('IdentityService.resetPassword error', error);
      return { success: false, error: error.message || String(error) };
    } finally {
      try { lock.releaseLock(); } catch (releaseErr) { console.warn('resetPassword: releaseLock failed', releaseErr); }
    }
  }

  function maskEmail(email) {
    if (!email) return '';
    const normalized = String(email).trim();
    const parts = normalized.split('@');
    if (parts.length !== 2) return normalized;
    const local = parts[0];
    const domain = parts[1];
    const maskedLocal = local.length <= 2
      ? local.charAt(0) + '*'
      : local.charAt(0) + Array(local.length - 1).fill('*').join('') + local.charAt(local.length - 1);
    return maskedLocal + '@' + domain;
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

    const rememberMe = !!options.rememberMe;
    const ipAddress = options.ipAddress || (options.metadata && options.metadata.ipAddress) || '';
    const userAgent = options.userAgent || (options.metadata && options.metadata.userAgent) || '';
    const metadata = options.metadata ? Object.assign({}, options.metadata) : {};
    if (ipAddress && !metadata.ipAddress) metadata.ipAddress = ipAddress;
    if (userAgent && !metadata.userAgent) metadata.userAgent = userAgent;

    if (options.campaignId && !metadata.requestedCampaignId) {
      metadata.requestedCampaignId = options.campaignId;
    }

    try {
      const auth = ensureAuthenticationService();
      return auth.login((email || '').trim(), password, rememberMe, metadata);
    } catch (error) {
      console.error('IdentityService.signIn error', error);
      return { success: false, error: error.message || String(error) };
    }
  }

  function verifyTwoFactorCode(challengeId, code, options) {
    if (!challengeId || !code) {
      return { success: false, error: 'Challenge ID and code are required', errorCode: 'INVALID_REQUEST' };
    }

    const metadata = options && options.metadata ? Object.assign({}, options.metadata) : {};
    if (options && options.campaignId && !metadata.requestedCampaignId) {
      metadata.requestedCampaignId = options.campaignId;
    }

    try {
      const auth = ensureAuthenticationService();
      return auth.verifyMfaCode(challengeId, code, metadata);
    } catch (error) {
      console.error('IdentityService.verifyTwoFactorCode error', error);
      return { success: false, error: error.message || String(error) };
    }
  }

  function signOut(sessionToken) {
    try {
      const auth = ensureAuthenticationService();
      return auth.logout(sessionToken);
    } catch (error) {
      return { success: false, error: error.message || String(error) };
    }
  }

  function getUserByEmail(email) {
    const userRow = findUserRowByEmail(email);
    return userRow ? sanitizeUser(userRow.user) : null;
  }

  function getUserById(userId) {
    const userRow = findUserRowById(userId);
    return userRow ? sanitizeUser(userRow.user) : null;
  }

  return {
    registerUser: registerUser,
    confirmEmail: confirmEmail,
    resendEmailConfirmation: resendEmailConfirmation,
    beginPasswordReset: beginPasswordReset,
    resetPassword: resetPassword,
    signIn: signIn,
    verifyTwoFactorCode: verifyTwoFactorCode,
    signOut: signOut,
    getUserByEmail: getUserByEmail,
    getUserById: getUserById
  };
})();

function identityRegisterUser(payload) {
  try {
    return IdentityService.registerUser(payload);
  } catch (error) {
    return { success: false, error: error.message || String(error) };
  }
}

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
