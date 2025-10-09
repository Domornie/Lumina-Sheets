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

  var cachedUserIdentityFields = null;

  function ensureUserIdentityFields() {
    if (Array.isArray(cachedUserIdentityFields) && cachedUserIdentityFields.length) {
      return cachedUserIdentityFields;
    }

    if (typeof USERS_HEADERS !== 'undefined' && Array.isArray(USERS_HEADERS) && USERS_HEADERS.length) {
      cachedUserIdentityFields = USERS_HEADERS.slice();
      return cachedUserIdentityFields;
    }

    console.warn('IdentityService: USERS_HEADERS not defined; using empty identity field list');
    cachedUserIdentityFields = [];
    return cachedUserIdentityFields;
  }

  function now() {
    return new Date();
  }

  function toSheetBoolean(value) {
    return value ? 'TRUE' : 'FALSE';
  }

  function parseBooleanFlag(value) {
    if (value === true) return true;
    if (value === false || value === null || typeof value === 'undefined') return false;
    if (typeof value === 'number') return value !== 0;

    const normalized = String(value).trim().toUpperCase();
    if (!normalized) return false;

    switch (normalized) {
      case 'TRUE':
      case 'YES':
      case 'Y':
      case '1':
      case 'ON':
        return true;
      default:
        return false;
    }
  }

  function normalizeEmail(email) {
    if (email === null || typeof email === 'undefined') {
      return '';
    }
    return String(email).trim().toLowerCase();
  }

  function coerceString(value) {
    if (value === null || typeof value === 'undefined') {
      return '';
    }
    if (value instanceof Date) {
      return value.toISOString();
    }
    return String(value);
  }

  function parseNumber(value) {
    if (value === null || typeof value === 'undefined' || value === '') {
      return null;
    }
    if (typeof value === 'number') {
      if (isNaN(value)) return null;
      return value;
    }
    const parsed = Number(String(value).replace(/,/g, '').trim());
    return isNaN(parsed) ? null : parsed;
  }

  function parseInteger(value) {
    const num = parseNumber(value);
    if (num === null) return null;
    const intVal = parseInt(num, 10);
    return isNaN(intVal) ? null : intVal;
  }

  function parseDelimitedList(value) {
    if (!value && value !== 0) {
      return [];
    }
    if (Array.isArray(value)) {
      return value
        .map(function (item) { return coerceString(item).trim(); })
        .filter(function (item) { return !!item; });
    }

    var raw = coerceString(value).trim();
    if (!raw) {
      return [];
    }

    try {
      if (/^\s*\[/.test(raw)) {
        var parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) {
          return parsed
            .map(function (item) { return coerceString(item).trim(); })
            .filter(function (item) { return !!item; });
        }
      }
    } catch (err) {
      console.warn('parseDelimitedList: JSON parse failed', err);
    }

    return raw
      .split(/[\r\n,;]+/)
      .map(function (item) { return item.trim(); })
      .filter(function (item) { return !!item; });
  }

  function parseRecoveryCodes(value) {
    var list = parseDelimitedList(value);
    return list.map(function (item) {
      return item.replace(/\s+/g, '').toUpperCase();
    });
  }

  function parseIsoDate(value) {
    if (!value && value !== 0) {
      return null;
    }
    if (value instanceof Date) {
      return value;
    }

    var str = coerceString(value).trim();
    if (!str) {
      return null;
    }

    var normalized = str.replace(/\s+/g, '').toLowerCase();
    if (normalized === '0'
      || normalized === '0000-00-00'
      || normalized === '0000-00-00t00:00:00z'
      || normalized === 'null'
      || normalized === 'false') {
      return null;
    }

    var parsed = new Date(str);
    if (isNaN(parsed.getTime())) {
      return null;
    }

    if (parsed.getTime() === 0) {
      return null;
    }

    if (parsed.getFullYear && parsed.getFullYear() <= 1901 && /1899|1900/.test(normalized)) {
      return null;
    }

    return parsed;
  }

  function ensureDateOnly(date) {
    if (!(date instanceof Date)) {
      return null;
    }
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  }

  function calculateProbationEnd(fields) {
    var explicitEnd = parseIsoDate(fields.ProbationEndDate || fields.ProbationEnd);
    if (explicitEnd) {
      return explicitEnd;
    }

    var hire = parseIsoDate(fields.HireDate);
    var months = parseInteger(fields.ProbationMonths);

    if (!hire || months === null || months < 1) {
      return null;
    }

    var endDate = new Date(hire.getTime());
    endDate.setMonth(endDate.getMonth() + months);
    return endDate;
  }

  function buildInsuranceStatus(fields) {
    var eligibleDate = parseIsoDate(fields.InsuranceEligibleDate || fields.InsuranceQualifiedDate);
    var qualifiedDate = parseIsoDate(fields.InsuranceQualifiedDate || fields.InsuranceEligibleDate);
    var enrolled = parseBooleanFlag(fields.InsuranceEnrolled);
    var signedUp = parseBooleanFlag(fields.InsuranceSignedUp);
    var eligible = parseBooleanFlag(fields.InsuranceEligible);
    var qualified = parseBooleanFlag(fields.InsuranceQualified);
    var cardReceived = parseIsoDate(fields.InsuranceCardReceivedDate);

    return {
      eligible: eligible || (!!eligibleDate && eligibleDate.getTime() <= now().getTime()),
      qualified: qualified || (!!qualifiedDate && qualifiedDate.getTime() <= now().getTime()),
      enrolled: enrolled,
      signedUp: signedUp,
      eligibleDate: eligibleDate,
      qualifiedDate: qualifiedDate,
      cardReceivedDate: cardReceived
    };
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

  function pickIdentityFields(record) {
    const picked = {};
    ensureUserIdentityFields().forEach(function (field) {
      if (Object.prototype.hasOwnProperty.call(record, field)) {
        picked[field] = record[field];
      } else {
        picked[field] = '';
      }
    });
    return picked;
  }

  function normalizeIdentityFields(raw) {
    const normalized = {};

    normalized.ID = coerceString(raw.ID).trim();
    normalized.UserName = coerceString(raw.UserName).trim();
    normalized.FullName = coerceString(raw.FullName).trim();
    normalized.Email = normalizeEmail(raw.Email);
    normalized.CampaignID = coerceString(raw.CampaignID).trim();
    normalized.PasswordHash = coerceString(raw.PasswordHash).trim();
    normalized.ResetRequired = parseBooleanFlag(raw.ResetRequired);
    normalized.EmailConfirmation = coerceString(raw.EmailConfirmation).trim();
    normalized.EmailConfirmed = parseBooleanFlag(raw.EmailConfirmed);
    normalized.PhoneNumber = coerceString(raw.PhoneNumber).trim();
    normalized.EmploymentStatus = coerceString(raw.EmploymentStatus).trim();
    normalized.HireDate = parseIsoDate(raw.HireDate);
    normalized.Country = coerceString(raw.Country).trim();
    normalized.LockoutEnd = parseIsoDate(raw.LockoutEnd);
    normalized.TwoFactorEnabled = parseBooleanFlag(raw.TwoFactorEnabled);
    normalized.CanLogin = parseBooleanFlag(raw.CanLogin !== '' ? raw.CanLogin : true);
    normalized.Roles = parseDelimitedList(raw.Roles);
    normalized.Pages = parseDelimitedList(raw.Pages);
    normalized.CreatedAt = parseIsoDate(raw.CreatedAt);
    normalized.UpdatedAt = parseIsoDate(raw.UpdatedAt);
    normalized.IsAdmin = parseBooleanFlag(raw.IsAdmin);
    normalized.NormalizedUserName = coerceString(raw.NormalizedUserName || raw.UserName).trim().toLowerCase();
    normalized.NormalizedEmail = coerceString(raw.NormalizedEmail || normalized.Email).trim().toLowerCase();
    normalized.PhoneNumberConfirmed = parseBooleanFlag(raw.PhoneNumberConfirmed);
    normalized.LockoutEnabled = parseBooleanFlag(raw.LockoutEnabled);
    normalized.AccessFailedCount = parseInteger(raw.AccessFailedCount) || 0;
    normalized.TwoFactorDelivery = coerceString(raw.TwoFactorDelivery).trim().toLowerCase();
    normalized.TwoFactorSecret = coerceString(raw.TwoFactorSecret).trim();
    normalized.TwoFactorRecoveryCodes = parseRecoveryCodes(raw.TwoFactorRecoveryCodes);
    normalized.SecurityStamp = coerceString(raw.SecurityStamp).trim();
    normalized.ConcurrencyStamp = coerceString(raw.ConcurrencyStamp).trim();
    normalized.EmailConfirmationTokenHash = coerceString(raw.EmailConfirmationTokenHash).trim();
    normalized.EmailConfirmationSentAt = parseIsoDate(raw.EmailConfirmationSentAt);
    normalized.EmailConfirmationExpiresAt = parseIsoDate(raw.EmailConfirmationExpiresAt);
    normalized.ResetPasswordToken = coerceString(raw.ResetPasswordToken).trim();
    normalized.ResetPasswordTokenHash = coerceString(raw.ResetPasswordTokenHash).trim();
    normalized.ResetPasswordSentAt = parseIsoDate(raw.ResetPasswordSentAt);
    normalized.ResetPasswordExpiresAt = parseIsoDate(raw.ResetPasswordExpiresAt);
    normalized.LastLogin = parseIsoDate(raw.LastLogin);
    normalized.LastLoginAt = parseIsoDate(raw.LastLoginAt);
    normalized.LastLoginIp = coerceString(raw.LastLoginIp).trim();
    normalized.LastLoginUserAgent = coerceString(raw.LastLoginUserAgent).trim();
    normalized.DeletedAt = parseIsoDate(raw.DeletedAt);
    normalized.TerminationDate = parseIsoDate(raw.TerminationDate);
    normalized.ProbationMonths = parseInteger(raw.ProbationMonths);
    normalized.ProbationEnd = parseIsoDate(raw.ProbationEnd);
    normalized.ProbationEndDate = parseIsoDate(raw.ProbationEndDate);
    normalized.InsuranceEligibleDate = parseIsoDate(raw.InsuranceEligibleDate);
    normalized.InsuranceQualifiedDate = parseIsoDate(raw.InsuranceQualifiedDate);
    normalized.InsuranceEligible = parseBooleanFlag(raw.InsuranceEligible);
    normalized.InsuranceQualified = parseBooleanFlag(raw.InsuranceQualified);
    normalized.InsuranceEnrolled = parseBooleanFlag(raw.InsuranceEnrolled);
    normalized.InsuranceSignedUp = parseBooleanFlag(raw.InsuranceSignedUp);
    normalized.InsuranceCardReceivedDate = parseIsoDate(raw.InsuranceCardReceivedDate);
    normalized.MFASecret = coerceString(raw.MFASecret).trim();
    normalized.MFABackupCodes = parseRecoveryCodes(raw.MFABackupCodes);
    normalized.MFADeliveryPreference = coerceString(raw.MFADeliveryPreference).trim().toLowerCase();
    normalized.MFAEnabled = parseBooleanFlag(raw.MFAEnabled);

    return normalized;
  }

  function summarizeIdentityForClient(identity) {
    if (!identity || !identity.fields) {
      return null;
    }

    var status = identity.status || {};
    var mfaStatus = status.mfa || {};
    var insurance = status.insurance || {};
    var probation = status.probation || {};
    var termination = status.termination || {};
    var deletion = status.deletion || {};
    var lockout = status.lockout || {};
    var lastLogin = status.lastLogin || {};

    return {
      id: identity.fields.ID,
      email: identity.fields.Email,
      normalizedEmail: identity.fields.NormalizedEmail,
      userName: identity.fields.UserName,
      normalizedUserName: identity.fields.NormalizedUserName,
      fullName: identity.fields.FullName,
      campaignId: identity.fields.CampaignID,
      roles: identity.fields.Roles.slice(),
      pages: identity.fields.Pages.slice(),
      isAdmin: identity.fields.IsAdmin,
      employmentStatus: identity.fields.EmploymentStatus,
      hireDate: identity.fields.HireDate,
      country: identity.fields.Country,
      createdAt: identity.fields.CreatedAt,
      updatedAt: identity.fields.UpdatedAt,
      phoneNumber: identity.fields.PhoneNumber,
      phoneNumberConfirmed: status.phone ? status.phone.confirmed : false,
      canLogin: status.canLogin,
      emailConfirmed: status.emailConfirmed,
      lockout: lockout,
      deletion: deletion,
      termination: termination,
      probation: probation,
      insurance: insurance,
      mfa: {
        enabled: mfaStatus.enabled,
        deliveryMethod: mfaStatus.deliveryMethod,
        hasRecoveryCodes: Array.isArray(mfaStatus.recoveryCodes) && mfaStatus.recoveryCodes.length > 0,
        phoneDelivery: mfaStatus.phoneDelivery,
        totpConfigured: mfaStatus.totpConfigured
      },
      lastLoginAt: lastLogin.at,
      lastLoginIp: lastLogin.ip,
      lastLoginUserAgent: lastLogin.userAgent,
      passwordResetRequired: status.passwordReset ? status.passwordReset.required : false,
      accessFailedCount: lockout.accessFailedCount || 0
    };
  }

  function buildIdentityState(record) {
    const raw = pickIdentityFields(record);
    const fields = normalizeIdentityFields(raw);

    const lockoutEnd = fields.LockoutEnd;
    const lockoutEnabled = fields.LockoutEnabled;
    const lockoutActive = lockoutEnabled && lockoutEnd && lockoutEnd.getTime() > now().getTime();

    const deletedAt = fields.DeletedAt;
    const terminatedAt = fields.TerminationDate;
    const terminated = !!terminatedAt && terminatedAt.getTime() <= now().getTime();

    const probationEnd = calculateProbationEnd(fields);
    const probationActive = !!probationEnd && ensureDateOnly(probationEnd).getTime() >= ensureDateOnly(now()).getTime();

    const insuranceStatus = buildInsuranceStatus(fields);

    const mfaRecoveryCodes = fields.TwoFactorRecoveryCodes.concat(fields.MFABackupCodes);
    const hasTotpSecret = !!fields.TwoFactorSecret || !!fields.MFASecret;
    const deliveryPreference = fields.MFADeliveryPreference || fields.TwoFactorDelivery;

    const status = {
      canLogin: fields.CanLogin,
      emailConfirmed: fields.EmailConfirmed,
      resetRequired: fields.ResetRequired,
      lockout: {
        enabled: lockoutEnabled,
        locked: lockoutActive,
        lockoutEnd: lockoutEnd,
        accessFailedCount: fields.AccessFailedCount
      },
      deletion: {
        deleted: !!deletedAt,
        deletedAt: deletedAt
      },
      termination: {
        terminated: terminated,
        terminationDate: terminatedAt
      },
      probation: {
        onProbation: probationActive,
        probationEndsAt: probationEnd,
        probationMonths: fields.ProbationMonths
      },
      insurance: insuranceStatus,
      mfa: {
        enabled: fields.TwoFactorEnabled || fields.MFAEnabled,
        legacyEnabled: fields.TwoFactorEnabled,
        mfaEnabled: fields.MFAEnabled,
        deliveryMethod: deliveryPreference,
        secretConfigured: hasTotpSecret,
        recoveryCodes: mfaRecoveryCodes,
        phoneDelivery: (deliveryPreference === 'sms' || deliveryPreference === 'phone') && !!fields.PhoneNumber,
        totpConfigured: hasTotpSecret
      },
      passwordReset: {
        required: fields.ResetRequired,
        token: fields.ResetPasswordToken,
        tokenHash: fields.ResetPasswordTokenHash,
        sentAt: fields.ResetPasswordSentAt,
        expiresAt: fields.ResetPasswordExpiresAt
      },
      emailConfirmation: {
        token: fields.EmailConfirmation,
        tokenHash: fields.EmailConfirmationTokenHash,
        sentAt: fields.EmailConfirmationSentAt,
        expiresAt: fields.EmailConfirmationExpiresAt,
        confirmed: fields.EmailConfirmed
      },
      phone: {
        number: fields.PhoneNumber,
        confirmed: fields.PhoneNumberConfirmed
      },
      lastLogin: {
        at: fields.LastLoginAt || fields.LastLogin,
        ip: fields.LastLoginIp,
        userAgent: fields.LastLoginUserAgent
      },
      security: {
        stamp: fields.SecurityStamp,
        concurrencyStamp: fields.ConcurrencyStamp
      }
    };

    return {
      found: true,
      raw: raw,
      fields: fields,
      status: status,
      summary: {
        id: fields.ID,
        email: fields.Email,
        userName: fields.UserName,
        fullName: fields.FullName,
        campaignId: fields.CampaignID,
        roles: fields.Roles.slice(),
        isAdmin: fields.IsAdmin
      }
    };
  }

  function buildIdentityStateFromUser(userRecord) {
    if (!userRecord) {
      return null;
    }
    return buildIdentityState(userRecord);
  }

  function evaluateIdentityForAuthentication(identity) {
    if (!identity || !identity.status) {
      return {
        allow: false,
        error: 'Account not found.',
        errorCode: 'IDENTITY_NOT_FOUND'
      };
    }

    const status = identity.status;
    const warnings = [];
    let allow = true;
    let error = null;
    let errorCode = null;

    if (!status.canLogin) {
      allow = false;
      error = 'Your account has been disabled. Please contact support.';
      errorCode = 'ACCOUNT_DISABLED';
    } else if (status.deletion && status.deletion.deleted) {
      allow = false;
      error = 'This account has been deleted.';
      errorCode = 'ACCOUNT_DELETED';
    } else if (status.termination && status.termination.terminated) {
      allow = false;
      error = 'This account is no longer active.';
      errorCode = 'ACCOUNT_TERMINATED';
    } else if (status.lockout && status.lockout.enabled && status.lockout.locked) {
      allow = false;
      error = 'Too many failed attempts. Please try again later.';
      errorCode = 'ACCOUNT_LOCKED';
    }

    if (!status.emailConfirmed) {
      warnings.push('EMAIL_NOT_CONFIRMED');
    }
    if (status.passwordReset && status.passwordReset.required) {
      warnings.push('PASSWORD_RESET_REQUIRED');
    }
    if (status.probation && status.probation.onProbation) {
      warnings.push('PROBATION_ACTIVE');
    }
    if (status.insurance && status.insurance.eligible && !status.insurance.enrolled) {
      warnings.push('INSURANCE_ELIGIBLE_NOT_ENROLLED');
    }
    if (status.insurance && status.insurance.qualified && !status.insurance.signedUp) {
      warnings.push('INSURANCE_QUALIFIED_NOT_SIGNED_UP');
    }
    if (status.phone && !status.phone.confirmed && status.phone.number) {
      warnings.push('PHONE_NOT_CONFIRMED');
    }
    if (status.mfa && status.mfa.enabled && !status.mfa.secretConfigured) {
      warnings.push('MFA_CONFIG_INCOMPLETE');
    }
    if (status.lockout && status.lockout.accessFailedCount >= 5) {
      warnings.push('HIGH_FAILED_LOGIN_ATTEMPTS');
    }

    return {
      allow: allow,
      error: error,
      errorCode: errorCode,
      warnings: warnings,
      status: status,
      summary: identity.summary
    };
  }

  function listIdentityFields() {
    return ensureUserIdentityFields().slice();
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

  function findUserRowById(id) {
    if (!id && id !== 0) {
      return null;
    }
    const normalized = coerceString(id).trim();
    if (!normalized) {
      return null;
    }
    return findUserRow(function (row) {
      return coerceString(row.ID).trim() === normalized;
    });
  }

  function mapIdentityFromMatch(match) {
    if (!match || !match.user) {
      return null;
    }
    const identity = buildIdentityState(match.user);
    identity.rowIndex = match.rowIndex;
    identity.headers = match.context && match.context.headers ? match.context.headers.slice() : [];
    identity.sheet = match.context && match.context.sheet ? match.context.sheet : null;
    return identity;
  }

  function getUserIdentityByEmail(email) {
    const match = findUserRowByEmail(email);
    if (!match) {
      return { success: false, error: 'User not found', errorCode: 'USER_NOT_FOUND' };
    }
    const identity = mapIdentityFromMatch(match);
    return {
      success: true,
      identity: identity,
      evaluation: evaluateIdentityForAuthentication(identity),
      summary: summarizeIdentityForClient(identity)
    };
  }

  function getUserIdentityById(id) {
    const match = findUserRowById(id);
    if (!match) {
      return { success: false, error: 'User not found', errorCode: 'USER_NOT_FOUND' };
    }
    const identity = mapIdentityFromMatch(match);
    return {
      success: true,
      identity: identity,
      evaluation: evaluateIdentityForAuthentication(identity),
      summary: summarizeIdentityForClient(identity)
    };
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

  function applyPasswordUpdate(rowContext, passwordUpdate) {
    if (!rowContext || !passwordUpdate) {
      return;
    }

    var columns = {};
    if (passwordUpdate.columns && typeof passwordUpdate.columns === 'object') {
      Object.keys(passwordUpdate.columns).forEach(function (key) {
        columns[key] = passwordUpdate.columns[key];
      });
    }

    if (typeof columns.PasswordHash === 'undefined' && typeof passwordUpdate.hash !== 'undefined') {
      columns.PasswordHash = passwordUpdate.hash;
    }

    Object.keys(columns).forEach(function (column) {
      if (!hasColumn(rowContext, column)) {
        return;
      }
      var value = columns[column];
      if (typeof value === 'undefined') {
        value = '';
      }
      setColumnValue(rowContext, column, value);
    });

    if (passwordUpdate.algorithm && hasColumn(rowContext, 'PasswordHashAlgorithm')) {
      setColumnValue(rowContext, 'PasswordHashAlgorithm', passwordUpdate.algorithm);
    }
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
        const confirmed = parseBooleanFlag(match.user.EmailConfirmed);
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
      const passwordUpdate = utils.createPasswordUpdate(newPassword);
      applyPasswordUpdate(match, passwordUpdate);

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
      let identity = null;
      let identityEvaluation = null;

      try {
        const identityLookupResult = getUserIdentityByEmail(normalizedEmail);
        if (identityLookupResult && identityLookupResult.success) {
          identity = identityLookupResult.identity;
          identityEvaluation = evaluateIdentityForAuthentication(identity);
          if (identityEvaluation && !identityEvaluation.allow) {
            return {
              success: false,
              error: identityEvaluation.error || 'Your account is not eligible to login.',
              errorCode: identityEvaluation.errorCode || 'IDENTITY_BLOCKED',
              identityWarnings: identityEvaluation.warnings || []
            };
          }
        }
      } catch (identityErr) {
        console.warn('IdentityService.signIn: identity preflight failed', identityErr);
      }

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
      if (result) {
        const authenticated = result.success || result.needsMfa || result.needsPasswordReset;
        if (authenticated && identity && !result.identity) {
          result.identity = summarizeIdentityForClient(identity);
        }
        if (identityEvaluation && identityEvaluation.warnings && identityEvaluation.warnings.length && authenticated) {
          result.identityWarnings = (result.identityWarnings || []).concat(identityEvaluation.warnings);
        }
        if (result.success && options && options.returnUrl) {
          const requestedReturn = sanitizeLoginReturnUrl(options.returnUrl);
          if (requestedReturn) {
            result.requestedReturnUrl = requestedReturn;
          }
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
    sanitizeLoginReturnUrl: sanitizeLoginReturnUrl,
    getUserIdentityByEmail: getUserIdentityByEmail,
    getUserIdentityById: getUserIdentityById,
    evaluateIdentityForAuthentication: evaluateIdentityForAuthentication,
    summarizeIdentityForClient: summarizeIdentityForClient,
    listIdentityFields: listIdentityFields,
    buildIdentityStateFromUser: buildIdentityStateFromUser
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

function identityGetUserIdentityByEmail(email) {
  try {
    return IdentityService.getUserIdentityByEmail(email);
  } catch (error) {
    return { success: false, error: error.message || String(error) };
  }
}

function identityGetUserIdentityById(id) {
  try {
    return IdentityService.getUserIdentityById(id);
  } catch (error) {
    return { success: false, error: error.message || String(error) };
  }
}

function identityEvaluateForAuthentication(identity) {
  try {
    return IdentityService.evaluateIdentityForAuthentication(identity);
  } catch (error) {
    return { allow: false, error: error.message || String(error), errorCode: 'IDENTITY_ERROR' };
  }
}

function identitySummarize(identity) {
  try {
    return IdentityService.summarizeIdentityForClient(identity);
  } catch (error) {
    return null;
  }
}

function identityListFields() {
  try {
    return IdentityService.listIdentityFields();
  } catch (error) {
    return [];
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
