/**
 * AuthenticationService.gs - Fixed Token-Based Authentication Service
 * Addresses common authentication issues in Lumina
 * 
 * Key Fixes:
 * - Consistent email normalization
 * - Robust password verification
 * - Better error handling
 * - Unified user lookup
 * - Proper empty password detection
 */

// ───────────────────────────────────────────────────────────────────────────────
// AUTHENTICATION CONFIGURATION
// ───────────────────────────────────────────────────────────────────────────────

const SESSION_TTL_MS = 30 * 60 * 1000; // 30 minutes
const REMEMBER_ME_TTL_MS = 24 * 60 * 60 * 1000; // 24 hours
const SESSION_EXPIRATION_ENABLED = false; // Disable automatic session expiration

const BASE_SESSION_COLUMNS = [
  'Token',
  'TokenHash',
  'TokenSalt',
  'UserId',
  'CreatedAt',
  'LastActivityAt'
];

const OPTIONAL_SESSION_COLUMNS = [
  'ExpiresAt',
  'IdleTimeoutMinutes',
  'RememberMe',
  'CampaignScope',
  'UserAgent',
  'IpAddress',
  'ServerIp'
];

const DEFAULT_SESSION_COLUMNS = BASE_SESSION_COLUMNS.concat(OPTIONAL_SESSION_COLUMNS);

const SESSION_COLUMNS = (function deriveSessionColumns() {
  const source = (typeof SESSIONS_HEADERS !== 'undefined' && Array.isArray(SESSIONS_HEADERS) && SESSIONS_HEADERS.length)
    ? SESSIONS_HEADERS.slice()
    : DEFAULT_SESSION_COLUMNS.slice();

  const unique = [];
  source.forEach(function (column) {
    const normalized = String(column || '').trim();
    if (!normalized) return;
    if (unique.indexOf(normalized) === -1) {
      unique.push(normalized);
    }
  });

  BASE_SESSION_COLUMNS.forEach(function (column) {
    if (unique.indexOf(column) === -1) {
      unique.push(column);
    }
  });

  return unique;
})();

const DEFAULT_IDLE_TIMEOUT_MINUTES = 30;

const MFA_CHALLENGE_TTL_SECONDS = 5 * 60; // 5 minutes
const MFA_MAX_ATTEMPTS = 5;
const MFA_MAX_DELIVERIES = 5;
const MFA_CODE_LENGTH = 6;
const MFA_STORAGE_PREFIX = 'AUTH_MFA_CHALLENGE:';

const DEVICE_VERIFICATION_CODE_LENGTH = 6;
const DEVICE_VERIFICATION_TTL_MS = 15 * 60 * 1000; // 15 minutes
const TRUSTED_DEVICE_TABLE = (typeof TRUSTED_DEVICES_SHEET === 'string' && TRUSTED_DEVICES_SHEET)
  ? TRUSTED_DEVICES_SHEET
  : 'TrustedDevices';
const TRUSTED_DEVICE_COLUMNS = [
  'ID',
  'UserId',
  'Fingerprint',
  'IpAddress',
  'ServerIp',
  'UserAgent',
  'Platform',
  'Languages',
  'TimezoneOffsetMinutes',
  'Status',
  'CreatedAt',
  'UpdatedAt',
  'ConfirmedAt',
  'LastSeenAt',
  'PendingVerificationId',
  'PendingVerificationExpiresAt',
  'PendingVerificationCodeHash',
  'PendingMetadataJson',
  'PendingRememberMe',
  'MetadataJson',
  'DeniedAt',
  'DenialReason'
];

const LOGIN_CONTEXT_CACHE_PREFIX = 'AUTH_LOGIN_CONTEXT:';
const LOGIN_CONTEXT_CACHE_TTL_SECONDS = 5 * 60; // 5 minutes

// ───────────────────────────────────────────────────────────────────────────────
// IMPROVED AUTHENTICATION SERVICE
// ───────────────────────────────────────────────────────────────────────────────

var AuthenticationService = (function () {

  // ─── Password utilities with error handling ─────────────────────────────────
  
  function getPasswordUtils() {
    try {
      if (typeof ensurePasswordUtilities === 'function') {
        return ensurePasswordUtilities();
      }
      if (typeof PasswordUtilities !== 'undefined' && PasswordUtilities) {
        return PasswordUtilities;
      }
      throw new Error('PasswordUtilities not available');
    } catch (error) {
      console.error('Error getting password utilities:', error);
      throw new Error('Password utilities not available');
    }
  }

  // ─── Consistent normalization helpers ─────────────────────────────────────────

  function normalizeEmail(email) {
    if (!email && email !== 0) return '';
    return String(email).trim().toLowerCase();
  }

  function normalizeString(str) {
    if (!str && str !== 0) return '';
    return String(str).trim();
  }

  function normalizeCampaignId(value) {
    return normalizeString(value);
  }

  function cleanCampaignList(list) {
    if (!Array.isArray(list)) return [];
    var seen = {};
    var result = [];
    for (var i = 0; i < list.length; i++) {
      var key = normalizeCampaignId(list[i]);
      if (!key || seen[key]) continue;
      seen[key] = true;
      result.push(key);
    }
    return result;
  }

  function parseCampaignScope(rawScope) {
    if (!rawScope && rawScope !== 0) return null;
    if (typeof rawScope === 'object') {
      return rawScope;
    }
    if (typeof rawScope === 'string') {
      try {
        return JSON.parse(rawScope);
      } catch (parseError) {
        console.warn('parseCampaignScope: Failed to parse scope JSON', parseError);
      }
    }
    return null;
  }

  function serializeCampaignScope(scope) {
    if (!scope || typeof scope !== 'object') return '';
    try {
      return JSON.stringify(scope);
    } catch (err) {
      console.warn('serializeCampaignScope: Failed to stringify scope', err);
      return '';
    }
  }

  function tenantSecurityAvailable() {
    return typeof TenantSecurity !== 'undefined'
      && TenantSecurity
      && typeof TenantSecurity.getAccessProfile === 'function';
  }

  function toBool(value) {
    if (value === true || value === false) return value;
    const str = normalizeString(value).toUpperCase();
    return str === 'TRUE' || str === '1' || str === 'YES' || str === 'Y';
  }

  const MFA_ALLOWED_METHODS = ['email', 'sms', 'totp'];

  function sanitizeClientMetadata(metadata) {
    if (!metadata || typeof metadata !== 'object') {
      return null;
    }

    const sanitized = {};
    if (metadata.userAgent) {
      sanitized.userAgent = String(metadata.userAgent).slice(0, 500);
    }
    if (metadata.ipAddress) {
      sanitized.ipAddress = String(metadata.ipAddress).slice(0, 100);
    }
    if (metadata.serverIp || metadata.serverObservedIp) {
      sanitized.serverIp = String(metadata.serverIp || metadata.serverObservedIp).slice(0, 100);
    }
    if (metadata.forwardedFor) {
      sanitized.forwardedFor = String(metadata.forwardedFor).slice(0, 250);
    }
    if (metadata.platform) {
      sanitized.platform = String(metadata.platform).slice(0, 100);
    }
    if (metadata.language) {
      sanitized.language = String(metadata.language).slice(0, 50);
    }
    if (Array.isArray(metadata.languages) && metadata.languages.length) {
      sanitized.languages = metadata.languages.slice(0, 5).map(function (lang) {
        return String(lang).slice(0, 50);
      }).filter(function (lang) { return lang.length > 0; });
    }
    if (typeof metadata.timezoneOffsetMinutes === 'number' && isFinite(metadata.timezoneOffsetMinutes)) {
      sanitized.timezoneOffsetMinutes = Math.round(metadata.timezoneOffsetMinutes);
    }
    if (metadata.originHost) {
      sanitized.originHost = String(metadata.originHost).slice(0, 200);
    }
    if (typeof metadata.deviceMemory === 'number' && isFinite(metadata.deviceMemory)) {
      sanitized.deviceMemory = Math.max(0, Math.round(metadata.deviceMemory));
    }
    if (typeof metadata.hardwareConcurrency === 'number' && isFinite(metadata.hardwareConcurrency)) {
      sanitized.hardwareConcurrency = Math.max(1, Math.round(metadata.hardwareConcurrency));
    }
    if (metadata.serverObservedAt) {
      sanitized.serverObservedAt = String(metadata.serverObservedAt);
    }
    if (metadata.observedAt) {
      sanitized.observedAt = String(metadata.observedAt);
    }

    if (metadata.logoutReason) {
      sanitized.logoutReason = String(metadata.logoutReason).slice(0, 20);
    }

    if (metadata.requestedReturnUrl) {
      try {
        if (typeof IdentityService !== 'undefined'
          && IdentityService
          && typeof IdentityService.sanitizeLoginReturnUrl === 'function') {
          const sanitizedReturn = IdentityService.sanitizeLoginReturnUrl(metadata.requestedReturnUrl);
          if (sanitizedReturn) {
            sanitized.requestedReturnUrl = sanitizedReturn;
          }
        }
      } catch (returnError) {
        console.warn('sanitizeClientMetadata: unable to sanitize requestedReturnUrl', returnError);
      }
    }

    return Object.keys(sanitized).length ? sanitized : null;
  }

  function resolveScriptBaseUrl() {
    try {
      if (typeof SCRIPT_URL === 'string' && SCRIPT_URL) {
        return SCRIPT_URL;
      }
    } catch (err) {
      console.warn('resolveScriptBaseUrl: SCRIPT_URL lookup failed', err);
    }

    try {
      if (typeof getBaseUrl === 'function') {
        const base = getBaseUrl();
        if (base) {
          return base;
        }
      }
    } catch (err) {
      console.warn('resolveScriptBaseUrl: getBaseUrl helper failed', err);
    }

    try {
      if (typeof ScriptApp !== 'undefined' && ScriptApp && ScriptApp.getService) {
        const serviceUrl = ScriptApp.getService().getUrl();
        if (serviceUrl) {
          return serviceUrl;
        }
      }
    } catch (err) {
      console.warn('resolveScriptBaseUrl: ScriptApp URL lookup failed', err);
    }

    return '';
  }

  function sanitizeReturnUrlCandidate(candidate) {
    if (!candidate && candidate !== 0) {
      return '';
    }

    try {
      const raw = String(candidate).trim();
      if (!raw) {
        return '';
      }

      if (/^javascript:/i.test(raw)) {
        return '';
      }

      let baseUrl = '';
      const resolvedBase = resolveScriptBaseUrl();
      if (resolvedBase) {
        baseUrl = resolvedBase;
      }

      let parsed;
      try {
        parsed = baseUrl ? new URL(raw, baseUrl) : new URL(raw);
      } catch (parseError) {
        if (baseUrl) {
          try {
            parsed = new URL(raw, baseUrl);
          } catch (fallbackError) {
            console.warn('sanitizeReturnUrlCandidate: unable to resolve URL', fallbackError);
            return '';
          }
        } else {
          console.warn('sanitizeReturnUrlCandidate: unable to parse URL', parseError);
          return '';
        }
      }

      if (!/^https?:$/i.test(parsed.protocol)) {
        return '';
      }

      if (baseUrl) {
        try {
          const base = new URL(baseUrl);
          if (parsed.host && base.host && parsed.host.toLowerCase() !== base.host.toLowerCase()) {
            return '';
          }
        } catch (hostError) {
          console.warn('sanitizeReturnUrlCandidate: host comparison failed', hostError);
        }
      }

      let sanitized = parsed.toString();
      if (sanitized.length > 500) {
        sanitized = sanitized.slice(0, 500);
      }

      return sanitized;
    } catch (error) {
      console.warn('sanitizeReturnUrlCandidate: fallback sanitation failed', error);
      try {
        if (typeof IdentityService !== 'undefined'
          && IdentityService
          && typeof IdentityService.sanitizeLoginReturnUrl === 'function') {
          return IdentityService.sanitizeLoginReturnUrl(candidate);
        }
      } catch (identityError) {
        console.warn('sanitizeReturnUrlCandidate: IdentityService fallback failed', identityError);
      }
      return '';
    }
  }

  // ─── Session storage helpers ────────────────────────────────────────────────

  function getSessionTableName() {
    return (typeof SESSIONS_SHEET === 'string' && SESSIONS_SHEET) ? SESSIONS_SHEET : 'Sessions';
  }

  function generateTokenSalt() {
    try {
      const source = Utilities.getUuid() + ':' + Utilities.getUuid() + ':' + Date.now();
      const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, source);
      return Utilities.base64EncodeWebSafe(digest);
    } catch (error) {
      console.warn('generateTokenSalt: Falling back to UUID salt generation', error);
      return (typeof Utilities !== 'undefined' && Utilities.getUuid)
        ? Utilities.getUuid().replace(/[^A-Za-z0-9]/g, '').slice(0, 32)
        : String(Math.random()).replace(/[^A-Za-z0-9]/g, '').slice(0, 32);
    }
  }

  function computeSessionTokenHash(token, salt) {
    if (!token || !salt) return '';
    try {
      const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + token);
      return Utilities.base64EncodeWebSafe(digest);
    } catch (error) {
      console.warn('computeSessionTokenHash: Failed to compute digest', error);
      return '';
    }
  }

  function setRecordValue(record, key, value) {
    if (!record || !key) return;
    const normalized = String(key).trim();
    if (!normalized) return;
    record[normalized] = value;
    const lower = normalized.toLowerCase();
    if (lower !== normalized) {
      record[lower] = value;
    }
  }

  function getRecordValue(record, key) {
    if (!record || !key) return null;
    const normalized = String(key).trim();
    if (!normalized) return null;
    if (Object.prototype.hasOwnProperty.call(record, normalized)) {
      return record[normalized];
    }
    const lower = normalized.toLowerCase();
    if (Object.prototype.hasOwnProperty.call(record, lower)) {
      return record[lower];
    }
    return null;
  }

  function parseDateValue(value) {
    if (value instanceof Date) {
      const ms = value.getTime();
      return isNaN(ms) ? null : ms;
    }
    if (!value && value !== 0) return null;
    const parsed = Date.parse(String(value));
    return isNaN(parsed) ? null : parsed;
  }

  function parseIdleTimeoutMinutes(value) {
    if (value === null || typeof value === 'undefined' || value === '') {
      return DEFAULT_IDLE_TIMEOUT_MINUTES;
    }
    const numeric = Number(value);
    if (!isFinite(numeric) || numeric <= 0) {
      return DEFAULT_IDLE_TIMEOUT_MINUTES;
    }
    return Math.max(1, Math.round(numeric));
  }

  function scrubLegacySessionTokens(sheet, headers) {
    if (!sheet || !Array.isArray(headers) || !headers.length) {
      return;
    }

    const tokenIndex = headers.indexOf('Token');
    if (tokenIndex === -1) {
      return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return;
    }

    const rows = lastRow - 1;
    const tokenColumnRange = sheet.getRange(2, tokenIndex + 1, rows, 1);
    let requiresScrub = false;
    try {
      const tokenColumnValues = tokenColumnRange.getValues();
      for (let i = 0; i < tokenColumnValues.length; i++) {
        const value = tokenColumnValues[i] && tokenColumnValues[i][0];
        if (value && String(value).trim()) {
          requiresScrub = true;
          break;
        }
      }
    } catch (columnReadError) {
      console.warn('scrubLegacySessionTokens: Unable to read token column', columnReadError);
      return;
    }

    if (!requiresScrub) {
      return;
    }

    const hashIndex = headers.indexOf('TokenHash');
    const saltIndex = headers.indexOf('TokenSalt');

    try {
      const dataRange = sheet.getRange(2, 1, rows, headers.length);
      const values = dataRange.getValues();
      let changed = false;

      for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
        const row = values[rowIndex];
        const tokenValue = row[tokenIndex];
        const normalizedToken = tokenValue && String(tokenValue).trim();
        if (!normalizedToken) {
          continue;
        }

        let rowChanged = false;
        let hashValue = hashIndex === -1 ? '' : String(row[hashIndex] || '').trim();
        let saltValue = saltIndex === -1 ? '' : String(row[saltIndex] || '').trim();
        const hadHash = !!hashValue;
        const hadSalt = !!saltValue;

        if (!hadHash || !hadSalt) {
          try {
            const freshSalt = generateTokenSalt();
            const freshHash = computeSessionTokenHash(normalizedToken, freshSalt);
            if (freshHash) {
              if (hashIndex !== -1) {
                row[hashIndex] = freshHash;
              }
              if (saltIndex !== -1) {
                row[saltIndex] = freshSalt;
              }
              rowChanged = true;
            }
          } catch (hashError) {
            console.warn('scrubLegacySessionTokens: Failed to backfill token hash', hashError);
          }
        }

        hashValue = hashIndex === -1 ? hashValue : String(row[hashIndex] || '').trim();
        saltValue = saltIndex === -1 ? saltValue : String(row[saltIndex] || '').trim();

        if ((!hashValue || !saltValue) && row[tokenIndex]) {
          row[tokenIndex] = '';
          rowChanged = true;
        }

        if (rowChanged) {
          changed = true;
        }
      }

      if (changed) {
        dataRange.setValues(values);
      }
    } catch (scrubError) {
      console.warn('scrubLegacySessionTokens: Failed to sanitize session sheet', scrubError);
    }
  }

  function ensureSessionSheetContext() {
    const tableName = getSessionTableName();
    let sheet = null;
    const headerTargetColumns = (SESSION_COLUMNS && SESSION_COLUMNS.length
      ? SESSION_COLUMNS.slice()
      : DEFAULT_SESSION_COLUMNS.slice());

    BASE_SESSION_COLUMNS.forEach(function (column) {
      if (headerTargetColumns.indexOf(column) === -1) {
        headerTargetColumns.push(column);
      }
    });

    if (typeof ensureSheetWithHeaders === 'function') {
      try {
        sheet = ensureSheetWithHeaders(tableName, headerTargetColumns);
      } catch (ensureError) {
        console.warn('ensureSessionSheetContext: ensureSheetWithHeaders failed', ensureError);
      }
    }

    if (!sheet) {
      if (typeof SpreadsheetApp === 'undefined') {
        throw new Error('SpreadsheetApp not available for session storage');
      }
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        throw new Error('Active spreadsheet not available for session storage');
      }
      sheet = ss.getSheetByName(tableName);
      if (!sheet) {
        sheet = ss.insertSheet(tableName);
        sheet.getRange(1, 1, 1, headerTargetColumns.length).setValues([headerTargetColumns]);
      }
    }

    let headerValues = [];
    try {
      const lastColumn = sheet.getLastColumn();
      headerValues = lastColumn > 0
        ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0].slice()
        : [];
    } catch (headerError) {
      console.warn('ensureSessionSheetContext: Failed to read existing headers', headerError);
      headerValues = [];
    }

    const normalizedHeaders = headerValues.map(function (value) { return String(value || '').trim(); });
    let headerUpdated = false;
    headerTargetColumns.forEach(function (column) {
      if (normalizedHeaders.indexOf(column) === -1) {
        headerValues.push(column);
        normalizedHeaders.push(column);
        headerUpdated = true;
      }
    });

    if (!headerValues.length) {
      headerValues = headerTargetColumns.slice();
      headerUpdated = true;
    }

    if (headerUpdated) {
      try {
        const width = headerValues.length;
        sheet.getRange(1, 1, 1, width).setValues([headerValues]);
      } catch (setHeaderError) {
        console.warn('ensureSessionSheetContext: Failed to update headers', setHeaderError);
      }
    }

    const headers = headerValues.map(function (value) { return String(value || '').trim(); });

    try {
      scrubLegacySessionTokens(sheet, headers);
    } catch (scrubError) {
      console.warn('ensureSessionSheetContext: Failed to scrub legacy tokens', scrubError);
    }

    return {
      tableName: tableName,
      sheet: sheet,
      headers: headers
    };
  }

  function buildHeaderMap(headers) {
    const map = {};
    if (!Array.isArray(headers)) return map;
    headers.forEach(function (header, index) {
      const normalized = String(header || '').trim();
      if (!normalized) return;
      if (!Object.prototype.hasOwnProperty.call(map, normalized)) {
        map[normalized] = index;
      }
      const lower = normalized.toLowerCase();
      if (!Object.prototype.hasOwnProperty.call(map, lower)) {
        map[lower] = index;
      }
    });
    return map;
  }

  function readSessionRecord(headers, rowValues) {
    const record = {};
    if (!Array.isArray(headers) || !Array.isArray(rowValues)) {
      return record;
    }
    headers.forEach(function (header, index) {
      const normalized = String(header || '').trim();
      if (!normalized) return;
      setRecordValue(record, normalized, rowValues[index]);
    });
    return record;
  }

  function sessionTokenMatches(record, sessionToken) {
    if (!record || !sessionToken) {
      return { matched: false };
    }

    const storedSalt = getRecordValue(record, 'TokenSalt');
    const storedHash = getRecordValue(record, 'TokenHash');

    if (storedSalt && storedHash) {
      try {
        const computed = computeSessionTokenHash(sessionToken, storedSalt);
        if (computed && normalizeString(computed) === normalizeString(storedHash)) {
          return { matched: true, method: 'hash' };
        }
      } catch (hashError) {
        console.warn('sessionTokenMatches: hash comparison failed', hashError);
      }
    }

    const legacyToken = normalizeString(getRecordValue(record, 'Token'));
    if (legacyToken && normalizeString(sessionToken) === legacyToken) {
      return { matched: true, method: 'legacy' };
    }

    return { matched: false };
  }

  function findSessionEntry(sessionToken) {
    if (!sessionToken) return null;
    try {
      const context = ensureSessionSheetContext();
      const sheet = context.sheet;
      const headers = context.headers;
      const columnCount = headers.length;
      const lastRow = sheet.getLastRow();

      if (lastRow < 2 || columnCount === 0) {
        return null;
      }

      const range = sheet.getRange(2, 1, lastRow - 1, columnCount);
      const values = range.getValues();
      const headerMap = buildHeaderMap(headers);

      for (let i = 0; i < values.length; i++) {
        const rowValues = values[i];
        const record = readSessionRecord(headers, rowValues);
        const match = sessionTokenMatches(record, sessionToken);
        if (match && match.matched) {
          return {
            tableName: context.tableName,
            sheet: sheet,
            headers: headers,
            headerMap: headerMap,
            rowIndex: i + 2,
            record: record,
            rowValues: Array.isArray(rowValues) ? rowValues.slice() : [],
            matchMethod: match.method || 'hash'
          };
        }
      }

      return null;
    } catch (error) {
      console.warn('findSessionEntry: Failed to locate session', error);
      return null;
    }
  }

  function updateSessionRow(entry) {
    if (!entry || !entry.sheet || !Array.isArray(entry.headers)) return;
    try {
      const headers = entry.headers;
      const rowValues = new Array(headers.length);

      for (let i = 0; i < headers.length; i++) {
        const header = String(headers[i] || '').trim();
        if (!header) {
          rowValues[i] = (entry.rowValues && typeof entry.rowValues[i] !== 'undefined') ? entry.rowValues[i] : '';
          continue;
        }

        if (Object.prototype.hasOwnProperty.call(entry.record, header)) {
          rowValues[i] = entry.record[header];
        } else {
          const lower = header.toLowerCase();
          if (Object.prototype.hasOwnProperty.call(entry.record, lower)) {
            rowValues[i] = entry.record[lower];
          } else if (entry.rowValues && typeof entry.rowValues[i] !== 'undefined') {
            rowValues[i] = entry.rowValues[i];
          } else {
            rowValues[i] = '';
          }
        }
      }

      entry.sheet.getRange(entry.rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
      entry.rowValues = rowValues;
    } catch (error) {
      console.warn('updateSessionRow: Failed to update session row', error);
    }
  }

  function removeSessionEntry(entry) {
    if (!entry || !entry.sheet) return false;
    try {
      entry.sheet.deleteRow(entry.rowIndex);
      if (typeof invalidateCache === 'function' && entry.tableName) {
        try {
          invalidateCache(entry.tableName);
        } catch (cacheError) {
          console.warn('removeSessionEntry: Cache invalidation failed', cacheError);
        }
      }
      return true;
    } catch (error) {
      console.warn('removeSessionEntry: Failed to delete session row', error);
      return false;
    }
  }

  function evaluateSessionEntry(entry, options) {
    if (!entry) {
      return { status: 'not_found', reason: 'NOT_FOUND' };
    }

    const nowMs = Date.now();
    const record = entry.record;
    const expiryTime = parseDateValue(getRecordValue(record, 'ExpiresAt'));
    const rememberFlag = toBool(getRecordValue(record, 'RememberMe'));
    const idleTimeoutMinutes = parseIdleTimeoutMinutes(getRecordValue(record, 'IdleTimeoutMinutes'));
    const lastActivityTime = parseDateValue(getRecordValue(record, 'LastActivityAt'))
      || parseDateValue(getRecordValue(record, 'CreatedAt'));

    if (SESSION_EXPIRATION_ENABLED) {
      if (!expiryTime || expiryTime < nowMs) {
        removeSessionEntry(entry);
        return {
          status: 'expired',
          reason: 'EXPIRED',
          idleTimeoutMinutes: idleTimeoutMinutes,
          lastActivityAt: lastActivityTime ? new Date(lastActivityTime).toISOString() : null
        };
      }

      if (lastActivityTime && (nowMs - lastActivityTime) > idleTimeoutMinutes * 60 * 1000) {
        removeSessionEntry(entry);
        return {
          status: 'expired',
          reason: 'IDLE_TIMEOUT',
          idleTimeoutMinutes: idleTimeoutMinutes,
          lastActivityAt: lastActivityTime ? new Date(lastActivityTime).toISOString() : null
        };
      }
    }

    let expiresAtIso = getRecordValue(record, 'ExpiresAt');
    let lastActivityIso = getRecordValue(record, 'LastActivityAt')
      || (lastActivityTime ? new Date(lastActivityTime).toISOString() : null);

    const touch = options && options.touch;
    const sessionToken = options && options.sessionToken;

    if (touch) {
      const nowIso = new Date(nowMs).toISOString();
      const ttl = rememberFlag ? REMEMBER_ME_TTL_MS : SESSION_TTL_MS;
      const nextExpiryIso = new Date(nowMs + ttl).toISOString();

      setRecordValue(record, 'LastActivityAt', nowIso);
      setRecordValue(record, 'ExpiresAt', nextExpiryIso);
      setRecordValue(record, 'IdleTimeoutMinutes', String(idleTimeoutMinutes));

      if (sessionToken && (entry.matchMethod === 'legacy' || !getRecordValue(record, 'TokenHash') || !getRecordValue(record, 'TokenSalt'))) {
        const salt = generateTokenSalt();
        const hash = computeSessionTokenHash(sessionToken, salt);
        if (hash) {
          setRecordValue(record, 'TokenSalt', salt);
          setRecordValue(record, 'TokenHash', hash);
        }
      }

      try {
        updateSessionRow(entry);
        if (typeof invalidateCache === 'function') {
          try {
            invalidateCache(entry.tableName);
          } catch (cacheError) {
            console.warn('evaluateSessionEntry: Cache invalidation failed', cacheError);
          }
        }
      } catch (updateError) {
        console.warn('evaluateSessionEntry: Failed to persist session updates', updateError);
      }

      expiresAtIso = nextExpiryIso;
      lastActivityIso = nowIso;
    }

    return {
      status: 'active',
      entry: entry,
      tableName: entry.tableName,
      idleTimeoutMinutes: idleTimeoutMinutes,
      lastActivityAt: lastActivityIso,
      expiresAt: expiresAtIso,
      rememberMe: rememberFlag
    };
  }

  function resolveSessionRecord(sessionToken, options) {
    if (!sessionToken) {
      return { status: 'not_found', reason: 'NOT_FOUND' };
    }

    const entry = findSessionEntry(sessionToken);
    const evaluation = evaluateSessionEntry(entry, Object.assign({}, options || {}, { sessionToken: sessionToken }));
    return evaluation;
  }

  function deriveLoginReturnUrlFromEvent(event) {
    try {
      if (!event || typeof event !== 'object') {
        return '';
      }

      const parameters = event.parameter || event.parameters || {};
      if (!parameters || typeof parameters !== 'object') {
        return '';
      }

      const directKeys = ['returnUrl', 'returnURL', 'ReturnUrl', 'ReturnURL'];
      for (let i = 0; i < directKeys.length; i++) {
        const key = directKeys[i];
        if (Object.prototype.hasOwnProperty.call(parameters, key) && parameters[key]) {
          const sanitizedDirect = sanitizeReturnUrlCandidate(parameters[key]);
          if (sanitizedDirect) {
            return sanitizedDirect;
          }
        }
      }

      const rawPage = parameters.page || parameters.Page || parameters.PAGE || '';
      const page = String(rawPage || '').trim();
      if (!page || page.toLowerCase() === 'login') {
        return '';
      }

      const additionalParams = {};
      let campaignId = '';
      Object.keys(parameters).forEach(function (key) {
        if (!key) return;
        if (/^page$/i.test(key)) return;
        if (/^token$/i.test(key)) return;
        if (/^returnurl$/i.test(key)) return;

        const value = parameters[key];
        if (value === null || typeof value === 'undefined' || value === '') {
          return;
        }

        if (!campaignId && /^campaign$/i.test(key)) {
          campaignId = value;
          return;
        }

        additionalParams[key] = value;
      });

      let builtUrl = '';
      try {
        if (typeof getAuthenticatedUrl === 'function') {
          builtUrl = getAuthenticatedUrl(page, campaignId, additionalParams);
        }
      } catch (buildError) {
        console.warn('deriveLoginReturnUrlFromEvent: getAuthenticatedUrl failed', buildError);
        builtUrl = '';
      }

      if (!builtUrl) {
        const base = resolveScriptBaseUrl();
        const parts = ['page=' + encodeURIComponent(page)];
        if (campaignId) {
          parts.push('campaign=' + encodeURIComponent(campaignId));
        }
        Object.keys(additionalParams).forEach(function (key) {
          parts.push(encodeURIComponent(key) + '=' + encodeURIComponent(additionalParams[key]));
        });

        if (base) {
          const separator = base.indexOf('?') === -1 ? '?' : (/[?&]$/.test(base) ? '' : '&');
          builtUrl = base + (parts.length ? separator + parts.join('&') : '');
        } else if (parts.length) {
          builtUrl = '?' + parts.join('&');
        }
      }

      return sanitizeReturnUrlCandidate(builtUrl);
    } catch (error) {
      console.warn('deriveLoginReturnUrlFromEvent: unable to determine return URL', error);
      return '';
    }
  }

  function buildSessionEntryFromRow(context, rowIndex, rowValues, headerMap) {
    if (!context || !context.sheet || !Array.isArray(context.headers)) {
      return null;
    }

    const record = readSessionRecord(context.headers, rowValues);
    return {
      tableName: context.tableName,
      sheet: context.sheet,
      headers: context.headers,
      headerMap: headerMap || buildHeaderMap(context.headers),
      rowIndex: rowIndex,
      record: record,
      rowValues: Array.isArray(rowValues) ? rowValues.slice() : [],
      matchMethod: 'user'
    };
  }

  function findActiveSessionForUser(userId, options) {
    const normalizedUserId = normalizeString(userId);
    if (!normalizedUserId) {
      return null;
    }

    try {
      const context = ensureSessionSheetContext();
      const sheet = context.sheet;
      const headers = context.headers;
      const lastRow = sheet.getLastRow();
      if (lastRow < 2 || !Array.isArray(headers) || !headers.length) {
        return null;
      }

      const headerMap = buildHeaderMap(headers);
      const columnCount = headers.length;
      const range = sheet.getRange(2, 1, lastRow - 1, columnCount);
      const values = range.getValues();

      let latest = null;
      let latestTimestamp = -Infinity;

      for (let i = 0; i < values.length; i++) {
        const rowValues = values[i];
        const entry = buildSessionEntryFromRow(context, i + 2, rowValues, headerMap);
        if (!entry) {
          continue;
        }

        const recordUserId = normalizeString(getRecordValue(entry.record, 'UserId'));
        if (recordUserId !== normalizedUserId) {
          continue;
        }

        const evaluation = evaluateSessionEntry(entry, options || {});
        if (!evaluation || evaluation.status !== 'active') {
          continue;
        }

        const activityTimestamp = Date.parse(evaluation.lastActivityAt || evaluation.expiresAt || '') || 0;
        if (activityTimestamp >= latestTimestamp) {
          latest = evaluation;
          latestTimestamp = activityTimestamp;
        }
      }

      return latest;
    } catch (error) {
      console.warn('findActiveSessionForUser: failed to locate active session', error);
      return null;
    }
  }

  function userHasActiveSession(userIdentifier) {
    try {
      if (!userIdentifier && userIdentifier !== 0) {
        return false;
      }

      let userId = '';
      if (typeof userIdentifier === 'object' && userIdentifier !== null) {
        userId = normalizeString(userIdentifier.ID || userIdentifier.Id || userIdentifier.userId || userIdentifier.UserId);
      } else {
        userId = normalizeString(userIdentifier);
      }

      if (!userId && typeof userIdentifier === 'object' && userIdentifier !== null) {
        const email = normalizeEmail(userIdentifier.Email || userIdentifier.email);
        if (email) {
          const user = findUserByEmail(email);
          if (user && user.ID) {
            userId = normalizeString(user.ID);
          }
        }
      }

      if (!userId) {
        return false;
      }

      const active = findActiveSessionForUser(userId, { touch: false });
      return !!(active && active.status === 'active');
    } catch (error) {
      console.warn('userHasActiveSession: unable to determine session state', error);
      return false;
    }
  }

  function buildSessionUserContext(entry, sessionToken, resolution) {
    if (!entry || !entry.record) {
      return null;
    }

    const record = entry.record;
    const userId = getRecordValue(record, 'UserId');
    if (!userId) {
      return null;
    }

    const user = findUserById(userId) || findUserByEmail(userId);
    if (!user) {
      return null;
    }

    const rawScope = parseCampaignScope(
      getRecordValue(record, 'CampaignScope') || getRecordValue(record, 'campaignScope')
    );
    const tenantPayload = buildTenantScopePayload(rawScope);
    const userPayload = buildUserPayload(user, tenantPayload);

    if (userPayload && userPayload.CampaignScope) {
      userPayload.CampaignScope.tenantContext = rawScope && rawScope.tenantContext ? rawScope.tenantContext : null;
      if (rawScope && Array.isArray(rawScope.assignments)) {
        userPayload.CampaignScope.assignments = rawScope.assignments.slice();
      }
      if (rawScope && Array.isArray(rawScope.permissions)) {
        userPayload.CampaignScope.permissions = rawScope.permissions.slice();
      }
    }

    if (userPayload) {
      userPayload.sessionToken = sessionToken;

      const expiresIso = resolution && resolution.expiresAt
        ? resolution.expiresAt
        : (getRecordValue(record, 'ExpiresAt') || null);
      if (expiresIso) {
        userPayload.sessionExpiry = expiresIso;
        userPayload.sessionExpiresAt = expiresIso;
      }

      const lastActivityIso = resolution && resolution.lastActivityAt
        ? resolution.lastActivityAt
        : (getRecordValue(record, 'LastActivityAt') || null);
      if (lastActivityIso) {
        userPayload.sessionLastActivityAt = lastActivityIso;
      }

      const idleTimeoutMinutes = resolution && typeof resolution.idleTimeoutMinutes !== 'undefined'
        ? resolution.idleTimeoutMinutes
        : parseIdleTimeoutMinutes(getRecordValue(record, 'IdleTimeoutMinutes'));
      if (idleTimeoutMinutes) {
        userPayload.sessionIdleTimeoutMinutes = idleTimeoutMinutes;
      }

      userPayload.sessionScope = rawScope || null;
      userPayload.NeedsCampaignAssignment = userPayload.CampaignScope
        ? !!userPayload.CampaignScope.needsCampaignAssignment
        : false;
    }

    return {
      user: userPayload,
      tenant: tenantPayload,
      rawScope: rawScope,
      rawUser: user
    };
  }

  function cleanupExpiredSessions() {
    try {
      const context = ensureSessionSheetContext();
      const sheet = context.sheet;
      const headers = context.headers;
      const columnCount = headers.length;
      const lastRow = sheet.getLastRow();

      if (lastRow < 2 || columnCount === 0) {
        console.log('cleanupExpiredSessions: No session rows to evaluate');
        return { success: true, removed: 0, evaluated: 0 };
      }

      const range = sheet.getRange(2, 1, lastRow - 1, columnCount);
      const values = range.getValues();
      const nowMs = Date.now();
      const rowsToDelete = [];
      const reasonCounts = {};

      for (let i = 0; i < values.length; i++) {
        const rowValues = values[i];
        const record = readSessionRecord(headers, rowValues);
        const expiryTime = parseDateValue(getRecordValue(record, 'ExpiresAt'));
        const idleTimeoutMinutes = parseIdleTimeoutMinutes(getRecordValue(record, 'IdleTimeoutMinutes'));
        const lastActivityTime = parseDateValue(getRecordValue(record, 'LastActivityAt'))
          || parseDateValue(getRecordValue(record, 'CreatedAt'));

        let shouldRemove = false;
        let reason = 'UNKNOWN';

        if (!expiryTime || expiryTime < nowMs) {
          shouldRemove = true;
          reason = 'EXPIRED';
        } else if (lastActivityTime && (nowMs - lastActivityTime) > idleTimeoutMinutes * 60 * 1000) {
          shouldRemove = true;
          reason = 'IDLE_TIMEOUT';
        }

        if (shouldRemove) {
          rowsToDelete.push(i + 2);
          reasonCounts[reason] = (reasonCounts[reason] || 0) + 1;
        }
      }

      let removed = 0;
      if (rowsToDelete.length) {
        rowsToDelete.sort(function (a, b) { return b - a; });
        rowsToDelete.forEach(function (rowIndex) {
          try {
            sheet.deleteRow(rowIndex);
            removed++;
          } catch (deleteError) {
            console.warn('cleanupExpiredSessions: Failed to delete row', rowIndex, deleteError);
          }
        });

        if (removed && typeof invalidateCache === 'function') {
          try {
            invalidateCache(context.tableName);
          } catch (cacheError) {
            console.warn('cleanupExpiredSessions: Cache invalidation failed', cacheError);
          }
        }
      }

      console.log('cleanupExpiredSessions: Removed ' + removed + ' sessions (reasons: ' + JSON.stringify(reasonCounts) + ').');
      return { success: true, removed: removed, evaluated: values.length, reasons: reasonCounts };
    } catch (error) {
      console.error('cleanupExpiredSessions: Error during cleanup', error);
      return { success: false, error: error.message };
    }
  }

  function sanitizeServerMetadata(metadata) {
    if (!metadata || typeof metadata !== 'object') {
      return null;
    }

    const sanitized = {};
    if (metadata.clientAddress) {
      sanitized.serverIp = String(metadata.clientAddress).slice(0, 100);
    }
    if (metadata.forwardedFor) {
      sanitized.forwardedFor = String(metadata.forwardedFor).slice(0, 250);
    }
    if (metadata.userAgent) {
      sanitized.serverUserAgent = String(metadata.userAgent).slice(0, 500);
    }
    if (metadata.host) {
      sanitized.host = String(metadata.host).slice(0, 200);
    }
    sanitized.serverObservedAt = new Date().toISOString();

    return Object.keys(sanitized).length ? sanitized : null;
  }

  function mergeClientAndServerMetadata(clientMetadata, serverMetadata) {
    const client = sanitizeClientMetadata(clientMetadata) || {};
    const server = sanitizeServerMetadata(serverMetadata) || {};
    const merged = {};

    Object.keys(client).forEach(function (key) {
      merged[key] = client[key];
    });

    Object.keys(server).forEach(function (key) {
      if (!Object.prototype.hasOwnProperty.call(merged, key) || !merged[key]) {
        merged[key] = server[key];
      }
    });

    if (server && server.serverIp) {
      merged.serverIp = server.serverIp;
    }

    return Object.keys(merged).length ? merged : null;
  }

  function resolveObservedIp(metadata) {
    if (!metadata || typeof metadata !== 'object') {
      return '';
    }

    const candidates = [metadata.serverIp, metadata.ipAddress, metadata.forwardedFor];
    for (let i = 0; i < candidates.length; i++) {
      const candidate = normalizeString(candidates[i]);
      if (!candidate) {
        continue;
      }
      const primary = candidate.indexOf(',') !== -1 ? candidate.split(',')[0].trim() : candidate;
      if (primary) {
        return primary;
      }
    }
    return '';
  }

  function getLoginContextCacheKey() {
    try {
      if (typeof Session !== 'undefined' && Session && typeof Session.getTemporaryActiveUserKey === 'function') {
        const key = Session.getTemporaryActiveUserKey();
        if (key) {
          return LOGIN_CONTEXT_CACHE_PREFIX + String(key);
        }
      }
    } catch (err) {
      console.warn('getLoginContextCacheKey: unable to derive key', err);
    }
    return null;
  }

  function persistLoginContext(metadata) {
    const key = getLoginContextCacheKey();
    if (!key) {
      return;
    }

    try {
      const serialized = JSON.stringify(metadata || {});
      if (typeof CacheService !== 'undefined' && CacheService) {
        try {
          CacheService.getUserCache().put(key, serialized, LOGIN_CONTEXT_CACHE_TTL_SECONDS);
        } catch (cacheError) {
          console.warn('persistLoginContext: cache put failed', cacheError);
        }
      }
      if (typeof PropertiesService !== 'undefined' && PropertiesService) {
        try {
          PropertiesService.getUserProperties().setProperty(key, serialized);
        } catch (propError) {
          console.warn('persistLoginContext: property set failed', propError);
        }
      }
    } catch (err) {
      console.warn('persistLoginContext: failed to serialize metadata', err);
    }
  }

  function consumeLoginContext() {
    const key = getLoginContextCacheKey();
    if (!key) {
      return null;
    }

    let raw = null;
    if (typeof CacheService !== 'undefined' && CacheService) {
      try {
        raw = CacheService.getUserCache().get(key);
        CacheService.getUserCache().remove(key);
      } catch (cacheError) {
        console.warn('consumeLoginContext: cache get/remove failed', cacheError);
      }
    }

    if (!raw && typeof PropertiesService !== 'undefined' && PropertiesService) {
      try {
        const props = PropertiesService.getUserProperties();
        raw = props.getProperty(key);
        props.deleteProperty(key);
      } catch (propError) {
        console.warn('consumeLoginContext: property read/delete failed', propError);
      }
    }

    if (!raw) {
      return null;
    }

    try {
      return JSON.parse(raw);
    } catch (parseError) {
      console.warn('consumeLoginContext: failed to parse metadata', parseError);
      return null;
    }
  }

  function captureLoginRequestContext(event) {
    try {
      const serverContext = event && event.context ? {
        clientAddress: event.context.clientAddress,
        forwardedFor: event.context.forwardedFor,
        host: event.context.host,
        userAgent: event.context.userAgent
      } : null;
      const sanitizedServer = sanitizeServerMetadata(serverContext);
      const requestedReturnUrl = deriveLoginReturnUrlFromEvent(event);

      let payload = null;
      if (sanitizedServer) {
        payload = Object.assign({}, sanitizedServer);
      }
      if (requestedReturnUrl) {
        payload = payload || {};
        payload.requestedReturnUrl = requestedReturnUrl;
      }

      if (payload) {
        persistLoginContext(payload);
      }
      return payload;
    } catch (error) {
      console.warn('captureLoginRequestContext: failed to capture context', error);
      return null;
    }
  }

  function maskEmail(email) {
    const normalized = normalizeEmail(email);
    if (!normalized) {
      return '';
    }

    const parts = normalized.split('@');
    if (parts.length !== 2) {
      return normalized;
    }

    const local = parts[0];
    const domain = parts[1];
    if (local.length <= 2) {
      return local.charAt(0) + '***@' + domain;
    }
    return local.charAt(0) + '***' + local.charAt(local.length - 1) + '@' + domain;
  }

  function buildDeviceFingerprint(userId, metadata) {
    if (!userId || !metadata || typeof metadata !== 'object') {
      return null;
    }

    try {
      const parts = [
        normalizeString(userId),
        normalizeString(metadata.userAgent),
        normalizeString(metadata.platform),
        normalizeString(metadata.language),
        Array.isArray(metadata.languages) ? metadata.languages.join(',') : '',
        metadata.timezoneOffsetMinutes !== null && typeof metadata.timezoneOffsetMinutes !== 'undefined'
          ? String(metadata.timezoneOffsetMinutes)
          : '',
        normalizeString(metadata.serverIp || metadata.ipAddress)
      ];

      const source = parts.join('|');
      if (!source.trim()) {
        return null;
      }

      const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, source);
      return Utilities.base64EncodeWebSafe(digest);
    } catch (error) {
      console.warn('buildDeviceFingerprint: failed to generate fingerprint', error);
      return null;
    }
  }

  function ensureTrustedDevicesSheet() {
    try {
      if (typeof ensureSheetWithHeaders === 'function') {
        return ensureSheetWithHeaders(TRUSTED_DEVICE_TABLE, TRUSTED_DEVICE_COLUMNS);
      }
    } catch (error) {
      console.warn('ensureTrustedDevicesSheet: ensureSheetWithHeaders failed', error);
    }

    if (typeof SpreadsheetApp === 'undefined') {
      throw new Error('SpreadsheetApp not available for trusted device storage');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(TRUSTED_DEVICE_TABLE);
    if (!sheet) {
      sheet = ss.insertSheet(TRUSTED_DEVICE_TABLE);
      sheet.getRange(1, 1, 1, TRUSTED_DEVICE_COLUMNS.length).setValues([TRUSTED_DEVICE_COLUMNS]);
      sheet.setFrozenRows(1);
    }

    const lastCol = sheet.getLastColumn();
    const headers = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    const normalizedHeaders = headers.map(function (value) { return String(value || '').trim(); });
    let modified = false;
    TRUSTED_DEVICE_COLUMNS.forEach(function (column) {
      if (normalizedHeaders.indexOf(column) === -1) {
        normalizedHeaders.push(column);
        modified = true;
      }
    });
    if (modified) {
      sheet.getRange(1, 1, 1, normalizedHeaders.length).setValues([normalizedHeaders]);
    }

    return sheet;
  }

  function readTrustedDevices() {
    const sheet = ensureTrustedDevicesSheet();
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    if (lastRow < 2 || lastColumn === 0) {
      return [];
    }

    const range = sheet.getRange(2, 1, lastRow - 1, lastColumn);
    const values = range.getValues();
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function (value) {
      return String(value || '').trim();
    });

    return values.map(function (row, index) {
      const record = {};
      headers.forEach(function (header, colIndex) {
        record[header] = row[colIndex];
      });
      record.__rowNumber = index + 2;
      return record;
    });
  }

  function writeTrustedDeviceRecord(record) {
    if (!record || typeof record !== 'object') {
      throw new Error('writeTrustedDeviceRecord: record must be an object');
    }

    const sheet = ensureTrustedDevicesSheet();
    const headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn() || TRUSTED_DEVICE_COLUMNS.length);
    const headers = headersRange.getValues()[0].map(function (value) { return String(value || '').trim(); });
    const rowValues = headers.map(function (header) {
      return Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '';
    });

    if (record.__rowNumber) {
      sheet.getRange(record.__rowNumber, 1, 1, headers.length).setValues([rowValues]);
      return record;
    }

    sheet.appendRow(rowValues);
    record.__rowNumber = sheet.getLastRow();
    return record;
  }

  function saveTrustedDeviceRecord(record, updates) {
    const nowIso = new Date().toISOString();
    const merged = Object.assign({}, record || {});
    Object.keys(updates || {}).forEach(function (key) {
      merged[key] = updates[key];
    });
    if (!merged.CreatedAt) {
      merged.CreatedAt = nowIso;
    }
    merged.UpdatedAt = nowIso;
    return writeTrustedDeviceRecord(merged);
  }

  function findTrustedDeviceRecord(userId, fingerprint) {
    if (!userId || !fingerprint) {
      return null;
    }

    const records = readTrustedDevices();
    for (let i = 0; i < records.length; i++) {
      const record = records[i];
      if (String(record.UserId) === String(userId) && String(record.Fingerprint) === String(fingerprint)) {
        return record;
      }
    }
    return null;
  }

  function findDeviceByVerificationId(verificationId) {
    if (!verificationId) {
      return null;
    }

    const records = readTrustedDevices();
    for (let i = 0; i < records.length; i++) {
      const record = records[i];
      if (String(record.PendingVerificationId || '') === String(verificationId)) {
        return record;
      }
    }
    return null;
  }

  function sendDeviceVerificationEmailSafe(user, metadata, verificationCode, expiresAtIso) {
    if (typeof sendDeviceVerificationEmail !== 'function') {
      console.warn('sendDeviceVerificationEmailSafe: EmailService not available');
      return { success: false, error: 'EMAIL_SERVICE_UNAVAILABLE' };
    }

    try {
      const result = sendDeviceVerificationEmail(user.Email, {
        fullName: user.FullName || user.UserName || user.Email,
        verificationCode: verificationCode,
        expiresAt: expiresAtIso,
        ipAddress: resolveObservedIp(metadata) || 'Not available',
        userAgent: metadata.userAgent || 'Unknown',
        platform: metadata.platform || '',
        originHost: metadata.originHost || '',
        languages: Array.isArray(metadata.languages) ? metadata.languages : (metadata.language ? [metadata.language] : [])
      });
      if (result === false || (result && result.success === false)) {
        return { success: false, error: (result && result.error) || 'EMAIL_SEND_FAILED' };
      }
      return { success: true };
    } catch (error) {
      console.error('sendDeviceVerificationEmailSafe: failed to send email', error);
      return { success: false, error: error.message };
    }
  }

  function evaluateTrustedDevice(user, metadata, rememberMe) {
    if (!metadata) {
      return { trusted: true, metadata: null };
    }

    const fingerprint = buildDeviceFingerprint(user.ID, metadata);
    if (!fingerprint) {
      return { trusted: true, metadata: metadata };
    }

    const now = new Date();
    const nowIso = now.toISOString();
    let record = findTrustedDeviceRecord(user.ID, fingerprint);

    const status = record && record.Status ? String(record.Status).toLowerCase() : '';
    if (record && status === 'trusted') {
      const updated = saveTrustedDeviceRecord(record, {
        LastSeenAt: nowIso,
        IpAddress: metadata.ipAddress || record.IpAddress || '',
        ServerIp: metadata.serverIp || metadata.ipAddress || record.ServerIp || '',
        UserAgent: metadata.userAgent || record.UserAgent || '',
        Platform: metadata.platform || record.Platform || '',
        Languages: Array.isArray(metadata.languages) ? metadata.languages.join(',') : (metadata.language || record.Languages || ''),
        TimezoneOffsetMinutes: typeof metadata.timezoneOffsetMinutes === 'number'
          ? String(metadata.timezoneOffsetMinutes)
          : (record.TimezoneOffsetMinutes || ''),
        MetadataJson: JSON.stringify(metadata || {}),
        PendingVerificationId: '',
        PendingVerificationExpiresAt: '',
        PendingVerificationCodeHash: '',
        PendingMetadataJson: '',
        PendingRememberMe: ''
      });
      return { trusted: true, metadata: metadata, record: updated };
    }

    const verificationId = Utilities.getUuid();
    const verificationCode = generateOneTimeNumericCode(DEVICE_VERIFICATION_CODE_LENGTH);
    const codeHash = hashMfaCode(verificationCode, verificationId);
    if (!codeHash) {
      console.error('evaluateTrustedDevice: unable to hash verification code');
      return {
        trusted: false,
        error: 'Failed to initiate verification.',
        errorCode: 'DEVICE_VERIFICATION_ERROR'
      };
    }

    const expiresAtIso = new Date(now.getTime() + DEVICE_VERIFICATION_TTL_MS).toISOString();

    const baseUpdates = {
      UserId: user.ID,
      Fingerprint: fingerprint,
      Status: 'pending',
      PendingVerificationId: verificationId,
      PendingVerificationExpiresAt: expiresAtIso,
      PendingVerificationCodeHash: codeHash,
      PendingMetadataJson: JSON.stringify(metadata || {}),
      PendingRememberMe: rememberMe ? 'TRUE' : 'FALSE',
      IpAddress: metadata.ipAddress || '',
      ServerIp: metadata.serverIp || metadata.ipAddress || '',
      UserAgent: metadata.userAgent || '',
      Platform: metadata.platform || '',
      Languages: Array.isArray(metadata.languages) ? metadata.languages.join(',') : (metadata.language || ''),
      TimezoneOffsetMinutes: typeof metadata.timezoneOffsetMinutes === 'number'
        ? String(metadata.timezoneOffsetMinutes)
        : (record && record.TimezoneOffsetMinutes ? record.TimezoneOffsetMinutes : ''),
      DeniedAt: '',
      DenialReason: ''
    };

    if (!record) {
      record = saveTrustedDeviceRecord({
        ID: Utilities.getUuid(),
        CreatedAt: nowIso
      }, baseUpdates);
    } else {
      record = saveTrustedDeviceRecord(record, baseUpdates);
    }

    const emailResult = sendDeviceVerificationEmailSafe(user, metadata, verificationCode, expiresAtIso);
    if (!emailResult || emailResult.success === false) {
      console.error('evaluateTrustedDevice: verification email failed', emailResult && emailResult.error);
      return {
        trusted: false,
        error: 'We were unable to send the verification email. Please try again later.',
        errorCode: 'DEVICE_EMAIL_FAILED'
      };
    }

    return {
      trusted: false,
      verification: {
        id: verificationId,
        expiresAt: expiresAtIso,
        maskedEmail: maskEmail(user.Email),
        ipAddress: metadata.serverIp || metadata.ipAddress || '',
        codeLength: DEVICE_VERIFICATION_CODE_LENGTH,
        message: 'We emailed a verification code to confirm this device.'
      }
    };
  }

  function parseMetadataJson(value) {
    if (!value) {
      return null;
    }

    try {
      return JSON.parse(value);
    } catch (error) {
      console.warn('parseMetadataJson: failed to parse metadata', error);
      return null;
    }
  }

  function mergeMetadataForSession(recordMetadata, clientMetadata, serverMetadata) {
    const merged = {};

    const sources = [
      sanitizeClientMetadata(recordMetadata) || {},
      sanitizeServerMetadata(serverMetadata) || {},
      sanitizeClientMetadata(clientMetadata) || {}
    ];

    for (let i = 0; i < sources.length; i++) {
      const source = sources[i];
      Object.keys(source).forEach(function (key) {
        if (!source[key] && source[key] !== 0) {
          return;
        }
        merged[key] = source[key];
      });
    }

    return Object.keys(merged).length ? merged : null;
  }

  function confirmDeviceVerification(verificationId, code, clientMetadata) {
    const normalizedId = normalizeString(verificationId);
    const normalizedCode = normalizeMfaCode(code);

    if (!normalizedId) {
      return {
        success: false,
        error: 'Verification reference is required.',
        errorCode: 'INVALID_VERIFICATION'
      };
    }

    if (!normalizedCode) {
      return {
        success: false,
        error: 'Please provide the verification code from your email.',
        errorCode: 'MISSING_CODE'
      };
    }

    const record = findDeviceByVerificationId(normalizedId);
    if (!record) {
      return {
        success: false,
        error: 'Verification request not found or already processed.',
        errorCode: 'INVALID_VERIFICATION'
      };
    }

    const expiresAt = record.PendingVerificationExpiresAt ? Date.parse(record.PendingVerificationExpiresAt) : NaN;
    if (!isNaN(expiresAt) && expiresAt < Date.now()) {
      saveTrustedDeviceRecord(record, {
        Status: 'expired',
        PendingVerificationId: '',
        PendingVerificationExpiresAt: '',
        PendingVerificationCodeHash: '',
        PendingMetadataJson: '',
        PendingRememberMe: ''
      });
      return {
        success: false,
        error: 'This verification request expired. Please try signing in again.',
        errorCode: 'VERIFICATION_EXPIRED'
      };
    }

    const expectedHash = record.PendingVerificationCodeHash;
    const providedHash = hashMfaCode(normalizedCode, normalizedId);

    if (!expectedHash || expectedHash !== providedHash) {
      return {
        success: false,
        error: 'The verification code is incorrect. Double-check the email and try again.',
        errorCode: 'INVALID_CODE'
      };
    }

    const storedMetadata = parseMetadataJson(record.PendingMetadataJson) || parseMetadataJson(record.MetadataJson) || {};
    const serverContext = consumeLoginContext();
    const mergedMetadata = mergeMetadataForSession(storedMetadata, clientMetadata, serverContext);

    const rememberMe = String(record.PendingRememberMe || '').toUpperCase() === 'TRUE';

    const user = findUserById(record.UserId);
    if (!user) {
      return {
        success: false,
        error: 'User not found for verification request.',
        errorCode: 'INVALID_USER'
      };
    }

    const canLogin = toBool(user.CanLogin);
    const emailConfirmed = toBool(user.EmailConfirmed);
    const resetRequired = toBool(user.ResetRequired);

    if (!canLogin) {
      return {
        success: false,
        error: 'Your account is disabled. Contact support for assistance.',
        errorCode: 'ACCOUNT_DISABLED'
      };
    }

    if (!emailConfirmed) {
      return {
        success: false,
        error: 'Please confirm your email address before signing in.',
        errorCode: 'EMAIL_NOT_CONFIRMED'
      };
    }

    if (resetRequired) {
      return {
        success: false,
        error: 'You must reset your password before accessing the system.',
        errorCode: 'PASSWORD_RESET_REQUIRED'
      };
    }

    const tenantAccess = resolveTenantAccess(user, null);
    if (!tenantAccess || !tenantAccess.success) {
      const tenantError = formatTenantAccessError(tenantAccess);
      return {
        success: false,
        error: tenantError.error,
        errorCode: tenantError.errorCode || 'TENANT_ACCESS_DENIED'
      };
    }

    const tenantSummary = Object.assign({}, tenantAccess.clientPayload, {
      tenantContext: tenantAccess.sessionScope && tenantAccess.sessionScope.tenantContext
        ? tenantAccess.sessionScope.tenantContext
        : null
    });
    if (Array.isArray(tenantAccess.warnings)) {
      tenantSummary.warnings = tenantAccess.warnings.slice();
    }
    tenantSummary.needsCampaignAssignment = tenantAccess.needsCampaignAssignment === true;

    const sessionResult = createSession(user.ID, rememberMe, tenantAccess.sessionScope, mergedMetadata);
    if (!sessionResult || !sessionResult.token) {
      return {
        success: false,
        error: 'We were unable to start your session. Please try again.',
        errorCode: 'SESSION_CREATION_FAILED'
      };
    }

    try {
      updateLastLogin(user.ID);
    } catch (lastLoginError) {
      console.warn('confirmDeviceVerification: Failed to update last login', lastLoginError);
    }

    const userPayload = buildUserPayload(user, tenantAccess.clientPayload);
    if (userPayload && userPayload.CampaignScope) {
      userPayload.CampaignScope.tenantContext = tenantAccess.sessionScope && tenantAccess.sessionScope.tenantContext
        ? tenantAccess.sessionScope.tenantContext
        : null;
      if (tenantAccess.sessionScope && Array.isArray(tenantAccess.sessionScope.assignments) && !userPayload.CampaignScope.assignments.length) {
        userPayload.CampaignScope.assignments = tenantAccess.sessionScope.assignments.slice();
      }
      if (tenantAccess.sessionScope && Array.isArray(tenantAccess.sessionScope.permissions) && !userPayload.CampaignScope.permissions.length) {
        userPayload.CampaignScope.permissions = tenantAccess.sessionScope.permissions.slice();
      }
    }

    const warnings = Array.isArray(tenantAccess.warnings) ? tenantAccess.warnings.slice() : [];
    const needsCampaignAssignment = tenantAccess.needsCampaignAssignment === true;

    const nowIso = new Date().toISOString();
    saveTrustedDeviceRecord(record, {
      Status: 'trusted',
      ConfirmedAt: record.ConfirmedAt || nowIso,
      LastSeenAt: nowIso,
      PendingVerificationId: '',
      PendingVerificationExpiresAt: '',
      PendingVerificationCodeHash: '',
      PendingMetadataJson: '',
      PendingRememberMe: '',
      MetadataJson: JSON.stringify(mergedMetadata || {}),
      IpAddress: (mergedMetadata && mergedMetadata.ipAddress) || record.IpAddress || '',
      ServerIp: (mergedMetadata && mergedMetadata.serverIp) || record.ServerIp || '',
      UserAgent: (mergedMetadata && mergedMetadata.userAgent) || record.UserAgent || '',
      Platform: (mergedMetadata && mergedMetadata.platform) || record.Platform || '',
      Languages: mergedMetadata && Array.isArray(mergedMetadata.languages)
        ? mergedMetadata.languages.join(',')
        : (mergedMetadata && mergedMetadata.language) || record.Languages || '',
      TimezoneOffsetMinutes: mergedMetadata && typeof mergedMetadata.timezoneOffsetMinutes === 'number'
        ? String(mergedMetadata.timezoneOffsetMinutes)
        : (record.TimezoneOffsetMinutes || '')
    });

    const loginMessage = needsCampaignAssignment
      ? 'Login approved. Your account is not yet assigned to any campaigns.'
      : 'Login successful';

    return {
      success: true,
      sessionToken: sessionResult.token,
      user: userPayload,
      message: loginMessage,
      rememberMe: !!rememberMe,
      sessionExpiresAt: sessionResult.expiresAt,
      sessionTtlSeconds: sessionResult.ttlSeconds,
      sessionIdleTimeoutMinutes: sessionResult.idleTimeoutMinutes,
      tenant: tenantSummary,
      campaignScope: userPayload ? userPayload.CampaignScope : null,
      warnings: warnings,
      needsCampaignAssignment: needsCampaignAssignment,
      trustedDeviceVerified: true
    };
  }

  function sendDeniedDeviceAlertEmailSafe(user, record, metadata) {
    if (typeof sendDeniedDeviceAlertEmail !== 'function') {
      console.warn('sendDeniedDeviceAlertEmailSafe: EmailService notifier unavailable');
      return;
    }

    try {
      sendDeniedDeviceAlertEmail({
        userEmail: user ? (user.Email || user.UserName || user.ID) : 'Unknown',
        userName: user ? (user.FullName || user.UserName || user.Email || user.ID) : 'Unknown User',
        ipAddress: (metadata && metadata.serverIp) || record.ServerIp || record.IpAddress || 'Unknown',
        clientIp: (metadata && metadata.ipAddress) || record.IpAddress || '',
        userAgent: (metadata && metadata.userAgent) || record.UserAgent || '',
        platform: (metadata && metadata.platform) || record.Platform || '',
        occurredAt: new Date().toISOString(),
        verificationId: record.PendingVerificationId || '',
        fingerprint: record.Fingerprint || ''
      });
    } catch (error) {
      console.error('sendDeniedDeviceAlertEmailSafe: Failed to notify admins', error);
    }
  }

  function denyDeviceVerification(verificationId, clientMetadata) {
    const normalizedId = normalizeString(verificationId);
    if (!normalizedId) {
      return {
        success: false,
        error: 'Verification reference is required.',
        errorCode: 'INVALID_VERIFICATION'
      };
    }

    const record = findDeviceByVerificationId(normalizedId);
    if (!record) {
      return {
        success: false,
        error: 'Verification request not found or already processed.',
        errorCode: 'INVALID_VERIFICATION'
      };
    }

    const storedMetadata = parseMetadataJson(record.PendingMetadataJson) || parseMetadataJson(record.MetadataJson) || {};
    const serverContext = consumeLoginContext();
    const mergedMetadata = mergeMetadataForSession(storedMetadata, clientMetadata, serverContext);

    const nowIso = new Date().toISOString();
    saveTrustedDeviceRecord(record, {
      Status: 'denied',
      PendingVerificationId: '',
      PendingVerificationExpiresAt: '',
      PendingVerificationCodeHash: '',
      PendingMetadataJson: '',
      PendingRememberMe: '',
      DeniedAt: nowIso,
      DenialReason: 'User denied via login prompt'
    });

    const user = findUserById(record.UserId);
    if (user) {
      sendDeniedDeviceAlertEmailSafe(user, record, mergedMetadata || storedMetadata);
    }

    return {
      success: true,
      message: 'Thanks for letting us know. We have blocked that sign-in attempt.'
    };
  }

  function normalizeMfaDeliveryPreference(value) {
    const normalized = normalizeString(value).toLowerCase();
    if (!normalized) return '';
    if (MFA_ALLOWED_METHODS.indexOf(normalized) !== -1) {
      return normalized;
    }
    return '';
  }

  function normalizeMfaCode(code) {
    if (code === null || typeof code === 'undefined') {
      return '';
    }
    return String(code).replace(/[^0-9a-z]/gi, '').trim();
  }

  function generateOneTimeNumericCode(length) {
    const digits = Math.max(4, length || MFA_CODE_LENGTH);
    let code = '';
    while (code.length < digits) {
      const randomChunk = Utilities.getUuid().replace(/[^0-9]/g, '');
      code += randomChunk;
    }
    return code.substring(0, digits);
  }

  function padNumber(value, width) {
    const str = String(value);
    if (str.length >= width) {
      return str;
    }
    return '0'.repeat(width - str.length) + str;
  }

  function hashMfaCode(code, challengeId) {
    const normalized = normalizeMfaCode(code);
    if (!normalized) {
      return null;
    }
    try {
      const digest = Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        normalized + '|' + String(challengeId || '')
      );
      return Utilities.base64Encode(digest);
    } catch (error) {
      console.warn('hashMfaCode failed:', error);
      return null;
    }
  }

  function constantTimeEquals(a, b) {
    try {
      const utils = getPasswordUtils();
      if (utils && typeof utils.constantTimeEquals === 'function') {
        return utils.constantTimeEquals(
          utils.normalizePasswordInput(a),
          utils.normalizePasswordInput(b)
        );
      }
    } catch (utilsError) {
      console.warn('constantTimeEquals: password utilities unavailable, falling back', utilsError);
    }

    if (a == null || b == null) {
      return false;
    }
    const strA = String(a);
    const strB = String(b);
    if (strA.length !== strB.length) {
      return false;
    }
    let result = 0;
    for (let i = 0; i < strA.length; i++) {
      result |= strA.charCodeAt(i) ^ strB.charCodeAt(i);
    }
    return result === 0;
  }

  function base32ToBytes(base32) {
    if (!base32) return [];
    const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567';
    const clean = String(base32).replace(/[^A-Z2-7]/gi, '').toUpperCase();
    let bits = '';
    for (let i = 0; i < clean.length; i++) {
      const val = alphabet.indexOf(clean.charAt(i));
      if (val === -1) {
        return [];
      }
      bits += padNumber(val.toString(2), 5);
    }
    const bytes = [];
    for (let j = 0; j + 8 <= bits.length; j += 8) {
      bytes.push(parseInt(bits.substring(j, j + 8), 2));
    }
    return bytes;
  }

  function generateTotpCode(secret, timestamp, digits, stepSeconds) {
    const keyBytes = base32ToBytes(secret);
    if (!keyBytes.length) {
      return null;
    }

    const step = Math.max(15, (stepSeconds || 30)) * 1000;
    const counter = Math.floor((timestamp || Date.now()) / step);
    const counterBytes = new Array(8).fill(0);
    let tempCounter = counter;
    for (let i = 7; i >= 0; i--) {
      counterBytes[i] = tempCounter & 0xff;
      tempCounter = tempCounter >> 8;
    }

    let signature;
    try {
      signature = Utilities.computeHmacSha1Signature(counterBytes, keyBytes);
    } catch (error) {
      console.warn('generateTotpCode: Failed to compute HMAC:', error);
      return null;
    }

    if (!signature || !signature.length) {
      return null;
    }

    const offset = signature[signature.length - 1] & 0x0f;
    const binary = ((signature[offset] & 0x7f) << 24)
      | ((signature[offset + 1] & 0xff) << 16)
      | ((signature[offset + 2] & 0xff) << 8)
      | (signature[offset + 3] & 0xff);

    const modulo = Math.pow(10, digits || MFA_CODE_LENGTH);
    const otp = binary % modulo;
    return padNumber(otp, digits || MFA_CODE_LENGTH);
  }

  function verifyTotpCode(secret, code, windowSize) {
    const normalizedCode = normalizeMfaCode(code);
    if (!normalizedCode) {
      return false;
    }

    const window = typeof windowSize === 'number' ? Math.max(0, windowSize) : 1;
    for (let errorWindow = -window; errorWindow <= window; errorWindow++) {
      const timestamp = Date.now() + (errorWindow * 30 * 1000);
      const expected = generateTotpCode(secret, timestamp, normalizedCode.length, 30);
      if (expected && constantTimeEquals(expected, normalizedCode)) {
        return true;
      }
    }
    return false;
  }

  function parseMfaBackupCodes(raw) {
    if (!raw && raw !== 0) {
      return [];
    }

    if (Array.isArray(raw)) {
      return raw
        .map(value => normalizeMfaCode(value))
        .filter(Boolean);
    }

    if (typeof raw === 'string') {
      const trimmed = raw.trim();
      if (!trimmed) {
        return [];
      }
      try {
        const parsed = JSON.parse(trimmed);
        if (Array.isArray(parsed)) {
          return parsed
            .map(value => normalizeMfaCode(value))
            .filter(Boolean);
        }
      } catch (_) {
        // Ignore JSON parse errors — treat as delimited string
      }
      return trimmed
        .split(/[\s,;]+/)
        .map(value => normalizeMfaCode(value))
        .filter(Boolean);
    }

    if (typeof raw === 'object' && raw) {
      if (Array.isArray(raw.codes)) {
        return raw.codes
          .map(value => normalizeMfaCode(value))
          .filter(Boolean);
      }
    }

    return [];
  }

  function getUserMfaConfig(user) {
    if (!user || typeof user !== 'object') {
      return {
        enabled: false,
        secret: '',
        backupCodes: [],
        deliveryPreference: ''
      };
    }

    const secret = normalizeString(user.MFASecret || user.MfaSecret || user.mfaSecret);
    const deliveryPreference = normalizeMfaDeliveryPreference(
      user.MFADeliveryPreference || user.MfaDeliveryPreference || user.mfaDeliveryPreference
    );
    const backupCodes = parseMfaBackupCodes(
      user.MFABackupCodes || user.MfaBackupCodes || user.mfaBackupCodes
    );
    const smsNumber = normalizeString(user.MFAPhone || user.mfaPhone || user.Phone || user.phoneNumber);
    const explicitEnabled = toBool(user.MFAEnabled || user.mfaEnabled || user.RequireMfa || user.requireMfa);

    return {
      enabled: explicitEnabled || !!secret || backupCodes.length > 0 || !!deliveryPreference,
      secret: secret,
      backupCodes: backupCodes,
      deliveryPreference: deliveryPreference || (secret ? 'totp' : 'email'),
      smsNumber: smsNumber
    };
  }

  function selectMfaDeliveryMethod(config, override) {
    const preferred = normalizeMfaDeliveryPreference(override) || config.deliveryPreference || 'email';
    if (preferred === 'totp' && !config.secret) {
      return config.backupCodes.length ? 'email' : 'email';
    }
    if (preferred === 'sms' && !config.smsNumber) {
      return config.secret ? 'totp' : 'email';
    }
    return preferred;
  }

  function maskEmailAddress(value) {
    const email = normalizeEmail(value);
    if (!email) return '';
    const parts = email.split('@');
    if (parts.length !== 2) {
      return email.replace(/.(?=.{2})/g, '*');
    }
    const local = parts[0];
    const domain = parts[1];
    if (local.length <= 2) {
      return local.charAt(0) + '***@' + domain;
    }
    return local.substring(0, 2) + '***@' + domain;
  }

  function maskPhoneNumber(value) {
    const digits = normalizeString(value).replace(/\D/g, '');
    if (!digits) return '';
    const visible = digits.slice(-4);
    return '***-***-' + visible;
  }

  function maskDeliveryDestination(method, user) {
    if (!user) return '';
    if (method === 'email') {
      return maskEmailAddress(user.Email || user.email || user.EmailAddress);
    }
    if (method === 'sms') {
      return maskPhoneNumber(user.MFAPhone || user.mfaPhone || user.Phone || user.phoneNumber);
    }
    return '';
  }

  function getMfaStorage() {
    let cache = null;
    let properties = null;

    try {
      if (typeof CacheService !== 'undefined' && CacheService) {
        cache = CacheService.getScriptCache();
      }
    } catch (error) {
      console.warn('getMfaStorage: CacheService unavailable', error);
    }

    try {
      if (typeof PropertiesService !== 'undefined' && PropertiesService) {
        properties = PropertiesService.getScriptProperties();
      }
    } catch (error) {
      console.warn('getMfaStorage: PropertiesService unavailable', error);
    }

    return {
      get: function (key) {
        if (cache) {
          const cached = cache.get(key);
          if (cached) {
            return cached;
          }
        }
        if (properties) {
          return properties.getProperty(key);
        }
        return null;
      },
      put: function (key, value, ttlSeconds) {
        const ttl = Math.max(60, Math.min(ttlSeconds || 300, 6 * 60 * 60));
        if (cache) {
          try {
            cache.put(key, value, ttl);
          } catch (error) {
            console.warn('getMfaStorage: Failed to put cache value', error);
          }
        }
        if (properties) {
          try {
            properties.setProperty(key, value);
          } catch (error) {
            console.warn('getMfaStorage: Failed to persist property', error);
          }
        }
      },
      remove: function (key) {
        if (cache) {
          try {
            cache.remove(key);
          } catch (error) {
            console.warn('getMfaStorage: Failed to remove cache entry', error);
          }
        }
        if (properties) {
          try {
            properties.deleteProperty(key);
          } catch (error) {
            console.warn('getMfaStorage: Failed to remove property', error);
          }
        }
      }
    };
  }

  function getMfaStorageKey(challengeId) {
    return MFA_STORAGE_PREFIX + String(challengeId || '').trim();
  }

  function loadMfaChallenge(challengeId) {
    if (!challengeId) {
      return null;
    }

    const storage = getMfaStorage();
    const raw = storage.get(getMfaStorageKey(challengeId));
    if (!raw) {
      return null;
    }

    try {
      const parsed = JSON.parse(raw);
      if (parsed && parsed.expiresAt && Date.now() > parsed.expiresAt) {
        storage.remove(getMfaStorageKey(challengeId));
        return null;
      }
      return parsed;
    } catch (error) {
      console.warn('loadMfaChallenge: Failed to parse stored challenge', error);
      storage.remove(getMfaStorageKey(challengeId));
      return null;
    }
  }

  function saveMfaChallenge(challenge, ttlSeconds) {
    if (!challenge || !challenge.id) {
      return;
    }

    const storage = getMfaStorage();
    const expiresAt = challenge.expiresAt || (Date.now() + MFA_CHALLENGE_TTL_SECONDS * 1000);
    const payload = Object.assign({}, challenge, { expiresAt: expiresAt });
    storage.put(getMfaStorageKey(challenge.id), JSON.stringify(payload), ttlSeconds || MFA_CHALLENGE_TTL_SECONDS + 120);
  }

  function deleteMfaChallenge(challengeId) {
    if (!challengeId) return;
    const storage = getMfaStorage();
    storage.remove(getMfaStorageKey(challengeId));
  }

  function ensureMfaUserColumns() {
    if (typeof SpreadsheetApp === 'undefined') {
      return;
    }

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) return;
      const sheet = ss.getSheetByName('Users');
      if (!sheet) return;

      const lastColumn = sheet.getLastColumn();
      const headers = sheet.getRange(1, 1, 1, lastColumn || 1).getValues()[0];
      const normalizedHeaders = headers.map(header => normalizeString(header).toLowerCase());
      const requiredColumns = ['mfasecret', 'mfabackupcodes', 'mfadeliverypreference', 'mfaenabled'];
      const headerLabels = {
        mfasecret: 'MFASecret',
        mfabackupcodes: 'MFABackupCodes',
        mfadeliverypreference: 'MFADeliveryPreference',
        mfaenabled: 'MFAEnabled'
      };

      requiredColumns.forEach(function (column) {
        if (normalizedHeaders.indexOf(column) === -1) {
          sheet.insertColumnAfter(sheet.getLastColumn() || 1);
          const newIndex = sheet.getLastColumn();
          sheet.getRange(1, newIndex).setValue(headerLabels[column] || column);
          normalizedHeaders.push(column);
        }
      });
    } catch (error) {
      console.warn('ensureMfaUserColumns failed:', error);
    }
  }

  function updateUserMfaFields(userId, updates) {
    if (!userId || !updates || typeof updates !== 'object') {
      return false;
    }

    if (typeof SpreadsheetApp === 'undefined') {
      return false;
    }

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) return false;
      const sheet = ss.getSheetByName('Users');
      if (!sheet) return false;

      const lastColumn = Math.max(sheet.getLastColumn(), 1);
      let headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
      headers = headers.map(header => String(header || ''));

      const normalizedMap = {};
      headers.forEach(function (header, index) {
        normalizedMap[normalizeString(header).toLowerCase()] = index + 1;
      });

      const ensureColumn = function (name) {
        const key = normalizeString(name).toLowerCase();
        if (!normalizedMap[key]) {
          sheet.insertColumnAfter(sheet.getLastColumn());
          const columnIndex = sheet.getLastColumn();
          sheet.getRange(1, columnIndex).setValue(name);
          normalizedMap[key] = columnIndex;
        }
        return normalizedMap[key];
      };

      const idColumnIndex = normalizedMap.id || normalizedMap['userid'] || normalizedMap['user id'];
      if (!idColumnIndex) {
        return false;
      }

      const dataRange = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), sheet.getLastColumn());
      const dataValues = dataRange.getValues();

      for (let rowIndex = 0; rowIndex < dataValues.length; rowIndex++) {
        if (String(dataValues[rowIndex][idColumnIndex - 1]) !== String(userId)) {
          continue;
        }

        const rowNumber = rowIndex + 2;
        Object.keys(updates).forEach(function (field) {
          const columnIndex = ensureColumn(field);
          sheet.getRange(rowNumber, columnIndex).setValue(updates[field]);
        });
        return true;
      }
    } catch (error) {
      console.warn('updateUserMfaFields failed:', error);
    }

    return false;
  }

  function consumeBackupCode(userId, code, config) {
    if (!config || !config.backupCodes || !config.backupCodes.length) {
      return false;
    }

    const normalized = normalizeMfaCode(code);
    if (!normalized) return false;

    const remaining = config.backupCodes.filter(existing => existing !== normalized);
    if (remaining.length === config.backupCodes.length) {
      return false;
    }

    const updated = remaining.join('\n');
    ensureMfaUserColumns();
    updateUserMfaFields(userId, { MFABackupCodes: updated });
    config.backupCodes = remaining;
    return true;
  }

  function deliverOutOfBandCode(method, user, code, expiresAt) {
    const fallbackMessage = 'Verification code sent.';
    if (method === 'email') {
      const recipient = user.Email || user.email || user.EmailAddress || user.username || '';
      if (!recipient) {
        return { success: false, error: 'No email address configured for MFA delivery.' };
      }

      const payload = {
        code: code,
        expiresAt: new Date(expiresAt).toISOString(),
        fullName: user.FullName || user.UserName || user.username || recipient,
        deliveryMethod: 'email'
      };

      if (typeof sendMfaCodeEmail === 'function') {
        try {
          const result = sendMfaCodeEmail(recipient, payload);
          if (!result || result.success === false) {
            return { success: false, error: (result && result.error) || 'Unable to send MFA email.' };
          }
          return {
            success: true,
            message: (result && result.message) || fallbackMessage
          };
        } catch (error) {
          console.warn('deliverOutOfBandCode: sendMfaCodeEmail failed', error);
          return { success: false, error: error.message || 'Failed to send MFA email.' };
        }
      }

      if (typeof MailApp !== 'undefined' && MailApp && typeof MailApp.sendEmail === 'function') {
        try {
          MailApp.sendEmail({
            to: recipient,
            subject: 'Your LuminaHQ verification code',
            htmlBody: '<p>Your verification code is <strong>' + code + '</strong>.</p><p>This code expires in 5 minutes.</p>',
            body: 'Your verification code is ' + code + '. It expires in 5 minutes.'
          });
          return { success: true, message: fallbackMessage };
        } catch (error) {
          console.warn('deliverOutOfBandCode: MailApp sendEmail failed', error);
          return { success: false, error: 'Unable to send MFA email.' };
        }
      }

      return { success: false, error: 'Email delivery service unavailable.' };
    }

    if (method === 'sms') {
      const phone = user.MFAPhone || user.mfaPhone || user.Phone || user.phoneNumber;
      if (!phone) {
        return { success: false, error: 'No phone number configured for SMS delivery.' };
      }

      if (typeof SmsService !== 'undefined' && SmsService && typeof SmsService.sendMfaCode === 'function') {
        try {
          const result = SmsService.sendMfaCode(phone, code, { expiresAt: expiresAt });
          if (!result || result.success === false) {
            return { success: false, error: (result && result.error) || 'Unable to send SMS code.' };
          }
          return {
            success: true,
            message: (result && result.message) || 'Verification code sent via SMS.'
          };
        } catch (error) {
          console.warn('deliverOutOfBandCode: SmsService failed', error);
          return { success: false, error: error.message || 'Unable to send SMS code.' };
        }
      }

      console.warn('deliverOutOfBandCode: SMS delivery requested but SmsService not available.');
      return { success: false, error: 'SMS delivery is not available.' };
    }

    return { success: false, error: 'Unsupported MFA delivery method.' };
  }

  function createMfaChallenge(user, tenantAccess, rememberMe, metadata, configOverride) {
    const config = configOverride || getUserMfaConfig(user);
    if (!config.enabled) {
      return { success: false, reason: 'MFA_NOT_ENABLED', config: config };
    }

    const userId = user.ID || user.Id || user.id;
    if (!userId) {
      return { success: false, reason: 'INVALID_USER', config: config };
    }

    const sanitizedMetadata = sanitizeClientMetadata(metadata);
    const now = Date.now();
    const challengeId = Utilities.getUuid();
    const deliveryMethod = selectMfaDeliveryMethod(config, null);

    const challenge = {
      id: challengeId,
      userId: userId,
      userEmail: normalizeEmail(user.Email || user.email || user.EmailAddress),
      rememberMe: !!rememberMe,
      metadata: sanitizedMetadata,
      createdAt: now,
      expiresAt: now + MFA_CHALLENGE_TTL_SECONDS * 1000,
      attempts: 0,
      maxAttempts: MFA_MAX_ATTEMPTS,
      deliveries: 0,
      maxDeliveries: MFA_MAX_DELIVERIES,
      deliveryMethod: deliveryMethod,
      maskedDestination: maskDeliveryDestination(deliveryMethod, user),
      totpEnabled: deliveryMethod === 'totp',
      backupCodesRemaining: config.backupCodes.length,
      tenant: {
        sessionScope: tenantAccess.sessionScope || null,
        clientPayload: tenantAccess.clientPayload || null,
        warnings: Array.isArray(tenantAccess.warnings) ? tenantAccess.warnings.slice() : [],
        needsCampaignAssignment: tenantAccess.needsCampaignAssignment === true
      }
    };

    saveMfaChallenge(challenge);

    return {
      success: true,
      challenge: challenge,
      config: config
    };
  }

  function issueMfaChallengeCode(challenge, user, config, deliveryOverride) {
    if (!challenge || !challenge.id) {
      return { success: false, error: 'Invalid MFA challenge.' };
    }

    const deliveries = challenge.deliveries || 0;
    const maxDeliveries = challenge.maxDeliveries || MFA_MAX_DELIVERIES;
    if (deliveries >= maxDeliveries) {
      return {
        success: false,
        error: 'Maximum number of MFA code deliveries reached.',
        deliveriesRemaining: 0
      };
    }

    const method = selectMfaDeliveryMethod(config, deliveryOverride || challenge.deliveryMethod);
    if (!method) {
      return { success: false, error: 'No MFA delivery method available.' };
    }

    const now = Date.now();
    let expiresAt = now + MFA_CHALLENGE_TTL_SECONDS * 1000;
    let message = '';
    let totp = false;

    if (method === 'totp') {
      challenge.codeHash = null;
      challenge.totpEnabled = true;
      challenge.expiresAt = expiresAt;
      totp = true;
      message = 'Open your authenticator app to retrieve the current verification code.';
    } else {
      const code = generateOneTimeNumericCode(MFA_CODE_LENGTH);
      const hashed = hashMfaCode(code, challenge.id);
      if (!hashed) {
        return { success: false, error: 'Failed to generate verification code.' };
      }
      challenge.codeHash = hashed;
      challenge.expiresAt = expiresAt;
      const deliveryResult = deliverOutOfBandCode(method, user, code, expiresAt);
      if (!deliveryResult || deliveryResult.success === false) {
        return {
          success: false,
          error: (deliveryResult && deliveryResult.error) || 'Failed to deliver verification code.'
        };
      }
      message = deliveryResult.message || 'Verification code sent.';
    }

    challenge.deliveryMethod = method;
    challenge.maskedDestination = maskDeliveryDestination(method, user);
    challenge.deliveries = deliveries + 1;
    challenge.lastDeliveryAt = now;
    challenge.backupCodesRemaining = config.backupCodes.length;
    saveMfaChallenge(challenge);

    return {
      success: true,
      method: method,
      challengeId: challenge.id,
      expiresAt: new Date(challenge.expiresAt).toISOString(),
      message: message,
      maskedDestination: challenge.maskedDestination,
      totp: totp,
      deliveriesRemaining: Math.max(0, (challenge.maxDeliveries || MFA_MAX_DELIVERIES) - challenge.deliveries),
      backupCodesRemaining: config.backupCodes.length
    };
  }

  function beginMfaChallenge(challengeId, options) {
    try {
      if (!challengeId) {
        return {
          success: false,
          error: 'MFA challenge id is required.',
          errorCode: 'MFA_CHALLENGE_REQUIRED'
        };
      }

      const challenge = loadMfaChallenge(challengeId);
      if (!challenge) {
        return {
          success: false,
          error: 'The verification challenge has expired. Please sign in again.',
          errorCode: 'MFA_CHALLENGE_NOT_FOUND',
          challengeExpired: true
        };
      }

      const user = findUserById(challenge.userId);
      if (!user) {
        deleteMfaChallenge(challengeId);
        return {
          success: false,
          error: 'User account could not be located for verification.',
          errorCode: 'USER_NOT_FOUND'
        };
      }

      const config = getUserMfaConfig(user);
      if (!config.enabled) {
        deleteMfaChallenge(challengeId);
        return {
          success: false,
          error: 'Multi-factor authentication is not configured for this account.',
          errorCode: 'MFA_NOT_ENABLED'
        };
      }

      const deliveryOverride = options && options.deliveryMethod ? options.deliveryMethod : null;
      const issueResult = issueMfaChallengeCode(challenge, user, config, deliveryOverride);
      if (!issueResult || issueResult.success === false) {
        return Object.assign({
          success: false,
          errorCode: issueResult && issueResult.errorCode ? issueResult.errorCode : 'MFA_DELIVERY_FAILED'
        }, issueResult || { error: 'Failed to send MFA code.' });
      }

      return Object.assign({
        success: true,
        challengeId: challengeId,
        maskedDestination: issueResult.maskedDestination,
        backupCodesRemaining: config.backupCodes.length
      }, issueResult);
    } catch (error) {
      console.error('beginMfaChallenge error:', error);
      return {
        success: false,
        error: error.message || 'Unable to deliver verification code.',
        errorCode: 'MFA_DELIVERY_ERROR'
      };
    }
  }

  function verifyMfaCode(challengeId, code, metadata) {
    try {
      const normalizedCode = normalizeMfaCode(code);
      if (!challengeId) {
        return {
          success: false,
          error: 'Verification challenge is required.',
          errorCode: 'MFA_CHALLENGE_REQUIRED'
        };
      }

      if (!normalizedCode) {
        return {
          success: false,
          error: 'Enter the verification code from your authenticator or message.',
          errorCode: 'MFA_CODE_REQUIRED'
        };
      }

      const challenge = loadMfaChallenge(challengeId);
      if (!challenge) {
        return {
          success: false,
          error: 'The verification session has expired. Please sign in again.',
          errorCode: 'MFA_CHALLENGE_NOT_FOUND',
          challengeExpired: true
        };
      }

      if (challenge.expiresAt && Date.now() > challenge.expiresAt) {
        deleteMfaChallenge(challengeId);
        return {
          success: false,
          error: 'The verification code has expired. Please start again.',
          errorCode: 'MFA_CODE_EXPIRED',
          challengeExpired: true
        };
      }

      const maxAttempts = challenge.maxAttempts || MFA_MAX_ATTEMPTS;
      const attempts = challenge.attempts || 0;
      if (attempts >= maxAttempts) {
        deleteMfaChallenge(challengeId);
        return {
          success: false,
          error: 'Too many invalid verification attempts. Please sign in again.',
          errorCode: 'MFA_TOO_MANY_ATTEMPTS',
          challengeExpired: true
        };
      }

      const user = findUserById(challenge.userId);
      if (!user) {
        deleteMfaChallenge(challengeId);
        return {
          success: false,
          error: 'User account could not be located for verification.',
          errorCode: 'USER_NOT_FOUND'
        };
      }

      const config = getUserMfaConfig(user);
      let verified = false;
      let usedBackup = false;

      if (!verified && challenge.totpEnabled && config.secret) {
        verified = verifyTotpCode(config.secret, normalizedCode, 1);
      }

      if (!verified && challenge.codeHash) {
        const hashed = hashMfaCode(normalizedCode, challenge.id);
        if (hashed && constantTimeEquals(hashed, challenge.codeHash)) {
          verified = true;
        }
      }

      if (!verified && config.backupCodes.length) {
        if (config.backupCodes.indexOf(normalizedCode) !== -1) {
          verified = true;
          usedBackup = true;
        }
      }

      if (!verified) {
        challenge.attempts = attempts + 1;
        saveMfaChallenge(challenge);
        const remaining = Math.max(0, (challenge.maxAttempts || MFA_MAX_ATTEMPTS) - challenge.attempts);
        return {
          success: false,
          error: 'The verification code you entered is not valid.',
          errorCode: 'MFA_CODE_INVALID',
          remainingAttempts: remaining
        };
      }

      if (usedBackup) {
        consumeBackupCode(challenge.userId, normalizedCode, config);
      }

      deleteMfaChallenge(challengeId);

      const sessionMetadata = sanitizeClientMetadata(metadata) || challenge.metadata || null;
      const sessionResult = createSession(
        challenge.userId,
        !!challenge.rememberMe,
        challenge.tenant ? challenge.tenant.sessionScope : null,
        sessionMetadata
      );

      if (!sessionResult || !sessionResult.token) {
        return {
          success: false,
          error: 'Failed to create a session after verification.',
          errorCode: 'SESSION_CREATION_FAILED'
        };
      }

      try {
        updateLastLogin(challenge.userId);
      } catch (error) {
        console.warn('verifyMfaCode: Failed to update last login', error);
      }

      const tenantPayload = challenge.tenant || {};
      const tenantSummary = Object.assign({}, tenantPayload.clientPayload || {}, {
        tenantContext: tenantPayload.sessionScope && tenantPayload.sessionScope.tenantContext
          ? tenantPayload.sessionScope.tenantContext
          : null
      });
      if (Array.isArray(tenantPayload.warnings)) {
        tenantSummary.warnings = tenantPayload.warnings.slice();
      }
      tenantSummary.needsCampaignAssignment = tenantPayload.needsCampaignAssignment === true;

      const userPayload = buildUserPayload(user, tenantPayload.clientPayload || null);

      if (userPayload && userPayload.CampaignScope) {
        userPayload.CampaignScope.tenantContext = tenantPayload.sessionScope && tenantPayload.sessionScope.tenantContext
          ? tenantPayload.sessionScope.tenantContext
          : null;
        if (tenantPayload.sessionScope && Array.isArray(tenantPayload.sessionScope.assignments) && !userPayload.CampaignScope.assignments.length) {
          userPayload.CampaignScope.assignments = tenantPayload.sessionScope.assignments.slice();
        }
        if (tenantPayload.sessionScope && Array.isArray(tenantPayload.sessionScope.permissions) && !userPayload.CampaignScope.permissions.length) {
          userPayload.CampaignScope.permissions = tenantPayload.sessionScope.permissions.slice();
        }
      }

      const warnings = Array.isArray(tenantPayload.warnings) ? tenantPayload.warnings.slice() : [];

      const landing = resolveLandingDestination(user, {
        user: userPayload,
        userPayload: userPayload,
        rawUser: user,
        tenantAccess: tenantPayload,
        tenant: { clientPayload: tenantSummary, sessionScope: tenantPayload.sessionScope },
        sessionScope: tenantPayload.sessionScope
      });
      const redirectSlug = landing && landing.slug ? landing.slug : 'dashboard';
      const redirectUrl = landing && landing.redirectUrl
        ? landing.redirectUrl
        : buildLandingRedirectUrlFromSlug(redirectSlug);

      return {
        success: true,
        sessionToken: sessionResult.token,
        sessionExpiresAt: sessionResult.expiresAt,
        sessionTtlSeconds: sessionResult.ttlSeconds,
        sessionIdleTimeoutMinutes: sessionResult.idleTimeoutMinutes,
        rememberMe: !!challenge.rememberMe,
        user: userPayload,
        tenant: tenantSummary,
        campaignScope: userPayload ? userPayload.CampaignScope : null,
        warnings: warnings,
        needsCampaignAssignment: tenantPayload.needsCampaignAssignment === true,
        message: 'Verification successful. You are now signed in.',
        redirectSlug: redirectSlug,
        redirectUrl: redirectUrl
      };
    } catch (error) {
      console.error('verifyMfaCode error:', error);
      return {
        success: false,
        error: error.message || 'Unable to verify the authentication code.',
        errorCode: 'MFA_VERIFY_ERROR'
      };
    }
  }

  // ─── Improved user lookup with fallbacks ─────────────────────────────────────

  function findUserByEmail(email) {
    const normalizedEmail = normalizeEmail(email);
    if (!normalizedEmail) {
      console.log('findUserByEmail: Empty email provided');
      return null;
    }

    console.log('findUserByEmail: Looking up user with email:', normalizedEmail);

    try {
      // Method 1: Try readSheet (most reliable)
      let users = [];
      try {
        users = readSheet('Users') || [];
        console.log(`findUserByEmail: Found ${users.length} users in sheet`);
      } catch (sheetError) {
        console.warn('findUserByEmail: Sheet read failed:', sheetError);
      }

      if (users.length > 0) {
        const user = users.find(u => {
          const userEmail = normalizeEmail(u.Email);
          return userEmail === normalizedEmail;
        });
        
        if (user) {
          console.log('findUserByEmail: Found user via sheet lookup:', user.FullName || user.UserName);
          return user;
        }
      }

      // Method 2: Try DatabaseManager if available
      if (typeof DatabaseManager !== 'undefined' && DatabaseManager && typeof DatabaseManager.table === 'function') {
        try {
          const table = DatabaseManager.table('Users');
          const dbUser = table.find({ where: { Email: normalizedEmail } });
          if (dbUser && dbUser.length > 0) {
            console.log('findUserByEmail: Found user via DatabaseManager:', dbUser[0].FullName || dbUser[0].UserName);
            return dbUser[0];
          }
        } catch (dbError) {
          console.warn('findUserByEmail: DatabaseManager lookup failed:', dbError);
        }
      }

      console.log('findUserByEmail: User not found with email:', normalizedEmail);
      return null;

    } catch (error) {
      console.error('findUserByEmail: Error during lookup:', error);
      return null;
    }
  }

  function findUserById(userId) {
    const normalizedId = normalizeString(userId);
    if (!normalizedId) {
      console.log('findUserById: Empty userId provided');
      return null;
    }

    try {
      let users = [];
      try {
        users = readSheet('Users') || [];
      } catch (sheetError) {
        console.warn('findUserById: Sheet read failed:', sheetError);
      }

      if (users.length > 0) {
        const user = users.find(u => normalizeString(u.ID) === normalizedId);
        if (user) {
          return user;
        }
      }

      if (typeof DatabaseManager !== 'undefined' && DatabaseManager && typeof DatabaseManager.table === 'function') {
        try {
          const table = DatabaseManager.table('Users');
          const dbUser = table.find({ where: { ID: normalizedId } });
          if (dbUser && dbUser.length > 0) {
            return dbUser[0];
          }
        } catch (dbError) {
          console.warn('findUserById: DatabaseManager lookup failed:', dbError);
        }
      }
    } catch (error) {
      console.error('findUserById: Error during lookup:', error);
    }

    console.log('findUserById: User not found with ID:', normalizedId);
    return null;
  }

  function buildTenantScopePayload(scope) {
    if (!scope || typeof scope !== 'object') {
      return {
        isGlobalAdmin: false,
        defaultCampaignId: '',
        activeCampaignId: '',
        allowedCampaignIds: [],
        managedCampaignIds: [],
        adminCampaignIds: [],
        assignments: [],
        permissions: [],
        warnings: [],
        needsCampaignAssignment: false
      };
    }
    return {
      isGlobalAdmin: !!scope.isGlobalAdmin,
      defaultCampaignId: normalizeCampaignId(scope.defaultCampaignId),
      activeCampaignId: normalizeCampaignId(scope.activeCampaignId),
      allowedCampaignIds: cleanCampaignList(scope.allowedCampaignIds || []),
      managedCampaignIds: cleanCampaignList(scope.managedCampaignIds || []),
      adminCampaignIds: cleanCampaignList(scope.adminCampaignIds || []),
      assignments: Array.isArray(scope.assignments) ? scope.assignments.slice() : [],
      permissions: Array.isArray(scope.permissions) ? scope.permissions.slice() : [],
      warnings: Array.isArray(scope.warnings) ? scope.warnings.slice() : [],
      needsCampaignAssignment: !!scope.needsCampaignAssignment
    };
  }

  function buildUnassignedTenantScope(options) {
    const scopeOptions = options || {};
    const assignments = Array.isArray(scopeOptions.assignments) ? scopeOptions.assignments.slice() : [];
    const permissions = Array.isArray(scopeOptions.permissions) ? scopeOptions.permissions.slice() : [];
    const defaultCampaignId = normalizeCampaignId(scopeOptions.defaultCampaignId);

    const warnings = ['NO_CAMPAIGN_ASSIGNMENTS'];

    return {
      isGlobalAdmin: !!scopeOptions.isGlobalAdmin,
      defaultCampaignId: defaultCampaignId,
      activeCampaignId: '',
      allowedCampaignIds: [],
      managedCampaignIds: [],
      adminCampaignIds: [],
      tenantContext: {
        tenantIds: [],
        allowedTenantIds: [],
        limitedAccess: true,
        requireAssignment: true,
        defaultTenantId: defaultCampaignId || ''
      },
      assignments: assignments,
      permissions: permissions,
      warnings: warnings,
      needsCampaignAssignment: true
    };
  }

  function resolveTenantAccess(user, requestedCampaignId) {
    const userId = normalizeString(user && (user.ID || user.Id));
    if (!userId) {
      return { success: false, reason: 'INVALID_USER' };
    }

    const requestedId = normalizeCampaignId(requestedCampaignId);
    const fallbackCampaignId = normalizeCampaignId(user && (user.CampaignID || user.campaignId || user.CampaignId));
    const isAdmin = toBool(user && user.IsAdmin);

    if (tenantSecurityAvailable()) {
      try {
        const profile = TenantSecurity.getAccessProfile(userId);
        if (!profile) {
          throw new Error('Access profile not returned');
        }

        const allowed = cleanCampaignList(profile.allowedCampaignIds || []);
        const managed = cleanCampaignList(profile.managedCampaignIds || []);
        const admin = cleanCampaignList(profile.adminCampaignIds || []);
        const defaultCampaignId = normalizeCampaignId(profile.defaultCampaignId) || fallbackCampaignId || (allowed[0] || '');
        let activeCampaignId = '';

        if (requestedId) {
          if (!profile.isGlobalAdmin && allowed.indexOf(requestedId) === -1) {
            return { success: false, reason: 'CAMPAIGN_ACCESS_DENIED', campaignId: requestedId };
          }
          activeCampaignId = requestedId;
        } else if (defaultCampaignId && (profile.isGlobalAdmin || allowed.indexOf(defaultCampaignId) !== -1)) {
          activeCampaignId = defaultCampaignId;
        }

        if (!activeCampaignId && allowed.length) {
          activeCampaignId = allowed[0];
        }

        if (!profile.isGlobalAdmin && allowed.length === 0) {
          const unassignedScope = buildUnassignedTenantScope({
            assignments: profile.assignments,
            permissions: profile.permissions,
            defaultCampaignId: defaultCampaignId,
            isGlobalAdmin: !!profile.isGlobalAdmin
          });
          const clientPayload = buildTenantScopePayload(unassignedScope);
          return {
            success: true,
            profile: profile,
            sessionScope: unassignedScope,
            clientPayload: clientPayload,
            warnings: unassignedScope.warnings.slice(),
            needsCampaignAssignment: true
          };
        }

        const tenantContext = profile.isGlobalAdmin
          ? (activeCampaignId
            ? { tenantId: activeCampaignId, campaignId: activeCampaignId, allowAllTenants: true }
            : { allowAllTenants: true })
          : (function () {
              const ctx = {
                tenantIds: allowed.slice(),
                allowedTenantIds: allowed.slice()
              };
              if (activeCampaignId) {
                ctx.tenantId = activeCampaignId;
                ctx.campaignId = activeCampaignId;
              }
              if (defaultCampaignId) {
                ctx.defaultTenantId = defaultCampaignId;
              }
              return ctx;
            })();

        const sessionScope = {
          isGlobalAdmin: !!profile.isGlobalAdmin,
          defaultCampaignId: defaultCampaignId || '',
          activeCampaignId: activeCampaignId || '',
          allowedCampaignIds: allowed.slice(),
          managedCampaignIds: managed.slice(),
          adminCampaignIds: admin.slice(),
          tenantContext: tenantContext,
          assignments: Array.isArray(profile.assignments) ? profile.assignments : [],
          permissions: Array.isArray(profile.permissions) ? profile.permissions : []
        };

        const clientPayload = buildTenantScopePayload(sessionScope);
        clientPayload.assignments = Array.isArray(profile.assignments) ? profile.assignments : [];
        clientPayload.permissions = Array.isArray(profile.permissions) ? profile.permissions : [];

        return {
          success: true,
          profile: profile,
          sessionScope: sessionScope,
          clientPayload: clientPayload,
          warnings: Array.isArray(sessionScope.warnings) ? sessionScope.warnings.slice() : [],
          needsCampaignAssignment: !!sessionScope.needsCampaignAssignment
        };
      } catch (err) {
        console.error('resolveTenantAccess: Failed to compute tenant scope for user', userId, err);
        return { success: false, reason: 'TENANT_PROFILE_ERROR', error: err };
      }
    }

    const fallbackAllowed = cleanCampaignList([
      fallbackCampaignId,
      requestedId
    ]);
    const activeFallback = requestedId || fallbackCampaignId || (fallbackAllowed[0] || '');

    if (!isAdmin && fallbackAllowed.length === 0) {
      const unassignedFallback = buildUnassignedTenantScope({
        defaultCampaignId: fallbackCampaignId,
        isGlobalAdmin: !!isAdmin
      });
      const unassignedPayload = buildTenantScopePayload(unassignedFallback);
      return {
        success: true,
        profile: null,
        sessionScope: unassignedFallback,
        clientPayload: unassignedPayload,
        warnings: unassignedFallback.warnings.slice(),
        needsCampaignAssignment: true
      };
    }

    const fallbackContext = isAdmin
      ? { allowAllTenants: true }
      : (function () {
          const ctx = {
            tenantIds: fallbackAllowed.slice(),
            allowedTenantIds: fallbackAllowed.slice()
          };
          if (activeFallback) {
            ctx.tenantId = activeFallback;
            ctx.campaignId = activeFallback;
          }
          if (fallbackCampaignId) {
            ctx.defaultTenantId = fallbackCampaignId;
          }
          return ctx;
        })();

    const fallbackScope = {
      isGlobalAdmin: !!isAdmin,
      defaultCampaignId: fallbackCampaignId || '',
      activeCampaignId: activeFallback || '',
      allowedCampaignIds: fallbackAllowed.slice(),
      managedCampaignIds: [],
      adminCampaignIds: isAdmin ? fallbackAllowed.slice() : [],
      tenantContext: fallbackContext,
      assignments: [],
      permissions: [],
      warnings: [],
      needsCampaignAssignment: false
    };

    const fallbackPayload = buildTenantScopePayload(fallbackScope);

    return {
      success: true,
      profile: null,
      sessionScope: fallbackScope,
      clientPayload: fallbackPayload,
      warnings: Array.isArray(fallbackScope.warnings) ? fallbackScope.warnings.slice() : [],
      needsCampaignAssignment: !!fallbackScope.needsCampaignAssignment

    };
  }

  function formatTenantAccessError(tenantAccess) {
    const defaultResponse = {
      error: 'We could not determine your campaign access. Please contact support.',
      errorCode: 'TENANT_SCOPE_ERROR'
    };

    if (!tenantAccess || tenantAccess.success) {
      return defaultResponse;
    }

    switch (tenantAccess.reason) {
      case 'NO_CAMPAIGN_ASSIGNMENTS':
        return {
          error: 'Your account is not assigned to any campaigns. Please contact your administrator.',
          errorCode: 'NO_CAMPAIGN_ACCESS'
        };
      case 'CAMPAIGN_ACCESS_DENIED':
        return {
          error: 'You do not have access to the requested campaign.',
          errorCode: 'CAMPAIGN_ACCESS_DENIED'
        };
      case 'INVALID_USER':
        return {
          error: 'Unable to verify your account. Please contact support.',
          errorCode: 'INVALID_USER'
        };
      case 'TENANT_PROFILE_ERROR':
        return {
          error: 'A configuration error prevented loading your campaign permissions. Please try again later or contact support.',
          errorCode: 'TENANT_PROFILE_ERROR'
        };
      default:
        return defaultResponse;
    }
  }

  function buildUserPayload(user, tenantPayload) {
    if (!user) return null;

    const payload = {
      ID: user.ID,
      UserName: user.UserName || '',
      FullName: user.FullName || user.UserName || '',
      Email: user.Email || '',
      CampaignID: user.CampaignID || '',
      IsAdmin: toBool(user.IsAdmin),
      CanLogin: toBool(user.CanLogin),
      EmailConfirmed: toBool(user.EmailConfirmed)
    };

    if (tenantPayload && typeof tenantPayload === 'object') {
      payload.CampaignScope = {
        isGlobalAdmin: !!tenantPayload.isGlobalAdmin,
        defaultCampaignId: tenantPayload.defaultCampaignId || '',
        activeCampaignId: tenantPayload.activeCampaignId || '',
        allowedCampaignIds: (tenantPayload.allowedCampaignIds || []).slice(),
        managedCampaignIds: (tenantPayload.managedCampaignIds || []).slice(),
        adminCampaignIds: (tenantPayload.adminCampaignIds || []).slice(),
        assignments: Array.isArray(tenantPayload.assignments) ? tenantPayload.assignments : [],
        permissions: Array.isArray(tenantPayload.permissions) ? tenantPayload.permissions : [],
        warnings: Array.isArray(tenantPayload.warnings) ? tenantPayload.warnings.slice() : [],
        needsCampaignAssignment: !!tenantPayload.needsCampaignAssignment
      };
      payload.DefaultCampaignId = payload.CampaignScope.defaultCampaignId;
      payload.ActiveCampaignId = payload.CampaignScope.activeCampaignId;
      payload.AllowedCampaignIds = payload.CampaignScope.allowedCampaignIds.slice();
      payload.ManagedCampaignIds = payload.CampaignScope.managedCampaignIds.slice();
      payload.AdminCampaignIds = payload.CampaignScope.adminCampaignIds.slice();
      payload.IsGlobalAdmin = payload.CampaignScope.isGlobalAdmin;
      payload.NeedsCampaignAssignment = payload.CampaignScope.needsCampaignAssignment;
    } else {
      payload.CampaignScope = buildTenantScopePayload(null);
      payload.DefaultCampaignId = payload.CampaignScope.defaultCampaignId;
      payload.ActiveCampaignId = payload.CampaignScope.activeCampaignId;
      payload.AllowedCampaignIds = payload.CampaignScope.allowedCampaignIds.slice();
      payload.ManagedCampaignIds = payload.CampaignScope.managedCampaignIds.slice();
      payload.AdminCampaignIds = payload.CampaignScope.adminCampaignIds.slice();
      payload.IsGlobalAdmin = payload.CampaignScope.isGlobalAdmin || payload.IsAdmin;
      payload.NeedsCampaignAssignment = payload.CampaignScope.needsCampaignAssignment;

    }

    return payload;
  }

  // ─── Landing destination helpers ───────────────────────────────────────────

  function landingNormalizeText(value) {
    if (value === null || typeof value === 'undefined') return '';
    const str = String(value).trim();
    if (!str || str.toLowerCase() === 'undefined' || str.toLowerCase() === 'null') {
      return '';
    }
    return str;
  }

  function landingLowerText(value) {
    const normalized = landingNormalizeText(value);
    return normalized ? normalized.toLowerCase() : '';
  }

  function landingMatchToken(value) {
    const lower = landingLowerText(value);
    if (!lower) return '';
    return lower.replace(/[^a-z0-9]/g, '');
  }

  function landingSanitizeSlug(value) {
    const lower = landingLowerText(value);
    if (!lower) return '';
    return lower
      .replace(/[^a-z0-9\-]/g, '-')
      .replace(/-+/g, '-')
      .replace(/^-+|-+$/g, '');
  }

  function gatherLandingSources(primary, context) {
    const sources = [];
    const seen = new Set();

    function add(source) {
      if (!source || typeof source !== 'object') return;
      if (seen.has(source)) return;
      seen.add(source);
      sources.push(source);
    }

    add(primary);

    if (context && typeof context === 'object') {
      add(context.user);
      add(context.userPayload);
      add(context.rawUser);

      if (context.contextUser) add(context.contextUser);
      if (context.contextPayload) add(context.contextPayload);
    }

    return sources;
  }

  function collectLandingRoleNames(primary, context) {
    const sources = gatherLandingSources(primary, context);
    const names = [];
    const seenNames = {};
    const collectedRoleIds = new Set();
    const userIds = new Set();

    function pushName(name) {
      const normalized = landingLowerText(name);
      if (!normalized) return;
      if (seenNames[normalized]) return;
      seenNames[normalized] = true;
      names.push(normalized);
    }

    function collectRoleIds(list) {
      if (!list && list !== 0) return;
      if (Array.isArray(list)) {
        list.forEach(collectRoleIds);
        return;
      }
      if (typeof list === 'object') {
        if (Object.prototype.hasOwnProperty.call(list, 'id')) {
          collectRoleIds(list.id);
        }
        if (Object.prototype.hasOwnProperty.call(list, 'ID')) {
          collectRoleIds(list.ID);
        }
        if (Object.prototype.hasOwnProperty.call(list, 'RoleId')) {
          collectRoleIds(list.RoleId);
        }
        if (Object.prototype.hasOwnProperty.call(list, 'RoleID')) {
          collectRoleIds(list.RoleID);
        }
        return;
      }
      const normalized = landingNormalizeText(list);
      if (!normalized) return;
      normalized.split(/[,;|]/).forEach(function (part) {
        const trimmed = landingNormalizeText(part);
        if (trimmed) {
          collectedRoleIds.add(trimmed);
        }
      });
    }

    sources.forEach(function (source) {
      if (!source || typeof source !== 'object') return;

      if (Array.isArray(source.roleNames)) {
        source.roleNames.forEach(pushName);
      }
      if (Array.isArray(source.RoleNames)) {
        source.RoleNames.forEach(pushName);
      }

      pushName(source.RoleName);
      pushName(source.roleName);
      pushName(source.PrimaryRole);
      pushName(source.primaryRole);

      collectRoleIds(source.RoleIds || source.roleIds || source.RoleIDs || source.roleIDs);
      collectRoleIds(source.Roles || source.roles);

      const possibleUserId = landingNormalizeText(
        source.ID || source.Id || source.id || source.UserId || source.UserID
      );
      if (possibleUserId) {
        userIds.add(possibleUserId);
      }
    });

    if (context && typeof context === 'object') {
      if (Array.isArray(context.roleNames)) {
        context.roleNames.forEach(pushName);
      }
      collectRoleIds(context.roleIds);

      if (Array.isArray(context.userIds)) {
        context.userIds.forEach(function (userId) {
          const normalized = landingNormalizeText(userId);
          if (normalized) {
            userIds.add(normalized);
          }
        });
      }
    }

    if (collectedRoleIds.size && typeof getRolesMapping === 'function') {
      try {
        const mapping = getRolesMapping();
        collectedRoleIds.forEach(function (roleId) {
          if (!roleId) return;
          const direct = mapping[roleId];
          if (direct) {
            pushName(direct);
            return;
          }
          const lowerKey = roleId.toLowerCase();
          if (mapping[lowerKey]) {
            pushName(mapping[lowerKey]);
          }
        });
      } catch (mappingError) {
        console.warn('collectLandingRoleNames: getRolesMapping failed', mappingError);
      }
    }

    if (typeof getUserRolesSafe === 'function') {
      userIds.forEach(function (userId) {
        try {
          const roles = getUserRolesSafe(userId) || [];
          roles.forEach(function (role) {
            if (!role) return;
            pushName(role.name || role.Name || role.title || role.Title);
          });
        } catch (roleError) {
          console.warn('collectLandingRoleNames: getUserRolesSafe failed', roleError);
        }
      });
    }

    return names;
  }

  function collectLandingPageTokens(primary, context) {
    const sources = gatherLandingSources(primary, context);
    const tokens = new Set();

    function addValue(value) {
      if (value === null || typeof value === 'undefined') return;

      if (Array.isArray(value)) {
        value.forEach(addValue);
        return;
      }

      if (typeof value === 'object') {
        if (Object.prototype.hasOwnProperty.call(value, 'slug')) addValue(value.slug);
        if (Object.prototype.hasOwnProperty.call(value, 'Slug')) addValue(value.Slug);
        if (Object.prototype.hasOwnProperty.call(value, 'key')) addValue(value.key);
        if (Object.prototype.hasOwnProperty.call(value, 'Key')) addValue(value.Key);
        if (Object.prototype.hasOwnProperty.call(value, 'page')) addValue(value.page);
        if (Object.prototype.hasOwnProperty.call(value, 'Page')) addValue(value.Page);
        if (Object.prototype.hasOwnProperty.call(value, 'defaultPage')) addValue(value.defaultPage);
        if (Object.prototype.hasOwnProperty.call(value, 'DefaultPage')) addValue(value.DefaultPage);
        return;
      }

      const text = landingNormalizeText(value);
      if (!text) return;

      const directToken = landingMatchToken(text);
      if (directToken) {
        tokens.add(directToken);
      }

      text.split(/[,;|]/).forEach(function (part) {
        const token = landingMatchToken(part);
        if (token) {
          tokens.add(token);
        }
      });
    }

    const candidateKeys = [
      'Pages', 'pages', 'Page', 'DefaultPage', 'defaultPage', 'HomePage', 'homePage',
      'LandingPage', 'landingPage', 'Landing', 'landing', 'LandingSlug', 'landingSlug',
      'PreferredLanding', 'preferredLanding', 'PreferredLandingPage', 'preferredLandingPage',
      'PreferredHome', 'preferredHome', 'PrimaryPage', 'primaryPage'
    ];

    sources.forEach(function (source) {
      if (!source || typeof source !== 'object') return;
      candidateKeys.forEach(function (key) {
        if (Object.prototype.hasOwnProperty.call(source, key)) {
          addValue(source[key]);
        }
      });

      addValue(source.allowedPages || source.AllowedPages);
      addValue(source.pagesAssigned || source.PagesAssigned);
      addValue(source.pageAssignments || source.PageAssignments);
    });

    if (context && typeof context === 'object') {
      addValue(context.pages);
      addValue(context.assignedPages);
      addValue(context.availablePages);
      addValue(context.userPages);

      if (context.tenantAccess && context.tenantAccess.clientPayload) {
        addValue(context.tenantAccess.clientPayload.pages);
        addValue(context.tenantAccess.clientPayload.allowedPages);
      }

      if (context.tenant && context.tenant.clientPayload) {
        addValue(context.tenant.clientPayload.pages);
        addValue(context.tenant.clientPayload.allowedPages);
      }
    }

    return tokens;
  }

  function extractFirstLandingValue(sources, keys) {
    for (let i = 0; i < sources.length; i++) {
      const source = sources[i];
      if (!source || typeof source !== 'object') continue;
      for (let j = 0; j < keys.length; j++) {
        const key = keys[j];
        if (!Object.prototype.hasOwnProperty.call(source, key)) continue;
        const text = landingNormalizeText(source[key]);
        if (text) {
          return text;
        }
      }
    }
    return '';
  }

  function collectCombinedLandingText(sources, keys) {
    const parts = [];
    const seen = new Set();

    sources.forEach(function (source) {
      if (!source || typeof source !== 'object') return;
      keys.forEach(function (key) {
        if (!Object.prototype.hasOwnProperty.call(source, key)) return;
        const text = landingLowerText(source[key]);
        if (!text || seen.has(text)) return;
        seen.add(text);
        parts.push(text);
      });
    });

    return parts.join(' ');
  }

  function determineLandingSlug(primary, context) {
    try {
      if (context && typeof context === 'object') {
        const explicitCandidate = landingSanitizeSlug(
          context.preferredSlug
            || context.landingSlug
            || (context.user && (context.user.landingSlug || context.user.LandingSlug))
            || (context.userPayload && (context.userPayload.landingSlug || context.userPayload.LandingSlug))
        );
        if (explicitCandidate) {
          return explicitCandidate;
        }
      }

      const sources = gatherLandingSources(primary, context);
      const explicitSourceCandidate = extractFirstLandingValue(sources, [
        'LandingSlug', 'landingSlug', 'PreferredLanding', 'preferredLanding',
        'PreferredLandingPage', 'preferredLandingPage', 'PreferredHome', 'preferredHome',
        'HomePage', 'homePage', 'DefaultPage', 'defaultPage', 'LandingPage', 'landingPage'
      ]);
      const explicitSourceSlug = landingSanitizeSlug(explicitSourceCandidate);

      if (explicitSourceSlug) {
        return explicitSourceSlug;
      }

      const roleNames = collectLandingRoleNames(primary, context) || [];
      const pageTokens = collectLandingPageTokens(primary, context) || new Set();

      const jobTitleText = landingLowerText(
        collectCombinedLandingText(sources, ['JobTitle', 'jobTitle', 'Title', 'title', 'Role', 'role', 'Position', 'position', 'PrimaryRole', 'primaryRole'])
      );
      const departmentText = landingLowerText(
        collectCombinedLandingText(sources, ['Department', 'department', 'Dept', 'dept', 'Division', 'division', 'Team', 'team', 'Group', 'group', 'Organization', 'organization', 'Organisation'])
      );
      const personaText = landingLowerText(
        collectCombinedLandingText(sources, ['Persona', 'persona', 'PrimaryPersona', 'primaryPersona', 'PersonaKey', 'personaKey', 'PersonaName', 'personaName', 'PersonaLabel', 'personaLabel', 'PersonaType', 'personaType', 'AccessPersona', 'accessPersona', 'WorkspacePersona', 'workspacePersona'])
      );
      const classificationText = landingLowerText(
        collectCombinedLandingText(sources, ['Classification', 'classification', 'EmploymentStatus', 'employmentStatus', 'EmploymentType', 'employmentType', 'EmployeeType', 'employeeType', 'StaffType', 'staffType', 'UserType', 'userType', 'AccountType', 'accountType', 'AccessLevel', 'accessLevel'])
      );

      const combinedPersonaText = [personaText, classificationText].filter(Boolean).join(' ');

      function hasRoleMatch(patterns) {
        return roleNames.some(function (role) {
          return patterns.some(function (pattern) {
            return role.indexOf(pattern) !== -1;
          });
        });
      }

      function hasPage(slug) {
        const token = landingMatchToken(slug);
        return token ? pageTokens.has(token) : false;
      }

      function textHas(normalizedText, patterns) {
        if (!normalizedText) return false;
        return patterns.some(function (pattern) {
          return normalizedText.indexOf(pattern) !== -1;
        });
      }

      const agentPatterns = ['agent'];
      const guestPatterns = ['guest', 'client', 'partner', 'collab'];

      const isAgent = hasRoleMatch(agentPatterns)
        || hasPage('agentexperience')
        || hasPage('workspaceagent')
        || hasPage('userprofile')
        || textHas(jobTitleText, agentPatterns)
        || textHas(departmentText, agentPatterns)
        || textHas(combinedPersonaText, agentPatterns);

      if (isAgent) {
        return 'userprofile';
      }

      const isGuest = hasPage('collaborationreporting')
        || hasRoleMatch(guestPatterns)
        || textHas(jobTitleText, guestPatterns)
        || textHas(departmentText, guestPatterns)
        || textHas(combinedPersonaText, guestPatterns);

      if (isGuest) {
        return 'collaborationreporting';
      }

      return 'dashboard';
    } catch (error) {
      console.warn('determineLandingSlug: failed to compute landing slug', error);
      return 'dashboard';
    }
  }

  function buildLandingRedirectUrlFromSlug(slug) {
    const sanitized = landingSanitizeSlug(slug);
    const finalSlug = sanitized || 'dashboard';
    return '?page=' + encodeURIComponent(finalSlug);
  }

  function resolveLandingDestination(primary, context) {
    const slug = determineLandingSlug(primary, context);
    return {
      slug: landingSanitizeSlug(slug) || 'dashboard',
      redirectUrl: buildLandingRedirectUrlFromSlug(slug)
    };
  }

  // ─── Improved password verification ─────────────────────────────────────────

  function verifyUserPassword(inputPassword, storedHash, userInfo = {}) {
    try {
      console.log('verifyUserPassword: Starting verification for user:', userInfo.email || 'unknown');
      
      // Check if password was provided
      if (!inputPassword && inputPassword !== 0) {
        console.log('verifyUserPassword: No password provided');
        return { success: false, reason: 'NO_PASSWORD_PROVIDED' };
      }

      // Check if user has a stored hash
      const normalizedHash = normalizeString(storedHash);
      if (!normalizedHash) {
        console.log('verifyUserPassword: No stored password hash');
        return { success: false, reason: 'NO_STORED_HASH' };
      }

      console.log('verifyUserPassword: Hash length:', normalizedHash.length);

      // Get password utilities
      let passwordUtils;
      try {
        passwordUtils = getPasswordUtils();
      } catch (utilsError) {
        console.error('verifyUserPassword: Password utilities error:', utilsError);
        return { success: false, reason: 'UTILS_ERROR', error: utilsError.message };
      }

      // Attempt verification with multiple methods for robustness
      const inputStr = String(inputPassword);
      
      // Method 1: Direct verification
      try {
        const isValid = passwordUtils.verifyPassword(inputStr, normalizedHash);
        console.log('verifyUserPassword: Direct verification result:', isValid);
        
        if (isValid) {
          return { success: true, method: 'direct' };
        }
      } catch (verifyError) {
        console.warn('verifyUserPassword: Direct verification failed:', verifyError);
      }

      // Method 2: Normalize hash first, then verify
      try {
        const normalizedStoredHash = passwordUtils.normalizeHash(normalizedHash);
        const newInputHash = passwordUtils.hashPassword(inputStr);
        const matches = passwordUtils.constantTimeEquals(newInputHash, normalizedStoredHash);
        console.log('verifyUserPassword: Normalized comparison result:', matches);
        
        if (matches) {
          return { success: true, method: 'normalized' };
        }
      } catch (normalizeError) {
        console.warn('verifyUserPassword: Normalized verification failed:', normalizeError);
      }

      // Method 3: Direct hash comparison
      try {
        const newInputHash = passwordUtils.hashPassword(inputStr);
        const matches = passwordUtils.constantTimeEquals(newInputHash, normalizedHash);
        console.log('verifyUserPassword: Direct hash comparison result:', matches);
        
        if (matches) {
          return { success: true, method: 'direct_hash' };
        }
      } catch (hashError) {
        console.warn('verifyUserPassword: Direct hash comparison failed:', hashError);
      }

      console.log('verifyUserPassword: All verification methods failed');
      return { success: false, reason: 'PASSWORD_MISMATCH' };

    } catch (error) {
      console.error('verifyUserPassword: Unexpected error:', error);
      return { success: false, reason: 'VERIFICATION_ERROR', error: error.message };
    }
  }

  // ─── Session management ─────────────────────────────────────────────────────

  function createSession(userId, rememberMe = false, campaignScope, metadata) {
    try {
      const token = Utilities.getUuid() + '_' + Date.now();
      const now = new Date();
      const ttl = rememberMe ? REMEMBER_ME_TTL_MS : SESSION_TTL_MS;
      let salt = generateTokenSalt();
      let tokenHash = computeSessionTokenHash(token, salt);
      if (!tokenHash) {
        salt = generateTokenSalt();
        tokenHash = computeSessionTokenHash(token, salt);
      }

      if (!tokenHash) {
        throw new Error('Unable to compute session token hash');
      }

      const idleTimeoutMinutes = parseIdleTimeoutMinutes(metadata && metadata.idleTimeoutMinutes);
      const nowIso = now.toISOString();
      const expiresAtIso = new Date(now.getTime() + ttl).toISOString();

      const scopeData = campaignScope && typeof campaignScope === 'object'
        ? campaignScope
        : null;

      const sessionRecord = {
        Token: token,
        TokenHash: tokenHash,
        TokenSalt: salt,
        UserId: userId,
        CreatedAt: nowIso,
        LastActivityAt: nowIso,
        ExpiresAt: expiresAtIso,
        IdleTimeoutMinutes: String(idleTimeoutMinutes),
        RememberMe: rememberMe ? 'TRUE' : 'FALSE',
        CampaignScope: serializeCampaignScope(scopeData),
        UserAgent: metadata && metadata.userAgent ? metadata.userAgent : 'Google Apps Script',
        IpAddress: metadata && metadata.ipAddress ? metadata.ipAddress : 'N/A',
        ServerIp: metadata && metadata.serverIp
          ? metadata.serverIp
          : (metadata && metadata.serverObservedIp ? metadata.serverObservedIp : 'N/A')
      };

      const tableName = (typeof SESSIONS_SHEET === 'string' && SESSIONS_SHEET) ? SESSIONS_SHEET : 'Sessions';
      let persisted = false;

      if (typeof DatabaseManager !== 'undefined' && DatabaseManager && typeof DatabaseManager.table === 'function') {
        try {
          DatabaseManager.table(tableName).insert(sessionRecord);
          persisted = true;
        } catch (dbError) {
          console.warn('createSession: DatabaseManager insert failed:', dbError);
        }
      }

      if (!persisted && typeof dbCreate === 'function') {
        try {
          dbCreate(tableName, sessionRecord);
          persisted = true;
        } catch (legacyError) {
          console.warn('createSession: dbCreate fallback failed:', legacyError);
        }
      }

      if (!persisted && typeof ensureSheetWithHeaders === 'function') {
        try {
          const sessionsSheet = ensureSheetWithHeaders(tableName, SESSION_COLUMNS);
          const rowValues = SESSION_COLUMNS.map(function (column) {
            return Object.prototype.hasOwnProperty.call(sessionRecord, column)
              ? sessionRecord[column]
              : '';
          });
          sessionsSheet.appendRow(rowValues);
          persisted = true;
        } catch (sheetError) {
          console.warn('createSession: ensureSheetWithHeaders fallback failed:', sheetError);
        }
      }

      if (!persisted && typeof SpreadsheetApp !== 'undefined') {
        try {
          const ss = SpreadsheetApp.getActiveSpreadsheet();
          if (ss) {
            let sheet = ss.getSheetByName(tableName);
            if (!sheet) {
              sheet = ss.insertSheet(tableName);
              sheet.getRange(1, 1, 1, SESSION_COLUMNS.length).setValues([SESSION_COLUMNS]);
            }
            const headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
            const normalizedHeaders = headersRange.map(function (value) { return String(value || '').trim(); });
            const row = normalizedHeaders.map(function (column) {
              return Object.prototype.hasOwnProperty.call(sessionRecord, column)
                ? sessionRecord[column]
                : '';
            });
            sheet.appendRow(row);
            persisted = true;
          }
        } catch (spreadsheetError) {
          console.warn('createSession: Spreadsheet fallback failed:', spreadsheetError);
        }
      }

      if (persisted && typeof invalidateCache === 'function') {
        try { invalidateCache(tableName); } catch (cacheError) { console.warn('createSession: Cache invalidation failed:', cacheError); }
      }

      return {
        token: token,
        record: sessionRecord,
        expiresAt: sessionRecord.ExpiresAt,
        ttlSeconds: Math.max(60, Math.floor(ttl / 1000)),
        campaignScope: scopeData,
        idleTimeoutMinutes: idleTimeoutMinutes
      };

    } catch (error) {
      console.error('createSession: Error creating session:', error);
      return null;
    }
  }

  // ─── Main login function ─────────────────────────────────────────────────────

  function login(email, password, rememberMe = false, clientMetadata) {
    console.log('=== AuthenticationService.login START ===');
    console.log('Email:', email ? 'PROVIDED' : 'EMPTY');
    console.log('Password:', password ? 'PROVIDED' : 'EMPTY');
    console.log('RememberMe:', rememberMe);

    try {
      // Input validation
      const normalizedEmail = normalizeEmail(email);
      const passwordStr = normalizeString(password);
      const sanitizedMetadata = sanitizeClientMetadata(clientMetadata);

      if (!normalizedEmail) {
        console.log('login: Invalid email provided');
        return {
          success: false,
          error: 'Email is required',
          errorCode: 'MISSING_EMAIL'
        };
      }

      if (!passwordStr) {
        console.log('login: Invalid password provided');
        return {
          success: false,
          error: 'Password is required',
          errorCode: 'MISSING_PASSWORD'
        };
      }

      console.log('login: Looking up user...');

      // Find user
      const user = findUserByEmail(normalizedEmail);
      if (!user) {
        console.log('login: User not found');
        return {
          success: false,
          error: 'Invalid email or password',
          errorCode: 'INVALID_CREDENTIALS'
        };
      }

      console.log('login: Found user:', user.FullName || user.UserName);

      // Check account status
      const canLogin = toBool(user.CanLogin);
      const emailConfirmed = toBool(user.EmailConfirmed);
      const resetRequired = toBool(user.ResetRequired);

      console.log('login: Account status - CanLogin:', canLogin, 'EmailConfirmed:', emailConfirmed, 'ResetRequired:', resetRequired);

      if (!canLogin) {
        console.log('login: Account disabled');
        return {
          success: false,
          error: 'Your account has been disabled. Please contact support.',
          errorCode: 'ACCOUNT_DISABLED'
        };
      }

      if (!emailConfirmed) {
        console.log('login: Email not confirmed');
        return {
          success: false,
          error: 'Please confirm your email address before logging in.',
          errorCode: 'EMAIL_NOT_CONFIRMED',
          needsEmailConfirmation: true
        };
      }

      // Check password
      console.log('login: Verifying password...');
      const passwordCheck = verifyUserPassword(passwordStr, user.PasswordHash, { email: normalizedEmail });
      
      if (!passwordCheck.success) {
        console.log('login: Password verification failed:', passwordCheck.reason);
        
        if (passwordCheck.reason === 'NO_STORED_HASH') {
          return {
            success: false,
            error: 'Please set up your password using the link from your welcome email.',
            errorCode: 'PASSWORD_NOT_SET',
            needsPasswordSetup: true
          };
        }
        
        return {
          success: false,
          error: 'Invalid email or password',
          errorCode: 'INVALID_CREDENTIALS'
        };
      }

      console.log('login: Password verified successfully using method:', passwordCheck.method);

      const tenantAccess = resolveTenantAccess(user, null);
      if (!tenantAccess || !tenantAccess.success) {
        console.log('login: Tenant access check failed:', tenantAccess ? tenantAccess.reason : 'unknown');
        const tenantError = formatTenantAccessError(tenantAccess);
        return {
          success: false,
          error: tenantError.error,
          errorCode: tenantError.errorCode
        };
      }

      const tenantSummary = Object.assign({}, tenantAccess.clientPayload, {
        tenantContext: tenantAccess.sessionScope && tenantAccess.sessionScope.tenantContext
          ? tenantAccess.sessionScope.tenantContext
          : null
      });
      if (Array.isArray(tenantAccess.warnings)) {
        tenantSummary.warnings = tenantAccess.warnings.slice();
      }
      tenantSummary.needsCampaignAssignment = tenantAccess.needsCampaignAssignment === true;

      // Handle reset required
      if (resetRequired) {
        console.log('login: Password reset required');
        const resetSession = createSession(user.ID, false, tenantAccess.sessionScope);
        return {
          success: false,
          error: 'You must change your password before continuing.',
          errorCode: 'PASSWORD_RESET_REQUIRED',
          resetToken: resetSession && resetSession.token ? resetSession.token : null,
          needsPasswordReset: true,
          tenant: tenantSummary,
          campaignScope: tenantSummary,
          warnings: Array.isArray(tenantAccess.warnings) ? tenantAccess.warnings.slice() : [],
          needsCampaignAssignment: tenantAccess.needsCampaignAssignment === true
        };
      }

      const mfaConfig = getUserMfaConfig(user);
      if (mfaConfig && mfaConfig.enabled) {
        console.log('login: MFA required for user:', user.Email || user.UserName || user.ID);
        const challengeResult = createMfaChallenge(user, tenantAccess, rememberMe, sanitizedMetadata, mfaConfig);

        if (!challengeResult || !challengeResult.success) {
          console.warn('login: Failed to create MFA challenge. Reason:', challengeResult && challengeResult.reason);
          return {
            success: false,
            error: 'We were unable to start the verification process. Please try again in a moment.',
            errorCode: 'MFA_CHALLENGE_FAILED'
          };
        }

        const challenge = challengeResult.challenge;
        return {
          success: false,
          needsMfa: true,
          errorCode: 'MFA_REQUIRED',
          message: 'Additional verification is required to finish signing in.',
          rememberMe: !!rememberMe,
          mfa: {
            challengeId: challenge.id,
            deliveryMethod: challenge.deliveryMethod,
            maskedDestination: challenge.maskedDestination,
            totp: challenge.totpEnabled,
            expiresAt: new Date(challenge.expiresAt).toISOString(),
            deliveriesRemaining: Math.max(0, (challenge.maxDeliveries || MFA_MAX_DELIVERIES) - (challenge.deliveries || 0)),
            backupCodesRemaining: mfaConfig.backupCodes.length
          }
        };
      }

      const deviceEvaluation = evaluateTrustedDevice(user, sanitizedMetadata, rememberMe);
      if (deviceEvaluation && deviceEvaluation.error) {
        console.warn('login: Device evaluation error:', deviceEvaluation.errorCode || deviceEvaluation.error);
        return {
          success: false,
          error: deviceEvaluation.error,
          errorCode: deviceEvaluation.errorCode || 'DEVICE_VERIFICATION_ERROR'
        };
      }

      if (deviceEvaluation && deviceEvaluation.trusted === false) {
        console.log('login: Device verification required for user');
        return {
          success: false,
          needsVerification: true,
          errorCode: 'DEVICE_VERIFICATION_REQUIRED',
          message: (deviceEvaluation.verification && deviceEvaluation.verification.message)
            || 'We need to confirm this device before completing your login.',
          verification: deviceEvaluation.verification || null,
          rememberMe: !!rememberMe
        };
      }

      // Create session
      console.log('login: Creating session...');
      const sessionResult = createSession(user.ID, rememberMe, tenantAccess.sessionScope, sanitizedMetadata);

      if (!sessionResult || !sessionResult.token) {
        console.log('login: Failed to create session');
        return {
          success: false,
          error: 'Failed to create session. Please try again.',
          errorCode: 'SESSION_CREATION_FAILED'
        };
      }

      console.log('login: Session created successfully');

      // Update last login
      try {
        updateLastLogin(user.ID);
      } catch (lastLoginError) {
        console.warn('login: Failed to update last login:', lastLoginError);
        // Don't fail login for this
      }

      // Build user payload
      const userPayload = buildUserPayload(user, tenantAccess.clientPayload);

      if (userPayload && userPayload.CampaignScope) {
        userPayload.CampaignScope.tenantContext = tenantAccess.sessionScope && tenantAccess.sessionScope.tenantContext
          ? tenantAccess.sessionScope.tenantContext
          : null;
        if (tenantAccess.sessionScope && Array.isArray(tenantAccess.sessionScope.assignments) && !userPayload.CampaignScope.assignments.length) {
          userPayload.CampaignScope.assignments = tenantAccess.sessionScope.assignments.slice();
        }
        if (tenantAccess.sessionScope && Array.isArray(tenantAccess.sessionScope.permissions) && !userPayload.CampaignScope.permissions.length) {
          userPayload.CampaignScope.permissions = tenantAccess.sessionScope.permissions.slice();
        }
      }

      const sessionToken = sessionResult.token;

      const warnings = Array.isArray(tenantAccess.warnings) ? tenantAccess.warnings.slice() : [];
      const needsCampaignAssignment = tenantAccess.needsCampaignAssignment === true;

      const loginMessage = needsCampaignAssignment
        ? 'Login successful, but your account is not yet assigned to any campaigns. You may have limited access until an administrator completes the assignment.'
        : 'Login successful';

      const landing = resolveLandingDestination(user, {
        user: userPayload,
        userPayload: userPayload,
        rawUser: user,
        tenantAccess: tenantAccess,
        tenant: { clientPayload: tenantSummary, sessionScope: tenantAccess.sessionScope },
        sessionScope: tenantAccess.sessionScope
      });
      const redirectSlug = landing && landing.slug ? landing.slug : 'dashboard';
      const redirectUrl = landing && landing.redirectUrl
        ? landing.redirectUrl
        : buildLandingRedirectUrlFromSlug(redirectSlug);

      console.log('login: Login successful for user:', userPayload.FullName);
      console.log('=== AuthenticationService.login SUCCESS ===');

      const result = {
        success: true,
        sessionToken: sessionToken,
        user: userPayload,
        message: loginMessage,
        rememberMe: !!rememberMe,
        sessionExpiresAt: sessionResult.expiresAt,
        sessionTtlSeconds: sessionResult.ttlSeconds,
        sessionIdleTimeoutMinutes: sessionResult.idleTimeoutMinutes,
        tenant: tenantSummary,
        campaignScope: userPayload ? userPayload.CampaignScope : null,
        warnings: warnings,
        needsCampaignAssignment: needsCampaignAssignment,
        redirectSlug: redirectSlug,
        redirectUrl: redirectUrl
      };

      if (sanitizedMetadata && sanitizedMetadata.requestedReturnUrl) {
        result.requestedReturnUrl = sanitizedMetadata.requestedReturnUrl;
      }

      return result;

    } catch (error) {
      console.error('login: Unexpected error:', error);
      console.log('=== AuthenticationService.login ERROR ===');
      
      // Write error to logs if function available
      if (typeof writeError === 'function') {
        writeError('AuthenticationService.login', error);
      }
      
      return {
        success: false,
        error: 'An error occurred during login. Please try again.',
        errorCode: 'SYSTEM_ERROR'
      };
    }
  }

  // ─── Session validation ─────────────────────────────────────────────────────

  function createSessionFor(userId, campaignId, rememberMe = false, metadata) {
    try {
      const user = findUserById(userId);
      if (!user) {
        return {
          success: false,
          error: 'User not found',
          errorCode: 'USER_NOT_FOUND'
        };
      }

      const tenantAccess = resolveTenantAccess(user, campaignId);
      if (!tenantAccess || !tenantAccess.success) {
        const tenantError = formatTenantAccessError(tenantAccess);
        return Object.assign({ success: false }, tenantError);
      }

      const sessionResult = createSession(user.ID || userId, rememberMe, tenantAccess.sessionScope, metadata);
      if (!sessionResult || !sessionResult.token) {
        return {
          success: false,
          error: 'Failed to create session. Please try again.',
          errorCode: 'SESSION_CREATION_FAILED'
        };
      }

      const userPayload = buildUserPayload(user, tenantAccess.clientPayload);

      if (userPayload && userPayload.CampaignScope) {
        userPayload.CampaignScope.tenantContext = tenantAccess.sessionScope && tenantAccess.sessionScope.tenantContext
          ? tenantAccess.sessionScope.tenantContext
          : null;
        if (tenantAccess.sessionScope && Array.isArray(tenantAccess.sessionScope.assignments) && !userPayload.CampaignScope.assignments.length) {
          userPayload.CampaignScope.assignments = tenantAccess.sessionScope.assignments.slice();
        }
        if (tenantAccess.sessionScope && Array.isArray(tenantAccess.sessionScope.permissions) && !userPayload.CampaignScope.permissions.length) {
          userPayload.CampaignScope.permissions = tenantAccess.sessionScope.permissions.slice();
        }
      }

      const tenantSummary = Object.assign({}, tenantAccess.clientPayload, {
        tenantContext: tenantAccess.sessionScope && tenantAccess.sessionScope.tenantContext
          ? tenantAccess.sessionScope.tenantContext
          : null
      });
      if (Array.isArray(tenantAccess.warnings)) {
        tenantSummary.warnings = tenantAccess.warnings.slice();
      }
      tenantSummary.needsCampaignAssignment = tenantAccess.needsCampaignAssignment === true;

      const warnings = Array.isArray(tenantAccess.warnings) ? tenantAccess.warnings.slice() : [];
      const needsCampaignAssignment = tenantAccess.needsCampaignAssignment === true;

      const landing = resolveLandingDestination(user, {
        user: userPayload,
        userPayload: userPayload,
        rawUser: user,
        tenantAccess: tenantAccess,
        tenant: { clientPayload: tenantSummary, sessionScope: tenantAccess.sessionScope },
        sessionScope: tenantAccess.sessionScope
      });
      const redirectSlug = landing && landing.slug ? landing.slug : 'dashboard';
      const redirectUrl = landing && landing.redirectUrl
        ? landing.redirectUrl
        : buildLandingRedirectUrlFromSlug(redirectSlug);

      return {
        success: true,
        sessionToken: sessionResult.token,
        sessionExpiresAt: sessionResult.expiresAt,
        sessionTtlSeconds: sessionResult.ttlSeconds,
        sessionIdleTimeoutMinutes: sessionResult.idleTimeoutMinutes,
        user: userPayload,
        tenant: tenantSummary,
        campaignScope: userPayload ? userPayload.CampaignScope : null,
        warnings: warnings,
        needsCampaignAssignment: needsCampaignAssignment,
        redirectSlug: redirectSlug,
        redirectUrl: redirectUrl
      };

    } catch (error) {
      console.error('createSessionFor: Error creating session for user', userId, error);
      return {
        success: false,
        error: error.message || 'Failed to create session',
        errorCode: 'SESSION_CREATION_ERROR'
      };
    }
  }

  function getSessionUser(sessionToken) {
    try {
      if (!sessionToken) return null;

      const resolution = resolveSessionRecord(sessionToken, { touch: true });

      if (!resolution || resolution.status !== 'active' || !resolution.entry) {
        if (resolution && resolution.status === 'expired') {
          console.log('getSessionUser: Session expired (' + resolution.reason + ')');
        } else {
          console.log('getSessionUser: Session not found');
        }
        return null;
      }

      const context = buildSessionUserContext(resolution.entry, sessionToken, resolution);
      if (!context || !context.user) {
        console.log('getSessionUser: User not found for session');
        removeSessionEntry(resolution.entry);
        return null;
      }

      return context.user;

    } catch (error) {
      console.error('getSessionUser: Error:', error);
      return null;
    }
  }

  // ─── Helper functions ─────────────────────────────────────────────────────

  function updateLastLogin(userId) {
    try {
      // This is a simplified version - you may need to adapt based on your sheet structure
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName('Users');
      if (!sheet) return;

      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const idIndex = headers.indexOf('ID');
      const lastLoginIndex = headers.indexOf('LastLogin');

      if (idIndex === -1) return;

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][idIndex]) === String(userId)) {
          if (lastLoginIndex !== -1) {
            sheet.getRange(i + 1, lastLoginIndex + 1).setValue(new Date());
          }
          break;
        }
      }
    } catch (error) {
      console.warn('updateLastLogin: Failed to update last login:', error);
    }
  }

  function logout(sessionToken) {
    try {
      if (!sessionToken) {
        return { success: true, message: 'No session to logout' };
      }

      const entry = findSessionEntry(sessionToken);
      if (entry) {
        removeSessionEntry(entry);
      }

      return { success: true, message: 'Logged out successfully' };

    } catch (error) {
      console.error('logout: Error:', error);
      return { success: false, error: error.message };
    }
  }

  function keepAlive(sessionToken) {
    try {
      const resolution = resolveSessionRecord(sessionToken, { touch: true });

      if (!resolution || resolution.status !== 'active' || !resolution.entry) {
        const reason = resolution ? resolution.reason : 'NOT_FOUND';
        const message = reason === 'IDLE_TIMEOUT'
          ? 'Session expired after 30 minutes of inactivity'
          : 'Session expired or invalid';
        return {
          success: false,
          expired: true,
          message: message,
          reason: reason,
          errorCode: reason ? 'SESSION_' + reason : 'SESSION_NOT_FOUND'
        };
      }

      const context = buildSessionUserContext(resolution.entry, sessionToken, resolution);
      if (!context || !context.user) {
        removeSessionEntry(resolution.entry);
        return {
          success: false,
          expired: true,
          message: 'Session expired or invalid',
          reason: 'USER_NOT_FOUND',
          errorCode: 'SESSION_USER_NOT_FOUND'
        };
      }

      const user = context.user;
      let ttlSeconds = null;
      if (user.sessionExpiresAt) {
        const expiryTime = Date.parse(user.sessionExpiresAt);
        if (!isNaN(expiryTime)) {
          ttlSeconds = Math.max(0, Math.floor((expiryTime - Date.now()) / 1000));
        }
      }

      const landing = resolveLandingDestination(context.rawUser || context.user, {
        user: context.user,
        userPayload: context.user,
        rawUser: context.rawUser,
        tenantAccess: { sessionScope: context.rawScope, clientPayload: context.tenant },
        tenant: { sessionScope: context.rawScope, clientPayload: context.tenant },
        sessionScope: context.rawScope
      });
      const redirectSlug = landing && landing.slug ? landing.slug : 'dashboard';
      const redirectUrl = landing && landing.redirectUrl
        ? landing.redirectUrl
        : buildLandingRedirectUrlFromSlug(redirectSlug);

      return {
        success: true,
        message: 'Session active',
        user: user,
        sessionToken: user.sessionToken,
        sessionExpiresAt: user.sessionExpiresAt || user.sessionExpiry || null,
        sessionTtlSeconds: ttlSeconds,
        tenant: user.CampaignScope || null,
        campaignScope: user.CampaignScope || null,
        warnings: user.CampaignScope && Array.isArray(user.CampaignScope.warnings) ? user.CampaignScope.warnings.slice() : [],
        needsCampaignAssignment: user.CampaignScope ? !!user.CampaignScope.needsCampaignAssignment : false,
        idleTimeoutMinutes: resolution.idleTimeoutMinutes,
        lastActivityAt: user.sessionLastActivityAt || null,
        redirectSlug: redirectSlug,
        redirectUrl: redirectUrl
      };
    } catch (error) {
      console.error('keepAlive: Error:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  // ─── Public API ─────────────────────────────────────────────────────────────

  return {
    login: login,
    logout: logout,
    createSessionFor: createSessionFor,
    getSessionUser: getSessionUser,
    keepAlive: keepAlive,
    findUserByEmail: findUserByEmail,
    findUserById: findUserById,
    verifyUserPassword: verifyUserPassword,
    getUserByEmail: findUserByEmail,
    findUserByPrincipal: findUserByEmail,
    beginMfaChallenge: beginMfaChallenge,
    verifyMfaCode: verifyMfaCode,
    resolveLandingDestination: resolveLandingDestination,
    getLandingSlug: determineLandingSlug,
    buildLandingRedirectUrl: buildLandingRedirectUrlFromSlug,
    captureLoginRequestContext: captureLoginRequestContext,
    consumeLoginRequestContext: consumeLoginContext,
    deriveLoginReturnUrlFromEvent: deriveLoginReturnUrlFromEvent,
    findActiveSessionForUser: findActiveSessionForUser,
    userHasActiveSession: userHasActiveSession,
    confirmDeviceVerification: confirmDeviceVerification,
    denyDeviceVerification: denyDeviceVerification,
    cleanupExpiredSessions: cleanupExpiredSessions
  };

})();

// ───────────────────────────────────────────────────────────────────────────────
// CLIENT-ACCESSIBLE FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

function loginUser(email, password, rememberMe = false, clientMetadata) {
  try {
    console.log('=== loginUser wrapper START ===');
    let mergedMetadata = null;
    try {
      if (clientMetadata && typeof clientMetadata === 'object') {
        mergedMetadata = Object.assign({}, clientMetadata);
      }

      if (typeof AuthenticationService !== 'undefined'
        && AuthenticationService
        && typeof AuthenticationService.consumeLoginRequestContext === 'function') {
        const serverContext = AuthenticationService.consumeLoginRequestContext();
        if (serverContext && typeof serverContext === 'object') {
          mergedMetadata = mergedMetadata || {};
          if (serverContext.serverIp) {
            mergedMetadata.serverIp = serverContext.serverIp;
            mergedMetadata.serverObservedIp = serverContext.serverIp;
          }
          if (serverContext.forwardedFor) {
            mergedMetadata.forwardedFor = serverContext.forwardedFor;
          }
          if (serverContext.serverUserAgent && !mergedMetadata.serverUserAgent) {
            mergedMetadata.serverUserAgent = serverContext.serverUserAgent;
          }
          if (serverContext.host && !mergedMetadata.host) {
            mergedMetadata.host = serverContext.host;
          }
          mergedMetadata.serverObservedAt = serverContext.serverObservedAt || new Date().toISOString();
        }
      }
    } catch (metadataMergeError) {
      console.warn('loginUser: Failed to merge server metadata', metadataMergeError);
    }

    const result = AuthenticationService.login(email, password, rememberMe, mergedMetadata || clientMetadata);
    console.log('=== loginUser wrapper END ===');
    return result;
  } catch (error) {
    console.error('loginUser wrapper error:', error);
    return {
      success: false,
      error: 'Login failed. Please try again.',
      errorCode: 'WRAPPER_ERROR'
    };
  }
}

function confirmDeviceVerification(verificationId, code, clientMetadata) {
  try {
    return AuthenticationService.confirmDeviceVerification(verificationId, code, clientMetadata);
  } catch (error) {
    console.error('confirmDeviceVerification wrapper error:', error);
    return {
      success: false,
      error: 'We were unable to confirm the device. Please try again.',
      errorCode: 'DEVICE_CONFIRM_ERROR'
    };
  }
}

function denyDeviceVerification(verificationId, clientMetadata) {
  try {
    return AuthenticationService.denyDeviceVerification(verificationId, clientMetadata);
  } catch (error) {
    console.error('denyDeviceVerification wrapper error:', error);
    return {
      success: false,
      error: 'We were unable to record your response. Please contact support if this persists.',
      errorCode: 'DEVICE_DENY_ERROR'
    };
  }
}

function beginMfaChallenge(challengeId, options) {
  try {
    return AuthenticationService.beginMfaChallenge(challengeId, options);
  } catch (error) {
    console.error('beginMfaChallenge wrapper error:', error);
    return {
      success: false,
      error: error.message || 'Unable to send verification code.',
      errorCode: 'MFA_DELIVERY_ERROR'
    };
  }
}

function verifyMfaCode(challengeId, code, clientMetadata) {
  try {
    return AuthenticationService.verifyMfaCode(challengeId, code, clientMetadata);
  } catch (error) {
    console.error('verifyMfaCode wrapper error:', error);
    return {
      success: false,
      error: error.message || 'Unable to verify the authentication code.',
      errorCode: 'MFA_VERIFY_ERROR'
    };
  }
}

function logoutUser(sessionToken) {
  try {
    return AuthenticationService.logout(sessionToken);
  } catch (error) {
    console.error('logoutUser wrapper error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function keepAliveSession(sessionToken) {
  try {
    return AuthenticationService.keepAlive(sessionToken);
  } catch (error) {
    console.error('keepAliveSession wrapper error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

// Legacy compatibility
function login(email, password) {
  return AuthenticationService.login(email, password);
}

function logout(sessionToken) {
  return AuthenticationService.logout(sessionToken);
}

function keepAlive(sessionToken) {
  return AuthenticationService.keepAlive(sessionToken);
}

function cleanupExpiredSessionsJob() {
  try {
    if (typeof AuthenticationService !== 'undefined'
      && AuthenticationService
      && typeof AuthenticationService.cleanupExpiredSessions === 'function') {
      const result = AuthenticationService.cleanupExpiredSessions();
      if (result && result.success) {
        const removed = typeof result.removed === 'number' ? result.removed : 0;
        const evaluated = typeof result.evaluated === 'number' ? result.evaluated : 0;
        const reasons = result.reasons ? JSON.stringify(result.reasons) : '{}';
        console.log('cleanupExpiredSessionsJob: removed ' + removed + ' of ' + evaluated + ' sessions (reasons: ' + reasons + ').');
      } else {
        console.warn('cleanupExpiredSessionsJob: cleanup did not succeed', result);
      }
      return result;
    }
    console.warn('cleanupExpiredSessionsJob: AuthenticationService not available');
    return { success: false, error: 'AuthenticationService unavailable' };
  } catch (error) {
    console.error('cleanupExpiredSessionsJob error:', error);
    return { success: false, error: error.message };
  }
}

console.log('Fixed AuthenticationService.gs loaded successfully');
console.log('Key improvements:');
console.log('- Consistent email normalization');
console.log('- Robust password verification with multiple fallback methods');
console.log('- Better error logging and debugging');
console.log('- Improved user lookup with fallbacks');
console.log('- Enhanced session management');


/**
 * Authentication Diagnostic Functions for Lumina
 * Use these functions to identify authentication issues
 */

function debugAuthenticationIssues(email, password) {
  try {
    const results = {
      timestamp: new Date().toISOString(),
      email: email,
      userLookup: null,
      passwordCheck: null,
      systemStatus: null,
      recommendations: []
    };

    // 1. Test User Lookup
    console.log('1. Testing user lookup for:', email);
    
    // Try different lookup methods
    const normalizedEmail = String(email || '').trim().toLowerCase();
    
    // Method 1: AuthenticationService lookup
    let userByAuth = null;
    if (typeof AuthenticationService !== 'undefined' && AuthenticationService.getUserByEmail) {
      try {
        userByAuth = AuthenticationService.getUserByEmail(normalizedEmail);
        console.log('AuthenticationService lookup result:', userByAuth ? 'Found' : 'Not found');
      } catch (e) {
        console.error('AuthenticationService lookup failed:', e);
      }
    }
    
    // Method 2: Direct sheet lookup
    let userBySheet = null;
    try {
      const users = readSheet('Users') || [];
      userBySheet = users.find(u => 
        String(u.Email || '').trim().toLowerCase() === normalizedEmail
      );
      console.log('Direct sheet lookup result:', userBySheet ? 'Found' : 'Not found');
    } catch (e) {
      console.error('Direct sheet lookup failed:', e);
    }
    
    // Method 3: findUserByPrincipal lookup
    let userByPrincipal = null;
    if (typeof AuthenticationService !== 'undefined' && AuthenticationService.findUserByPrincipal) {
      try {
        userByPrincipal = AuthenticationService.findUserByPrincipal(normalizedEmail);
        console.log('findUserByPrincipal lookup result:', userByPrincipal ? 'Found' : 'Not found');
      } catch (e) {
        console.error('findUserByPrincipal lookup failed:', e);
      }
    }

    results.userLookup = {
      normalizedEmail: normalizedEmail,
      authServiceResult: userByAuth ? 'Found' : 'Not found',
      directSheetResult: userBySheet ? 'Found' : 'Not found',
      principalResult: userByPrincipal ? 'Found' : 'Not found',
      consistencyCheck: (!!userByAuth === !!userBySheet && !!userBySheet === !!userByPrincipal)
    };

    // Use the first successful lookup for further testing
    const user = userByAuth || userBySheet || userByPrincipal;
    
    if (!user) {
      results.recommendations.push('USER_NOT_FOUND: Check if user exists in Users sheet with correct email');
      results.recommendations.push('Check email case sensitivity and whitespace');
      return results;
    }

    console.log('2. Found user:', user.FullName || user.UserName || user.Email);

    // 2. Test Password Verification
    console.log('3. Testing password verification...');
    
    const storedHash = user.PasswordHash || '';
    const hasPassword = storedHash && storedHash.trim() !== '';
    
    results.passwordCheck = {
      hasStoredHash: hasPassword,
      storedHashLength: storedHash.length,
      storedHashSample: storedHash ? storedHash.substring(0, 10) + '...' : 'empty',
      canLogin: user.CanLogin,
      emailConfirmed: user.EmailConfirmed,
      resetRequired: user.ResetRequired
    };

    if (!hasPassword) {
      results.recommendations.push('PASSWORD_NOT_SET: User has no password hash - needs password setup');
      results.recommendations.push('Check if user completed initial password setup process');
    } else {
      // Test password verification methods
      let verificationResults = {};
      
      // Method 1: PasswordUtilities.verifyPassword
      if (typeof PasswordUtilities !== 'undefined') {
        try {
          const isValid1 = PasswordUtilities.verifyPassword(password, storedHash);
          verificationResults.passwordUtilsResult = isValid1;
          console.log('PasswordUtilities.verifyPassword result:', isValid1);
        } catch (e) {
          verificationResults.passwordUtilsError = e.message;
        }
      }

      // Method 2: Test hash generation
      if (typeof PasswordUtilities !== 'undefined') {
        try {
          const newHash = PasswordUtilities.hashPassword(password);
          verificationResults.newHashMatches = (newHash === storedHash);
          verificationResults.newHashSample = newHash.substring(0, 10) + '...';
          console.log('Generated hash matches stored:', newHash === storedHash);
        } catch (e) {
          verificationResults.hashGenError = e.message;
        }
      }

      // Method 3: Test normalized hash
      if (typeof PasswordUtilities !== 'undefined') {
        try {
          const normalizedStored = PasswordUtilities.normalizeHash(storedHash);
          const newHash = PasswordUtilities.hashPassword(password);
          verificationResults.normalizedComparison = (newHash === normalizedStored);
          console.log('Normalized hash comparison:', newHash === normalizedStored);
        } catch (e) {
          verificationResults.normalizeError = e.message;
        }
      }

      results.passwordCheck.verificationTests = verificationResults;

      if (!verificationResults.passwordUtilsResult && !verificationResults.newHashMatches) {
        results.recommendations.push('PASSWORD_MISMATCH: Password verification failed - check password or hash corruption');
      }
    }

    // 3. Account Status Checks
    console.log('4. Checking account status...');
    
    const accountStatus = {
      canLogin: String(user.CanLogin).toUpperCase() === 'TRUE',
      emailConfirmed: String(user.EmailConfirmed).toUpperCase() === 'TRUE',
      resetRequired: String(user.ResetRequired).toUpperCase() === 'TRUE',
      isAdmin: String(user.IsAdmin).toUpperCase() === 'TRUE',
      campaignId: user.CampaignID || '',
      lockoutEnd: user.LockoutEnd || null
    };

    results.systemStatus = accountStatus;

    if (!accountStatus.canLogin) {
      results.recommendations.push('ACCOUNT_DISABLED: User CanLogin is FALSE');
    }
    if (!accountStatus.emailConfirmed) {
      results.recommendations.push('EMAIL_NOT_CONFIRMED: User EmailConfirmed is FALSE');
    }
    if (accountStatus.resetRequired) {
      results.recommendations.push('RESET_REQUIRED: User ResetRequired is TRUE');
    }

    // 4. System-wide checks
    console.log('5. Running system-wide checks...');
    
    const systemChecks = {
      authServiceAvailable: typeof AuthenticationService !== 'undefined',
      passwordUtilsAvailable: typeof PasswordUtilities !== 'undefined',
      usersSheetExists: false,
      usersSheetRowCount: 0
    };

    try {
      const users = readSheet('Users') || [];
      systemChecks.usersSheetExists = true;
      systemChecks.usersSheetRowCount = users.length;
    } catch (e) {
      results.recommendations.push('SHEET_ACCESS_ERROR: Cannot read Users sheet');
    }

    results.systemStatus.systemChecks = systemChecks;

    // 5. Generate specific recommendations
    if (results.recommendations.length === 0) {
      results.recommendations.push('ALL_CHECKS_PASSED: Authentication should work - investigate client-side issues');
    }

    return results;

  } catch (error) {
    console.error('Error in debugAuthenticationIssues:', error);
    return {
      error: error.message,
      stack: error.stack,
      recommendations: ['DIAGNOSTIC_ERROR: Cannot complete authentication diagnosis']
    };
  }
}

function testPasswordHashing(plainPassword) {
  try {
    console.log('Testing password hashing for password:', plainPassword ? 'PROVIDED' : 'EMPTY');
    
    const results = {
      timestamp: new Date().toISOString(),
      tests: {}
    };

    if (typeof PasswordUtilities !== 'undefined') {
      // Test 1: Basic hashing
      const hash1 = PasswordUtilities.hashPassword(plainPassword);
      const hash2 = PasswordUtilities.hashPassword(plainPassword);
      
      results.tests.basicHashing = {
        hash1: hash1,
        hash2: hash2,
        consistent: hash1 === hash2,
        length: hash1.length
      };

      // Test 2: Verification
      const verifies = PasswordUtilities.verifyPassword(plainPassword, hash1);
      results.tests.verification = {
        verifies: verifies
      };

      // Test 3: Normalization
      const normalized = PasswordUtilities.normalizeHash(hash1);
      results.tests.normalization = {
        original: hash1,
        normalized: normalized,
        same: hash1 === normalized
      };

      // Test 4: Edge cases
      results.tests.edgeCases = {
        emptyPassword: PasswordUtilities.hashPassword(''),
        spacePassword: PasswordUtilities.hashPassword(' '),
        nullPassword: PasswordUtilities.hashPassword(null)
      };

    } else {
      results.error = 'PasswordUtilities not available';
    }

    return results;

  } catch (error) {
    console.error('Error in testPasswordHashing:', error);
    return {
      error: error.message,
      stack: error.stack
    };
  }
}

function fixAuthenticationIssues(email, options = {}) {
  try {
    const {
      resetPassword = false,
      enableLogin = false,
      confirmEmail = false,
      generateNewHash = false,
      newPassword = null
    } = options;

    const results = {
      timestamp: new Date().toISOString(),
      email: email,
      actions: [],
      errors: []
    };

    // 1. Find user
    const normalizedEmail = String(email || '').trim().toLowerCase();
    const users = readSheet('Users') || [];
    const userIndex = users.findIndex(u => 
      String(u.Email || '').trim().toLowerCase() === normalizedEmail
    );

    if (userIndex === -1) {
      results.errors.push('User not found');
      return results;
    }

    const user = users[userIndex];
    results.userId = user.ID;

    // 2. Get sheet reference for updates
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const getColumnIndex = (columnName) => {
      return headers.indexOf(columnName);
    };

    const rowNumber = userIndex + 2; // +1 for 0-based index, +1 for header row

    // 3. Apply fixes
    if (enableLogin) {
      const canLoginCol = getColumnIndex('CanLogin');
      if (canLoginCol !== -1) {
        sheet.getRange(rowNumber, canLoginCol + 1).setValue('TRUE');
        results.actions.push('Set CanLogin to TRUE');
      }
    }

    if (confirmEmail) {
      const emailConfirmedCol = getColumnIndex('EmailConfirmed');
      if (emailConfirmedCol !== -1) {
        sheet.getRange(rowNumber, emailConfirmedCol + 1).setValue('TRUE');
        results.actions.push('Set EmailConfirmed to TRUE');
      }
    }

    if (resetPassword) {
      const resetRequiredCol = getColumnIndex('ResetRequired');
      if (resetRequiredCol !== -1) {
        sheet.getRange(rowNumber, resetRequiredCol + 1).setValue('FALSE');
        results.actions.push('Set ResetRequired to FALSE');
      }
    }

    if (generateNewHash && newPassword && typeof PasswordUtilities !== 'undefined') {
      const passwordHashCol = getColumnIndex('PasswordHash');
      if (passwordHashCol !== -1) {
        const newHash = PasswordUtilities.hashPassword(newPassword);
        sheet.getRange(rowNumber, passwordHashCol + 1).setValue(newHash);
        results.actions.push('Generated new password hash');
        results.newHashSample = newHash.substring(0, 10) + '...';
      }
    }

    // 4. Update timestamp
    const updatedAtCol = getColumnIndex('UpdatedAt');
    if (updatedAtCol !== -1) {
      sheet.getRange(rowNumber, updatedAtCol + 1).setValue(new Date());
      results.actions.push('Updated timestamp');
    }

    // 5. Clear cache
    if (typeof invalidateCache === 'function') {
      invalidateCache('Users');
      results.actions.push('Cleared Users cache');
    }

    return results;

  } catch (error) {
    console.error('Error in fixAuthenticationIssues:', error);
    return {
      error: error.message,
      stack: error.stack
    };
  }
}

// Helper function to check all users for common authentication issues
function scanAllUsersForAuthIssues() {
  try {
    const users = readSheet('Users') || [];
    const issues = {
      noPasswordHash: [],
      cannotLogin: [],
      emailNotConfirmed: [],
      resetRequired: [],
      emptyEmail: [],
      duplicateEmails: [],
      totalUsers: users.length
    };

    const emailCounts = {};

    users.forEach((user, index) => {
      const email = String(user.Email || '').trim().toLowerCase();
      
      // Track email duplicates
      if (email) {
        emailCounts[email] = (emailCounts[email] || 0) + 1;
      } else {
        issues.emptyEmail.push({
          index: index + 2, // Sheet row number
          id: user.ID,
          userName: user.UserName,
          fullName: user.FullName
        });
      }

      // Check password hash
      if (!user.PasswordHash || String(user.PasswordHash).trim() === '') {
        issues.noPasswordHash.push({
          index: index + 2,
          id: user.ID,
          email: user.Email,
          userName: user.UserName,
          fullName: user.FullName
        });
      }

      // Check login capability
      if (String(user.CanLogin).toUpperCase() !== 'TRUE') {
        issues.cannotLogin.push({
          index: index + 2,
          id: user.ID,
          email: user.Email,
          userName: user.UserName,
          fullName: user.FullName,
          canLogin: user.CanLogin
        });
      }

      // Check email confirmation
      if (String(user.EmailConfirmed).toUpperCase() !== 'TRUE') {
        issues.emailNotConfirmed.push({
          index: index + 2,
          id: user.ID,
          email: user.Email,
          userName: user.UserName,
          fullName: user.FullName,
          emailConfirmed: user.EmailConfirmed
        });
      }

      // Check reset required
      if (String(user.ResetRequired).toUpperCase() === 'TRUE') {
        issues.resetRequired.push({
          index: index + 2,
          id: user.ID,
          email: user.Email,
          userName: user.UserName,
          fullName: user.FullName
        });
      }
    });

    // Find duplicate emails
    Object.entries(emailCounts).forEach(([email, count]) => {
      if (count > 1) {
        const duplicates = users
          .map((user, index) => ({ user, index: index + 2 }))
          .filter(item => String(item.user.Email || '').trim().toLowerCase() === email)
          .map(item => ({
            index: item.index,
            id: item.user.ID,
            userName: item.user.UserName,
            fullName: item.user.FullName
          }));
        
        issues.duplicateEmails.push({
          email: email,
          count: count,
          users: duplicates
        });
      }
    });

    return issues;

  } catch (error) {
    console.error('Error in scanAllUsersForAuthIssues:', error);
    return {
      error: error.message,
      stack: error.stack
    };
  }
}

// Client-accessible wrapper functions
function clientDebugAuth(email, password) {
  try {
    return debugAuthenticationIssues(email, password);
  } catch (error) {
    return { error: error.message };
  }
}

function clientTestPasswordHashing(password) {
  try {
    return testPasswordHashing(password);
  } catch (error) {
    return { error: error.message };
  }
}

function clientFixAuthIssues(email, options) {
  try {
    return fixAuthenticationIssues(email, options);
  } catch (error) {
    return { error: error.message };
  }
}

function clientScanAuthIssues() {
  try {
    return scanAllUsersForAuthIssues();
  } catch (error) {
    return { error: error.message };
  }
}

console.log('Authentication diagnostic functions loaded');
console.log('Available functions:');
console.log('- debugAuthenticationIssues(email, password)');
console.log('- testPasswordHashing(password)');
console.log('- fixAuthenticationIssues(email, options)');
console.log('- scanAllUsersForAuthIssues()');
console.log('- clientDebugAuth(email, password) - for google.script.run');
console.log('- clientTestPasswordHashing(password) - for google.script.run');
console.log('- clientFixAuthIssues(email, options) - for google.script.run');
console.log('- clientScanAuthIssues() - for google.script.run');
