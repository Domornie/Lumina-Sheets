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

const SESSION_TTL_MS = 60 * 60 * 1000; // 1 hour
const REMEMBER_ME_TTL_MS = 24 * 60 * 60 * 1000; // 24 hours
const SESSION_COLUMNS = (typeof SESSIONS_HEADERS !== 'undefined' && Array.isArray(SESSIONS_HEADERS) && SESSIONS_HEADERS.length)
  ? SESSIONS_HEADERS.slice()
  : ['Token', 'UserId', 'CreatedAt', 'ExpiresAt', 'RememberMe', 'CampaignScope', 'UserAgent', 'IpAddress'];

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
      const expiresAt = new Date(now.getTime() + ttl);

      const scopeData = campaignScope && typeof campaignScope === 'object'
        ? campaignScope
        : null;

      const sessionRecord = {
        Token: token,
        UserId: userId,
        CreatedAt: now.toISOString(),
        ExpiresAt: expiresAt.toISOString(),
        RememberMe: rememberMe ? 'TRUE' : 'FALSE',
        CampaignScope: serializeCampaignScope(scopeData),
        UserAgent: metadata && metadata.userAgent ? metadata.userAgent : 'Google Apps Script',
        IpAddress: metadata && metadata.ipAddress ? metadata.ipAddress : 'N/A'
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
        campaignScope: scopeData
      };

    } catch (error) {
      console.error('createSession: Error creating session:', error);
      return null;
    }
  }

  // ─── Main login function ─────────────────────────────────────────────────────

  function login(email, password, rememberMe = false) {
    console.log('=== AuthenticationService.login START ===');
    console.log('Email:', email ? 'PROVIDED' : 'EMPTY');
    console.log('Password:', password ? 'PROVIDED' : 'EMPTY');
    console.log('RememberMe:', rememberMe);

    try {
      // Input validation
      const normalizedEmail = normalizeEmail(email);
      const passwordStr = normalizeString(password);

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

      // Create session
      console.log('login: Creating session...');
      const sessionResult = createSession(user.ID, rememberMe, tenantAccess.sessionScope);

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

      console.log('login: Login successful for user:', userPayload.FullName);
      console.log('=== AuthenticationService.login SUCCESS ===');

      return {
        success: true,
        sessionToken: sessionToken,
        user: userPayload,
        message: loginMessage,
        rememberMe: !!rememberMe,
        sessionExpiresAt: sessionResult.expiresAt,
        sessionTtlSeconds: sessionResult.ttlSeconds,
        tenant: tenantSummary,
        campaignScope: userPayload ? userPayload.CampaignScope : null,
        warnings: warnings,
        needsCampaignAssignment: needsCampaignAssignment
      };

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

      return {
        success: true,
        sessionToken: sessionResult.token,
        sessionExpiresAt: sessionResult.expiresAt,
        sessionTtlSeconds: sessionResult.ttlSeconds,
        user: userPayload,
        tenant: tenantSummary,
        campaignScope: userPayload ? userPayload.CampaignScope : null,
        warnings: warnings,
        needsCampaignAssignment: needsCampaignAssignment
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

      // Find session
      let sessions = [];
      try {
        sessions = readSheet('Sessions') || [];
      } catch (error) {
        console.warn('getSessionUser: Failed to read sessions:', error);
        return null;
      }

      const session = sessions.find(s => 
        normalizeString(s.Token) === normalizeString(sessionToken)
      );

      if (!session) {
        console.log('getSessionUser: Session not found');
        return null;
      }

      // Check expiry
      const expiryTime = new Date(session.ExpiresAt).getTime();
      const now = Date.now();

      if (!expiryTime || isNaN(expiryTime) || expiryTime < now) {
        console.log('getSessionUser: Session expired');
        return null;
      }

      // Get user
      const user = findUserById(session.UserId) || findUserByEmail(session.UserId);
      if (!user) {
        console.log('getSessionUser: User not found for session');
        return null;
      }

      const rawScope = parseCampaignScope(session.CampaignScope || session.campaignScope);
      const tenantPayload = buildTenantScopePayload(rawScope);
      const userPayload = buildUserPayload(user, tenantPayload);

      if (userPayload && userPayload.CampaignScope) {
        userPayload.CampaignScope.tenantContext = rawScope && rawScope.tenantContext ? rawScope.tenantContext : null;
      }

      if (rawScope && Array.isArray(rawScope.assignments) && userPayload && userPayload.CampaignScope) {
        userPayload.CampaignScope.assignments = rawScope.assignments.slice();
      }

      if (rawScope && Array.isArray(rawScope.permissions) && userPayload && userPayload.CampaignScope) {
        userPayload.CampaignScope.permissions = rawScope.permissions.slice();
      }

      if (userPayload) {
        userPayload.sessionToken = sessionToken;
        userPayload.sessionExpiry = session.ExpiresAt;
        userPayload.sessionExpiresAt = session.ExpiresAt;
        userPayload.sessionScope = rawScope || null;
        userPayload.NeedsCampaignAssignment = userPayload.CampaignScope ? !!userPayload.CampaignScope.needsCampaignAssignment : false;
      }

      return userPayload;

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

      // Remove session from sheet
      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('Sessions');
        if (sheet) {
          const data = sheet.getDataRange().getValues();
          const headers = data[0];
          const tokenIndex = headers.indexOf('Token');

          if (tokenIndex !== -1) {
            for (let i = data.length - 1; i >= 1; i--) {
              if (normalizeString(data[i][tokenIndex]) === normalizeString(sessionToken)) {
                sheet.deleteRow(i + 1);
                break;
              }
            }
          }
        }
      } catch (error) {
        console.warn('logout: Failed to remove session from sheet:', error);
      }

      return { success: true, message: 'Logged out successfully' };

    } catch (error) {
      console.error('logout: Error:', error);
      return { success: false, error: error.message };
    }
  }

  function keepAlive(sessionToken) {
    try {
      const user = getSessionUser(sessionToken);
      
      if (!user) {
        return {
          success: false,
          expired: true,
          message: 'Session expired or invalid'
        };
      }

      let ttlSeconds = null;
      if (user.sessionExpiresAt) {
        const expiryTime = Date.parse(user.sessionExpiresAt);
        if (!isNaN(expiryTime)) {
          ttlSeconds = Math.max(0, Math.floor((expiryTime - Date.now()) / 1000));
        }
      }

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
        needsCampaignAssignment: user.CampaignScope ? !!user.CampaignScope.needsCampaignAssignment : false
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
    findUserByPrincipal: findUserByEmail
  };

})();

// ───────────────────────────────────────────────────────────────────────────────
// CLIENT-ACCESSIBLE FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

function loginUser(email, password, rememberMe = false) {
  try {
    console.log('=== loginUser wrapper START ===');
    const result = AuthenticationService.login(email, password, rememberMe);
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
