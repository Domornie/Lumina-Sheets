/**
 * LuminaIdentityService.js
 *
 * Centralized identity orchestration for Lumina HQ.
 * Provides canonical identity headers, sheet synchronization,
 * authentication evaluation, authorization context, and structured logging.
 */

var IdentityService = (function attachLuminaIdentity(existing) {
  var exports = (existing && typeof existing === 'object') ? existing : {};

  var IDENTITY_SYSTEM_NAME = 'Lumina Identity';
  var IDENTITY_VERSION = 'lumina-identity/v1';

  function getCanonicalIdentityHeaders_() {
    if (typeof LUMINA_IDENTITY_HEADERS !== 'undefined' && Array.isArray(LUMINA_IDENTITY_HEADERS) && LUMINA_IDENTITY_HEADERS.length) {
      return LUMINA_IDENTITY_HEADERS.slice();
    }
    if (typeof getCanonicalUserHeaders === 'function') {
      try {
        var fallback = getCanonicalUserHeaders({ preferIdentityService: false });
        if (Array.isArray(fallback) && fallback.length) return fallback.slice();
      } catch (headerErr) {
        console.warn('LuminaIdentityService.getCanonicalIdentityHeaders_: fallback failed', headerErr);
      }
    }
    return [
      'ID', 'UserName', 'Email', 'FullName', 'FirstName', 'LastName', 'Roles', 'CampaignID', 'EmploymentStatus',
      'CanLogin', 'IsAdmin', 'TwoFactorEnabled', 'CreatedAt', 'UpdatedAt', 'LastLoginAt'
    ];
  }

  function ensureIdentityHeaders(options) {
    try {
      if (typeof synchronizeLuminaIdentityHeaders === 'function') {
        return synchronizeLuminaIdentityHeaders(options || {});
      }

      if (typeof ensureSheetWithHeaders === 'function') {
        var canonical = Array.isArray(USERS_HEADERS) && USERS_HEADERS.length ? USERS_HEADERS : getCanonicalIdentityHeaders_();
        var sheetName = (options && options.sheetName) || LUMINA_IDENTITY_SHEET || USERS_SHEET;
        if (sheetName) {
          return ensureSheetWithHeaders(sheetName, canonical);
        }
      }
    } catch (syncErr) {
      console.warn('LuminaIdentityService.ensureIdentityHeaders: failed', syncErr);
    }
    return null;
  }

  function ensureInfrastructure(options) {
    var summary = {
      ensured: [],
      errors: []
    };

    try {
      if (typeof ensureIdentitySheetStructures === 'function') {
        var res = ensureIdentitySheetStructures(options || {});
        if (Array.isArray(res)) summary.ensured = summary.ensured.concat(res);
      }
    } catch (infraErr) {
      summary.errors.push(infraErr);
      safeWriteError && safeWriteError('LuminaIdentityService.ensureInfrastructure', infraErr);
    }

    try {
      if (typeof ensureUserInsuranceSheet_ === 'function') {
        ensureUserInsuranceSheet_();
      }
    } catch (insuranceErr) {
      summary.errors.push(insuranceErr);
      safeWriteError && safeWriteError('LuminaIdentityService.ensureInfrastructure.insurance', insuranceErr);
    }

    try { ensureIdentityHeaders(options); } catch (hdrErr) {
      summary.errors.push(hdrErr);
      safeWriteError && safeWriteError('LuminaIdentityService.ensureInfrastructure.headers', hdrErr);
    }

    return summary;
  }

  function safeReadSheet_(name, opts) {
    if (!name) return [];
    try {
      if (typeof readSheet === 'function') {
        var data = readSheet(name, opts || {});
        return Array.isArray(data) ? data : [];
      }
    } catch (readErr) {
      safeWriteError && safeWriteError('LuminaIdentityService.readSheet(' + name + ')', readErr);
    }
    return [];
  }

  function normalizeString_(value) {
    if (value === null || typeof value === 'undefined') return '';
    return String(value).trim();
  }

  function normalizeId_(value) {
    var str = normalizeString_(value);
    return str || '';
  }

  function normalizeEmail_(value) {
    var str = normalizeString_(value);
    return str ? str.toLowerCase() : '';
  }

  function toBoolean_(value) {
    if (value === true || value === false) return value;
    if (typeof value === 'number') return value !== 0;
    if (typeof value === 'string') {
      var normalized = value.trim().toLowerCase();
      if (!normalized) return false;
      return ['true', 'yes', '1', 'y', 't'].indexOf(normalized) !== -1;
    }
    return false;
  }

  function parseDate_(value) {
    if (!value && value !== 0) return null;
    if (value instanceof Date) return new Date(value.getTime());
    var num = Date.parse(value);
    return isNaN(num) ? null : new Date(num);
  }

  function safeDateIso_(value) {
    var date = parseDate_(value);
    return date ? date.toISOString() : '';
  }

  function computeNameParts_(fields) {
    var first = normalizeString_(fields.FirstName);
    var middle = normalizeString_(fields.MiddleName);
    var last = normalizeString_(fields.LastName);
    var full = normalizeString_(fields.FullName);

    if (!first || !last) {
      var source = full || normalizeString_(fields.DisplayName) || normalizeString_(fields.UserName) || '';
      if (source) {
        var segments = source.split(/\s+/).filter(Boolean);
        if (!first && segments.length) first = segments[0];
        if (!last && segments.length > 1) last = segments[segments.length - 1];
        if (!middle && segments.length > 2) middle = segments.slice(1, segments.length - 1).join(' ');
      }
    }

    if (!full) {
      full = [first, middle, last].filter(Boolean).join(' ').trim();
    }

    var display = normalizeString_(fields.DisplayName) || full || normalizeString_(fields.UserName) || normalizeString_(fields.Email);

    return {
      first: first,
      middle: middle,
      last: last,
      full: full,
      display: display,
      preferred: normalizeString_(fields.PreferredName) || first || display
    };
  }

  function buildCampaignContext_(userId, fields, reference) {
    var assignments = [];
    var campaignIds = new Set();
    var campaignNames = [];
    var primaryId = normalizeId_(fields.PrimaryCampaignId || fields.CampaignID || fields.CampaignId);
    var primaryName = '';

    var userCampaigns = Array.isArray(reference.userCampaigns) ? reference.userCampaigns : [];
    for (var i = 0; i < userCampaigns.length; i++) {
      var row = userCampaigns[i];
      if (normalizeId_(row.UserId || row.UserID || row.User) !== userId) continue;
      var campaignId = normalizeId_(row.CampaignId || row.CampaignID || row.Campaign);
      if (!campaignId) continue;
      campaignIds.add(campaignId);
      var campaign = reference.campaignsById[campaignId];
      var campaignName = campaign ? normalizeString_(campaign.Name || campaign.DisplayName) : '';
      if (campaignName) campaignNames.push(campaignName);
      var isPrimary = toBoolean_(row.IsPrimary);
      if (!primaryId && isPrimary) primaryId = campaignId;
      assignments.push({
        campaignId: campaignId,
        campaignName: campaignName,
        role: row.Role || row.RoleName || '',
        permissionLevel: row.PermissionLevel || '',
        isPrimary: isPrimary,
        createdAt: row.CreatedAt || row.CreatedOn || '',
        updatedAt: row.UpdatedAt || row.UpdatedOn || ''
      });
    }

    if (primaryId && !campaignIds.has(primaryId)) {
      campaignIds.add(primaryId);
    }

    if (primaryId) {
      var primaryCampaign = reference.campaignsById[primaryId];
      if (primaryCampaign) primaryName = normalizeString_(primaryCampaign.Name || primaryCampaign.DisplayName || primaryCampaign.Title);
      if (primaryName && campaignNames.indexOf(primaryName) === -1) campaignNames.unshift(primaryName);
    }

    var idsArray = Array.from(campaignIds);

    return {
      primaryId: primaryId,
      primaryName: primaryName,
      ids: idsArray,
      names: campaignNames,
      assignments: assignments
    };
  }

  function buildRoleContext_(userId, reference) {
    var assignments = [];
    var roleIds = new Set();
    var roleNames = [];

    var userRoles = Array.isArray(reference.userRoles) ? reference.userRoles : [];
    for (var i = 0; i < userRoles.length; i++) {
      var row = userRoles[i];
      if (normalizeId_(row.UserId || row.UserID || row.User) !== userId) continue;
      var roleId = normalizeId_(row.RoleId || row.RoleID || row.Role);
      var role = reference.rolesById[roleId];
      var roleName = role ? normalizeString_(role.Name || role.DisplayName) : normalizeString_(row.RoleName || row.Role);
      if (roleId) roleIds.add(roleId);
      if (roleName) roleNames.push(roleName);
      assignments.push({
        roleId: roleId,
        roleName: roleName,
        scope: row.Scope || role && role.Scope || '',
        assignedBy: row.AssignedBy || '',
        createdAt: row.CreatedAt || row.CreatedOn || '',
        updatedAt: row.UpdatedAt || row.UpdatedOn || ''
      });
    }

    return {
      ids: Array.from(roleIds),
      names: roleNames,
      assignments: assignments
    };
  }

  function buildClaimContext_(userId, reference) {
    var claims = [];
    var rows = Array.isArray(reference.claims) ? reference.claims : [];
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (normalizeId_(row.UserId || row.UserID || row.User) !== userId) continue;
      var claimType = normalizeString_(row.ClaimType || row.Type || row.Name);
      if (!claimType) continue;
      claims.push({
        type: claimType,
        createdAt: row.CreatedAt || row.CreatedOn || '',
        updatedAt: row.UpdatedAt || row.UpdatedOn || ''
      });
    }
    return claims;
  }

  function buildInsuranceContext_(userId, fields, reference) {
    var insurance = reference.insuranceByUserId[userId] || null;
    if (!insurance) {
      insurance = {
        InsuranceEligibleDate: fields.InsuranceEligibleDate || '',
        InsuranceQualifiedDate: fields.InsuranceQualifiedDate || fields.InsuranceEligibleDate || '',
        InsuranceEligible: fields.InsuranceEligible || fields.InsuranceQualified || false,
        InsuranceQualified: fields.InsuranceQualified || fields.InsuranceEligible || false,
        InsuranceEnrolled: fields.InsuranceEnrolled || fields.InsuranceSignedUp || false,
        InsuranceCardReceivedDate: fields.InsuranceCardReceivedDate || ''
      };
    }
    return Object.assign({}, insurance);
  }

  function computeAuthenticationLevel_(security) {
    if (!security) return 'unknown';
    if (!security.canLogin) return 'disabled';
    if (security.isAdmin) return 'privileged';
    if (security.twoFactorEnabled || security.mfaEnabled) return 'multifactor';
    if (security.emailConfirmed) return 'verified';
    return 'authenticated';
  }

  function buildSecurityContext_(fields, reference, userId) {
    var sessions = Array.isArray(reference.sessions) ? reference.sessions : [];
    var activeSessions = 0;
    var now = Date.now();
    for (var i = 0; i < sessions.length; i++) {
      var row = sessions[i];
      if (normalizeId_(row.UserId || row.UserID || row.User) !== userId) continue;
      var expiresAt = parseDate_(row.ExpiresAt || row.Expiration || row.ExpiredAt);
      if (!expiresAt || expiresAt.getTime() >= now) {
        activeSessions += 1;
      }
    }

    var security = {
      canLogin: toBoolean_(fields.CanLogin !== undefined ? fields.CanLogin : true),
      isAdmin: toBoolean_(fields.IsAdmin),
      emailConfirmed: toBoolean_(fields.EmailConfirmed),
      twoFactorEnabled: toBoolean_(fields.TwoFactorEnabled) || toBoolean_(fields.MFAEnabled),
      twoFactorDelivery: fields.TwoFactorDelivery || fields.MFADeliveryPreference || '',
      mfaEnabled: toBoolean_(fields.MFAEnabled),
      hasPassword: Boolean(fields.PasswordHash || fields.PasswordHashHex || fields.PasswordHashBase64),
      resetRequired: toBoolean_(fields.ResetRequired),
      lockoutEnd: fields.LockoutEnd || '',
      lockoutEnabled: toBoolean_(fields.LockoutEnabled),
      accessFailedCount: Number(fields.AccessFailedCount || 0),
      lastLoginAt: fields.LastLoginAt || fields.LastLogin || '',
      lastLoginIp: fields.LastLoginIp || '',
      lastLoginUserAgent: fields.LastLoginUserAgent || '',
      securityStamp: fields.SecurityStamp || '',
      concurrencyStamp: fields.ConcurrencyStamp || '',
      activeSessionCount: activeSessions,
      authenticationLevel: 'unknown'
    };

    security.authenticationLevel = computeAuthenticationLevel_(security);
    return security;
  }

  function buildEmploymentContext_(fields) {
    return {
      status: fields.EmploymentStatus || '',
      type: fields.EmploymentType || '',
      department: fields.Department || '',
      jobTitle: fields.JobTitle || '',
      managerId: fields.ManagerId || fields.ManagerUserID || '',
      managerName: fields.ManagerName || '',
      hireDate: fields.HireDate || '',
      seniorityDate: fields.SeniorityDate || '',
      terminationDate: fields.TerminationDate || '',
      probationMonths: fields.ProbationMonths || '',
      probationEnd: fields.ProbationEnd || '',
      probationEndDate: fields.ProbationEndDate || ''
    };
  }

  function createReferenceContext_() {
    var context = {};
    context.roles = safeReadSheet_(ROLES_SHEET, { useCache: true });
    context.userRoles = safeReadSheet_(USER_ROLES_SHEET, { useCache: true });
    context.claims = safeReadSheet_(USER_CLAIMS_SHEET, { useCache: true });
    context.userCampaigns = safeReadSheet_(USER_CAMPAIGNS_SHEET, { useCache: true });
    context.campaigns = safeReadSheet_(CAMPAIGNS_SHEET, { useCache: true });
    context.sessions = safeReadSheet_(SESSIONS_SHEET, { useCache: false });
    context.insurance = safeReadSheet_(USER_INSURANCE_SHEET, { useCache: true });

    context.rolesById = {};
    for (var i = 0; i < context.roles.length; i++) {
      var role = context.roles[i];
      var roleId = normalizeId_(role.ID || role.Id || role.RoleId || role.Guid);
      if (roleId) context.rolesById[roleId] = role;
    }

    context.campaignsById = {};
    for (var j = 0; j < context.campaigns.length; j++) {
      var campaign = context.campaigns[j];
      var campaignId = normalizeId_(campaign.ID || campaign.Id || campaign.CampaignId);
      if (campaignId) context.campaignsById[campaignId] = campaign;
    }

    context.insuranceByUserId = {};
    for (var k = 0; k < context.insurance.length; k++) {
      var record = context.insurance[k];
      var userId = normalizeId_(record.UserId || record.UserID || record.ID || record.Id);
      if (!userId || context.insuranceByUserId[userId]) continue;
      context.insuranceByUserId[userId] = record;
    }

    return context;
  }

  function cloneObject_(obj) {
    var clone = {};
    if (obj && typeof obj === 'object') {
      Object.keys(obj).forEach(function (key) {
        clone[key] = obj[key];
      });
    }
    return clone;
  }

  function buildIdentityStateFromUser(userRecord, reference) {
    if (!userRecord) return null;
    var referenceContext = reference && typeof reference === 'object' ? reference : createReferenceContext_();

    var projected = {};
    if (typeof projectRecordToCanonicalUser === 'function') {
      try {
        projected = projectRecordToCanonicalUser(userRecord, { preferIdentityService: false }) || {};
      } catch (projectionErr) {
        console.warn('LuminaIdentityService.buildIdentityStateFromUser: projection failed', projectionErr);
      }
    }

    var fields = cloneObject_(projected);
    Object.keys(userRecord).forEach(function (key) {
      if (!Object.prototype.hasOwnProperty.call(fields, key)) {
        fields[key] = userRecord[key];
      }
    });

    fields.IdentityVersion = IDENTITY_VERSION;
    fields.IdentityEvaluatedAt = new Date();

    var names = computeNameParts_(fields);
    fields.FirstName = names.first;
    fields.MiddleName = names.middle;
    fields.LastName = names.last;
    fields.FullName = names.full;
    fields.DisplayName = names.display;
    fields.PreferredName = names.preferred;

    var userId = normalizeId_(fields.ID || fields.Id || fields.UserId || fields.UserID);
    fields.ID = userId || fields.ID || fields.Id || fields.UserId || fields.UserID;

    var campaigns = buildCampaignContext_(userId, fields, referenceContext);
    fields.PrimaryCampaignId = campaigns.primaryId;
    fields.CampaignIds = campaigns.ids.join(', ');
    fields.CampaignNames = campaigns.names.join(', ');
    fields.CampaignAssignments = JSON.stringify(campaigns.assignments || []);

    var roles = buildRoleContext_(userId, referenceContext);
    fields.RoleIds = roles.ids.join(', ');
    fields.RoleNames = roles.names.join(', ');

    var claims = buildClaimContext_(userId, referenceContext);
    fields.Claims = claims.map(function (claim) { return claim.type; }).join(', ');
    fields.ClaimTypes = fields.Claims;

    var insurance = buildInsuranceContext_(userId, fields, referenceContext);
    Object.keys(insurance).forEach(function (key) {
      if (!Object.prototype.hasOwnProperty.call(fields, key) || !fields[key]) {
        fields[key] = insurance[key];
      }
    });

    var security = buildSecurityContext_(fields, referenceContext, userId);
    fields.ActiveSessionCount = security.activeSessionCount;
    fields.SessionIdleTimeout = fields.SessionIdleTimeout || fields.IdleTimeoutMinutes || '';

    var employment = buildEmploymentContext_(fields);

    var identity = {
      system: IDENTITY_SYSTEM_NAME,
      version: IDENTITY_VERSION,
      id: userId,
      userName: normalizeString_(fields.UserName),
      normalizedUserName: normalizeString_(fields.NormalizedUserName) || normalizeString_(fields.UserName).toUpperCase(),
      email: normalizeString_(fields.Email),
      normalizedEmail: normalizeString_(fields.NormalizedEmail) || normalizeEmail_(fields.Email),
      firstName: names.first,
      middleName: names.middle,
      lastName: names.last,
      fullName: names.full,
      displayName: names.display,
      preferredName: names.preferred,
      pronouns: normalizeString_(fields.Pronouns),
      avatarUrl: normalizeString_(fields.AvatarUrl),
      campaigns: campaigns,
      roles: roles,
      claims: { list: claims, types: claims.map(function (c) { return c.type; }) },
      insurance: insurance,
      employment: employment,
      security: security,
      fields: fields,
      headers: getCanonicalIdentityHeaders_()
    };

    return identity;
  }

  function summarizeIdentityForClient(identity) {
    if (!identity) return null;
    var security = identity.security || {};
    var employment = identity.employment || {};
    var insurance = identity.insurance || {};
    var campaigns = identity.campaigns || { ids: [], assignments: [] };
    var roles = identity.roles || { names: [], ids: [] };

    return {
      id: identity.id,
      displayName: identity.displayName,
      email: identity.email,
      userName: identity.userName,
      authenticationLevel: security.authenticationLevel,
      canLogin: security.canLogin,
      isAdmin: security.isAdmin,
      roleNames: roles.names.slice(),
      roleIds: roles.ids.slice(),
      roles: roles.names.slice(),
      claims: identity.claims ? identity.claims.types.slice() : [],
      campaignIds: campaigns.ids.slice(),
      primaryCampaignId: campaigns.primaryId || '',
      primaryCampaignName: campaigns.primaryName || '',
      campaignAssignments: campaigns.assignments.slice(),
      employmentStatus: employment.status || '',
      jobTitle: employment.jobTitle || '',
      department: employment.department || '',
      hireDate: employment.hireDate || '',
      terminationDate: employment.terminationDate || '',
      insuranceEligibleDate: insurance.InsuranceEligibleDate || '',
      insuranceQualified: toBoolean_(insurance.InsuranceQualified || insurance.InsuranceEligible),
      activeSessionCount: security.activeSessionCount || 0,
      updatedAt: identity.fields.UpdatedAt || identity.fields.UpdatedOn || '',
      createdAt: identity.fields.CreatedAt || identity.fields.CreatedOn || ''
    };
  }

  function evaluateIdentityForAuthentication(identity) {
    if (!identity) {
      return {
        status: 'unknown',
        authenticationLevel: 'unknown',
        warnings: ['Identity not available']
      };
    }

    var security = identity.security || {};
    var campaigns = identity.campaigns || { ids: [] };
    var roles = identity.roles || { ids: [] };
    var warnings = [];

    if (!security.canLogin) warnings.push('Login disabled');
    if (!security.emailConfirmed) warnings.push('Email not confirmed');
    if (!roles.ids.length) warnings.push('No roles assigned');
    if (!campaigns.ids.length) warnings.push('No campaign assignment');
    if (!security.twoFactorEnabled) warnings.push('Multi-factor authentication not enabled');

    return {
      status: security.canLogin ? 'active' : 'restricted',
      authenticationLevel: security.authenticationLevel || 'unknown',
      canLogin: !!security.canLogin,
      isAdmin: !!security.isAdmin,
      hasTwoFactor: !!security.twoFactorEnabled,
      roleCount: roles.ids.length,
      campaignCount: campaigns.ids.length,
      warnings: warnings,
      activeSessionCount: security.activeSessionCount || 0
    };
  }

  function resolveUserRecordById_(userId, reference) {
    var id = normalizeId_(userId);
    if (!id) return null;
    var users = safeReadSheet_(LUMINA_IDENTITY_SHEET || USERS_SHEET, { useCache: true });
    for (var i = 0; i < users.length; i++) {
      var record = users[i];
      var recordId = normalizeId_(record.ID || record.Id || record.UserId || record.UserID);
      if (recordId === id) return record;
    }
    return null;
  }

  function resolveUserRecordByEmail_(email) {
    var normalizedEmail = normalizeEmail_(email);
    if (!normalizedEmail) return null;
    var users = safeReadSheet_(LUMINA_IDENTITY_SHEET || USERS_SHEET, { useCache: true });
    for (var i = 0; i < users.length; i++) {
      var record = users[i];
      var recordEmail = normalizeEmail_(record.Email || record.UserEmail);
      if (recordEmail && recordEmail === normalizedEmail) return record;
    }
    return null;
  }

  function buildIdentityResult_(identity) {
    if (!identity) return null;
    var summary = summarizeIdentityForClient(identity);
    var evaluation = evaluateIdentityForAuthentication(identity);
    var result = {
      identity: identity,
      identitySummary: summary,
      summary: summary,
      identityEvaluation: evaluation,
      evaluation: evaluation,
      warnings: (evaluation && Array.isArray(evaluation.warnings)) ? evaluation.warnings.slice() : [],
      headers: identity.headers.slice(),
      identityHeaders: identity.headers.slice(),
      identityFields: identity.fields
    };
    identity.fields.IdentityEvaluationWarnings = result.warnings.join('; ');
    return result;
  }

  function getUserIdentityById(userId) {
    var record = resolveUserRecordById_(userId);
    if (!record) return null;
    var context = createReferenceContext_();
    var identity = buildIdentityStateFromUser(record, context);
    return buildIdentityResult_(identity);
  }

  function getUserIdentityByEmail(email) {
    var record = resolveUserRecordByEmail_(email);
    if (!record) return null;
    var context = createReferenceContext_();
    var identity = buildIdentityStateFromUser(record, context);
    return buildIdentityResult_(identity);
  }

  function resolveUserIdFromIdentifier_(identifier, context) {
    if (!identifier && identifier !== 0) return '';
    if (typeof identifier === 'object') {
      if (identifier.id || identifier.ID) return normalizeId_(identifier.id || identifier.ID);
      if (identifier.userId || identifier.UserId || identifier.UserID) return normalizeId_(identifier.userId || identifier.UserId || identifier.UserID);
      if (identifier.email || identifier.Email) {
        var record = resolveUserRecordByEmail_(identifier.email || identifier.Email);
        if (record) return normalizeId_(record.ID || record.Id || record.UserId || record.UserID);
      }
    }

    var str = String(identifier);
    if (!str) return '';
    if (str.indexOf('@') !== -1) {
      var recordFromEmail = resolveUserRecordByEmail_(str);
      return recordFromEmail ? normalizeId_(recordFromEmail.ID || recordFromEmail.Id || recordFromEmail.UserId || recordFromEmail.UserID) : '';
    }
    return normalizeId_(str);
  }

  function hasActiveSession(userIdentifier) {
    var context = createReferenceContext_();
    var userId = resolveUserIdFromIdentifier_(userIdentifier, context);
    if (!userId) return false;
    var security = buildSecurityContext_({
      CanLogin: true,
      EmailConfirmed: true
    }, context, userId);
    return security.activeSessionCount > 0;
  }

  function ensureIdentityLogSheet_() {
    try {
      if (typeof ensureSheetWithHeaders === 'function') {
        return ensureSheetWithHeaders(LUMINA_IDENTITY_LOGS_SHEET, LUMINA_IDENTITY_LOGS_HEADERS);
      }
      if (typeof ensureSheetStructureFromDefinition_ === 'function') {
        return ensureSheetStructureFromDefinition_({
          name: LUMINA_IDENTITY_LOGS_SHEET,
          headers: LUMINA_IDENTITY_LOGS_HEADERS
        });
      }
    } catch (logErr) {
      console.warn('LuminaIdentityService.ensureIdentityLogSheet_: failed', logErr);
    }
    return null;
  }

  function logIdentityEvent(eventType, payload) {
    var type = normalizeString_(eventType);
    if (!type) {
      throw new Error('Identity log event type is required');
    }

    var sh = ensureIdentityLogSheet_();
    if (!sh) {
      throw new Error('Unable to access identity log sheet');
    }

    var headers = Array.isArray(LUMINA_IDENTITY_LOGS_HEADERS) ? LUMINA_IDENTITY_LOGS_HEADERS : ['Timestamp', 'EventType', 'UserId', 'Message'];
    var row = new Array(headers.length).fill('');
    var identity = payload && payload.identity ? payload.identity : null;
    if (!identity && payload && payload.user) {
      identity = buildIdentityStateFromUser(payload.user);
    }

    var summary = payload && payload.summary ? payload.summary : (identity ? summarizeIdentityForClient(identity) : null);
    var security = identity ? identity.security : null;

    for (var i = 0; i < headers.length; i++) {
      var header = headers[i];
      switch (header) {
        case 'Timestamp':
          row[i] = new Date();
          break;
        case 'EventType':
          row[i] = type;
          break;
        case 'UserId':
          row[i] = identity ? identity.id : (payload && payload.userId ? payload.userId : '');
          break;
        case 'UserName':
          row[i] = identity ? identity.userName : (summary ? summary.userName : '');
          break;
        case 'DisplayName':
          row[i] = identity ? identity.displayName : (summary ? summary.displayName : '');
          break;
        case 'Email':
          row[i] = identity ? identity.email : (summary ? summary.email : '');
          break;
        case 'CampaignContext':
          row[i] = identity && identity.campaigns ? JSON.stringify(identity.campaigns.assignments || []) : '';
          break;
        case 'RoleContext':
          row[i] = identity && identity.roles ? JSON.stringify(identity.roles.assignments || []) : '';
          break;
        case 'ClaimContext':
          row[i] = identity && identity.claims ? JSON.stringify(identity.claims.list || []) : '';
          break;
        case 'AuthenticationLevel':
          row[i] = security ? security.authenticationLevel : '';
          break;
        case 'SessionState':
          row[i] = security ? security.activeSessionCount : '';
          break;
        case 'Source':
          row[i] = payload && payload.source ? payload.source : IDENTITY_SYSTEM_NAME;
          break;
        case 'Message':
          row[i] = payload && payload.message ? payload.message : '';
          break;
        case 'Metadata':
          row[i] = payload && payload.metadata ? JSON.stringify(payload.metadata) : '';
          break;
        default:
          row[i] = '';
      }
    }

    sh.appendRow(row);
    try { invalidateCache && invalidateCache(LUMINA_IDENTITY_LOGS_SHEET); } catch (_) { /* ignore */ }
    return { success: true };
  }

  exports.listIdentityFields = getCanonicalIdentityHeaders_;
  exports.ensureIdentityHeaders = ensureIdentityHeaders;
  exports.ensureInfrastructure = ensureInfrastructure;
  exports.buildIdentityStateFromUser = buildIdentityStateFromUser;
  exports.summarizeIdentityForClient = summarizeIdentityForClient;
  exports.evaluateIdentityForAuthentication = evaluateIdentityForAuthentication;
  exports.getUserIdentityById = getUserIdentityById;
  exports.getUserIdentityByEmail = getUserIdentityByEmail;
  exports.hasActiveSession = hasActiveSession;
  exports.logIdentityEvent = logIdentityEvent;

  return exports;
})(typeof IdentityService !== 'undefined' ? IdentityService : {});

