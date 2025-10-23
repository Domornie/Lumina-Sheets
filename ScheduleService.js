/**
 * COMPLETE Enhanced Schedule Management Backend Service
 * Version 4.1 - Integrated with ScheduleUtilities and MainUtilities
 * Now properly uses dedicated spreadsheet support and shared functions
 */

// ────────────────────────────────────────────────────────────────────────────
// CONFIGURATION - Uses ScheduleUtilities constants
// ────────────────────────────────────────────────────────────────────────────

const SCHEDULE_SETTINGS = (typeof getScheduleConfig === 'function')
  ? getScheduleConfig()
  : {
      PRIMARY_COUNTRY: 'JM',
      SUPPORTED_COUNTRIES: ['JM', 'US', 'DO', 'PH'],
      DEFAULT_SHIFT_CAPACITY: 10,
      DEFAULT_BREAK_MINUTES: 15,
      DEFAULT_LUNCH_MINUTES: 60,
      CACHE_DURATION: 300
    };

const DEFAULT_SCHEDULE_TIME_ZONE = (typeof Session !== 'undefined' && typeof Session.getScriptTimeZone === 'function')
  ? Session.getScriptTimeZone()
  : 'UTC';

function resolveSchedulePeriodStart(record, timeZone = DEFAULT_SCHEDULE_TIME_ZONE) {
  if (!record || typeof record !== 'object') {
    return '';
  }

  const candidates = [
    record.PeriodStart,
    record.StartDate,
    record.ScheduleStart,
    record.AssignmentStart,
    record.Date,
    record.ScheduleDate,
    record.Day
  ];

  for (let i = 0; i < candidates.length; i++) {
    const normalized = normalizeDateForSheet(candidates[i], timeZone);
    if (normalized) {
      return normalized;
    }
  }

  return '';
}

function resolveSchedulePeriodEnd(record, fallbackStart = '', timeZone = DEFAULT_SCHEDULE_TIME_ZONE) {
  if (!record || typeof record !== 'object') {
    return '';
  }

  const candidates = [
    record.PeriodEnd,
    record.EndDate,
    record.ScheduleEnd,
    record.AssignmentEnd,
    record.Date,
    record.ScheduleDate,
    record.Day,
    fallbackStart
  ];

  for (let i = 0; i < candidates.length; i++) {
    const normalized = normalizeDateForSheet(candidates[i], timeZone);
    if (normalized) {
      return normalized;
    }
  }

  return '';
}

function resolveSchedulePeriodStartDate(record, timeZone = DEFAULT_SCHEDULE_TIME_ZONE) {
  const start = resolveSchedulePeriodStart(record, timeZone);
  if (!start) {
    return null;
  }

  const startDate = new Date(start);
  return isNaN(startDate.getTime()) ? null : startDate;
}

function resolveSchedulePeriodEndDate(record, timeZone = DEFAULT_SCHEDULE_TIME_ZONE) {
  const start = resolveSchedulePeriodStart(record, timeZone);
  const end = resolveSchedulePeriodEnd(record, start, timeZone);
  if (!end) {
    return null;
  }

  const endDate = new Date(end);
  return isNaN(endDate.getTime()) ? null : endDate;
}

function normalizeSchedulePeriodRecord(record, timeZone = DEFAULT_SCHEDULE_TIME_ZONE) {
  if (!record || typeof record !== 'object') {
    return record;
  }

  const normalizedStart = resolveSchedulePeriodStart(record, timeZone);
  const normalizedEnd = resolveSchedulePeriodEnd(record, normalizedStart, timeZone);

  if (!normalizedStart && !normalizedEnd) {
    return record;
  }

  const normalizedRecord = Object.assign({}, record);

  if (normalizedStart) {
    normalizedRecord.PeriodStart = normalizedStart;
    normalizedRecord.Date = normalizedStart;
  }

  if (normalizedEnd) {
    normalizedRecord.PeriodEnd = normalizedEnd;
  }

  return normalizedRecord;
}

function buildScheduleCompositeKey(record, timeZone = DEFAULT_SCHEDULE_TIME_ZONE) {
  const normalizedRecord = normalizeSchedulePeriodRecord(record, timeZone);
  const userPart = normalizeUserKey(
    (normalizedRecord && (normalizedRecord.UserName || normalizedRecord.UserID || normalizedRecord.userName || normalizedRecord.userId))
      || ''
  );

  const start = normalizedRecord ? normalizedRecord.PeriodStart || '' : '';
  const end = normalizedRecord ? normalizedRecord.PeriodEnd || start : '';

  return `${userPart}::${start}::${end}`;
}

function getSchedulePeriodSortValue(record, timeZone = DEFAULT_SCHEDULE_TIME_ZONE) {
  const startDate = resolveSchedulePeriodStartDate(record, timeZone);
  return startDate ? startDate.getTime() : 0;
}

// ────────────────────────────────────────────────────────────────────────────
// CORE SCHEDULE STORAGE HELPERS
// ────────────────────────────────────────────────────────────────────────────

function ensureShiftAssignmentsSheet() {
  return ensureScheduleSheetWithHeaders(SHIFT_ASSIGNMENTS_SHEET, SHIFT_ASSIGNMENTS_HEADERS);
}

function ensureAuditLogSheet() {
  return ensureScheduleSheetWithHeaders(AUDIT_LOG_SHEET, AUDIT_LOG_HEADERS);
}

function appendAuditLogEntry(action, entityType, entityId, beforeObj, afterObj, notes) {
  try {
    const sheet = ensureAuditLogSheet();
    const actor = typeof getCurrentUser === 'function' ? getCurrentUser() : null;
    const actorName = actor && (actor.Email || actor.email || actor.UserName || actor.name) || 'System';
    const timestamp = new Date();

    const row = [
      timestamp,
      actorName,
      action,
      entityType,
      entityId || '',
      beforeObj ? JSON.stringify(beforeObj) : '',
      afterObj ? JSON.stringify(afterObj) : '',
      notes || ''
    ];

    sheet.appendRow(row);
  } catch (error) {
    console.warn('Failed to append audit log entry:', error && error.message ? error.message : error);
  }
}

function readShiftAssignments() {
  return readScheduleSheet(SHIFT_ASSIGNMENTS_SHEET) || [];
}

function normalizeAssignmentRecord(record) {
  if (!record || typeof record !== 'object') {
    return null;
  }

  const normalized = Object.assign({}, record);
  const startDate = normalizeDateForSheet(record.StartDate || record.PeriodStart || record.Date, DEFAULT_SCHEDULE_TIME_ZONE);
  const endDate = normalizeDateForSheet(record.EndDate || record.PeriodEnd || record.Date, DEFAULT_SCHEDULE_TIME_ZONE);
  if (startDate) {
    normalized.StartDate = startDate;
  }
  if (endDate) {
    normalized.EndDate = endDate;
  }

  normalized.Status = (record.Status || 'Pending').toString().toUpperCase();
  normalized.AllowSwap = scheduleFlagToBool(record.AllowSwap || record.AllowSwaps || record.allowSwap);
  normalized.Premiums = record.Premiums || '';
  normalized.BreaksConfigJSON = record.BreaksConfigJSON || record.BreaksJson || '';

  if (!normalized.AssignmentId && record.ID) {
    normalized.AssignmentId = record.ID;
  }

  if (!normalized.UserName && record.UserID) {
    const users = readSheet(USERS_SHEET) || [];
    const match = users.find(u => String(u.ID) === String(record.UserID));
    if (match) {
      normalized.UserName = match.UserName || match.FullName || '';
    }
  }

  normalized.StartDateObj = normalized.StartDate ? new Date(normalized.StartDate) : null;
  normalized.EndDateObj = normalized.EndDate ? new Date(normalized.EndDate) : null;

  return normalized;
}

function writeShiftAssignments(assignments, actorId, notes, statusOverride) {
  if (!Array.isArray(assignments) || !assignments.length) {
    return { success: false, error: 'No assignments to write' };
  }

  const sheet = ensureShiftAssignmentsSheet();
  const now = new Date();
  const actor = actorId || (typeof getCurrentUser === 'function' ? getCurrentUser()?.Email : 'System');

  const rows = assignments.map(assignment => {
    const normalized = Object.assign({}, assignment);
    normalized.AssignmentId = normalized.AssignmentId || Utilities.getUuid();
    normalized.CreatedAt = normalized.CreatedAt || now;
    normalized.CreatedBy = normalized.CreatedBy || actor;
    normalized.UpdatedAt = now;
    normalized.UpdatedBy = actor;
    if (statusOverride) {
      normalized.Status = statusOverride;
    } else {
      normalized.Status = normalized.Status || 'PENDING';
    }

    return SHIFT_ASSIGNMENTS_HEADERS.map(header => Object.prototype.hasOwnProperty.call(normalized, header) ? normalized[header] : '');
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, SHIFT_ASSIGNMENTS_HEADERS.length).setValues(rows);
  SpreadsheetApp.flush();

  assignments.forEach(assignment => {
    appendAuditLogEntry(
      'CREATE',
      'ShiftAssignment',
      assignment.AssignmentId,
      null,
      assignment,
      notes || ''
    );
  });

  return { success: true, count: rows.length };
}

function updateShiftAssignmentRow(assignmentId, updater) {
  const sheet = ensureShiftAssignmentsSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { success: false, error: 'No assignments found' };
  }

  const headers = data[0];
  const idIndex = headers.indexOf('AssignmentId');
  if (idIndex === -1) {
    return { success: false, error: 'Assignment sheet missing AssignmentId column' };
  }

  for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
    if (String(data[rowIndex][idIndex]) === String(assignmentId)) {
      const rowObject = {};
      headers.forEach((header, columnIndex) => {
        rowObject[header] = data[rowIndex][columnIndex];
      });

      const before = Object.assign({}, rowObject);
      const updated = updater(rowObject) || rowObject;

      const rowValues = SHIFT_ASSIGNMENTS_HEADERS.map(header => Object.prototype.hasOwnProperty.call(updated, header) ? updated[header] : '');
      sheet.getRange(rowIndex + 1, 1, 1, SHIFT_ASSIGNMENTS_HEADERS.length).setValues([rowValues]);
      SpreadsheetApp.flush();

      appendAuditLogEntry('UPDATE', 'ShiftAssignment', assignmentId, before, updated, 'Assignment updated');

      return { success: true, assignment: updated };
    }
  }

  return { success: false, error: 'Assignment not found' };
}

function buildDateSeries(startDateStr, endDateStr) {
  const start = new Date(startDateStr);
  const end = new Date(endDateStr);
  if (isNaN(start.getTime()) || isNaN(end.getTime()) || start > end) {
    return [];
  }

  const dates = [];
  const current = new Date(start.getTime());
  while (current <= end) {
    dates.push(Utilities.formatDate(current, DEFAULT_SCHEDULE_TIME_ZONE, 'yyyy-MM-dd'));
    current.setDate(current.getDate() + 1);
  }

  return dates;
}

function loadHolidayMap(startDateStr, endDateStr) {
  const holidays = readScheduleSheet(HOLIDAYS_SHEET) || [];
  const holidayMap = new Map();
  if (!holidays.length) {
    return holidayMap;
  }

  const dateRange = buildDateSeries(startDateStr, endDateStr);
  const dateSet = new Set(dateRange);

  holidays.forEach(holiday => {
    const dateStr = normalizeDateForSheet(holiday.Date, DEFAULT_SCHEDULE_TIME_ZONE);
    if (!dateStr || (dateSet.size && !dateSet.has(dateStr))) {
      return;
    }
    const entry = holidayMap.get(dateStr) || [];
    entry.push({
      name: holiday.Name || '',
      region: holiday.Region || '',
      isWorkingDay: scheduleFlagToBool(holiday.IsWorkingDayOverride, false)
    });
    holidayMap.set(dateStr, entry);
  });

  return holidayMap;
}

function isWeekendDate(dateStr) {
  const date = new Date(dateStr);
  if (isNaN(date.getTime())) {
    return false;
  }
  const day = date.getDay();
  return day === 0 || day === 6;
}

function createSeededRandom(seedValue) {
  let seed = 0;
  if (typeof seedValue === 'number') {
    seed = seedValue;
  } else if (seedValue) {
    const text = String(seedValue);
    for (let i = 0; i < text.length; i++) {
      seed = (seed << 5) - seed + text.charCodeAt(i);
      seed |= 0;
    }
  } else {
    seed = Date.now();
  }

  return function seededRandom() {
    seed = (seed * 9301 + 49297) % 233280;
    return seed / 233280;
  };
}

function shuffleWithSeed(array, seedValue) {
  const shuffled = array.slice();
  const random = createSeededRandom(seedValue);
  for (let i = shuffled.length - 1; i > 0; i--) {
    const j = Math.floor(random() * (i + 1));
    [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
  }
  return shuffled;
}

function storeSchedulePreview(previewData) {
  const cache = CacheService.getScriptCache();
  const token = Utilities.getUuid();
  cache.put(`schedule_preview_${token}`, JSON.stringify(previewData), 600);
  return token;
}

function loadSchedulePreview(token) {
  if (!token) {
    return null;
  }
  const cache = CacheService.getScriptCache();
  const payload = cache.get(`schedule_preview_${token}`);
  if (!payload) {
    return null;
  }
  try {
    return JSON.parse(payload);
  } catch (error) {
    console.warn('Failed to parse schedule preview payload:', error);
    return null;
  }
}

// ────────────────────────────────────────────────────────────────────────────
// USER MANAGEMENT FUNCTIONS - Integrated with MainUtilities
// ────────────────────────────────────────────────────────────────────────────

/**
 * Resolve the active schedule context for the requesting user or manager.
 * Provides manager/campaign identifiers, identity metadata, and managed roster.
 */
function clientGetScheduleContext(managerIdCandidate, campaignIdCandidate) {
  const providedManagerId = normalizeUserIdValue(managerIdCandidate);
  const providedCampaignId = normalizeCampaignIdValue(campaignIdCandidate);
  const timestamp = new Date().toISOString();

  const context = {
    success: false,
    providedManagerId,
    providedCampaignId,
    managerId: '',
    campaignId: '',
    user: null,
    managedUserIds: [],
    managedUserCount: 0,
    managedCampaigns: [],
    identity: null,
    authenticated: false,
    timestamp
  };

  try {
    const allUsers = readSheet(USERS_SHEET) || [];
    const usersById = new Map();
    const usersByUsername = new Map();

    allUsers.forEach(user => {
      if (!user || typeof user !== 'object') {
        return;
      }

      const normalizedId = normalizeUserIdValue(user.ID);
      if (normalizedId) {
        usersById.set(normalizedId, user);
      }

      const normalizedUsername = normalizeUserIdValue(user.UserName || user.Username);
      if (normalizedUsername) {
        const usernameKey = normalizedUsername.toLowerCase();
        if (!usersByUsername.has(usernameKey)) {
          usersByUsername.set(usernameKey, user);
        }
      }
    });

    const currentUser = (typeof getCurrentUser === 'function') ? getCurrentUser() : null;
    const currentUserId = normalizeUserIdValue(currentUser && (currentUser.ID || currentUser.UserID));

    const managerIdSources = [];
    let resolvedManagerId = providedManagerId;

    if (resolvedManagerId) {
      managerIdSources.push('parameter');
    }

    if (!resolvedManagerId && currentUserId) {
      resolvedManagerId = currentUserId;
      managerIdSources.push('current-user');
    }

    let resolvedManagerUser = resolvedManagerId ? usersById.get(resolvedManagerId) : null;

    if (!resolvedManagerUser && resolvedManagerId) {
      const key = resolvedManagerId.toLowerCase ? resolvedManagerId.toLowerCase() : String(resolvedManagerId || '').toLowerCase();
      const usernameMatch = usersByUsername.get(key);
      if (usernameMatch) {
        resolvedManagerUser = usernameMatch;
        resolvedManagerId = normalizeUserIdValue(usernameMatch.ID) || resolvedManagerId;
        managerIdSources.push('username-match');
      }
    }

    if (!resolvedManagerUser && currentUserId) {
      resolvedManagerUser = usersById.get(currentUserId) || currentUser || null;
    }

    const campaignIdSources = [];
    let resolvedCampaignId = providedCampaignId;

    if (resolvedCampaignId) {
      campaignIdSources.push('parameter');
    }

    const appendCampaignCandidate = (value, source) => {
      if (resolvedCampaignId) {
        return;
      }
      const normalized = normalizeCampaignIdValue(value);
      if (normalized) {
        resolvedCampaignId = normalized;
        campaignIdSources.push(source);
      }
    };

    appendCampaignCandidate(resolvedManagerUser && (resolvedManagerUser.CampaignID || resolvedManagerUser.campaignID || resolvedManagerUser.Campaign || resolvedManagerUser.campaign), 'manager-profile');
    appendCampaignCandidate(currentUser && (currentUser.CampaignID || currentUser.campaignID || currentUser.Campaign || currentUser.campaign), 'current-user');

    const managedSet = resolvedManagerId ? buildManagedUserSet(resolvedManagerId) : new Set();
    const managedUserIds = Array.from(managedSet).map(normalizeUserIdValue).filter(Boolean);

    let managedCampaigns = [];
    try {
      if (resolvedManagerId && typeof getUserManagedCampaigns === 'function') {
        const campaigns = getUserManagedCampaigns(resolvedManagerId) || [];
        managedCampaigns = campaigns
          .filter(Boolean)
          .map(campaign => ({
            id: normalizeCampaignIdValue(campaign.ID || campaign.Id || campaign.id),
            name: campaign.Name || campaign.name || '',
            isPrimary: scheduleFlagToBool(campaign.IsPrimary || campaign.isPrimary)
          }));

        if (!resolvedCampaignId) {
          const primary = managedCampaigns.find(campaign => campaign.isPrimary);
          if (primary && primary.id) {
            resolvedCampaignId = primary.id;
            campaignIdSources.push('managed-campaign');
          }
        }
      }
    } catch (campaignError) {
      console.warn('clientGetScheduleContext: unable to resolve managed campaigns', campaignError);
    }

    const roles = collectUserRoleCandidates(resolvedManagerUser || currentUser || {});
    const normalizedRoles = roles
      .map(role => String(role || '').trim())
      .filter(Boolean);

    const identity = {
      authenticated: !!currentUserId,
      resolvedAt: timestamp,
      managerId: resolvedManagerId || '',
      campaignId: resolvedCampaignId || '',
      providedManagerId,
      providedCampaignId,
      managerIdSources,
      campaignIdSources,
      managedUserCount: managedUserIds.length,
      roles: normalizedRoles,
      isAdmin: scheduleFlagToBool((resolvedManagerUser && resolvedManagerUser.IsAdmin) || (currentUser && currentUser.IsAdmin)),
      userId: currentUserId || resolvedManagerId || '',
      userName: (currentUser && (currentUser.UserName || currentUser.Username)) || '',
      fullName: (currentUser && (currentUser.FullName || currentUser.Name)) || '',
      email: (currentUser && (currentUser.Email || currentUser.email)) || ''
    };

    const clientUser = Object.assign({}, resolvedManagerUser || currentUser || {}, {
      ID: normalizeUserIdValue((resolvedManagerUser && resolvedManagerUser.ID) || (currentUser && currentUser.ID) || resolvedManagerId),
      CampaignID: resolvedCampaignId || (resolvedManagerUser && resolvedManagerUser.CampaignID) || '',
      Roles: normalizedRoles,
      IsAdmin: identity.isAdmin,
      managedUserCount: managedUserIds.length
    });

    if (!clientUser.UserName && clientUser.Username) {
      clientUser.UserName = clientUser.Username;
    }
    if (!clientUser.FullName && clientUser.Name) {
      clientUser.FullName = clientUser.Name;
    }

    context.permissions = {
      canManageSchedules: identity.isAdmin || managedUserIds.length > 0,
      canApproveSchedules: identity.isAdmin || managedUserIds.length > 0,
      canImport: identity.isAdmin,
      canEditShiftSlots: identity.isAdmin || normalizedRoles.some(role => role.toLowerCase() === 'workforce' || role.toLowerCase() === 'scheduler')
    };

    context.success = true;
    context.authenticated = identity.authenticated;
    context.managerId = resolvedManagerId || '';
    context.campaignId = resolvedCampaignId || '';
    context.user = clientUser;
    identity.permissions = context.permissions;

    context.identity = identity;
    context.managedUserIds = managedUserIds;
    context.managedUserCount = managedUserIds.length;
    context.managedCampaigns = managedCampaigns;

    return context;
  } catch (error) {
    console.error('❌ Error resolving schedule context:', error);
    context.error = error && error.message ? error.message : String(error || 'Unknown error');
    try {
      safeWriteError && safeWriteError('clientGetScheduleContext', error);
    } catch (_) {
      // ignore logging failures
    }
    return context;
  }
}

/**
 * Get users for schedule management with manager filtering
 * Uses MainUtilities user functions with campaign support
 */
function clientGetScheduleUsers(requestingUserId, campaignId = null) {
  try {
    const normalizedCampaignId = normalizeCampaignIdValue(campaignId);
    console.log('🔍 Getting schedule users for:', requestingUserId, 'campaign:', normalizedCampaignId || '(not provided)');

    // Use MainUtilities to get all users
    const allUsers = readSheet(USERS_SHEET) || [];
    if (allUsers.length === 0) {
      console.warn('No users found in Users sheet');
      return [];
    }

    const normalizedManagerId = normalizeUserIdValue(requestingUserId);
    let requestingUser = null;
    if (normalizedManagerId) {
      requestingUser = allUsers.find(u => normalizeUserIdValue(u && u.ID) === normalizedManagerId) || null;
    }

    let effectiveCampaignId = normalizedCampaignId;
    if (!effectiveCampaignId && requestingUser) {
      const managerCampaignCandidates = [
        requestingUser.CampaignID,
        requestingUser.campaignID,
        requestingUser.CampaignId,
        requestingUser.campaignId,
        requestingUser.Campaign,
        requestingUser.campaign
      ];

      for (let i = 0; i < managerCampaignCandidates.length; i++) {
        const candidate = normalizeCampaignIdValue(managerCampaignCandidates[i]);
        if (candidate) {
          effectiveCampaignId = candidate;
          break;
        }
      }
    }

    let filteredUsers = allUsers;

    // Filter by campaign if specified - use MainUtilities campaign functions
    if (effectiveCampaignId) {
      filteredUsers = filterUsersByCampaign(allUsers, effectiveCampaignId);
    }

    // Apply manager permissions using MainUtilities functions
    if (normalizedManagerId) {
      if (requestingUser) {
        const isAdmin = scheduleFlagToBool(requestingUser.IsAdmin);

        if (!isAdmin) {
          const managedUserIds = buildManagedUserSet(normalizedManagerId);

          filteredUsers = filteredUsers.filter(user => managedUserIds.has(normalizeUserIdValue(user && user.ID)));
        }
      } else {
        console.warn('Requesting user not found when applying manager filter:', requestingUserId);
      }
    }

    // Transform to schedule-friendly format
    const scheduleUsers = filteredUsers
      .filter(user => user && user.ID && (user.UserName || user.FullName))
      .filter(user => !isScheduleNameRestricted(user))
      .filter(user => !isScheduleRoleRestricted(user))
      .filter(user => isUserConsideredActive(user))
      .map(user => {
        const campaignName = getCampaignById(user.CampaignID)?.Name || '';
        return {
          ID: user.ID,
          UserName: user.UserName || user.FullName,
          FullName: user.FullName || user.UserName,
          Email: user.Email || '',
          CampaignID: user.CampaignID || '',
          campaignName: campaignName,
          EmploymentStatus: user.EmploymentStatus || 'Active',
          HireDate: user.HireDate || '',
          TerminationDate: user.TerminationDate || user.terminationDate || '',
          isActive: isUserConsideredActive(user)
        };
      });

    console.log(`✅ Returning ${scheduleUsers.length} schedule users`);
    return scheduleUsers;

  } catch (error) {
    console.error('❌ Error getting schedule users:', error);
    safeWriteError('clientGetScheduleUsers', error);
    return [];
  }
}

/**
 * Get users for attendance (all active users)
 */
function clientGetAttendanceUsers(requestingUserId, campaignId = null) {
  try {
    console.log('📋 Getting attendance users');
    
    // Use the existing function but return just names for compatibility
    const scheduleUsers = clientGetScheduleUsers(requestingUserId, campaignId);
    const userNames = scheduleUsers
      .map(user => user.UserName || user.FullName)
      .filter(name => name && name.trim())
      .sort();

    console.log(`✅ Returning ${userNames.length} attendance users`);
    return userNames;

  } catch (error) {
    console.error('❌ Error getting attendance users:', error);
    safeWriteError('clientGetAttendanceUsers', error);
    return [];
  }
}

/**
 * Get managed users list - delegates to MainUtilities
 */
function clientGetManagedUsersList(managerId) {
  try {
    if (!managerId) return [];

    const normalizedManagerId = normalizeUserIdValue(managerId);
    const userLookup = buildScheduleUserLookupIndex();
    const managedUsers = [];
    const seen = new Set();

    const pushUser = (user, campaignInfo = {}) => {
      if (!user || typeof user !== 'object') {
        return;
      }

      const candidateIds = extractUserIdsFromCandidates([user], userLookup);
      const normalizedId = candidateIds.length
        ? normalizeUserIdValue(candidateIds[0])
        : normalizeUserIdValue(user.ID || user.UserID || user.id || user.userId);

      if (!normalizedId || normalizedId === normalizedManagerId || seen.has(normalizedId)) {
        return;
      }

      seen.add(normalizedId);

      const campaignId = normalizeCampaignIdValue(
        campaignInfo.campaignId
          || user.CampaignID
          || user.campaignID
          || user.CampaignId
          || user.campaignId
      );

      let campaignName = campaignInfo.campaignName
        || user.campaignName
        || user.CampaignName
        || user.campaign;

      if (!campaignName && campaignId && typeof getCampaignById === 'function') {
        try {
          const campaignRecord = getCampaignById(campaignId);
          if (campaignRecord) {
            campaignName = campaignRecord.Name || campaignRecord.name || '';
          }
        } catch (campaignError) {
          console.warn('Unable to resolve campaign details for roster entry', campaignId, campaignError);
        }
      }

      managedUsers.push({
        ID: normalizedId,
        UserName: user.UserName || user.Username || user.username || user.FullName || '',
        FullName: user.FullName || user.fullName || user.UserName || user.Username || '',
        Email: user.Email || user.email || '',
        CampaignID: campaignId || '',
        campaignName: campaignName || '',
        EmploymentStatus: user.EmploymentStatus || 'Active'
      });
    };

    let managedCampaigns = [];
    if (typeof getUserManagedCampaigns === 'function') {
      try {
        const rawManaged = getUserManagedCampaigns(normalizedManagerId) || [];
        managedCampaigns = Array.isArray(rawManaged) ? rawManaged : [];
      } catch (campaignError) {
        console.warn('Unable to resolve managed campaigns for roster', normalizedManagerId, campaignError);
      }
    }

    managedCampaigns.forEach(campaign => {
      const campaignId = normalizeCampaignIdValue(
        campaign && (campaign.ID || campaign.Id || campaign.id || campaign.CampaignID || campaign.CampaignId)
      );

      if (!campaignId) {
        return;
      }

      let campaignUsers = [];
      if (typeof getUsersByCampaign === 'function') {
        try {
          campaignUsers = getUsersByCampaign(campaignId) || [];
        } catch (campaignError) {
          console.warn('Unable to read campaign roster for manager', normalizedManagerId, campaignId, campaignError);
        }
      }

      if ((!Array.isArray(campaignUsers) || !campaignUsers.length) && userLookup.users.length) {
        campaignUsers = userLookup.users.filter(user => doesUserBelongToCampaign(user, campaignId));
      }

      const campaignName = campaign && (campaign.Name || campaign.name || '');
      campaignUsers.forEach(user => pushUser(user, { campaignId, campaignName }));
    });

    if (!managedUsers.length) {
      const fallback = collectCampaignUsersForManager(normalizedManagerId, { allUsers: userLookup.users });
      fallback.users.forEach(user => pushUser(user, {
        campaignId: fallback.campaignId,
        campaignName: fallback.campaignName
      }));
    }

    return managedUsers;

  } catch (error) {
    console.error('Error getting managed users:', error);
    safeWriteError('clientGetManagedUsersList', error);
    return [];
  }
}

// ────────────────────────────────────────────────────────────────────────────
// SHIFT SLOTS MANAGEMENT - Uses ScheduleUtilities
// ────────────────────────────────────────────────────────────────────────────

/**
 * Create shift slot with proper validation - uses ScheduleUtilities
 */

function clientCreateShiftSlot(slotData) {
  try {
    console.log('🕒 Creating shift slot:', slotData);

    if (!slotData || !slotData.name || !slotData.startTime || !slotData.endTime) {
      return {
        success: false,
        error: 'Slot name, start time, and end time are required'
      };
    }

    const validation = validateShiftSlot(slotData);
    if (!validation.isValid) {
      return {
        success: false,
        error: validation.errors.join('; ')
      };
    }

    const sheet = ensureScheduleSheetWithHeaders(SHIFT_SLOTS_SHEET, SHIFT_SLOTS_HEADERS);
    const existingSlots = readScheduleSheet(SHIFT_SLOTS_SHEET) || [];

    const actor = typeof getCurrentUser === 'function' ? getCurrentUser() : null;
    const actorName = actor && (actor.Email || actor.UserName || actor.name) || 'System';
    const now = new Date();

    const slotId = Utilities.getUuid();
    const normalizedDays = normalizeDaySelection(slotData.daysOfWeek || slotData.DaysOfWeek || slotData.days);
    const daysCsv = normalizedDays.length ? convertDaysToCsv(normalizedDays) : 'Mon,Tue,Wed,Thu,Fri';
    const startTime = normalizeTimeTo12Hour(slotData.startTime);
    const endTime = normalizeTimeTo12Hour(slotData.endTime);

    if (!startTime || !endTime) {
      return {
        success: false,
        error: 'Start and end times must be valid 12-hour values.'
      };
    }

    const campaign = (slotData.campaign || slotData.Campaign || 'General').toString().trim();
    const slotName = slotData.name.toString().trim();

    const duplicate = existingSlots.some(slot => {
      const name = (slot.SlotName || slot.Name || '').toString().trim().toLowerCase();
      const campaignName = (slot.Campaign || slot.Department || '').toString().trim().toLowerCase();
      const statusValue = (slot.Status || '').toString().trim();
      const isActive = statusValue ? statusValue.toUpperCase() !== 'ARCHIVED' : scheduleFlagToBool(slot.IsActive, true);
      return isActive && name === slotName.toLowerCase() && campaignName === campaign.toLowerCase();
    });

    if (duplicate) {
      return {
        success: false,
        error: `A shift slot named "${slotName}" already exists for campaign "${campaign}".`
      };
    }

    const slotRecord = {
      ID: slotId,
      Name: slotName,
      StartTime: startTime,
      EndTime: endTime,
      DaysOfWeek: daysCsv,
      Department: campaign,
      Location: slotData.location || slotData.Location || 'Office',
      Description: slotData.description || '',
      CreatedBy: actorName,
      Notes: slotData.notes || '',
      Status: 'Active',
      CreatedAt: now,
      UpdatedAt: now,
      UpdatedBy: actorName,
      // compatibility aliases
      SlotId: slotId,
      SlotName: slotName,
      Campaign: campaign,
      DaysCSV: daysCsv
    };

    const rowData = SHIFT_SLOTS_HEADERS.map(header => Object.prototype.hasOwnProperty.call(slotRecord, header) ? slotRecord[header] : '');
    sheet.appendRow(rowData);
    SpreadsheetApp.flush();

    appendAuditLogEntry('CREATE', 'ShiftSlot', slotId, null, slotRecord, 'Created shift slot');
    invalidateScheduleCaches();

    console.log('✅ Shift slot created:', slotId);
    return {
      success: true,
      slotId: slotId,
      slot: slotRecord
    };

  } catch (error) {
    console.error('Error creating shift slot:', error);
    safeWriteError('clientCreateShiftSlot', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function buildScheduleUserLookupIndex() {
  const lookup = {
    users: [],
    byId: new Map(),
    byEmail: new Map(),
    byUserName: new Map(),
    byFullName: new Map()
  };

  try {
    const users = readSheet(USERS_SHEET) || [];
    lookup.users = users;

    users.forEach(user => {
      if (!user || typeof user !== 'object') {
        return;
      }

      const normalizedId = normalizeUserIdValue(user.ID || user.UserID || user.id || user.userId);
      const normalizedEmail = (user.Email || user.email || '').toString().trim().toLowerCase();
      const normalizedUserName = (user.UserName || user.Username || user.username || '').toString().trim().toLowerCase();
      const normalizedFullName = (user.FullName || user.fullName || '').toString().trim().toLowerCase();

      if (normalizedId) {
        lookup.byId.set(normalizedId, normalizedId);
      }
      if (normalizedEmail && !lookup.byEmail.has(normalizedEmail)) {
        lookup.byEmail.set(normalizedEmail, normalizedId || normalizedEmail);
      }
      if (normalizedUserName && !lookup.byUserName.has(normalizedUserName)) {
        lookup.byUserName.set(normalizedUserName, normalizedId || normalizedUserName);
      }
      if (normalizedFullName && !lookup.byFullName.has(normalizedFullName)) {
        lookup.byFullName.set(normalizedFullName, normalizedId || normalizedFullName);
      }
    });
  } catch (error) {
    console.warn('Unable to build schedule user lookup index:', error && error.message ? error.message : error);
  }

  return lookup;
}

function resolveUserIdViaLookup(candidate, lookup) {
  if (candidate === null || typeof candidate === 'undefined') {
    return '';
  }

  if (Array.isArray(candidate)) {
    for (let index = 0; index < candidate.length; index++) {
      const resolved = resolveUserIdViaLookup(candidate[index], lookup);
      if (resolved) {
        return resolved;
      }
    }
    return '';
  }

  if (typeof candidate === 'object') {
    const objectCandidates = [
      candidate.ID, candidate.Id, candidate.id,
      candidate.UserID, candidate.UserId, candidate.userId,
      candidate.ManagedUserID, candidate.ManagedUserId, candidate.managedUserId,
      candidate.ManagerID, candidate.ManagerId, candidate.managerId,
      candidate.Email, candidate.email,
      candidate.UserEmail, candidate.userEmail,
      candidate.ManagedEmail, candidate.managedEmail,
      candidate.UserName, candidate.Username, candidate.username,
      candidate.ManagedUserName, candidate.managedUserName, candidate.ManagedUsername, candidate.managedUsername,
      candidate.FullName, candidate.fullName,
      candidate.Name, candidate.name
    ];

    for (let index = 0; index < objectCandidates.length; index++) {
      const resolved = resolveUserIdViaLookup(objectCandidates[index], lookup);
      if (resolved) {
        return resolved;
      }
    }

    return '';
  }

  const raw = String(candidate).trim();
  if (!raw) {
    return '';
  }

  const normalizedId = normalizeUserIdValue(raw);
  if (lookup && lookup.byId && lookup.byId.has(normalizedId)) {
    return lookup.byId.get(normalizedId) || normalizedId;
  }

  const lower = raw.toLowerCase();
  if (lookup && lookup.byEmail && lookup.byEmail.has(lower)) {
    return lookup.byEmail.get(lower) || lower;
  }
  if (lookup && lookup.byUserName && lookup.byUserName.has(lower)) {
    return lookup.byUserName.get(lower) || lower;
  }
  if (lookup && lookup.byFullName && lookup.byFullName.has(lower)) {
    return lookup.byFullName.get(lower) || lower;
  }

  return normalizedId;
}

function extractUserIdsFromCandidates(candidates, lookup) {
  const ids = [];

  const visit = (value) => {
    if (value === null || typeof value === 'undefined') {
      return;
    }

    if (Array.isArray(value)) {
      value.forEach(visit);
      return;
    }

    if (typeof value === 'object') {
      const objectCandidates = [
        value.ID, value.Id, value.id,
        value.UserID, value.UserId, value.userId,
        value.ManagedUserID, value.ManagedUserId, value.managedUserId,
        value.ManagerID, value.ManagerId, value.managerId,
        value.Email, value.email,
        value.UserEmail, value.userEmail,
        value.ManagedEmail, value.managedEmail,
        value.UserName, value.Username, value.username,
        value.ManagedUserName, value.managedUserName, value.ManagedUsername, value.managedUsername,
        value.FullName, value.fullName,
        value.Name, value.name
      ];

      const objectLists = [
        value.Users, value.users,
        value.ManagedUsers, value.managedUsers,
        value.UserIDs, value.UserIds, value.userIds,
        value.ManagedIds, value.managedIds, value.ManagedIDs, value.managedIDs,
        value.TeamMembers, value.teamMembers
      ];

      objectCandidates.forEach(visit);
      objectLists.forEach(visit);
      return;
    }

    const raw = String(value);
    if (/[;,|]/.test(raw)) {
      raw.split(/[;,|]/).forEach(part => visit(part));
      return;
    }

    const resolved = resolveUserIdViaLookup(raw, lookup);
    if (resolved) {
      ids.push(resolved);
    }
  };

  (Array.isArray(candidates) ? candidates : [candidates]).forEach(visit);

  return Array.from(new Set(ids.filter(Boolean)));
}

function collectCampaignUsersForManager(managerId, options = {}) {
  const normalizedManagerId = normalizeUserIdValue(managerId);
  const result = {
    users: [],
    campaignId: '',
    campaignName: ''
  };

  if (!normalizedManagerId) {
    return result;
  }

  const providedUsers = Array.isArray(options.allUsers) ? options.allUsers : null;
  let allUsers = providedUsers || [];

  if (!allUsers.length) {
    try {
      allUsers = readSheet(USERS_SHEET) || [];
    } catch (error) {
      console.warn('Unable to read users for campaign roster fallback:', error && error.message ? error.message : error);
      allUsers = [];
    }
  }

  let managerRecord = null;
  if (allUsers.length) {
    managerRecord = allUsers.find(user => normalizeUserIdValue(user && user.ID) === normalizedManagerId) || null;
  }

  const candidateCampaignIds = [];
  if (managerRecord) {
    candidateCampaignIds.push(
      managerRecord.CampaignID,
      managerRecord.campaignID,
      managerRecord.CampaignId,
      managerRecord.campaignId,
      managerRecord.DefaultCampaignID,
      managerRecord.defaultCampaignId
    );
  }

  if (typeof getUserCampaignsSafe === 'function') {
    try {
      const joinedCampaigns = getUserCampaignsSafe(normalizedManagerId) || [];
      joinedCampaigns.forEach(entry => {
        if (!entry) {
          return;
        }
        candidateCampaignIds.push(
          entry.campaignId,
          entry.CampaignId,
          entry.campaignID,
          entry.CampaignID,
          entry.id,
          entry.Id,
          entry.ID
        );
      });
    } catch (error) {
      console.warn('Unable to resolve campaign membership for manager', normalizedManagerId, error);
    }
  }

  let resolvedCampaignId = '';
  for (let index = 0; index < candidateCampaignIds.length; index++) {
    const normalized = normalizeCampaignIdValue(candidateCampaignIds[index]);
    if (normalized) {
      resolvedCampaignId = normalized;
      break;
    }
  }

  if (!resolvedCampaignId) {
    return result;
  }

  let campaignUsers = [];
  if (typeof getUsersByCampaign === 'function') {
    try {
      campaignUsers = getUsersByCampaign(resolvedCampaignId) || [];
    } catch (error) {
      console.warn('Unable to read campaign users for roster fallback', resolvedCampaignId, error);
    }
  }

  if ((!Array.isArray(campaignUsers) || !campaignUsers.length) && allUsers.length) {
    campaignUsers = allUsers.filter(user => doesUserBelongToCampaign(user, resolvedCampaignId));
  }

  const campaignRecord = typeof getCampaignById === 'function'
    ? getCampaignById(resolvedCampaignId)
    : null;

  result.users = Array.isArray(campaignUsers) ? campaignUsers.filter(Boolean) : [];
  result.campaignId = resolvedCampaignId;
  result.campaignName = campaignRecord ? (campaignRecord.Name || campaignRecord.name || '') : '';

  return result;
}

function getDirectManagedUserIds(managerId) {
  const normalizedManagerId = normalizeUserIdValue(managerId);
  const managedUsers = new Set();

  if (!normalizedManagerId) {
    return managedUsers;
  }

  const userLookup = buildScheduleUserLookupIndex();

  const appendFromRows = (rows) => {
    if (!Array.isArray(rows)) {
      return;
    }

    rows.forEach(row => {
      if (!row || typeof row !== 'object') {
        return;
      }

      const managerCandidates = extractUserIdsFromCandidates([
        row.ManagerUserID, row.ManagerUserId, row.managerUserId,
        row.ManagerID, row.ManagerId, row.managerId, row.manager_id,
        row.UserManagerID, row.UserManagerId, row.userManagerId,
        row.ManagerEmail, row.managerEmail, row.ManagerEmailAddress, row.managerEmailAddress,
        row.ManagerUserName, row.managerUserName, row.ManagerUsername, row.managerUsername,
        row.ManagerName, row.managerName,
        row.Manager, row.manager,
        row.SupervisorID, row.SupervisorId, row.supervisorId,
        row.SupervisorEmail, row.supervisorEmail
      ], userLookup);

      const managedCandidates = extractUserIdsFromCandidates([
        row.UserID, row.UserId, row.userId,
        row.ManagedUserID, row.ManagedUserId, row.managedUserId,
        row.ManagedUserID, row.managed_user_id,
        row.ManagedID, row.ManagedId, row.managedId,
        row.ManagedUsers, row.managedUsers,
        row.UserEmail, row.userEmail, row.Email, row.email,
        row.ManagedEmail, row.managedEmail, row.ManagedEmailAddress, row.managedEmailAddress,
        row.UserName, row.Username, row.username,
        row.ManagedUserName, row.managedUserName, row.ManagedUsername, row.managedUsername,
        row.ManagedName, row.managedName,
        row.Name, row.name,
        row.TeamMemberID, row.TeamMemberId, row.teamMemberId,
        row.TeamMembers, row.teamMembers,
        row.AgentID, row.AgentId, row.agentId,
        row.AgentEmail, row.agentEmail,
        row.AgentName, row.agentName
      ], userLookup);

      const managerMatch = managerCandidates.find(candidate => candidate === normalizedManagerId);

      if (managerMatch && managedCandidates.length) {
        managedCandidates.forEach(candidate => {
          if (candidate && candidate !== normalizedManagerId) {
            managedUsers.add(candidate);
          }
        });
      }

      // Some datasets may store the relationship reversed
      const reversedManager = managedCandidates.find(candidate => candidate === normalizedManagerId);
      if (reversedManager) {
        managerCandidates.forEach(candidate => {
          if (candidate && candidate !== normalizedManagerId) {
            managedUsers.add(candidate);
          }
        });
      }
    });
  };

  try {
    if (typeof readManagerAssignments_ === 'function') {
      appendFromRows(readManagerAssignments_());
    }
  } catch (error) {
    safeWriteError && safeWriteError('getDirectManagedUserIds.readManagerAssignments', error);
  }

  const candidateSheets = Array.from(new Set([
    typeof getManagerUsersSheetName_ === 'function' ? getManagerUsersSheetName_() : null,
    typeof G !== 'undefined' && G ? G.MANAGER_USERS_SHEET : null,
    typeof USER_MANAGERS_SHEET !== 'undefined' ? USER_MANAGERS_SHEET : null,
    'MANAGER_USERS',
    'ManagerUsers',
    'manager_users',
    'UserManagers'
  ].filter(Boolean)));

  candidateSheets.forEach(sheetName => {
    try {
      appendFromRows(readSheet(sheetName));
    } catch (error) {
      console.warn(`Unable to read manager assignments from ${sheetName}:`, error && error.message ? error.message : error);
    }
  });

  let hasManagedUsers = false;
  managedUsers.forEach(id => {
    if (id && id !== normalizedManagerId) {
      hasManagedUsers = true;
    }
  });

  if (!hasManagedUsers) {
    const fallback = collectCampaignUsersForManager(normalizedManagerId, { allUsers: userLookup.users });
    const fallbackIds = extractUserIdsFromCandidates(fallback.users, userLookup);
    fallbackIds.forEach(id => {
      if (id && id !== normalizedManagerId) {
        managedUsers.add(id);
      }
    });
  }

  return managedUsers;
}

function clientCreateEnhancedShiftSlot(slotData) {
  return clientCreateShiftSlot(slotData);
}

function buildManagedUserSet(managerId) {
  const managedUserIds = getDirectManagedUserIds(managerId);
  const normalizedManagerId = normalizeUserIdValue(managerId);

  if (normalizedManagerId) {
    managedUserIds.add(normalizedManagerId);
  }

  try {
    if (typeof getUserManagedCampaigns === 'function' && typeof getUsersByCampaign === 'function') {
      const campaigns = getUserManagedCampaigns(normalizedManagerId) || [];
      campaigns.forEach(campaign => {
        try {
          const campaignUsers = getUsersByCampaign(campaign.ID) || [];
          campaignUsers.forEach(user => {
            const normalizedId = normalizeUserIdValue(user.ID);
            if (normalizedId) {
              managedUserIds.add(normalizedId);
            }
          });
        } catch (campaignErr) {
          console.warn('Failed to append campaign users for campaign', campaign && campaign.ID, campaignErr);
        }
      });
    }
  } catch (error) {
    console.warn('Unable to expand managed users via campaigns:', error);
  }

  let hasManagedUsers = false;
  managedUserIds.forEach(id => {
    if (id && id !== normalizedManagerId) {
      hasManagedUsers = true;
    }
  });

  if (!hasManagedUsers) {
    try {
      const fallback = collectCampaignUsersForManager(normalizedManagerId);
      const fallbackLookup = buildScheduleUserLookupIndex();
      const fallbackIds = extractUserIdsFromCandidates(fallback.users, fallbackLookup);
      fallbackIds.forEach(id => {
        if (id && id !== normalizedManagerId) {
          managedUserIds.add(id);
        }
      });
    } catch (fallbackError) {
      console.warn('Unable to expand managed users via fallback campaign roster:', fallbackError);
    }
  }

  return managedUserIds;
}

function isUserConsideredActive(user) {
  if (!user) {
    return false;
  }

  const status = typeof user.EmploymentStatus === 'string'
    ? user.EmploymentStatus.trim().toLowerCase()
    : '';

  if (!status) {
    return true;
  }

  if (['active', 'activated'].includes(status)) {
    return true;
  }

  if (['terminated', 'inactive', 'disabled', 'separated'].includes(status)) {
    return false;
  }

  return true;
}

/**
 * Get all shift slots - uses ScheduleUtilities
 */

function clientGetAllShiftSlots() {
  try {
    console.log('📊 Getting all shift slots');

    let slots = readScheduleSheet(SHIFT_SLOTS_SHEET) || [];
    if (!slots.length) {
      createDefaultShiftSlots();
      slots = readScheduleSheet(SHIFT_SLOTS_SHEET) || [];
    }

    const normalizedSlots = slots.map(slot => {
      const slotId = (slot.SlotId || slot.ID || slot.Id || slot.slotId || '').toString().trim() || Utilities.getUuid();
      const slotName = (slot.SlotName || slot.Name || '').toString().trim();
      const campaign = (slot.Campaign || slot.Department || '').toString().trim();
      const location = (slot.Location || '').toString().trim() || 'Office';
      const startTime = normalizeTimeTo12Hour(slot.StartTime || slot.startTime || '');
      const endTime = normalizeTimeTo12Hour(slot.EndTime || slot.endTime || '');
      const daysArray = parseDaysCsv(slot.DaysCSV || slot.DaysOfWeek || '');
      const statusValue = (slot.Status || '').toString().trim().toUpperCase();
      const status = statusValue || (scheduleFlagToBool(slot.IsActive, true) ? 'Active' : 'Archived');

      return {
        ID: slotId,
        SlotId: slotId,
        Name: slotName,
        SlotName: slotName,
        Campaign: campaign,
        Department: campaign,
        Location: location,
        StartTime: startTime,
        EndTime: endTime,
        DaysOfWeekArray: daysArray,
        DaysOfWeek: daysArray.join(','),
        Description: slot.Description || '',
        Notes: slot.Notes || '',
        Status: status,
        CreatedAt: slot.CreatedAt || '',
        CreatedBy: slot.CreatedBy || '',
        UpdatedAt: slot.UpdatedAt || '',
        UpdatedBy: slot.UpdatedBy || ''
      };
    });

    normalizedSlots.sort((a, b) => {
      const campaignCompare = (a.Campaign || '').localeCompare(b.Campaign || '');
      if (campaignCompare !== 0) {
        return campaignCompare;
      }
      return (a.SlotName || '').localeCompare(b.SlotName || '');
    });

    console.log(`✅ Returning ${normalizedSlots.length} normalized shift slots`);
    return normalizedSlots;

  } catch (error) {
    console.error('❌ Error getting shift slots:', error);
    safeWriteError('clientGetAllShiftSlots', error);
    return [];
  }
}

// ────────────────────────────────────────────────────────────────────────────
// SCHEDULE GENERATION - Enhanced with ScheduleUtilities integration
// ────────────────────────────────────────────────────────────────────────────

function normalizeGenerationOptions(options = {}) {
  const coerceNumber = (value, fallback, { min = null, max = null } = {}) => {
    const fallbackNumber = Number(fallback);
    let resolved = Number(value);

    if (!Number.isFinite(resolved)) {
      resolved = Number.isFinite(fallbackNumber) ? fallbackNumber : 0;
    }

    if (typeof min === 'number' && resolved < min) {
      resolved = min;
    }

    if (typeof max === 'number' && resolved > max) {
      resolved = max;
    }

    return resolved;
  };

  const coerceBoolean = (value, fallback = false) => {
    if (typeof value === 'boolean') {
      return value;
    }

    if (typeof value === 'number') {
      return value !== 0;
    }

    if (typeof value === 'string') {
      const normalized = value.trim().toLowerCase();
      if (!normalized) {
        return fallback;
      }

      if (['true', '1', 'yes', 'y', 'on'].includes(normalized)) {
        return true;
      }

      if (['false', '0', 'no', 'n', 'off'].includes(normalized)) {
        return false;
      }
    }

    return fallback;
  };

  const capacityOptions = options && typeof options === 'object' ? options.capacity || {} : {};
  const breaksOptions = options && typeof options === 'object' ? options.breaks || {} : {};
  const overtimeOptions = options && typeof options === 'object' ? options.overtime || {} : {};
  const advancedOptions = options && typeof options === 'object' ? options.advanced || {} : {};

  const maxCapacity = coerceNumber(capacityOptions.max ?? options.maxCapacity, 10, { min: 1 });
  const minCoverage = coerceNumber(capacityOptions.min ?? options.minCoverage, 3, { min: 0 });

  const break1Duration = coerceNumber(breaksOptions.first ?? options.break1Duration ?? options.breakDuration, 15, { min: 0 });
  const lunchDuration = coerceNumber(breaksOptions.lunch ?? options.lunchDuration, 30, { min: 0 });
  const break2Duration = coerceNumber(breaksOptions.second ?? options.break2Duration, 15, { min: 0 });
  const enableStaggered = coerceBoolean(breaksOptions.enableStaggered ?? options.enableStaggeredBreaks, true);
  const breakGroups = coerceNumber(breaksOptions.groups ?? options.breakGroups, 3, { min: 1 });
  const staggerInterval = coerceNumber(breaksOptions.interval ?? options.staggerInterval, 15, { min: 1 });
  const minCoveragePct = coerceNumber(breaksOptions.minCoveragePct ?? options.minCoveragePct, 70, { min: 0, max: 100 });

  const overtimeEnabled = coerceBoolean(overtimeOptions.enabled ?? options.overtimeEnabled, false);
  const maxDailyOT = coerceNumber(overtimeOptions.maxDaily ?? options.maxDailyOT, overtimeEnabled ? 2 : 0, { min: 0 });
  const maxWeeklyOT = coerceNumber(overtimeOptions.maxWeekly ?? options.maxWeeklyOT, overtimeEnabled ? 10 : 0, { min: 0 });
  const otApproval = (overtimeOptions.approval ?? options.otApproval ?? 'supervisor') || 'supervisor';
  const otRate = coerceNumber(overtimeOptions.rate ?? options.otRate, 1.5, { min: 1 });
  const otPolicy = (overtimeOptions.policy ?? options.otPolicy ?? 'MANDATORY') || 'MANDATORY';

  const allowSwaps = coerceBoolean(advancedOptions.allowSwaps ?? options.allowSwaps, true);
  const weekendPremium = coerceBoolean(advancedOptions.weekendPremium ?? options.weekendPremium, false);
  const holidayPremium = coerceBoolean(advancedOptions.holidayPremium ?? options.holidayPremium, true);
  const autoAssignment = coerceBoolean(advancedOptions.autoAssignment ?? options.autoAssignment, false);
  const restPeriod = coerceNumber(advancedOptions.restPeriod ?? options.restPeriod, 8, { min: 0 });
  const notificationLead = coerceNumber(advancedOptions.notificationLead ?? options.notificationLead, 24, { min: 0 });
  const handoverTime = coerceNumber(advancedOptions.handoverTime ?? options.handoverTime, 15, { min: 0 });

  const normalized = {
    capacity: {
      max: maxCapacity,
      min: minCoverage
    },
    breaks: {
      first: break1Duration,
      lunch: lunchDuration,
      second: break2Duration,
      enableStaggered,
      groups: breakGroups,
      interval: staggerInterval,
      minCoveragePct
    },
    overtime: {
      enabled: overtimeEnabled,
      maxDaily: maxDailyOT,
      maxWeekly: maxWeeklyOT,
      approval: otApproval,
      rate: otRate,
      policy: otPolicy
    },
    advanced: {
      allowSwaps,
      weekendPremium,
      holidayPremium,
      autoAssignment,
      restPeriod,
      notificationLead,
      handoverTime
    }
  };

  normalized.snapshot = {
    capacity: Object.assign({}, normalized.capacity),
    breaks: Object.assign({}, normalized.breaks),
    overtime: Object.assign({}, normalized.overtime),
    advanced: Object.assign({}, normalized.advanced)
  };

  return normalized;
}

function applyGenerationOptionsToSlot(slot, generationOptions) {
  const applied = Object.assign({}, slot || {});
  const { capacity, breaks, overtime, advanced } = generationOptions || {};

  if (capacity) {
    if (typeof capacity.max === 'number') {
      applied.MaxCapacity = capacity.max;
    }
    if (typeof capacity.min === 'number') {
      applied.MinCoverage = capacity.min;
    }
  }

  if (breaks) {
    if (typeof breaks.first === 'number') {
      applied.BreakDuration = breaks.first;
      applied.Break1Duration = breaks.first;
    }
    if (typeof breaks.second === 'number') {
      applied.Break2Duration = breaks.second;
    }
    if (typeof breaks.lunch === 'number') {
      applied.LunchDuration = breaks.lunch;
    }
    if (typeof breaks.enableStaggered === 'boolean') {
      applied.EnableStaggeredBreaks = breaks.enableStaggered;
    }
    if (typeof breaks.groups === 'number') {
      applied.BreakGroups = breaks.groups;
    }
    if (typeof breaks.interval === 'number') {
      applied.StaggerInterval = breaks.interval;
    }
    if (typeof breaks.minCoveragePct === 'number') {
      applied.MinCoveragePct = breaks.minCoveragePct;
    }
  }

  if (overtime) {
    if (typeof overtime.enabled === 'boolean') {
      applied.EnableOvertime = overtime.enabled;
    }
    if (typeof overtime.maxDaily === 'number') {
      applied.MaxDailyOT = overtime.maxDaily;
    }
    if (typeof overtime.maxWeekly === 'number') {
      applied.MaxWeeklyOT = overtime.maxWeekly;
    }
    if (overtime.approval) {
      applied.OTApproval = overtime.approval;
    }
    if (typeof overtime.rate === 'number') {
      applied.OTRate = overtime.rate;
    }
    if (overtime.policy) {
      applied.OTPolicy = overtime.policy;
    }
  }

  if (advanced) {
    if (typeof advanced.allowSwaps === 'boolean') {
      applied.AllowSwaps = advanced.allowSwaps;
    }
    if (typeof advanced.weekendPremium === 'boolean') {
      applied.WeekendPremium = advanced.weekendPremium;
    }
    if (typeof advanced.holidayPremium === 'boolean') {
      applied.HolidayPremium = advanced.holidayPremium;
    }
    if (typeof advanced.autoAssignment === 'boolean') {
      applied.AutoAssignment = advanced.autoAssignment;
    }
    if (typeof advanced.restPeriod === 'number') {
      applied.RestPeriod = advanced.restPeriod;
    }
    if (typeof advanced.notificationLead === 'number') {
      applied.NotificationLead = advanced.notificationLead;
    }
    if (typeof advanced.handoverTime === 'number') {
      applied.HandoverTime = advanced.handoverTime;
    }
  }

  return applied;
}

function parseDateTimeForGeneration(dateStr, timeStr) {
  if (!dateStr) {
    return null;
  }

  try {
    const base = new Date(dateStr);
    if (isNaN(base.getTime())) {
      return null;
    }

    if (!timeStr) {
      return base;
    }

    const parts = String(timeStr).split(':');
    const hours = Number(parts[0]);
    const minutes = Number(parts[1]);
    const seconds = Number(parts[2] || 0);

    if (!Number.isFinite(hours) || !Number.isFinite(minutes) || !Number.isFinite(seconds)) {
      return base;
    }

    const dateTime = new Date(base.getTime());
    dateTime.setHours(hours, minutes, seconds, 0);
    return dateTime;
  } catch (error) {
    return null;
  }
}

function buildSlotAssignmentKey(slotId, dateStr) {
  if (!dateStr) {
    return null;
  }

  return `${slotId || 'UNASSIGNED'}::${dateStr}`;
}

function determineCapacityLimit(slot, generationOptions) {
  const slotCapacity = Number(slot && slot.MaxCapacity);
  const generationCapacity = generationOptions && generationOptions.capacity ? Number(generationOptions.capacity.max) : NaN;

  if (Number.isFinite(generationCapacity) && generationCapacity > 0) {
    if (Number.isFinite(slotCapacity) && slotCapacity > 0) {
      return Math.min(generationCapacity, slotCapacity);
    }
    return generationCapacity;
  }

  if (Number.isFinite(slotCapacity) && slotCapacity > 0) {
    return slotCapacity;
  }

  return null;
}

/**
 * Enhanced schedule generation with comprehensive validation
 */
/**
 * Enhanced schedule generation with comprehensive validation
 */

function clientGenerateSchedulesEnhanced(startDate, endDate, userNames, shiftSlotIds, templateId, generatedBy, options = {}) {
  try {
    const normalizedStart = normalizeDateForSheet(startDate, DEFAULT_SCHEDULE_TIME_ZONE);
    const normalizedEnd = normalizeDateForSheet(endDate, DEFAULT_SCHEDULE_TIME_ZONE);

    if (!normalizedStart || !normalizedEnd) {
      return {
        success: false,
        error: 'Start and end dates are required for schedule generation.'
      };
    }

    const startDateObj = new Date(normalizedStart);
    const endDateObj = new Date(normalizedEnd);
    if (startDateObj > endDateObj) {
      return {
        success: false,
        error: 'End date must be on or after the start date.'
      };
    }

    const campaignId = normalizeCampaignIdValue(options.campaignId || '');
    const detectConflicts = options.detectConflicts !== false;
    const includeHolidays = options.includeHolidays !== false;
    const advancedOptions = options.advanced || {};
    const capacityOptions = options.capacity || {};
    const breaksOptions = options.breaks || {};
    const overtimeOptions = options.overtime || {};
    const allowSwaps = scheduleFlagToBool(advancedOptions.allowSwaps, true);
    const restHours = Number(advancedOptions.restPeriod || 0);
    const notificationLead = Number(advancedOptions.notificationLead || 0);
    const handoverMinutes = Number(advancedOptions.handoverTime || 0);
    const overtimeEnabled = scheduleFlagToBool(overtimeOptions.enabled, false);
    const overtimeMinutes = overtimeEnabled ? Math.round(Number(overtimeOptions.maxDaily || 0) * 60) : '';
    const maxCapacity = Number(capacityOptions.max || options.maxCapacity || 0) || null;
    const minCoverage = Number(capacityOptions.min || options.minCoverage || 0) || 0;
    const minCoveragePct = Number(breaksOptions.minCoveragePct || options.minCoveragePct || 0) || 0;

    let selectedSlots = clientGetAllShiftSlots();
    selectedSlots = selectedSlots.filter(slot => (slot.Status || 'Active').toUpperCase() !== 'ARCHIVED');
    if (campaignId) {
      selectedSlots = selectedSlots.filter(slot => (slot.Campaign || '').toString().toLowerCase() === campaignId.toLowerCase());
    }
    if (Array.isArray(shiftSlotIds) && shiftSlotIds.length) {
      selectedSlots = selectedSlots.filter(slot => shiftSlotIds.includes(slot.SlotId));
    }

    if (!selectedSlots.length) {
      return {
        success: false,
        error: 'No active shift slots matched the selection for this campaign.'
      };
    }

    const slotMap = new Map(selectedSlots.map(slot => [slot.SlotId, slot]));
    const scheduleUsers = clientGetScheduleUsers(generatedBy || 'system', campaignId || null);
    const userKeyMap = new Map();
    const userIdMap = new Map();
    scheduleUsers.forEach(user => {
      userKeyMap.set(normalizeUserKey(user.UserName || user.FullName), user);
      userIdMap.set(String(user.ID), user);
    });

    let targetUsers = [];
    const unresolvedUsers = [];
    if (Array.isArray(userNames) && userNames.length) {
      userNames.forEach(entry => {
        if (!entry) {
          return;
        }
        const nameKey = normalizeUserKey(entry);
        const idKey = String(entry);
        const user = userKeyMap.get(nameKey) || userIdMap.get(idKey);
        if (user) {
          targetUsers.push(user);
        } else {
          unresolvedUsers.push(entry);
        }
      });
    } else {
      targetUsers = scheduleUsers.slice();
    }

    const filteredUsers = targetUsers.filter(user => {
      if (!user || !user.ID) {
        return false;
      }
      if (user.isActive === false) {
        return false;
      }
      if (campaignId && (user.CampaignID || '').toString().toLowerCase() !== campaignId.toLowerCase()) {
        return false;
      }
      if (user.HireDate) {
        const hireDate = new Date(user.HireDate);
        if (!isNaN(hireDate.getTime()) && hireDate > endDateObj) {
          return false;
        }
      }
      return true;
    });

    if (!filteredUsers.length) {
      return {
        success: false,
        error: 'No eligible users were found for the selected campaign and date range.'
      };
    }

    const seed = options.seed || `${campaignId || 'ALL'}-${normalizedStart}-${normalizedEnd}-${(shiftSlotIds || []).join('|')}`;
    const orderedUsers = shuffleWithSeed(filteredUsers, seed);
    const slotCounts = new Map();
    const assignments = [];
    const skippedUsers = [];
    const now = new Date();
    const actor = generatedBy || (typeof getCurrentUser === 'function' ? (getCurrentUser()?.Email || 'System') : 'System');

    orderedUsers.forEach((user, index) => {
      let assignedSlot = null;
      for (let attempt = 0; attempt < selectedSlots.length; attempt++) {
        const slot = selectedSlots[(index + attempt) % selectedSlots.length];
        const slotCount = slotCounts.get(slot.SlotId) || 0;
        if (maxCapacity && slotCount >= maxCapacity) {
          continue;
        }
        assignedSlot = slot;
        break;
      }

      if (!assignedSlot) {
        skippedUsers.push({
          userId: user.ID,
          userName: user.UserName || user.FullName,
          reason: 'Max capacity reached for selected slots'
        });
        return;
      }

      slotCounts.set(assignedSlot.SlotId, (slotCounts.get(assignedSlot.SlotId) || 0) + 1);

      assignments.push({
        AssignmentId: Utilities.getUuid(),
        UserId: user.ID,
        UserName: user.UserName || user.FullName,
        Campaign: campaignId || user.CampaignID || '',
        SlotId: assignedSlot.SlotId,
        SlotName: assignedSlot.SlotName || assignedSlot.Name,
        StartDate: normalizedStart,
        EndDate: normalizedEnd,
        Status: 'PENDING',
        AllowSwap: allowSwaps,
        Premiums: '',
        BreaksConfigJSON: JSON.stringify({
          break1: breaksOptions.first || 15,
          break2: breaksOptions.second || 0,
          lunch: breaksOptions.lunch || 30,
          enableStaggered: scheduleFlagToBool(breaksOptions.enableStaggered, false),
          groups: breaksOptions.groups || '',
          interval: breaksOptions.interval || '',
          minCoveragePct: breaksOptions.minCoveragePct || '',
          unproductive: (breaksOptions.first || 0) + (breaksOptions.second || 0) + (breaksOptions.lunch || 0)
        }),
        OvertimeMinutes: overtimeMinutes || '',
        RestPeriodHours: restHours || '',
        NotificationLeadHours: notificationLead || '',
        HandoverMinutes: handoverMinutes || '',
        Notes: options.notes || '',
        CreatedAt: now,
        CreatedBy: actor,
        UpdatedAt: now,
        UpdatedBy: actor
      });
    });

    const dateSeries = buildDateSeries(normalizedStart, normalizedEnd);
    const holidayMap = includeHolidays ? loadHolidayMap(normalizedStart, normalizedEnd) : new Map();

    const existingAssignments = readShiftAssignments()
      .map(normalizeAssignmentRecord)
      .filter(record => record && record.AssignmentId)
      .filter(record => (record.Status || '').toUpperCase() !== 'ARCHIVED' && (record.Status || '').toUpperCase() !== 'REJECTED');

    const relevantExisting = existingAssignments.filter(record => {
      if (campaignId && (record.Campaign || '').toString().toLowerCase() !== campaignId.toLowerCase()) {
        return false;
      }
      return !(record.EndDate < normalizedStart || record.StartDate > normalizedEnd);
    });

    const conflicts = [];
    const assignmentPremiums = new Map();

    const checkRestPeriod = (existing, generatedSlot, assignment) => {
      if (!restHours || !generatedSlot) {
        return false;
      }
      const candidateSlot = slotMap.get(existing.SlotId);
      if (!candidateSlot) {
        return false;
      }
      const existingStart = new Date(`${existing.StartDate}T00:00:00`);
      const existingEnd = new Date(`${existing.EndDate}T00:00:00`);
      const generatedStart = new Date(`${assignment.StartDate}T00:00:00`);
      const generatedEnd = new Date(`${assignment.EndDate}T00:00:00`);
      const existingStartMinutes = parseTimeToMinutes(candidateSlot.StartTime || candidateSlot.startTime || '');
      const existingEndMinutes = parseTimeToMinutes(candidateSlot.EndTime || candidateSlot.endTime || '');
      const generatedStartMinutes = parseTimeToMinutes(generatedSlot.StartTime || generatedSlot.startTime || '');
      const generatedEndMinutes = parseTimeToMinutes(generatedSlot.EndTime || generatedSlot.endTime || '');

      if (Number.isFinite(existingEndMinutes)) {
        existingEnd.setHours(0, existingEndMinutes, 0, 0);
        if (Number.isFinite(existingStartMinutes) && existingEndMinutes <= existingStartMinutes) {
          existingEnd.setDate(existingEnd.getDate() + 1);
        }
      }

      if (Number.isFinite(generatedStartMinutes)) {
        generatedStart.setHours(0, generatedStartMinutes, 0, 0);
      }
      if (Number.isFinite(generatedEndMinutes)) {
        generatedEnd.setHours(0, generatedEndMinutes, 0, 0);
        if (Number.isFinite(generatedStartMinutes) && generatedEndMinutes <= generatedStartMinutes) {
          generatedEnd.setDate(generatedEnd.getDate() + 1);
        }
      }

      const diffHours = (generatedStart.getTime() - existingEnd.getTime()) / (1000 * 60 * 60);
      return diffHours < restHours;
    };

    const normalizedAssignments = assignments.filter(assignment => {
      const slot = slotMap.get(assignment.SlotId);
      if (!slot) {
        conflicts.push({
          userId: assignment.UserId,
          userName: assignment.UserName,
          type: 'MISSING_SLOT',
          error: 'Assigned slot could not be found',
          periodStart: assignment.StartDate,
          periodEnd: assignment.EndDate
        });
        return false;
      }

      const existingForUser = relevantExisting.filter(existing => {
        const existingUserKey = normalizeUserKey(existing.UserName || '');
        const assignmentUserKey = normalizeUserKey(assignment.UserName || '');
        const sameUser = existing.UserId && assignment.UserId
          ? String(existing.UserId) === String(assignment.UserId)
          : existingUserKey && existingUserKey === assignmentUserKey;
        if (!sameUser) {
          return false;
        }
        const overlaps = !(existing.EndDate < assignment.StartDate || existing.StartDate > assignment.EndDate);
        if (!overlaps) {
          return false;
        }
        return true;
      });

      if (existingForUser.length) {
        existingForUser.forEach(existing => {
          conflicts.push({
            userId: assignment.UserId,
            userName: assignment.UserName,
            type: 'USER_DOUBLE_BOOKING',
            existingAssignmentId: existing.AssignmentId,
            periodStart: existing.StartDate,
            periodEnd: existing.EndDate,
            error: 'User already has an assignment that overlaps this period'
          });
        });
        if (detectConflicts) {
          return false;
        }
      }

      if (restHours > 0) {
        const restConflict = existingForUser.some(existing => checkRestPeriod(existing, slot, assignment));
        if (restConflict) {
          conflicts.push({
            userId: assignment.UserId,
            userName: assignment.UserName,
            type: 'REST_VIOLATION',
            periodStart: assignment.StartDate,
            periodEnd: assignment.EndDate,
            error: `Rest period requirement of ${restHours} hours would be violated`
          });
          if (detectConflicts) {
            return false;
          }
        }
      }

      const premiumSet = new Set();
      const datesForAssignment = dateSeries.filter(date => date >= assignment.StartDate && date <= assignment.EndDate);
      const hasWeekend = datesForAssignment.some(isWeekendDate);
      if (hasWeekend && scheduleFlagToBool(advancedOptions.weekendPremium, false)) {
        premiumSet.add('Weekend');
      }
      const hasHoliday = datesForAssignment.some(date => {
        const entries = holidayMap.get(date) || [];
        return entries.some(entry => (entry.region || '').toLowerCase() === 'jamaica');
      });
      if (hasHoliday && scheduleFlagToBool(advancedOptions.holidayPremium, true)) {
        premiumSet.add('Holiday');
      }
      if (overtimeEnabled) {
        premiumSet.add('Overtime');
      }
      assignmentPremiums.set(assignment.AssignmentId, Array.from(premiumSet));
      assignment.Premiums = Array.from(premiumSet).join(',');
      return true;
    });

    const coverageDetails = dateSeries.map(date => {
      let total = 0;
      const breakdown = {};
      normalizedAssignments.forEach(assignment => {
        if (assignment.StartDate <= date && assignment.EndDate >= date) {
          total += 1;
          breakdown[assignment.SlotId] = (breakdown[assignment.SlotId] || 0) + 1;
        }
      });

      let target = minCoverage;
      if (minCoveragePct > 0) {
        const base = maxCapacity || normalizedAssignments.length || selectedSlots.length;
        const pctTarget = Math.ceil(base * (minCoveragePct / 100));
        target = Math.max(target, pctTarget);
      }

      const holidayEntries = holidayMap.get(date) || [];
      const weekend = isWeekendDate(date);
      return {
        date,
        total,
        minRequired: target,
        shortfall: target > total ? target - total : 0,
        excess: target && total > target ? total - target : 0,
        weekend,
        holidayRegions: holidayEntries.map(entry => entry.region || ''),
        slotBreakdown: breakdown,
        premium: {
          weekend: weekend && scheduleFlagToBool(advancedOptions.weekendPremium, false),
          holiday: holidayEntries.some(entry => (entry.region || '').toLowerCase() === 'jamaica') && scheduleFlagToBool(advancedOptions.holidayPremium, true)
        }
      };
    });

    const daysWithShortfall = coverageDetails.filter(day => day.shortfall > 0).length;
    const coverageMetDays = coverageDetails.length ? coverageDetails.length - daysWithShortfall : 0;
    const coveragePercent = coverageDetails.length ? Math.round((coverageMetDays / coverageDetails.length) * 100) : 100;

    const previewSummary = {
      periodStart: normalizedStart,
      periodEnd: normalizedEnd,
      totalAssignments: normalizedAssignments.length,
      coverageDetails,
      coveragePercent,
      shortfallDays: daysWithShortfall,
      skippedUsers,
      conflicts,
      unresolvedUsers
    };

    if (options.commitToken) {
      const cached = loadSchedulePreview(options.commitToken);
      if (!cached || !Array.isArray(cached.assignments)) {
        return {
          success: false,
          error: 'Preview token expired or not found. Please regenerate the schedule preview.'
        };
      }
      const commitResult = writeShiftAssignments(cached.assignments, actor, options.notes || 'Auto-assigned schedule generation', 'PENDING');
      CacheService.getScriptCache().put(`schedule_preview_${options.commitToken}`, '', 1);
      return {
        success: true,
        generated: commitResult.count || cached.assignments.length,
        periodStart: cached.metadata?.periodStart || normalizedStart,
        periodEnd: cached.metadata?.periodEnd || normalizedEnd,
        coverage: cached.metadata?.coverage || previewSummary,
        conflicts: cached.metadata?.conflicts || [],
        skipped: cached.metadata?.skippedUsers || []
      };
    }

    const previewToken = storeSchedulePreview({
      assignments: normalizedAssignments,
      metadata: {
        periodStart: normalizedStart,
        periodEnd: normalizedEnd,
        coverage: previewSummary,
        conflicts,
        skippedUsers
      }
    });

    const assignmentSummary = normalizedAssignments.map(assignment => ({
      AssignmentId: assignment.AssignmentId,
      UserId: assignment.UserId,
      UserName: assignment.UserName,
      SlotId: assignment.SlotId,
      SlotName: assignment.SlotName,
      StartDate: assignment.StartDate,
      EndDate: assignment.EndDate,
      Premiums: assignmentPremiums.get(assignment.AssignmentId) || []
    }));

    return {
      success: true,
      previewToken,
      generated: normalizedAssignments.length,
      preview: previewSummary,
      assignments: assignmentSummary,
      conflicts,
      skippedUsers,
      unresolvedUsers
    };

  } catch (error) {
    console.error('❌ Enhanced schedule generation failed:', error);
    safeWriteError('clientGenerateSchedulesEnhanced', error);
    return {
      success: false,
      error: error.message,
      generated: 0,
      conflicts: []
    };
  }
}


function saveSchedulesToSheet(schedules) {
  try {
    if (!Array.isArray(schedules) || !schedules.length) {
      return;
    }

    const actor = 'Import';
    const assignments = schedules.map(schedule => ({
      AssignmentId: schedule.AssignmentId || schedule.ID || Utilities.getUuid(),
      UserId: schedule.UserID || schedule.UserId || '',
      UserName: schedule.UserName || '',
      Campaign: schedule.Campaign || schedule.Department || '',
      SlotId: schedule.SlotID || schedule.SlotId || '',
      SlotName: schedule.SlotName || schedule.Name || '',
      StartDate: normalizeDateForSheet(schedule.PeriodStart || schedule.Date, DEFAULT_SCHEDULE_TIME_ZONE),
      EndDate: normalizeDateForSheet(schedule.PeriodEnd || schedule.Date, DEFAULT_SCHEDULE_TIME_ZONE) || normalizeDateForSheet(schedule.PeriodStart || schedule.Date, DEFAULT_SCHEDULE_TIME_ZONE),
      Status: schedule.Status || 'PENDING',
      AllowSwap: scheduleFlagToBool(schedule.AllowSwaps || schedule.AllowSwap, false),
      Premiums: [
        scheduleFlagToBool(schedule.WeekendPremium, false) ? 'Weekend' : '',
        scheduleFlagToBool(schedule.HolidayPremium, false) ? 'Holiday' : '',
        scheduleFlagToBool(schedule.EnableOvertime || schedule.EnableOT, false) ? 'Overtime' : ''
      ].filter(Boolean).join(','),
      BreaksConfigJSON: schedule.GenerationConfig || schedule.BreaksConfigJSON || '',
      OvertimeMinutes: schedule.MaxDailyOT ? Math.round(Number(schedule.MaxDailyOT) * 60) : '',
      RestPeriodHours: schedule.RestPeriodHours || schedule.RestPeriod || '',
      NotificationLeadHours: schedule.NotificationLeadHours || schedule.NotificationLead || '',
      HandoverMinutes: schedule.HandoverTimeMinutes || schedule.HandoverTime || '',
      Notes: schedule.Notes || '',
      CreatedAt: new Date(),
      CreatedBy: actor,
      UpdatedAt: new Date(),
      UpdatedBy: actor
    }));

    writeShiftAssignments(assignments, actor, 'Legacy schedule import', 'PENDING');

  } catch (error) {
    console.error('Error saving schedules to sheet:', error);
    safeWriteError('saveSchedulesToSheet', error);
    throw error;
  }
}

/**
 * Get all schedules with filtering - uses ScheduleUtilities
 */

function clientGetAllSchedules(filters = {}) {
  try {
    console.log('📋 Getting all assignments with filters:', filters);
    const assignments = readShiftAssignments().map(normalizeAssignmentRecord);
    const slotMap = new Map(clientGetAllShiftSlots().map(slot => [slot.SlotId, slot]));

    let filtered = assignments;

    if (filters.startDate) {
      filtered = filtered.filter(record => !record.EndDate || record.EndDate >= filters.startDate);
    }
    if (filters.endDate) {
      filtered = filtered.filter(record => !record.StartDate || record.StartDate <= filters.endDate);
    }
    if (filters.userId) {
      filtered = filtered.filter(record => String(record.UserId || '') === String(filters.userId));
    }
    if (filters.userName) {
      filtered = filtered.filter(record => (record.UserName || '').toString() === filters.userName);
    }
    if (filters.status) {
      filtered = filtered.filter(record => (record.Status || '').toString().toUpperCase() === filters.status.toUpperCase());
    }
    if (filters.campaign) {
      filtered = filtered.filter(record => (record.Campaign || '').toString().toLowerCase() === filters.campaign.toLowerCase());
    }
    if (filters.slotId) {
      filtered = filtered.filter(record => record.SlotId === filters.slotId);
    }

    const normalized = filtered.map(record => {
      const slot = slotMap.get(record.SlotId) || {};
      return {
        ID: record.AssignmentId,
        AssignmentId: record.AssignmentId,
        UserId: record.UserId,
        UserName: record.UserName,
        SlotId: record.SlotId,
        SlotName: record.SlotName || slot.SlotName || slot.Name || '',
        Campaign: record.Campaign || slot.Campaign || '',
        Location: slot.Location || '',
        StartDate: record.StartDate,
        EndDate: record.EndDate,
        Status: record.Status || 'PENDING',
        AllowSwap: scheduleFlagToBool(record.AllowSwap, false),
        Premiums: record.Premiums || '',
        Notes: record.Notes || '',
        StartTime: slot.StartTime || '',
        EndTime: slot.EndTime || ''
      };
    });

    normalized.sort((a, b) => {
      const startCompare = (b.StartDate || '').localeCompare(a.StartDate || '');
      if (startCompare !== 0) {
        return startCompare;
      }
      return (a.UserName || '').localeCompare(b.UserName || '');
    });

    return {
      success: true,
      schedules: normalized,
      total: normalized.length,
      filters
    };

  } catch (error) {
    console.error('❌ Error getting assignments:', error);
    safeWriteError('clientGetAllSchedules', error);
    return {
      success: false,
      error: error.message,
      schedules: []
    };
  }
}

/**
 * Core schedule import implementation shared by all callers
 */

function internalClientImportSchedules(importRequest = {}) {
  try {
    const schedules = Array.isArray(importRequest.schedules) ? importRequest.schedules : [];
    if (schedules.length === 0) {
      throw new Error('No schedules were provided for import.');
    }

    saveSchedulesToSheet(schedules);

    return {
      success: true,
      imported: schedules.length
    };

  } catch (error) {
    console.error('❌ Error importing schedules:', error);
    safeWriteError('internalClientImportSchedules', error);
    return {
      success: false,
      error: error.message || 'Unknown schedule import error'
    };
  }
}
function clientImportSchedules(importRequest = {}) {
  return internalClientImportSchedules(importRequest);
}

/**
 * Fetch schedule data directly from a Google Sheet link for importing
 */
function clientFetchScheduleSheetData(request = {}) {
  try {
    const options = typeof request === 'string' ? { url: request } : (request || {});
    const sheetUrl = (options.url || options.sheetUrl || '').trim();
    const sheetName = (options.sheetName || options.tabName || '').trim();
    const sheetRange = (options.range || options.sheetRange || '').trim();
    const spreadsheetId = (options.id || options.sheetId || options.spreadsheetId || '').trim();
    const gidValue = options.gid || options.sheetGid || options.sheetNumericId;

    if (!sheetUrl && !spreadsheetId) {
      throw new Error('A Google Sheets link or ID is required to import schedules.');
    }

    let spreadsheet = null;
    if (spreadsheetId) {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    }

    if (!spreadsheet) {
      const candidateUrl = sheetUrl;
      if (candidateUrl) {
        try {
          spreadsheet = SpreadsheetApp.openByUrl(candidateUrl);
        } catch (urlError) {
          const extractedId = extractSpreadsheetId(candidateUrl);
          if (extractedId) {
            spreadsheet = SpreadsheetApp.openById(extractedId);
          } else {
            throw urlError;
          }
        }
      }
    }

    if (!spreadsheet) {
      throw new Error('Unable to open the provided Google Sheets link.');
    }

    let sheet = null;
    if (sheetName) {
      sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        throw new Error(`Could not find a sheet named "${sheetName}" in ${spreadsheet.getName()}.`);
      }
    }

    if (!sheet && gidValue !== undefined && gidValue !== null && gidValue !== '') {
      const numericId = Number(gidValue);
      if (!Number.isNaN(numericId)) {
        sheet = spreadsheet.getSheets().find(tab => tab.getSheetId() === numericId) || null;
      }
    }

    if (!sheet) {
      const sheets = spreadsheet.getSheets();
      if (!sheets || sheets.length === 0) {
        throw new Error('The spreadsheet does not contain any sheets to import.');
      }
      sheet = sheets[0];
    }

    const range = sheetRange ? sheet.getRange(sheetRange) : sheet.getDataRange();
    const values = range.getDisplayValues();

    if (!values || values.length === 0) {
      return {
        success: true,
        rows: [],
        spreadsheetName: spreadsheet.getName(),
        sheetName: sheet.getName(),
        sheetId: sheet.getSheetId(),
        range: range.getA1Notation(),
        rowCount: 0,
        columnCount: 0
      };
    }

    return {
      success: true,
      rows: values,
      spreadsheetName: spreadsheet.getName(),
      sheetName: sheet.getName(),
      sheetId: sheet.getSheetId(),
      range: range.getA1Notation(),
      rowCount: values.length,
      columnCount: values[0] ? values[0].length : 0
    };
  } catch (error) {
    console.error('❌ Error fetching schedule data from Google Sheets:', error);
    safeWriteError('clientFetchScheduleSheetData', error);
    return {
      success: false,
      error: error.message
    };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// ATTENDANCE DASHBOARD WITH AI INSIGHTS - Enhanced
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get comprehensive attendance dashboard data with AI insights
 */
function clientGetAttendanceDashboard(startDate, endDate, campaignId = null) {
  try {
    console.log('📊 Generating attendance dashboard');

    // Use ScheduleUtilities to read attendance data
    const attendanceData = readScheduleSheet(ATTENDANCE_STATUS_SHEET) || [];
    
    // Filter by date range
    const filteredData = attendanceData.filter(record => {
      if (!record.Date) return false;
      const recordDate = new Date(record.Date);
      const start = new Date(startDate);
      const end = new Date(endDate);
      return recordDate >= start && recordDate <= end;
    });

    // Get users for context using our enhanced user functions
    const users = clientGetScheduleUsers('system', campaignId);
    const userMap = new Map(users.map(u => [u.UserName, u]));

    // Calculate metrics
    const metrics = calculateAttendanceMetrics(filteredData);
    const userStats = calculateUserAttendanceStats(filteredData, userMap);
    const trends = calculateAttendanceTrends(filteredData);
    const aiInsights = generateAIInsights(metrics, userStats, trends);

    return {
      success: true,
      dashboard: {
        period: { startDate, endDate },
        totalUsers: users.length,
        totalRecords: filteredData.length,
        metrics: metrics,
        userStats: userStats,
        trends: trends,
        insights: aiInsights,
        generatedAt: new Date().toISOString()
      }
    };

  } catch (error) {
    console.error('Error generating attendance dashboard:', error);
    safeWriteError('clientGetAttendanceDashboard', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Calculate attendance metrics
 */
function calculateAttendanceMetrics(attendanceData) {
  const statusCounts = {
    Present: 0,
    Absent: 0,
    Late: 0,
    'Sick Leave': 0,
    'Bereavement': 0,
    'Vacation': 0,
    'Leave Of Absence': 0,
    'No Call No Show': 0,
    Other: 0
  };

  attendanceData.forEach(record => {
    const status = record.Status || 'Other';
    if (statusCounts.hasOwnProperty(status)) {
      statusCounts[status]++;
    } else {
      statusCounts.Other++;
    }
  });

  const total = Object.values(statusCounts).reduce((sum, count) => sum + count, 0);

  const percentages = {};
  Object.keys(statusCounts).forEach(status => {
    percentages[status] = total > 0 ? Math.round((statusCounts[status] / total) * 100) : 0;
  });

  return {
    counts: statusCounts,
    percentages: percentages,
    total: total,
    attendanceRate: percentages.Present,
    absenceRate: percentages.Absent + percentages['No Call No Show']
  };
}

/**
 * Calculate user-specific attendance statistics
 */
function calculateUserAttendanceStats(attendanceData, userMap) {
  const userStats = {};

  attendanceData.forEach(record => {
    const userName = record.UserName;
    if (!userName) return;

    if (!userStats[userName]) {
      userStats[userName] = {
        userName: userName,
        totalRecords: 0,
        present: 0,
        absent: 0,
        late: 0,
        sick: 0,
        other: 0,
        attendanceRate: 0,
        user: userMap.get(userName) || null
      };
    }

    const stats = userStats[userName];
    stats.totalRecords++;

    switch (record.Status) {
      case 'Present': stats.present++; break;
      case 'Absent': stats.absent++; break;
      case 'Late': stats.late++; break;
      case 'Sick Leave': stats.sick++; break;
      default: stats.other++; break;
    }
  });

  // Calculate attendance rates
  Object.values(userStats).forEach(stats => {
    if (stats.totalRecords > 0) {
      stats.attendanceRate = Math.round((stats.present / stats.totalRecords) * 100);
    }
  });

  // Sort by attendance rate (best first)
  return Object.values(userStats).sort((a, b) => b.attendanceRate - a.attendanceRate);
}

/**
 * Calculate attendance trends using ScheduleUtilities week functions
 */
function calculateAttendanceTrends(attendanceData) {
  const dailyStats = {};
  const weeklyStats = {};

  attendanceData.forEach(record => {
    const date = record.Date;
    if (!date) return;

    const dayKey = date;
    const weekKey = weekStringFromDate(new Date(date)); // Use ScheduleUtilities function

    // Daily stats
    if (!dailyStats[dayKey]) {
      dailyStats[dayKey] = { date: dayKey, present: 0, absent: 0, total: 0 };
    }
    dailyStats[dayKey].total++;
    if (record.Status === 'Present') {
      dailyStats[dayKey].present++;
    } else {
      dailyStats[dayKey].absent++;
    }

    // Weekly stats
    if (!weeklyStats[weekKey]) {
      weeklyStats[weekKey] = { week: weekKey, present: 0, absent: 0, total: 0 };
    }
    weeklyStats[weekKey].total++;
    if (record.Status === 'Present') {
      weeklyStats[weekKey].present++;
    } else {
      weeklyStats[weekKey].absent++;
    }
  });

  // Calculate attendance rates
  Object.values(dailyStats).forEach(stat => {
    stat.attendanceRate = stat.total > 0 ? Math.round((stat.present / stat.total) * 100) : 0;
  });

  Object.values(weeklyStats).forEach(stat => {
    stat.attendanceRate = stat.total > 0 ? Math.round((stat.present / stat.total) * 100) : 0;
  });

  return {
    daily: Object.values(dailyStats).sort((a, b) => a.date.localeCompare(b.date)),
    weekly: Object.values(weeklyStats).sort((a, b) => a.week.localeCompare(b.week))
  };
}

/**
 * Generate AI insights from attendance data
 */
function generateAIInsights(metrics, userStats, trends) {
  const insights = [];

  // Overall attendance insights
  if (metrics.attendanceRate >= 95) {
    insights.push({
      type: 'positive',
      category: 'Overall Performance',
      message: `Excellent attendance rate of ${metrics.attendanceRate}%. The team is highly reliable.`,
      priority: 'low'
    });
  } else if (metrics.attendanceRate >= 85) {
    insights.push({
      type: 'neutral',
      category: 'Overall Performance',
      message: `Good attendance rate of ${metrics.attendanceRate}%. Some room for improvement.`,
      priority: 'medium'
    });
  } else {
    insights.push({
      type: 'warning',
      category: 'Overall Performance',
      message: `Attendance rate of ${metrics.attendanceRate}% is below optimal. Consider implementing attendance improvement strategies.`,
      priority: 'high'
    });
  }

  // User-specific insights
  const topPerformers = userStats.slice(0, 3);
  const poorPerformers = userStats.filter(u => u.attendanceRate < 85).slice(0, 3);

  if (topPerformers.length > 0) {
    insights.push({
      type: 'positive',
      category: 'Top Performers',
      message: `Top attendance: ${topPerformers.map(u => `${u.userName} (${u.attendanceRate}%)`).join(', ')}`,
      priority: 'low'
    });
  }

  if (poorPerformers.length > 0) {
    insights.push({
      type: 'warning',
      category: 'Attendance Concerns',
      message: `Users needing attention: ${poorPerformers.map(u => `${u.userName} (${u.attendanceRate}%)`).join(', ')}`,
      priority: 'high'
    });
  }

  // Health insights
  if (metrics.percentages['Sick Leave'] > 10) {
    insights.push({
      type: 'warning',
      category: 'Health Trends',
      message: `High sick leave rate (${metrics.percentages['Sick Leave']}%). Consider wellness programs or workplace health assessment.`,
      priority: 'medium'
    });
  }

  // Policy compliance insights
  if (metrics.percentages['No Call No Show'] > 2) {
    insights.push({
      type: 'critical',
      category: 'Policy Compliance',
      message: `${metrics.percentages['No Call No Show']}% no call/no show rate requires immediate attention. Review attendance policies.`,
      priority: 'critical'
    });
  }

  // Trend insights
  if (trends.weekly.length >= 2) {
    const recentWeeks = trends.weekly.slice(-2);
    const trend = recentWeeks[1].attendanceRate - recentWeeks[0].attendanceRate;
    
    if (trend > 5) {
      insights.push({
        type: 'positive',
        category: 'Trends',
        message: `Attendance improving! Up ${trend}% from previous week.`,
        priority: 'low'
      });
    } else if (trend < -5) {
      insights.push({
        type: 'warning',
        category: 'Trends',
        message: `Attendance declining. Down ${Math.abs(trend)}% from previous week.`,
        priority: 'high'
      });
    }
  }

  return insights;
}

// ────────────────────────────────────────────────────────────────────────────
// HOLIDAYS MANAGEMENT WITH MULTI-COUNTRY SUPPORT - Enhanced
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get updated holidays for supported countries with Jamaica priority
 */
function getUpdatedHolidays(countryCode, year) {
  const holidayData = {
    'JM': [ // Jamaica - Primary country
      { name: 'New Year\'s Day', date: `${year}-01-01` },
      { name: 'Ash Wednesday', date: `${year}-02-14` },
      { name: 'Good Friday', date: `${year}-04-07` },
      { name: 'Easter Monday', date: `${year}-04-10` },
      { name: 'Labour Day', date: `${year}-05-23` },
      { name: 'Emancipation Day', date: `${year}-08-01` },
      { name: 'Independence Day', date: `${year}-08-06` },
      { name: 'National Heroes Day', date: `${year}-10-16` },
      { name: 'Christmas Day', date: `${year}-12-25` },
      { name: 'Boxing Day', date: `${year}-12-26` }
    ],
    'US': [ // United States
      { name: 'New Year\'s Day', date: `${year}-01-01` },
      { name: 'Martin Luther King Jr. Day', date: `${year}-01-15` },
      { name: 'Presidents\' Day', date: `${year}-02-19` },
      { name: 'Memorial Day', date: `${year}-05-27` },
      { name: 'Independence Day', date: `${year}-07-04` },
      { name: 'Labor Day', date: `${year}-09-02` },
      { name: 'Columbus Day', date: `${year}-10-14` },
      { name: 'Veterans Day', date: `${year}-11-11` },
      { name: 'Thanksgiving Day', date: `${year}-11-28` },
      { name: 'Christmas Day', date: `${year}-12-25` }
    ],
    'DO': [ // Dominican Republic
      { name: 'New Year\'s Day', date: `${year}-01-01` },
      { name: 'Epiphany', date: `${year}-01-06` },
      { name: 'Lady of Altagracia Day', date: `${year}-01-21` },
      { name: 'Juan Pablo Duarte Day', date: `${year}-01-26` },
      { name: 'Independence Day', date: `${year}-02-27` },
      { name: 'Good Friday', date: `${year}-04-07` },
      { name: 'Labour Day', date: `${year}-05-01` },
      { name: 'Corpus Christi', date: `${year}-06-15` },
      { name: 'Restoration Day', date: `${year}-08-16` },
      { name: 'Our Lady of Mercedes Day', date: `${year}-09-24` },
      { name: 'Constitution Day', date: `${year}-11-06` },
      { name: 'Christmas Day', date: `${year}-12-25` }
    ],
    'PH': [ // Philippines
      { name: 'New Year\'s Day', date: `${year}-01-01` },
      { name: 'People Power Anniversary', date: `${year}-02-25` },
      { name: 'Maundy Thursday', date: `${year}-04-06` },
      { name: 'Good Friday', date: `${year}-04-07` },
      { name: 'Araw ng Kagitingan', date: `${year}-04-09` },
      { name: 'Labor Day', date: `${year}-05-01` },
      { name: 'Independence Day', date: `${year}-06-12` },
      { name: 'National Heroes Day', date: `${year}-08-26` },
      { name: 'Bonifacio Day', date: `${year}-11-30` },
      { name: 'Christmas Day', date: `${year}-12-25` },
      { name: 'Rizal Day', date: `${year}-12-30` }
    ]
  };

  return holidayData[countryCode] || [];
}

/**
 * Get holidays for supported countries with Jamaica priority
 */
function clientGetCountryHolidays(countryCode, year) {
  try {
    console.log('🎉 Getting holidays for:', countryCode, year);

    if (!SCHEDULE_SETTINGS.SUPPORTED_COUNTRIES.includes(countryCode)) {
      return {
        success: false,
        error: `Country ${countryCode} not supported. Supported countries: ${SCHEDULE_SETTINGS.SUPPORTED_COUNTRIES.join(', ')}`,
        holidays: []
      };
    }

    const holidays = getUpdatedHolidays(countryCode, year);
    const isPrimary = countryCode === SCHEDULE_SETTINGS.PRIMARY_COUNTRY;

    return {
      success: true,
      holidays: holidays,
      country: countryCode,
      year: year,
      isPrimary: isPrimary,
      note: isPrimary ? 'Primary country (Jamaica) - takes precedence' : 'Secondary country'
    };

  } catch (error) {
    console.error('Error getting country holidays:', error);
    safeWriteError('clientGetCountryHolidays', error);
    return {
      success: false,
      error: error.message,
      holidays: []
    };
  }
}

  function clientAddManualShiftSlots(request = {}) {
  try {
    const timeZone = typeof Session !== 'undefined' ? Session.getScriptTimeZone() : DEFAULT_SCHEDULE_TIME_ZONE;
    const normalizedStart = normalizeDateForSheet(request.startDate || request.date, timeZone);
    const normalizedEnd = normalizeDateForSheet(request.endDate || request.date, timeZone) || normalizedStart;

    if (!normalizedStart || !normalizedEnd) {
      return {
        success: false,
        error: 'Start and end dates are required for manual assignment.'
      };
    }

    if (new Date(normalizedStart) > new Date(normalizedEnd)) {
      return {
        success: false,
        error: 'End date must be on or after the start date.'
      };
    }

    const slotId = request.slotId || request.slot || '';
    if (!slotId) {
      return {
        success: false,
        error: 'Select a shift slot before creating assignments.'
      };
    }

    const availableSlots = clientGetAllShiftSlots();
    const slot = availableSlots.find(slotRecord => {
      if (!slotRecord || typeof slotRecord !== 'object') {
        return false;
      }
      const candidates = [
        slotRecord.SlotId, slotRecord.SlotID, slotRecord.slotId,
        slotRecord.ID, slotRecord.Id, slotRecord.id
      ];
      return candidates.some(candidate => candidate && String(candidate) === String(slotId));
    });
    if (!slot) {
      return {
        success: false,
        error: 'The selected shift slot could not be found.'
      };
    }

    const userEntries = Array.isArray(request.users) ? request.users : [];
    if (!userEntries.length) {
      return {
        success: false,
        error: 'Choose at least one user for manual assignment.'
      };
    }

    const replaceExisting = scheduleFlagToBool(request.replaceExisting, false);
    const campaignId = normalizeCampaignIdValue(request.campaignId || slot.Campaign || '');
    const actor = request.createdBy || (typeof getCurrentUser === 'function' ? (getCurrentUser()?.Email || 'System') : 'System');

    const scheduleUsers = clientGetScheduleUsers(actor, campaignId || null);
    const userKeyMap = new Map();
    const userIdMap = new Map();
    scheduleUsers.forEach(user => {
      userKeyMap.set(normalizeUserKey(user.UserName || user.FullName), user);
      userIdMap.set(String(user.ID), user);
    });

    const existingAssignments = readShiftAssignments()
      .map(normalizeAssignmentRecord)
      .filter(record => record && record.AssignmentId)
      .filter(record => (record.Status || '').toUpperCase() !== 'ARCHIVED');

    const conflicts = [];
    const createdAssignments = [];
    const failedUsers = [];
    const archivedAssignments = [];
    const now = new Date();

    userEntries.forEach(entry => {
      if (!entry) {
        return;
      }
      const nameKey = normalizeUserKey(entry.UserName || entry.FullName || entry.name || entry);
      const idKey = String(entry.ID || entry.id || entry.userId || entry);
      const user = userKeyMap.get(nameKey) || userIdMap.get(idKey);
      if (!user) {
        failedUsers.push({
          entry,
          userId: idKey,
          userName: entry.UserName || entry.FullName || entry.name || '',
          reason: 'User not found in schedule directory'
        });
        return;
      }

      const overlap = existingAssignments.filter(record => {
        const sameUser = record.UserId && user.ID
          ? String(record.UserId) === String(user.ID)
          : normalizeUserKey(record.UserName || '') === normalizeUserKey(user.UserName || user.FullName);
        if (!sameUser) {
          return false;
        }
        return !(record.EndDate < normalizedStart || record.StartDate > normalizedEnd);
      });

      if (overlap.length && !replaceExisting) {
        overlap.forEach(conflict => {
          conflicts.push({
            userId: user.ID,
            userName: user.UserName || user.FullName,
            type: 'USER_DOUBLE_BOOKING',
            existingAssignmentId: conflict.AssignmentId,
            periodStart: conflict.StartDate,
            periodEnd: conflict.EndDate,
            error: 'Existing assignment overlaps the selected range'
          });
        });
      }

      if (replaceExisting && overlap.length) {
        overlap.forEach(conflict => {
          updateShiftAssignmentRow(conflict.AssignmentId, row => {
            row.Status = 'ARCHIVED';
            row.UpdatedAt = now;
            row.UpdatedBy = actor;
            return row;
          });
          archivedAssignments.push(conflict.AssignmentId);
        });
      }

      createdAssignments.push({
        AssignmentId: Utilities.getUuid(),
        UserId: user.ID,
        UserName: user.UserName || user.FullName,
        Campaign: campaignId || user.CampaignID || '',
        SlotId: slot.SlotId,
        SlotName: slot.SlotName || slot.Name,
        StartDate: normalizedStart,
        EndDate: normalizedEnd,
        Status: 'PENDING',
        AllowSwap: scheduleFlagToBool(request.allowSwaps, true),
        Premiums: '',
        BreaksConfigJSON: JSON.stringify({
          break1: 15,
          break2: 15,
          lunch: 30,
          enableStaggered: false,
          groups: '',
          interval: '',
          unproductive: 60
        }),
        OvertimeMinutes: '',
        RestPeriodHours: '',
        NotificationLeadHours: '',
        HandoverMinutes: '',
        Notes: request.notes || '',
        CreatedAt: now,
        CreatedBy: actor,
        UpdatedAt: now,
        UpdatedBy: actor
      });
    });

    if (!createdAssignments.length) {
      return {
        success: false,
        error: conflicts.length ? 'Assignments blocked by existing conflicts.' : 'No assignments were created.',
        conflicts,
        failed: failedUsers
      };
    }

    const writeResult = writeShiftAssignments(createdAssignments, actor, request.notes || 'Manual assignment', 'PENDING');

    const outputAssignments = createdAssignments.map(item => ({
      AssignmentId: item.AssignmentId,
      UserId: item.UserId,
      UserName: item.UserName,
      SlotId: item.SlotId,
      SlotName: item.SlotName,
      StartDate: item.StartDate,
      EndDate: item.EndDate,
      Notes: item.Notes || ''
    }));

    const slotNameLabel = slot.SlotName || slot.Name || 'Shift Slot';
    const startLabel = normalizedStart;
    const endLabel = normalizedEnd;
    const rangeLabel = startLabel === endLabel
      ? startLabel
      : `${startLabel} to ${endLabel}`;
    const userCountLabel = outputAssignments.length === 1 ? 'user' : 'users';
    const message = `Assigned ${outputAssignments.length} ${userCountLabel} to ${slotNameLabel} for ${rangeLabel}.`;

    return {
      success: true,
      created: writeResult.count || createdAssignments.length,
      conflicts,
      failed: failedUsers,
      archived: archivedAssignments,
      message,
      assignments: outputAssignments,
      details: outputAssignments
    };

  } catch (error) {
    console.error('❌ Error creating manual shift assignments:', error);
    safeWriteError('clientAddManualShiftSlots', error);
    return {
      success: false,
      error: error.message
    };
  }
}

  function clientMarkAttendanceStatus(userName, date, status, notes = '') {
  try {
    console.log('📝 Marking attendance status:', { userName, date, status, notes });

    // Use ScheduleUtilities to ensure proper sheet structure
    const sheet = ensureScheduleSheetWithHeaders(ATTENDANCE_STATUS_SHEET, ATTENDANCE_STATUS_HEADERS);

    // Check if entry already exists
    const existingData = readScheduleSheet(ATTENDANCE_STATUS_SHEET) || [];
    const existingEntry = existingData.find(entry =>
      entry.UserName === userName && entry.Date === date
    );

    const now = new Date();

    if (existingEntry) {
      // Update existing entry
      const data = sheet.getDataRange().getValues();
      const headers = data[0];

      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === existingEntry.ID) {
          sheet.getRange(i + 1, headers.indexOf('Status') + 1).setValue(status);
          sheet.getRange(i + 1, headers.indexOf('Notes') + 1).setValue(notes);
          sheet.getRange(i + 1, headers.indexOf('UpdatedAt') + 1).setValue(now);
          break;
        }
      }
    } else {
      // Create new entry using proper header order
      const entry = {
        ID: Utilities.getUuid(),
        UserID: getUserIdByName(userName) || userName,
        UserName: userName,
        Date: date,
        Status: status,
        Notes: notes,
        MarkedBy: Session.getActiveUser().getEmail(),
        CreatedAt: now,
        UpdatedAt: now
      };

      const rowData = ATTENDANCE_STATUS_HEADERS.map(header => entry[header] || '');
      sheet.appendRow(rowData);
    }

    SpreadsheetApp.flush();
    invalidateScheduleCaches();

    return {
      success: true,
      message: `Attendance status updated to ${status} for ${userName} on ${date}`
    };

  } catch (error) {
    console.error('Error marking attendance status:', error);
    safeWriteError('clientMarkAttendanceStatus', error);
    return {
      success: false,
      error: error.message
    };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// SYSTEM DIAGNOSTICS - Enhanced with ScheduleUtilities integration
// ────────────────────────────────────────────────────────────────────────────

/**
 * Run comprehensive system diagnostics
 */
function clientRunSystemDiagnostics() {
  try {
    console.log('🔍 Running comprehensive system diagnostics');

    const diagnostics = {
      timestamp: new Date().toISOString(),
      system: 'Enhanced Schedule Management v4.1 - Integrated',
      spreadsheetConfig: {},
      userSystem: {},
      shiftSlots: {},
      holidays: {},
      attendance: {},
      scheduleUtilities: {},
      issues: [],
      recommendations: []
    };

    // Test spreadsheet configuration using ScheduleUtilities
    try {
      const config = validateScheduleSpreadsheetConfig();
      diagnostics.spreadsheetConfig = {
        ...config,
        canAccess: config.canAccess,
        usingDedicatedSpreadsheet: config.hasScheduleSpreadsheetId
      };
    } catch (error) {
      diagnostics.spreadsheetConfig = {
        canAccess: false,
        error: error.message
      };
      diagnostics.issues.push({
        severity: 'HIGH',
        component: 'Spreadsheet Configuration',
        message: 'Cannot validate spreadsheet configuration: ' + error.message
      });
    }

    // Test user system using MainUtilities integration
    try {
      const users = clientGetScheduleUsers('test-user');
      const attendanceUsers = clientGetAttendanceUsers('test-user');
      
      diagnostics.userSystem = {
        scheduleUsersCount: users.length,
        attendanceUsersCount: attendanceUsers.length,
        working: users.length > 0,
        mainUtilitiesIntegration: true
      };

      if (users.length === 0) {
        diagnostics.issues.push({
          severity: 'HIGH',
          component: 'User System',
          message: 'No users found for scheduling - check MainUtilities integration'
        });
      }
    } catch (error) {
      diagnostics.userSystem = {
        working: false,
        error: error.message,
        mainUtilitiesIntegration: false
      };
      diagnostics.issues.push({
        severity: 'HIGH',
        component: 'User System',
        message: 'MainUtilities integration failed: ' + error.message
      });
    }

    // Test shift slots using ScheduleUtilities
    try {
      const slots = clientGetAllShiftSlots();
      diagnostics.shiftSlots = {
        count: slots.length,
        working: slots.length > 0,
        scheduleUtilitiesIntegration: true
      };

      if (slots.length === 0) {
        diagnostics.issues.push({
          severity: 'MEDIUM',
          component: 'Shift Slots',
          message: 'No shift slots available - defaults will be created'
        });
      }
    } catch (error) {
      diagnostics.shiftSlots = {
        working: false,
        error: error.message,
        scheduleUtilitiesIntegration: false
      };
      diagnostics.issues.push({
        severity: 'HIGH',
        component: 'Shift Slots',
        message: 'ScheduleUtilities integration failed for shift slots: ' + error.message
      });
    }

    // Test holiday system
    try {
      const holidays = clientGetCountryHolidays('JM', 2025);
      diagnostics.holidays = {
        jamaicaHolidays: holidays.success ? holidays.holidays.length : 0,
        supportedCountries: SCHEDULE_SETTINGS.SUPPORTED_COUNTRIES,
        working: holidays.success,
        primaryCountry: SCHEDULE_SETTINGS.PRIMARY_COUNTRY
      };
    } catch (error) {
      diagnostics.holidays = {
        working: false,
        error: error.message
      };
      diagnostics.issues.push({
        severity: 'MEDIUM',
        component: 'Holiday System',
        message: 'Holiday system failed: ' + error.message
      });
    }

    // Test attendance system
    try {
      const dashboard = clientGetAttendanceDashboard('2025-01-01', '2025-01-31');
      diagnostics.attendance = {
        working: dashboard.success,
        hasData: dashboard.success && dashboard.dashboard.totalRecords > 0,
        scheduleUtilitiesIntegration: true
      };
    } catch (error) {
      diagnostics.attendance = {
        working: false,
        error: error.message,
        scheduleUtilitiesIntegration: false
      };
      diagnostics.issues.push({
        severity: 'MEDIUM',
        component: 'Attendance System',
        message: 'Attendance system failed: ' + error.message
      });
    }

    // Test ScheduleUtilities functions
    try {
      const testResult = testScheduleUtilities();
      diagnostics.scheduleUtilities = {
        available: true,
        testsPassed: testResult.success,
        testDetails: testResult.summary || testResult.error
      };
    } catch (error) {
      diagnostics.scheduleUtilities = {
        available: false,
        error: error.message
      };
      diagnostics.issues.push({
        severity: 'HIGH',
        component: 'ScheduleUtilities',
        message: 'ScheduleUtilities not available: ' + error.message
      });
    }

    // Test schedule analytics + health scoring
    try {
      if (typeof evaluateSchedulePerformance === 'function') {
        const analyticsSample = loadScheduleDataBundle(null, { limitSamples: 20 });
        const sampleEvaluation = evaluateSchedulePerformance(
          analyticsSample.scheduleRows.slice(0, 25),
          analyticsSample.demandRows.slice(0, 25),
          analyticsSample.agentProfiles.slice(0, 50),
          { intervalMinutes: 30 }
        );

        diagnostics.scheduleAnalytics = {
          working: true,
          sampleHealthScore: sampleEvaluation.healthScore,
          sampleServiceLevel: sampleEvaluation.summary ? sampleEvaluation.summary.serviceLevel : null
        };
      } else {
        diagnostics.scheduleAnalytics = {
          working: false,
          error: 'Schedule analytics utilities not available'
        };

        diagnostics.issues.push({
          severity: 'MEDIUM',
          component: 'Schedule Analytics',
          message: 'Schedule analytics utilities are not loaded from ScheduleUtilities'
        });
      }
    } catch (error) {
      diagnostics.scheduleAnalytics = {
        working: false,
        error: error.message
      };

      diagnostics.issues.push({
        severity: 'MEDIUM',
        component: 'Schedule Analytics',
        message: 'Schedule analytics evaluation failed: ' + error.message
      });
    }

    // Generate recommendations
    if (diagnostics.issues.length === 0) {
      diagnostics.recommendations.push('System is working well. All components are functional with proper utility integration.');
    } else {
      const highIssues = diagnostics.issues.filter(i => i.severity === 'HIGH');
      if (highIssues.length > 0) {
        diagnostics.recommendations.push(`${highIssues.length} critical issues need immediate attention`);
      }
      
      if (!diagnostics.userSystem.working) {
        diagnostics.recommendations.push('Check MainUtilities integration and Users sheet configuration');
      }
      
      if (!diagnostics.shiftSlots.working || diagnostics.shiftSlots.count === 0) {
        diagnostics.recommendations.push('Create shift slots using the Shift Slots tab');
      }

      if (!diagnostics.spreadsheetConfig.canAccess) {
        diagnostics.recommendations.push('Check spreadsheet access and ScheduleUtilities configuration');
      }
    }

    return {
      success: true,
      diagnostics: diagnostics,
      overallHealth: diagnostics.issues.filter(i => i.severity === 'HIGH').length === 0 ? 'HEALTHY' : 'NEEDS_ATTENTION'
    };

  } catch (error) {
    console.error('Error running diagnostics:', error);
    safeWriteError('clientRunSystemDiagnostics', error);
    return {
      success: false,
      error: error.message,
      diagnostics: null
    };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// SCHEDULE ACTIONS - Approve/Reject functions for frontend
// ────────────────────────────────────────────────────────────────────────────

/**
 * Approve schedules
 */

function clientApproveSchedules(scheduleIds, approvingUserId, notes = '') {
  try {
    if (!Array.isArray(scheduleIds) || !scheduleIds.length) {
      return {
        success: false,
        error: 'Select at least one assignment to approve.'
      };
    }

    const actor = approvingUserId || (typeof getCurrentUser === 'function' ? (getCurrentUser()?.Email || 'System') : 'System');
    const results = [];

    scheduleIds.forEach(id => {
      const update = updateShiftAssignmentRow(id, row => {
        const updated = Object.assign({}, row);
        updated.Status = 'APPROVED';
        updated.UpdatedAt = new Date();
        updated.UpdatedBy = actor;
        if (notes) {
          updated.Notes = updated.Notes ? `${updated.Notes}
${notes}` : notes;
        }
        return updated;
      });
      if (update.success) {
        results.push(update.assignment);
      }
    });

    return {
      success: true,
      approved: results.length,
      assignments: results
    };

  } catch (error) {
    console.error('Error approving assignments:', error);
    safeWriteError('clientApproveSchedules', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Reject schedules
 */

function clientRejectSchedules(scheduleIds, rejectingUserId, reason = '') {
  try {
    if (!Array.isArray(scheduleIds) || !scheduleIds.length) {
      return {
        success: false,
        error: 'Select at least one assignment to reject.'
      };
    }

    const actor = rejectingUserId || (typeof getCurrentUser === 'function' ? (getCurrentUser()?.Email || 'System') : 'System');
    const results = [];

    scheduleIds.forEach(id => {
      const update = updateShiftAssignmentRow(id, row => {
        const updated = Object.assign({}, row);
        updated.Status = 'REJECTED';
        updated.UpdatedAt = new Date();
        updated.UpdatedBy = actor;
        if (reason) {
          updated.Notes = updated.Notes ? `${updated.Notes}
Rejected: ${reason}` : `Rejected: ${reason}`;
        }
        return updated;
      });
      if (update.success) {
        results.push(update.assignment);
      }
    });

    return {
      success: true,
      rejected: results.length,
      assignments: results
    };

  } catch (error) {
    console.error('Error rejecting assignments:', error);
    safeWriteError('clientRejectSchedules', error);
    return {
      success: false,
      error: error.message
    };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// LEGACY COMPATIBILITY AND UTILITY FUNCTIONS
// ────────────────────────────────────────────────────────────────────────────

/**
 * Legacy helper functions for backward compatibility
 */
function getUserIdByName(userName) {
  try {
    const users = readSheet(USERS_SHEET) || [];
    const user = users.find(u =>
      u.UserName === userName ||
      u.FullName === userName ||
      (u.UserName && u.UserName.toLowerCase() === userName.toLowerCase()) ||
      (u.FullName && u.FullName.toLowerCase() === userName.toLowerCase())
    );
    return user ? user.ID : userName;
  } catch (error) {
    console.warn('Error getting user ID by name:', error);
    return userName;
  }
}

/**
 * Check if schedule exists for user within a period - uses ScheduleUtilities
 */

function checkExistingSchedule(userName, periodStart, periodEnd) {
  try {
    const assignments = readShiftAssignments().map(normalizeAssignmentRecord);
    const start = normalizeDateForSheet(periodStart, DEFAULT_SCHEDULE_TIME_ZONE);
    const end = normalizeDateForSheet(periodEnd || periodStart, DEFAULT_SCHEDULE_TIME_ZONE) || start;

    if (!start || !end) {
      return null;
    }

    const userKey = normalizeUserKey(userName);
    return assignments.find(record => {
      if (!record) {
        return false;
      }
      const recordUserKey = normalizeUserKey(record.UserName || '');
      const sameUser = recordUserKey === userKey || (record.UserId && userName && String(record.UserId) === String(userName));
      if (!sameUser) {
        return false;
      }
      return !(record.EndDate < start || record.StartDate > end);
    }) || null;

  } catch (error) {
    console.warn('Error checking existing assignment:', error);
    return null;
  }
}
/**
 * Check if date is a holiday - uses ScheduleUtilities
 */
function checkIfHoliday(dateStr) {
  try {
    const holidays = readScheduleSheet(HOLIDAYS_SHEET) || [];
    return holidays.some(h => h.Date === dateStr);
  } catch (error) {
    console.warn('Error checking holiday:', error);
    return false;
  }
}

function normalizeImportedScheduleRecord(raw, metadata, userLookup, nowIso, timeZone) {
  if (!raw) {
    return null;
  }

  const periodStart = normalizeDateForSheet(
    raw.PeriodStart
      || raw.StartDate
      || raw.AssignmentStart
      || raw.ScheduleStart
      || raw.Date
      || raw.ScheduleDate,
    timeZone
  );

  const dateStr = normalizeDateForSheet(raw.Date || raw.ScheduleDate || periodStart, timeZone);
  const periodEnd = normalizeDateForSheet(
    raw.PeriodEnd
      || raw.EndDate
      || raw.AssignmentEnd
      || raw.ScheduleEnd
      || raw.Date
      || raw.ScheduleDate
      || periodStart,
    timeZone
  );

  const primaryDate = periodStart || dateStr;
  const userName = (raw.UserName || '').toString().trim();

  if (!userName || !primaryDate) {
    return null;
  }

  const userKey = normalizeUserKey(userName);
  const matchedUser = userLookup[userKey];

  const notes = [];
  if (metadata && metadata.sourceMonth) {
    const monthName = getMonthNameFromNumber(metadata.sourceMonth);
    const yearPart = metadata.sourceYear ? ` ${metadata.sourceYear}` : '';
    notes.push(`Imported from ${monthName || 'prior schedule'}${yearPart}`.trim());
  }

  if (raw.SourceDayLabel) {
    notes.push(`Original Day: ${raw.SourceDayLabel}`);
  }

  if (raw.SourceCell && !raw.StartTime) {
    notes.push(`Source: ${raw.SourceCell}`);
  }

  if (raw.Break2Start || raw.Break2End) {
    const break2Start = raw.Break2Start || '';
    const break2End = raw.Break2End || '';
    notes.push(`Break 2: ${break2Start}${break2End ? ` - ${break2End}` : ''}`.trim());
  }

  if (raw.Notes) {
    notes.push(raw.Notes);
  }

  const defaultPriority = typeof metadata.defaultPriority === 'number' ? metadata.defaultPriority : 2;

  return {
    ID: raw.ID || Utilities.getUuid(),
    UserID: raw.UserID || (matchedUser ? matchedUser.ID : ''),
    UserName: matchedUser ? (matchedUser.UserName || matchedUser.FullName) : userName,
    Date: primaryDate,
    PeriodStart: periodStart || primaryDate,
    PeriodEnd: periodEnd || primaryDate,
    SlotID: raw.SlotID || '',
    SlotName: raw.SlotName || `Imported ${raw.SourceDayLabel || 'Shift'}`,
    StartTime: raw.StartTime || '',
    EndTime: raw.EndTime || '',
    OriginalStartTime: raw.OriginalStartTime || raw.StartTime || '',
    OriginalEndTime: raw.OriginalEndTime || raw.EndTime || '',
    BreakStart: raw.BreakStart || '',
    BreakEnd: raw.BreakEnd || '',
    LunchStart: raw.LunchStart || '',
    LunchEnd: raw.LunchEnd || '',
    IsDST: raw.IsDST || '',
    Status: raw.Status || 'PENDING',
    GeneratedBy: raw.GeneratedBy || metadata.importedBy || 'Schedule Importer',
    ApprovedBy: raw.ApprovedBy || '',
    NotificationSent: raw.NotificationSent || '',
    CreatedAt: raw.CreatedAt || nowIso,
    UpdatedAt: nowIso,
    RecurringScheduleID: raw.RecurringScheduleID || '',
    SwapRequestID: raw.SwapRequestID || '',
    Priority: typeof raw.Priority === 'number' ? raw.Priority : defaultPriority,
    Notes: notes.filter(Boolean).join(' | '),
    Location: raw.Location || metadata.location || '',
    Department: raw.Department || metadata.department || ''
  };
}

function buildScheduleUserLookup() {
  try {
    const users = clientGetScheduleUsers('system') || [];
    const lookup = {};

    users.forEach(user => {
      const candidateNames = [
        user.UserName,
        user.FullName,
        user.Email ? user.Email.split('@')[0] : null
      ].filter(Boolean);

      candidateNames.forEach(name => {
        const key = normalizeUserKey(name);
        if (key && !lookup[key]) {
          lookup[key] = user;
        }
      });
    });

    return lookup;
  } catch (error) {
    console.warn('Unable to build user lookup for schedule import:', error);
    return {};
  }
}

function normalizeUserKey(value) {
  return (value || '').toString().trim().toLowerCase().replace(/\s+/g, ' ');
}

function normalizeDateForSheet(value, timeZone) {
  if (value === null || value === undefined) {
    return '';
  }

  if (value instanceof Date) {
    if (isNaN(value.getTime())) {
      return '';
    }
    return Utilities.formatDate(value, timeZone, 'yyyy-MM-dd');
  }

  if (typeof value === 'number') {
    const dateFromNumber = new Date(value);
    if (!isNaN(dateFromNumber.getTime())) {
      return Utilities.formatDate(dateFromNumber, timeZone, 'yyyy-MM-dd');
    }
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) {
      return '';
    }

    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
      return trimmed;
    }

    const parsed = new Date(trimmed);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, timeZone, 'yyyy-MM-dd');
    }

    const maybeNumber = Number(trimmed);
    if (!Number.isNaN(maybeNumber) && maybeNumber > 0) {
      const baseDate = new Date('1899-12-30T00:00:00Z');
      baseDate.setDate(baseDate.getDate() + maybeNumber);
      return Utilities.formatDate(baseDate, timeZone, 'yyyy-MM-dd');
    }
  }

  return '';
}

function calculateDaySpanCount(startDate, endDate, minDate, maxDate) {
  let start = startDate ? new Date(startDate) : null;
  let end = endDate ? new Date(endDate) : null;

  if ((!start || isNaN(start.getTime())) && minDate instanceof Date && !isNaN(minDate.getTime())) {
    start = new Date(minDate);
  }

  if ((!end || isNaN(end.getTime())) && maxDate instanceof Date && !isNaN(maxDate.getTime())) {
    end = new Date(maxDate);
  }

  if (!start || !end || isNaN(start.getTime()) || isNaN(end.getTime())) {
    return 0;
  }

  const millisecondsPerDay = 24 * 60 * 60 * 1000;
  const diff = end.getTime() - start.getTime();
  const days = Math.floor(diff / millisecondsPerDay) + 1;
  return days > 0 ? days : 0;
}

function convertLegacyShiftSlotRecord(raw) {
  if (!raw || typeof raw !== 'object') {
    return null;
  }

  const resolve = (candidates, fallback = '') => {
    for (let i = 0; i < candidates.length; i++) {
      const value = raw[candidates[i]];
      if (value !== undefined && value !== null && String(value).trim() !== '') {
        return value;
      }
    }
    return fallback;
  };

  const daysOfWeek = resolve(['DaysOfWeek', 'Days', 'DayCodes', 'DayIndexes']);
  const parsedDays = Array.isArray(daysOfWeek)
    ? daysOfWeek
    : typeof daysOfWeek === 'string'
      ? daysOfWeek.split(/[;,]/).map(d => parseInt(String(d).trim(), 10)).filter(d => !isNaN(d))
      : [];

  const uuid = (typeof Utilities !== 'undefined' && Utilities.getUuid)
    ? Utilities.getUuid()
    : `legacy-slot-${Math.random().toString(36).slice(2)}`;

  return {
    ID: resolve(['ID', 'SlotID', 'Slot Id', 'Guid', 'Uuid'], uuid),
    Name: resolve(['Name', 'SlotName', 'Title', 'ShiftName', 'Shift']),
    StartTime: resolve(['StartTime', 'Start', 'Start Time', 'ShiftStart']),
    EndTime: resolve(['EndTime', 'End', 'End Time', 'ShiftEnd']),
    DaysOfWeek: parsedDays.length ? parsedDays.join(',') : '1,2,3,4,5',
    DaysOfWeekArray: parsedDays.length ? parsedDays : undefined,
    Department: resolve(['Department', 'Team', 'Campaign', 'Program'], 'General'),
    Location: resolve(['Location', 'Site'], 'Office'),
    MaxCapacity: resolve(['MaxCapacity', 'Capacity', 'Max Agents', 'Headcount'], ''),
    MinCoverage: resolve(['MinCoverage', 'MinimumCoverage', 'Min Agents'], ''),
    Priority: resolve(['Priority', 'Rank', 'Weight'], 2),
    Description: resolve(['Description', 'Notes'], ''),
    BreakDuration: resolve(['BreakDuration', 'Break Minutes', 'BreakLength'], ''),
    LunchDuration: resolve(['LunchDuration', 'Lunch Minutes', 'LunchLength'], ''),
    Break1Duration: resolve(['Break1Duration', 'BreakDuration', 'Break1'], ''),
    Break2Duration: resolve(['Break2Duration', 'Break2'], ''),
    EnableStaggeredBreaks: resolve(['EnableStaggeredBreaks', 'StaggerBreaks', 'Staggered'], false),
    BreakGroups: resolve(['BreakGroups', 'StaggerGroups'], ''),
    StaggerInterval: resolve(['StaggerInterval', 'StaggerMinutes'], ''),
    MinCoveragePct: resolve(['MinCoveragePct', 'CoveragePct'], ''),
    EnableOvertime: resolve(['EnableOvertime', 'AllowOT', 'Overtime'], false),
    MaxDailyOT: resolve(['MaxDailyOT', 'DailyOTHours', 'DailyOvertime'], ''),
    MaxWeeklyOT: resolve(['MaxWeeklyOT', 'WeeklyOTHours', 'WeeklyOvertime'], ''),
    OTApproval: resolve(['OTApproval', 'OvertimeApproval'], ''),
    OTRate: resolve(['OTRate', 'OvertimeRate'], ''),
    OTPolicy: resolve(['OTPolicy', 'OvertimePolicy'], ''),
    AllowSwaps: resolve(['AllowSwaps', 'SwapAllowed'], ''),
    WeekendPremium: resolve(['WeekendPremium', 'Weekend'], ''),
    HolidayPremium: resolve(['HolidayPremium', 'Holiday'], ''),
    AutoAssignment: resolve(['AutoAssignment', 'AutoAssign'], ''),
    RestPeriod: resolve(['RestPeriod', 'RestHours'], ''),
    NotificationLead: resolve(['NotificationLead', 'NotifyHours'], ''),
    HandoverTime: resolve(['HandoverTime', 'Handover'], ''),
    OvertimePolicy: resolve(['OvertimePolicy', 'OTPolicy'], ''),
    IsActive: resolve(['IsActive', 'Active', 'Enabled'], true),
    CreatedBy: resolve(['CreatedBy', 'Author', 'Owner'], 'Legacy Import'),
    CreatedAt: resolve(['CreatedAt', 'Created', 'Created On'], ''),
    UpdatedAt: resolve(['UpdatedAt', 'Updated', 'Updated On'], '')
  };
}

function convertLegacyScheduleRecord(raw) {
  if (!raw || typeof raw !== 'object') {
    return null;
  }

  const resolve = (candidates, fallback = '') => {
    for (let i = 0; i < candidates.length; i++) {
      const key = candidates[i];
      if (key == null) {
        continue;
      }
      const value = raw[key];
      if (value !== undefined && value !== null && String(value).trim() !== '') {
        return value;
      }
    }
    return fallback;
  };

  const userName = resolve(['UserName', 'Agent', 'AgentName', 'Name', 'User']);
  const userId = resolve(['UserID', 'UserId', 'AgentID', 'AgentId', 'EmployeeID']);
  const scheduleDate = resolve(['Date', 'ScheduleDate', 'ShiftDate', 'Day']);
  const scheduleEnd = resolve(['PeriodEnd', 'EndDate', 'ShiftEndDate', 'AssignmentEnd', 'ScheduleEnd'], scheduleDate);
  const slotName = resolve(['SlotName', 'Shift', 'ShiftName', 'Schedule']);

  const timezone = (typeof Session !== 'undefined' && Session.getScriptTimeZone)
    ? Session.getScriptTimeZone()
    : 'UTC';

  const normalizeDate = (value) => {
    if (!value) return '';
    if (value instanceof Date && !isNaN(value.getTime())) {
      return Utilities.formatDate(value, timezone, 'yyyy-MM-dd');
    }

    const parsed = new Date(value);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, timezone, 'yyyy-MM-dd');
    }

    return value;
  };

  const uuid = (typeof Utilities !== 'undefined' && Utilities.getUuid)
    ? Utilities.getUuid()
    : `legacy-schedule-${Math.random().toString(36).slice(2)}`;

  const normalizedStart = normalizeDate(scheduleDate);
  const normalizedEnd = normalizeDate(scheduleEnd) || normalizedStart;

  return {
    ID: resolve(['ID', 'ScheduleID', 'Schedule Id', 'RecordID'], uuid),
    UserID: userId || normalizeUserIdValue(userName),
    UserName: userName || userId,
    Date: normalizedStart,
    PeriodStart: normalizedStart,
    PeriodEnd: normalizedEnd,
    SlotID: resolve(['SlotID', 'ShiftID', 'TemplateID'], ''),
    SlotName: slotName || 'Shift',
    StartTime: resolve(['StartTime', 'Start', 'ShiftStart', 'Begin']),
    EndTime: resolve(['EndTime', 'End', 'ShiftEnd', 'Finish']),
    OriginalStartTime: resolve(['OriginalStartTime', 'StartTime', 'Start']),
    OriginalEndTime: resolve(['OriginalEndTime', 'EndTime', 'End']),
    BreakStart: resolve(['BreakStart', 'BreakStartTime']),
    BreakEnd: resolve(['BreakEnd', 'BreakEndTime']),
    LunchStart: resolve(['LunchStart', 'LunchStartTime']),
    LunchEnd: resolve(['LunchEnd', 'LunchEndTime']),
    IsDST: resolve(['IsDST', 'DST', 'DaylightSavings'], false),
    Status: (resolve(['Status', 'State'], 'PENDING') || 'PENDING').toString().toUpperCase(),
    GeneratedBy: resolve(['GeneratedBy', 'CreatedBy', 'Author'], 'Legacy Import'),
    ApprovedBy: resolve(['ApprovedBy', 'Supervisor']),
    NotificationSent: resolve(['NotificationSent', 'Notified'], false),
    CreatedAt: resolve(['CreatedAt', 'Created', 'Created On'], ''),
    UpdatedAt: resolve(['UpdatedAt', 'Updated', 'Updated On'], ''),
    RecurringScheduleID: resolve(['RecurringScheduleID', 'RecurringID']),
    SwapRequestID: resolve(['SwapRequestID', 'SwapID']),
    Priority: resolve(['Priority', 'Rank'], 2),
    Notes: resolve(['Notes', 'Comments']),
    Location: resolve(['Location', 'Site', 'Office']),
    Department: resolve(['Department', 'Campaign', 'Program'])
  };
}

function calculateWeekSpanCount(startDate, endDate, minDate, maxDate) {
  const daySpan = calculateDaySpanCount(startDate, endDate, minDate, maxDate);
  if (!daySpan || daySpan <= 0) {
    return 0;
  }

  return Math.ceil(daySpan / 7);
}

function getMonthNameFromNumber(monthNumber) {
  const months = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];

  const index = Number(monthNumber) - 1;
  return months[index] || '';
}

function extractSpreadsheetId(input) {
  if (!input) {
    return '';
  }

  const stringValue = String(input).trim();
  if (!stringValue) {
    return '';
  }

  const directMatch = stringValue.match(/[-\w]{25,}/);
  if (directMatch && directMatch[0]) {
    return directMatch[0];
  }

  const urlMatch = stringValue.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (urlMatch && urlMatch[1]) {
    return urlMatch[1];
  }

  const queryMatch = stringValue.match(/[?&]id=([a-zA-Z0-9-_]+)/);
  if (queryMatch && queryMatch[1]) {
    return queryMatch[1];
  }

  return '';
}

function scheduleToNumber(value, fallback = 0) {
  const numeric = Number(value);
  return isFinite(numeric) ? numeric : fallback;
}

function safeNormalizeScheduleDate(value) {
  try {
    if (typeof normalizeScheduleDate === 'function') {
      return normalizeScheduleDate(value);
    }
  } catch (error) {
    console.warn('safeNormalizeScheduleDate: normalizeScheduleDate failed', error);
  }

  if (value instanceof Date) {
    return new Date(value.getTime());
  }

  if (typeof value === 'number') {
    if (value > 100000000000) {
      return new Date(value);
    }
    return new Date(value * 24 * 60 * 60 * 1000);
  }

  if (typeof value === 'string') {
    const text = value.trim();
    if (!text) {
      return null;
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
      return new Date(`${text}T00:00:00`);
    }
    const parsed = new Date(text);
    if (!isNaN(parsed.getTime())) {
      return parsed;
    }
  }

  return null;
}

function safeNormalizeScheduleTimeToMinutes(value) {
  try {
    if (typeof normalizeScheduleTimeToMinutes === 'function') {
      const normalized = normalizeScheduleTimeToMinutes(value);
      if (typeof normalized === 'number' && isFinite(normalized)) {
        return normalized;
      }
    }
  } catch (error) {
    console.warn('safeNormalizeScheduleTimeToMinutes: normalizeScheduleTimeToMinutes failed', error);
  }

  if (value instanceof Date) {
    return value.getHours() * 60 + value.getMinutes();
  }

  if (typeof value === 'number') {
    if (value > 100000000000) {
      const date = new Date(value);
      return date.getHours() * 60 + date.getMinutes();
    }

    if (value >= 0 && value <= 1) {
      return Math.round(value * 24 * 60);
    }

    if (value > 1 && value < 24) {
      return Math.round(value * 60);
    }

    return Math.round(value);
  }

  if (typeof value === 'string') {
    const text = value.trim();
    if (!text) {
      return null;
    }

    const ampmMatch = text.match(/^(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}))?\s*(AM|PM)$/i);
    if (ampmMatch) {
      let hours = parseInt(ampmMatch[1], 10);
      const minutes = parseInt(ampmMatch[2] || '0', 10);
      const period = ampmMatch[4].toUpperCase();
      if (period === 'PM' && hours < 12) {
        hours += 12;
      }
      if (period === 'AM' && hours === 12) {
        hours = 0;
      }
      return hours * 60 + minutes;
    }

    const isoMatch = text.match(/^(\d{4}-\d{2}-\d{2})[ T](\d{2}):(\d{2})/);
    if (isoMatch) {
      return parseInt(isoMatch[2], 10) * 60 + parseInt(isoMatch[3], 10);
    }

    const hhmmMatch = text.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (hhmmMatch) {
      return parseInt(hhmmMatch[1], 10) * 60 + parseInt(hhmmMatch[2], 10);
    }
  }

  return null;
}

function safeNormalizeSchedulePercentage(value, fallback = 0) {
  try {
    if (typeof normalizeSchedulePercentage === 'function') {
      return normalizeSchedulePercentage(value, fallback);
    }
  } catch (error) {
    console.warn('safeNormalizeSchedulePercentage: normalizeSchedulePercentage failed', error);
  }

  const numeric = Number(value);
  if (!isFinite(numeric)) {
    return fallback;
  }
  if (Math.abs(numeric) > 1) {
    return numeric / 100;
  }
  return numeric;
}

function minutesToTimeString(minutes) {
  if (!isFinite(minutes)) {
    return '';
  }

  const normalized = ((minutes % (24 * 60)) + (24 * 60)) % (24 * 60);
  const hours = Math.floor(normalized / 60);
  const mins = Math.round(normalized % 60);
  return `${String(hours).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
}

function formatDateForOutput(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return '';
  }
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function mapScheduleRowToAgentShift(row) {
  if (!row || typeof row !== 'object') {
    return null;
  }

  const date = safeNormalizeScheduleDate(row.Date || row.ScheduleDate || row.PeriodStart || row.StartDate || row.Day);
  const startMinutes = safeNormalizeScheduleTimeToMinutes(row.StartTime || row.PeriodStart || row.ScheduleStart || row.ShiftStart);
  const endMinutes = safeNormalizeScheduleTimeToMinutes(row.EndTime || row.PeriodEnd || row.ScheduleEnd || row.ShiftEnd);

  const startDateTime = date ? combineDateAndMinutes(date, startMinutes !== null && startMinutes !== undefined ? startMinutes : 0) : null;
  const endDateTime = date ? combineDateAndMinutes(date, endMinutes !== null && endMinutes !== undefined ? endMinutes : (startMinutes || 0)) : null;

  const fallbackId = (typeof Utilities !== 'undefined' && Utilities.getUuid)
    ? Utilities.getUuid()
    : `schedule_${Date.now()}_${Math.floor(Math.random() * 1000)}`;

  return {
    id: row.ID || row.Id || row.ScheduleID || row.ScheduleId || row.scheduleId || fallbackId,
    date: date ? formatDateForOutput(date) : '',
    dateIso: date ? date.toISOString() : '',
    dayOfWeek: date ? date.toLocaleDateString('en-US', { weekday: 'short' }) : '',
    startTime: startMinutes !== null && startMinutes !== undefined ? minutesToTimeString(startMinutes) : '',
    endTime: endMinutes !== null && endMinutes !== undefined ? minutesToTimeString(endMinutes) : '',
    startTimestamp: startDateTime ? startDateTime.getTime() : null,
    endTimestamp: endDateTime ? endDateTime.getTime() : null,
    startDateTime: startDateTime ? startDateTime.toISOString() : '',
    endDateTime: endDateTime ? endDateTime.toISOString() : '',
    shiftSlot: row.SlotName || row.SlotID || row.Slot || '',
    status: String(row.Status || row.State || 'pending').toLowerCase(),
    location: row.Location || '',
    department: row.Department || '',
    skill: row.Skill || row.Queue || '',
    notes: row.Notes || '',
    raw: row
  };
}

function buildScheduleRowLookup(rows) {
  const map = new Map();
  (rows || []).forEach(row => {
    const idCandidates = [row.ID, row.Id, row.ScheduleID, row.ScheduleId, row.scheduleId];
    for (let i = 0; i < idCandidates.length; i++) {
      const id = idCandidates[i];
      if (!id) continue;
      const key = String(id).trim();
      if (key) {
        map.set(key, row);
        break;
      }
    }
  });
  return map;
}

function buildAgentProfileLookup(profiles) {
  const map = new Map();
  (profiles || []).forEach(profile => {
    const id = normalizeUserIdValue(profile && (profile.ID || profile.Id || profile.UserID || profile.UserId));
    if (id) {
      map.set(id, profile);
    }
  });
  return map;
}

function resolveAgentScheduleWindow(agentId, context, options = {}) {
  const windowDays = Number(options.windowDays) > 0 ? Number(options.windowDays) : 30;
  const startDate = safeNormalizeScheduleDate(options.startDate) || new Date();
  startDate.setHours(0, 0, 0, 0);

  const endDate = safeNormalizeScheduleDate(options.endDate) || new Date(startDate.getTime() + windowDays * 24 * 60 * 60 * 1000);
  endDate.setHours(23, 59, 59, 999);

  const bundle = loadScheduleDataBundle(context.campaignId || context.providedCampaignId, {
    managedUserIds: [agentId],
    startDate,
    endDate
  });

  const agentSchedules = bundle.scheduleRows.filter(row => normalizeUserIdValue(row.UserID || row.UserId || row.AgentID || row.AgentId) === agentId);

  return {
    agentSchedules,
    bundle,
    startDate,
    endDate
  };
}

function formatShiftSwapRequestForAgent(row, agentId, lookups = {}) {
  if (!row) {
    return null;
  }

  const scheduleLookup = lookups.scheduleLookup || new Map();
  const profileLookup = lookups.profileLookup || new Map();

  const requestorId = normalizeUserIdValue(row.RequestorUserID || row.RequestorUserId);
  const targetId = normalizeUserIdValue(row.TargetUserID || row.TargetUserId);

  if (agentId && requestorId !== agentId && targetId !== agentId) {
    return null;
  }

  const isRequestor = requestorId === agentId;
  const counterpartId = isRequestor ? targetId : requestorId;

  const requestorScheduleRow = scheduleLookup.get(String(row.RequestorScheduleID || row.RequestorScheduleId || '').trim());
  const targetScheduleRow = scheduleLookup.get(String(row.TargetScheduleID || row.TargetScheduleId || '').trim());

  const myScheduleRow = isRequestor ? requestorScheduleRow : targetScheduleRow;
  const theirScheduleRow = isRequestor ? targetScheduleRow : requestorScheduleRow;

  const myShift = mapScheduleRowToAgentShift(myScheduleRow);
  const theirShift = mapScheduleRowToAgentShift(theirScheduleRow);

  const swapDate = safeNormalizeScheduleDate(row.SwapDate || (myShift && myShift.date ? myShift.date : null));
  const counterpartProfile = profileLookup.get(counterpartId);
  const counterpartName = (counterpartProfile && (counterpartProfile.FullName || counterpartProfile.UserName || counterpartProfile.Email))
    || (isRequestor ? row.TargetUserName : row.RequestorUserName)
    || 'Teammate';

  const createdAtRaw = row.CreatedAt || row.RequestedAt || swapDate || null;
  const createdAtDate = createdAtRaw instanceof Date ? createdAtRaw : (createdAtRaw ? new Date(createdAtRaw) : null);

  const statusValue = String(row.Status || row.status || (typeof SHIFT_SWAP_STATUS !== 'undefined' ? SHIFT_SWAP_STATUS.PENDING : 'PENDING')).toUpperCase();

  return {
    id: row.ID || row.Id || row.id || '',
    status: statusValue.toLowerCase(),
    statusRaw: statusValue,
    myShiftId: myShift ? myShift.id : (isRequestor ? (row.RequestorScheduleID || row.RequestorScheduleId || '') : (row.TargetScheduleID || row.TargetScheduleId || '')),
    theirShiftId: theirShift ? theirShift.id : (!isRequestor ? (row.RequestorScheduleID || row.RequestorScheduleId || '') : (row.TargetScheduleID || row.TargetScheduleId || '')),
    myShiftDate: myShift && myShift.date ? myShift.date : (swapDate ? formatDateForOutput(swapDate) : ''),
    myShiftTime: myShift && myShift.startTime ? `${myShift.startTime}${myShift.endTime ? ` - ${myShift.endTime}` : ''}` : '',
    theirShiftDate: theirShift && theirShift.date ? theirShift.date : '',
    theirShiftTime: theirShift && theirShift.startTime ? `${theirShift.startTime}${theirShift.endTime ? ` - ${theirShift.endTime}` : ''}` : '',
    swapWith: counterpartId || '',
    swapWithName: counterpartName,
    reason: row.Reason || row.reason || '',
    requestedAt: createdAtDate ? createdAtDate.toISOString() : '',
    raw: row
  };
}

function combineDateAndMinutes(date, minutes) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return null;
  }
  const clone = new Date(date.getTime());
  clone.setHours(0, 0, 0, 0);
  const normalizedMinutes = Number(minutes);
  if (isFinite(normalizedMinutes)) {
    clone.setMinutes(normalizedMinutes);
  }
  return clone;
}

function loadScheduleDataBundle(campaignId, options = {}) {
  const normalizedCampaignId = normalizeCampaignIdValue(campaignId);
  const managedUserIds = Array.isArray(options.managedUserIds)
    ? options.managedUserIds.map(normalizeUserIdValue).filter(Boolean)
    : [];
  const managedUserSet = new Set(managedUserIds);

  const startDate = options.startDate ? safeNormalizeScheduleDate(options.startDate) : (options.dateRange && options.dateRange.start ? safeNormalizeScheduleDate(options.dateRange.start) : null);
  const endDate = options.endDate ? safeNormalizeScheduleDate(options.endDate) : (options.dateRange && options.dateRange.end ? safeNormalizeScheduleDate(options.dateRange.end) : null);
  const inclusiveEnd = endDate ? new Date(endDate.getTime()) : null;
  if (inclusiveEnd) {
    inclusiveEnd.setHours(23, 59, 59, 999);
  }

  const limitSamples = scheduleToNumber(options.limitSamples, 0);

  const loadSheet = (sheetName) => {
    if (typeof readScheduleSheet !== 'function') {
      return [];
    }
    try {
      let rows = readScheduleSheet(sheetName) || [];
      if (limitSamples && rows.length > limitSamples) {
        rows = rows.slice(0, limitSamples);
      }
      return rows;
    } catch (error) {
      console.warn('loadScheduleDataBundle: unable to read sheet', sheetName, error);
      return [];
    }
  };

    const scheduleRows = readShiftAssignments().map(normalizeAssignmentRecord);
  const demandRows = loadSheet(DEMAND_SHEET);
  const ftePlanRows = loadSheet(FTE_PLAN_SHEET);

  let agentProfiles = [];
  try {
    agentProfiles = readSheet(USERS_SHEET) || [];
    if (limitSamples && agentProfiles.length > limitSamples * 2) {
      agentProfiles = agentProfiles.slice(0, limitSamples * 2);
    }
  } catch (error) {
    console.warn('loadScheduleDataBundle: unable to read users sheet', error);
  }

  const matchesCampaign = (record) => {
    if (!normalizedCampaignId) {
      return true;
    }
    const candidates = [
      record && record.Campaign,
      record && record.CampaignID,
      record && record.CampaignId,
      record && record.campaign,
      record && record.campaignId,
      record && record.campaignID,
      record && record.AssignedCampaign
    ];
    return candidates.some(value => normalizeCampaignIdValue(value) === normalizedCampaignId);
  };

  const matchesDateRange = (record) => {
    if (!startDate && !endDate) {
      return true;
    }

    const dateCandidates = [
      record && record.Date,
      record && record.ScheduleDate,
      record && record.PeriodStart,
      record && record.StartDate,
      record && record.Day,
      record && record.IntervalStart,
      record && record.intervalStart
    ];

    let recordDate = null;
    for (let i = 0; i < dateCandidates.length; i++) {
      const candidate = safeNormalizeScheduleDate(dateCandidates[i]);
      if (candidate) {
        recordDate = candidate;
        break;
      }
    }

    if (!recordDate) {
      return true;
    }

    const timeValue = recordDate.getTime();
    if (startDate && timeValue < startDate.getTime()) {
      return false;
    }
    if (inclusiveEnd && timeValue > inclusiveEnd.getTime()) {
      return false;
    }
    return true;
  };

  const matchesUserFilter = (record) => {
    if (!managedUserSet.size) {
      return true;
    }
    const userId = normalizeUserIdValue(record && (record.UserID || record.UserId || record.AgentID || record.AgentId));
    return managedUserSet.has(userId);
  };

  const filterRows = (rows, options = {}) => rows.filter(row => matchesCampaign(row) && matchesDateRange(row) && (options.skipUserFilter || matchesUserFilter(row)));

  const filteredSchedules = filterRows(scheduleRows);
  const filteredDemand = filterRows(demandRows, { skipUserFilter: true });
  const filteredFtePlans = filterRows(ftePlanRows, { skipUserFilter: true });
  const filteredProfiles = agentProfiles.filter(profile => matchesCampaign(profile));

  return {
    campaignId: normalizedCampaignId,
    scheduleRows: filteredSchedules,
    demandRows: filteredDemand,
    ftePlanRows: filteredFtePlans,
    agentProfiles: filteredProfiles,
    startDate,
    endDate
  };
}

function buildScheduleRecommendations(evaluation, bundle) {
  const recommendations = [];
  if (!evaluation || !evaluation.summary) {
    return recommendations;
  }

  const coverage = evaluation.coverage || {};
  const fairness = evaluation.fairness || {};
  const compliance = evaluation.compliance || {};

  if (coverage.serviceLevel < 80) {
    const topInterval = (coverage.backlogRiskIntervals || [])[0];
    if (topInterval) {
      recommendations.push(`Add staffing to ${topInterval.intervalKey} for skill ${topInterval.skill || 'general'} (deficit ${topInterval.deficit} FTE).`);
    } else {
      recommendations.push('Increase staffing in critical intervals to protect service level.');
    }
  }

  if (coverage.peakCoverage < 85) {
    recommendations.push('Rebalance opening and closing coverage to meet first/last hour SLAs.');
  }

  if (fairness.rotationHealth < 75) {
    recommendations.push('Review weekend and night rotation to improve fairness.');
  }

  if (compliance.complianceScore < 85) {
    recommendations.push('Resolve compliance issues (breaks, rest periods, overtime) before publishing schedules.');
  }

  if (!bundle || !bundle.scheduleRows || bundle.scheduleRows.length === 0) {
    recommendations.push('No schedules found for the selected filters. Import or generate schedules to proceed.');
  }

  return recommendations;
}

function persistScheduleHealthSnapshot(context, evaluation, bundle, options = {}) {
  if (!context || !evaluation || !evaluation.summary) {
    return;
  }

  if (typeof ensureScheduleSheetWithHeaders !== 'function') {
    return;
  }

  try {
    const sheet = ensureScheduleSheetWithHeaders(SCHEDULE_HEALTH_SHEET, SCHEDULE_HEALTH_HEADERS);
    const id = (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.getUuid === 'function')
      ? Utilities.getUuid()
      : `health_${Date.now()}`;

    const totalMinutes = (bundle.scheduleRows || []).reduce((sum, row) => {
      const start = safeNormalizeScheduleTimeToMinutes(row.StartTime || row.PeriodStart || row.ScheduleStart || row.ShiftStart);
      const end = safeNormalizeScheduleTimeToMinutes(row.EndTime || row.PeriodEnd || row.ScheduleEnd || row.ShiftEnd);
      if (start === null || end === null) {
        return sum;
      }
      let diff = end - start;
      if (diff < 0) {
        diff += 24 * 60;
      }
      return sum + diff;
    }, 0);

    const agentSet = new Set((bundle.scheduleRows || []).map(row => normalizeUserIdValue(row.UserID || row.UserId || row.AgentID || row.AgentId)).filter(Boolean));
    const totalHours = totalMinutes / 60;
    const standardHours = scheduleToNumber(options.standardHoursPerAgent || 8, 8);
    const overtimeHours = Math.max(0, totalHours - (agentSet.size * standardHours));

    const summary = evaluation.summary;
    const fairness = evaluation.fairness || {};
    const compliance = evaluation.compliance || {};
    const coverage = evaluation.coverage || {};

    const costPerStaffedHour = options.costPerStaffedHour || '';
    const rowValues = [
      id,
      context.campaignId || context.providedCampaignId || '',
      evaluation.generatedAt,
      summary.serviceLevel || 0,
      summary.asa || 0,
      summary.abandonRate || 0,
      summary.occupancy || 0,
      summary.occupancy || 0,
      Number(overtimeHours.toFixed(2)),
      costPerStaffedHour,
      fairness.fairnessIndex || 0,
      fairness.preferenceSatisfaction || 0,
      compliance.complianceScore || 0,
      summary.scheduleEfficiency || coverage.coverageScore || 0,
      `SL ${summary.serviceLevel || 0}%, Fairness ${fairness.fairnessIndex || 0}, Compliance ${compliance.complianceScore || 0}`
    ];

    sheet.appendRow(rowValues);
  } catch (error) {
    console.warn('persistScheduleHealthSnapshot failed:', error);
  }
}

function clientGetAttendanceDataRange(startDate, endDate, campaignId = null) {
  try {
    const attendanceData = readScheduleSheet(ATTENDANCE_STATUS_SHEET) || [];

    const normalizeDate = (value) => {
      if (value instanceof Date) {
        return new Date(value.getTime());
      }
      if (typeof value === 'number') {
        const parsed = new Date(value);
        return isNaN(parsed.getTime()) ? null : parsed;
      }
      if (typeof value === 'string' && value.trim().length > 0) {
        const parsed = new Date(value);
        return isNaN(parsed.getTime()) ? null : parsed;
      }
      return null;
    };

    const rangeStart = normalizeDate(startDate);
    const rangeEnd = normalizeDate(endDate);

    const filtered = attendanceData.filter(record => {
      const recordDate = normalizeDate(record.Date || record.date);
      if (!recordDate) {
        return false;
      }

      if (rangeStart && recordDate < rangeStart) {
        return false;
      }
      if (rangeEnd && recordDate > rangeEnd) {
        return false;
      }

      if (campaignId) {
        const recordCampaign = record.CampaignID || record.CampaignId || record.Campaign || null;
        if (recordCampaign && recordCampaign !== campaignId) {
          return false;
        }
      }

      return true;
    });

    const toIsoDate = (date) => {
      if (!(date instanceof Date) || isNaN(date.getTime())) {
        return '';
      }
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    };

    const records = filtered
      .map(record => {
        const date = normalizeDate(record.Date || record.date);
        const isoDate = toIsoDate(date);
        const userName = record.UserName || record.User || record.user || '';
        const status = record.Status || record.status || record.state || '';

        if (!userName || !isoDate || !status) {
          return null;
        }

        return {
          userName,
          status,
          date: isoDate,
          notes: record.Notes || record.notes || ''
        };
      })
      .filter(Boolean);

    return {
      success: true,
      records
    };
  } catch (error) {
    console.error('Error retrieving attendance data range:', error);
    safeWriteError('clientGetAttendanceDataRange', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function clientGetScheduleDashboard(managerIdCandidate, campaignIdCandidate, options = {}) {
  try {
    console.log('📊 Building schedule dashboard for manager/campaign:', managerIdCandidate, campaignIdCandidate);

    if (typeof evaluateSchedulePerformance !== 'function') {
      throw new Error('Schedule analytics utilities are not available. Ensure ScheduleUtilities is loaded.');
    }

    const context = clientGetScheduleContext(managerIdCandidate || null, campaignIdCandidate || null);
    if (!context || !context.success) {
      return {
        success: false,
        error: context && context.error ? context.error : 'Unable to resolve schedule context',
        context
      };
    }

    const dateRange = options.dateRange || {};
    const bundle = loadScheduleDataBundle(context.campaignId || context.providedCampaignId || null, {
      managedUserIds: (context.managedUserIds || []).concat(context.managerId ? [context.managerId] : []),
      startDate: options.startDate || dateRange.start,
      endDate: options.endDate || dateRange.end
    });

    const metricsOptions = Object.assign({}, options.metrics || {}, {
      intervalMinutes: options.intervalMinutes || (options.metrics && options.metrics.intervalMinutes) || 30,
      targetServiceLevel: options.targetServiceLevel || (options.metrics && options.metrics.targetServiceLevel) || 0.8,
      baselineASA: options.baselineASA || (options.metrics && options.metrics.baselineASA) || 45,
      openingHour: options.openingHour || (options.metrics && options.metrics.openingHour),
      closingHour: options.closingHour || (options.metrics && options.metrics.closingHour)
    });

    const evaluation = evaluateSchedulePerformance(bundle.scheduleRows, bundle.demandRows, bundle.agentProfiles, metricsOptions);

    const agentSet = new Set(bundle.scheduleRows.map(row => normalizeUserIdValue(row.UserID || row.UserId || row.AgentID || row.AgentId)).filter(Boolean));
    const totalMinutes = bundle.scheduleRows.reduce((sum, row) => {
      const start = safeNormalizeScheduleTimeToMinutes(row.StartTime || row.PeriodStart || row.ScheduleStart || row.ShiftStart);
      const end = safeNormalizeScheduleTimeToMinutes(row.EndTime || row.PeriodEnd || row.ScheduleEnd || row.ShiftEnd);
      if (start === null || end === null) {
        return sum;
      }
      let diff = end - start;
      if (diff < 0) {
        diff += 24 * 60;
      }
      return sum + diff;
    }, 0);

    const totalHours = Number((totalMinutes / 60).toFixed(2));
    const rosterSummary = {
      agentCount: agentSet.size,
      totalHours,
      averageHoursPerAgent: agentSet.size ? Number((totalHours / agentSet.size).toFixed(2)) : 0
    };

    const fteTotals = bundle.ftePlanRows.reduce((acc, row) => {
      acc.planned += scheduleToNumber(row.PlannedFTE || row.Planned || row.FTEPlanned, 0);
      acc.actual += scheduleToNumber(row.ActualFTE || row.Actual || row.FTEActual, 0);
      return acc;
    }, { planned: 0, actual: 0 });

    const recommendations = buildScheduleRecommendations(evaluation, bundle);

    const response = {
      success: true,
      context: {
        managerId: context.managerId,
        campaignId: context.campaignId,
        providedManagerId: context.providedManagerId,
        providedCampaignId: context.providedCampaignId,
        permissions: context.permissions,
        managedUserCount: context.managedUserCount
      },
      generatedAt: evaluation.generatedAt,
      healthScore: evaluation.healthScore,
      summary: evaluation.summary,
      coverage: evaluation.coverage,
      fairness: evaluation.fairness,
      compliance: evaluation.compliance,
      totals: {
        requiredFTE: evaluation.coverage.totalRequiredFTE,
        staffedFTE: evaluation.coverage.totalStaffedFTE,
        plannedFTE: Number(fteTotals.planned.toFixed(2)),
        actualFTE: Number(fteTotals.actual.toFixed(2)),
        varianceFTE: Number((fteTotals.actual - fteTotals.planned).toFixed(2)),
        rosterHours: totalHours,
        agentCount: rosterSummary.agentCount
      },
      roster: rosterSummary,
      backlogIntervals: evaluation.coverage.backlogRiskIntervals,
      recommendations,
      demandSamples: bundle.demandRows.slice(0, 50)
    };

    if (!options.skipPersistence) {
      persistScheduleHealthSnapshot(context, evaluation, bundle, options);
    }

    return response;
  } catch (error) {
    console.error('Error generating schedule dashboard:', error);
    safeWriteError && safeWriteError('clientGetScheduleDashboard', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function normalizeManagedRosterPayload(payload) {
  const result = {
    recognized: false,
    users: [],
    error: null
  };

  if (payload == null) {
    return result;
  }

  if (Array.isArray(payload)) {
    result.recognized = true;
    result.users = payload.filter(user => user && typeof user === 'object');
    return result;
  }

  if (payload && typeof payload === 'object') {
    if (Array.isArray(payload.users)) {
      result.recognized = true;
      result.users = payload.users.filter(user => user && typeof user === 'object');
      if (payload.success === false && payload.error) {
        result.error = String(payload.error);
      }
      return result;
    }

    if (Array.isArray(payload.managedUsers)) {
      result.recognized = true;
      result.users = payload.managedUsers.filter(user => user && typeof user === 'object');
      if (payload.success === false && payload.error) {
        result.error = String(payload.error);
      }
      return result;
    }

    if (payload.success === false && payload.error) {
      result.recognized = true;
      result.error = String(payload.error);
    }
  }

  return result;
}

function resolveUnifiedManagedRoster(managerId) {
  const normalizedManagerId = normalizeUserIdValue(managerId);
  const response = {
    users: [],
    source: '',
    warnings: [],
    managedUserIds: []
  };

  if (!normalizedManagerId) {
    response.warnings.push('Manager identifier unavailable for roster resolution.');
    return response;
  }

  const attempts = [
    { name: 'clientGetManagedUsersList', fn: () => clientGetManagedUsersList(normalizedManagerId) },
    {
      name: 'clientGetManagedUsers',
      fn: () => (typeof clientGetManagedUsers === 'function' ? clientGetManagedUsers(normalizedManagerId) : null)
    }
  ];

  for (let index = 0; index < attempts.length; index++) {
    const attempt = attempts[index];
    if (typeof attempt.fn !== 'function') {
      continue;
    }

    try {
      const raw = attempt.fn();
      const parsed = normalizeManagedRosterPayload(raw);

      if (!parsed.recognized) {
        continue;
      }

      if (parsed.error) {
        response.warnings.push(`${attempt.name}: ${parsed.error}`);
      }

      response.users = parsed.users;
      response.source = attempt.name;
      response.managedUserIds = parsed.users
        .map(user => normalizeUserIdValue(user && (user.ID || user.UserID || user.id || user.userId)))
        .filter(Boolean);
      return response;
    } catch (error) {
      response.warnings.push(`${attempt.name}: ${error && error.message ? error.message : error}`);
    }
  }

  return response;
}

function buildUnifiedUserCollection(...collections) {
  const map = new Map();

  const addUser = (user) => {
    if (!user || typeof user !== 'object') {
      return;
    }

    const normalizedId = normalizeUserIdValue(user.ID || user.UserID || user.id || user.userId);
    const normalizedUserName = (user.UserName || user.username || '').toString().trim().toLowerCase();
    const normalizedEmail = (user.Email || user.email || '').toString().trim().toLowerCase();

    const key = normalizedId
      ? `id:${normalizedId}`
      : (normalizedUserName ? `username:${normalizedUserName}` : (normalizedEmail ? `email:${normalizedEmail}` : null));

    if (!key) {
      return;
    }

    const existing = map.get(key) || {};

    const normalized = Object.assign({}, existing, user, {
      ID: normalizedId || existing.ID || '',
      UserName: user.UserName || user.username || existing.UserName || existing.username || '',
      FullName: user.FullName || user.fullName || existing.FullName || existing.fullName || user.UserName || existing.UserName || '',
      Email: user.Email || user.email || existing.Email || existing.email || '',
      CampaignID: user.CampaignID || user.campaignID || existing.CampaignID || existing.campaignID || '',
      campaignName: user.campaignName || user.CampaignName || existing.campaignName || existing.CampaignName || '',
      EmploymentStatus: user.EmploymentStatus || existing.EmploymentStatus || 'Active',
      HireDate: user.HireDate || existing.HireDate || '',
      TerminationDate: user.TerminationDate || user.terminationDate || existing.TerminationDate || existing.terminationDate || '',
      isActive: typeof user.isActive === 'boolean'
        ? user.isActive
        : (typeof existing.isActive === 'boolean' ? existing.isActive : isUserConsideredActive(user)),
      roleNames: Array.isArray(user.roleNames)
        ? user.roleNames.slice()
        : (Array.isArray(existing.roleNames) ? existing.roleNames.slice() : [])
    });

    map.set(key, normalized);
  };

  collections
    .filter(collection => Array.isArray(collection) && collection.length)
    .forEach(collection => collection.forEach(addUser));

  const merged = Array.from(map.values());
  merged.sort((a, b) => {
    const nameA = (a.FullName || a.UserName || '').toString().toLowerCase();
    const nameB = (b.FullName || b.UserName || '').toString().toLowerCase();
    return nameA.localeCompare(nameB);
  });

  return merged;
}

function resolveUnifiedScheduleRange(request = {}, timeZone = DEFAULT_SCHEDULE_TIME_ZONE) {
  const now = new Date();
  const fallbackStart = normalizeDateForSheet(new Date(now.getFullYear(), now.getMonth(), 1), timeZone);
  const fallbackEnd = normalizeDateForSheet(new Date(now.getFullYear(), now.getMonth() + 1, 0), timeZone);

  const candidateStart = request.scheduleStart || request.startDate || request.filterStartDate || request.schedulesStart;
  const candidateEnd = request.scheduleEnd || request.endDate || request.filterEndDate || request.schedulesEnd;

  const startDate = normalizeDateForSheet(candidateStart, timeZone) || fallbackStart;
  const endDate = normalizeDateForSheet(candidateEnd, timeZone) || fallbackEnd;

  return {
    startDate,
    endDate,
    fallbackStart,
    fallbackEnd
  };
}

function resolveUnifiedAttendanceRange(request = {}, scheduleRange = {}, timeZone = DEFAULT_SCHEDULE_TIME_ZONE) {
  const monthCandidate = Number(request.attendanceMonth || request.month);
  const yearCandidate = Number(request.attendanceYear || request.year);

  let resolvedYear = Number.isFinite(yearCandidate) && yearCandidate > 1900 ? yearCandidate : null;
  let resolvedMonth = Number.isFinite(monthCandidate) && monthCandidate >= 1 && monthCandidate <= 12 ? monthCandidate : null;

  if (!resolvedYear && scheduleRange.startDate) {
    const parsed = new Date(scheduleRange.startDate);
    if (!isNaN(parsed.getTime())) {
      resolvedYear = parsed.getFullYear();
    }
  }

  if (!resolvedMonth && scheduleRange.startDate) {
    const parsed = new Date(scheduleRange.startDate);
    if (!isNaN(parsed.getTime())) {
      resolvedMonth = parsed.getMonth() + 1;
    }
  }

  if (!resolvedYear) {
    resolvedYear = new Date().getFullYear();
  }

  if (!resolvedMonth) {
    resolvedMonth = new Date().getMonth() + 1;
  }

  const monthStart = new Date(resolvedYear, resolvedMonth - 1, 1);
  const monthEnd = new Date(resolvedYear, resolvedMonth, 0);

  const startDate = normalizeDateForSheet(request.attendanceStart || monthStart, timeZone)
    || normalizeDateForSheet(monthStart, timeZone);
  const endDate = normalizeDateForSheet(request.attendanceEnd || monthEnd, timeZone)
    || normalizeDateForSheet(monthEnd, timeZone);

  const yearStart = normalizeDateForSheet(`${resolvedYear}-01-01`, timeZone);
  const yearEnd = normalizeDateForSheet(`${resolvedYear}-12-31`, timeZone);

  return {
    startDate,
    endDate,
    month: resolvedMonth,
    year: resolvedYear,
    yearRange: { start: yearStart, end: yearEnd }
  };
}

function clientGetScheduleUnifiedState(request = {}) {
  try {
    const options = (request && typeof request === 'object') ? request : {};
    const candidateManagerId = normalizeUserIdValue(
      options.managerId || options.userId || options.requestingUserId || options.identityUserId
    );
    const candidateCampaignId = normalizeCampaignIdValue(
      options.campaignId || options.teamId || options.programId || options.identityCampaignId
    );

    const context = clientGetScheduleContext(candidateManagerId || null, candidateCampaignId || null);
    if (!context || !context.success) {
      return {
        success: false,
        error: context && context.error ? context.error : 'Unable to resolve schedule context',
        context
      };
    }

    const resolvedManagerId = normalizeUserIdValue(
      options.managerId
      || context.managerId
      || context.providedManagerId
      || (context.user && (context.user.ID || context.user.UserID))
      || candidateManagerId
      || context.identity?.userId
    );

    const resolvedCampaignId = normalizeCampaignIdValue(
      options.campaignId
      || context.campaignId
      || context.providedCampaignId
      || candidateCampaignId
    );

    const scheduleRange = resolveUnifiedScheduleRange(options, DEFAULT_SCHEDULE_TIME_ZONE);
    const attendanceRange = resolveUnifiedAttendanceRange(options, scheduleRange, DEFAULT_SCHEDULE_TIME_ZONE);

    const scheduleUsers = clientGetScheduleUsers(resolvedManagerId || 'system', resolvedCampaignId || null) || [];
    const roster = resolveUnifiedManagedRoster(resolvedManagerId || candidateManagerId || context.identity?.userId || '');

    const scheduleFilters = {
      startDate: scheduleRange.startDate,
      endDate: scheduleRange.endDate,
      campaign: resolvedCampaignId || undefined
    };

    const assignments = options.includeSchedules === false
      ? { success: true, schedules: [], total: 0, filters: scheduleFilters }
      : clientGetAllSchedules(scheduleFilters);

    const assignmentUsers = (typeof collectUsersFromScheduleAssignments === 'function')
      ? collectUsersFromScheduleAssignments(assignments, [scheduleUsers, roster.users])
      : [];

    const shiftSlots = options.includeShiftSlots === false ? [] : clientGetAllShiftSlots();

    const dashboard = options.includeScheduleDashboard === false
      ? null
      : clientGetScheduleDashboard(resolvedManagerId || null, resolvedCampaignId || null, {
          startDate: scheduleRange.startDate,
          endDate: scheduleRange.endDate,
          intervalMinutes: options.intervalMinutes || 30,
          openingHour: options.openingHour || 8,
          closingHour: options.closingHour || 21,
          skipPersistence: options.skipDashboardPersistence === true
        });

    const attendanceUsers = options.includeAttendanceUsers === false
      ? []
      : clientGetAttendanceUsers(resolvedManagerId || null, resolvedCampaignId || null);

    const attendanceUserRecords = (typeof buildUserRecordsFromNames === 'function')
      ? buildUserRecordsFromNames(attendanceUsers)
      : attendanceUsers.map(name => ({
        ID: '',
        UserID: '',
        UserName: String(name || ''),
        FullName: String(name || ''),
        Email: '',
        CampaignID: '',
        campaignName: '',
        EmploymentStatus: 'Active',
        isActive: true
      }));

    const attendanceYearResponse = options.includeAttendance === false
      ? { success: true, records: [] }
      : clientGetAttendanceDataRange(attendanceRange.yearRange.start, attendanceRange.yearRange.end, resolvedCampaignId || null);

    const yearlyAttendanceRecords = attendanceYearResponse && attendanceYearResponse.success
      ? attendanceYearResponse.records || []
      : [];

    const monthlyAttendanceRecords = yearlyAttendanceRecords.filter(record => {
      if (!record || !record.date) {
        return false;
      }
      return (!attendanceRange.startDate || record.date >= attendanceRange.startDate)
        && (!attendanceRange.endDate || record.date <= attendanceRange.endDate);
    });

    const attendanceDashboard = options.includeAttendanceDashboard === false
      ? null
      : clientGetAttendanceDashboard(attendanceRange.yearRange.start, attendanceRange.yearRange.end, resolvedCampaignId || null);

    const holidayCountry = options.holidayCountry || context.identity?.country || SCHEDULE_SETTINGS.PRIMARY_COUNTRY;
    const holidayYear = options.holidayYear
      || (scheduleRange.startDate ? Number(String(scheduleRange.startDate).slice(0, 4)) : null)
      || new Date().getFullYear();

    const holidays = options.includeHolidays === false
      ? null
      : clientGetCountryHolidays(holidayCountry, holidayYear);

    const combinedUsers = buildUnifiedUserCollection(
      scheduleUsers,
      roster.users,
      options.combinedUsers,
      assignmentUsers,
      attendanceUserRecords
    );

    const managedUserIdSet = new Set();
    const appendManagedUserId = (value) => {
      const normalized = normalizeUserIdValue(value);
      if (normalized) {
        managedUserIdSet.add(normalized);
      }
    };

    (Array.isArray(roster.managedUserIds) ? roster.managedUserIds : []).forEach(appendManagedUserId);
    (Array.isArray(context.managedUserIds) ? context.managedUserIds : []).forEach(appendManagedUserId);
    roster.users.forEach(user => appendManagedUserId(user && (user.ID || user.UserID || user.id || user.userId)));
    scheduleUsers.forEach(user => appendManagedUserId(user && (user.ID || user.UserID || user.id || user.userId)));
    assignmentUsers.forEach(user => appendManagedUserId(user && (user.ID || user.UserID || user.id || user.userId)));

    if (resolvedManagerId) {
      managedUserIdSet.delete(resolvedManagerId);
    }

    const managedUserIds = Array.from(managedUserIdSet);

    const userSources = {
      schedule: scheduleUsers.length,
      roster: roster.users.length,
      assignments: assignmentUsers.length,
      attendance: attendanceUserRecords.length
    };

    return {
      success: true,
      generatedAt: new Date().toISOString(),
      managerId: resolvedManagerId || '',
      campaignId: resolvedCampaignId || '',
      context,
      users: {
        combined: combinedUsers,
        schedule: scheduleUsers,
        roster: roster.users,
        assignments: assignmentUsers,
        attendance: attendanceUserRecords,
        rosterSource: roster.source,
        managedUserIds,
        rosterManagedUserIds: Array.isArray(roster.managedUserIds) ? roster.managedUserIds.slice() : [],
        contextManagedUserIds: Array.isArray(context.managedUserIds) ? context.managedUserIds.slice() : [],
        warnings: roster.warnings,
        sources: userSources
      },
      schedule: {
        range: scheduleRange,
        assignments,
        shiftSlots,
        dashboard
      },
      attendance: {
        range: attendanceRange,
        users: attendanceUsers,
        monthlyRecords: monthlyAttendanceRecords,
        yearlyRecords: yearlyAttendanceRecords,
        dashboard: attendanceDashboard
      },
      holidays
    };
  } catch (error) {
    console.error('Error building unified schedule state:', error);
    safeWriteError && safeWriteError('clientGetScheduleUnifiedState', error);
    return {
      success: false,
      error: error && error.message ? error.message : String(error || 'Unknown error')
    };
  }
}

function applyScenarioAdjustments(bundle, scenario = {}) {
  const volumeMultiplier = scenario.volumeMultiplier || (scenario.volumeDelta ? 1 + scenario.volumeDelta : 1);
  const ahtMultiplier = scenario.ahtMultiplier || (scenario.ahtDelta ? 1 + scenario.ahtDelta : 1);
  const shrinkageDelta = scenario.shrinkageDelta || 0;
  const absenceRate = scenario.absenceRate || 0;
  const additionalOvertimeMinutes = scheduleToNumber(scenario.additionalOvertimeMinutes || scenario.overtimeMinutes, 0);

  const adjustedDemand = (bundle.demandRows || []).map(row => {
    const clone = Object.assign({}, row);
    if (clone.ForecastContacts !== undefined) {
      clone.ForecastContacts = scheduleToNumber(clone.ForecastContacts, 0) * volumeMultiplier;
    }
    if (clone.ForecastAHT !== undefined) {
      clone.ForecastAHT = scheduleToNumber(clone.ForecastAHT, 0) * ahtMultiplier;
    }
    const shrinkage = safeNormalizeSchedulePercentage(clone.Shrinkage, 0.3) + shrinkageDelta;
    clone.Shrinkage = Math.max(0, shrinkage);
    return clone;
  });

  const adjustedSchedules = (bundle.scheduleRows || []).map(row => {
    const clone = Object.assign({}, row);
    const start = safeNormalizeScheduleTimeToMinutes(clone.StartTime || clone.PeriodStart || clone.ScheduleStart || clone.ShiftStart);
    const end = safeNormalizeScheduleTimeToMinutes(clone.EndTime || clone.PeriodEnd || clone.ScheduleEnd || clone.ShiftEnd);
    if (additionalOvertimeMinutes && end !== null) {
      const newEnd = end + additionalOvertimeMinutes;
      clone.EndTime = minutesToTimeString(newEnd);
      clone.PeriodEnd = clone.EndTime;
    }

    if (absenceRate > 0) {
      clone.FTE = scheduleToNumber(clone.FTE || 1, 1) * Math.max(0, 1 - absenceRate);
    }

    return clone;
  });

  return { scheduleRows: adjustedSchedules, demandRows: adjustedDemand };
}

function clientSimulateScheduleScenario(scenario = {}) {
  try {
    console.log('🧪 Simulating schedule scenario:', scenario && scenario.name ? scenario.name : '(ad-hoc scenario)');

    if (typeof evaluateSchedulePerformance !== 'function') {
      throw new Error('Schedule analytics utilities are not available. Ensure ScheduleUtilities is loaded.');
    }

    const context = clientGetScheduleContext(scenario.managerId || scenario.manager || null, scenario.campaignId || scenario.campaign || null);
    if (!context || !context.success) {
      return {
        success: false,
        error: context && context.error ? context.error : 'Unable to resolve schedule context',
        context
      };
    }

    const bundle = loadScheduleDataBundle(context.campaignId || context.providedCampaignId, {
      managedUserIds: scenario.managedUserIds || context.managedUserIds,
      startDate: scenario.startDate,
      endDate: scenario.endDate
    });

    const metricsOptions = Object.assign({ intervalMinutes: scenario.intervalMinutes || 30 }, scenario.metrics || {});

    const baseline = evaluateSchedulePerformance(bundle.scheduleRows, bundle.demandRows, bundle.agentProfiles, metricsOptions);
    const adjusted = applyScenarioAdjustments(bundle, scenario);
    const projection = evaluateSchedulePerformance(adjusted.scheduleRows, adjusted.demandRows, bundle.agentProfiles, metricsOptions);

    return {
      success: true,
      context: {
        managerId: context.managerId,
        campaignId: context.campaignId
      },
      scenario,
      baseline,
      projection,
      delta: {
        serviceLevel: projection.summary.serviceLevel - baseline.summary.serviceLevel,
        healthScore: projection.healthScore - baseline.healthScore,
        compliance: projection.summary.complianceScore - baseline.summary.complianceScore,
        fairness: projection.summary.fairnessIndex - baseline.summary.fairnessIndex
      },
      recommendations: buildScheduleRecommendations(projection, bundle)
    };
  } catch (error) {
    console.error('Error simulating schedule scenario:', error);
    safeWriteError && safeWriteError('clientSimulateScheduleScenario', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function clientGetAgentScheduleSnapshot(agentIdCandidate, startDateCandidate, endDateCandidate, campaignIdCandidate = null, options = {}) {
  try {
    const resolvedAgentId = normalizeUserIdValue(agentIdCandidate) || normalizeUserIdValue(options.agentId);
    const currentUser = typeof getCurrentUser === 'function' ? getCurrentUser() : null;
    const fallbackAgentId = normalizeUserIdValue(currentUser && (currentUser.ID || currentUser.UserID));
    const agentId = resolvedAgentId || fallbackAgentId;

    if (!agentId) {
      return {
        success: false,
        error: 'Agent could not be resolved from parameters or current user context.'
      };
    }

    const context = clientGetScheduleContext(agentId, campaignIdCandidate || options.campaignId || null);
    if (!context || !context.success) {
      return {
        success: false,
        error: context && context.error ? context.error : 'Unable to resolve agent context',
        context
      };
    }

    const startDate = safeNormalizeScheduleDate(startDateCandidate || options.startDate) || new Date();
    const endDate = safeNormalizeScheduleDate(endDateCandidate || options.endDate) || new Date(startDate.getTime() + 14 * 24 * 60 * 60 * 1000);

    const bundle = loadScheduleDataBundle(context.campaignId || context.providedCampaignId, {
      managedUserIds: [agentId],
      startDate,
      endDate
    });

    const agentSchedules = bundle.scheduleRows.filter(row => normalizeUserIdValue(row.UserID || row.UserId || row.AgentID || row.AgentId) === agentId);
    const agentProfile = bundle.agentProfiles.find(profile => normalizeUserIdValue(profile.ID || profile.UserID || profile.UserId) === agentId) || (context.user && normalizeUserIdValue(context.user.ID) === agentId ? context.user : null);

    const now = new Date();
    const upcomingShifts = agentSchedules
      .map(row => {
        const date = safeNormalizeScheduleDate(row.Date || row.ScheduleDate || row.PeriodStart || row.StartDate || row.Day);
        const startMinutes = safeNormalizeScheduleTimeToMinutes(row.StartTime || row.PeriodStart || row.ScheduleStart || row.ShiftStart);
        const startDateTime = date ? combineDateAndMinutes(date, startMinutes || 0) : null;
        return { row, date, startMinutes, startDateTime };
      })
      .filter(item => item.startDateTime && item.startDateTime >= now)
      .sort((a, b) => a.startDateTime - b.startDateTime);

    const historyShifts = agentSchedules
      .map(row => {
        const date = safeNormalizeScheduleDate(row.Date || row.ScheduleDate || row.PeriodStart || row.StartDate || row.Day);
        const startMinutes = safeNormalizeScheduleTimeToMinutes(row.StartTime || row.PeriodStart || row.ScheduleStart || row.ShiftStart);
        const startDateTime = date ? combineDateAndMinutes(date, startMinutes || 0) : null;
        return { row, date, startMinutes, startDateTime };
      })
      .filter(item => item.startDateTime && item.startDateTime < now)
      .sort((a, b) => b.startDateTime - a.startDateTime);

    const totalMinutes = agentSchedules.reduce((sum, row) => {
      const start = safeNormalizeScheduleTimeToMinutes(row.StartTime || row.PeriodStart || row.ScheduleStart || row.ShiftStart);
      const end = safeNormalizeScheduleTimeToMinutes(row.EndTime || row.PeriodEnd || row.ScheduleEnd || row.ShiftEnd);
      if (start === null || end === null) {
        return sum;
      }
      let diff = end - start;
      if (diff < 0) {
        diff += 24 * 60;
      }
      return sum + diff;
    }, 0);

    const weekendShifts = agentSchedules.filter(row => {
      const date = safeNormalizeScheduleDate(row.Date || row.ScheduleDate || row.PeriodStart || row.StartDate || row.Day);
      return date && WEEKEND.includes(date.getDay());
    }).length;

    const nightShifts = agentSchedules.filter(row => {
      const start = safeNormalizeScheduleTimeToMinutes(row.StartTime || row.PeriodStart || row.ScheduleStart || row.ShiftStart);
      const end = safeNormalizeScheduleTimeToMinutes(row.EndTime || row.PeriodEnd || row.ScheduleEnd || row.ShiftEnd);
      return (start !== null && start >= (options.nightThresholdStart || 20 * 60)) || (end !== null && end <= (options.nightThresholdEnd || 6 * 60));
    }).length;

    const averageStartMinutes = agentSchedules.length
      ? agentSchedules.reduce((sum, row) => sum + (safeNormalizeScheduleTimeToMinutes(row.StartTime || row.PeriodStart || row.ScheduleStart || row.ShiftStart) || 0), 0) / agentSchedules.length
      : null;

    let preferenceScore = null;
    let complianceScore = null;
    let fairnessSummary = null;
    if (typeof calculateFairnessMetrics === 'function') {
      const fairness = calculateFairnessMetrics(agentSchedules, agentProfile ? [agentProfile] : [], options.metrics || {});
      fairnessSummary = fairness && fairness.agentSummaries && fairness.agentSummaries.length ? fairness.agentSummaries[0] : null;
      if (fairnessSummary && typeof fairnessSummary.preferenceScore === 'number') {
        preferenceScore = fairnessSummary.preferenceScore;
      }
    }

    if (typeof calculateComplianceMetrics === 'function') {
      const compliance = calculateComplianceMetrics(agentSchedules, agentProfile ? [agentProfile] : [], {
        allowedBreakOverlap: options.allowedBreakOverlap || 3,
        maxHoursPerDay: options.maxHoursPerDay || 12,
        minRestHours: options.minRestHours || 10
      });
      complianceScore = compliance && typeof compliance.complianceScore === 'number' ? compliance.complianceScore : null;
    }

    const nextShift = upcomingShifts.length ? upcomingShifts[0] : null;
    const alerts = [];

    let pendingSwaps = 0;
    if (typeof listShiftSwapRequests === 'function') {
      try {
        const swapRows = listShiftSwapRequests({ userId: agentId });
        pendingSwaps = (swapRows || []).filter(row => {
          const status = String(row.Status || row.status || (typeof SHIFT_SWAP_STATUS !== 'undefined' ? SHIFT_SWAP_STATUS.PENDING : 'PENDING')).toUpperCase();
          return status === (typeof SHIFT_SWAP_STATUS !== 'undefined' ? SHIFT_SWAP_STATUS.PENDING : 'PENDING');
        }).length;
      } catch (swapError) {
        console.warn('clientGetAgentScheduleSnapshot: unable to load swap requests', swapError);
      }
    }

    if (nightShifts >= 3) {
      alerts.push('Multiple night shifts scheduled this period. Ensure adequate rest between shifts.');
    }
    if (weekendShifts >= 3) {
      alerts.push('Heavy weekend coverage detected. Consider requesting swaps if needed.');
    }
    if (complianceScore !== null && complianceScore < 85) {
      alerts.push('Compliance score below target. Review breaks, lunches, and rest periods.');
    }
    if (pendingSwaps > 0) {
      alerts.push(`You have ${pendingSwaps} pending swap request${pendingSwaps === 1 ? '' : 's'}.`);
    }

    const formatShiftOutput = (item) => ({
      id: item.row.ID || item.row.Id || item.row.id || '',
      date: item.date ? formatDateForOutput(item.date) : '',
      dayOfWeek: item.date ? item.date.toLocaleDateString('en-US', { weekday: 'short' }) : '',
      startTime: item.startMinutes !== null && item.startMinutes !== undefined ? minutesToTimeString(item.startMinutes) : '',
      endTime: (() => {
        const end = safeNormalizeScheduleTimeToMinutes(item.row.EndTime || item.row.PeriodEnd || item.row.ScheduleEnd || item.row.ShiftEnd);
        return end !== null && end !== undefined ? minutesToTimeString(end) : '';
      })(),
      location: item.row.Location || '',
      skill: item.row.Skill || item.row.Queue || '',
      status: item.row.Status || item.row.State || '',
      notes: item.row.Notes || ''
    });

    const summary = {
      agentId,
      agentName: (agentProfile && (agentProfile.FullName || agentProfile.UserName || agentProfile.Name)) || (context.user && (context.user.FullName || context.user.UserName)) || '',
      agentEmail: (agentProfile && (agentProfile.Email || agentProfile.email)) || (context.user && (context.user.Email || context.user.email)) || '',
      totalShifts: agentSchedules.length,
      totalScheduledHours: Number((totalMinutes / 60).toFixed(2)),
      weekendShifts,
      nightShifts,
      averageStartTime: averageStartMinutes !== null ? minutesToTimeString(averageStartMinutes) : '',
      preferenceScore,
      complianceScore,
      pendingSwaps,
      upcomingHolidays: 0,
      nextShift: nextShift ? formatShiftOutput(nextShift) : null
    };

    return {
      success: true,
      agentId,
      campaignId: context.campaignId || context.providedCampaignId || '',
      summary,
      upcomingShifts: upcomingShifts.slice(0, options.limitUpcoming || 5).map(formatShiftOutput),
      recentShifts: historyShifts.slice(0, options.limitHistory || 5).map(formatShiftOutput),
      alerts,
      context: {
        permissions: context.permissions,
        managedUserCount: context.managedUserCount
      }
    };
  } catch (error) {
    console.error('Error generating agent schedule snapshot:', error);
    safeWriteError && safeWriteError('clientGetAgentScheduleSnapshot', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function clientGetAgentSchedule(agentIdCandidate, options = {}) {
  try {
    const resolvedAgentId = normalizeUserIdValue(agentIdCandidate) || normalizeUserIdValue(options.agentId);
    const currentUser = typeof getCurrentUser === 'function' ? getCurrentUser() : null;
    const fallbackAgentId = normalizeUserIdValue(currentUser && (currentUser.ID || currentUser.UserID));
    const agentId = resolvedAgentId || fallbackAgentId;

    if (!agentId) {
      return {
        success: false,
        error: 'Agent could not be resolved from parameters or current user context.'
      };
    }

    const context = clientGetScheduleContext(agentId, options.campaignId || null);
    if (!context || !context.success) {
      return {
        success: false,
        error: context && context.error ? context.error : 'Unable to resolve schedule context.'
      };
    }

    const windowOptions = Object.assign({}, options);
    const { agentSchedules, startDate, endDate } = resolveAgentScheduleWindow(agentId, context, windowOptions);

    const schedules = agentSchedules
      .map(mapScheduleRowToAgentShift)
      .filter(Boolean)
      .sort((a, b) => {
        const aKey = typeof a.startTimestamp === 'number' ? a.startTimestamp : Number.MAX_SAFE_INTEGER;
        const bKey = typeof b.startTimestamp === 'number' ? b.startTimestamp : Number.MAX_SAFE_INTEGER;
        return aKey - bKey;
      });

    return {
      success: true,
      agentId,
      campaignId: context.campaignId || context.providedCampaignId || '',
      schedules,
      summary: {
        total: schedules.length,
        startDate: formatDateForOutput(startDate),
        endDate: formatDateForOutput(endDate)
      }
    };
  } catch (error) {
    console.error('Error fetching agent schedule:', error);
    safeWriteError && safeWriteError('clientGetAgentSchedule', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function clientGetAgentUpcomingShifts(agentIdCandidate, options = {}) {
  try {
    const scheduleResponse = clientGetAgentSchedule(agentIdCandidate, Object.assign({}, options, {
      windowDays: options.windowDays || 60
    }));

    if (!scheduleResponse || scheduleResponse.success === false) {
      return scheduleResponse;
    }

    const now = Date.now();
    const limit = Number(options.limit) > 0 ? Number(options.limit) : 10;

    const upcoming = (scheduleResponse.schedules || [])
      .filter(shift => typeof shift.startTimestamp === 'number' ? shift.startTimestamp >= now : true)
      .sort((a, b) => {
        const aKey = typeof a.startTimestamp === 'number' ? a.startTimestamp : Number.MAX_SAFE_INTEGER;
        const bKey = typeof b.startTimestamp === 'number' ? b.startTimestamp : Number.MAX_SAFE_INTEGER;
        return aKey - bKey;
      })
      .slice(0, limit);

    return {
      success: true,
      agentId: scheduleResponse.agentId,
      campaignId: scheduleResponse.campaignId,
      shifts: upcoming
    };
  } catch (error) {
    console.error('Error fetching upcoming shifts:', error);
    safeWriteError && safeWriteError('clientGetAgentUpcomingShifts', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function clientGetAgentSwapRequests(agentIdCandidate, options = {}) {
  try {
    const resolvedAgentId = normalizeUserIdValue(agentIdCandidate) || normalizeUserIdValue(options.agentId);
    const currentUser = typeof getCurrentUser === 'function' ? getCurrentUser() : null;
    const fallbackAgentId = normalizeUserIdValue(currentUser && (currentUser.ID || currentUser.UserID));
    const agentId = resolvedAgentId || fallbackAgentId;

    if (!agentId) {
      return {
        success: false,
        error: 'Agent could not be resolved.'
      };
    }

    const context = clientGetScheduleContext(agentId, options.campaignId || null);
    if (!context || !context.success) {
      return {
        success: false,
        error: context && context.error ? context.error : 'Unable to resolve schedule context.'
      };
    }

    const windowDays = Number(options.windowDays) > 0 ? Number(options.windowDays) : 60;
    const startDate = safeNormalizeScheduleDate(options.startDate) || new Date(Date.now() - 14 * 24 * 60 * 60 * 1000);
    startDate.setHours(0, 0, 0, 0);
    const endDate = safeNormalizeScheduleDate(options.endDate) || new Date(startDate.getTime() + windowDays * 24 * 60 * 60 * 1000);
    endDate.setHours(23, 59, 59, 999);

    const bundle = loadScheduleDataBundle(context.campaignId || context.providedCampaignId, {
      startDate,
      endDate
    });

    const scheduleLookup = buildScheduleRowLookup(bundle.scheduleRows || []);
    const profileLookup = buildAgentProfileLookup(bundle.agentProfiles || []);

    const rawRequests = typeof listShiftSwapRequests === 'function'
      ? listShiftSwapRequests({ userId: agentId })
      : [];

    const formatted = (rawRequests || [])
      .map(row => formatShiftSwapRequestForAgent(row, agentId, { scheduleLookup, profileLookup }))
      .filter(Boolean)
      .sort((a, b) => {
        const aDate = a.requestedAt ? new Date(a.requestedAt).getTime() : 0;
        const bDate = b.requestedAt ? new Date(b.requestedAt).getTime() : 0;
        return bDate - aDate;
      });

    return {
      success: true,
      agentId,
      campaignId: context.campaignId || context.providedCampaignId || '',
      requests: formatted
    };
  } catch (error) {
    console.error('Error loading agent swap requests:', error);
    safeWriteError && safeWriteError('clientGetAgentSwapRequests', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function clientGetAvailableSwapAgents(agentIdCandidate, options = {}) {
  try {
    const resolvedAgentId = normalizeUserIdValue(agentIdCandidate) || normalizeUserIdValue(options.agentId);
    const currentUser = typeof getCurrentUser === 'function' ? getCurrentUser() : null;
    const fallbackAgentId = normalizeUserIdValue(currentUser && (currentUser.ID || currentUser.UserID));
    const agentId = resolvedAgentId || fallbackAgentId;

    if (!agentId) {
      return {
        success: false,
        error: 'Agent could not be resolved.'
      };
    }

    const context = clientGetScheduleContext(agentId, options.campaignId || null);
    if (!context || !context.success) {
      return {
        success: false,
        error: context && context.error ? context.error : 'Unable to resolve schedule context.'
      };
    }

    const scheduleUsers = clientGetScheduleUsers(agentId, context.campaignId || context.providedCampaignId || null);

    const agents = (scheduleUsers || [])
      .map(user => {
        const id = normalizeUserIdValue(user && (user.ID || user.Id || user.UserID || user.UserId));
        if (!id || id === agentId) {
          return null;
        }
        return {
          id,
          name: user.FullName || user.UserName || user.Email || `Agent ${id}`,
          email: user.Email || '',
          team: user.Team || user.Department || ''
        };
      })
      .filter(Boolean)
      .sort((a, b) => a.name.localeCompare(b.name));

    return {
      success: true,
      agents
    };
  } catch (error) {
    console.error('Error loading available swap agents:', error);
    safeWriteError && safeWriteError('clientGetAvailableSwapAgents', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function clientSubmitShiftSwapRequest(agentIdCandidate, request = {}) {
  try {
    const resolvedAgentId = normalizeUserIdValue(agentIdCandidate) || normalizeUserIdValue(request.agentId);
    const currentUser = typeof getCurrentUser === 'function' ? getCurrentUser() : null;
    const fallbackAgentId = normalizeUserIdValue(currentUser && (currentUser.ID || currentUser.UserID));
    const agentId = resolvedAgentId || fallbackAgentId;

    if (!agentId) {
      return {
        success: false,
        error: 'Agent context is required to submit a swap request.'
      };
    }

    const targetId = normalizeUserIdValue(request.swapWith || request.targetUserId);
    if (!targetId) {
      return {
        success: false,
        error: 'Please select an agent to swap with.'
      };
    }

    const myShiftId = String(request.myShiftId || request.requestorScheduleId || '').trim();
    if (!myShiftId) {
      return {
        success: false,
        error: 'Select the shift you would like to swap.'
      };
    }

    const context = clientGetScheduleContext(agentId, request.campaignId || null);
    if (!context || !context.success) {
      return {
        success: false,
        error: context && context.error ? context.error : 'Unable to resolve schedule context.'
      };
    }

    const windowDays = Number(request.windowDays) > 0 ? Number(request.windowDays) : 60;
    const startDate = safeNormalizeScheduleDate(request.startDate) || new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
    startDate.setHours(0, 0, 0, 0);
    const endDate = safeNormalizeScheduleDate(request.endDate) || new Date(startDate.getTime() + windowDays * 24 * 60 * 60 * 1000);
    endDate.setHours(23, 59, 59, 999);

    const bundle = loadScheduleDataBundle(context.campaignId || context.providedCampaignId, {
      startDate,
      endDate
    });

    const scheduleLookup = buildScheduleRowLookup(bundle.scheduleRows || []);
    const profileLookup = buildAgentProfileLookup(bundle.agentProfiles || []);

    const myScheduleRow = scheduleLookup.get(myShiftId);
    if (!myScheduleRow) {
      return {
        success: false,
        error: 'Unable to locate the selected shift. Please refresh and try again.'
      };
    }

    const theirShiftId = String(request.theirShiftId || request.targetScheduleId || '').trim();
    const theirScheduleRow = theirShiftId ? scheduleLookup.get(theirShiftId) : null;

    const requestorProfile = profileLookup.get(agentId) || context.user || {};
    const targetProfile = profileLookup.get(targetId) || null;

    const swapDate = safeNormalizeScheduleDate(request.swapDate)
      || safeNormalizeScheduleDate(myScheduleRow.Date || myScheduleRow.ScheduleDate || myScheduleRow.PeriodStart || myScheduleRow.StartDate || myScheduleRow.Day)
      || new Date();

    const reason = String(request.reason || '').trim();

    const entry = createShiftSwapRequestEntry({
      requestorUserId: agentId,
      requestorUserName: requestorProfile.FullName || requestorProfile.UserName || requestorProfile.Email || 'Agent',
      targetUserId: targetId,
      targetUserName: targetProfile ? (targetProfile.FullName || targetProfile.UserName || targetProfile.Email) : (request.targetUserName || ''),
      requestorScheduleId: myShiftId,
      targetScheduleId: theirShiftId || '',
      swapDate,
      reason,
      status: (typeof SHIFT_SWAP_STATUS !== 'undefined' ? SHIFT_SWAP_STATUS.PENDING : 'PENDING')
    });

    const formatted = formatShiftSwapRequestForAgent(entry, agentId, { scheduleLookup, profileLookup });

    return {
      success: true,
      requestId: entry.ID || entry.Id || entry.id || '',
      request: formatted
    };
  } catch (error) {
    console.error('Error submitting shift swap request:', error);
    safeWriteError && safeWriteError('clientSubmitShiftSwapRequest', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function clientCancelShiftSwapRequest(requestId, agentIdCandidate = null) {
  try {
    const normalizedRequestId = String(requestId || '').trim();
    if (!normalizedRequestId) {
      return {
        success: false,
        error: 'Swap request ID is required.'
      };
    }

    const resolvedAgentId = normalizeUserIdValue(agentIdCandidate);
    const currentUser = typeof getCurrentUser === 'function' ? getCurrentUser() : null;
    const fallbackAgentId = normalizeUserIdValue(currentUser && (currentUser.ID || currentUser.UserID));
    const agentId = resolvedAgentId || fallbackAgentId;

    const requests = typeof listShiftSwapRequests === 'function' ? listShiftSwapRequests() : [];
    const targetRequest = (requests || []).find(row => String(row.ID || row.Id || row.id || '').trim() === normalizedRequestId);

    if (!targetRequest) {
      return {
        success: false,
        error: 'Swap request not found.'
      };
    }

    const requestorId = normalizeUserIdValue(targetRequest.RequestorUserID || targetRequest.RequestorUserId);
    const targetId = normalizeUserIdValue(targetRequest.TargetUserID || targetRequest.TargetUserId);

    if (agentId && agentId !== requestorId && agentId !== targetId) {
      return {
        success: false,
        error: 'You are not authorized to update this swap request.'
      };
    }

    updateShiftSwapRequestEntry(normalizedRequestId, {
      Status: (typeof SHIFT_SWAP_STATUS !== 'undefined' ? SHIFT_SWAP_STATUS.CANCELLED : 'CANCELLED'),
      DecisionNotes: 'Cancelled by agent',
      UpdatedAt: new Date()
    });

    return {
      success: true
    };
  } catch (error) {
    console.error('Error cancelling swap request:', error);
    safeWriteError && safeWriteError('clientCancelShiftSwapRequest', error);
    return {
      success: false,
      error: error.message
    };
  }
}

console.log('✅ Enhanced Schedule Management Backend v4.1 loaded successfully');
console.log('🔧 Features: ScheduleUtilities integration, MainUtilities user management, dedicated spreadsheet support');
console.log('🎯 Ready for production use with comprehensive diagnostics and proper utility integration');
console.log('📊 Integrated: User/Campaign management from MainUtilities, Sheet management from ScheduleUtilities');
