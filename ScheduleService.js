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

const DAY_NAME_TO_INDEX_MAP = {
  sun: 0,
  sunday: 0,
  mon: 1,
  monday: 1,
  tue: 2,
  tues: 2,
  tuesday: 2,
  wed: 3,
  weds: 3,
  wednesday: 3,
  thu: 4,
  thur: 4,
  thurs: 4,
  thursday: 4,
  fri: 5,
  friday: 5,
  sat: 6,
  saturday: 6
};

function convertCsvToDayIndexes(value) {
  if (value === null || typeof value === 'undefined') {
    return [];
  }

  const values = [];

  if (Array.isArray(value)) {
    values.push(...value);
  } else if (typeof value === 'string') {
    const normalized = value
      .replace(/[\[\](){}]/g, ' ')
      .replace(/weekdays?/gi, '1,2,3,4,5')
      .replace(/weekends?/gi, '0,6');

    normalized
      .split(/[^0-9a-zA-Z]+/)
      .map(token => token.trim())
      .filter(Boolean)
      .forEach(token => values.push(token));
  } else if (typeof value === 'number') {
    values.push(value);
  } else if (typeof value === 'object') {
    // If we received an object with a DaysOfWeek or similar property, reuse it.
    const candidate = value.DaysOfWeekArray || value.DaysOfWeek || value.daysOfWeek || value.days;
    if (candidate) {
      return convertCsvToDayIndexes(candidate);
    }
  }

  const indexes = new Set();

  values.forEach(entry => {
    if (entry === null || typeof entry === 'undefined') {
      return;
    }

    let index = null;

    if (typeof entry === 'number' && Number.isFinite(entry)) {
      index = Math.round(entry);
    } else {
      const token = String(entry).trim();
      if (!token) {
        return;
      }

      const numeric = Number(token);
      if (!Number.isNaN(numeric) && Number.isFinite(numeric)) {
        index = Math.round(numeric);
      } else {
        const normalizedToken = token.toLowerCase();

        if (normalizedToken.includes('-')) {
          const [startToken, endToken] = normalizedToken.split('-').map(part => part.trim());
          const startIndex = convertCsvToDayIndexes(startToken)[0];
          const endIndex = convertCsvToDayIndexes(endToken)[0];

          if (Number.isInteger(startIndex) && Number.isInteger(endIndex)) {
            const count = ((endIndex - startIndex + 7) % 7) + 1;
            for (let offset = 0; offset < count; offset++) {
              indexes.add((startIndex + offset) % 7);
            }
            return;
          }
        }

        if (DAY_NAME_TO_INDEX_MAP.hasOwnProperty(normalizedToken)) {
          index = DAY_NAME_TO_INDEX_MAP[normalizedToken];
        }
      }
    }

    if (Number.isInteger(index) && index >= 0 && index <= 6) {
      indexes.add(index);
    }
  });

  return Array.from(indexes).sort((a, b) => a - b);
}

function convertDayIndexesToCsv(days) {
  const normalized = convertCsvToDayIndexes(days);
  return normalized.length ? normalized.join(',') : '';
}

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
    
    // Use MainUtilities function for managed campaigns
    const managedCampaigns = getUserManagedCampaigns(managerId);
    const managedUsers = [];
    
    managedCampaigns.forEach(campaign => {
      const campaignUsers = getUsersByCampaign(campaign.ID);
      campaignUsers.forEach(user => {
        if (String(user.ID) !== String(managerId)) { // Don't include self
          managedUsers.push({
            ID: user.ID,
            UserName: user.UserName,
            FullName: user.FullName,
            Email: user.Email,
            CampaignID: user.CampaignID,
            campaignName: campaign.Name,
            EmploymentStatus: user.EmploymentStatus
          });
        }
      });
    });

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
    console.log('🕒 Creating shift slot (range-based model):', slotData);

    const rawName = slotData && (slotData.name || slotData.slotName || slotData.SlotName);
    const slotName = rawName ? String(rawName).trim() : '';
    const rawCampaign = slotData && (slotData.campaign || slotData.Campaign || slotData.department || slotData.Department || '');
    const campaign = rawCampaign ? String(rawCampaign).trim() : '';
    const rawLocation = slotData && (slotData.location || slotData.Location || 'Office');
    const location = rawLocation ? String(rawLocation).trim() : 'Office';
    const startMinutes = parseTimeToMinutes(slotData && slotData.startTime ? slotData.startTime : slotData && slotData.StartTime);
    const endMinutes = parseTimeToMinutes(slotData && slotData.endTime ? slotData.endTime : slotData && slotData.EndTime);
    const daysArray = normalizeDayTokens(
      (slotData && slotData.daysOfWeek)
        || (slotData && slotData.DaysOfWeek)
        || (slotData && slotData.days)
        || 'Mon,Tue,Wed,Thu,Fri'
    );

    if (!slotName) {
      return { success: false, error: 'Slot name is required.' };
    }
    if (!campaign) {
      return { success: false, error: 'Campaign/Department selection is required for the slot.' };
    }
    if (!Number.isFinite(startMinutes) || !Number.isFinite(endMinutes)) {
      return { success: false, error: 'Start and end times must be valid 12-hour time values.' };
    }
    if (startMinutes === endMinutes) {
      return { success: false, error: 'Start time and end time cannot be the same.' };
    }
    if (!Array.isArray(daysArray) || daysArray.length === 0) {
      return { success: false, error: 'Select at least one day of the week for the shift slot.' };
    }

    const slots = loadShiftSlotsNormalized();
    const normalizedName = slotName.toLowerCase();
    const normalizedCampaign = campaign.toLowerCase();
    const duplicate = slots.find(slot =>
      slot && String(slot.Campaign || '').toLowerCase() === normalizedCampaign && String(slot.SlotName || '').toLowerCase() === normalizedName && slot.Status !== 'Archived'
    );
    if (duplicate) {
      return { success: false, error: 'A shift slot with this name already exists for the selected campaign.' };
    }

    const actor = resolveCurrentActorEmail();
    const now = new Date();
    const slotId = Utilities.getUuid();

    const defaults = {
      priority: Number(slotData && slotData.priority) || 2,
      maxCapacity: Number(slotData && slotData.maxCapacity) || Number(slotData && slotData.MaxCapacity) || null,
      minCoverage: Number(slotData && slotData.minCoverage) || Number(slotData && slotData.MinCoverage) || null,
      breakDuration: Number(slotData && slotData.breakDuration) || Number(slotData && slotData.BreakDuration) || 15,
      break1Duration: Number(slotData && slotData.break1Duration) || Number(slotData && slotData.Break1Duration) || 15,
      break2Duration: Number(slotData && slotData.break2Duration) || Number(slotData && slotData.Break2Duration) || 0,
      lunchDuration: Number(slotData && slotData.lunchDuration) || Number(slotData && slotData.LunchDuration) || 30,
      enableStaggeredBreaks: Boolean(slotData && (slotData.enableStaggeredBreaks ?? slotData.EnableStaggeredBreaks ?? true)),
      breakGroups: Number(slotData && slotData.breakGroups) || Number(slotData && slotData.BreakGroups) || 3,
      staggerInterval: Number(slotData && slotData.staggerInterval) || Number(slotData && slotData.StaggerInterval) || 15,
      minCoveragePct: Number(slotData && slotData.minCoveragePct) || Number(slotData && slotData.MinCoveragePct) || null,
      enableOvertime: Boolean(slotData && (slotData.enableOvertime ?? slotData.EnableOvertime ?? false)),
      maxDailyOT: Number(slotData && slotData.maxDailyOT) || Number(slotData && slotData.MaxDailyOT) || 0,
      maxWeeklyOT: Number(slotData && slotData.maxWeeklyOT) || Number(slotData && slotData.MaxWeeklyOT) || 0,
      otApproval: slotData && (slotData.otApproval || slotData.OTApproval) || '',
      otRate: Number(slotData && slotData.otRate) || Number(slotData && slotData.OTRate) || 1.5,
      otPolicy: slotData && (slotData.otPolicy || slotData.OTPolicy) || '',
      allowSwaps: Boolean(slotData && (slotData.allowSwaps ?? slotData.AllowSwaps ?? true)),
      weekendPremium: Boolean(slotData && (slotData.weekendPremium ?? slotData.WeekendPremium ?? false)),
      holidayPremium: Boolean(slotData && (slotData.holidayPremium ?? slotData.HolidayPremium ?? true)),
      autoAssignment: Boolean(slotData && (slotData.autoAssignment ?? slotData.AutoAssignment ?? false)),
      restPeriod: Number(slotData && slotData.restPeriod) || Number(slotData && slotData.RestPeriod) || 0,
      notificationLead: Number(slotData && slotData.notificationLead) || Number(slotData && slotData.NotificationLead) || 0,
      handoverTime: Number(slotData && slotData.handoverTime) || Number(slotData && slotData.HandoverTime) || 0,
      overtimePolicy: slotData && (slotData.overtimePolicy || slotData.OvertimePolicy) || '',
      location: location
    };

    const normalizedSlot = {
      SlotId: slotId,
      SlotName: slotName,
      Campaign: campaign,
      Location: location,
      StartTime: formatMinutesToTime12Hour(startMinutes),
      EndTime: formatMinutesToTime12Hour(endMinutes),
      StartMinutes: startMinutes,
      EndMinutes: endMinutes,
      DaysOfWeekArray: daysArray,
      DaysCSV: convertDayIndexesToCsv(daysArray),
      Description: slotData && (slotData.description || slotData.Description) || '',
      Notes: slotData && (slotData.notes || slotData.Notes) || '',
      Status: 'Active',
      CreatedAt: now,
      CreatedBy: slotData && (slotData.createdBy || slotData.CreatedBy) || actor,
      UpdatedAt: now,
      UpdatedBy: actor,
      GenerationDefaults: defaults
    };

    const updatedSlots = slots.concat([normalizedSlot]);
    writeShiftSlotsNormalized(updatedSlots);
    logScheduleAudit('CREATE', 'ShiftSlot', slotId, null, normalizedSlot, 'Shift slot created', { campaign, location });

    const responseSlot = normalizeShiftSlotRecord(convertSlotRecordToRow(normalizedSlot));
    responseSlot.GenerationDefaults = defaults;

    return {
      success: true,
      message: 'Shift slot created successfully',
      slot: responseSlot
    };

  } catch (error) {
    console.error('❌ Error creating shift slot:', error);
    try {
      safeWriteError('clientCreateShiftSlot', error);
    } catch (loggingError) {
      console.error('Error logging shift slot failure:', loggingError);
    }
    return {
      success: false,
      error: error && error.message ? error.message : String(error || 'Unknown error')
    };
  }
}

function getDirectManagedUserIds(managerId) {
  const normalizedManagerId = normalizeUserIdValue(managerId);
  const managedUsers = new Set();

  if (!normalizedManagerId) {
    return managedUsers;
  }

  const appendFromRows = (rows) => {
    if (!Array.isArray(rows)) {
      return;
    }

    rows.forEach(row => {
      if (!row || typeof row !== 'object') {
        return;
      }

      const managerCandidates = [
        row.ManagerUserID, row.ManagerUserId, row.managerUserId,
        row.ManagerID, row.ManagerId, row.managerId, row.manager_id,
        row.UserManagerID, row.UserManagerId, row.userManagerId
      ].map(normalizeUserIdValue).filter(Boolean);

      const managedCandidates = [
        row.UserID, row.UserId, row.userId,
        row.ManagedUserID, row.ManagedUserId, row.managedUserId,
        row.ManagedUserID, row.managed_user_id,
        row.ManagedID, row.ManagedId
      ].map(normalizeUserIdValue).filter(Boolean);

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

  return managedUserIds;
}

// ────────────────────────────────────────────────────────────────────────────
// SHIFT & ASSIGNMENT HELPERS (NEW DATA MODEL)
// ────────────────────────────────────────────────────────────────────────────

function ensureShiftSlotSheet_() {
  return ensureScheduleSheetWithHeaders(SHIFT_SLOTS_SHEET, SHIFT_SLOTS_HEADERS);
}

function ensureShiftAssignmentSheet_() {
  return ensureScheduleSheetWithHeaders(SHIFT_ASSIGNMENTS_SHEET, SHIFT_ASSIGNMENTS_HEADERS);
}

function parseSlotGenerationDefaults_(rawValue) {
  if (!rawValue) {
    return {};
  }

  if (typeof rawValue === 'object') {
    return rawValue || {};
  }

  if (typeof rawValue === 'string') {
    const trimmed = rawValue.trim();
    if (!trimmed) {
      return {};
    }
    try {
      return JSON.parse(trimmed) || {};
    } catch (error) {
      console.warn('Unable to parse slot defaults JSON:', error);
    }
  }

  return {};
}

function serializeSlotGenerationDefaults_(defaults) {
  if (!defaults || typeof defaults !== 'object') {
    return '';
  }
  try {
    return JSON.stringify(defaults);
  } catch (error) {
    console.warn('Unable to serialize generation defaults:', error);
    return '';
  }
}

function normalizeSlotStatus_(status, fallbackActive = true) {
  if (status === null || typeof status === 'undefined') {
    return fallbackActive ? 'Active' : 'Archived';
  }

  if (typeof status === 'boolean') {
    return status ? 'Active' : 'Archived';
  }

  const normalized = String(status).trim().toLowerCase();
  if (!normalized) {
    return fallbackActive ? 'Active' : 'Archived';
  }

  if (['archived', 'inactive', 'disabled'].includes(normalized)) {
    return 'Archived';
  }

  return 'Active';
}

function normalizeShiftSlotRecord(row) {
  if (!row || typeof row !== 'object') {
    return null;
  }

  const slotId = row.SlotId || row.SlotID || row.ID || row.Id || row.id || '';
  const startTime = row.StartTime || row.Start || row.Start_Time || '';
  const endTime = row.EndTime || row.End || row.End_Time || '';
  const daysCsv = row.DaysCSV || row.DaysOfWeek || row.Days || '';
  const defaults = parseSlotGenerationDefaults_(row.GenerationDefaults);
  const daysOfWeekArray = convertCsvToDayIndexes(daysCsv);
  const startMinutes = parseTimeToMinutes(startTime);
  const endMinutes = parseTimeToMinutes(endTime);
  const status = normalizeSlotStatus_(row.Status !== undefined ? row.Status : row.IsActive !== false);

  const slot = {
    ID: slotId,
    SlotId: slotId,
    SlotName: row.SlotName || row.Name || '',
    Name: row.SlotName || row.Name || '',
    Campaign: row.Campaign || row.Department || '',
    Department: row.Campaign || row.Department || '',
    Location: row.Location || defaults.location || 'Office',
    StartTime: startTime || (Number.isFinite(startMinutes) ? formatMinutesToTime12Hour(startMinutes) : ''),
    EndTime: endTime || (Number.isFinite(endMinutes) ? formatMinutesToTime12Hour(endMinutes) : ''),
    StartMinutes: Number.isFinite(startMinutes) ? startMinutes : null,
    EndMinutes: Number.isFinite(endMinutes) ? endMinutes : null,
    DaysCSV: convertDayIndexesToCsv(daysOfWeekArray),
    DaysOfWeek: convertDayIndexesToCsv(daysOfWeekArray),
    DaysOfWeekArray: daysOfWeekArray,
    Description: row.Description || '',
    Notes: row.Notes || '',
    Status: status,
    IsActive: status !== 'Archived',
    CreatedAt: row.CreatedAt || '',
    CreatedBy: row.CreatedBy || '',
    UpdatedAt: row.UpdatedAt || '',
    UpdatedBy: row.UpdatedBy || '',
    GenerationDefaults: defaults,
    Priority: defaults.priority,
    MaxCapacity: defaults.maxCapacity,
    MinCoverage: defaults.minCoverage,
    BreakDuration: defaults.breakDuration,
    Break1Duration: defaults.break1Duration,
    Break2Duration: defaults.break2Duration,
    LunchDuration: defaults.lunchDuration,
    EnableStaggeredBreaks: defaults.enableStaggeredBreaks,
    BreakGroups: defaults.breakGroups,
    StaggerInterval: defaults.staggerInterval,
    MinCoveragePct: defaults.minCoveragePct,
    EnableOvertime: defaults.enableOvertime,
    MaxDailyOT: defaults.maxDailyOT,
    MaxWeeklyOT: defaults.maxWeeklyOT,
    OTApproval: defaults.otApproval,
    OTRate: defaults.otRate,
    OTPolicy: defaults.otPolicy,
    AllowSwaps: defaults.allowSwaps,
    WeekendPremium: defaults.weekendPremium,
    HolidayPremium: defaults.holidayPremium,
    AutoAssignment: defaults.autoAssignment,
    RestPeriod: defaults.restPeriod,
    NotificationLead: defaults.notificationLead,
    HandoverTime: defaults.handoverTime,
    OvertimePolicy: defaults.overtimePolicy
  };

  return slot;
}

function convertSlotRecordToRow(slot) {
  if (!slot || typeof slot !== 'object') {
    return null;
  }

  const defaults = Object.assign({}, slot.GenerationDefaults || {});

  const headerRow = {
    SlotId: slot.SlotId || slot.ID || Utilities.getUuid(),
    SlotName: slot.SlotName || slot.Name || '',
    Campaign: slot.Campaign || slot.Department || '',
    Location: slot.Location || defaults.location || 'Office',
    StartTime: slot.StartTime || (Number.isFinite(slot.StartMinutes) ? formatMinutesToTime12Hour(slot.StartMinutes) : ''),
    EndTime: slot.EndTime || (Number.isFinite(slot.EndMinutes) ? formatMinutesToTime12Hour(slot.EndMinutes) : ''),
    DaysCSV: convertDayIndexesToCsv(slot.DaysOfWeekArray || slot.DaysOfWeek || slot.DaysCSV || []),
    Description: slot.Description || '',
    Notes: slot.Notes || '',
    Status: normalizeSlotStatus_(slot.Status, slot.IsActive !== false),
    CreatedAt: slot.CreatedAt || '',
    CreatedBy: slot.CreatedBy || '',
    UpdatedAt: slot.UpdatedAt || '',
    UpdatedBy: slot.UpdatedBy || '',
    GenerationDefaults: serializeSlotGenerationDefaults_(defaults)
  };

  return headerRow;
}

function loadShiftSlotsNormalized() {
  const rawRows = readScheduleSheet(SHIFT_SLOTS_SHEET) || [];
  if (!rawRows.length) {
    return [];
  }

  return rawRows
    .map(normalizeShiftSlotRecord)
    .filter(Boolean)
    .map(slot => {
      if (!slot.CreatedAt) {
        slot.CreatedAt = new Date();
      }
      return slot;
    });
}

function writeShiftSlotsNormalized(slots) {
  const records = Array.isArray(slots) ? slots : [];
  const rows = records
    .map(convertSlotRecordToRow)
    .filter(Boolean);
  writeToScheduleSheet(SHIFT_SLOTS_SHEET, rows);
  invalidateScheduleCaches();
}

function buildSlotLookupById(slots) {
  const lookup = new Map();
  (slots || []).forEach(slot => {
    if (!slot || !slot.SlotId) {
      return;
    }
    lookup.set(String(slot.SlotId), slot);
  });
  return lookup;
}

function buildUserLookupById_() {
  const users = readSheet(USERS_SHEET) || [];
  const map = new Map();
  users.forEach(user => {
    if (!user || !user.ID) {
      return;
    }
    map.set(String(user.ID), user);
  });
  return map;
}

function buildUserLookupByName_() {
  const users = readSheet(USERS_SHEET) || [];
  const map = new Map();
  users.forEach(user => {
    if (!user) {
      return;
    }
    const candidates = [user.UserName, user.Username, user.FullName, user.Email]
      .map(value => normalizeUserKey(value))
      .filter(Boolean);
    candidates.forEach(key => {
      if (!map.has(key)) {
        map.set(key, user);
      }
    });
  });
  return map;
}

function normalizeAssignmentStatus_(status) {
  if (!status) {
    return 'PENDING';
  }
  const normalized = String(status).trim().toUpperCase();
  if (['APPROVED', 'REJECTED', 'ARCHIVED', 'PENDING'].includes(normalized)) {
    return normalized;
  }
  return 'PENDING';
}

function parsePremiumCsv_(value) {
  if (!value) {
    return [];
  }
  if (Array.isArray(value)) {
    return value.map(v => String(v).trim()).filter(Boolean);
  }
  if (typeof value === 'string') {
    return value.split(/[\s,;|]+/g).map(v => v.trim()).filter(Boolean);
  }
  return [];
}

function serializePremiumCsv_(premiums) {
  if (!premiums || !premiums.length) {
    return '';
  }
  const unique = Array.from(new Set(premiums.map(p => String(p).trim()).filter(Boolean)));
  return unique.join(',');
}

function normalizeShiftAssignmentRecord(row, slotLookup, userLookupById, userLookupByName) {
  if (!row || typeof row !== 'object') {
    return null;
  }

  const assignmentId = row.AssignmentId || row.ID || row.Id || row.id || '';
  const slotId = row.SlotId || row.SlotID || row.slotId || '';
  const slot = slotLookup.get(String(slotId)) || null;
  const userId = row.UserId || row.UserID || row.userId || row.ID_User || '';
  const status = normalizeAssignmentStatus_(row.Status);
  const premiums = parsePremiumCsv_(row.Premiums);
  const breaksConfig = parseSlotGenerationDefaults_(row.BreaksConfigJSON || row.BreaksConfig || row.BreakConfiguration);
  const startDate = normalizeDateForSheet(row.StartDate || row.PeriodStart || row.Date || row.ScheduleStart, DEFAULT_SCHEDULE_TIME_ZONE);
  const endDate = normalizeDateForSheet(row.EndDate || row.PeriodEnd || row.ScheduleEnd || row.AssignmentEnd || row.End, DEFAULT_SCHEDULE_TIME_ZONE) || startDate;

  let userRecord = userLookupById.get(String(userId));
  if (!userRecord && row.UserName) {
    const nameKey = normalizeUserKey(row.UserName);
    if (nameKey && userLookupByName.has(nameKey)) {
      userRecord = userLookupByName.get(nameKey);
    }
  }

  const assignment = {
    AssignmentId: assignmentId,
    ID: assignmentId,
    SlotId: slotId,
    SlotName: slot ? slot.SlotName : (row.SlotName || ''),
    UserId: userId,
    UserName: row.UserName || (userRecord && (userRecord.UserName || userRecord.FullName)) || '',
    Campaign: row.Campaign || (slot && slot.Campaign) || '',
    StartDate: startDate,
    EndDate: endDate,
    Status: status,
    AllowSwap: row.AllowSwap === true || String(row.AllowSwap || '').toLowerCase() === 'true',
    Premiums: premiums,
    BreaksConfigJSON: breaksConfig,
    OvertimeMinutes: Number(row.OvertimeMinutes) || 0,
    RestPeriodHours: Number(row.RestPeriodHours) || 0,
    NotificationLeadHours: Number(row.NotificationLeadHours) || 0,
    HandoverMinutes: Number(row.HandoverMinutes) || 0,
    Notes: row.Notes || '',
    CreatedAt: row.CreatedAt || '',
    CreatedBy: row.CreatedBy || '',
    UpdatedAt: row.UpdatedAt || '',
    UpdatedBy: row.UpdatedBy || '',
    RollbackGroupId: row.RollbackGroupId || '',
    Slot: slot || null,
    UserRecord: userRecord || null
  };

  if (slot) {
    assignment.Location = slot.Location;
    assignment.StartTime = slot.StartTime;
    assignment.EndTime = slot.EndTime;
    assignment.DaysCSV = slot.DaysCSV;
    assignment.DaysOfWeekArray = slot.DaysOfWeekArray;
  }

  return assignment;
}

function loadShiftAssignmentsNormalized() {
  const slotList = loadShiftSlotsNormalized();
  const slotLookup = buildSlotLookupById(slotList);
  const usersById = buildUserLookupById_();
  const usersByName = buildUserLookupByName_();
  const rows = readScheduleSheet(SHIFT_ASSIGNMENTS_SHEET) || [];
  if (!rows.length) {
    return [];
  }

  return rows
    .map(row => normalizeShiftAssignmentRecord(row, slotLookup, usersById, usersByName))
    .filter(Boolean);
}

function serializeAssignmentRecord_(assignment) {
  if (!assignment || typeof assignment !== 'object') {
    return null;
  }

  const row = {
    AssignmentId: assignment.AssignmentId || assignment.ID || Utilities.getUuid(),
    UserId: assignment.UserId || (assignment.UserRecord && assignment.UserRecord.ID) || '',
    Campaign: assignment.Campaign || (assignment.Slot && assignment.Slot.Campaign) || '',
    SlotId: assignment.SlotId || (assignment.Slot && assignment.Slot.SlotId) || '',
    StartDate: assignment.StartDate || '',
    EndDate: assignment.EndDate || assignment.StartDate || '',
    Status: normalizeAssignmentStatus_(assignment.Status),
    AllowSwap: assignment.AllowSwap ? true : false,
    Premiums: serializePremiumCsv_(assignment.Premiums),
    BreaksConfigJSON: serializeSlotGenerationDefaults_(assignment.BreaksConfigJSON || assignment.breaksConfiguration || {}),
    OvertimeMinutes: Number(assignment.OvertimeMinutes) || 0,
    RestPeriodHours: Number(assignment.RestPeriodHours) || 0,
    NotificationLeadHours: Number(assignment.NotificationLeadHours) || 0,
    HandoverMinutes: Number(assignment.HandoverMinutes) || 0,
    Notes: assignment.Notes || '',
    CreatedAt: assignment.CreatedAt || new Date(),
    CreatedBy: assignment.CreatedBy || resolveCurrentActorEmail(),
    UpdatedAt: assignment.UpdatedAt || new Date(),
    UpdatedBy: assignment.UpdatedBy || resolveCurrentActorEmail(),
    RollbackGroupId: assignment.RollbackGroupId || ''
  };

  return row;
}

function writeShiftAssignments(assignments) {
  const rows = (assignments || [])
    .map(serializeAssignmentRecord_)
    .filter(Boolean);
  writeToScheduleSheet(SHIFT_ASSIGNMENTS_SHEET, rows);
  invalidateScheduleCaches();
}

function appendShiftAssignments(newAssignments) {
  const sheet = ensureShiftAssignmentSheet_();
  const rows = (newAssignments || []).map(serializeAssignmentRecord_).filter(Boolean);
  if (!rows.length) {
    return;
  }
  const headerCount = SHIFT_ASSIGNMENTS_HEADERS.length;
  const values = rows.map(row => SHIFT_ASSIGNMENTS_HEADERS.map(header => row[header] !== undefined ? row[header] : ''));
  sheet.insertRowsAfter(sheet.getLastRow() || 1, values.length);
  sheet.getRange(sheet.getLastRow() - values.length + 1, 1, values.length, headerCount).setValues(values);
  invalidateScheduleCaches();
}

function updateAssignmentsByIds_(assignmentIds, updater) {
  const ids = Array.isArray(assignmentIds) ? assignmentIds.filter(Boolean).map(String) : [];
  if (!ids.length) {
    return { updated: [], assignments: [] };
  }

  const idSet = new Set(ids);
  const assignments = loadShiftAssignmentsNormalized();
  const updated = [];
  let changed = false;

  assignments.forEach(assignment => {
    if (!assignment || !idSet.has(String(assignment.AssignmentId))) {
      return;
    }
    const before = Object.assign({}, assignment);
    updater(assignment, before);
    updated.push({ before, after: assignment });
    changed = true;
  });

  if (changed) {
    writeShiftAssignments(assignments);
  }

  return { updated, assignments };
}

function normalizeCoverageTarget_(options = {}, slotDefaults = {}) {
  const capacity = Number(options.capacity && options.capacity.max);
  const slotCapacity = Number(slotDefaults.maxCapacity);
  let maxCapacity = Number.isFinite(capacity) && capacity > 0 ? capacity : null;
  if (Number.isFinite(slotCapacity) && slotCapacity > 0) {
    maxCapacity = maxCapacity ? Math.min(maxCapacity, slotCapacity) : slotCapacity;
  }

  const minCoverage = Number(options.capacity && options.capacity.min);
  const slotMinCoverage = Number(slotDefaults.minCoverage);
  let minimum = Number.isFinite(minCoverage) && minCoverage > 0 ? minCoverage : null;
  if (Number.isFinite(slotMinCoverage) && slotMinCoverage > 0) {
    minimum = minimum ? Math.max(minimum, slotMinCoverage) : slotMinCoverage;
  }

  const minCoveragePct = Number(options.breaks && options.breaks.minCoveragePct);
  const slotMinCoveragePct = Number(slotDefaults.minCoveragePct);
  let minimumPct = Number.isFinite(minCoveragePct) ? minCoveragePct : null;
  if (Number.isFinite(slotMinCoveragePct)) {
    minimumPct = Number.isFinite(minimumPct) ? Math.max(minimumPct, slotMinCoveragePct) : slotMinCoveragePct;
  }

  return {
    maxCapacity: maxCapacity,
    minCoverage: minimum,
    minCoveragePct: minimumPct
  };
}

function computeMinimumCoverageTarget_(target) {
  if (!target) {
    return 0;
  }

  const minAbsolute = Number(target.minCoverage);
  const capacity = Number(target.maxCapacity);
  const minPct = Number(target.minCoveragePct);

  let minimum = Number.isFinite(minAbsolute) ? minAbsolute : 0;
  if (Number.isFinite(minPct) && Number.isFinite(capacity) && capacity > 0) {
    const pctRequirement = Math.ceil((minPct / 100) * capacity);
    minimum = Math.max(minimum, pctRequirement);
  }
  return minimum;
}

function buildHolidayLookup_(startDate, endDate) {
  const holidays = readScheduleSheet(HOLIDAYS_SHEET) || [];
  if (!holidays.length) {
    return new Map();
  }

  const lookup = new Map();
  const start = startDate ? new Date(startDate) : null;
  const end = endDate ? new Date(endDate) : null;

  holidays.forEach(row => {
    if (!row || !row.Date) {
      return;
    }
    const dateStr = normalizeDateForSheet(row.Date, DEFAULT_SCHEDULE_TIME_ZONE);
    if (!dateStr) {
      return;
    }
    const dateObj = new Date(dateStr);
    if ((start && dateObj < start) || (end && dateObj > end)) {
      return;
    }
    if (!lookup.has(dateStr)) {
      lookup.set(dateStr, []);
    }
    lookup.get(dateStr).push({
      date: dateStr,
      name: row.Name || row.HolidayName || '',
      region: row.Region || 'Jamaica',
      isWorkingDayOverride: row.IsWorkingDayOverride === true || String(row.IsWorkingDayOverride || '').toLowerCase() === 'true',
      notes: row.Notes || ''
    });
  });

  return lookup;
}

function isWeekendDay_(dateStr) {
  if (!dateStr) {
    return false;
  }
  const date = new Date(dateStr + 'T00:00:00');
  const day = date.getDay();
  return day === 0 || day === 6;
}

function expandAssignmentToDays_(assignment, slot) {
  const startDate = new Date(assignment.StartDate + 'T00:00:00');
  const endDate = new Date(assignment.EndDate + 'T00:00:00');
  const days = [];
  if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
    return days;
  }
  const slotDays = slot && Array.isArray(slot.DaysOfWeekArray) ? slot.DaysOfWeekArray : [];
  const cursor = new Date(startDate.getTime());
  while (cursor <= endDate) {
    const dayIndex = cursor.getDay();
    if (slotDays.length === 0 || slotDays.includes(dayIndex)) {
      const iso = Utilities.formatDate(cursor, DEFAULT_SCHEDULE_TIME_ZONE, 'yyyy-MM-dd');
      days.push(iso);
    }
    cursor.setDate(cursor.getDate() + 1);
  }
  return days;
}

function buildAssignmentCoverage_(assignments, slotLookup, startDate, endDate) {
  const coverage = new Map();
  (assignments || []).forEach(assignment => {
    const slot = slotLookup.get(String(assignment.SlotId));
    const days = expandAssignmentToDays_(assignment, slot);
    days.forEach(dateStr => {
      if (!coverage.has(dateStr)) {
        coverage.set(dateStr, { count: 0, bySlot: new Map() });
      }
      const entry = coverage.get(dateStr);
      entry.count += 1;
      const slotCount = entry.bySlot.get(assignment.SlotId) || 0;
      entry.bySlot.set(assignment.SlotId, slotCount + 1);
    });
  });
  return coverage;
}

function collectAttendanceConflicts_(users, startDate, endDate) {
  try {
    if (typeof clientGetAttendanceDataRange !== 'function') {
      return new Map();
    }
    const response = clientGetAttendanceDataRange(startDate, endDate, null);
    if (!response || !response.success) {
      return new Map();
    }
    const conflictStatuses = new Set(['PTO', 'Leave', 'Vacation', 'Sick Leave', 'Holiday', 'Absent']);
    const map = new Map();
    (response.records || []).forEach(record => {
      const userName = record.User || record.UserName || record.Agent;
      const normalizedUser = normalizeUserKey(userName);
      if (!normalizedUser || !conflictStatuses.has(String(record.State || record.Status || '').trim())) {
        return;
      }
      if (!map.has(normalizedUser)) {
        map.set(normalizedUser, []);
      }
      map.get(normalizedUser).push(record);
    });
    return map;
  } catch (error) {
    console.warn('Unable to collect attendance conflicts:', error);
    return new Map();
  }
}

function computeRestPeriodViolation_(existingAssignments, userId, startDate, restPeriodHours) {
  if (!restPeriodHours || restPeriodHours <= 0) {
    return null;
  }
  const restMs = restPeriodHours * 60 * 60 * 1000;
  const start = new Date(startDate + 'T00:00:00');
  if (isNaN(start.getTime())) {
    return null;
  }
  const windowStart = new Date(start.getTime() - restMs);
  const conflicts = [];
  existingAssignments.forEach(assignment => {
    if (!assignment || assignment.UserId !== userId) {
      return;
    }
    const assignmentEnd = new Date(assignment.EndDate + 'T23:59:59');
    if (assignmentEnd >= windowStart && assignmentEnd <= start) {
      conflicts.push(assignment);
    }
  });
  if (!conflicts.length) {
    return null;
  }
  return conflicts[0];
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
    console.log('📊 Getting all shift slots (range-based model)');

    let slots = loadShiftSlotsNormalized();
    const sources = new Set(['ShiftSlots']);

    if (!slots.length) {
      console.warn('No modern shift slots found. Attempting legacy conversion.');
      const legacySheets = ['Shift Slots', 'Shift Slot', 'ShiftTemplates', 'Shift Template', 'Shifts'];
      const legacyRecords = [];
      legacySheets.forEach(sheetName => {
        try {
          const rows = readScheduleSheet(sheetName) || [];
          if (rows.length) {
            rows
              .map(convertLegacyShiftSlotRecord)
              .filter(Boolean)
              .forEach(slot => legacyRecords.push(slot));
            sources.add(sheetName);
          }
        } catch (error) {
          console.warn(`Unable to read legacy shift slots from ${sheetName}:`, error && error.message ? error.message : error);
        }
      });

      if (legacyRecords.length) {
        slots = legacyRecords.map(slot => normalizeShiftSlotRecord(convertSlotRecordToRow(slot))).filter(Boolean);
        writeShiftSlotsNormalized(slots);
      }
    }

    const normalizedSlots = slots.map(slot => {
      const enriched = Object.assign({}, slot);
      enriched.ID = slot.SlotId;
      enriched.Name = slot.SlotName;
      enriched.Department = slot.Campaign;
      enriched.DaysOfWeek = slot.DaysCSV;
      enriched.DaysOfWeekArray = slot.DaysOfWeekArray || convertCsvToDayIndexes(slot.DaysCSV);
      if (!enriched.GenerationDefaults || typeof enriched.GenerationDefaults !== 'object') {
        enriched.GenerationDefaults = {};
      }
      enriched.MaxCapacity = enriched.GenerationDefaults.maxCapacity;
      enriched.MinCoverage = enriched.GenerationDefaults.minCoverage;
      enriched.Priority = enriched.GenerationDefaults.priority;
      enriched.BreakDuration = enriched.GenerationDefaults.breakDuration;
      enriched.Break1Duration = enriched.GenerationDefaults.break1Duration;
      enriched.Break2Duration = enriched.GenerationDefaults.break2Duration;
      enriched.LunchDuration = enriched.GenerationDefaults.lunchDuration;
      enriched.EnableStaggeredBreaks = enriched.GenerationDefaults.enableStaggeredBreaks;
      enriched.BreakGroups = enriched.GenerationDefaults.breakGroups;
      enriched.StaggerInterval = enriched.GenerationDefaults.staggerInterval;
      enriched.MinCoveragePct = enriched.GenerationDefaults.minCoveragePct;
      enriched.EnableOvertime = enriched.GenerationDefaults.enableOvertime;
      enriched.MaxDailyOT = enriched.GenerationDefaults.maxDailyOT;
      enriched.MaxWeeklyOT = enriched.GenerationDefaults.maxWeeklyOT;
      enriched.OTApproval = enriched.GenerationDefaults.otApproval;
      enriched.OTRate = enriched.GenerationDefaults.otRate;
      enriched.OTPolicy = enriched.GenerationDefaults.otPolicy;
      enriched.AllowSwaps = enriched.GenerationDefaults.allowSwaps;
      enriched.WeekendPremium = enriched.GenerationDefaults.weekendPremium;
      enriched.HolidayPremium = enriched.GenerationDefaults.holidayPremium;
      enriched.AutoAssignment = enriched.GenerationDefaults.autoAssignment;
      enriched.RestPeriod = enriched.GenerationDefaults.restPeriod;
      enriched.NotificationLead = enriched.GenerationDefaults.notificationLead;
      enriched.HandoverTime = enriched.GenerationDefaults.handoverTime;
      enriched.OvertimePolicy = enriched.GenerationDefaults.overtimePolicy;
      enriched.__sources = Array.from(sources);
      return enriched;
    });

    console.log(`📦 Loaded ${normalizedSlots.length} shift slots.`);

    return {
      success: true,
      slots: normalizedSlots,
      total: normalizedSlots.length,
      metadata: {
        sources: Array.from(sources),
        totalCount: normalizedSlots.length
      }
    };

  } catch (error) {
    console.error('❌ Error getting shift slots:', error);
    safeWriteError('clientGetAllShiftSlots', error);
    return {
      success: false,
      error: error.message || 'Failed to load shift slots',
      slots: [],
      total: 0,
      metadata: { sources: [] }
    };
  }
}


// ────────────────────────────────────────────────────────────────────────────
// SCHEDULE GENERATION - Enhanced with ScheduleUtilities integration
// ────────────────────────────────────────────────────────────────────────────

function coerceNumberOrNull(value) {
  if (value === null || typeof value === 'undefined') {
    return null;
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) {
      return null;
    }
    const parsed = Number(trimmed);
    return Number.isFinite(parsed) ? parsed : null;
  }

  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }

  return null;
}

function coerceBooleanOrNull(value) {
  if (value === null || typeof value === 'undefined') {
    return null;
  }

  if (typeof value === 'boolean') {
    return value;
  }

  if (typeof value === 'number') {
    return value !== 0;
  }

  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (!normalized) {
      return null;
    }

    if (['true', '1', 'yes', 'y', 'on'].includes(normalized)) {
      return true;
    }

    if (['false', '0', 'no', 'n', 'off'].includes(normalized)) {
      return false;
    }
  }

  return null;
}

function resolveNumberFromSource(source, keys) {
  if (!source || typeof source !== 'object') {
    return null;
  }

  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (!Object.prototype.hasOwnProperty.call(source, key)) {
      continue;
    }

    const resolved = coerceNumberOrNull(source[key]);
    if (resolved !== null) {
      return resolved;
    }
  }

  return null;
}

function resolveBooleanFromSource(source, keys) {
  if (!source || typeof source !== 'object') {
    return null;
  }

  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (!Object.prototype.hasOwnProperty.call(source, key)) {
      continue;
    }

    const resolved = coerceBooleanOrNull(source[key]);
    if (resolved !== null) {
      return resolved;
    }
  }

  return null;
}

function resolveStringFromSource(source, keys) {
  if (!source || typeof source !== 'object') {
    return null;
  }

  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (!Object.prototype.hasOwnProperty.call(source, key)) {
      continue;
    }

    const value = source[key];
    if (value === null || typeof value === 'undefined') {
      continue;
    }

    const trimmed = String(value).trim();
    if (trimmed) {
      return trimmed;
    }
  }

  return null;
}

function sanitizeGenerationConfig(value) {
  if (value === null || typeof value === 'undefined') {
    return undefined;
  }

  if (Array.isArray(value)) {
    const sanitizedArray = value
      .map(entry => sanitizeGenerationConfig(entry))
      .filter(entry => entry !== undefined);
    return sanitizedArray.length ? sanitizedArray : undefined;
  }

  if (typeof value === 'object') {
    const result = {};
    Object.keys(value).forEach(key => {
      const sanitized = sanitizeGenerationConfig(value[key]);
      if (sanitized !== undefined) {
        result[key] = sanitized;
      }
    });
    return Object.keys(result).length ? result : undefined;
  }

  return value;
}

function buildGenerationDefaultsFromSource(source = {}) {
  if (!source || typeof source !== 'object') {
    return null;
  }

  const defaults = {
    capacity: {
      max: resolveNumberFromSource(source, ['MaxCapacity', 'maxCapacity', 'capacityMax']),
      min: resolveNumberFromSource(source, ['MinCoverage', 'minCoverage', 'capacityMin'])
    },
    breaks: {
      first: resolveNumberFromSource(source, ['Break1Duration', 'break1Duration', 'BreakDuration', 'breakDuration']),
      second: resolveNumberFromSource(source, ['Break2Duration', 'break2Duration']),
      lunch: resolveNumberFromSource(source, ['LunchDuration', 'lunchDuration']),
      enableStaggered: resolveBooleanFromSource(source, ['EnableStaggeredBreaks', 'enableStaggeredBreaks']),
      groups: resolveNumberFromSource(source, ['BreakGroups', 'breakGroups']),
      interval: resolveNumberFromSource(source, ['StaggerInterval', 'staggerInterval']),
      minCoveragePct: resolveNumberFromSource(source, ['MinCoveragePct', 'minCoveragePct'])
    },
    overtime: {
      enabled: resolveBooleanFromSource(source, ['EnableOvertime', 'enableOvertime', 'overtimeEnabled']),
      maxDaily: resolveNumberFromSource(source, ['MaxDailyOT', 'maxDailyOT']),
      maxWeekly: resolveNumberFromSource(source, ['MaxWeeklyOT', 'maxWeeklyOT']),
      approval: resolveStringFromSource(source, ['OTApproval', 'otApproval']),
      rate: resolveNumberFromSource(source, ['OTRate', 'otRate']),
      policy: resolveStringFromSource(source, ['OTPolicy', 'otPolicy', 'OvertimePolicy', 'overtimePolicy'])
    },
    advanced: {
      allowSwaps: resolveBooleanFromSource(source, ['AllowSwaps', 'allowSwaps']),
      weekendPremium: resolveBooleanFromSource(source, ['WeekendPremium', 'weekendPremium']),
      holidayPremium: resolveBooleanFromSource(source, ['HolidayPremium', 'holidayPremium']),
      autoAssignment: resolveBooleanFromSource(source, ['AutoAssignment', 'autoAssignment']),
      restPeriod: resolveNumberFromSource(source, ['RestPeriod', 'restPeriod']),
      notificationLead: resolveNumberFromSource(source, ['NotificationLead', 'notificationLead']),
      handoverTime: resolveNumberFromSource(source, ['HandoverTime', 'handoverTime'])
    }
  };

  const sanitized = sanitizeGenerationConfig(defaults);
  return sanitized || null;
}

function parseGenerationDefaults(value, contextLabel) {
  if (!value) {
    return null;
  }

  if (typeof value === 'object') {
    const sanitized = sanitizeGenerationConfig(value);
    return sanitized || null;
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) {
      return null;
    }

    try {
      const parsed = JSON.parse(trimmed);
      const sanitized = sanitizeGenerationConfig(parsed);
      return sanitized || null;
    } catch (error) {
      console.warn('Unable to parse generation defaults JSON' + (contextLabel ? ` for ${contextLabel}` : '') + ':', error);
      return null;
    }
  }

  return null;
}

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
  const applyRules = (rules) => {
    if (!rules || typeof rules !== 'object') {
      return;
    }

    const { capacity, breaks, overtime, advanced } = rules;

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
  };

  const slotDefaults = parseGenerationDefaults(slot && slot.GenerationDefaults, slot && slot.ID) || buildGenerationDefaultsFromSource(slot);
  applyRules(slotDefaults);
  applyRules(generationOptions || {});

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
      return { success: false, error: 'Valid start and end dates are required.' };
    }

    if (new Date(normalizedEnd) < new Date(normalizedStart)) {
      return { success: false, error: 'End date must be on or after the start date.' };
    }

    const slotList = loadShiftSlotsNormalized();
    if (!slotList.length) {
      return { success: false, error: 'No shift slots are available. Please create a slot first.' };
    }

    const requestedSlotIds = Array.isArray(shiftSlotIds) ? shiftSlotIds.filter(Boolean) : [];
    let selectedSlots = slotList.filter(slot => {
      if (!requestedSlotIds.length) {
        return true;
      }
      return requestedSlotIds.includes(slot.SlotId) || requestedSlotIds.includes(slot.ID);
    });

    if (!selectedSlots.length) {
      return { success: false, error: 'No matching shift slots were found for the selection.' };
    }

    let campaignId = normalizeCampaignIdValue(options.campaignId);
    if (!campaignId) {
      const campaignCandidates = selectedSlots
        .map(slot => normalizeCampaignIdValue(slot.Campaign))
        .filter(Boolean);
      if (campaignCandidates.length) {
        campaignId = campaignCandidates[0];
      }
    }

    if (!campaignId) {
      return { success: false, error: 'A campaign is required to generate schedules. Please select a campaign or ensure slots have a campaign assigned.' };
    }

    selectedSlots = selectedSlots.filter(slot => normalizeCampaignIdValue(slot.Campaign) === campaignId);
    if (!selectedSlots.length) {
      return { success: false, error: 'Selected shift slots do not belong to the requested campaign.' };
    }

    const slotLookup = buildSlotLookupById(selectedSlots);
    const targetDefaults = selectedSlots.length ? (selectedSlots[0].GenerationDefaults || {}) : {};
    const coverageTarget = normalizeCoverageTarget_(options, targetDefaults);
    const minimumCoverageRequired = computeMinimumCoverageTarget_(coverageTarget);

    const allAssignments = loadShiftAssignmentsNormalized();
    const existingAssignments = allAssignments.filter(assignment => {
      if (!assignment || assignment.Status === 'ARCHIVED') {
        return false;
      }
      if (normalizeCampaignIdValue(assignment.Campaign) !== campaignId) {
        return false;
      }
      const assignmentStart = new Date(assignment.StartDate + 'T00:00:00');
      const assignmentEnd = new Date(assignment.EndDate + 'T23:59:59');
      const periodStart = new Date(normalizedStart + 'T00:00:00');
      const periodEnd = new Date(normalizedEnd + 'T23:59:59');
      return assignmentEnd >= periodStart && assignmentStart <= periodEnd;
    });

    const generatedByUser = generatedBy || 'system';
    const scheduleUsers = clientGetScheduleUsers(generatedByUser, campaignId) || [];
    if (!scheduleUsers.length) {
      return { success: false, error: 'No eligible users were found for the selected campaign.' };
    }

    const userNameLookup = new Map();
    scheduleUsers.forEach(user => {
      const keys = [user.UserName, user.FullName, user.Email].map(value => normalizeUserKey(value));
      keys.filter(Boolean).forEach(key => {
        if (!userNameLookup.has(key)) {
          userNameLookup.set(key, user);
        }
      });
    });

    const requestedUserKeys = Array.isArray(userNames)
      ? userNames.map(name => normalizeUserKey(name)).filter(Boolean)
      : [];

    const includeAllUsers = requestedUserKeys.length === 0 || options.advanced?.autoAssignment === true;

    const eligibleUsers = scheduleUsers.filter(user => {
      if (!includeAllUsers) {
        const keyMatches = [user.UserName, user.FullName, user.Email]
          .map(value => normalizeUserKey(value))
          .filter(Boolean)
          .some(key => requestedUserKeys.includes(key));
        if (!keyMatches) {
          return false;
        }
      }

      if (user.isActive === false) {
        return false;
      }

      const hireDate = user.HireDate ? normalizeDateForSheet(user.HireDate, DEFAULT_SCHEDULE_TIME_ZONE) : null;
      if (hireDate && new Date(hireDate) > new Date(normalizedEnd + 'T23:59:59')) {
        return false;
      }

      const terminationDate = user.TerminationDate ? normalizeDateForSheet(user.TerminationDate, DEFAULT_SCHEDULE_TIME_ZONE) : null;
      if (terminationDate && new Date(terminationDate + 'T23:59:59') < new Date(normalizedStart + 'T00:00:00')) {
        return false;
      }

      return true;
    });

    if (!eligibleUsers.length) {
      return { success: false, error: 'No eligible users matched the selection or campaign criteria.' };
    }

    const holidayLookup = buildHolidayLookup_(normalizedStart, normalizedEnd);
    const attendanceConflicts = collectAttendanceConflicts_(eligibleUsers.map(user => user.UserName || user.FullName || user.Email), normalizedStart, normalizedEnd);

    const restPeriodHours = Number(options.advanced?.restPeriod) || targetDefaults.restPeriod || 0;
    const allowSwaps = options.advanced?.allowSwaps !== false;
    const weekendPremiumEnabled = options.advanced?.weekendPremium === true;
    const holidayPremiumEnabled = options.advanced?.holidayPremium !== false;
    const includeHolidays = options.includeHolidays !== false;

    const breaksConfig = {
      breakDuration: Number(options.breaks?.base) || targetDefaults.breakDuration || 15,
      break1Duration: Number(options.breaks?.first) || targetDefaults.break1Duration || 15,
      break2Duration: Number(options.breaks?.second) || targetDefaults.break2Duration || 0,
      lunchDuration: Number(options.breaks?.lunch) || targetDefaults.lunchDuration || 30,
      enableStaggered: options.breaks?.enableStaggered !== false && (targetDefaults.enableStaggeredBreaks !== false),
      breakGroups: Number(options.breaks?.groups) || targetDefaults.breakGroups || 3,
      staggerInterval: Number(options.breaks?.interval) || targetDefaults.staggerInterval || 15,
      minCoveragePct: Number(options.breaks?.minCoveragePct) || targetDefaults.minCoveragePct || null
    };

    const overtimeMinutes = (() => {
      if (options.overtime?.enabled !== true) {
        return 0;
      }
      const hours = Number(options.overtime?.maxDaily);
      if (Number.isFinite(hours) && hours > 0) {
        return Math.round(hours * 60);
      }
      return 0;
    })();

    const notificationLead = Number(options.advanced?.notificationLead) || targetDefaults.notificationLead || 0;
    const handoverMinutes = Number(options.advanced?.handoverTime) || targetDefaults.handoverTime || 0;

    const existingCoverage = buildAssignmentCoverage_(existingAssignments, slotLookup, normalizedStart, normalizedEnd);
    const plannedCoverage = new Map();
    const plannedAssignments = [];
    const conflicts = [];

    const seedString = options.randomSeed || `${normalizedStart}-${normalizedEnd}-${eligibleUsers.length}`;
    let seedAccumulator = 0;
    for (let i = 0; i < seedString.length; i++) {
      seedAccumulator = (seedAccumulator + seedString.charCodeAt(i)) % 2147483647;
    }

    const shuffledUsers = eligibleUsers.slice();
    for (let i = shuffledUsers.length - 1; i > 0; i--) {
      seedAccumulator = (seedAccumulator * 48271) % 2147483647;
      const random = seedAccumulator / 2147483647;
      const j = Math.floor(random * (i + 1));
      const temp = shuffledUsers[i];
      shuffledUsers[i] = shuffledUsers[j];
      shuffledUsers[j] = temp;
    }

    const groupId = Utilities.getUuid();

    const addCoverage = (slotId, dates) => {
      dates.forEach(dateStr => {
        if (!plannedCoverage.has(dateStr)) {
          plannedCoverage.set(dateStr, new Map());
        }
        const slotMap = plannedCoverage.get(dateStr);
        slotMap.set(slotId, (slotMap.get(slotId) || 0) + 1);
      });
    };

    const getCombinedCoverage = (slotId, dateStr) => {
      const existing = existingCoverage.has(dateStr) ? (existingCoverage.get(dateStr).bySlot.get(slotId) || 0) : 0;
      const planned = plannedCoverage.has(dateStr) ? (plannedCoverage.get(dateStr).get(slotId) || 0) : 0;
      return existing + planned;
    };

    shuffledUsers.forEach(user => {
      const userId = user.ID;
      const userName = user.FullName || user.UserName || user.Email;
      const normalizedUserKey = normalizeUserKey(user.UserName || user.FullName || user.Email);
      const restConflict = computeRestPeriodViolation_(existingAssignments.concat(plannedAssignments), userId, normalizedStart, restPeriodHours);
      if (restConflict) {
        conflicts.push({
          user: userName,
          periodStart: normalizedStart,
          periodEnd: normalizedEnd,
          type: 'REST_PERIOD',
          error: `Rest period violation with assignment ${restConflict.AssignmentId}`
        });
        return;
      }

      const slotCandidates = selectedSlots
        .slice()
        .sort((a, b) => {
          const aKey = getCombinedCoverage(a.SlotId, normalizedStart);
          const bKey = getCombinedCoverage(b.SlotId, normalizedStart);
          if (aKey !== bKey) {
            return aKey - bKey;
          }
          return (a.Priority || 0) - (b.Priority || 0);
        });

      let assignedSlot = null;
      let assignmentCoverageDates = [];

      for (let i = 0; i < slotCandidates.length; i++) {
        const candidate = slotCandidates[i];
        if (!candidate) {
          continue;
        }
        const slotCoverageTarget = normalizeCoverageTarget_(options, candidate.GenerationDefaults || {});
        const capacityLimit = slotCoverageTarget.maxCapacity;

        const tentativeAssignment = {
          StartDate: normalizedStart,
          EndDate: normalizedEnd,
          SlotId: candidate.SlotId
        };

        const dates = expandAssignmentToDays_(tentativeAssignment, candidate);
        if (!dates.length) {
          continue;
        }

        let capacityExceeded = false;
        for (let d = 0; d < dates.length; d++) {
          const dateStr = dates[d];
          const totalForDate = getCombinedCoverage(candidate.SlotId, dateStr);
          if (Number.isFinite(capacityLimit) && capacityLimit > 0 && totalForDate >= capacityLimit) {
            capacityExceeded = true;
            break;
          }
        }

        if (capacityExceeded) {
          continue;
        }

        const attendanceRecords = attendanceConflicts.get(normalizedUserKey) || [];
        if (attendanceRecords.length) {
          conflicts.push({
            user: userName,
            periodStart: normalizedStart,
            periodEnd: normalizedEnd,
            type: 'ATTENDANCE',
            error: `Attendance records conflict with proposed assignment (${attendanceRecords.length} record(s))`
          });
          continue;
        }

        assignedSlot = candidate;
        assignmentCoverageDates = dates;
        break;
      }

      if (!assignedSlot) {
        conflicts.push({
          user: userName,
          periodStart: normalizedStart,
          periodEnd: normalizedEnd,
          type: 'CAPACITY',
          error: 'Unable to place user without exceeding slot capacity or violating policies.'
        });
        return;
      }

      addCoverage(assignedSlot.SlotId, assignmentCoverageDates);

      const premiums = [];
      if (weekendPremiumEnabled) {
        const hasWeekend = assignmentCoverageDates.some(date => isWeekendDay_(date));
        if (hasWeekend) {
          premiums.push('Weekend');
        }
      }

      if (holidayPremiumEnabled) {
        const hasHoliday = assignmentCoverageDates.some(date => {
          const holidays = holidayLookup.get(date) || [];
          return holidays.some(holiday => String(holiday.region || '').toLowerCase() === 'jamaica');
        });
        if (hasHoliday) {
          premiums.push('Holiday');
        }
      }

      if (overtimeMinutes > 0) {
        premiums.push('Overtime');
      }

      const assignment = {
        AssignmentId: Utilities.getUuid(),
        UserId: userId,
        UserName: userName,
        Campaign: campaignId,
        SlotId: assignedSlot.SlotId,
        StartDate: normalizedStart,
        EndDate: normalizedEnd,
        Status: 'PENDING',
        AllowSwap: allowSwaps,
        Premiums: premiums,
        BreaksConfigJSON: {
          breakDuration: breaksConfig.breakDuration,
          break1Duration: breaksConfig.break1Duration,
          break2Duration: breaksConfig.break2Duration,
          lunchDuration: breaksConfig.lunchDuration,
          enableStaggered: breaksConfig.enableStaggered,
          breakGroups: breaksConfig.breakGroups,
          staggerInterval: breaksConfig.staggerInterval,
          minCoveragePct: breaksConfig.minCoveragePct
        },
        OvertimeMinutes: overtimeMinutes,
        RestPeriodHours: restPeriodHours,
        NotificationLeadHours: notificationLead,
        HandoverMinutes: handoverMinutes,
        Notes: options.notes || '',
        CreatedAt: new Date(),
        CreatedBy: generatedByUser,
        UpdatedAt: new Date(),
        UpdatedBy: generatedByUser,
        RollbackGroupId: groupId,
        Slot: assignedSlot,
        UserRecord: user
      };

      plannedAssignments.push(assignment);
    });

    const combinedAssignments = existingAssignments.concat(plannedAssignments);
    const combinedCoverage = buildAssignmentCoverage_(combinedAssignments, slotLookup, normalizedStart, normalizedEnd);

    const coverageSummary = [];
    const periodStartDate = new Date(normalizedStart + 'T00:00:00');
    const periodEndDate = new Date(normalizedEnd + 'T00:00:00');
    const coverageShortfalls = [];

    for (let cursor = new Date(periodStartDate.getTime()); cursor <= periodEndDate; cursor.setDate(cursor.getDate() + 1)) {
      const dateStr = Utilities.formatDate(cursor, DEFAULT_SCHEDULE_TIME_ZONE, 'yyyy-MM-dd');
      const coverageEntry = combinedCoverage.get(dateStr);
      const assignedCount = coverageEntry ? coverageEntry.count : 0;
      const minRequired = minimumCoverageRequired;
      const isHoliday = (holidayLookup.get(dateStr) || []).length > 0;
      const holidayNames = (holidayLookup.get(dateStr) || []).map(entry => entry.name).filter(Boolean);
      const summaryEntry = {
        date: dateStr,
        assigned: assignedCount,
        required: minRequired,
        variance: assignedCount - minRequired,
        isHoliday,
        holidayNames,
        isWeekend: isWeekendDay_(dateStr),
        status: assignedCount < minRequired ? 'SHORTFALL' : 'MET'
      };
      if (assignedCount < minRequired) {
        coverageShortfalls.push(summaryEntry);
      }
      coverageSummary.push(summaryEntry);
    }

    const previewAssignments = plannedAssignments.map(assignment => {
      const slot = assignment.Slot || slotLookup.get(String(assignment.SlotId));
      return {
        assignmentId: assignment.AssignmentId,
        userName: assignment.UserName,
        slotName: slot ? slot.SlotName : assignment.SlotId,
        startDate: assignment.StartDate,
        endDate: assignment.EndDate,
        startTime: slot ? slot.StartTime : '',
        endTime: slot ? slot.EndTime : '',
        premiums: assignment.Premiums,
        allowSwap: assignment.AllowSwap,
        status: assignment.Status
      };
    });

    if (options.previewOnly === true) {
      return {
        success: true,
        previewOnly: true,
        generated: plannedAssignments.length,
        userCount: eligibleUsers.length,
        periodStart: normalizedStart,
        periodEnd: normalizedEnd,
        assignments: previewAssignments,
        details: previewAssignments,
        coverage: coverageSummary,
        shortfalls: coverageShortfalls,
        conflicts
      };
    }

    if (plannedAssignments.length) {
      appendShiftAssignments(plannedAssignments);
      plannedAssignments.forEach(assignment => {
        logScheduleAudit('CREATE', 'ShiftAssignment', assignment.AssignmentId, null, assignment, 'Auto-generated shift assignment', { campaign: campaignId, slotId: assignment.SlotId });
      });
    }

    return {
      success: true,
      generated: plannedAssignments.length,
      userCount: eligibleUsers.length,
      periodStart: normalizedStart,
      periodEnd: normalizedEnd,
      assignments: previewAssignments,
      details: previewAssignments,
      coverage: coverageSummary,
      shortfalls: coverageShortfalls,
      conflicts,
      rollbackGroupId: groupId
    };

  } catch (error) {
    console.error('❌ Error generating schedules:', error);
    safeWriteError('clientGenerateSchedulesEnhanced', error);
    return {
      success: false,
      error: error && error.message ? error.message : 'Failed to generate schedules'
    };
  }
}


/**
 * Save schedules to sheet using ScheduleUtilities
 */
function saveSchedulesToSheet(schedules) {
  try {
    // Use ScheduleUtilities to ensure proper sheet and headers
    const sheet = ensureScheduleSheetWithHeaders(SCHEDULE_GENERATION_SHEET, SCHEDULE_GENERATION_HEADERS);

    schedules.forEach(schedule => {
      const normalized = normalizeSchedulePeriodRecord(schedule);
      // Create row data using proper header order from ScheduleUtilities
      const rowData = SCHEDULE_GENERATION_HEADERS.map(header => {
        if (normalized && Object.prototype.hasOwnProperty.call(normalized, header)) {
          return normalized[header];
        }
        return '';
      });
      sheet.appendRow(rowData);
    });

    SpreadsheetApp.flush();
    invalidateScheduleCaches();

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
    console.log('📋 Getting all schedules with filters:', filters);

    const assignments = loadShiftAssignmentsNormalized();
    if (!assignments.length) {
      console.warn('No range-based assignments found. Falling back to legacy schedule reader.');
      return legacyGetAllSchedules_(filters);
    }

    const slotLookup = buildSlotLookupById(loadShiftSlotsNormalized());
    const startDateFilter = filters.startDate ? new Date(normalizeDateForSheet(filters.startDate, DEFAULT_SCHEDULE_TIME_ZONE)) : null;
    const endDateFilter = filters.endDate ? new Date(normalizeDateForSheet(filters.endDate, DEFAULT_SCHEDULE_TIME_ZONE)) : null;
    const statusFilterRaw = filters.status ? String(filters.status).trim().toUpperCase() : '';
    const statusFilter = statusFilterRaw && statusFilterRaw !== 'ALL' ? statusFilterRaw : '';
    const userIdFilter = filters.userId ? String(filters.userId).trim() : '';
    const userNameFilter = filters.userName ? normalizeUserKey(filters.userName) : '';
    const campaignFilter = normalizeCampaignIdValue(filters.campaign || filters.campaignId || filters.Department || filters.department);
    const slotIdFilter = filters.slotId || filters.slot || filters.SlotId || '';
    const slotNameFilter = filters.slotName ? String(filters.slotName).trim().toLowerCase() : '';

    const filteredAssignments = assignments.filter(assignment => {
      if (!assignment) {
        return false;
      }

      const startDate = assignment.StartDate ? new Date(assignment.StartDate + 'T00:00:00') : null;
      const endDate = assignment.EndDate ? new Date(assignment.EndDate + 'T23:59:59') : startDate;

      if (startDateFilter && endDate && endDate < startDateFilter) {
        return false;
      }

      if (endDateFilter && startDate && startDate > endDateFilter) {
        return false;
      }

      if (statusFilter && String(assignment.Status || '').toUpperCase() !== statusFilter) {
        return false;
      }

      if (userIdFilter && String(assignment.UserId || '') !== userIdFilter) {
        return false;
      }

      if (userNameFilter) {
        const assignmentNameKey = normalizeUserKey(assignment.UserName || (assignment.UserRecord && (assignment.UserRecord.FullName || assignment.UserRecord.UserName || assignment.UserRecord.Email)));
        if (assignmentNameKey !== userNameFilter) {
          return false;
        }
      }

      if (campaignFilter && normalizeCampaignIdValue(assignment.Campaign || (assignment.Slot && assignment.Slot.Campaign)) !== campaignFilter) {
        return false;
      }

      if (slotIdFilter && String(assignment.SlotId || '') !== String(slotIdFilter)) {
        return false;
      }

      if (slotNameFilter) {
        const slot = assignment.Slot || slotLookup.get(assignment.SlotId) || null;
        const slotName = slot ? String(slot.SlotName || '').toLowerCase() : String(assignment.SlotName || '').toLowerCase();
        if (!slotName.includes(slotNameFilter)) {
          return false;
        }
      }

      return true;
    });

    const formatTimeForDisplay = (slot, key, minutesKey) => {
      if (!slot) {
        return '';
      }
      if (slot[key]) {
        return slot[key];
      }
      const minutes = slot[minutesKey];
      if (Number.isFinite(minutes)) {
        return formatMinutesToTime12Hour(minutes);
      }
      return '';
    };

    const formattedAssignments = filteredAssignments.map(assignment => {
      const slot = assignment.Slot || slotLookup.get(assignment.SlotId) || null;
      const premiums = Array.isArray(assignment.Premiums)
        ? assignment.Premiums
        : parsePremiumCsv_(assignment.Premiums);

      const startTime = formatTimeForDisplay(slot, 'StartTime', 'StartMinutes');
      const endTime = formatTimeForDisplay(slot, 'EndTime', 'EndMinutes');

      const view = {
        ID: assignment.AssignmentId,
        AssignmentId: assignment.AssignmentId,
        UserId: assignment.UserId,
        UserName: assignment.UserName || (assignment.UserRecord && (assignment.UserRecord.FullName || assignment.UserRecord.UserName || assignment.UserRecord.Email)) || '',
        Campaign: assignment.Campaign || (slot && slot.Campaign) || '',
        SlotId: assignment.SlotId,
        SlotName: (slot && slot.SlotName) || assignment.SlotName || '',
        StartDate: assignment.StartDate,
        EndDate: assignment.EndDate,
        PeriodStart: assignment.StartDate,
        PeriodEnd: assignment.EndDate,
        StartTime: startTime,
        EndTime: endTime,
        Status: assignment.Status || 'PENDING',
        AllowSwap: assignment.AllowSwap === true,
        Premiums: premiums,
        BreaksConfigJSON: assignment.BreaksConfigJSON || {},
        OvertimeMinutes: assignment.OvertimeMinutes || 0,
        RestPeriodHours: assignment.RestPeriodHours || 0,
        NotificationLeadHours: assignment.NotificationLeadHours || 0,
        HandoverMinutes: assignment.HandoverMinutes || 0,
        Notes: assignment.Notes || '',
        CreatedAt: assignment.CreatedAt || '',
        CreatedBy: assignment.CreatedBy || '',
        UpdatedAt: assignment.UpdatedAt || '',
        UpdatedBy: assignment.UpdatedBy || '',
        RollbackGroupId: assignment.RollbackGroupId || '',
        PremiumFlags: {
          weekend: premiums.includes('Weekend'),
          holiday: premiums.includes('Holiday'),
          overtime: premiums.includes('Overtime')
        },
        Source: 'ShiftAssignments'
      };

      return view;
    });

    formattedAssignments.sort((a, b) => {
      const startA = a.StartDate ? new Date(a.StartDate).getTime() : 0;
      const startB = b.StartDate ? new Date(b.StartDate).getTime() : 0;
      if (startA !== startB) {
        return startB - startA;
      }
      return String(b.UserName || '').localeCompare(String(a.UserName || ''));
    });

    return {
      success: true,
      schedules: formattedAssignments,
      total: formattedAssignments.length,
      filters,
      metadata: {
        source: 'ShiftAssignments',
        fallbackUsed: false
      }
    };
  } catch (error) {
    console.error('❌ Error getting schedules:', error);
    safeWriteError('clientGetAllSchedules', error);
    return {
      success: false,
      error: error.message,
      schedules: [],
      total: 0,
      filters
    };
  }
}

function legacyGetAllSchedules_(filters = {}) {
  try {
    let schedules = readScheduleSheet(SCHEDULE_GENERATION_SHEET) || [];

    if (!schedules.length) {
      const legacySheets = ['Schedules', 'Schedule', 'AgentSchedules'];
      for (let i = 0; i < legacySheets.length && !schedules.length; i++) {
        const legacyRows = readSheet(legacySheets[i]);
        if (Array.isArray(legacyRows) && legacyRows.length) {
          console.log(`Discovered legacy schedule data in ${legacySheets[i]}`);
          schedules = legacyRows.map(convertLegacyScheduleRecord).filter(Boolean);
        }
      }
    }

    const normalizedSchedules = schedules.map(record => normalizeSchedulePeriodRecord(record));

    let filteredSchedules = normalizedSchedules.slice();

    if (filters.startDate) {
      const startDate = new Date(filters.startDate);
      if (!isNaN(startDate.getTime())) {
        filteredSchedules = filteredSchedules.filter(s => {
          const scheduleEnd = resolveSchedulePeriodEndDate(s) || resolveSchedulePeriodStartDate(s);
          return !scheduleEnd || scheduleEnd >= startDate;
        });
      }
    }

    if (filters.endDate) {
      const endDate = new Date(filters.endDate);
      if (!isNaN(endDate.getTime())) {
        filteredSchedules = filteredSchedules.filter(s => {
          const scheduleStart = resolveSchedulePeriodStartDate(s);
          return !scheduleStart || scheduleStart <= endDate;
        });
      }
    }

    if (filters.userId) {
      filteredSchedules = filteredSchedules.filter(s => s.UserID === filters.userId);
    }

    if (filters.userName) {
      filteredSchedules = filteredSchedules.filter(s => s.UserName === filters.userName);
    }

    if (filters.status) {
      filteredSchedules = filteredSchedules.filter(s => String(s.Status || '').toUpperCase() === String(filters.status).toUpperCase());
    }

    if (filters.department) {
      filteredSchedules = filteredSchedules.filter(s => s.Department === filters.department);
    }

    filteredSchedules.sort((a, b) => getSchedulePeriodSortValue(b) - getSchedulePeriodSortValue(a));

    return {
      success: true,
      schedules: filteredSchedules,
      total: filteredSchedules.length,
      filters,
      metadata: {
        source: 'LegacyScheduleSheet',
        fallbackUsed: true
      }
    };
  } catch (error) {
    console.error('❌ Error getting schedules via legacy fallback:', error);
    safeWriteError('legacyGetAllSchedules_', error);
    return {
      success: false,
      error: error.message,
      schedules: [],
      total: 0,
      filters
    };
  }
}

function internalClientImportSchedules(importRequest = {}) {
  try {
    const schedules = Array.isArray(importRequest.schedules) ? importRequest.schedules : [];
    if (schedules.length === 0) {
      throw new Error('No schedules were provided for import.');
    }

    const metadata = importRequest.metadata || {};
    const timeZone = DEFAULT_SCHEDULE_TIME_ZONE;
    const now = new Date();
    const nowIso = Utilities.formatDate(now, timeZone, "yyyy-MM-dd'T'HH:mm:ss");

    const userLookup = buildScheduleUserLookup();
    const normalizedNew = schedules
      .map(raw => normalizeImportedScheduleRecord(raw, metadata, userLookup, nowIso, timeZone))
      .filter(record => record)
      .map(record => normalizeSchedulePeriodRecord(record, timeZone));

    if (normalizedNew.length === 0) {
      throw new Error('No valid schedules were found in the uploaded file.');
    }

    const existingRecords = readScheduleSheet(SCHEDULE_GENERATION_SHEET) || [];
    const normalizedExisting = existingRecords.map(record => normalizeSchedulePeriodRecord(record, timeZone));
    const replaceExisting = metadata.replaceExisting === true;

    let minStart = null;
    let maxEnd = null;

    normalizedNew.forEach(record => {
      const startDate = resolveSchedulePeriodStartDate(record, timeZone);
      const endDate = resolveSchedulePeriodEndDate(record, timeZone) || startDate;

      if (startDate && (!minStart || startDate < minStart)) {
        minStart = new Date(startDate);
      }
      if (endDate && (!maxEnd || endDate > maxEnd)) {
        maxEnd = new Date(endDate);
      }
    });

    if (metadata.startDate) {
      metadata.startDate = normalizeDateForSheet(metadata.startDate, timeZone);
    }
    if (metadata.endDate) {
      metadata.endDate = normalizeDateForSheet(metadata.endDate, timeZone);
    }
    if (metadata.startWeekDate) {
      metadata.startWeekDate = normalizeDateForSheet(metadata.startWeekDate, timeZone);
    }
    if (metadata.endWeekDate) {
      metadata.endWeekDate = normalizeDateForSheet(metadata.endWeekDate, timeZone);
    }

    if (!metadata.startDate && metadata.startWeekDate) {
      metadata.startDate = metadata.startWeekDate;
    }
    if (!metadata.endDate && metadata.endWeekDate) {
      metadata.endDate = metadata.endWeekDate;
    }

    const newKeys = new Set(normalizedNew.map(record => buildScheduleCompositeKey(record, timeZone)));
    let replacedCount = 0;

    const retainedRecords = normalizedExisting.filter(existing => {
      const key = buildScheduleCompositeKey(existing, timeZone);
      if (newKeys.has(key)) {
        replacedCount++;
        return false;
      }

      if (replaceExisting && minStart && maxEnd) {
        const existingStart = resolveSchedulePeriodStartDate(existing, timeZone);
        const existingEnd = resolveSchedulePeriodEndDate(existing, timeZone) || existingStart;
        if (existingStart && existingEnd && existingEnd >= minStart && existingStart <= maxEnd) {
          replacedCount++;
          return false;
        }
      }

      return true;
    });

    const combinedRecords = retainedRecords.concat(normalizedNew);

    combinedRecords.sort((a, b) => {
      const diff = getSchedulePeriodSortValue(a, timeZone) - getSchedulePeriodSortValue(b, timeZone);
      if (diff !== 0) {
        return diff;
      }

      const nameA = (a.UserName || '').toString();
      const nameB = (b.UserName || '').toString();
      return nameA.localeCompare(nameB);
    });

    writeToScheduleSheet(SCHEDULE_GENERATION_SHEET, combinedRecords);
    invalidateScheduleCaches();

    const normalizedStart = minStart ? normalizeDateForSheet(minStart, timeZone) : '';
    const normalizedEnd = maxEnd ? normalizeDateForSheet(maxEnd, timeZone) : '';

    const summary = typeof metadata.summary === 'object' && metadata.summary !== null
      ? metadata.summary
      : {};

    if (metadata.startDate && !summary.startDate) {
      summary.startDate = metadata.startDate;
    } else if (normalizedStart && !summary.startDate) {
      summary.startDate = normalizedStart;
    }

    if (metadata.endDate && !summary.endDate) {
      summary.endDate = metadata.endDate;
    } else if (normalizedEnd && !summary.endDate) {
      summary.endDate = normalizedEnd;
    }

    if (typeof summary.totalAssignments !== 'number') {
      summary.totalAssignments = normalizedNew.length;
    }
    if (typeof summary.totalShifts !== 'number') {
      summary.totalShifts = normalizedNew.length;
    }

    const daySpan = calculateDaySpanCount(metadata.startDate, metadata.endDate, minStart, maxEnd);
    if ((typeof summary.dayCount !== 'number' || summary.dayCount <= 0) && daySpan) {
      summary.dayCount = daySpan;
    }

    const weekSpan = calculateWeekSpanCount(
      metadata.startWeekDate || metadata.startDate,
      metadata.endWeekDate || metadata.endDate,
      minStart,
      maxEnd
    );
    if ((typeof summary.weekCount !== 'number' || summary.weekCount <= 0) && weekSpan) {
      summary.weekCount = weekSpan;
    }

    metadata.summary = summary;

    if (!metadata.dayCount && daySpan) {
      metadata.dayCount = daySpan;
    }

    if (!metadata.weekCount && weekSpan) {
      metadata.weekCount = weekSpan;
    }

    return {
      success: true,
      importedCount: normalizedNew.length,
      replacedCount,
      totalAfterImport: combinedRecords.length,
      range: {
        start: normalizedStart,
        end: normalizedEnd
      },
      metadata
    };

  } catch (error) {
    console.error('❌ Error importing schedules:', error);
    safeWriteError('clientImportSchedules', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Import schedules from uploaded data
 */
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

/**
 * Manually create shift slot assignments for specific users
 */
function clientAddManualShiftSlots(request = {}) {
  try {
    const timeZone = DEFAULT_SCHEDULE_TIME_ZONE;
    const actor = resolveCurrentActorEmail();
    const now = new Date();
    const nowIso = Utilities.formatDate(now, timeZone, "yyyy-MM-dd'T'HH:mm:ss");

    const normalizedStartDate = normalizeDateForSheet(request.startDate || request.date, timeZone);
    const normalizedEndDate = normalizeDateForSheet(request.endDate || request.startDate || request.date, timeZone) || normalizedStartDate;

    if (!normalizedStartDate || !normalizedEndDate) {
      return {
        success: false,
        error: 'Start date and end date are required.'
      };
    }

    if (new Date(normalizedEndDate) < new Date(normalizedStartDate)) {
      return {
        success: false,
        error: 'End date must be on or after the start date.'
      };
    }

    const slotCandidates = [request.slotId, request.slotID, request.SlotId, request.SlotID]
      .concat(request.slotName ? [request.slotName] : [])
      .map(value => (value === null || typeof value === 'undefined') ? '' : String(value).trim())
      .filter(Boolean);

    if (!slotCandidates.length) {
      return {
        success: false,
        error: 'A shift slot is required for manual assignment.'
      };
    }

    const allSlots = loadShiftSlotsNormalized();
    const slotLookup = buildSlotLookupById(allSlots);
    let selectedSlot = null;

    for (let i = 0; i < slotCandidates.length && !selectedSlot; i++) {
      const candidate = slotCandidates[i];
      if (slotLookup.has(candidate)) {
        selectedSlot = slotLookup.get(candidate);
        break;
      }
      selectedSlot = allSlots.find(slot => {
        if (!slot) {
          return false;
        }
        const slotIdMatch = [slot.SlotId, slot.ID].some(value => String(value || '') === candidate);
        if (slotIdMatch) {
          return true;
        }
        return String(slot.SlotName || '').toLowerCase() === candidate.toLowerCase();
      }) || null;
    }

    if (!selectedSlot) {
      return {
        success: false,
        error: 'The selected shift slot could not be found. Please refresh the slot list and try again.'
      };
    }

    const campaignId = normalizeCampaignIdValue(request.campaignId || request.campaign || selectedSlot.Campaign);
    if (!campaignId) {
      return {
        success: false,
        error: 'A campaign is required for manual assignments.'
      };
    }

    const requestedUserIds = Array.isArray(request.users) ? request.users.filter(Boolean).map(String) : [];
    if (!requestedUserIds.length) {
      return {
        success: false,
        error: 'Select at least one user before assigning a shift.'
      };
    }

    const usersById = buildUserLookupById_();
    const userNamesForAttendance = [];
    const validUsers = [];

    requestedUserIds.forEach(userId => {
      const userRecord = usersById.get(String(userId));
      if (userRecord) {
        validUsers.push(userRecord);
        const attendanceName = userRecord.UserName || userRecord.FullName || userRecord.Email;
        if (attendanceName) {
          userNamesForAttendance.push(attendanceName);
        }
      }
    });

    if (!validUsers.length) {
      return {
        success: false,
        error: 'None of the selected users were found in the directory.'
      };
    }

    const attendanceConflicts = collectAttendanceConflicts_(userNamesForAttendance, normalizedStartDate, normalizedEndDate);
    const existingAssignments = loadShiftAssignmentsNormalized();
    const holidayLookup = buildHolidayLookup_(normalizedStartDate, normalizedEndDate);
    const restPeriodHours = Number(selectedSlot.RestPeriod) || Number(selectedSlot.GenerationDefaults?.restPeriod) || 0;
    const allowSwap = selectedSlot.AllowSwaps === true;
    const weekendPremiumEnabled = selectedSlot.WeekendPremium === true;
    const holidayPremiumEnabled = selectedSlot.HolidayPremium === true;

    const breaksConfig = {
      breakDuration: Number(selectedSlot.BreakDuration) || Number(selectedSlot.GenerationDefaults?.breakDuration) || 15,
      break1Duration: Number(selectedSlot.Break1Duration) || Number(selectedSlot.GenerationDefaults?.break1Duration) || 15,
      break2Duration: Number(selectedSlot.Break2Duration) || Number(selectedSlot.GenerationDefaults?.break2Duration) || 0,
      lunchDuration: Number(selectedSlot.LunchDuration) || Number(selectedSlot.GenerationDefaults?.lunchDuration) || 30,
      enableStaggered: selectedSlot.EnableStaggeredBreaks === true,
      breakGroups: Number(selectedSlot.BreakGroups) || Number(selectedSlot.GenerationDefaults?.breakGroups) || 0,
      staggerInterval: Number(selectedSlot.StaggerInterval) || Number(selectedSlot.GenerationDefaults?.staggerInterval) || 0,
      minCoveragePct: Number(selectedSlot.MinCoveragePct) || Number(selectedSlot.GenerationDefaults?.minCoveragePct) || null
    };

    const resolveOvertimeMinutes = () => {
      const defaults = selectedSlot.GenerationDefaults || {};
      if (Number.isFinite(defaults.overtimeMinutes) && defaults.overtimeMinutes > 0) {
        return defaults.overtimeMinutes;
      }
      if (Number.isFinite(defaults.maxDailyOT) && defaults.maxDailyOT > 0) {
        return Math.round(defaults.maxDailyOT * 60);
      }
      return 0;
    };

    const overtimeMinutes = resolveOvertimeMinutes();
    const notificationLead = Number(selectedSlot.NotificationLead) || Number(selectedSlot.GenerationDefaults?.notificationLead) || 0;
    const handoverMinutes = Number(selectedSlot.HandoverTime) || Number(selectedSlot.GenerationDefaults?.handoverTime) || 0;

    const replaceExisting = request.replaceExisting === true;
    const archiveIds = new Set();
    const failedUsers = [];
    const newAssignments = [];
    const details = [];
    const groupId = Utilities.getUuid();

    const notesParts = [];
    const label = (request.slotLabel || '').trim();
    const manualNotes = (request.notes || '').trim();
    const rawSourceMonth = request.sourceMonth;
    const parsedSourceMonth = rawSourceMonth === null || typeof rawSourceMonth === 'undefined' || rawSourceMonth === ''
      ? null
      : Number(rawSourceMonth);

    if (label) {
      notesParts.push(`Label: ${label}`);
    }
    if (Number.isFinite(parsedSourceMonth)) {
      notesParts.push(`Source Month: ${parsedSourceMonth}`);
    }
    if (manualNotes) {
      notesParts.push(manualNotes);
    }

    const combinedNotes = notesParts.join(' | ');

    const normalizeTimeForDisplay = (slot, key, minutesKey) => {
      if (!slot) {
        return '';
      }
      const direct = slot[key];
      if (direct) {
        return direct;
      }
      const minutes = slot[minutesKey];
      if (Number.isFinite(minutes)) {
        return formatMinutesToTime12Hour(minutes);
      }
      return '';
    };

    const slotStartTime = normalizeTimeForDisplay(selectedSlot, 'StartTime', 'StartMinutes');
    const slotEndTime = normalizeTimeForDisplay(selectedSlot, 'EndTime', 'EndMinutes');

    validUsers.forEach(userRecord => {
      if (!userRecord || !userRecord.ID) {
        return;
      }

      const userId = String(userRecord.ID);
      const normalizedCampaign = normalizeCampaignIdValue(selectedSlot.Campaign);

      if (!isUserConsideredActive(userRecord)) {
        failedUsers.push({ userId, userName: userRecord.FullName || userRecord.UserName || userRecord.Email || userId, reason: 'User is not active.' });
        return;
      }

      if (!doesUserBelongToCampaign(userRecord, normalizedCampaign)) {
        failedUsers.push({ userId, userName: userRecord.FullName || userRecord.UserName || userRecord.Email || userId, reason: 'User is not assigned to the selected campaign.' });
        return;
      }

      const hireDate = normalizeDateForSheet(userRecord.HireDate || userRecord.hireDate || userRecord.StartDate || userRecord.Hire_Date, timeZone);
      if (hireDate && new Date(normalizedStartDate) < new Date(hireDate)) {
        failedUsers.push({ userId, userName: userRecord.FullName || userRecord.UserName || userRecord.Email || userId, reason: `Start date precedes hire date (${hireDate}).` });
        return;
      }

      const normalizedUserKey = normalizeUserKey(userRecord.UserName || userRecord.FullName || userRecord.Email);
      const attendanceRecords = normalizedUserKey ? (attendanceConflicts.get(normalizedUserKey) || []) : [];
      if (attendanceRecords.length) {
        failedUsers.push({ userId, userName: userRecord.FullName || userRecord.UserName || userRecord.Email || userId, reason: 'Attendance records conflict with the requested range.' });
        return;
      }

      const overlappingAssignments = existingAssignments.filter(assignment => {
        if (!assignment || assignment.Status === 'ARCHIVED') {
          return false;
        }
        if (String(assignment.UserId) !== userId) {
          return false;
        }
        const assignmentStart = new Date((assignment.StartDate || assignment.PeriodStart) + 'T00:00:00');
        const assignmentEnd = new Date((assignment.EndDate || assignment.PeriodEnd || assignment.StartDate) + 'T23:59:59');
        const requestedStart = new Date(normalizedStartDate + 'T00:00:00');
        const requestedEnd = new Date(normalizedEndDate + 'T23:59:59');
        return assignmentEnd >= requestedStart && assignmentStart <= requestedEnd;
      });

      if (overlappingAssignments.length && !replaceExisting) {
        failedUsers.push({ userId, userName: userRecord.FullName || userRecord.UserName || userRecord.Email || userId, reason: 'Existing assignment overlaps with the requested range.' });
        return;
      }

      overlappingAssignments.forEach(assignment => {
        if (assignment && assignment.AssignmentId) {
          archiveIds.add(String(assignment.AssignmentId));
        }
      });

      const restConflict = computeRestPeriodViolation_(existingAssignments, userId, normalizedStartDate, restPeriodHours);
      if (restConflict && !archiveIds.has(String(restConflict.AssignmentId))) {
        failedUsers.push({ userId, userName: userRecord.FullName || userRecord.UserName || userRecord.Email || userId, reason: 'Rest period policy violation with an existing assignment.' });
        return;
      }

      const assignment = {
        AssignmentId: Utilities.getUuid(),
        UserId: userId,
        UserName: userRecord.FullName || userRecord.UserName || userRecord.Email || userId,
        Campaign: normalizedCampaign,
        SlotId: selectedSlot.SlotId,
        SlotName: selectedSlot.SlotName,
        StartDate: normalizedStartDate,
        EndDate: normalizedEndDate,
        Status: 'PENDING',
        AllowSwap: allowSwap,
        BreaksConfigJSON: breaksConfig,
        OvertimeMinutes: overtimeMinutes,
        RestPeriodHours: restPeriodHours || 0,
        NotificationLeadHours: notificationLead,
        HandoverMinutes: handoverMinutes,
        Notes: combinedNotes,
        CreatedAt: now,
        CreatedBy: actor,
        UpdatedAt: now,
        UpdatedBy: actor,
        RollbackGroupId: groupId,
        Slot: selectedSlot,
        UserRecord: userRecord
      };

      const coverageDates = expandAssignmentToDays_(assignment, selectedSlot);
      const premiums = [];

      if (weekendPremiumEnabled && coverageDates.some(isWeekendDay_)) {
        premiums.push('Weekend');
      }

      if (holidayPremiumEnabled) {
        const hasPremiumHoliday = coverageDates.some(dateStr => {
          const holidays = holidayLookup.get(dateStr) || [];
          return holidays.some(holiday => String(holiday.region || '').toLowerCase() === 'jamaica');
        });
        if (hasPremiumHoliday) {
          premiums.push('Holiday');
        }
      }

      if (overtimeMinutes > 0) {
        premiums.push('Overtime');
      }

      assignment.Premiums = premiums;

      newAssignments.push(assignment);
      details.push({
        assignmentId: assignment.AssignmentId,
        userId: assignment.UserId,
        userName: assignment.UserName,
        slotId: assignment.SlotId,
        slotName: assignment.SlotName,
        startDate: assignment.StartDate,
        endDate: assignment.EndDate,
        startTime: slotStartTime,
        endTime: slotEndTime
      });
    });

    if (!newAssignments.length) {
      return {
        success: false,
        error: 'No assignments were created.',
        failed: failedUsers
      };
    }

    let replacedCount = 0;
    if (archiveIds.size) {
      const archiveList = Array.from(archiveIds);
      const archiveResult = updateAssignmentsByIds_(archiveList, (assignment) => {
        assignment.Status = 'ARCHIVED';
        assignment.UpdatedAt = now;
        assignment.UpdatedBy = actor;
        assignment.RollbackGroupId = groupId;
        assignment.Notes = assignment.Notes ? `${assignment.Notes} | Replaced by manual assignment ${groupId}` : `Replaced by manual assignment ${groupId}`;
      });
      replacedCount = archiveResult.updated.length;
      archiveResult.updated.forEach(change => {
        logScheduleAudit('ARCHIVE', 'ShiftAssignment', change.after.AssignmentId, change.before, change.after, 'Manual assignment replaceExisting', { groupId, replaceExisting: true });
      });
    }

    appendShiftAssignments(newAssignments);

    newAssignments.forEach(assignment => {
      logScheduleAudit('CREATE', 'ShiftAssignment', assignment.AssignmentId, null, assignment, 'Manual shift assignment', { groupId, replaceExisting });
    });

    const uniqueUsers = new Set(newAssignments.map(record => record.UserId));
    const dateLabel = normalizedStartDate === normalizedEndDate ? normalizedStartDate : `${normalizedStartDate} → ${normalizedEndDate}`;
    const message = `Added ${newAssignments.length} manual shift assignment${newAssignments.length === 1 ? '' : 's'} for ${selectedSlot.SlotName} (${dateLabel}).`;

    return {
      success: true,
      message,
      created: newAssignments.length,
      replaced: replacedCount,
      failed: failedUsers,
      details,
      slotName: selectedSlot.SlotName,
      slotId: selectedSlot.SlotId,
      startDate: normalizedStartDate,
      endDate: normalizedEndDate,
      usersAffected: uniqueUsers.size,
      rollbackGroupId: groupId,
      createdAt: nowIso
    };
  } catch (error) {
    console.error('Error manually adding shift slots:', error);
    safeWriteError('clientAddManualShiftSlots', error);
    return {
      success: false,
      error: error && error.message ? error.message : 'Failed to add manual shift slots.'
    };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// ATTENDANCE MANAGEMENT - Uses ScheduleUtilities
// ────────────────────────────────────────────────────────────────────────────

/**
 * Mark attendance status for a user on a specific date
 */
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
      const slotsResult = clientGetAllShiftSlots();
      const slots = Array.isArray(slotsResult)
        ? slotsResult
        : (slotsResult && Array.isArray(slotsResult.slots) ? slotsResult.slots : []);
      const slotsSuccess = Array.isArray(slotsResult) ? true : slotsResult && slotsResult.success !== false;

      diagnostics.shiftSlots = {
        count: slots.length,
        working: slotsSuccess && slots.length > 0,
        scheduleUtilitiesIntegration: slotsSuccess,
        metadata: slotsResult && !Array.isArray(slotsResult) ? slotsResult.metadata : undefined
      };

      if (!slotsSuccess) {
        diagnostics.issues.push({
          severity: 'HIGH',
          component: 'Shift Slots',
          message: 'Shift slot loader returned an error: ' + (slotsResult && slotsResult.error ? slotsResult.error : 'Unknown error')
        });
      } else if (slots.length === 0) {
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
    const ids = Array.isArray(scheduleIds) ? scheduleIds.filter(Boolean).map(String) : [];
    if (!ids.length) {
      return {
        success: false,
        error: 'No schedule identifiers were provided.'
      };
    }

    const actor = resolveCurrentActorEmail();
    const approver = approvingUserId || actor;
    const noteText = notes ? String(notes).trim() : '';
    const now = new Date();

    const { updated } = updateAssignmentsByIds_(ids, (assignment) => {
      assignment.Status = 'APPROVED';
      assignment.UpdatedAt = now;
      assignment.UpdatedBy = approver;
      if (noteText) {
        assignment.Notes = assignment.Notes ? `${assignment.Notes} | Approved: ${noteText}` : `Approved: ${noteText}`;
      }
    });

    if (!updated.length) {
      return {
        success: false,
        error: 'No assignments were approved. They may already be processed or missing.'
      };
    }

    updated.forEach(change => {
      logScheduleAudit('APPROVE', 'ShiftAssignment', change.after.AssignmentId, change.before, change.after, noteText, {
        action: 'approve',
        approver
      });
    });

    return {
      success: true,
      message: `Approved ${updated.length} assignments`,
      approved: updated.length,
      updatedIds: updated.map(change => change.after.AssignmentId)
    };
  } catch (error) {
    console.error('Error approving schedules:', error);
    safeWriteError('clientApproveSchedules', error);
    return {
      success: false,
      error: error.message || 'Failed to approve schedules'
    };
  }
}

function clientRejectSchedules(scheduleIds, rejectingUserId, reason = '') {
  try {
    const ids = Array.isArray(scheduleIds) ? scheduleIds.filter(Boolean).map(String) : [];
    if (!ids.length) {
      return {
        success: false,
        error: 'No schedule identifiers were provided.'
      };
    }

    const actor = resolveCurrentActorEmail();
    const rejector = rejectingUserId || actor;
    const reasonText = reason ? String(reason).trim() : '';
    const now = new Date();

    const { updated } = updateAssignmentsByIds_(ids, (assignment) => {
      assignment.Status = 'REJECTED';
      assignment.UpdatedAt = now;
      assignment.UpdatedBy = rejector;
      if (reasonText) {
        assignment.Notes = assignment.Notes ? `${assignment.Notes} | Rejected: ${reasonText}` : `Rejected: ${reasonText}`;
      }
    });

    if (!updated.length) {
      return {
        success: false,
        error: 'No assignments were rejected. They may already be processed or missing.'
      };
    }

    updated.forEach(change => {
      logScheduleAudit('REJECT', 'ShiftAssignment', change.after.AssignmentId, change.before, change.after, reasonText, {
        action: 'reject',
        rejector
      });
    });

    return {
      success: true,
      message: `Rejected ${updated.length} assignments`,
      rejected: updated.length,
      updatedIds: updated.map(change => change.after.AssignmentId)
    };
  } catch (error) {
    console.error('Error rejecting schedules:', error);
    safeWriteError('clientRejectSchedules', error);
    return {
      success: false,
      error: error.message || 'Failed to reject schedules'
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
    const schedules = readScheduleSheet(SCHEDULE_GENERATION_SHEET) || [];
    const normalizeDate = (typeof normalizeScheduleDate === 'function')
      ? normalizeScheduleDate
      : value => {
          if (!value) {
            return null;
          }
          const date = new Date(value);
          return isNaN(date.getTime()) ? null : date;
        };

    const requestedStart = normalizeDate(periodStart);
    const requestedEnd = normalizeDate(periodEnd || periodStart);

    if (!requestedStart || !requestedEnd) {
      return null;
    }

    return schedules.find(schedule => {
      if (schedule.UserName !== userName) {
        return false;
      }

      const existingStart = normalizeDate(schedule.PeriodStart || schedule.Date);
      const existingEnd = normalizeDate(schedule.PeriodEnd || schedule.Date || schedule.PeriodStart);

      if (!existingStart || !existingEnd) {
        return false;
      }

      return existingStart <= requestedEnd && existingEnd >= requestedStart;
    }) || null;
  } catch (error) {
    console.warn('Error checking existing schedule:', error);
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

  const scheduleRows = loadSheet(SCHEDULE_GENERATION_SHEET);
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
