/**
 * COMPLETE Enhanced Schedule Management Backend Service
 * Version 4.1 - Integrated with ScheduleUtilities and MainUtilities
 * Now properly uses dedicated spreadsheet support and shared functions
 */

// ────────────────────────────────────────────────────────────────────────────
// CONFIGURATION - Uses ScheduleUtilities constants
// ────────────────────────────────────────────────────────────────────────────

const SCHEDULE_CONFIG = {
  PRIMARY_COUNTRY: 'JM', // Jamaica takes priority
  SUPPORTED_COUNTRIES: ['JM', 'US', 'DO', 'PH'], // Jamaica, US, Dominican Republic, Philippines
  DEFAULT_SHIFT_CAPACITY: 10,
  DEFAULT_BREAK_MINUTES: 15,
  DEFAULT_LUNCH_MINUTES: 60,
  CACHE_DURATION: 300 // 5 minutes
};

function scheduleFlagToBool(value) {
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

function normalizeCampaignIdValue(value) {
  if (value === null || typeof value === 'undefined') {
    return '';
  }

  if (Array.isArray(value)) {
    for (let i = 0; i < value.length; i++) {
      const normalized = normalizeCampaignIdValue(value[i]);
      if (normalized) {
        return normalized;
      }
    }
    return '';
  }

  if (typeof value === 'object') {
    const objectCandidates = [
      value.ID,
      value.Id,
      value.id,
      value.CampaignID,
      value.campaignID,
      value.CampaignId,
      value.campaignId,
      value.value
    ];

    for (let i = 0; i < objectCandidates.length; i++) {
      const normalized = normalizeCampaignIdValue(objectCandidates[i]);
      if (normalized) {
        return normalized;
      }
    }

    return '';
  }

  const text = String(value).trim();
  if (!text || text.toLowerCase() === 'undefined' || text.toLowerCase() === 'null') {
    return '';
  }

  return text;
}

function doesUserBelongToCampaign(user, campaignId) {
  const normalizedCampaignId = normalizeCampaignIdValue(campaignId);
  if (!normalizedCampaignId || !user) {
    return false;
  }

  const candidateValues = [
    user.CampaignID,
    user.campaignID,
    user.CampaignId,
    user.campaignId,
    user.Campaign,
    user.campaign,
    user.primaryCampaignId,
    user.PrimaryCampaignId,
    user.primaryCampaignID,
    user.PrimaryCampaignID
  ];

  for (let i = 0; i < candidateValues.length; i++) {
    const candidate = normalizeCampaignIdValue(candidateValues[i]);
    if (candidate && candidate === normalizedCampaignId) {
      return true;
    }
  }

  return false;
}

function collectUserRoleCandidates(user) {
  const roles = [];

  const appendValue = (value) => {
    if (value === null || typeof value === 'undefined') {
      return;
    }

    if (Array.isArray(value)) {
      value.forEach(appendValue);
      return;
    }

    if (typeof value === 'object') {
      appendValue(value.name || value.Name || value.roleName || value.RoleName);
      appendValue(value.value);
      return;
    }

    const text = String(value);
    if (!text) {
      return;
    }

    text.split(/[,;/|]+/).forEach((part) => {
      const trimmed = part.trim();
      if (trimmed) {
        roles.push(trimmed);
      }
    });
  };

  appendValue(user && user.roleNames);
  appendValue(user && user.RoleNames);
  appendValue(user && user.roles);
  appendValue(user && user.Roles);
  appendValue(user && user.role);
  appendValue(user && user.Role);
  appendValue(user && user.primaryRole);
  appendValue(user && user.PrimaryRole);
  appendValue(user && user.primaryRoles);
  appendValue(user && user.PrimaryRoles);
  appendValue(user && user.csvRoles);
  appendValue(user && user.CsvRoles);
  appendValue(user && user.RoleName);
  appendValue(user && user.roleName);

  return roles;
}

function isScheduleRoleRestricted(user) {
  const restrictedRoles = ['client', 'guest'];
  const roleNames = collectUserRoleCandidates(user)
    .map(role => String(role || '').trim().toLowerCase())
    .filter(Boolean);

  return roleNames.some(role => restrictedRoles.includes(role));
}

function isScheduleNameRestricted(user) {
  const restrictedNames = ['client', 'guest'];
  const nameCandidates = [
    user && user.UserName,
    user && user.Username,
    user && user.username,
    user && user.FullName,
    user && user.Name,
    user && user.DisplayName
  ];

  return nameCandidates.some(name => {
    if (!name) {
      return false;
    }
    const normalized = String(name).trim().toLowerCase();
    return normalized && restrictedNames.includes(normalized);
  });
}

function filterUsersByCampaign(users, campaignId) {
  const normalizedCampaignId = normalizeCampaignIdValue(campaignId);
  if (!normalizedCampaignId) {
    return Array.isArray(users) ? users.slice() : [];
  }

  let filteredUsers = Array.isArray(users) ? users.filter(Boolean) : [];

  try {
    if (typeof getUsersByCampaign === 'function') {
      const campaignUsers = getUsersByCampaign(normalizedCampaignId) || [];
      if (Array.isArray(campaignUsers) && campaignUsers.length) {
        const campaignUserIds = new Set(
          campaignUsers
            .map(user => normalizeUserIdValue(user && user.ID))
            .filter(Boolean)
        );

        if (campaignUserIds.size) {
          filteredUsers = filteredUsers.filter(user => campaignUserIds.has(normalizeUserIdValue(user && user.ID)));
          return filteredUsers;
        }
      }
    }
  } catch (error) {
    console.warn('Unable to resolve campaign membership via getUsersByCampaign:', normalizedCampaignId, error);
  }

  return filteredUsers.filter(user => doesUserBelongToCampaign(user, normalizedCampaignId));
}

// ────────────────────────────────────────────────────────────────────────────
// USER MANAGEMENT FUNCTIONS - Integrated with MainUtilities
// ────────────────────────────────────────────────────────────────────────────

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
    console.log('🕒 Creating shift slot:', slotData);

    // Validate required fields
    if (!slotData.name || !slotData.startTime || !slotData.endTime) {
      return {
        success: false,
        error: 'Slot name, start time, and end time are required'
      };
    }

    // Use ScheduleUtilities validation
    const validation = validateShiftSlot(slotData);
    if (!validation.isValid) {
      return {
        success: false,
        error: validation.errors.join('; ')
      };
    }

    // Use ScheduleUtilities to ensure sheet exists with proper headers
    const sheet = ensureScheduleSheetWithHeaders(SHIFT_SLOTS_SHEET, SHIFT_SLOTS_HEADERS);
    const now = new Date();
    const slotId = Utilities.getUuid();

    const toNumber = (value, fallback = '') => {
      if (value === null || value === undefined || value === '') {
        return fallback;
      }

      const num = Number(value);
      return Number.isFinite(num) ? num : fallback;
    };

    const toBoolean = (value, fallback = false) => {
      if (typeof value === 'boolean') return value;
      if (typeof value === 'number') return value !== 0;
      if (typeof value === 'string') {
        const normalized = value.trim().toLowerCase();
        if (!normalized) return fallback;
        return ['true', 'yes', '1', 'y'].includes(normalized);
      }
      return fallback;
    };

    // Process days of week
    let daysOfWeek = '1,2,3,4,5'; // Default to weekdays
    if (slotData.daysOfWeek && Array.isArray(slotData.daysOfWeek)) {
      daysOfWeek = slotData.daysOfWeek.join(',');
    }

    const maxCapacity = toNumber(
      slotData.maxCapacity,
      SCHEDULE_CONFIG.DEFAULT_SHIFT_CAPACITY
    );
    const breakDuration = toNumber(
      slotData.breakDuration !== undefined ? slotData.breakDuration : slotData.break1Duration,
      SCHEDULE_CONFIG.DEFAULT_BREAK_MINUTES
    );
    const lunchDuration = toNumber(
      slotData.lunchDuration,
      SCHEDULE_CONFIG.DEFAULT_LUNCH_MINUTES
    );

    const slot = {
      ID: slotId,
      Name: slotData.name,
      StartTime: slotData.startTime,
      EndTime: slotData.endTime,
      DaysOfWeek: daysOfWeek,
      Department: slotData.department || 'General',
      Location: slotData.location || 'Office',
      MaxCapacity: maxCapacity,
      MinCoverage: toNumber(slotData.minCoverage, ''),
      Priority: toNumber(slotData.priority, 2),
      Description: slotData.description || '',
      BreakDuration: breakDuration,
      LunchDuration: lunchDuration,
      Break1Duration: toNumber(slotData.break1Duration, breakDuration),
      Break2Duration: toNumber(slotData.break2Duration, 0),
      EnableStaggeredBreaks: toBoolean(slotData.enableStaggeredBreaks, true),
      BreakGroups: toNumber(slotData.breakGroups, 3),
      StaggerInterval: toNumber(slotData.staggerInterval, 15),
      MinCoveragePct: toNumber(slotData.minCoveragePct, 70),
      EnableOvertime: toBoolean(slotData.enableOvertime, false),
      MaxDailyOT: toNumber(slotData.maxDailyOT, 0),
      MaxWeeklyOT: toNumber(slotData.maxWeeklyOT, 0),
      OTApproval: slotData.otApproval || slotData.overtimeApproval || 'supervisor',
      OTRate: toNumber(slotData.otRate, 1.5),
      OTPolicy: slotData.otPolicy || slotData.overtimePolicy || 'MANDATORY',
      AllowSwaps: toBoolean(slotData.allowSwaps, true),
      WeekendPremium: toBoolean(slotData.weekendPremium, false),
      HolidayPremium: toBoolean(slotData.holidayPremium, true),
      AutoAssignment: toBoolean(slotData.autoAssignment, false),
      RestPeriod: toNumber(slotData.restPeriod, 8),
      NotificationLead: toNumber(slotData.notificationLead, 24),
      HandoverTime: toNumber(slotData.handoverTime, 15),
      OvertimePolicy: slotData.overtimePolicy || slotData.otPolicy || 'LIMITED_30',
      IsActive: true,
      CreatedBy: slotData.createdBy || 'System',
      CreatedAt: now,
      UpdatedAt: now
    };

    // Create row data using proper header order
    const rowData = SHIFT_SLOTS_HEADERS.map(header =>
      Object.prototype.hasOwnProperty.call(slot, header) ? slot[header] : ''
    );
    sheet.appendRow(rowData);
    SpreadsheetApp.flush();

    // Invalidate cache
    invalidateScheduleCaches();

    console.log('✅ Shift slot created successfully:', slotId);

    return {
      success: true,
      message: 'Shift slot created successfully',
      slot: slot
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

function normalizeUserIdValue(value) {
  if (value === null || value === undefined) {
    return '';
  }

  const str = typeof value === 'number' && Number.isFinite(value)
    ? String(Math.trunc(value))
    : String(value);

  return str.trim();
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

    // Use ScheduleUtilities to read from schedule sheet
    let slots = readScheduleSheet(SHIFT_SLOTS_SHEET) || [];
    
    // If no slots exist, create defaults using ScheduleUtilities function
    if (slots.length === 0) {
      console.log('No shift slots found, creating defaults');
      createDefaultShiftSlots();
      slots = readScheduleSheet(SHIFT_SLOTS_SHEET) || [];
    }

    if (!slots.length) {
      const legacySlotSheets = ['Shift Slots', 'Shifts', 'ShiftTemplates'];
      for (let i = 0; i < legacySlotSheets.length && !slots.length; i++) {
        const legacyRows = readSheet(legacySlotSheets[i]);
        if (Array.isArray(legacyRows) && legacyRows.length) {
          console.log(`Found legacy shift slot data in ${legacySlotSheets[i]}`);
          slots = legacyRows.map(convertLegacyShiftSlotRecord).filter(Boolean);
        }
      }
    }

    const normalizeBoolean = value => {
      if (value === true || value === false) return value;
      if (typeof value === 'number') return value !== 0;
      if (typeof value === 'string') {
        const normalized = value.trim().toLowerCase();
        if (!normalized) return false;
        return ['true', 'yes', '1', 'y'].includes(normalized);
      }
      return false;
    };

    const normalizeNumber = value => {
      if (value === null || typeof value === 'undefined' || value === '') {
        return '';
      }
      if (typeof value === 'number') return value;
      const parsed = Number(value);
      return Number.isFinite(parsed) ? parsed : value;
    };

    return slots.map(slot => {
      const normalizedSlot = {
        ...slot,
        DaysOfWeekArray: Array.isArray(slot.DaysOfWeekArray) && slot.DaysOfWeekArray.length
          ? slot.DaysOfWeekArray
          : slot.DaysOfWeek
            ? String(slot.DaysOfWeek).split(',').map(d => parseInt(String(d).trim(), 10)).filter(d => !isNaN(d))
            : [1, 2, 3, 4, 5]
      };

      normalizedSlot.EnableStaggeredBreaks = normalizeBoolean(slot.EnableStaggeredBreaks);
      normalizedSlot.EnableOvertime = normalizeBoolean(slot.EnableOvertime);
      normalizedSlot.AllowSwaps = normalizeBoolean(slot.AllowSwaps);
      normalizedSlot.WeekendPremium = normalizeBoolean(slot.WeekendPremium);
      normalizedSlot.HolidayPremium = normalizeBoolean(slot.HolidayPremium);
      normalizedSlot.AutoAssignment = normalizeBoolean(slot.AutoAssignment);
      const isActive = slot.IsActive === '' ? true : normalizeBoolean(slot.IsActive);
      normalizedSlot.IsActive = isActive;

      normalizedSlot.MaxCapacity = normalizeNumber(slot.MaxCapacity);
      normalizedSlot.MinCoverage = normalizeNumber(slot.MinCoverage);
      normalizedSlot.Priority = normalizeNumber(slot.Priority);
      normalizedSlot.BreakDuration = normalizeNumber(slot.BreakDuration);
      normalizedSlot.LunchDuration = normalizeNumber(slot.LunchDuration);
      normalizedSlot.Break1Duration = normalizeNumber(slot.Break1Duration);
      normalizedSlot.Break2Duration = normalizeNumber(slot.Break2Duration);
      normalizedSlot.BreakGroups = normalizeNumber(slot.BreakGroups);
      normalizedSlot.StaggerInterval = normalizeNumber(slot.StaggerInterval);
      normalizedSlot.MinCoveragePct = normalizeNumber(slot.MinCoveragePct);
      normalizedSlot.MaxDailyOT = normalizeNumber(slot.MaxDailyOT);
      normalizedSlot.MaxWeeklyOT = normalizeNumber(slot.MaxWeeklyOT);
      normalizedSlot.OTRate = normalizeNumber(slot.OTRate);
      normalizedSlot.RestPeriod = normalizeNumber(slot.RestPeriod);
      normalizedSlot.NotificationLead = normalizeNumber(slot.NotificationLead);
      normalizedSlot.HandoverTime = normalizeNumber(slot.HandoverTime);

      if (slot.CreatedAt) {
        const createdDate = new Date(slot.CreatedAt);
        normalizedSlot.CreatedAt = isNaN(createdDate.getTime()) ? slot.CreatedAt : createdDate;
      }

      if (slot.UpdatedAt) {
        const updatedDate = new Date(slot.UpdatedAt);
        normalizedSlot.UpdatedAt = isNaN(updatedDate.getTime()) ? slot.UpdatedAt : updatedDate;
      }

      return normalizedSlot;
    });

  } catch (error) {
    console.error('❌ Error getting shift slots:', error);
    safeWriteError('clientGetAllShiftSlots', error);
    
    // Fallback: ensure at least default slots exist
    try {
      createDefaultShiftSlots();
      return readScheduleSheet(SHIFT_SLOTS_SHEET) || [];
    } catch (fallbackError) {
      return [];
    }
  }
}

// ────────────────────────────────────────────────────────────────────────────
// SCHEDULE GENERATION - Enhanced with ScheduleUtilities integration
// ────────────────────────────────────────────────────────────────────────────

/**
 * Enhanced schedule generation with comprehensive validation
 */
/**
 * Enhanced schedule generation with comprehensive validation
 */
function clientGenerateSchedulesEnhanced(startDate, endDate, userNames, shiftSlotIds, templateId, generatedBy, options = {}) {
  try {
    console.log('🚀 Enhanced schedule generation started');
    console.log('Parameters:', { startDate, endDate, userNames, shiftSlotIds, templateId, generatedBy, options });

    // Use ScheduleUtilities validation
    const validation = validateScheduleParameters(startDate, endDate, userNames);
    if (!validation.isValid) {
      throw new Error('Invalid parameters: ' + validation.errors.join('; '));
    }

    if (!generatedBy) {
      generatedBy = 'System';
    }

    const start = new Date(startDate);
    const end = new Date(endDate);

    // Get users to schedule
    let usersToSchedule = [];
    if (userNames && userNames.length > 0) {
      usersToSchedule = userNames;
    } else {
      // If no users specified, get all active users for the requesting user
      const allUsers = clientGetAttendanceUsers(generatedBy, options.campaignId);
      if (!allUsers || allUsers.length === 0) {
        throw new Error('No users found for scheduling. Please check user data.');
      }
      usersToSchedule = allUsers;
    }

    console.log(`📝 Scheduling for ${usersToSchedule.length} users`);

    // Get shift slots - either selected ones or all available
    let shiftSlots = [];
    if (shiftSlotIds && shiftSlotIds.length > 0) {
      // Get only the selected shift slots
      console.log(`🎯 Using ${shiftSlotIds.length} selected shift slots:`, shiftSlotIds);
      const allSlots = clientGetAllShiftSlots();
      shiftSlots = allSlots.filter(slot => shiftSlotIds.includes(slot.ID));
      
      if (shiftSlots.length === 0) {
        throw new Error('None of the selected shift slots were found. Please refresh and try again.');
      }
      
      console.log(`✅ Found ${shiftSlots.length} matching shift slots`);
    } else {
      // Use all available shift slots
      shiftSlots = clientGetAllShiftSlots();
      console.log(`📋 Using all available shift slots (${shiftSlots.length} total)`);
    }

    if (!shiftSlots || shiftSlots.length === 0) {
      throw new Error('No shift slots available. Please create shift slots first or select specific slots.');
    }

    console.log(`⏰ Working with ${shiftSlots.length} shift slot(s)`);

    // Generate schedules
    const generatedSchedules = [];
    const conflicts = [];
    const dstChanges = [];

    // Loop through each date
    for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
      const currentDate = new Date(d);
      const dateStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const dayOfWeek = currentDate.getDay(); // 0 = Sunday, 1 = Monday, etc.

      console.log(`📅 Processing date: ${dateStr} (Day: ${dayOfWeek})`);

      // Check for holidays using ScheduleUtilities
      const isHoliday = checkIfHoliday(dateStr);
      if (isHoliday && !options.includeHolidays) {
        console.log(`🎉 Skipping holiday: ${dateStr}`);
        continue;
      }

      // Check DST status using ScheduleUtilities
      const dstStatus = checkDSTStatus(dateStr);
      if (dstStatus.isDSTChange) {
        dstChanges.push({
          date: dateStr,
          changeType: dstStatus.changeType,
          adjustment: dstStatus.timeAdjustment
        });
      }

      // Generate schedules for each user
      usersToSchedule.forEach(userName => {
        try {
          // Get suitable shift slots for this user and day from the selected/available slots
          const suitableSlots = shiftSlots.filter(slot => {
            if (!slot.IsActive) return false;

            // Check if slot is active on this day of week
            const daysOfWeek = slot.DaysOfWeek ? slot.DaysOfWeek.split(',').map(d => parseInt(d)) : [1, 2, 3, 4, 5];
            return daysOfWeek.includes(dayOfWeek);
          });

          if (suitableSlots.length === 0) {
            console.log(`⚠️ No suitable slots for ${userName} on ${dateStr} from selected slots`);
            conflicts.push({
              user: userName,
              date: dateStr,
              error: 'No suitable shift slots available for this day from selected slots',
              type: 'NO_SUITABLE_SLOTS'
            });
            return;
          }

          // Select best suitable slot (can be enhanced with more logic)
          // For now, prefer slots with higher capacity or priority
          const selectedSlot = suitableSlots.sort((a, b) => {
            const priorityA = a.Priority || 2;
            const priorityB = b.Priority || 2;
            if (priorityA !== priorityB) return priorityB - priorityA; // Higher priority first
            
            const capacityA = a.MaxCapacity || 10;
            const capacityB = b.MaxCapacity || 10;
            return capacityB - capacityA; // Higher capacity first
          })[0];

          // Check for conflicts using ScheduleUtilities
          const existingSchedule = checkExistingSchedule(userName, dateStr);
          if (existingSchedule && !options.overrideExisting) {
            conflicts.push({
              user: userName,
              date: dateStr,
              error: 'User already has a schedule for this date',
              type: 'USER_DOUBLE_BOOKING'
            });
            return;
          }

          // Create schedule record using ScheduleUtilities time functions
          const schedule = {
            ID: Utilities.getUuid(),
            UserID: getUserIdByName(userName),
            UserName: userName,
            Date: dateStr,
            SlotID: selectedSlot.ID,
            SlotName: selectedSlot.Name,
            StartTime: selectedSlot.StartTime,
            EndTime: selectedSlot.EndTime,
            OriginalStartTime: selectedSlot.StartTime,
            OriginalEndTime: selectedSlot.EndTime,
            BreakStart: calculateBreakStart(selectedSlot),
            BreakEnd: calculateBreakEnd(selectedSlot),
            LunchStart: calculateLunchStart(selectedSlot),
            LunchEnd: calculateLunchEnd(selectedSlot),
            IsDST: dstStatus.isDST,
            Status: 'PENDING',
            GeneratedBy: generatedBy,
            ApprovedBy: null,
            NotificationSent: false,
            CreatedAt: new Date(),
            UpdatedAt: new Date(),
            RecurringScheduleID: null,
            SwapRequestID: null,
            Priority: options.priority || 2,
            Notes: options.notes || `Generated from selected slot: ${selectedSlot.Name}`,
            Location: selectedSlot.Location || '',
            Department: selectedSlot.Department || ''
          };

          generatedSchedules.push(schedule);
          console.log(`✅ Generated schedule for ${userName} on ${dateStr} using slot: ${selectedSlot.Name}`);

        } catch (userError) {
          conflicts.push({
            user: userName,
            date: dateStr,
            error: userError.message,
            type: 'GENERATION_ERROR'
          });
        }
      });
    }

    // Save generated schedules using ScheduleUtilities
    if (generatedSchedules.length > 0) {
      saveSchedulesToSheet(generatedSchedules);
      console.log(`💾 Saved ${generatedSchedules.length} schedules`);
    }

    // Return comprehensive result with shift slot information
    const result = {
      success: true,
      generated: generatedSchedules.length,
      conflicts: conflicts,
      dstChanges: dstChanges,
      message: `Successfully generated ${generatedSchedules.length} schedules using ${shiftSlots.length} shift slot(s)`,
      schedules: generatedSchedules.slice(0, 10), // Return first 10 for preview
      userCount: usersToSchedule.length,
      shiftSlotsUsed: shiftSlots.length,
      selectedSlots: shiftSlotIds && shiftSlotIds.length > 0 ? shiftSlotIds : null
    };

    console.log('✅ Enhanced schedule generation completed:', result);
    return result;

  } catch (error) {
    console.error('❌ Enhanced schedule generation failed:', error);
    safeWriteError('clientGenerateSchedulesEnhanced', error);
    return {
      success: false,
      error: error.message,
      generated: 0,
      conflicts: [],
      dstChanges: []
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
      // Create row data using proper header order from ScheduleUtilities
      const rowData = SCHEDULE_GENERATION_HEADERS.map(header => schedule[header] || '');
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

    // Use ScheduleUtilities to read schedules
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

    console.log(`📊 Total schedules in sheet: ${schedules.length}`);

    let filteredSchedules = schedules;

    // Apply filters
    if (filters.startDate) {
      const startDate = new Date(filters.startDate);
      filteredSchedules = filteredSchedules.filter(s => new Date(s.Date) >= startDate);
    }

    if (filters.endDate) {
      const endDate = new Date(filters.endDate);
      filteredSchedules = filteredSchedules.filter(s => new Date(s.Date) <= endDate);
    }

    if (filters.userId) {
      filteredSchedules = filteredSchedules.filter(s => s.UserID === filters.userId);
    }

    if (filters.userName) {
      filteredSchedules = filteredSchedules.filter(s => s.UserName === filters.userName);
    }

    if (filters.status) {
      filteredSchedules = filteredSchedules.filter(s => s.Status === filters.status);
    }

    if (filters.department) {
      filteredSchedules = filteredSchedules.filter(s => s.Department === filters.department);
    }

    // Sort by date (newest first)
    filteredSchedules.sort((a, b) => new Date(b.Date) - new Date(a.Date));

    console.log(`✅ Returning ${filteredSchedules.length} filtered schedules`);

    return {
      success: true,
      schedules: filteredSchedules,
      total: filteredSchedules.length,
      filters: filters
    };

  } catch (error) {
    console.error('❌ Error getting schedules:', error);
    safeWriteError('clientGetAllSchedules', error);
    return {
      success: false,
      error: error.message,
      schedules: [],
      total: 0
    };
  }
}

/**
 * Import schedules from uploaded data
 */
function clientImportSchedules(importRequest = {}) {
  try {
    const schedules = Array.isArray(importRequest.schedules) ? importRequest.schedules : [];
    if (schedules.length === 0) {
      throw new Error('No schedules were provided for import.');
    }

    const metadata = importRequest.metadata || {};
    const timeZone = typeof Session !== 'undefined' ? Session.getScriptTimeZone() : 'UTC';
    const now = new Date();
    const nowIso = Utilities.formatDate(now, timeZone, "yyyy-MM-dd'T'HH:mm:ss");

    const userLookup = buildScheduleUserLookup();
    const normalizedNew = schedules
      .map(raw => normalizeImportedScheduleRecord(raw, metadata, userLookup, nowIso, timeZone))
      .filter(record => record);

    if (normalizedNew.length === 0) {
      throw new Error('No valid schedules were found in the uploaded file.');
    }

    const existingRecords = readScheduleSheet(SCHEDULE_GENERATION_SHEET) || [];
    const replaceExisting = metadata.replaceExisting === true;

    const dateObjects = normalizedNew
      .map(record => new Date(record.Date))
      .filter(date => !isNaN(date.getTime()));

    let minDate = null;
    let maxDate = null;
    if (dateObjects.length > 0) {
      minDate = new Date(Math.min.apply(null, dateObjects));
      maxDate = new Date(Math.max.apply(null, dateObjects));
    }

    if (metadata.startDate) {
      metadata.startDate = normalizeDateForSheet(metadata.startDate, timeZone);
    } else if (metadata.startWeekDate) {
      metadata.startDate = normalizeDateForSheet(metadata.startWeekDate, timeZone);
    }

    if (metadata.endDate) {
      metadata.endDate = normalizeDateForSheet(metadata.endDate, timeZone);
    } else if (metadata.endWeekDate) {
      metadata.endDate = normalizeDateForSheet(metadata.endWeekDate, timeZone);
    }

    const newKeys = new Set(normalizedNew.map(record => `${normalizeUserKey(record.UserName || record.UserID)}::${record.Date}`));
    let replacedCount = 0;

    const retainedRecords = existingRecords.filter(existing => {
      const existingDate = normalizeDateForSheet(existing.Date, timeZone);
      if (!existingDate) {
        return true;
      }

      const key = `${normalizeUserKey(existing.UserName || existing.UserID)}::${existingDate}`;

      if (replaceExisting && minDate && maxDate) {
        const existingDateObj = new Date(existingDate);
        if (!isNaN(existingDateObj.getTime()) && existingDateObj >= minDate && existingDateObj <= maxDate) {
          replacedCount++;
          return false;
        }
      }

      if (newKeys.has(key)) {
        replacedCount++;
        return false;
      }

      return true;
    });

    const normalizedMin = minDate ? normalizeDateForSheet(minDate, timeZone) : '';
    const normalizedMax = maxDate ? normalizeDateForSheet(maxDate, timeZone) : '';

    const summary = typeof metadata.summary === 'object' && metadata.summary !== null ? metadata.summary : {};
    if (metadata.startDate && !summary.startDate) {
      summary.startDate = metadata.startDate;
    } else if (normalizedMin && !summary.startDate) {
      summary.startDate = normalizedMin;
    }
    if (metadata.endDate && !summary.endDate) {
      summary.endDate = metadata.endDate;
    } else if (normalizedMax && !summary.endDate) {
      summary.endDate = normalizedMax;
    }
    if (typeof summary.totalAssignments !== 'number') {
      summary.totalAssignments = normalizedNew.length;
    }
    if (typeof summary.totalShifts !== 'number') {
      summary.totalShifts = normalizedNew.length;
    }
    if (typeof summary.dayCount !== 'number' || summary.dayCount <= 0) {
      summary.dayCount = calculateDaySpanCount(metadata.startDate, metadata.endDate, minDate, maxDate);
    }
    metadata.summary = summary;

    if (!metadata.dayCount) {
      const computedDays = calculateDaySpanCount(metadata.startDate, metadata.endDate, minDate, maxDate);
      if (computedDays) {
        metadata.dayCount = computedDays;
      }
    }

    const combinedRecords = retainedRecords.concat(normalizedNew);

    combinedRecords.sort((a, b) => {
      const dateA = new Date(a.Date || 0);
      const dateB = new Date(b.Date || 0);
      if (dateA.getTime() !== dateB.getTime()) {
        return dateA - dateB;
      }
      const nameA = (a.UserName || '').toString();
      const nameB = (b.UserName || '').toString();
      return nameA.localeCompare(nameB);
    });

    writeToScheduleSheet(SCHEDULE_GENERATION_SHEET, combinedRecords);
    invalidateScheduleCaches();

    return {
      success: true,
      importedCount: normalizedNew.length,
      replacedCount,
      totalAfterImport: combinedRecords.length,
      range: {
        start: normalizedMin,
        end: normalizedMax
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

/**
 * Import schedules from uploaded data
 */
function clientImportSchedules(importRequest = {}) {
  try {
    const schedules = Array.isArray(importRequest.schedules) ? importRequest.schedules : [];
    if (schedules.length === 0) {
      throw new Error('No schedules were provided for import.');
    }

    const metadata = importRequest.metadata || {};
    const timeZone = typeof Session !== 'undefined' ? Session.getScriptTimeZone() : 'UTC';
    const now = new Date();
    const nowIso = Utilities.formatDate(now, timeZone, "yyyy-MM-dd'T'HH:mm:ss");

    const userLookup = buildScheduleUserLookup();
    const normalizedNew = schedules
      .map(raw => normalizeImportedScheduleRecord(raw, metadata, userLookup, nowIso, timeZone))
      .filter(record => record);

    if (normalizedNew.length === 0) {
      throw new Error('No valid schedules were found in the uploaded file.');
    }

    const existingRecords = readScheduleSheet(SCHEDULE_GENERATION_SHEET) || [];
    const replaceExisting = metadata.replaceExisting === true;

    const dateObjects = normalizedNew
      .map(record => new Date(record.Date))
      .filter(date => !isNaN(date.getTime()));

    let minDate = null;
    let maxDate = null;
    if (dateObjects.length > 0) {
      minDate = new Date(Math.min.apply(null, dateObjects));
      maxDate = new Date(Math.max.apply(null, dateObjects));
    }

    if (metadata.startDate) {
      metadata.startDate = normalizeDateForSheet(metadata.startDate, timeZone);
    } else if (metadata.startWeekDate) {
      metadata.startDate = normalizeDateForSheet(metadata.startWeekDate, timeZone);
    }

    if (metadata.endDate) {
      metadata.endDate = normalizeDateForSheet(metadata.endDate, timeZone);
    } else if (metadata.endWeekDate) {
      metadata.endDate = normalizeDateForSheet(metadata.endWeekDate, timeZone);
    }

    const newKeys = new Set(normalizedNew.map(record => `${normalizeUserKey(record.UserName || record.UserID)}::${record.Date}`));
    let replacedCount = 0;

    const retainedRecords = existingRecords.filter(existing => {
      const existingDate = normalizeDateForSheet(existing.Date, timeZone);
      if (!existingDate) {
        return true;
      }

      const key = `${normalizeUserKey(existing.UserName || existing.UserID)}::${existingDate}`;

      if (replaceExisting && minDate && maxDate) {
        const existingDateObj = new Date(existingDate);
        if (!isNaN(existingDateObj.getTime()) && existingDateObj >= minDate && existingDateObj <= maxDate) {
          replacedCount++;
          return false;
        }
      }

      if (newKeys.has(key)) {
        replacedCount++;
        return false;
      }

      return true;
    });

    const normalizedMin = minDate ? normalizeDateForSheet(minDate, timeZone) : '';
    const normalizedMax = maxDate ? normalizeDateForSheet(maxDate, timeZone) : '';

    const summary = typeof metadata.summary === 'object' && metadata.summary !== null ? metadata.summary : {};
    if (metadata.startDate && !summary.startDate) {
      summary.startDate = metadata.startDate;
    } else if (normalizedMin && !summary.startDate) {
      summary.startDate = normalizedMin;
    }
    if (metadata.endDate && !summary.endDate) {
      summary.endDate = metadata.endDate;
    } else if (normalizedMax && !summary.endDate) {
      summary.endDate = normalizedMax;
    }
    if (typeof summary.totalAssignments !== 'number') {
      summary.totalAssignments = normalizedNew.length;
    }
    if (typeof summary.totalShifts !== 'number') {
      summary.totalShifts = normalizedNew.length;
    }
    if (typeof summary.dayCount !== 'number' || summary.dayCount <= 0) {
      summary.dayCount = calculateDaySpanCount(metadata.startDate, metadata.endDate, minDate, maxDate);
    }
    metadata.summary = summary;

    if (!metadata.dayCount) {
      const computedDays = calculateDaySpanCount(metadata.startDate, metadata.endDate, minDate, maxDate);
      if (computedDays) {
        metadata.dayCount = computedDays;
      }
    }

    const combinedRecords = retainedRecords.concat(normalizedNew);

    combinedRecords.sort((a, b) => {
      const dateA = new Date(a.Date || 0);
      const dateB = new Date(b.Date || 0);
      if (dateA.getTime() !== dateB.getTime()) {
        return dateA - dateB;
      }
      const nameA = (a.UserName || '').toString();
      const nameB = (b.UserName || '').toString();
      return nameA.localeCompare(nameB);
    });

    writeToScheduleSheet(SCHEDULE_GENERATION_SHEET, combinedRecords);
    invalidateScheduleCaches();

    return {
      success: true,
      importedCount: normalizedNew.length,
      replacedCount,
      totalAfterImport: combinedRecords.length,
      range: {
        start: normalizedMin,
        end: normalizedMax
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
  try {
    const schedules = Array.isArray(importRequest.schedules) ? importRequest.schedules : [];
    if (schedules.length === 0) {
      throw new Error('No schedules were provided for import.');
    }

    const metadata = importRequest.metadata || {};
    const timeZone = typeof Session !== 'undefined' ? Session.getScriptTimeZone() : 'UTC';
    const now = new Date();
    const nowIso = Utilities.formatDate(now, timeZone, "yyyy-MM-dd'T'HH:mm:ss");

    const userLookup = buildScheduleUserLookup();
    const normalizedNew = schedules
      .map(raw => normalizeImportedScheduleRecord(raw, metadata, userLookup, nowIso, timeZone))
      .filter(record => record);

    if (normalizedNew.length === 0) {
      throw new Error('No valid schedules were found in the uploaded file.');
    }

    const existingRecords = readScheduleSheet(SCHEDULE_GENERATION_SHEET) || [];
    const replaceExisting = metadata.replaceExisting === true;

    const dateObjects = normalizedNew
      .map(record => new Date(record.Date))
      .filter(date => !isNaN(date.getTime()));

    let minDate = null;
    let maxDate = null;
    if (dateObjects.length > 0) {
      minDate = new Date(Math.min.apply(null, dateObjects));
      maxDate = new Date(Math.max.apply(null, dateObjects));
    }

    if (metadata.startDate) {
      metadata.startDate = normalizeDateForSheet(metadata.startDate, timeZone);
    } else if (metadata.startWeekDate) {
      metadata.startDate = normalizeDateForSheet(metadata.startWeekDate, timeZone);
    }

    if (metadata.endDate) {
      metadata.endDate = normalizeDateForSheet(metadata.endDate, timeZone);
    } else if (metadata.endWeekDate) {
      metadata.endDate = normalizeDateForSheet(metadata.endWeekDate, timeZone);
    }

    const newKeys = new Set(normalizedNew.map(record => `${normalizeUserKey(record.UserName || record.UserID)}::${record.Date}`));
    let replacedCount = 0;

    const retainedRecords = existingRecords.filter(existing => {
      const existingDate = normalizeDateForSheet(existing.Date, timeZone);
      if (!existingDate) {
        return true;
      }

      const key = `${normalizeUserKey(existing.UserName || existing.UserID)}::${existingDate}`;

      if (replaceExisting && minDate && maxDate) {
        const existingDateObj = new Date(existingDate);
        if (!isNaN(existingDateObj.getTime()) && existingDateObj >= minDate && existingDateObj <= maxDate) {
          replacedCount++;
          return false;
        }
      }

      if (newKeys.has(key)) {
        replacedCount++;
        return false;
      }

      return true;
    });

    const normalizedMin = minDate ? normalizeDateForSheet(minDate, timeZone) : '';
    const normalizedMax = maxDate ? normalizeDateForSheet(maxDate, timeZone) : '';

    const summary = typeof metadata.summary === 'object' && metadata.summary !== null ? metadata.summary : {};
    if (metadata.startDate && !summary.startDate) {
      summary.startDate = metadata.startDate;
    } else if (normalizedMin && !summary.startDate) {
      summary.startDate = normalizedMin;
    }
    if (metadata.endDate && !summary.endDate) {
      summary.endDate = metadata.endDate;
    } else if (normalizedMax && !summary.endDate) {
      summary.endDate = normalizedMax;
    }
    if (typeof summary.totalAssignments !== 'number') {
      summary.totalAssignments = normalizedNew.length;
    }
    if (typeof summary.totalShifts !== 'number') {
      summary.totalShifts = normalizedNew.length;
    }
    if (typeof summary.dayCount !== 'number' || summary.dayCount <= 0) {
      summary.dayCount = calculateDaySpanCount(metadata.startDate, metadata.endDate, minDate, maxDate);
    }
    metadata.summary = summary;

    if (!metadata.dayCount) {
      const computedDays = calculateDaySpanCount(metadata.startDate, metadata.endDate, minDate, maxDate);
      if (computedDays) {
        metadata.dayCount = computedDays;
      }
    }

    const combinedRecords = retainedRecords.concat(normalizedNew);

    combinedRecords.sort((a, b) => {
      const dateA = new Date(a.Date || 0);
      const dateB = new Date(b.Date || 0);
      if (dateA.getTime() !== dateB.getTime()) {
        return dateA - dateB;
      }
      const nameA = (a.UserName || '').toString();
      const nameB = (b.UserName || '').toString();
      return nameA.localeCompare(nameB);
    });

    writeToScheduleSheet(SCHEDULE_GENERATION_SHEET, combinedRecords);
    invalidateScheduleCaches();

    return {
      success: true,
      importedCount: normalizedNew.length,
      replacedCount,
      totalAfterImport: combinedRecords.length,
      range: {
        start: normalizedMin,
        end: normalizedMax
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
  try {
    const schedules = Array.isArray(importRequest.schedules) ? importRequest.schedules : [];
    if (schedules.length === 0) {
      throw new Error('No schedules were provided for import.');
    }

    const metadata = importRequest.metadata || {};
    const timeZone = typeof Session !== 'undefined' ? Session.getScriptTimeZone() : 'UTC';
    const now = new Date();
    const nowIso = Utilities.formatDate(now, timeZone, "yyyy-MM-dd'T'HH:mm:ss");

    const userLookup = buildScheduleUserLookup();
    const normalizedNew = schedules
      .map(raw => normalizeImportedScheduleRecord(raw, metadata, userLookup, nowIso, timeZone))
      .filter(record => record);

    if (normalizedNew.length === 0) {
      throw new Error('No valid schedules were found in the uploaded file.');
    }

    const existingRecords = readScheduleSheet(SCHEDULE_GENERATION_SHEET) || [];
    const replaceExisting = metadata.replaceExisting === true;

    const dateObjects = normalizedNew
      .map(record => new Date(record.Date))
      .filter(date => !isNaN(date.getTime()));

    let minDate = null;
    let maxDate = null;
    if (dateObjects.length > 0) {
      minDate = new Date(Math.min.apply(null, dateObjects));
      maxDate = new Date(Math.max.apply(null, dateObjects));
    }

    if (metadata.startWeekDate) {
      metadata.startWeekDate = normalizeDateForSheet(metadata.startWeekDate, timeZone);
    }
    if (metadata.endWeekDate) {
      metadata.endWeekDate = normalizeDateForSheet(metadata.endWeekDate, timeZone);
    }

    const newKeys = new Set(normalizedNew.map(record => `${normalizeUserKey(record.UserName || record.UserID)}::${record.Date}`));
    let replacedCount = 0;

    const retainedRecords = existingRecords.filter(existing => {
      const existingDate = normalizeDateForSheet(existing.Date, timeZone);
      if (!existingDate) {
        return true;
      }

      const key = `${normalizeUserKey(existing.UserName || existing.UserID)}::${existingDate}`;

      if (replaceExisting && minDate && maxDate) {
        const existingDateObj = new Date(existingDate);
        if (!isNaN(existingDateObj.getTime()) && existingDateObj >= minDate && existingDateObj <= maxDate) {
          replacedCount++;
          return false;
        }
      }

      if (newKeys.has(key)) {
        replacedCount++;
        return false;
      }

      return true;
    });

    const normalizedMin = minDate ? normalizeDateForSheet(minDate, timeZone) : '';
    const normalizedMax = maxDate ? normalizeDateForSheet(maxDate, timeZone) : '';

    const summary = typeof metadata.summary === 'object' && metadata.summary !== null ? metadata.summary : {};
    if (normalizedMin && !summary.startDate) {
      summary.startDate = normalizedMin;
    }
    if (normalizedMax && !summary.endDate) {
      summary.endDate = normalizedMax;
    }
    if (typeof summary.totalAssignments !== 'number') {
      summary.totalAssignments = normalizedNew.length;
    }
    if (typeof summary.totalShifts !== 'number') {
      summary.totalShifts = normalizedNew.length;
    }
    metadata.summary = summary;

    if (!metadata.weekCount) {
      const computedWeeks = calculateWeekSpanCount(metadata.startWeekDate, metadata.endWeekDate, minDate, maxDate);
      if (computedWeeks) {
        metadata.weekCount = computedWeeks;
      }
    }

    const combinedRecords = retainedRecords.concat(normalizedNew);

    combinedRecords.sort((a, b) => {
      const dateA = new Date(a.Date || 0);
      const dateB = new Date(b.Date || 0);
      if (dateA.getTime() !== dateB.getTime()) {
        return dateA - dateB;
      }
      const nameA = (a.UserName || '').toString();
      const nameB = (b.UserName || '').toString();
      return nameA.localeCompare(nameB);
    });

    writeToScheduleSheet(SCHEDULE_GENERATION_SHEET, combinedRecords);
    invalidateScheduleCaches();

    return {
      success: true,
      importedCount: normalizedNew.length,
      replacedCount,
      totalAfterImport: combinedRecords.length,
      range: {
        start: normalizedMin,
        end: normalizedMax
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

    if (!SCHEDULE_CONFIG.SUPPORTED_COUNTRIES.includes(countryCode)) {
      return {
        success: false,
        error: `Country ${countryCode} not supported. Supported countries: ${SCHEDULE_CONFIG.SUPPORTED_COUNTRIES.join(', ')}`,
        holidays: []
      };
    }

    const holidays = getUpdatedHolidays(countryCode, year);
    const isPrimary = countryCode === SCHEDULE_CONFIG.PRIMARY_COUNTRY;

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
        supportedCountries: SCHEDULE_CONFIG.SUPPORTED_COUNTRIES,
        working: holidays.success,
        primaryCountry: SCHEDULE_CONFIG.PRIMARY_COUNTRY
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
    console.log('✅ Approving schedules:', scheduleIds);
    
    const sheet = getScheduleSpreadsheet().getSheetByName(SCHEDULE_GENERATION_SHEET);
    if (!sheet) {
      throw new Error('Schedules sheet not found');
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const statusCol = headers.indexOf('Status') + 1;
    const approvedByCol = headers.indexOf('ApprovedBy') + 1;
    const updatedAtCol = headers.indexOf('UpdatedAt') + 1;

    let updated = 0;

    for (let i = 1; i < data.length; i++) {
      const scheduleId = data[i][0]; // ID is first column
      if (scheduleIds.includes(scheduleId)) {
        sheet.getRange(i + 1, statusCol).setValue('APPROVED');
        sheet.getRange(i + 1, approvedByCol).setValue(approvingUserId || 'System');
        sheet.getRange(i + 1, updatedAtCol).setValue(new Date());
        updated++;
      }
    }

    SpreadsheetApp.flush();
    invalidateScheduleCaches();

    return {
      success: true,
      message: `Approved ${updated} schedules`,
      approved: updated
    };

  } catch (error) {
    console.error('Error approving schedules:', error);
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
    console.log('❌ Rejecting schedules:', scheduleIds);
    
    const sheet = getScheduleSpreadsheet().getSheetByName(SCHEDULE_GENERATION_SHEET);
    if (!sheet) {
      throw new Error('Schedules sheet not found');
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const statusCol = headers.indexOf('Status') + 1;
    const notesCol = headers.indexOf('Notes') + 1;
    const updatedAtCol = headers.indexOf('UpdatedAt') + 1;

    let updated = 0;

    for (let i = 1; i < data.length; i++) {
      const scheduleId = data[i][0]; // ID is first column
      if (scheduleIds.includes(scheduleId)) {
        sheet.getRange(i + 1, statusCol).setValue('REJECTED');
        if (reason) {
          const existingNotes = data[i][notesCol - 1] || '';
          const newNotes = existingNotes + (existingNotes ? '; ' : '') + 'Rejected: ' + reason;
          sheet.getRange(i + 1, notesCol).setValue(newNotes);
        }
        sheet.getRange(i + 1, updatedAtCol).setValue(new Date());
        updated++;
      }
    }

    SpreadsheetApp.flush();
    invalidateScheduleCaches();

    return {
      success: true,
      message: `Rejected ${updated} schedules`,
      rejected: updated
    };

  } catch (error) {
    console.error('Error rejecting schedules:', error);
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
 * Check if schedule exists for user on date - uses ScheduleUtilities
 */
function checkExistingSchedule(userName, date) {
  try {
    const schedules = readScheduleSheet(SCHEDULE_GENERATION_SHEET) || [];
    return schedules.find(s => s.UserName === userName && s.Date === date);
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

  const dateStr = normalizeDateForSheet(raw.Date, timeZone);
  const userName = (raw.UserName || '').toString().trim();

  if (!userName || !dateStr) {
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
    Date: dateStr,
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

  return {
    ID: resolve(['ID', 'ScheduleID', 'Schedule Id', 'RecordID'], uuid),
    UserID: userId || normalizeUserIdValue(userName),
    UserName: userName || userId,
    Date: normalizeDate(scheduleDate),
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

console.log('✅ Enhanced Schedule Management Backend v4.1 loaded successfully');
console.log('🔧 Features: ScheduleUtilities integration, MainUtilities user management, dedicated spreadsheet support');
console.log('🎯 Ready for production use with comprehensive diagnostics and proper utility integration');
console.log('📊 Integrated: User/Campaign management from MainUtilities, Sheet management from ScheduleUtilities');
