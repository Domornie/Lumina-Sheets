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

// ────────────────────────────────────────────────────────────────────────────
// USER MANAGEMENT FUNCTIONS - Integrated with MainUtilities
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get users for schedule management with manager filtering
 * Uses MainUtilities user functions with campaign support
 */
function clientGetScheduleUsers(requestingUserId, campaignId = null) {
  try {
    console.log('🔍 Getting schedule users for:', requestingUserId, 'campaign:', campaignId);

    // Use MainUtilities to get all users
    const allUsers = readSheet(USERS_SHEET) || [];
    if (allUsers.length === 0) {
      console.warn('No users found in Users sheet');
      return [];
    }

    let filteredUsers = allUsers;

    // Filter by campaign if specified - use MainUtilities campaign functions
    if (campaignId) {
      const campaignUsers = getUsersByCampaign(campaignId);
      const campaignUserIds = new Set(campaignUsers.map(u => u.ID));
      filteredUsers = allUsers.filter(user => campaignUserIds.has(user.ID));
    }

    // Apply manager permissions using MainUtilities functions
    if (requestingUserId) {
      const requestingUser = allUsers.find(u => String(u.ID) === String(requestingUserId));
      
      if (requestingUser) {
        // Use MainUtilities admin check
        if (requestingUser.IsAdmin === 'TRUE' || requestingUser.IsAdmin === true) {
          // Admin can see all users - no additional filtering needed
        } else {
          // Check managed campaigns using MainUtilities
          const managedCampaigns = getUserManagedCampaigns(requestingUserId);
          const managedCampaignIds = new Set(managedCampaigns.map(c => c.ID));
          
          // Get users from managed campaigns plus requesting user
          const userCampaigns = getUserCampaignsSafe(requestingUserId).map(uc => uc.campaignId);
          const accessibleUsers = new Set([requestingUserId]);
          
          // Add users from managed campaigns
          managedCampaignIds.forEach(campaignId => {
            const campaignUsers = getUsersByCampaign(campaignId);
            campaignUsers.forEach(u => accessibleUsers.add(u.ID));
          });
          
          filteredUsers = filteredUsers.filter(user => accessibleUsers.has(user.ID));
        }
      }
    }

    // Transform to schedule-friendly format
    const scheduleUsers = filteredUsers
      .filter(user => user && user.ID && (user.UserName || user.FullName))
      .filter(user => user.EmploymentStatus === 'Active' || !user.EmploymentStatus)
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
          canLogin: user.CanLogin === 'TRUE' || user.CanLogin === true,
          isActive: true
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
    safeWriteError('clientCreateShiftSlot', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function clientCreateEnhancedShiftSlot(slotData) {
  return clientCreateShiftSlot(slotData);
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
        DaysOfWeekArray: slot.DaysOfWeek ?
          slot.DaysOfWeek.split(',').map(d => parseInt(d.trim(), 10)).filter(d => !isNaN(d)) :
          [1, 2, 3, 4, 5]
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
    const schedules = readScheduleSheet(SCHEDULE_GENERATION_SHEET) || [];

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

console.log('✅ Enhanced Schedule Management Backend v4.1 loaded successfully');
console.log('🔧 Features: ScheduleUtilities integration, MainUtilities user management, dedicated spreadsheet support');
console.log('🎯 Ready for production use with comprehensive diagnostics and proper utility integration');
console.log('📊 Integrated: User/Campaign management from MainUtilities, Sheet management from ScheduleUtilities');
