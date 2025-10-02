/**
 * ScheduleUtilities.gs
 * Constants, sheet definitions, and utility functions for Schedule Management
 * Updated to include dedicated schedule spreadsheet management
 */

// ────────────────────────────────────────────────────────────────────────────
// SCHEDULE SPREADSHEET CONFIGURATION
// ────────────────────────────────────────────────────────────────────────────

// Schedule Management Spreadsheet ID - UPDATE THIS WITH YOUR ACTUAL SCHEDULE SPREADSHEET ID
const SCHEDULE_SPREADSHEET_ID = '1owlD-RdNBYgpnccOPp6zP4ndzg0W4IWtoQeWC3NIUL0'; // TODO: Set your dedicated schedule management spreadsheet ID here

// If no dedicated schedule spreadsheet ID is set, these functions will fall back to the main spreadsheet
const FALLBACK_TO_MAIN_SPREADSHEET = true;

/**
 * Get the Schedule Management spreadsheet
 * Returns the dedicated schedule spreadsheet if ID is configured, otherwise falls back to main/active spreadsheet
 */
function getScheduleSpreadsheet() {
  try {
    // First, try to use the dedicated schedule spreadsheet ID
    if (SCHEDULE_SPREADSHEET_ID && SCHEDULE_SPREADSHEET_ID.trim() !== '') {
      console.log('Using dedicated schedule spreadsheet:', SCHEDULE_SPREADSHEET_ID);
      return SpreadsheetApp.openById(SCHEDULE_SPREADSHEET_ID);
    }

    // Fallback: Try to use the main spreadsheet ID if available
    if (FALLBACK_TO_MAIN_SPREADSHEET) {
      // Check if MAIN_SPREADSHEET_ID is available from the global scope
      if (typeof G !== 'undefined' && G.MAIN_SPREADSHEET_ID && G.MAIN_SPREADSHEET_ID.trim() !== '') {
        console.log('Using main spreadsheet for schedules:', G.MAIN_SPREADSHEET_ID);
        return SpreadsheetApp.openById(G.MAIN_SPREADSHEET_ID);
      }

      // Final fallback: Use the active spreadsheet
      console.log('Using active spreadsheet for schedules (fallback)');
      return SpreadsheetApp.getActiveSpreadsheet();
    } else {
      throw new Error('SCHEDULE_SPREADSHEET_ID not configured and fallback disabled');
    }

  } catch (error) {
    console.error('Error accessing schedule spreadsheet:', error);

    // Emergency fallback to active spreadsheet
    if (FALLBACK_TO_MAIN_SPREADSHEET) {
      console.log('Emergency fallback to active spreadsheet');
      return SpreadsheetApp.getActiveSpreadsheet();
    } else {
      throw new Error(`Cannot access schedule spreadsheet: ${error.message}`);
    }
  }
}

/**
 * Validate schedule spreadsheet configuration
 * Returns information about the current spreadsheet configuration
 */
function validateScheduleSpreadsheetConfig() {
  try {
    const config = {
      hasScheduleSpreadsheetId: SCHEDULE_SPREADSHEET_ID && SCHEDULE_SPREADSHEET_ID.trim() !== '',
      scheduleSpreadsheetId: SCHEDULE_SPREADSHEET_ID,
      fallbackEnabled: FALLBACK_TO_MAIN_SPREADSHEET,
      currentSpreadsheet: null,
      spreadsheetName: '',
      canAccess: false
    };

    try {
      const ss = getScheduleSpreadsheet();
      config.currentSpreadsheet = ss.getId();
      config.spreadsheetName = ss.getName();
      config.canAccess = true;
    } catch (error) {
      config.canAccess = false;
      config.error = error.message;
    }

    return config;
  } catch (error) {
    return {
      hasScheduleSpreadsheetId: false,
      canAccess: false,
      error: error.message
    };
  }
}

/**
 * Setup function to configure the schedule spreadsheet ID
 * Call this function to set or update the schedule spreadsheet ID
 */
function setScheduleSpreadsheetId(spreadsheetId) {
  try {
    // This is a helper function - in practice, you'll need to manually update the constant above
    // or store this in a configuration sheet
    console.log('To set the schedule spreadsheet ID, update SCHEDULE_SPREADSHEET_ID constant to:', spreadsheetId);

    // Validate the spreadsheet ID by trying to access it
    const testSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const name = testSpreadsheet.getName();

    return {
      success: true,
      message: `Schedule spreadsheet ID validated. Spreadsheet name: ${name}`,
      instructions: 'Update the SCHEDULE_SPREADSHEET_ID constant in ScheduleUtilities.gs'
    };

  } catch (error) {
    return {
      success: false,
      error: `Invalid spreadsheet ID: ${error.message}`
    };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// SHEET CONSTANTS AND IDs
// ────────────────────────────────────────────────────────────────────────────

// Main schedule management sheets
const SCHEDULE_GENERATION_SHEET = "GeneratedSchedules";
const SHIFT_SLOTS_SHEET = "ShiftSlots";
const SHIFT_SWAPS_SHEET = "ShiftSwaps";
const SCHEDULE_TEMPLATES_SHEET = "ScheduleTemplates";
const SCHEDULE_NOTIFICATIONS_SHEET = "ScheduleNotifications";
const SCHEDULE_ADHERENCE_SHEET = "ScheduleAdherence";
const SCHEDULE_CONFLICTS_SHEET = "ScheduleConflicts";
const RECURRING_SCHEDULES_SHEET = "RecurringSchedules";

// Attendance and holiday sheets
const ATTENDANCE_STATUS_SHEET = "AttendanceStatus";
const USER_HOLIDAY_PAY_STATUS_SHEET = "UserHolidayPayStatus";
const HOLIDAYS_SHEET = "Holidays";

// ────────────────────────────────────────────────────────────────────────────
// SHEET HEADERS DEFINITIONS
// ────────────────────────────────────────────────────────────────────────────

const SCHEDULE_GENERATION_HEADERS = [
  'ID', 'UserID', 'UserName', 'Date', 'SlotID', 'SlotName', 'StartTime', 'EndTime',
  'OriginalStartTime', 'OriginalEndTime', 'BreakStart', 'BreakEnd', 'LunchStart', 'LunchEnd',
  'IsDST', 'Status', 'GeneratedBy', 'ApprovedBy', 'NotificationSent', 'CreatedAt', 'UpdatedAt',
  'RecurringScheduleID', 'SwapRequestID', 'Priority', 'Notes', 'Location', 'Department'
];

const SHIFT_SLOTS_HEADERS = [
  'ID', 'Name', 'StartTime', 'EndTime', 'DaysOfWeek', 'Department', 'Location',
  'MaxCapacity', 'MinCoverage', 'Priority', 'Description',
  'BreakDuration', 'LunchDuration', 'Break1Duration', 'Break2Duration',
  'EnableStaggeredBreaks', 'BreakGroups', 'StaggerInterval', 'MinCoveragePct',
  'EnableOvertime', 'MaxDailyOT', 'MaxWeeklyOT', 'OTApproval', 'OTRate', 'OTPolicy',
  'AllowSwaps', 'WeekendPremium', 'HolidayPremium', 'AutoAssignment',
  'RestPeriod', 'NotificationLead', 'HandoverTime',
  'OvertimePolicy', 'IsActive', 'CreatedBy', 'CreatedAt', 'UpdatedAt'
];

const SHIFT_SWAPS_HEADERS = [
  'ID', 'RequestorUserID', 'RequestorUserName', 'TargetUserID', 'TargetUserName',
  'RequestorScheduleID', 'TargetScheduleID', 'SwapDate', 'Reason', 'Status',
  'ApprovedBy', 'RejectedBy', 'DecisionNotes', 'CreatedAt', 'UpdatedAt'
];

const SCHEDULE_TEMPLATES_HEADERS = [
  'ID', 'TemplateName', 'Description', 'SlotConfiguration', 'BreakConfiguration',
  'LunchConfiguration', 'RecurrencePattern', 'IsActive', 'CreatedBy', 'CreatedAt'
];

const SCHEDULE_NOTIFICATIONS_HEADERS = [
  'ID', 'ScheduleID', 'UserID', 'NotificationType', 'SentAt', 'Method', 'Status', 'RetryCount'
];

const SCHEDULE_ADHERENCE_HEADERS = [
  'ID', 'ScheduleID', 'UserID', 'Date', 'ScheduledStart', 'ActualStart',
  'ScheduledEnd', 'ActualEnd', 'MinutesLate', 'MinutesEarly', 'AdherenceScore',
  'BreakAdherence', 'LunchAdherence', 'CreatedAt'
];

const SCHEDULE_CONFLICTS_HEADERS = [
  'ID', 'ScheduleID1', 'ScheduleID2', 'ConflictType', 'Severity', 'Resolution', 'CreatedAt'
];

const RECURRING_SCHEDULES_HEADERS = [
  'ID', 'UserID', 'UserName', 'SlotID', 'SlotName', 'RecurrencePattern',
  'StartDate', 'EndDate', 'IsActive', 'CreatedBy', 'CreatedAt', 'UpdatedAt'
];

const ATTENDANCE_STATUS_HEADERS = [
  'ID', 'UserID', 'UserName', 'Date', 'Status', 'Notes', 'MarkedBy', 'CreatedAt', 'UpdatedAt'
];

const USER_HOLIDAY_PAY_STATUS_HEADERS = [
  'ID', 'UserID', 'UserName', 'CountryCode', 'IsPaid', 'Notes', 'CreatedAt', 'UpdatedAt'
];

const HOLIDAYS_HEADERS = [
  'ID', 'HolidayName', 'Date', 'AllDay', 'Notes', 'CreatedAt', 'UpdatedAt'
];

// ────────────────────────────────────────────────────────────────────────────
// BUSINESS LOGIC CONSTANTS
// ────────────────────────────────────────────────────────────────────────────

const SCHEDULE_STATUS = {
  PENDING: 'PENDING',
  APPROVED: 'APPROVED',
  REJECTED: 'REJECTED',
  CANCELLED: 'CANCELLED'
};

const SWAP_STATUS = {
  PENDING: 'PENDING',
  APPROVED: 'APPROVED',
  REJECTED: 'REJECTED',
  CANCELLED: 'CANCELLED'
};

const PRIORITY_LEVELS = {
  LOW: 1,
  NORMAL: 2,
  HIGH: 3,
  CRITICAL: 4
};

const ATTENDANCE_STATUSES = [
  'Present',
  'Absent',
  'Late',
  'Early Leave',
  'Sick Leave',
  'Bereavement',
  'Vacation',
  'Personal Leave',
  'Emergency Leave',
  'Training',
  'Holiday'
];

const OVERTIME_POLICIES = {
  NONE: 'NONE',
  LIMITED_15: 'LIMITED_15',
  LIMITED_30: 'LIMITED_30',
  LIMITED_60: 'LIMITED_60',
  UNLIMITED: 'UNLIMITED',
  APPROVAL_REQUIRED: 'APPROVAL_REQUIRED'
};

const NOTIFICATION_TYPES = {
  SCHEDULE_CREATED: 'SCHEDULE_CREATED',
  SCHEDULE_APPROVED: 'SCHEDULE_APPROVED',
  SCHEDULE_REJECTED: 'SCHEDULE_REJECTED',
  SWAP_REQUEST: 'SWAP_REQUEST',
  SWAP_APPROVED: 'SWAP_APPROVED',
  SWAP_REJECTED: 'SWAP_REJECTED',
  REMINDER: 'REMINDER'
};

const CONFLICT_TYPES = {
  USER_DOUBLE_BOOKING: 'USER_DOUBLE_BOOKING',
  SLOT_OVERCAPACITY: 'SLOT_OVERCAPACITY',
  TIME_OVERLAP: 'TIME_OVERLAP',
  INVALID_ASSIGNMENT: 'INVALID_ASSIGNMENT'
};

const RECURRENCE_PATTERNS = {
  DAILY: 'daily',
  WEEKLY: 'weekly',
  MONTHLY: 'monthly',
  CUSTOM: 'custom'
};

// Days of week constants (Sunday = 0, Monday = 1, etc.)
const DAYS_OF_WEEK = {
  SUNDAY: 0,
  MONDAY: 1,
  TUESDAY: 2,
  WEDNESDAY: 3,
  THURSDAY: 4,
  FRIDAY: 5,
  SATURDAY: 6
};

const WEEKDAYS = [1, 2, 3, 4, 5]; // Monday through Friday
const WEEKEND = [0, 6]; // Sunday and Saturday

// Configuration constants
const DEFAULT_BREAK_DURATION = 15; // minutes
const DEFAULT_LUNCH_DURATION = 60; // minutes
const DEFAULT_SHIFT_CAPACITY = 10; // maximum users per slot
const PUNCTUALITY_GRACE_PERIOD = 5; // minutes late before considered late

// ────────────────────────────────────────────────────────────────────────────
// ENHANCED SHEET MANAGEMENT UTILITIES WITH SCHEDULE SPREADSHEET SUPPORT
// ────────────────────────────────────────────────────────────────────────────
/**
 * Return the correct header array for a given schedule sheet name
 */
function getHeadersForSheet(sheetName) {
  const map = {};
  map[SCHEDULE_GENERATION_SHEET] = SCHEDULE_GENERATION_HEADERS;
  map[SHIFT_SLOTS_SHEET] = SHIFT_SLOTS_HEADERS;
  map[SHIFT_SWAPS_SHEET] = SHIFT_SWAPS_HEADERS;
  map[SCHEDULE_TEMPLATES_SHEET] = SCHEDULE_TEMPLATES_HEADERS;
  map[SCHEDULE_NOTIFICATIONS_SHEET] = SCHEDULE_NOTIFICATIONS_HEADERS;
  map[SCHEDULE_ADHERENCE_SHEET] = SCHEDULE_ADHERENCE_HEADERS;
  map[SCHEDULE_CONFLICTS_SHEET] = SCHEDULE_CONFLICTS_HEADERS;
  map[RECURRING_SCHEDULES_SHEET] = RECURRING_SCHEDULES_HEADERS;
  map[ATTENDANCE_STATUS_SHEET] = ATTENDANCE_STATUS_HEADERS;
  map[USER_HOLIDAY_PAY_STATUS_SHEET] = USER_HOLIDAY_PAY_STATUS_HEADERS;
  map[HOLIDAYS_SHEET] = HOLIDAYS_HEADERS;
  return map[sheetName] || null;
}

/**
 * Read data from a schedule sheet with caching
 */
function readScheduleSheet(sheetName) {
  try {
    // Try to get from cache first
    const cacheKey = `schedule_${sheetName}`;
    const cached = getFromCache(cacheKey);
    if (cached) {
      return cached;
    }

    const ss = getScheduleSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      console.warn(`Schedule sheet not found: ${sheetName}`);
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return [];
    }

    const headers = data[0];
    const rows = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });

    // Cache the result
    setInCache(cacheKey, rows);
    return rows;

  } catch (error) {
    console.error(`Error reading schedule sheet ${sheetName}:`, error);
    return [];
  }
}

/**
 * Write data to a schedule sheet (auto-creates sheet + headers if missing)
 */
function writeToScheduleSheet(sheetName, data) {
  try {
    const ss = getScheduleSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);

    // Ensure sheet exists with correct headers
    const headers = getHeadersForSheet(sheetName);
    if (!headers) {
      throw new Error(`No header definition found for sheet: ${sheetName}`);
    }
    sheet = ensureScheduleSheetWithHeaders(sheetName, headers);

    // Clear existing rows below header
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }

    // Write rows if provided
    if (Array.isArray(data) && data.length > 0) {
      const rows = data.map(obj => headers.map(h => (obj[h] !== undefined ? obj[h] : '')));
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    // Invalidate cache
    removeFromCache(`schedule_${sheetName}`);
    return true;
  } catch (error) {
    console.error(`Error writing to schedule sheet ${sheetName}:`, error);
    throw error;
  }
}


/**
 * Create (if needed) and enforce headers + formatting on a sheet.
 * Writes headers when:
 *  - sheet is newly created
 *  - A1 is blank
 *  - first row doesn't match expected headers (length or values)
 */
function ensureScheduleSheetWithHeaders(sheetName, headers) {
  const ss = getScheduleSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  // Create if missing
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  // Read first row safely (always returns at least 1x1)
  const firstRow = sheet.getRange(1, 1, 1, Math.max(headers.length, sheet.getMaxColumns())).getValues()[0];

  // Determine if headers need to be (re)written
  const currentHeaders = firstRow.slice(0, headers.length);
  const a1Empty = String(currentHeaders[0] || '').trim() === '';
  const lengthMismatch = sheet.getLastColumn() < headers.length;
  const valueMismatch = headers.some((h, i) => String(currentHeaders[i] || '').trim() !== String(h));

  if (a1Empty || lengthMismatch || valueMismatch) {
    // Ensure enough columns exist
    if (sheet.getMaxColumns() < headers.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
    }
    // Write headers
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  // Formatting (idempotent)
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('white')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Freeze header
  sheet.setFrozenRows(1);

  // Set sensible widths (then try auto-resize)
  for (let c = 1; c <= headers.length; c++) sheet.setColumnWidth(c, 120);
  try { sheet.autoResizeColumns(1, headers.length); } catch (e) { /* non-fatal */ }

  return sheet;
}

/**
 * Enhanced setupScheduleManagementSheets function with improved formatting
 * Ensure all schedule management sheets exist with proper headers, formatting, and frozen rows
 */
function setupScheduleManagementSheets() {
  try {
    console.log('Setting up Schedule Management sheets with enhanced formatting...');

    // Validate spreadsheet configuration
    const config = validateScheduleSpreadsheetConfig();
    if (!config.canAccess) {
      throw new Error(`Cannot access schedule spreadsheet: ${config.error || 'Unknown error'}`);
    }

    console.log(`Using spreadsheet: ${config.spreadsheetName} (${config.currentSpreadsheet})`);

    const sheetsCreated = [];
    const sheetsUpdated = [];

    // Core schedule sheets with enhanced formatting
    console.log('Creating core schedule management sheets...');

    const coreSheets = [
      { name: SCHEDULE_GENERATION_SHEET, headers: SCHEDULE_GENERATION_HEADERS },
      { name: SHIFT_SLOTS_SHEET, headers: SHIFT_SLOTS_HEADERS },
      { name: SHIFT_SWAPS_SHEET, headers: SHIFT_SWAPS_HEADERS },
      { name: SCHEDULE_TEMPLATES_SHEET, headers: SCHEDULE_TEMPLATES_HEADERS },
      { name: SCHEDULE_NOTIFICATIONS_SHEET, headers: SCHEDULE_NOTIFICATIONS_HEADERS },
      { name: SCHEDULE_ADHERENCE_SHEET, headers: SCHEDULE_ADHERENCE_HEADERS },
      { name: SCHEDULE_CONFLICTS_SHEET, headers: SCHEDULE_CONFLICTS_HEADERS },
      { name: RECURRING_SCHEDULES_SHEET, headers: RECURRING_SCHEDULES_HEADERS }
    ];

    coreSheets.forEach(sheetConfig => {
      try {
        const ss = getScheduleSpreadsheet();
        const existedBefore = ss.getSheetByName(sheetConfig.name) !== null;

        ensureScheduleSheetWithHeaders(sheetConfig.name, sheetConfig.headers);

        if (existedBefore) {
          sheetsUpdated.push(sheetConfig.name);
        } else {
          sheetsCreated.push(sheetConfig.name);
        }
      } catch (error) {
        console.error(`Failed to setup sheet ${sheetConfig.name}:`, error);
        throw error;
      }
    });

    // Attendance and holiday sheets
    console.log('Creating attendance and holiday sheets...');

    const attendanceSheets = [
      { name: ATTENDANCE_STATUS_SHEET, headers: ATTENDANCE_STATUS_HEADERS },
      { name: USER_HOLIDAY_PAY_STATUS_SHEET, headers: USER_HOLIDAY_PAY_STATUS_HEADERS },
      { name: HOLIDAYS_SHEET, headers: HOLIDAYS_HEADERS }
    ];

    attendanceSheets.forEach(sheetConfig => {
      try {
        const ss = getScheduleSpreadsheet();
        const existedBefore = ss.getSheetByName(sheetConfig.name) !== null;

        ensureScheduleSheetWithHeaders(sheetConfig.name, sheetConfig.headers);

        if (existedBefore) {
          sheetsUpdated.push(sheetConfig.name);
        } else {
          sheetsCreated.push(sheetConfig.name);
        }
      } catch (error) {
        console.error(`Failed to setup sheet ${sheetConfig.name}:`, error);
        throw error;
      }
    });

    // Create default shift slots if none exist
    console.log('Checking for default shift slots...');
    createDefaultShiftSlots();

    // Clear any relevant caches
    invalidateScheduleCaches();

    const totalSheets = sheetsCreated.length + sheetsUpdated.length;

    console.log('Schedule Management sheets setup complete');
    console.log(`Created: ${sheetsCreated.length} sheets`);
    console.log(`Updated: ${sheetsUpdated.length} sheets`);
    console.log(`Total processed: ${totalSheets} sheets`);

    if (sheetsCreated.length > 0) {
      console.log('New sheets created:', sheetsCreated.join(', '));
    }

    if (sheetsUpdated.length > 0) {
      console.log('Sheets updated with formatting:', sheetsUpdated.join(', '));
    }

    return {
      success: true,
      spreadsheetId: config.currentSpreadsheet,
      spreadsheetName: config.spreadsheetName,
      usingDedicatedSpreadsheet: config.hasScheduleSpreadsheetId,
      sheetsCreated: sheetsCreated,
      sheetsUpdated: sheetsUpdated,
      totalSheets: totalSheets,
      message: `Successfully processed ${totalSheets} schedule management sheets`
    };

  } catch (error) {
    console.error('Setup failed:', error);
    safeWriteError('setupScheduleManagementSheets', error);
    return {
      success: false,
      error: error.message,
      details: 'Check console for detailed error information'
    };
  }
}

/**
 * Fix formatting on existing sheets (if needed)
 * Call this if sheets were created before the formatting enhancements
 */
function fixExistingScheduleSheetsFormatting() {
  try {
    console.log('Fixing formatting on existing schedule sheets...');

    const allScheduleSheets = [
      SCHEDULE_GENERATION_SHEET,
      SHIFT_SLOTS_SHEET,
      SHIFT_SWAPS_SHEET,
      SCHEDULE_TEMPLATES_SHEET,
      SCHEDULE_NOTIFICATIONS_SHEET,
      SCHEDULE_ADHERENCE_SHEET,
      SCHEDULE_CONFLICTS_SHEET,
      RECURRING_SCHEDULES_SHEET,
      ATTENDANCE_STATUS_SHEET,
      USER_HOLIDAY_PAY_STATUS_SHEET,
      HOLIDAYS_SHEET
    ];

    const ss = getScheduleSpreadsheet();
    const results = [];

    allScheduleSheets.forEach(sheetName => {
      try {
        const sheet = ss.getSheetByName(sheetName);
        if (sheet) {
          const lastColumn = sheet.getLastColumn();
          if (lastColumn > 0) {
            const headerRange = sheet.getRange(1, 1, 1, lastColumn);

            // Apply formatting
            headerRange.setFontWeight('bold');
            headerRange.setBackground('#4285f4');
            headerRange.setFontColor('white');
            headerRange.setHorizontalAlignment('center');
            headerRange.setVerticalAlignment('middle');

            // Freeze header row
            sheet.setFrozenRows(1);

            // Auto-resize columns
            try {
              sheet.autoResizeColumns(1, lastColumn);
            } catch (resizeError) {
              console.warn(`Auto-resize failed for ${sheetName}:`, resizeError);
            }

            results.push({ sheet: sheetName, status: 'success' });
            console.log(`Fixed formatting for: ${sheetName}`);
          } else {
            results.push({ sheet: sheetName, status: 'skipped', reason: 'No data' });
          }
        } else {
          results.push({ sheet: sheetName, status: 'not_found' });
        }
      } catch (error) {
        results.push({ sheet: sheetName, status: 'error', error: error.message });
        console.error(`Error fixing ${sheetName}:`, error);
      }
    });

    const successCount = results.filter(r => r.status === 'success').length;
    console.log(`Formatting fix complete: ${successCount}/${allScheduleSheets.length} sheets processed`);

    return {
      success: true,
      results: results,
      summary: `Fixed ${successCount} sheets`
    };

  } catch (error) {
    console.error('Error fixing sheet formatting:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Complete setup function that ensures everything is properly configured
 */
function completeScheduleSetup() {
  try {
    console.log('Running complete schedule setup...');

    // First, run the main setup
    const setupResult = setupScheduleManagementSheets();

    if (!setupResult.success) {
      return setupResult;
    }

    // If some sheets were updated (existed before), fix their formatting
    if (setupResult.sheetsUpdated && setupResult.sheetsUpdated.length > 0) {
      console.log('Applying formatting fixes to updated sheets...');
      fixExistingScheduleSheetsFormatting();
    }

    // Run validation
    const debugResult = debugScheduleConfiguration();

    return {
      success: true,
      setup: setupResult,
      debug: debugResult,
      message: 'Complete schedule setup finished successfully'
    };

  } catch (error) {
    console.error('Complete setup failed:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Create default shift slots for initial setup
 */
function createDefaultShiftSlots() {
  try {
    const slots = readScheduleSheet(SHIFT_SLOTS_SHEET) || [];
    if (slots.length > 0) {
      console.log('Shift slots already exist, skipping default creation');
      return;
    }

    const defaultSlots = [
      {
        ID: Utilities.getUuid(),
        Name: 'Morning Shift',
        StartTime: '08:00',
        EndTime: '16:00',
        DaysOfWeek: '1,2,3,4,5',
        Department: 'General',
        Location: 'Office',
        MaxCapacity: 10,
        MinCoverage: 5,
        Priority: 2,
        Description: 'Standard morning shift (8 AM - 4 PM)',
        BreakDuration: 15,
        LunchDuration: 60,
        Break1Duration: 15,
        Break2Duration: 15,
        EnableStaggeredBreaks: true,
        BreakGroups: 3,
        StaggerInterval: 15,
        MinCoveragePct: 70,
        EnableOvertime: false,
        MaxDailyOT: 0,
        MaxWeeklyOT: 0,
        OTApproval: 'supervisor',
        OTRate: 1.5,
        OTPolicy: 'MANDATORY',
        AllowSwaps: true,
        WeekendPremium: false,
        HolidayPremium: true,
        AutoAssignment: true,
        RestPeriod: 8,
        NotificationLead: 24,
        HandoverTime: 15,
        OvertimePolicy: 'LIMITED_30',
        IsActive: true,
        CreatedBy: 'System',
        CreatedAt: new Date(),
        UpdatedAt: new Date()
      },
      {
        ID: Utilities.getUuid(),
        Name: 'Evening Shift',
        StartTime: '16:00',
        EndTime: '00:00',
        DaysOfWeek: '1,2,3,4,5',
        Department: 'General',
        Location: 'Office',
        MaxCapacity: 8,
        MinCoverage: 4,
        Priority: 2,
        Description: 'Standard evening shift (4 PM - 12 AM)',
        BreakDuration: 15,
        LunchDuration: 60,
        Break1Duration: 15,
        Break2Duration: 15,
        EnableStaggeredBreaks: true,
        BreakGroups: 3,
        StaggerInterval: 15,
        MinCoveragePct: 70,
        EnableOvertime: false,
        MaxDailyOT: 0,
        MaxWeeklyOT: 0,
        OTApproval: 'supervisor',
        OTRate: 1.5,
        OTPolicy: 'MANDATORY',
        AllowSwaps: true,
        WeekendPremium: false,
        HolidayPremium: true,
        AutoAssignment: true,
        RestPeriod: 8,
        NotificationLead: 24,
        HandoverTime: 15,
        OvertimePolicy: 'LIMITED_30',
        IsActive: true,
        CreatedBy: 'System',
        CreatedAt: new Date(),
        UpdatedAt: new Date()
      },
      {
        ID: Utilities.getUuid(),
        Name: 'Day Shift',
        StartTime: '09:00',
        EndTime: '17:00',
        DaysOfWeek: '1,2,3,4,5',
        Department: 'Customer Service',
        Location: 'Remote',
        MaxCapacity: 15,
        MinCoverage: 6,
        Priority: 2,
        Description: 'Standard day shift (9 AM - 5 PM)',
        BreakDuration: 15,
        LunchDuration: 60,
        Break1Duration: 15,
        Break2Duration: 15,
        EnableStaggeredBreaks: true,
        BreakGroups: 3,
        StaggerInterval: 15,
        MinCoveragePct: 70,
        EnableOvertime: false,
        MaxDailyOT: 0,
        MaxWeeklyOT: 0,
        OTApproval: 'supervisor',
        OTRate: 1.5,
        OTPolicy: 'MANDATORY',
        AllowSwaps: true,
        WeekendPremium: false,
        HolidayPremium: true,
        AutoAssignment: true,
        RestPeriod: 8,
        NotificationLead: 24,
        HandoverTime: 15,
        OvertimePolicy: 'LIMITED_30',
        IsActive: true,
        CreatedBy: 'System',
        CreatedAt: new Date(),
        UpdatedAt: new Date()
      }
    ];

    const sheet = ensureScheduleSheetWithHeaders(SHIFT_SLOTS_SHEET, SHIFT_SLOTS_HEADERS);
    defaultSlots.forEach(slot => {
      const rowData = SHIFT_SLOTS_HEADERS.map(header =>
        Object.prototype.hasOwnProperty.call(slot, header) ? slot[header] : ''
      );
      sheet.appendRow(rowData);
    });

    // Invalidate cache
    const cacheKey = `schedule_${SHIFT_SLOTS_SHEET}`;
    removeFromCache(cacheKey);

    console.log('Created default shift slots');

  } catch (error) {
    console.error('Error creating default shift slots:', error);
    safeWriteError('createDefaultShiftSlots', error);
  }
}

// ────────────────────────────────────────────────────────────────────────────
// CACHE MANAGEMENT UTILITIES FOR SCHEDULE SYSTEM
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get data from cache
 */
function getFromCache(key) {
  try {
    if (typeof CacheService !== 'undefined') {
      const cache = CacheService.getScriptCache();
      const cached = cache.get(key);
      return cached ? JSON.parse(cached) : null;
    }
    return null;
  } catch (error) {
    console.warn('Cache get error:', error);
    return null;
  }
}

/**
 * Set data in cache
 */
function setInCache(key, data, durationSeconds = 300) {
  try {
    if (typeof CacheService !== 'undefined') {
      const cache = CacheService.getScriptCache();
      cache.put(key, JSON.stringify(data), durationSeconds);
    }
  } catch (error) {
    console.warn('Cache set error:', error);
  }
}

/**
 * Remove data from cache
 */
function removeFromCache(key) {
  try {
    if (typeof CacheService !== 'undefined') {
      const cache = CacheService.getScriptCache();
      cache.remove(key);
    }
  } catch (error) {
    console.warn('Cache remove error:', error);
  }
}

/**
 * Invalidate schedule-related caches
 */
function invalidateScheduleCaches() {
  const scheduleSheets = [
    SCHEDULE_GENERATION_SHEET,
    SHIFT_SLOTS_SHEET,
    SHIFT_SWAPS_SHEET,
    SCHEDULE_TEMPLATES_SHEET,
    ATTENDANCE_STATUS_SHEET,
    USER_HOLIDAY_PAY_STATUS_SHEET,
    HOLIDAYS_SHEET
  ];

  scheduleSheets.forEach(sheetName => {
    try {
      const cacheKey = `schedule_${sheetName}`;
      removeFromCache(cacheKey);
    } catch (error) {
      console.warn(`Failed to invalidate cache for ${sheetName}:`, error);
    }
  });
}

// ────────────────────────────────────────────────────────────────────────────
// DATE AND TIME UTILITIES
// ────────────────────────────────────────────────────────────────────────────

/**
 * Format date for display
 */
function formatDisplayDate(dateStr) {
  try {
    if (!dateStr) return '';
    const date = new Date(dateStr);
    return date.toLocaleDateString();
  } catch (error) {
    return dateStr;
  }
}

/**
 * Format time for display
 */
function formatDisplayTime(timeStr) {
  try {
    if (!timeStr) return '';
    // Handle both "HH:MM" and "HH:MM:SS" formats
    const timeParts = timeStr.split(':');
    if (timeParts.length >= 2) {
      const hours = parseInt(timeParts[0]);
      const minutes = timeParts[1];
      const ampm = hours >= 12 ? 'PM' : 'AM';
      const displayHours = hours % 12 || 12;
      return `${displayHours}:${minutes} ${ampm}`;
    }
    return timeStr;
  } catch (error) {
    return timeStr;
  }
}

/**
 * Calculate break start time
 */
function calculateBreakStart(slot) {
  try {
    const startTime = slot.StartTime;
    const [hours, minutes] = startTime.split(':').map(Number);
    const breakHour = hours + 2; // 2 hours after start
    return `${breakHour.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
  } catch (error) {
    return slot.StartTime;
  }
}

/**
 * Calculate break end time
 */
function calculateBreakEnd(slot) {
  try {
    const breakStart = calculateBreakStart(slot);
    const [hours, minutes] = breakStart.split(':').map(Number);
    const breakDuration = slot.BreakDuration || DEFAULT_BREAK_DURATION;
    const endMinutes = minutes + breakDuration;
    const endHour = hours + Math.floor(endMinutes / 60);
    const finalMinutes = endMinutes % 60;
    return `${endHour.toString().padStart(2, '0')}:${finalMinutes.toString().padStart(2, '0')}`;
  } catch (error) {
    return slot.StartTime;
  }
}

/**
 * Calculate lunch start time
 */
function calculateLunchStart(slot) {
  try {
    const startTime = slot.StartTime;
    const [hours, minutes] = startTime.split(':').map(Number);
    const lunchHour = hours + 4; // 4 hours after start
    return `${lunchHour.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
  } catch (error) {
    return slot.StartTime;
  }
}

/**
 * Calculate lunch end time
 */
function calculateLunchEnd(slot) {
  try {
    const lunchStart = calculateLunchStart(slot);
    const [hours, minutes] = lunchStart.split(':').map(Number);
    const lunchDuration = slot.LunchDuration || DEFAULT_LUNCH_DURATION;
    const endMinutes = minutes + lunchDuration;
    const endHour = hours + Math.floor(endMinutes / 60);
    const finalMinutes = endMinutes % 60;
    return `${endHour.toString().padStart(2, '0')}:${finalMinutes.toString().padStart(2, '0')}`;
  } catch (error) {
    return slot.StartTime;
  }
}

/**
 * Check if a date is a holiday
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

/**
 * Basic DST status check
 */
function checkDSTStatus(dateStr) {
  try {
    const date = new Date(dateStr);
    const year = date.getFullYear();

    // Simple US DST check (second Sunday in March to first Sunday in November)
    const dstStart = new Date(year, 2, 8 + (7 - new Date(year, 2, 8).getDay()) % 7);
    const dstEnd = new Date(year, 10, 1 + (7 - new Date(year, 10, 1).getDay()) % 7);

    const isDST = date >= dstStart && date < dstEnd;

    return {
      isDST: isDST,
      isDSTChange: false,
      changeType: null,
      timeAdjustment: 0
    };
  } catch (error) {
    return {
      isDST: false,
      isDSTChange: false,
      changeType: null,
      timeAdjustment: 0
    };
  }
}

/**
 * Get week string from date (ISO week format)
 */
function weekStringFromDate(date) {
  try {
    if (!date || !(date instanceof Date)) {
      date = new Date();
    }

    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    const weekNum = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    return `${d.getUTCFullYear()}-W${weekNum.toString().padStart(2, '0')}`;
  } catch (error) {
    console.error('Error generating week string:', error);
    return new Date().getFullYear() + '-W01';
  }
}

// ────────────────────────────────────────────────────────────────────────────
// DATA VALIDATION UTILITIES
// ────────────────────────────────────────────────────────────────────────────

/**
 * Validate schedule generation parameters
 */
function validateScheduleParameters(startDate, endDate, users) {
  const errors = [];

  if (!startDate) errors.push('Start date is required');
  if (!endDate) errors.push('End date is required');

  if (startDate && endDate) {
    const start = new Date(startDate);
    const end = new Date(endDate);

    if (start >= end) {
      errors.push('End date must be after start date');
    }

    if (start < new Date(Date.now() - 24 * 60 * 60 * 1000)) {
      errors.push('Start date cannot be in the past');
    }
  }

  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

/**
 * Validate shift slot data
 */
function validateShiftSlot(slotData) {
  const errors = [];

  if (!slotData.name) errors.push('Slot name is required');
  if (!slotData.startTime) errors.push('Start time is required');
  if (!slotData.endTime) errors.push('End time is required');

  if (slotData.startTime && slotData.endTime) {
    const start = new Date(`2000-01-01 ${slotData.startTime}`);
    const end = new Date(`2000-01-01 ${slotData.endTime}`);

    if (start >= end) {
      errors.push('End time must be after start time');
    }
  }

  if (slotData.maxCapacity && slotData.maxCapacity < 1) {
    errors.push('Max capacity must be at least 1');
  }

  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

// ────────────────────────────────────────────────────────────────────────────
// USER AND DATA LOOKUP UTILITIES
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get user ID by name
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
 * Get user name by ID
 */
function getUserNameById(userId) {
  try {
    const users = readSheet(USERS_SHEET) || [];
    const user = users.find(u => u.ID === userId);
    return user ? (user.FullName || user.UserName) : userId;
  } catch (error) {
    console.warn('Error getting user name by ID:', error);
    return userId;
  }
}

/**
 * Get all users for attendance (includes ALL users, not just active ones)
 */
function getAttendanceUserList() {
  try {
    const users = readSheet(USERS_SHEET) || [];
    // Return ALL users, not just login-enabled ones
    return users
      .filter(user => user.UserName || user.FullName) // Only filter out users without any name
      .map(user => user.UserName || user.FullName)
      .filter(name => name)
      .sort();
  } catch (error) {
    console.error('Error getting attendance user list:', error);
    safeWriteError('getAttendanceUserList', error);
    return [];
  }
}

/**
 * Check if schedule exists for user on date
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

// ────────────────────────────────────────────────────────────────────────────
// STATUS AND DISPLAY UTILITIES
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get CSS class for status badge
 */
function getStatusBadgeClass(status) {
  switch (status) {
    case 'APPROVED': return 'status-approved';
    case 'REJECTED': return 'status-rejected';
    case 'PENDING': return 'status-pending';
    case 'CANCELLED': return 'status-cancelled';
    default: return 'status-pending';
  }
}

/**
 * Get CSS class for priority level
 */
function getPriorityClass(priority) {
  switch (priority) {
    case 4: return 'bg-danger text-white';
    case 3: return 'bg-warning text-dark';
    case 2: return 'bg-primary text-white';
    default: return 'bg-secondary text-white';
  }
}

/**
 * Get text for priority level
 */
function getPriorityText(priority) {
  switch (priority) {
    case 4: return 'Critical';
    case 3: return 'High';
    case 2: return 'Normal';
    default: return 'Low';
  }
}

/**
 * Get CSS class for attendance status
 */
function getAttendanceStatusClass(status) {
  switch (status.toLowerCase()) {
    case 'present': return 'bg-success text-white';
    case 'absent': return 'bg-danger text-white';
    case 'late': return 'bg-warning text-dark';
    case 'early leave': return 'bg-warning text-dark';
    case 'sick leave': return 'bg-info text-white';
    case 'bereavement': return 'bg-dark text-white';
    case 'vacation': return 'bg-primary text-white';
    case 'personal leave': return 'bg-secondary text-white';
    case 'emergency leave': return 'bg-danger text-white';
    case 'training': return 'bg-success text-white';
    case 'holiday': return 'bg-info text-white';
    default: return 'bg-light text-dark';
  }
}

/**
 * Get short code for attendance status
 */
function getAttendanceStatusCode(status) {
  switch (status.toLowerCase()) {
    case 'present': return 'P';
    case 'absent': return 'A';
    case 'late': return 'L';
    case 'early leave': return 'E';
    case 'sick leave': return 'S';
    case 'bereavement': return 'B';
    case 'vacation': return 'V';
    case 'personal leave': return 'PL';
    case 'emergency leave': return 'EM';
    case 'training': return 'T';
    case 'holiday': return 'H';
    default: return '?';
  }
}

// ────────────────────────────────────────────────────────────────────────────
// HOLIDAY DATA UTILITIES
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get hardcoded holidays for common countries
 */
function getHardcodedHolidays(countryCode, year) {
  const holidaySets = {
    'US': [
      { name: 'New Year\'s Day', date: `${year}-01-01` },
      { name: 'Martin Luther King Jr. Day', date: `${year}-01-15` }, // Third Monday in January (approximate)
      { name: 'Presidents\' Day', date: `${year}-02-19` }, // Third Monday in February (approximate)
      { name: 'Memorial Day', date: `${year}-05-27` }, // Last Monday in May (approximate)
      { name: 'Independence Day', date: `${year}-07-04` },
      { name: 'Labor Day', date: `${year}-09-02` }, // First Monday in September (approximate)
      { name: 'Columbus Day', date: `${year}-10-14` }, // Second Monday in October (approximate)
      { name: 'Veterans Day', date: `${year}-11-11` },
      { name: 'Thanksgiving Day', date: `${year}-11-28` }, // Fourth Thursday in November (approximate)
      { name: 'Christmas Day', date: `${year}-12-25` }
    ],
    'CA': [
      { name: 'New Year\'s Day', date: `${year}-01-01` },
      { name: 'Family Day', date: `${year}-02-19` }, // Third Monday in February (approximate)
      { name: 'Good Friday', date: `${year}-04-07` }, // Approximate
      { name: 'Easter Monday', date: `${year}-04-10` }, // Approximate
      { name: 'Victoria Day', date: `${year}-05-20` }, // Monday before May 25 (approximate)
      { name: 'Canada Day', date: `${year}-07-01` },
      { name: 'Civic Holiday', date: `${year}-08-05` }, // First Monday in August (approximate)
      { name: 'Labour Day', date: `${year}-09-02` }, // First Monday in September (approximate)
      { name: 'Thanksgiving Day', date: `${year}-10-14` }, // Second Monday in October (approximate)
      { name: 'Remembrance Day', date: `${year}-11-11` },
      { name: 'Christmas Day', date: `${year}-12-25` },
      { name: 'Boxing Day', date: `${year}-12-26` }
    ],
    'UK': [
      { name: 'New Year\'s Day', date: `${year}-01-01` },
      { name: 'Good Friday', date: `${year}-04-07` }, // Approximate
      { name: 'Easter Monday', date: `${year}-04-10` }, // Approximate
      { name: 'Early May Bank Holiday', date: `${year}-05-06` }, // First Monday in May (approximate)
      { name: 'Spring Bank Holiday', date: `${year}-05-27` }, // Last Monday in May (approximate)
      { name: 'Summer Bank Holiday', date: `${year}-08-26` }, // Last Monday in August (approximate)
      { name: 'Christmas Day', date: `${year}-12-25` },
      { name: 'Boxing Day', date: `${year}-12-26` }
    ],
    'AU': [
      { name: 'New Year\'s Day', date: `${year}-01-01` },
      { name: 'Australia Day', date: `${year}-01-26` },
      { name: 'Good Friday', date: `${year}-04-07` }, // Approximate
      { name: 'Easter Saturday', date: `${year}-04-08` }, // Approximate
      { name: 'Easter Monday', date: `${year}-04-10` }, // Approximate
      { name: 'ANZAC Day', date: `${year}-04-25` },
      { name: 'Queen\'s Birthday', date: `${year}-06-10` }, // Second Monday in June (approximate)
      { name: 'Labour Day', date: `${year}-10-07` }, // First Monday in October (approximate)
      { name: 'Christmas Day', date: `${year}-12-25` },
      { name: 'Boxing Day', date: `${year}-12-26` }
    ],
    'PH': [
      { name: 'New Year\'s Day', date: `${year}-01-01` },
      { name: 'Maundy Thursday', date: `${year}-04-06` }, // Approximate
      { name: 'Good Friday', date: `${year}-04-07` }, // Approximate
      { name: 'Araw ng Kagitingan', date: `${year}-04-09` },
      { name: 'Labor Day', date: `${year}-05-01` },
      { name: 'Independence Day', date: `${year}-06-12` },
      { name: 'National Heroes Day', date: `${year}-08-26` }, // Last Monday in August (approximate)
      { name: 'All Saints\' Day', date: `${year}-11-01` },
      { name: 'Bonifacio Day', date: `${year}-11-30` },
      { name: 'Christmas Day', date: `${year}-12-25` },
      { name: 'Rizal Day', date: `${year}-12-30` }
    ],
    'IN': [
      { name: 'New Year\'s Day', date: `${year}-01-01` },
      { name: 'Republic Day', date: `${year}-01-26` },
      { name: 'Holi', date: `${year}-03-13` }, // Approximate (varies by lunar calendar)
      { name: 'Good Friday', date: `${year}-04-07` }, // Approximate
      { name: 'Independence Day', date: `${year}-08-15` },
      { name: 'Gandhi Jayanti', date: `${year}-10-02` },
      { name: 'Dussehra', date: `${year}-10-15` }, // Approximate (varies by lunar calendar)
      { name: 'Diwali', date: `${year}-11-04` }, // Approximate (varies by lunar calendar)
      { name: 'Christmas Day', date: `${year}-12-25` }
    ]
  };

  return holidaySets[countryCode] || [];
}

/**
 * Get country names mapping
 */
function getCountryNames() {
  return {
    'US': 'United States',
    'CA': 'Canada',
    'UK': 'United Kingdom',
    'AU': 'Australia',
    'DE': 'Germany',
    'FR': 'France',
    'JP': 'Japan',
    'IN': 'India',
    'PH': 'Philippines',
    'MX': 'Mexico'
  };
}

// ────────────────────────────────────────────────────────────────────────────
// ERROR HANDLING UTILITIES
// ────────────────────────────────────────────────────────────────────────────

/**
 * Safe wrapper for error writing
 */
function safeWriteError(context, error) {
  try {
    if (typeof writeError === 'function') {
      writeError(context, error);
    } else {
      console.error(`${context}:`, error);
    }
  } catch (e) {
    console.error('Error logging failed:', e);
  }
}

/**
 * Create standardized error response
 */
function createErrorResponse(message, details = null) {
  return {
    success: false,
    error: message,
    details: details,
    timestamp: new Date().toISOString()
  };
}

/**
 * Create standardized success response
 */
function createSuccessResponse(message, data = null) {
  return {
    success: true,
    message: message,
    data: data,
    timestamp: new Date().toISOString()
  };
}

// ────────────────────────────────────────────────────────────────────────────
// DEBUGGING AND CONFIGURATION UTILITIES
// ────────────────────────────────────────────────────────────────────────────

/**
 * Debug function to check schedule system configuration
 */
function debugScheduleConfiguration() {
  try {
    console.log('=== Schedule System Configuration Debug ===');

    const config = validateScheduleSpreadsheetConfig();
    console.log('Spreadsheet Configuration:', config);

    // Test spreadsheet access
    try {
      const ss = getScheduleSpreadsheet();
      console.log('✅ Spreadsheet Access: SUCCESS');
      console.log(`   Name: ${ss.getName()}`);
      console.log(`   ID: ${ss.getId()}`);
      console.log(`   URL: ${ss.getUrl()}`);
    } catch (error) {
      console.log('❌ Spreadsheet Access: FAILED');
      console.log(`   Error: ${error.message}`);
    }

    // Check sheet existence
    const requiredSheets = [
      SCHEDULE_GENERATION_SHEET,
      SHIFT_SLOTS_SHEET,
      SHIFT_SWAPS_SHEET,
      SCHEDULE_TEMPLATES_SHEET
    ];

    console.log('\n=== Required Sheets Check ===');
    const ss = getScheduleSpreadsheet();
    requiredSheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      console.log(`${sheetName}: ${sheet ? '✅ EXISTS' : '❌ MISSING'}`);
    });

    // Check data
    console.log('\n=== Data Check ===');
    const slotsCount = (readScheduleSheet(SHIFT_SLOTS_SHEET) || []).length;
    const schedulesCount = (readScheduleSheet(SCHEDULE_GENERATION_SHEET) || []).length;
    console.log(`Shift Slots: ${slotsCount}`);
    console.log(`Generated Schedules: ${schedulesCount}`);

    return {
      success: true,
      configuration: config,
      sheetsExist: requiredSheets.every(name => ss.getSheetByName(name) !== null),
      dataCount: { shifts: slotsCount, schedules: schedulesCount }
    };

  } catch (error) {
    console.error('Debug configuration error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test function to validate all schedule utilities are working
 */
function testScheduleUtilities() {
  try {
    console.log('=== Testing Schedule Utilities ===');

    const tests = [];

    // Test 1: Spreadsheet access
    try {
      const ss = getScheduleSpreadsheet();
      tests.push({ name: 'Spreadsheet Access', status: 'PASS', details: ss.getName() });
    } catch (error) {
      tests.push({ name: 'Spreadsheet Access', status: 'FAIL', details: error.message });
    }

    // Test 2: Sheet creation
    try {
      const sheet = ensureScheduleSheetWithHeaders('TestSheet', ['ID', 'Name', 'Value']);
      tests.push({ name: 'Sheet Creation', status: 'PASS', details: `Created sheet: ${sheet.getName()}` });
    } catch (error) {
      tests.push({ name: 'Sheet Creation', status: 'FAIL', details: error.message });
    }

    // Test 3: Data operations
    try {
      const testData = [{ ID: '1', Name: 'Test', Value: '123' }];
      writeToScheduleSheet('TestSheet', testData);
      const readData = readScheduleSheet('TestSheet');
      const success = readData.length === 1 && readData[0].Name === 'Test';
      tests.push({
        name: 'Data Operations',
        status: success ? 'PASS' : 'FAIL',
        details: success ? 'Read/Write successful' : 'Data mismatch'
      });
    } catch (error) {
      tests.push({ name: 'Data Operations', status: 'FAIL', details: error.message });
    }

    // Test 4: Date utilities
    try {
      const formatted = formatDisplayDate('2025-01-15');
      const timeFormatted = formatDisplayTime('14:30');
      tests.push({
        name: 'Date/Time Utilities',
        status: 'PASS',
        details: `Date: ${formatted}, Time: ${timeFormatted}`
      });
    } catch (error) {
      tests.push({ name: 'Date/Time Utilities', status: 'FAIL', details: error.message });
    }

    console.log('\n=== Test Results ===');
    tests.forEach(test => {
      console.log(`${test.status === 'PASS' ? '✅' : '❌'} ${test.name}: ${test.details}`);
    });

    const allPassed = tests.every(test => test.status === 'PASS');

    return {
      success: allPassed,
      tests: tests,
      summary: `${tests.filter(t => t.status === 'PASS').length}/${tests.length} tests passed`
    };

  } catch (error) {
    console.error('Test utilities error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

console.log('✅ Enhanced ScheduleUtilities.gs loaded successfully');
console.log('🗂️ Features: Dedicated schedule spreadsheet support, enhanced caching, comprehensive utilities');
console.log('⚙️ Configuration: Update SCHEDULE_SPREADSHEET_ID constant to use a dedicated spreadsheet');
console.log('🔧 Debug: Run validateScheduleSpreadsheetConfig() or debugScheduleConfiguration() to check setup');