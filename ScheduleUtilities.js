/**
 * ScheduleUtilities.js
 * Constants, sheet definitions, and utility functions for Schedule Management
 * Updated to include dedicated schedule spreadsheet management
 */

// ────────────────────────────────────────────────────────────────────────────
// SCHEDULE SPREADSHEET CONFIGURATION
// ────────────────────────────────────────────────────────────────────────────

// Schedule Management Spreadsheet ID.
// Update this constant (or the corresponding script property) with your actual schedule spreadsheet ID.
const SCHEDULE_SPREADSHEET_ID = '';

// If no dedicated schedule spreadsheet ID is set, these functions will fall back to the main spreadsheet
const FALLBACK_TO_MAIN_SPREADSHEET = true;

/**
 * Get the Schedule Management spreadsheet
 * Returns the dedicated schedule spreadsheet if ID is configured, otherwise falls back to main/active spreadsheet
 */
function normalizeSpreadsheetId(spreadsheetId) {
  return (typeof spreadsheetId === 'string' ? spreadsheetId.trim() : '') || '';
}

function isPlaceholderSpreadsheetId(spreadsheetId) {
  if (!spreadsheetId) {
    return true;
  }

  const normalized = spreadsheetId.toLowerCase();
  return normalized.includes('todo') || normalized.includes('replace') || normalized.includes('your');
}

function getScheduleSpreadsheetIdCandidates() {
  const candidates = [];

  const normalizedConstantId = normalizeSpreadsheetId(SCHEDULE_SPREADSHEET_ID);
  if (normalizedConstantId && !isPlaceholderSpreadsheetId(normalizedConstantId)) {
    candidates.push(normalizedConstantId);
  }

  try {
    if (typeof PropertiesService !== 'undefined') {
      const properties = PropertiesService.getScriptProperties();
      const propertyIds = [
        normalizeSpreadsheetId(properties.getProperty('SCHEDULE_SPREADSHEET_ID')),
        normalizeSpreadsheetId(properties.getProperty('MAIN_SPREADSHEET_ID'))
      ];
      propertyIds.filter(Boolean).forEach(id => candidates.push(id));
    }
  } catch (error) {
    console.warn('Unable to read schedule spreadsheet ID from script properties:', error && error.message ? error.message : error);
  }

  if (typeof G !== 'undefined' && G) {
    const globalIds = [
      normalizeSpreadsheetId(G.SCHEDULE_SPREADSHEET_ID),
      normalizeSpreadsheetId(G.MAIN_SPREADSHEET_ID)
    ];
    globalIds.filter(Boolean).forEach(id => candidates.push(id));
  }

  return Array.from(new Set(candidates.filter(Boolean)));
}

function getScheduleSpreadsheet() {
  const candidates = getScheduleSpreadsheetIdCandidates();

  for (let index = 0; index < candidates.length; index++) {
    const candidateId = candidates[index];
    try {
      console.log('Using schedule spreadsheet candidate:', candidateId);
      return SpreadsheetApp.openById(candidateId);
    } catch (error) {
      console.warn(`Unable to open schedule spreadsheet ${candidateId}:`, error && error.message ? error.message : error);
    }
  }

  if (FALLBACK_TO_MAIN_SPREADSHEET) {
    try {
      const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      if (activeSpreadsheet) {
        console.log('Using active spreadsheet for schedules (fallback)');
        return activeSpreadsheet;
      }
    } catch (error) {
      console.warn('Active spreadsheet unavailable for schedules:', error && error.message ? error.message : error);
    }
  }

  throw new Error('Schedule spreadsheet not configured or accessible. Set SCHEDULE_SPREADSHEET_ID or configure the script property SCHEDULE_SPREADSHEET_ID / MAIN_SPREADSHEET_ID.');
}

/**
 * Validate schedule spreadsheet configuration
 * Returns information about the current spreadsheet configuration
 */
function validateScheduleSpreadsheetConfig() {
  try {
    const candidates = getScheduleSpreadsheetIdCandidates();
    const config = {
      hasScheduleSpreadsheetId: candidates.length > 0,
      scheduleSpreadsheetId: SCHEDULE_SPREADSHEET_ID,
      fallbackEnabled: FALLBACK_TO_MAIN_SPREADSHEET,
      candidateSpreadsheetIds: candidates,
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
const SHIFT_ASSIGNMENTS_SHEET = "ShiftAssignments";
const SHIFT_SWAPS_SHEET = "ShiftSwaps";
const SCHEDULE_TEMPLATES_SHEET = "ScheduleTemplates";
const SCHEDULE_NOTIFICATIONS_SHEET = "ScheduleNotifications";
const SCHEDULE_ADHERENCE_SHEET = "ScheduleAdherence";
const SCHEDULE_CONFLICTS_SHEET = "ScheduleConflicts";
const RECURRING_SCHEDULES_SHEET = "RecurringSchedules";
const DEMAND_SHEET = "Demand";
const FTE_PLAN_SHEET = "FTEPlan";
const SCHEDULE_FORECAST_METADATA_SHEET = "ScheduleForecastMetadata";
const SCHEDULE_HEALTH_SHEET = "ScheduleHealth";

// Attendance and holiday sheets
const ATTENDANCE_STATUS_SHEET = "AttendanceStatus";
const USER_HOLIDAY_PAY_STATUS_SHEET = "UserHolidayPayStatus";
const HOLIDAYS_SHEET = "Holidays";

const SCHEDULE_SHEET_REGISTRY = Object.freeze({
  SCHEDULES: 'Schedules',
  SHIFTS: 'Shifts',
  SCHEDULE_GENERATION: SCHEDULE_GENERATION_SHEET,
  SHIFT_SLOTS: SHIFT_SLOTS_SHEET,
  SHIFT_ASSIGNMENTS: SHIFT_ASSIGNMENTS_SHEET,
  SHIFT_SWAPS: SHIFT_SWAPS_SHEET,
  SCHEDULE_TEMPLATES: SCHEDULE_TEMPLATES_SHEET,
  SCHEDULE_NOTIFICATIONS: SCHEDULE_NOTIFICATIONS_SHEET,
  SCHEDULE_ADHERENCE: SCHEDULE_ADHERENCE_SHEET,
  SCHEDULE_CONFLICTS: SCHEDULE_CONFLICTS_SHEET,
  RECURRING_SCHEDULES: RECURRING_SCHEDULES_SHEET,
  DEMAND: DEMAND_SHEET,
  FTE_PLAN: FTE_PLAN_SHEET,
  SCHEDULE_FORECAST_METADATA: SCHEDULE_FORECAST_METADATA_SHEET,
  SCHEDULE_HEALTH: SCHEDULE_HEALTH_SHEET,
  ATTENDANCE_STATUS: ATTENDANCE_STATUS_SHEET,
  USER_HOLIDAY_PAY_STATUS: USER_HOLIDAY_PAY_STATUS_SHEET,
  HOLIDAYS: HOLIDAYS_SHEET
});

function getScheduleTimeZone() {
  if (typeof DEFAULT_SCHEDULE_TIME_ZONE !== 'undefined') {
    return DEFAULT_SCHEDULE_TIME_ZONE;
  }

  if (typeof Session !== 'undefined' && typeof Session.getScriptTimeZone === 'function') {
    try {
      return Session.getScriptTimeZone();
    } catch (error) {
      console.warn('Unable to resolve script time zone:', error && error.message ? error.message : error);
    }
  }

  return 'UTC';
}

function normalizeScheduleRowPeriod(record, timeZone) {
  if (!record || typeof record !== 'object') {
    return record;
  }

  const resolveDateString = (value) => {
    if (typeof normalizeDateForSheet === 'function') {
      const normalized = normalizeDateForSheet(value, timeZone);
      if (normalized) {
        return normalized;
      }
    }

    if (value instanceof Date && !isNaN(value.getTime())) {
      return Utilities.formatDate(value, timeZone, 'yyyy-MM-dd');
    }

    if (typeof value === 'number') {
      const numericDate = new Date(value);
      if (!isNaN(numericDate.getTime())) {
        return Utilities.formatDate(numericDate, timeZone, 'yyyy-MM-dd');
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
  };

  const start = resolveDateString(
    record.PeriodStart
      || record.StartDate
      || record.ScheduleStart
      || record.AssignmentStart
      || record.Date
      || record.ScheduleDate
      || record.Day
  );

  const end = resolveDateString(
    record.PeriodEnd
      || record.EndDate
      || record.ScheduleEnd
      || record.AssignmentEnd
      || record.Date
      || record.ScheduleDate
      || record.Day
      || start
  );

  if (!start && !end) {
    return record;
  }

  const normalized = Object.assign({}, record);

  if (start) {
    normalized.PeriodStart = start;
    normalized.Date = start;
  }

  if (end) {
    normalized.PeriodEnd = end;
  }

  return normalized;
}

function getScheduleSheetNames() {
  return Object.assign({}, SCHEDULE_SHEET_REGISTRY);
}

function getScheduleSheetName(key) {
  if (!key) {
    return '';
  }

  const normalizedKey = String(key).trim().toUpperCase();
  const names = SCHEDULE_SHEET_REGISTRY;

  if (Object.prototype.hasOwnProperty.call(names, normalizedKey)) {
    return names[normalizedKey];
  }

  return '';
}

// ────────────────────────────────────────────────────────────────────────────
// SHEET HEADERS DEFINITIONS
// ────────────────────────────────────────────────────────────────────────────

const SCHEDULE_GENERATION_HEADERS = [
  'ID', 'UserID', 'UserName', 'Date', 'PeriodStart', 'PeriodEnd', 'SlotID', 'SlotName', 'StartTime', 'EndTime',
  'OriginalStartTime', 'OriginalEndTime', 'BreakStart', 'BreakEnd', 'LunchStart', 'LunchEnd',
  'IsDST', 'Status', 'GeneratedBy', 'ApprovedBy', 'NotificationSent', 'CreatedAt', 'UpdatedAt',
  'RecurringScheduleID', 'SwapRequestID', 'Priority', 'Notes', 'Location', 'Department',
  'MaxCapacity', 'MinCoverage', 'BreakDuration', 'Break1Duration', 'Break2Duration', 'LunchDuration',
  'EnableStaggeredBreaks', 'BreakGroups', 'StaggerInterval', 'MinCoveragePct',
  'EnableOvertime', 'MaxDailyOT', 'MaxWeeklyOT', 'OTApproval', 'OTRate', 'OTPolicy',
  'AllowSwaps', 'WeekendPremium', 'HolidayPremium', 'AutoAssignment',
  'RestPeriodHours', 'NotificationLeadHours', 'HandoverTimeMinutes', 'NotificationTarget', 'GenerationConfig'
];

const SHIFT_SLOTS_HEADERS = [
  'ID', 'Name', 'StartTime', 'EndTime', 'DaysOfWeek', 'Department', 'Location',
  'MaxCapacity', 'MinCoverage', 'Priority', 'Description',
  'BreakDuration', 'LunchDuration', 'Break1Duration', 'Break2Duration',
  'EnableStaggeredBreaks', 'BreakGroups', 'StaggerInterval', 'MinCoveragePct',
  'EnableOvertime', 'MaxDailyOT', 'MaxWeeklyOT', 'OTApproval', 'OTRate', 'OTPolicy',
  'AllowSwaps', 'WeekendPremium', 'HolidayPremium', 'AutoAssignment',
  'RestPeriod', 'NotificationLead', 'HandoverTime',
  'OvertimePolicy', 'IsActive', 'CreatedBy', 'CreatedAt', 'UpdatedAt',
  'GenerationDefaults'
];

const SHIFT_ASSIGNMENTS_HEADERS = [
  'AssignmentId', 'UserId', 'UserName', 'Campaign', 'SlotId', 'SlotName',
  'StartDate', 'EndDate', 'Status', 'AllowSwap', 'Premiums', 'BreaksConfigJSON',
  'OvertimeMinutes', 'RestPeriodHours', 'NotificationLeadHours', 'HandoverMinutes',
  'Notes', 'CreatedAt', 'CreatedBy', 'UpdatedAt', 'UpdatedBy', 'RollbackGroupId'
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

const DEMAND_HEADERS = [
  'ID', 'Campaign', 'Skill', 'IntervalStart', 'IntervalEnd',
  'ForecastContacts', 'ForecastAHT', 'TargetSL', 'TargetASA',
  'Shrinkage', 'RequiredFTE', 'Notes', 'CreatedAt', 'UpdatedAt'
];

const FTE_PLAN_HEADERS = [
  'ID', 'Campaign', 'Skill', 'IntervalStart', 'IntervalEnd',
  'PlannedFTE', 'ActualFTE', 'Variance', 'CoverageStatus',
  'CreatedAt', 'UpdatedAt', 'CreatedBy', 'Notes'
];

const SCHEDULE_FORECAST_METADATA_HEADERS = [
  'ID', 'Campaign', 'GeneratedAt', 'ForecastWindowStart', 'ForecastWindowEnd',
  'ModelType', 'Parameters', 'Notes', 'Author'
];

const SCHEDULE_HEALTH_HEADERS = [
  'ID', 'Campaign', 'GeneratedAt', 'ServiceLevel', 'ASA', 'AbandonRate',
  'Occupancy', 'Utilization', 'OvertimeHours', 'CostPerHour',
  'FairnessIndex', 'PreferenceSatisfaction', 'ComplianceScore',
  'ScheduleEfficiency', 'Summary'
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

// Schedule configuration defaults and helpers
var SCHEDULE_CONFIG = typeof SCHEDULE_CONFIG !== 'undefined' ? SCHEDULE_CONFIG : Object.freeze({
  PRIMARY_COUNTRY: 'JM',
  SUPPORTED_COUNTRIES: ['JM', 'US', 'DO', 'PH'],
  DEFAULT_SHIFT_CAPACITY: 10,
  DEFAULT_BREAK_MINUTES: 15,
  DEFAULT_LUNCH_MINUTES: 60,
  CACHE_DURATION: 300
});

function getScheduleConfig() {
  return Object.assign({}, SCHEDULE_CONFIG);
}

function getScheduleConfigValue(key, fallback = null) {
  if (SCHEDULE_CONFIG && Object.prototype.hasOwnProperty.call(SCHEDULE_CONFIG, key)) {
    return SCHEDULE_CONFIG[key];
  }
  return fallback;
}

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

function normalizeUserIdValue(value) {
  if (value === null || value === undefined) {
    return '';
  }

  const str = typeof value === 'number' && Number.isFinite(value)
    ? String(Math.trunc(value))
    : String(value);

  return str.trim();
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

const SHIFT_SWAP_STATUS = {
  PENDING: 'PENDING',
  APPROVED: 'APPROVED',
  REJECTED: 'REJECTED',
  CANCELLED: 'CANCELLED'
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
  map[SHIFT_ASSIGNMENTS_SHEET] = SHIFT_ASSIGNMENTS_HEADERS;
  map[SHIFT_SWAPS_SHEET] = SHIFT_SWAPS_HEADERS;
  map[SCHEDULE_TEMPLATES_SHEET] = SCHEDULE_TEMPLATES_HEADERS;
  map[SCHEDULE_NOTIFICATIONS_SHEET] = SCHEDULE_NOTIFICATIONS_HEADERS;
  map[SCHEDULE_ADHERENCE_SHEET] = SCHEDULE_ADHERENCE_HEADERS;
  map[SCHEDULE_CONFLICTS_SHEET] = SCHEDULE_CONFLICTS_HEADERS;
  map[RECURRING_SCHEDULES_SHEET] = RECURRING_SCHEDULES_HEADERS;
  map[DEMAND_SHEET] = DEMAND_HEADERS;
  map[FTE_PLAN_SHEET] = FTE_PLAN_HEADERS;
  map[SCHEDULE_FORECAST_METADATA_SHEET] = SCHEDULE_FORECAST_METADATA_HEADERS;
  map[SCHEDULE_HEALTH_SHEET] = SCHEDULE_HEALTH_HEADERS;
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
      const timeZone = getScheduleTimeZone();
      if (Array.isArray(cached)) {
        const normalizedCache = cached.map(row => normalizeScheduleRowPeriod(row, timeZone));
        const requiresUpdate = cached.some(row => row && !row.PeriodStart && (row.Date || row.StartDate || row.ScheduleDate));
        if (requiresUpdate) {
          setInCache(cacheKey, normalizedCache);
        }
        return normalizedCache;
      }
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
    const timeZone = getScheduleTimeZone();
    const rows = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });

    const normalizedRows = rows.map(row => normalizeScheduleRowPeriod(row, timeZone));

    // Cache the result
    setInCache(cacheKey, normalizedRows);
    return normalizedRows;

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
      const timeZone = getScheduleTimeZone();
      const normalizedData = data.map(obj => normalizeScheduleRowPeriod(obj, timeZone));
      const rows = normalizedData.map(obj => headers.map(h => (obj && obj[h] !== undefined ? obj[h] : '')));
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
      { name: SHIFT_ASSIGNMENTS_SHEET, headers: SHIFT_ASSIGNMENTS_HEADERS },
      { name: SHIFT_SWAPS_SHEET, headers: SHIFT_SWAPS_HEADERS },
      { name: SCHEDULE_TEMPLATES_SHEET, headers: SCHEDULE_TEMPLATES_HEADERS },
      { name: SCHEDULE_NOTIFICATIONS_SHEET, headers: SCHEDULE_NOTIFICATIONS_HEADERS },
      { name: SCHEDULE_ADHERENCE_SHEET, headers: SCHEDULE_ADHERENCE_HEADERS },
      { name: SCHEDULE_CONFLICTS_SHEET, headers: SCHEDULE_CONFLICTS_HEADERS },
      { name: RECURRING_SCHEDULES_SHEET, headers: RECURRING_SCHEDULES_HEADERS }
    ];

    const analyticsSheets = [
      { name: DEMAND_SHEET, headers: DEMAND_HEADERS },
      { name: FTE_PLAN_SHEET, headers: FTE_PLAN_HEADERS },
      { name: SCHEDULE_FORECAST_METADATA_SHEET, headers: SCHEDULE_FORECAST_METADATA_HEADERS },
      { name: SCHEDULE_HEALTH_SHEET, headers: SCHEDULE_HEALTH_HEADERS }
    ];

    const setupSheetDefinition = (sheetConfig) => {
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
    };

    coreSheets.forEach(setupSheetDefinition);
    analyticsSheets.forEach(setupSheetDefinition);

    // Attendance and holiday sheets
    console.log('Creating attendance and holiday sheets...');

    const attendanceSheets = [
      { name: ATTENDANCE_STATUS_SHEET, headers: ATTENDANCE_STATUS_HEADERS },
      { name: USER_HOLIDAY_PAY_STATUS_SHEET, headers: USER_HOLIDAY_PAY_STATUS_HEADERS },
      { name: HOLIDAYS_SHEET, headers: HOLIDAYS_HEADERS }
    ];

    attendanceSheets.forEach(setupSheetDefinition);

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
      SHIFT_ASSIGNMENTS_SHEET,
      SHIFT_SWAPS_SHEET,
      SCHEDULE_TEMPLATES_SHEET,
      SCHEDULE_NOTIFICATIONS_SHEET,
      SCHEDULE_ADHERENCE_SHEET,
      SCHEDULE_CONFLICTS_SHEET,
      RECURRING_SCHEDULES_SHEET,
      DEMAND_SHEET,
      FTE_PLAN_SHEET,
      SCHEDULE_FORECAST_METADATA_SHEET,
      SCHEDULE_HEALTH_SHEET,
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

    const defaultGenerationTemplate = {
      capacity: { max: 10, min: 5 },
      breaks: {
        first: 15,
        lunch: 60,
        second: 15,
        enableStaggered: true,
        groups: 3,
        interval: 15,
        minCoveragePct: 70
      },
      overtime: {
        enabled: false,
        maxDaily: 0,
        maxWeekly: 0,
        approval: 'supervisor',
        rate: 1.5,
        policy: 'MANDATORY'
      },
      advanced: {
        allowSwaps: true,
        weekendPremium: false,
        holidayPremium: true,
        autoAssignment: true,
        restPeriod: 8,
        notificationLead: 24,
        handoverTime: 15
      }
    };

    const applyOverrides = (target, overrides) => {
      if (!overrides || typeof overrides !== 'object') {
        return target;
      }

      Object.keys(overrides).forEach(key => {
        const value = overrides[key];
        if (value && typeof value === 'object' && !Array.isArray(value)) {
          target[key] = target[key] || {};
          applyOverrides(target[key], value);
        } else if (value !== undefined) {
          target[key] = value;
        }
      });

      return target;
    };

    const createSlotRecord = (config, overrides = {}) => {
      const generationDefaults = JSON.parse(JSON.stringify(defaultGenerationTemplate));
      applyOverrides(generationDefaults, overrides);

      return Object.assign({
        ID: Utilities.getUuid(),
        DaysOfWeek: '1,2,3,4,5',
        Priority: 2,
        Description: '',
        IsActive: true,
        CreatedBy: 'System',
        CreatedAt: new Date(),
        UpdatedAt: new Date(),
        GenerationDefaults: JSON.stringify(generationDefaults)
      }, config);
    };

    const defaultSlots = [
      createSlotRecord({
        Name: 'Morning Shift',
        StartTime: '08:00',
        EndTime: '16:00',
        Department: 'General',
        Location: 'Office',
        Description: 'Standard morning shift (8 AM - 4 PM)'
      }, {
        capacity: { max: 10, min: 5 }
      }),
      createSlotRecord({
        Name: 'Evening Shift',
        StartTime: '16:00',
        EndTime: '00:00',
        Department: 'General',
        Location: 'Office',
        Description: 'Standard evening shift (4 PM - 12 AM)'
      }, {
        capacity: { max: 8, min: 4 }
      }),
      createSlotRecord({
        Name: 'Day Shift',
        StartTime: '09:00',
        EndTime: '17:00',
        Department: 'Customer Service',
        Location: 'Remote',
        Description: 'Standard day shift (9 AM - 5 PM)'
      }, {
        capacity: { max: 15, min: 6 }
      })
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
    DEMAND_SHEET,
    FTE_PLAN_SHEET,
    SCHEDULE_FORECAST_METADATA_SHEET,
    SCHEDULE_HEALTH_SHEET,
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
// SHIFT SWAP DATA HELPERS
// ────────────────────────────────────────────────────────────────────────────

function ensureShiftSwapsSheet() {
  return ensureScheduleSheetWithHeaders(SHIFT_SWAPS_SHEET, SHIFT_SWAPS_HEADERS);
}

function listShiftSwapRequests(options = {}) {
  const rows = readScheduleSheet(SHIFT_SWAPS_SHEET) || [];
  if (!rows.length) {
    return [];
  }

  const statusFilter = Array.isArray(options.status)
    ? options.status.map(value => String(value || '').toUpperCase()).filter(Boolean)
    : (options.status ? [String(options.status).toUpperCase()] : null);

  const userFilter = Array.isArray(options.userIds)
    ? options.userIds.map(value => String(value || '').trim()).filter(Boolean)
    : (options.userId ? [String(options.userId).trim()] : null);

  return rows.filter(row => {
    if (statusFilter && statusFilter.length) {
      const status = String(row.Status || row.status || SHIFT_SWAP_STATUS.PENDING).toUpperCase();
      if (!statusFilter.includes(status)) {
        return false;
      }
    }

    if (userFilter && userFilter.length) {
      const requestorId = String(row.RequestorUserID || row.RequestorUserId || row.requestorUserId || '').trim();
      const targetId = String(row.TargetUserID || row.TargetUserId || row.targetUserId || '').trim();
      if (!userFilter.includes(requestorId) && !userFilter.includes(targetId)) {
        return false;
      }
    }

    return true;
  });
}

function createShiftSwapRequestEntry(request = {}) {
  const sheet = ensureShiftSwapsSheet();
  const id = request.ID || request.Id || request.id || (typeof Utilities !== 'undefined' && Utilities.getUuid ? Utilities.getUuid() : `swap_${Date.now()}`);
  const now = new Date();

  const normalized = {
    ID: id,
    RequestorUserID: request.requestorUserId || request.RequestorUserID || '',
    RequestorUserName: request.requestorUserName || request.RequestorUserName || '',
    TargetUserID: request.targetUserId || request.TargetUserID || '',
    TargetUserName: request.targetUserName || request.TargetUserName || '',
    RequestorScheduleID: request.requestorScheduleId || request.RequestorScheduleID || '',
    TargetScheduleID: request.targetScheduleId || request.TargetScheduleID || '',
    SwapDate: request.swapDate || request.SwapDate || '',
    Reason: request.reason || request.Reason || '',
    Status: (request.status || SHIFT_SWAP_STATUS.PENDING),
    ApprovedBy: request.approvedBy || request.ApprovedBy || '',
    RejectedBy: request.rejectedBy || request.RejectedBy || '',
    DecisionNotes: request.decisionNotes || request.DecisionNotes || '',
    CreatedAt: request.createdAt || request.CreatedAt || now,
    UpdatedAt: request.updatedAt || request.UpdatedAt || now
  };

  const rowValues = SHIFT_SWAPS_HEADERS.map(header => Object.prototype.hasOwnProperty.call(normalized, header) ? normalized[header] : '');
  sheet.appendRow(rowValues);

  removeFromCache && removeFromCache(`schedule_${SHIFT_SWAPS_SHEET}`);
  return normalized;
}

function updateShiftSwapRequestEntry(requestId, updates = {}) {
  if (!requestId) {
    return null;
  }

  const sheet = ensureShiftSwapsSheet();
  const idValue = String(requestId).trim();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }

  const headers = SHIFT_SWAPS_HEADERS.slice();
  const idColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let targetRowNumber = null;
  for (let index = 0; index < idColumn.length; index++) {
    if (String(idColumn[index][0] || '').trim() === idValue) {
      targetRowNumber = index + 2;
      break;
    }
  }

  if (!targetRowNumber) {
    return null;
  }

  const currentValues = sheet.getRange(targetRowNumber, 1, 1, headers.length).getValues()[0];
  const currentRecord = {};
  headers.forEach((header, columnIndex) => {
    currentRecord[header] = currentValues[columnIndex];
  });

  const now = new Date();
  const merged = Object.assign({}, currentRecord, updates, { ID: currentRecord.ID || idValue });

  if (merged.status && !merged.Status) {
    merged.Status = merged.status;
  }

  const normalizedStatus = merged.Status || SHIFT_SWAP_STATUS.PENDING;
  merged.Status = String(normalizedStatus).toUpperCase();
  merged.UpdatedAt = merged.UpdatedAt || merged.updatedAt || now;

  const rowValues = headers.map(header => Object.prototype.hasOwnProperty.call(merged, header) ? merged[header] : currentRecord[header]);
  sheet.getRange(targetRowNumber, 1, 1, headers.length).setValues([rowValues]);

  removeFromCache && removeFromCache(`schedule_${SHIFT_SWAPS_SHEET}`);

  const resultRecord = {};
  headers.forEach((header, columnIndex) => {
    resultRecord[header] = rowValues[columnIndex];
  });
  return resultRecord;
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
function checkExistingSchedule(userName, periodStart, periodEnd) {
  try {
    const schedules = readScheduleSheet(SCHEDULE_GENERATION_SHEET) || [];
    const requestedStart = normalizeScheduleDate(periodStart);
    const requestedEnd = normalizeScheduleDate(periodEnd || periodStart);

    if (!requestedStart || !requestedEnd) {
      return null;
    }

    return schedules.find(schedule => {
      if (schedule.UserName !== userName) {
        return false;
      }

      const existingStart = normalizeScheduleDate(schedule.PeriodStart || schedule.Date);
      const existingEnd = normalizeScheduleDate(schedule.PeriodEnd || schedule.Date || schedule.PeriodStart);

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

function normalizeScheduleDate(value) {
  if (!value) {
    return null;
  }

  try {
    const date = new Date(value);
    if (isNaN(date.getTime())) {
      return null;
    }
    return date;
  } catch (error) {
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
// SCHEDULE ANALYTICS AND HEALTH METRICS
// ────────────────────────────────────────────────────────────────────────────

function normalizeScheduleUserId(value) {
  if (value === null || typeof value === 'undefined') {
    return '';
  }

  if (Array.isArray(value)) {
    for (let i = 0; i < value.length; i++) {
      const normalized = normalizeScheduleUserId(value[i]);
      if (normalized) {
        return normalized;
      }
    }
    return '';
  }

  if (typeof value === 'object') {
    const candidates = [value.ID, value.Id, value.id, value.UserID, value.UserId, value.userId, value.value];
    for (let i = 0; i < candidates.length; i++) {
      const normalized = normalizeScheduleUserId(candidates[i]);
      if (normalized) {
        return normalized;
      }
    }
    return '';
  }

  const text = String(value).trim();
  return text;
}

function normalizeScheduleDate(value) {
  if (!value && value !== 0) {
    return null;
  }

  if (value instanceof Date) {
    return new Date(value.getTime());
  }

  if (typeof value === 'number') {
    if (value > 100000000000) {
      return new Date(value);
    }

    const baseDate = new Date(Date.UTC(1899, 11, 30));
    baseDate.setUTCDate(baseDate.getUTCDate() + Math.floor(value));
    return new Date(Date.UTC(baseDate.getUTCFullYear(), baseDate.getUTCMonth(), baseDate.getUTCDate()));
  }

  if (typeof value === 'string') {
    const text = value.trim();
    if (!text) {
      return null;
    }

    if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
      return new Date(`${text}T00:00:00`);
    }

    if (/^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}/.test(text)) {
      const parsed = new Date(text.replace(' ', 'T'));
      if (!isNaN(parsed.getTime())) {
        return parsed;
      }
    }

    const parsedDate = new Date(text);
    if (!isNaN(parsedDate.getTime())) {
      return parsedDate;
    }
  }

  return null;
}

function normalizeScheduleTimeToMinutes(value) {
  if (value === null || typeof value === 'undefined' || value === '') {
    return null;
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

  const text = String(value).trim();
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

  const isoMatch = text.match(/^(\d{4}-\d{2}-\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?/);
  if (isoMatch) {
    const hours = parseInt(isoMatch[2], 10);
    const minutes = parseInt(isoMatch[3], 10);
    return hours * 60 + minutes;
  }

  const hhmmMatch = text.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (hhmmMatch) {
    const hours = parseInt(hhmmMatch[1], 10);
    const minutes = parseInt(hhmmMatch[2], 10);
    return hours * 60 + minutes;
  }

  const numeric = parseFloat(text);
  if (!isNaN(numeric)) {
    return normalizeScheduleTimeToMinutes(numeric);
  }

  return null;
}

function normalizeSchedulePercentage(value, fallback = 0) {
  if (value === null || typeof value === 'undefined' || value === '') {
    return fallback;
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

function buildIntervalKeyFromDate(date, minutesFromMidnight, intervalMinutes = 30) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return null;
  }

  const normalized = new Date(date.getTime());
  normalized.setHours(0, 0, 0, 0);

  const roundedMinutes = Math.max(0, Math.round(minutesFromMidnight / intervalMinutes) * intervalMinutes);
  const hours = Math.floor(roundedMinutes / 60);
  const minutes = roundedMinutes % 60;

  normalized.setHours(hours, minutes, 0, 0);

  const year = normalized.getFullYear();
  const month = String(normalized.getMonth() + 1).padStart(2, '0');
  const day = String(normalized.getDate()).padStart(2, '0');
  const hour = String(hours).padStart(2, '0');
  const minute = String(minutes).padStart(2, '0');

  return `${year}-${month}-${day}T${hour}:${minute}`;
}

function expandScheduleIntervals(scheduleRow, intervalMinutes = 30, options = {}) {
  const dateValue = scheduleRow.Date || scheduleRow.ScheduleDate || scheduleRow.PeriodStart || scheduleRow.StartDate || scheduleRow.Day;
  const date = normalizeScheduleDate(dateValue);
  if (!date) {
    return [];
  }

  const startMinutes = normalizeScheduleTimeToMinutes(scheduleRow.StartTime || scheduleRow.PeriodStart || scheduleRow.ScheduleStart || scheduleRow.ShiftStart);
  const endMinutes = normalizeScheduleTimeToMinutes(scheduleRow.EndTime || scheduleRow.PeriodEnd || scheduleRow.ScheduleEnd || scheduleRow.ShiftEnd);

  if (startMinutes === null || endMinutes === null) {
    return [];
  }

  const breakRanges = [];

  const addBreakRange = (startValue, endValue) => {
    const breakStart = normalizeScheduleTimeToMinutes(startValue);
    const breakEnd = normalizeScheduleTimeToMinutes(endValue);
    if (breakStart === null || breakEnd === null) {
      return;
    }
    const start = Math.min(breakStart, breakEnd);
    const end = Math.max(breakStart, breakEnd);
    if (end > start) {
      breakRanges.push([start, end]);
    }
  };

  addBreakRange(scheduleRow.BreakStart || scheduleRow.Break1Start, scheduleRow.BreakEnd || scheduleRow.Break1End);
  addBreakRange(scheduleRow.Break2Start, scheduleRow.Break2End);
  addBreakRange(scheduleRow.LunchStart, scheduleRow.LunchEnd);
  addBreakRange(scheduleRow.MealStart, scheduleRow.MealEnd);

  const coverage = [];
  const bufferMinutes = Number(options.breakBufferMinutes || 0);

  for (let minute = startMinutes; minute < endMinutes; minute += intervalMinutes) {
    const intervalCenter = minute + intervalMinutes / 2;
    const onBreak = breakRanges.some(range => intervalCenter >= range[0] - bufferMinutes && intervalCenter < range[1] + bufferMinutes);
    if (!onBreak) {
      const key = buildIntervalKeyFromDate(date, minute, intervalMinutes);
      if (key) {
        coverage.push(key);
      }
    }
  }

  return coverage;
}

function aggregateIntervalsFromDemandRow(demandRow, intervalMinutes = 30) {
  const intervals = [];
  const intervalStartValue = demandRow.IntervalStart || demandRow.intervalStart || demandRow.Interval || demandRow.Start;
  const intervalEndValue = demandRow.IntervalEnd || demandRow.intervalEnd || demandRow.End;

  const startDate = normalizeScheduleDate(intervalStartValue || demandRow.Date || demandRow.Day);
  const endDate = normalizeScheduleDate(intervalEndValue);

  if (!startDate) {
    return intervals;
  }

  const startMinutes = normalizeScheduleTimeToMinutes(intervalStartValue);
  const endMinutes = normalizeScheduleTimeToMinutes(intervalEndValue);

  if (startMinutes === null) {
    const derivedMinutes = normalizeScheduleTimeToMinutes(demandRow.StartTime || demandRow.IntervalTime || demandRow.Time);
    if (derivedMinutes !== null) {
      intervals.push(buildIntervalKeyFromDate(startDate, derivedMinutes, intervalMinutes));
      return intervals;
    }
    intervals.push(buildIntervalKeyFromDate(startDate, 0, intervalMinutes));
    return intervals;
  }

  const finalMinutes = endMinutes !== null ? endMinutes : startMinutes + intervalMinutes;
  for (let minute = startMinutes; minute < finalMinutes; minute += intervalMinutes) {
    const key = buildIntervalKeyFromDate(startDate, minute, intervalMinutes);
    if (key) {
      intervals.push(key);
    }
  }

  return intervals;
}

function toNumeric(value, fallback = 0) {
  if (value === null || typeof value === 'undefined' || value === '') {
    return fallback;
  }

  const numberValue = Number(value);
  return isFinite(numberValue) ? numberValue : fallback;
}

function average(values) {
  if (!Array.isArray(values) || values.length === 0) {
    return 0;
  }
  const total = values.reduce((sum, value) => sum + value, 0);
  return total / values.length;
}

function standardDeviation(values, meanValue) {
  if (!Array.isArray(values) || values.length === 0) {
    return 0;
  }
  const mean = typeof meanValue === 'number' ? meanValue : average(values);
  const variance = values.reduce((sum, value) => sum + Math.pow(value - mean, 2), 0) / values.length;
  return Math.sqrt(variance);
}

function calculateCoverageMetrics(scheduleRows, demandRows, options = {}) {
  const intervalMinutes = Number(options.intervalMinutes || 30);
  const peakWeight = Number(options.peakWindowWeight || 1.25);
  const openingHour = Number(options.openingHour || 8);
  const closingHour = Number(options.closingHour || 21);
  const baselineASA = Number(options.baselineASA || 45);
  const targetServiceLevel = Number(options.targetServiceLevel || 0.8);

  const demandByInterval = new Map();
  const coverageByInterval = new Map();
  const intervalDemandMetadata = new Map();

  if (Array.isArray(demandRows)) {
    demandRows.forEach(row => {
      const intervals = aggregateIntervalsFromDemandRow(row, intervalMinutes);
      const requiredFTE = toNumeric(row.RequiredFTE || row.requiredFTE, null);
      const contacts = toNumeric(row.ForecastContacts || row.Contacts || row.Volume, 0);
      const aht = toNumeric(row.ForecastAHT || row.AHT || row.HandleTime, 0);
      const shrinkage = normalizeSchedulePercentage(row.Shrinkage, options.defaultShrinkage || 0.3);

      const computedFTE = requiredFTE !== null && requiredFTE !== undefined && requiredFTE !== ''
        ? Number(requiredFTE)
        : ((contacts * aht) / (intervalMinutes * 60 || 1)) / Math.max(1 - shrinkage, 0.5);

      intervals.forEach((key, index) => {
        if (!key) {
          return;
        }

        const existing = demandByInterval.get(key) || { requiredFTE: 0, contacts: 0 };
        existing.requiredFTE += computedFTE || 0;
        existing.contacts += contacts;
        existing.weight = (existing.weight || 0) + (computedFTE || 1);
        demandByInterval.set(key, existing);
        intervalDemandMetadata.set(key, {
          campaign: row.Campaign || row.campaign || row.CampaignID || row.CampaignId || '',
          skill: row.Skill || row.skill || row.Queue || '',
          targetSL: toNumeric(row.TargetSL || row.TargetServiceLevel, targetServiceLevel * 100),
          targetASA: toNumeric(row.TargetASA || row.ASA || baselineASA),
          shrinkage: shrinkage
        });
      });
    });
  }

  const scheduleRowsArray = Array.isArray(scheduleRows) ? scheduleRows : [];
  scheduleRowsArray.forEach(row => {
    const intervals = expandScheduleIntervals(row, intervalMinutes, options);
    if (!intervals.length) {
      return;
    }

    const userId = normalizeScheduleUserId(row.UserID || row.UserId || row.AgentID || row.AgentId);
    const fteWeight = toNumeric(row.FTE || row.Fte || row.FTEEquivalent || row.WorkdayTarget, 1);
    const skill = row.Skill || row.Queue || row.Channel || '';
    const campaign = row.Campaign || row.CampaignID || row.CampaignId || '';

    intervals.forEach(intervalKey => {
      if (!intervalKey) {
        return;
      }

      if (!coverageByInterval.has(intervalKey)) {
        coverageByInterval.set(intervalKey, {
          staffedFTE: 0,
          agentSet: new Set(),
          skillSet: new Set(),
          campaignSet: new Set()
        });
      }

      const coverage = coverageByInterval.get(intervalKey);
      coverage.staffedFTE += fteWeight || 1;
      if (userId) {
        coverage.agentSet.add(userId);
      }
      if (skill) {
        coverage.skillSet.add(String(skill));
      }
      if (campaign) {
        coverage.campaignSet.add(String(campaign));
      }
    });
  });

  const intervalSummaries = [];
  let totalRequired = 0;
  let totalStaffed = 0;
  let weightedCoverage = 0;
  let weightedRequirement = 0;
  let peakCoverageWeighted = 0;
  let peakWeightTotal = 0;
  let underIntervals = 0;
  let overIntervals = 0;

  const backlogRiskIntervals = [];

  demandByInterval.forEach((demand, key) => {
    const required = Number(demand.requiredFTE || 0);
    const coverage = coverageByInterval.get(key) || { staffedFTE: 0, agentSet: new Set(), skillSet: new Set(), campaignSet: new Set() };
    const staffed = Number(coverage.staffedFTE || 0);

    totalRequired += required;
    totalStaffed += staffed;

    const coverageRatio = required > 0 ? staffed / required : (staffed > 0 ? 1 : 1);
    const normalizedCoverage = Math.min(Math.max(coverageRatio, 0), 2);

    const intervalDate = normalizeScheduleDate(key);
    const intervalHour = intervalDate instanceof Date ? intervalDate.getHours() : Number(String(key).substring(11, 13));
    const isOpening = intervalHour <= openingHour;
    const isClosing = intervalHour >= (closingHour - 1);

    const weight = demand.weight || (required > 0 ? required : 1);
    const adjustedWeight = weight * (isOpening || isClosing ? peakWeight : 1);

    weightedCoverage += Math.min(normalizedCoverage, 1) * adjustedWeight;
    weightedRequirement += adjustedWeight;

    if (isOpening || isClosing) {
      peakCoverageWeighted += Math.min(normalizedCoverage, 1) * adjustedWeight;
      peakWeightTotal += adjustedWeight;
    }

    if (coverageRatio < 0.95) {
      underIntervals += 1;
      backlogRiskIntervals.push({
        intervalKey: key,
        requiredFTE: Number(required.toFixed(2)),
        staffedFTE: Number(staffed.toFixed(2)),
        coverageRatio: Number(coverageRatio.toFixed(3)),
        deficit: Number((required - staffed).toFixed(2)),
        skill: (intervalDemandMetadata.get(key) || {}).skill || ''
      });
    } else if (coverageRatio > 1.15) {
      overIntervals += 1;
    }

    intervalSummaries.push({
      intervalKey: key,
      requiredFTE: Number(required.toFixed(2)),
      staffedFTE: Number(staffed.toFixed(2)),
      coverageRatio: Number(coverageRatio.toFixed(3)),
      variance: Number((staffed - required).toFixed(2)),
      agents: Array.from(coverage.agentSet || []),
      skillsCovered: Array.from(coverage.skillSet || []),
      campaigns: Array.from(coverage.campaignSet || []),
      isOpeningWindow: isOpening,
      isClosingWindow: isClosing,
      metadata: intervalDemandMetadata.get(key) || {}
    });
  });

  coverageByInterval.forEach((coverage, key) => {
    if (demandByInterval.has(key)) {
      return;
    }
    intervalSummaries.push({
      intervalKey: key,
      requiredFTE: 0,
      staffedFTE: Number((coverage.staffedFTE || 0).toFixed(2)),
      coverageRatio: coverage.staffedFTE > 0 ? 2 : 0,
      variance: Number((coverage.staffedFTE || 0).toFixed(2)),
      agents: Array.from(coverage.agentSet || []),
      skillsCovered: Array.from(coverage.skillSet || []),
      campaigns: Array.from(coverage.campaignSet || []),
      metadata: intervalDemandMetadata.get(key) || {}
    });
  });

  intervalSummaries.sort((a, b) => a.intervalKey.localeCompare(b.intervalKey));
  backlogRiskIntervals.sort((a, b) => b.deficit - a.deficit);

  const averageCoverageRatio = totalRequired > 0 ? totalStaffed / totalRequired : 1;
  const serviceLevel = Math.round(Math.min(1, weightedRequirement ? weightedCoverage / weightedRequirement : 1) * 100);
  const asa = Math.round(Math.max(5, baselineASA / Math.max(averageCoverageRatio, 0.4)));
  const abandonRate = Number(Math.max(0, Math.min(40, (1 - Math.min(averageCoverageRatio, 1)) * 25)).toFixed(2));
  const occupancy = Number(Math.max(50, Math.min(98, averageCoverageRatio * 90)).toFixed(2));
  const peakCoverage = peakWeightTotal > 0 ? Number((peakCoverageWeighted / peakWeightTotal * 100).toFixed(2)) : serviceLevel;

  return {
    intervalMinutes,
    intervalSummaries,
    totalRequiredFTE: Number(totalRequired.toFixed(2)),
    totalStaffedFTE: Number(totalStaffed.toFixed(2)),
    coverageScore: Number(Math.min(100, Math.max(0, (weightedRequirement ? (weightedCoverage / weightedRequirement) : 1) * 100)).toFixed(2)),
    averageCoverageRatio: Number(averageCoverageRatio.toFixed(3)),
    serviceLevel,
    asa,
    abandonRate,
    occupancy,
    peakCoverage,
    underStaffedIntervals: underIntervals,
    overStaffedIntervals: overIntervals,
    backlogRiskIntervals: backlogRiskIntervals.slice(0, 10)
  };
}

function evaluatePreferenceAlignment(stats, profile, options = {}) {
  if (!profile) {
    return 100;
  }

  const preferenceText = String(profile.PreferenceNotes || profile.Preferences || profile.preferences || '').toLowerCase();
  if (!preferenceText) {
    return 100;
  }

  let score = 100;
  const totalShifts = Math.max(stats.totalShifts || 0, 1);

  if (/no weekend|avoid weekend|weekend off/.test(preferenceText)) {
    score -= Math.min(40, (stats.weekendShifts / totalShifts) * 120);
  }

  if (/no night|avoid night|prefer day|day shift/.test(preferenceText)) {
    score -= Math.min(35, (stats.nightShifts / totalShifts) * 120);
  }

  if (/prefer morning|prefer opening|early shift/.test(preferenceText)) {
    const ratio = stats.openingShifts / totalShifts;
    score -= Math.max(0, 30 - ratio * 100);
  }

  if (/prefer evening|prefer closing|late shift/.test(preferenceText)) {
    const ratio = stats.closingShifts / totalShifts;
    score -= Math.max(0, 30 - ratio * 100);
  }

  return Math.max(0, Math.min(100, Math.round(score)));
}

function calculateFairnessMetrics(scheduleRows, agentProfiles = [], options = {}) {
  const statsByUser = new Map();
  const profileByUser = new Map();

  (agentProfiles || []).forEach(profile => {
    const userId = normalizeScheduleUserId(profile.ID || profile.UserID || profile.UserId || profile.Email || profile.Username);
    if (userId) {
      profileByUser.set(userId, profile);
    }
  });

  const ensureStats = (userId) => {
    if (!statsByUser.has(userId)) {
      statsByUser.set(userId, {
        userId,
        userName: '',
        totalShifts: 0,
        totalMinutes: 0,
        weekendShifts: 0,
        nightShifts: 0,
        openingShifts: 0,
        closingShifts: 0,
        consecutiveDayBlocks: 0,
        skills: new Set(),
        locations: new Set(),
        preferenceScore: 100,
        fairnessScore: 100
      });
    }
    return statsByUser.get(userId);
  };

  const nightStart = Number(options.nightThresholdStart || 20 * 60);
  const nightEnd = Number(options.nightThresholdEnd || 6 * 60);
  const openingThreshold = Number(options.openingThreshold || 8 * 60);
  const closingThreshold = Number(options.closingThreshold || 21 * 60);

  const scheduleRowsArray = Array.isArray(scheduleRows) ? scheduleRows : [];
  scheduleRowsArray.forEach(row => {
    const userId = normalizeScheduleUserId(row.UserID || row.UserId || row.AgentID || row.AgentId);
    if (!userId) {
      return;
    }

    const stats = ensureStats(userId);
    const date = normalizeScheduleDate(row.Date || row.ScheduleDate || row.PeriodStart || row.StartDate || row.Day);
    if (!date) {
      return;
    }

    const startMinutes = normalizeScheduleTimeToMinutes(row.StartTime || row.PeriodStart || row.ScheduleStart || row.ShiftStart);
    const endMinutes = normalizeScheduleTimeToMinutes(row.EndTime || row.PeriodEnd || row.ScheduleEnd || row.ShiftEnd);

    if (row.UserName || row.FullName || row.Name) {
      stats.userName = stats.userName || row.UserName || row.FullName || row.Name;
    }

    stats.totalShifts += 1;
    if (startMinutes !== null && endMinutes !== null) {
      let duration = endMinutes - startMinutes;
      if (duration < 0) {
        duration += 24 * 60;
      }
      stats.totalMinutes += Math.max(0, duration);
    }

    const dayOfWeek = date.getDay();
    if (WEEKEND.includes(dayOfWeek)) {
      stats.weekendShifts += 1;
    }

    if ((startMinutes !== null && startMinutes >= nightStart) || (endMinutes !== null && endMinutes <= nightEnd)) {
      stats.nightShifts += 1;
    }

    if (startMinutes !== null && startMinutes <= openingThreshold) {
      stats.openingShifts += 1;
    }

    if (endMinutes !== null && endMinutes >= closingThreshold) {
      stats.closingShifts += 1;
    }

    if (row.Skill) {
      stats.skills.add(String(row.Skill));
    }
    if (row.Location) {
      stats.locations.add(String(row.Location));
    }
  });

  const userStats = Array.from(statsByUser.values());
  const weekendCounts = userStats.map(stats => stats.weekendShifts);
  const nightCounts = userStats.map(stats => stats.nightShifts);
  const closingCounts = userStats.map(stats => stats.closingShifts);

  const weekendMean = average(weekendCounts);
  const nightMean = average(nightCounts);
  const closingMean = average(closingCounts);

  const weekendStd = standardDeviation(weekendCounts, weekendMean);
  const nightStd = standardDeviation(nightCounts, nightMean);
  const closingStd = standardDeviation(closingCounts, closingMean);

  const normalizeVariance = (stdDev, mean) => {
    if (!isFinite(stdDev) || mean <= 0) {
      return 0;
    }
    return Math.min(1.5, stdDev / Math.max(mean, 0.5));
  };

  const fairnessVariance = (normalizeVariance(weekendStd, weekendMean) + normalizeVariance(nightStd, nightMean) + normalizeVariance(closingStd, closingMean)) / 3;
  const fairnessIndex = Math.max(0, Math.min(100, Math.round(100 - fairnessVariance * 50)));

  let preferenceTotal = 0;

  userStats.forEach(stats => {
    const profile = profileByUser.get(stats.userId);
    stats.preferenceScore = evaluatePreferenceAlignment(stats, profile, options);
    preferenceTotal += stats.preferenceScore;

    const fairnessPenalty = (
      Math.abs(stats.weekendShifts - weekendMean) / Math.max(weekendMean || 1, 1) +
      Math.abs(stats.nightShifts - nightMean) / Math.max(nightMean || 1, 1) +
      Math.abs(stats.closingShifts - closingMean) / Math.max(closingMean || 1, 1)
    ) / 3;

    stats.fairnessScore = Math.max(0, Math.min(100, Math.round(100 - fairnessPenalty * 40)));
    stats.skills = Array.from(stats.skills);
    stats.locations = Array.from(stats.locations);
  });

  const preferenceSatisfaction = userStats.length ? Math.round(preferenceTotal / userStats.length) : 100;
  const rotationHealth = Math.max(0, Math.min(100, Math.round(100 - normalizeVariance(weekendStd, weekendMean) * 60)));

  return {
    fairnessIndex,
    preferenceSatisfaction,
    rotationHealth,
    weekendBalance: {
      average: Number(weekendMean.toFixed(2)),
      stdDev: Number(weekendStd.toFixed(2)),
      counts: weekendCounts
    },
    nightBalance: {
      average: Number(nightMean.toFixed(2)),
      stdDev: Number(nightStd.toFixed(2)),
      counts: nightCounts
    },
    closingBalance: {
      average: Number(closingMean.toFixed(2)),
      stdDev: Number(closingStd.toFixed(2)),
      counts: closingCounts
    },
    agentSummaries: userStats
  };
}

function calculateComplianceMetrics(scheduleRows, agentProfiles = [], options = {}) {
  const maxHoursPerDay = Number(options.maxHoursPerDay || 12);
  const minRestHours = Number(options.minRestHours || 10);
  const allowedBreakOverlap = Number(options.allowedBreakOverlap || 3);
  const maxConsecutiveDays = Number(options.maxConsecutiveDays || 6);

  const violationDetails = [];
  const shiftsByUser = new Map();
  const breakOverlapCounter = new Map();

  const addViolation = (type, details) => {
    violationDetails.push(Object.assign({ type, timestamp: new Date().toISOString() }, details || {}));
  };

  const scheduleRowsArray = Array.isArray(scheduleRows) ? scheduleRows : [];
  scheduleRowsArray.forEach(row => {
    const userId = normalizeScheduleUserId(row.UserID || row.UserId || row.AgentID || row.AgentId);
    if (!userId) {
      return;
    }

    const date = normalizeScheduleDate(row.Date || row.ScheduleDate || row.PeriodStart || row.StartDate || row.Day);
    if (!date) {
      return;
    }

    const startMinutes = normalizeScheduleTimeToMinutes(row.StartTime || row.PeriodStart || row.ScheduleStart || row.ShiftStart);
    const endMinutes = normalizeScheduleTimeToMinutes(row.EndTime || row.PeriodEnd || row.ScheduleEnd || row.ShiftEnd);
    if (startMinutes === null || endMinutes === null) {
      return;
    }

    const shiftStart = new Date(date.getTime());
    shiftStart.setHours(0, 0, 0, 0);
    shiftStart.setMinutes(shiftStart.getMinutes() + startMinutes);

    const shiftEnd = new Date(date.getTime());
    shiftEnd.setHours(0, 0, 0, 0);
    shiftEnd.setMinutes(shiftEnd.getMinutes() + endMinutes);
    if (shiftEnd <= shiftStart) {
      shiftEnd.setDate(shiftEnd.getDate() + 1);
    }

    const durationHours = (shiftEnd.getTime() - shiftStart.getTime()) / (1000 * 60 * 60);
    if (durationHours > maxHoursPerDay) {
      addViolation('SHIFT_DURATION', {
        userId,
        userName: row.UserName || row.FullName || row.Name,
        scheduledHours: Number(durationHours.toFixed(2)),
        maxHoursPerDay
      });
    }

    const breaks = [];
    const addBreak = (startValue, endValue, type) => {
      const breakStartMinutes = normalizeScheduleTimeToMinutes(startValue);
      const breakEndMinutes = normalizeScheduleTimeToMinutes(endValue);
      if (breakStartMinutes === null || breakEndMinutes === null) {
        return;
      }
      breaks.push({
        startMinutes: breakStartMinutes,
        endMinutes: breakEndMinutes,
        type: type || 'BREAK'
      });
      const key = `${date.toISOString().slice(0, 10)}|${breakStartMinutes}`;
      breakOverlapCounter.set(key, (breakOverlapCounter.get(key) || 0) + 1);
    };

    addBreak(row.BreakStart || row.Break1Start, row.BreakEnd || row.Break1End, 'BREAK');
    addBreak(row.Break2Start, row.Break2End, 'BREAK');
    addBreak(row.LunchStart, row.LunchEnd, 'LUNCH');

    if (breaks.length === 0) {
      addViolation('MISSING_BREAK', {
        userId,
        userName: row.UserName || row.FullName || row.Name,
        date: date.toISOString().slice(0, 10)
      });
    }

    if (!shiftsByUser.has(userId)) {
      shiftsByUser.set(userId, []);
    }

    shiftsByUser.get(userId).push({
      start: shiftStart,
      end: shiftEnd,
      date,
      breaks,
      userName: row.UserName || row.FullName || row.Name
    });
  });

  let restViolations = 0;
  let consecutiveViolations = 0;

  shiftsByUser.forEach((shifts, userId) => {
    shifts.sort((a, b) => a.start.getTime() - b.start.getTime());

    for (let i = 1; i < shifts.length; i++) {
      const restHours = (shifts[i].start.getTime() - shifts[i - 1].end.getTime()) / (1000 * 60 * 60);
      if (restHours < minRestHours) {
        restViolations += 1;
        addViolation('REST_PERIOD', {
          userId,
          userName: shifts[i].userName,
          restHours: Number(restHours.toFixed(2)),
          requiredRestHours: minRestHours,
          previousShiftEnd: shifts[i - 1].end.toISOString(),
          nextShiftStart: shifts[i].start.toISOString()
        });
      }
    }

    let currentStreak = 1;
    for (let i = 1; i < shifts.length; i++) {
      const previousDay = shifts[i - 1].date;
      const currentDay = shifts[i].date;
      const dayDifference = Math.floor((currentDay - previousDay) / (24 * 60 * 60 * 1000));
      if (dayDifference === 1) {
        currentStreak += 1;
      } else if (dayDifference > 1) {
        currentStreak = 1;
      }

      if (currentStreak > maxConsecutiveDays) {
        consecutiveViolations += 1;
        addViolation('CONSECUTIVE_DAYS', {
          userId,
          userName: shifts[i].userName,
          consecutiveDays: currentStreak,
          maxConsecutiveDays
        });
        currentStreak = 1;
      }
    }
  });

  let maxBreakOverlap = 0;
  breakOverlapCounter.forEach(count => {
    if (count > maxBreakOverlap) {
      maxBreakOverlap = count;
    }
  });

  const breakOverlapViolations = Math.max(0, maxBreakOverlap - allowedBreakOverlap);

  const totalViolations = violationDetails.length;
  const penalty = totalViolations * 3 + restViolations * 2 + breakOverlapViolations * 5 + consecutiveViolations * 2;
  const complianceScore = Math.max(0, Math.min(100, Math.round(100 - penalty)));

  const recommendations = [];
  if (restViolations > 0) {
    recommendations.push(`Increase rest periods to at least ${minRestHours} hours between consecutive shifts.`);
  }
  if (breakOverlapViolations > 0) {
    recommendations.push(`Stagger breaks to ensure no more than ${allowedBreakOverlap} agents are off simultaneously.`);
  }
  if (violationDetails.some(v => v.type === 'MISSING_BREAK')) {
    recommendations.push('Insert the required paid breaks and unpaid lunch for every shift.');
  }
  if (violationDetails.some(v => v.type === 'SHIFT_DURATION')) {
    recommendations.push(`Review long shifts exceeding ${maxHoursPerDay} hours and split or reassign as needed.`);
  }

  return {
    complianceScore,
    totalViolations,
    restViolations,
    consecutiveViolations,
    breakOverlap: {
      maxSimultaneous: maxBreakOverlap,
      allowed: allowedBreakOverlap
    },
    violationDetails,
    recommendations
  };
}

function evaluateSchedulePerformance(scheduleRows, demandRows, agentProfiles = [], options = {}) {
  const coverage = calculateCoverageMetrics(scheduleRows, demandRows, options);
  const fairness = calculateFairnessMetrics(scheduleRows, agentProfiles, options);
  const compliance = calculateComplianceMetrics(scheduleRows, agentProfiles, options);

  const weights = Object.assign({ coverage: 0.45, fairness: 0.25, compliance: 0.2, preference: 0.1 }, options.healthWeights || {});

  const coverageComponent = (coverage.coverageScore || 0) / 100;
  const fairnessComponent = (fairness.fairnessIndex || 0) / 100;
  const complianceComponent = (compliance.complianceScore || 0) / 100;
  const preferenceComponent = (fairness.preferenceSatisfaction || 0) / 100;

  const healthScore = Math.round(
    (coverageComponent * weights.coverage) +
    (fairnessComponent * weights.fairness) +
    (complianceComponent * weights.compliance) +
    (preferenceComponent * weights.preference)
  * 100
  );

  return {
    generatedAt: new Date().toISOString(),
    healthScore,
    coverage,
    fairness,
    compliance,
    summary: {
      serviceLevel: coverage.serviceLevel,
      asa: coverage.asa,
      abandonRate: coverage.abandonRate,
      occupancy: coverage.occupancy,
      fairnessIndex: fairness.fairnessIndex,
      preferenceSatisfaction: fairness.preferenceSatisfaction,
      complianceScore: compliance.complianceScore,
      scheduleEfficiency: coverage.coverageScore
    }
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

console.log('✅ Enhanced ScheduleUtilities.js loaded successfully');
console.log('🗂️ Features: Dedicated schedule spreadsheet support, enhanced caching, comprehensive utilities');
console.log('⚙️ Configuration: Update SCHEDULE_SPREADSHEET_ID constant to use a dedicated spreadsheet');
console.log('🔧 Debug: Run validateScheduleSpreadsheetConfig() or debugScheduleConfiguration() to check setup');
