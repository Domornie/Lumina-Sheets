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

const SCHEDULE_SYSTEM_ACTOR_NAME = (SCHEDULE_SETTINGS && (
  SCHEDULE_SETTINGS.SYSTEM_ACTOR_NAME
    || SCHEDULE_SETTINGS.SYSTEM_NAME
    || SCHEDULE_SETTINGS.SYSTEM_USER_NAME
)) || 'Lumina Schedule System';

const SCHEDULE_SYSTEM_ACTOR_EMAIL = (SCHEDULE_SETTINGS && (
  SCHEDULE_SETTINGS.SYSTEM_ACTOR_EMAIL
    || SCHEDULE_SETTINGS.SYSTEM_EMAIL
    || SCHEDULE_SETTINGS.SYSTEM_USER_EMAIL
)) || 'system@lumina-hq.com';

const SCHEDULE_SYSTEM_ACTOR_ALIASES = [
  'system',
  'systemuser',
  'system-user',
  'system_user',
  'systemaccount',
  'system-account',
  'system scheduler',
  'schedule-system',
  'schedule_system',
  (SCHEDULE_SYSTEM_ACTOR_EMAIL || '').toLowerCase()
].filter(Boolean);

let scheduleActorDirectory = null;
let scheduleActorDirectoryBuiltAt = 0;
const SCHEDULE_ACTOR_DIRECTORY_TTL_MS = 5 * 60 * 1000;

function normalizeActorId(value) {
  if (typeof normalizeUserIdValue === 'function') {
    return normalizeUserIdValue(value);
  }
  if (value === null || typeof value === 'undefined') {
    return '';
  }
  return String(value).trim();
}

function formatScheduleActorLabel(name, email) {
  const normalizedName = name ? String(name).trim() : '';
  const normalizedEmail = email ? String(email).trim().toLowerCase() : '';

  if (normalizedName && normalizedEmail) {
    return `${normalizedName} <${normalizedEmail}>`;
  }
  if (normalizedEmail) {
    return normalizedEmail;
  }
  if (normalizedName) {
    return normalizedName;
  }
  const fallbackName = SCHEDULE_SYSTEM_ACTOR_NAME || 'Lumina Schedule System';
  const fallbackEmail = SCHEDULE_SYSTEM_ACTOR_EMAIL || 'system@lumina-hq.com';
  return `${fallbackName} <${(fallbackEmail || '').toLowerCase()}>`;
}

function buildScheduleActorDirectory(forceReload = false) {
  const now = Date.now ? Date.now() : new Date().getTime();
  if (!forceReload && scheduleActorDirectory && (now - scheduleActorDirectoryBuiltAt) < SCHEDULE_ACTOR_DIRECTORY_TTL_MS) {
    return scheduleActorDirectory;
  }

  const directory = {
    byId: new Map(),
    byEmail: new Map()
  };

  try {
    const users = (typeof readSheet === 'function') ? (readSheet(USERS_SHEET) || []) : [];
    users.forEach(user => {
      if (!user || typeof user !== 'object') {
        return;
      }
      const id = normalizeActorId(user.ID || user.Id || user.UserID || user.UserId || user.id || user.userId);
      const email = (user.Email || user.email || '').toString().trim().toLowerCase();
      const name = (user.FullName || user.fullName || user.UserName || user.Username || user.Name || '').toString().trim();
      const record = {
        id,
        email,
        name
      };

      if (id) {
        directory.byId.set(id, record);
      }
      if (email) {
        directory.byEmail.set(email, record);
      }
    });
  } catch (error) {
    console.warn('Unable to build schedule actor directory:', error);
  }

  scheduleActorDirectory = directory;
  scheduleActorDirectoryBuiltAt = now;
  return scheduleActorDirectory;
}

function resolveScheduleActor(candidate) {
  const systemEmail = (SCHEDULE_SYSTEM_ACTOR_EMAIL || '').toLowerCase();
  const systemName = SCHEDULE_SYSTEM_ACTOR_NAME || 'Lumina Schedule System';
  const systemLabel = formatScheduleActorLabel(systemName, systemEmail);
  const systemInfo = {
    id: 'system',
    email: systemEmail,
    name: systemName,
    label: systemLabel,
    lookupKey: 'system'
  };

  const directory = buildScheduleActorDirectory(false);

  if (candidate && typeof candidate === 'object') {
    const candidateId = normalizeActorId(
      candidate.ID
        || candidate.Id
        || candidate.UserID
        || candidate.UserId
        || candidate.id
        || candidate.userId
    );
    const candidateEmail = (candidate.Email || candidate.email || '').toString().trim().toLowerCase();
    const candidateName = (candidate.FullName || candidate.fullName || candidate.Name || candidate.name || candidate.UserName || candidate.Username || candidate.userName || '').toString().trim();

    const directoryRecord = candidateEmail && directory.byEmail ? directory.byEmail.get(candidateEmail) : null;
    const idRecord = (!directoryRecord && candidateId && directory.byId) ? directory.byId.get(candidateId) : directoryRecord;

    const resolvedEmail = candidateEmail || (directoryRecord && directoryRecord.email) || (idRecord && idRecord.email) || '';
    const resolvedId = candidateId || (idRecord && idRecord.id) || (directoryRecord && directoryRecord.id) || '';
    const resolvedName = candidateName || (idRecord && idRecord.name) || (directoryRecord && directoryRecord.name) || '';

    return {
      id: resolvedId || '',
      email: resolvedEmail || '',
      name: resolvedName || '',
      label: formatScheduleActorLabel(resolvedName || '', resolvedEmail || ''),
      lookupKey: resolvedId || resolvedEmail || systemInfo.lookupKey
    };
  }

  const candidateStr = candidate === null || typeof candidate === 'undefined'
    ? ''
    : String(candidate).trim();

  if (!candidateStr) {
    return systemInfo;
  }

  const lowerCandidate = candidateStr.toLowerCase();
  if (SCHEDULE_SYSTEM_ACTOR_ALIASES.includes(lowerCandidate)) {
    return systemInfo;
  }

  let resolvedId = '';
  let resolvedEmail = '';
  let resolvedName = '';

  if (candidateStr.includes('<') && candidateStr.includes('>')) {
    const emailMatch = candidateStr.match(/<([^>]+)>/);
    if (emailMatch && emailMatch[1]) {
      resolvedEmail = emailMatch[1].trim().toLowerCase();
    }
    const namePart = candidateStr.replace(/<[^>]+>/g, '').trim();
    if (namePart) {
      resolvedName = namePart;
    }
  }

  if (!resolvedEmail && candidateStr.includes('@')) {
    resolvedEmail = candidateStr.toLowerCase();
  }

  resolvedId = normalizeActorId(candidateStr);

  if (!resolvedName && !candidateStr.includes('@') && !candidateStr.includes('<')) {
    resolvedName = candidateStr;
  }

  if (resolvedEmail && directory.byEmail && directory.byEmail.has(resolvedEmail)) {
    const record = directory.byEmail.get(resolvedEmail);
    resolvedId = resolvedId || (record && record.id) || '';
    resolvedName = resolvedName || (record && record.name) || '';
  }

  if (resolvedId && directory.byId && directory.byId.has(resolvedId)) {
    const record = directory.byId.get(resolvedId);
    resolvedEmail = resolvedEmail || (record && record.email) || '';
    resolvedName = resolvedName || (record && record.name) || '';
  }

  if (!resolvedName && !resolvedEmail) {
    return systemInfo;
  }

  return {
    id: resolvedId || '',
    email: resolvedEmail || '',
    name: resolvedName || '',
    label: formatScheduleActorLabel(resolvedName || '', resolvedEmail || ''),
    lookupKey: resolvedId || resolvedEmail || systemInfo.lookupKey
  };
}

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
  const actorCandidate = actorId || (typeof getCurrentUser === 'function' ? getCurrentUser() : null);
  const actorInfo = resolveScheduleActor(actorCandidate);
  const actorLabel = actorInfo.label;

  const rows = assignments.map(assignment => {
    const normalized = Object.assign({}, assignment);
    normalized.AssignmentId = normalized.AssignmentId || Utilities.getUuid();
    normalized.CreatedAt = normalized.CreatedAt || now;
    normalized.CreatedBy = normalized.CreatedBy || actorLabel;
    normalized.UpdatedAt = now;
    normalized.UpdatedBy = actorLabel;
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

  invalidateScheduleCaches();

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

function safeParseJson(value, fallback = null) {
  if (!value && value !== 0) {
    return fallback;
  }

  if (typeof value === 'object') {
    return value;
  }

  if (typeof value !== 'string') {
    return fallback;
  }

  try {
    return JSON.parse(value);
  } catch (error) {
    console.warn('Unable to parse JSON value:', error);
    return fallback;
  }
}

function coerceScheduleNumber(value, fallback = 0, { min = null, max = null } = {}) {
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
}

function coerceScheduleBoolean(value, fallback = false) {
  if (value === null || typeof value === 'undefined' || value === '') {
    return Boolean(fallback);
  }
  return scheduleFlagToBool(value);
}

function serializeSlotConfiguration(config = {}) {
  const payload = {
    capacity: Object.assign({}, config.capacity || {}),
    breaks: Object.assign({}, config.breaks || {}),
    overtime: Object.assign({}, config.overtime || {}),
    advanced: Object.assign({}, config.advanced || {})
  };
  try {
    return JSON.stringify(payload);
  } catch (error) {
    console.warn('Unable to serialize slot configuration:', error);
    return JSON.stringify({});
  }
}

function normalizeSlotConfiguration(slot = {}, fallbackOptions = {}) {
  const fallbackNormalized = normalizeGenerationOptions(fallbackOptions || {});
  const config = {
    capacity: Object.assign({}, fallbackNormalized.capacity || {}),
    breaks: Object.assign({}, fallbackNormalized.breaks || {}),
    overtime: Object.assign({}, fallbackNormalized.overtime || {}),
    advanced: Object.assign({}, fallbackNormalized.advanced || {})
  };

  const configurationCandidates = [
    slot.SlotConfiguration,
    slot.ConfigurationJSON,
    slot.ConfigurationJson,
    slot.ConfigJSON,
    slot.Configuration,
    slot.SettingsJSON,
    slot.GenerationConfig,
    slot.configuration,
    slot.config
  ];

  for (let index = 0; index < configurationCandidates.length; index++) {
    const candidate = configurationCandidates[index];
    if (!candidate && candidate !== 0) {
      continue;
    }

    const parsed = safeParseJson(candidate, null);
    if (parsed) {
      const normalized = normalizeGenerationOptions(parsed);
      config.capacity = Object.assign(config.capacity, normalized.capacity || {});
      config.breaks = Object.assign(config.breaks, normalized.breaks || {});
      config.overtime = Object.assign(config.overtime, normalized.overtime || {});
      config.advanced = Object.assign(config.advanced, normalized.advanced || {});
      break;
    }
  }

  const resolveNumberCandidate = (values, options = {}) => {
    for (let i = 0; i < values.length; i++) {
      const candidate = values[i];
      if (candidate === null || typeof candidate === 'undefined' || candidate === '') {
        continue;
      }
      const resolved = Number(candidate);
      if (!Number.isFinite(resolved)) {
        continue;
      }
      let output = resolved;
      if (typeof options.min === 'number' && output < options.min) {
        output = options.min;
      }
      if (typeof options.max === 'number' && output > options.max) {
        output = options.max;
      }
      return output;
    }
    return null;
  };

  const resolveStringCandidate = (values, fallback = '') => {
    for (let i = 0; i < values.length; i++) {
      const candidate = values[i];
      if (candidate === null || typeof candidate === 'undefined') {
        continue;
      }
      const value = String(candidate).trim();
      if (value) {
        return value;
      }
    }
    return fallback;
  };

  const resolveBooleanCandidate = (values, fallback) => {
    for (let i = 0; i < values.length; i++) {
      const candidate = values[i];
      if (candidate === null || typeof candidate === 'undefined' || candidate === '') {
        continue;
      }
      return scheduleFlagToBool(candidate);
    }
    return Boolean(fallback);
  };

  const capacityMax = resolveNumberCandidate([
    slot.CapacityMax,
    slot.MaxCapacity,
    slot.capacityMax,
    slot.maxCapacity
  ], { min: 0 });
  if (capacityMax !== null) {
    config.capacity.max = capacityMax;
  }

  const capacityMin = resolveNumberCandidate([
    slot.CapacityMin,
    slot.MinCoverage,
    slot.capacityMin,
    slot.minCoverage
  ], { min: 0 });
  if (capacityMin !== null) {
    config.capacity.min = capacityMin;
  }

  const break1 = resolveNumberCandidate([
    slot.Break1Minutes,
    slot.Break1Duration,
    slot.BreakDuration,
    slot.break1Minutes,
    slot.break1Duration
  ], { min: 0 });
  if (break1 !== null) {
    config.breaks.first = break1;
  }

  const break2 = resolveNumberCandidate([
    slot.Break2Minutes,
    slot.Break2Duration,
    slot.break2Minutes,
    slot.break2Duration
  ], { min: 0 });
  if (break2 !== null) {
    config.breaks.second = break2;
  }

  const lunch = resolveNumberCandidate([
    slot.LunchMinutes,
    slot.LunchDuration,
    slot.lunchMinutes,
    slot.lunchDuration
  ], { min: 0 });
  if (lunch !== null) {
    config.breaks.lunch = lunch;
  }

  const enableStaggered = resolveBooleanCandidate([
    slot.EnableStaggeredBreaks,
    slot.enableStaggeredBreaks
  ], config.breaks.enableStaggered);
  config.breaks.enableStaggered = enableStaggered;

  const breakGroups = resolveNumberCandidate([
    slot.BreakGroups,
    slot.breakGroups
  ], { min: 1 });
  if (breakGroups !== null) {
    config.breaks.groups = breakGroups;
  }

  const staggerInterval = resolveNumberCandidate([
    slot.StaggerIntervalMinutes,
    slot.StaggerInterval,
    slot.staggerInterval
  ], { min: 1 });
  if (staggerInterval !== null) {
    config.breaks.interval = staggerInterval;
  }

  const minCoveragePct = resolveNumberCandidate([
    slot.MinCoveragePct,
    slot.minCoveragePct
  ], { min: 0, max: 100 });
  if (minCoveragePct !== null) {
    config.breaks.minCoveragePct = minCoveragePct;
  }

  config.advanced.allowSwaps = resolveBooleanCandidate([
    slot.AllowSwaps,
    slot.AllowSwap,
    slot.allowSwaps,
    slot.allowSwap
  ], config.advanced.allowSwaps);

  config.advanced.weekendPremium = resolveBooleanCandidate([
    slot.WeekendPremium,
    slot.weekendPremium
  ], config.advanced.weekendPremium);

  config.advanced.holidayPremium = resolveBooleanCandidate([
    slot.HolidayPremium,
    slot.holidayPremium
  ], config.advanced.holidayPremium);

  config.advanced.autoAssignment = resolveBooleanCandidate([
    slot.AutoAssignment,
    slot.autoAssignment
  ], config.advanced.autoAssignment);

  const restPeriod = resolveNumberCandidate([
    slot.RestPeriodHours,
    slot.RestPeriod,
    slot.restPeriod
  ], { min: 0 });
  if (restPeriod !== null) {
    config.advanced.restPeriod = restPeriod;
  }

  const notificationLead = resolveNumberCandidate([
    slot.NotificationLeadHours,
    slot.NotificationLead,
    slot.notificationLead
  ], { min: 0 });
  if (notificationLead !== null) {
    config.advanced.notificationLead = notificationLead;
  }

  const handoverTime = resolveNumberCandidate([
    slot.HandoverMinutes,
    slot.HandoverTime,
    slot.handoverTime
  ], { min: 0 });
  if (handoverTime !== null) {
    config.advanced.handoverTime = handoverTime;
  }

  config.overtime.enabled = resolveBooleanCandidate([
    slot.OvertimeEnabled,
    slot.EnableOvertime,
    slot.enableOvertime
  ], config.overtime.enabled);

  const maxDailyOt = resolveNumberCandidate([
    slot.MaxDailyOT,
    slot.maxDailyOT,
    slot.maxDailyOt
  ], { min: 0 });
  if (maxDailyOt !== null) {
    config.overtime.maxDaily = maxDailyOt;
  }

  const maxWeeklyOt = resolveNumberCandidate([
    slot.MaxWeeklyOT,
    slot.maxWeeklyOT,
    slot.maxWeeklyOt
  ], { min: 0 });
  if (maxWeeklyOt !== null) {
    config.overtime.maxWeekly = maxWeeklyOt;
  }

  const overtimeApproval = resolveStringCandidate([
    slot.OTApproval,
    slot.otApproval
  ], config.overtime.approval || 'supervisor');
  config.overtime.approval = overtimeApproval || config.overtime.approval;

  const overtimePolicy = resolveStringCandidate([
    slot.OTPolicy,
    slot.OvertimePolicy,
    slot.otPolicy
  ], config.overtime.policy || 'MANDATORY');
  config.overtime.policy = overtimePolicy || config.overtime.policy;

  const overtimeRate = resolveNumberCandidate([
    slot.OTRate,
    slot.otRate
  ], { min: 1 });
  if (overtimeRate !== null) {
    config.overtime.rate = overtimeRate;
  }

  const sanitized = normalizeGenerationOptions(config);
  sanitized.serialized = serializeSlotConfiguration(sanitized);
  return sanitized;
}

function buildAssignmentBreakConfig(slotConfig = {}, dstAdjustments = [], overrides = {}) {
  const breaks = Object.assign({
    first: 0,
    second: 0,
    lunch: 0,
    enableStaggered: false,
    groups: '',
    interval: '',
    minCoveragePct: ''
  }, slotConfig && slotConfig.breaks ? slotConfig.breaks : {});

  const break1 = coerceScheduleNumber(overrides.break1 ?? breaks.first ?? 0, 0, { min: 0 });
  const break2 = coerceScheduleNumber(overrides.break2 ?? breaks.second ?? 0, 0, { min: 0 });
  const lunch = coerceScheduleNumber(overrides.lunch ?? breaks.lunch ?? 0, 0, { min: 0 });
  const enableStaggered = coerceScheduleBoolean(overrides.enableStaggered ?? breaks.enableStaggered, false);
  const groups = overrides.groups ?? breaks.groups ?? '';
  const interval = overrides.interval ?? breaks.interval ?? '';
  const minCoveragePct = overrides.minCoveragePct ?? breaks.minCoveragePct ?? '';

  const payload = {
    break1,
    break2,
    lunch,
    enableStaggered,
    groups,
    interval,
    minCoveragePct,
    unproductive: (Number.isFinite(break1) ? break1 : 0)
      + (Number.isFinite(break2) ? break2 : 0)
      + (Number.isFinite(lunch) ? lunch : 0)
  };

  if (Array.isArray(dstAdjustments) && dstAdjustments.length) {
    payload.dstAdjustments = dstAdjustments;
  }

  return payload;
}

function getSafeScheduleTimeZone() {
  if (typeof getScheduleTimeZone === 'function') {
    try {
      const tz = getScheduleTimeZone();
      if (tz) {
        return tz;
      }
    } catch (error) {
      console.warn('Falling back to default schedule timezone:', error && error.message ? error.message : error);
    }
  }

  return DEFAULT_SCHEDULE_TIME_ZONE || 'UTC';
}

function normalizeSlotDaysArray(slot) {
  if (!slot) {
    return [];
  }

  if (Array.isArray(slot.DaysOfWeekArray) && slot.DaysOfWeekArray.length) {
    return slot.DaysOfWeekArray.slice();
  }

  if (slot.DaysOfWeek) {
    return parseDaysCsv(slot.DaysOfWeek);
  }

  if (slot.DaysCSV) {
    return parseDaysCsv(slot.DaysCSV);
  }

  if (slot.daysOfWeek) {
    return normalizeDaySelection(slot.daysOfWeek);
  }

  return [];
}

function convertDateToScheduleDayIndex(dateStr) {
  if (!dateStr) {
    return null;
  }

  const date = new Date(`${dateStr}T00:00:00Z`);
  if (isNaN(date.getTime())) {
    return null;
  }

  const day = date.getUTCDay();
  return (day + 6) % 7;
}

function buildDstAdjustmentsForSlot(slot, dateSeries, timeZone) {
  if (!slot || !Array.isArray(dateSeries) || !dateSeries.length) {
    return [];
  }

  const startMinutes = parseTimeToMinutes(slot.StartTime || slot.startTime || '');
  const endMinutes = parseTimeToMinutes(slot.EndTime || slot.endTime || '');

  if (!Number.isFinite(startMinutes) || !Number.isFinite(endMinutes)) {
    return [];
  }

  const slotDays = normalizeSlotDaysArray(slot);
  const normalizedStartTime = formatMinutesTo12Hour(startMinutes);
  const normalizedEndTime = formatMinutesTo12Hour(endMinutes);

  const adjustments = [];

  dateSeries.forEach(dateStr => {
    const dayIndex = convertDateToScheduleDayIndex(dateStr);
    if (slotDays.length && (dayIndex === null || !slotDays.includes(dayIndex))) {
      return;
    }

    const dstInfo = typeof checkDSTStatus === 'function'
      ? checkDSTStatus(dateStr, timeZone)
      : { isDST: false, isDSTChange: false, changeType: null, timeAdjustment: 0 };

    if (!dstInfo.isDST && !dstInfo.isDSTChange) {
      return;
    }

    let adjustmentMinutes = 0;
    let adjustedEndMinutes = endMinutes;
    let adjustedStartMinutes = startMinutes;

    if (dstInfo.isDSTChange && dstInfo.timeAdjustment) {
      adjustmentMinutes = -dstInfo.timeAdjustment;
      adjustedEndMinutes = endMinutes + adjustmentMinutes;
      if (dstInfo.changeType === 'END') {
        adjustedStartMinutes = startMinutes + adjustmentMinutes;
      }
    }

    adjustments.push({
      date: dateStr,
      isDST: !!dstInfo.isDST,
      isDSTChange: !!dstInfo.isDSTChange,
      changeType: dstInfo.changeType || '',
      adjustmentMinutes: adjustmentMinutes,
      originalStartTime: normalizedStartTime,
      originalEndTime: normalizedEndTime,
      adjustedStartTime: formatMinutesTo12Hour(adjustedStartMinutes),
      adjustedEndTime: formatMinutesTo12Hour(adjustedEndMinutes)
    });
  });

  return adjustments;
}

function summarizeDstAdjustments(adjustments) {
  if (!Array.isArray(adjustments) || !adjustments.length) {
    return '';
  }

  const parts = adjustments.map(entry => {
    const direction = entry.adjustmentMinutes > 0 ? `+${entry.adjustmentMinutes}` : String(entry.adjustmentMinutes);
    const change = entry.changeType ? ` (${entry.changeType})` : '';
    return `${entry.date}${change}: ${direction} mins`;
  });

  return `DST adjustments applied - ${parts.join('; ')}`;
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
    let normalizedManagerId = normalizeUserIdValue(requestingUserId);
    const originalManagerId = normalizedManagerId;
    const systemManagerAliases = [
      'system',
      'systemuser',
      'system-user',
      'system_user',
      'systemaccount',
      'system-account'
    ];

    if (normalizedManagerId && systemManagerAliases.includes(normalizedManagerId.toLowerCase())) {
      normalizedManagerId = '';
      console.log('ℹ️ System-level schedule request detected. Applying global roster without manager filtering.');
    }

    console.log(
      '🔍 Getting schedule users for:',
      normalizedManagerId || originalManagerId || '(system)',
      'campaign:',
      normalizedCampaignId || '(not provided)'
    );

    const userLookup = buildScheduleUserLookupIndex();
    const allUsers = Array.isArray(userLookup.users) ? userLookup.users : [];
    if (allUsers.length === 0) {
      console.warn('No users found in Users sheet');
      return [];
    }

    const rosterContext = normalizedManagerId
      ? resolveUnifiedManagedRoster(normalizedManagerId)
      : { users: [], managedUserIds: [] };
    const managedRosterUsers = Array.isArray(rosterContext.users) ? rosterContext.users : [];
    const managedIdSet = new Set(
      Array.isArray(rosterContext.managedUserIds)
        ? rosterContext.managedUserIds
            .map(id => normalizeUserIdValue(id))
            .filter(id => id && id !== normalizedManagerId)
        : []
    );

    if (normalizedManagerId && !managedIdSet.size) {
      const supplemental = buildManagedUserSet(normalizedManagerId);
      supplemental.forEach(id => {
        const normalized = normalizeUserIdValue(id);
        if (normalized && normalized !== normalizedManagerId) {
          managedIdSet.add(normalized);
        }
      });
    }

    let requestingUser = null;
    if (normalizedManagerId) {
      requestingUser = allUsers.find(u => normalizeUserIdValue(u && (u.ID || u.UserID)) === normalizedManagerId) || null;
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

    if (normalizedManagerId && requestingUser && !scheduleFlagToBool(requestingUser.IsAdmin)) {
      if (managedIdSet.size) {
        const matchedUsers = allUsers.filter(user => managedIdSet.has(normalizeUserIdValue(user && (user.ID || user.UserID))));

        if (matchedUsers.length) {
          filteredUsers = matchedUsers;
        } else if (managedRosterUsers.length) {
          filteredUsers = managedRosterUsers;
        } else {
          filteredUsers = [];
          console.warn('Managed roster ids resolved but no matching users found for manager', normalizedManagerId);
        }
      } else {
        filteredUsers = [];
        console.warn('No managed users associated with manager', normalizedManagerId, '- returning empty roster.');
      }
    } else if (normalizedManagerId && !requestingUser) {
      filteredUsers = [];
      console.warn('Requesting user not found when applying manager filter:', requestingUserId);
    } else if (effectiveCampaignId) {
      filteredUsers = filterUsersByCampaign(allUsers, effectiveCampaignId);
    }

    if (!filteredUsers.length && effectiveCampaignId && (!normalizedManagerId || scheduleFlagToBool(requestingUser && requestingUser.IsAdmin))) {
      filteredUsers = filterUsersByCampaign(allUsers, effectiveCampaignId);
    }

    const scheduleUsers = filteredUsers
      .filter(user => user && (user.ID || user.UserID) && (user.UserName || user.FullName || user.Username))
      .filter(user => !isScheduleNameRestricted(user))
      .filter(user => !isScheduleRoleRestricted(user))
      .filter(user => isUserConsideredActive(user))
      .map(user => normalizeScheduleUserRecord(user, userLookup))
      .filter(Boolean);

    console.log(`✅ Returning ${scheduleUsers.length} schedule users`);
    return scheduleUsers;

  } catch (error) {
    console.error('❌ Error getting schedule users:', error);
    safeWriteError('clientGetScheduleUsers', error);
    return [];
  }
}

function normalizeScheduleUserRecord(user, lookup = null) {
  if (!user || typeof user !== 'object') {
    return null;
  }

  const normalizedId = normalizeUserIdValue(user.ID || user.UserID || user.id || user.userId || user.UserName || user.username);
  if (!normalizedId) {
    return null;
  }

  let baseRecord = user;
  if (lookup && typeof lookup === 'object' && Array.isArray(lookup.users)) {
    const lookupRecord = lookup.users.find(entry => normalizeUserIdValue(entry && (entry.ID || entry.UserID)) === normalizedId);
    if (lookupRecord) {
      baseRecord = Object.assign({}, lookupRecord, baseRecord);
    }
  }

  const campaignId = normalizeCampaignIdValue(
    baseRecord.CampaignID
      || baseRecord.campaignID
      || baseRecord.CampaignId
      || baseRecord.campaignId
  );

  let campaignName = baseRecord.campaignName
    || baseRecord.CampaignName
    || baseRecord.campaign
    || baseRecord.Campaign
    || '';

  if (!campaignName && campaignId && typeof getCampaignById === 'function') {
    try {
      const campaignRecord = getCampaignById(campaignId);
      if (campaignRecord) {
        campaignName = campaignRecord.Name || campaignRecord.name || campaignName;
      }
    } catch (campaignError) {
      console.warn('Unable to resolve campaign name for user', campaignId, campaignError);
    }
  }

  return {
    ID: normalizedId,
    UserName: baseRecord.UserName || baseRecord.Username || baseRecord.username || baseRecord.FullName || '',
    FullName: baseRecord.FullName || baseRecord.fullName || baseRecord.UserName || baseRecord.Username || '',
    Email: baseRecord.Email || baseRecord.email || '',
    CampaignID: campaignId || '',
    campaignName: campaignName || '',
    EmploymentStatus: baseRecord.EmploymentStatus || baseRecord.employmentStatus || 'Active',
    HireDate: baseRecord.HireDate || baseRecord.hireDate || '',
    TerminationDate: baseRecord.TerminationDate || baseRecord.terminationDate || '',
    isActive: isUserConsideredActive(baseRecord),
    roleNames: baseRecord.roleNames || baseRecord.RoleNames || []
  };
}

/**
 * Get users for attendance (all active users)
 */
function clientGetAttendanceUsers(requestingUserId, campaignId = null) {
  try {
    console.log('📋 Getting attendance users');

    const normalizedManagerId = normalizeUserIdValue(requestingUserId);
    const scheduleUsers = clientGetScheduleUsers(requestingUserId, campaignId) || [];
    const scheduleById = new Map();
    scheduleUsers.forEach(user => {
      const normalizedId = normalizeUserIdValue(user && (user.ID || user.UserID));
      if (normalizedId) {
        scheduleById.set(normalizedId, user);
      }
    });

    const managedRoster = normalizedManagerId
      ? clientGetManagedUsersList(normalizedManagerId)
      : [];

    const managedById = new Map();
    managedRoster.forEach(user => {
      const normalizedId = normalizeUserIdValue(user && (user.ID || user.UserID));
      if (normalizedId && normalizedId !== normalizedManagerId && !managedById.has(normalizedId)) {
        managedById.set(normalizedId, user);
      }
    });

    const finalRoster = new Map();

    if (normalizedManagerId && managedById.size) {
      managedById.forEach((rosterUser, userId) => {
        const bestRecord = scheduleById.get(userId) || rosterUser;
        if (!bestRecord || normalizeUserIdValue(bestRecord && (bestRecord.ID || bestRecord.UserID)) === normalizedManagerId) {
          return;
        }

        finalRoster.set(userId, bestRecord);
      });
    } else {
      scheduleById.forEach((user, userId) => {
        if (userId && (!normalizedManagerId || userId !== normalizedManagerId)) {
          finalRoster.set(userId, user);
        }
      });
    }

    if (!finalRoster.size && normalizedManagerId) {
      const lookup = buildScheduleUserLookupIndex();
      const managerRecord = scheduleById.get(normalizedManagerId)
        || managedById.get(normalizedManagerId)
        || (Array.isArray(lookup.users)
          ? lookup.users.find(user => normalizeUserIdValue(user && (user.ID || user.UserID)) === normalizedManagerId)
          : null);

      if (managerRecord) {
        finalRoster.set(normalizedManagerId, managerRecord);
      }
    }

    const userNames = Array.from(finalRoster.values())
      .map(candidate => {
        if (!candidate) {
          return '';
        }

        const name = candidate.FullName || candidate.UserName || candidate.Email || '';
        return name ? name.toString().trim() : '';
      })
      .filter(Boolean)
      .sort((a, b) => a.localeCompare(b));

    const sourceLabel = normalizedManagerId && managedById.size ? 'managed roster' : 'schedule users';
    console.log(`✅ Returning ${userNames.length} attendance users (source: ${sourceLabel})`);
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

    let hasDirectAssignments = false;
    if (typeof clientGetManagedUsers === 'function') {
      try {
        const managedResponse = clientGetManagedUsers(normalizedManagerId);
        if (managedResponse && managedResponse.success !== false) {
          const assignedUsers = Array.isArray(managedResponse.users)
            ? managedResponse.users
            : [];
          assignedUsers.forEach(user => {
            hasDirectAssignments = true;
            pushUser(user, {
              campaignId: user && (user.CampaignID || user.campaignID || user.campaignId || user.CampaignId),
              campaignName: user && (user.campaignName || user.CampaignName || user.campaign)
            });
          });
        }
      } catch (managedError) {
        console.warn('Unable to load direct managed users for roster', normalizedManagerId, managedError);
      }
    }

    if (!hasDirectAssignments) {
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
    const actorInfo = resolveScheduleActor(actor);
    const actorLabel = actorInfo.label;
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

    const campaign = (slotData.campaign || slotData.Campaign || slotData.department || slotData.Department || 'General').toString().trim();
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
      CreatedBy: actorLabel,
      Notes: slotData.notes || '',
      Status: 'Active',
      CreatedAt: now,
      UpdatedAt: now,
      UpdatedBy: actorLabel,
      ConfigurationJSON: '',
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

function normalizeShiftSlotIdentifier(slot) {
  if (slot === null || typeof slot === 'undefined') {
    return '';
  }

  if (typeof slot === 'object') {
    const candidateKeys = [
      slot.ID, slot.Id, slot.id,
      slot.SlotID, slot.SlotId, slot.slotId,
      slot.Guid, slot.GUID, slot.UUID, slot.Uuid,
      slot.Name, slot.SlotName
    ];

    for (let index = 0; index < candidateKeys.length; index++) {
      const value = candidateKeys[index];
      if (value === null || typeof value === 'undefined') {
        continue;
      }

      const normalized = String(value).trim();
      if (normalized) {
        return normalized;
      }
    }

    return '';
  }

  return String(slot).trim();
}

function clientDeleteShiftSlot(slot) {
  try {
    const normalizedId = normalizeShiftSlotIdentifier(slot);

    if (!normalizedId) {
      return {
        success: false,
        error: 'A valid shift slot identifier is required to delete a shift slot.'
      };
    }

    const sheet = ensureScheduleSheetWithHeaders(SHIFT_SLOTS_SHEET, SHIFT_SLOTS_HEADERS);
    const data = sheet.getDataRange().getValues();

    if (!Array.isArray(data) || data.length <= 1) {
      return {
        success: false,
        error: 'No shift slots are available to delete.'
      };
    }

    const headers = data[0].map(header => (header || '').toString());
    const headerLookup = headers.map(header => header.trim().toLowerCase());
    const idIndex = headerLookup.indexOf('id');
    const nameIndex = headerLookup.indexOf('name');
    const slotNameIndex = headerLookup.indexOf('slotname');

    let rowToDelete = -1;
    let deletedSlotRecord = null;
    const normalizedLowerId = normalizedId.toLowerCase();

    for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
      const rowValues = data[rowIndex];
      const record = {};

      headers.forEach((header, columnIndex) => {
        record[header] = rowValues[columnIndex];
      });

      const candidateIds = [];

      if (idIndex !== -1) {
        candidateIds.push(rowValues[idIndex]);
      }

      if (slotNameIndex !== -1) {
        candidateIds.push(rowValues[slotNameIndex]);
      }

      if (nameIndex !== -1 && nameIndex !== slotNameIndex) {
        candidateIds.push(rowValues[nameIndex]);
      }

      candidateIds.push(record.ID, record.Id, record.id);
      candidateIds.push(record.SlotID, record.SlotId, record.slotId);
      candidateIds.push(record.Guid, record.GUID, record.UUID, record.Uuid);

      const hasMatch = candidateIds.some(candidate => {
        if (candidate === null || typeof candidate === 'undefined') {
          return false;
        }

        const text = String(candidate).trim();
        if (!text) {
          return false;
        }

        return text.toLowerCase() === normalizedLowerId;
      });

      if (hasMatch) {
        rowToDelete = rowIndex + 1; // account for header row offset
        deletedSlotRecord = record;
        break;
      }
    }

    if (rowToDelete === -1 || !deletedSlotRecord) {
      const fallbackSeparatorIndex = normalizedId.indexOf('|');
      if (fallbackSeparatorIndex !== -1) {
        const fallbackName = normalizedId.slice(0, fallbackSeparatorIndex).trim().toLowerCase();
        const fallbackTimeRange = normalizedId.slice(fallbackSeparatorIndex + 1).trim();
        const [fallbackStartRaw, fallbackEndRaw] = fallbackTimeRange.split('-').map(part => part ? part.trim() : '');
        const fallbackStart = normalizeTimeTo12Hour(fallbackStartRaw || '') || '';
        const fallbackEnd = normalizeTimeTo12Hour(fallbackEndRaw || '') || '';

        for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
          const rowValues = data[rowIndex];
          const record = {};

          headers.forEach((header, columnIndex) => {
            record[header] = rowValues[columnIndex];
          });

          const recordName = (record.Name || record.SlotName || '').toString().trim().toLowerCase();
          if (!recordName || recordName !== fallbackName) {
            continue;
          }

          const recordStart = normalizeTimeTo12Hour(record.StartTime || record.startTime || record['Start Time'] || '') || '';
          const recordEnd = normalizeTimeTo12Hour(record.EndTime || record.endTime || record['End Time'] || '') || '';

          const startMatches = fallbackStart ? recordStart === fallbackStart : true;
          const endMatches = fallbackEnd ? recordEnd === fallbackEnd : true;

          if (startMatches && endMatches) {
            rowToDelete = rowIndex + 1;
            deletedSlotRecord = record;
            break;
          }
        }
      }

      if (rowToDelete === -1 || !deletedSlotRecord) {
        return {
          success: false,
          error: 'Shift slot not found. It may have already been deleted.',
          slotId: normalizedId
        };
      }
    }

    sheet.deleteRow(rowToDelete);
    SpreadsheetApp.flush();

    appendAuditLogEntry('DELETE', 'ShiftSlot', normalizedId, deletedSlotRecord, null, 'Deleted shift slot');
    invalidateScheduleCaches();

    return {
      success: true,
      message: 'Shift slot deleted successfully.',
      slotId: normalizedId,
      slot: deletedSlotRecord
    };

  } catch (error) {
    console.error('Error deleting shift slot:', error);
    safeWriteError('clientDeleteShiftSlot', error);
    return {
      success: false,
      error: error && error.message ? error.message : 'Failed to delete shift slot.'
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
      const slotId = (
        slot.SlotId ||
        slot.SlotID ||
        slot['Slot ID'] ||
        slot['Slot Id'] ||
        slot.ID ||
        slot.Id ||
        slot.slotId ||
        slot.id ||
        ''
      ).toString().trim() || Utilities.getUuid();
      const slotName = (slot.SlotName || slot.Name || '').toString().trim();
      const campaign = (slot.Campaign || slot.Department || '').toString().trim();
      const location = (slot.Location || '').toString().trim() || 'Office';
      const startTime = normalizeTimeTo12Hour(slot.StartTime || slot.startTime || slot['Start Time'] || slot.Start || '');
      const endTime = normalizeTimeTo12Hour(slot.EndTime || slot.endTime || slot['End Time'] || slot.End || '');
      const daysArray = parseDaysCsv(slot.DaysCSV || slot.DaysOfWeek || '');
      const statusValue = (slot.Status || '').toString().trim().toUpperCase();
      const status = statusValue || (scheduleFlagToBool(slot.IsActive, true) ? 'Active' : 'Archived');

      const slotConfig = normalizeSlotConfiguration(slot || {}, {});
      const configurationJson = slot.ConfigurationJSON || slotConfig.serialized || serializeSlotConfiguration(slotConfig);
      const overtimeMinutes = slotConfig.overtime && slotConfig.overtime.enabled
        ? Math.round(Number(slotConfig.overtime.maxDaily || 0) * 60)
        : 0;

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
        UpdatedBy: slot.UpdatedBy || '',
        ConfigurationJSON: configurationJson,
        SlotConfiguration: slotConfig,
        CapacityMax: slotConfig.capacity ? slotConfig.capacity.max : '',
        CapacityMin: slotConfig.capacity ? slotConfig.capacity.min : '',
        Break1Minutes: slotConfig.breaks ? slotConfig.breaks.first : '',
        Break2Minutes: slotConfig.breaks ? slotConfig.breaks.second : '',
        LunchMinutes: slotConfig.breaks ? slotConfig.breaks.lunch : '',
        EnableStaggeredBreaks: slotConfig.breaks ? slotConfig.breaks.enableStaggered : false,
        BreakGroups: slotConfig.breaks ? slotConfig.breaks.groups : '',
        StaggerIntervalMinutes: slotConfig.breaks ? slotConfig.breaks.interval : '',
        MinCoveragePct: slotConfig.breaks ? slotConfig.breaks.minCoveragePct : '',
        AllowSwaps: slotConfig.advanced ? slotConfig.advanced.allowSwaps : true,
        WeekendPremium: slotConfig.advanced ? slotConfig.advanced.weekendPremium : false,
        HolidayPremium: slotConfig.advanced ? slotConfig.advanced.holidayPremium : false,
        AutoAssignment: slotConfig.advanced ? slotConfig.advanced.autoAssignment : false,
        RestPeriodHours: slotConfig.advanced ? slotConfig.advanced.restPeriod : '',
        NotificationLeadHours: slotConfig.advanced ? slotConfig.advanced.notificationLead : '',
        HandoverMinutes: slotConfig.advanced ? slotConfig.advanced.handoverTime : '',
        OvertimeEnabled: slotConfig.overtime ? slotConfig.overtime.enabled : false,
        MaxDailyOT: slotConfig.overtime ? slotConfig.overtime.maxDaily : '',
        MaxWeeklyOT: slotConfig.overtime ? slotConfig.overtime.maxWeekly : '',
        OvertimeApproval: slotConfig.overtime ? slotConfig.overtime.approval : '',
        OvertimeRate: slotConfig.overtime ? slotConfig.overtime.rate : '',
        OvertimePolicy: slotConfig.overtime ? slotConfig.overtime.policy : '',
        OvertimeMinutes: overtimeMinutes
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

    const scheduleTimeZone = getSafeScheduleTimeZone();

    const startDateObj = new Date(normalizedStart);
    const endDateObj = new Date(normalizedEnd);
    if (startDateObj > endDateObj) {
      return {
        success: false,
        error: 'End date must be on or after the start date.'
      };
    }

    const campaignId = normalizeCampaignIdValue(options.campaignId || '');
    const actorInfo = resolveScheduleActor(generatedBy || null);
    const actorLabel = actorInfo.label;
    const actorLookupKey = actorInfo.lookupKey || 'system';
    const detectConflicts = options.detectConflicts !== false;
    const includeHolidays = options.includeHolidays !== false;
    const normalizedGeneration = normalizeGenerationOptions(options || {});
    const advancedOptions = options.advanced || {};
    const capacityOptions = options.capacity || {};
    const breaksOptions = options.breaks || {};
    const overtimeOptions = options.overtime || {};

    const globalAllowSwaps = coerceScheduleBoolean(advancedOptions.allowSwaps, normalizedGeneration.advanced.allowSwaps);
    const globalRestHours = coerceScheduleNumber(advancedOptions.restPeriod, normalizedGeneration.advanced.restPeriod, { min: 0 });
    const globalNotificationLead = coerceScheduleNumber(advancedOptions.notificationLead, normalizedGeneration.advanced.notificationLead, { min: 0 });
    const globalHandoverMinutes = coerceScheduleNumber(advancedOptions.handoverTime, normalizedGeneration.advanced.handoverTime, { min: 0 });
    const globalOvertimeEnabled = coerceScheduleBoolean(overtimeOptions.enabled, normalizedGeneration.overtime.enabled);
    const globalOvertimeMaxDaily = coerceScheduleNumber(overtimeOptions.maxDaily, normalizedGeneration.overtime.maxDaily, { min: 0 });

    const maxCapacityValue = coerceScheduleNumber(
      Object.prototype.hasOwnProperty.call(capacityOptions, 'max') ? capacityOptions.max : options.maxCapacity,
      normalizedGeneration.capacity.max,
      { min: 0 }
    );
    const maxCapacity = maxCapacityValue > 0 ? maxCapacityValue : null;

    const minCoverage = coerceScheduleNumber(
      Object.prototype.hasOwnProperty.call(capacityOptions, 'min') ? capacityOptions.min : options.minCoverage,
      normalizedGeneration.capacity.min,
      { min: 0 }
    );

    const minCoveragePct = coerceScheduleNumber(
      Object.prototype.hasOwnProperty.call(breaksOptions, 'minCoveragePct') ? breaksOptions.minCoveragePct : options.minCoveragePct,
      normalizedGeneration.breaks.minCoveragePct,
      { min: 0, max: 100 }
    );

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

    const slotConfigMap = new Map();
    const slotRestRequirements = new Map();
    selectedSlots = selectedSlots.map(slot => {
      const config = normalizeSlotConfiguration(slot, normalizedGeneration);
      const configurationJson = config.serialized || serializeSlotConfiguration(config);
      if (slot && slot.SlotId) {
        slotConfigMap.set(slot.SlotId, config);
        const restRequirement = coerceScheduleNumber(config.advanced ? config.advanced.restPeriod : undefined, globalRestHours, { min: 0 });
        slotRestRequirements.set(slot.SlotId, restRequirement);
      }
      return Object.assign({}, slot, {
        SlotConfiguration: config,
        ConfigurationJSON: configurationJson
      });
    });

    const slotMap = new Map(selectedSlots.map(slot => [slot.SlotId, slot]));
    const scheduleUsers = clientGetScheduleUsers(actorLookupKey || 'system', campaignId || null);
    const userKeyMap = new Map();
    const userIdMap = new Map();
    scheduleUsers.forEach(user => {
      if (!user) {
        return;
      }

      const idCandidates = [
        user.ID,
        user.Id,
        user.UserID,
        user.UserId,
        user.id,
        user.userId,
        user.AgentID,
        user.AgentId
      ];
      idCandidates.forEach(candidate => {
        const normalized = normalizeUserIdValue(candidate);
        if (normalized && !userIdMap.has(normalized)) {
          userIdMap.set(normalized, user);
        }
      });

      const keyCandidates = [
        user.UserName,
        user.Username,
        user.username,
        user.FullName,
        user.fullName,
        user.Email,
        user.email
      ];
      keyCandidates.forEach(candidate => {
        const key = normalizeUserKey(candidate);
        if (key && !userKeyMap.has(key)) {
          userKeyMap.set(key, user);
        }
      });
    });

    let targetUsers = [];
    const unresolvedUsers = [];
    const explicitlyRequestedIds = new Set();
    const explicitlyRequestedNameKeys = new Set();
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
          const normalizedId = normalizeUserIdValue(user.ID || user.UserID || user.id || user.userId);
          if (normalizedId) {
            explicitlyRequestedIds.add(normalizedId);
          }
          const resolvedKey = normalizeUserKey(user.UserName || user.FullName || user.Username || user.Email);
          if (resolvedKey) {
            explicitlyRequestedNameKeys.add(resolvedKey);
          }
        } else {
          unresolvedUsers.push(entry);
        }
      });
    } else {
      targetUsers = scheduleUsers.slice();
    }

    const normalizedCampaignId = campaignId ? campaignId.toLowerCase() : '';
    const filteredUsers = targetUsers.filter(user => {
      if (!user) {
        return false;
      }

      const normalizedId = normalizeUserIdValue(
        user.ID || user.Id || user.UserID || user.UserId || user.id || user.userId
      );
      if (!normalizedId) {
        return false;
      }

      const activeFlag = coerceScheduleBoolean(
        Object.prototype.hasOwnProperty.call(user, 'isActive') ? user.isActive
          : (Object.prototype.hasOwnProperty.call(user, 'IsActive') ? user.IsActive : undefined),
        true
      );
      if (!activeFlag) {
        return false;
      }

      const employmentStatus = (user.EmploymentStatus || user.Status || user.employmentStatus || '').toString().trim().toLowerCase();
      if (employmentStatus && ['inactive', 'terminated', 'disabled', 'separated', 'archived'].includes(employmentStatus)) {
        return false;
      }

      const explicitRequest = (normalizedId && explicitlyRequestedIds.has(normalizedId))
        || explicitlyRequestedNameKeys.has(normalizeUserKey(user.UserName || user.FullName || user.Username || user.Email));

      if (normalizedCampaignId) {
        const userCampaignId = normalizeCampaignIdValue(
          user.CampaignID
            || user.campaignID
            || user.CampaignId
            || user.campaignId
            || user.Campaign
            || user.campaign
            || user.primaryCampaignId
            || user.PrimaryCampaignId
        );

        if (userCampaignId) {
          if (userCampaignId.toString().trim().toLowerCase() !== normalizedCampaignId && !explicitRequest) {
            return false;
          }
        } else if (!explicitRequest) {
          return false;
        }
      }

      const hireDateCandidate = user.HireDate || user.hireDate || user.StartDate || user.startDate;
      if (hireDateCandidate) {
        const hireDate = new Date(hireDateCandidate);
        if (!isNaN(hireDate.getTime()) && hireDate > endDateObj) {
          return false;
        }
      }

      const terminationCandidate = user.TerminationDate || user.terminationDate || user.EndDate || user.endDate;
      if (terminationCandidate) {
        const terminationDate = new Date(terminationCandidate);
        if (!isNaN(terminationDate.getTime()) && terminationDate < startDateObj && !explicitRequest) {
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

    const dateSeries = buildDateSeries(normalizedStart, normalizedEnd);

    const seed = options.seed || `${campaignId || 'ALL'}-${normalizedStart}-${normalizedEnd}-${(shiftSlotIds || []).join('|')}`;
    const orderedUsers = shuffleWithSeed(filteredUsers, seed);
    const slotCounts = new Map();
    const assignments = [];
    const skippedUsers = [];
    const now = new Date();
    const slotRandom = createSeededRandom(`${seed}-slot-pick`);
    orderedUsers.forEach(user => {
      const availableSlots = selectedSlots.filter(slot => {
        if (!slot) {
          return false;
        }
        if (!maxCapacity) {
          return true;
        }
        const slotCount = slotCounts.get(slot.SlotId) || 0;
        return slotCount < maxCapacity;
      });

      if (!availableSlots.length) {
        skippedUsers.push({
          userId: user.ID,
          userName: user.UserName || user.FullName,
          reason: 'Max capacity reached for selected slots'
        });
        return;
      }

      const randomIndex = Math.floor(slotRandom() * availableSlots.length);
      const assignedSlot = availableSlots[randomIndex];

      if (!assignedSlot) {
        skippedUsers.push({
          userId: user.ID,
          userName: user.UserName || user.FullName,
          reason: 'Max capacity reached for selected slots'
        });
        return;
      }

      slotCounts.set(assignedSlot.SlotId, (slotCounts.get(assignedSlot.SlotId) || 0) + 1);

      const slotConfig = slotConfigMap.get(assignedSlot.SlotId) || normalizeSlotConfiguration(assignedSlot, normalizedGeneration);
      const dstAdjustments = buildDstAdjustmentsForSlot(assignedSlot, dateSeries, scheduleTimeZone);
      const breakConfig = buildAssignmentBreakConfig(slotConfig, dstAdjustments);

      const allowSwapForSlot = coerceScheduleBoolean(
        slotConfig.advanced ? slotConfig.advanced.allowSwaps : undefined,
        globalAllowSwaps
      );
      const restHoursForSlot = coerceScheduleNumber(
        slotConfig.advanced ? slotConfig.advanced.restPeriod : undefined,
        globalRestHours,
        { min: 0 }
      );
      const notificationLeadForSlot = coerceScheduleNumber(
        slotConfig.advanced ? slotConfig.advanced.notificationLead : undefined,
        globalNotificationLead,
        { min: 0 }
      );
      const handoverMinutesForSlot = coerceScheduleNumber(
        slotConfig.advanced ? slotConfig.advanced.handoverTime : undefined,
        globalHandoverMinutes,
        { min: 0 }
      );
      const overtimeEnabledForSlot = coerceScheduleBoolean(
        slotConfig.overtime ? slotConfig.overtime.enabled : undefined,
        globalOvertimeEnabled
      );
      const overtimeMaxDailyForSlot = coerceScheduleNumber(
        slotConfig.overtime ? slotConfig.overtime.maxDaily : undefined,
        globalOvertimeMaxDaily,
        { min: 0 }
      );
      const overtimeMinutesForSlot = overtimeEnabledForSlot ? Math.round(overtimeMaxDailyForSlot * 60) : '';

      if (assignedSlot && assignedSlot.SlotId) {
        const existingRestRequirement = slotRestRequirements.get(assignedSlot.SlotId) || 0;
        slotRestRequirements.set(assignedSlot.SlotId, Math.max(existingRestRequirement, restHoursForSlot));
      }

      const dstNotes = summarizeDstAdjustments(dstAdjustments);
      const assignmentNotes = [options.notes || '', dstNotes].filter(Boolean).join(' | ');

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
        AllowSwap: allowSwapForSlot,
        Premiums: '',
        BreaksConfigJSON: JSON.stringify(breakConfig),
        OvertimeMinutes: overtimeMinutesForSlot || '',
        RestPeriodHours: restHoursForSlot || '',
        NotificationLeadHours: notificationLeadForSlot || '',
        HandoverMinutes: handoverMinutesForSlot || '',
        Notes: assignmentNotes,
        CreatedAt: now,
        CreatedBy: actorLabel,
        UpdatedAt: now,
        UpdatedBy: actorLabel
      });
    });

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

    const getRestRequirement = slotId => {
      if (!slotId) {
        return globalRestHours || 0;
      }
      if (slotRestRequirements.has(slotId)) {
        return slotRestRequirements.get(slotId) || 0;
      }
      const slotRecord = slotMap.get(slotId);
      if (slotRecord) {
        const config = slotConfigMap.get(slotId) || normalizeSlotConfiguration(slotRecord, normalizedGeneration);
        const restRequirement = coerceScheduleNumber(config.advanced ? config.advanced.restPeriod : undefined, globalRestHours, { min: 0 });
        slotRestRequirements.set(slotId, restRequirement);
        return restRequirement || 0;
      }
      return globalRestHours || 0;
    };

    const checkRestPeriod = (existing, generatedSlot, assignment) => {
      const restRequirement = Math.max(
        getRestRequirement(existing.SlotId),
        getRestRequirement(assignment.SlotId)
      );
      if (!restRequirement || !generatedSlot) {
        return false;
      }
      const candidateSlot = slotMap.get(existing.SlotId) || {};
      const existingStart = new Date(`${existing.StartDate}T00:00:00`);
      const existingEnd = new Date(`${existing.EndDate}T00:00:00`);
      const generatedStart = new Date(`${assignment.StartDate}T00:00:00`);
      const generatedEnd = new Date(`${assignment.EndDate}T00:00:00`);
      const existingStartMinutes = parseTimeToMinutes(candidateSlot.StartTime || candidateSlot.startTime || existing.StartTime || '');
      const existingEndMinutes = parseTimeToMinutes(candidateSlot.EndTime || candidateSlot.endTime || existing.EndTime || '');
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
      return diffHours < restRequirement;
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

      const restRequirement = Math.max(globalRestHours || 0, getRestRequirement(assignment.SlotId));
      if (restRequirement > 0) {
        const restConflict = existingForUser.some(existing => checkRestPeriod(existing, slot, assignment));
        if (restConflict) {
          conflicts.push({
            userId: assignment.UserId,
            userName: assignment.UserName,
            type: 'REST_VIOLATION',
            periodStart: assignment.StartDate,
            periodEnd: assignment.EndDate,
            error: `Rest period requirement of ${restRequirement} hours would be violated`
          });
          if (detectConflicts) {
            return false;
          }
        }
      }

      const premiumSet = new Set();
      const slotConfig = slotConfigMap.get(assignment.SlotId) || slot.SlotConfiguration || normalizeSlotConfiguration(slot, normalizedGeneration);
      const advancedConfig = slotConfig && slotConfig.advanced ? slotConfig.advanced : {};
      const overtimeConfig = slotConfig && slotConfig.overtime ? slotConfig.overtime : {};
      const datesForAssignment = dateSeries.filter(date => date >= assignment.StartDate && date <= assignment.EndDate);
      const hasWeekend = datesForAssignment.some(isWeekendDate);
      if (hasWeekend && coerceScheduleBoolean(advancedConfig.weekendPremium, false)) {
        premiumSet.add('Weekend');
      }
      const hasHoliday = datesForAssignment.some(date => {
        const entries = holidayMap.get(date) || [];
        return entries.some(entry => (entry.region || '').toLowerCase() === 'jamaica');
      });
      if (hasHoliday && coerceScheduleBoolean(advancedConfig.holidayPremium, true)) {
        premiumSet.add('Holiday');
      }
      if (coerceScheduleBoolean(overtimeConfig.enabled, globalOvertimeEnabled)) {
        premiumSet.add('Overtime');
      }
      assignmentPremiums.set(assignment.AssignmentId, Array.from(premiumSet));
      assignment.Premiums = Array.from(premiumSet).join(',');
      return true;
    });

    const weekendPremiumGlobal = coerceScheduleBoolean(normalizedGeneration.advanced.weekendPremium, false);
    const holidayPremiumGlobal = coerceScheduleBoolean(normalizedGeneration.advanced.holidayPremium, true);

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
          weekend: weekend && weekendPremiumGlobal,
          holiday: holidayEntries.some(entry => (entry.region || '').toLowerCase() === 'jamaica') && holidayPremiumGlobal
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
      const commitResult = writeShiftAssignments(cached.assignments, actorLabel, options.notes || 'Auto-assigned schedule generation', 'PENDING');
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
    Training: 0,
    'Sick Leave': 0,
    'Bereavement': 0,
    'Vacation': 0,
    'Leave Of Absence': 0,
    'Maternity Leave': 0,
    'No Call No Show': 0,
    Other: 0
  };

  const statusKeyMap = {
    present: 'Present',
    punctual: 'Present',
    late: 'Late',
    training: 'Training',
    'sick leave': 'Sick Leave',
    sick: 'Sick Leave',
    bereavement: 'Bereavement',
    vacation: 'Vacation',
    'leave of absence': 'Leave Of Absence',
    'leave of absent': 'Leave Of Absence',
    'maternity leave': 'Maternity Leave',
    'no call no show': 'No Call No Show',
    'no call/no show': 'No Call No Show',
    'no call/no-show': 'No Call No Show'
  };

  attendanceData.forEach(record => {
    const normalized = (record.Status || 'Other').toString().trim().toLowerCase();
    const key = statusKeyMap[normalized] || (statusCounts.hasOwnProperty(record.Status) ? record.Status : 'Other');
    if (statusCounts.hasOwnProperty(key)) {
      statusCounts[key]++;
    } else {
      statusCounts.Other++;
    }
  });

  const total = Object.values(statusCounts).reduce((sum, count) => sum + count, 0);
  const totalAbsences = statusCounts.Absent
    + statusCounts['Sick Leave']
    + statusCounts['Leave Of Absence']
    + statusCounts['No Call No Show'];

  const percentages = {};
  Object.keys(statusCounts).forEach(status => {
    percentages[status] = total > 0 ? Math.round((statusCounts[status] / total) * 100) : 0;
  });

  const baseAttendanceRate = total > 0 ? Math.round(((total - totalAbsences) / total) * 100) : 0;
  const latePenalty = Math.min(statusCounts.Late, 100);
  const attendanceRate = Math.max(0, Math.min(100, baseAttendanceRate - latePenalty));
  const absenceRate = total > 0 ? Math.round((totalAbsences / total) * 100) : 0;

  return {
    counts: statusCounts,
    percentages: percentages,
    total: total,
    attendanceRate,
    absenceRate
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

    const normalized = (record.Status || '').toString().trim().toLowerCase();
    const isNoCall = normalized === 'no call no show' || normalized === 'no call/no show' || normalized === 'no call/no-show';
    const isSick = normalized === 'sick' || normalized === 'sick leave';
    const isLeaveOfAbsence = normalized === 'leave of absence' || normalized === 'leave of absent';

    switch (normalized) {
      case 'present':
      case 'punctual':
        stats.present++;
        break;
      case 'late':
        stats.present++;
        stats.late++;
        break;
      case 'training':
        stats.present++;
        stats.other++;
        break;
      case 'absent':
        stats.absent++;
        break;
      default:
        if (isSick) {
          stats.sick++;
          stats.absent++;
        } else if (isNoCall || isLeaveOfAbsence) {
          stats.absent++;
        } else {
          stats.other++;
        }
        break;
    }
  });

  // Calculate attendance rates
  Object.values(userStats).forEach(stats => {
    if (stats.totalRecords > 0) {
      const baseRate = Math.round((stats.present / stats.totalRecords) * 100);
      const latePenalty = Math.min(stats.late, 100);
      stats.attendanceRate = Math.max(0, Math.min(100, baseRate - latePenalty));
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
    const normalized = (record.Status || '').toString().trim().toLowerCase();
    const presentStatuses = new Set(['present', 'punctual', 'training', 'late']);
    const absenceStatuses = new Set([
      'absent',
      'sick',
      'sick leave',
      'leave of absence',
      'leave of absent',
      'no call no show',
      'no call/no show',
      'no call/no-show'
    ]);

    // Daily stats
    if (!dailyStats[dayKey]) {
      dailyStats[dayKey] = { date: dayKey, present: 0, absent: 0, total: 0 };
    }
    dailyStats[dayKey].total++;
    if (presentStatuses.has(normalized)) {
      dailyStats[dayKey].present++;
    } else if (absenceStatuses.has(normalized)) {
      dailyStats[dayKey].absent++;
    }

    // Weekly stats
    if (!weeklyStats[weekKey]) {
      weeklyStats[weekKey] = { week: weekKey, present: 0, absent: 0, total: 0 };
    }
    weeklyStats[weekKey].total++;
    if (presentStatuses.has(normalized)) {
      weeklyStats[weekKey].present++;
    } else if (absenceStatuses.has(normalized)) {
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
// ATTENDANCE EXPORTS (Dashboard vs Calendar)
// ────────────────────────────────────────────────────────────────────────────

function normalizeDateRangeForExport(periodType, startDate, endDate) {
  const normalize = (value) => {
    if (value instanceof Date) {
      return isNaN(value.getTime()) ? null : new Date(value.getTime());
    }

    if (typeof value === 'number' && Number.isFinite(value)) {
      const parsed = new Date(value);
      return isNaN(parsed.getTime()) ? null : parsed;
    }

    if (typeof value === 'string') {
      const trimmed = value.trim();
      if (!trimmed) return null;
      const parsed = new Date(trimmed);
      return isNaN(parsed.getTime()) ? null : parsed;
    }

    return null;
  };

  const safeStart = normalize(startDate);
  const safeEnd = normalize(endDate);

  if (!safeStart || !safeEnd || safeEnd < safeStart) {
    throw new Error('A valid start and end date are required for export.');
  }

  const getMondayStart = (date) => {
    const base = new Date(date.getTime());
    const day = (base.getDay() + 6) % 7; // Sunday -> 6, Monday -> 0
    base.setDate(base.getDate() - day);
    base.setHours(0, 0, 0, 0);
    return base;
  };

  const normalizedStart = new Date(safeStart.getTime());
  const normalizedEnd = new Date(safeEnd.getTime());
  normalizedStart.setHours(0, 0, 0, 0);
  normalizedEnd.setHours(23, 59, 59, 999);

  const type = typeof periodType === 'string' && periodType.trim() ? periodType.trim() : 'Custom';
  const normalizedType = type.toLowerCase();

  // Enforce Monday-Sunday alignment for weekly exports
  if (normalizedType === 'week' || normalizedType === 'weekly') {
    const weekStart = getMondayStart(normalizedStart);
    normalizedStart.setTime(weekStart.getTime());
    normalizedEnd.setTime(weekStart.getTime());
    normalizedEnd.setDate(normalizedEnd.getDate() + 6);
    normalizedEnd.setHours(23, 59, 59, 999);
  }

  if (normalizedType === 'biweekly' || normalizedType === 'bi-weekly') {
    const weekStart = getMondayStart(normalizedStart);
    normalizedStart.setTime(weekStart.getTime());
    normalizedEnd.setTime(weekStart.getTime());
    normalizedEnd.setDate(normalizedEnd.getDate() + 13);
    normalizedEnd.setHours(23, 59, 59, 999);
  }

  const pad = (num) => String(num).padStart(2, '0');
  const startIso = `${normalizedStart.getFullYear()}-${pad(normalizedStart.getMonth() + 1)}-${pad(normalizedStart.getDate())}`;
  const endIso = `${normalizedEnd.getFullYear()}-${pad(normalizedEnd.getMonth() + 1)}-${pad(normalizedEnd.getDate())}`;
  const label = (() => {
    switch (normalizedType) {
      case 'week':
      case 'weekly':
        return `Week of ${startIso}`;
      case 'biweekly':
      case 'bi-weekly':
        return `Bi-Weekly starting ${startIso}`;
      case 'month':
      case 'monthly':
        return `${safeStart.toLocaleDateString('en-US', { month: 'long', year: 'numeric' })}`;
      case 'quarter':
      case 'quarterly':
        return `Quarter of ${safeStart.getFullYear()}`;
      case 'year':
      case 'yearly':
        return `${safeStart.getFullYear()}`;
      default:
        return `${startIso} to ${endIso}`;
    }
  })();

  return { type, start: normalizedStart, end: normalizedEnd, startIso, endIso, label };
}

function ensureAttendanceExportSheet(name) {
  const fileName = `${name} - ${new Date().toISOString().replace(/[:]/g, '-')}`;
  const spreadsheet = SpreadsheetApp.create(fileName);
  const sheet = spreadsheet.getSheets()[0];
  sheet.clear();
  return { spreadsheet, sheet };
}

function applyDashboardExportFormatting(sheet, headers, rowCount) {
  const totalRows = Math.max(1, rowCount + 1);
  sheet.setFrozenRows(1);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  const styleHeader = () => headerRange
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#1f4e79')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  styleHeader();

  if (rowCount > 0) {
    const dataRange = sheet.getRange(2, 1, rowCount, headers.length);
    dataRange.setHorizontalAlignment('center').setVerticalAlignment('middle');

    sheet.getRange(2, 1, rowCount, 2).setBackground('#f4f5f7'); // Agent + Period
    sheet.getRange(2, 3, rowCount, 7).setBackground('#e8f4fd'); // Totals section
    sheet.getRange(2, 10, rowCount, 1).setBackground('#f3f6fa'); // On-Time count
    sheet.getRange(2, 11, rowCount, headers.length - 10).setBackground('#fdf7e3'); // Percentages

    sheet.getRange(1, 1, totalRows, headers.length)
      .applyRowBanding(SpreadsheetApp.BandingTheme.BLUE, true, false);

    styleHeader();
  }

  sheet.setRowHeight(1, 30);
  sheet.setColumnWidths(1, 1, 180); // Agent
  sheet.setColumnWidths(2, 1, 180); // Period
  sheet.setColumnWidths(3, 7, 110); // Totals section
  sheet.setColumnWidths(10, headers.length - 9, 120); // Percentages

  if (rowCount > 0) {
    const rules = sheet.getConditionalFormatRules() || [];
    const addRule = (builder) => rules.push(builder.build());

    const negativeCols = [11, 12]; // Absent %, Late %
    negativeCols.forEach(col => {
      const range = sheet.getRange(2, col, rowCount, 1);
      addRule(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(15)
        .setBackground('#f8d7da')
        .setFontColor('#6b0000')
        .setRanges([range]));
      addRule(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(5, 15)
        .setBackground('#fff4ce')
        .setFontColor('#8a6d00')
        .setRanges([range]));
      addRule(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThanOrEqualTo(5)
        .setBackground('#d9ead3')
        .setFontColor('#114b00')
        .setRanges([range]));
    });

    const positiveCols = [13, 16]; // On-Time %, Attendance Score %
    positiveCols.forEach(col => {
      const range = sheet.getRange(2, col, rowCount, 1);
      addRule(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(90)
        .setBackground('#d9ead3')
        .setFontColor('#114b00')
        .setRanges([range]));
      addRule(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(70, 90)
        .setBackground('#fff4ce')
        .setFontColor('#8a6d00')
        .setRanges([range]));
      addRule(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(70)
        .setBackground('#f8d7da')
        .setFontColor('#6b0000')
        .setRanges([range]));
    });

    const neutralCols = [14, 15]; // Sick %, Vacation %
    neutralCols.forEach(col => {
      const range = sheet.getRange(2, col, rowCount, 1);
      addRule(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0)
        .setBackground('#e8f5e9')
        .setFontColor('#1b5e20')
        .setRanges([range]));
    });

    sheet.setConditionalFormatRules(rules);
  }
}

function applyCalendarExportFormatting(sheet, headers, rowCount, dayColumnStart) {
  const totalRows = Math.max(1, rowCount + 1);
  sheet.setFrozenRows(1);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  const styleHeader = () => headerRange
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#264653')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  styleHeader();

  if (rowCount > 0) {
    const dataRange = sheet.getRange(2, 1, rowCount, headers.length);
    dataRange.setHorizontalAlignment('center').setVerticalAlignment('middle');

    sheet.getRange(2, 1, rowCount, 2).setBackground('#f4f5f7'); // Agent + Period
    const percentStartIndex = headers.findIndex(h => h.includes('%')) + 1;
    const summaryCountStart = 3;
    const summaryCountCols = Math.max(0, (percentStartIndex || headers.length) - summaryCountStart);
    const percentageCols = Math.max(0, dayColumnStart - percentStartIndex);

    if (summaryCountCols > 0) {
      sheet.getRange(2, summaryCountStart, rowCount, summaryCountCols).setBackground('#e8f4fd');
    }
    if (percentStartIndex > 0 && percentageCols > 0) {
      sheet.getRange(2, percentStartIndex, rowCount, percentageCols).setBackground('#fdf7e3');
    }
    if (headers.length >= dayColumnStart) {
      const dayCols = headers.length - dayColumnStart + 1;
      sheet.getRange(2, dayColumnStart, rowCount, dayCols).setBackground('#f7f7f7');
    }

    sheet.getRange(1, 1, totalRows, headers.length)
      .applyRowBanding(SpreadsheetApp.BandingTheme.CYAN, true, false);

    styleHeader();
  }

  sheet.setRowHeight(1, 30);
  sheet.setColumnWidths(1, 1, 170); // Agent
  sheet.setColumnWidths(2, 1, 170); // Period
  const summaryColumnCount = Math.max(1, dayColumnStart - 3);
  sheet.setColumnWidths(3, summaryColumnCount, 110); // Summary + percentages
  if (headers.length >= dayColumnStart) {
    const dayCols = headers.length - dayColumnStart + 1;
    sheet.setColumnWidths(dayColumnStart, dayCols, 75);
  }

  if (rowCount > 0) {
    const rules = sheet.getConditionalFormatRules() || [];
    const addRule = (builder) => rules.push(builder.build());

    const percentageCols = headers
      .map((header, index) => ({ header, index: index + 1 }))
      .filter(entry => entry.header.includes('%'))
      .map(entry => entry.index)
      .filter(index => index < dayColumnStart);

    percentageCols.forEach(col => {
      const range = sheet.getRange(2, col, rowCount, 1);
      addRule(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(90)
        .setBackground('#d9ead3')
        .setFontColor('#114b00')
        .setRanges([range]));
      addRule(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(70, 90)
        .setBackground('#fff4ce')
        .setFontColor('#8a6d00')
        .setRanges([range]));
      addRule(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(70)
        .setBackground('#f8d7da')
        .setFontColor('#6b0000')
        .setRanges([range]));
    });

    if (headers.length >= dayColumnStart) {
      const dayCols = headers.length - dayColumnStart + 1;
      const dayRange = sheet.getRange(2, dayColumnStart, rowCount, dayCols);
      const statusColors = [
        { match: 'Present', bg: '#d9ead3', fg: '#114b00' },
        { match: 'Late', bg: '#fff4ce', fg: '#8a6d00' },
        { match: 'Absent', bg: '#f8d7da', fg: '#6b0000' },
        { match: 'No Call No Show', bg: '#f8d7da', fg: '#6b0000' },
        { match: 'Sick Leave', bg: '#e9d5ff', fg: '#4a148c' },
        { match: 'Vacation', bg: '#d0e7ff', fg: '#0b5394' },
        { match: 'Holiday', bg: '#e2e3e5', fg: '#343a40' }
      ];

      statusColors.forEach(({ match, bg, fg }) => {
        addRule(SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(match)
          .setBackground(bg)
          .setFontColor(fg)
          .setRanges([dayRange]));
      });
    }

    sheet.setConditionalFormatRules(rules);
  }
}

function buildUserDisplayNameMap() {
  const displayNameMap = new Map();

  try {
    const users = readSheet(USERS_SHEET) || [];
    users.forEach(user => {
      const userName = (user.UserName || user.Username || user.userName || '').toString().trim();
      if (!userName) {
        return;
      }

      const fullName = (user.FullName || user.fullName || '').toString().trim();
      displayNameMap.set(userName, fullName || userName);
    });
  } catch (error) {
    console.warn('Unable to build user display name map:', error && error.message ? error.message : error);
  }

  return displayNameMap;
}

function pad2(num) {
  return String(num).padStart(2, '0');
}

function isWeekdayDate(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return false;
  }

  const day = date.getDay();
  return day !== 0 && day !== 6;
}

function isWeekdayIsoDate(iso) {
  if (!iso) return false;
  const date = new Date(`${iso}T00:00:00Z`);
  if (isNaN(date.getTime())) return false;
  const day = date.getUTCDay();
  return day !== 0 && day !== 6;
}

function toIsoDateString(value) {
  const date = value instanceof Date ? value : new Date(value);
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return null;
  }

  const y = date.getUTCFullYear();
  const m = pad2(date.getUTCMonth() + 1);
  const d = pad2(date.getUTCDate());
  return `${y}-${m}-${d}`;
}

function getWeekdayIsoDatesInRange(start, end) {
  const dates = [];
  if (!(start instanceof Date) || !(end instanceof Date)) {
    return dates;
  }

  const startUtc = Date.UTC(start.getUTCFullYear(), start.getUTCMonth(), start.getUTCDate());
  const endUtc = Date.UTC(end.getUTCFullYear(), end.getUTCMonth(), end.getUTCDate());

  for (let ts = startUtc; ts <= endUtc; ts += 24 * 60 * 60 * 1000) {
    const cursor = new Date(ts);
    const iso = toIsoDateString(cursor);
    if (iso && isWeekdayIsoDate(iso)) {
      dates.push(iso);
    }
  }

  return dates;
}

function exportAttendanceDashboard(periodType, startDate, endDate) {
  try {
    const range = normalizeDateRangeForExport(periodType, startDate, endDate);
    const attendanceData = readScheduleSheet(ATTENDANCE_STATUS_SHEET) || [];
    const displayNameMap = buildUserDisplayNameMap();
    const workingDates = getWeekdayIsoDatesInRange(range.start, range.end);
    const totalWorkingDays = workingDates.length || 1;
    const positiveStatuses = new Set(['present', 'punctual', 'training']);
    const absenceStatuses = new Set([
      'absent',
      'no call no show',
      'no call/no show',
      'no call/no-show',
      'sick',
      'sick leave',
      'leave of absent',
      'leave of absence'
    ]);
    const normalizeStatus = (status) => (status || '').toString().trim().toLowerCase();

    const filtered = attendanceData.filter(record => {
      const iso = record && record.Date ? toIsoDateString(record.Date) : null;
      return iso && isWeekdayIsoDate(iso) && iso >= range.startIso && iso <= range.endIso;
    });

    const userMap = new Map();
    filtered.forEach(record => {
      const user = record.UserName || record.User || record.userName;
      if (user) {
        const fullName = (record.FullName || record.fullName || '').toString().trim();
        if (fullName && !displayNameMap.has(user)) {
          displayNameMap.set(user, fullName);
        }
        userMap.set(user, user);
      }
    });

    const headers = [
      'Agent', 'Period', 'Total Days', 'Present/Worked', 'Absent', 'Sick', 'Vacation', 'Holiday', 'Late', 'On-Time',
      'Absent %', 'Late %', 'On-Time %', 'Sick %', 'Vacation %', 'Attendance Score %'
    ];

    const rows = Array.from(userMap.keys()).sort().map(user => {
      const stats = filtered.filter(r => (r.UserName || r.User || r.userName) === user);
      const dailyStatuses = new Map();

      stats.forEach(record => {
        const iso = record && record.Date ? toIsoDateString(record.Date) : null;
        if (!iso || !workingDates.includes(iso)) return;
        dailyStatuses.set(iso, record);
      });

      const missingDays = Math.max(0, totalWorkingDays - dailyStatuses.size);
      const totals = {
        present: 0,
        late: 0,
        absent: missingDays,
        sick: 0,
        vacation: 0,
        holiday: 0
      };

      Array.from(dailyStatuses.values()).forEach(record => {
        const rawStatus = record.Status || record.status || '';
        const status = normalizeStatus(rawStatus);

        if (status === 'late') {
          totals.present += 1;
          totals.late += 1;
        } else if (positiveStatuses.has(status)) {
          totals.present += 1;
        } else if (absenceStatuses.has(status)) {
          totals.absent += 1;
          if (status === 'sick' || status === 'sick leave') {
            totals.sick += 1;
          }
        } else if (status === 'vacation') {
          totals.vacation += 1;
        } else if (status === 'holiday') {
          totals.holiday += 1;
        }
      });

      const totalDays = totalWorkingDays;
      const onTime = Math.max(0, totals.present - totals.late);
      const pct = (count) => totalDays > 0 ? Math.round((count / totalDays) * 10000) / 100 : 0;
      const baseAttendanceScore = pct(Math.max(0, totalDays - totals.absent));
      const latePenalty = Math.min(totals.late, 100);
      const attendanceScore = Math.max(0, Math.min(100, baseAttendanceScore - latePenalty));

      return [
        displayNameMap.get(user) || user,
        range.label,
        totalDays,
        totals.present,
        totals.absent,
        totals.sick,
        totals.vacation,
        totals.holiday,
        totals.late,
        onTime,
        pct(totals.absent),
        pct(totals.late),
        pct(onTime),
        pct(totals.sick),
        pct(totals.vacation),
        attendanceScore
      ];
    });

    const { spreadsheet, sheet } = ensureAttendanceExportSheet('Attendance Dashboard Summary');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    applyDashboardExportFormatting(sheet, headers, rows.length);

    return {
      success: true,
      fileId: spreadsheet.getId(),
      fileUrl: spreadsheet.getUrl(),
      fileName: spreadsheet.getName()
    };
  } catch (error) {
    console.error('Error exporting attendance dashboard summary:', error);
    safeWriteError('exportAttendanceDashboard', error);
    return { success: false, error: error.message };
  }
}

function exportAttendanceCalendar(periodType, startDate, endDate) {
  try {
    const range = normalizeDateRangeForExport(periodType, startDate, endDate);
    const attendanceData = readScheduleSheet(ATTENDANCE_STATUS_SHEET) || [];
    const displayNameMap = buildUserDisplayNameMap();
    const dates = getWeekdayIsoDatesInRange(range.start, range.end);
    const positiveStatuses = new Set(['present', 'punctual', 'training']);
    const absentStatuses = new Set([
      'absent',
      'no call no show',
      'no call/no show',
      'no call/no-show',
      'sick',
      'sick leave',
      'leave of absent',
      'leave of absence'
    ]);
    const normalizeStatus = (status) => (status || '').toString().trim().toLowerCase();

    const filtered = attendanceData.filter(record => {
      const iso = record && record.Date ? toIsoDateString(record.Date) : null;
      return iso && isWeekdayIsoDate(iso) && iso >= range.startIso && iso <= range.endIso;
    });

    const users = new Set();
    filtered.forEach(record => {
      const user = record.UserName || record.User || record.userName;
      if (user) {
        const fullName = (record.FullName || record.fullName || '').toString().trim();
        if (fullName && !displayNameMap.has(user)) {
          displayNameMap.set(user, fullName);
        }
        users.add(user);
      }
    });

    const headers = [
      'Agent', 'Period', 'Total Days', 'Present', 'Absent', 'Sick', 'Vacation', 'Holiday', 'Late',
      'Bereavement', 'Maternity Leave', 'Training',
      'Present %', 'Absent %', 'Sick %', 'Vacation %', 'Holiday %', 'Late %', 'Bereavement %', 'Maternity Leave %', 'Training %',
      'Attendance Rate %', 'Punctual Rate %', 'Attendance Score %',
      ...dates
    ];

    const rows = Array.from(users).sort().map(user => {
      const userRecords = filtered.filter(r => (r.UserName || r.User || r.userName) === user);
      const dayStatus = new Map();
      dates.forEach(dateKey => dayStatus.set(dateKey, 'Absent'));

      const normalizedRecords = userRecords
        .map(record => ({
          record,
          iso: record && record.Date ? toIsoDateString(record.Date) : null
        }))
        .filter(entry => entry.iso && dates.includes(entry.iso));

      const uniqueIsos = new Set(normalizedRecords.map(entry => entry.iso));
      const missingDays = Math.max(0, dates.length - uniqueIsos.size);
      const totals = {
        present: 0,
        absent: missingDays,
        sick: 0,
        vacation: 0,
        holiday: 0,
        late: 0,
        bereavement: 0,
        maternity: 0,
        training: 0
      };

      normalizedRecords.forEach(({ record, iso }) => {
        const rawStatus = record.Status || record.status || '';
        const status = normalizeStatus(rawStatus);
        dayStatus.set(iso, rawStatus);

        if (positiveStatuses.has(status)) {
          totals.present += 1;
          if (status === 'training') {
            totals.training += 1;
          }
        } else if (absentStatuses.has(status)) {
          totals.absent += 1;
          if (status === 'sick' || status === 'sick leave') {
            totals.sick += 1;
          }
        } else if (status === 'late') {
          totals.present += 1;
          totals.late += 1;
        } else if (status === 'vacation') {
          totals.vacation += 1;
        } else if (status === 'holiday') {
          totals.holiday += 1;
        } else if (status === 'bereavement') {
          totals.bereavement += 1;
        } else if (status === 'maternity leave') {
          totals.maternity += 1;
        }
      });

      const totalDays = dates.length || 1;
      const absentDays = totals.absent;
      const pct = (count) => totalDays > 0 ? Math.round((count / totalDays) * 10000) / 100 : 0;
      const calendarStatuses = dates.map(dateKey => dayStatus.get(dateKey) || 'Absent');
      const presentDays = Math.max(0, totalDays - absentDays);
      const attendanceRate = pct(presentDays);
      const punctualDays = Math.max(0, totals.present - totals.late);
      const attendanceScore = pct(Math.max(0, totalDays - (totals.late + totals.absent + totals.sick + totals.vacation)));

      return [
        displayNameMap.get(user) || user,
        range.label,
        totalDays,
        presentDays,
        totals.absent,
        totals.sick,
        totals.vacation,
        totals.holiday,
        totals.late,
        totals.bereavement,
        totals.maternity,
        totals.training,
        attendanceRate,
        pct(totals.absent),
        pct(totals.sick),
        pct(totals.vacation),
        pct(totals.holiday),
        pct(totals.late),
        pct(totals.bereavement),
        pct(totals.maternity),
        pct(totals.training),
        attendanceRate,
        pct(punctualDays),
        attendanceScore,
        ...calendarStatuses
      ];
    });

    const { spreadsheet, sheet } = ensureAttendanceExportSheet('Attendance Calendar Export');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    const dayColumnStart = headers.length - dates.length + 1;
    applyCalendarExportFormatting(sheet, headers, rows.length, dayColumnStart);

    return {
      success: true,
      fileId: spreadsheet.getId(),
      fileUrl: spreadsheet.getUrl(),
      fileName: spreadsheet.getName()
    };
  } catch (error) {
    console.error('Error exporting attendance calendar:', error);
    safeWriteError('exportAttendanceCalendar', error);
    return { success: false, error: error.message };
  }
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

    const scheduleConfig = (typeof SCHEDULE_SETTINGS === 'object' && SCHEDULE_SETTINGS)
      || (typeof getScheduleConfig === 'function' ? getScheduleConfig() : {});
    const supportedCountriesSource = Array.isArray(scheduleConfig.SUPPORTED_COUNTRIES) && scheduleConfig.SUPPORTED_COUNTRIES.length
      ? scheduleConfig.SUPPORTED_COUNTRIES
      : ['JM', 'US', 'DO', 'PH'];
    const supportedCountries = supportedCountriesSource
      .map(country => typeof country === 'string' ? country.trim().toUpperCase() : String(country || '').trim().toUpperCase())
      .filter(country => country);

    const primaryCountryCandidate = typeof scheduleConfig.PRIMARY_COUNTRY === 'string' && scheduleConfig.PRIMARY_COUNTRY.trim()
      ? scheduleConfig.PRIMARY_COUNTRY.trim().toUpperCase()
      : 'JM';
    const fallbackCountry = supportedCountries.includes(primaryCountryCandidate)
      ? primaryCountryCandidate
      : (supportedCountries[0] || 'JM');

    const requestedCountry = typeof countryCode === 'string'
      ? countryCode.trim().toUpperCase()
      : String(countryCode || '').trim().toUpperCase();
    const normalizedCountry = supportedCountries.includes(requestedCountry)
      ? requestedCountry
      : fallbackCountry;
    const fallbackApplied = !requestedCountry || normalizedCountry !== requestedCountry;

    const parsedYear = parseInt(String(year), 10);
    const normalizedYear = Number.isFinite(parsedYear) && parsedYear > 1900
      ? parsedYear
      : new Date().getFullYear();

    const holidays = getUpdatedHolidays(normalizedCountry, normalizedYear);
    const isPrimary = normalizedCountry === primaryCountryCandidate;

    const noteParts = [
      isPrimary
        ? 'Primary country (Jamaica) - takes precedence'
        : 'Secondary country'
    ];

    if (fallbackApplied) {
      if (requestedCountry) {
        noteParts.push(`Requested country ${requestedCountry} is not supported. Using ${normalizedCountry} instead.`);
      } else {
        noteParts.push(`No country provided. Defaulting to ${normalizedCountry}.`);
      }
    }

    return {
      success: true,
      holidays: holidays,
      country: normalizedCountry,
      requestedCountry: requestedCountry || '',
      supportedCountries: supportedCountries,
      primaryCountry: primaryCountryCandidate,
      fallbackApplied: fallbackApplied,
      year: normalizedYear,
      isPrimary: isPrimary,
      note: noteParts.join(' ')
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
    const actorInfo = resolveScheduleActor(request.createdBy || null);
    const actorLabel = actorInfo.label;
    const actorLookupKey = actorInfo.lookupKey || 'system';

    const scheduleTimeZone = getSafeScheduleTimeZone();
    const scheduleUsers = clientGetScheduleUsers(actorLookupKey, campaignId || null);
    const userKeyMap = new Map();
    const userIdMap = new Map();
    scheduleUsers.forEach(user => {
      if (!user) {
        return;
      }

      const idCandidates = [
        user.ID,
        user.Id,
        user.UserID,
        user.UserId,
        user.id,
        user.userId
      ];
      idCandidates.forEach(candidate => {
        const normalized = normalizeUserIdValue(candidate);
        if (normalized && !userIdMap.has(normalized)) {
          userIdMap.set(normalized, user);
        }
      });

      const keyCandidates = [
        user.UserName,
        user.Username,
        user.username,
        user.FullName,
        user.fullName,
        user.Email,
        user.email
      ];
      keyCandidates.forEach(candidate => {
        const key = normalizeUserKey(candidate);
        if (key && !userKeyMap.has(key)) {
          userKeyMap.set(key, user);
        }
      });
    });

    const slotConfig = normalizeSlotConfiguration(slot, slot.SlotConfiguration || {});
    const allowSwapOverride = Object.prototype.hasOwnProperty.call(request, 'allowSwaps') ? request.allowSwaps : undefined;
    const allowSwapForSlot = coerceScheduleBoolean(allowSwapOverride, slotConfig.advanced ? slotConfig.advanced.allowSwaps : true);
    const restHoursForSlot = coerceScheduleNumber(slotConfig.advanced ? slotConfig.advanced.restPeriod : undefined, 0, { min: 0 });
    const notificationLeadForSlot = coerceScheduleNumber(slotConfig.advanced ? slotConfig.advanced.notificationLead : undefined, 0, { min: 0 });
    const handoverMinutesForSlot = coerceScheduleNumber(slotConfig.advanced ? slotConfig.advanced.handoverTime : undefined, 0, { min: 0 });
    const overtimeEnabledForSlot = coerceScheduleBoolean(slotConfig.overtime ? slotConfig.overtime.enabled : undefined, false);
    const overtimeMaxDailyForSlot = coerceScheduleNumber(slotConfig.overtime ? slotConfig.overtime.maxDaily : undefined, 0, { min: 0 });
    const overtimeMinutesForSlot = overtimeEnabledForSlot ? Math.round(overtimeMaxDailyForSlot * 60) : '';

    const existingAssignments = readShiftAssignments()
      .map(normalizeAssignmentRecord)
      .filter(record => record && record.AssignmentId)
      .filter(record => (record.Status || '').toUpperCase() !== 'ARCHIVED');

    const conflicts = [];
    const createdAssignments = [];
    const failedUsers = [];
    const archivedAssignments = [];
    const now = new Date();
    const dateSeries = buildDateSeries(normalizedStart, normalizedEnd);
    const weekendPremiumEnabled = coerceScheduleBoolean(slotConfig.advanced ? slotConfig.advanced.weekendPremium : false, false);
    const considerHolidayPremium = coerceScheduleBoolean(slotConfig.advanced ? slotConfig.advanced.holidayPremium : true, true);
    const holidayMap = considerHolidayPremium ? loadHolidayMap(normalizedStart, normalizedEnd) : new Map();

    userEntries.forEach(entry => {
      if (!entry) {
        return;
      }
      const nameKey = normalizeUserKey(entry.UserName || entry.FullName || entry.name || entry);
      const idKey = normalizeUserIdValue(entry.ID || entry.id || entry.userId || entry);
      const user = userKeyMap.get(nameKey) || (idKey ? userIdMap.get(idKey) : null);
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
            row.UpdatedBy = actorLabel;
            return row;
          });
          archivedAssignments.push(conflict.AssignmentId);
        });
      }

      const dstAdjustments = buildDstAdjustmentsForSlot(slot, dateSeries, scheduleTimeZone);
      const breakConfig = buildAssignmentBreakConfig(slotConfig, dstAdjustments);
      const premiumSet = new Set();

      if (weekendPremiumEnabled && dateSeries.some(isWeekendDate)) {
        premiumSet.add('Weekend');
      }

      if (considerHolidayPremium && holidayMap.size) {
        const hasHoliday = dateSeries.some(date => {
          const entries = holidayMap.get(date) || [];
          return entries.some(entry => (entry.region || '').toLowerCase() === 'jamaica');
        });
        if (hasHoliday) {
          premiumSet.add('Holiday');
        }
      }

      if (overtimeEnabledForSlot) {
        premiumSet.add('Overtime');
      }

      const dstNotes = summarizeDstAdjustments(dstAdjustments);
      const assignmentNotes = [request.notes || '', dstNotes].filter(Boolean).join(' | ');

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
        AllowSwap: allowSwapForSlot,
        Premiums: Array.from(premiumSet).join(','),
        BreaksConfigJSON: JSON.stringify(breakConfig),
        OvertimeMinutes: overtimeMinutesForSlot || '',
        RestPeriodHours: restHoursForSlot || '',
        NotificationLeadHours: notificationLeadForSlot || '',
        HandoverMinutes: handoverMinutesForSlot || '',
        Notes: assignmentNotes,
        CreatedAt: now,
        CreatedBy: actorLabel,
        UpdatedAt: now,
        UpdatedBy: actorLabel
      });
    });

    if (!createdAssignments.length) {
      const reasonMessages = [];
      if (conflicts.length) {
        reasonMessages.push('Assignments were blocked by existing conflicts.');
      }
      if (failedUsers.length) {
        const uniqueReasons = Array.from(new Set(failedUsers.map(entry => entry.reason).filter(Boolean)));
        if (uniqueReasons.length) {
          reasonMessages.push(uniqueReasons.join(' '));
        } else {
          reasonMessages.push(`${failedUsers.length} user${failedUsers.length === 1 ? '' : 's'} could not be assigned to the slot.`);
        }
      }
      const failureMessage = reasonMessages.length
        ? `No assignments were created successfully. ${reasonMessages.join(' ')}`
        : 'No assignments were created successfully. No eligible users met the slot configuration for the selected date range.';
      return {
        success: false,
        error: failureMessage,
        conflicts,
        failed: failedUsers
      };
    }

    const writeResult = writeShiftAssignments(createdAssignments, actorLabel, request.notes || 'Manual assignment', 'PENDING');

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

function normalizeAttendanceDateValue(value, timeZone = getScheduleTimeZone()) {
  const zone = timeZone || Session.getScriptTimeZone() || 'UTC';

  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, zone, 'yyyy-MM-dd');
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    const parsed = new Date(value);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, zone, 'yyyy-MM-dd');
    }
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) {
      return '';
    }

    const isoMatch = trimmed.match(/^(\d{4}-\d{2}-\d{2})(?:[T\s].*)?$/);
    if (isoMatch && isoMatch[1]) {
      return isoMatch[1];
    }

    const parsed = new Date(trimmed);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, zone, 'yyyy-MM-dd');
    }
  }

  return '';
}

function clientMarkAttendanceStatus(userName, date, status, notes = '') {
  try {
    console.log('📝 Marking attendance status:', { userName, date, status, notes });

    const timeZone = getScheduleTimeZone();
    const normalizedDate = normalizeAttendanceDateValue(date, timeZone);

    if (!normalizedDate) {
      throw new Error('A valid date is required to update attendance status.');
    }

    // Use ScheduleUtilities to ensure proper sheet structure
    const sheet = ensureScheduleSheetWithHeaders(ATTENDANCE_STATUS_SHEET, ATTENDANCE_STATUS_HEADERS);

    // Check if entry already exists (normalize both sides to avoid duplicate rows)
    const existingData = readScheduleSheet(ATTENDANCE_STATUS_SHEET) || [];
    const existingEntry = existingData.find(entry => {
      const entryDate = normalizeAttendanceDateValue(entry.Date, timeZone);
      const entryUser = String(entry.UserName || '').trim();
      return entryUser === String(userName || '').trim() && entryDate === normalizedDate;
    });

    const now = new Date();

    if (existingEntry) {
      // Update existing entry
      const data = sheet.getDataRange().getValues();
      const headers = data[0];

      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === existingEntry.ID) {
          sheet.getRange(i + 1, headers.indexOf('Date') + 1).setValue(normalizedDate);
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
        Date: normalizedDate,
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
      message: `Attendance status updated to ${status} for ${userName} on ${normalizedDate}`
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

function clientBulkMarkAttendanceStatus(request) {
  try {
    const payload = request && typeof request === 'object' ? request : {};
    const status = (payload.status || '').toString().trim();
    if (!status) {
      throw new Error('A status value is required for bulk attendance updates.');
    }

    const entries = Array.isArray(payload.entries) ? payload.entries : [];
    if (!entries.length) {
      throw new Error('At least one participant/date combination is required.');
    }

    const noteProvided = Object.prototype.hasOwnProperty.call(payload, 'notes');
    const noteValue = noteProvided ? (payload.notes || '').toString().trim() : '';

    const sheet = ensureScheduleSheetWithHeaders(ATTENDANCE_STATUS_SHEET, ATTENDANCE_STATUS_HEADERS);
    const data = sheet.getDataRange().getValues();
    const headers = data.length ? data[0] : ATTENDANCE_STATUS_HEADERS.slice();

    const headerMap = {};
    headers.forEach((header, index) => {
      headerMap[header] = index + 1;
    });

    const userNameIndex = headerMap.UserName;
    const dateIndex = headerMap.Date;
    const statusIndex = headerMap.Status;

    if (!userNameIndex || !dateIndex || !statusIndex) {
      throw new Error('Attendance status sheet is missing required headers.');
    }

    const notesIndex = headerMap.Notes || null;
    const updatedAtIndex = headerMap.UpdatedAt || null;

    const existingRowMap = new Map();
    if (data.length > 1) {
      for (let row = 1; row < data.length; row++) {
        const rowUser = (data[row][userNameIndex - 1] || '').toString().trim();
        const rowDate = normalizeDateForSheet(data[row][dateIndex - 1], DEFAULT_SCHEDULE_TIME_ZONE);
        if (!rowUser || !rowDate) {
          continue;
        }
        const key = `${rowUser}::${rowDate}`;
        const existingStatus = (data[row][statusIndex - 1] || '').toString().trim();
        const existingNotes = notesIndex ? (data[row][notesIndex - 1] || '').toString().trim() : '';
        existingRowMap.set(key, {
          row: row + 1,
          status: existingStatus,
          notes: existingNotes
        });
      }
    }

    const normalizedEntries = [];
    const skipped = [];
    const seenKeys = new Set();
    let duplicateCount = 0;

    entries.forEach(entry => {
      const userCandidate = entry && (entry.userName || entry.UserName || entry.user || entry.User || '');
      const normalizedUser = (userCandidate || '').toString().trim();
      const dateCandidate = entry && (entry.date || entry.Date || entry.day || entry.Day || entry.timestamp || entry.Timestamp || '');
      const normalizedDate = normalizeDateForSheet(dateCandidate, DEFAULT_SCHEDULE_TIME_ZONE);

      if (!normalizedUser || !normalizedDate) {
        skipped.push({
          userName: normalizedUser || userCandidate || '',
          date: normalizedDate || dateCandidate || '',
          reason: 'Missing participant or date'
        });
        return;
      }

      const key = `${normalizedUser}::${normalizedDate}`;
      if (seenKeys.has(key)) {
        duplicateCount++;
        skipped.push({
          userName: normalizedUser,
          date: normalizedDate,
          reason: 'Duplicate entry ignored'
        });
        return;
      }

      seenKeys.add(key);
      normalizedEntries.push({
        userName: normalizedUser,
        date: normalizedDate
      });
    });

    if (!normalizedEntries.length) {
      throw new Error('No valid attendance entries were provided.');
    }

    const now = new Date();
    const actorEmail = Session.getActiveUser().getEmail();
    const newRows = [];
    const appliedEntries = [];
    let updatedCount = 0;
    let insertedCount = 0;
    let unchangedCount = 0;

    normalizedEntries.forEach(entry => {
      const key = `${entry.userName}::${entry.date}`;
      const existing = existingRowMap.get(key);

      if (existing) {
        const currentStatus = (existing.status || '').toString().trim();
        const currentNotes = (existing.notes || '').toString().trim();
        const shouldUpdateNotes = noteProvided && currentNotes !== noteValue;

        if (currentStatus === status && !shouldUpdateNotes) {
          unchangedCount++;
          appliedEntries.push(entry);
          return;
        }

        sheet.getRange(existing.row, statusIndex).setValue(status);
        if (noteProvided && notesIndex) {
          sheet.getRange(existing.row, notesIndex).setValue(noteValue);
        }
        if (updatedAtIndex) {
          sheet.getRange(existing.row, updatedAtIndex).setValue(now);
        }
        updatedCount++;
        appliedEntries.push(entry);
        return;
      }

      const newEntry = {
        ID: Utilities.getUuid(),
        UserID: getUserIdByName(entry.userName) || entry.userName,
        UserName: entry.userName,
        Date: entry.date,
        Status: status,
        Notes: noteProvided ? noteValue : '',
        MarkedBy: actorEmail,
        CreatedAt: now,
        UpdatedAt: now
      };

      newRows.push(newEntry);
      appliedEntries.push(entry);
      insertedCount++;
    });

    if (newRows.length) {
      const startRow = sheet.getLastRow() + 1;
      const requiredRows = startRow + newRows.length - 1;
      const maxRows = sheet.getMaxRows();
      if (requiredRows > maxRows) {
        sheet.insertRowsAfter(maxRows, requiredRows - maxRows);
      }

      const rowsData = newRows.map(entry => ATTENDANCE_STATUS_HEADERS.map(header => entry[header] || ''));
      sheet.getRange(startRow, 1, rowsData.length, ATTENDANCE_STATUS_HEADERS.length).setValues(rowsData);
    }

    SpreadsheetApp.flush();
    invalidateScheduleCaches();

    const appliedCount = appliedEntries.length;
    const summaryParts = [];
    if (updatedCount) {
      summaryParts.push(`${updatedCount} updated`);
    }
    if (insertedCount) {
      summaryParts.push(`${insertedCount} new`);
    }
    if (unchangedCount) {
      summaryParts.push(`${unchangedCount} unchanged`);
    }
    if (skipped.length) {
      summaryParts.push(`${skipped.length} skipped`);
    }
    if (duplicateCount) {
      summaryParts.push(`${duplicateCount} duplicate${duplicateCount === 1 ? '' : 's'}`);
    }

    const messageBase = `Applied ${status} to ${appliedCount} record${appliedCount === 1 ? '' : 's'}`;
    const message = summaryParts.length ? `${messageBase} (${summaryParts.join(', ')}).` : `${messageBase}.`;

    return {
      success: true,
      status,
      applied: appliedEntries,
      updatedCount,
      insertedCount,
      unchangedCount,
      skipped,
      duplicates: duplicateCount,
      message
    };
  } catch (error) {
    console.error('Error applying bulk attendance status:', error);
    safeWriteError('clientBulkMarkAttendanceStatus', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function clientBulkClearAttendanceStatus(request) {
  try {
    const payload = request && typeof request === 'object' ? request : {};
    const entries = Array.isArray(payload.entries) ? payload.entries : [];

    if (!entries.length) {
      throw new Error('At least one participant/date combination is required to clear attendance statuses.');
    }

    const sheet = ensureScheduleSheetWithHeaders(ATTENDANCE_STATUS_SHEET, ATTENDANCE_STATUS_HEADERS);
    const data = sheet.getDataRange().getValues();
    const headers = data.length ? data[0] : ATTENDANCE_STATUS_HEADERS.slice();

    const headerMap = {};
    headers.forEach((header, index) => {
      headerMap[header] = index + 1;
    });

    const userNameIndex = headerMap.UserName;
    const dateIndex = headerMap.Date;
    const statusIndex = headerMap.Status || null;

    if (!userNameIndex || !dateIndex) {
      throw new Error('Attendance status sheet is missing required headers.');
    }

    const normalizedEntries = [];
    const skipped = [];
    const seenKeys = new Set();
    let duplicateCount = 0;

    entries.forEach(entry => {
      const userCandidate = entry && (entry.userName || entry.UserName || entry.user || entry.User || '');
      const dateCandidate = entry && (entry.date || entry.Date || entry.day || entry.Day || entry.timestamp || entry.Timestamp || '');
      const normalizedUser = (userCandidate || '').toString().trim();
      const normalizedDate = normalizeDateForSheet(dateCandidate, DEFAULT_SCHEDULE_TIME_ZONE);

      if (!normalizedUser || !normalizedDate) {
        skipped.push({
          userName: normalizedUser || userCandidate || '',
          date: normalizedDate || dateCandidate || '',
          reason: 'Missing participant or date'
        });
        return;
      }

      const key = `${normalizedUser}::${normalizedDate}`;
      if (seenKeys.has(key)) {
        duplicateCount++;
        return;
      }

      seenKeys.add(key);
      normalizedEntries.push({
        userName: normalizedUser,
        date: normalizedDate
      });
    });

    if (!normalizedEntries.length) {
      return {
        success: true,
        cleared: [],
        missing: [],
        skipped,
        duplicates: duplicateCount,
        message: 'No valid attendance entries were provided to clear.'
      };
    }

    if (data.length <= 1) {
      return {
        success: true,
        cleared: [],
        missing: normalizedEntries,
        skipped,
        duplicates: duplicateCount,
        message: 'No attendance statuses exist to clear.'
      };
    }

    const targets = new Set(normalizedEntries.map(entry => `${entry.userName}::${entry.date}`));
    const rowsToDelete = [];
    const clearedEntries = [];

    for (let row = 1; row < data.length; row++) {
      const rowUser = (data[row][userNameIndex - 1] || '').toString().trim();
      const rowDate = normalizeDateForSheet(data[row][dateIndex - 1], DEFAULT_SCHEDULE_TIME_ZONE);

      if (!rowUser || !rowDate) {
        continue;
      }

      const key = `${rowUser}::${rowDate}`;
      if (!targets.has(key)) {
        continue;
      }

      const rowStatus = statusIndex ? (data[row][statusIndex - 1] || '').toString().trim() : '';
      rowsToDelete.push({
        rowNumber: row + 1,
        key,
        userName: rowUser,
        date: rowDate,
        status: rowStatus
      });
      clearedEntries.push({
        userName: rowUser,
        date: rowDate,
        status: rowStatus
      });
    }

    if (!rowsToDelete.length) {
      return {
        success: true,
        cleared: [],
        missing: normalizedEntries,
        skipped,
        duplicates: duplicateCount,
        message: 'No matching attendance statuses were found to clear.'
      };
    }

    rowsToDelete.sort((a, b) => b.rowNumber - a.rowNumber);
    rowsToDelete.forEach(item => sheet.deleteRow(item.rowNumber));

    SpreadsheetApp.flush();
    invalidateScheduleCaches();

    const foundKeys = new Set(rowsToDelete.map(item => item.key));
    const missing = normalizedEntries.filter(entry => !foundKeys.has(`${entry.userName}::${entry.date}`));

    const summaryParts = [];
    if (missing.length) {
      summaryParts.push(`${missing.length} without status`);
    }
    if (skipped.length) {
      summaryParts.push(`${skipped.length} skipped`);
    }
    if (duplicateCount) {
      summaryParts.push(`${duplicateCount} duplicate${duplicateCount === 1 ? '' : 's'}`);
    }

    const clearedCount = clearedEntries.length;
    const messageBase = `Cleared ${clearedCount} attendance status${clearedCount === 1 ? '' : 'es'}`;
    const message = summaryParts.length ? `${messageBase} (${summaryParts.join(', ')}).` : `${messageBase}.`;

    return {
      success: true,
      cleared: clearedEntries,
      missing,
      skipped,
      duplicates: duplicateCount,
      message
    };
  } catch (error) {
    console.error('Error clearing bulk attendance statuses:', error);
    safeWriteError('clientBulkClearAttendanceStatus', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function clientRemoveAttendanceStatus(userName, date) {
  try {
    console.log('🧹 Clearing attendance status:', { userName, date });

      const sheet = ensureScheduleSheetWithHeaders(ATTENDANCE_STATUS_SHEET, ATTENDANCE_STATUS_HEADERS);
      const data = sheet.getDataRange().getValues();

      if (!data.length) {
        return {
          success: true,
          message: 'No attendance statuses found to clear.'
        };
      }

      const headers = data[0];
      const userNameIndex = headers.indexOf('UserName');
      const dateIndex = headers.indexOf('Date');

      if (userNameIndex === -1 || dateIndex === -1) {
        throw new Error('Attendance status sheet is missing required headers.');
      }

      const normalizedUser = String(userName || '').trim();
      const normalizedDate = normalizeAttendanceDateValue(date, getScheduleTimeZone());

      if (!normalizedUser || !normalizedDate) {
        throw new Error('Both userName and date are required to clear attendance status.');
      }

      let removed = false;

      for (let row = data.length - 1; row >= 1; row--) {
        const rowUser = String(data[row][userNameIndex] || '').trim();
        const rowDate = normalizeAttendanceDateValue(data[row][dateIndex], getScheduleTimeZone());

        if (rowUser === normalizedUser && rowDate === normalizedDate) {
          sheet.deleteRow(row + 1);
          removed = true;
          break;
        }
      }

      if (removed) {
        SpreadsheetApp.flush();
        invalidateScheduleCaches();
        return {
          success: true,
          message: `Attendance status cleared for ${normalizedUser} on ${normalizedDate}`
        };
      }

      return {
        success: true,
        message: 'No matching attendance status found to clear.'
      };

    } catch (error) {
      console.error('Error clearing attendance status:', error);
      safeWriteError('clientRemoveAttendanceStatus', error);
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
    const timeZone = getScheduleTimeZone();

    const normalizeDate = (value) => {
      const iso = normalizeAttendanceDateValue(value, timeZone);
      if (!iso) {
        return null;
      }
      const [year, month, day] = iso.split('-').map(Number);
      if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
        return null;
      }
      return new Date(Date.UTC(year, month - 1, day));
    };

    const normalizeTimestamp = (value) => {
      if (value instanceof Date && !isNaN(value.getTime())) {
        return value.getTime();
      }
      if (typeof value === 'number' && Number.isFinite(value)) {
        const parsed = new Date(value);
        return isNaN(parsed.getTime()) ? null : parsed.getTime();
      }
      if (typeof value === 'string' && value.trim()) {
        const parsed = new Date(value.trim());
        return isNaN(parsed.getTime()) ? null : parsed.getTime();
      }
      return null;
    };

    const rangeStart = normalizeDate(startDate);
    const rangeEnd = normalizeDate(endDate);

    const deduped = new Map();

    attendanceData.forEach(record => {
      const isoDate = normalizeAttendanceDateValue(record.Date || record.date, timeZone);
      if (!isoDate) {
        return;
      }

      const recordDate = normalizeDate(isoDate);
      if (!recordDate) {
        return;
      }

      if (rangeStart && recordDate < rangeStart) {
        return;
      }
      if (rangeEnd && recordDate > rangeEnd) {
        return;
      }

      if (campaignId) {
        const recordCampaign = record.CampaignID || record.CampaignId || record.Campaign || null;
        if (recordCampaign && recordCampaign !== campaignId) {
          return;
        }
      }

      const userName = String(record.UserName || record.User || record.user || '').trim();
      const status = (record.Status || record.status || record.state || '').toString();
      if (!userName || !status) {
        return;
      }

      const key = `${userName.toLowerCase()}__${isoDate}`;
      const existing = deduped.get(key);
      const timestamp = normalizeTimestamp(record.UpdatedAt || record.CreatedAt || record.Updated || record.Created) || 0;

      if (!existing || timestamp >= existing.timestamp) {
        deduped.set(key, {
          userName,
          status,
          date: isoDate,
          notes: record.Notes || record.notes || '',
          timestamp
        });
      }
    });

    const records = Array.from(deduped.values())
      .filter(entry => entry && entry.userName && entry.date && entry.status)
      .map(entry => ({
        userName: entry.userName,
        status: entry.status,
        date: entry.date,
        notes: entry.notes || ''
      }));

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

  const userLookup = buildScheduleUserLookupIndex();

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
      break;
    } catch (error) {
      response.warnings.push(`${attempt.name}: ${error && error.message ? error.message : error}`);
    }
  }

  const rosterIdSet = new Set(
    response.users
      .map(user => normalizeUserIdValue(user && (user.ID || user.UserID || user.id || user.userId)))
      .filter(Boolean)
  );

  const managedSet = buildManagedUserSet(normalizedManagerId);
  managedSet.forEach(id => {
    if (id) {
      rosterIdSet.add(id);
    }
  });

  const normalizedManagedIds = Array.from(managedSet)
    .map(id => normalizeUserIdValue(id))
    .filter(id => id && id !== normalizedManagerId);

  const hasVisibleRoster = Array.from(rosterIdSet).some(id => id && id !== normalizedManagerId);

  if (!hasVisibleRoster && normalizedManagedIds.length) {
    const fallback = collectCampaignUsersForManager(normalizedManagerId, { allUsers: userLookup.users });
    if (Array.isArray(fallback.users) && fallback.users.length) {
      const filteredFallbackUsers = fallback.users.filter(user => {
        if (!user || typeof user !== 'object') {
          return false;
        }

        const id = normalizeUserIdValue(user.ID || user.UserID || user.id || user.userId);
        if (!id) {
          return false;
        }

        return normalizedManagedIds.includes(id);
      });

      filteredFallbackUsers.forEach(user => {
        const id = normalizeUserIdValue(user && (user.ID || user.UserID));
        if (id) {
          rosterIdSet.add(id);
        }
      });

      if (filteredFallbackUsers.length) {
        response.source = response.source
          ? `${response.source}+campaignRosterFallback`
          : 'campaignRosterFallback';
        response.warnings.push('Managed roster did not return agents; using campaign roster fallback.');
        response.users = filteredFallbackUsers;
      }
    }
  }

  if (!response.users.length && rosterIdSet.size) {
    response.users = Array.from(rosterIdSet).map(id => ({ ID: id }));
  }

  const normalizedRoster = response.users
    .map(user => normalizeScheduleUserRecord(user, userLookup))
    .filter(Boolean);

  const dedupedRoster = new Map();
  normalizedRoster.forEach(user => {
    const id = normalizeUserIdValue(user && user.ID);
    if (id && !dedupedRoster.has(id)) {
      dedupedRoster.set(id, user);
    }
  });

  let filteredRoster = Array.from(dedupedRoster.values());
  if (normalizedManagedIds.length) {
    filteredRoster = filteredRoster.filter(user => {
      const id = normalizeUserIdValue(user && (user.ID || user.UserID || user.id || user.userId));
      return id && id !== normalizedManagerId && normalizedManagedIds.includes(id);
    });
  } else {
    filteredRoster = filteredRoster.filter(user => {
      const id = normalizeUserIdValue(user && (user.ID || user.UserID || user.id || user.userId));
      return id && id !== normalizedManagerId;
    });
  }

  response.users = filteredRoster;

  const managedIdsSource = normalizedManagedIds.length
    ? normalizedManagedIds
    : Array.from(rosterIdSet).map(id => normalizeUserIdValue(id)).filter(id => id && id !== normalizedManagerId);
  response.managedUserIds = Array.from(new Set(managedIdsSource));

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
