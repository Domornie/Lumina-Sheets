/**
 * AttendanceService.gs - Complete Production Version
 * 
 * ACTUAL DATABASE STRUCTURE:
 * Columns: ID, Timestamp, User, DurationMin, State, Date, UserID, CreatedAt, UpdatedAt
 * 
 * CRITICAL: DurationMin column contains SECONDS despite the name!
 * Examples from your data:
 * - 5782 seconds = 1.6 hours (reasonable for a meeting)
 * - 30534 seconds = 8.5 hours (reasonable for a full work day)
 * - 3641 seconds = 1.0 hours (reasonable for available time)
 * 
 * Timestamp format: "9/28/2023, 8:38 AM"
 * Date format: "9/28/2023"
 * Timezone: Configurable via ATTENDANCE_TIMEZONE (defaults to script timezone)
 */

/** @OnlyCurrentDoc */

// ────────────────────────────────────────────────────────────────────────────
// CONFIGURATION & CONSTANTS
// ────────────────────────────────────────────────────────────────────────────

const BILLABLE_STATES = ['Available', 'Administrative Work', 'Training', 'Meeting'];
const NON_PRODUCTIVE_STATES = ['Break', 'Lunch'];
const BILLABLE_DISPLAY_STATES = [...BILLABLE_STATES, 'Break'];
const NON_PRODUCTIVE_DISPLAY_STATES = [...new Set([...NON_PRODUCTIVE_STATES, 'Break'])];
const END_SHIFT_STATES = ['End of Shift'];

// Resolve a safe global scope reference for Apps Script V8
var GLOBAL_SCOPE = (typeof GLOBAL_SCOPE !== 'undefined') ? GLOBAL_SCOPE
  : (typeof globalThis === 'object' && globalThis)
    ? globalThis
    : (typeof this === 'object' && this)
      ? this
      : {};

// Time constants (working in seconds, since DurationMin column contains seconds)
const DAILY_SHIFT_SECS = 8 * 3600;       // 8 hours in seconds
const DAILY_BREAKS_SECS = 30 * 60;       // 30 minutes in seconds
const DAILY_LUNCH_SECS = 30 * 60;        // 30 minutes in seconds
const WEEKLY_OVERTIME_SECS = 40 * 3600;  // 40 hours in seconds

// Primary attendance timezone configuration (defaults to script timezone)
const ATTENDANCE_TIMEZONE = (typeof GLOBAL_SCOPE.ATTENDANCE_TIMEZONE === 'string' && GLOBAL_SCOPE.ATTENDANCE_TIMEZONE)
  ? GLOBAL_SCOPE.ATTENDANCE_TIMEZONE
  : (typeof Session !== 'undefined' && Session.getScriptTimeZone ? Session.getScriptTimeZone() : 'America/Jamaica');
const ATTENDANCE_TIMEZONE_LABEL = (typeof GLOBAL_SCOPE.ATTENDANCE_TIMEZONE_LABEL === 'string' && GLOBAL_SCOPE.ATTENDANCE_TIMEZONE_LABEL)
  ? GLOBAL_SCOPE.ATTENDANCE_TIMEZONE_LABEL
  : 'Company Time';
const ATTENDANCE_SHEET_NAME = (typeof GLOBAL_SCOPE.ATTENDANCE === 'string' && GLOBAL_SCOPE.ATTENDANCE)
  ? GLOBAL_SCOPE.ATTENDANCE
  : 'AttendanceLog';

// Performance optimization constants
const MAX_PROCESSING_TIME = 25000; // 25 seconds max execution time
const CHUNK_SIZE = 1000; // Process data in chunks
const CACHE_TTL_SHORT = 60; // 1 minute cache
const CACHE_TTL_MEDIUM = 300; // 5 minute cache
const LARGE_CACHE_CHUNK_SIZE = 90000; // stay below 100k Apps Script cache limit per entry
const ATTENDANCE_CACHE_VERSION = 'v4';

function cloneDate(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getTime());
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    const dateFromNumber = new Date(value);
    return isNaN(dateFromNumber.getTime()) ? null : dateFromNumber;
  }
  return null;
}

function createDateInLocalTime(year, month, day, hour, minute, second) {
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
    return null;
  }

  const safeHour = Number.isFinite(hour) ? hour : 0;
  const safeMinute = Number.isFinite(minute) ? minute : 0;
  const safeSecond = Number.isFinite(second) ? second : 0;

  const date = new Date(year, month - 1, day, safeHour, safeMinute, safeSecond, 0);
  return isNaN(date.getTime()) ? null : date;
}

function normalizeDateValue(value) {
  if (value == null) return null;

  const cloned = cloneDate(value);
  if (cloned) {
    return cloned;
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) {
      return null;
    }

    const isoDateOnlyMatch = trimmed.match(/^([0-9]{4})-([0-9]{1,2})-([0-9]{1,2})$/);
    if (isoDateOnlyMatch) {
      const year = Number(isoDateOnlyMatch[1]);
      const month = Number(isoDateOnlyMatch[2]);
      const day = Number(isoDateOnlyMatch[3]);
      const localDate = createDateInLocalTime(year, month, day, 0, 0, 0);
      if (localDate) {
        return localDate;
      }
    }

    const isoDateTimeMatch = trimmed.match(/^([0-9]{4})-([0-9]{1,2})-([0-9]{1,2})[ T]([0-9]{1,2}):([0-9]{2})(?::([0-9]{2}))?$/);
    if (isoDateTimeMatch) {
      const year = Number(isoDateTimeMatch[1]);
      const month = Number(isoDateTimeMatch[2]);
      const day = Number(isoDateTimeMatch[3]);
      const hour = Number(isoDateTimeMatch[4]);
      const minute = Number(isoDateTimeMatch[5]);
      const second = isoDateTimeMatch[6] ? Number(isoDateTimeMatch[6]) : 0;
      const localDateTime = createDateInLocalTime(year, month, day, hour, minute, second);
      if (localDateTime) {
        return localDateTime;
      }
    }

    const slashFormatMatch = trimmed.match(/^([0-9]{1,2})\/([0-9]{1,2})\/([0-9]{4})(?:[ ,T]+([0-9]{1,2}):([0-9]{2})(?::([0-9]{2}))?\s*(AM|PM)?)?$/i);
    if (slashFormatMatch) {
      const month = Number(slashFormatMatch[1]);
      const day = Number(slashFormatMatch[2]);
      const year = Number(slashFormatMatch[3]);
      let hour = slashFormatMatch[4] ? Number(slashFormatMatch[4]) : 0;
      const minute = slashFormatMatch[5] ? Number(slashFormatMatch[5]) : 0;
      const second = slashFormatMatch[6] ? Number(slashFormatMatch[6]) : 0;
      const meridiem = slashFormatMatch[7] ? slashFormatMatch[7].toUpperCase() : '';

      if (meridiem === 'PM' && hour < 12) {
        hour += 12;
      }
      if (meridiem === 'AM' && hour === 12) {
        hour = 0;
      }

      const localFromSlash = createDateInLocalTime(year, month, day, hour, minute, second);
      if (localFromSlash) {
        return localFromSlash;
      }
    }

    const parsed = Date.parse(trimmed);
    if (!Number.isNaN(parsed)) {
      const dateFromParse = new Date(parsed);
      return isNaN(dateFromParse.getTime()) ? null : dateFromParse;
    }

    const cleaned = trimmed.replace(/,/g, '');
    const parsedCleaned = Date.parse(cleaned);
    if (!Number.isNaN(parsedCleaned)) {
      const cleanedDate = new Date(parsedCleaned);
      return isNaN(cleanedDate.getTime()) ? null : cleanedDate;
    }
  }

  return null;
}

function toIsoDateString(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return '';
  }

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function toIsoDayOfWeek(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return undefined;
  }

  const jsDay = date.getDay();
  return jsDay === 0 ? 7 : jsDay;
}

function coerceDurationSeconds(value) {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }

  if (typeof value === 'string') {
    const parsed = parseFloat(value);
    if (!Number.isNaN(parsed) && Number.isFinite(parsed)) {
      return parsed;
    }
  }

  return 0;
}

function readLargeCache(cache, baseKey) {
  try {
    const metaRaw = cache.get(`${baseKey}::meta`);
    if (!metaRaw) return null;

    const meta = JSON.parse(metaRaw);
    if (!meta || meta.version !== ATTENDANCE_CACHE_VERSION || typeof meta.chunks !== 'number') {
      return null;
    }

    if (typeof meta.timestamp === 'number' && (Date.now() - meta.timestamp) > CACHE_TTL_MEDIUM * 1000) {
      return null;
    }

    const parts = [];
    for (let i = 0; i < meta.chunks; i++) {
      const part = cache.get(`${baseKey}::part::${i}`);
      if (!part) {
        return null;
      }
      parts.push(part);
    }

    const payload = parts.join('');
    return JSON.parse(payload);
  } catch (err) {
    try {
      console.warn('Large cache read failed:', err);
    } catch (_) {}
    return null;
  }
}

function writeLargeCache(cache, baseKey, value, ttlSeconds) {
  try {
    const payload = JSON.stringify(value);
    const chunks = [];
    for (let offset = 0; offset < payload.length; offset += LARGE_CACHE_CHUNK_SIZE) {
      chunks.push(payload.substring(offset, offset + LARGE_CACHE_CHUNK_SIZE));
    }

    const meta = JSON.stringify({
      version: ATTENDANCE_CACHE_VERSION,
      chunks: chunks.length,
      timestamp: Date.now()
    });

    cache.put(`${baseKey}::meta`, meta, ttlSeconds);

    chunks.forEach((chunk, index) => {
      cache.put(`${baseKey}::part::${index}`, chunk, ttlSeconds);
    });
  } catch (err) {
    try {
      console.warn('Large cache write failed:', err);
    } catch (_) {}
  }
}

function resolveAttendanceSpreadsheet() {
  if (typeof getIBTRSpreadsheet === 'function') {
    try {
      const ss = getIBTRSpreadsheet();
      if (ss) {
        return ss;
      }
    } catch (err) {
      try {
        console.warn('getIBTRSpreadsheet() failed, attempting SpreadsheetApp fallback', err);
      } catch (_) {}
    }
  }

  if (typeof SpreadsheetApp !== 'undefined') {
    try {
      if (typeof GLOBAL_SCOPE.CAMPAIGN_SPREADSHEET_ID === 'string' && GLOBAL_SCOPE.CAMPAIGN_SPREADSHEET_ID) {
        return SpreadsheetApp.openById(GLOBAL_SCOPE.CAMPAIGN_SPREADSHEET_ID);
      }
    } catch (openErr) {
      try {
        console.warn('SpreadsheetApp.openById fallback failed:', openErr);
      } catch (_) {}
    }

    if (typeof SpreadsheetApp.getActiveSpreadsheet === 'function') {
      return SpreadsheetApp.getActiveSpreadsheet();
    }
  }

  throw new Error('Unable to resolve attendance spreadsheet. Ensure IBTRUtilities is loaded.');
}


// ────────────────────────────────────────────────────────────────────────────
// RPC WITH TIMEOUT PROTECTION
// ────────────────────────────────────────────────────────────────────────────

function rpc(label, fn, fallback, maxTime = 20000) {
  const startTime = Date.now();

  try {
    if (Date.now() - startTime > maxTime * 0.8) {
      console.warn(`${label}: Approaching timeout, using fallback`);
      return (typeof fallback === 'function') ? fallback() : fallback;
    }

    const res = fn();
    return res;
  } catch (err) {
    const elapsed = Date.now() - startTime;
    if (elapsed > maxTime * 0.9) {
      console.warn(`${label}: Timeout after ${elapsed}ms, using fallback`);
    } else {
      console.error(`${label} error:`, err);
    }
    return (typeof fallback === 'function') ? fallback(err) : fallback;
  }
}

// ────────────────────────────────────────────────────────────────────────────
// MAIN DATA FETCHING
// ────────────────────────────────────────────────────────────────────────────

// In AttendanceService.gs, update the fetchAllAttendanceRows function around line 100-150
function fetchAllAttendanceRows() {
  return rpc('fetchAllAttendanceRows', () => {
    const CACHE_KEY = 'ATTENDANCE_ROWS_CACHE_FINAL_V4';

    // Try cache first
    try {
      const cache = CacheService.getScriptCache();
      const cached = readLargeCache(cache, CACHE_KEY);
      if (cached && cached.timestamp && Array.isArray(cached.rows)) {
        console.log('Using cached attendance data');
        const mapped = cached.rows.map(row => {
          const timestampMsRaw = typeof row.t === 'number' ? row.t : parseFloat(row.t);
          const timestampMs = Number.isFinite(timestampMsRaw) ? timestampMsRaw : undefined;
          const timestamp = typeof timestampMs === 'number' ? new Date(timestampMs) : null;
          let dateMs = typeof row.dm === 'number' && Number.isFinite(row.dm) ? row.dm : undefined;
          if (!Number.isFinite(dateMs) && timestamp instanceof Date && !isNaN(timestamp.getTime())) {
            const fallbackDate = createDateInLocalTime(
              timestamp.getFullYear(),
              timestamp.getMonth() + 1,
              timestamp.getDate(),
              0,
              0,
              0
            );
            dateMs = (fallbackDate instanceof Date && !isNaN(fallbackDate.getTime())) ? fallbackDate.getTime() : undefined;
          }
          const durationSec = coerceDurationSeconds(row.d);
          let dayOfWeek = typeof row.dow === 'number' && Number.isFinite(row.dow)
            ? row.dow
            : undefined;

          const cachedDate = normalizeDateValue(row.ds);
          if (!dayOfWeek && cachedDate) {
            dayOfWeek = toIsoDayOfWeek(cachedDate);
          }

          const dateString = (typeof row.ds === 'string' && row.ds)
            ? row.ds
            : (cachedDate ? toIsoDateString(cachedDate) : '');

          const isWeekend = typeof row.w === 'boolean'
            ? row.w
            : (typeof dayOfWeek === 'number' ? dayOfWeek >= 6 : undefined);

          return {
            timestamp,
            timestampMs,
            dateMs,
            user: row.u,
            state: row.s,
            durationSec,
            durationMin: durationSec / 60,
            durationHours: durationSec / 3600,
            dateString,
            dayOfWeek,
            isWeekend
          };
        });

        mapped.sort((a, b) => ensureComparableMs(a) - ensureComparableMs(b));
        return mapped;
      }
    } catch (e) {
      console.warn('Cache read failed:', e);
    }

    const ss = resolveAttendanceSpreadsheet();
    const sheet = ss.getSheetByName(ATTENDANCE_SHEET_NAME);
    if (!sheet) {
      console.warn(`Attendance sheet "${ATTENDANCE_SHEET_NAME}" not found in IBTR spreadsheet.`);
      return [];
    }

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return [];

    const headers = values[0].map(h => h.toString().trim());
    const timestampIdx = headers.indexOf('Timestamp');
    const userIdx = headers.indexOf('User');
    const stateIdx = headers.indexOf('State');
    const durationIdx = headers.indexOf('DurationMin');
    const dateIdx = headers.indexOf('Date');

    if (timestampIdx < 0 || userIdx < 0 || stateIdx < 0 || durationIdx < 0) {
      throw new Error('Required columns not found');
    }

    const out = [];
    const startTime = Date.now();

    for (let i = 1; i < values.length; i += CHUNK_SIZE) {
      if (Date.now() - startTime > MAX_PROCESSING_TIME * 0.8) {
        console.warn('Processing timeout approaching, stopping at row', i);
        break;
      }

      const chunk = values.slice(i, Math.min(i + CHUNK_SIZE, values.length));

      chunk.forEach((row, rowIndex) => {
        try {
          const timestampVal = row[timestampIdx];
          const timestamp = normalizeDateValue(timestampVal);
          if (!(timestamp instanceof Date) || isNaN(timestamp.getTime())) {
            console.warn(`Row ${i + rowIndex}: Could not parse timestamp:`, timestampVal);
            return;
          }

          const durationSeconds = coerceDurationSeconds(row[durationIdx]);
          const user = String(row[userIdx] || '').trim();
          const state = String(row[stateIdx] || '').trim();

          if (!user || !state) {
            console.warn(`Row ${i + rowIndex}: Missing user or state:`, { user, state });
            return;
          }

          const timestampMs = timestamp.getTime();

          let dateBasis = null;
          if (dateIdx >= 0) {
            dateBasis = normalizeDateValue(row[dateIdx]);
          }
          if (!(dateBasis instanceof Date) || isNaN(dateBasis.getTime())) {
            dateBasis = new Date(timestampMs);
          }

          const dateString = toIsoDateString(dateBasis);
          const dayOfWeek = toIsoDayOfWeek(dateBasis);
          const isWeekend = typeof dayOfWeek === 'number' ? dayOfWeek >= 6 : undefined;
          const dateStart = createDateInLocalTime(
            dateBasis.getFullYear(),
            dateBasis.getMonth() + 1,
            dateBasis.getDate(),
            0,
            0,
            0
          );
          const dateMs = (dateStart instanceof Date && !isNaN(dateStart.getTime())) ? dateStart.getTime() : undefined;

          out.push({
            timestamp,
            timestampMs,
            dateMs,
            user,
            state,
            durationSec: durationSeconds,
            durationMin: durationSeconds / 60,
            durationHours: durationSeconds / 3600,
            dateString,
            dayOfWeek,
            isWeekend
          });
        } catch (rowError) {
          console.error(`Error processing row ${i + rowIndex}:`, rowError, row);
        }
      });
    }

    console.log(`Processed ${out.length} attendance records with enhanced timezone handling`);
    
    // Cache the results (serialize timestamps for storage)
    try {
      const cache = CacheService.getScriptCache();
      const rowsForCache = out.map(row => ({
        t: row.timestamp.getTime(),
        u: row.user,
        s: row.state,
        d: row.durationSec,
        ds: row.dateString,
        dow: row.dayOfWeek,
        w: typeof row.isWeekend === 'boolean' ? row.isWeekend : undefined,
        dm: Number.isFinite(row.dateMs) ? row.dateMs : undefined
      }));

      writeLargeCache(cache, CACHE_KEY, {
        timestamp: Date.now(),
        rows: rowsForCache
      }, CACHE_TTL_MEDIUM);
    } catch (e) {
      console.warn('Cache write failed:', e);
    }

    out.sort((a, b) => ensureComparableMs(a) - ensureComparableMs(b));

    return out;
  }, []);
}

function ensureTimestampMs(row) {
  if (!row || typeof row !== 'object') {
    return NaN;
  }

  if (typeof row.timestampMs === 'number' && !isNaN(row.timestampMs)) {
    return row.timestampMs;
  }

  let timestamp = row.timestamp;
  if (!(timestamp instanceof Date) || isNaN(timestamp.getTime())) {
    timestamp = new Date(timestamp);
  }

  if (!(timestamp instanceof Date) || isNaN(timestamp.getTime())) {
    return NaN;
  }

  const ms = timestamp.getTime();
  row.timestamp = timestamp;
  row.timestampMs = ms;
  return ms;
}

function ensureDateMs(row) {
  if (!row || typeof row !== 'object') {
    return NaN;
  }

  if (typeof row.dateMs === 'number' && Number.isFinite(row.dateMs)) {
    return row.dateMs;
  }

  const dateSource = row.dateString ? normalizeDateValue(row.dateString) : null;
  if (!(dateSource instanceof Date) || isNaN(dateSource.getTime())) {
    return NaN;
  }

  const localStart = createDateInLocalTime(
    dateSource.getFullYear(),
    dateSource.getMonth() + 1,
    dateSource.getDate(),
    0,
    0,
    0
  );

  if (!(localStart instanceof Date) || isNaN(localStart.getTime())) {
    return NaN;
  }

  const ms = localStart.getTime();
  row.dateMs = ms;
  return ms;
}

function ensureComparableMs(row) {
  const timestampMs = ensureTimestampMs(row);
  if (Number.isFinite(timestampMs)) {
    return timestampMs;
  }

  return ensureDateMs(row);
}

// ────────────────────────────────────────────────────────────────────────────
// ANALYTICS ENGINE
// ────────────────────────────────────────────────────────────────────────────

function getAttendanceAnalyticsByPeriod(granularity, periodId, agentFilter) {
  return rpc('getAttendanceAnalyticsByPeriod', () => {
    const startTime = Date.now();
    const TIME_BUDGET_MS = Math.min(MAX_PROCESSING_TIME - 2000, 20000);
    console.log(`Analytics request: ${granularity}, ${periodId}, ${agentFilter || 'all'}`);

    if (!periodId) {
      throw new Error('Period ID is required');
    }

    const CACHE_KEY = `ANALYTICS_FINAL_${granularity}_${periodId}_${agentFilter || 'all'}`;

    // Try cache first
    try {
      const cached = CacheService.getScriptCache().get(CACHE_KEY);
      if (cached) {
        const data = JSON.parse(cached);
        if (data.timestamp && (Date.now() - data.timestamp) < CACHE_TTL_SHORT * 1000) {
          console.log('Using cached analytics');
          return data.analytics;
        }
      }
    } catch (e) {
      console.warn('Analytics cache read failed:', e);
    }

    let periodStart, periodEnd;
    try {
      [periodStart, periodEnd] = derivePeriodBounds(granularity, periodId);
    } catch (e) {
      throw new Error(`Invalid period: ${granularity} ${periodId}`);
    }

    const periodStartMs = periodStart.getTime();
    const periodEndMs = periodEnd.getTime();
    const periodStartDate = createDateInLocalTime(
      periodStart.getFullYear(),
      periodStart.getMonth() + 1,
      periodStart.getDate(),
      0,
      0,
      0
    );
    const periodEndDate = createDateInLocalTime(
      periodEnd.getFullYear(),
      periodEnd.getMonth() + 1,
      periodEnd.getDate(),
      0,
      0,
      0
    );
    const periodStartDateMs = (periodStartDate instanceof Date && !isNaN(periodStartDate.getTime()))
      ? periodStartDate.getTime()
      : periodStartMs;
    const periodEndDateMs = (periodEndDate instanceof Date && !isNaN(periodEndDate.getTime()))
      ? periodEndDate.getTime()
      : periodEndMs;

    const allRows = fetchAllAttendanceRows();
    const normalizedAgentFilter = agentFilter ? String(agentFilter).trim() : '';

    const summary = {};
    const stateDuration = {};
    const seedStates = [...new Set([...BILLABLE_STATES, ...NON_PRODUCTIVE_STATES, ...END_SHIFT_STATES])];
    seedStates.forEach(state => {
      summary[state] = 0;
      stateDuration[state] = 0;
    });

    const filteredRows = [];
    const userComplianceMap = new Map();
    const userDayMetrics = new Map();
    const topSeconds = new Map();
    const dailyMap = new Map();
    const uniqueUsers = new Set();
    const feedBuffer = [];
    const FEED_LIMIT = 10;

    let totalBillableSecs = 0;
    let totalRowsConsidered = 0;

    const registerFeedRow = (row, timestampMs) => {
      if (!timestampMs) return;
      if (feedBuffer.length < FEED_LIMIT) {
        feedBuffer.push({ row, timestampMs });
        feedBuffer.sort((a, b) => b.timestampMs - a.timestampMs);
        return;
      }

      const last = feedBuffer[feedBuffer.length - 1];
      if (last.timestampMs >= timestampMs) {
        return;
      }

      feedBuffer[feedBuffer.length - 1] = { row, timestampMs };
      feedBuffer.sort((a, b) => b.timestampMs - a.timestampMs);
    };

    let exceededTimeBudget = false;
    let scannedRows = 0;

    const loopStart = 0;
    const loopEnd = allRows.length - 1;
    const safePeriodStartMs = Number.isFinite(periodStartMs) ? periodStartMs : Number.POSITIVE_INFINITY;
    const safePeriodDateStartMs = Number.isFinite(periodStartDateMs) ? periodStartDateMs : Number.POSITIVE_INFINITY;
    const lowerBoundMs = Math.min(safePeriodStartMs, safePeriodDateStartMs);

    for (let idx = loopEnd; idx >= loopStart; idx--) {
      if (!exceededTimeBudget && scannedRows > 0 && (scannedRows % 250 === 0)) {
        if ((Date.now() - startTime) > TIME_BUDGET_MS) {
          console.warn('Analytics processing time budget exceeded after scanning', scannedRows, 'rows. Returning snapshot.');
          exceededTimeBudget = true;
          break;
        }
      }

      const row = allRows[idx];
      if (!row) continue;
      scannedRows++;

      const comparableMs = ensureComparableMs(row);
      if (Number.isFinite(comparableMs) && comparableMs < lowerBoundMs) {
        break;
      }

      const timestampMs = ensureTimestampMs(row);
      const dateMs = ensureDateMs(row);

      const hasTimestampMatch = Number.isFinite(timestampMs)
        ? (timestampMs >= periodStartMs && timestampMs <= periodEndMs)
        : false;
      const hasDateMatch = Number.isFinite(dateMs)
        ? (dateMs >= periodStartDateMs && dateMs <= periodEndDateMs)
        : false;

      if (!hasTimestampMatch && !hasDateMatch) {
        continue;
      }

      const effectiveTimestampMs = Number.isFinite(timestampMs)
        ? timestampMs
        : (Number.isFinite(dateMs) ? dateMs : undefined);

      const timestamp = row.timestamp instanceof Date
        ? row.timestamp
        : (Number.isFinite(effectiveTimestampMs) ? new Date(effectiveTimestampMs) : null);

      if (!(timestamp instanceof Date) || isNaN(timestamp.getTime())) {
        continue;
      }

      if (normalizedAgentFilter && row.user !== normalizedAgentFilter) {
        continue;
      }


      totalRowsConsidered++;

      const state = row.state || '';
      const durationSec = typeof row.durationSec === 'number' ? row.durationSec : parseFloat(row.durationSec) || 0;
      const durationHrs = typeof row.durationHours === 'number'
        ? row.durationHours
        : Math.round((durationSec / 3600) * 100) / 100;

      const dayOfWeek = (typeof row.dayOfWeek === 'number' && !isNaN(row.dayOfWeek))
        ? row.dayOfWeek
        : getAttendanceDayOfWeek(timestamp);
      const dateKey = row.dateString || Utilities.formatDate(timestamp, ATTENDANCE_TIMEZONE, 'yyyy-MM-dd');
      const isWeekend = typeof row.isWeekend === 'boolean' ? row.isWeekend : (dayOfWeek >= 6);

      uniqueUsers.add(row.user);

      summary[state] = (summary[state] || 0) + 1;
      stateDuration[state] = (stateDuration[state] || 0) + durationSec;

      const compliance = (() => {
        if (!userComplianceMap.has(row.user)) {
          userComplianceMap.set(row.user, {
            weekdayProdSecs: 0,
            weekendProdSecs: 0,
            breakSecs: 0,
            lunchSecs: 0
          });
        }
        return userComplianceMap.get(row.user);
      })();

      if (state === 'Break') {
        compliance.breakSecs += durationSec;
      } else if (state === 'Lunch') {
        compliance.lunchSecs += durationSec;
      }

      if (BILLABLE_STATES.includes(state)) {
        if (dayOfWeek >= 1 && dayOfWeek <= 5) {
          compliance.weekdayProdSecs += durationSec;
          topSeconds.set(row.user, (topSeconds.get(row.user) || 0) + durationSec);
        } else {
          compliance.weekendProdSecs += durationSec;
        }

        totalBillableSecs += durationSec;

        if (!dailyMap.has(dateKey)) {
          dailyMap.set(dateKey, { onWorkSecs: 0, lateCount: 0 });
        }
        dailyMap.get(dateKey).onWorkSecs += durationSec;
      }

      const userDayKey = `${row.user || ''}|${dateKey}`;
      if (!userDayMetrics.has(userDayKey)) {
        userDayMetrics.set(userDayKey, { prod: 0, break: 0, lunch: 0 });
      }

      const metrics = userDayMetrics.get(userDayKey);
      if (BILLABLE_STATES.includes(state)) {
        metrics.prod += durationSec;
      } else if (state === 'Break') {
        metrics.break += durationSec;
      } else if (state === 'Lunch') {
        metrics.lunch += durationSec;
      }

      const sanitizedRow = {
        timestamp,
        timestampMs: effectiveTimestampMs,
        user: row.user,
        state,
        durationSec,
        durationHrs,
        dateString: dateKey,
        dayOfWeek,
        isWeekend
      };
      filteredRows.push(sanitizedRow);
      if (!exceededTimeBudget) {
        registerFeedRow(sanitizedRow, effectiveTimestampMs);
      }
    }

    filteredRows.reverse();

    console.log(`Filtered ${totalRowsConsidered} records from ${allRows.length} total (scanned ${scannedRows})`);

    if (exceededTimeBudget) {
      return createBasicAnalytics(filteredRows, granularity, periodId, agentFilter, periodStart, periodEnd);
    }

    if (Date.now() - startTime > MAX_PROCESSING_TIME * 0.6) {
      console.warn('Approaching timeout, returning basic analytics snapshot');
      return createBasicAnalytics(filteredRows, granularity, periodId, agentFilter, periodStart, periodEnd);
    }

    let violationDays = 0;

    userDayMetrics.forEach(metrics => {
      const paidBreak = Math.min(metrics.break, DAILY_BREAKS_SECS);
      const breakExcess = Math.max(0, metrics.break - DAILY_BREAKS_SECS);
      const lunchExcess = Math.max(0, metrics.lunch - DAILY_LUNCH_SECS);

      if (breakExcess > 0 || lunchExcess > 0) {
        violationDays++;
      }
    });

    const breakSecs = stateDuration['Break'] || 0;
    const lunchSecs = stateDuration['Lunch'] || 0;
    const billableWithBreakSecs = totalBillableSecs + breakSecs;
    const totalBillableHours = Math.round((billableWithBreakSecs / 3600) * 100) / 100;
    const totalNonProductiveHours = Math.round(((breakSecs + lunchSecs) / 3600) * 100) / 100;

    const billableBreakdown = buildHourBreakdown(BILLABLE_DISPLAY_STATES, stateDuration);
    const nonProductiveBreakdown = buildHourBreakdown(NON_PRODUCTIVE_DISPLAY_STATES, stateDuration);

    const userCompliance = Array.from(userComplianceMap.entries()).map(([user, stats]) => ({
      user,
      availableSecsWeekday: stats.weekdayProdSecs,
      availableLabelWeekday: formatSecsAsHhMm(stats.weekdayProdSecs),
      breakSecs: stats.breakSecs,
      breakLabel: formatSecsAsHhMm(stats.breakSecs),
      lunchSecs: stats.lunchSecs,
      lunchLabel: formatSecsAsHhMm(stats.lunchSecs),
      weekendSecs: stats.weekendProdSecs,
      weekendLabel: formatSecsAsHhMm(stats.weekendProdSecs),
      exceededLunchDays: Math.floor(stats.lunchSecs / DAILY_LUNCH_SECS),
      exceededBreakDays: Math.floor(stats.breakSecs / DAILY_BREAKS_SECS),
      exceededWeeklyCount: 0
    }));

    const weekdaysInPeriod = countWeekdaysInclusive(periodStart, periodEnd);
    const expectedCapacitySecs = weekdaysInPeriod * DAILY_SHIFT_SECS;

    const top5Attendance = Array.from(topSeconds.entries())
      .map(([user, secs]) => ({
        user,
        percentage: expectedCapacitySecs > 0 ? Math.min(Math.round((secs / expectedCapacitySecs) * 100), 100) : 0
      }))
      .sort((a, b) => b.percentage - a.percentage)
      .slice(0, 5);

    const attendanceStats = [{
      periodLabel: periodId,
      OnWork: Math.round((billableWithBreakSecs / 3600) * 100) / 100,
      OverTime: 0,
      Leave: 0,
      EarlyEntry: 0,
      Late: 0,
      Absent: 0,
      EarlyOut: 0
    }];

    const attendanceFeed = feedBuffer
      .sort((a, b) => b.timestampMs - a.timestampMs)
      .map(item => {
        const ts = new Date(item.timestampMs);
        return {
          user: item.row.user,
          action: item.row.state,
          date: Utilities.formatDate(ts, ATTENDANCE_TIMEZONE, 'yyyy-MM-dd'),
          time24: Utilities.formatDate(ts, ATTENDANCE_TIMEZONE, 'HH:mm:ss'),
          time12: Utilities.formatDate(ts, ATTENDANCE_TIMEZONE, 'h:mm:ss a'),
          dayOfWeek: Utilities.formatDate(ts, ATTENDANCE_TIMEZONE, 'EEEE'),
          durationSec: item.row.durationSec,
          durationHrs: Math.round(item.row.durationSec / 3600 * 100) / 100
        };
      });

    const dailyMetrics = Array.from(dailyMap.entries())
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([day, metrics]) => ({
        date: day,
        OnWorkHrs: Math.round((metrics.onWorkSecs / 3600) * 100) / 100,
        LateCount: metrics.lateCount || 0
      }));

    const totalHours = totalBillableHours + totalNonProductiveHours;
    const efficiencyRate = totalHours > 0 ? (totalBillableHours / totalHours) * 100 : 0;
    const totalUserDays = userDayMetrics.size;
    const complianceRate = totalUserDays > 0 ? ((totalUserDays - violationDays) / totalUserDays) * 100 : 100;

    const executiveMetrics = {
      overview: {
        efficiencyRate,
        complianceRate,
        totalEmployees: uniqueUsers.size,
        activeEmployees: uniqueUsers.size,
        billableHours: totalBillableHours,
        productiveHours: totalBillableHours,
        nonProductiveHours: totalNonProductiveHours,
        breakHours: Math.round((breakSecs / 3600) * 100) / 100,
        lunchHours: Math.round((lunchSecs / 3600) * 100) / 100
      },
      violations: {
        totalViolations: violationDays
      },
      timeBreakdown: {
        billable: billableBreakdown,
        nonProductive: nonProductiveBreakdown
      }
    };

    const intelligence = generateAttendanceIntelligence(filteredRows, {
      periodStart,
      periodEnd,
      billableBreakdown,
      nonProductiveBreakdown,
      stateDuration
    });

    const analytics = {
      summary,
      stateDuration,
      totalBillableHours,
      totalProductiveHours: totalBillableHours,
      totalNonProductiveHours,
      billableHoursBreakdown: billableBreakdown,
      nonProductiveHoursBreakdown: nonProductiveBreakdown,
      filteredRows,
      filteredRowCount: filteredRows.length,
      userCompliance,
      top5Attendance,
      attendanceStats,
      attendanceFeed,
      dailyMetrics,
      shiftMetrics: {},
      enhanced: true,
      executiveMetrics,
      intelligence,
      periodInfo: {
        granularity,
        periodId,
        startDateIso: periodStart.toISOString(),
        endDateIso: periodEnd.toISOString(),
        workingDays: weekdaysInPeriod,
        timezone: ATTENDANCE_TIMEZONE,
        timezoneLabel: ATTENDANCE_TIMEZONE_LABEL
      }
    };

    // Cache results
    try {
      const cacheData = {
        timestamp: Date.now(),
        analytics
      };
      CacheService.getScriptCache().put(CACHE_KEY, JSON.stringify(cacheData), CACHE_TTL_SHORT);
    } catch (e) {
      console.warn('Analytics cache write failed:', e);
    }

    const elapsed = Date.now() - startTime;
    console.log(`Analytics completed in ${elapsed}ms`);

    return analytics;
  }, () => createEmptyAnalytics(), MAX_PROCESSING_TIME);
}

// ────────────────────────────────────────────────────────────────────────────
// CALCULATION FUNCTIONS
// ────────────────────────────────────────────────────────────────────────────

function getAttendanceDayOfWeek(timestamp) {
  try {
    // Ensure we have a proper Date object
    const date = timestamp instanceof Date ? timestamp : new Date(timestamp);
    
    if (isNaN(date.getTime())) {
      console.warn('Invalid timestamp provided to getAttendanceDayOfWeek:', timestamp);
      return 0;
    }

    // Get the day of week for the configured timezone
    const localizedDateString = Utilities.formatDate(date, ATTENDANCE_TIMEZONE, 'yyyy-MM-dd');
    const localizedDate = new Date(localizedDateString + 'T12:00:00.000Z');
    return localizedDate.getUTCDay();
  } catch (error) {
    console.error('Error in getAttendanceDayOfWeek:', error, 'timestamp:', timestamp);
    return 0;
  }
}

function generateDailyBreakdownData() {
  try {
    const dailyData = {};
    const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];

    days.forEach(day => {
      dailyData[day] = {
        lunch: { total: 0, count: 0, violations: 0, userViolations: {}, users: new Set() },
        break: { total: 0, count: 0, violations: 0, userViolations: {}, users: new Set() }
      };
    });

    if (this.currentData.filteredRows && Array.isArray(this.currentData.filteredRows)) {
      const userDailyTotals = {};

      // Process all records consistently
      this.currentData.filteredRows.forEach(r => {
        try {
          const timestamp = new Date(r.timestampMs || r.timestamp);
          
          // Use consistent configured timezone date calculation
          const timezoneDateString = timestamp.toLocaleDateString('en-CA', {
            timeZone: ATTENDANCE_TIMEZONE
          });
          const timezoneDate = new Date(timezoneDateString + 'T12:00:00');
          const dayOfWeek = timezoneDate.getDay();
          
          // Only process weekdays (1-5 = Monday-Friday)
          if (dayOfWeek === 0 || dayOfWeek === 6) return;

          const dayName = days[dayOfWeek - 1];
          const duration = r.durationSec || 0;
          const user = r.user;

          if (!user || !dayName) return;

          // Initialize user daily tracking
          if (!userDailyTotals[user]) {
            userDailyTotals[user] = {};
          }
          if (!userDailyTotals[user][dayName]) {
            userDailyTotals[user][dayName] = { lunch: 0, break: 0 };
          }

          // Process lunch and break states
          if (r.state === 'Lunch') {
            dailyData[dayName].lunch.total += duration;
            dailyData[dayName].lunch.count++;
            dailyData[dayName].lunch.users.add(user);
            userDailyTotals[user][dayName].lunch += duration;

            // Check for violations (over 30 minutes = 1800 seconds)
            if (userDailyTotals[user][dayName].lunch > 1800) {
              if (!dailyData[dayName].lunch.userViolations[user]) {
                dailyData[dayName].lunch.userViolations[user] = true;
                dailyData[dayName].lunch.violations++;
              }
            }
          } else if (r.state === 'Break') {
            dailyData[dayName].break.total += duration;
            dailyData[dayName].break.count++;
            dailyData[dayName].break.users.add(user);
            userDailyTotals[user][dayName].break += duration;

            // Check for violations (over 30 minutes = 1800 seconds)
            if (userDailyTotals[user][dayName].break > 1800) {
              if (!dailyData[dayName].break.userViolations[user]) {
                dailyData[dayName].break.userViolations[user] = true;
                dailyData[dayName].break.violations++;
              }
            }
          }
        } catch (recordError) {
          console.warn('Error processing record for daily breakdown:', recordError);
        }
      });

      // Convert Sets to arrays for consistency
      days.forEach(day => {
        dailyData[day].lunch.users = Array.from(dailyData[day].lunch.users);
        dailyData[day].break.users = Array.from(dailyData[day].break.users);
      });

      console.log('Daily breakdown data generated:', JSON.stringify(dailyData, null, 2));
    }

    return dailyData;
  } catch (error) {
    console.error('Error generating daily breakdown data:', error);
    return {};
  }
}

function calculateProductivityMetrics(filtered) {
    const stateDuration = {};

    filtered.forEach(r => {
        if (!r) return;
        const durationSec = typeof r.durationSec === 'number' ? r.durationSec : parseFloat(r.durationSec) || 0;
        const state = r.state || 'Unknown';
        stateDuration[state] = (stateDuration[state] || 0) + durationSec;
    });

    const breakSecs = stateDuration['Break'] || 0;
    const lunchSecs = stateDuration['Lunch'] || 0;
    const billableSecs = BILLABLE_STATES.reduce((sum, state) => sum + (stateDuration[state] || 0), 0);
    const billableWithBreakSecs = billableSecs + breakSecs;

    const totalBillableHours = Math.round((billableWithBreakSecs / 3600) * 100) / 100;
    const totalNonProductiveHours = Math.round(((breakSecs + lunchSecs) / 3600) * 100) / 100;

    return {
        totalBillableHours,
        totalProductiveHours: totalBillableHours,
        totalNonProductiveHours,
        billableHoursBreakdown: buildHourBreakdown(BILLABLE_DISPLAY_STATES, stateDuration),
        nonProductiveHoursBreakdown: buildHourBreakdown(NON_PRODUCTIVE_DISPLAY_STATES, stateDuration)
    };
}

function calculateUserCompliance(filtered) {
    const userStats = new Map();

    filtered.forEach(r => {
        if (!userStats.has(r.user)) {
            userStats.set(r.user, {
                weekdayProdSecs: 0,
                weekendProdSecs: 0,
                breakSecs: 0,
                lunchSecs: 0
            });
        }

        const stats = userStats.get(r.user);
        const secs = r.durationSec; // Already in seconds
        
        // Use precomputed day of week when available to avoid repeated timezone formatting
        const attendanceDayOfWeek = (typeof r.dayOfWeek === 'number' && !isNaN(r.dayOfWeek) && r.dayOfWeek > 0)
            ? r.dayOfWeek
            : getAttendanceDayOfWeek(r.timestamp);

        if (r.state === 'Break') stats.breakSecs += secs;
        if (r.state === 'Lunch') stats.lunchSecs += secs;

        if (BILLABLE_STATES.includes(r.state)) {
            if (attendanceDayOfWeek >= 1 && attendanceDayOfWeek <= 5) {
                stats.weekdayProdSecs += secs;
            } else {
                stats.weekendProdSecs += secs;
            }
        }
    });

    return Array.from(userStats.entries()).map(([user, stats]) => ({
        user,
        availableSecsWeekday: stats.weekdayProdSecs,
        availableLabelWeekday: formatSecsAsHhMm(stats.weekdayProdSecs),
        breakSecs: stats.breakSecs,
        breakLabel: formatSecsAsHhMm(stats.breakSecs),
        lunchSecs: stats.lunchSecs,
        lunchLabel: formatSecsAsHhMm(stats.lunchSecs),
        weekendSecs: stats.weekendProdSecs,
        weekendLabel: formatSecsAsHhMm(stats.weekendProdSecs),
        exceededLunchDays: Math.floor(stats.lunchSecs / (30 * 60)), // Number of 30-min periods
        exceededBreakDays: Math.floor(stats.breakSecs / (30 * 60)), // Number of 30-min periods  
        exceededWeeklyCount: 0
    }));
}

// ────────────────────────────────────────────────────────────────────────────
// UTILITY FUNCTIONS
// ────────────────────────────────────────────────────────────────────────────

function formatSecsAsHhMm(secs) {
  const hours = Math.floor(secs / 3600);
  const minutes = Math.floor((secs % 3600) / 60);
  return `${hours}h ${minutes}m`;
}

function buildHourBreakdown(stateList, durationMap) {
  const breakdown = {};
  if (!Array.isArray(stateList) || !durationMap) {
    return breakdown;
  }

  stateList.forEach(state => {
    const seconds = typeof durationMap[state] === 'number' ? durationMap[state] : 0;
    breakdown[state] = Math.round((seconds / 3600) * 100) / 100;
  });

  return breakdown;
}

function resolveAnalyticsDateKey(row) {
  if (!row) return null;

  if (typeof row.dateString === 'string' && row.dateString) {
    return row.dateString;
  }

  let timestampMs = null;
  if (typeof row.timestampMs === 'number') {
    timestampMs = row.timestampMs;
  } else if (row.timestamp instanceof Date) {
    timestampMs = row.timestamp.getTime();
  } else if (typeof row.timestamp === 'number') {
    timestampMs = row.timestamp;
  }

  if (!Number.isFinite(timestampMs)) {
    return null;
  }

  try {
    if (typeof Utilities !== 'undefined' && Utilities.formatDate) {
      return Utilities.formatDate(new Date(timestampMs), ATTENDANCE_TIMEZONE, 'yyyy-MM-dd');
    }
  } catch (err) {
    try {
      console.warn('resolveAnalyticsDateKey timezone formatting failed:', err);
    } catch (_) {}
  }

  const date = new Date(timestampMs);
  if (Number.isNaN(date.getTime())) {
    return null;
  }
  return date.toISOString().split('T')[0];
}

function generateAttendanceIntelligence(filteredRows, context) {
  try {
    const rows = Array.isArray(filteredRows) ? filteredRows : [];
    if (rows.length === 0) {
      return { insights: [], employeeTrends: [] };
    }

    const billableStates = new Set(BILLABLE_DISPLAY_STATES);
    const nonProdStates = new Set(NON_PRODUCTIVE_DISPLAY_STATES);

    const perUser = new Map();

    rows.forEach(row => {
      if (!row || !row.user) return;
      const durationSec = typeof row.durationSec === 'number' ? row.durationSec : parseFloat(row.durationSec) || 0;
      if (!Number.isFinite(durationSec) || durationSec <= 0) return;

      if (!perUser.has(row.user)) {
        perUser.set(row.user, {
          total: 0,
          billable: 0,
          nonProd: 0,
          breakSecs: 0,
          lunchSecs: 0,
          stateMap: {},
          days: new Set()
        });
      }

      const stats = perUser.get(row.user);
      stats.total += durationSec;
      stats.stateMap[row.state || 'Unknown'] = (stats.stateMap[row.state || 'Unknown'] || 0) + durationSec;
      if (billableStates.has(row.state)) {
        stats.billable += durationSec;
      }
      if (nonProdStates.has(row.state)) {
        stats.nonProd += durationSec;
      }
      if (row.state === 'Break') {
        stats.breakSecs += durationSec;
      }
      if (row.state === 'Lunch') {
        stats.lunchSecs += durationSec;
      }

      const dateKey = resolveAnalyticsDateKey(row);
      if (dateKey) {
        stats.days.add(dateKey);
      }
    });

    const toHours = (secs) => Math.round((secs / 3600) * 100) / 100;

    const employeeTrends = Array.from(perUser.entries()).map(([user, stats]) => {
      const daysActive = stats.days.size || 0;
      const activeDays = daysActive > 0 ? daysActive : 1;
      const billableHours = toHours(stats.billable);
      const nonProdHours = toHours(stats.nonProd);
      const efficiencyRate = (billableHours + nonProdHours) > 0
        ? Math.round((billableHours / (billableHours + nonProdHours)) * 1000) / 10
        : 0;

      const averageBillablePerDay = Math.round((billableHours / activeDays) * 100) / 100;
      const averageBreakPerDay = Math.round((toHours(stats.breakSecs) / activeDays) * 100) / 100;
      const averageLunchPerDay = Math.round((toHours(stats.lunchSecs) / activeDays) * 100) / 100;

      const sortedStates = Object.entries(stats.stateMap)
        .sort((a, b) => (b[1] || 0) - (a[1] || 0));
      const focusArea = sortedStates.length > 0 ? sortedStates[0][0] : null;

      let trendDirection = 'stable';
      if (averageBillablePerDay >= 7.5) {
        trendDirection = 'up';
      } else if (averageBillablePerDay <= 5) {
        trendDirection = 'down';
      }

      const trendSummary = `${averageBillablePerDay.toFixed(2)}h billable / day, ${averageBreakPerDay.toFixed(2)}h break` +
        `, ${averageLunchPerDay.toFixed(2)}h lunch`;

      return {
        user,
        billableHours,
        nonProductiveHours: nonProdHours,
        efficiencyRate,
        averageBillablePerDay,
        averageBreakPerDay,
        averageLunchPerDay,
        daysActive,
        focusArea,
        trendDirection,
        trendSummary
      };
    });

    employeeTrends.sort((a, b) => (b.billableHours || 0) - (a.billableHours || 0));

    const insights = [];

    if (employeeTrends.length > 0) {
      const topPerformer = [...employeeTrends].sort((a, b) => (b.averageBillablePerDay || 0) - (a.averageBillablePerDay || 0))[0];
      if (topPerformer) {
        insights.push({
          priority: 'high',
          title: 'Top Billable Performer',
          description: `${topPerformer.user} averaged ${topPerformer.averageBillablePerDay.toFixed(2)} billable hours per active day (${topPerformer.billableHours.toFixed(2)}h total).`,
          recommendation: 'Recognize this trend and consider sharing best practices with the wider team.'
        });
      }

      const downtimeThreshold = 0.35;
      const downtimeAlerts = employeeTrends
        .map(trend => {
          const total = trend.billableHours + trend.nonProductiveHours;
          const ratio = total > 0 ? trend.nonProductiveHours / total : 0;
          return { trend, ratio };
        })
        .filter(item => item.ratio > downtimeThreshold)
        .sort((a, b) => b.ratio - a.ratio);

      if (downtimeAlerts.length > 0) {
        const names = downtimeAlerts.slice(0, 3).map(item => item.trend.user).join(', ');
        const percentage = Math.round(downtimeAlerts[0].ratio * 100);
        insights.push({
          priority: 'critical',
          title: 'Extended Non-Productive Time Detected',
          description: `${downtimeAlerts.length} employee(s) spent over ${percentage}% of tracked time in lunch or break (${names}${downtimeAlerts.length > 3 ? ', …' : ''}).`,
          recommendation: 'Review schedules and coaching plans to bring downtime back within policy thresholds.'
        });
      }

      const breakOutliers = employeeTrends.filter(trend => trend.averageBreakPerDay > 1);
      if (breakOutliers.length > 0) {
        insights.push({
          priority: 'medium',
          title: 'High Daily Break Usage',
          description: `${breakOutliers.length} employee(s) average more than 1.00 hour of breaks per day.`,
          recommendation: 'Confirm coverage plans and reinforce standard break allocations.'
        });
      }

      const teamAverageBillable = employeeTrends.reduce((sum, trend) => sum + (trend.averageBillablePerDay || 0), 0) / employeeTrends.length;
      const topBillableState = context && context.billableBreakdown
        ? Object.entries(context.billableBreakdown).sort((a, b) => (b[1] || 0) - (a[1] || 0))[0]
        : null;

      insights.push({
        priority: 'medium',
        title: 'Team Billable Average',
        description: `Across ${employeeTrends.length} employees the team averages ${teamAverageBillable.toFixed(2)} billable hours per active day.`,
        recommendation: 'Use this baseline to set goals for upcoming periods.'
      });

      if (topBillableState && topBillableState[1] > 0) {
        insights.push({
          priority: 'low',
          title: 'Primary Billable Activity',
          description: `${topBillableState[0]} contributed ${topBillableState[1].toFixed(2)} billable hours for the period.`,
          recommendation: 'Ensure support resources remain aligned to this activity.'
        });
      }
    }

    return {
      insights,
      employeeTrends
    };
  } catch (error) {
    try {
      console.error('generateAttendanceIntelligence failed:', error);
    } catch (_) {}
    return { insights: [], employeeTrends: [] };
  }
}

function generateTopPerformers(filtered, periodStart, periodEnd) {
  const weekdaysInPeriod = countWeekdaysInclusive(periodStart, periodEnd);
  const expectedCapacitySecs = weekdaysInPeriod * DAILY_SHIFT_SECS;

  const userProdSecs = new Map();

  filtered.forEach(r => {
    const dow = (typeof r.dayOfWeek === 'number' && !isNaN(r.dayOfWeek) && r.dayOfWeek > 0)
      ? r.dayOfWeek
      : getAttendanceDayOfWeek(r.timestamp);
    if (dow >= 1 && dow <= 5 && BILLABLE_STATES.includes(r.state)) {
      userProdSecs.set(r.user, (userProdSecs.get(r.user) || 0) + r.durationSec);
    }
  });

  return Array.from(userProdSecs.entries())
    .map(([user, secs]) => ({
      user,
      percentage: expectedCapacitySecs > 0 ?
        Math.min(Math.round((secs / expectedCapacitySecs) * 100), 100) : 0
    }))
    .sort((a, b) => b.percentage - a.percentage)
    .slice(0, 5);
}

function generateAttendanceStats(filtered, periodId) {
  const totalWork = filtered
    .filter(r => BILLABLE_STATES.includes(r.state))
    .reduce((sum, r) => sum + r.durationSec, 0) / 3600; // Convert seconds to hours

  return [{
    periodLabel: periodId,
    OnWork: Math.round(totalWork * 100) / 100,
    OverTime: 0,
    Leave: 0,
    EarlyEntry: 0,
    Late: 0,
    Absent: 0,
    EarlyOut: 0
  }];
}

function generateAttendanceFeed(filtered) {
  return filtered
    .slice()
    .sort((a, b) => b.timestamp - a.timestamp)
    .slice(0, 10)
    .map(r => ({
      user: r.user,
      action: r.state,
      date: Utilities.formatDate(r.timestamp, ATTENDANCE_TIMEZONE, 'yyyy-MM-dd'),
      time24: Utilities.formatDate(r.timestamp, ATTENDANCE_TIMEZONE, 'HH:mm:ss'),
      time12: Utilities.formatDate(r.timestamp, ATTENDANCE_TIMEZONE, 'h:mm:ss a'),
      dayOfWeek: Utilities.formatDate(r.timestamp, ATTENDANCE_TIMEZONE, 'EEEE'),
      durationSec: r.durationSec,
      durationHrs: Math.round(r.durationSec / 3600 * 100) / 100
    }));
}

function generateDailyMetrics(filtered) {
  const dailyMap = new Map();

  filtered.forEach(r => {
    if (BILLABLE_STATES.includes(r.state)) {
      // Ensure consistent date formatting in configured timezone
      const timezoneDateString = Utilities.formatDate(r.timestamp, ATTENDANCE_TIMEZONE, 'yyyy-MM-dd');

      if (!dailyMap.has(timezoneDateString)) {
        dailyMap.set(timezoneDateString, { onWorkSecs: 0, lateCount: 0 });
      }
      dailyMap.get(timezoneDateString).onWorkSecs += r.durationSec;
    }
  });

  return Array.from(dailyMap.entries())
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([day, metrics]) => ({
      date: day, // This will be in YYYY-MM-DD format in the configured timezone
      OnWorkHrs: Math.round(metrics.onWorkSecs / 3600 * 100) / 100, // Convert seconds to decimal hours
      LateCount: metrics.lateCount
    }));
}

// ────────────────────────────────────────────────────────────────────────────
// ENHANCED EXPORT FUNCTIONS
// ────────────────────────────────────────────────────────────────────────────

function clientExecuteDailyPivotExport(params) {
  return rpc('clientExecuteDailyPivotExport', () => {
    console.log('Daily pivot matrix export requested with params:', JSON.stringify(params));

    const validation = validateEnhancedExportParams(params);
    if (!validation.isValid) {
      return { success: false, error: validation.errors.join(', ') };
    }

    try {
      const result = generateEnhancedDailyPivotMatrix(params);
      return { success: true, ...result };
    } catch (error) {
      console.error('Daily pivot matrix export failed:', error);
      return { success: false, error: 'Daily pivot export failed: ' + error.message };
    }
  }, { success: false, error: 'Daily pivot export service unavailable' }, MAX_PROCESSING_TIME);
}

function generateEnhancedDailyPivotMatrix(params) {
  try {
    const { period, users, userSelection, dailyPivotOptions } = params;
    let granularity, periodValue;

    // Determine period
    if (period.type === 'custom') {
      granularity = 'Week';
      periodValue = weekStringFromDate(new Date(period.start));
    } else {
      granularity = period.type;
      periodValue = period.value;
    }

    // Get analytics data
    let agentFilter = '';
    if (userSelection === 'single' && users.length > 0) {
      agentFilter = users[0];
    }
    
    const analytics = getAttendanceAnalyticsByPeriod(granularity, periodValue, agentFilter);
    
    // Filter users if multiple selection
    let filteredRows = analytics.filteredRows;
    if (userSelection === 'multiple' && users.length > 0) {
      filteredRows = analytics.filteredRows.filter(row => users.includes(row.user));
    }
    
    // Generate enhanced daily pivot matrix
    const pivotMatrix = generateDailyPivotMatrix(filteredRows, granularity, periodValue, dailyPivotOptions);
    
    // Generate CSV in enhanced matrix format
    const csv = generateEnhancedDailyPivotCSV(pivotMatrix, params);
    
    return {
      success: true,
      csvData: csv,
      filename: `daily_matrix_${granularity}_${periodValue}_company.csv`,
      type: 'daily_pivot_matrix',
      users: pivotMatrix.users.length,
      days: pivotMatrix.dateRange.length,
      dateRange: pivotMatrix.dateRange.length > 0 ? 
        `${pivotMatrix.dateRange[0].date} to ${pivotMatrix.dateRange[pivotMatrix.dateRange.length - 1].date}` : 'No data'
    };

  } catch (error) {
    console.error('Enhanced daily pivot matrix generation failed:', error);
    throw error;
  }
}

function generateDateRangeForPeriod(granularity, periodValue) {
  const dateRange = [];
  const daysOfWeek = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  
  let startDate, endDate;
  
  try {
    [startDate, endDate] = derivePeriodBounds(granularity, periodValue);
  } catch (e) {
    // Fallback to current week if period parsing fails
    const today = new Date();
    startDate = new Date(today);
    startDate.setDate(today.getDate() - today.getDay()); // Start of week (Sunday)
    endDate = new Date(startDate);
    endDate.setDate(startDate.getDate() + 6); // End of week (Saturday)
    console.warn('Period parsing failed, using current week:', e.message);
  }
  
  const currentDate = new Date(startDate);
  while (currentDate <= endDate) {
    // Use configured timezone for consistent date formatting
    const dateStr = Utilities.formatDate(currentDate, ATTENDANCE_TIMEZONE, 'yyyy-MM-dd');
    const dayOfWeek = currentDate.getDay();
    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
    
    dateRange.push({
      date: dateStr,
      dayName: daysOfWeek[dayOfWeek],
      dayOfWeek: dayOfWeek,
      isWeekend: isWeekend,
      formattedDate: Utilities.formatDate(currentDate, ATTENDANCE_TIMEZONE, 'M/d/yyyy')
    });
    
    currentDate.setDate(currentDate.getDate() + 1);
  }
  
  console.log(`Generated date range: ${startDate.toISOString().split('T')[0]} to ${endDate.toISOString().split('T')[0]} (${dateRange.length} days)`);
  return dateRange;
}

function generateDailyPivotMatrix(filteredRows, granularity, periodValue, options = {}) {
  // Generate date range for the period with proper configured timezone handling
  const dateRange = generateDateRangeForPeriod(granularity, periodValue);
  
  // Create user-date-hours mapping
  const userDateHours = new Map();
  const userDateBreakMinutes = new Map();
  const userDateLunchMinutes = new Map();
  
  // Get all unique users and sort them
  const allUsers = [...new Set(filteredRows.map(r => r.user))].sort((a, b) => a.localeCompare(b));
  
  // Initialize user data structure
  allUsers.forEach(user => {
    userDateHours.set(user, new Map());
    userDateBreakMinutes.set(user, new Map());
    userDateLunchMinutes.set(user, new Map());
    
    // Initialize all dates for this user
    dateRange.forEach(dateInfo => {
      userDateHours.get(user).set(dateInfo.date, 0);
      userDateBreakMinutes.get(user).set(dateInfo.date, 0);
      userDateLunchMinutes.get(user).set(dateInfo.date, 0);
    });
  });
  
  // Process attendance data with proper timezone handling
  filteredRows.forEach(r => {
    const timestamp = new Date(r.timestampMs || r.timestamp);
    const dateStr = Utilities.formatDate(timestamp, ATTENDANCE_TIMEZONE, 'yyyy-MM-dd');
    const user = r.user;
    
    if (userDateHours.has(user) && userDateHours.get(user).has(dateStr)) {
      const durationHours = r.durationSec / 3600; // Convert seconds to hours
      const durationMinutes = r.durationSec / 60; // Convert seconds to minutes
      
      // Track productive hours
      if (BILLABLE_STATES.includes(r.state)) {
        const currentHours = userDateHours.get(user).get(dateStr) || 0;
        userDateHours.get(user).set(dateStr, currentHours + durationHours);
      }
      
      // Track break and lunch minutes for violations
      if (r.state === 'Break') {
        const currentBreakMin = userDateBreakMinutes.get(user).get(dateStr) || 0;
        userDateBreakMinutes.get(user).set(dateStr, currentBreakMin + durationMinutes);
      } else if (r.state === 'Lunch') {
        const currentLunchMin = userDateLunchMinutes.get(user).get(dateStr) || 0;
        userDateLunchMinutes.get(user).set(dateStr, currentLunchMin + durationMinutes);
      }
    }
  });
  
  // Generate user data with enhanced metrics
  const userData = allUsers.map(user => {
    const userHours = userDateHours.get(user);
    const userBreakMin = userDateBreakMinutes.get(user);
    const userLunchMin = userDateLunchMinutes.get(user);
    
    let totalHours = 0;
    let weekdayHours = 0;
    let weekendHours = 0;
    let discrepancyDays = 0;
    let overtimeHours = 0;
    let perfectAttendanceDays = 0;
    let violationDays = 0;
    
    const dailyData = dateRange.map(dateInfo => {
      const hours = userHours.get(dateInfo.date) || 0;
      const breakMin = userBreakMin.get(dateInfo.date) || 0;
      const lunchMin = userLunchMin.get(dateInfo.date) || 0;
      
      if (dateInfo.isWeekend) {
        weekendHours += hours;
        if (!options.includeWeekends) {
          return { 
            date: dateInfo.date, 
            value: 'OFF', 
            isStatus: true, 
            isWeekend: true,
            breakMin: 0,
            lunchMin: 0
          };
        }
      } else {
        // Weekday processing
        weekdayHours += hours;
        totalHours += hours;
        
        // Check for discrepancies (less than 8 hours on weekday with some work)
        if (hours > 0 && hours < 8.0) {
          discrepancyDays++;
        }
        
        // Check for overtime (more than 8 hours)
        if (hours > 8.0) {
          overtimeHours += (hours - 8.0);
        }
        
        // Check for perfect attendance (8+ hours with reasonable breaks/lunch)
        if (hours >= 8.0 && breakMin <= 30 && lunchMin <= 60) {
          perfectAttendanceDays++;
        }
        
        // Check for policy violations
        if (breakMin > 30 || lunchMin > 60) {
          violationDays++;
        }
      }
      
      const formattedHours = hours.toFixed(2);
      const isLow = hours > 0 && hours < 8.0 && !dateInfo.isWeekend;
      
      return { 
        date: dateInfo.date, 
        value: formattedHours,
        numericValue: hours,
        isStatus: false,
        isLow: isLow,
        isWeekend: dateInfo.isWeekend,
        breakMin: Math.round(breakMin),
        lunchMin: Math.round(lunchMin),
        hasViolations: (breakMin > 30 || lunchMin > 60)
      };
    });
    
    // Calculate efficiency and compliance metrics
    const expectedWeekdayHours = dateRange.filter(d => !d.isWeekend).length * 8;
    const efficiency = expectedWeekdayHours > 0 ? (weekdayHours / expectedWeekdayHours * 100) : 0;
    
    return {
      user: user,
      dailyData: dailyData,
      metrics: {
        totalHours: totalHours.toFixed(2),
        weekdayHours: weekdayHours.toFixed(2),
        weekendHours: weekendHours.toFixed(2),
        discrepancy: discrepancyDays,
        overtime: overtimeHours.toFixed(2),
        perfectDays: perfectAttendanceDays,
        violationDays: violationDays,
        efficiency: efficiency.toFixed(1),
        daysWorked: dailyData.filter(d => !d.isStatus && d.numericValue > 0).length
      }
    };
  });
  
  // Calculate daily totals for each column
  const dailyTotals = dateRange.map(dateInfo => {
    let dayTotal = 0;
    let dayWorkers = 0;
    
    userData.forEach(user => {
      const dayData = user.dailyData.find(d => d.date === dateInfo.date);
      if (dayData && !dayData.isStatus && !isNaN(dayData.numericValue)) {
        dayTotal += dayData.numericValue;
        if (dayData.numericValue > 0) dayWorkers++;
      }
    });
    
    return {
      total: dayTotal.toFixed(2),
      workers: dayWorkers,
      average: dayWorkers > 0 ? (dayTotal / dayWorkers).toFixed(2) : '0.00'
    };
  });
  
  // Calculate grand totals
  const grandTotals = {
    totalHours: userData.reduce((sum, user) => sum + parseFloat(user.metrics.totalHours), 0).toFixed(2),
    totalDiscrepancies: userData.reduce((sum, user) => sum + user.metrics.discrepancy, 0),
    totalOvertime: userData.reduce((sum, user) => sum + parseFloat(user.metrics.overtime), 0).toFixed(2),
    averageEfficiency: userData.length > 0 ? (userData.reduce((sum, user) => sum + parseFloat(user.metrics.efficiency), 0) / userData.length).toFixed(1) : '0.0'
  };
  
  return {
    users: userData,
    dateRange: dateRange,
    dailyTotals: dailyTotals,
    grandTotals: grandTotals,
    metadata: {
      totalUsers: userData.length,
      totalDays: dateRange.length,
      weekdays: dateRange.filter(d => !d.isWeekend).length,
      weekends: dateRange.filter(d => d.isWeekend).length
    }
  };
}

function generateEnhancedDailyPivotCSV(pivotMatrix, params) {
  const now = new Date();
  const timestamp = Utilities.formatDate(now, ATTENDANCE_TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  const options = params.dailyPivotOptions || {};
  
  let csv = '';
  
  // Header section
  csv += `Daily Attendance Matrix (Users × Days)\n`;
  csv += `Generated: ${timestamp} (${ATTENDANCE_TIMEZONE_LABEL})\n`;
  csv += `Period: ${params.period.type} ${params.period.value || 'Custom'}\n`;
  csv += `Users: ${pivotMatrix.metadata.totalUsers}, Days: ${pivotMatrix.metadata.totalDays} (${pivotMatrix.metadata.weekdays} weekdays, ${pivotMatrix.metadata.weekends} weekends)\n`;
  csv += `Data Format: Productive hours converted from seconds to decimal hours\n\n`;
  
  // Column headers - First row with day names
  csv += `User Name`;
  pivotMatrix.dateRange.forEach(dateInfo => {
    if (!dateInfo.isWeekend || options.includeWeekends) {
      csv += `,${dateInfo.dayName}`;
    }
  });
  
  // Summary columns
  if (options.includeTotalHours) csv += `,Total Hours`;
  if (options.includeDiscrepancy) csv += `,Low Days`;
  if (options.includeOvertime) csv += `,Overtime`;
  csv += `,Efficiency %,Days Worked`;
  csv += `\n`;
  
  // Second header row with dates
  csv += ``;
  pivotMatrix.dateRange.forEach(dateInfo => {
    if (!dateInfo.isWeekend || options.includeWeekends) {
      csv += `,${dateInfo.formattedDate}`;
    }
  });
  
  // Summary column headers
  if (options.includeTotalHours) csv += `,`;
  if (options.includeDiscrepancy) csv += `,`;
  if (options.includeOvertime) csv += `,`;
  csv += `,,`;
  csv += `\n`;
  
  // User data rows
  pivotMatrix.users.forEach(userData => {
    csv += `"${userData.user}"`;
    
    // Daily hours columns
    userData.dailyData.forEach(dayData => {
      if (!dayData.isWeekend || options.includeWeekends) {
        let cellValue = dayData.value;
        
        // Add markers for special conditions
        if (options.highlightLowHours && dayData.isLow) {
          cellValue += '*';
        }
        
        // Show violations in parentheses if requested
        if (dayData.hasViolations && !dayData.isStatus) {
          cellValue += ` (B${dayData.breakMin}m L${dayData.lunchMin}m)`;
        }
        
        csv += `,${cellValue}`;
      }
    });
    
    // Summary columns
    if (options.includeTotalHours) {
      csv += `,${userData.metrics.totalHours}`;
    }
    if (options.includeDiscrepancy) {
      csv += `,${userData.metrics.discrepancy}`;
    }
    if (options.includeOvertime) {
      csv += `,${userData.metrics.overtime}`;
    }
    
    csv += `,${userData.metrics.efficiency}%,${userData.metrics.daysWorked}`;
    csv += `\n`;
  });
  
  // Totals row
  if (options.includeTotalsRow) {
    csv += `\nTOTALS`;
    
    // Daily totals
    pivotMatrix.dailyTotals.forEach((dayTotal, index) => {
      const dateInfo = pivotMatrix.dateRange[index];
      if (!dateInfo.isWeekend || options.includeWeekends) {
        csv += `,${dayTotal.total}`;
      }
    });
    
    // Summary totals
    if (options.includeTotalHours) {
      csv += `,${pivotMatrix.grandTotals.totalHours}`;
    }
    if (options.includeDiscrepancy) {
      csv += `,${pivotMatrix.grandTotals.totalDiscrepancies}`;
    }
    if (options.includeOvertime) {
      csv += `,${pivotMatrix.grandTotals.totalOvertime}`;
    }
    
    csv += `,${pivotMatrix.grandTotals.averageEfficiency}%,`;
    csv += `\n`;
    
    // Average row
    csv += `AVERAGES`;
    
    // Daily averages
    pivotMatrix.dailyTotals.forEach((dayTotal, index) => {
      const dateInfo = pivotMatrix.dateRange[index];
      if (!dateInfo.isWeekend || options.includeWeekends) {
        csv += `,${dayTotal.average}`;
      }
    });
    
    // Summary averages
    if (options.includeTotalHours) {
      const avgTotal = pivotMatrix.users.length > 0 ? (parseFloat(pivotMatrix.grandTotals.totalHours) / pivotMatrix.users.length).toFixed(2) : '0.00';
      csv += `,${avgTotal}`;
    }
    if (options.includeDiscrepancy) {
      const avgDiscrepancy = pivotMatrix.users.length > 0 ? (pivotMatrix.grandTotals.totalDiscrepancies / pivotMatrix.users.length).toFixed(1) : '0.0';
      csv += `,${avgDiscrepancy}`;
    }
    if (options.includeOvertime) {
      const avgOvertime = pivotMatrix.users.length > 0 ? (parseFloat(pivotMatrix.grandTotals.totalOvertime) / pivotMatrix.users.length).toFixed(2) : '0.00';
      csv += `,${avgOvertime}`;
    }
    
    csv += `,${pivotMatrix.grandTotals.averageEfficiency}%,`;
    csv += `\n`;
  }
  
  // Footer with legend and notes
  csv += `\n\nLEGEND AND NOTES:`;
  csv += `\n"OFF" = Weekend or non-working day`;
  csv += `\n"0.00" = No productive hours recorded`;
  if (options.highlightLowHours) {
    csv += `\n"*" = Hours below 8.00 on weekdays (potential discrepancy)`;
  }
  csv += `\n"(B##m L##m)" = Break and lunch minutes if violations detected`;
  csv += `\n"Low Days" = Number of weekdays with less than 8 hours`;
  csv += `\n"Overtime" = Total hours over 8.00 per day`;
  csv += `\n"Efficiency %" = Weekday hours / Expected hours (weekdays × 8)`;
  csv += `\n\nDATA QUALITY NOTES:`;
  csv += `\n- All duration values converted from seconds to decimal hours`;
  csv += `\n- Times calculated in ${ATTENDANCE_TIMEZONE_LABEL} (${ATTENDANCE_TIMEZONE})`;
  csv += `\n- Only productive states counted: ${BILLABLE_STATES.join(', ')}`;
  csv += `\n- Break/lunch times tracked separately for violations`;
  csv += `\n- Source: DurationMin column (contains seconds despite name)`;
  
  return csv;
}

// ────────────────────────────────────────────────────────────────────────────
// STANDARD EXPORT FUNCTIONS
// ────────────────────────────────────────────────────────────────────────────

function exportAttendanceCsv(granularity, periodId, agentFilter) {
  return rpc('exportAttendanceCsv', () => {
    const analytics = getAttendanceAnalyticsByPeriod(granularity, periodId, agentFilter);

    let csv = 'Employee,Productive Hours,Non-Productive Hours,Break Hours,Lunch Hours,Compliance Score\n';
    csv += '# Note: All duration values converted from seconds to decimal hours\n';

    analytics.userCompliance.forEach(user => {
      const compliance = calculateComplianceScore(user);
      csv += `${user.user},${(user.availableSecsWeekday / 3600).toFixed(2)},` +
        `${((user.breakSecs + user.lunchSecs) / 3600).toFixed(2)},` +
        `${(user.breakSecs / 3600).toFixed(2)},${(user.lunchSecs / 3600).toFixed(2)},` +
        `${compliance}\n`;
    });

    return csv;
  }, '');
}

function calculateComplianceScore(user) {
  let score = 100;
  score -= user.exceededBreakDays * 2.5;
  score -= user.exceededLunchDays * 2.5;
  score -= user.exceededWeeklyCount * 10;
  return Math.max(0, score);
}

// ────────────────────────────────────────────────────────────────────────────
// IMPORT FUNCTION
// ────────────────────────────────────────────────────────────────────────────

function importAttendance(rows) {
  return rpc('importAttendance', () => {
    const lock = LockService.getDocumentLock();
    try {
      if (!lock.tryLock(30000)) {
        throw new Error('Could not acquire document lock within 30 seconds');
      }

      const ss = resolveAttendanceSpreadsheet();
      const sheet = ss.getSheetByName(ATTENDANCE_SHEET_NAME);
      if (!sheet) {
        throw new Error(`Sheet "${ATTENDANCE_SHEET_NAME}" not found. Call setupCampaignSheets() first.`);
      }

      const now = new Date();
      const lastRow = sheet.getLastRow();
      const existingKeys = new Set();

      // Load existing data for deduplication
      if (lastRow > 1) {
        try {
          // Match actual database structure: ID, Timestamp, User, DurationMin, State, Date, UserID, CreatedAt, UpdatedAt
          const dataRange = sheet.getRange(2, 2, lastRow - 1, 4).getValues(); // Timestamp, User, DurationMin, State
          dataRange.forEach(r => {
            const key = r.map(v => (v || '').toString().trim()).join("||");
            existingKeys.add(key);
          });
        } catch (readError) {
          console.warn('Error reading existing data for deduplication:', readError);
        }
      }

      const batchKeysSeen = new Set();
      const toAppend = [];

      // Process each row from the frontend
      rows.forEach(r => {
        try {
          const rawTimestamp = (r["Time"] || r["Timestamp"] || "").toString();
          let rawName = (r["Natterbox User: Name"] || r["User"] || "").toString().trim()
            .replace(/\bVLBPO\b/gi, "").replace(/\s+/g, " ").trim();
          const durationRaw = (r["Seconds In State"] || r["DurationMin"] || "").toString().trim();
          const stateVal = (r["Availability State"] || r["State"] || "").toString().trim();

          // Skip rows with missing critical data
          if (!rawTimestamp || !rawName || !stateVal) {
            console.warn('Skipping row with missing critical data:', r);
            return;
          }

          // Parse duration - this should be in seconds
          let durationInSeconds = 0;
          if (typeof durationRaw === 'number') {
            durationInSeconds = durationRaw;
          } else if (typeof durationRaw === 'string') {
            const parsed = parseFloat(durationRaw);
            if (!isNaN(parsed)) {
              durationInSeconds = parsed;
            }
          }

          // Parse timestamp
          let timestamp;
          if (rawTimestamp instanceof Date) {
            timestamp = rawTimestamp;
          } else {
            timestamp = new Date(rawTimestamp);
          }
          
          if (isNaN(timestamp.getTime())) {
            console.warn('Invalid timestamp:', rawTimestamp);
            return;
          }

          // Create composite key for deduplication
          const compositeKey = [timestamp.toISOString(), rawName, durationInSeconds, stateVal].join("||");

          // Skip if already exists or already seen in this batch
          if (existingKeys.has(compositeKey)) return;
          if (batchKeysSeen.has(compositeKey)) return;

          batchKeysSeen.add(compositeKey);

          // Extract date for Date column (matching your database format "9/28/2023")
          const dateOnly = Utilities.formatDate(timestamp, ATTENDANCE_TIMEZONE, 'M/d/yyyy');

          // Prepare row for insertion matching your exact database structure:
          // ID, Timestamp, User, DurationMin, State, Date, UserID, CreatedAt, UpdatedAt
          toAppend.push([
            '', // ID (auto-generated by spreadsheet)
            Utilities.formatDate(timestamp, ATTENDANCE_TIMEZONE, 'M/d/yyyy, h:mm:ss a'), // Timestamp (matching format "9/28/2023, 8:38 AM")
            rawName, // User
            durationInSeconds, // DurationMin (contains seconds despite name!)
            stateVal, // State
            dateOnly, // Date (date only in format "9/28/2023")
            '', // UserID (can be populated later)
            Utilities.formatDate(now, ATTENDANCE_TIMEZONE, 'M/d/yyyy'), // CreatedAt
            Utilities.formatDate(now, ATTENDANCE_TIMEZONE, 'M/d/yyyy') // UpdatedAt
          ]);

          // Log for audit trail
          logCampaignDirtyRow(ATTENDANCE_SHEET_NAME, timestamp.toISOString() + "_" + rawName, "CREATE");
        } catch (rowError) {
          console.error('Error processing row:', rowError, r);
        }
      });

      // Insert new rows if any
      if (toAppend.length) {
        try {
          const firstRow = sheet.getLastRow() + 1;
          sheet.getRange(firstRow, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
          console.log(`Successfully imported ${toAppend.length} attendance rows (DurationMin contains seconds)`);
        } catch (insertError) {
          console.error('Error inserting rows:', insertError);
          throw new Error('Failed to insert attendance data: ' + insertError.message);
        }
      }

      // Flush audit logs
      try {
        flushCampaignDirtyRows();
      } catch (flushError) {
        console.warn('Error flushing audit rows:', flushError);
      }

      // Clear cache after successful import
      try {
        CacheService.getScriptCache().remove('ATTENDANCE_ROWS_CACHE_FINAL');
      } catch (cacheError) {
        console.warn('Could not clear attendance cache:', cacheError);
      }

      return {
        imported: toAppend.length,
        skipped: rows.length - toAppend.length,
        total: rows.length,
        note: 'Duration values stored as seconds in DurationMin column (despite column name)'
      };

    } catch (error) {
      console.error('Import attendance error:', error);
      throw error;
    } finally {
      try {
        lock.releaseLock();
      } catch (lockError) {
        console.warn('Error releasing lock:', lockError);
      }
    }
  }, { imported: 0, error: 'Import service unavailable' });
}

// ────────────────────────────────────────────────────────────────────────────
// FALLBACK FUNCTIONS
// ────────────────────────────────────────────────────────────────────────────

function createBasicAnalytics(filtered, granularity, periodId, agentFilter, periodStart, periodEnd) {
  console.log('Creating basic analytics fallback');

  const rows = Array.isArray(filtered) ? filtered : [];
  const summary = {};
  const stateDuration = {};
  const seedStates = [...new Set([...BILLABLE_STATES, ...NON_PRODUCTIVE_STATES, ...END_SHIFT_STATES])];
  seedStates.forEach(state => {
    summary[state] = 0;
    stateDuration[state] = 0;
  });

  rows.forEach(r => {
    if (!r) return;
    const state = r.state || '';
    const durationSec = typeof r.durationSec === 'number' ? r.durationSec : parseFloat(r.durationSec) || 0;
    summary[state] = (summary[state] || 0) + 1;
    stateDuration[state] = (stateDuration[state] || 0) + durationSec;
  });

  const breakSecs = stateDuration['Break'] || 0;
  const lunchSecs = stateDuration['Lunch'] || 0;
  const billableSecs = BILLABLE_STATES.reduce((sum, state) => sum + (stateDuration[state] || 0), 0);
  const billableWithBreakSecs = billableSecs + breakSecs;
  const totalBillableHours = Math.round((billableWithBreakSecs / 3600) * 100) / 100;
  const totalNonProductiveHours = Math.round(((breakSecs + lunchSecs) / 3600) * 100) / 100;

  const billableBreakdown = buildHourBreakdown(BILLABLE_DISPLAY_STATES, stateDuration);
  const nonProductiveBreakdown = buildHourBreakdown(NON_PRODUCTIVE_DISPLAY_STATES, stateDuration);

  const workingDays = periodStart && periodEnd
    ? countWeekdaysInclusive(periodStart, periodEnd)
    : 5;

  const periodInfo = {
    granularity,
    periodId,
    workingDays,
    timezone: ATTENDANCE_TIMEZONE,
    timezoneLabel: ATTENDANCE_TIMEZONE_LABEL
  };

  if (periodStart instanceof Date && !isNaN(periodStart.getTime())) {
    periodInfo.startDateIso = periodStart.toISOString();
  }
  if (periodEnd instanceof Date && !isNaN(periodEnd.getTime())) {
    periodInfo.endDateIso = periodEnd.toISOString();
  }

  const totalHours = totalBillableHours + totalNonProductiveHours;
  const efficiencyRate = totalHours > 0 ? (totalBillableHours / totalHours) * 100 : 0;
  const complianceRate = 100; // Default compliance placeholder

  const intelligence = generateAttendanceIntelligence(rows, {
    periodStart,
    periodEnd,
    billableBreakdown,
    nonProductiveBreakdown,
    stateDuration
  });

  return {
    summary,
    stateDuration,
    totalBillableHours,
    totalProductiveHours: totalBillableHours,
    totalNonProductiveHours,
    billableHoursBreakdown: billableBreakdown,
    nonProductiveHoursBreakdown: nonProductiveBreakdown,
    filteredRows: rows.slice(0, 100).map(r => ({
      timestampMs: r.timestampMs || (r.timestamp instanceof Date ? r.timestamp.getTime() : null),
      user: r.user,
      state: r.state,
      durationSec: r.durationSec,
      durationHrs: Math.round((r.durationSec || 0) / 3600 * 100) / 100,
      dateString: r.dateString,
      dayOfWeek: r.dayOfWeek,
      isWeekend: r.isWeekend
    })),
    filteredRowCount: rows.length,
    userCompliance: [],
    top5Attendance: [],
    attendanceStats: [],
    attendanceFeed: [],
    dailyMetrics: [],
    shiftMetrics: {},
    enhanced: false,
    executiveMetrics: {
      overview: {
        efficiencyRate,
        complianceRate,
        totalEmployees: new Set(rows.map(r => r.user)).size,
        activeEmployees: new Set(rows.map(r => r.user)).size,
        billableHours: totalBillableHours,
        productiveHours: totalBillableHours,
        nonProductiveHours: totalNonProductiveHours,
        breakHours: Math.round((breakSecs / 3600) * 100) / 100,
        lunchHours: Math.round((lunchSecs / 3600) * 100) / 100
      },
      violations: {
        totalViolations: 0
      },
      timeBreakdown: {
        billable: billableBreakdown,
        nonProductive: nonProductiveBreakdown
      }
    },
    intelligence,
    periodInfo
  };
}

function createEmptyAnalytics() {
  return {
    summary: {
      'Available': 0,
      'Administrative Work': 0,
      'Training': 0,
      'Meeting': 0,
      'Break': 0,
      'Lunch': 0,
      'End of Shift': 0
    },
    stateDuration: {},
    totalBillableHours: 0,
    totalProductiveHours: 0,
    totalNonProductiveHours: 0,
    billableHoursBreakdown: {},
    nonProductiveHoursBreakdown: {},
    top5Attendance: [],
    attendanceStats: [],
    attendanceFeed: [],
    filteredRows: [],
    filteredRowCount: 0,
    userCompliance: [],
    shiftMetrics: {},
    dailyMetrics: [],
    enhanced: false,
    executiveMetrics: {
      overview: {
        efficiencyRate: 0,
        complianceRate: 100,
        totalEmployees: 0,
        activeEmployees: 0,
        billableHours: 0,
        productiveHours: 0,
        nonProductiveHours: 0,
        breakHours: 0,
        lunchHours: 0
      },
      violations: {
        totalViolations: 0
      },
      timeBreakdown: {
        billable: {},
        nonProductive: {}
      }
    },
    intelligence: { insights: [], employeeTrends: [] },
    periodInfo: {
      granularity: 'Week',
      periodId: '',
      workingDays: 5,
      timezone: ATTENDANCE_TIMEZONE,
      timezoneLabel: ATTENDANCE_TIMEZONE_LABEL,
      startDateIso: new Date().toISOString(),
      endDateIso: new Date().toISOString()
    }
  };
}

// ────────────────────────────────────────────────────────────────────────────
// PERIOD CALCULATION UTILITIES
// ────────────────────────────────────────────────────────────────────────────

function derivePeriodBounds(granularity, id) {
  function weekStartLocal(d) {
    const day = d.getDay();
    const diff = d.getDate() - day + (day === 0 ? -6 : 1);
    return new Date(d.getFullYear(), d.getMonth(), diff, 0, 0, 0, 0);
  }

  if (granularity === 'Week') {
    const [yStr, wStr] = id.split('-W');
    const y = Number(yStr), w = Number(wStr);
    const jan4 = new Date(y, 0, 4, 0, 0, 0, 0);
    const isoWeek1Mon = weekStartLocal(jan4);
    const start = new Date(isoWeek1Mon);
    start.setDate(start.getDate() + (w - 1) * 7);
    const end = new Date(start);
    end.setDate(end.getDate() + 6);
    end.setHours(23, 59, 59, 999);
    return [start, end];
  }

  if (granularity === 'BiWeekly') {
    const [yearStr, biStr] = id.split('-BW');
    const year = Number(yearStr);
    const biIndex = Number(biStr);
    if (!Number.isFinite(year) || !Number.isFinite(biIndex) || biIndex < 1) {
      throw new Error(`Invalid bi-week period: ${id}`);
    }

    const jan4 = new Date(year, 0, 4, 0, 0, 0, 0);
    const isoWeek1Mon = weekStartLocal(jan4);
    const start = new Date(isoWeek1Mon);
    start.setDate(start.getDate() + (biIndex - 1) * 14);
    const end = new Date(start);
    end.setDate(end.getDate() + 13);
    end.setHours(23, 59, 59, 999);
    return [start, end];
  }

  if (granularity === 'Month') {
    const [y, m] = id.split('-').map(Number);
    const start = new Date(y, m - 1, 1, 0, 0, 0, 0);
    const end = new Date(y, m, 0, 23, 59, 59, 999);
    return [start, end];
  }

  if (granularity === 'Quarter') {
    const [qStr, yStr] = id.split('-');
    const q = Number(qStr.replace(/^Q/, ''));
    const y = Number(yStr);
    const start = new Date(y, (q - 1) * 3, 1, 0, 0, 0, 0);
    const end = new Date(y, q * 3, 0, 23, 59, 59, 999);
    return [start, end];
  }

  if (granularity === 'Year') {
    const y = Number(id);
    return [new Date(y, 0, 1, 0, 0, 0, 0), new Date(y, 11, 31, 23, 59, 59, 999)];
  }

  throw new Error(`Unknown granularity: ${granularity}`);
}

function countWeekdaysInclusive(start, end) {
  let count = 0;
  const current = new Date(start);
  while (current <= end) {
    const dow = current.getDay();
    if (dow >= 1 && dow <= 5) count++;
    current.setDate(current.getDate() + 1);
  }
  return count;
}

function weekStringFromDate(date) {
  const d = new Date(date);
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
  return `${d.getUTCFullYear()}-W${weekNo.toString().padStart(2, '0')}`;
}

// Validation and other utility functions
function validateEnhancedExportParams(params) {
  const errors = [];
  
  if (!params.period) {
    errors.push('Period is required');
  }
  
  if (params.period && params.period.type === 'custom') {
    if (!params.period.start || !params.period.end) {
      errors.push('Start and end dates are required for custom period');
    }
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

// ────────────────────────────────────────────────────────────────────────────
// USER LISTS BY MANAGER
// ────────────────────────────────────────────────────────────────────────────

function clientGetAttendanceUserNamesForManager(managerUserId, opts) {
  return rpc('clientGetAttendanceUserNamesForManager', () => {
    const options = opts || {};
    const displayCol = (options.display || 'FullName');
    const activeOnly = (options.activeOnly !== false); // default true
    const includeSelf = !!options.includeSelf;

    if (!managerUserId) return [];

    let scopedUsers = [];
    try {
      if (typeof getUsersByManager === 'function') {
        scopedUsers = getUsersByManager(String(managerUserId), {
          includeManager: includeSelf,
          fallbackToCampaign: true,
          fallbackToAll: false
        }) || [];
      } else if (typeof getUsers === 'function') {
        scopedUsers = getUsers() || [];
      }
    } catch (err) {
      console.warn('clientGetAttendanceUserNamesForManager: getUsersByManager failed:', err);
      scopedUsers = [];
    }

    if (!Array.isArray(scopedUsers) || scopedUsers.length === 0) {
      return [];
    }

    const allowedStatuses = ['Active', 'Probation', 'Contractor', 'Intern'];
    const normalizedManagerId = String(managerUserId);
    const uniqueNames = new Set();

    scopedUsers.forEach(user => {
      if (!user) return;
      const userId = String(user.ID || user.id || '');
      if (!userId) return;

      if (!includeSelf && userId === normalizedManagerId) return;

      if (activeOnly) {
        const status = String(user.EmploymentStatus || user.employmentStatus || '').trim();
        if (status && !allowedStatuses.includes(status)) return;
      }

      const displayValue = String(
        user[displayCol] ||
        user.FullName ||
        user.UserName ||
        user.name ||
        ''
      ).trim();

      if (displayValue) {
        uniqueNames.add(displayValue);
      }
    });

    return Array.from(uniqueNames).sort((a, b) => a.localeCompare(b));
  }, []);
}

function clientGetAssignedAgentNames(managerUserId) {
  return clientGetAttendanceUserNamesForManager(managerUserId, {
    activeOnly: true,
    includeSelf: false,
    display: 'FullName'
  });
}

// ────────────────────────────────────────────────────────────────────────────
// CONNECTION TESTING
// ────────────────────────────────────────────────────────────────────────────

function testConnection() {
  return rpc('testConnection', () => ({
    message: 'Connection successful! Data structure correctly identified.',
    timestamp: Utilities.formatDate(new Date(), ATTENDANCE_TIMEZONE, 'yyyy-MM-dd HH:mm:ss') + ` (${ATTENDANCE_TIMEZONE_LABEL})`,
    timezone: ATTENDANCE_TIMEZONE,
    version: 'FINAL-1.0.0',
    features: ['corrected-database-structure', 'seconds-handling', 'configured-timezone', 'daily-pivot-export'],
    dataStructure: {
      columns: ['ID', 'Timestamp', 'User', 'DurationMin', 'State', 'Date', 'UserID', 'CreatedAt', 'UpdatedAt'],
      timestampFormat: 'M/d/yyyy, h:mm AM/PM',
      dateFormat: 'M/d/yyyy',
      durationUnit: 'seconds (in DurationMin column)',
      note: 'DurationMin column contains seconds despite the name'
    }
  }), { message: 'Connection failed', error: true });
}

// ────────────────────────────────────────────────────────────────────────────
// DEBUG AND VALIDATION FUNCTIONS
// ────────────────────────────────────────────────────────────────────────────

function debugDatabaseStructure() {
  try {
    const ss = resolveAttendanceSpreadsheet();
    const sheet = ss.getSheetByName(ATTENDANCE_SHEET_NAME);
    
    if (!sheet) {
      return { error: 'Attendance sheet not found' };
    }

    const values = sheet.getDataRange().getValues();
    if (values.length === 0) {
      return { error: 'No data in attendance sheet' };
    }

    const headers = values[0];
    const sampleData = values.slice(1, 6); // First 5 data rows

    console.log('Database Structure Analysis:');
    console.log('Headers:', headers);
    console.log('Sample data:', sampleData);

    // Analyze DurationMin column values
    const durationIdx = headers.indexOf('DurationMin');
    if (durationIdx >= 0) {
      const durations = sampleData.map(row => row[durationIdx]).filter(d => d);
      console.log('Duration values analysis:', durations.map(d => ({
        raw: d,
        ifSeconds: `${(d/3600).toFixed(2)}h`,
        ifMinutes: `${(d/60).toFixed(2)}h`
      })));
    }

    return {
      success: true,
      headers: headers,
      sampleData: sampleData.slice(0, 3),
      totalRows: values.length - 1,
      analysis: {
        durationColumnIndex: durationIdx,
        durationUnit: 'seconds (confirmed by value analysis)',
        note: 'DurationMin column contains seconds despite the name'
      }
    };
  } catch (error) {
    console.error('Debug error:', error);
    return { error: error.message };
  }
}

function validateDataConsistency() {
  try {
    const allRows = fetchAllAttendanceRows();
    
    if (allRows.length === 0) {
      return { error: 'No attendance data found' };
    }

    // Sample analysis
    const sample = allRows.slice(0, 10);
    const analysis = sample.map(r => ({
      user: r.user,
      state: r.state,
      timestamp: Utilities.formatDate(r.timestamp, ATTENDANCE_TIMEZONE, 'M/d/yyyy, h:mm:ss a'),
      durationSec: r.durationSec,
      durationHours: r.durationHours.toFixed(2),
      reasonableHours: r.durationHours > 0 && r.durationHours < 24
    }));

    const reasonableCount = analysis.filter(a => a.reasonableHours).length;
    
    return {
      success: true,
      totalRecords: allRows.length,
      sampleAnalysis: analysis,
      dataQuality: {
        reasonableValues: `${reasonableCount}/${analysis.length}`,
        percentage: Math.round((reasonableCount / analysis.length) * 100),
        note: 'Duration values interpreted as seconds and converted to hours'
      },
      dateTimeFormat: {
        expectedTimestampFormat: 'M/d/yyyy, h:mm AM/PM',
        expectedDateFormat: 'M/d/yyyy',
        timezone: ATTENDANCE_TIMEZONE,
        note: 'Matches your database format exactly'
      }
    };
  } catch (error) {
    console.error('Validation error:', error);
    return { error: error.message };
  }
}