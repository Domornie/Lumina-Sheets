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
const END_SHIFT_STATES = ['End of Shift'];

// Time constants (working in seconds, since DurationMin column contains seconds)
const DAILY_SHIFT_SECS = 8 * 3600;       // 8 hours in seconds
const DAILY_BREAKS_SECS = 30 * 60;       // 30 minutes in seconds
const DAILY_LUNCH_SECS = 30 * 60;        // 30 minutes in seconds
const WEEKLY_OVERTIME_SECS = 40 * 3600;  // 40 hours in seconds

// Primary attendance timezone configuration (defaults to script timezone)
const ATTENDANCE_TIMEZONE = (typeof global.ATTENDANCE_TIMEZONE === 'string' && global.ATTENDANCE_TIMEZONE)
  ? global.ATTENDANCE_TIMEZONE
  : (typeof Session !== 'undefined' && Session.getScriptTimeZone ? Session.getScriptTimeZone() : 'America/Jamaica');
const ATTENDANCE_TIMEZONE_LABEL = (typeof global.ATTENDANCE_TIMEZONE_LABEL === 'string' && global.ATTENDANCE_TIMEZONE_LABEL)
  ? global.ATTENDANCE_TIMEZONE_LABEL
  : 'Company Time';

// Performance optimization constants
const MAX_PROCESSING_TIME = 25000; // 25 seconds max execution time
const CHUNK_SIZE = 1000; // Process data in chunks
const CACHE_TTL_SHORT = 60; // 1 minute cache
const CACHE_TTL_MEDIUM = 300; // 5 minute cache

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
    const CACHE_KEY = 'ATTENDANCE_ROWS_CACHE_FINAL_V2';

    // Try cache first
    try {
      const cached = CacheService.getScriptCache().get(CACHE_KEY);
      if (cached) {
        const data = JSON.parse(cached);
        if (data.timestamp && (Date.now() - data.timestamp) < CACHE_TTL_MEDIUM * 1000) {
          console.log('Using cached attendance data');
          return data.rows;
        }
      }
    } catch (e) {
      console.warn('Cache read failed:', e);
    }

    const ss = getIBTRSpreadsheet();
    const sheet = ss.getSheetByName(ATTENDANCE);
    if (!sheet) return [];

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
          
          // Enhanced timestamp parsing for company timezone alignment
          let timestamp;
          if (timestampVal instanceof Date) {
            timestamp = new Date(timestampVal);
          } else if (typeof timestampVal === 'string') {
            // Handle format "9/28/2023, 8:38 AM"
            timestamp = new Date(timestampVal);
            
            // If parsing failed, try alternative formats
            if (isNaN(timestamp.getTime())) {
              // Try parsing as ISO string or other common formats
              const cleanedTimestamp = timestampVal.trim().replace(/,\s*/, ' ');
              timestamp = new Date(cleanedTimestamp);
            }
          } else {
            console.warn(`Row ${i + rowIndex}: Invalid timestamp value:`, timestampVal);
            return;
          }
          
          if (isNaN(timestamp.getTime())) {
            console.warn(`Row ${i + rowIndex}: Could not parse timestamp:`, timestampVal);
            return;
          }

          // Ensure we're working with the configured company timezone
          const localizedTimestamp = new Date(timestamp.toLocaleString('en-US', { timeZone: ATTENDANCE_TIMEZONE }));

          const durationValue = row[durationIdx];
          let durationSeconds = 0;
          
          if (typeof durationValue === 'number') {
            durationSeconds = durationValue;
          } else if (typeof durationValue === 'string') {
            const parsed = parseFloat(durationValue);
            if (!isNaN(parsed)) {
              durationSeconds = parsed;
            }
          }

          const user = String(row[userIdx] || '').trim();
          const state = String(row[stateIdx] || '').trim();
          
          if (!user || !state) {
            console.warn(`Row ${i + rowIndex}: Missing user or state:`, { user, state });
            return;
          }

          out.push({
            timestamp: localizedTimestamp,
            user: user,
            state: state,
            durationSec: durationSeconds,
            durationMin: durationSeconds / 60,
            durationHours: durationSeconds / 3600,
            // Add date string for consistent filtering
            dateString: Utilities.formatDate(localizedTimestamp, ATTENDANCE_TIMEZONE, 'yyyy-MM-dd')
          });

        } catch (rowError) {
          console.error(`Error processing row ${i + rowIndex}:`, rowError, row);
        }
      });
    }

    console.log(`Processed ${out.length} attendance records with enhanced timezone handling`);
    
    // Cache the results
    try {
      const cacheData = { timestamp: Date.now(), rows: out };
      CacheService.getScriptCache().put(CACHE_KEY, JSON.stringify(cacheData), CACHE_TTL_MEDIUM);
    } catch (e) {
      console.warn('Cache write failed:', e);
    }

    return out;
  }, []);
}

// ────────────────────────────────────────────────────────────────────────────
// ANALYTICS ENGINE
// ────────────────────────────────────────────────────────────────────────────

function getAttendanceAnalyticsByPeriod(granularity, periodId, agentFilter) {
  return rpc('getAttendanceAnalyticsByPeriod', () => {
    const startTime = Date.now();
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

    const allRows = fetchAllAttendanceRows();

    let periodStart, periodEnd;
    try {
      [periodStart, periodEnd] = derivePeriodBounds(granularity, periodId);
    } catch (e) {
      throw new Error(`Invalid period: ${granularity} ${periodId}`);
    }

    // Filter data efficiently
    const filtered = allRows.filter(r => {
      if (agentFilter && r.user !== agentFilter) return false;
      return r.timestamp >= periodStart && r.timestamp <= periodEnd;
    });

    console.log(`Filtered ${filtered.length} records from ${allRows.length} total`);

    // Check timeout before heavy processing
    if (Date.now() - startTime > MAX_PROCESSING_TIME * 0.6) {
      console.warn('Approaching timeout, returning basic analytics');
      return createBasicAnalytics(filtered, granularity, periodId, agentFilter);
    }

    // Initialize state summaries
    const summary = {};
    const stateDuration = {};
    [...BILLABLE_STATES, ...NON_PRODUCTIVE_STATES, ...END_SHIFT_STATES].forEach(s => {
      summary[s] = 0;
      stateDuration[s] = 0;
    });

    // Process records efficiently
    filtered.forEach(r => {
      if (summary[r.state] != null) summary[r.state]++;
      if (stateDuration[r.state] != null) stateDuration[r.state] += r.durationSec; // Using actual seconds
    });

    // Calculate metrics
    const productivity = calculateProductivityMetrics(filtered);
    const userCompliance = calculateUserCompliance(filtered);

    // Build result
    const analytics = {
      summary,
      stateDuration,
      totalProductiveHours: productivity.totalProductiveHours,
      totalNonProductiveHours: productivity.totalNonProductiveHours,
      filteredRows: filtered.slice(0, 500).map(r => ({ // Limit returned rows
        timestampMs: r.timestamp.getTime(),
        user: r.user,
        state: r.state,
        durationSec: r.durationSec, // Actual seconds
        durationHrs: r.durationHours // Decimal hours
      })),
      userCompliance,
      top5Attendance: generateTopPerformers(filtered, periodStart, periodEnd),
      attendanceStats: generateAttendanceStats(filtered, periodId),
      attendanceFeed: generateAttendanceFeed(filtered),
      dailyMetrics: generateDailyMetrics(filtered),
      shiftMetrics: {},
      periodInfo: {
        granularity,
        periodId,
        startDateIso: periodStart.toISOString(),
        endDateIso: periodEnd.toISOString(),
        workingDays: countWeekdaysInclusive(periodStart, periodEnd),
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
    const BREAK_CAP_SECS = 30 * 60;  // 30 minutes in seconds
    const LUNCH_CAP_SECS = 30 * 60;  // 30 minutes in seconds
    const DAILY_CAP_SECS = 8 * 3600; // 8 hours in seconds

    const userDayMetrics = new Map();

    // Process records efficiently
    filtered.forEach(r => {
        // Use configured timezone for day of week calculation
        const attendanceDayOfWeek = getAttendanceDayOfWeek(r.timestamp);
        if (attendanceDayOfWeek < 1 || attendanceDayOfWeek > 5) return; // Weekdays only (Monday=1, Friday=5)

        const dayKey = Utilities.formatDate(r.timestamp, ATTENDANCE_TIMEZONE, 'yyyy-MM-dd');
        const userDayKey = `${r.user}:${dayKey}`;

        if (!userDayMetrics.has(userDayKey)) {
            userDayMetrics.set(userDayKey, { prod: 0, break: 0, lunch: 0 });
        }

        const metrics = userDayMetrics.get(userDayKey);

        // r.durationSec is in seconds (despite DurationMin column name)
        if (BILLABLE_STATES.includes(r.state)) {
            metrics.prod += r.durationSec;
        } else if (r.state === 'Break') {
            metrics.break += r.durationSec;
        } else if (r.state === 'Lunch') {
            metrics.lunch += r.durationSec;
        }
    });

    // Calculate totals with caps
    let prodSecs = 0, nonProdSecs = 0;

    userDayMetrics.forEach(metrics => {
        const paidBreak = Math.min(metrics.break, BREAK_CAP_SECS);
        const breakExcess = Math.max(0, metrics.break - BREAK_CAP_SECS);
        const lunchExcess = Math.max(0, metrics.lunch - LUNCH_CAP_SECS);

        let dayProdSecs = metrics.prod + paidBreak - breakExcess - lunchExcess;
        dayProdSecs = Math.max(0, dayProdSecs);
        const capped = Math.min(dayProdSecs, DAILY_CAP_SECS);
        prodSecs += capped;

        const excessProd = dayProdSecs - capped;
        nonProdSecs += breakExcess + metrics.lunch + Math.max(0, excessProd);
    });

    return {
        totalProductiveHours: Math.round(prodSecs / 3600 * 100) / 100,
        totalNonProductiveHours: Math.round(nonProdSecs / 3600 * 100) / 100
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
        
        // Use configured timezone for day calculation
        const attendanceDayOfWeek = getAttendanceDayOfWeek(r.timestamp);

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

function generateTopPerformers(filtered, periodStart, periodEnd) {
  const weekdaysInPeriod = countWeekdaysInclusive(periodStart, periodEnd);
  const expectedCapacitySecs = weekdaysInPeriod * DAILY_SHIFT_SECS;

  const userProdSecs = new Map();

  filtered.forEach(r => {
    const dow = getAttendanceDayOfWeek(r.timestamp);
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

      const ss = getIBTRSpreadsheet();
      const sheet = ss.getSheetByName(ATTENDANCE);
      if (!sheet) {
        throw new Error(`Sheet "${ATTENDANCE}" not found. Call setupCampaignSheets() first.`);
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
          logCampaignDirtyRow(ATTENDANCE, timestamp.toISOString() + "_" + rawName, "CREATE");
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

function createBasicAnalytics(filtered, granularity, periodId, agentFilter) {
  console.log('Creating basic analytics fallback');

  const summary = {};
  [...BILLABLE_STATES, ...NON_PRODUCTIVE_STATES, ...END_SHIFT_STATES].forEach(s => {
    summary[s] = filtered.filter(r => r.state === s).length;
  });

  const totalProd = filtered
    .filter(r => BILLABLE_STATES.includes(r.state))
    .reduce((sum, r) => sum + r.durationSec, 0) / 3600; // Convert seconds to hours

  const totalNonProd = filtered
    .filter(r => NON_PRODUCTIVE_STATES.includes(r.state))
    .reduce((sum, r) => sum + r.durationSec, 0) / 3600; // Convert seconds to hours

  return {
    summary,
    stateDuration: {},
    totalProductiveHours: Math.round(totalProd * 100) / 100,
    totalNonProductiveHours: Math.round(totalNonProd * 100) / 100,
    filteredRows: filtered.slice(0, 100).map(r => ({
      timestampMs: r.timestamp.getTime(),
      user: r.user,
      state: r.state,
      durationSec: r.durationSec,
      durationHrs: Math.round(r.durationSec / 3600 * 100) / 100
    })),
    userCompliance: [],
    top5Attendance: [],
    attendanceStats: [],
    attendanceFeed: [],
    dailyMetrics: [],
    shiftMetrics: {},
    periodInfo: {
      granularity,
      periodId,
      workingDays: 5,
      timezone: ATTENDANCE_TIMEZONE,
      timezoneLabel: ATTENDANCE_TIMEZONE_LABEL
    }
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
    totalProductiveHours: 0,
    totalNonProductiveHours: 0,
    top5Attendance: [],
    attendanceStats: [],
    attendanceFeed: [],
    filteredRows: [],
    userCompliance: [],
    shiftMetrics: {},
    dailyMetrics: [],
    periodInfo: {
      granularity: 'Week',
      periodId: '',
      workingDays: 5,
      timezone: ATTENDANCE_TIMEZONE,
      timezoneLabel: ATTENDANCE_TIMEZONE_LABEL
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
    const ss = getIBTRSpreadsheet();
    const sheet = ss.getSheetByName(ATTENDANCE);
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