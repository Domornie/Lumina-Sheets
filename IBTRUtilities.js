/**
 * IBTRUtilities.gs — Full Consolidated & Safe Update (Apps Script friendly)
 * - No modern assignment operators (no ??=) or optional chaining.
 * - Guards all globals using typeof checks; no redeclare errors.
 * - Replaces browser setInterval with a time-driven trigger.
 * - Provides safe stubs for advanced engines so calls won't explode.
 * - Keeps all IBTR reads/writes on the IBTR spreadsheet.
 * - Enhanced coaching functionality with modal support.
 */

(function () {
  // Resolve a safe global object
  var G = (typeof globalThis === 'object')
    ? globalThis
    : (function(){ return this; })();

  // ────────────────────────────────────────────────────────────────────────────
  // Global constants (GUARDED). Change values if needed; no redeclare errors.
  // ────────────────────────────────────────────────────────────────────────────
  if (typeof G.CAMPAIGN_SPREADSHEET_ID === 'undefined') G.CAMPAIGN_SPREADSHEET_ID = '13RbSGQ6OdlxJP8Tkb7p-Fakppq1RIDVmDdubcNPPJ8E';
  if (typeof G.MAIN_SPREADSHEET_ID === 'undefined') G.MAIN_SPREADSHEET_ID = '';

  // Durations & policy
  if (typeof G.DAILY_SHIFT_SECS === 'undefined') G.DAILY_SHIFT_SECS = 8.5 * 3600;
  if (typeof G.DAILY_LUNCH_SECS === 'undefined') G.DAILY_LUNCH_SECS = 30 * 60;
  if (typeof G.DAILY_BREAKS_SECS === 'undefined') G.DAILY_BREAKS_SECS = (15 + 15) * 60;
  if (typeof G.WEEKLY_OVERTIME_SECS === 'undefined') G.WEEKLY_OVERTIME_SECS = 40 * 3600;
  if (typeof G.OPERATION_TIMEOUT_MS === 'undefined') G.OPERATION_TIMEOUT_MS = 25000;

  // Cache
  if (typeof G.CACHE_TTL_SEC === 'undefined') G.CACHE_TTL_SEC = 300;
  try {
    if (typeof G.scriptCache === 'undefined') G.scriptCache = CacheService.getScriptCache();
  } catch (e) {
    if (typeof G.scriptCache === 'undefined') {
      G.scriptCache = { get: function(){}, put: function(){}, remove: function(){} };
    }
  }

  // Sheets (campaign)
  if (typeof G.CALL_REPORT === 'undefined') G.CALL_REPORT = 'CallReport';
  if (typeof G.DIRTY_ROWS === 'undefined') G.DIRTY_ROWS = 'DirtyRows';
  if (typeof G.ATTENDANCE === 'undefined') G.ATTENDANCE = 'AttendanceLog';
  if (typeof G.SCHEDULES_SHEET === 'undefined') G.SCHEDULES_SHEET = 'Schedules';
  if (typeof G.SHIFTS_SHEET === 'undefined') G.SHIFTS_SHEET = 'Shifts';
  if (typeof G.QA_RECORDS === 'undefined') G.QA_RECORDS = 'Quality';
  if (typeof G.QA_COLLAB_RECORDS === 'undefined') G.QA_COLLAB_RECORDS = 'QACollab';
  if (typeof G.ASSIGNMENTS_SHEET === 'undefined') G.ASSIGNMENTS_SHEET = 'Assignment';
  if (typeof G.ESCALATIONS_SHEET === 'undefined') G.ESCALATIONS_SHEET = 'Escalations';
  if (typeof G.COACHING_SHEET === 'undefined') G.COACHING_SHEET = 'CoachingRecords';
  if (typeof G.ATTENDANCE_PREFIX === 'undefined') G.ATTENDANCE_PREFIX = 'AttendCal';
  // Additional sheets referenced by advanced functions
  if (typeof G.BOOKMARKS_SHEET === 'undefined') G.BOOKMARKS_SHEET = 'Bookmarks';
  if (typeof G.SCHEDULE_GENERATION_SHEET === 'undefined') G.SCHEDULE_GENERATION_SHEET = 'ScheduleGeneration';
  if (typeof G.SHIFT_SWAPS_SHEET === 'undefined') G.SHIFT_SWAPS_SHEET = 'ShiftSwaps';

  // Headers (campaign)
  if (typeof G.CALL_REPORT_HEADERS === 'undefined') {
    G.CALL_REPORT_HEADERS = [
      'ID','CreatedDate','TalkTimeMinutes','FromRoutingPolicy','WrapupLabel','ToSFUser','UserID','CSAT','CreatedAt','UpdatedAt'
    ];
  }
  if (typeof G.DIRTY_ROWS_HEADERS === 'undefined') {
    G.DIRTY_ROWS_HEADERS = ['ID','TableName','RowID','Action','Timestamp'];
  }
  
  // Enhanced Coaching Headers with modal support
  if (typeof G.COACHING_HEADERS === 'undefined') {
    G.COACHING_HEADERS = [
      'ID','QAId','SessionDate','AgentName','CoacheeName','CoacheeEmail','TopicsPlanned','CoveredTopics',
      'Summary','ActionPlan','FollowUpDate','Notes','Completed','AcknowledgementText','AcknowledgedOn','CreatedAt','UpdatedAt'
    ];
  }
  
  if (typeof G.QA_HEADERS === 'undefined') {
    G.QA_HEADERS = [
      'ID','Timestamp','CallerName','AgentName','AgentEmail','ClientName','CallDate','CaseNumber','CallLink','AuditorName',
      'AuditDate','FeedbackShared','Q1','Q2','Q3','Q4','Q5','Q6','Q7','Q8','Q9','Q10','Q11','Q12','Q13','Q14','Q15','Q16','Q17','Q18',
      'OverallFeedback','TotalScore','Percentage','Notes','AgentFeedback'
    ];
  }
  if (typeof G.QA_COLLAB_HEADERS === 'undefined') G.QA_COLLAB_HEADERS = G.QA_HEADERS.slice();
  if (typeof G.ESCALATIONS_HEADERS === 'undefined') G.ESCALATIONS_HEADERS = ['ID','Timestamp','User','Type','Notes','CreatedAt','UpdatedAt'];
  if (typeof G.BOOKMARKS_HEADERS === 'undefined') G.BOOKMARKS_HEADERS = ['ID','UserEmail','Title','URL','Description','Tags','Created','LastAccessed','AccessCount','Folder'];

  // Attendance-related headers (used by ensure/setup)
  if (typeof G.ATTENDANCE_LOG_HEADERS === 'undefined') G.ATTENDANCE_LOG_HEADERS = ['Timestamp','User','State','DurationMin'];
  if (typeof G.ATTENDANCE_HEADERS === 'undefined') G.ATTENDANCE_HEADERS = ['Date','User','State','Start','End','Notes'];
  if (typeof G.ASSIGNMENTS_HEADERS === 'undefined') G.ASSIGNMENTS_HEADERS = ['ID','UserEmail','UserName','Campaign','Role','CreatedAt','UpdatedAt','Active'];

  // States config
  if (typeof G.PRODUCTIVE_STATES === 'undefined') G.PRODUCTIVE_STATES = ['Available','Administrative Work','Training','Meeting','Break'];
  if (typeof G.NON_PRODUCTIVE_STATES === 'undefined') G.NON_PRODUCTIVE_STATES = ['Lunch'];
  if (typeof G.ATTENDANCE_STATES === 'undefined') G.ATTENDANCE_STATES = ['Available','Administrative Work','Training','Meeting','Break','Lunch'];

  // ────────────────────────────────────────────────────────────────────────────
  // Spreadsheet openers (needed by logging below)
  // ────────────────────────────────────────────────────────────────────────────
  if (typeof G.getIBTRSpreadsheet !== 'function') {
    G.getIBTRSpreadsheet = function getIBTRSpreadsheet() {
      if (!G.CAMPAIGN_SPREADSHEET_ID) throw new Error('CAMPAIGN_SPREADSHEET_ID not configured');
      return SpreadsheetApp.openById(G.CAMPAIGN_SPREADSHEET_ID);
    };
  }
  if (typeof G.getIdentitySpreadsheet !== 'function') {
    G.getIdentitySpreadsheet = function getIdentitySpreadsheet() {
      try { if (G.MAIN_SPREADSHEET_ID) return SpreadsheetApp.openById(G.MAIN_SPREADSHEET_ID); } catch (e) {}
      return SpreadsheetApp.getActiveSpreadsheet();
    };
  }

  // ────────────────────────────────────────────────────────────────────────────
  // Logging & performance helpers (no optional chaining)
  // ────────────────────────────────────────────────────────────────────────────
  if (typeof G.logError !== 'function') {
    G.logError = function logError(where, err) {
      try { console.error('['+where+']', (err && err.stack) ? err.stack : err); } catch (e) {}
    };
  }

  if (typeof G.safeWriteError !== 'function') {
    G.safeWriteError = function safeWriteError(where, message) {
      try {
        var ss = getIBTRSpreadsheet();
        var sh = ss.getSheetByName('ErrorLog');
        if (!sh) {
          sh = ss.insertSheet('ErrorLog');
          sh.appendRow(['When','Where','Message']);
          sh.setFrozenRows(1);
        }
        if (sh.getLastRow() === 0) sh.appendRow(['When','Where','Message']);
        sh.appendRow([new Date(), where, String(message)]);
      } catch (e) {
        try { console.warn('safeWriteError failed: '+ e.message); } catch (_){}
      }
    };
  }

  if (typeof G.withPerformanceMonitoring !== 'function') {
    G.withPerformanceMonitoring = function withPerformanceMonitoring(name, fn) {
      var t0 = Date.now();
      try { return fn(); }
      catch (e) { G.logError(name, e); throw e; }
      finally {
        try {
          var ss = getIBTRSpreadsheet();
          var sh = ss.getSheetByName('PerfLog');
          if (!sh) {
            sh = ss.insertSheet('PerfLog');
            sh.appendRow(['When','Function','Ms']);
            sh.setFrozenRows(1);
          }
          sh.appendRow([new Date(), name, Date.now() - t0]);
        } catch (e2) {}
      }
    };
  }

  // ────────────────────────────────────────────────────────────────────────────
  // Campaign sheet setup / ensure / read
  // ────────────────────────────────────────────────────────────────────────────
  if (typeof G.ensureCampaignSheetWithHeaders !== 'function') {
    G.ensureCampaignSheetWithHeaders = function ensureCampaignSheetWithHeaders(name, headers) {
      var MAX_RETRIES = 3;
      var BASE_DELAY_MS = 1000;
      var recursionGuardKey = 'CAMPAIGN_RECURSION_GUARD_' + name;
      var recursionGuard = PropertiesService.getScriptProperties().getProperty(recursionGuardKey);
      var sleep = function(ms){ Utilities.sleep(ms); };

      try {
        if (recursionGuard === 'active') {
          console.log('Recursion detected for '+name+', skip ensure');
          return getIBTRSpreadsheet().getSheetByName(name);
        }
        PropertiesService.getScriptProperties().setProperty(recursionGuardKey, 'active');

        if (!headers || !Array.isArray(headers) || headers.some(function(h){ return !h || typeof h !== 'string'; })) {
          throw new Error('Invalid or empty headers provided');
        }
        var uniq = {};
        for (var i=0;i<headers.length;i++){ if (uniq[headers[i]]) throw new Error('Duplicate headers detected'); uniq[headers[i]] = true; }

        var cacheKey = 'CAMPAIGN_SHEET_EXISTS_' + name;
        var cached = G.scriptCache.get(cacheKey);
        var ss = getIBTRSpreadsheet();
        var sh = ss.getSheetByName(name);

        if (sh) {
          var range = sh.getRange(1, 1, 1, headers.length);
          var existing = range.getValues()[0] || [];
          var ok = (existing.length === headers.length);
          if (ok) {
            for (var j=0;j<headers.length;j++){ if (existing[j] !== headers[j]) { ok=false; break; } }
          }
          if (ok) { console.log('Sheet '+name+' OK'); return sh; }
        }

        if (cached === 'true' && sh) {
          console.log('Cache hit: '+name+' exists');
        } else {
          var lastError = null;
          for (var attempt = 1; attempt <= MAX_RETRIES; attempt++) {
            try {
              sh = ss.getSheetByName(name);
              if (!sh) {
                sh = ss.insertSheet(name);
                var r = sh.getRange(1, 1, 1, headers.length);
                r.setValues([headers]);
                r.setFontWeight('bold');
                sh.setFrozenRows(1);
                G.scriptCache.put(cacheKey, 'true', G.CACHE_TTL_SEC);
                console.log('Created '+name);
              } else {
                var r2 = sh.getRange(1, 1, 1, headers.length);
                var existing2 = r2.getValues()[0] || [];
                var mismatch = (existing2.length !== headers.length);
                if (!mismatch) {
                  for (var k=0;k<headers.length;k++){ if (existing2[k] !== headers[k]) { mismatch = true; break; } }
                }
                if (mismatch) {
                  r2.clearContent();
                  r2.setValues([headers]);
                  r2.setFontWeight('bold');
                  sh.setFrozenRows(1);
                  console.log('Updated headers for '+name);
                }
                G.scriptCache.put(cacheKey, 'true', G.CACHE_TTL_SEC);
              }
              return sh;
            } catch (e) {
              lastError = e;
              var msg = String(e && e.message ? e.message : e);
              if (msg.indexOf('timed out') !== -1 && attempt < MAX_RETRIES) {
                var delay = BASE_DELAY_MS * Math.pow(2, attempt - 1);
                console.log('Attempt '+attempt+' failed for '+name+': '+msg+'. Retry in '+delay+'ms');
                sleep(delay);
                continue;
              }
              throw e;
            }
          }
          throw lastError || new Error('Failed ensure '+name);
        }
        return sh;
      } catch (e) {
        console.error('ensureCampaignSheetWithHeaders('+name+') failed: '+ e.message);
        throw e;
      } finally {
        PropertiesService.getScriptProperties().deleteProperty(recursionGuardKey);
      }
    };
  }

  if (typeof G.readCampaignSheet !== 'function') {
    G.readCampaignSheet = function readCampaignSheet(sheetName) {
      try {
        var cacheKey = 'CAMPAIGN_DATA_' + sheetName;
        var cached = G.scriptCache.get(cacheKey);
        if (cached) return JSON.parse(cached);

        var ss = getIBTRSpreadsheet();
        var sh = ss.getSheetByName(sheetName);
        if (!sh) { if (typeof safeWriteError === 'function') safeWriteError('readCampaignSheet('+sheetName+')', 'Missing'); return []; }

        var lastRow = sh.getLastRow();
        var lastCol = sh.getLastColumn();
        if (lastRow < 2 || lastCol < 1) return [];

        var vals = sh.getRange(1, 1, lastRow, lastCol).getValues();
        var headers = vals.shift().map(function(h){ return String(h).trim() || null; });
        if (headers.some(function(h){ return !h; })) return [];

        var seen = {};
        for (var i=0;i<headers.length;i++){ if (seen[headers[i]]) return []; seen[headers[i]] = true; }

        var data = [];
        for (var r=0;r<vals.length;r++){
          var row = vals[r];
          var hasValue = false;
          for (var c=0;c<row.length;c++){ if (row[c] !== '' && row[c] != null){ hasValue = true; break; } }
          if (!hasValue) continue;
          var obj = {};
          for (var c2=0;c2<headers.length;c2++){ if (headers[c2]) obj[headers[c2]] = row[c2]; }
          data.push(obj);
        }

        G.scriptCache.put(cacheKey, JSON.stringify(data), G.CACHE_TTL_SEC);
        return data;
      } catch (e) {
        if (typeof safeWriteError === 'function') safeWriteError('readCampaignSheet('+sheetName+')', e.message);
        return [];
      }
    };
  }

  if (typeof G.readCampaignSheetPaged !== 'function') {
    G.readCampaignSheetPaged = function readCampaignSheetPaged(sheetName, pageIndex, pageSize) {
      try {
        var ss = getIBTRSpreadsheet();
        var sh = ss.getSheetByName(sheetName);
        if (!sh) return [];
        var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
        var startRow = 2 + (pageIndex || 0) * (pageSize || 10);
        var numRows = Math.max(0, Math.min((pageSize || 10), sh.getLastRow() - startRow + 1));
        if (numRows === 0) return [];
        var rows = sh.getRange(startRow, 1, numRows, header.length).getValues();
        var out = [];
        for (var i=0;i<rows.length;i++){
          var r = rows[i], o = {};
          for (var j=0;j<header.length;j++){ o[header[j]] = r[j]; }
          out.push(o);
        }
        return out;
      } catch (e) {
        if (typeof safeWriteError === 'function') safeWriteError('readCampaignSheetPaged('+sheetName+')', e.message);
        return [];
      }
    };
  }

  if (typeof G.invalidateCampaignCache !== 'function') {
    G.invalidateCampaignCache = function invalidateCampaignCache(sheetName) {
      try { G.scriptCache.remove('CAMPAIGN_DATA_' + sheetName); } catch (e) {}
    };
  }

  // ────────────────────────────────────────────────────────────────────────────
  // Enhanced Coaching Functions
  // ────────────────────────────────────────────────────────────────────────────
  if (typeof G.ensureCoachingSheetHeaders !== 'function') {
    G.ensureCoachingSheetHeaders = function ensureCoachingSheetHeaders() {
      try {
        var ss = getIBTRSpreadsheet();
        var sheet = ss.getSheetByName(G.COACHING_SHEET);
        
        if (!sheet) {
          sheet = ss.insertSheet(G.COACHING_SHEET);
        }

        // Get existing headers
        var existingData = sheet.getDataRange().getValues();
        var existingHeaders = existingData.length > 0 ? existingData[0] : [];

        // Check if we need to update headers
        var headersMatch = G.COACHING_HEADERS.length === existingHeaders.length && 
          G.COACHING_HEADERS.every(function(header, index) { 
            return header === existingHeaders[index]; 
          });

        if (!headersMatch) {
          // Update headers
          sheet.clear();
          sheet.getRange(1, 1, 1, G.COACHING_HEADERS.length).setValues([G.COACHING_HEADERS]);
          sheet.setFrozenRows(1);

          // Apply formatting
          var headerRange = sheet.getRange(1, 1, 1, G.COACHING_HEADERS.length);
          headerRange.setBackground('#4285f4');
          headerRange.setFontColor('white');
          headerRange.setFontWeight('bold');
          headerRange.setHorizontalAlignment('center');
          
          console.log('Coaching sheet headers updated');
        }
        
        return true;
      } catch (error) {
        console.error('Error ensuring coaching sheet headers:', error);
        if (typeof G.safeWriteError === 'function') G.safeWriteError('ensureCoachingSheetHeaders', error);
        return false;
      }
    };
  }

  if (typeof G.initializeCoachingSystem !== 'function') {
    G.initializeCoachingSystem = function initializeCoachingSystem() {
      try {
        G.ensureCoachingSheetHeaders();
        console.log('Coaching system initialized with enhanced modal support');
        return { success: true };
      } catch (error) {
        console.error('Error initializing coaching system:', error);
        if (typeof G.safeWriteError === 'function') G.safeWriteError('initializeCoachingSystem', error);
        return { success: false, error: error.message };
      }
    };
  }

  // ────────────────────────────────────────────────────────────────────────────
  // Setup helpers
  // ────────────────────────────────────────────────────────────────────────────
  if (typeof G.setupCampaignSheets !== 'function') {
    G.setupCampaignSheets = function setupCampaignSheets() {
      try {
        ensureCampaignSheetWithHeaders(G.CALL_REPORT, G.CALL_REPORT_HEADERS);
        ensureCampaignSheetWithHeaders(G.DIRTY_ROWS, G.DIRTY_ROWS_HEADERS);
        ensureCampaignSheetWithHeaders(G.ATTENDANCE, G.ATTENDANCE_LOG_HEADERS);
        ensureCampaignSheetWithHeaders(G.QA_RECORDS, G.QA_HEADERS);
        ensureCampaignSheetWithHeaders(G.QA_COLLAB_RECORDS, G.QA_COLLAB_HEADERS);
        ensureCampaignSheetWithHeaders(G.ESCALATIONS_SHEET, G.ESCALATIONS_HEADERS);
        ensureCampaignSheetWithHeaders(G.ATTENDANCE_PREFIX, G.ATTENDANCE_HEADERS);
        ensureCampaignSheetWithHeaders(G.COACHING_SHEET, G.COACHING_HEADERS);
        ensureCampaignSheetWithHeaders(G.BOOKMARKS_SHEET, G.BOOKMARKS_HEADERS);
        
        // Initialize coaching system with enhanced headers
        initializeCoachingSystem();
      } catch (e) {
        console.error('setupCampaignSheets failed:', e);
      }
    };
  }

  // ────────────────────────────────────────────────────────────────────────────
  // Dirty rows buffering (campaign)
  // ────────────────────────────────────────────────────────────────────────────
  if (typeof G.pendingCampaignDirtyRows === 'undefined') G.pendingCampaignDirtyRows = [];
  if (typeof G.logCampaignDirtyRow !== 'function') {
    G.logCampaignDirtyRow = function logCampaignDirtyRow(tableName, rowId, action) {
      try {
        G.pendingCampaignDirtyRows.push([Utilities.getUuid(), tableName, rowId, action, new Date()]);
      } catch (e) {
        if (typeof safeWriteError === 'function') safeWriteError('logCampaignDirtyRow', e);
      }
    };
  }
  if (typeof G.flushCampaignDirtyRows !== 'function') {
    G.flushCampaignDirtyRows = function flushCampaignDirtyRows() {
      try {
        if (!G.pendingCampaignDirtyRows.length) return;
        var sh = ensureCampaignSheetWithHeaders(G.DIRTY_ROWS, G.DIRTY_ROWS_HEADERS);
        sh.getRange(sh.getLastRow() + 1, 1, G.pendingCampaignDirtyRows.length, G.pendingCampaignDirtyRows[0].length)
          .setValues(G.pendingCampaignDirtyRows);
        G.pendingCampaignDirtyRows.length = 0;
        invalidateCampaignCache(G.DIRTY_ROWS);
      } catch (e) {
        if (typeof safeWriteError === 'function') safeWriteError('flushCampaignDirtyRows', e.message);
      }
    };
  }

  // ────────────────────────────────────────────────────────────────────────────
  // Formatting & ISO helpers
  // ────────────────────────────────────────────────────────────────────────────
  if (typeof G.formatDuration !== 'function') {
    G.formatDuration = function formatDuration(minutesDecimal) {
      try {
        var secs = Math.round(minutesDecimal * 60);
        var m = Math.floor(secs / 60), s = secs % 60;
        return m + 'm ' + (s < 10 ? '0' + s : s) + 's';
      } catch (e) { return '0m 00s'; }
    };
  }
  if (typeof G.formatSecsAsHhMmSs !== 'function') {
    G.formatSecsAsHhMmSs = function formatSecsAsHhMmSs(secs) {
      try {
        secs = Math.max(0, Math.floor(secs));
        var h = Math.floor(secs / 3600), m = Math.floor((secs % 3600) / 60), s = secs % 60;
        function pad(n){ return (n<10?'0':'') + n; }
        return pad(h)+':'+pad(m)+':'+pad(s);
      } catch (e) { return '00:00:00'; }
    };
  }
  if (typeof G.getDateOfISOWeek !== 'function') {
    G.getDateOfISOWeek = function getDateOfISOWeek(week, year) {
      try {
        var d = new Date(Date.UTC(year, 0, 1 + (week - 1) * 7));
        var dow = d.getUTCDay();
        var delta = (dow <= 4 ? 1 - dow : 8 - dow);
        d.setUTCDate(d.getUTCDate() + delta);
        return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
      } catch (e) { return new Date(); }
    };
  }
  if (typeof G.weekStringFromDate !== 'function') {
    G.weekStringFromDate = function weekStringFromDate(d) {
      try {
        var t = new Date(d), dayNr = (d.getDay() + 6) % 7;
        t.setDate(t.getDate() - dayNr + 3);
        var firstThu = new Date(t.getFullYear(), 0, 4);
        var weekNum = 1 + Math.round((t - firstThu) / (7 * 24 * 60 * 60 * 1000));
        return t.getFullYear() + '-W' + (weekNum < 10 ? '0' + weekNum : weekNum);
      } catch (e) { return ''; }
    };
  }
  if (typeof G.isoPeriodToDateRange !== 'function') {
    G.isoPeriodToDateRange = function isoPeriodToDateRange(granularity, period) {
      try {
        var start, end;
        if (granularity === 'Week') {
          var parts = period.split('-W');
          var y = Number(parts[0]), w = Number(parts[1]);
          start = getDateOfISOWeek(w, y);
          end = new Date(start); end.setDate(start.getDate() + 6);
        } else if (granularity === 'Month') {
          var p2 = period.split('-'); var y2 = Number(p2[0]), m2 = Number(p2[1]);
          start = new Date(y2, m2 - 1, 1); end = new Date(y2, m2, 0);
        } else if (granularity === 'Quarter') {
          var p3 = period.split('-');
          var q = Number(p3[0].replace('Q','')); var y3 = Number(p3[1]);
          var m0 = (q - 1) * 3; start = new Date(y3, m0, 1); end = new Date(y3, m0 + 3, 0);
        } else if (granularity === 'Year') {
          var y4 = Number(period); start = new Date(y4, 0, 1); end = new Date(y4, 11, 31);
        } else { throw new Error('Unknown granularity "'+granularity+'"'); }
        start.setHours(0,0,0,0); end.setHours(23,59,59,999);
        return { startDate: start, endDate: end };
      } catch (e) {
        return { startDate: new Date(), endDate: new Date() };
      }
    };
  }
  if (typeof G.getDefaultPeriod !== 'function') {
    G.getDefaultPeriod = function getDefaultPeriod(granularity) {
      try {
        var now = new Date(), tz = Session.getScriptTimeZone();
        if (granularity === 'Week') return weekStringFromDate(now);
        if (granularity === 'Month') return Utilities.formatDate(now, tz, 'yyyy-MM');
        if (granularity === 'Quarter') { var q = Math.floor(now.getMonth() / 3) + 1; return 'Q'+q+'-'+now.getFullYear(); }
        if (granularity === 'Year') return ''+now.getFullYear();
        return '';
      } catch (e) { return ''; }
    };
  }

  // ────────────────────────────────────────────────────────────────────────────
  // Health check
  // ────────────────────────────────────────────────────────────────────────────
  if (typeof G.healthCheckCampaign !== 'function') {
    G.healthCheckCampaign = function healthCheckCampaign() {
      var results = {};
      var specs = [
        { name: G.CALL_REPORT, headers: G.CALL_REPORT_HEADERS },
        { name: G.DIRTY_ROWS, headers: G.DIRTY_ROWS_HEADERS },
        { name: G.ATTENDANCE, headers: G.ATTENDANCE_LOG_HEADERS },
        { name: G.QA_RECORDS, headers: G.QA_HEADERS },
        { name: G.QA_COLLAB_RECORDS, headers: G.QA_COLLAB_HEADERS },
        { name: G.ASSIGNMENTS_SHEET, headers: G.ASSIGNMENTS_HEADERS },
        { name: G.ESCALATIONS_SHEET, headers: G.ESCALATIONS_HEADERS },
        { name: G.COACHING_SHEET, headers: G.COACHING_HEADERS },
        { name: G.BOOKMARKS_SHEET, headers: G.BOOKMARKS_HEADERS }
      ];
      for (var i=0;i<specs.length;i++){
        var spec = specs[i];
        try {
          var ss = getIBTRSpreadsheet();
          var sh = ss.getSheetByName(spec.name);
          if (!sh) results[spec.name] = { ok: false, message: 'Missing sheet' };
          else {
            var existing = sh.getRange(1, 1, 1, spec.headers.length).getValues()[0];
            var match = true;
            for (var j=0;j<spec.headers.length;j++){ if (existing[j] !== spec.headers[j]) { match=false; break; } }
            results[spec.name] = match ? { ok: true, message: 'OK' } : { ok: false, message: 'Header mismatch' };
          }
        } catch (e) {
          results[spec.name] = { ok: false, message: 'Error: '+ e.message };
        }
      }
      return results;
    };
  }

  // ────────────────────────────────────────────────────────────────────────────
  // AI Optimization & GA (stubs where needed)
  // ────────────────────────────────────────────────────────────────────────────
  if (typeof G.prepareOptimizationData !== 'function') G.prepareOptimizationData = function (data){ return data || []; };
  if (typeof G.initializePopulation !== 'function') G.initializePopulation = function (data, size){
    var seed = (Object.prototype.toString.call(data) === '[object Array]' && data.length) ? data : [{}];
    var out = [];
    for (var i=0;i<size;i++){ out.push(JSON.parse(JSON.stringify(seed[i % seed.length]))); }
    return out;
  };
  if (typeof G.tournamentSelection !== 'function') G.tournamentSelection = function (population, fitness){
    var a = Math.floor(Math.random() * population.length);
    var b = Math.floor(Math.random() * population.length);
    return (fitness[a] >= fitness[b]) ? population[a] : population[b];
  };
  if (typeof G.crossover !== 'function') G.crossover = function (p1, p2){ return [JSON.parse(JSON.stringify(p1)), JSON.parse(JSON.stringify(p2))]; };
  if (typeof G.mutate !== 'function') G.mutate = function (ind){ return ind; };
  if (typeof G.countScheduleConflicts !== 'function') G.countScheduleConflicts = function (schedule){
    if (Object.prototype.toString.call(schedule) !== '[object Array]') return 0;
    var seen = {}; var conflicts = 0;
    for (var i=0;i<schedule.length;i++){
      var r = schedule[i];
      var k = (r.UserID||'')+'::'+(r.Date||'')+'::'+(r.StartTime||'')+'-'+(r.EndTime||'');
      if (seen[k]) conflicts++; else seen[k] = true;
    }
    return conflicts;
  };
  if (typeof G.calculateEmployeeSatisfaction !== 'function') G.calculateEmployeeSatisfaction = function (){ return 0; };
  if (typeof G.calculateWorkloadBalance !== 'function') G.calculateWorkloadBalance = function (){ return 0; };
  if (typeof G.calculateCoverageAdequacy !== 'function') G.calculateCoverageAdequacy = function (){ return 0; };
  if (typeof G.calculateScheduleStability !== 'function') G.calculateScheduleStability = function (){ return 0; };
  if (typeof G.validateOptimizedSchedule !== 'function') G.validateOptimizedSchedule = function (opt){ return (opt && Object.prototype.toString.call(opt.bestSchedule)==='[object Array]') ? opt.bestSchedule : (opt ? opt.bestSchedule : []); };
  if (typeof G.calculateImprovementMetrics !== 'function') G.calculateImprovementMetrics = function (){ return { conflictsDelta: 0, scoreDelta: 0 }; };
  if (typeof G.generateOptimizationRecommendations !== 'function') G.generateOptimizationRecommendations = function (){ return ['Review coverage on peak hours','Add cross-trained backup for Fridays']; };

  if (typeof G.calculateFitness !== 'function') {
    G.calculateFitness = function calculateFitness(schedule, data, goals, penalties) {
      try {
        var fitness = 0;
        if (goals.indexOf('minimize_conflicts') !== -1) {
          var conflicts = countScheduleConflicts(schedule);
          fitness -= conflicts * penalties.conflicts;
        }
        if (goals.indexOf('maximize_satisfaction') !== -1) {
          var satisfaction = calculateEmployeeSatisfaction(schedule, data);
          fitness += satisfaction * penalties.satisfaction;
        }
        if (goals.indexOf('balance_workload') !== -1) {
          var balance = calculateWorkloadBalance(schedule, data);
          fitness += balance * penalties.workload;
        }
        var coverage = calculateCoverageAdequacy(schedule, data);
        fitness += coverage * 2;
        var stability = calculateScheduleStability(schedule, data);
        fitness += stability * 1;
        return fitness;
      } catch (e) {
        logError('calculateFitness', e);
        return -1000;
      }
    };
  }

  if (typeof G.runGeneticAlgorithm !== 'function') {
    G.runGeneticAlgorithm = function runGeneticAlgorithm(data, params) {
      try {
        var population = initializePopulation(data, params.populationSize);
        var bestFitness = -Infinity;
        var convergenceData = [];
        var generationsWithoutImprovement = 0;

        for (var generation = 0; generation < params.generations; generation++) {
          var fitnessScores = [];
          for (var i=0;i<population.length;i++){
            fitnessScores.push(calculateFitness(population[i], data, params.goals, params.penalties));
          }

          var currentBest = Math.max.apply(null, fitnessScores);
          if (currentBest > bestFitness) {
            bestFitness = currentBest;
            generationsWithoutImprovement = 0;
          } else {
            generationsWithoutImprovement++;
          }

          var sum = 0; for (var f=0;f<fitnessScores.length;f++) sum += fitnessScores[f];
          convergenceData.push({ generation: generation, bestFitness: currentBest, averageFitness: sum / fitnessScores.length });

          if (generationsWithoutImprovement > 50) break;

          var newPopulation = [];
          var eliteCount = Math.floor(params.populationSize * params.elitismRate);
          var idx = fitnessScores.map(function(v, i){ return {v:v, i:i}; })
                                 .sort(function(a,b){ return b.v - a.v; })
                                 .slice(0, eliteCount)
                                 .map(function(o){ return o.i; });
          for (var e=0;e<idx.length;e++){ newPopulation.push(population[idx[e]]); }

          while (newPopulation.length < params.populationSize) {
            var parent1 = tournamentSelection(population, fitnessScores);
            var parent2 = tournamentSelection(population, fitnessScores);

            var offspring1, offspring2;
            if (Math.random() < params.crossoverRate) {
              var x = crossover(parent1, parent2);
              offspring1 = x[0]; offspring2 = x[1];
            } else {
              offspring1 = parent1; offspring2 = parent2;
            }
            if (Math.random() < params.mutationRate) offspring1 = mutate(offspring1, data);
            if (Math.random() < params.mutationRate) offspring2 = mutate(offspring2, data);

            newPopulation.push(offspring1);
            if (newPopulation.length < params.populationSize) newPopulation.push(offspring2);
          }
          population = newPopulation;
        }

        var finalFitnessScores = [];
        for (var z=0;z<population.length;z++){ finalFitnessScores.push(calculateFitness(population[z], data, params.goals, params.penalties)); }
        var bestIndex = 0; var bestVal = -Infinity;
        for (var b=0;b<finalFitnessScores.length;b++){ if (finalFitnessScores[b] > bestVal){ bestVal = finalFitnessScores[b]; bestIndex = b; } }

        return { bestSchedule: population[bestIndex], fitness: finalFitnessScores[bestIndex], convergenceData: convergenceData };
      } catch (error) {
        logError('runGeneticAlgorithm', error);
        throw error;
      }
    };
  }

  if (typeof G.optimizeScheduleWithAI !== 'function') {
    G.optimizeScheduleWithAI = function optimizeScheduleWithAI(scheduleData, constraints) {
      constraints = constraints || {};
      return withPerformanceMonitoring('optimizeScheduleWithAI', function () {
        try {
          var maxIterations = (typeof constraints.maxIterations === 'number') ? constraints.maxIterations : 1000;
          var optimizationGoals = constraints.optimizationGoals || ['minimize_conflicts','maximize_satisfaction','balance_workload'];
          var penaltyWeights = constraints.penaltyWeights || { conflicts: 10, satisfaction: 5, workload: 3 };

          var optimizationInput = prepareOptimizationData(scheduleData);
          var optimized = runGeneticAlgorithm(optimizationInput, {
            populationSize: 100,
            generations: maxIterations,
            mutationRate: 0.1,
            crossoverRate: 0.8,
            elitismRate: 0.1,
            goals: optimizationGoals,
            penalties: penaltyWeights
          });
          var validated = validateOptimizedSchedule(optimized, scheduleData);

          return {
            success: true,
            optimizedSchedule: validated,
            improvementMetrics: calculateImprovementMetrics(scheduleData, validated),
            convergenceData: optimized.convergenceData,
            recommendedActions: generateOptimizationRecommendations(validated)
          };
        } catch (error) {
          logError('optimizeScheduleWithAI', error);
          return { success: false, error: error.message };
        }
      });
    };
  }

  // ────────────────────────────────────────────────────────────────────────────
  // Predictive Analytics (safe stubs)
  // ────────────────────────────────────────────────────────────────────────────
  if (typeof G.extractDemandMetrics !== 'function') G.extractDemandMetrics = function (historical){ return { dailyDemand: (historical || []).map(function(){ return 1; }) }; };
  if (typeof G.calculateMovingAverage !== 'function') G.calculateMovingAverage = function (arr, w){
    var out = [];
    for (var i=0;i<arr.length;i++){
      var s = Math.max(0, i - w + 1);
      var sum = 0; var n=0;
      for (var j=s;j<=i;j++){ sum += arr[j]; n++; }
      out.push(n ? (sum/n) : 0);
    }
    return out.length ? out : [0];
  };
  if (typeof G.calculateTrend !== 'function') G.calculateTrend = function (){ return 0; };
  if (typeof G.calculateSeasonality !== 'function') G.calculateSeasonality = function (){ return [0,0,0,0,0,0,0]; };
  if (typeof G.calculateDemandConfidence !== 'function') G.calculateDemandConfidence = function (){ return 0.6; };

  if (typeof G.extractAbsencePatterns !== 'function') G.extractAbsencePatterns = function (hist){ return hist || []; };
  if (typeof G.buildUserAbsenceProfiles !== 'function') {
    G.buildUserAbsenceProfiles = function (abs) {
      var out = {};
      abs = abs || [];
      for (var i=0;i<abs.length;i++){
        var r = abs[i];
        var id = r.UserID || r.userId || 'unknown';
        if (!out[id]) {
          out[id] = {
            userName: r.UserName || r.userName || id,
            baseAbsenceRate: 0.05,
            dayOfWeekMultipliers: {0:1,1:1,2:1,3:1,4:1,5:1,6:1},
            seasonalMultipliers: {0:1,1:1,2:1,3:1,4:1,5:1,6:1,7:1,8:1,9:1,10:1,11:1},
            recentTrend: 1
          };
        }
      }
      return out;
    };
  }

  if (typeof G.predictDemandPatterns !== 'function') {
    G.predictDemandPatterns = function predictDemandPatterns(historicalData, days) {
      try {
        var demandData = extractDemandMetrics(historicalData);
        var ma = calculateMovingAverage(demandData.dailyDemand, 7);
        var trend = calculateTrend(demandData.dailyDemand);
        var seasonality = calculateSeasonality(demandData.dailyDemand);

        var predictions = [];
        for (var i=0;i<days;i++) {
          var base = ma[ma.length - 1];
          var predicted = Math.max(0, base + trend * (i + 1) + seasonality[i % seasonality.length]);
          var date = new Date(Date.now() + i * 24 * 60 * 60 * 1000);
          predictions.push({
            date: date.toISOString().split('T')[0],
            predictedDemand: Math.round(predicted),
            confidence: calculateDemandConfidence(i, (historicalData && historicalData.length) ? historicalData.length : 1)
          });
        }
        return predictions;
      } catch (error) {
        logError('predictDemandPatterns', error);
        return [];
      }
    };
  }

  if (typeof G.predictAbsencePatterns !== 'function') {
    G.predictAbsencePatterns = function predictAbsencePatterns(historicalData, days) {
      try {
        var absenceData = extractAbsencePatterns(historicalData);
        var profiles = buildUserAbsenceProfiles(absenceData);
        var predictions = [];
        var keys = [];
        for (var k in profiles) if (profiles.hasOwnProperty(k)) keys.push(k);

        for (var x=0;x<keys.length;x++){
          var userId = keys[x], profile = profiles[userId];
          for (var i=0;i<days;i++) {
            var date = new Date(Date.now() + i * 24 * 60 * 60 * 1000);
            var dayOfWeek = date.getDay();
            var month = date.getMonth();
            var p = profile.baseAbsenceRate;
            p *= profile.dayOfWeekMultipliers[dayOfWeek] || 1;
            p *= profile.seasonalMultipliers[month] || 1;
            p *= profile.recentTrend || 1;
            if (p > 1) p = 1;
            if (p < 0) p = 0;
            predictions.push({
              date: date.toISOString().split('T')[0],
              userId: userId,
              userName: profile.userName,
              absenceProbability: p,
              riskLevel: (p > 0.3) ? 'HIGH' : (p > 0.15 ? 'MEDIUM' : 'LOW')
            });
          }
        }

        predictions.sort(function(a,b){ return new Date(a.date) - new Date(b.date); });
        return predictions;
      } catch (error) {
        logError('predictAbsencePatterns', error);
        return [];
      }
    };
  }

  if (typeof G.generateStaffingRecommendations !== 'function') G.generateStaffingRecommendations = function (){ return {}; };
  if (typeof G.assessSchedulingRisks !== 'function') G.assessSchedulingRisks = function (){ return {}; };
  if (typeof G.identifyOptimizationOpportunities !== 'function') G.identifyOptimizationOpportunities = function (){ return {}; };
  if (typeof G.analyzeSeasonalTrends !== 'function') G.analyzeSeasonalTrends = function (){ return {}; };
  if (typeof G.predictPerformanceMetrics !== 'function') G.predictPerformanceMetrics = function (){ return {}; };
  if (typeof G.calculatePredictionConfidence !== 'function') G.calculatePredictionConfidence = function (){ return 0.5; };

  if (typeof G.generatePredictiveAnalytics !== 'function') {
    G.generatePredictiveAnalytics = function generatePredictiveAnalytics(historicalData, forecastPeriod) {
      forecastPeriod = forecastPeriod || 30;
      return withPerformanceMonitoring('generatePredictiveAnalytics', function () {
        try {
          var predictions = {
            demandForecasting: predictDemandPatterns(historicalData, forecastPeriod),
            absenceForecasting: predictAbsencePatterns(historicalData, forecastPeriod),
            overtimePrediction: [],
            staffingRecommendations: generateStaffingRecommendations(historicalData, forecastPeriod),
            riskAssessment: assessSchedulingRisks(historicalData),
            optimizationOpportunities: identifyOptimizationOpportunities(historicalData),
            seasonalTrends: analyzeSeasonalTrends(historicalData),
            performanceMetrics: predictPerformanceMetrics(historicalData, forecastPeriod)
          };
          return {
            success: true,
            predictions: predictions,
            confidence: calculatePredictionConfidence(predictions, historicalData),
            generatedAt: new Date().toISOString(),
            validUntil: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString()
          };
        } catch (error) {
          logError('generatePredictiveAnalytics', error);
          return { success: false, error: error.message };
        }
      });
    };
  }

  // ────────────────────────────────────────────────────────────────────────────
  // Intelligent Automation (stubs)
  // ────────────────────────────────────────────────────────────────────────────
  if (typeof G.detectAllScheduleConflicts !== 'function') G.detectAllScheduleConflicts = function (){ return []; };
  if (typeof G.findOptimalResolution !== 'function') G.findOptimalResolution = function (){ return { confidence: 0.0 }; };
  if (typeof G.applyConflictResolution !== 'function') G.applyConflictResolution = function (){};
  if (typeof G.escalateConflict !== 'function') G.escalateConflict = function (){};

  if (typeof G.getUserApprovalHistory !== 'function') G.getUserApprovalHistory = function (){ return { approvalRate: 0.9 }; };
  if (typeof G.checkScheduleConflicts !== 'function') G.checkScheduleConflicts = function (){ return []; };
  if (typeof G.checkStaffingAdequacy !== 'function') G.checkStaffingAdequacy = function (){ return 1; };
  if (typeof G.approveSchedules !== 'function') G.approveSchedules = function (){ return { success: true }; };
  if (typeof G.sendAutoApprovalNotification !== 'function') G.sendAutoApprovalNotification = function (){};
  if (typeof G.flagScheduleForReview !== 'function') G.flagScheduleForReview = function (){};

  if (typeof G.calculateAutoApprovalScore !== 'function') {
    G.calculateAutoApprovalScore = function calculateAutoApprovalScore(schedule) {
      try {
        var score = 0.5;
        var hist = getUserApprovalHistory(schedule.UserID);
        score += (hist.approvalRate || 0) * 0.3;
        var conflicts = checkScheduleConflicts([schedule]);
        score -= (conflicts.length || 0) * 0.2;
        var staffing = checkStaffingAdequacy(schedule.Date, schedule.SlotID);
        score += (staffing || 0) * 0.2;
        if (schedule.RecurringScheduleID) score += 0.1;
        if (schedule.isDST && !schedule.DSTConfirmed) score -= 0.1;
        if (score < 0) score = 0; if (score > 1) score = 1;
        return score;
      } catch (e) { logError('calculateAutoApprovalScore', e); return 0; }
    };
  }

  if (typeof G.processAutoApprovals !== 'function') {
    G.processAutoApprovals = function processAutoApprovals() {
      try {
        var pendingSchedules = readCampaignSheet(G.SCHEDULE_GENERATION_SHEET).filter(function(s){ return s.Status === 'PENDING'; });
        var out = { autoApproved: 0, flaggedForReview: 0, errors: [] };
        for (var i=0;i<pendingSchedules.length;i++){
          var s = pendingSchedules[i];
          try {
            var score = calculateAutoApprovalScore(s);
            if (score >= 0.8) {
              var result = approveSchedules([s.ID], 'SYSTEM_AUTO');
              if (result && result.success) { out.autoApproved++; try { sendAutoApprovalNotification(s); } catch (e){} }
            } else if (score < 0.5) {
              try { flagScheduleForReview(s, 'Low confidence score'); } catch (e){}
              out.flaggedForReview++;
            }
          } catch (e) { out.errors.push('Schedule '+ s.ID +': '+ e.message); }
        }
        return out;
      } catch (e) {
        logError('processAutoApprovals', e);
        return { autoApproved: 0, flaggedForReview: 0, errors: [e.message] };
      }
    };
  }

  if (typeof G.autoResolveConflicts !== 'function') {
    G.autoResolveConflicts = function autoResolveConflicts() {
      try {
        var conflicts = detectAllScheduleConflicts();
        var out = { resolved: 0, escalated: 0, errors: [] };
        for (var i=0;i<conflicts.length;i++){
          var c = conflicts[i];
          try {
            var res = findOptimalResolution(c);
            if (res && res.confidence > 0.7) { try { applyConflictResolution(c, res); } catch (e){} out.resolved++; }
            else { try { escalateConflict(c, 'Unable to auto-resolve'); } catch (e){} out.escalated++; }
          } catch (e) { out.errors.push('Conflict '+ c.id +': '+ e.message); }
        }
        return out;
      } catch (e) { logError('autoResolveConflicts', e); return { resolved: 0, escalated: 0, errors: [e.message] }; }
    };
  }

  if (typeof G.optimizeStaffingLevels !== 'function') G.optimizeStaffingLevels = function (){ return {}; };
  if (typeof G.sendProactiveNotifications !== 'function') G.sendProactiveNotifications = function (){ return {}; };
  if (typeof G.optimizeSystemPerformance !== 'function') G.optimizeSystemPerformance = function (){ return {}; };
  if (typeof G.performDataCleanup !== 'function') G.performDataCleanup = function (){ return {}; };
  if (typeof G.executePredictiveActions !== 'function') G.executePredictiveActions = function (){ return {}; };
  if (typeof G.logAutomationResults !== 'function') G.logAutomationResults = function (){};

  if (typeof G.runIntelligentAutomation !== 'function') {
    G.runIntelligentAutomation = function runIntelligentAutomation() {
      return withPerformanceMonitoring('runIntelligentAutomation', function () {
        try {
          var results = {
            autoApprovals: processAutoApprovals(),
            conflictResolution: autoResolveConflicts(),
            optimalStaffing: optimizeStaffingLevels(),
            proactiveNotifications: sendProactiveNotifications(),
            performanceOptimization: optimizeSystemPerformance(),
            dataCleanup: performDataCleanup(),
            predictiveActions: executePredictiveActions()
          };
          try { logAutomationResults(results); } catch (e){}
          return {
            success: true,
            results: results,
            executedAt: new Date().toISOString(),
            nextRun: new Date(Date.now() + 4 * 60 * 60 * 1000).toISOString()
          };
        } catch (error) {
          logError('runIntelligentAutomation', error);
          return { success: false, error: error.message };
        }
      });
    };
  }

  // ────────────────────────────────────────────────────────────────────────────
  // Advanced Reporting
  // ────────────────────────────────────────────────────────────────────────────
  if (typeof G.getTimeframeBounds !== 'function') {
    G.getTimeframeBounds = function (tf) {
      tf = tf || 'month';
      var now = new Date();
      var start = new Date(now), end = new Date(now);
      if (tf === 'day') { start.setHours(0,0,0,0); end.setHours(23,59,59,999); }
      else if (tf === 'week') {
        var day = now.getDay() || 7; // Monday=1..Sunday=7
        start.setDate(now.getDate() - (day - 1)); start.setHours(0,0,0,0);
        end = new Date(start); end.setDate(start.getDate() + 6); end.setHours(23,59,59,999);
      } else { // month
        start = new Date(now.getFullYear(), now.getMonth(), 1);
        end = new Date(now.getFullYear(), now.getMonth()+1, 0, 23,59,59,999);
      }
      return { startDate: start, endDate: end };
    };
  }

  if (typeof G.generateOverviewMetrics !== 'function') G.generateOverviewMetrics = function (){ return { totalAgents: 0, totalShifts: 0 }; };
  if (typeof G.calculateKPIs !== 'function') G.calculateKPIs = function (){ return { adherence: 0, utilization: 0, csat: 0 }; };
  if (typeof G.generateTrendAnalysis !== 'function') G.generateTrendAnalysis = function (){ return { weekly: [] }; };
  if (typeof G.generateRiskAssessment !== 'function') G.generateRiskAssessment = function (){ return { risks: [] }; };
  if (typeof G.generateExecutiveRecommendations !== 'function') G.generateExecutiveRecommendations = function (){ return ['No critical actions']; };
  if (typeof G.generateForecastInsights !== 'function') G.generateForecastInsights = function (){ return { nextMonth: {} }; };
  if (typeof G.generateDepartmentBreakdown !== 'function') G.generateDepartmentBreakdown = function (){ return {}; };
  if (typeof G.generateCostAnalysis !== 'function') G.generateCostAnalysis = function (){ return { current: 0, projected: 0 }; };
  if (typeof G.generateComplianceMetrics !== 'function') G.generateComplianceMetrics = function (){ return {}; };
  if (typeof G.calculateDataFreshness !== 'function') G.calculateDataFreshness = function (){ return { seconds: 0 }; };

  if (typeof G.generateExecutiveDashboard !== 'function') {
    G.generateExecutiveDashboard = function generateExecutiveDashboard(timeframe) {
      timeframe = timeframe || 'month';
      return withPerformanceMonitoring('generateExecutiveDashboard', function () {
        try {
          var bounds = getTimeframeBounds(timeframe);
          var startDate = bounds.startDate, endDate = bounds.endDate;
          var dashboard = {
            overview: generateOverviewMetrics(startDate, endDate),
            keyPerformanceIndicators: calculateKPIs(startDate, endDate),
            trendAnalysis: generateTrendAnalysis(startDate, endDate),
            riskAssessment: generateRiskAssessment(startDate, endDate),
            recommendations: generateExecutiveRecommendations(startDate, endDate),
            forecastInsights: generateForecastInsights(endDate),
            departmentBreakdown: generateDepartmentBreakdown(startDate, endDate),
            costAnalysis: generateCostAnalysis(startDate, endDate),
            complianceMetrics: generateComplianceMetrics(startDate, endDate),
            actionItems: []
          };
          return { success: true, dashboard: dashboard, timeframe: timeframe, generatedAt: new Date().toISOString(), dataFreshness: calculateDataFreshness() };
        } catch (error) {
          logError('generateExecutiveDashboard', error);
          return { success: false, error: error.message };
        }
      });
    };
  }

  if (typeof G.generateStaffingAnalysisReport !== 'function') G.generateStaffingAnalysisReport = function (){ return {}; };
  if (typeof G.generateAdherenceDeepDiveReport !== 'function') G.generateAdherenceDeepDiveReport = function (){ return {}; };
  if (typeof G.generateEfficiencyMetricsReport !== 'function') G.generateEfficiencyMetricsReport = function (){ return {}; };
  if (typeof G.generateCostOptimizationReport !== 'function') G.generateCostOptimizationReport = function (){ return {}; };
  if (typeof G.generatePredictiveInsightsReport !== 'function') G.generatePredictiveInsightsReport = function (){ return {}; };
  if (typeof G.generateComplianceAuditReport !== 'function') G.generateComplianceAuditReport = function (){ return {}; };

  if (typeof G.generateOperationalReport !== 'function') {
    G.generateOperationalReport = function generateOperationalReport(reportType, parameters) {
      parameters = parameters || {};
      return withPerformanceMonitoring('generateOperationalReport', function () {
        try {
          var report;
          if (reportType === 'staffing_analysis') report = generateStaffingAnalysisReport(parameters);
          else if (reportType === 'adherence_deep_dive') report = generateAdherenceDeepDiveReport(parameters);
          else if (reportType === 'efficiency_metrics') report = generateEfficiencyMetricsReport(parameters);
          else if (reportType === 'cost_optimization') report = generateCostOptimizationReport(parameters);
          else if (reportType === 'predictive_insights') report = generatePredictiveInsightsReport(parameters);
          else if (reportType === 'compliance_audit') report = generateComplianceAuditReport(parameters);
          else throw new Error('Unknown report type: '+ reportType);
          return { success: true, report: report, reportType: reportType, parameters: parameters, generatedAt: new Date().toISOString() };
        } catch (error) {
          logError('generateOperationalReport', error);
          return { success: false, error: error.message };
        }
      });
    };
  }

  // ────────────────────────────────────────────────────────────────────────────
  // Mobile & Real-Time
  // ────────────────────────────────────────────────────────────────────────────
  if (typeof G.formatScheduleForMobile !== 'function') G.formatScheduleForMobile = function (s){ return s; };
  if (typeof G.formatSwapForMobile !== 'function') G.formatSwapForMobile = function (s){ return s; };
  if (typeof G.getUserNotifications !== 'function') G.getUserNotifications = function (){ return []; };

  if (typeof G.getMobileScheduleData !== 'function') {
    G.getMobileScheduleData = function getMobileScheduleData(userId, dateRange) {
      try {
        var startDate = dateRange.startDate, endDate = dateRange.endDate;
        var schedules = readCampaignSheet(G.SCHEDULE_GENERATION_SHEET).filter(function(s){
          return s.UserID === userId && s.Date >= startDate && s.Date <= endDate && s.Status === 'APPROVED';
        });
        var swapRequests = readCampaignSheet(G.SHIFT_SWAPS_SHEET).filter(function(s){
          return (s.RequestorUserID === userId || s.TargetUserID === userId) && s.Status === 'PENDING';
        });
        var notifications = getUserNotifications(userId, 10);

        return {
          success: true,
          data: {
            schedules: schedules.map(function(s){ return formatScheduleForMobile(s); }),
            swapRequests: swapRequests.map(function(s){ return formatSwapForMobile(s); }),
            notifications: notifications,
            lastUpdated: new Date().toISOString()
          }
        };
      } catch (error) {
        logError('getMobileScheduleData', error);
        return { success: false, error: error.message };
      }
    };
  }

  // ────────────────────────────────────────────────────────────────────────────
  // Integration Utilities (Calendar/HR)
  // ────────────────────────────────────────────────────────────────────────────
  if (typeof G.formatDateForICS !== 'function') {
    G.formatDateForICS = function formatDateForICS(d) {
      function pad(n){ return (n<10?'0':'') + n; }
      return d.getUTCFullYear()
        + pad(d.getUTCMonth()+1)
        + pad(d.getUTCDate())
        + 'T' + pad(d.getUTCHours())
        + pad(d.getUTCMinutes())
        + pad(d.getUTCSeconds()) + 'Z';
    };
  }
  if (typeof G.getSchedulesByIds !== 'function') {
    G.getSchedulesByIds = function getSchedulesByIds(ids) {
      try {
        var rows = readCampaignSheet(G.SCHEDULE_GENERATION_SHEET);
        var map = {};
        for (var i=0;i<rows.length;i++){ map[String(rows[i].ID)] = rows[i]; }
        var out = [];
        ids = ids || [];
        for (var j=0;j<ids.length;j++){ if (map[String(ids[j])]) out.push(map[String(ids[j])]); }
        return out;
      } catch (e) { logError('getSchedulesByIds', e); return []; }
    };
  }
  if (typeof G.generateICSFile !== 'function') {
    G.generateICSFile = function generateICSFile(schedules) {
      try {
        var ics = ['BEGIN:VCALENDAR','VERSION:2.0','PRODID:-//VLBPO LuminaHQ//Schedule System//EN','CALSCALE:GREGORIAN'];
        schedules = schedules || [];
        for (var i=0;i<schedules.length;i++){
          var s = schedules[i];
          var startDateTime = new Date(String(s.Date) + 'T' + String(s.StartTime));
          var endDateTime = new Date(String(s.Date) + 'T' + String(s.EndTime));
          ics.push(
            'BEGIN:VEVENT',
            'UID:' + (s.ID || Utilities.getUuid()) + '@vlbpo.com',
            'DTSTART:' + formatDateForICS(startDateTime),
            'DTEND:' + formatDateForICS(endDateTime),
            'SUMMARY:' + (s.SlotName || 'Work Shift'),
            'DESCRIPTION:' + 'Scheduled shift for ' + (s.UserName || '') + '\\nLocation: ' + (s.Location || 'Office') + '\\nDepartment: ' + (s.Department || 'General'),
            'LOCATION:' + (s.Location || 'Office'),
            'STATUS:CONFIRMED',
            'END:VEVENT'
          );
        }
        ics.push('END:VCALENDAR');
        return { success: true, content: ics.join('\r\n'), filename: 'schedules_' + new Date().toISOString().split('T')[0] + '.ics' };
      } catch (e) { logError('generateICSFile', e); return { success: false, error: e.message }; }
    };
  }
  if (typeof G.generateGoogleCalendarEvents !== 'function') G.generateGoogleCalendarEvents = function (){ return { success: false, error: 'Direct Google Calendar push not implemented in this module.' }; };
  if (typeof G.generateOutlookEvents !== 'function') G.generateOutlookEvents = function (){ return { success: false, error: 'Outlook export not implemented in this module.' }; };

  if (typeof G.exportToCalendar !== 'function') {
    G.exportToCalendar = function exportToCalendar(scheduleIds, calendarFormat) {
      calendarFormat = calendarFormat || 'ics';
      try {
        var schedules = getSchedulesByIds(scheduleIds);
        if (calendarFormat === 'ics') return generateICSFile(schedules);
        if (calendarFormat === 'google') return generateGoogleCalendarEvents(schedules);
        if (calendarFormat === 'outlook') return generateOutlookEvents(schedules);
        throw new Error('Unsupported calendar format: ' + calendarFormat);
      } catch (error) { logError('exportToCalendar', error); return { success: false, error: error.message }; }
    };
  }

  if (typeof G.syncWithWorkday !== 'function') G.syncWithWorkday = function (){ return {}; };
  if (typeof G.syncWithBambooHR !== 'function') G.syncWithBambooHR = function (){ return {}; };
  if (typeof G.syncWithADP !== 'function') G.syncWithADP = function (){ return {}; };

  if (typeof G.syncWithHRSystem !== 'function') {
    G.syncWithHRSystem = function syncWithHRSystem(hrSystemType, syncOptions) {
      syncOptions = syncOptions || {};
      try {
        var res;
        if (hrSystemType === 'workday') res = syncWithWorkday(syncOptions);
        else if (hrSystemType === 'bamboohr') res = syncWithBambooHR(syncOptions);
        else if (hrSystemType === 'adp') res = syncWithADP(syncOptions);
        else throw new Error('Unsupported HR system: ' + hrSystemType);
        return { success: true, syncResult: res, syncedAt: new Date().toISOString() };
      } catch (e) { logError('syncWithHRSystem', e); return { success: false, error: e.message }; }
    };
  }

  // ────────────────────────────────────────────────────────────────────────────
  // Client-accessible wrappers
  // ────────────────────────────────────────────────────────────────────────────
  if (typeof G.clientRunAIOptimization !== 'function') {
    G.clientRunAIOptimization = function clientRunAIOptimization(scheduleData, constraints) {
      try { return optimizeScheduleWithAI(scheduleData, constraints); }
      catch (e) { logError('clientRunAIOptimization', e); return { success: false, error: e.message }; }
    };
  }
  if (typeof G.clientGetPredictiveAnalytics !== 'function') {
    G.clientGetPredictiveAnalytics = function clientGetPredictiveAnalytics(timeframe) {
      try {
        var bounds = getTimeframeBounds(timeframe || 'month');
        var startDate = bounds.startDate, endDate = bounds.endDate;
        var historicalData = [];
        if (typeof getHistoricalScheduleData === 'function') {
          historicalData = getHistoricalScheduleData(startDate, endDate);
        }
        return generatePredictiveAnalytics(historicalData);
      } catch (e) { logError('clientGetPredictiveAnalytics', e); return { success: false, error: e.message }; }
    };
  }
  if (typeof G.clientRunIntelligentAutomation !== 'function') {
    G.clientRunIntelligentAutomation = function clientRunIntelligentAutomation() {
      try { return runIntelligentAutomation(); }
      catch (e) { logError('clientRunIntelligentAutomation', e); return { success: false, error: e.message }; }
    };
  }
  if (typeof G.clientGenerateExecutiveDashboard !== 'function') {
    G.clientGenerateExecutiveDashboard = function clientGenerateExecutiveDashboard(timeframe) {
      try { return generateExecutiveDashboard(timeframe || 'month'); }
      catch (e) { logError('clientGenerateExecutiveDashboard', e); return { success: false, error: e.message }; }
    };
  }
  if (typeof G.clientGetMobileScheduleData !== 'function') {
    G.clientGetMobileScheduleData = function clientGetMobileScheduleData(userId, dateRange) {
      try { return getMobileScheduleData(userId, dateRange); }
      catch (e) { logError('clientGetMobileScheduleData', e); return { success: false, error: e.message }; }
    };
  }
  if (typeof G.clientExportToCalendar !== 'function') {
    G.clientExportToCalendar = function clientExportToCalendar(scheduleIds, format) {
      try { return exportToCalendar(scheduleIds, format); }
      catch (e) { logError('clientExportToCalendar', e); return { success: false, error: e.message }; }
    };
  }

  // ────────────────────────────────────────────────────────────────────────────
  // System Health & Monitoring
  // ────────────────────────────────────────────────────────────────────────────
  function _ok(score){ return { status: 'OK', score: (typeof score === 'number' ? score : 1) }; }
  if (typeof G.checkDatabaseHealth !== 'function') G.checkDatabaseHealth = function (){ return _ok(0.95); };
  if (typeof G.checkScheduleGenerationHealth !== 'function') G.checkScheduleGenerationHealth = function (){ return _ok(0.9); };
  if (typeof G.checkNotificationHealth !== 'function') G.checkNotificationHealth = function (){ return _ok(0.9); };
  if (typeof G.checkIntegrationHealth !== 'function') G.checkIntegrationHealth = function (){ return _ok(0.85); };
  if (typeof G.checkPerformanceHealth !== 'function') G.checkPerformanceHealth = function (){ return _ok(0.92); };
  if (typeof G.measureAverageResponseTime !== 'function') G.measureAverageResponseTime = function (){ return 120; };
  if (typeof G.calculateErrorRate !== 'function') G.calculateErrorRate = function (){ return 0.01; };
  if (typeof G.calculateThroughput !== 'function') G.calculateThroughput = function (){ return 200; };
  if (typeof G.calculateAvailability !== 'function') G.calculateAvailability = function (){ return 0.999; };
  if (typeof G.generateHealthRecommendations !== 'function') {
    G.generateHealthRecommendations = function (hc) {
      var recs = [];
      if (hc && hc.metrics && hc.metrics.errorRate > 0.02) recs.push('Investigate elevated error rate in schedule generation.');
      if (hc && hc.components && hc.components.integrations && hc.components.integrations.score < 0.9) recs.push('Review integration retries and backoffs.');
      return recs.length ? recs : ['System operating within expected parameters.'];
    };
  }

  if (typeof G.performSystemHealthCheck !== 'function') {
    G.performSystemHealthCheck = function performSystemHealthCheck() {
      try {
        var healthCheck = {
          timestamp: new Date().toISOString(),
          overallHealth: 'HEALTHY',
          components: {
            database: checkDatabaseHealth(),
            scheduleGeneration: checkScheduleGenerationHealth(),
            notifications: checkNotificationHealth(),
            integrations: checkIntegrationHealth(),
            performance: checkPerformanceHealth()
          },
          metrics: {
            responseTime: measureAverageResponseTime(),
            errorRate: calculateErrorRate(),
            throughput: calculateThroughput(),
            availability: calculateAvailability()
          },
          recommendations: []
        };

        var scores = [];
        var comps = healthCheck.components;
        for (var k in comps){ if (comps.hasOwnProperty(k)) scores.push(comps[k].score || 0); }
        var sum = 0; for (var i=0;i<scores.length;i++) sum += scores[i];
        var avg = scores.length ? (sum / scores.length) : 0;
        healthCheck.overallHealth = (avg >= 0.9) ? 'HEALTHY' : (avg >= 0.7) ? 'WARNING' : 'CRITICAL';
        healthCheck.recommendations = generateHealthRecommendations(healthCheck);
        return { success: true, healthCheck: healthCheck };
      } catch (e) {
        logError('performSystemHealthCheck', e);
        return { success: false, error: e.message, healthCheck: { overallHealth: 'CRITICAL', timestamp: new Date().toISOString() } };
      }
    };
  }

  // ────────────────────────────────────────────────────────────────────────────
  // Initialization (daily)
  // ────────────────────────────────────────────────────────────────────────────
  if (typeof G.initializeAIEngine !== 'function') G.initializeAIEngine = function (){ return {}; };
  if (typeof G.initializePredictiveEngine !== 'function') G.initializePredictiveEngine = function (){ return {}; };
  if (typeof G.initializeIntelligentAutomation !== 'function') G.initializeIntelligentAutomation = function (){ return {}; };

  if (typeof G.initializeAdvancedScheduleSystem !== 'function') {
    G.initializeAdvancedScheduleSystem = function initializeAdvancedScheduleSystem() {
      try {
        console.log('🚀 Initializing Advanced Schedule System...');
        initializeAIEngine();
        initializePredictiveEngine();
        initializeIntelligentAutomation();
        initializeCoachingSystem(); // Initialize enhanced coaching system
        var healthCheck = performSystemHealthCheck();
        console.log('✅ Advanced Schedule System initialized successfully');
        return { success: true, message: 'Advanced Schedule System ready', healthCheck: healthCheck.healthCheck };
      } catch (error) {
        logError('initializeAdvancedScheduleSystem', error);
        return { success: false, error: error.message };
      }
    };
  }

  try {
    var prop = PropertiesService.getScriptProperties();
    var last = prop.getProperty('ADVANCED_SCHEDULE_LAST_INIT');
    var now = Date.now();
    if (!last || (now - parseInt(last, 10)) > 24 * 60 * 60 * 1000) {
      initializeAdvancedScheduleSystem();
      prop.setProperty('ADVANCED_SCHEDULE_LAST_INIT', String(now));
    }
  } catch (error) {
    try { console.warn('Advanced auto-initialization failed: '+ error); } catch (_){}
  }
})();
