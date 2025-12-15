// Centralized data export dispatcher for cross-domain reporting

function listDataExportOptions(context) {
  const globalScope = (typeof GLOBAL_SCOPE === 'object' && GLOBAL_SCOPE) ? GLOBAL_SCOPE : this;
  const options = [
    {
      key: 'unified-export-sheet',
      title: 'Unified Multi-export Workbook',
      description: 'Compiles attendance, adherence, and export catalog data onto a single, well-formatted sheet.',
      requires: 'exportUnifiedWorkbook'
    },
    {
      key: 'attendance-daily-pivot',
      title: 'Attendance Daily Pivot',
      description: 'Enhanced attendance pivot by user with custom ranges and granular periods.',
      requires: 'clientExecuteDailyPivotExport'
    },
    {
      key: 'attendance-csv',
      title: 'Attendance CSV',
      description: 'Standard attendance CSV export filtered by campaign, agent, and period.',
      requires: 'exportAttendanceCsv'
    },
    {
      key: 'attendance-dashboard',
      title: 'Attendance Dashboard Summary',
      description: 'Summary export of attendance dashboard metrics for a date window.',
      requires: 'exportAttendanceDashboard'
    },
    {
      key: 'attendance-calendar',
      title: 'Attendance Calendar',
      description: 'Calendar-style attendance export for visual schedule coverage.',
      requires: 'exportAttendanceCalendar'
    },
    {
      key: 'attendance-adherence-sheet',
      title: 'Adherence & Compliance Sheet',
      description: 'Generates adherence and compliance workbook for the selected range.',
      requires: 'exportAdherenceComplianceSheet'
    },
    {
      key: 'adherence-data',
      title: 'Adherence Dataset',
      description: 'Returns adherence dataset for downstream CSV/JSON handling.',
      requires: 'getAdherenceComplianceExportData'
    },
    {
      key: 'qa-agent-matrix',
      title: 'QA Agent Matrix',
      description: 'Cross-campaign QA agent performance matrix with filters.',
      requires: 'clientExportAgentMatrix'
    },
    {
      key: 'call-analytics-csv',
      title: 'Call Analytics CSV',
      description: 'Exports call analytics by period and agent.',
      requires: 'exportCallAnalyticsCsv'
    },
    {
      key: 'call-csat-csv',
      title: 'Call CSAT CSV',
      description: 'Exports CSAT summaries by campaign and agent.',
      requires: 'exportCallCsatCsv'
    },
    {
      key: 'call-performance-matrix',
      title: 'Call Performance Matrix',
      description: 'Performance matrix export with KPI columns.',
      requires: 'exportCallPerformanceMatrixCsv'
    }
  ];

  return options.filter(function (opt) {
    return typeof globalScope[opt.requires] === 'function' || typeof this[opt.requires] === 'function';
  }, this);
}

function clientRunDataExport(request) {
  return runDataExport(request || {});
}

function runDataExport(payload) {
  var type = (payload && payload.type) || '';
  var normalized = normalizeExportPayload_(payload);

  switch (type) {
    case 'attendance-daily-pivot':
      return executeDailyPivot_(normalized);
    case 'attendance-csv':
      return exportAttendanceCsv(normalized.granularity, normalized.period, normalized.userId || '', normalized.policyOptions || {});
    case 'attendance-dashboard':
      return exportAttendanceDashboard(normalized.periodType || 'Custom', normalized.startDateIso, normalized.endDateIso);
    case 'attendance-calendar':
      return exportAttendanceCalendar(normalized.periodType || 'Custom', normalized.startDateIso, normalized.endDateIso);
    case 'attendance-adherence-sheet':
      return exportAdherenceComplianceSheet(normalized);
    case 'adherence-data':
      return getAdherenceComplianceExportData({
        startDateIso: normalized.startDateIso,
        endDateIso: normalized.endDateIso,
        periodType: normalized.periodType || normalized.granularity || 'custom',
        users: normalized.userId ? [normalized.userId] : [],
        timezone: normalized.timezone
      });
    case 'qa-agent-matrix':
      return clientExportAgentMatrix({
        granularity: normalized.granularity,
        period: normalized.period,
        agent: normalized.userId,
        campaignId: normalized.campaignId,
        settings: { format: 'xlsx' }
      });
    case 'call-analytics-csv':
      return exportCallAnalyticsCsv(normalized.granularity, normalized.period, normalized.userId || '');
    case 'call-csat-csv':
      return exportCallCsatCsv(normalized.granularity, normalized.period, normalized.userId || '');
    case 'call-performance-matrix':
      return exportCallPerformanceMatrixCsv(normalized.granularity, normalized.period, normalized.userId || '');
    case 'unified-export-sheet':
      return exportUnifiedWorkbook(normalized);
    default:
      return { success: false, error: 'Unsupported export type requested.' };
  }
}

function normalizeExportPayload_(payload) {
  var granularity = (payload.granularity || payload.periodType || 'Week').toString();
  var periodValue = (payload.period || payload.periodId || '').toString();
  var startDateIso = normalizeDateString_(payload.startDate);
  var endDateIso = normalizeDateString_(payload.endDate);

  if (!startDateIso && granularity.toLowerCase() === 'custom') {
    startDateIso = normalizeDateString_(new Date());
  }

  if (!endDateIso && granularity.toLowerCase() === 'custom') {
    endDateIso = normalizeDateString_(new Date());
  }

  return {
    type: payload.type || '',
    campaignId: payload.campaignId || payload.campaign || '',
    userId: payload.userId || payload.agent || '',
    focusType: payload.focusType || '',
    granularity: granularity,
    period: periodValue,
    periodType: granularity,
    startDateIso: startDateIso,
    endDateIso: endDateIso,
    timezone: payload.timezone || (typeof Session !== 'undefined' && Session.getScriptTimeZone ? Session.getScriptTimeZone() : 'America/Jamaica'),
    policyOptions: payload.policyOptions || {}
  };
}

function normalizeDateString_(value) {
  if (!value) return '';
  var d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function executeDailyPivot_(normalized) {
  var periodType = normalized.granularity === 'Custom' ? 'custom' : normalized.granularity;
  var periodPayload = { type: periodType, value: normalized.period };

  if (periodType.toLowerCase() === 'custom') {
    periodPayload.start = normalized.startDateIso;
    periodPayload.end = normalized.endDateIso;
  }

  var params = {
    period: periodPayload,
    users: normalized.userId ? [normalized.userId] : [],
    userSelection: normalized.userId ? 'specific' : 'all',
    dailyPivotOptions: {},
    hourPolicy: normalized.policyOptions || {}
  };

  var result = clientExecuteDailyPivotExport(params);
  if (result && result.success) {
    return {
      success: true,
      url: result.url || result.fileUrl,
      name: result.name || result.fileName,
      message: 'Daily pivot export created.'
    };
  }

  return result || { success: false, error: 'Unable to generate daily pivot export.' };
}

function exportUnifiedWorkbook(normalized) {
  try {
    var timezone = normalized.timezone || (typeof Session !== 'undefined' && Session.getScriptTimeZone ? Session.getScriptTimeZone() : 'America/Jamaica');
    var periodLabel = normalized.period
      || (normalized.startDateIso && normalized.endDateIso ? normalized.startDateIso + ' → ' + normalized.endDateIso : normalized.granularity);
    var name = 'Unified Export - ' + Utilities.formatDate(new Date(), timezone, 'yyyyMMdd_HHmm');
    var ss = SpreadsheetApp.create(name);
    var sheet = ss.getActiveSheet();
    sheet.setName('Unified Export');

    var currentRow = 1;
    currentRow = writeUnifiedHeader_(sheet, currentRow, normalized, periodLabel);
    currentRow = appendAttendanceSection_(sheet, currentRow, normalized, periodLabel);
    currentRow = appendAdherenceSection_(sheet, currentRow, normalized);
    currentRow = appendExportCatalogSection_(sheet, currentRow);

    sheet.autoResizeColumns(1, Math.max(1, sheet.getLastColumn()));

    return {
      success: true,
      url: ss.getUrl(),
      name: ss.getName(),
      message: 'Unified workbook created with attendance, adherence, and export catalog insights.'
    };
  } catch (error) {
    console.error('exportUnifiedWorkbook failed:', error);
    return { success: false, error: error && error.message ? error.message : 'Unable to create unified export workbook.' };
  }
}

function writeUnifiedHeader_(sheet, startRow, normalized, periodLabel) {
  var headerValues = [
    ['Unified Data Export', ''],
    ['Period', periodLabel],
    ['Granularity', normalized.granularity],
    ['Campaign', normalized.campaignId || 'All campaigns'],
    ['User filter', normalized.userId || 'All users']
  ];

  var headerRange = sheet.getRange(startRow, 1, headerValues.length, 2);
  headerRange.setValues(headerValues);
  headerRange.setFontSize(11);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#eef2ff');
  headerRange.setBorder(true, true, true, true, true, true);
  headerRange.setHorizontalAlignment('left');

  sheet.getRange(startRow, 1).setFontSize(14);
  sheet.getRange(startRow, 1, 1, 2).mergeAcross();

  return startRow + headerValues.length + 2;
}

function appendAttendanceSection_(sheet, startRow, normalized, periodLabel) {
  var analytics = {};
  try {
    analytics = getAttendanceAnalyticsByPeriod(normalized.granularity, normalized.period, normalized.userId || '', normalized.policyOptions || {}) || {};
  } catch (attendanceError) {
    console.error('appendAttendanceSection_ attendanceError:', attendanceError);
  }

  var userCompliance = Array.isArray(analytics.userCompliance) ? analytics.userCompliance : [];
  var rows = userCompliance.map(function (user) {
    var compliance = typeof calculateComplianceScore === 'function' ? calculateComplianceScore(user) : '';
    var baseBillableSecs = Number.isFinite(user.baseBillableSecs) ? user.baseBillableSecs : 0;
    var breakCreditSecs = Number.isFinite(user.breakCreditSecs) ? user.breakCreditSecs : 0;
    var lunchAdjustmentSecs = Number.isFinite(user.lunchAdjustmentSecs) ? user.lunchAdjustmentSecs : 0;
    var adjustedBillableSecs = Number.isFinite(user.adjustedBillableSecs)
      ? user.adjustedBillableSecs
      : Math.max(0, baseBillableSecs + breakCreditSecs + lunchAdjustmentSecs);

    var formatHours = function (secs) {
      var hours = Number.isFinite(secs) ? secs / 3600 : 0;
      return (Math.abs(hours) < 0.005 ? 0 : hours).toFixed(2);
    };

    var formatMinutes = function (secs) {
      var minutes = Number.isFinite(secs) ? secs / 60 : 0;
      return (Math.abs(minutes) < 0.005 ? 0 : minutes).toFixed(2);
    };

    return [
      user.user,
      formatHours(adjustedBillableSecs),
      compliance,
      formatMinutes(user.breakOverSecs),
      formatMinutes(user.lunchOverSecs)
    ];
  });

  var sectionTitle = 'Attendance summary (' + (periodLabel || normalized.granularity) + ')';
  var headers = ['User', 'Adjusted Billable Hours', 'Compliance Score', 'Break Over (mins)', 'Lunch Over (mins)'];

  return writeUnifiedSection_(sheet, startRow, sectionTitle, headers, rows, 'Attendance performance spotlight');
}

function appendAdherenceSection_(sheet, startRow, normalized) {
  var dataset = {};
  try {
    dataset = getAdherenceComplianceExportData({
      startDateIso: normalized.startDateIso,
      endDateIso: normalized.endDateIso,
      periodType: normalized.periodType || normalized.granularity,
      users: normalized.userId ? [normalized.userId] : [],
      timezone: normalized.timezone
    }) || {};
  } catch (adherenceError) {
    console.error('appendAdherenceSection_ adherenceError:', adherenceError);
  }

  var rows = Array.isArray(dataset.rows) ? dataset.rows : [];
  var compactRows = rows.map(function (entry) {
    return [
      entry.user,
      entry.average,
      entry.compliantDays,
      entry.flaggedDays,
      entry.overtimeDays,
      entry.daysWorked
    ];
  });

  var headers = ['User', 'Average %', 'Compliant Days', 'Flagged Days', 'Overtime Days', 'Days Worked'];
  return writeUnifiedSection_(sheet, startRow, 'Adherence scorecard', headers, compactRows, 'Daily adherence heatmap summary');
}

function appendExportCatalogSection_(sheet, startRow) {
  var options = listDataExportOptions({}) || [];
  var rows = options.map(function (opt) {
    return [opt.title, opt.description, opt.key];
  });
  var headers = ['Export', 'Description', 'Key'];
  return writeUnifiedSection_(sheet, startRow, 'Export catalog', headers, rows, 'Reference for every export available in the hub');
}

function writeUnifiedSection_(sheet, startRow, title, headers, rows, subtitle) {
  var columnCount = Math.max(1, headers.length);
  var titleRow = [padRowToColumns_(title, columnCount)];
  var titleRange = sheet.getRange(startRow, 1, 1, columnCount);
  titleRange.setValues(titleRow);
  titleRange.mergeAcross();
  titleRange.setFontSize(12).setFontWeight('bold').setBackground('#e0f2fe');

  if (subtitle) {
    var subtitleRange = sheet.getRange(startRow + 1, 1, 1, columnCount);
    subtitleRange.setValues([padRowToColumns_(subtitle, columnCount)]);
    subtitleRange.mergeAcross();
    subtitleRange.setFontColor('#475569');
  }

  var headerRange = sheet.getRange(startRow + 2, 1, 1, columnCount);
  headerRange.setValues([padRowToColumns_(headers, columnCount)]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1f4e79');
  headerRange.setFontColor('#ffffff');

  var dataRowCount = Math.max(1, rows.length);
  var dataRange = sheet.getRange(startRow + 3, 1, dataRowCount, columnCount);
  if (rows.length) {
    var normalizedRows = rows.map(function (row) { return padRowToColumns_(row, columnCount); });
    dataRange.setValues(normalizedRows);
  } else {
    dataRange.setValues([padRowToColumns_('No data available for this period', columnCount)]);
  }

  dataRange.setHorizontalAlignment('center');
  dataRange.setBorder(true, true, true, true, true, true);

  return startRow + 3 + dataRowCount + 2;
}

function padRowToColumns_(rowOrValue, columnCount) {
  var row = Array.isArray(rowOrValue) ? rowOrValue.slice(0, columnCount) : [rowOrValue];
  while (row.length < columnCount) {
    row.push('');
  }
  return row;
}
