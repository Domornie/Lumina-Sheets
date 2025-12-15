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

    sheet.autoResizeColumns(1, Math.max(1, sheet.getLastColumn()));

    return {
      success: true,
      url: ss.getUrl(),
      name: ss.getName(),
      message: 'Unified workbook created with attendance dashboard and adherence exports.'
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
  var table = buildAttendanceDashboardTable_(normalized.periodType || normalized.granularity, normalized.startDateIso, normalized.endDateIso);
  var subtitle = table.subtitle || ('Period: ' + (periodLabel || normalized.granularity));
  return writeUnifiedSection_(sheet, startRow, 'Attendance Dashboard Summary', table.headers, table.rows, subtitle);
}

function appendAdherenceSection_(sheet, startRow, normalized) {
  var table = buildAdherenceComplianceTable_({
    startDateIso: normalized.startDateIso,
    endDateIso: normalized.endDateIso,
    periodType: normalized.periodType || normalized.granularity,
    users: normalized.userId ? [normalized.userId] : [],
    timezone: normalized.timezone
  });

  return writeUnifiedSection_(sheet, startRow, 'Adherence & Compliance Export', table.headers, table.rows, table.subtitle);
}

function buildAttendanceDashboardTable_(periodType, startDate, endDate) {
  var range = typeof normalizeDateRangeForExport === 'function'
    ? normalizeDateRangeForExport(periodType, startDate, endDate)
    : null;

  if (!range) {
    return { headers: [], rows: [], subtitle: 'Period unavailable' };
  }

  var attendanceData = readScheduleSheet(ATTENDANCE_STATUS_SHEET) || [];
  var displayNameMap = typeof buildUserDisplayNameMap === 'function' ? buildUserDisplayNameMap() : new Map();
  var workingDates = typeof getWeekdayIsoDatesInRange === 'function' ? getWeekdayIsoDatesInRange(range.start, range.end) : [];
  var totalWorkingDays = workingDates.length || 1;

  var positiveStatuses = new Set(['present', 'punctual', 'training']);
  var negativeStatuses = new Set([
    'absent',
    'bereavement',
    'late',
    'no call no show',
    'no call/no show',
    'no call/no-show',
    'vacation',
    'sick',
    'sick leave',
    'leave of absent',
    'leave of absence',
    'maternity leave',
    'personal leave'
  ]);
  var normalizeStatus = function (status) { return (status || '').toString().trim().toLowerCase(); };

  var filtered = attendanceData.filter(function (record) {
    var iso = record && record.Date ? toIsoDateString(record.Date) : null;
    return iso && isWeekdayIsoDate(iso) && iso >= range.startIso && iso <= range.endIso;
  });

  var userMap = new Map();
  filtered.forEach(function (record) {
    var user = record.UserName || record.User || record.userName;
    if (user) {
      var fullName = (record.FullName || record.fullName || '').toString().trim();
      if (fullName && !displayNameMap.has(user)) {
        displayNameMap.set(user, fullName);
      }
      userMap.set(user, user);
    }
  });

  var headers = [
    'Agent', 'Period', 'Total Days', 'Present/Worked', 'Absent', 'Sick', 'Vacation', 'Holiday', 'Late', 'On-Time',
    'Absent %', 'Late %', 'On-Time %', 'Sick %', 'Vacation %', 'Attendance Score %'
  ];

  var rows = Array.from(userMap.keys()).sort().map(function (user) {
    var stats = filtered.filter(function (r) { return (r.UserName || r.User || r.userName) === user; });
    var dailyStatuses = new Map();

    stats.forEach(function (record) {
      var iso = record && record.Date ? toIsoDateString(record.Date) : null;
      if (!iso || workingDates.indexOf(iso) === -1) return;
      dailyStatuses.set(iso, record);
    });

    var missingDays = Math.max(0, totalWorkingDays - dailyStatuses.size);
    var totals = {
      present: 0,
      late: 0,
      absent: missingDays,
      sick: 0,
      vacation: 0,
      holiday: 0
    };

    Array.from(dailyStatuses.values()).forEach(function (record) {
      var rawStatus = record.Status || record.status || '';
      var status = normalizeStatus(rawStatus);

      if (positiveStatuses.has(status)) {
        totals.present += 1;
      } else if (negativeStatuses.has(status)) {
        if (status === 'late') {
          totals.late += 1;
        } else if (status === 'vacation') {
          totals.vacation += 1;
        } else if (status === 'sick' || status === 'sick leave') {
          totals.sick += 1;
        } else {
          totals.absent += 1;
        }
      } else if (status === 'holiday') {
        totals.holiday += 1;
      }
    });

    var totalDays = totalWorkingDays;
    var negativeImpactDays = totals.late + totals.absent + totals.sick + totals.vacation;
    var onTime = totals.present;
    var pct = function (count) { return totalDays > 0 ? Math.round((count / totalDays) * 10000) / 100 : 0; };
    var attendanceScore = pct(Math.max(0, totalDays - negativeImpactDays));

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

  return {
    headers: headers,
    rows: rows,
    subtitle: 'Period: ' + range.label
  };
}

function buildAdherenceComplianceTable_(payload) {
  var data = {};
  try {
    data = getAdherenceComplianceExportData(payload || {});
  } catch (err) {
    console.error('buildAdherenceComplianceTable_ error:', err);
    data = { success: false };
  }

  if (!data || data.success === false) {
    return { headers: [], rows: [], subtitle: 'No adherence data available' };
  }

  var headers = ['Agent', 'Period'];
  var dayLabels = (data.dateKeys || []).map(function (day) {
    var dt = normalizeDateValue(day);
    return dt instanceof Date && !isNaN(dt.getTime())
      ? Utilities.formatDate(dt, ATTENDANCE_TIMEZONE, 'EEE MMM d')
      : day;
  });
  headers = headers.concat(dayLabels, ['Total %', 'Low Days', 'Overtime Days', 'Average %', 'Days Worked', 'Status']);

  var rows = (data.rows || []).map(function (row) {
    var status = row.average >= 100
      ? 'Excellent'
      : row.average >= data.goal
        ? 'Pass'
        : row.average >= 80
          ? 'Needs Attention'
          : 'Non-Compliant';

    var dayValues = (row.dayPercents || []).map(function (val) { return typeof val === 'number' ? val : ''; });
    return [row.user]
      .concat((data.startDateIso || '') + ' to ' + (data.endDateIso || ''))
      .concat(dayValues)
      .concat([
        row.totalPercent,
        row.flaggedDays,
        row.overtimeDays,
        row.average,
        row.daysWorked,
        status
      ]);
  });

  var totalPercentSum = (data.rows || []).reduce(function (sum, r) {
    return sum + (Number.isFinite(r.totalPercent) ? r.totalPercent : 0);
  }, 0);

  var dailyAverages = dayLabels.map(function (_, index) {
    var values = (data.rows || [])
      .map(function (r) { return typeof (r.dayPercents || [])[index] === 'number' ? r.dayPercents[index] : null; })
      .filter(function (v) { return typeof v === 'number'; });
    if (!values.length) return '';
    return Math.round((values.reduce(function (a, b) { return a + b; }, 0) / values.length) * 100) / 100;
  });

  var dailyAverageRow = ['Daily Averages']
    .concat((data.startDateIso || '') + ' to ' + (data.endDateIso || ''))
    .concat(dailyAverages)
    .concat([
      dailyAverages.filter(function (v) { return typeof v === 'number'; }).reduce(function (sum, v) { return sum + v; }, 0) || '',
      '',
      '',
      dailyAverages.filter(function (v) { return typeof v === 'number'; }).length
        ? Math.round((dailyAverages.filter(function (v) { return typeof v === 'number'; }).reduce(function (a, b) { return a + b; }, 0) / dailyAverages.filter(function (v) { return typeof v === 'number'; }).length) * 100) / 100
        : '',
      '',
      ''
    ]);

  var summaryRow = ['Totals / Averages']
    .concat((data.startDateIso || '') + ' to ' + (data.endDateIso || ''))
    .concat(new Array(dayLabels.length).fill(''))
    .concat([
      totalPercentSum,
      data.summary ? data.summary.totalFlaggedDays : '',
      data.summary ? data.summary.totalOvertimeDays : '',
      data.summary ? data.summary.averageScore : '',
      data.summary ? data.summary.totalWorkedDays : '',
      data.summary && data.summary.averageScore >= data.goal ? 'Pass' : 'Needs Attention'
    ]);

  if (rows.length) {
    rows.push(dailyAverageRow);
  }
  rows.push(summaryRow);

  return {
    headers: headers,
    rows: rows,
    subtitle: 'Period: ' + ((data.startDateIso || '') + ' to ' + (data.endDateIso || ''))
  };
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
