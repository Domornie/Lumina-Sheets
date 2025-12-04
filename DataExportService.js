// Centralized data export dispatcher for cross-domain reporting

function listDataExportOptions(context) {
  const globalScope = (typeof GLOBAL_SCOPE === 'object' && GLOBAL_SCOPE) ? GLOBAL_SCOPE : this;
  const options = [
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
      return getAdherenceComplianceExportData(normalized.startDateIso, normalized.endDateIso, normalized.userId ? [normalized.userId] : [], normalized.timezone);
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
