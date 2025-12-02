/**
 * PayrollManagementService
 *
 * Provides live payroll data for the Payroll Management System UI by reading
 * directly from existing Sheets without creating or modifying schema.
 *
 * Sheets consumed (if present):
 * - "PayrollData": granular weekly payroll records
 * - "PayrollDiscrepancies": logged discrepancy items
 */

/** @OnlyCurrentDoc */

function getPayrollManagementData(options) {
  options = options || {};
  var targetCampaigns = [];

  try {
    if (typeof getCampaignConfigs === 'function') {
      var configuredCampaigns = getCampaignConfigs();
      if (Array.isArray(configuredCampaigns) && configuredCampaigns.length) {
        targetCampaigns = configuredCampaigns.filter(function (cfg) {
          if (!cfg) return false;
          if (cfg.isCatalogOnly && !cfg.utilitiesFileId && !cfg.scheduleFileId) {
            return false; // skip catalog placeholders until configured
          }
          return true;
        });

        if (!targetCampaigns.length) {
          targetCampaigns = configuredCampaigns;
        }
      }
    }
  } catch (err) {
    try { console.warn('Unable to load campaign configs for payroll:', err); } catch (_) {}
  }

  // Filter by requested campaign(s) when provided
  if (targetCampaigns.length && (options.campaignId || options.campaignIds)) {
    var requested = options.campaignIds || [options.campaignId];
    var requestedSet = {};
    (requested || []).forEach(function (cid) { if (cid) requestedSet[String(cid).trim().toLowerCase()] = true; });
    targetCampaigns = targetCampaigns.filter(function (cfg) {
      return requestedSet[String((cfg && cfg.id) || '').trim().toLowerCase()];
    });
  }

  var payrollRows = [];
  var discrepancyRows = [];

  function appendFromSpreadsheet(ss, campaignId) {
    if (!ss) return;
    try {
      var payrollSheet = ss.getSheetByName('PayrollData');
      var discrepancySheet = ss.getSheetByName('PayrollDiscrepancies');
      var payroll = payrollSheet ? readSheetAsObjects_(payrollSheet) : [];
      var discrepancies = discrepancySheet ? readSheetAsObjects_(discrepancySheet) : [];

      if (campaignId) {
        payroll = payroll.map(function (row) { row.CampaignID = row.CampaignID || campaignId; return row; });
        discrepancies = discrepancies.map(function (row) { row.CampaignID = row.CampaignID || campaignId; return row; });
      }

      payrollRows = payrollRows.concat(payroll);
      discrepancyRows = discrepancyRows.concat(discrepancies);
    } catch (error) {
      try { console.warn('Payroll spreadsheet read failed for', campaignId || 'default', error); } catch (_) {}
      if (typeof safeWriteError === 'function') {
        try { safeWriteError('getPayrollManagementData.read', error); } catch (_) {}
      }
    }
  }

  if (targetCampaigns.length) {
    targetCampaigns.forEach(function (campaign) {
      var ss = null;
      try {
        if (typeof getCampaignScheduleSpreadsheet === 'function') {
          ss = getCampaignScheduleSpreadsheet(campaign.id || campaign.ID || campaign.key);
        }
      } catch (err) {
        try { console.warn('Campaign schedule lookup failed for payroll', campaign && campaign.id, err); } catch (_) {}
      }

      if (!ss) {
        try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (_) {}
      }

      appendFromSpreadsheet(ss, campaign.id || campaign.ID || campaign.key || '');
    });
  } else {
    // Legacy single-spreadsheet behavior
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    appendFromSpreadsheet(ss, '');
  }

  var weeklyPeriods = buildWeeklyPeriods_(payrollRows);
  var periodIndex = indexPeriodsById_(weeklyPeriods);

  return {
    periods: {
      weekly: weeklyPeriods,
      biweekly: buildBiweeklyPeriods_(weeklyPeriods),
      monthly: buildMonthlyPeriods_(weeklyPeriods)
    },
    payroll: buildPayrollRecords_(payrollRows, periodIndex),
    discrepancies: buildDiscrepancyLog_(discrepancyRows)
  };
}

function readSheetAsObjects_(sheet) {
  var range = sheet.getDataRange();
  var values = range.getValues();
  if (values.length < 2) return [];

  var headers = values[0].map(function (h) { return String(h || '').trim(); });
  return values.slice(1).filter(function (row) {
    return row.some(function (cell) { return cell !== '' && cell !== null; });
  }).map(function (row) {
    var obj = {};
    headers.forEach(function (header, idx) {
      obj[header] = row[idx];
    });
    return obj;
  });
}

function buildWeeklyPeriods_(rows) {
  var unique = {};
  rows.forEach(function (row) {
    var period = extractPeriodFromRow_(row);
    if (!period.id) return;

    if (!unique[period.id]) {
      unique[period.id] = {
        id: period.id,
        start: period.start ? formatDate_(period.start) : null,
        end: period.end ? formatDate_(period.end) : null,
        label: buildPeriodLabel_(period.id, period.start, period.end)
      };
    }
  });

  return Object.keys(unique).sort().map(function (id) { return unique[id]; });
}

function indexPeriodsById_(periods) {
  var index = {};
  periods.forEach(function (period) { index[period.id] = period; });
  return index;
}

function buildBiweeklyPeriods_(weeklyPeriods) {
  var sorted = weeklyPeriods.slice();
  sorted.sort(function (a, b) { return comparePeriodStart_(a, b); });
  var biweekly = [];

  for (var i = 0; i < sorted.length; i += 2) {
    var first = sorted[i];
    var second = sorted[i + 1];
    if (!first || !second) break;

    var id = first.id + '+' + second.id;
    biweekly.push({
      id: id,
      label: buildCombinedLabel_(first, second, 'Bi-weekly'),
      weekIds: [first.id, second.id],
      start: first.start,
      end: second.end
    });
  }

  return biweekly;
}

function buildMonthlyPeriods_(weeklyPeriods) {
  var months = {};
  weeklyPeriods.forEach(function (period) {
    var start = normalizeDate_(period.start) || normalizeDate_(period.end);
    if (!start) return;
    var key = start.getFullYear() + '-' + pad2_(start.getMonth() + 1);
    if (!months[key]) {
      months[key] = { id: key, label: formatMonth_(start), weekIds: [] };
    }
    months[key].weekIds.push(period.id);
  });

  return Object.keys(months).sort().map(function (id) { return months[id]; });
}

function buildPayrollRecords_(rows, periodIndex) {
  var grouped = {};

  rows.forEach(function (row) {
    var agentId = pickFirstValue_(row, ['AgentId', 'Agent ID', 'UserID', 'User Id', 'ID', 'Employee ID', 'EmployeeID']);
    var agentName = pickFirstValue_(row, ['AgentName', 'Agent Name', 'User', 'Name', 'Full Name', 'Employee Name']);
    var campaign = pickFirstValue_(row, ['Campaign', 'Campaign Name', 'Program', 'Client', 'Account', 'Line of Business']);
    var period = extractPeriodFromRow_(row);
    if (!agentId || !period.id) return;

    var record = grouped[agentId] || (grouped[agentId] = {
      agentId: String(agentId),
      agentName: agentName ? String(agentName) : '',
      campaign: campaign ? String(campaign) : '',
      hourlyRate: toNumber_(pickFirstValue_(row, ['HourlyRate', 'Hourly Rate', 'Rate', 'Pay Rate', 'Base Rate'])) || 0,
      weeklyRecords: []
    });

    record.weeklyRecords.push({
      weekId: period.id,
      regularHours: toNumber_(pickFirstValue_(row, ['RegularHours', 'Regular Hours', 'Regular'])) || 0,
      overtimeHours: toNumber_(pickFirstValue_(row, ['OvertimeHours', 'Overtime Hours', 'OT'])) || 0,
      sickHours: toNumber_(pickFirstValue_(row, ['SickHours', 'Sick Hours', 'Sick'])) || 0,
      vacationHours: toNumber_(pickFirstValue_(row, ['VacationHours', 'Vacation Hours', 'PTO'])) || 0,
      holidayHours: toNumber_(pickFirstValue_(row, ['HolidayHours', 'Holiday Hours'])) || 0,
      leaveHours: toNumber_(pickFirstValue_(row, ['LeaveHours', 'Leave Hours', 'LOA'])) || 0,
      lateOccurrences: toNumber_(pickFirstValue_(row, ['LateOccurrences', 'Late Occurrences', 'Late'])) || 0,
      holidayName: pickFirstValue_(row, ['HolidayName', 'Holiday Name']) || null,
      holidayMultiplier: toNumber_(pickFirstValue_(row, ['HolidayMultiplier', 'Holiday Multiplier', 'Multiplier'])) || 0,
      start: periodIndex[period.id] ? periodIndex[period.id].start : (period.start ? formatDate_(period.start) : null),
      end: periodIndex[period.id] ? periodIndex[period.id].end : (period.end ? formatDate_(period.end) : null)
    });
  });

  return Object.keys(grouped).map(function (key) {
    var record = grouped[key];
    record.weeklyRecords.sort(function (a, b) { return comparePeriodId_(a.weekId, b.weekId); });
    if (!record.campaign) {
      record.campaign = 'Unassigned';
    }
    return record;
  });
}

function extractPeriodFromRow_(row) {
  var rawId = pickFirstValue_(row, ['WeekId', 'Week ID', 'Week', 'PeriodId', 'Period ID', 'PayPeriod', 'Pay Period', 'Pay Period ID']);
  var start = normalizeDate_(pickFirstValue_(row, ['Start', 'StartDate', 'Start Date', 'Week Start', 'Period Start', 'Pay Period Start']));
  var end = normalizeDate_(pickFirstValue_(row, ['End', 'EndDate', 'End Date', 'Week End', 'Period End', 'Pay Period End']));

  if (!rawId) {
    var parsed = parsePeriodLabel_(pickFirstValue_(row, ['Period', 'Payroll Period', 'Pay Period Label']));
    if (parsed) {
      rawId = parsed.id;
      start = start || parsed.start;
      end = end || parsed.end;
    }
  }

  var id = rawId ? String(rawId) : null;
  if (!id && (start || end)) {
    var anchor = start || end;
    id = formatDate_(anchor);
  }

  return { id: id, start: start, end: end };
}

function parsePeriodLabel_(value) {
  if (!value) return null;
  var text = String(value);

  var rangeMatch = text.match(/(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}).*?(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/);
  if (rangeMatch) {
    var start = normalizeDate_(rangeMatch[1]);
    var end = normalizeDate_(rangeMatch[2]);
    var id = start ? formatDate_(start) : text;
    return { id: id, start: start, end: end };
  }

  var isoRange = text.match(/(\d{4}-\d{2}-\d{2}).*?(\d{4}-\d{2}-\d{2})/);
  if (isoRange) {
    var startIso = normalizeDate_(isoRange[1]);
    var endIso = normalizeDate_(isoRange[2]);
    return { id: isoRange[1], start: startIso, end: endIso };
  }

  return { id: text, start: null, end: null };
}

function buildDiscrepancyLog_(rows) {
  return rows.map(function (row, idx) {
    return {
      id: pickFirstValue_(row, ['ID', 'Id', 'Ticket', 'Reference']) || 'DISC-' + pad3_(idx + 1),
      createdAt: normalizeDate_(pickFirstValue_(row, ['CreatedAt', 'Created At', 'LoggedAt', 'Timestamp'])) || null,
      agentId: pickFirstValue_(row, ['AgentId', 'Agent ID', 'UserID', 'User Id', 'ID']) || '',
      issue: pickFirstValue_(row, ['Issue', 'Subject', 'Title']) || '',
      type: pickFirstValue_(row, ['Type', 'Category']) || '',
      hours: toNumber_(pickFirstValue_(row, ['Hours', 'ImpactedHours', 'Impacted Hours'])) || 0,
      status: pickFirstValue_(row, ['Status', 'State']) || '',
      clientApproved: toBoolean_(pickFirstValue_(row, ['ClientApproved', 'Client Approved', 'Approved'])) || false,
      notes: pickFirstValue_(row, ['Notes', 'Comments', 'Details']) || ''
    };
  }).map(function (row) {
    row.createdAt = row.createdAt ? row.createdAt.toISOString() : null;
    return row;
  });
}

function pickFirstValue_(row, keys) {
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    if (Object.prototype.hasOwnProperty.call(row, key)) {
      return row[key];
    }
  }
  return null;
}

function toNumber_(value) {
  var num = Number(value);
  return isNaN(num) ? null : num;
}

function toBoolean_(value) {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value !== 0;
  if (typeof value === 'string') {
    var normalized = value.trim().toLowerCase();
    if (!normalized) return false;
    return ['true', 'yes', 'y', '1', 'approved'].indexOf(normalized) !== -1;
  }
  return false;
}

function normalizeDate_(value) {
  if (!value) return null;
  if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
  var parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function formatDate_(date) {
  return [date.getFullYear(), pad2_(date.getMonth() + 1), pad2_(date.getDate())].join('-');
}

function formatMonth_(date) {
  var monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  return monthNames[date.getMonth()] + ' ' + date.getFullYear();
}

function buildPeriodLabel_(weekId, start, end) {
  if (start && end) {
    return formatDate_(start) + ' - ' + formatDate_(end);
  }
  if (start) {
    return weekId + ' (' + formatDate_(start) + ')';
  }
  return String(weekId);
}

function buildCombinedLabel_(first, second, prefix) {
  var start = first && first.start ? first.start : null;
  var end = second && second.end ? second.end : null;
  if (start && end) {
    return prefix + ': ' + start + ' - ' + end;
  }
  return prefix + ' ' + (first.id + ' / ' + second.id);
}

function pad2_(value) {
  return ('0' + value).slice(-2);
}

function pad3_(value) {
  return ('00' + value).slice(-3);
}

function comparePeriodStart_(a, b) {
  var aDate = normalizeDate_(a.start) || normalizeDate_(a.end);
  var bDate = normalizeDate_(b.start) || normalizeDate_(b.end);
  if (aDate && bDate) return aDate.getTime() - bDate.getTime();
  return String(a.id).localeCompare(String(b.id));
}

function comparePeriodId_(a, b) {
  return String(a).localeCompare(String(b));
}
