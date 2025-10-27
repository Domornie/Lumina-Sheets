/**
 * CollaborationReportingService.gs
 * Lightweight data loader for the Collaboration & Reporting Hub.
 * Reads directly from the core campaign sheets so the client can
 * render a clear snapshot without relying on additional services.
 */
function clientGetCollaborationReportingData() {
  var ss;
  if (typeof getIBTRSpreadsheet === 'function') {
    ss = getIBTRSpreadsheet();
  } else {
    ss = SpreadsheetApp.getActive();
  }

  var response = {
    generatedAt: new Date().toISOString(),
    qa: collabLoadQaSummary_(ss),
    attendance: collabLoadAttendanceSummary_(ss),
    coaching: collabLoadCoachingSummary_(ss),
    escalations: collabLoadEscalationSummary_(ss)
  };

  return response;
}

function collabLoadQaSummary_(ss) {
  var sheetName = (typeof QA_COLLAB_RECORDS !== 'undefined') ? QA_COLLAB_RECORDS : 'QACollab';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return {
      summary: {
        totalAudits: 0,
        avgScore: null,
        feedbackShared: 0,
        uniqueAgents: 0
      },
      recent: []
    };
  }

  var values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) {
    return {
      summary: {
        totalAudits: 0,
        avgScore: null,
        feedbackShared: 0,
        uniqueAgents: 0
      },
      recent: []
    };
  }

  var headers = values[0];
  var rows = values.slice(1);
  var timeZone = Session.getScriptTimeZone();

  var scoreSum = 0;
  var scoreCount = 0;
  var feedbackShared = 0;
  var agentSet = {};
  var recent = [];

  rows.forEach(function (row) {
    var record = collabMapRow_(headers, row);
    var score = collabNormalizeScore_(record.Percentage, record.TotalScore);
    if (score !== null) {
      scoreSum += score;
      scoreCount += 1;
    }
    if (String(record.FeedbackShared || '').toLowerCase() === 'yes') {
      feedbackShared += 1;
    }
    if (record.AgentName) {
      agentSet[String(record.AgentName)] = true;
    }

    recent.push({
      agent: record.AgentName || '',
      client: record.ClientName || '',
      score: score,
      auditDate: collabFormatDate_(record.AuditDate, timeZone, true),
      feedbackShared: String(record.FeedbackShared || ''),
      notes: record.Notes || record.OverallFeedback || ''
    });
  });

  recent.sort(function (a, b) {
    var da = a.auditDate ? new Date(a.auditDate).getTime() : 0;
    var db = b.auditDate ? new Date(b.auditDate).getTime() : 0;
    return db - da;
  });
  recent = recent.slice(0, 15);

  return {
    summary: {
      totalAudits: rows.length,
      avgScore: scoreCount ? collabRound_(scoreSum / scoreCount, 1) : null,
      feedbackShared: feedbackShared,
      uniqueAgents: Object.keys(agentSet).length
    },
    recent: recent
  };
}

function collabLoadAttendanceSummary_(ss) {
  var sheetName = (typeof ATTENDANCE === 'undefined') ? 'AttendanceLog' : ATTENDANCE;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return {
      summary: {
        entries: 0,
        usersTracked: 0,
        totalHours: 0
      },
      recent: [],
      states: []
    };
  }

  var values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) {
    return {
      summary: {
        entries: 0,
        usersTracked: 0,
        totalHours: 0
      },
      recent: [],
      states: []
    };
  }

  var headers = values[0];
  var rows = values.slice(1);
  var timeZone = Session.getScriptTimeZone();
  var durationIndex = headers.indexOf('DurationMin');
  var stateIndex = headers.indexOf('State');
  var userIndex = headers.indexOf('User');
  var timestampIndex = headers.indexOf('Timestamp');

  var totalsByState = {};
  var userSet = {};
  var totalMinutes = 0;
  var recent = [];

  rows.forEach(function (row) {
    var durationMin = (durationIndex > -1) ? Number(row[durationIndex]) : 0;
    if (!isFinite(durationMin)) durationMin = 0;
    totalMinutes += durationMin;

    var state = (stateIndex > -1) ? String(row[stateIndex] || '') : '';
    if (state) {
      if (!totalsByState[state]) totalsByState[state] = 0;
      totalsByState[state] += durationMin;
    }

    var user = (userIndex > -1) ? String(row[userIndex] || '') : '';
    if (user) userSet[user] = true;

    recent.push({
      user: user,
      state: state,
      duration: durationMin,
      timestamp: collabFormatDate_(row[timestampIndex], timeZone)
    });
  });

  recent.sort(function (a, b) {
    var ta = a.timestamp ? new Date(a.timestamp).getTime() : 0;
    var tb = b.timestamp ? new Date(b.timestamp).getTime() : 0;
    return tb - ta;
  });
  recent = recent.slice(0, 15);

  var states = [];
  for (var stateKey in totalsByState) {
    if (totalsByState.hasOwnProperty(stateKey)) {
      states.push({
        name: stateKey,
        hours: collabRound_(totalsByState[stateKey] / 60, 2)
      });
    }
  }
  states.sort(function (a, b) { return b.hours - a.hours; });

  return {
    summary: {
      entries: rows.length,
      usersTracked: Object.keys(userSet).length,
      totalHours: collabRound_(totalMinutes / 60, 2)
    },
    recent: recent,
    states: states
  };
}

function collabLoadCoachingSummary_(ss) {
  var sheetName = (typeof COACHING_SHEET !== 'undefined') ? COACHING_SHEET : 'CoachingRecords';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return {
      summary: {
        sessions: 0,
        acknowledgements: 0
      },
      recent: []
    };
  }

  var values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) {
    return {
      summary: {
        sessions: 0,
        acknowledgements: 0
      },
      recent: []
    };
  }

  var headers = values[0];
  var rows = values.slice(1);
  var timeZone = Session.getScriptTimeZone();
  var acknowledgedIndex = headers.indexOf('AcknowledgedOn');

  var acknowledgements = 0;
  var recent = [];

  rows.forEach(function (row) {
    if (acknowledgedIndex > -1 && row[acknowledgedIndex]) {
      acknowledgements += 1;
    }
    var record = collabMapRow_(headers, row);
    recent.push({
      agent: record.AgentName || record.CoacheeName || '',
      coach: record.CoacheeName ? record.AgentName : '',
      sessionDate: collabFormatDate_(record.SessionDate, timeZone, true),
      followUp: collabFormatDate_(record.FollowUpDate, timeZone, true),
      notes: record.Summary || record.ActionPlan || ''
    });
  });

  recent.sort(function (a, b) {
    var da = a.sessionDate ? new Date(a.sessionDate).getTime() : 0;
    var db = b.sessionDate ? new Date(b.sessionDate).getTime() : 0;
    return db - da;
  });
  recent = recent.slice(0, 10);

  return {
    summary: {
      sessions: rows.length,
      acknowledgements: acknowledgements
    },
    recent: recent
  };
}

function collabLoadEscalationSummary_(ss) {
  var sheetName = (typeof ESCALATIONS_SHEET !== 'undefined') ? ESCALATIONS_SHEET : 'Escalations';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return {
      summary: {
        openItems: 0
      },
      recent: []
    };
  }

  var values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) {
    return {
      summary: {
        openItems: 0
      },
      recent: []
    };
  }

  var headers = values[0];
  var rows = values.slice(1);
  var timeZone = Session.getScriptTimeZone();
  var typeIndex = headers.indexOf('Type');
  var userIndex = headers.indexOf('User');
  var notesIndex = headers.indexOf('Notes');
  var createdIndex = headers.indexOf('CreatedAt');

  var openItems = rows.length;
  var recent = [];

  rows.forEach(function (row) {
    recent.push({
      type: typeIndex > -1 ? String(row[typeIndex] || '') : '',
      owner: userIndex > -1 ? String(row[userIndex] || '') : '',
      createdAt: collabFormatDate_(row[createdIndex], timeZone),
      notes: notesIndex > -1 ? String(row[notesIndex] || '') : ''
    });
  });

  recent.sort(function (a, b) {
    var da = a.createdAt ? new Date(a.createdAt).getTime() : 0;
    var db = b.createdAt ? new Date(b.createdAt).getTime() : 0;
    return db - da;
  });
  recent = recent.slice(0, 10);

  return {
    summary: {
      openItems: openItems
    },
    recent: recent
  };
}

function collabMapRow_(headers, row) {
  var obj = {};
  for (var i = 0; i < headers.length; i++) {
    obj[headers[i]] = row[i];
  }
  return obj;
}

function collabFormatDate_(value, timeZone, dateOnly) {
  if (!value) return '';
  if (value instanceof Date) {
    var format = dateOnly ? "yyyy-MM-dd" : "yyyy-MM-dd'T'HH:mm:ss";
    return Utilities.formatDate(value, timeZone, format);
  }
  if (typeof value === 'number') {
    return collabFormatDate_(new Date(value), timeZone, dateOnly);
  }
  if (typeof value === 'string') {
    return value;
  }
  return '';
}

function collabNormalizeScore_(percentage, totalScore) {
  if (percentage === null || percentage === '' || typeof percentage === 'undefined') {
    if (totalScore === null || typeof totalScore === 'undefined' || totalScore === '') {
      return null;
    }
    var numericScore = Number(totalScore);
    return isFinite(numericScore) ? numericScore : null;
  }

  var value = percentage;
  if (typeof value === 'string') {
    value = value.replace('%', '');
  }
  var num = Number(value);
  if (!isFinite(num)) return null;
  if (num <= 1) {
    num = num * 100;
  }
  return num;
}

function collabRound_(value, decimals) {
  var factor = Math.pow(10, decimals || 0);
  return Math.round(value * factor) / factor;
}
