/**
 * CollaborationReportingService.gs
 * Provides data orchestration for the Collaboration & Reporting hub view.
 * Aggregates QA, attendance, executive, and collaboration chat insights
 * from existing workflow services while ensuring safe fallbacks when data
 * sources are unavailable.
 */
function clientGetCollaborationReportingData(options) {
  options = options || {};
  var currentUser = (typeof getCurrentUser === 'function') ? getCurrentUser() : null;
  if (!currentUser || !currentUser.ID) {
    throw new Error('Unable to resolve the current signed-in user.');
  }

  var userId = String(currentUser.ID);
  var workspace = null;

  if (typeof CallCenterWorkflowService !== 'undefined' && CallCenterWorkflowService &&
      typeof CallCenterWorkflowService.getWorkspace === 'function') {
    try {
      workspace = CallCenterWorkflowService.getWorkspace(userId, {
        activeOnly: true,
        maxMessages: 50
      });
    } catch (err) {
      console.error('clientGetCollaborationReportingData.getWorkspace failed:', err);
    }
  }

  var qaRecords = [];
  if (typeof getAllCollabQA === 'function') {
    try {
      qaRecords = getAllCollabQA();
    } catch (err) {
      console.error('clientGetCollaborationReportingData.getAllCollabQA failed:', err);
    }
  }

  var campaigns = sanitizeWorkspaceCampaigns_(workspace);

  return {
    user: {
      id: userId,
      name: currentUser.FullName || currentUser.UserName || '',
      email: currentUser.Email || '',
      campaignId: currentUser.CampaignID || currentUser.CampaignId || '',
      roles: currentUser.roleNames || []
    },
    qa: buildCollaborationQaPayload_(qaRecords, workspace, campaigns),
    attendance: buildCollaborationAttendancePayload_(workspace),
    executive: buildCollaborationExecutivePayload_(userId, workspace),
    chat: buildCollaborationChatPayload_(workspace),
    teams: buildCollaborationTeamsPayload_(workspace),
    campaigns: campaigns,
    generatedAt: new Date().toISOString()
  };
}

function clientSubmitCollaborationNote(payload) {
  payload = payload || {};
  var currentUser = (typeof getCurrentUser === 'function') ? getCurrentUser() : null;
  if (!currentUser || !currentUser.ID) {
    throw new Error('Unable to resolve the current signed-in user.');
  }

  var userId = String(currentUser.ID);
  var channelId = collabToStr_(payload.channelId) || 'qa-collaboration';
  var campaignId = collabToStr_(payload.campaignId || currentUser.CampaignID || currentUser.CampaignId || '');

  if (!channelId) {
    throw new Error('A collaboration channel is required to post updates.');
  }

  var parts = [];
  if (payload.agent) parts.push('Agent: ' + payload.agent);
  if (payload.campaignName) parts.push('Campaign: ' + payload.campaignName);
  if (payload.reviewer) parts.push('Reviewer: ' + payload.reviewer);
  if (payload.focusArea) parts.push('Focus: ' + payload.focusArea);
  if (payload.status) parts.push('Status: ' + payload.status);
  if (payload.score != null && payload.score !== '') {
    parts.push('Score: ' + payload.score);
  }
  if (payload.nextTouch) parts.push('Next touch: ' + payload.nextTouch);
  if (payload.collaborators && payload.collaborators.length) {
    parts.push('Collaborators: ' + payload.collaborators.join(', '));
  }
  if (payload.channel) parts.push('Channel: ' + payload.channel);

  var highlights = collabToStr_(payload.highlights);
  var messageLines = [];
  messageLines.push('QA collaboration logged by ' + (currentUser.FullName || currentUser.UserName || 'team member') + '.');
  if (parts.length) messageLines.push(parts.join(' | '));
  if (highlights) messageLines.push(highlights);

  if (typeof CallCenterWorkflowService === 'undefined' || !CallCenterWorkflowService ||
      typeof CallCenterWorkflowService.postCollaborationMessage !== 'function') {
    throw new Error('Collaboration messaging service is unavailable.');
  }

  CallCenterWorkflowService.postCollaborationMessage(userId, channelId, messageLines.join('\n'), {
    campaignId: campaignId
  });

  return { success: true };
}

function clientPostCollaborationThreadMessage(request) {
  request = request || {};
  var currentUser = (typeof getCurrentUser === 'function') ? getCurrentUser() : null;
  if (!currentUser || !currentUser.ID) {
    throw new Error('Unable to resolve the current signed-in user.');
  }

  var channelId = collabToStr_(request.channelId);
  var message = collabToStr_(request.message);
  if (!channelId) throw new Error('A collaboration channel is required.');
  if (!message) throw new Error('Message body cannot be empty.');

  if (typeof CallCenterWorkflowService === 'undefined' || !CallCenterWorkflowService ||
      typeof CallCenterWorkflowService.postCollaborationMessage !== 'function') {
    throw new Error('Collaboration messaging service is unavailable.');
  }

  var campaignId = collabToStr_(request.campaignId || currentUser.CampaignID || currentUser.CampaignId || '');
  var result = CallCenterWorkflowService.postCollaborationMessage(String(currentUser.ID), channelId, message, {
    campaignId: campaignId
  });

  return {
    success: true,
    message: result
  };
}

function buildCollaborationQaPayload_(records, workspace, campaigns) {
  var sorted = Array.isArray(records) ? records.slice() : [];
  sorted.sort(function (a, b) {
    var da = collabToDate_(a && (a.AuditDate || a.Timestamp || a.UpdatedAt || a.CreatedAt));
    var db = collabToDate_(b && (b.AuditDate || b.Timestamp || b.UpdatedAt || b.CreatedAt));
    var ta = da ? da.getTime() : 0;
    var tb = db ? db.getTime() : 0;
    return tb - ta;
  });

  var limited = sorted.slice(0, 50);
  var normalizedCampaigns = Array.isArray(campaigns) ? campaigns : sanitizeWorkspaceCampaigns_(workspace);
  var campaignById = {};
  var campaignByName = {};
  normalizedCampaigns.forEach(function (campaign) {
    if (!campaign) return;
    var cid = collabToStr_(campaign.id || '');
    var name = collabToStr_(campaign.name || '');
    if (cid) campaignById[cid] = { id: cid, name: name || cid };
    if (name) campaignByName[name.toLowerCase()] = { id: cid || name, name: name };
  });

  var mapped = [];
  limited.forEach(function (row) {
    var mappedRow = mapCollaborationQaRecord_(row);
    if (mappedRow) mapped.push(mappedRow);
  });

  var directory = {
    agents: [],
    reviewers: [],
    collaborators: [],
    campaigns: []
  };

  var agentSet = {};
  var reviewerSet = {};
  var collaboratorSet = {};
  var campaignDirectoryMap = {};

  function registerCampaignOption(id, name) {
    var normalizedId = id ? String(id).toLowerCase() : '';
    var normalizedName = name ? String(name).toLowerCase() : '';
    var existing = null;
    if (normalizedId && campaignDirectoryMap['id:' + normalizedId]) {
      existing = campaignDirectoryMap['id:' + normalizedId];
    } else if (normalizedName && campaignDirectoryMap['name:' + normalizedName]) {
      existing = campaignDirectoryMap['name:' + normalizedName];
    }
    if (!existing) {
      existing = {
        id: id || name || '',
        name: name || id || ''
      };
    } else {
      if (id && !existing.id) existing.id = id;
      if (name && !existing.name) existing.name = name;
    }
    if (normalizedId) campaignDirectoryMap['id:' + normalizedId] = existing;
    if (normalizedName) campaignDirectoryMap['name:' + normalizedName] = existing;
  }

  normalizedCampaigns.forEach(function (campaign) {
    if (!campaign) return;
    registerCampaignOption(campaign.id, campaign.name);
  });

  mapped.forEach(function (item) {
    if (item.agent) agentSet[item.agent] = true;
    if (item.reviewer) reviewerSet[item.reviewer] = true;
    if (item.campaignId) {
      item.campaignId = collabToStr_(item.campaignId);
    }
    if (item.campaignId && campaignById[item.campaignId]) {
      var match = campaignById[item.campaignId];
      if (!item.campaignName) item.campaignName = match.name;
      item.campaignId = match.id;
    } else if (item.campaignName) {
      var nameKey = item.campaignName.toLowerCase();
      var lookup = campaignByName[nameKey];
      if (lookup) {
        item.campaignId = lookup.id;
        item.campaignName = lookup.name;
      }
    }
    if (!item.campaignId && item.campaignName) {
      item.campaignId = item.campaignName;
    }
    if (item.campaignId || item.campaignName) {
      registerCampaignOption(item.campaignId, item.campaignName);
    }
    if (item.collaborators && item.collaborators.length) {
      item.collaborators.forEach(function (collab) {
        if (collab) collaboratorSet[collab] = true;
      });
    }
  });

  if (workspace && workspace.users && Array.isArray(workspace.users.records)) {
    workspace.users.records.forEach(function (user) {
      var name = collabToStr_(user.FullName || user.UserName || user.Email || user.Name || '');
      var id = collabToStr_(user.ID || user.Id || user.UserID || user.UserId || '');
      if (name) agentSet[name] = true;
      if (name) reviewerSet[name] = true;
      if (id && name) collaboratorSet[name] = true;
    });
  }

  directory.agents = Object.keys(agentSet).sort().map(function (name) {
    return { id: name, name: name };
  });
  directory.reviewers = Object.keys(reviewerSet).sort().map(function (name) {
    return { id: name, name: name };
  });
  directory.collaborators = Object.keys(collaboratorSet).sort().map(function (name) {
    return { id: name, name: name };
  });

  var addedCampaigns = {};
  Object.keys(campaignDirectoryMap).forEach(function (key) {
    var entry = campaignDirectoryMap[key];
    if (!entry) return;
    var dedupeKey = (entry.id || '') + '|' + (entry.name || '');
    if (addedCampaigns[dedupeKey]) return;
    addedCampaigns[dedupeKey] = true;
    directory.campaigns.push({
      id: entry.id || entry.name,
      name: entry.name || entry.id || ''
    });
  });

  directory.campaigns.sort(function (a, b) {
    var nameA = (a.name || '').toLowerCase();
    var nameB = (b.name || '').toLowerCase();
    if (nameA < nameB) return -1;
    if (nameA > nameB) return 1;
    return 0;
  });

  var metrics = computeCollaborationQaMetrics_(mapped, workspace);
  var trend = buildCollaborationQaTrend_(mapped);

  return {
    records: mapped,
    directory: directory,
    metrics: metrics,
    trend: trend
  };
}

function mapCollaborationQaRecord_(row) {
  if (!row) return null;
  var id = collabToStr_(row.ID || row.Id || row.id || row.CaseNumber);
  if (!id) {
    var fallbackDate = collabToDate_(row.Timestamp || row.AuditDate || row.CallDate);
    id = (row.AgentName || 'QA') + '-' + (fallbackDate ? fallbackDate.getTime() : new Date().getTime());
  }

  var score = collabToNumber_(row.Percentage);
  if (score !== null && score <= 1) score = score * 100;
  if (score === null) {
    score = collabToNumber_(row.TotalScore);
  }
  if (score !== null) score = collabRound_(score, 1);

  var focus = collabToStr_(row.FocusArea || row.Focus || row.Category || row.OverallFeedback || '');
  var statusRaw = row.Status || row.FeedbackShared || row.ReviewStatus;
  var status = normalizeCollaborationStatus_(statusRaw);
  var timestamp = collabToDate_(row.AuditDate || row.Timestamp || row.UpdatedAt || row.CreatedAt);
  var callDate = collabToDate_(row.CallDate || row.AuditDate || row.Timestamp);
  var highlights = collabToStr_(row.Notes || row.Highlights || row.OverallFeedback || '');
  var nextTouch = collabToDate_(row.FollowUpDate || row.NextTouch || row.NextReviewDate);
  var channel = collabToStr_(row.Channel || row.InteractionChannel || row.InteractionType || '');
  var collaboratorRaw = row.Collaborators || row.SharedWith || row.CollaborationParticipants || row.Audience || row.Tags;
  var collaborators = [];
  if (Array.isArray(collaboratorRaw)) {
    collaboratorRaw.forEach(function (item) {
      var name = collabToStr_(item);
      if (name) collaborators.push(name);
    });
  } else if (collaboratorRaw) {
    collabToStr_(collaboratorRaw).split(/[;,]/).forEach(function (part) {
      var name = part.trim();
      if (name) collaborators.push(name);
    });
  }

  if (!collaborators.length) {
    var agentFeedback = collabToStr_(row.AgentFeedback || '');
    if (agentFeedback) collaborators.push(agentFeedback);
  }

  var agentId = collabNormalizeUserId_(
    row.AgentUserID || row.AgentUserId || row.AgentId ||
    row.UserID || row.UserId || row.EmployeeID || row.EmployeeId
  );
  var agentEmail = collabNormalizeEmail_(
    row.AgentEmail || row.Email || row.AgentWorkEmail || row.UserEmail
  );

  return {
    id: id,
    agent: collabToStr_(row.AgentName || row.Agent || ''),
    agentId: agentId,
    agentEmail: agentEmail,
    campaignId: collabToStr_(row.CampaignId || row.CampaignID || row.ClientId || row.ClientID || ''),
    campaignName: collabToStr_(row.ClientName || row.Campaign || row.CampaignName || ''),
    reviewer: collabToStr_(row.AuditorName || row.Reviewer || row.ReviewedBy || ''),
    collaborators: collaborators,
    score: score,
    focus: focus,
    status: status,
    updated: timestamp ? timestamp.toISOString() : '',
    callDate: callDate ? callDate.toISOString() : '',
    highlights: highlights,
    channel: channel,
    nextTouch: nextTouch ? nextTouch.toISOString().split('T')[0] : ''
  };
}

function computeCollaborationQaMetrics_(records, workspace) {
  var metrics = {
    averageScore: null,
    followUps: 0,
    collaborativeThreads: 0,
    coverageRate: null,
    deltaVsTarget: null
  };

  if (!records || !records.length) return metrics;

  var scoreSum = 0;
  var scoreCount = 0;
  var agentSet = {};

  records.forEach(function (record) {
    if (record.score !== null && record.score !== undefined && !isNaN(record.score)) {
      scoreSum += record.score;
      scoreCount += 1;
    }
    if (record.status === 'Follow-up') metrics.followUps += 1;
    if (record.collaborators && record.collaborators.length > 1) metrics.collaborativeThreads += 1;
    if (record.agent) agentSet[record.agent] = true;
  });

  if (scoreCount) {
    metrics.averageScore = collabRound_(scoreSum / scoreCount, 1);
  }

  var totalAgents = null;
  if (workspace && workspace.users && typeof workspace.users.total === 'number') {
    totalAgents = workspace.users.total;
  }
  var uniqueAgents = Object.keys(agentSet).length;
  if (totalAgents) {
    metrics.coverageRate = collabRound_((uniqueAgents / Math.max(totalAgents, 1)) * 100, 1);
  } else if (uniqueAgents) {
    metrics.coverageRate = collabRound_(Math.min(uniqueAgents * 10, 100), 1);
  }

  var target = 92;
  if (workspace && workspace.reporting && Array.isArray(workspace.reporting.metrics)) {
    var qaMetric = workspace.reporting.metrics.find(function (item) {
      return item && item.key === 'qaAverage';
    });
    if (qaMetric && qaMetric.value != null && !isNaN(qaMetric.value)) {
      var value = qaMetric.value;
      if (value <= 1) value = value * 100;
      metrics.averageScore = collabRound_(value, 1);
    }
  }

  if (metrics.averageScore != null) {
    metrics.deltaVsTarget = collabRound_(metrics.averageScore - target, 1);
  }

  return metrics;
}

function buildCollaborationQaTrend_(records) {
  var buckets = {};
  (records || []).forEach(function (record) {
    if (!record.callDate) return;
    var date = collabToDate_(record.callDate);
    if (!date) return;
    var key = collabIsoWeek_(date);
    if (!key) return;
    if (!buckets[key]) {
      buckets[key] = { sum: 0, count: 0 };
    }
    if (record.score != null && !isNaN(record.score)) {
      buckets[key].sum += record.score;
      buckets[key].count += 1;
    }
  });

  var labels = Object.keys(buckets).sort();
  var values = labels.map(function (label) {
    var bucket = buckets[label];
    return bucket.count ? collabRound_(bucket.sum / bucket.count, 1) : null;
  });

  if (labels.length > 12) {
    labels = labels.slice(-12);
    values = values.slice(-12);
  }

  return {
    labels: labels,
    values: values
  };
}

function buildCollaborationAttendancePayload_(workspace) {
  var payload = {
    summary: {
      attendanceRate: null,
      absenceRate: null,
      averageAdherence: null
    },
    campaigns: [],
    history: {}
  };

  if (!workspace || !workspace.performance || !workspace.performance.attendance) {
    return payload;
  }

  var attendance = workspace.performance.attendance;
  var summary = attendance.summary || {};

  if (summary.attendanceRate != null) {
    var rate = summary.attendanceRate;
    if (rate <= 1) rate = rate * 100;
    payload.summary.attendanceRate = collabRound_(rate, 1);
  }

  if (summary.statusCounts) {
    var total = 0;
    var absent = 0;
    Object.keys(summary.statusCounts).forEach(function (key) {
      var count = summary.statusCounts[key] || 0;
      total += count;
      if (/absent|no show|callout/i.test(key)) absent += count;
    });
    if (total > 0) {
      payload.summary.absenceRate = collabRound_((absent / total) * 100, 1);
    }
  }

  var rows = Array.isArray(attendance.rows) ? attendance.rows : [];
  var campaigns = {};

  rows.forEach(function (row) {
    var campaign = collabToStr_(row.Campaign || row.CampaignName || row.Client || row.Team || 'All Campaigns');
    var date = collabToDate_(row.Date || row.EventDate || row.Timestamp || row.CreatedAt);
    if (!date) return;
    var weekKey = collabIsoWeek_(date);
    if (!weekKey) return;
    if (!campaigns[campaign]) campaigns[campaign] = {};
    if (!campaigns[campaign][weekKey]) {
      campaigns[campaign][weekKey] = {
        week: weekKey,
        present: 0,
        total: 0,
        absent: 0,
        adherenceSum: 0,
        adherenceCount: 0
      };
    }
    var bucket = campaigns[campaign][weekKey];
    bucket.total += 1;
    var status = collabToStr_(row.Status || row.State || row.Result);
    if (/absent|no show|callout/i.test(status)) {
      bucket.absent += 1;
    } else {
      bucket.present += 1;
    }
    var adherence = collabToNumber_(row.AdherenceScore || row.Adherence || row.Score || row.Percent || row.Percentage);
    if (adherence !== null) {
      if (adherence <= 1) adherence = adherence * 100;
      bucket.adherenceSum += adherence;
      bucket.adherenceCount += 1;
    }
  });

  var campaignList = Object.keys(campaigns);
  campaignList.sort();
  campaignList.forEach(function (campaign) {
    var weeks = Object.keys(campaigns[campaign]).sort();
    payload.history[campaign] = weeks.map(function (week) {
      var bucket = campaigns[campaign][week];
      var adherence = bucket.adherenceCount ? collabRound_(bucket.adherenceSum / bucket.adherenceCount, 1) : null;
      if (adherence === null && bucket.total > 0) {
        adherence = collabRound_((bucket.present / bucket.total) * 100, 1);
      }
      var absenteeism = bucket.total > 0 ? collabRound_((bucket.absent / bucket.total) * 100, 1) : null;
      return {
        week: week,
        adherence: adherence,
        absenteeism: absenteeism,
        present: bucket.present,
        total: bucket.total
      };
    });
  });

  payload.campaigns = campaignList.map(function (name) {
    return { id: name, name: name };
  });

  var adherenceTotals = [];
  campaignList.forEach(function (campaign) {
    payload.history[campaign].forEach(function (entry) {
      if (entry.adherence != null) adherenceTotals.push(entry.adherence);
    });
  });
  if (adherenceTotals.length) {
    var sum = adherenceTotals.reduce(function (acc, val) { return acc + val; }, 0);
    payload.summary.averageAdherence = collabRound_(sum / adherenceTotals.length, 1);
  }

  return payload;
}

function sanitizeWorkspaceCampaigns_(workspace) {
  if (!workspace || !Array.isArray(workspace.campaigns)) {
    return [];
  }

  var seen = {};
  var list = [];

  workspace.campaigns.forEach(function (campaign) {
    if (!campaign) return;
    var id = collabToStr_(campaign.id || campaign.ID || campaign.Id || campaign.CampaignId || campaign.CampaignID || '');
    var name = collabToStr_(campaign.name || campaign.Name || campaign.ClientName || '');
    var key = id || name;
    if (!key || seen[key]) return;
    seen[key] = true;
    list.push({
      id: id || name,
      name: name || id || '',
      description: collabToStr_(campaign.description || campaign.Description || ''),
      isManaged: !!campaign.isManaged,
      isAdmin: !!campaign.isAdmin,
      isDefault: !!campaign.isDefault
    });
  });

  list.sort(function (a, b) {
    var nameA = (a.name || '').toLowerCase();
    var nameB = (b.name || '').toLowerCase();
    if (nameA < nameB) return -1;
    if (nameA > nameB) return 1;
    return 0;
  });

  return list;
}

function buildCollaborationExecutivePayload_(userId, workspace) {
  var payload = {
    summary: null,
    campaigns: [],
    brief: [],
    timeframeLabel: '',
    payPeriod: null
  };

  var analytics = null;
  if (typeof CallCenterWorkflowService !== 'undefined' && CallCenterWorkflowService &&
      typeof CallCenterWorkflowService.getExecutiveAnalytics === 'function') {
    try {
      analytics = CallCenterWorkflowService.getExecutiveAnalytics(userId, {});
    } catch (err) {
      console.error('buildCollaborationExecutivePayload_.getExecutiveAnalytics failed:', err);
    }
  }

  if (!analytics && workspace && workspace.reporting) {
    analytics = {
      summary: {
        totalCampaigns: workspace.reporting.totals ? workspace.reporting.totals.campaigns : 0,
        totalAgents: workspace.reporting.totals ? workspace.reporting.totals.agents : 0,
        averageQaScore: workspace.reporting.metrics ? workspace.reporting.metrics.filter(function (m) { return m.key === 'qaAverage'; }).map(function (m) { return m.value; })[0] : null,
        averageAttendanceRate: workspace.reporting.metrics ? workspace.reporting.metrics.filter(function (m) { return m.key === 'attendanceRate'; }).map(function (m) { return m.value; })[0] : null
      },
      campaigns: []
    };
  }

  if (analytics && analytics.summary) {
    var qaValue = analytics.summary.averageQaScore;
    if (qaValue != null && qaValue <= 1) qaValue = qaValue * 100;
    var attValue = analytics.summary.averageAttendanceRate;
    if (attValue != null && attValue <= 1) attValue = attValue * 100;
    payload.summary = {
      campaigns: analytics.summary.totalCampaigns || 0,
      agents: analytics.summary.totalAgents || 0,
      qaAverage: qaValue != null ? collabRound_(qaValue, 1) : null,
      attendanceAverage: attValue != null ? collabRound_(attValue, 1) : null
    };
  }

  if (analytics && Array.isArray(analytics.campaigns)) {
    analytics.campaigns.forEach(function (entry) {
      if (!entry || !entry.campaign || !entry.snapshot) return;
      var snapshot = entry.snapshot;
      var qaSummary = snapshot.performance && snapshot.performance.qa && snapshot.performance.qa.summary;
      var attendanceSummary = snapshot.performance && snapshot.performance.attendance && snapshot.performance.attendance.summary;
      var qa = qaSummary && qaSummary.averageScore != null ? qaSummary.averageScore : null;
      if (qa != null && qa <= 1) qa = qa * 100;
      var attendance = attendanceSummary && attendanceSummary.attendanceRate != null ? attendanceSummary.attendanceRate : null;
      if (attendance != null && attendance <= 1) attendance = attendance * 100;
      payload.campaigns.push({
        id: entry.campaign.id || entry.campaign.ID || '',
        title: entry.campaign.name || entry.campaign.Name || 'Campaign',
        qa: qa != null ? collabRound_(qa, 1) : null,
        attendance: attendance != null ? collabRound_(attendance, 1) : null,
        alerts: buildExecutiveNotes_(snapshot)
      });
    });
  }

  if (payload.campaigns.length && !payload.summary && workspace && workspace.reporting && workspace.reporting.metrics) {
    var qaMetric = workspace.reporting.metrics.find(function (metric) { return metric && metric.key === 'qaAverage'; });
    var attMetric = workspace.reporting.metrics.find(function (metric) { return metric && metric.key === 'attendanceRate'; });
    payload.summary = {
      campaigns: payload.campaigns.length,
      agents: workspace.reporting.totals ? workspace.reporting.totals.agents : null,
      qaAverage: qaMetric ? collabRound_(qaMetric.value <= 1 ? qaMetric.value * 100 : qaMetric.value, 1) : null,
      attendanceAverage: attMetric ? collabRound_(attMetric.value <= 1 ? attMetric.value * 100 : attMetric.value, 1) : null
    };
  }

  payload.brief = buildExecutiveBrief_(payload.campaigns);
  var payPeriodDetails = collabGetCurrentPayPeriodDetails_();
  if (payPeriodDetails) {
    payload.payPeriod = payPeriodDetails;
    payload.timeframeLabel = payPeriodDetails.label;
  } else {
    payload.timeframeLabel = buildIsoTimeframeLabel_();
  }

  return payload;
}

function buildCollaborationTeamsPayload_(workspace) {
  var payload = {
    overview: {
      totalTeams: 0,
      totalAgents: 0,
      qualityAverage: null,
      attendanceAverage: null,
      callVolume: 0,
      csatAverage: null
    },
    managers: [],
    guests: [],
    managerTabs: []
  };

  var userDirectory = collabBuildUserDirectory_(workspace);
  payload.guests = collabCollectGuestDirectory_(userDirectory);

  var assignments = collabCollectManagerAssignments_();
  if (!assignments.order.length) {
    return payload;
  }

  var attendanceRows = (workspace && workspace.performance && workspace.performance.attendance && workspace.performance.attendance.rows) || [];
  var attendanceIndex = collabBuildAttendanceMetricsIndex_(attendanceRows);
  var callIndex = collabBuildCallMetricsIndex_();

  var overallAgentSet = {};
  var overallQuality = { sum: 0, count: 0 };
  var overallAttendance = { sum: 0, count: 0 };
  var overallCallVolume = 0;
  var overallCsat = { yes: 0, total: 0 };

  var managerIds = assignments.order.slice();
  managerIds.sort(function (a, b) {
    var nameA = (userDirectory[a] && userDirectory[a].name) ? userDirectory[a].name.toLowerCase() : a.toLowerCase();
    var nameB = (userDirectory[b] && userDirectory[b].name) ? userDirectory[b].name.toLowerCase() : b.toLowerCase();
    if (nameA < nameB) return -1;
    if (nameA > nameB) return 1;
    return 0;
  });

  managerIds.forEach(function (managerId) {
    var assignment = assignments.map[managerId];
    if (!assignment) return;
    var managedIds = Object.keys(assignment.userIds);
    if (!managedIds.length) return;

    var managerInfo = userDirectory[managerId] || null;
    var qaSummary = null;
    if (typeof clientGetManagerTeamSummary === 'function') {
      try {
        qaSummary = clientGetManagerTeamSummary(managerId);
      } catch (err) {
        console.error('buildCollaborationTeamsPayload_.getManagerTeamSummary failed:', err);
      }
    }
    var qaIndex = collabIndexManagerQaSummary_(qaSummary);

    var teamUsers = [];
    var teamQuality = { sum: 0, count: 0 };
    var teamAttendance = { sum: 0, count: 0 };
    var teamCallVolume = 0;
    var teamCsat = { yes: 0, total: 0 };

    managedIds.sort(function (a, b) {
      var entryA = userDirectory[a];
      var entryB = userDirectory[b];
      var nameA = (entryA && entryA.name) ? entryA.name.toLowerCase() : a.toLowerCase();
      var nameB = (entryB && entryB.name) ? entryB.name.toLowerCase() : b.toLowerCase();
      if (nameA < nameB) return -1;
      if (nameA > nameB) return 1;
      return 0;
    });

    managedIds.forEach(function (userId) {
      overallAgentSet[userId] = true;
      var directoryEntry = userDirectory[userId] || {};
      var qaEntry = collabLookupMetricForUser_(userId, directoryEntry.displayEmail, qaIndex);
      var attendanceEntry = collabLookupMetricForUser_(userId, directoryEntry.displayEmail, attendanceIndex);
      var callEntry = collabLookupMetricForUser_(userId, directoryEntry.displayEmail, callIndex);

      var quality = (qaEntry && qaEntry.average != null && !isNaN(qaEntry.average)) ? Number(qaEntry.average) : null;
      if (quality != null) {
        teamQuality.sum += quality;
        teamQuality.count += 1;
        overallQuality.sum += quality;
        overallQuality.count += 1;
      }

      var attendanceRate = null;
      if (attendanceEntry && attendanceEntry.total > 0) {
        var rate = ((attendanceEntry.present || 0) / Math.max(attendanceEntry.total, 1)) * 100;
        attendanceRate = collabRound_(rate, 1);
        teamAttendance.sum += attendanceRate;
        teamAttendance.count += 1;
        overallAttendance.sum += attendanceRate;
        overallAttendance.count += 1;
      }

      var callCount = null;
      var csatValue = null;
      if (callEntry) {
        callCount = callEntry.callCount != null ? Number(callEntry.callCount) : 0;
        if (!isNaN(callCount)) {
          teamCallVolume += callCount;
          overallCallVolume += callCount;
        }
        if (callEntry.csatTotal > 0) {
          var ratio = callEntry.csatYes / callEntry.csatTotal;
          if (!isNaN(ratio)) {
            csatValue = collabRound_(ratio * 100, 1);
            teamCsat.yes += callEntry.csatYes;
            teamCsat.total += callEntry.csatTotal;
            overallCsat.yes += callEntry.csatYes;
            overallCsat.total += callEntry.csatTotal;
          }
        }
      }

      var displayName = directoryEntry.name;
      if (!displayName && qaEntry && qaEntry.name) displayName = qaEntry.name;
      if (!displayName) displayName = 'Team Member';

      var email = directoryEntry.displayEmail || (qaEntry && qaEntry.email) || '';

      teamUsers.push({
        id: userId,
        name: displayName,
        email: email,
        roles: directoryEntry.roles || (qaEntry && qaEntry.roles) || [],
        campaignName: directoryEntry.campaignName || '',
        quality: quality,
        attendance: attendanceRate,
        callCount: callCount,
        csat: csatValue
      });
    });

    teamUsers.sort(function (a, b) {
      var nameA = (a.name || '').toLowerCase();
      var nameB = (b.name || '').toLowerCase();
      if (nameA < nameB) return -1;
      if (nameA > nameB) return 1;
      return 0;
    });

    var managerName = managerInfo && managerInfo.name ? managerInfo.name : managerId;
    var managerEmail = managerInfo && managerInfo.displayEmail ? managerInfo.displayEmail : '';

    var qualityAverage = teamQuality.count ? collabRound_(teamQuality.sum / teamQuality.count, 1) : null;
    var attendanceAverage = teamAttendance.count ? collabRound_(teamAttendance.sum / teamAttendance.count, 1) : null;
    var csatAverage = teamCsat.total ? collabRound_((teamCsat.yes / teamCsat.total) * 100, 1) : null;

    payload.managers.push({
      id: managerId,
      name: managerName,
      email: managerEmail,
      teamSize: managedIds.length,
      qualityAverage: qualityAverage,
      attendanceAverage: attendanceAverage,
      callVolume: teamCallVolume,
      csatAverage: csatAverage
    });

    payload.managerTabs.push({
      managerId: managerId,
      name: managerName,
      email: managerEmail,
      summary: {
        teamSize: managedIds.length,
        qualityAverage: qualityAverage,
        attendanceAverage: attendanceAverage,
        callVolume: teamCallVolume,
        csatAverage: csatAverage
      },
      users: teamUsers
    });
  });

  payload.managers.sort(function (a, b) {
    var nameA = (a.name || '').toLowerCase();
    var nameB = (b.name || '').toLowerCase();
    if (nameA < nameB) return -1;
    if (nameA > nameB) return 1;
    return 0;
  });

  payload.managerTabs.sort(function (a, b) {
    var nameA = (a.name || '').toLowerCase();
    var nameB = (b.name || '').toLowerCase();
    if (nameA < nameB) return -1;
    if (nameA > nameB) return 1;
    return 0;
  });

  payload.overview.totalTeams = payload.managerTabs.length;
  payload.overview.totalAgents = Object.keys(overallAgentSet).length;
  payload.overview.qualityAverage = overallQuality.count ? collabRound_(overallQuality.sum / overallQuality.count, 1) : null;
  payload.overview.attendanceAverage = overallAttendance.count ? collabRound_(overallAttendance.sum / overallAttendance.count, 1) : null;
  payload.overview.callVolume = overallCallVolume;
  payload.overview.csatAverage = overallCsat.total ? collabRound_((overallCsat.yes / overallCsat.total) * 100, 1) : null;

  return payload;
}

function collabBuildUserDirectory_(workspace) {
  var directory = {};
  if (!workspace || !workspace.users || !Array.isArray(workspace.users.records)) {
    return directory;
  }

  workspace.users.records.forEach(function (record) {
    if (!record) return;
    var id = collabNormalizeUserId_(record.ID || record.Id || record.UserID || record.UserId || record.AgentId || record.AgentID);
    if (!id) return;
    var roles = Array.isArray(record.roleNames) ? record.roleNames.slice() : [];
    var displayEmail = collabToStr_(record.Email || record.UserEmail || record.AgentEmail || '');
    directory[id] = {
      id: id,
      name: collabToStr_(record.FullName || record.UserName || record.Name || displayEmail || ''),
      displayEmail: displayEmail,
      email: collabNormalizeEmail_(displayEmail),
      roles: roles,
      campaignId: collabToStr_(record.CampaignID || record.CampaignId || record.Campaign || record.campaignId || ''),
      campaignName: collabToStr_(record.CampaignName || record.campaignName || record.Campaign || ''),
      isGuest: roles.some(function (role) { return String(role || '').toLowerCase().indexOf('guest') >= 0; }),
      isManager: roles.some(function (role) { return /manager/i.test(String(role || '')); })
    };
  });

  return directory;
}

function collabCollectGuestDirectory_(directory) {
  var guests = [];
  for (var id in directory) {
    if (!directory.hasOwnProperty(id)) continue;
    var entry = directory[id];
    if (!entry || !entry.isGuest) continue;
    guests.push({
      id: id,
      name: entry.name || 'Guest User',
      email: entry.displayEmail || '',
      campaignName: entry.campaignName || '',
      roles: entry.roles || []
    });
  }

  guests.sort(function (a, b) {
    var nameA = (a.name || '').toLowerCase();
    var nameB = (b.name || '').toLowerCase();
    if (nameA < nameB) return -1;
    if (nameA > nameB) return 1;
    return 0;
  });

  return guests;
}

function collabCollectManagerAssignments_() {
  var result = { order: [], map: {} };
  if (typeof readManagerAssignments_ !== 'function') {
    return result;
  }

  var rows = [];
  try {
    rows = readManagerAssignments_() || [];
  } catch (err) {
    console.error('collabCollectManagerAssignments_ failed:', err);
    rows = [];
  }

  rows.forEach(function (row) {
    if (!row) return;
    var managerId = collabNormalizeUserId_(row.ManagerUserID || row.ManagerId || row.ManagerID || row.managerId);
    var userId = collabNormalizeUserId_(row.UserID || row.ManagedUserID || row.UserId || row.ManagedUserId || row.userId);
    if (!managerId || !userId) return;
    if (!result.map[managerId]) {
      result.map[managerId] = { userIds: {}, campaigns: {} };
      result.order.push(managerId);
    }
    result.map[managerId].userIds[userId] = true;
    var campaignId = collabToStr_(row.CampaignID || row.CampaignId || row.Campaign || '');
    if (campaignId) {
      result.map[managerId].campaigns[campaignId] = true;
    }
  });

  return result;
}

function collabBuildAttendanceMetricsIndex_(rows) {
  var index = { byId: {}, byEmail: {} };
  (rows || []).forEach(function (row) {
    if (!row) return;
    var id = collabNormalizeUserId_(row.UserID || row.UserId || row.AgentId || row.EmployeeID || row.EmployeeId);
    var email = collabNormalizeEmail_(row.UserEmail || row.Email || row.AgentEmail || row.AgentWorkEmail);
    var entry = collabEnsureMetricEntry_(index, id, email, { total: 0, present: 0, absent: 0 });
    if (!entry) return;
    entry.total += 1;
    var status = collabToStr_(row.Status || row.State || row.Result || row.AttendanceStatus || row.Outcome || '');
    if (/absent|no show|noshow|callout|absence/i.test(status)) {
      entry.absent += 1;
    } else {
      entry.present += 1;
    }
  });
  return index;
}

function collabBuildCallMetricsIndex_() {
  var index = { byId: {}, byEmail: {} };
  if (typeof readSheet !== 'function') {
    return index;
  }

  var sheetName = null;
  if (typeof G !== 'undefined' && G && typeof G.CALL_REPORT === 'string' && G.CALL_REPORT) {
    sheetName = G.CALL_REPORT;
  } else if (typeof CALL_REPORT_SHEET === 'string' && CALL_REPORT_SHEET) {
    sheetName = CALL_REPORT_SHEET;
  } else {
    sheetName = 'CallReport';
  }

  var rows = [];
  try {
    rows = readSheet(sheetName) || [];
  } catch (err) {
    try { console.warn('collabBuildCallMetricsIndex_ unable to read sheet', err); } catch (_) {}
    return index;
  }

  rows.forEach(function (row) {
    if (!row) return;
    var id = collabNormalizeUserId_(row.UserID || row.UserId || row.AgentId || row.AgentUserID);
    var email = collabNormalizeEmail_(row.UserEmail || row.AgentEmail || row.AgentWorkEmail || row.ToSFUser);
    var entry = collabEnsureMetricEntry_(index, id, email, { callCount: 0, csatYes: 0, csatTotal: 0 });
    if (!entry) return;
    entry.callCount += 1;
    var csatRaw = collabToStr_(row.CSAT || row.Csat || row.Question1 || row.SurveyResult || '');
    if (csatRaw) {
      entry.csatTotal += 1;
      var normalized = csatRaw.toLowerCase ? csatRaw.toLowerCase() : String(csatRaw).toLowerCase();
      if (/^y(es)?$/.test(normalized) || /^1$/.test(normalized) || normalized === 'true' || normalized === 'positive' || normalized === 'pass') {
        entry.csatYes += 1;
      }
    }
  });

  return index;
}

function collabEnsureMetricEntry_(index, id, email, defaults) {
  if (!index) return null;
  var keyId = collabNormalizeUserId_(id);
  var keyEmail = collabNormalizeEmail_(email);
  var entry = null;
  if (keyId && index.byId[keyId]) entry = index.byId[keyId];
  if (!entry && keyEmail && index.byEmail[keyEmail]) entry = index.byEmail[keyEmail];
  if (!entry) {
    entry = collabCreateMetricEntry_(defaults);
  }
  if (keyId) index.byId[keyId] = entry;
  if (keyEmail) index.byEmail[keyEmail] = entry;
  return entry;
}

function collabCreateMetricEntry_(template) {
  var entry = {};
  for (var key in template) {
    if (template.hasOwnProperty(key)) entry[key] = template[key];
  }
  return entry;
}

function collabLookupMetricForUser_(userId, email, index) {
  if (!index) return null;
  var keyId = collabNormalizeUserId_(userId);
  if (keyId && index.byId[keyId]) return index.byId[keyId];
  var keyEmail = collabNormalizeEmail_(email);
  if (keyEmail && index.byEmail[keyEmail]) return index.byEmail[keyEmail];
  return null;
}

function collabIndexManagerQaSummary_(summary) {
  var index = { byId: {}, byEmail: {} };
  if (!summary || !Array.isArray(summary.users)) {
    return index;
  }

  summary.users.forEach(function (user) {
    if (!user) return;
    var id = collabNormalizeUserId_(user.id || user.ID || user.UserID || user.UserId);
    var email = collabNormalizeEmail_(user.email || user.Email || user.userEmail);
    var entry = {
      id: id,
      name: collabToStr_(user.name || user.FullName || user.UserName || user.email || ''),
      email: collabToStr_(user.email || user.Email || ''),
      roles: Array.isArray(user.roles) ? user.roles.slice() : [],
      average: (user.averageScore != null && !isNaN(user.averageScore)) ? Number(user.averageScore) : null,
      evaluations: user.evaluations || 0
    };
    if (id) index.byId[id] = entry;
    if (email) index.byEmail[email.toLowerCase()] = entry;
  });

  return index;
}

function collabNormalizeUserId_(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return String(value.getTime());
  var text = String(value);
  return text ? text.trim() : '';
}

function collabNormalizeEmail_(value) {
  var text = collabToStr_(value);
  return text ? text.toLowerCase() : '';
}

function buildExecutiveNotes_(snapshot) {
  var notes = [];
  if (!snapshot) return notes;
  var coaching = snapshot.coaching && snapshot.coaching.summary;
  if (coaching) {
    if (coaching.pendingCount) notes.push(coaching.pendingCount + ' coaching items pending');
    if (coaching.overdueCount) notes.push(coaching.overdueCount + ' coaching follow-ups overdue');
  }
  var qaSummary = snapshot.performance && snapshot.performance.qa && snapshot.performance.qa.summary;
  if (qaSummary && qaSummary.failingEvaluations) {
    notes.push(qaSummary.failingEvaluations + ' QA evaluations below threshold');
  }
  var attendanceSummary = snapshot.performance && snapshot.performance.attendance && snapshot.performance.attendance.summary;
  if (attendanceSummary && attendanceSummary.absentCount) {
    notes.push(attendanceSummary.absentCount + ' attendance exceptions');
  }
  return notes;
}

function buildExecutiveBrief_(campaigns) {
  var brief = [];
  campaigns.forEach(function (campaign) {
    var parts = [];
    if (campaign.qa != null) parts.push('QA ' + campaign.qa + '%');
    if (campaign.attendance != null) parts.push('Attendance ' + campaign.attendance + '%');
    if (campaign.alerts && campaign.alerts.length) {
      campaign.alerts.forEach(function (note) {
        brief.push({
          title: campaign.title,
          metrics: parts.join(' · '),
          note: note
        });
      });
    } else {
      brief.push({
        title: campaign.title,
        metrics: parts.join(' · '),
        note: ''
      });
    }
  });
  return brief;
}

function buildTimeframeLabel_() {
  var payPeriod = collabGetCurrentPayPeriodDetails_();
  if (payPeriod && payPeriod.label) {
    return payPeriod.label;
  }
  return buildIsoTimeframeLabel_();
}

function buildIsoTimeframeLabel_() {
  var now = new Date();
  var month = now.getMonth();
  var quarter = Math.floor(month / 3) + 1;
  var week = collabIsoWeek_(now).split('-W')[1];
  return 'Q' + quarter + ' · Week ' + week;
}

function collabGetCurrentPayPeriodDetails_() {
  var periods = collabGetBiweeklyPayPeriods_();
  if (!periods || !periods.length) {
    return null;
  }

  var todayValue = collabDateValue_(new Date());
  if (todayValue === null) {
    return null;
  }

  var chosen = null;
  for (var i = 0; i < periods.length; i++) {
    var period = periods[i];
    if (todayValue >= period.startValue && todayValue <= period.endValue) {
      chosen = period;
      break;
    }
  }

  if (!chosen) {
    if (todayValue < periods[0].startValue) {
      chosen = periods[0];
    } else {
      chosen = periods[periods.length - 1];
    }
  }

  return collabDecoratePayPeriod_(chosen);
}

function collabDecoratePayPeriod_(period) {
  if (!period) {
    return null;
  }

  var rangeLabel = collabFormatDateRange_(period.start, period.end);
  var badgeLabel = 'Pay Period ' + period.number;
  var payDateLabel = collabFormatFullDate_(period.payDate);

  return {
    number: period.number,
    label: badgeLabel + (rangeLabel ? ' · ' + rangeLabel : ''),
    badgeLabel: badgeLabel,
    rangeLabel: rangeLabel,
    payDateLabel: payDateLabel,
    startIso: collabToIsoDateString_(period.start),
    endIso: collabToIsoDateString_(period.end),
    payDateIso: collabToIsoDateString_(period.payDate)
  };
}

var collabBiweeklyPayPeriodsCache_ = null;

function collabGetBiweeklyPayPeriods_() {
  if (collabBiweeklyPayPeriodsCache_) {
    return collabBiweeklyPayPeriodsCache_;
  }

  // WBPO 2025 biweekly pay calendar shared by operations.
  var periods = [
    collabCreatePayPeriod_(1, collabMakeDate_(2024, 12, 22), collabMakeDate_(2025, 1, 4), collabMakeDate_(2025, 1, 10)),
    collabCreatePayPeriod_(2, collabMakeDate_(2025, 1, 5), collabMakeDate_(2025, 1, 18), collabMakeDate_(2025, 1, 24)),
    collabCreatePayPeriod_(3, collabMakeDate_(2025, 1, 19), collabMakeDate_(2025, 2, 1), collabMakeDate_(2025, 2, 7)),
    collabCreatePayPeriod_(4, collabMakeDate_(2025, 2, 2), collabMakeDate_(2025, 2, 15), collabMakeDate_(2025, 2, 21)),
    collabCreatePayPeriod_(5, collabMakeDate_(2025, 2, 16), collabMakeDate_(2025, 3, 1), collabMakeDate_(2025, 3, 7)),
    collabCreatePayPeriod_(6, collabMakeDate_(2025, 3, 2), collabMakeDate_(2025, 3, 15), collabMakeDate_(2025, 3, 21)),
    collabCreatePayPeriod_(7, collabMakeDate_(2025, 3, 16), collabMakeDate_(2025, 3, 29), collabMakeDate_(2025, 4, 4)),
    collabCreatePayPeriod_(8, collabMakeDate_(2025, 3, 30), collabMakeDate_(2025, 4, 12), collabMakeDate_(2025, 4, 18)),
    collabCreatePayPeriod_(9, collabMakeDate_(2025, 4, 13), collabMakeDate_(2025, 4, 26), collabMakeDate_(2025, 5, 2)),
    collabCreatePayPeriod_(10, collabMakeDate_(2025, 4, 27), collabMakeDate_(2025, 5, 10), collabMakeDate_(2025, 5, 16)),
    collabCreatePayPeriod_(11, collabMakeDate_(2025, 5, 11), collabMakeDate_(2025, 5, 24), collabMakeDate_(2025, 5, 30)),
    collabCreatePayPeriod_(12, collabMakeDate_(2025, 5, 25), collabMakeDate_(2025, 6, 7), collabMakeDate_(2025, 6, 13)),
    collabCreatePayPeriod_(13, collabMakeDate_(2025, 6, 8), collabMakeDate_(2025, 6, 21), collabMakeDate_(2025, 6, 27)),
    collabCreatePayPeriod_(14, collabMakeDate_(2025, 6, 22), collabMakeDate_(2025, 7, 5), collabMakeDate_(2025, 7, 11)),
    collabCreatePayPeriod_(15, collabMakeDate_(2025, 7, 6), collabMakeDate_(2025, 7, 19), collabMakeDate_(2025, 7, 25)),
    collabCreatePayPeriod_(16, collabMakeDate_(2025, 7, 20), collabMakeDate_(2025, 8, 2), collabMakeDate_(2025, 8, 8)),
    collabCreatePayPeriod_(17, collabMakeDate_(2025, 8, 3), collabMakeDate_(2025, 8, 16), collabMakeDate_(2025, 8, 22)),
    collabCreatePayPeriod_(18, collabMakeDate_(2025, 8, 17), collabMakeDate_(2025, 8, 30), collabMakeDate_(2025, 9, 5)),
    collabCreatePayPeriod_(19, collabMakeDate_(2025, 8, 31), collabMakeDate_(2025, 9, 13), collabMakeDate_(2025, 9, 19)),
    collabCreatePayPeriod_(20, collabMakeDate_(2025, 9, 14), collabMakeDate_(2025, 9, 27), collabMakeDate_(2025, 10, 3)),
    collabCreatePayPeriod_(21, collabMakeDate_(2025, 9, 28), collabMakeDate_(2025, 10, 11), collabMakeDate_(2025, 10, 17)),
    collabCreatePayPeriod_(22, collabMakeDate_(2025, 10, 12), collabMakeDate_(2025, 10, 25), collabMakeDate_(2025, 10, 31)),
    collabCreatePayPeriod_(23, collabMakeDate_(2025, 10, 26), collabMakeDate_(2025, 11, 8), collabMakeDate_(2025, 11, 14)),
    collabCreatePayPeriod_(24, collabMakeDate_(2025, 11, 9), collabMakeDate_(2025, 11, 22), collabMakeDate_(2025, 11, 28)),
    collabCreatePayPeriod_(25, collabMakeDate_(2025, 11, 23), collabMakeDate_(2025, 12, 6), collabMakeDate_(2025, 12, 12)),
    collabCreatePayPeriod_(26, collabMakeDate_(2025, 12, 7), collabMakeDate_(2025, 12, 20), collabMakeDate_(2025, 12, 26)),
    collabCreatePayPeriod_(27, collabMakeDate_(2025, 12, 21), collabMakeDate_(2026, 1, 3), collabMakeDate_(2026, 1, 9))
  ];

  collabBiweeklyPayPeriodsCache_ = periods;
  return periods;
}

function collabCreatePayPeriod_(number, start, end, payDate) {
  return {
    number: number,
    start: start,
    end: end,
    payDate: payDate,
    startValue: collabDateValue_(start),
    endValue: collabDateValue_(end)
  };
}

function collabMakeDate_(year, month, day) {
  return new Date(Date.UTC(year, month - 1, day, 12));
}

function collabDateValue_(date) {
  if (!(date instanceof Date)) {
    return null;
  }
  return Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
}

function collabToIsoDateString_(date) {
  if (!(date instanceof Date)) {
    return '';
  }
  return date.getUTCFullYear() + '-' + collabPad2_(date.getUTCMonth() + 1) + '-' + collabPad2_(date.getUTCDate());
}

function collabPad2_(value) {
  return value < 10 ? '0' + value : String(value);
}

function collabFormatDateRange_(start, end) {
  var hasStart = start instanceof Date;
  var hasEnd = end instanceof Date;
  if (!hasStart && !hasEnd) {
    return '';
  }
  if (!hasStart) {
    return collabFormatFullDate_(end);
  }
  if (!hasEnd) {
    return collabFormatFullDate_(start);
  }

  var startMonth = collabMonthName_(start);
  var endMonth = collabMonthName_(end);
  var startDay = start.getUTCDate();
  var endDay = end.getUTCDate();
  var startYear = start.getUTCFullYear();
  var endYear = end.getUTCFullYear();

  if (startYear === endYear) {
    if (startMonth === endMonth) {
      return startMonth + ' ' + startDay + '–' + endDay + ', ' + startYear;
    }
    return startMonth + ' ' + startDay + ' – ' + endMonth + ' ' + endDay + ', ' + startYear;
  }

  return startMonth + ' ' + startDay + ', ' + startYear + ' – ' + endMonth + ' ' + endDay + ', ' + endYear;
}

function collabFormatFullDate_(date) {
  if (!(date instanceof Date)) {
    return '';
  }
  var monthName = collabMonthName_(date);
  if (!monthName) {
    return '';
  }
  return monthName + ' ' + date.getUTCDate() + ', ' + date.getUTCFullYear();
}

function collabMonthName_(date) {
  if (!(date instanceof Date)) {
    return '';
  }
  var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var index = date.getUTCMonth();
  if (index < 0 || index >= months.length) {
    return '';
  }
  return months[index];
}

function buildCollaborationChatPayload_(workspace) {
  var payload = {
    personas: [],
    threads: {}
  };

  if (!workspace || !workspace.collaboration || !Array.isArray(workspace.collaboration.messages)) {
    payload.personas = [{ key: 'general', label: 'Collaboration Feed', icon: 'fa-comments' }];
    payload.threads.general = [];
    return payload;
  }

  var userNameIndex = {};
  if (workspace.users && Array.isArray(workspace.users.records)) {
    workspace.users.records.forEach(function (user) {
      var id = collabToStr_(user.ID || user.Id || user.UserID || user.UserId);
      var name = collabToStr_(user.FullName || user.UserName || user.Email || user.Name || '');
      if (id && name) userNameIndex[id] = name;
    });
  }

  var threadMap = {};
  workspace.collaboration.messages.forEach(function (message) {
    if (!message) return;
    var channelId = collabToStr_(message.ChannelId || message.Channel || 'general');
    if (!threadMap[channelId]) {
      threadMap[channelId] = {
        id: channelId,
        channelId: channelId,
        title: collabToStr_(message.ChannelName || message.Subject || channelId),
        audience: collabToStr_(message.Audience || 'Team'),
        updated: '',
        messages: [],
        campaignId: collabToStr_(message.CampaignID || message.CampaignId || message.campaignId || '')
      };
    }
    var thread = threadMap[channelId];
    var timestamp = collabToDate_(message.Timestamp || message.CreatedAt || message.SentAt || message.UpdatedAt);
    if (timestamp && (!thread.updated || collabToDate_(thread.updated) < timestamp)) {
      thread.updated = timestamp.toISOString();
    }
    var userId = collabToStr_(message.UserId || message.UserID || message.AuthorId || message.SenderId || '');
    var authorName = userNameIndex[userId] || collabToStr_(message.Author || message.Sender || message.CreatedBy || 'Team');
    var tags = [];
    if (Array.isArray(message.Tags)) {
      message.Tags.forEach(function (tag) {
        var name = collabToStr_(tag);
        if (name) tags.push(name);
      });
    } else if (message.Tags) {
      collabToStr_(message.Tags).split(/[;,]/).forEach(function (part) {
        var name = part.trim();
        if (name) tags.push(name);
      });
    }
    thread.messages.push({
      author: authorName || 'Team',
      time: timestamp ? timestamp.toISOString() : '',
      text: collabToStr_(message.Message || message.Body || message.Text || ''),
      tags: tags
    });
  });

  var threads = Object.keys(threadMap).map(function (key) {
    var thread = threadMap[key];
    thread.messages.sort(function (a, b) {
      var ta = collabToDate_(a.time);
      var tb = collabToDate_(b.time);
      var va = ta ? ta.getTime() : 0;
      var vb = tb ? tb.getTime() : 0;
      return va - vb;
    });
    return thread;
  });

  threads.sort(function (a, b) {
    var ta = collabToDate_(a.updated);
    var tb = collabToDate_(b.updated);
    var va = ta ? ta.getTime() : 0;
    var vb = tb ? tb.getTime() : 0;
    return vb - va;
  });

  payload.personas = [{ key: 'general', label: 'Collaboration Feed', icon: 'fa-comments' }];
  payload.threads.general = threads;
  return payload;
}

function normalizeCollaborationStatus_(value) {
  var str = collabToStr_(value).toLowerCase();
  if (!str) return 'Draft';
  if (str === 'yes' || str === 'true' || str.indexOf('publish') >= 0) return 'Published';
  if (str.indexOf('follow') >= 0) return 'Follow-up';
  return 'Draft';
}

function collabToNumber_(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number') {
    return isNaN(value) ? null : value;
  }
  if (typeof value === 'string') {
    var normalized = value.replace(/[^0-9.\-]/g, '');
    if (!normalized) return null;
    var parsed = Number(normalized);
    return isNaN(parsed) ? null : parsed;
  }
  return null;
}

function collabToDate_(value) {
  if (!value && value !== 0) return null;
  if (value instanceof Date) {
    return new Date(value.getTime());
  }
  if (typeof value === 'number') {
    var numeric = new Date(value);
    return isNaN(numeric.getTime()) ? null : numeric;
  }
  if (typeof value === 'string') {
    var trimmed = value.trim();
    if (!trimmed) return null;
    var parsed = new Date(trimmed);
    if (!isNaN(parsed.getTime())) return parsed;
  }
  return null;
}

function collabToStr_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function collabRound_(value, digits) {
  if (value === null || value === undefined || isNaN(value)) return null;
  var factor = Math.pow(10, digits || 0);
  return Math.round(value * factor) / factor;
}

function collabIsoWeek_(date) {
  if (!date) return '';
  var target = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  var dayNr = (target.getUTCDay() + 6) % 7;
  target.setUTCDate(target.getUTCDate() - dayNr + 3);
  var firstThursday = new Date(Date.UTC(target.getUTCFullYear(), 0, 4));
  var week = 1 + Math.round(((target - firstThursday) / 86400000 - 3) / 7);
  var year = target.getUTCFullYear();
  return year + '-W' + String(week).padStart(2, '0');
}
