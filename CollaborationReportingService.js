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

  return {
    user: {
      id: userId,
      name: currentUser.FullName || currentUser.UserName || '',
      email: currentUser.Email || '',
      campaignId: currentUser.CampaignID || currentUser.CampaignId || '',
      roles: currentUser.roleNames || []
    },
    qa: buildCollaborationQaPayload_(qaRecords, workspace),
    attendance: buildCollaborationAttendancePayload_(workspace),
    executive: buildCollaborationExecutivePayload_(userId, workspace),
    chat: buildCollaborationChatPayload_(workspace),
    campaigns: workspace && Array.isArray(workspace.campaigns) ? workspace.campaigns : [],
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

function buildCollaborationQaPayload_(records, workspace) {
  var sorted = Array.isArray(records) ? records.slice() : [];
  sorted.sort(function (a, b) {
    var da = collabToDate_(a && (a.AuditDate || a.Timestamp || a.UpdatedAt || a.CreatedAt));
    var db = collabToDate_(b && (b.AuditDate || b.Timestamp || b.UpdatedAt || b.CreatedAt));
    var ta = da ? da.getTime() : 0;
    var tb = db ? db.getTime() : 0;
    return tb - ta;
  });

  var limited = sorted.slice(0, 50);
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
  var campaignSet = {};

  mapped.forEach(function (item) {
    if (item.agent) agentSet[item.agent] = true;
    if (item.reviewer) reviewerSet[item.reviewer] = true;
    if (item.campaignName) campaignSet[item.campaignName] = true;
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

  if (workspace && Array.isArray(workspace.campaigns)) {
    workspace.campaigns.forEach(function (campaign) {
      var name = collabToStr_(campaign.name || campaign.Name || '');
      if (name) campaignSet[name] = true;
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
  directory.campaigns = Object.keys(campaignSet).sort().map(function (name) {
    return { id: name, name: name };
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

  return {
    id: id,
    agent: collabToStr_(row.AgentName || row.Agent || ''),
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

function buildCollaborationExecutivePayload_(userId, workspace) {
  var payload = {
    summary: null,
    campaigns: [],
    brief: [],
    timeframeLabel: ''
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
  payload.timeframeLabel = buildTimeframeLabel_();

  return payload;
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
  var now = new Date();
  var month = now.getMonth();
  var quarter = Math.floor(month / 3) + 1;
  var week = collabIsoWeek_(now).split('-W')[1];
  return 'Q' + quarter + ' · Week ' + week;
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
