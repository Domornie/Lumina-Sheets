/**
 * ManagerExecutiveService.gs
 * High-level orchestration layer powering the Manager & Executive experience.
 * Surfaces campaign dashboards, coaching insights, communications and
 * administrative tooling from existing workflow services in a single payload.
 */
function clientGetManagerExecutiveExperience(options) {
  options = options || {};
  var currentUser = (typeof getCurrentUser === 'function') ? getCurrentUser() : null;
  if (!currentUser || !currentUser.ID) {
    throw new Error('Unable to resolve the current signed-in user.');
  }

  var userId = String(currentUser.ID);
  var roleNames = Array.isArray(currentUser.roleNames) ? currentUser.roleNames.slice() : [];
  var normalizedRoles = roleNames.map(function (role) {
    return String(role || '').toLowerCase();
  });
  var isExecutive = normalizedRoles.some(function (role) {
    return role.indexOf('executive') >= 0 || role.indexOf('admin') >= 0;
  }) || currentUser.IsAdmin === true || String(currentUser.IsAdmin).toUpperCase() === 'TRUE';

  var response = {
    user: {
      id: userId,
      name: currentUser.FullName || currentUser.UserName || '',
      email: currentUser.Email || currentUser.email || '',
      campaignId: currentUser.CampaignID || currentUser.CampaignId || '',
      campaignName: currentUser.campaignName || '',
      roles: roleNames
    },
    isExecutive: isExecutive,
    campaigns: [],
    dashboards: [],
    executive: null,
    agentsByCampaign: {},
    hero: null,
    coachingInsights: null,
    communications: null,
    admin: null,
    generatedAt: new Date().toISOString()
  };

  var campaignAccess = [];
  if (typeof CallCenterWorkflowService !== 'undefined' && CallCenterWorkflowService &&
      typeof CallCenterWorkflowService.listCampaignAccess === 'function') {
    try {
      campaignAccess = CallCenterWorkflowService.listCampaignAccess(userId) || [];
    } catch (err) {
      console.error('clientGetManagerExecutiveExperience.listCampaignAccess failed:', err);
    }
  }
  response.campaigns = campaignAccess;

  var dashboards = [];
  campaignAccess.forEach(function (campaign) {
    if (!campaign || !campaign.id) return;
    try {
      var snapshot = CallCenterWorkflowService.getManagerCampaignDashboard(
        userId,
        campaign.id,
        Object.assign({}, options, { includeRoster: true })
      );
      dashboards.push({
        campaign: campaign,
        snapshot: snapshot
      });
      if (snapshot && snapshot.roster && Array.isArray(snapshot.roster.records)) {
        response.agentsByCampaign[campaign.id] = snapshot.roster.records.map(function (record) {
          return {
            id: String(record.ID || record.Id || record.UserID || record.UserId || record.AgentId || record.AgentID || ''),
            name: record.FullName || record.UserName || record.AgentName || record.Name || record.Email || 'Agent',
            email: record.Email || record.UserEmail || record.AgentEmail || '',
            employmentStatus: record.EmploymentStatus || record.Status || '',
            role: record.Role || record.RoleName || '',
            campaignId: campaign.id,
            campaignName: campaign.name || ''
          };
        });
      }
    } catch (err) {
      console.error('clientGetManagerExecutiveExperience.getManagerCampaignDashboard failed:', err);
      dashboards.push({
        campaign: campaign,
        error: err && err.message ? err.message : String(err)
      });
    }
  });
  response.dashboards = dashboards;

  if (isExecutive && typeof CallCenterWorkflowService.getExecutiveAnalytics === 'function') {
    try {
      response.executive = CallCenterWorkflowService.getExecutiveAnalytics(userId, options) || null;
    } catch (err) {
      console.error('clientGetManagerExecutiveExperience.getExecutiveAnalytics failed:', err);
      response.executive = { error: err && err.message ? err.message : String(err) };
    }
  }

  response.hero = buildManagerExecutiveHero_(dashboards, response.executive);
  response.coachingInsights = buildManagerExecutiveCoaching_(dashboards);
  response.communications = buildManagerExecutiveCommunications_(dashboards);
  response.admin = buildManagerExecutiveAdminContext_(campaignAccess);

  return response;
}

function buildManagerExecutiveHero_(dashboards, executiveAnalytics) {
  var totals = {
    campaigns: 0,
    agents: 0,
    qaScoreSum: 0,
    qaScoreCount: 0,
    attendanceSum: 0,
    attendanceCount: 0,
    pendingCoaching: 0,
    overdueCoaching: 0,
    acknowledgedCoaching: 0
  };

  dashboards.forEach(function (entry) {
    if (!entry || !entry.snapshot || !entry.campaign || entry.error) return;
    totals.campaigns += 1;
    var snapshot = entry.snapshot;
    if (snapshot.roster && typeof snapshot.roster.total === 'number') {
      totals.agents += snapshot.roster.total;
    }
    var qaSummary = snapshot.performance && snapshot.performance.qa && snapshot.performance.qa.summary;
    if (qaSummary && qaSummary.averageScore != null && !isNaN(qaSummary.averageScore)) {
      totals.qaScoreSum += Number(qaSummary.averageScore);
      totals.qaScoreCount += 1;
    }
    var attendanceSummary = snapshot.performance && snapshot.performance.attendance && snapshot.performance.attendance.summary;
    if (attendanceSummary && attendanceSummary.attendanceRate != null && !isNaN(attendanceSummary.attendanceRate)) {
      totals.attendanceSum += Number(attendanceSummary.attendanceRate);
      totals.attendanceCount += 1;
    }
    var coachingSummary = snapshot.coaching && snapshot.coaching.summary;
    if (coachingSummary) {
      totals.pendingCoaching += Number(coachingSummary.pendingCount || 0);
      totals.overdueCoaching += Number(coachingSummary.overdueCount || 0);
      totals.acknowledgedCoaching += Number(coachingSummary.acknowledgedCount || 0);
    }
  });

  if (executiveAnalytics && executiveAnalytics.summary) {
    var exec = executiveAnalytics.summary;
    if (exec.totalCampaigns && exec.totalCampaigns > totals.campaigns) {
      totals.campaigns = exec.totalCampaigns;
    }
    if (exec.totalAgents && exec.totalAgents > totals.agents) {
      totals.agents = exec.totalAgents;
    }
    if (exec.averageQaScore != null && !isNaN(exec.averageQaScore)) {
      totals.qaScoreSum = Number(exec.averageQaScore);
      totals.qaScoreCount = 1;
    }
    if (exec.averageAttendanceRate != null && !isNaN(exec.averageAttendanceRate)) {
      totals.attendanceSum = Number(exec.averageAttendanceRate);
      totals.attendanceCount = 1;
    }
  }

  var qaAvg = totals.qaScoreCount ? Math.round((totals.qaScoreSum / totals.qaScoreCount) * 100) / 100 : null;
  var attendanceAvg = totals.attendanceCount ? Math.round((totals.attendanceSum / totals.attendanceCount) * 100) / 100 : null;

  return {
    campaigns: totals.campaigns,
    agents: totals.agents,
    qaAverage: qaAvg,
    attendanceAverage: attendanceAvg,
    pendingCoaching: totals.pendingCoaching,
    overdueCoaching: totals.overdueCoaching,
    acknowledgedCoaching: totals.acknowledgedCoaching,
    metrics: [
      { key: 'campaigns', label: 'Active Campaigns', value: totals.campaigns, icon: 'fa-layer-group', format: 'number' },
      { key: 'agents', label: 'Agents Managed', value: totals.agents, icon: 'fa-users', format: 'number' },
      { key: 'qaAverage', label: 'Average QA', value: qaAvg, icon: 'fa-star', format: 'percentage' },
      { key: 'attendanceAverage', label: 'Attendance Rate', value: attendanceAvg, icon: 'fa-calendar-check', format: 'percentage' },
      { key: 'pendingCoaching', label: 'Pending Coaching', value: totals.pendingCoaching, icon: 'fa-user-clock', format: 'number' },
      { key: 'overdueCoaching', label: 'Overdue Follow-ups', value: totals.overdueCoaching, icon: 'fa-exclamation-triangle', format: 'number' }
    ]
  };
}

function buildManagerExecutiveCoaching_(dashboards) {
  var insights = {
    pending: [],
    summaryByCampaign: {},
    totals: {
      sessions: 0,
      pending: 0,
      acknowledged: 0,
      overdue: 0
    }
  };

  dashboards.forEach(function (entry) {
    if (!entry || !entry.snapshot || !entry.campaign || entry.error) return;
    var campaignId = entry.campaign.id;
    var campaignName = entry.campaign.name || '';
    var snapshot = entry.snapshot;
    var summary = snapshot.coaching && snapshot.coaching.summary;
    if (summary) {
      insights.summaryByCampaign[campaignId] = {
        campaignId: campaignId,
        campaignName: campaignName,
        totalSessions: Number(summary.totalSessions || 0),
        pendingCount: Number(summary.pendingCount || 0),
        acknowledgedCount: Number(summary.acknowledgedCount || 0),
        overdueCount: Number(summary.overdueCount || 0)
      };
      insights.totals.sessions += Number(summary.totalSessions || 0);
      insights.totals.pending += Number(summary.pendingCount || 0);
      insights.totals.acknowledged += Number(summary.acknowledgedCount || 0);
      insights.totals.overdue += Number(summary.overdueCount || 0);
    }
    if (snapshot.coaching && Array.isArray(snapshot.coaching.pending)) {
      snapshot.coaching.pending.forEach(function (row) {
        insights.pending.push({
          id: String(row.ID || row.Id || row.CoachingId || row.CoachingID || row.SessionId || ''),
          campaignId: campaignId,
          campaignName: campaignName,
          agentId: String(row.UserID || row.UserId || row.AgentId || row.AgentID || ''),
          agentName: row.AgentName || row.Agent || row.UserName || row.FullName || row.Name || 'Agent',
          status: row.Status || row.AcknowledgementStatus || row.State || 'Pending',
          focusArea: row.FocusArea || row.Topic || row.Category || '',
          dueDate: formatIsoDate_(row.DueDate || row.FollowUpDate || row.AcknowledgeBy || ''),
          createdAt: formatIsoDateTime_(row.CreatedAt || row.Created || row.SessionDate || ''),
          notes: row.Summary || row.Notes || row.ActionPlan || '',
          requireAcknowledgement: row.AcknowledgementRequired === true || /true/i.test(String(row.AcknowledgementRequired || ''))
        });
      });
    }
  });

  insights.pending.sort(function (a, b) {
    if (a.dueDate && b.dueDate) return a.dueDate.localeCompare(b.dueDate);
    if (a.dueDate) return -1;
    if (b.dueDate) return 1;
    return (a.createdAt || '').localeCompare(b.createdAt || '');
  });

  return insights;
}

function buildManagerExecutiveCommunications_(dashboards) {
  var feed = [];
  dashboards.forEach(function (entry) {
    if (!entry || !entry.snapshot || !entry.campaign || entry.error) return;
    var campaignId = entry.campaign.id;
    var campaignName = entry.campaign.name || '';
    var communications = entry.snapshot.communications;
    if (!communications || !Array.isArray(communications.records)) return;
    communications.records.forEach(function (record) {
      feed.push({
        campaignId: campaignId,
        campaignName: campaignName,
        id: String(record.ID || record.Id || record.NotificationId || record.NotificationID || ''),
        title: record.Title || record.Subject || 'Campaign Update',
        body: record.Message || record.Body || '',
        severity: record.Severity || record.Level || 'Info',
        createdAt: formatIsoDateTime_(record.CreatedAt || record.Timestamp || record.SentAt || ''),
        author: record.CreatedBy || record.Sender || ''
      });
    });
  });

  feed.sort(function (a, b) {
    return (b.createdAt || '').localeCompare(a.createdAt || '');
  });

  return {
    records: feed.slice(0, 25),
    total: feed.length
  };
}

function buildManagerExecutiveAdminContext_(campaigns) {
  var roles = [];
  if (typeof getAllRoles === 'function') {
    try {
      roles = (getAllRoles() || []).map(function (role) {
        return { id: role.id || role.ID, name: role.name || role.Name || '' };
      });
    } catch (err) {
      console.error('buildManagerExecutiveAdminContext_:getAllRoles failed:', err);
    }
  }

  var permissionLevels = [
    { key: 'VIEWER', label: 'Viewer (Read-only)' },
    { key: 'EDITOR', label: 'Editor (Standard manager)' },
    { key: 'GUEST', label: 'Guest (Client view)' },
    { key: 'ADMIN', label: 'Administrator' }
  ];

  return {
    campaigns: campaigns,
    roles: roles,
    permissionLevels: permissionLevels
  };
}

function clientCreateManagerExecutiveCoaching(payload) {
  payload = payload || {};
  var currentUser = (typeof getCurrentUser === 'function') ? getCurrentUser() : null;
  if (!currentUser || !currentUser.ID) {
    throw new Error('Unable to resolve current manager context.');
  }
  if (!payload.agentId) {
    throw new Error('An agent selection is required to create coaching.');
  }

  var managerId = String(currentUser.ID);
  var campaignId = payload.campaignId || currentUser.CampaignID || currentUser.CampaignId || '';
  var now = new Date();
  var sessionDate = payload.sessionDate || formatIsoDate_(now);

  var record = {
    UserID: payload.agentId,
    UserId: payload.agentId,
    AgentId: payload.agentId,
    CampaignID: campaignId,
    CampaignId: campaignId,
    SessionDate: sessionDate,
    TopicsCovered: Array.isArray(payload.topics) ? payload.topics.join(', ') : (payload.topics || ''),
    FocusArea: payload.focusArea || '',
    Summary: payload.summary || payload.notes || '',
    ActionPlan: payload.actionPlan || '',
    FollowUpDate: payload.followUpDate || '',
    FollowUpNotes: payload.followUpNotes || '',
    DueDate: payload.followUpDate || payload.dueDate || '',
    Status: payload.requireAcknowledgement ? 'Pending Acknowledgement' : (payload.status || 'Pending'),
    AcknowledgementRequired: payload.requireAcknowledgement ? 'TRUE' : 'FALSE',
    Source: 'manager-executive-experience'
  };

  try {
    var created = CallCenterWorkflowService.createCoachingSession(managerId, record);
    return { success: true, coachingId: created && created.ID ? created.ID : record.ID, record: created || record };
  } catch (err) {
    console.error('clientCreateManagerExecutiveCoaching failed:', err);
    throw err;
  }
}

function clientUpdateManagerCoachingStatus(payload) {
  payload = payload || {};
  var currentUser = (typeof getCurrentUser === 'function') ? getCurrentUser() : null;
  if (!currentUser || !currentUser.ID) {
    throw new Error('Unable to resolve current manager context.');
  }
  var coachingId = String(payload.coachingId || payload.id || '');
  if (!coachingId) {
    throw new Error('A coaching identifier is required.');
  }

  var managerId = String(currentUser.ID);
  var campaignId = payload.campaignId || currentUser.CampaignID || currentUser.CampaignId || '';
  var updates = {};
  if (payload.status) updates.Status = payload.status;
  if (payload.followUpDate) {
    updates.FollowUpDate = payload.followUpDate;
    updates.DueDate = payload.followUpDate;
  }
  if (payload.followUpNotes) updates.FollowUpNotes = payload.followUpNotes;
  if (payload.notes) updates.Notes = payload.notes;
  if (payload.acknowledgedAt) updates.AcknowledgedAt = payload.acknowledgedAt;
  if (payload.requireAcknowledgement != null) {
    updates.AcknowledgementRequired = payload.requireAcknowledgement ? 'TRUE' : 'FALSE';
  }
  if (payload.requestAcknowledgement) {
    updates.AcknowledgementRequestedAt = new Date().toISOString();
    if (!updates.Status) updates.Status = 'Awaiting Acknowledgement';
  }

  try {
    var updated = CallCenterWorkflowService.updateCoachingSession(managerId, coachingId, updates, campaignId);
    if (payload.requestAcknowledgement && payload.agentId) {
      try {
        CallCenterWorkflowService.sendCampaignCommunication(managerId, campaignId, {
          title: 'Coaching acknowledgement requested',
          message: 'Please review and acknowledge your latest coaching session.',
          severity: 'Info',
          userIds: [payload.agentId]
        });
      } catch (notifyErr) {
        console.warn('clientUpdateManagerCoachingStatus notification failed:', notifyErr);
      }
    }
    return { success: true, coachingId: coachingId, updates: updates, result: updated };
  } catch (err) {
    console.error('clientUpdateManagerCoachingStatus failed:', err);
    throw err;
  }
}

function clientSendManagerExecutiveBroadcast(payload) {
  payload = payload || {};
  var currentUser = (typeof getCurrentUser === 'function') ? getCurrentUser() : null;
  if (!currentUser || !currentUser.ID) {
    throw new Error('Unable to resolve current manager context.');
  }
  if (!payload.campaignId) {
    throw new Error('A campaign selection is required to send an update.');
  }
  if (!payload.message) {
    throw new Error('A message body is required.');
  }

  var managerId = String(currentUser.ID);
  var result = CallCenterWorkflowService.sendCampaignCommunication(managerId, payload.campaignId, {
    title: payload.title || 'Performance Update',
    message: payload.message,
    severity: payload.severity || 'Info',
    userIds: Array.isArray(payload.userIds) ? payload.userIds : undefined,
    metadata: payload.metadata || { source: 'manager-executive-experience' }
  });
  return result || { success: true };
}

function clientOnboardManagedUser(payload) {
  payload = payload || {};
  if (!payload.email) {
    throw new Error('Email address is required for onboarding.');
  }
  if (!payload.campaignId) {
    throw new Error('Campaign selection is required for onboarding.');
  }

  var userData = {
    fullName: payload.fullName || '',
    userName: payload.userName || (payload.email ? payload.email.split('@')[0] : ''),
    email: payload.email,
    campaignId: payload.campaignId,
    employmentStatus: payload.employmentStatus || 'Active',
    roles: Array.isArray(payload.roleIds) ? payload.roleIds : [],
    pages: Array.isArray(payload.pageKeys) ? payload.pageKeys : [],
    canLogin: payload.sendInvite === true,
    mergeIfExists: payload.mergeIfExists === true,
    permissionLevel: payload.permissionLevel || 'VIEWER',
    canManageUsers: payload.canManageUsers === true,
    canManagePages: payload.canManagePages === true
  };

  if (typeof clientRegisterUser !== 'function') {
    throw new Error('User registration service is unavailable.');
  }

  var result = clientRegisterUser(userData);
  if (!result || result.success === false) {
    return result;
  }

  var onboardedUserId = result.userId;
  if (payload.grantGuestAccess && onboardedUserId && typeof setCampaignUserPermissions === 'function') {
    try {
      setCampaignUserPermissions(payload.campaignId, onboardedUserId, 'GUEST', false, false);
    } catch (err) {
      console.warn('clientOnboardManagedUser guest access assignment failed:', err);
    }
  }

  return result;
}

function clientSetCampaignGuestAccess(payload) {
  payload = payload || {};
  if (!payload.userId) throw new Error('A user selection is required.');
  if (!payload.campaignId) throw new Error('A campaign selection is required.');
  if (typeof setCampaignUserPermissions !== 'function') {
    throw new Error('Campaign permissions service is unavailable.');
  }

  var permissionLevel = payload.permissionLevel || 'GUEST';
  var canManageUsers = payload.canManageUsers === true;
  var canManagePages = payload.canManagePages === true;
  return setCampaignUserPermissions(payload.campaignId, payload.userId, permissionLevel, canManageUsers, canManagePages);
}

function formatIsoDate_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }
  var str = String(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
  var date = new Date(str);
  if (isNaN(date)) return '';
  return date.toISOString().slice(0, 10);
}

function formatIsoDateTime_(value) {
  if (!value) return '';
  if (value instanceof Date) return value.toISOString();
  var str = String(value);
  var date = new Date(str);
  if (isNaN(date)) return '';
  return date.toISOString();
}
