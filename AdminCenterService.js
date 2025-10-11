(function (global) {
  'use strict';

  var root = (typeof global !== 'undefined' && global) || (typeof this !== 'undefined' && this) || {};
  var G = (typeof root.G === 'object' && root.G) || (typeof G === 'object' && G) || {};

  function toNumber(value) {
    if (value === null || typeof value === 'undefined') return null;
    var num = Number(value);
    return isNaN(num) ? null : num;
  }

  function toIsoString(date) {
    if (date instanceof Date && !isNaN(date.getTime())) {
      return date.toISOString();
    }
    return new Date().toISOString();
  }

  function safeArray(value) {
    return Array.isArray(value) ? value : [];
  }

  function truthy(value) {
    if (value === null || typeof value === 'undefined') return false;
    if (typeof value === 'boolean') return value;
    if (typeof value === 'number') return value !== 0;
    var str = String(value).toLowerCase();
    if (!str) return false;
    return str === 'true' || str === '1' || str === 'yes' || str === 'y';
  }

  function normalizeStatus(value) {
    if (value === null || typeof value === 'undefined') return '';
    return String(value).trim().toLowerCase();
  }

  function isActiveStatus(record) {
    var statusKeys = ['EmploymentStatus', 'Status', 'State', 'Active'];
    for (var i = 0; i < statusKeys.length; i++) {
      var key = statusKeys[i];
      if (!record || !Object.prototype.hasOwnProperty.call(record, key)) continue;
      var val = record[key];
      if (typeof val === 'boolean') return val;
      var normalized = normalizeStatus(val);
      if (!normalized) continue;
      if (normalized === 'active' || normalized === 'enabled' || normalized === 'current') return true;
      if (normalized === 'inactive' || normalized === 'terminated' || normalized === 'disabled') return false;
    }
    return false;
  }

  function isAdminRecord(record) {
    if (!record) return false;
    var keys = ['isAdminBool', 'IsAdmin', 'isAdmin'];
    for (var i = 0; i < keys.length; i++) {
      var key = keys[i];
      if (!Object.prototype.hasOwnProperty.call(record, key)) continue;
      if (truthy(record[key])) return true;
    }
    if (Array.isArray(record.Roles)) {
      for (var j = 0; j < record.Roles.length; j++) {
        if (/admin/i.test(String(record.Roles[j]))) return true;
      }
    }
    return false;
  }

  function safeWorkspace(userId, options) {
    try {
      if (root.CallCenterWorkflowService && typeof root.CallCenterWorkflowService.getWorkspace === 'function') {
        return root.CallCenterWorkflowService.getWorkspace(userId, options || {});
      }
    } catch (err) {
      try { root.safeWriteError && root.safeWriteError('AdminCenter.getWorkspace', err); } catch (_) {}
    }
    return null;
  }

  function safeCampaignStats() {
    try {
      if (typeof root.csGetCampaignStatsV2 === 'function') {
        return root.csGetCampaignStatsV2();
      }
      if (typeof root.csGetCampaignStats === 'function') {
        return root.csGetCampaignStats();
      }
    } catch (err) {
      try { root.safeWriteError && root.safeWriteError('AdminCenter.campaignStats', err); } catch (_) {}
    }
    return [];
  }

  function deriveCollaborationVelocity(collaboration, campaigns) {
    var totalMessages = 0;
    if (collaboration && typeof collaboration.totalMessages === 'number') {
      totalMessages = collaboration.totalMessages;
    } else if (collaboration && Array.isArray(collaboration.messages)) {
      totalMessages = collaboration.messages.length;
    }
    var campaignCount = campaigns && campaigns.length ? campaigns.length : 1;
    return Math.round((totalMessages / campaignCount) * 100) / 100;
  }

  function buildSummary(workspace, campaignStats) {
    var users = (workspace && workspace.users && safeArray(workspace.users.records)) || [];
    var campaigns = safeArray(workspace && workspace.campaigns);
    var performance = workspace && workspace.performance || {};
    var coaching = workspace && workspace.coaching || {};
    var scheduling = workspace && workspace.scheduling || {};
    var collaboration = workspace && workspace.collaboration || {};

    var qaSummary = performance.qa && performance.qa.summary || {};
    var attendanceSummary = performance.attendance && performance.attendance.summary || {};
    var adherenceSummary = performance.adherence && performance.adherence.summary || {};
    var coachingSummary = coaching.summary || {};

    var totalUsers = users.length;
    var adminCount = 0;
    var activeCount = 0;
    for (var i = 0; i < users.length; i++) {
      if (isAdminRecord(users[i])) adminCount += 1;
      if (isActiveStatus(users[i])) activeCount += 1;
    }

    var attendanceRate = toNumber(attendanceSummary.attendanceRate);
    var qaAverage = toNumber(qaSummary.averageScore);
    var adherenceAverage = toNumber(adherenceSummary.averageScore);
    var pendingCoaching = toNumber(coachingSummary.pendingCount);
    var overdueCoaching = toNumber(coachingSummary.overdueCount);

    var adminCoverage = totalUsers > 0 ? Math.round((adminCount / totalUsers) * 10000) / 100 : null;
    var scheduleToday = safeArray(scheduling.today).length;
    var scheduleUpcoming = safeArray(scheduling.upcoming).length;
    var collaborationVelocity = deriveCollaborationVelocity(collaboration, campaigns);

    return {
      totals: {
        users: totalUsers,
        activeUsers: activeCount,
        admins: adminCount,
        campaigns: campaigns.length,
        qaEvaluations: toNumber(qaSummary.totalEvaluations) || 0,
        attendanceEvents: toNumber(attendanceSummary.totalEvents) || 0,
        coachingSessions: toNumber(coachingSummary.totalSessions) || 0,
        todayShifts: scheduleToday,
        upcomingShifts: scheduleUpcoming,
        collaborationMessages: collaboration && typeof collaboration.totalMessages === 'number'
          ? collaboration.totalMessages
          : safeArray(collaboration.messages).length
      },
      health: {
        qaAverage: qaAverage,
        attendanceRate: attendanceRate,
        adherenceAverage: adherenceAverage,
        pendingCoaching: pendingCoaching,
        overdueCoaching: overdueCoaching,
        adminCoverage: adminCoverage,
        collaborationVelocity: collaborationVelocity
      },
      campaignStats: Array.isArray(campaignStats) ? campaignStats : [],
      workspaceAvailable: !!workspace
    };
  }

  function evaluateMetric(metricKey, summary) {
    var health = summary.health || {};
    switch (metricKey) {
      case 'qaAverage':
        return classifyThreshold(health.qaAverage, { good: 90, warning: 80 }, true);
      case 'attendanceRate':
        return classifyThreshold(health.attendanceRate, { good: 97, warning: 93 }, true);
      case 'adherenceAverage':
        return classifyThreshold(health.adherenceAverage, { good: 95, warning: 90 }, true);
      case 'pendingCoaching':
        return classifyThreshold(health.pendingCoaching, { good: 5, warning: 10 }, false);
      case 'adminCoverage':
        return classifyThreshold(health.adminCoverage, { good: 12, warning: 8 }, true);
      case 'collaborationVelocity':
        return classifyThreshold(health.collaborationVelocity, { good: 20, warning: 10 }, true);
      default:
        return { level: 'unknown', label: 'No Signal' };
    }
  }

  function classifyThreshold(value, thresholds, higherIsBetter) {
    if (value === null || typeof value === 'undefined') {
      return { level: 'unknown', label: 'No Data' };
    }
    if (higherIsBetter) {
      if (value >= thresholds.good) return { level: 'good', label: 'Healthy' };
      if (value >= thresholds.warning) return { level: 'warning', label: 'Watch' };
      return { level: 'critical', label: 'Critical' };
    }
    if (value <= thresholds.good) return { level: 'good', label: 'Healthy' };
    if (value <= thresholds.warning) return { level: 'warning', label: 'Watch' };
    return { level: 'critical', label: 'Critical' };
  }

  function buildMetric(label, value, options) {
    var metric = {
      label: label,
      value: value
    };
    if (options) {
      if (options.unit) metric.unit = options.unit;
      if (options.helpText) metric.helpText = options.helpText;
      if (options.key) metric.key = options.key;
      if (options.status) metric.status = options.status;
    }
    return metric;
  }

  function buildModules(summary, workspace) {
    var modules = [];
    var totals = summary.totals || {};
    var health = summary.health || {};
    var campaigns = safeArray(workspace && workspace.campaigns);

    var identityModule = {
      key: 'identity',
      title: 'Identity & Access',
      icon: 'fas fa-user-shield',
      description: 'Oversee administrators, user lifecycle, and permissions across Lumina.',
      status: evaluateMetric('adminCoverage', summary),
      metrics: [
        buildMetric('Total Users', totals.users, { unit: '', helpText: 'Users discovered in the centralized directory.' }),
        buildMetric('Active Workforce', totals.activeUsers, { unit: '', helpText: 'Users marked as active or enabled.' }),
        buildMetric('Administrators', totals.admins, {
          unit: '',
          helpText: 'Users flagged with administrator privileges.',
          key: 'adminCoverage',
          status: evaluateMetric('adminCoverage', summary)
        })
      ],
      links: [
        { label: 'Manage Users', page: 'manageuser', icon: 'fas fa-users-cog' },
        { label: 'Manage Roles', page: 'manageroles', icon: 'fas fa-user-cog' }
      ]
    };
    modules.push(identityModule);

    var performanceModule = {
      key: 'performance',
      title: 'Quality & Performance',
      icon: 'fas fa-chart-line',
      description: 'Track QA scoring, attendance, and adherence to elevate customer outcomes.',
      status: evaluateMetric('qaAverage', summary),
      metrics: [
        buildMetric('QA Average', health.qaAverage, {
          unit: '%',
          helpText: 'Average QA score across evaluations.',
          key: 'qaAverage',
          status: evaluateMetric('qaAverage', summary)
        }),
        buildMetric('Attendance Rate', health.attendanceRate, {
          unit: '%',
          helpText: 'Present vs. total attendance events.',
          key: 'attendanceRate',
          status: evaluateMetric('attendanceRate', summary)
        }),
        buildMetric('Adherence Average', health.adherenceAverage, {
          unit: '%',
          helpText: 'Average schedule adherence across records.',
          key: 'adherenceAverage',
          status: evaluateMetric('adherenceAverage', summary)
        }),
        buildMetric('Pending Coaching', health.pendingCoaching, {
          unit: 'sessions',
          helpText: 'Coaching sessions awaiting acknowledgement or completion.',
          key: 'pendingCoaching',
          status: evaluateMetric('pendingCoaching', summary)
        })
      ],
      links: [
        { label: 'QA Dashboard', page: 'qadashboard', icon: 'fas fa-clipboard-check' },
        { label: 'Coaching Dashboard', page: 'coachingdashboard', icon: 'fas fa-chalkboard-teacher' }
      ]
    };
    modules.push(performanceModule);

    var workforceModule = {
      key: 'workforce',
      title: 'Workforce Operations',
      icon: 'fas fa-people-arrows',
      description: 'Manage scheduling coverage, upcoming shifts, and coaching follow-ups.',
      status: evaluateMetric('attendanceRate', summary),
      metrics: [
        buildMetric('Shifts Today', totals.todayShifts, { unit: 'shifts', helpText: 'Scheduled shifts occurring today.' }),
        buildMetric('Upcoming Shifts', totals.upcomingShifts, { unit: 'shifts', helpText: 'Future scheduled shifts queued.' }),
        buildMetric('Coaching Overdue', health.overdueCoaching, {
          unit: 'sessions',
          helpText: 'Coaching commitments past due date.',
          key: 'pendingCoaching',
          status: evaluateMetric('pendingCoaching', summary)
        })
      ],
      links: [
        { label: 'Schedule Management', page: 'schedulemanagement', icon: 'fas fa-calendar-alt' },
        { label: 'Attendance Reports', page: 'attendancereports', icon: 'fas fa-user-clock' }
      ]
    };
    modules.push(workforceModule);

    var collaborationModule = {
      key: 'collaboration',
      title: 'Collaboration & Broadcasts',
      icon: 'fas fa-comments',
      description: 'Monitor communications velocity and broadcast updates to every campaign.',
      status: evaluateMetric('collaborationVelocity', summary),
      metrics: [
        buildMetric('Messages (Last Snapshot)', totals.collaborationMessages, {
          unit: 'messages',
          helpText: 'Recent collaboration entries included in the workspace snapshot.'
        }),
        buildMetric('Avg. Messages per Campaign', health.collaborationVelocity, {
          unit: 'msg/campaign',
          helpText: 'Engagement signal normalized per active campaign.',
          key: 'collaborationVelocity',
          status: evaluateMetric('collaborationVelocity', summary)
        }),
        buildMetric('Active Campaigns', totals.campaigns, { unit: '', helpText: 'Campaigns accessible to this administrator.' })
      ],
      links: [
        { label: 'Notifications', page: 'notifications', icon: 'fas fa-bullhorn' },
        { label: 'Team Chat', page: 'chat', icon: 'fas fa-comments' }
      ]
    };
    modules.push(collaborationModule);

    var platformModule = {
      key: 'platform',
      title: 'Platform Governance',
      icon: 'fas fa-solar-panel',
      description: 'Orchestrate caches, navigation, and cross-campaign governance.',
      status: evaluateMetric('adminCoverage', summary),
      metrics: [
        buildMetric('Campaign Snapshots', summary.campaignStats.length, {
          unit: 'campaigns',
          helpText: 'Campaigns with aggregated stats available in the admin center.'
        }),
        buildMetric('QA Records Cached', summary.totals.qaEvaluations, {
          unit: 'evaluations',
          helpText: 'Evaluations counted in the latest workspace snapshot.'
        }),
        buildMetric('Attendance Events Cached', summary.totals.attendanceEvents, {
          unit: 'events',
          helpText: 'Attendance entries contributing to the current health score.'
        })
      ],
      links: [
        { label: 'Campaign Management', page: 'managecampaign', icon: 'fas fa-bullhorn' },
        { label: 'Lumina User Guide', page: 'lumina-hq-user-guide', icon: 'fas fa-book-open' }
      ]
    };
    modules.push(platformModule);

    return modules;
  }

  function buildAiInsights(summary, modules) {
    var insights = [];
    var health = summary.health || {};
    var totals = summary.totals || {};

    if (health.qaAverage !== null && health.qaAverage < 85) {
      insights.push({
        title: 'Quality Dip Detected',
        severity: health.qaAverage < 80 ? 'critical' : 'warning',
        message: 'Average QA performance has fallen to ' + health.qaAverage + '%. Focus reviews on high-risk campaigns and deploy targeted coaching.',
        recommendation: 'Prioritize QA remediation workflows within the Quality & Performance module.'
      });
    }

    if (health.attendanceRate !== null && health.attendanceRate < 95) {
      insights.push({
        title: 'Attendance Risk',
        severity: health.attendanceRate < 90 ? 'critical' : 'warning',
        message: 'Attendance rate currently sits at ' + health.attendanceRate + '%. Investigate callouts and schedule adjustments before service levels slip.',
        recommendation: 'Review Attendance Reports and ensure coverage in Schedule Management.'
      });
    }

    if (health.pendingCoaching !== null && health.pendingCoaching > Math.max(5, Math.round(totals.activeUsers * 0.15))) {
      insights.push({
        title: 'Coaching Backlog',
        severity: 'warning',
        message: health.pendingCoaching + ' coaching sessions are still pending acknowledgement. Agents may be waiting on critical feedback.',
        recommendation: 'Drive completion via the Coaching Dashboard and send reminders from the Workforce Operations module.'
      });
    }

    if (health.adminCoverage !== null && health.adminCoverage < 7) {
      insights.push({
        title: 'Low Admin Coverage',
        severity: 'warning',
        message: 'Only ' + health.adminCoverage + '% of users hold administrator privileges. Ensure redundancy for critical operations.',
        recommendation: 'Nominate backup administrators or delegate advanced roles in Identity & Access.'
      });
    }

    if (health.collaborationVelocity !== null && health.collaborationVelocity < 5) {
      insights.push({
        title: 'Quiet Collaboration Channels',
        severity: 'info',
        message: 'Collaboration activity is trending low at ' + health.collaborationVelocity + ' messages per campaign. Consider sharing updates to maintain engagement.',
        recommendation: 'Publish a broadcast through Notifications or encourage campaign leads to share wins in chat.'
      });
    }

    if (!insights.length) {
      insights.push({
        title: 'All Systems Steady',
        severity: 'success',
        message: 'Key health indicators are within target ranges. Continue monitoring dashboards for early signals.',
        recommendation: 'Leverage automation controls below to keep caches warm and data flowing.'
      });
    }

    return insights;
  }

  var CONTROL_ACTIONS = [
    {
      key: 'identity-refresh-user-cache',
      moduleKey: 'identity',
      label: 'Refresh User Directory Cache',
      description: 'Clears cached user records to ensure the admin center pulls the latest identities.',
      icon: 'fas fa-sync',
      variant: 'primary',
      execute: function () {
        if (typeof root.invalidateCache === 'function') {
          var sheetName = (typeof root.USERS_SHEET === 'string' && root.USERS_SHEET) || (G && G.USERS_SHEET) || 'Users';
          root.invalidateCache(sheetName);
          return { status: 'ok', message: 'User cache invalidated for "' + sheetName + '".' };
        }
        return { status: 'skipped', message: 'Cache service not available in this deployment.' };
      }
    },
    {
      key: 'performance-refresh-qa-cache',
      moduleKey: 'performance',
      label: 'Rebuild QA Snapshot',
      description: 'Flushes QA evaluation caches and reinitializes the workflow registry.',
      icon: 'fas fa-clipboard-check',
      variant: 'secondary',
      execute: function () {
        var responses = [];
        if (typeof root.invalidateCache === 'function') {
          var qaSheet = (typeof root.QA_RECORDS === 'string' && root.QA_RECORDS) || (G && G.QA_RECORDS) || 'Quality';
          root.invalidateCache(qaSheet);
          responses.push('QA cache cleared for "' + qaSheet + '".');
        }
        if (root.CallCenterWorkflowService && typeof root.CallCenterWorkflowService.initialize === 'function') {
          var registry = root.CallCenterWorkflowService.initialize();
          var keys = registry ? Object.keys(registry) : [];
          responses.push('Workflow registry refreshed' + (keys.length ? ' (' + keys.length + ' tables).' : '.'));
        }
        return {
          status: responses.length ? 'ok' : 'skipped',
          message: responses.join(' ')
        };
      }
    },
    {
      key: 'workforce-refresh-schedule-cache',
      moduleKey: 'workforce',
      label: 'Refresh Schedule Caches',
      description: 'Clears schedule caches to pull updated shift assignments.',
      icon: 'fas fa-calendar-week',
      variant: 'secondary',
      execute: function () {
        if (typeof root.invalidateScheduleCaches === 'function') {
          root.invalidateScheduleCaches();
          return { status: 'ok', message: 'Schedule caches invalidated successfully.' };
        }
        if (typeof root.invalidateCache === 'function') {
          var scheduleSheet = (typeof G.SCHEDULES_SHEET === 'string' && G.SCHEDULES_SHEET) || 'Schedules';
          root.invalidateCache(scheduleSheet);
          return { status: 'ok', message: 'Schedule sheet cache cleared for "' + scheduleSheet + '".' };
        }
        return { status: 'skipped', message: 'No schedule cache helpers available.' };
      }
    },
    {
      key: 'collaboration-clear-notification-cache',
      moduleKey: 'collaboration',
      label: 'Clear Broadcast Cache',
      description: 'Invalidates notification caches so new announcements publish instantly.',
      icon: 'fas fa-bullhorn',
      variant: 'secondary',
      execute: function () {
        if (typeof root.invalidateCache === 'function') {
          var sheet = (typeof root.NOTIFICATIONS_SHEET === 'string' && root.NOTIFICATIONS_SHEET) || (G && G.NOTIFICATIONS_SHEET) || 'Notifications';
          root.invalidateCache(sheet);
          return { status: 'ok', message: 'Notification cache cleared for "' + sheet + '".' };
        }
        return { status: 'skipped', message: 'Notification cache helper unavailable.' };
      }
    },
    {
      key: 'platform-refresh-navigation-cache',
      moduleKey: 'platform',
      label: 'Refresh Navigation Cache',
      description: 'Invalidates campaign navigation caches to surface new pages immediately.',
      icon: 'fas fa-compass',
      variant: 'secondary',
      execute: function (payload, userId) {
        var refreshed = 0;
        if (typeof root.CallCenterWorkflowService === 'object' && root.CallCenterWorkflowService && typeof root.CallCenterWorkflowService.listCampaignAccess === 'function' && userId) {
          try {
            var campaigns = root.CallCenterWorkflowService.listCampaignAccess(userId) || [];
            for (var i = 0; i < campaigns.length; i++) {
              var cid = campaigns[i] && (campaigns[i].id || campaigns[i].ID || campaigns[i].Id);
              if (!cid) continue;
              if (typeof root.invalidateNavigationCache === 'function') {
                root.invalidateNavigationCache(cid);
                refreshed += 1;
              }
            }
          } catch (navErr) {
            try { root.safeWriteError && root.safeWriteError('AdminCenter.navigation', navErr); } catch (_) {}
          }
        }

        if (!refreshed && typeof root.invalidateNavigationCache === 'function') {
          var fallbackCampaign = (payload && payload.campaignId) || (G && G.DEFAULT_CAMPAIGN_ID) || null;
          if (fallbackCampaign) {
            root.invalidateNavigationCache(fallbackCampaign);
            refreshed = 1;
          }
        }

        if (typeof root.invalidateCache === 'function') {
          var sheets = ['PAGES_SHEET', 'CAMPAIGN_PAGES_SHEET', 'PAGE_CATEGORIES_SHEET'];
          for (var s = 0; s < sheets.length; s++) {
            var key = sheets[s];
            var name = (typeof root[key] === 'string' && root[key]) || (G && G[key]) || null;
            if (name) {
              try { root.invalidateCache(name); } catch (_) {}
            }
          }
        }

        if (refreshed) {
          return { status: 'ok', message: 'Navigation caches refreshed for ' + refreshed + ' campaign(s).' };
        }
        return { status: 'skipped', message: 'Navigation cache helper unavailable or no campaigns resolved.' };
      }
    }
  ];

  function exportAction(action) {
    return {
      key: action.key,
      moduleKey: action.moduleKey,
      label: action.label,
      description: action.description,
      icon: action.icon,
      variant: action.variant || 'secondary'
    };
  }

  function runControlAction(key, payload, userId) {
    for (var i = 0; i < CONTROL_ACTIONS.length; i++) {
      var action = CONTROL_ACTIONS[i];
      if (action.key === key) {
        try {
          var result = action.execute(payload || {}, userId);
          if (!result || typeof result !== 'object') {
            return { status: 'ok', message: 'Action executed.' };
          }
          return result;
        } catch (err) {
          try { root.safeWriteError && root.safeWriteError('AdminCenter.action.' + key, err); } catch (_) {}
          return { status: 'error', message: err && err.message ? err.message : 'Unexpected error' };
        }
      }
    }
    return { status: 'error', message: 'Unknown action: ' + key };
  }

  var AdminCenterService = {
    getAdminCenterSnapshot: function (userId, options) {
      var workspace = safeWorkspace(userId, options || {});
      var campaignStats = safeCampaignStats();
      var summary = buildSummary(workspace, campaignStats);
      var modules = buildModules(summary, workspace);
      var insights = buildAiInsights(summary, modules);
      return {
        generatedAt: toIsoString(new Date()),
        summary: summary,
        modules: modules,
        aiInsights: insights,
        controlActions: CONTROL_ACTIONS.map(exportAction)
      };
    },
    listControlActions: function () {
      return CONTROL_ACTIONS.map(exportAction);
    },
    runControlAction: runControlAction
  };

  root.AdminCenterService = AdminCenterService;
  root.clientGetAdminCenterSnapshot = function (userId, options) {
    return AdminCenterService.getAdminCenterSnapshot(userId, options || {});
  };
  root.clientListAdminControlActions = function () {
    return AdminCenterService.listControlActions();
  };
  root.clientRunAdminControlAction = function (actionKey, payload, userId) {
    return AdminCenterService.runControlAction(actionKey, payload || {}, userId);
  };

})(typeof globalThis !== 'undefined' ? globalThis : this);
