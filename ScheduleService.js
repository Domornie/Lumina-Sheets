/**
 * COMPLETE Enhanced Schedule Management Backend Service
 * Version 4.1 - Integrated with ScheduleUtilities and MainUtilities
 * Now properly uses dedicated spreadsheet support and shared functions
 */

// ────────────────────────────────────────────────────────────────────────────
// CONFIGURATION - Uses ScheduleUtilities constants
// ────────────────────────────────────────────────────────────────────────────

const SCHEDULE_SETTINGS = (typeof getScheduleConfig === 'function')
  ? getScheduleConfig()
  : {
      PRIMARY_COUNTRY: 'JM',
      SUPPORTED_COUNTRIES: ['JM', 'US', 'DO', 'PH'],
      DEFAULT_SHIFT_CAPACITY: 10,
      DEFAULT_BREAK_MINUTES: 15,
      DEFAULT_LUNCH_MINUTES: 60,
      CACHE_DURATION: 300
    };

const DEFAULT_SCHEDULE_TIME_ZONE = (typeof Session !== 'undefined' && typeof Session.getScriptTimeZone === 'function')
  ? Session.getScriptTimeZone()
  : 'UTC';

function resolveSchedulePeriodStart(record, timeZone = DEFAULT_SCHEDULE_TIME_ZONE) {
  if (!record || typeof record !== 'object') {
    return '';
  }

  const candidates = [
    record.PeriodStart,
    record.StartDate,
    record.ScheduleStart,
    record.AssignmentStart,
    record.Date,
    record.ScheduleDate,
    record.Day
  ];

  for (let i = 0; i < candidates.length; i++) {
    const normalized = normalizeDateForSheet(candidates[i], timeZone);
    if (normalized) {
      return normalized;
    }
  }

  return '';
}

function resolveSchedulePeriodEnd(record, fallbackStart = '', timeZone = DEFAULT_SCHEDULE_TIME_ZONE) {
  if (!record || typeof record !== 'object') {
    return '';
  }

  const candidates = [
    record.PeriodEnd,
    record.EndDate,
    record.ScheduleEnd,
    record.AssignmentEnd,
    record.Date,
    record.ScheduleDate,
    record.Day,
    fallbackStart
  ];

  for (let i = 0; i < candidates.length; i++) {
    const normalized = normalizeDateForSheet(candidates[i], timeZone);
    if (normalized) {
      return normalized;
    }
  }

  return '';
}

function resolveSchedulePeriodStartDate(record, timeZone = DEFAULT_SCHEDULE_TIME_ZONE) {
  const start = resolveSchedulePeriodStart(record, timeZone);
  if (!start) {
    return null;
  }

  const startDate = new Date(start);
  return isNaN(startDate.getTime()) ? null : startDate;
}

function resolveSchedulePeriodEndDate(record, timeZone = DEFAULT_SCHEDULE_TIME_ZONE) {
  const start = resolveSchedulePeriodStart(record, timeZone);
  const end = resolveSchedulePeriodEnd(record, start, timeZone);
  if (!end) {
    return null;
  }

  const endDate = new Date(end);
  return isNaN(endDate.getTime()) ? null : endDate;
}

function normalizeSchedulePeriodRecord(record, timeZone = DEFAULT_SCHEDULE_TIME_ZONE) {
  if (!record || typeof record !== 'object') {
    return record;
  }

  const normalizedStart = resolveSchedulePeriodStart(record, timeZone);
  const normalizedEnd = resolveSchedulePeriodEnd(record, normalizedStart, timeZone);

  if (!normalizedStart && !normalizedEnd) {
    return record;
  }

  const normalizedRecord = Object.assign({}, record);

  if (normalizedStart) {
    normalizedRecord.PeriodStart = normalizedStart;
    normalizedRecord.Date = normalizedStart;
  }

  if (normalizedEnd) {
    normalizedRecord.PeriodEnd = normalizedEnd;
  }

  return normalizedRecord;
}

function buildScheduleCompositeKey(record, timeZone = DEFAULT_SCHEDULE_TIME_ZONE) {
  const normalizedRecord = normalizeSchedulePeriodRecord(record, timeZone);
  const userPart = normalizeUserKey(
    (normalizedRecord && (normalizedRecord.UserName || normalizedRecord.UserID || normalizedRecord.userName || normalizedRecord.userId))
      || ''
  );

  const start = normalizedRecord ? normalizedRecord.PeriodStart || '' : '';
  const end = normalizedRecord ? normalizedRecord.PeriodEnd || start : '';

  return `${userPart}::${start}::${end}`;
}

function getSchedulePeriodSortValue(record, timeZone = DEFAULT_SCHEDULE_TIME_ZONE) {
  const startDate = resolveSchedulePeriodStartDate(record, timeZone);
  return startDate ? startDate.getTime() : 0;
}

// ────────────────────────────────────────────────────────────────────────────
// USER MANAGEMENT FUNCTIONS - Integrated with MainUtilities
// ────────────────────────────────────────────────────────────────────────────

/**
 * Resolve the active schedule context for the requesting user or manager.
 * Provides manager/campaign identifiers, identity metadata, and managed roster.
 */
function clientGetScheduleContext(managerIdCandidate, campaignIdCandidate) {
  const providedManagerId = normalizeUserIdValue(managerIdCandidate);
  const providedCampaignId = normalizeCampaignIdValue(campaignIdCandidate);
  const timestamp = new Date().toISOString();

  const context = {
    success: false,
    providedManagerId,
    providedCampaignId,
    managerId: '',
    campaignId: '',
    user: null,
    managedUserIds: [],
    managedUserCount: 0,
    managedCampaigns: [],
    identity: null,
    authenticated: false,
    timestamp
  };

  try {
    const allUsers = readSheet(USERS_SHEET) || [];
    const usersById = new Map();
    const usersByUsername = new Map();

    allUsers.forEach(user => {
      if (!user || typeof user !== 'object') {
        return;
      }

      const normalizedId = normalizeUserIdValue(user.ID);
      if (normalizedId) {
        usersById.set(normalizedId, user);
      }

      const normalizedUsername = normalizeUserIdValue(user.UserName || user.Username);
      if (normalizedUsername) {
        const usernameKey = normalizedUsername.toLowerCase();
        if (!usersByUsername.has(usernameKey)) {
          usersByUsername.set(usernameKey, user);
        }
      }
    });

    const currentUser = (typeof getCurrentUser === 'function') ? getCurrentUser() : null;
    const currentUserId = normalizeUserIdValue(currentUser && (currentUser.ID || currentUser.UserID));

    const managerIdSources = [];
    let resolvedManagerId = providedManagerId;

    if (resolvedManagerId) {
      managerIdSources.push('parameter');
    }

    if (!resolvedManagerId && currentUserId) {
      resolvedManagerId = currentUserId;
      managerIdSources.push('current-user');
    }

    let resolvedManagerUser = resolvedManagerId ? usersById.get(resolvedManagerId) : null;

    if (!resolvedManagerUser && resolvedManagerId) {
      const key = resolvedManagerId.toLowerCase ? resolvedManagerId.toLowerCase() : String(resolvedManagerId || '').toLowerCase();
      const usernameMatch = usersByUsername.get(key);
      if (usernameMatch) {
        resolvedManagerUser = usernameMatch;
        resolvedManagerId = normalizeUserIdValue(usernameMatch.ID) || resolvedManagerId;
        managerIdSources.push('username-match');
      }
    }

    if (!resolvedManagerUser && currentUserId) {
      resolvedManagerUser = usersById.get(currentUserId) || currentUser || null;
    }

    const campaignIdSources = [];
    let resolvedCampaignId = providedCampaignId;

    if (resolvedCampaignId) {
      campaignIdSources.push('parameter');
    }

    const appendCampaignCandidate = (value, source) => {
      if (resolvedCampaignId) {
        return;
      }
      const normalized = normalizeCampaignIdValue(value);
      if (normalized) {
        resolvedCampaignId = normalized;
        campaignIdSources.push(source);
      }
    };

    appendCampaignCandidate(resolvedManagerUser && (resolvedManagerUser.CampaignID || resolvedManagerUser.campaignID || resolvedManagerUser.Campaign || resolvedManagerUser.campaign), 'manager-profile');
    appendCampaignCandidate(currentUser && (currentUser.CampaignID || currentUser.campaignID || currentUser.Campaign || currentUser.campaign), 'current-user');

    const managedSet = resolvedManagerId ? buildManagedUserSet(resolvedManagerId) : new Set();
    const managedUserIds = Array.from(managedSet).map(normalizeUserIdValue).filter(Boolean);

    let managedCampaigns = [];
    try {
      if (resolvedManagerId && typeof getUserManagedCampaigns === 'function') {
        const campaigns = getUserManagedCampaigns(resolvedManagerId) || [];
        managedCampaigns = campaigns
          .filter(Boolean)
          .map(campaign => ({
            id: normalizeCampaignIdValue(campaign.ID || campaign.Id || campaign.id),
            name: campaign.Name || campaign.name || '',
            isPrimary: scheduleFlagToBool(campaign.IsPrimary || campaign.isPrimary)
          }));

        if (!resolvedCampaignId) {
          const primary = managedCampaigns.find(campaign => campaign.isPrimary);
          if (primary && primary.id) {
            resolvedCampaignId = primary.id;
            campaignIdSources.push('managed-campaign');
          }
        }
      }
    } catch (campaignError) {
      console.warn('clientGetScheduleContext: unable to resolve managed campaigns', campaignError);
    }

    const roles = collectUserRoleCandidates(resolvedManagerUser || currentUser || {});
    const normalizedRoles = roles
      .map(role => String(role || '').trim())
      .filter(Boolean);

    const identity = {
      authenticated: !!currentUserId,
      resolvedAt: timestamp,
      managerId: resolvedManagerId || '',
      campaignId: resolvedCampaignId || '',
      providedManagerId,
      providedCampaignId,
      managerIdSources,
      campaignIdSources,
      managedUserCount: managedUserIds.length,
      roles: normalizedRoles,
      isAdmin: scheduleFlagToBool((resolvedManagerUser && resolvedManagerUser.IsAdmin) || (currentUser && currentUser.IsAdmin)),
      userId: currentUserId || resolvedManagerId || '',
      userName: (currentUser && (currentUser.UserName || currentUser.Username)) || '',
      fullName: (currentUser && (currentUser.FullName || currentUser.Name)) || '',
      email: (currentUser && (currentUser.Email || currentUser.email)) || ''
    };

    const clientUser = Object.assign({}, resolvedManagerUser || currentUser || {}, {
      ID: normalizeUserIdValue((resolvedManagerUser && resolvedManagerUser.ID) || (currentUser && currentUser.ID) || resolvedManagerId),
      CampaignID: resolvedCampaignId || (resolvedManagerUser && resolvedManagerUser.CampaignID) || '',
      Roles: normalizedRoles,
      IsAdmin: identity.isAdmin,
      managedUserCount: managedUserIds.length
    });

    if (!clientUser.UserName && clientUser.Username) {
      clientUser.UserName = clientUser.Username;
    }
    if (!clientUser.FullName && clientUser.Name) {
      clientUser.FullName = clientUser.Name;
    }

    context.permissions = {
      canManageSchedules: identity.isAdmin || managedUserIds.length > 0,
      canApproveSchedules: identity.isAdmin || managedUserIds.length > 0,
      canImport: identity.isAdmin
    };

    context.success = true;
    context.authenticated = identity.authenticated;
    context.managerId = resolvedManagerId || '';
    context.campaignId = resolvedCampaignId || '';
    context.user = clientUser;
    identity.permissions = context.permissions;

    context.identity = identity;
    context.managedUserIds = managedUserIds;
    context.managedUserCount = managedUserIds.length;
    context.managedCampaigns = managedCampaigns;

    return context;
  } catch (error) {
    console.error('❌ Error resolving schedule context:', error);
    context.error = error && error.message ? error.message : String(error || 'Unknown error');
    try {
      safeWriteError && safeWriteError('clientGetScheduleContext', error);
    } catch (_) {
      // ignore logging failures
    }
    return context;
  }
}

/**
 * Get users for schedule management with manager filtering
 * Uses MainUtilities user functions with campaign support
 */
function clientGetScheduleUsers(requestingUserId, campaignId = null) {
  try {
    const normalizedCampaignId = normalizeCampaignIdValue(campaignId);
    console.log('🔍 Getting schedule users for:', requestingUserId, 'campaign:', normalizedCampaignId || '(not provided)');

    // Use MainUtilities to get all users
    const allUsers = readSheet(USERS_SHEET) || [];
    if (allUsers.length === 0) {
      console.warn('No users found in Users sheet');
      return [];
    }

    const normalizedManagerId = normalizeUserIdValue(requestingUserId);
    let requestingUser = null;
    if (normalizedManagerId) {
      requestingUser = allUsers.find(u => normalizeUserIdValue(u && u.ID) === normalizedManagerId) || null;
    }

    let effectiveCampaignId = normalizedCampaignId;
    if (!effectiveCampaignId && requestingUser) {
      const managerCampaignCandidates = [
        requestingUser.CampaignID,
        requestingUser.campaignID,
        requestingUser.CampaignId,
        requestingUser.campaignId,
        requestingUser.Campaign,
        requestingUser.campaign
      ];

      for (let i = 0; i < managerCampaignCandidates.length; i++) {
        const candidate = normalizeCampaignIdValue(managerCampaignCandidates[i]);
        if (candidate) {
          effectiveCampaignId = candidate;
          break;
        }
      }
    }

    let filteredUsers = allUsers;

    // Filter by campaign if specified - use MainUtilities campaign functions
    if (effectiveCampaignId) {
      filteredUsers = filterUsersByCampaign(allUsers, effectiveCampaignId);
    }

    // Apply manager permissions using MainUtilities functions
    if (normalizedManagerId) {
      if (requestingUser) {
        const isAdmin = scheduleFlagToBool(requestingUser.IsAdmin);

        if (!isAdmin) {
          const managedUserIds = buildManagedUserSet(normalizedManagerId);

          filteredUsers = filteredUsers.filter(user => managedUserIds.has(normalizeUserIdValue(user && user.ID)));
        }
      } else {
        console.warn('Requesting user not found when applying manager filter:', requestingUserId);
      }
    }

    // Transform to schedule-friendly format
    const scheduleUsers = filteredUsers
      .filter(user => user && user.ID && (user.UserName || user.FullName))
      .filter(user => !isScheduleNameRestricted(user))
      .filter(user => !isScheduleRoleRestricted(user))
      .filter(user => isUserConsideredActive(user))
      .map(user => {
        const campaignName = getCampaignById(user.CampaignID)?.Name || '';
        return {
          ID: user.ID,
          UserName: user.UserName || user.FullName,
          FullName: user.FullName || user.UserName,
          Email: user.Email || '',
          CampaignID: user.CampaignID || '',
          campaignName: campaignName,
          EmploymentStatus: user.EmploymentStatus || 'Active',
          HireDate: user.HireDate || '',
          TerminationDate: user.TerminationDate || user.terminationDate || '',
          isActive: isUserConsideredActive(user)
        };
      });

    console.log(`✅ Returning ${scheduleUsers.length} schedule users`);
    return scheduleUsers;

  } catch (error) {
    console.error('❌ Error getting schedule users:', error);
    safeWriteError('clientGetScheduleUsers', error);
    return [];
  }
}

/**
 * Get users for attendance (all active users)
 */
function clientGetAttendanceUsers(requestingUserId, campaignId = null) {
  try {
    console.log('📋 Getting attendance users');
    
    // Use the existing function but return just names for compatibility
    const scheduleUsers = clientGetScheduleUsers(requestingUserId, campaignId);
    const userNames = scheduleUsers
      .map(user => user.UserName || user.FullName)
      .filter(name => name && name.trim())
      .sort();

    console.log(`✅ Returning ${userNames.length} attendance users`);
    return userNames;

  } catch (error) {
    console.error('❌ Error getting attendance users:', error);
    safeWriteError('clientGetAttendanceUsers', error);
    return [];
  }
}

/**
 * Get managed users list - delegates to MainUtilities
 */
function clientGetManagedUsersList(managerId) {
  try {
    if (!managerId) return [];

    // Use MainUtilities function for managed campaigns
    const managedCampaigns = getUserManagedCampaigns(managerId);
    const managedUsers = [];
    
    managedCampaigns.forEach(campaign => {
      const campaignUsers = getUsersByCampaign(campaign.ID);
      campaignUsers.forEach(user => {
        if (String(user.ID) !== String(managerId)) { // Don't include self
          managedUsers.push({
            ID: user.ID,
            UserName: user.UserName,
            FullName: user.FullName,
            Email: user.Email,
            CampaignID: user.CampaignID,
            campaignName: campaign.Name,
            EmploymentStatus: user.EmploymentStatus
          });
        }
      });
    });

    return managedUsers;

  } catch (error) {
    console.error('Error getting managed users:', error);
    safeWriteError('clientGetManagedUsersList', error);
    return [];
  }
}

function buildScheduleRecommendations(evaluation, bundle) {
  const recommendations = [];
  if (!evaluation || !evaluation.summary) {
    return recommendations;
  }

  const coverage = evaluation.coverage || {};
  const fairness = evaluation.fairness || {};
  const compliance = evaluation.compliance || {};

  if (Number(coverage.serviceLevel || 0) < 80) {
    const topInterval = (coverage.backlogRiskIntervals || [])[0];
    if (topInterval) {
      recommendations.push(`Add staffing to ${topInterval.intervalKey} for skill ${topInterval.skill || 'general'} (deficit ${topInterval.deficit} FTE).`);
    } else {
      recommendations.push('Increase staffing in critical intervals to protect service level.');
    }
  }

  if (Number(coverage.peakCoverage || 0) < 85) {
    recommendations.push('Rebalance opening and closing coverage to meet first/last hour SLAs.');
  }

  if (Number(fairness.rotationHealth || 0) < 75) {
    recommendations.push('Review weekend and night rotation to improve fairness.');
  }

  if (Number(compliance.complianceScore || 0) < 85) {
    recommendations.push('Resolve compliance issues (breaks, rest periods, overtime) before publishing schedules.');
  }

  if (!bundle || !bundle.scheduleRows || bundle.scheduleRows.length === 0) {
    recommendations.push('No schedules found for the selected filters. Import or generate schedules to proceed.');
  }

  return recommendations;
}

function persistScheduleHealthSnapshot(context, evaluation, bundle, options = {}) {
  if (!context || !evaluation || !evaluation.summary) {
    return;
  }

  if (typeof ensureScheduleSheetWithHeaders !== 'function') {
    return;
  }

  try {
    const sheet = ensureScheduleSheetWithHeaders(SCHEDULE_HEALTH_SHEET, SCHEDULE_HEALTH_HEADERS);
    const id = (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.getUuid === 'function')
      ? Utilities.getUuid()
      : `health_${Date.now()}`;

    const totalMinutes = (bundle.scheduleRows || []).reduce((sum, row) => {
      const start = safeNormalizeScheduleTimeToMinutes(row.StartTime || row.PeriodStart || row.ScheduleStart || row.ShiftStart);
      const end = safeNormalizeScheduleTimeToMinutes(row.EndTime || row.PeriodEnd || row.ScheduleEnd || row.ShiftEnd);
      if (start === null || end === null) {
        return sum;
      }
      let diff = end - start;
      if (diff < 0) {
        diff += 24 * 60;
      }
      return sum + diff;
    }, 0);

    const agentSet = new Set((bundle.scheduleRows || []).map(row => normalizeUserIdValue(row.UserID || row.UserId || row.AgentID || row.AgentId)).filter(Boolean));
    const totalHours = totalMinutes / 60;
    const standardHours = scheduleToNumber(options.standardHoursPerAgent || 8, 8);
    const overtimeHours = Math.max(0, totalHours - (agentSet.size * standardHours));

    const summary = evaluation.summary;
    const coverage = evaluation.coverage || {};
    const fairness = evaluation.fairness || {};
    const compliance = evaluation.compliance || {};
    const row = [
      id,
      new Date(),
      context.managerId || '',
      context.campaignId || context.providedCampaignId || '',
      Number(summary.healthScore || evaluation.healthScore || 0),
      Number(coverage.serviceLevel || 0),
      Number(fairness.rotationHealth || 0),
      Number(compliance.complianceScore || 0),
      totalHours,
      overtimeHours,
      agentSet.size,
      JSON.stringify(summary || {}),
      JSON.stringify(evaluation.coverage || {}),
      JSON.stringify(evaluation.fairness || {}),
      JSON.stringify(evaluation.compliance || {}),
      JSON.stringify(bundle.scheduleRows ? bundle.scheduleRows.slice(0, 20) : []),
      JSON.stringify(bundle.demandRows ? bundle.demandRows.slice(0, 20) : []),
      JSON.stringify(bundle.agentProfiles ? bundle.agentProfiles.slice(0, 20) : [])
    ];

    sheet.appendRow(row);
  } catch (error) {
    console.error('Error persisting schedule health snapshot:', error);
    safeWriteError && safeWriteError('persistScheduleHealthSnapshot', error);
  }
}

// ────────────────────────────────────────────────────────────────────────────
function clientGetAttendanceDataRange(startDate, endDate, campaignId = null) {
  try {
    const attendanceData = readScheduleSheet(ATTENDANCE_STATUS_SHEET) || [];

    const normalizeDate = (value) => {
      if (value instanceof Date) {
        return new Date(value.getTime());
      }
      if (typeof value === 'number') {
        const parsed = new Date(value);
        return isNaN(parsed.getTime()) ? null : parsed;
      }
      if (typeof value === 'string' && value.trim().length > 0) {
        const parsed = new Date(value);
        return isNaN(parsed.getTime()) ? null : parsed;
      }
      return null;
    };

    const rangeStart = normalizeDate(startDate);
    const rangeEnd = normalizeDate(endDate);

    const filtered = attendanceData.filter(record => {
      const recordDate = normalizeDate(record.Date || record.date);
      if (!recordDate) {
        return false;
      }

      if (rangeStart && recordDate < rangeStart) {
        return false;
      }
      if (rangeEnd && recordDate > rangeEnd) {
        return false;
      }

      if (campaignId) {
        const recordCampaign = record.CampaignID || record.CampaignId || record.Campaign || null;
        if (recordCampaign && recordCampaign !== campaignId) {
          return false;
        }
      }

      return true;
    });

    const toIsoDate = (date) => {
      if (!(date instanceof Date) || isNaN(date.getTime())) {
        return '';
      }
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    };

    const records = filtered
      .map(record => {
        const date = normalizeDate(record.Date || record.date);
        const isoDate = toIsoDate(date);
        const userName = record.UserName || record.User || record.user || '';
        const status = record.Status || record.status || record.state || '';

        if (!userName || !isoDate || !status) {
          return null;
        }

        return {
          userName,
          status,
          date: isoDate,
          notes: record.Notes || record.notes || ''
        };
      })
      .filter(Boolean);

    return {
      success: true,
      records
    };
  } catch (error) {
    console.error('Error retrieving attendance data range:', error);
    safeWriteError('clientGetAttendanceDataRange', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function clientGetScheduleDashboard(managerIdCandidate, campaignIdCandidate, options = {}) {
  try {
    console.log('📊 Building schedule dashboard for manager/campaign:', managerIdCandidate, campaignIdCandidate);

    if (typeof evaluateSchedulePerformance !== 'function') {
      throw new Error('Schedule analytics utilities are not available. Ensure ScheduleUtilities is loaded.');
    }

    const context = clientGetScheduleContext(managerIdCandidate || null, campaignIdCandidate || null);
    if (!context || !context.success) {
      return {
        success: false,
        error: context && context.error ? context.error : 'Unable to resolve schedule context',
        context
      };
    }

    const dateRange = options.dateRange || {};
    const bundle = loadScheduleDataBundle(context.campaignId || context.providedCampaignId || null, {
      managedUserIds: (context.managedUserIds || []).concat(context.managerId ? [context.managerId] : []),
      startDate: options.startDate || dateRange.start,
      endDate: options.endDate || dateRange.end
    });

    const metricsOptions = Object.assign({}, options.metrics || {}, {
      intervalMinutes: options.intervalMinutes || (options.metrics && options.metrics.intervalMinutes) || 30,
      targetServiceLevel: options.targetServiceLevel || (options.metrics && options.metrics.targetServiceLevel) || 0.8,
      baselineASA: options.baselineASA || (options.metrics && options.metrics.baselineASA) || 45,
      openingHour: options.openingHour || (options.metrics && options.metrics.openingHour),
      closingHour: options.closingHour || (options.metrics && options.metrics.closingHour)
    });

    const evaluation = evaluateSchedulePerformance(bundle.scheduleRows, bundle.demandRows, bundle.agentProfiles, metricsOptions);

    const agentSet = new Set(bundle.scheduleRows.map(row => normalizeUserIdValue(row.UserID || row.UserId || row.AgentID || row.AgentId)).filter(Boolean));
    const totalMinutes = bundle.scheduleRows.reduce((sum, row) => {
      const start = safeNormalizeScheduleTimeToMinutes(row.StartTime || row.PeriodStart || row.ScheduleStart || row.ShiftStart);
      const end = safeNormalizeScheduleTimeToMinutes(row.EndTime || row.PeriodEnd || row.ScheduleEnd || row.ShiftEnd);
      if (start === null || end === null) {
        return sum;
      }
      let diff = end - start;
      if (diff < 0) {
        diff += 24 * 60;
      }
      return sum + diff;
    }, 0);

    const totalHours = Number((totalMinutes / 60).toFixed(2));
    const rosterSummary = {
      agentCount: agentSet.size,
      totalHours,
      averageHoursPerAgent: agentSet.size ? Number((totalHours / agentSet.size).toFixed(2)) : 0
    };

    const fteTotals = bundle.ftePlanRows.reduce((acc, row) => {
      acc.planned += scheduleToNumber(row.PlannedFTE || row.Planned || row.FTEPlanned, 0);
      acc.actual += scheduleToNumber(row.ActualFTE || row.Actual || row.FTEActual, 0);
      return acc;
    }, { planned: 0, actual: 0 });

    const recommendations = buildScheduleRecommendations(evaluation, bundle);

    const response = {
      success: true,
      context: {
        managerId: context.managerId,
        campaignId: context.campaignId,
        providedManagerId: context.providedManagerId,
        providedCampaignId: context.providedCampaignId,
        permissions: context.permissions,
        managedUserCount: context.managedUserCount
      },
      generatedAt: evaluation.generatedAt,
      healthScore: evaluation.healthScore,
      summary: evaluation.summary,
      coverage: evaluation.coverage,
      fairness: evaluation.fairness,
      compliance: evaluation.compliance,
      totals: {
        requiredFTE: evaluation.coverage.totalRequiredFTE,
        staffedFTE: evaluation.coverage.totalStaffedFTE,
        plannedFTE: Number(fteTotals.planned.toFixed(2)),
        actualFTE: Number(fteTotals.actual.toFixed(2)),
        varianceFTE: Number((fteTotals.actual - fteTotals.planned).toFixed(2)),
        rosterHours: totalHours,
        agentCount: rosterSummary.agentCount
      },
      roster: rosterSummary,
      backlogIntervals: evaluation.coverage.backlogRiskIntervals,
      recommendations,
      demandSamples: bundle.demandRows.slice(0, 50)
    };

    if (!options.skipPersistence) {
      persistScheduleHealthSnapshot(context, evaluation, bundle, options);
    }

    return response;
  } catch (error) {
    console.error('Error generating schedule dashboard:', error);
    safeWriteError && safeWriteError('clientGetScheduleDashboard', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function applyScenarioAdjustments(bundle, scenario = {}) {
  const volumeMultiplier = scenario.volumeMultiplier || (scenario.volumeDelta ? 1 + scenario.volumeDelta : 1);
  const ahtMultiplier = scenario.ahtMultiplier || (scenario.ahtDelta ? 1 + scenario.ahtDelta : 1);
  const shrinkageDelta = scenario.shrinkageDelta || 0;
  const absenceRate = scenario.absenceRate || 0;
  const additionalOvertimeMinutes = scheduleToNumber(scenario.additionalOvertimeMinutes || scenario.overtimeMinutes, 0);

  const adjustedDemand = (bundle.demandRows || []).map(row => {
    const clone = Object.assign({}, row);
    if (clone.ForecastContacts !== undefined) {
      clone.ForecastContacts = scheduleToNumber(clone.ForecastContacts, 0) * volumeMultiplier;
    }
    if (clone.ForecastAHT !== undefined) {
      clone.ForecastAHT = scheduleToNumber(clone.ForecastAHT, 0) * ahtMultiplier;
    }
    const shrinkage = safeNormalizeSchedulePercentage(clone.Shrinkage, 0.3) + shrinkageDelta;
    clone.Shrinkage = Math.max(0, shrinkage);
    return clone;
  });

  const adjustedSchedules = (bundle.scheduleRows || []).map(row => {
    const clone = Object.assign({}, row);
    const start = safeNormalizeScheduleTimeToMinutes(clone.StartTime || clone.PeriodStart || clone.ScheduleStart || clone.ShiftStart);
    const end = safeNormalizeScheduleTimeToMinutes(clone.EndTime || clone.PeriodEnd || clone.ScheduleEnd || clone.ShiftEnd);
    if (additionalOvertimeMinutes && end !== null) {
      const newEnd = end + additionalOvertimeMinutes;
      clone.EndTime = minutesToTimeString(newEnd);
      clone.PeriodEnd = clone.EndTime;
    }

    if (absenceRate > 0) {
      clone.FTE = scheduleToNumber(clone.FTE || 1, 1) * Math.max(0, 1 - absenceRate);
    }

    return clone;
  });

  return { scheduleRows: adjustedSchedules, demandRows: adjustedDemand };
}

function clientSimulateScheduleScenario(scenario = {}) {
  try {
    console.log('🧪 Simulating schedule scenario:', scenario && scenario.name ? scenario.name : '(ad-hoc scenario)');

    if (typeof evaluateSchedulePerformance !== 'function') {
      throw new Error('Schedule analytics utilities are not available. Ensure ScheduleUtilities is loaded.');
    }

    const context = clientGetScheduleContext(scenario.managerId || scenario.manager || null, scenario.campaignId || scenario.campaign || null);
    if (!context || !context.success) {
      return {
        success: false,
        error: context && context.error ? context.error : 'Unable to resolve schedule context',
        context
      };
    }

    const bundle = loadScheduleDataBundle(context.campaignId || context.providedCampaignId, {
      managedUserIds: scenario.managedUserIds || context.managedUserIds,
      startDate: scenario.startDate,
      endDate: scenario.endDate
    });

    const metricsOptions = Object.assign({ intervalMinutes: scenario.intervalMinutes || 30 }, scenario.metrics || {});

    const baseline = evaluateSchedulePerformance(bundle.scheduleRows, bundle.demandRows, bundle.agentProfiles, metricsOptions);
    const adjusted = applyScenarioAdjustments(bundle, scenario);
    const projection = evaluateSchedulePerformance(adjusted.scheduleRows, adjusted.demandRows, bundle.agentProfiles, metricsOptions);

    return {
      success: true,
      context: {
        managerId: context.managerId,
        campaignId: context.campaignId
      },
      scenario,
      baseline,
      projection,
      delta: {
        serviceLevel: projection.summary.serviceLevel - baseline.summary.serviceLevel,
        healthScore: projection.healthScore - baseline.healthScore,
        compliance: projection.summary.complianceScore - baseline.summary.complianceScore,
        fairness: projection.summary.fairnessIndex - baseline.summary.fairnessIndex
      },
      recommendations: buildScheduleRecommendations(projection, bundle)
    };
  } catch (error) {
    console.error('Error simulating schedule scenario:', error);
    safeWriteError && safeWriteError('clientSimulateScheduleScenario', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function clientGetAgentScheduleSnapshot(agentIdCandidate, startDateCandidate, endDateCandidate, campaignIdCandidate = null, options = {}) {
  try {
    const resolvedAgentId = normalizeUserIdValue(agentIdCandidate) || normalizeUserIdValue(options.agentId);
    const currentUser = typeof getCurrentUser === 'function' ? getCurrentUser() : null;
    const fallbackAgentId = normalizeUserIdValue(currentUser && (currentUser.ID || currentUser.UserID));
    const agentId = resolvedAgentId || fallbackAgentId;

    if (!agentId) {
      return {
        success: false,
        error: 'Agent could not be resolved from parameters or current user context.'
      };
    }

    const context = clientGetScheduleContext(agentId, campaignIdCandidate || options.campaignId || null);
    if (!context || !context.success) {
      return {
        success: false,
        error: context && context.error ? context.error : 'Unable to resolve agent context',
        context
      };
    }

    const startDate = safeNormalizeScheduleDate(startDateCandidate || options.startDate) || new Date();
    const endDate = safeNormalizeScheduleDate(endDateCandidate || options.endDate) || new Date(startDate.getTime() + 14 * 24 * 60 * 60 * 1000);

    const bundle = loadScheduleDataBundle(context.campaignId || context.providedCampaignId, {
      managedUserIds: [agentId],
      startDate,
      endDate
    });

    const agentSchedules = bundle.scheduleRows.filter(row => normalizeUserIdValue(row.UserID || row.UserId || row.AgentID || row.AgentId) === agentId);
    const agentProfile = bundle.agentProfiles.find(profile => normalizeUserIdValue(profile.ID || profile.UserID || profile.UserId) === agentId) || (context.user && normalizeUserIdValue(context.user.ID) === agentId ? context.user : null);

    const now = new Date();
    const upcomingShifts = agentSchedules
      .map(row => {
        const date = safeNormalizeScheduleDate(row.Date || row.ScheduleDate || row.PeriodStart || row.StartDate || row.Day);
        const startMinutes = safeNormalizeScheduleTimeToMinutes(row.StartTime || row.PeriodStart || row.ScheduleStart || row.ShiftStart);
        const startDateTime = date ? combineDateAndMinutes(date, startMinutes || 0) : null;
        return { row, date, startMinutes, startDateTime };
      })
      .filter(item => item.startDateTime && item.startDateTime >= now)
      .sort((a, b) => a.startDateTime - b.startDateTime);

    const historyShifts = agentSchedules
      .map(row => {
        const date = safeNormalizeScheduleDate(row.Date || row.ScheduleDate || row.PeriodStart || row.StartDate || row.Day);
        const startMinutes = safeNormalizeScheduleTimeToMinutes(row.StartTime || row.PeriodStart || row.ScheduleStart || row.ShiftStart);
        const startDateTime = date ? combineDateAndMinutes(date, startMinutes || 0) : null;
        return { row, date, startMinutes, startDateTime };
      })
      .filter(item => item.startDateTime && item.startDateTime < now)
      .sort((a, b) => b.startDateTime - a.startDateTime);

    const totalMinutes = agentSchedules.reduce((sum, row) => {
      const start = safeNormalizeScheduleTimeToMinutes(row.StartTime || row.PeriodStart || row.ScheduleStart || row.ShiftStart);
      const end = safeNormalizeScheduleTimeToMinutes(row.EndTime || row.PeriodEnd || row.ScheduleEnd || row.ShiftEnd);
      if (start === null || end === null) {
        return sum;
      }
      let diff = end - start;
      if (diff < 0) {
        diff += 24 * 60;
      }
      return sum + diff;
    }, 0);

    const weekendShifts = agentSchedules.filter(row => {
      const date = safeNormalizeScheduleDate(row.Date || row.ScheduleDate || row.PeriodStart || row.StartDate || row.Day);
      return date && WEEKEND.includes(date.getDay());
    }).length;

    const nightShifts = agentSchedules.filter(row => {
      const start = safeNormalizeScheduleTimeToMinutes(row.StartTime || row.PeriodStart || row.ScheduleStart || row.ShiftStart);
      const end = safeNormalizeScheduleTimeToMinutes(row.EndTime || row.PeriodEnd || row.ScheduleEnd || row.ShiftEnd);
      return (start !== null && start >= (options.nightThresholdStart || 20 * 60)) || (end !== null && end <= (options.nightThresholdEnd || 6 * 60));
    }).length;

    const averageStartMinutes = agentSchedules.length
      ? agentSchedules.reduce((sum, row) => sum + (safeNormalizeScheduleTimeToMinutes(row.StartTime || row.PeriodStart || row.ScheduleStart || row.ShiftStart) || 0), 0) / agentSchedules.length
      : null;

    let preferenceScore = null;
    let complianceScore = null;
    let fairnessSummary = null;
    if (typeof calculateFairnessMetrics === 'function') {
      const fairness = calculateFairnessMetrics(agentSchedules, agentProfile ? [agentProfile] : [], options.metrics || {});
      fairnessSummary = fairness && fairness.agentSummaries && fairness.agentSummaries.length ? fairness.agentSummaries[0] : null;
      if (fairnessSummary && typeof fairnessSummary.preferenceScore === 'number') {
        preferenceScore = fairnessSummary.preferenceScore;
      }
    }

    if (typeof calculateComplianceMetrics === 'function') {
      const compliance = calculateComplianceMetrics(agentSchedules, agentProfile ? [agentProfile] : [], {
        allowedBreakOverlap: options.allowedBreakOverlap || 3,
        maxHoursPerDay: options.maxHoursPerDay || 12,
        minRestHours: options.minRestHours || 10
      });
      complianceScore = compliance && typeof compliance.complianceScore === 'number' ? compliance.complianceScore : null;
    }

    const nextShift = upcomingShifts.length ? upcomingShifts[0] : null;
    const alerts = [];

    if (nightShifts >= 3) {
      alerts.push('Multiple night shifts scheduled this period. Ensure adequate rest between shifts.');
    }
    if (weekendShifts >= 3) {
      alerts.push('Heavy weekend coverage detected. Consider requesting swaps if needed.');
    }
    if (complianceScore !== null && complianceScore < 85) {
      alerts.push('Compliance score below target. Review breaks, lunches, and rest periods.');
    }
    const formatShiftOutput = (item) => ({
      id: item.row.ID || item.row.Id || item.row.id || '',
      date: item.date ? formatDateForOutput(item.date) : '',
      dayOfWeek: item.date ? item.date.toLocaleDateString('en-US', { weekday: 'short' }) : '',
      startTime: item.startMinutes !== null && item.startMinutes !== undefined ? minutesToTimeString(item.startMinutes) : '',
      endTime: (() => {
        const end = safeNormalizeScheduleTimeToMinutes(item.row.EndTime || item.row.PeriodEnd || item.row.ScheduleEnd || item.row.ShiftEnd);
        return end !== null && end !== undefined ? minutesToTimeString(end) : '';
      })(),
      location: item.row.Location || '',
      skill: item.row.Skill || item.row.Queue || '',
      status: item.row.Status || item.row.State || '',
      notes: item.row.Notes || ''
    });

    const summary = {
      agentId,
      agentName: (agentProfile && (agentProfile.FullName || agentProfile.UserName || agentProfile.Name)) || (context.user && (context.user.FullName || context.user.UserName)) || '',
      agentEmail: (agentProfile && (agentProfile.Email || agentProfile.email)) || (context.user && (context.user.Email || context.user.email)) || '',
      totalShifts: agentSchedules.length,
      totalScheduledHours: Number((totalMinutes / 60).toFixed(2)),
      weekendShifts,
      nightShifts,
      averageStartTime: averageStartMinutes !== null ? minutesToTimeString(averageStartMinutes) : '',
      preferenceScore,
      complianceScore,
      upcomingHolidays: 0,
      nextShift: nextShift ? formatShiftOutput(nextShift) : null
    };

    return {
      success: true,
      agentId,
      campaignId: context.campaignId || context.providedCampaignId || '',
      summary,
      upcomingShifts: upcomingShifts.slice(0, options.limitUpcoming || 5).map(formatShiftOutput),
      recentShifts: historyShifts.slice(0, options.limitHistory || 5).map(formatShiftOutput),
      alerts,
      context: {
        permissions: context.permissions,
        managedUserCount: context.managedUserCount
      }
    };
  } catch (error) {
    console.error('Error generating agent schedule snapshot:', error);
    safeWriteError && safeWriteError('clientGetAgentScheduleSnapshot', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function clientGetAgentSchedule(agentIdCandidate, options = {}) {
  try {
    const resolvedAgentId = normalizeUserIdValue(agentIdCandidate) || normalizeUserIdValue(options.agentId);
    const currentUser = typeof getCurrentUser === 'function' ? getCurrentUser() : null;
    const fallbackAgentId = normalizeUserIdValue(currentUser && (currentUser.ID || currentUser.UserID));
    const agentId = resolvedAgentId || fallbackAgentId;

    if (!agentId) {
      return {
        success: false,
        error: 'Agent could not be resolved from parameters or current user context.'
      };
    }

    const context = clientGetScheduleContext(agentId, options.campaignId || null);
    if (!context || !context.success) {
      return {
        success: false,
        error: context && context.error ? context.error : 'Unable to resolve schedule context.'
      };
    }

    const windowOptions = Object.assign({}, options);
    const { agentSchedules, startDate, endDate } = resolveAgentScheduleWindow(agentId, context, windowOptions);

    const schedules = agentSchedules
      .map(mapScheduleRowToAgentShift)
      .filter(Boolean)
      .sort((a, b) => {
        const aKey = typeof a.startTimestamp === 'number' ? a.startTimestamp : Number.MAX_SAFE_INTEGER;
        const bKey = typeof b.startTimestamp === 'number' ? b.startTimestamp : Number.MAX_SAFE_INTEGER;
        return aKey - bKey;
      });

    return {
      success: true,
      agentId,
      campaignId: context.campaignId || context.providedCampaignId || '',
      schedules,
      summary: {
        total: schedules.length,
        startDate: formatDateForOutput(startDate),
        endDate: formatDateForOutput(endDate)
      }
    };
  } catch (error) {
    console.error('Error fetching agent schedule:', error);
    safeWriteError && safeWriteError('clientGetAgentSchedule', error);
    return {
      success: false,
      error: error.message
    };
  }
}

console.log('✅ Enhanced Schedule Management Backend v4.1 loaded successfully');
console.log('🔧 Features: ScheduleUtilities integration, MainUtilities user management, dedicated spreadsheet support');
console.log('🎯 Ready for production use with comprehensive diagnostics and proper utility integration');
console.log('📊 Integrated: User/Campaign management from MainUtilities, Sheet management from ScheduleUtilities');
