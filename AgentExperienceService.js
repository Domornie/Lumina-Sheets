function clientGetAgentExperienceDashboard(options) {
  options = options || {};
  try {
    const currentUser = typeof getCurrentUser === 'function' ? getCurrentUser() : null;
    if (!currentUser || !currentUser.ID) {
      throw new Error('Unable to resolve the current signed-in user.');
    }

    const agent = buildAgentExperienceContext_(currentUser);
    const daysAhead = Number(options.daysAhead) > 0 ? Number(options.daysAhead) : 14;
    const lookbackDays = Number(options.lookbackDays) > 0 ? Number(options.lookbackDays) : 90;

    const schedule = getAgentScheduleSnapshot_(agent, daysAhead);
    const qa = getAgentQaSnapshot_(agent, lookbackDays);
    const coaching = getAgentCoachingSnapshot_(agent, lookbackDays);
    const attendance = getAgentAttendanceSnapshot_(agent, Math.max(lookbackDays, 30));
    const messaging = getAgentNotifications_(agent, options.notificationLimit || 30);
    const recognition = buildAgentRecognitionHighlights_(agent, { schedule, qa, coaching, attendance, messaging });
    const hero = buildAgentHeroSummary_(agent, { schedule, qa, coaching, attendance, messaging });

    return {
      agent: agent,
      hero: hero,
      schedule: schedule,
      qa: qa,
      coaching: coaching,
      attendance: attendance,
      recognition: recognition,
      messaging: messaging,
      generatedAt: new Date().toISOString()
    };
  } catch (error) {
    console.error('clientGetAgentExperienceDashboard failed:', error);
    throw error;
  }
}

function clientSubmitAgentExperienceMessage(audience, messageBody) {
  try {
    const currentUser = typeof getCurrentUser === 'function' ? getCurrentUser() : null;
    if (!currentUser || !currentUser.ID) {
      throw new Error('Unable to resolve current user context.');
    }

    const trimmed = String(messageBody || '').trim();
    if (!trimmed) {
      throw new Error('Message body is required.');
    }

    const agent = buildAgentExperienceContext_(currentUser);
    const audienceKey = String(audience || 'manager').toLowerCase();
    const category = deriveNotificationCategoryFromAudience_(audienceKey);
    const now = new Date();
    const dataPayload = {
      from: agent.fullName || agent.email || 'Agent',
      audience: audienceKey,
      category: category,
      source: 'agent-experience'
    };

    const sheet = ensureSheetWithHeaders(NOTIFICATIONS_SHEET, NOTIFICATIONS_HEADERS);
    const id = Utilities.getUuid();
    const rowValues = NOTIFICATIONS_HEADERS.map(function(header) {
      switch (header) {
        case 'ID': return id;
        case 'UserId':
        case 'UserID': return agent.id;
        case 'Type': return 'AgentExperienceReply';
        case 'Severity': return 'Info';
        case 'Title': return 'Agent Update';
        case 'Message': return trimmed;
        case 'Data': return JSON.stringify(dataPayload);
        case 'Read': return 'TRUE';
        case 'ActionTaken': return '';
        case 'CreatedAt': return now;
        case 'ReadAt': return now;
        case 'ExpiresAt': return '';
        default: return '';
      }
    });

    sheet.appendRow(rowValues);
    invalidateCache && invalidateCache(NOTIFICATIONS_SHEET);

    return {
      success: true,
      message: {
        id: id,
        category: category,
        title: 'Sent to ' + formatAudienceLabel_(audienceKey),
        body: trimmed,
        from: agent.fullName || 'You',
        createdAt: now.toISOString(),
        read: true
      }
    };
  } catch (error) {
    console.error('clientSubmitAgentExperienceMessage failed:', error);
    throw error;
  }
}

function clientAcknowledgeCoachingRecord(coachingId, acknowledgementText) {
  try {
    if (!coachingId) {
      throw new Error('Coaching record ID is required.');
    }

    const note = String(acknowledgementText || '').trim();
    const safeNote = note ? agentExperienceEscapeHtml_(note) : 'Acknowledged via Agent Experience dashboard';
    const html = '<p>' + safeNote + '</p>';
    const timestamp = acknowledgeCoaching(coachingId, html);

    return {
      success: true,
      acknowledgedOn: timestamp
    };
  } catch (error) {
    console.error('clientAcknowledgeCoachingRecord failed:', error);
    throw error;
  }
}

function buildAgentExperienceContext_(user) {
  const fullName = String(user.FullName || user.UserName || '').trim();
  const email = normalizeEmail_(user.Email || user.email || '');
  return {
    id: String(user.ID || user.Id || '').trim(),
    fullName: fullName,
    firstName: fullName ? fullName.split(/\s+/)[0] : '',
    email: email,
    campaignId: String(user.CampaignID || user.CampaignId || '').trim(),
    campaignName: String(user.campaignName || '').trim(),
    roles: Array.isArray(user.roleNames) ? user.roleNames.slice() : []
  };
}

function getAgentScheduleSnapshot_(agent, daysAhead) {
  if (typeof readScheduleSheet !== 'function') {
    return { items: [], metrics: {} };
  }

  try {
    const rows = readScheduleSheet(SCHEDULE_GENERATION_SHEET) || [];
    const today = new Date();
    const endWindow = new Date(today.getTime() + daysAhead * 24 * 60 * 60 * 1000);
    const matchesId = agent.id ? agent.id : null;
    const nameLower = agent.fullName ? agent.fullName.toLowerCase() : '';
    const emailLower = agent.email ? agent.email.toLowerCase() : '';

    const items = [];
    let totalHours = 0;
    const statusCounts = { onsite: 0, remote: 0, training: 0, flex: 0 };

    rows.forEach(function(row) {
      const dateValue = parseDateValue_(row.Date || row.date);
      if (!dateValue) return;
      if (dateValue < today || dateValue > endWindow) return;

      const userId = String(row.UserID || row.UserId || row.userid || '').trim();
      const userName = String(row.UserName || row.User || '').trim().toLowerCase();
      const userEmail = normalizeEmail_(row.UserEmail || row.Email || '');

      if (matchesId && userId === matchesId) {
        // ok
      } else if (userEmail && emailLower && userEmail === emailLower) {
        // ok
      } else if (nameLower && userName === nameLower) {
        // ok
      } else {
        return;
      }

      const startTime = parseTimeValue_(row.StartTime || row.startTime);
      const endTime = parseTimeValue_(row.EndTime || row.endTime);
      const hours = startTime && endTime ? (endTime - startTime) / (1000 * 60 * 60) : 0;
      if (hours > 0) {
        totalHours += hours;
      }

      const status = String(row.Status || row.status || '').toLowerCase();
      if (status.indexOf('remote') >= 0) statusCounts.remote++;
      else if (status.indexOf('training') >= 0) statusCounts.training++;
      else if (status.indexOf('flex') >= 0 || status.indexOf('rest') >= 0) statusCounts.flex++;
      else statusCounts.onsite++;

      items.push({
        id: String(row.ID || row.Id || Utilities.getUuid()),
        dateIso: dateValue.toISOString(),
        displayDate: formatDisplayDate_(dateValue),
        startDisplay: formatDisplayTime_(startTime),
        endDisplay: formatDisplayTime_(endTime),
        slotName: row.SlotName || row.SlotID || '',
        department: row.Department || '',
        status: status || 'onsite',
        location: row.Location || '',
        notes: row.Notes || ''
      });
    });

    items.sort(function(a, b) {
      return new Date(a.dateIso).getTime() - new Date(b.dateIso).getTime();
    });

    const nextShift = items.length ? items[0] : null;
    const weekLabel = 'Week of ' + Utilities.formatDate(today, 'America/Jamaica', 'MMM d');

    return {
      items: items,
      metrics: {
        upcomingCount: items.length,
        totalHours: totalHours,
        statusCounts: statusCounts,
        nextShift: nextShift,
        weekLabel: weekLabel
      }
    };
  } catch (error) {
    console.error('getAgentScheduleSnapshot_ error:', error);
    return { items: [], metrics: {} };
  }
}

function getAgentQaSnapshot_(agent, lookbackDays) {
  if (typeof getAllQA !== 'function') {
    return { metrics: {}, trend: {}, categoryBreakdown: [], recentAudits: [] };
  }

  try {
    const allRecords = getAllQA() || [];
    const now = new Date();
    const lookbackStart = new Date(now.getTime() - lookbackDays * 24 * 60 * 60 * 1000);
    const emailLower = agent.email || '';
    const nameLower = agent.fullName ? agent.fullName.toLowerCase() : '';

    const relevant = [];

    allRecords.forEach(function(record) {
      const recordEmail = normalizeEmail_(record.AgentEmail || record.agentEmail || '');
      const recordName = String(record.AgentName || record.agentName || '').trim().toLowerCase();
      if (emailLower && recordEmail && recordEmail === emailLower) {
        // include
      } else if (nameLower && recordName === nameLower) {
        // include
      } else {
        return;
      }

      const callDate = parseDateValue_(record.CallDate || record.callDate || record.Timestamp);
      if (callDate && callDate >= lookbackStart && callDate <= now) {
        relevant.push({ record: record, callDate: callDate });
      }
    });

    relevant.sort(function(a, b) { return b.callDate - a.callDate; });

    const percentages = relevant.map(function(item) {
      const raw = Number(item.record.Percentage || item.record.percentage || 0);
      return raw > 1 ? raw : raw * 100;
    });

    const averageScore = percentages.length ? (percentages.reduce(function(sum, value) { return sum + value; }, 0) / percentages.length) : 0;
    const passCount = relevant.filter(function(item, index) {
      const score = percentages[index];
      return score >= 80;
    }).length;

    const focusArea = determineFocusArea_(relevant);
    const delta = computeScoreDelta_(relevant);
    const rankLabel = computeAgentRankLabel_(allRecords, averageScore, agent);

    const trend30 = buildQaTrend_(relevant, 30, 'week');
    const trend90 = buildQaTrend_(relevant, 90, 'month');
    const categoryBreakdown = buildQaCategoryBreakdown_(relevant);

    const recentAudits = relevant.slice(0, 5).map(function(item, index) {
      const record = item.record;
      const score = percentages[index] || 0;
      return {
        id: String(record.ID || record.Id || Utilities.getUuid()),
        callDate: item.callDate.toISOString(),
        score: score,
        auditor: record.AuditorName || record.auditorName || '',
        clientName: record.ClientName || record.clientName || '',
        notes: record.OverallFeedback || record.overallFeedback || record.Notes || record.notes || ''
      };
    });

    const lastAudit = recentAudits.length ? recentAudits[0] : null;

    return {
      metrics: {
        averageScore: averageScore / 100,
        passRate: percentages.length ? passCount / percentages.length : 0,
        evaluationsCount: percentages.length,
        focusArea: focusArea,
        delta: delta,
        rankLabel: rankLabel,
        lastAuditScore: lastAudit ? lastAudit.score / 100 : null,
        lastAuditDate: lastAudit ? lastAudit.callDate : null,
        target: 0.95
      },
      trend: {
        30: trend30,
        90: trend90
      },
      categoryBreakdown: categoryBreakdown,
      recentAudits: recentAudits
    };
  } catch (error) {
    console.error('getAgentQaSnapshot_ error:', error);
    return { metrics: {}, trend: {}, categoryBreakdown: [], recentAudits: [] };
  }
}

function getAgentCoachingSnapshot_(agent, lookbackDays) {
  if (typeof getAllCoaching !== 'function') {
    return { records: [], pendingCount: 0, acknowledgedCount: 0, totalCount: 0 };
  }

  try {
    const allRecords = getAllCoaching() || [];
    const now = new Date();
    const lookbackStart = new Date(now.getTime() - lookbackDays * 24 * 60 * 60 * 1000);
    const emailLower = agent.email || '';
    const nameLower = agent.fullName ? agent.fullName.toLowerCase() : '';

    const mapped = [];

    allRecords.forEach(function(record) {
      const coacheeEmail = normalizeEmail_(record.CoacheeEmail || record.coacheeEmail || '');
      const coacheeName = String(record.CoacheeName || record.coacheeName || '').trim().toLowerCase();
      if (emailLower && coacheeEmail && coacheeEmail === emailLower) {
        // ok
      } else if (nameLower && coacheeName === nameLower) {
        // ok
      } else {
        return;
      }

      const sessionDate = parseDateValue_(record.SessionDate || record.sessionDate);
      if (sessionDate && sessionDate < lookbackStart) {
        return;
      }

      const followUpDate = parseDateValue_(record.FollowUpDate || record.followUpDate);
      const acknowledgedOn = parseDateValue_(record.AcknowledgedOn || record.acknowledgedOn);
      const acknowledged = !!acknowledgedOn || Boolean(record.AcknowledgementText);

      mapped.push({
        id: String(record.ID || record.Id || Utilities.getUuid()),
        sessionDate: sessionDate ? sessionDate.toISOString() : null,
        followUpDate: followUpDate ? followUpDate.toISOString() : null,
        coach: record.AgentName || record.agentName || '',
        topics: parseTopics_(record.TopicsPlanned || record.topicsPlanned),
        summary: record.Summary || record.summary || '',
        actionPlan: record.ActionPlan || record.actionPlan || '',
        acknowledged: acknowledged,
        acknowledgedOn: acknowledgedOn ? acknowledgedOn.toISOString() : null,
        acknowledgementText: record.AcknowledgementText || record.acknowledgementText || ''
      });
    });

    mapped.sort(function(a, b) {
      return new Date(b.sessionDate || 0) - new Date(a.sessionDate || 0);
    });

    const pendingCount = mapped.filter(function(r) { return !r.acknowledged; }).length;
    const acknowledgedCount = mapped.length - pendingCount;

    return {
      records: mapped,
      pendingCount: pendingCount,
      acknowledgedCount: acknowledgedCount,
      totalCount: mapped.length
    };
  } catch (error) {
    console.error('getAgentCoachingSnapshot_ error:', error);
    return { records: [], pendingCount: 0, acknowledgedCount: 0, totalCount: 0 };
  }
}

function getAgentAttendanceSnapshot_(agent, lookbackDays) {
  if (typeof fetchAllAttendanceRows !== 'function') {
    return { metrics: {} };
  }

  try {
    const rows = fetchAllAttendanceRows() || [];
    const now = new Date();
    const start = new Date(now.getTime() - lookbackDays * 24 * 60 * 60 * 1000);
    const emailLower = agent.email || '';
    const nameLower = agent.fullName ? agent.fullName.toLowerCase() : '';

    const dayMap = new Map();
    let productiveSeconds = 0;

    rows.forEach(function(row) {
      const timestamp = row.timestamp instanceof Date ? row.timestamp : new Date(row.timestamp || row.Timestamp || row.Date || row.date);
      if (!(timestamp instanceof Date) || isNaN(timestamp.getTime())) return;
      if (timestamp < start || timestamp > now) return;

      const userId = String(row.userId || row.UserId || row.UserID || '').trim();
      const userName = String(row.user || row.User || '').trim().toLowerCase();
      const userEmail = normalizeEmail_(row.userEmail || row.UserEmail || '');

      if (agent.id && userId === agent.id) {
        // ok
      } else if (emailLower && userEmail && emailLower === userEmail) {
        // ok
      } else if (nameLower && userName === nameLower) {
        // ok
      } else {
        return;
      }

      const dayKey = Utilities.formatDate(timestamp, 'America/Jamaica', 'yyyy-MM-dd');
      if (!dayMap.has(dayKey)) {
        dayMap.set(dayKey, { productive: 0, total: 0 });
      }
      const entry = dayMap.get(dayKey);
      const durationSec = Number(row.durationSec || row.DurationSec || row.DurationMin || 0);
      entry.total += durationSec;
      if (Array.isArray(BILLABLE_STATES) && BILLABLE_STATES.indexOf(row.state || row.State) >= 0) {
        entry.productive += durationSec;
        productiveSeconds += durationSec;
      }
    });

    const dayStats = Array.from(dayMap.entries()).map(function(pair) {
      return { day: pair[0], productive: pair[1].productive, total: pair[1].total };
    }).sort(function(a, b) {
      return new Date(b.day) - new Date(a.day);
    });

    const presenceDays = dayStats.filter(function(d) { return d.productive > 0; }).length;
    const totalTrackedDays = dayStats.length;
    const attendanceRate = totalTrackedDays ? presenceDays / totalTrackedDays : 0;
    const consecutivePresenceDays = computeConsecutivePresenceDays_(dayStats);

    return {
      metrics: {
        attendanceRate: attendanceRate,
        presenceDays: presenceDays,
        totalTrackedDays: totalTrackedDays,
        totalProductiveHours: productiveSeconds / 3600,
        consecutivePresenceDays: consecutivePresenceDays,
        lookbackStart: start.toISOString(),
        lookbackEnd: now.toISOString()
      }
    };
  } catch (error) {
    console.error('getAgentAttendanceSnapshot_ error:', error);
    return { metrics: {} };
  }
}

function getAgentNotifications_(agent, limit) {
  try {
    const notifications = typeof readSheet === 'function' ? (readSheet(NOTIFICATIONS_SHEET) || []) : [];
    const id = agent.id;
    const emailLower = agent.email || '';
    const nameLower = agent.fullName ? agent.fullName.toLowerCase() : '';

    const matched = [];

    notifications.forEach(function(row) {
      const userId = String(row.UserId || row.UserID || row.User || '').trim();
      const userEmail = normalizeEmail_(row.UserEmail || row.Email || '');
      const userName = String(row.UserName || '').trim().toLowerCase();
      const data = parseNotificationData_(row.Data || row.data);

      if (id && userId && userId === id) {
        // ok
      } else if (emailLower && userEmail && userEmail === emailLower) {
        // ok
      } else if (nameLower && userName && userName === nameLower) {
        // ok
      } else if (data && data.agentId && id && String(data.agentId) === id) {
        // ok
      } else {
        return;
      }

      const createdAt = parseDateValue_(row.CreatedAt || row.createdAt || row.Timestamp || row.timestamp);
      const category = (data && data.category) ? data.category : deriveNotificationCategory_(row.Type || row.type || '', row.Severity || row.severity || '');

      matched.push({
        id: String(row.ID || row.Id || Utilities.getUuid()),
        category: category,
        title: row.Title || row.title || formatAudienceLabel_(data && data.audience),
        body: row.Message || row.message || '',
        from: (data && data.from) || row.CreatedBy || row.Sender || 'System',
        createdAt: createdAt ? createdAt.toISOString() : null,
        read: parseBoolean_(row.Read || row.read)
      });
    });

    matched.sort(function(a, b) {
      return new Date(b.createdAt || 0) - new Date(a.createdAt || 0);
    });

    const limited = matched.slice(0, limit);
    const unreadCount = limited.filter(function(m) { return !m.read; }).length;
    const byCategory = limited.reduce(function(acc, item) {
      acc[item.category] = (acc[item.category] || 0) + 1;
      return acc;
    }, {});

    return {
      threads: limited,
      unreadCount: unreadCount,
      byCategory: byCategory
    };
  } catch (error) {
    console.error('getAgentNotifications_ error:', error);
    return { threads: [], unreadCount: 0, byCategory: {} };
  }
}

function buildAgentRecognitionHighlights_(agent, context) {
  const highlights = [];

  if (context.attendance && context.attendance.metrics && context.attendance.metrics.totalTrackedDays) {
    highlights.push({
      title: 'Attendance Excellence',
      icon: 'fa-regular fa-calendar-check',
      badge: formatPercent_(context.attendance.metrics.attendanceRate) + ' attendance',
      leaders: [
        { name: agent.firstName || 'You', detail: context.attendance.metrics.consecutivePresenceDays + ' day streak' },
        { name: 'Productive Hours', detail: formatHours_(context.attendance.metrics.totalProductiveHours) }
      ]
    });
  }

  if (context.qa && context.qa.metrics && context.qa.metrics.averageScore) {
    highlights.push({
      title: 'QA Standout',
      icon: 'fa-solid fa-shield-heart',
      badge: formatPercent_(context.qa.metrics.averageScore) + ' composite',
      leaders: [
        { name: 'Latest Audit', detail: context.qa.metrics.lastAuditScore ? formatPercent_(context.qa.metrics.lastAuditScore) + ' recent' : 'Awaiting recent audit' },
        { name: 'Focus Area', detail: context.qa.metrics.focusArea || 'Balanced performance' }
      ]
    });
  }

  if (context.coaching && context.coaching.totalCount) {
    highlights.push({
      title: 'Coaching Momentum',
      icon: 'fa-solid fa-user-graduate',
      badge: (context.coaching.totalCount - context.coaching.pendingCount) + ' of ' + context.coaching.totalCount + ' completed',
      leaders: [
        { name: 'Pending', detail: context.coaching.pendingCount + ' acknowledgement' + (context.coaching.pendingCount === 1 ? '' : 's') },
        { name: 'Last Session', detail: context.coaching.records.length ? formatDisplayDate_(new Date(context.coaching.records[0].sessionDate)) : 'No recent session' }
      ]
    });
  }

  if (context.schedule && context.schedule.metrics && context.schedule.metrics.upcomingCount) {
    highlights.push({
      title: 'Schedule Preview',
      icon: 'fa-solid fa-business-time',
      badge: context.schedule.metrics.upcomingCount + ' upcoming shifts',
      leaders: [
        { name: 'Next Shift', detail: context.schedule.metrics.nextShift ? context.schedule.metrics.nextShift.displayDate + ' • ' + context.schedule.metrics.nextShift.startDisplay : 'Awaiting assignment' },
        { name: 'This Week', detail: context.schedule.metrics.weekLabel || 'Current week' }
      ]
    });
  }

  return highlights;
}

function buildAgentHeroSummary_(agent, context) {
  const scheduleMetrics = context.schedule && context.schedule.metrics ? context.schedule.metrics : {};
  const qaMetrics = context.qa && context.qa.metrics ? context.qa.metrics : {};
  const attendanceMetrics = context.attendance && context.attendance.metrics ? context.attendance.metrics : {};
  const coachingMetrics = context.coaching || {};
  const messaging = context.messaging || {};

  const tags = [];
  if (scheduleMetrics.weekLabel) {
    tags.push({ icon: 'fa-regular fa-calendar', label: scheduleMetrics.weekLabel });
  }
  if (qaMetrics.averageScore) {
    tags.push({ icon: 'fa-solid fa-chart-line', label: formatPercent_(qaMetrics.averageScore) + ' QA' });
  }
  if (typeof messaging.unreadCount === 'number') {
    tags.push({ icon: 'fa-regular fa-comments', label: messaging.unreadCount + ' unread message' + (messaging.unreadCount === 1 ? '' : 's') });
  }

  return {
    subtitle: 'Stay on top of schedules, quality targets, and coaching actions in one view.',
    tags: tags,
    stats: {
      qa: {
        value: qaMetrics.averageScore || null,
        target: qaMetrics.target || null,
        delta: qaMetrics.delta || null
      },
      attendance: {
        value: attendanceMetrics.attendanceRate || null,
        periodLabel: attendanceMetrics.totalTrackedDays ? 'Last ' + attendanceMetrics.totalTrackedDays + ' days tracked' : 'Attendance window'
      },
      coaching: {
        pending: coachingMetrics.pendingCount || 0,
        completed: (coachingMetrics.totalCount || 0) - (coachingMetrics.pendingCount || 0)
      },
      shifts: {
        upcoming: scheduleMetrics.upcomingCount || 0,
        nextShift: scheduleMetrics.nextShift ? scheduleMetrics.nextShift.displayDate + ' • ' + scheduleMetrics.nextShift.startDisplay : 'Pending'
      }
    }
  };
}

function buildQaTrend_(records, days, grouping) {
  const now = new Date();
  const start = new Date(now.getTime() - days * 24 * 60 * 60 * 1000);
  const buckets = {};

  records.forEach(function(item, index) {
    const date = item.callDate;
    if (!date || date < start) return;
    const score = Number(item.record.Percentage || item.record.percentage || 0);
    const percent = score > 1 ? score : score * 100;

    let key;
    if (grouping === 'week') {
      key = Utilities.formatDate(date, 'America/Jamaica', 'YYYY-ww');
    } else {
      key = Utilities.formatDate(date, 'America/Jamaica', 'yyyy-MM');
    }
    if (!buckets[key]) {
      buckets[key] = [];
    }
    buckets[key].push(percent);
  });

  return Object.keys(buckets).sort().map(function(key) {
    const values = buckets[key];
    const avg = values.reduce(function(sum, value) { return sum + value; }, 0) / values.length;
    const label = grouping === 'week' ? formatWeekLabel_(key) : formatMonthLabel_(key);
    return { label: label, score: Math.round(avg * 10) / 10 };
  });
}

function buildQaCategoryBreakdown_(records) {
  if (typeof qaCategories_ !== 'function' || typeof qaWeights_ !== 'function') {
    return [];
  }

  const categories = qaCategories_();
  const weights = qaWeights_();
  const totals = {};
  const counts = {};

  records.forEach(function(item) {
    const record = item.record;
    Object.keys(categories).forEach(function(category) {
      const questions = categories[category];
      let earned = 0;
      let applicable = 0;
      questions.forEach(function(question) {
        const key = question.toUpperCase();
        const value = String(record[key] || record[key.toLowerCase()] || '').toLowerCase();
        const weight = weights[key.toLowerCase()] || weights[key] || 1;
        if (!value || value === 'n/a' || value === 'na') return;
        applicable += weight;
        if (value === 'yes') earned += weight;
      });
      if (!counts[category]) {
        counts[category] = 0;
        totals[category] = 0;
      }
      if (applicable > 0) {
        totals[category] += earned / applicable;
        counts[category]++;
      }
    });
  });

  return Object.keys(totals).map(function(category) {
    const avg = counts[category] ? totals[category] / counts[category] : 0;
    return { label: category, value: avg };
  }).sort(function(a, b) { return b.value - a.value; });
}

function determineFocusArea_(records) {
  const breakdown = buildQaCategoryBreakdown_(records);
  if (!breakdown.length) return '';
  const sorted = breakdown.slice().sort(function(a, b) { return a.value - b.value; });
  return sorted[0].label;
}

function computeScoreDelta_(records) {
  if (records.length < 4) return null;
  const recent = records.slice(0, 2).map(function(item) {
    const raw = Number(item.record.Percentage || item.record.percentage || 0);
    return raw > 1 ? raw : raw * 100;
  });
  const previous = records.slice(2, 4).map(function(item) {
    const raw = Number(item.record.Percentage || item.record.percentage || 0);
    return raw > 1 ? raw : raw * 100;
  });
  if (!previous.length) return null;
  const recentAvg = recent.reduce(function(sum, value) { return sum + value; }, 0) / recent.length;
  const previousAvg = previous.reduce(function(sum, value) { return sum + value; }, 0) / previous.length;
  return (recentAvg - previousAvg) / 100;
}

function computeAgentRankLabel_(allRecords, agentAverageScore, agent) {
  if (!allRecords || !allRecords.length || !agentAverageScore) return '';
  const agentScores = {};

  allRecords.forEach(function(record) {
    const email = normalizeEmail_(record.AgentEmail || record.agentEmail || '');
    const name = String(record.AgentName || record.agentName || '').trim().toLowerCase();
    const key = email || name;
    if (!key) return;
    if (!agentScores[key]) {
      agentScores[key] = { total: 0, count: 0 };
    }
    const raw = Number(record.Percentage || record.percentage || 0);
    const percent = raw > 1 ? raw : raw * 100;
    if (!isNaN(percent) && percent > 0) {
      agentScores[key].total += percent;
      agentScores[key].count++;
    }
  });

  const averages = Object.keys(agentScores).map(function(key) {
    const entry = agentScores[key];
    return entry.count ? entry.total / entry.count : 0;
  }).filter(function(value) { return value > 0; }).sort(function(a, b) { return b - a; });

  if (!averages.length) return '';
  const agentKey = (agent.email || '').toLowerCase() || (agent.fullName || '').toLowerCase();
  const agentPercent = agentAverageScore > 1 ? agentAverageScore : agentAverageScore * 100;
  let position = averages.length;
  for (let i = 0; i < averages.length; i++) {
    if (agentPercent >= averages[i] - 0.1) {
      position = i + 1;
      break;
    }
  }
  const percentile = Math.round((position / averages.length) * 100);
  if (percentile <= 0) return 'Top performer';
  if (percentile <= 10) return 'Top 10%';
  if (percentile <= 25) return 'Top 25%';
  if (percentile <= 50) return 'Top 50%';
  return 'On track';
}

function computeConsecutivePresenceDays_(dayStats) {
  if (!Array.isArray(dayStats) || !dayStats.length) return 0;
  let streak = 0;
  let previousDate = null;
  const oneDayMs = 24 * 60 * 60 * 1000;

  dayStats.forEach(function(entry) {
    const date = new Date(entry.day + 'T00:00:00Z');
    if (isNaN(date.getTime())) return;

    if (entry.productive <= 0) {
      previousDate = null;
      return;
    }

    if (!previousDate) {
      streak = 1;
    } else {
      const diffDays = Math.round(Math.abs(previousDate - date) / oneDayMs);
      streak = diffDays === 1 ? streak + 1 : 1;
    }

    previousDate = date;
  });

  return streak;
}

function deriveNotificationCategory_(type, severity) {
  const lower = String(type || '').toLowerCase();
  if (lower.indexOf('coach') >= 0) return 'coaching';
  if (lower.indexOf('recognition') >= 0) return 'recognition';
  if (lower.indexOf('schedule') >= 0) return 'updates';
  if (lower.indexOf('qa') >= 0) return 'coaching';
  const sev = String(severity || '').toLowerCase();
  if (sev === 'positive') return 'recognition';
  return 'updates';
}

function deriveNotificationCategoryFromAudience_(audience) {
  if (audience === 'coach') return 'coaching';
  if (audience === 'qa') return 'updates';
  return 'updates';
}

function parseNotificationData_(data) {
  if (!data) return null;
  try {
    if (typeof data === 'object') return data;
    return JSON.parse(data);
  } catch (error) {
    return null;
  }
}

function parseTopics_(topics) {
  if (!topics) return [];
  if (Array.isArray(topics)) return topics;
  try {
    const parsed = JSON.parse(topics);
    if (Array.isArray(parsed)) return parsed;
  } catch (error) {
    // ignore
  }
  return String(topics).split(/[,\n]/).map(function(item) { return item.trim(); }).filter(Boolean);
}

function parseDateValue_(value) {
  if (!value) return null;
  if (value instanceof Date) {
    if (!isNaN(value.getTime())) return value;
    return null;
  }
  const parsed = new Date(value);
  if (!isNaN(parsed.getTime())) return parsed;
  return null;
}

function parseTimeValue_(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  if (typeof value === 'number') {
    const base = new Date();
    base.setHours(0, 0, 0, 0);
    base.setMilliseconds(value);
    return base;
  }
  if (typeof value === 'string') {
    const match = value.match(/^(\d{1,2}):(\d{2})(?:\s*(AM|PM))?/i);
    if (match) {
      let hours = Number(match[1]);
      const minutes = Number(match[2]);
      const meridian = match[3] ? match[3].toUpperCase() : null;
      if (meridian === 'PM' && hours < 12) hours += 12;
      if (meridian === 'AM' && hours === 12) hours = 0;
      const date = new Date();
      date.setHours(hours, minutes, 0, 0);
      return date;
    }
  }
  return null;
}

function formatDisplayDate_(date) {
  if (!date) return '';
  return Utilities.formatDate(date, 'America/Jamaica', 'EEE, MMM d');
}

function formatDisplayTime_(date) {
  if (!date) return '';
  return Utilities.formatDate(date, 'America/Jamaica', 'hh:mm a');
}

function formatAudienceLabel_(audience) {
  switch (audience) {
    case 'coach': return 'Coach';
    case 'qa': return 'QA Analyst';
    case 'manager':
    default: return 'Manager';
  }
}

function formatPercent_(value) {
  if (value === null || value === undefined || isNaN(value)) return '--';
  const percent = value <= 1 ? value * 100 : value;
  return (Math.round(percent * 10) / 10) + '%';
}

function formatHours_(value) {
  if (!value) return '0 hrs';
  return (Math.round(value * 10) / 10) + ' hrs';
}

function formatWeekLabel_(key) {
  if (!key) return 'Current Week';
  const parts = key.split('-');
  if (parts.length !== 2) return key;
  const year = parts[0];
  const week = parts[1];
  return 'Week ' + parseInt(week, 10) + ' • ' + year;
}

function formatMonthLabel_(key) {
  if (!key) return '';
  const parts = key.split('-');
  if (parts.length !== 2) return key;
  const year = parts[0];
  const month = parts[1];
  const date = new Date(Number(year), Number(month) - 1, 1);
  return Utilities.formatDate(date, 'America/Jamaica', 'MMM yyyy');
}

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function parseBoolean_(value) {
  if (typeof value === 'boolean') return value;
  const str = String(value || '').toLowerCase();
  return str === 'true' || str === '1' || str === 'yes';
}

function agentExperienceEscapeHtml_(value) {
  const str = String(value || '');
  return str.replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
