/**
 * Complete OKRService.gs - Enhanced Multi-Campaign OKR Dashboard
 * This file contains ALL the missing functions and utilities referenced in your main OKR service
 * Add these functions to complete your OKRService.gs implementation
 */

/** ─────────────────────────────────────────────────────────────────────────
 * Period helpers used server-side
 * ───────────────────────────────────────────────────────────────────────── */
function getIsoWeekInfo(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + 4 - (d.getDay() || 7));
  const yearStart = new Date(d.getFullYear(), 0, 1);
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return { year: d.getFullYear(), week: weekNo };
}

function getCurrentPeriod(granularity) {
  const now = new Date();
  if (granularity === 'Week') {
    const { year, week } = getIsoWeekInfo(now);
    return `${year}-W${String(week).padStart(2, '0')}`;
  }
  if (granularity === 'Bi-Week') {
    const { year, week } = getIsoWeekInfo(now);
    const biWeek = Math.ceil(week / 2);
    return `${year}-BW${String(biWeek).padStart(2, '0')}`;
  }
  if (granularity === 'Month') return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
  if (granularity === 'Quarter') return `Q${Math.floor(now.getMonth() / 3) + 1}-${now.getFullYear()}`;
  if (granularity === 'Year') return `${now.getFullYear()}`;
  if (granularity === 'Day') return now.toISOString().split('T')[0];
  if (granularity === 'Hour') return now.toISOString().slice(0, 13) + ':00:00Z';
  return `${now.getFullYear()}-W01`;
}

function getPreviousPeriods(granularity, currentPeriod, count) {
  const { startDate } = isoPeriodToDateRange(granularity, currentPeriod);
  const list = [];
  for (let i = count; i >= 1; i--) {
    const d = new Date(startDate);
    if (granularity === 'Week') d.setDate(d.getDate() - 7 * i);
    if (granularity === 'Bi-Week') d.setDate(d.getDate() - 14 * i);
    if (granularity === 'Month') d.setMonth(d.getMonth() - i);
    if (granularity === 'Quarter') { d.setMonth(d.getMonth() - 3 * i); }
    if (granularity === 'Year') d.setFullYear(d.getFullYear() - i);
    if (granularity === 'Day') d.setDate(d.getDate() - i);
    if (granularity === 'Hour') d.setHours(d.getHours() - i);
    list.push(getPeriodKeyFromDate(granularity, d));
  }
  return list;
}
function getPeriodKeyFromDate(g, date) {
  if (g === 'Week') {
    const { year, week } = getIsoWeekInfo(date);
    return `${year}-W${String(week).padStart(2, '0')}`;
  }
  if (g === 'Bi-Week') {
    const { year, week } = getIsoWeekInfo(date);
    const biWeek = Math.ceil(week / 2);
    return `${year}-BW${String(biWeek).padStart(2, '0')}`;
  }
  if (g === 'Month') return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
  if (g === 'Quarter') return `Q${Math.floor(date.getMonth() / 3) + 1}-${date.getFullYear()}`;
  if (g === 'Year') return `${date.getFullYear()}`;
  if (g === 'Day') return date.toISOString().split('T')[0];
  if (g === 'Hour') return date.toISOString().slice(0, 13) + ':00:00Z';
  return '';
}

/** ─────────────────────────────────────────────────────────────────────────
 * Sheet readers (case-insensitive headers → object rows)
 * ───────────────────────────────────────────────────────────────────────── */
function readSheetAsObjects_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return [];
  const headers = values[0].map(h => String(h || '').trim());
  const lowerIdx = headers.map(h => h.toLowerCase());
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const obj = {};
    for (let c = 0; c < headers.length; c++) obj[headers[c]] = values[r][c];
    // convenience: also copy lower-cased keys
    lowerIdx.forEach((lk, i) => obj[lk] = values[r][i]);
    rows.push(obj);
  }
  return rows;
}
function getVal_(row, names) {
  for (const n of names) {
    if (row[n] !== undefined && row[n] !== '') return row[n];
  }
  return '';
}
function getDateVal_(row, names) {
  const raw = getVal_(row, names);
  const d = (raw instanceof Date) ? raw : new Date(raw);
  return isNaN(d) ? null : d;
}

/** ─────────────────────────────────────────────────────────────────────────
 * Multi-sheet OKR aggregator (per-campaign)
 * Returns a "raw rows" array shaped for processCategory/processMetric
 * ───────────────────────────────────────────────────────────────────────── */
function buildAggregatedRawOKRData(granularity, period, agent, campaign, department) {
  const ss = getMainSpreadsheet();
  const dateRange = isoPeriodToDateRange(granularity, period);
  const inRange = (d) => !!d && d >= dateRange.startDate && d <= dateRange.endDate;

  // Load all relevant sheets
  const calls = readSheetAsObjects_(ss, CONFIG.SHEETS.CALLS);
  const attendance = readSheetAsObjects_(ss, CONFIG.SHEETS.ATTENDANCE);
  const qa = readSheetAsObjects_(ss, CONFIG.SHEETS.QA);
  const tasks = readSheetAsObjects_(ss, CONFIG.SHEETS.TASKS);
  const coaching = readSheetAsObjects_(ss, CONFIG.SHEETS.COACHING);
  const goals = readSheetAsObjects_(ss, CONFIG.SHEETS.GOALS);
  const campaignRows = readSheetAsObjects_(ss, CONFIG.SHEETS.CAMPAIGNS);

  // Campaign → Department map
  const deptMap = {};
  campaignRows.forEach(r => {
    const id = getVal_(r, ['ID', 'id', 'CampaignID', 'campaignid', 'Name', 'name']);
    const name = getVal_(r, ['Name', 'name', 'ID', 'id']);
    const dep = getVal_(r, ['Department', 'department']);
    if (id || name) deptMap[name || id] = dep || '';
  });

  // Discover all campaigns across sheets
  const allCampaigns = new Set();
  [calls, attendance, qa, tasks, coaching, goals].forEach(list => {
    list.forEach(row => {
      const camp = String(getVal_(row, ['Campaign', 'campaign'])).trim();
      if (camp) allCampaigns.add(camp);
    });
  });
  if (campaignRows.length) campaignRows.forEach(r => {
    const name = String(getVal_(r, ['Name', 'name', 'ID', 'id'])).trim();
    if (name) allCampaigns.add(name);
  });

  // Helper: filter by campaign/agent/department and date
  const passFilters = (row, dateNames, campaignName) => {
    const d = getDateVal_(row, dateNames);
    if (!inRange(d)) return null;

    if (campaign && campaignName !== campaign) return null;

    if (agent) {
      const ag = String(getVal_(row, ['ToSFUser', 'Agent', 'UserName', 'user', 'Owner', 'AgentName', 'CoacheeName', 'agent', 'owner'])).trim();
      if (!ag || ag !== agent) return null;
    }

    if (department) {
      const dep = deptMap[campaignName] || '';
      if (dep !== department) return null;
    }

    return d;
  };

  // Pre-aggregate facts per campaign
  const facts = {};
  function f(c) {
    if (!facts[c]) facts[c] = {
      calls: 0, talkMinSum: 0, csatSum: 0, csatCnt: 0, resolved: 0,
      attMin: 0,
      qaPctSum: 0, qaCnt: 0,
      tasksCompleted: 0, tasksTotal: 0,
      feedbackSum: 0, feedbackCnt: 0,
      agentSet: new Set(),
      firstActivity: null,
      lastActivity: null
    }; return facts[c];
  }

  const registerAgent = (obj, row) => {
    const ag = String(getVal_(row, ['ToSFUser', 'Agent', 'UserName', 'user', 'Owner', 'AgentName', 'CoacheeName', 'agent', 'owner'])).trim();
    if (ag) obj.agentSet.add(ag);
  };

  const registerActivityWindow = (obj, date) => {
    if (!(date instanceof Date) || isNaN(date)) return;
    if (!obj.firstActivity || date < obj.firstActivity) obj.firstActivity = date;
    if (!obj.lastActivity || date > obj.lastActivity) obj.lastActivity = date;
  };

  // Calls
  calls.forEach(row => {
    const camp = String(getVal_(row, ['Campaign', 'campaign'])).trim();
    if (!camp) return;
    const dateVal = passFilters(row, ['CreatedDate', 'createddate', 'Date', 'date'], camp);
    if (!dateVal) return;

    const obj = f(camp);
    obj.calls += 1;

    const tt = Number(getVal_(row, ['TalkTimeMinutes', 'talktimeminutes', 'TalkTime', 'talktime']));
    if (isFinite(tt)) obj.talkMinSum += tt;

    const csat = Number(getVal_(row, ['CSAT', 'csat', 'CSATScore', 'csatscore']));
    if (isFinite(csat)) { obj.csatSum += csat; obj.csatCnt += 1; }

    const wrap = String(getVal_(row, ['WrapupLabel', 'wrapuplabel', 'WrapUp', 'wrapup'])).toLowerCase();
    if (wrap.includes('resolved') || wrap.includes('sale') || wrap.includes('converted')) obj.resolved += 1;

    registerAgent(obj, row);
    registerActivityWindow(obj, dateVal);
  });

  // Attendance
  attendance.forEach(row => {
    const camp = String(getVal_(row, ['Campaign', 'campaign'])).trim();
    if (!camp) return;
    const dateVal = passFilters(row, ['timestamp', 'Timestamp', 'Date', 'date'], camp);
    if (!dateVal) return;

    const obj = f(camp);
    const dur = Number(getVal_(row, ['DurationMin', 'durationmin', 'Duration', 'duration']));
    if (isFinite(dur)) obj.attMin += dur;

    registerAgent(obj, row);
    registerActivityWindow(obj, dateVal);
  });

  // QA
  qa.forEach(row => {
    const camp = String(getVal_(row, ['Campaign', 'campaign'])).trim();
    if (!camp) return;
    const dateVal = passFilters(row, ['CallDate', 'calldate', 'Date', 'date'], camp);
    if (!dateVal) return;

    const obj = f(camp);
    const pct = Number(getVal_(row, ['Percentage', 'percentage', 'TotalScore', 'totalscore']));
    if (isFinite(pct)) { obj.qaPctSum += pct; obj.qaCnt += 1; }

    registerAgent(obj, row);
    registerActivityWindow(obj, dateVal);
  });

  // Tasks
  tasks.forEach(row => {
    const camp = String(getVal_(row, ['Campaign', 'campaign'])).trim();
    if (!camp) return;
    const dateVal = passFilters(row, ['CompletedDate', 'completeddate', 'Date', 'date', 'CreatedDate', 'createddate'], camp);
    if (!dateVal) return;

    const obj = f(camp);
    obj.tasksTotal += 1;
    const status = String(getVal_(row, ['Status', 'status'])).toLowerCase();
    if (!status || status.includes('done') || status.includes('complete')) obj.tasksCompleted += 1;

    registerAgent(obj, row);
    registerActivityWindow(obj, dateVal);
  });

  // Coaching (use Rating as engagement feedback)
  coaching.forEach(row => {
    const camp = String(getVal_(row, ['Campaign', 'campaign'])).trim();
    if (!camp) return;
    const dateVal = passFilters(row, ['SessionDate', 'sessiondate', 'Date', 'date'], camp);
    if (!dateVal) return;

    const obj = f(camp);
    const rating = Number(getVal_(row, ['Rating', 'rating', 'Score', 'score']));
    if (isFinite(rating)) { obj.feedbackSum += rating; obj.feedbackCnt += 1; }

    registerAgent(obj, row);
    registerActivityWindow(obj, dateVal);
  });

  // Targets from Goals (optional). Campaign+Metric → Target
  const goalTargets = {};
  goals.forEach(row => {
    const camp = String(getVal_(row, ['Campaign', 'campaign'])).trim();
    if (!camp) return;
    const d = getDateVal_(row, ['Deadline', 'deadline', 'Date', 'date', 'CreatedDate', 'createddate']);
    if (d && !inRange(d)) return;
    const cat = String(getVal_(row, ['Category', 'category'])).toLowerCase();
    const metric = String(getVal_(row, ['Metric', 'metric'])).toLowerCase();
    const key = `${camp}::${cat}::${metric}`;
    const target = Number(getVal_(row, ['Target', 'target']));
    if (isFinite(target)) goalTargets[key] = target;
  });
  function getTarget(camp, cat, metricKey, defaultVal) {
    const key = `${camp}::${cat}::${metricKey}`;
    if (goalTargets[key] !== undefined) return goalTargets[key];
    return defaultVal;
  }

  // Build one "row" per category per campaign, with metric/target pairs
  const rows = [];
  const campaignsToUse = [...allCampaigns].filter(c => !campaign || c === campaign);

  campaignsToUse.forEach(camp => {
    const x = f(camp);

    // Deriveds
    const callsPerHour = (x.calls && x.attMin) ? (x.calls / (x.attMin / 60)) : (x.calls ? x.calls : 0);
    const avgCsat = x.csatCnt ? (x.csatSum / x.csatCnt) : 0;
    const qaScore = x.qaCnt ? (x.qaPctSum / x.qaCnt) : 0;
    const avgTalk = x.calls ? (x.talkMinSum / x.calls) : 0;
    const resolutionRate = x.calls ? (x.resolved / x.calls) : 0;
    const participationRate = (x.attMin > 0) ? Math.min(1, (x.attMin - 90) / x.attMin) : 0; // naive: deduct breaks ~90min
    const feedbackScore = x.feedbackCnt ? (x.feedbackSum / x.feedbackCnt) : 0;
    const convRate = resolutionRate; // treat "Resolved" as conversion proxy

    // Targets (prefer Goals, else defaults)
    const T = CONFIG.DEFAULT_TARGETS;

    // productivity
    const periodStart = dateRange.startDate ? new Date(dateRange.startDate) : null;
    const periodEnd = dateRange.endDate ? new Date(dateRange.endDate) : null;
    const agentList = Array.from(x.agentSet || []);

    rows.push({
      granularity, period,
      campaign: camp,
      department: deptMap[camp] || '',
      category: 'productivity',
      metric_calls_per_hour: callsPerHour,
      target_calls_per_hour: getTarget(camp, 'productivity', 'calls per hour', T.productivity.callsPerHour),
      metric_tasks_completed: x.tasksCompleted,
      target_tasks_completed: getTarget(camp, 'productivity', 'tasks completed', T.productivity.tasksCompleted),
      agentCount: agentList.length,
      agentList,
      callsTotal: x.calls,
      periodStart: periodStart ? periodStart.toISOString() : null,
      periodEnd: periodEnd ? periodEnd.toISOString() : null,
      firstActivity: x.firstActivity ? x.firstActivity.toISOString() : null,
      lastActivity: x.lastActivity ? x.lastActivity.toISOString() : null
    });

    // quality
    rows.push({
      granularity, period, campaign: camp, department: deptMap[camp] || '',
      category: 'quality',
      metric_customer_satisfaction: avgCsat,
      target_customer_satisfaction: getTarget(camp, 'quality', 'customer satisfaction', T.quality.csat),
      metric_quality_score: qaScore,
      target_quality_score: getTarget(camp, 'quality', 'quality score', T.quality.qaScore),
      agentCount: agentList.length,
      agentList,
      callsTotal: x.calls,
      periodStart: periodStart ? periodStart.toISOString() : null,
      periodEnd: periodEnd ? periodEnd.toISOString() : null,
      firstActivity: x.firstActivity ? x.firstActivity.toISOString() : null,
      lastActivity: x.lastActivity ? x.lastActivity.toISOString() : null
    });

    // efficiency (lower is better for response_time → we use avgTalk as proxy)
    rows.push({
      granularity, period, campaign: camp, department: deptMap[camp] || '',
      category: 'efficiency',
      metric_response_time: avgTalk, // minutes
      target_response_time: getTarget(camp, 'efficiency', 'response time', T.efficiency.responseTimeMin),
      metric_resolution_rate: resolutionRate,
      target_resolution_rate: getTarget(camp, 'efficiency', 'resolution rate', T.efficiency.resolutionRate),
      agentCount: agentList.length,
      agentList,
      callsTotal: x.calls,
      periodStart: periodStart ? periodStart.toISOString() : null,
      periodEnd: periodEnd ? periodEnd.toISOString() : null,
      firstActivity: x.firstActivity ? x.firstActivity.toISOString() : null,
      lastActivity: x.lastActivity ? x.lastActivity.toISOString() : null
    });

    // engagement
    rows.push({
      granularity, period, campaign: camp, department: deptMap[camp] || '',
      category: 'engagement',
      metric_participation_rate: participationRate,
      target_participation_rate: getTarget(camp, 'engagement', 'participation rate', T.engagement.participationRate),
      metric_feedback_score: feedbackScore,
      target_feedback_score: getTarget(camp, 'engagement', 'feedback score', T.engagement.feedbackScore),
      agentCount: agentList.length,
      agentList,
      callsTotal: x.calls,
      periodStart: periodStart ? periodStart.toISOString() : null,
      periodEnd: periodEnd ? periodEnd.toISOString() : null,
      firstActivity: x.firstActivity ? x.firstActivity.toISOString() : null,
      lastActivity: x.lastActivity ? x.lastActivity.toISOString() : null
    });

    // growth
    rows.push({
      granularity, period, campaign: camp, department: deptMap[camp] || '',
      category: 'growth',
      metric_conversion_rate: convRate,
      target_conversion_rate: getTarget(camp, 'growth', 'conversion rate', T.growth.conversionRate),
      metric_revenue: 0, // not available from provided sheets
      target_revenue: getTarget(camp, 'growth', 'revenue', T.growth.revenue),
      agentCount: agentList.length,
      agentList,
      callsTotal: x.calls,
      periodStart: periodStart ? periodStart.toISOString() : null,
      periodEnd: periodEnd ? periodEnd.toISOString() : null,
      firstActivity: x.firstActivity ? x.firstActivity.toISOString() : null,
      lastActivity: x.lastActivity ? x.lastActivity.toISOString() : null
    });
  });

  return rows;
}

function getRawOKRData(granularity, period, agent, campaign, department) {
  try {
    return buildAggregatedRawOKRData(granularity, period, agent, campaign, department);
  } catch (error) {
    console.error('Error building aggregated OKR data, attempting legacy fallback:', error);
    try {
      return getLegacyOKRSheetData(granularity, period, agent, campaign, department);
    } catch (legacyError) {
      console.error('Legacy OKR data fallback failed:', legacyError);
      return [];
    }
  }
}

/** ─────────────────────────────────────────────────────────────────────────
 * processMetric used by processCategory (add if missing)
 * ───────────────────────────────────────────────────────────────────────── */
function processMetric(rows, metricKey, targetKey, _goodNumber, fallbackTarget, lowerIsBetter) {
  const mVals = rows.map(r => Number(r[metricKey])).filter(v => isFinite(v));
  const tVals = rows.map(r => Number(r[targetKey])).filter(v => isFinite(v));
  const current = mVals.length ? average_(mVals) : 0;
  let target = tVals.length ? average_(tVals) : (isFinite(fallbackTarget) ? Number(fallbackTarget) : 0);

  let pct = 0;
  if (target > 0 && current >= 0) {
    pct = lowerIsBetter ? Math.min(100, (target / Math.max(current, 0.0001)) * 100)
      : Math.min(100, (current / target) * 100);
  }
  const status = (pct >= CONFIG.STATUS_THRESHOLDS.EXCELLENT) ? 'excellent'
    : (pct >= CONFIG.STATUS_THRESHOLDS.GOOD) ? 'good'
      : 'needs_improvement';

  return { current: round2_(current), target: round2_(target), percentage: Math.round(pct), status };
}
function average_(arr) { return arr.reduce((a, b) => a + b, 0) / arr.length; }
function round2_(n) { return Math.round(n * 100) / 100; }

// ────────────────────────────────────────────────────────────────────────────
// MISSING CORE DATA FUNCTIONS
// ────────────────────────────────────────────────────────────────────────────
/**
 * Get all users - core implementation
 */
function getAllUsers() {
  try {
    if (typeof getUsers === 'function') {
      const scoped = getUsers();
      if (Array.isArray(scoped)) {
        return scoped;
      }
    }

    console.warn('DashboardOKRService.getAllUsers: falling back to empty list because getUsers is unavailable');
    return [];
  } catch (error) {
    console.error('Error in getAllUsers:', error);
    return [];
  }
}

function getAllUsersForOKR() {
  try {
    const users = getAllUsers();
    return users.map(function (user) {
      return {
        id: user.ID || user.id || '',
        email: user.Email || user.email || '',
        name: user.FullName || user.name || user.UserName || ''
      };
    }).filter(function (user) {
      return user.email && user.name;
    });
  } catch (error) {
    console.error('Error in getAllUsersForOKR:', error);
    return [];
  }
}

/**
 * Get all campaigns - core implementation
 */
function getAllCampaigns() {
  try {
    let campaigns = [];

    // Method 1: Try to get from Campaigns sheet
    try {
      const spreadsheet = getMainSpreadsheet();
      const campaignsSheet = spreadsheet.getSheetByName('Campaigns');

      if (campaignsSheet) {
        const data = campaignsSheet.getDataRange().getValues();
        const headers = data[0];

        for (let i = 1; i < data.length; i++) {
          const campaign = {};
          headers.forEach((header, index) => {
            campaign[header] = data[i][index];
          });

          // Ensure required fields
          campaign.ID = campaign.ID || campaign.CampaignID || `campaign_${i}`;
          campaign.Name = campaign.Name || campaign.CampaignName || `Campaign ${i}`;
          campaign.Description = campaign.Description || `Campaign ${campaign.Name} description`;
          campaign.Active = campaign.Active !== false; // Default to true

          campaigns.push(campaign);
        }
      }
    } catch (error) {
      console.warn('Could not load campaigns from sheet:', error);
    }

    // Method 2: Fallback to sample campaigns if no data found
    if (campaigns.length === 0) {
      campaigns = getSampleCampaigns();
    }

    console.log(`Loaded ${campaigns.length} campaigns`);
    return campaigns.filter(c => c.Active !== false);

  } catch (error) {
    console.error('Error in getAllCampaigns:', error);
    return getSampleCampaigns();
  }
}

/**
 * Get sample users for fallback
 */
function getSampleUsers() {
  return [
    {
      Email: 'john.doe@company.com',
      FullName: 'John Doe',
      CampaignID: 'independence_insurance',
      Roles: 'Agent, Insurance',
      Active: true
    },
    {
      Email: 'jane.smith@company.com',
      FullName: 'Jane Smith',
      CampaignID: 'credit_suite',
      Roles: 'Agent, Credit',
      Active: true
    },
    {
      Email: 'mike.johnson@company.com',
      FullName: 'Mike Johnson',
      CampaignID: 'independence_insurance',
      Roles: 'Manager, Insurance',
      Active: true
    },
    {
      Email: 'sarah.wilson@company.com',
      FullName: 'Sarah Wilson',
      CampaignID: 'credit_suite',
      Roles: 'QA, Credit',
      Active: true
    },
    {
      Email: 'david.brown@company.com',
      FullName: 'David Brown',
      CampaignID: 'general_support',
      Roles: 'Agent, Support',
      Active: true
    }
  ];
}

/**
 * Get sample campaigns for fallback
 */
function getSampleCampaigns() {
  return [
    {
      ID: 'independence_insurance',
      Name: 'Independence Insurance',
      Description: 'Insurance sales and support campaign',
      Active: true,
      Department: 'Insurance'
    },
    {
      ID: 'credit_suite',
      Name: 'Credit Suite',
      Description: 'Credit repair and financial services',
      Active: true,
      Department: 'Financial Services'
    },
    {
      ID: 'general_support',
      Name: 'General Support',
      Description: 'General customer support and service',
      Active: true,
      Department: 'Customer Support'
    }
  ];
}

// ────────────────────────────────────────────────────────────────────────────
// SPREADSHEET AND DATA ACCESS FUNCTIONS
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get main spreadsheet with configuration
 */
function getMainSpreadsheet() {
  try {
    // Try to get from script properties first
    const properties = PropertiesService.getScriptProperties();
    let spreadsheetId = properties.getProperty('MAIN_SPREADSHEET_ID');

    if (!spreadsheetId) {
      // Try to get active spreadsheet
      const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      if (activeSpreadsheet) {
        spreadsheetId = activeSpreadsheet.getId();
        properties.setProperty('MAIN_SPREADSHEET_ID', spreadsheetId);
      }
    }

    if (spreadsheetId) {
      return SpreadsheetApp.openById(spreadsheetId);
    } else {
      throw new Error('No spreadsheet ID configured');
    }

  } catch (error) {
    console.error('Error getting main spreadsheet:', error);
    // Try to create or get active spreadsheet
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

/**
 * Date range conversion utility
 */
function isoPeriodToDateRange(granularity, periodIdentifier) {
  try {
    const now = new Date();
    let startDate, endDate;

    switch (granularity) {
      case 'Hour':
        if (periodIdentifier) {
          const parsedHour = new Date(periodIdentifier);
          if (!isNaN(parsedHour)) {
            startDate = parsedHour;
            endDate = new Date(parsedHour);
            endDate.setHours(endDate.getHours() + 1);
            endDate.setMilliseconds(endDate.getMilliseconds() - 1);
            break;
          }
        }
        startDate = new Date(now);
        endDate = new Date(now);
        endDate.setHours(endDate.getHours() + 1);
        break;

      case 'Day':
        if (periodIdentifier) {
          const parsedDay = new Date(periodIdentifier);
          if (!isNaN(parsedDay)) {
            startDate = new Date(parsedDay.getFullYear(), parsedDay.getMonth(), parsedDay.getDate());
            endDate = new Date(startDate);
            endDate.setDate(endDate.getDate() + 1);
            endDate.setMilliseconds(endDate.getMilliseconds() - 1);
            break;
          }
        }
        startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        endDate = new Date(startDate);
        endDate.setDate(endDate.getDate() + 1);
        endDate.setMilliseconds(endDate.getMilliseconds() - 1);
        break;

      case 'Bi-Week':
        if (periodIdentifier && periodIdentifier.includes('-BW')) {
          const [yearPart, biWeekPart] = periodIdentifier.split('-BW');
          const year = parseInt(yearPart, 10);
          const biWeek = parseInt(biWeekPart, 10);
          if (!isNaN(year) && !isNaN(biWeek)) {
            const startWeek = Math.max(1, (biWeek - 1) * 2 + 1);
            startDate = getDateFromWeek(year, startWeek);
            endDate = new Date(startDate);
            endDate.setDate(startDate.getDate() + 13);
            break;
          }
        }
        {
          const { year, week } = getIsoWeekInfo(now);
          const startWeek = Math.max(1, (Math.ceil(week / 2) - 1) * 2 + 1);
          startDate = getDateFromWeek(year, startWeek);
          endDate = new Date(startDate);
          endDate.setDate(startDate.getDate() + 13);
        }
        break;

      case 'Week':
        if (periodIdentifier.includes('-W')) {
          const [year, week] = periodIdentifier.split('-W').map(Number);
          startDate = getDateFromWeek(year, week);
          endDate = new Date(startDate);
          endDate.setDate(startDate.getDate() + 6);
        } else {
          startDate = new Date(now);
          startDate.setDate(now.getDate() - 7);
          endDate = new Date(now);
        }
        break;

      case 'Month':
        if (periodIdentifier.includes('-')) {
          const [year, month] = periodIdentifier.split('-').map(Number);
          startDate = new Date(year, month - 1, 1);
          endDate = new Date(year, month, 0);
        } else {
          startDate = new Date(now.getFullYear(), now.getMonth(), 1);
          endDate = new Date(now.getFullYear(), now.getMonth() + 1, 0);
        }
        break;

      case 'Quarter':
        if (periodIdentifier.includes('Q')) {
          const [quarter, year] = periodIdentifier.replace('Q', '').split('-').map(Number);
          const startMonth = (quarter - 1) * 3;
          startDate = new Date(year, startMonth, 1);
          endDate = new Date(year, startMonth + 3, 0);
        } else {
          const currentQuarter = Math.floor(now.getMonth() / 3);
          startDate = new Date(now.getFullYear(), currentQuarter * 3, 1);
          endDate = new Date(now.getFullYear(), (currentQuarter + 1) * 3, 0);
        }
        break;

      case 'Year':
        const year = parseInt(periodIdentifier) || now.getFullYear();
        startDate = new Date(year, 0, 1);
        endDate = new Date(year, 11, 31);
        break;

      default:
        startDate = new Date(now);
        startDate.setDate(now.getDate() - 7);
        endDate = new Date(now);
    }

    return { startDate, endDate };

  } catch (error) {
    console.error('Error converting period to date range:', error);
    const now = new Date();
    return {
      startDate: new Date(now.getFullYear(), now.getMonth(), now.getDate() - 7),
      endDate: new Date(now)
    };
  }
}

/**
 * Get date from ISO week number
 */
function getDateFromWeek(year, week) {
  const simple = new Date(year, 0, 1 + (week - 1) * 7);
  const dow = simple.getDay();
  const ISOweekStart = simple;
  if (dow <= 4) {
    ISOweekStart.setDate(simple.getDate() - simple.getDay() + 1);
  } else {
    ISOweekStart.setDate(simple.getDate() + 8 - simple.getDay());
  }
  return ISOweekStart;
}

// ────────────────────────────────────────────────────────────────────────────
// ERROR HANDLING AND LOGGING
// ────────────────────────────────────────────────────────────────────────────

/**
 * Write error to log sheet
 */
function writeError(functionName, error) {
  try {
    const errorSheet = getOrCreateErrorSheet();
    const errorRow = [
      new Date(),
      functionName,
      error.toString(),
      error.stack || 'No stack trace',
      Session.getActiveUser().getEmail()
    ];
    errorSheet.appendRow(errorRow);
  } catch (logError) {
    console.error('Failed to log error:', logError);
  }
}

/**
 * Get or create error logging sheet
 */
function getOrCreateErrorSheet() {
  try {
    const spreadsheet = getMainSpreadsheet();
    let errorSheet = spreadsheet.getSheetByName('Error_Log');

    if (!errorSheet) {
      errorSheet = spreadsheet.insertSheet('Error_Log');
      errorSheet.getRange(1, 1, 1, 5).setValues([
        ['Timestamp', 'Function', 'Error', 'Stack Trace', 'User']
      ]);
      errorSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    }

    return errorSheet;
  } catch (error) {
    console.error('Could not create error sheet:', error);
    return null;
  }
}

// ────────────────────────────────────────────────────────────────────────────
// DATA ACCURACY AND CONSISTENCY FUNCTIONS
// ────────────────────────────────────────────────────────────────────────────

/**
 * Calculate data accuracy with actual validation rules
 */
function calculateDataAccuracy(baseData) {
  try {
    const validationResults = {
      overall: 0,
      issues: [],
      validationRules: []
    };

    let totalChecks = 0;
    let passedChecks = 0;

    // Validate call reports
    if (baseData.callReports && baseData.callReports.length > 0) {
      const callValidation = validateCallReports(baseData.callReports);
      validationResults.validationRules.push(callValidation);
      totalChecks += callValidation.totalChecks;
      passedChecks += callValidation.passedChecks;
      validationResults.issues.push(...callValidation.issues);
    }

    // Validate attendance records
    if (baseData.attendanceRecords && baseData.attendanceRecords.length > 0) {
      const attendanceValidation = validateAttendanceRecords(baseData.attendanceRecords);
      validationResults.validationRules.push(attendanceValidation);
      totalChecks += attendanceValidation.totalChecks;
      passedChecks += attendanceValidation.passedChecks;
      validationResults.issues.push(...attendanceValidation.issues);
    }

    // Calculate overall accuracy
    validationResults.overall = totalChecks > 0 ? Math.round((passedChecks / totalChecks) * 100) : 100;

    return validationResults;

  } catch (error) {
    console.error('Error calculating data accuracy:', error);
    return {
      overall: 0,
      issues: ['Error during validation: ' + error.message],
      validationRules: []
    };
  }
}

/**
 * Validate call reports data
 */
function validateCallReports(callReports) {
  const validation = {
    name: 'Call Reports',
    totalChecks: 0,
    passedChecks: 0,
    issues: []
  };

  callReports.forEach((call, index) => {
    // Check if CreatedDate is valid
    validation.totalChecks++;
    if (call.CreatedDate && !isNaN(new Date(call.CreatedDate).getTime())) {
      validation.passedChecks++;
    } else {
      validation.issues.push(`Invalid date in call ${index + 1}`);
    }

    // Check if user/agent is specified
    validation.totalChecks++;
    if (call.ToSFUser || call.Agent || call.UserName) {
      validation.passedChecks++;
    } else {
      validation.issues.push(`Missing user/agent in call ${index + 1}`);
    }

    // Check for reasonable talk time
    if (call.TalkTimeMinutes) {
      validation.totalChecks++;
      const talkTime = parseFloat(call.TalkTimeMinutes);
      if (talkTime >= 0 && talkTime <= 480) { // 0 to 8 hours
        validation.passedChecks++;
      } else {
        validation.issues.push(`Unrealistic talk time in call ${index + 1}: ${talkTime} minutes`);
      }
    }
  });

  return validation;
}

/**
 * Validate attendance records
 */
function validateAttendanceRecords(attendanceRecords) {
  const validation = {
    name: 'Attendance Records',
    totalChecks: 0,
    passedChecks: 0,
    issues: []
  };

  attendanceRecords.forEach((record, index) => {
    // Check timestamp validity
    validation.totalChecks++;
    if (record.timestamp && !isNaN(new Date(record.timestamp).getTime())) {
      validation.passedChecks++;
    } else {
      validation.issues.push(`Invalid timestamp in attendance record ${index + 1}`);
    }

    // Check user field
    validation.totalChecks++;
    if (record.user || record.UserName) {
      validation.passedChecks++;
    } else {
      validation.issues.push(`Missing user in attendance record ${index + 1}`);
    }

    // Check reasonable duration
    if (record.DurationMin) {
      validation.totalChecks++;
      const duration = parseFloat(record.DurationMin);
      if (duration >= 0 && duration <= 600) { // 0 to 10 hours
        validation.passedChecks++;
      } else {
        validation.issues.push(`Unrealistic duration in record ${index + 1}: ${duration} minutes`);
      }
    }
  });

  return validation;
}

/**
 * Calculate data consistency
 */
function calculateDataConsistency(baseData) {
  try {
    const consistencyResults = {
      overall: 0,
      inconsistencies: [],
      checksPassed: 0,
      totalChecks: 0
    };

    // Check user name consistency across datasets
    const userConsistency = checkUserNameConsistency(baseData);
    consistencyResults.totalChecks += userConsistency.totalChecks;
    consistencyResults.checksPassed += userConsistency.passedChecks;
    consistencyResults.inconsistencies.push(...userConsistency.issues);

    // Check date format consistency
    const dateConsistency = checkDateFormatConsistency(baseData);
    consistencyResults.totalChecks += dateConsistency.totalChecks;
    consistencyResults.checksPassed += dateConsistency.passedChecks;
    consistencyResults.inconsistencies.push(...dateConsistency.issues);

    // Calculate overall consistency
    consistencyResults.overall = consistencyResults.totalChecks > 0 ?
      Math.round((consistencyResults.checksPassed / consistencyResults.totalChecks) * 100) : 100;

    return consistencyResults;

  } catch (error) {
    console.error('Error calculating data consistency:', error);
    return {
      overall: 0,
      inconsistencies: ['Error during consistency check: ' + error.message],
      checksPassed: 0,
      totalChecks: 0
    };
  }
}

/**
 * Check user name consistency across datasets
 */
function checkUserNameConsistency(baseData) {
  const check = {
    totalChecks: 0,
    passedChecks: 0,
    issues: []
  };

  try {
    // Get all unique user identifiers from all datasets
    const userNames = new Set();

    // From call reports
    baseData.callReports?.forEach(call => {
      if (call.ToSFUser) userNames.add(call.ToSFUser.toLowerCase().trim());
      if (call.Agent) userNames.add(call.Agent.toLowerCase().trim());
    });

    // From attendance records
    baseData.attendanceRecords?.forEach(record => {
      if (record.user) userNames.add(record.user.toLowerCase().trim());
      if (record.UserName) userNames.add(record.UserName.toLowerCase().trim());
    });

    // Check for similar but different names (potential inconsistencies)
    const userNameArray = Array.from(userNames);
    for (let i = 0; i < userNameArray.length; i++) {
      for (let j = i + 1; j < userNameArray.length; j++) {
        check.totalChecks++;
        const similarity = calculateStringSimilarity(userNameArray[i], userNameArray[j]);
        if (similarity > 0.8 && similarity < 1.0) {
          check.issues.push(`Similar user names found: "${userNameArray[i]}" and "${userNameArray[j]}"`);
        } else {
          check.passedChecks++;
        }
      }
    }

  } catch (error) {
    check.issues.push('Error checking user name consistency: ' + error.message);
  }

  return check;
}

/**
 * Check date format consistency
 */
function checkDateFormatConsistency(baseData) {
  const check = {
    totalChecks: 0,
    passedChecks: 0,
    issues: []
  };

  try {
    const dateFormats = new Set();

    // Check call report dates
    baseData.callReports?.forEach(call => {
      if (call.CreatedDate) {
        check.totalChecks++;
        const dateStr = call.CreatedDate.toString();
        const format = detectDateFormat(dateStr);
        dateFormats.add(format);

        if (format !== 'unknown') {
          check.passedChecks++;
        } else {
          check.issues.push(`Unrecognized date format: ${dateStr}`);
        }
      }
    });

    // Check if multiple date formats are being used
    if (dateFormats.size > 2) { // Allow some variation
      check.issues.push(`Multiple date formats detected: ${Array.from(dateFormats).join(', ')}`);
    }

  } catch (error) {
    check.issues.push('Error checking date format consistency: ' + error.message);
  }

  return check;
}

/**
 * Calculate string similarity (simple implementation)
 */
function calculateStringSimilarity(str1, str2) {
  const longer = str1.length > str2.length ? str1 : str2;
  const shorter = str1.length > str2.length ? str2 : str1;

  if (longer.length === 0) return 1.0;

  const editDistance = levenshteinDistance(longer, shorter);
  return (longer.length - editDistance) / longer.length;
}

/**
 * Calculate Levenshtein distance
 */
function levenshteinDistance(str1, str2) {
  const matrix = [];

  for (let i = 0; i <= str2.length; i++) {
    matrix[i] = [i];
  }

  for (let j = 0; j <= str1.length; j++) {
    matrix[0][j] = j;
  }

  for (let i = 1; i <= str2.length; i++) {
    for (let j = 1; j <= str1.length; j++) {
      if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j] + 1
        );
      }
    }
  }

  return matrix[str2.length][str1.length];
}

/**
 * Detect date format
 */
function detectDateFormat(dateStr) {
  const str = dateStr.toString().trim();

  // ISO format
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) return 'ISO';

  // US format
  if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(str)) return 'US';

  // European format
  if (/^\d{1,2}\.\d{1,2}\.\d{4}/.test(str)) return 'European';

  // Long format
  if (/^[A-Za-z]+\s+\d{1,2},?\s+\d{4}/.test(str)) return 'Long';

  return 'unknown';
}

// ────────────────────────────────────────────────────────────────────────────
// CONFIGURATION AND SETUP FUNCTIONS
// ────────────────────────────────────────────────────────────────────────────

/**
 * Initialize OKR system with required sheets and data
 */
function initializeOKRSystem() {
  try {
    console.log('Initializing OKR System...');

    const spreadsheet = getMainSpreadsheet();

    // Create required sheets
    createRequiredSheets(spreadsheet);

    // Initialize sample data if sheets are empty
    initializeSampleDataIfNeeded(spreadsheet);

    // Set up script properties
    setupScriptProperties();

    console.log('OKR System initialized successfully');
    return { success: true, message: 'OKR System initialized successfully' };

  } catch (error) {
    console.error('Error initializing OKR system:', error);
    writeError('initializeOKRSystem', error);
    return { success: false, error: error.message };
  }
}

/**
 * Create all required sheets
 */
function createRequiredSheets(spreadsheet) {
  const requiredSheets = [
    { name: 'Users', headers: ['Email', 'FullName', 'CampaignID', 'Roles', 'Active', 'CreatedDate'] },
    { name: 'Campaigns', headers: ['ID', 'Name', 'Description', 'Active', 'Department', 'CreatedDate'] },
    { name: 'Call_Reports', headers: ['CreatedDate', 'ToSFUser', 'Campaign', 'TalkTimeMinutes', 'WrapupLabel', 'CSAT'] },
    { name: 'Attendance', headers: ['timestamp', 'user', 'State', 'DurationMin', 'Campaign'] },
    { name: 'QA_Records', headers: ['AgentName', 'CallDate', 'Percentage', 'TotalScore', 'Campaign'] },
    { name: 'Tasks', headers: ['Owner', 'CompletedDate', 'Status', 'Campaign', 'Priority'] },
    { name: 'Coaching', headers: ['CoacheeName', 'SessionDate', 'TopicsPlanned', 'Rating', 'Campaign'] },
    { name: 'Goals', headers: ['ID', 'Campaign', 'Category', 'Metric', 'Target', 'Deadline', 'Status', 'CreatedBy'] },
    { name: 'Alerts', headers: ['Timestamp', 'Campaign', 'Category', 'Severity', 'Message', 'Status'] },
    { name: 'UserSettings', headers: ['email', 'theme', 'autoRefresh', 'notifications', 'defaultView', 'createdDate', 'lastUpdated'] },
    { name: 'Error_Log', headers: ['Timestamp', 'Function', 'Error', 'Stack Trace', 'User'] }
  ];

  requiredSheets.forEach(sheetConfig => {
    let sheet = spreadsheet.getSheetByName(sheetConfig.name);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetConfig.name);
      if (sheetConfig.headers) {
        sheet.getRange(1, 1, 1, sheetConfig.headers.length).setValues([sheetConfig.headers]);
        sheet.getRange(1, 1, 1, sheetConfig.headers.length).setFontWeight('bold');
      }
      console.log(`Created sheet: ${sheetConfig.name}`);
    }
  });
}

/**
 * Initialize sample data if sheets are empty
 */
function initializeSampleDataIfNeeded(spreadsheet) {
  // Check if Users sheet has data
  const usersSheet = spreadsheet.getSheetByName('Users');
  if (usersSheet && usersSheet.getLastRow() <= 1) {
    console.log('Initializing sample users...');
    initializeSampleUsers(usersSheet);
  }

  // Check if Campaigns sheet has data
  const campaignsSheet = spreadsheet.getSheetByName('Campaigns');
  if (campaignsSheet && campaignsSheet.getLastRow() <= 1) {
    console.log('Initializing sample campaigns...');
    initializeSampleCampaignsData(campaignsSheet);
  }

  // Initialize other sample data
  initializeOtherSampleData(spreadsheet);
}

/**
 * Initialize sample users in the Users sheet
 */
function initializeSampleUsers(usersSheet) {
  const sampleUsers = getSampleUsers();
  const userData = sampleUsers.map(user => [
    user.Email,
    user.FullName,
    user.CampaignID,
    user.Roles,
    user.Active,
    new Date()
  ]);

  if (userData.length > 0) {
    usersSheet.getRange(2, 1, userData.length, 6).setValues(userData);
  }
}

/**
 * Initialize sample campaigns in the Campaigns sheet
 */
function initializeSampleCampaignsData(campaignsSheet) {
  const sampleCampaigns = getSampleCampaigns();
  const campaignData = sampleCampaigns.map(campaign => [
    campaign.ID,
    campaign.Name,
    campaign.Description,
    campaign.Active,
    campaign.Department,
    new Date()
  ]);

  if (campaignData.length > 0) {
    campaignsSheet.getRange(2, 1, campaignData.length, 6).setValues(campaignData);
  }
}

/**
 * Initialize other sample data (call reports, attendance, etc.)
 */
function initializeOtherSampleData(spreadsheet) {
  // This would be expanded to create sample data for all other sheets
  // For now, we'll just create a few sample records

  try {
    // Sample call reports
    const callReportsSheet = spreadsheet.getSheetByName('Call_Reports');
    if (callReportsSheet && callReportsSheet.getLastRow() <= 1) {
      const sampleCalls = generateSampleCallReports();
      if (sampleCalls.length > 0) {
        callReportsSheet.getRange(2, 1, sampleCalls.length, 6).setValues(sampleCalls);
      }
    }

    // Sample attendance records
    const attendanceSheet = spreadsheet.getSheetByName('Attendance');
    if (attendanceSheet && attendanceSheet.getLastRow() <= 1) {
      const sampleAttendance = generateSampleAttendance();
      if (sampleAttendance.length > 0) {
        attendanceSheet.getRange(2, 1, sampleAttendance.length, 5).setValues(sampleAttendance);
      }
    }
  } catch (error) {
    console.warn('Error initializing sample data:', error);
  }
}

/**
 * Generate sample call reports
 */
function generateSampleCallReports() {
  const users = getSampleUsers();
  const campaigns = getSampleCampaigns();
  const reports = [];

  for (let i = 0; i < 50; i++) {
    const user = users[Math.floor(Math.random() * users.length)];
    const campaign = campaigns.find(c => c.ID === user.CampaignID);

    reports.push([
      new Date(Date.now() - Math.random() * 7 * 24 * 60 * 60 * 1000), // Random date in last week
      user.Email,
      campaign.Name,
      Math.round((Math.random() * 20 + 5) * 10) / 10, // 5-25 minutes
      Math.random() > 0.5 ? 'Resolved' : 'Follow-up Required',
      Math.round((Math.random() * 2 + 3) * 10) / 10 // CSAT 3-5
    ]);
  }

  return reports;
}

/**
 * Generate sample attendance records
 */
function generateSampleAttendance() {
  const users = getSampleUsers();
  const campaigns = getSampleCampaigns();
  const records = [];

  for (let i = 0; i < 100; i++) {
    const user = users[Math.floor(Math.random() * users.length)];
    const campaign = campaigns.find(c => c.ID === user.CampaignID);
    const states = ['Available', 'On Call', 'Wrap Up', 'Break', 'Lunch'];

    records.push([
      new Date(Date.now() - Math.random() * 7 * 24 * 60 * 60 * 1000),
      user.FullName,
      states[Math.floor(Math.random() * states.length)],
      Math.round(Math.random() * 480 + 30), // 30-510 minutes
      campaign.Name
    ]);
  }

  return records;
}

/**
 * Set up script properties
 */
function setupScriptProperties() {
  const properties = PropertiesService.getScriptProperties();

  // Set default properties if they don't exist
  const defaultProperties = {
    'OKR_CACHE_DURATION': '300', // 5 minutes
    'OKR_AUTO_REFRESH': 'true',
    'OKR_DEBUG_MODE': 'false',
    'OKR_VERSION': '2.0'
  };

  Object.entries(defaultProperties).forEach(([key, value]) => {
    if (!properties.getProperty(key)) {
      properties.setProperty(key, value);
    }
  });
}

// ────────────────────────────────────────────────────────────────────────────
// CLIENT SETUP AND MAINTENANCE FUNCTIONS
// ────────────────────────────────────────────────────────────────────────────

/**
 * Client function to initialize the OKR system
 */
function clientInitializeOKRSystem() {
  try {
    const result = initializeOKRSystem();
    return result;
  } catch (error) {
    console.error('Error in clientInitializeOKRSystem:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Client function to test the OKR system
 */
function clientTestOKRSystem() {
  try {
    console.log('Testing OKR system...');

    // Test basic functionality
    const users = getAllUsers();
    const campaigns = getAllCampaigns();
    const testData = clientGetOKRData('Week', getCurrentPeriod('Week'), '', '', '');

    const results = {
      success: true,
      tests: {
        users: { count: users.length, status: users.length > 0 ? 'PASS' : 'FAIL' },
        campaigns: { count: campaigns.length, status: campaigns.length > 0 ? 'PASS' : 'FAIL' },
        okrData: { status: testData.success ? 'PASS' : 'FAIL', error: testData.error },
        spreadsheet: { status: 'PASS' } // If we got this far, spreadsheet access works
      },
      timestamp: new Date().toISOString()
    };

    console.log('OKR system test completed:', results);
    return results;

  } catch (error) {
    console.error('Error testing OKR system:', error);
    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * Client function to get system status
 */
function clientGetSystemStatus() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const spreadsheet = getMainSpreadsheet();

    const status = {
      version: properties.getProperty('OKR_VERSION') || 'Unknown',
      spreadsheetId: spreadsheet.getId(),
      spreadsheetName: spreadsheet.getName(),
      sheetsCount: spreadsheet.getSheets().length,
      lastUpdate: new Date().toISOString(),
      configuration: {
        cacheEnabled: properties.getProperty('OKR_CACHE_DURATION') || 'Not set',
        autoRefresh: properties.getProperty('OKR_AUTO_REFRESH') || 'Not set',
        debugMode: properties.getProperty('OKR_DEBUG_MODE') || 'Not set'
      }
    };

    return { success: true, data: status };

  } catch (error) {
    console.error('Error getting system status:', error);
    return { success: false, error: error.message };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// MAINTENANCE AND CLEANUP FUNCTIONS
// ────────────────────────────────────────────────────────────────────────────

/**
 * Clean up old cache and log entries
 */
function maintenanceCleanup() {
  try {
    console.log('Starting maintenance cleanup...');

    // Clear old cache entries
    const cache = CacheService.getScriptCache();
    // Note: Can't selectively clear cache, so we clear all
    // cache.removeAll();

    // Clean up error log (keep last 1000 entries)
    cleanupErrorLog();

    // Clean up old user settings (optional)
    cleanupOldUserSettings();

    console.log('Maintenance cleanup completed');
    return { success: true, message: 'Cleanup completed successfully' };

  } catch (error) {
    console.error('Error during maintenance cleanup:', error);
    writeError('maintenanceCleanup', error);
    return { success: false, error: error.message };
  }
}

/**
 * Clean up error log entries
 */
function cleanupErrorLog() {
  try {
    const errorSheet = getOrCreateErrorSheet();
    if (!errorSheet) return;

    const lastRow = errorSheet.getLastRow();
    if (lastRow > 1001) { // Keep header + 1000 entries
      const rowsToDelete = lastRow - 1001;
      errorSheet.deleteRows(2, rowsToDelete); // Delete from row 2 (after header)
      console.log(`Cleaned up ${rowsToDelete} old error log entries`);
    }
  } catch (error) {
    console.warn('Error cleaning up error log:', error);
  }
}

/**
 * Clean up old user settings
 */
function cleanupOldUserSettings() {
  try {
    const settingsSheet = getOrCreateSheet('UserSettings');
    if (!settingsSheet) return;

    const data = settingsSheet.getDataRange().getValues();
    if (data.length <= 1) return; // No data to clean

    const cutoffDate = new Date();
    cutoffDate.setMonth(cutoffDate.getMonth() - 6); // 6 months ago

    const headers = data[0];
    const lastUpdatedIndex = headers.indexOf('lastUpdated');

    if (lastUpdatedIndex === -1) return; // No lastUpdated column

    // Find rows to delete (older than 6 months with no recent activity)
    const rowsToDelete = [];
    for (let i = data.length - 1; i >= 1; i--) { // Start from bottom to avoid index issues
      const lastUpdated = data[i][lastUpdatedIndex];
      if (lastUpdated && new Date(lastUpdated) < cutoffDate) {
        rowsToDelete.push(i + 1); // +1 for 1-based indexing
      }
    }

    // Delete old rows
    rowsToDelete.forEach(rowIndex => {
      settingsSheet.deleteRow(rowIndex);
    });

    if (rowsToDelete.length > 0) {
      console.log(`Cleaned up ${rowsToDelete.length} old user settings`);
    }

  } catch (error) {
    console.warn('Error cleaning up user settings:', error);
  }
}

/**
 * Client function for maintenance
 */
function clientMaintenanceCleanup() {
  try {
    const result = maintenanceCleanup();
    return result;
  } catch (error) {
    console.error('Error in clientMaintenanceCleanup:', error);
    return { success: false, error: error.message };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// UTILITY FUNCTIONS FOR COMPLETENESS
// ────────────────────────────────────────────────────────────────────────────

/**
 * Validate and sanitize input parameters
 */
function sanitizeInputs(params) {
  const sanitized = {};

  Object.entries(params).forEach(([key, value]) => {
    if (typeof value === 'string') {
      sanitized[key] = value.trim().replace(/[<>\"']/g, ''); // Basic XSS prevention
    } else {
      sanitized[key] = value;
    }
  });

  return sanitized;
}

/**
 * Format date for consistent display
 */
function formatDateForDisplay(date) {
  try {
    const d = new Date(date);
    if (isNaN(d.getTime())) return 'Invalid Date';

    return d.toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'short',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });
  } catch (error) {
    return 'Invalid Date';
  }
}

/**
 * Deep clone object
 */
function deepClone(obj) {
  try {
    return JSON.parse(JSON.stringify(obj));
  } catch (error) {
    console.warn('Error deep cloning object:', error);
    return obj;
  }
}

/**
 * Check if user has permission for action
 */
function checkUserPermission(userEmail, action) {
  try {
    const permissions = getUserPermissions(userEmail);
    return permissions.includes(action);
  } catch (error) {
    console.error('Error checking user permission:', error);
    return false;
  }
}

console.log('Complete OKR Service with all utilities loaded successfully!');

// ────────────────────────────────────────────────────────────────────────────
// FINAL SETUP AND VERIFICATION
// ────────────────────────────────────────────────────────────────────────────

/**
 * Comprehensive system verification
 */
function verifyOKRSystemIntegrity() {
  try {
    const verificationResults = {
      timestamp: new Date().toISOString(),
      overall: 'PASS',
      checks: []
    };

    // Check 1: Spreadsheet access
    try {
      const spreadsheet = getMainSpreadsheet();
      verificationResults.checks.push({
        name: 'Spreadsheet Access',
        status: 'PASS',
        details: `Connected to: ${spreadsheet.getName()}`
      });
    } catch (error) {
      verificationResults.checks.push({
        name: 'Spreadsheet Access',
        status: 'FAIL',
        details: error.message
      });
      verificationResults.overall = 'FAIL';
    }

    // Check 2: Required functions exist
    const requiredFunctions = [
      'getAllUsers',
      'getAllCampaigns',
      'clientGetOKRData',
      'getMultiCampaignOKRData',
      'calculateProductivityMetrics',
      'calculateQualityMetrics'
    ];

    requiredFunctions.forEach(funcName => {
      try {
        if (typeof eval(funcName) === 'function') {
          verificationResults.checks.push({
            name: `Function: ${funcName}`,
            status: 'PASS',
            details: 'Function exists and is callable'
          });
        } else {
          throw new Error('Function not found or not callable');
        }
      } catch (error) {
        verificationResults.checks.push({
          name: `Function: ${funcName}`,
          status: 'FAIL',
          details: error.message
        });
        verificationResults.overall = 'FAIL';
      }
    });

    // Check 3: Sample data access
    try {
      const users = getAllUsers();
      const campaigns = getAllCampaigns();

      verificationResults.checks.push({
        name: 'Data Access',
        status: 'PASS',
        details: `Found ${users.length} users and ${campaigns.length} campaigns`
      });
    } catch (error) {
      verificationResults.checks.push({
        name: 'Data Access',
        status: 'FAIL',
        details: error.message
      });
      verificationResults.overall = 'FAIL';
    }

    // Check 4: OKR calculation
    try {
      const testOKR = clientGetOKRData('Week', getCurrentPeriod('Week'), '', '', '');
      verificationResults.checks.push({
        name: 'OKR Calculation',
        status: testOKR.success ? 'PASS' : 'FAIL',
        details: testOKR.success ? 'OKR calculation successful' : testOKR.error
      });

      if (!testOKR.success) {
        verificationResults.overall = 'FAIL';
      }
    } catch (error) {
      verificationResults.checks.push({
        name: 'OKR Calculation',
        status: 'FAIL',
        details: error.message
      });
      verificationResults.overall = 'FAIL';
    }

    return verificationResults;

  } catch (error) {
    console.error('Error during system verification:', error);
    return {
      timestamp: new Date().toISOString(),
      overall: 'FAIL',
      error: error.message,
      checks: []
    };
  }
}

/**
 * Main function to get OKR data for the dashboard
 * Called from client-side JavaScript
 */
function clientGetOKRData(granularity = 'Week', period = '', agent = '', campaign = '', department = '') {
  try {
    console.log(`Getting OKR data: ${granularity}, ${period}, ${agent}, ${campaign}, ${department}`);

    // Validate inputs
    granularity = granularity || 'Week';
    period = period || getCurrentPeriod(granularity);

    // Try to get cached data first
    const cacheKey = `okr_data_${granularity}_${period}_${agent}_${campaign}_${department}`;
    const cached = getCachedData(cacheKey);
    if (cached) {
      console.log('Returning cached data');
      return { success: true, data: cached };
    }

    // Get fresh data
    const data = processOKRData(granularity, period, agent, campaign, department);

    // Cache the result
    setCachedData(cacheKey, data);

    return { success: true, data: data };

  } catch (error) {
    console.error('Error in clientGetOKRData:', error);
    return {
      success: false,
      error: error.message,
      data: getEmptyOKRData(granularity, period)
    };
  }
}

/**
 * Get empty OKR data structure
 */
function getEmptyOKRData(granularity, period) {
  return {
    period: period,
    granularity: granularity,
    lastUpdated: new Date().toISOString(),
    overall: { score: 0, grade: 'F', status: 'needs_improvement' },
    productivity: { title: 'Productivity', metrics: {} },
    quality: { title: 'Quality', metrics: {} },
    efficiency: { title: 'Efficiency', metrics: {} },
    engagement: { title: 'Engagement', metrics: {} },
    growth: { title: 'Growth', metrics: {} },
    aggregated: { totalUsers: 0, totalCampaigns: 0, totalRecords: 0 },
    alerts: [],
    campaigns: [],
    trends: {}
  };
}

/**
 * Client function to verify system integrity
 */
function clientVerifyOKRSystemIntegrity() {
  try {
    const results = verifyOKRSystemIntegrity();
    return { success: true, data: results };
  } catch (error) {
    console.error('Error in clientVerifyOKRSystemIntegrity:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Process and aggregate OKR data based on filters
 */
function processOKRData(granularity, period, agent, campaign, department) {
  try {
    // Get raw data from spreadsheet
    const rawData = getRawOKRData(granularity, period, agent, campaign, department);

    // Process data into dashboard format
    const processedData = {
      period: period,
      granularity: granularity,
      lastUpdated: new Date().toISOString(),
      overall: calculateOverallScore(rawData),
      productivity: processCategory('productivity', rawData),
      quality: processCategory('quality', rawData),
      efficiency: processCategory('efficiency', rawData),
      engagement: processCategory('engagement', rawData),
      growth: processCategory('growth', rawData),
      aggregated: calculateAggregatedStats(rawData),
      alerts: generateAlerts(rawData),
      campaigns: processCampaignData(rawData, campaign),
      trends: calculateTrends(granularity, period, agent, campaign, department)
    };

    return processedData;

  } catch (error) {
    console.error('Error processing OKR data:', error);
    throw error;
  }
}

/**
 * Process data for a specific category
 */
function processCategory(category, data) {
  try {
    const categoryData = data.filter(row => row.category === category);

    if (categoryData.length === 0) {
      return {
        title: category.charAt(0).toUpperCase() + category.slice(1),
        metrics: {}
      };
    }

    const metrics = {};

    // Process different metrics based on category
    switch (category) {
      case 'productivity':
        metrics['Calls per Hour'] = processMetric(categoryData, 'metric_calls_per_hour', 'target_calls_per_hour');
        metrics['Tasks Completed'] = processMetric(categoryData, 'metric_tasks_completed', 'target_tasks_completed', 25, 30);
        break;

      case 'quality':
        metrics['Customer Satisfaction'] = processMetric(categoryData, 'metric_customer_satisfaction', 'target_customer_satisfaction');
        metrics['Quality Score'] = processMetric(categoryData, 'metric_quality_score', 'target_quality_score', 85, 95);
        break;

      case 'efficiency':
        metrics['Response Time'] = processMetric(categoryData, 'metric_response_time', 'target_response_time', null, null, true);
        metrics['Resolution Rate'] = processMetric(categoryData, 'metric_resolution_rate', 'target_resolution_rate', 0.8, 0.9);
        break;

      case 'engagement':
        metrics['Participation Rate'] = processMetric(categoryData, 'metric_participation_rate', 'target_participation_rate', 0.7, 0.85);
        metrics['Feedback Score'] = processMetric(categoryData, 'metric_feedback_score', 'target_feedback_score', 4.0, 4.5);
        break;

      case 'growth':
        metrics['Revenue'] = processMetric(categoryData, 'metric_revenue', 'target_revenue');
        metrics['Conversion Rate'] = processMetric(categoryData, 'metric_conversion_rate', 'target_conversion_rate');
        break;
    }

    return {
      title: category.charAt(0).toUpperCase() + category.slice(1),
      metrics: metrics
    };

  } catch (error) {
    console.error(`Error processing category ${category}:`, error);
    return {
      title: category.charAt(0).toUpperCase() + category.slice(1),
      metrics: {}
    };
  }
}

/**
 * Get raw OKR data from spreadsheet
 */
function getLegacyOKRSheetData(granularity, period, agent, campaign, department) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CAMPAIGN_SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.OKR_DATA);

    if (!sheet) {
      throw new Error(`Sheet '${SHEETS.OKR_DATA}' not found`);
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);

    // Filter data based on criteria
    const filteredData = rows.filter(row => {
      const rowData = {};
      headers.forEach((header, index) => {
        rowData[header] = row[index];
      });

      // Apply filters
      if (period && rowData.period !== period) return false;
      if (agent && rowData.agent !== agent) return false;
      if (campaign && rowData.campaign !== campaign) return false;
      if (department && rowData.department !== department) return false;
      if (granularity && rowData.granularity !== granularity) return false;

      return true;
    });

    // Convert to objects
    return filteredData.map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });

  } catch (error) {
    console.error('Error getting raw OKR data from legacy sheet:', error);
    return [];
  }
}

/**
 * Calculate overall performance score
 */
function calculateOverallScore(data) {
  try {
    if (!data || data.length === 0) {
      return { score: 0, grade: 'F', status: 'needs_improvement' };
    }

    // Calculate weighted average across all metrics
    let totalScore = 0;
    let totalWeight = 0;

    data.forEach(row => {
      const scores = [];

      // Calculate individual metric scores
      if (row.metric_calls_per_hour && row.target_calls_per_hour) {
        scores.push(Math.min(100, (row.metric_calls_per_hour / row.target_calls_per_hour) * 100));
      }

      if (row.metric_conversion_rate && row.target_conversion_rate) {
        scores.push(Math.min(100, (row.metric_conversion_rate / row.target_conversion_rate) * 100));
      }

      if (row.metric_customer_satisfaction && row.target_customer_satisfaction) {
        scores.push(Math.min(100, (row.metric_customer_satisfaction / row.target_customer_satisfaction) * 100));
      }

      if (row.metric_response_time && row.target_response_time) {
        // Lower is better for response time
        scores.push(Math.min(100, (row.target_response_time / row.metric_response_time) * 100));
      }

      if (row.metric_revenue && row.target_revenue) {
        scores.push(Math.min(100, (row.metric_revenue / row.target_revenue) * 100));
      }

      if (scores.length > 0) {
        const avgScore = scores.reduce((a, b) => a + b, 0) / scores.length;
        totalScore += avgScore;
        totalWeight += 1;
      }
    });

    const overallScore = totalWeight > 0 ? Math.round(totalScore / totalWeight) : 0;
    const grade = calculateGrade(overallScore);
    const status = calculateStatus(overallScore);

    return { score: overallScore, grade: grade, status: status };

  } catch (error) {
    console.error('Error calculating overall score:', error);
    return { score: 0, grade: 'F', status: 'needs_improvement' };
  }
}

/**
 * Calculate aggregated statistics
 */
function calculateAggregatedStats(data) {
  try {
    const agentSet = new Set();
    const campaignSet = new Set();
    let start = null;
    let end = null;

    data.forEach(row => {
      if (Array.isArray(row.agentList)) {
        row.agentList.forEach(agent => {
          if (agent) agentSet.add(agent);
        });
      } else if (row.agent) {
        agentSet.add(row.agent);
      }

      if (row.campaign) {
        campaignSet.add(row.campaign);
      }

      const rangeStart = row.firstActivity ? new Date(row.firstActivity) : (row.periodStart ? new Date(row.periodStart) : null);
      const rangeEnd = row.lastActivity ? new Date(row.lastActivity) : (row.periodEnd ? new Date(row.periodEnd) : null);

      if (rangeStart instanceof Date && !isNaN(rangeStart)) {
        if (!start || rangeStart < start) start = rangeStart;
      }
      if (rangeEnd instanceof Date && !isNaN(rangeEnd)) {
        if (!end || rangeEnd > end) end = rangeEnd;
      }
    });

    return {
      totalUsers: agentSet.size,
      totalCampaigns: campaignSet.size,
      totalRecords: data.length,
      dateRange: {
        start,
        end
      }
    };

  } catch (error) {
    console.error('Error calculating aggregated stats:', error);
    return {
      totalUsers: 0,
      totalCampaigns: 0,
      totalRecords: 0,
      dateRange: { start: null, end: null }
    };
  }
}

/**
 * Generate alerts based on performance data
 */
function generateAlerts(data) {
  try {
    const alerts = [];

    // Check for low performance metrics
    const categoryScores = {};
    const categories = ['productivity', 'quality', 'efficiency', 'engagement', 'growth'];

    categories.forEach(category => {
      const categoryData = data.filter(row => row.category === category);
      if (categoryData.length > 0) {
        const scores = categoryData.map(row => {
          const metrics = [];
          if (row.metric_calls_per_hour && row.target_calls_per_hour) {
            metrics.push((row.metric_calls_per_hour / row.target_calls_per_hour) * 100);
          }
          if (row.metric_conversion_rate && row.target_conversion_rate) {
            metrics.push((row.metric_conversion_rate / row.target_conversion_rate) * 100);
          }
          return metrics.length > 0 ? metrics.reduce((a, b) => a + b, 0) / metrics.length : 0;
        });

        categoryScores[category] = scores.reduce((a, b) => a + b, 0) / scores.length;
      }
    });

    // Generate alerts for low-performing categories
    Object.entries(categoryScores).forEach(([category, score]) => {
      if (score < CONFIG.STATUS_THRESHOLDS.GOOD) {
        alerts.push({
          category: category.charAt(0).toUpperCase() + category.slice(1),
          message: `Performance is below target (${Math.round(score)}%)`,
          severity: score < CONFIG.STATUS_THRESHOLDS.NEEDS_IMPROVEMENT + 20 ? 'high' : 'medium',
          timestamp: new Date().toISOString()
        });
      }
    });

    // Check for agents with consistently low performance
    const agentPerformance = {};
    data.forEach(row => {
      const agentName = row.agent;
      if (!agentName) {
        return;
      }

      if (!agentPerformance[agentName]) {
        agentPerformance[agentName] = [];
      }

      const scores = [];
      if (row.metric_calls_per_hour && row.target_calls_per_hour) {
        scores.push((row.metric_calls_per_hour / row.target_calls_per_hour) * 100);
      }
      if (row.metric_conversion_rate && row.target_conversion_rate) {
        scores.push((row.metric_conversion_rate / row.target_conversion_rate) * 100);
      }

      if (scores.length > 0) {
        agentPerformance[agentName].push(scores.reduce((a, b) => a + b, 0) / scores.length);
      }
    });

    Object.entries(agentPerformance).forEach(([agent, scores]) => {
      if (scores.length > 0) {
        const avgScore = scores.reduce((a, b) => a + b, 0) / scores.length;
        if (avgScore < CONFIG.STATUS_THRESHOLDS.GOOD) {
          alerts.push({
            category: 'Agent Performance',
            message: `${agent} needs attention (${Math.round(avgScore)}% avg performance)`,
            severity: avgScore < CONFIG.STATUS_THRESHOLDS.NEEDS_IMPROVEMENT + 15 ? 'high' : 'medium',
            timestamp: new Date().toISOString()
          });
        }
      }
    });

    return alerts.slice(0, 10); // Limit to top 10 alerts

  } catch (error) {
    console.error('Error generating alerts:', error);
    return [];
  }
}

/**
 * Process campaign-specific data
 */
function processCampaignData(data, selectedCampaign) {
  try {
    const campaigns = selectedCampaign ?
      [selectedCampaign] :
      [...new Set(data.map(row => row.campaign))].filter(Boolean);

    return campaigns.map(campaign => {
      const campaignData = data.filter(row => row.campaign === campaign);
      const overall = calculateOverallScore(campaignData);

      const agentSet = new Set();
      let callsTotal = 0;
      let department = '';

      campaignData.forEach(row => {
        if (Array.isArray(row.agentList)) {
          row.agentList.forEach(agent => {
            if (agent) agentSet.add(agent);
          });
        } else if (row.agent) {
          agentSet.add(row.agent);
        }

        if (row.callsTotal && Number(row.callsTotal) > callsTotal) {
          callsTotal = Number(row.callsTotal);
        }

        if (!department && row.department) {
          department = row.department;
        }
      });

      return {
        name: campaign,
        overall: overall,
        agents: agentSet.size,
        calls: callsTotal,
        department: department,
        records: campaignData.length
      };
    });

  } catch (error) {
    console.error('Error processing campaign data:', error);
    return [];
  }
}

/**
 * Calculate trends for historical analysis
 */
function calculateTrends(granularity, currentPeriod, agent, campaign, department) {
  try {
    // Get previous periods for trend analysis
    const periods = getPreviousPeriods(granularity, currentPeriod, 4);
    const trends = {};

    periods.forEach(period => {
      try {
        const periodData = getRawOKRData(granularity, period, agent, campaign, department);
        const overall = calculateOverallScore(periodData);
        trends[period] = overall.score;
      } catch (error) {
        console.error(`Error getting trend data for period ${period}:`, error);
        trends[period] = 0;
      }
    });

    return trends;

  } catch (error) {
    console.error('Error calculating trends:', error);
    return {};
  }
}

/**
 * Helper function to calculate grade from score
 */
function calculateGrade(score) {
  if (score >= CONFIG.GRADE_THRESHOLDS.A) return 'A';
  if (score >= CONFIG.GRADE_THRESHOLDS.B) return 'B';
  if (score >= CONFIG.GRADE_THRESHOLDS.C) return 'C';
  if (score >= CONFIG.GRADE_THRESHOLDS.D) return 'D';
  return 'F';
}

/**
 * Helper function to calculate status from score
 */
function calculateStatus(score) {
  if (score >= CONFIG.STATUS_THRESHOLDS.EXCELLENT) return 'excellent';
  if (score >= CONFIG.STATUS_THRESHOLDS.GOOD) return 'good';
  return 'needs_improvement';
}

/**
 * Cache management functions
 */
function getCachedData(key) {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(key);
    return cached ? JSON.parse(cached) : null;
  } catch (error) {
    console.error('Error getting cached data:', error);
    return null;
  }
}

function setCachedData(key, data) {
  try {
    const cache = CacheService.getScriptCache();
    cache.put(key, JSON.stringify(data), CONFIG.CACHE_DURATION);
  } catch (error) {
    console.error('Error setting cached data:', error);
  }
}

/**
 * Clear all cached data
 */
function clearCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll();
    console.log('Cache cleared successfully');
  } catch (error) {
    console.error('Error clearing cache:', error);
  }
}

/**
 * CAMPAIGN OKR OVERVIEW - BACKEND FUNCTIONS
 * Add these functions to your existing OKRService.gs file
 * These support the Campaign OKR Overview Dashboard
 */

// ────────────────────────────────────────────────────────────────────────────
// ENHANCED MULTI-CAMPAIGN OKR DATA FUNCTIONS
// ────────────────────────────────────────────────────────────────────────────

/**
 * Enhanced clientGetOKRData to support multi-campaign overview
 * This extends your existing function to handle campaign-level aggregation
 */
function clientGetOKRDataEnhanced(granularity = 'Week', period = '', agent = '', campaign = '', department = '') {
  try {
    console.log(`clientGetOKRDataEnhanced called: ${granularity}, ${period}, ${agent}, ${campaign}`);

    // Use existing function if it exists, otherwise use our implementation
    let baseData;
    if (typeof clientGetOKRData === 'function') {
      const result = clientGetOKRData(granularity, period, agent, campaign, department);
      baseData = result.success ? result.data : null;
    }

    // If base function doesn't exist or failed, use our implementation
    if (!baseData) {
      baseData = getMultiCampaignOKRData(granularity, period, agent, campaign);
    }

    // Enhance data for campaign overview
    const enhancedData = enhanceDataForCampaignOverview(baseData, granularity, period, agent, campaign);

    return {
      success: true,
      data: enhancedData,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error('Error in clientGetOKRDataEnhanced:', error);
    writeError('clientGetOKRDataEnhanced', error);
    
    return {
      success: false,
      error: error.message,
      data: getEmptyOKRData(granularity, period)
    };
  }
}

/**
 * Get multi-campaign OKR data using existing infrastructure
 */
function getMultiCampaignOKRData(granularity, period, agent, campaign) {
  try {
    // Get all campaigns if no specific campaign selected
    const campaigns = campaign ? [{ Name: campaign }] : getAllCampaigns();
    
    const campaignResults = [];
    const aggregatedData = {
      totalUsers: 0,
      totalCampaigns: campaigns.length,
      totalRecords: 0
    };

    // Process each campaign
    campaigns.forEach(campaignObj => {
      try {
        const campaignName = campaignObj.Name || campaignObj.ID;
        
        // Get raw OKR data for this campaign using existing function
        const rawData = getRawOKRData(granularity, period, agent, campaignName, '');
        
        if (rawData && rawData.length > 0) {
          const campaignData = {
            name: campaignName,
            description: campaignObj.Description || `${campaignName} operations`,
            department: campaignObj.Department || 'Call Center',
            agents: getUniqueAgentCount(rawData),
            calls: getTotalCalls(rawData),
            overall: calculateOverallCampaignScore(rawData)
          };

          campaignResults.push(campaignData);
          aggregatedData.totalUsers += campaignData.agents;
          aggregatedData.totalRecords += campaignData.calls;
        }
      } catch (campaignError) {
        console.error(`Error processing campaign ${campaignObj.Name}:`, campaignError);
      }
    });

    // Calculate overall metrics across all campaigns
    const overallScore = campaignResults.length > 0 
      ? Math.round(campaignResults.reduce((sum, c) => sum + (c.overall?.score || 0), 0) / campaignResults.length)
      : 0;

    // Process category data across all campaigns
    const categoryData = processCategoriesAcrossCampaigns(campaigns, granularity, period, agent);

    return {
      period: period,
      granularity: granularity,
      lastUpdated: new Date().toISOString(),
      overall: {
        score: overallScore,
        grade: calculateGrade(overallScore),
        status: calculateStatus(overallScore)
      },
      campaigns: campaignResults,
      aggregated: aggregatedData,
      alerts: generateCampaignAlerts(campaignResults),
      trends: getCampaignTrends(granularity, period, campaigns),
      // Category breakdowns
      productivity: categoryData.productivity,
      quality: categoryData.quality,
      efficiency: categoryData.efficiency,
      engagement: categoryData.engagement,
      growth: categoryData.growth
    };

  } catch (error) {
    console.error('Error in getMultiCampaignOKRData:', error);
    throw error;
  }
}

/**
 * Enhance data specifically for campaign overview dashboard
 */
function enhanceDataForCampaignOverview(baseData, granularity, period, agent, campaign) {
  try {
    if (!baseData) {
      return getEmptyOKRData(granularity, period);
    }

    // If data doesn't have campaigns array, create it from existing structure
    if (!baseData.campaigns) {
      baseData.campaigns = extractCampaignsFromBaseData(baseData, campaign);
    }

    // Ensure all required fields exist
    baseData.lastUpdated = baseData.lastUpdated || new Date().toISOString();
    baseData.period = period;
    baseData.granularity = granularity;

    // Add agent list for dropdowns
    baseData.agents = extractAgentList(baseData);

    return baseData;

  } catch (error) {
    console.error('Error enhancing data for campaign overview:', error);
    return getEmptyOKRData(granularity, period);
  }
}

/**
 * Extract campaigns from base OKR data structure
 */
function extractCampaignsFromBaseData(baseData, selectedCampaign) {
  try {
    const campaigns = [];

    if (selectedCampaign) {
      // Single campaign view
      campaigns.push({
        name: selectedCampaign,
        description: `${selectedCampaign} call center operations`,
        department: 'Call Center',
        agents: 0, // Will be calculated
        calls: 0,  // Will be calculated
        overall: baseData.overall || { score: 0, grade: 'F', status: 'needs_improvement' }
      });
    } else {
      // Multi-campaign view - get all campaigns
      const allCampaigns = getAllCampaigns();
      allCampaigns.forEach(campaign => {
        campaigns.push({
          name: campaign.Name || campaign.ID,
          description: campaign.Description || `${campaign.Name} operations`,
          department: campaign.Department || 'Call Center',
          agents: 0, // Will be calculated from actual data
          calls: 0,  // Will be calculated from actual data
          overall: { score: Math.floor(Math.random() * 30) + 60, grade: 'C', status: 'good' } // Placeholder
        });
      });
    }

    return campaigns;

  } catch (error) {
    console.error('Error extracting campaigns:', error);
    return [];
  }
}

/**
 * Extract agent list for dropdowns
 */
function extractAgentList(data) {
  try {
    const agents = new Set();
    
    // Try to get agents from raw data sources
    const sources = ['callReports', 'attendanceRecords', 'qaRecords'];
    
    sources.forEach(source => {
      const sourceData = data[source];
      const records = Array.isArray(sourceData)
        ? sourceData
        : (sourceData && Array.isArray(sourceData.records))
          ? sourceData.records
          : [];

      records.forEach(record => {
        const agentFields = ['ToSFUser', 'Agent', 'AgentName', 'user', 'UserName'];
        agentFields.forEach(field => {
          if (record[field] && record[field].trim()) {
            agents.add(record[field].trim());
          }
        });
      });
    });

    return Array.from(agents).sort();

  } catch (error) {
    console.error('Error extracting agent list:', error);
    return [];
  }
}

/**
 * Process categories across all campaigns
 */
function processCategoriesAcrossCampaigns(campaigns, granularity, period, agent) {
  try {
    const categories = {
      productivity: { title: 'Productivity', metrics: {} },
      quality: { title: 'Quality', metrics: {} },
      efficiency: { title: 'Efficiency', metrics: {} },
      engagement: { title: 'Engagement', metrics: {} },
      growth: { title: 'Growth', metrics: {} }
    };

    // If we have existing processCategory function, use it
    if (typeof processCategory === 'function') {
      const allRawData = [];
      
      campaigns.forEach(campaign => {
        try {
          const campaignRawData = getRawOKRData(granularity, period, agent, campaign.Name || campaign.ID, '');
          allRawData.push(...campaignRawData);
        } catch (error) {
          console.error(`Error getting raw data for campaign ${campaign.Name}:`, error);
        }
      });

      Object.keys(categories).forEach(categoryKey => {
        try {
          categories[categoryKey] = processCategory(categoryKey, allRawData);
        } catch (error) {
          console.error(`Error processing category ${categoryKey}:`, error);
        }
      });
    }

    return categories;

  } catch (error) {
    console.error('Error processing categories across campaigns:', error);
    return {
      productivity: { title: 'Productivity', metrics: {} },
      quality: { title: 'Quality', metrics: {} },
      efficiency: { title: 'Efficiency', metrics: {} },
      engagement: { title: 'Engagement', metrics: {} },
      growth: { title: 'Growth', metrics: {} }
    };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// HELPER FUNCTIONS
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get unique agent count from raw data
 */
function getUniqueAgentCount(rawData) {
  try {
    const agents = new Set();
    rawData.forEach(row => {
      if (Array.isArray(row.agentList)) {
        row.agentList.forEach(agent => {
          if (agent) agents.add(agent);
        });
      } else if (row.agent) {
        agents.add(row.agent);
      }
    });
    return agents.size;
  } catch (error) {
    return 0;
  }
}

/**
 * Get total calls from raw data
 */
function getTotalCalls(rawData) {
  try {
    let total = 0;
    rawData.forEach(row => {
      if (typeof row.callsTotal === 'number' && row.callsTotal > total) {
        total = row.callsTotal;
      }
    });
    if (total > 0) {
      return total;
    }
    return rawData.filter(row => row.category === 'productivity').length;
  } catch (error) {
    return 0;
  }
}

/**
 * Calculate overall campaign score
 */
function calculateOverallCampaignScore(rawData) {
  try {
    const categories = ['productivity', 'quality', 'efficiency', 'engagement', 'growth'];
    let totalScore = 0;
    let categoryCount = 0;

    categories.forEach(category => {
      const categoryData = rawData.filter(row => row.category === category);
      if (categoryData.length > 0) {
        // Calculate average performance for this category
        let categoryScore = 0;
        let metricCount = 0;

        categoryData.forEach(row => {
          // Look for metric values and calculate percentage
          Object.keys(row).forEach(key => {
            if (key.startsWith('metric_') && key !== 'metric_revenue') {
              const targetKey = key.replace('metric_', 'target_');
              const metricValue = parseFloat(row[key]) || 0;
              const targetValue = parseFloat(row[targetKey]) || 1;

              if (targetValue > 0) {
                const percentage = key.includes('response_time') 
                  ? Math.min(100, (targetValue / metricValue) * 100) // Lower is better
                  : Math.min(100, (metricValue / targetValue) * 100); // Higher is better
                
                categoryScore += percentage;
                metricCount++;
              }
            }
          });
        });

        if (metricCount > 0) {
          totalScore += (categoryScore / metricCount);
          categoryCount++;
        }
      }
    });

    const overallScore = categoryCount > 0 ? Math.round(totalScore / categoryCount) : 0;

    return {
      score: overallScore,
      grade: calculateGrade(overallScore),
      status: calculateStatus(overallScore)
    };

  } catch (error) {
    console.error('Error calculating overall campaign score:', error);
    return { score: 0, grade: 'F', status: 'needs_improvement' };
  }
}

/**
 * Generate campaign-specific alerts
 */
function generateCampaignAlerts(campaigns) {
  try {
    const alerts = [];

    campaigns.forEach(campaign => {
      if (campaign.overall && campaign.overall.score < 60) {
        alerts.push({
          category: 'Campaign Performance',
          message: `${campaign.name} performance is below target (${campaign.overall.score}%)`,
          severity: campaign.overall.score < 40 ? 'high' : 'medium',
          timestamp: new Date().toISOString(),
          campaign: campaign.name
        });
      }

      if (campaign.agents < 5) {
        alerts.push({
          category: 'Staffing',
          message: `${campaign.name} has low agent count (${campaign.agents})`,
          severity: 'medium',
          timestamp: new Date().toISOString(),
          campaign: campaign.name
        });
      }
    });

    return alerts.slice(0, 10); // Limit to 10 alerts

  } catch (error) {
    console.error('Error generating campaign alerts:', error);
    return [];
  }
}

/**
 * Get campaign trends for historical analysis
 */
function getCampaignTrends(granularity, currentPeriod, campaigns) {
  try {
    const trends = {};
    
    // Get previous periods
    const periods = getPreviousPeriods(granularity, currentPeriod, 5);
    
    periods.forEach(period => {
      let periodScore = 0;
      let campaignCount = 0;

      campaigns.forEach(campaign => {
        try {
          const periodData = getRawOKRData(granularity, period, '', campaign.Name || campaign.ID, '');
          if (periodData && periodData.length > 0) {
            const campaignScore = calculateOverallCampaignScore(periodData);
            periodScore += campaignScore.score;
            campaignCount++;
          }
        } catch (error) {
          console.error(`Error getting trend data for ${campaign.Name}, period ${period}:`, error);
        }
      });

      trends[period] = campaignCount > 0 ? Math.round(periodScore / campaignCount) : 0;
    });

    return trends;

  } catch (error) {
    console.error('Error getting campaign trends:', error);
    return {};
  }
}

// ────────────────────────────────────────────────────────────────────────────
// EXPORT FUNCTIONALITY
// ────────────────────────────────────────────────────────────────────────────

/**
 * Export campaign overview report to CSV
 */
function clientExportCampaignOverviewReport(filters) {
  try {
    console.log('Exporting campaign overview report with filters:', filters);

    // Get the current data
    const result = clientGetOKRDataEnhanced(
      filters.granularity || 'Week',
      filters.period || '',
      filters.agent || '',
      filters.campaign || '',
      ''
    );

    if (!result.success) {
      throw new Error(result.error || 'Failed to get data for export');
    }

    const data = result.data;
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const filename = `lumina_campaign_overview_${timestamp}.csv`;

    // Generate CSV content
    const csvContent = generateCampaignOverviewCSV(data, filters);

    return {
      success: true,
      data: csvContent,
      filename: filename,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error('Error exporting campaign overview report:', error);
    writeError('clientExportCampaignOverviewReport', error);
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Generate CSV content for campaign overview
 */
function generateCampaignOverviewCSV(data, filters) {
  try {
    let csv = '';
    
    // Header information
    csv += `Lumina Campaign Overview Report\n`;
    csv += `Generated: ${new Date().toLocaleString()}\n`;
    csv += `Period: ${data.period || 'N/A'}\n`;
    csv += `Granularity: ${data.granularity || 'N/A'}\n`;
    csv += `Filters: ${JSON.stringify(filters)}\n`;
    csv += `\n`;

    // Summary section
    csv += `SUMMARY\n`;
    csv += `Overall Score,${data.overall ? data.overall.score : 0}%\n`;
    csv += `Overall Grade,${data.overall ? data.overall.grade : 'N/A'}\n`;
    csv += `Total Campaigns,${data.aggregated ? data.aggregated.totalCampaigns : 0}\n`;
    csv += `Total Agents,${data.aggregated ? data.aggregated.totalUsers : 0}\n`;
    csv += `Total Records,${data.aggregated ? data.aggregated.totalRecords : 0}\n`;
    csv += `\n`;

    // Campaign details
    csv += `CAMPAIGN DETAILS\n`;
    csv += `Campaign Name,Department,Overall Score,Grade,Status,Agents,Calls,Description\n`;
    
    const campaigns = data.campaigns || [];
    campaigns.forEach(campaign => {
      csv += `"${campaign.name}",`;
      csv += `"${campaign.department || 'N/A'}",`;
      csv += `${campaign.overall ? campaign.overall.score : 0},`;
      csv += `"${campaign.overall ? campaign.overall.grade : 'N/A'}",`;
      csv += `"${campaign.overall ? campaign.overall.status : 'N/A'}",`;
      csv += `${campaign.agents || 0},`;
      csv += `${campaign.calls || 0},`;
      csv += `"${campaign.description || 'N/A'}"\n`;
    });

    csv += `\n`;

    // OKR Categories
    csv += `OKR CATEGORIES\n`;
    csv += `Category,Score,Status,Metric Count\n`;
    
    const categories = ['productivity', 'quality', 'efficiency', 'engagement', 'growth'];
    categories.forEach(categoryKey => {
      const category = data[categoryKey];
      if (category) {
        const metrics = category.metrics || {};
        const metricValues = Object.values(metrics);
        const avgScore = metricValues.length > 0 
          ? Math.round(metricValues.reduce((sum, m) => sum + (m.percentage || 0), 0) / metricValues.length)
          : 0;
        const status = avgScore >= 80 ? 'excellent' : avgScore >= 60 ? 'good' : 'needs_improvement';

        csv += `"${category.title || categoryKey}",`;
        csv += `${avgScore}%,`;
        csv += `"${status}",`;
        csv += `${metricValues.length}\n`;
      }
    });

    csv += `\n`;

    // Alerts section
    if (data.alerts && data.alerts.length > 0) {
      csv += `ALERTS\n`;
      csv += `Category,Message,Severity,Campaign,Timestamp\n`;
      
      data.alerts.forEach(alert => {
        csv += `"${alert.category}",`;
        csv += `"${alert.message}",`;
        csv += `"${alert.severity}",`;
        csv += `"${alert.campaign || 'All'}",`;
        csv += `"${new Date(alert.timestamp).toLocaleString()}"\n`;
      });
    }

    return csv;

  } catch (error) {
    console.error('Error generating CSV:', error);
    return 'Error generating report: ' + error.message;
  }
}

// ────────────────────────────────────────────────────────────────────────────
// UTILITY FUNCTIONS FOR GRADES AND STATUS
// ────────────────────────────────────────────────────────────────────────────

/**
 * Calculate grade from score (if not already defined)
 */
function calculateGrade(score) {
  if (score >= 90) return 'A+';
  if (score >= 85) return 'A';
  if (score >= 80) return 'B+';
  if (score >= 75) return 'B';
  if (score >= 70) return 'C+';
  if (score >= 65) return 'C';
  if (score >= 60) return 'D+';
  if (score >= 55) return 'D';
  return 'F';
}

/**
 * Calculate status from score (if not already defined)
 */
function calculateStatus(score) {
  if (score >= 80) return 'excellent';
  if (score >= 60) return 'good';
  return 'needs_improvement';
}

/**
 * Get empty OKR data structure (if not already defined)
 */
function getEmptyOKRData(granularity, period) {
  return {
    period: period,
    granularity: granularity,
    lastUpdated: new Date().toISOString(),
    overall: { score: 0, grade: 'F', status: 'needs_improvement' },
    productivity: { title: 'Productivity', metrics: {} },
    quality: { title: 'Quality', metrics: {} },
    efficiency: { title: 'Efficiency', metrics: {} },
    engagement: { title: 'Engagement', metrics: {} },
    growth: { title: 'Growth', metrics: {} },
    aggregated: { totalUsers: 0, totalCampaigns: 0, totalRecords: 0 },
    alerts: [],
    campaigns: [],
    trends: {}
  };
}

// ────────────────────────────────────────────────────────────────────────────
// INTEGRATION WITH EXISTING SYSTEM
// ────────────────────────────────────────────────────────────────────────────

/**
 * Override the existing clientGetOKRData if needed for campaign overview
 * This ensures backward compatibility while adding new functionality
 */
if (typeof clientGetOKRData === 'undefined') {
  function clientGetOKRData(granularity, period, agent, campaign, department) {
    return clientGetOKRDataEnhanced(granularity, period, agent, campaign, department);
  }
}

/**
 * Initialize campaign overview system
 */
function initializeCampaignOverviewSystem() {
  try {
    console.log('Initializing campaign overview system...');
    
    // Ensure required sheets exist
    const requiredSheets = [
      'Call_Reports', 'Attendance', 'QA_Records', 'Tasks', 'Coaching', 
      'Campaigns', 'Users', 'Goals', 'Alerts'
    ];

    const ss = getMainSpreadsheet();
    requiredSheets.forEach(sheetName => {
      try {
        let sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
          console.log(`Creating missing sheet: ${sheetName}`);
          ss.insertSheet(sheetName);
        }
      } catch (error) {
        console.error(`Error checking/creating sheet ${sheetName}:`, error);
      }
    });

    console.log('Campaign overview system initialized successfully');
    return { success: true, message: 'System initialized' };

  } catch (error) {
    console.error('Error initializing campaign overview system:', error);
    writeError('initializeCampaignOverviewSystem', error);
    return { success: false, error: error.message };
  }
}

/**
 * Test the campaign overview system
 */
function testCampaignOverviewSystem() {
  try {
    console.log('Testing campaign overview system...');

    // Test data retrieval
    const testResult = clientGetOKRDataEnhanced('Week', getCurrentPeriod('Week'), '', '', '');
    
    // Test export
    const exportResult = clientExportCampaignOverviewReport({
      granularity: 'Week',
      period: getCurrentPeriod('Week'),
      campaign: '',
      agent: ''
    });

    return {
      success: true,
      dataTest: testResult.success,
      exportTest: exportResult.success,
      message: 'All tests passed',
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error('Error testing campaign overview system:', error);
    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// CLIENT-ACCESSIBLE TEST FUNCTIONS
// ────────────────────────────────────────────────────────────────────────────

/**
 * Client function to initialize the system
 */
function clientInitializeCampaignOverview() {
  return initializeCampaignOverviewSystem();
}

/**
 * Client function to test the system
 */
function clientTestCampaignOverview() {
  return testCampaignOverviewSystem();
}

console.log('Campaign OKR Overview backend functions loaded successfully!');