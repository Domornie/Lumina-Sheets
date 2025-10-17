/**
 * CallReportService_IBTR.gs — IBTR-scoped Call Report service
 * Uses getIBTRSpreadsheet() and campaign helpers from IBTRUtilities.gs
 * Sheet headers assumed from CALL_REPORT_HEADERS in IBTRUtilities:
 * ["ID","CreatedDate","TalkTimeMinutes","To Answer Time","FromRoutingPolicy","WrapupLabel","ToSFUser","UserID","CSAT","CreatedAt","UpdatedAt"]
 */

const __PAGE_SIZE_FALLBACK = 50;

function __ensureDate(value) {
  if (value instanceof Date && !isNaN(value)) return value;
  if (value === null || value === undefined || value === '') return null;
  const parsed = new Date(value);
  return isNaN(parsed) ? null : parsed;
}

function __parseAnswerSeconds(rawValue, createdDate) {
  if (rawValue === null || rawValue === undefined || rawValue === '') return null;

  const created = __ensureDate(createdDate);

  if (rawValue instanceof Date && !isNaN(rawValue)) {
    if (!created) return null;
    const diff = (rawValue.getTime() - created.getTime()) / 1000;
    return isFinite(diff) ? Math.max(0, Math.round(diff * 100) / 100) : null;
  }

  if (typeof rawValue === 'number' && isFinite(rawValue)) {
    let seconds = rawValue;
    if (Math.abs(seconds) > 86400 * 365) seconds = seconds / 1000;
    return Math.max(0, Math.round(seconds * 100) / 100);
  }

  const str = String(rawValue).trim();
  if (!str) return null;

  const numeric = Number(str);
  if (!isNaN(numeric) && isFinite(numeric)) {
    let seconds = numeric;
    if (Math.abs(seconds) > 86400 * 365) seconds = seconds / 1000;
    return Math.max(0, Math.round(seconds * 100) / 100);
  }

  const timeParts = str.split(':');
  if (timeParts.length >= 2 && timeParts.length <= 3) {
    let seconds = 0;
    for (let i = 0; i < timeParts.length; i++) {
      const part = Number(timeParts[i]);
      if (isNaN(part)) {
        seconds = null;
        break;
      }
      seconds = (seconds * 60) + part;
    }
    if (seconds !== null) return Math.max(0, Math.round(seconds * 100) / 100);
  }

  const parsedDate = new Date(str);
  if (!isNaN(parsedDate) && created) {
    const diff = (parsedDate.getTime() - created.getTime()) / 1000;
    return isFinite(diff) ? Math.max(0, Math.round(diff * 100) / 100) : null;
  }

  const secondsMatch = str.match(/(\d+(?:\.\d+)?)/);
  if (secondsMatch) {
    const seconds = Number(secondsMatch[1]);
    if (!isNaN(seconds)) return Math.max(0, Math.round(seconds * 100) / 100);
  }

  return null;
}

function __formatAnswerSeconds(seconds) {
  if (seconds === null || seconds === undefined || seconds === '') return '';
  const numeric = Number(seconds);
  if (isNaN(numeric) || !isFinite(numeric)) return '';
  const totalSeconds = Math.max(0, Math.round(numeric));
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const secs = totalSeconds % 60;
  const hh = String(hours).padStart(2, '0');
  const mm = String(minutes).padStart(2, '0');
  const ss = String(secs).padStart(2, '0');
  return hh + ':' + mm + ':' + ss;
}

function __coerceAnswerCell(rawValue, createdDate) {
  const seconds = __parseAnswerSeconds(rawValue, createdDate);
  return seconds === null ? '' : __formatAnswerSeconds(seconds);
}

const __ANSWER_HEADER_ALIASES = ['To Answer Time', 'ToAnswerTime'];

function __isAnswerHeader(name) {
  if (!name) return false;
  for (let i = 0; i < __ANSWER_HEADER_ALIASES.length; i++) {
    if (name === __ANSWER_HEADER_ALIASES[i]) return true;
  }
  return false;
}

function __getAnswerFieldValue(record) {
  if (!record || typeof record !== 'object') return undefined;
  for (let i = 0; i < __ANSWER_HEADER_ALIASES.length; i++) {
    const key = __ANSWER_HEADER_ALIASES[i];
    if (Object.prototype.hasOwnProperty.call(record, key) && record[key] !== undefined) {
      return record[key];
    }
  }
  return undefined;
}

function __applyAnswerFieldAliases(record, value) {
  if (!record || typeof record !== 'object') return;
  for (let i = 0; i < __ANSWER_HEADER_ALIASES.length; i++) {
    record[__ANSWER_HEADER_ALIASES[i]] = value;
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// Internal: get the CallReport sheet (ensure headers exist)
function __getCallReportSheet() {
  const ss = getIBTRSpreadsheet();
  const sh = ss.getSheetByName(CALL_REPORT) ||
             ensureCampaignSheetWithHeaders(CALL_REPORT, CALL_REPORT_HEADERS);
  return sh;
}

// Internal: read all rows to objects using header row
function __readAllCallReportRows() {
  const sh = __getCallReportSheet();
  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < CALL_REPORT_HEADERS.length) return [];

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0].map(String);
  const rows = sh.getRange(2, 1, lr - 1, lc).getValues();

  return rows.map(r => {
    const obj = {};
    headers.forEach((h, i) => (obj[h] = r[i]));
    // normalize CreatedDate to Date if parsable
    if (obj.CreatedDate && !(obj.CreatedDate instanceof Date)) {
      const d = new Date(obj.CreatedDate);
      if (!isNaN(d)) obj.CreatedDate = d;
    }
    const answerValue = __getAnswerFieldValue(obj);
    if (answerValue !== undefined) {
      __applyAnswerFieldAliases(obj, answerValue);
    }
    return obj;
  });
}

// Internal: find row number by UUID in column A (ID). Returns 0 if not found.
function __findRowById(uuid) {
  if (!uuid) return 0;
  const sh = __getCallReportSheet();
  const lr = sh.getLastRow();
  if (lr < 2) return 0;
  const ids = sh.getRange(2, 1, lr - 1, 1).getValues().map(r => String(r[0] || ''));
  const idx = ids.findIndex(v => v === String(uuid));
  return idx >= 0 ? (idx + 2) : 0; // +2 because range starts at row 2
}

// ───────────────────────────────────────────────────────────────────────────────
// Public API
// ───────────────────────────────────────────────────────────────────────────────

// Read every call report row
function getAllReports() {
  return __readAllCallReportRows();
}

// Unique agent list from ToSFUser
function getUserList() {
  const all = __readAllCallReportRows();
  const set = new Set();
  all.forEach(r => { const a = r.ToSFUser || ''; if (a) set.add(a); });
  return Array.from(set).sort();
}

// Filtered/paged reports; filterObj = { userId?, policy? }
function getFilteredReports(filterObj, pageNum, pageSize) {
  let list = __readAllCallReportRows();

  if (filterObj?.userId) {
    list = list.filter(r => r.ToSFUser === filterObj.userId);
  }
  if (filterObj?.policy) {
    list = list.filter(r => r.FromRoutingPolicy === filterObj.policy);
  }

  const totalCount = list.length;
  const page = parseInt(pageNum, 10) || 1;
  const size = parseInt(pageSize, 10) || (typeof PAGE_SIZE === 'number' ? PAGE_SIZE : __PAGE_SIZE_FALLBACK);
  const start = Math.max(0, (page - 1) * size);
  const slice = list.slice(start, start + size);

  return { reports: slice, totalCount };
}

// Get report by UUID in ID column (not sheet row)
function getCallReportById(id) {
  const rowNum = __findRowById(id);
  if (!rowNum) return null;

  const sh = __getCallReportSheet();
  const lc = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const dataRow = sh.getRange(rowNum, 1, 1, lc).getValues()[0];

  const report = {};
  headers.forEach((h, i) => (report[h] = dataRow[i]));
  return report;
}

// Delete report by UUID
function deleteCallReport(id) {
  const rowNum = __findRowById(id);
  if (!rowNum) throw new Error('Invalid report ID.');
  const sh = __getCallReportSheet();
  sh.deleteRow(rowNum);
  logCampaignDirtyRow(CALL_REPORT, id, 'DELETE');
  flushCampaignDirtyRows();
}

// Create or update report by UUID
// reportData keys should match headers; when creating, CreatedDate will be set if not provided
function createOrUpdateCallReport(reportData) {
  const sh = __getCallReportSheet();
  const lc = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];

  // Coerce CSAT to Yes/No/""
  const rawCsat = (reportData.CSAT || '').toString().trim().toLowerCase();
  if (rawCsat === 'yes' || rawCsat === 'true' || rawCsat === '1') reportData.CSAT = 'Yes';
  else if (rawCsat === 'no' || rawCsat === 'false' || rawCsat === '0') reportData.CSAT = 'No';
  else reportData.CSAT = '';

  const now = new Date();

  if (reportData.ID) {
    // UPDATE by UUID
    const rowNum = __findRowById(reportData.ID);
    if (!rowNum) throw new Error('Invalid report ID for update.');
    const createdDateIdx = headers.indexOf('CreatedDate');
    const existingCreatedDate = createdDateIdx >= 0
      ? sh.getRange(rowNum, createdDateIdx + 1).getValue()
      : null;
    const answerRaw = __getAnswerFieldValue(reportData);
    headers.forEach((h, idx) => {
      if (h === 'ID' || h === 'CreatedDate' || h === 'CreatedAt') return;
      if (__isAnswerHeader(h)) {
        const baseDate = Object.prototype.hasOwnProperty.call(reportData, 'CreatedDate')
          ? reportData.CreatedDate
          : existingCreatedDate;
        if (answerRaw !== undefined) {
          const coerced = __coerceAnswerCell(answerRaw, baseDate);
          __applyAnswerFieldAliases(reportData, coerced);
          sh.getRange(rowNum, idx + 1).setValue(coerced);
        }
        return;
      }
      if (Object.prototype.hasOwnProperty.call(reportData, h)) {
        sh.getRange(rowNum, idx + 1).setValue(reportData[h]);
      }
    });
    // bump UpdatedAt if it exists
    const updatedAtIdx = headers.indexOf('UpdatedAt');
    if (updatedAtIdx >= 0) sh.getRange(rowNum, updatedAtIdx + 1).setValue(now);

    logCampaignDirtyRow(CALL_REPORT, reportData.ID, 'UPDATE');
    flushCampaignDirtyRows();
    return String(reportData.ID);
  }

  // CREATE new
  const uuid = Utilities.getUuid();
  const createAnswerRaw = __getAnswerFieldValue(reportData);
  const coercedCreateAnswer = __coerceAnswerCell(createAnswerRaw, reportData.CreatedDate || now);
  if (createAnswerRaw !== undefined) {
    __applyAnswerFieldAliases(reportData, coercedCreateAnswer);
  }
  const row = headers.map(h => {
    if (h === 'ID') return uuid;
    if (h === 'CreatedDate') return reportData.CreatedDate ? new Date(reportData.CreatedDate) : now;
    if (h === 'CreatedAt') return now;
    if (h === 'UpdatedAt') return now;
    if (__isAnswerHeader(h)) return coercedCreateAnswer;
    if (Object.prototype.hasOwnProperty.call(reportData, h)) {
      const value = reportData[h];
      return value === undefined ? '' : value;
    }
    return '';
  });
  sh.appendRow(row);
  logCampaignDirtyRow(CALL_REPORT, uuid, 'CREATE');
  flushCampaignDirtyRows();
  return String(uuid);
}

/**
 * getAnalyticsByPeriod(granularity, periodIdentifier, agentFilter)
 * Returns exactly what your view expects + activeAgents (the ones with activity in range)
 */
function getAnalyticsByPeriod(granularity, periodIdentifier, agentFilter) {
  const { startDate, endDate } = isoPeriodToDateRange(granularity, periodIdentifier);
  const tz = Session.getScriptTimeZone();

  // Filter by date range (CreatedDate, date-only)
  const dateFiltered = __readAllCallReportRows().filter(r => {
    const dt = r.CreatedDate instanceof Date ? r.CreatedDate : new Date(r.CreatedDate);
    if (isNaN(dt)) return false;
    const d = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
    return d >= startDate && d <= endDate;
  });

  // Agents active in this window
  const activeAgents = Array.from(new Set(dateFiltered.map(r => r.ToSFUser || '—'))).sort();

  // Apply agent filter if provided
  const filtered = (agentFilter && agentFilter !== '')
    ? dateFiltered.filter(r => r.ToSFUser === agentFilter)
    : dateFiltered;

  // repMetrics
  const repMap = {};
  const answerSecondsList = [];
  filtered.forEach(r => {
    const agent = r.ToSFUser || '—';
    const talk  = parseFloat(r.TalkTimeMinutes) || 0;
    if (!repMap[agent]) {
      repMap[agent] = {
        totalCalls: 0,
        totalTalk: 0,
        totalAnswerSeconds: 0,
        answeredCount: 0,
        fastAnswerCount: 0
      };
    }
    repMap[agent].totalCalls += 1;
    repMap[agent].totalTalk  += talk;

    const answerSeconds = __parseAnswerSeconds(__getAnswerFieldValue(r), r.CreatedDate);
    if (answerSeconds !== null) {
      repMap[agent].totalAnswerSeconds += answerSeconds;
      repMap[agent].answeredCount += 1;
      if (answerSeconds <= 30) repMap[agent].fastAnswerCount += 1;
      answerSecondsList.push(answerSeconds);
    }
  });
  const repMetrics = Object.entries(repMap).map(([agent, v]) => {
    const averageAnswerSeconds = v.answeredCount > 0 ? v.totalAnswerSeconds / v.answeredCount : null;
    const fastAnswerRate = v.answeredCount > 0 ? (v.fastAnswerCount / v.answeredCount) * 100 : null;
    return {
      agent,
      totalCalls: v.totalCalls,
      totalTalk: v.totalTalk,
      totalAnswerSeconds: v.totalAnswerSeconds,
      answeredCount: v.answeredCount,
      averageAnswerSeconds,
      fastAnswerRate,
      fastAnswerCount: v.fastAnswerCount
    };
  });

  const answerTimeStats = (function () {
    const answeredCount = answerSecondsList.length;
    if (!answeredCount) {
      return {
        answeredCount: 0,
        averageSeconds: 0,
        medianSeconds: 0,
        p90Seconds: 0,
        fastAnswerRate: 0,
        underMinuteRate: 0,
        slowAnswerRate: 0,
        fastAnswerCount: 0,
        underMinuteCount: 0,
        slowAnswerCount: 0,
        totalSeconds: 0,
        buckets: [],
        fastestResponder: null,
        slowestResponder: null
      };
    }

    const sorted = answerSecondsList.slice().sort((a, b) => a - b);
    const totalSeconds = answerSecondsList.reduce((sum, value) => sum + value, 0);
    const averageSeconds = totalSeconds / answeredCount;
    const medianSeconds = answeredCount % 2 === 0
      ? (sorted[answeredCount / 2 - 1] + sorted[answeredCount / 2]) / 2
      : sorted[Math.floor(answeredCount / 2)];
    const p90Index = Math.min(sorted.length - 1, Math.floor(0.9 * (sorted.length - 1)));
    const p90Seconds = sorted[p90Index];

    let fastCount = 0;
    let underMinuteCount = 0;
    let slowCount = 0;
    const bucketCounts = {
      fast15: 0,
      fast30: 0,
      medium60: 0,
      medium120: 0,
      slow: 0
    };

    answerSecondsList.forEach(seconds => {
      if (seconds <= 15) {
        bucketCounts.fast15 += 1;
        fastCount += 1;
        underMinuteCount += 1;
      } else if (seconds <= 30) {
        bucketCounts.fast30 += 1;
        fastCount += 1;
        underMinuteCount += 1;
      } else if (seconds <= 60) {
        bucketCounts.medium60 += 1;
        underMinuteCount += 1;
      } else if (seconds <= 120) {
        bucketCounts.medium120 += 1;
      } else {
        bucketCounts.slow += 1;
        slowCount += 1;
      }
    });

    const fastAnswerRate = (fastCount / answeredCount) * 100;
    const underMinuteRate = (underMinuteCount / answeredCount) * 100;
    const slowAnswerRate = (slowCount / answeredCount) * 100;

    const buckets = [
      { id: '0_15', label: '0-15 sec', count: bucketCounts.fast15 },
      { id: '16_30', label: '16-30 sec', count: bucketCounts.fast30 },
      { id: '31_60', label: '31-60 sec', count: bucketCounts.medium60 },
      { id: '61_120', label: '1-2 min', count: bucketCounts.medium120 },
      { id: 'gt_120', label: 'Over 2 min', count: bucketCounts.slow }
    ];

    const responderPool = repMetrics
      .filter(r => r.answeredCount >= 3 && r.averageAnswerSeconds !== null)
      .sort((a, b) => a.averageAnswerSeconds - b.averageAnswerSeconds);
    const fastestResponder = responderPool[0] || null;
    const slowestResponder = responderPool.length ? responderPool[responderPool.length - 1] : null;

    return {
      answeredCount,
      averageSeconds,
      medianSeconds,
      p90Seconds,
      fastAnswerRate,
      underMinuteRate,
      slowAnswerRate,
      fastAnswerCount: fastCount,
      underMinuteCount,
      slowAnswerCount: slowCount,
      totalSeconds,
      buckets,
      fastestResponder: fastestResponder
        ? {
            agent: fastestResponder.agent,
            averageAnswerSeconds: fastestResponder.averageAnswerSeconds,
            answeredCount: fastestResponder.answeredCount
          }
        : null,
      slowestResponder: slowestResponder
        ? {
            agent: slowestResponder.agent,
            averageAnswerSeconds: slowestResponder.averageAnswerSeconds,
            answeredCount: slowestResponder.answeredCount
          }
        : null
    };
  })();

  // policyDist
  const policyMap = {};
  filtered.forEach(r => {
    const pol = r.FromRoutingPolicy || '—';
    policyMap[pol] = (policyMap[pol] || 0) + 1;
  });
  const policyDist = Object.entries(policyMap).map(([policy, count]) => ({ policy, count }));

  // wrapDist
  const wrapMap = {};
  filtered.forEach(r => {
    const wl = r.WrapupLabel || '—';
    wrapMap[wl] = (wrapMap[wl] || 0) + 1;
  });
  const wrapDist = Object.entries(wrapMap).map(([wrapupLabel, count]) => ({ wrapupLabel, count }));

  // csatDist
  const csatMap = { Yes: 0, No: 0 };
  filtered.forEach(r => {
    const v = (r.CSAT || '').toString().trim().toLowerCase();
    if (v === 'yes') csatMap.Yes++;
    else if (v === 'no') csatMap.No++;
  });
  const csatDist = [{ csat: 'Yes', count: csatMap.Yes }, { csat: 'No', count: csatMap.No }];

  // weekday buckets for trends
  const msPerDay = 24 * 60 * 60 * 1000;
  const days = Math.floor((endDate - startDate) / msPerDay) + 1;
  const buckets = {};
  for (let i = 0; i < days; i++) {
    const d = new Date(startDate);
    d.setDate(startDate.getDate() + i);
    const dow = d.getDay();
    if (dow === 0 || dow === 6) continue; // skip Sunday/Saturday
    const k = Utilities.formatDate(d, tz, 'MM/dd');
    buckets[k] = { count: 0, talk: 0 };
  }

  filtered.forEach(r => {
    const dt = r.CreatedDate instanceof Date ? r.CreatedDate : new Date(r.CreatedDate);
    if (isNaN(dt)) return;
    const d = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
    const key = Utilities.formatDate(d, tz, 'MM/dd');
    if (!buckets[key]) return;
    buckets[key].count += 1;
    buckets[key].talk += parseFloat(r.TalkTimeMinutes) || 0;
  });

  const callTrend = Object.keys(buckets).map(k => ({ periodLabel: k, callCount: buckets[k].count }));
  const talkTrend = Object.keys(buckets).map(k => ({ periodLabel: k, totalTalk: buckets[k].talk }));

  const fifteenMinuteBuckets = Array.from({ length: 24 * 4 }, () => 0);
  const fifteenMinuteTalk = Array.from({ length: 24 * 4 }, () => 0);

  filtered.forEach(r => {
    const dt = r.CreatedDate instanceof Date ? r.CreatedDate : new Date(r.CreatedDate);
    if (isNaN(dt)) return;
    const hour = Number(Utilities.formatDate(dt, tz, 'H'));
    const minute = Number(Utilities.formatDate(dt, tz, 'm'));
    if (isNaN(hour) || isNaN(minute)) return;
    const slotIndex = hour * 4 + Math.floor(minute / 15);
    if (slotIndex < 0 || slotIndex >= fifteenMinuteBuckets.length) return;
    fifteenMinuteBuckets[slotIndex] += 1;
    fifteenMinuteTalk[slotIndex] += parseFloat(r.TalkTimeMinutes) || 0;
  });

  const totalCallsAll = filtered.length;
  const activeSlotCount = fifteenMinuteBuckets.filter(count => count > 0).length || 1;
  const averageActiveLoad = activeSlotCount ? totalCallsAll / activeSlotCount : 0;

  const toTimeLabel = totalMinutes => {
    const normalized = ((totalMinutes % 1440) + 1440) % 1440;
    const hour = Math.floor(normalized / 60);
    const minute = normalized % 60;
    const hour12 = ((hour + 11) % 12) + 1;
    const ampm = hour < 12 ? 'AM' : 'PM';
    return `${hour12}:${String(minute).padStart(2, '0')} ${ampm}`;
  };

  const formatSlotRange = (startSlot, endSlot) => {
    const startMinutes = startSlot * 15;
    const endMinutes = endSlot * 15;
    return `${toTimeLabel(startMinutes)} – ${toTimeLabel(endMinutes)}`;
  };

  const classifyLoad = avgPerSlot => {
    if (!avgPerSlot || !isFinite(avgPerSlot) || averageActiveLoad === 0) return 'low';
    if (avgPerSlot <= averageActiveLoad * 0.75) return 'low';
    if (avgPerSlot <= averageActiveLoad * 1.25) return 'moderate';
    return 'high';
  };

  const intensityLabel = key => {
    if (key === 'high') return 'High load';
    if (key === 'moderate') return 'Moderate load';
    return 'Low load';
  };

  const intervalVolume = fifteenMinuteBuckets.map((count, idx) => ({
    slotIndex: idx,
    windowLabel: formatSlotRange(idx, idx + 1),
    callCount: count,
    averageTalk: fifteenMinuteTalk[idx],
    intensity: classifyLoad(count),
    intensityLabel: intensityLabel(classifyLoad(count))
  }));

  const hourlyVolume = Array.from({ length: 24 }, (_, hour) => {
    const start = hour * 4;
    const end = start + 4;
    const callCount = fifteenMinuteBuckets.slice(start, end).reduce((sum, v) => sum + v, 0);
    const talkTotal = fifteenMinuteTalk.slice(start, end).reduce((sum, v) => sum + v, 0);
    return {
      hour,
      label: `${String(hour).padStart(2, '0')}:00`,
      windowLabel: `${toTimeLabel(hour * 60)} – ${toTimeLabel((hour + 1) * 60)}`,
      callCount,
      averageTalk: callCount > 0 ? talkTotal / callCount : 0,
      intensity: classifyLoad(callCount / 4 || 0),
      intensityLabel: intensityLabel(classifyLoad(callCount / 4 || 0))
    };
  });

  const peakTimeWindows = intervalVolume
    .filter(entry => entry.callCount > 0)
    .sort((a, b) => b.callCount - a.callCount)
    .slice(0, 5)
    .map(entry => ({
      slotIndex: entry.slotIndex,
      windowLabel: entry.windowLabel,
      callCount: entry.callCount,
      shareOfDay: totalCallsAll > 0 ? (entry.callCount / totalCallsAll) * 100 : 0,
      intensity: entry.intensity,
      intensityLabel: entry.intensityLabel
    }));

  const findBestWindow = (startIdx, endIdx, length) => {
    if (startIdx < 0) startIdx = 0;
    if (endIdx >= fifteenMinuteBuckets.length) endIdx = fifteenMinuteBuckets.length - 1;
    if (length <= 0 || startIdx > endIdx) return null;
    let best = null;
    for (let idx = startIdx; idx <= endIdx - length + 1; idx++) {
      let sum = 0;
      for (let j = 0; j < length; j++) {
        sum += fifteenMinuteBuckets[idx + j];
      }
      if (!best || sum < best.callCount || (sum === best.callCount && idx < best.startSlot)) {
        best = { startSlot: idx, endSlot: idx + length, callCount: sum };
      }
    }
    return best;
  };

  const globalPeakSlots = new Set(peakTimeWindows.map(p => p.slotIndex));

  const convertWindow = (window, shiftTotal) => {
    if (!window) {
      return {
        rangeLabel: '—',
        callCount: 0,
        shareOfShift: 0,
        intensity: 'low',
        intensityLabel: 'Low load'
      };
    }
    const slots = Math.max(window.endSlot - window.startSlot, 1);
    const avgPerSlot = window.callCount / slots;
    const intensityKey = classifyLoad(avgPerSlot);
    return {
      rangeLabel: formatSlotRange(window.startSlot, window.endSlot),
      callCount: window.callCount,
      shareOfShift: shiftTotal > 0 ? (window.callCount / shiftTotal) * 100 : 0,
      intensity: intensityKey,
      intensityLabel: intensityLabel(intensityKey)
    };
  };

  const scheduleTemplates = [
    { id: 'early', name: 'Early Morning', startSlot: (8 * 4) + 2, startLabel: '8:30 AM' },
    { id: 'mid', name: 'Mid-Morning', startSlot: 9 * 4, startLabel: '9:00 AM' },
    { id: 'late', name: 'Late Morning', startSlot: (9 * 4) + 2, startLabel: '9:30 AM' },
    { id: 'afternoon', name: 'Afternoon', startSlot: (9 * 4) + 2, startLabel: '9:30 AM', note: 'Extends coverage deepest into the afternoon window.' }
  ];

  const scheduleRecommendations = scheduleTemplates.map(template => {
    const shiftLengthSlots = 8 * 4;
    const startSlot = template.startSlot;
    const endSlot = startSlot + shiftLengthSlots;
    const shiftIntervals = intervalVolume.filter(entry => entry.slotIndex >= startSlot && entry.slotIndex < endSlot);
    const shiftTotalCalls = shiftIntervals.reduce((sum, entry) => sum + entry.callCount, 0);
    const averageHourlyLoad = shiftTotalCalls / 8;

    const lunchWindow = findBestWindow(startSlot + 12, Math.min(endSlot - 4, startSlot + 20), 4);
    const firstBreakWindow = findBestWindow(startSlot + 4, Math.min(endSlot - 1, startSlot + 12), 1);
    const secondBreakWindow = findBestWindow(Math.max(startSlot + 20, startSlot + 8), endSlot - 1, 1);

    const peakWithinShift = shiftIntervals
      .slice()
      .sort((a, b) => b.callCount - a.callCount)[0] || null;

    const overlapCount = Array.from(globalPeakSlots).filter(slot => slot >= startSlot && slot < endSlot).length;
    const overlapPercent = globalPeakSlots.size > 0 ? (overlapCount / globalPeakSlots.size) * 100 : 0;

    const coverageNote = overlapPercent >= 75
      ? 'Strong overlap with historic spikes—keep full coverage during the watch window.'
      : overlapPercent >= 40
        ? 'Moderate overlap—stagger lunches as recommended to stay ahead of the peak.'
        : 'Light overlap with peaks—ideal window for flexible coverage.';

    return {
      id: template.id,
      name: template.name,
      startLabel: template.startLabel,
      shiftRangeLabel: formatSlotRange(startSlot, endSlot),
      totalCalls: shiftTotalCalls,
      averageHourlyLoad: averageHourlyLoad || 0,
      overlapPercent,
      lunch: convertWindow(lunchWindow, shiftTotalCalls),
      breaks: [
        convertWindow(firstBreakWindow, shiftTotalCalls),
        convertWindow(secondBreakWindow, shiftTotalCalls)
      ],
      peakGuard: peakWithinShift
        ? {
            rangeLabel: peakWithinShift.windowLabel,
            callCount: peakWithinShift.callCount,
            shareOfShift: shiftTotalCalls > 0 ? (peakWithinShift.callCount / shiftTotalCalls) * 100 : 0,
            intensity: peakWithinShift.intensity,
            intensityLabel: peakWithinShift.intensityLabel
          }
        : convertWindow(null, shiftTotalCalls),
      coverageNote: template.note ? `${template.note} ${coverageNote}`.trim() : coverageNote
    };
  });

  return {
    repMetrics,
    policyDist,
    wrapDist,
    callTrend,
    talkTrend,
    csatDist,
    activeAgents,
    hourlyVolume,
    intervalVolume,
    peakTimeWindows,
    scheduleRecommendations,
    answerTimeStats
  };
}

/**
 * importCallReports(rows) — IBTR-scoped; de-dupes via composite key and logs dirty in batch
 */
function importCallReports(rows) {
  const sh = __getCallReportSheet();
  const now = new Date();

  // Build Set of existing composite keys from B–I
  const lr = sh.getLastRow();
  const existingKeys = new Set();
  if (lr > 1) {
    const raw = sh.getRange(2, 2, lr - 1, 8).getValues(); // B..I
    raw.forEach(row => {
      const rawDate = row[0];
      const rawTalk = row[1];
      const rawAnswer = row[2];
      const policy  = (row[3] || '').toString().trim();
      const wrap    = (row[4] || '').toString().trim();
      const user    = (row[5] || '').toString().trim().replace(/\s*VLBPO\s*/gi, ' ').trim();
      const csat    = (row[7] || '').toString().trim();

      let dateIso;
      if (rawDate instanceof Date && !isNaN(rawDate)) dateIso = rawDate.toISOString();
      else {
        const t = new Date(rawDate);
        dateIso = !isNaN(t) ? t.toISOString() : String(rawDate || '').trim();
      }
      const talkStr = Number(rawTalk) >= 0 ? String(Number(rawTalk)) : String(rawTalk || '0');
      const answerSeconds = __parseAnswerSeconds(rawAnswer, rawDate);
      const answerKey = answerSeconds !== null ? String(answerSeconds) : '';
      const key = [
        dateIso,
        talkStr,
        answerKey,
        policy,
        wrap,
        user,
        (csat.toLowerCase() === 'yes' ? 'Yes' : csat.toLowerCase() === 'no' ? 'No' : '')
      ].join('||');
      existingKeys.add(key);
    });
  }

  // Build new rows
  const seen = new Set();
  const toAppend = [];
  let skipped = 0;

  rows.forEach(r => {
    // Normalize incoming
    const d = new Date(r.Date);
    const dateIso = !isNaN(d) ? d.toISOString() : String(r.Date || '').trim();
    const talk = Number(r.TalkTimeMinutes) || 0;
    const talkStr = String(talk);
    const policy = (r.Policy || '').toString().trim();
    const wrap   = (r.Wrapup || '').toString().trim();
    const user   = (r.UserName || '').toString().trim().replace(/\s*VLBPO\s*/gi, ' ').trim();
    const userId = (r.UserID || '').toString().trim();
    const lc = (r.CSAT || '').toString().trim().toLowerCase();
    const csat = lc === 'yes' ? 'Yes' : lc === 'no' ? 'No' : '';
    const incomingAnswer = __getAnswerFieldValue(r);
    const answerValue = __coerceAnswerCell(incomingAnswer, r.Date || dateIso);
    __applyAnswerFieldAliases(r, answerValue);
    const answerSeconds = __parseAnswerSeconds(answerValue, r.Date || dateIso);
    const answerKey = answerSeconds !== null ? String(answerSeconds) : '';

    const key = [dateIso, talkStr, answerKey, policy, wrap, user, csat].join('||');
    if (existingKeys.has(key) || seen.has(key)) {
      skipped += 1;
      return;
    }
    seen.add(key);

    const uuid = Utilities.getUuid();
    toAppend.push([
      uuid,                 // A: ID
      new Date(dateIso),    // B: CreatedDate
      talk,                 // C: TalkTimeMinutes
      answerValue,          // D: To Answer Time (HH:MM:SS)
      policy,               // E: FromRoutingPolicy
      wrap,                 // F: WrapupLabel
      user,                 // G: ToSFUser
      userId,               // H: UserID
      csat,                 // I: CSAT
      now,                  // J: CreatedAt
      now                   // K: UpdatedAt
    ]);
    logCampaignDirtyRow(CALL_REPORT, uuid, 'CREATE');
  });

  if (toAppend.length) {
    const first = sh.getLastRow() + 1;
    sh.getRange(first, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
  }
  flushCampaignDirtyRows();
  return { imported: toAppend.length, skipped };
}

// Remove "VLBPO" tokens from ToSFUser column
function cleanCallReportUserNames() {
  const sh = __getCallReportSheet();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = headers.indexOf('ToSFUser');
  if (idx < 0) throw new Error('"ToSFUser" column not found.');

  const lr = sh.getLastRow();
  if (lr < 2) return 0;

  const range = sh.getRange(2, idx + 1, lr - 1, 1);
  const vals = range.getValues();
  const cleaned = vals.map(r => [String(r[0] || '').replace(/\s*VLBPO\s*/gi, ' ').trim()]);
  range.setValues(cleaned);
  return cleaned.length;
}

/**
 * exportCallAnalyticsCsv(granularity, periodIdentifier, agentFilter)
 * 6 blocks, blank-line separated (same as your current UI expects)
 */
function exportCallAnalyticsCsv(granularity, periodIdentifier, agentFilter) {
  const a = getAnalyticsByPeriod(granularity, periodIdentifier, agentFilter);

  const toCsv = rows => rows.map(r => r.map(c => `"${c}"`).join(',')).join('\r\n');

  const repRows = [['Agent','TotalCalls','TotalTalk(min)']]
    .concat(a.repMetrics.map(r => [r.agent, r.totalCalls, r.totalTalk]));

  const policyRows = [['Policy','Count']]
    .concat(a.policyDist.map(o => [o.policy, o.count]));

  const wrapRows = [['Wrapup','Count']]
    .concat(a.wrapDist.map(o => [o.wrapupLabel, o.count]));

  const callTrendRows = [['Period','CallCount']]
    .concat(a.callTrend.map(o => [o.periodLabel, o.callCount]));

  const talkTrendRows = [['Period','TotalTalk(min)']]
    .concat(a.talkTrend.map(o => [o.periodLabel, o.totalTalk]));

  const csatRows = [['CSAT','Count']]
    .concat(a.csatDist.map(o => [o.csat, o.count]));

  return [toCsv(repRows), toCsv(policyRows), toCsv(wrapRows), toCsv(callTrendRows), toCsv(talkTrendRows), toCsv(csatRows)].join('\r\n\r\n');
}
