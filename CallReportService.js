/**
 * CallReportService_IBTR.gs — IBTR-scoped Call Report service
 * Uses getIBTRSpreadsheet() and campaign helpers from IBTRUtilities.gs
 * Sheet headers assumed from CALL_REPORT_HEADERS in IBTRUtilities:
 * ["ID","CreatedDate","TalkTimeMinutes","FromRoutingPolicy","WrapupLabel","ToSFUser","UserID","CSAT","CreatedAt","UpdatedAt"]
 */

const __PAGE_SIZE_FALLBACK = 50;

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
    headers.forEach((h, idx) => {
      if (h === 'ID' || h === 'CreatedDate' || h === 'CreatedAt') return;
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
  const row = headers.map(h => {
    if (h === 'ID') return uuid;
    if (h === 'CreatedDate') return reportData.CreatedDate ? new Date(reportData.CreatedDate) : now;
    if (h === 'CreatedAt') return now;
    if (h === 'UpdatedAt') return now;
    return reportData[h] || '';
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
  filtered.forEach(r => {
    const agent = r.ToSFUser || '—';
    const talk  = parseFloat(r.TalkTimeMinutes) || 0;
    if (!repMap[agent]) repMap[agent] = { totalCalls: 0, totalTalk: 0 };
    repMap[agent].totalCalls += 1;
    repMap[agent].totalTalk  += talk;
  });
  const repMetrics = Object.entries(repMap).map(([agent, v]) => ({
    agent,
    totalCalls: v.totalCalls,
    totalTalk:  v.totalTalk
  }));

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

  return { repMetrics, policyDist, wrapDist, callTrend, talkTrend, csatDist, activeAgents };
}

/**
 * importCallReports(rows) — IBTR-scoped; de-dupes via composite key and logs dirty in batch
 */
function importCallReports(rows) {
  const sh = __getCallReportSheet();
  const now = new Date();

  // Build Set of existing composite keys from B–H
  const lr = sh.getLastRow();
  const existingKeys = new Set();
  if (lr > 1) {
    const raw = sh.getRange(2, 2, lr - 1, 7).getValues(); // B..H
    raw.forEach(row => {
      const rawDate = row[0];
      const rawTalk = row[1];
      const policy  = (row[2] || '').toString().trim();
      const wrap    = (row[3] || '').toString().trim();
      const user    = (row[4] || '').toString().trim().replace(/\s*VLBPO\s*/gi, ' ').trim();
      const csat    = (row[6] || '').toString().trim();

      let dateIso;
      if (rawDate instanceof Date && !isNaN(rawDate)) dateIso = rawDate.toISOString();
      else {
        const t = new Date(rawDate);
        dateIso = !isNaN(t) ? t.toISOString() : String(rawDate || '').trim();
      }
      const talkStr = Number(rawTalk) >= 0 ? String(Number(rawTalk)) : String(rawTalk || '0');
      const key = [dateIso, talkStr, policy, wrap, user, (csat.toLowerCase() === 'yes' ? 'Yes' : csat.toLowerCase() === 'no' ? 'No' : '')].join('||');
      existingKeys.add(key);
    });
  }

  // Build new rows
  const seen = new Set();
  const toAppend = [];

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

    const key = [dateIso, talkStr, policy, wrap, user, csat].join('||');
    if (existingKeys.has(key) || seen.has(key)) return;
    seen.add(key);

    const uuid = Utilities.getUuid();
    toAppend.push([
      uuid,                 // A: ID
      new Date(dateIso),    // B: CreatedDate
      talk,                 // C: TalkTimeMinutes
      policy,               // D: FromRoutingPolicy
      wrap,                 // E: WrapupLabel
      user,                 // F: ToSFUser
      userId,               // G: UserID
      csat,                 // H: CSAT
      now,                  // I: CreatedAt
      now                   // J: UpdatedAt
    ]);
    logCampaignDirtyRow(CALL_REPORT, uuid, 'CREATE');
  });

  if (toAppend.length) {
    const first = sh.getLastRow() + 1;
    sh.getRange(first, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
  }
  flushCampaignDirtyRows();
  return { imported: toAppend.length };
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
