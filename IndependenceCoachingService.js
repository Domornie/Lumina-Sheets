/**
 * CoachingService.gs — IndependenceCoaching service functions
 * Updated to align with Independence* utilities and sheet schema.
 * - Uses INDEPENDENCE_SHEET_ID and IndependenceCoaching sheet names
 * - Honors INDEPENDENCE_COACHING_HEADERS / QA_RECORDS from utilities
 * - Robust date formatting and topics parsing (JSON or CSV)
 * - Safer lookups and error logging via safeWriteError (if available)
 */

/** Internal: open the shared Independence spreadsheet */
function getIndependenceSpreadsheet_() {
  try {
    if (typeof INDEPENDENCE_SHEET_ID === 'string' && INDEPENDENCE_SHEET_ID.length > 10) {
      return SpreadsheetApp.openById(INDEPENDENCE_SHEET_ID);
    }
    // Fallback to legacy accessor if present
    if (typeof getIBTRSpreadsheet === 'function') return getIBTRSpreadsheet();
    throw new Error('Spreadsheet not available (INDEPENDENCE_SHEET_ID missing and getIBTRSpreadsheet() not defined).');
  } catch (e) {
    if (typeof safeWriteError === 'function') safeWriteError('getIndependenceSpreadsheet_', e);
    throw e;
  }
}

/** Internal: format Date → 'yyyy-MM-dd' using script TZ */
function fmtIsoDateYMD_(val) {
  if (!(val instanceof Date)) return val;
  return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/** Internal: normalize end/start of day (strip time) */
function stripTime_(d, endOfDay) {
  const x = new Date(d);
  if (endOfDay) x.setHours(23, 59, 59, 999);
  else x.setHours(0, 0, 0, 0);
  return x;
}

/** Internal: ISO week label like '2025-W34' */
function weekStringFromDate_(d) {
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const dayNum = date.getUTCDay() || 7; // Monday=1..Sunday=7
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
  return `${date.getUTCFullYear()}-W${String(weekNo).padStart(2, '0')}`;
}

/** Internal: resolve date range (uses isoPeriodToDateRange if available) */
function resolveDateRange_(granularity, period) {
  if (typeof isoPeriodToDateRange === 'function') return isoPeriodToDateRange(granularity, period);

  // Fallback parser for 'YYYY', 'YYYY-MM', 'YYYY-Www'
  const today = new Date();
  let start, end;

  if (/^\d{4}$/.test(period)) {
    const y = parseInt(period, 10);
    start = new Date(y, 0, 1);
    end   = new Date(y, 11, 31, 23, 59, 59, 999);
  } else if (/^\d{4}-\d{2}$/.test(period)) {
    const [y, m] = period.split('-').map(n => parseInt(n, 10));
    start = new Date(y, m - 1, 1);
    end   = new Date(y, m, 0, 23, 59, 59, 999);
  } else if (/^\d{4}-W\d{2}$/.test(period)) {
    const [y, wStr] = period.split('-W');
    const yNum = parseInt(y, 10);
    const w = parseInt(wStr, 10);
    const jan4 = new Date(yNum, 0, 4);
    const dayOfWeek = (jan4.getDay() + 6) % 7; // Monday=0
    start = new Date(jan4);
    start.setDate(jan4.getDate() - dayOfWeek + (w - 1) * 7);
    end = new Date(start);
    end.setDate(start.getDate() + 6);
    end.setHours(23, 59, 59, 999);
  } else {
    start = new Date(today.getFullYear(), today.getMonth(), 1);
    end   = new Date(today.getFullYear(), today.getMonth() + 1, 0, 23, 59, 59, 999);
  }
  return { startDate: stripTime_(start), endDate: stripTime_(end, true) };
}

/** Internal: parse topics from JSON or CSV into a string[] */
function parseTopics_(value) {
  if (!value && value !== 0) return [];
  if (Array.isArray(value)) return value.map(v => String(v).trim()).filter(Boolean);
  try {
    const parsed = JSON.parse(String(value));
    if (Array.isArray(parsed)) return parsed.map(v => String(v).trim()).filter(Boolean);
  } catch (_) { /* fallthrough to csv */ }
  return String(value)
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);
}

/** Internal: ensure headers exist (adds missing columns to the right if needed) */
function ensureHeaders_(sh, canonicalHeaders) {
  const headerRange = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), canonicalHeaders.length));
  const current = headerRange.getValues()[0];
  if (!current[0]) {
    sh.getRange(1, 1, 1, canonicalHeaders.length)
      .setValues([canonicalHeaders])
      .setFontWeight('bold')
      .setBackground('#003177')
      .setFontColor('white');
    sh.setFrozenRows(1);
    return canonicalHeaders.slice();
  }
  const missing = canonicalHeaders.filter(h => current.indexOf(h) === -1);
  if (missing.length > 0) {
    const merged = current.slice(0, canonicalHeaders.length);
    for (let i = 0; i < canonicalHeaders.length; i++) {
      merged[i] = canonicalHeaders[i] || merged[i] || '';
    }
    sh.getRange(1, 1, 1, merged.length).setValues([merged]);
    return merged;
  }
  return current;
}

/** @returns {Array<Object>} each object has keys from INDEPENDENCE_COACHING_HEADERS */
function getAllCoaching() {
  try {
    const ss = getIndependenceSpreadsheet_();
    const sh = ss.getSheetByName(typeof INDEPENDENCE_COACHING_SHEET === 'string' ? INDEPENDENCE_COACHING_SHEET : 'IndependenceCoaching');
    if (!sh) return [];

    const data = sh.getDataRange().getValues();
    if (data.length < 2) return [];

    const headers = data.shift();
    const tz = Session.getScriptTimeZone();
    const dateFields = (typeof COACHING_DATE_COLUMNS !== 'undefined' && Array.isArray(COACHING_DATE_COLUMNS))
      ? COACHING_DATE_COLUMNS
      : ['SessionDate', 'FollowUpDate', 'AcknowledgedOn', 'CreatedAt', 'UpdatedAt'];

    return data.map(row => {
      const obj = {};
      row.forEach((cell, i) => {
        const key = headers[i];
        if (dateFields.indexOf(key) !== -1 && cell instanceof Date) {
          obj[key] = Utilities.formatDate(cell, tz, 'yyyy-MM-dd');
        } else {
          obj[key] = cell;
        }
      });
      return obj;
    });
  } catch (error) {
    if (typeof safeWriteError === 'function') safeWriteError('getAllCoaching', error);
    return [];
  }
}

// ─────────── Get one record by ID ───────────
/** @param {string} id */
function getCoachingRecordById(id) {
  const all = getAllCoaching();
  const rec = all.find(r => String(r.ID) === String(id));
  if (!rec) throw new Error('Coaching record not found: ' + id);
  return rec;
}

/**
 * Returns an array of {id, date, percentage, agentName, agentEmail}
 * for populating the QA dropdown.
 */
function getQAItems() {
  try {
    const ss = getIndependenceSpreadsheet_();
    const qaSheetName = (typeof QA_RECORDS === 'string' && QA_RECORDS) ? QA_RECORDS : 'IndependenceQA';
    const sh = ss.getSheetByName(qaSheetName);
    if (!sh) return [];

    const data = sh.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data.shift();

    // find the columns (support both "PercentageScore" and legacy "Percentage")
    const idIdx    = headers.indexOf('ID');
    const dateIdx  = headers.indexOf('CallDate');
    let pctIdx     = headers.indexOf('PercentageScore');
    if (pctIdx < 0) pctIdx = headers.indexOf('Percentage');
    const agIdx    = headers.indexOf('AgentName');
    const emailIdx = headers.indexOf('AgentEmail');

    const tz = Session.getScriptTimeZone();

    return data.map(row => {
      // format the date
      const rawDate = row[dateIdx];
      const dateStr = rawDate instanceof Date
        ? Utilities.formatDate(rawDate, tz, 'yyyy-MM-dd')
        : String(rawDate || '');

      // coerce the percentage
      let pct = row[pctIdx];
      if (typeof pct !== 'number') pct = parseFloat(pct) || 0;

      return {
        id: row[idIdx],
        date: dateStr,
        percentage: pct,
        agentName: row[agIdx],
        agentEmail: emailIdx >= 0 ? row[emailIdx] : ''
      };
    });
  } catch (error) {
    if (typeof safeWriteError === 'function') safeWriteError('getQAItems', error);
    return [];
  }
}

/**
 * Compute metrics for coaching dashboard.
 * @param {'Week'|'Month'|'Quarter'|'Year'} granularity
 * @param {string} period e.g., '2025-W34', '2025-08', '2025'
 * @param {string=} agentFilter
 */
function getDashboardCoaching(granularity, period, agentFilter) {
  try {
    const all = getAllCoaching();
    const { startDate, endDate } = resolveDateRange_(granularity, period);
    const tz = Session.getScriptTimeZone();

    // 1) Filter
    const filtered = all.filter(r => {
      if (!r.SessionDate) return false;
      const d = new Date(String(r.SessionDate) + 'T00:00:00');
      if (isNaN(d)) return false;
      return d >= startDate && d <= endDate && (!agentFilter || String(r.AgentName) === String(agentFilter));
    });

    // 2) Trend buckets
    const buckets = {};
    filtered.forEach(r => {
      const d = new Date(String(r.SessionDate) + 'T00:00:00');
      if (isNaN(d)) return;
      let label;
      if (granularity === 'Week')       label = weekStringFromDate_(d);
      else if (granularity === 'Month') label = Utilities.formatDate(d, tz, 'yyyy-MM');
      else if (granularity === 'Quarter') {
        const q = Math.floor(d.getMonth() / 3) + 1;
        label = `Q${q}-${d.getFullYear()}`;
      } else                             label = `${d.getFullYear()}`;
      buckets[label] = (buckets[label] || 0) + 1;
    });
    const sessionsTrend = Object.entries(buckets)
      .sort((a, b) => (a[0] < b[0] ? -1 : 1))
      .map(([periodLabel, count]) => ({ periodLabel, count }));

    // 3) Topics distribution (prefers CoveredTopics JSON; falls back to TopicsPlanned CSV/JSON)
    const topicCounts = {};
    filtered.forEach(r => {
      let topics = [];
      if (r.CoveredTopics) topics = parseTopics_(r.CoveredTopics);
      else if (r.TopicsPlanned) topics = parseTopics_(r.TopicsPlanned);
      topics.forEach(t => {
        topicCounts[t] = (topicCounts[t] || 0) + 1;
      });
    });
    const topicsDist = Object.entries(topicCounts).map(([topic, count]) => ({ topic, count }));

    // 4) Upcoming follow-ups (next 20)
    const today = stripTime_(new Date());
    const upcoming = all
      .filter(r => {
        if (!r.FollowUpDate) return false;
        const f = new Date(String(r.FollowUpDate) + 'T00:00:00');
        if (isNaN(f)) return false;
        return f >= today && (!agentFilter || String(r.AgentName) === String(agentFilter));
      })
      .sort((a, b) => new Date(a.FollowUpDate) - new Date(b.FollowUpDate))
      .slice(0, 20);

    return { sessionsTrend, topicsDist, upcoming };
  } catch (error) {
    if (typeof safeWriteError === 'function') safeWriteError('getDashboardCoaching', error);
    return { sessionsTrend: [], topicsDist: [], upcoming: [] };
  }
}

/**
 * Create/save a coaching session row.
 * @param {{
 *   qaId?: string,
 *   coachingDate: string,   // 'YYYY-MM-DD'
 *   coachName: string,
 *   employeeName: string,
 *   coacheeEmail: string,
 *   topicsPlanned: string,  // JSON string or CSV
 *   summary?: string,
 *   plan: string,
 *   followUpDate?: string,  // 'YYYY-MM-DD'
 *   followUpNotes?: string
 * }} sess
 * @return {{id: string, ts: string}}
 */
function saveCoachingSession(sess) {
  try {
    const ss = getIndependenceSpreadsheet_();
    const sh = ss.getSheetByName(typeof INDEPENDENCE_COACHING_SHEET === 'string' ? INDEPENDENCE_COACHING_SHEET : 'IndependenceCoaching');
    if (!sh) throw new Error('Missing Coaching sheet: ' + INDEPENDENCE_COACHING_SHEET);

    // Ensure headers
    const headers = ensureHeaders_(sh, (typeof INDEPENDENCE_COACHING_HEADERS !== 'undefined' ? INDEPENDENCE_COACHING_HEADERS : [
      "ID","QAId","SessionDate","AgentName","CoacheeName","CoacheeEmail","TopicsPlanned","CoveredTopics","ActionPlan","FollowUpDate","Notes","Completed","AcknowledgementText","AcknowledgedOn","CreatedAt","UpdatedAt"
    ]));

    const nowIso = new Date().toISOString();
    const prefix = (typeof COACHING_ID_PREFIX === 'string' && COACHING_ID_PREFIX) ? COACHING_ID_PREFIX : 'COACH';
    const id = `${prefix}_${Date.now()}_${Math.floor(Math.random() * 1000).toString().padStart(3, '0')}`;

    const rowMap = {
      ID: id,
      QAId: sess.qaId || '',
      SessionDate: sess.coachingDate || '',
      AgentName: sess.coachName || '',
      CoacheeName: sess.employeeName || '',
      CoacheeEmail: sess.coacheeEmail || '',
      TopicsPlanned: sess.topicsPlanned || '',
      CoveredTopics: JSON.stringify([]),
      ActionPlan: sess.plan || '',
      FollowUpDate: sess.followUpDate || '',
      Notes: sess.followUpNotes || '',
      Completed: '',
      AcknowledgementText: '',
      AcknowledgedOn: '',
      CreatedAt: nowIso,
      UpdatedAt: nowIso
    };

    const row = headers.map(h => (rowMap[h] !== undefined ? rowMap[h] : ''));
    sh.appendRow(row);

    return { id, ts: nowIso };
  } catch (error) {
    if (typeof safeWriteError === 'function') safeWriteError('saveCoachingSession', error);
    throw error;
  }
}

/**
 * Update CoveredTopics for a coaching session.
 * @param {string} id
 * @param {string[]} coveredArray
 * @return {string} JSON-stringified coveredArray
 */
function updateCoveredTopics(id, coveredArray) {
  try {
    const ss = getIndependenceSpreadsheet_();
    const sh = ss.getSheetByName(typeof INDEPENDENCE_COACHING_SHEET === 'string' ? INDEPENDENCE_COACHING_SHEET : 'IndependenceCoaching');
    if (!sh) throw new Error('Missing Coaching sheet: ' + INDEPENDENCE_COACHING_SHEET);

    const data = sh.getDataRange().getValues();
    if (data.length < 2) throw new Error('Empty coaching sheet');

    const headers = data.shift();
    const idCol  = headers.indexOf('ID');
    let covCol   = headers.indexOf('CoveredTopics');
    const updCol = headers.indexOf('UpdatedAt');

    if (idCol < 0) throw new Error('Missing ID column');
    if (updCol < 0) throw new Error('Missing UpdatedAt column');
    if (covCol < 0) {
      sh.getRange(1, headers.length + 1).setValue('CoveredTopics');
      covCol = headers.length;
    }

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][idCol]) === String(id)) {
        const rowIdx  = i + 2;
        const jsonStr = JSON.stringify(Array.isArray(coveredArray) ? coveredArray : []);
        sh.getRange(rowIdx, covCol + 1).setValue(jsonStr);
        sh.getRange(rowIdx, updCol + 1).setValue(new Date().toISOString());
        return jsonStr;
      }
    }
    throw new Error('Coaching record not found: ' + id);
  } catch (error) {
    if (typeof safeWriteError === 'function') safeWriteError('updateCoveredTopics', error);
    throw error;
  }
}

/**
 * Marks the session complete, emails the coachee, and includes the ack-form link.
 */
function completeCoachingAndNotify(id) {
  try {
    const baseUrl = ScriptApp.getService().getUrl();
    const ackUrl  = `${baseUrl}&page=ackform&id=${encodeURIComponent(id)}`;

    const ss   = getIndependenceSpreadsheet_();
    const sh   = ss.getSheetByName(typeof INDEPENDENCE_COACHING_SHEET === 'string' ? INDEPENDENCE_COACHING_SHEET : 'IndependenceCoaching');
    if (!sh) throw new Error('Missing Coaching sheet: ' + INDEPENDENCE_COACHING_SHEET);

    const data = sh.getDataRange().getValues();
    if (data.length < 2) throw new Error('No coaching rows to notify.');

    const headers = data.shift();
    const idCol    = headers.indexOf('ID');
    const emailCol = headers.indexOf('CoacheeEmail');
    const nameCol  = headers.indexOf('CoacheeName');
    const coachCol = headers.indexOf('AgentName');
    const dateCol  = headers.indexOf('SessionDate');
    const updCol   = headers.indexOf('UpdatedAt');
    const compCol  = headers.indexOf('Completed');

    if (emailCol < 0) throw new Error('Cannot send notification — no “CoacheeEmail” column in sheet.');

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][idCol]) === String(id)) {
        const rowIdx    = i + 2;
        const recipient = data[i][emailCol];
        const coachee   = data[i][nameCol];
        const coachName = data[i][coachCol];
        const sessionOn = data[i][dateCol];

        if (compCol >= 0) sh.getRange(rowIdx, compCol + 1).setValue('Yes');
        if (updCol  >= 0) sh.getRange(rowIdx, updCol  + 1).setValue(new Date().toISOString());

        MailApp.sendEmail({
          to: recipient,
          subject: `Coaching Session on ${sessionOn} Completed`,
          htmlBody: `
            <div style="font-family: Arial, sans-serif; background-color: #f4f6f8; padding: 20px;">
              <table width="100%" cellpadding="0" cellspacing="0" style="max-width:600px; margin:auto; background:#ffffff; border-radius:8px; overflow:hidden; box-shadow:0 2px 8px rgba(0,0,0,0.1);">
                <tr>
                  <td style="background-color:#4e73df; padding:16px; text-align:center;">
                    <h1 style="color:#ffffff; font-size:24px; margin:0;">Session Complete</h1>
                  </td>
                </tr>
                <tr>
                  <td style="padding:24px; color:#333333; line-height:1.5;">
                    <p style="margin-top:0;">Hi <strong>${coachee}</strong>,</p>
                    <p>Your coaching session held on <strong>${sessionOn}</strong> has been marked as complete.</p>
                    <p style="margin:24px 0;">
                      <a href="${ackUrl}" target="_blank" style="display:inline-block; padding:10px 20px; background:#4e73df; color:#fff; text-decoration:none; border-radius:4px; font-weight:bold;">Acknowledge Coaching</a>
                    </p>
                    <hr style="border:none; border-top:1px solid #e1e5ea; margin:24px 0;">
                    <p style="margin:0;">Feel free to reply to this email if you have any questions or need further support.</p>
                    <p style="margin-top:24px;">Thanks,<br><em>${coachName}</em></p>
                  </td>
                </tr>
                <tr>
                  <td style="background:#f4f6f8; text-align:center; padding:12px; font-size:12px; color:#777777;">
                    &copy; ${new Date().getFullYear()} Your Company Name. All rights reserved.
                  </td>
                </tr>
              </table>
            </div>`
        });

        return;
      }
    }
    throw new Error('Coaching record not found: ' + id);
  } catch (error) {
    if (typeof safeWriteError === 'function') safeWriteError('completeCoachingAndNotify', error);
    throw error;
  }
}

/**
 * Saves the acknowledgement into both CoachingRecords and Quality sheets.
 *
 * @param {string} id       Coaching-record ID
 * @param {string} ackHtml  The rich-text HTML acknowledgement
 * @return {string} ISO timestamp when saved
 */
function acknowledgeCoaching(id, ackHtml) {
  try {
    const ss = getIndependenceSpreadsheet_();
    const crSh = ss.getSheetByName(typeof INDEPENDENCE_COACHING_SHEET === 'string' ? INDEPENDENCE_COACHING_SHEET : 'IndependenceCoaching');
    if (!crSh) throw new Error('Missing Coaching sheet: ' + INDEPENDENCE_COACHING_SHEET);

    const crData = crSh.getDataRange().getValues();
    if (crData.length < 2) throw new Error('Empty Coaching sheet');

    const crHdr = crData.shift();
    const crIdCol      = crHdr.indexOf('ID');
    const crAckTextCol = crHdr.indexOf('AcknowledgementText');
    const crAckOnCol   = crHdr.indexOf('AcknowledgedOn');
    const crUpdCol     = crHdr.indexOf('UpdatedAt');
    const crQaIdCol    = crHdr.indexOf('QAId');

    if (crIdCol < 0 || crAckTextCol < 0 || crAckOnCol < 0 || crQaIdCol < 0) {
      throw new Error('Missing one of: ID, AcknowledgementText, AcknowledgedOn, QAId in Coaching sheet');
    }

    const now = new Date().toISOString();
    let qaId = '';

    for (let i = 0; i < crData.length; i++) {
      if (String(crData[i][crIdCol]) === String(id)) {
        const rowIdx = i + 2; // account for header
        crSh.getRange(rowIdx, crAckTextCol + 1).setValue(ackHtml);
        crSh.getRange(rowIdx, crAckOnCol   + 1).setValue(now);
        crSh.getRange(rowIdx, crUpdCol     + 1).setValue(now);
        qaId = String(crData[i][crQaIdCol] || '');
        break;
      }
    }
    if (!qaId) throw new Error('Coaching record not found or missing QAId: ' + id);

    // Update QA sheet
    const qaSheetName = (typeof QA_RECORDS === 'string' && QA_RECORDS) ? QA_RECORDS : 'IndependenceQA';
    const qaSh = ss.getSheetByName(qaSheetName);
    if (!qaSh) return now; // silently succeed for coaching side

    const qaData = qaSh.getDataRange().getValues();
    if (qaData.length < 2) return now;
    const qaHdr = qaData.shift();

    const qaIdCol       = qaHdr.indexOf('ID');
    const feedbackOnCol = qaHdr.indexOf('FeedbackShared');
    const agentFbCol    = qaHdr.indexOf('AgentFeedback'); // Ensure your QA sheet has this column

    if (qaIdCol < 0 || feedbackOnCol < 0 || agentFbCol < 0) {
      if (typeof safeWriteError === 'function') {
        safeWriteError('acknowledgeCoaching', new Error('Missing ID/FeedbackShared/AgentFeedback in QA sheet'));
      }
      return now;
    }

    for (let j = 0; j < qaData.length; j++) {
      if (String(qaData[j][qaIdCol]) === String(qaId)) {
        const qaRow = j + 2;
        qaSh.getRange(qaRow, feedbackOnCol + 1).setValue(now);
        qaSh.getRange(qaRow, agentFbCol    + 1).setValue(ackHtml);
        break;
      }
    }

    return now;
  } catch (error) {
    if (typeof safeWriteError === 'function') safeWriteError('acknowledgeCoaching', error);
    throw error;
  }
}
