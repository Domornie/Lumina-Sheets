

/** @returns {Array<Object>} each object has keys from COACHING_HEADERS */
function getAllCoaching() {
  setupSheets();
  const sh   = getIBTRSpreadsheet().getSheetByName(COACHING_SHEET);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data.shift();
  const tz = Session.getScriptTimeZone();

  return data.map(row => {
    const obj = {};
    row.forEach((cell, i) => {
      const key = headers[i];
      // Format date fields uniformly
      if (
        ['SessionDate','FollowUpDate','CreatedAt','UpdatedAt']
          .includes(key) &&
        cell instanceof Date
      ) {
        obj[key] = Utilities.formatDate(cell, tz, 'yyyy-MM-dd');
      } else {
        obj[key] = cell;
      }
    });
    return obj;
  });
}

// ─────────── 4) Get one record by ID ───────────
/** @param {string} id */
function getCoachingRecordById(id) {
  const all = getAllCoaching();
  const rec = all.find(r => r.ID === id);
  if (!rec) throw new Error('Coaching record not found: ' + id);
  return rec;
}

/**
 * Returns an array of {id, date, percentage, agentName}
 * for populating the QA dropdown.
 */
function getQAItems() {
  const ss = getIBTRSpreadsheet();
  const sh = ss.getSheetByName(QA_RECORDS);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data.shift();

  // find the columns
  const idIdx = headers.indexOf('ID');
  const dateIdx = headers.indexOf('CallDate');
  const pctIdx = headers.indexOf('Percentage');
  const agIdx = headers.indexOf('AgentName');
  const emailIdx = headers.indexOf('AgentEmail');  // ← new

  const tz = Session.getScriptTimeZone();

  return data.map(row => {
    // format the date
    const rawDate = row[dateIdx];
    const dateStr = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, tz, 'yyyy-MM-dd')
      : String(rawDate);

    // coerce the percentage
    let pct = row[pctIdx];
    if (typeof pct !== 'number') {
      pct = parseFloat(pct) || 0;
    }

    return {
      id: row[idIdx],
      date: dateStr,
      percentage: pct,
      agentName: row[agIdx],
      agentEmail: emailIdx >= 0 ? row[emailIdx] : ''  // ← new
    };
  });
}

/**
 * Compute metrics for coaching dashboard.
 */
function getDashboardCoaching(granularity, period, agentFilter) {
  const all = getAllCoaching();
  const { startDate, endDate } = isoPeriodToDateRange(granularity, period);
  const tz = Session.getScriptTimeZone();

  // 1) Filter
  const filtered = all.filter(r => {
    const d = new Date(r.SessionDate + 'T00:00:00');
    return d >= startDate && d <= endDate
      && (!agentFilter || r.AgentName === agentFilter);
  });

  // 2) Trend buckets
  const buckets = {};
  filtered.forEach(r => {
    const d = new Date(r.SessionDate);
    let label;
    if (granularity === 'Week')      label = weekStringFromDate(d);
    else if (granularity === 'Month') label = Utilities.formatDate(d, tz, 'yyyy-MM');
    else if (granularity === 'Quarter') {
      const q = Math.floor(d.getMonth()/3) + 1;
      label = `Q${q}-${d.getFullYear()}`;
    } else                            label = `${d.getFullYear()}`;
    buckets[label] = (buckets[label]||0) + 1;
  });
  const sessionsTrend = Object.entries(buckets)
    .sort((a,b) => a[0]<b[0] ? -1 : 1)
    .map(([periodLabel,count]) => ({ periodLabel, count }));

  // 3) Topics distribution
  const topicCounts = {};
  filtered.forEach(r => {
    (r.TopicsCovered||'').split(',').forEach(t => {
      t = t.trim(); if (!t) return;
      topicCounts[t] = (topicCounts[t]||0)+1;
    });
  });
  const topicsDist = Object.entries(topicCounts)
    .map(([topic,count]) => ({ topic, count }));

  // 4) Upcoming follow-ups
  const today = new Date();
  const upcoming = all
    .filter(r => {
      const f = new Date(r.FollowUpDate + 'T00:00:00');
      return f >= today && (!agentFilter || r.AgentName === agentFilter);
    })
    .sort((a,b) => new Date(a.FollowUpDate) - new Date(b.FollowUpDate))
    .slice(0,20);

  return { sessionsTrend, topicsDist, upcoming };
}

/**
 * @param {{
 *   coachingDate:    string,  // 'YYYY-MM-DD'
 *   coachName:       string,
 *   employeeName:    string,
 *   coacheeEmail:    string,  // ← new
 *   topicsPlanned:   string,  // JSON string
 *   summary:         string,
 *   plan:            string,
 *   followUpDate:    string,
 *   followUpNotes:   string
 * }} sess
 */
function saveCoachingSession(sess) {
  setupSheets();
  const sh = getIBTRSpreadsheet()
    .getSheetByName(COACHING_SHEET);

  const now = new Date().toISOString();
  const id  = Utilities.getUuid();

  // build a map that matches COACHING_HEADERS
  const rowMap = {
    ID: id,
    QAId: sess.qaId,
    SessionDate: sess.coachingDate,
    AgentName: sess.coachName,
    CoacheeName: sess.employeeName,
    CoacheeEmail: sess.coacheeEmail,      // ← new
    TopicsPlanned: sess.topicsPlanned,
    CoveredTopics: JSON.stringify([]),
    ActionPlan: sess.plan,
    FollowUpDate: sess.followUpDate,
    Notes: sess.followUpNotes,
    CreatedAt: now,
    UpdatedAt: now
  };

  const row = COACHING_HEADERS.map(h => rowMap[h] || '');
  sh.appendRow(row);
  return { id, ts: now };
}

/**
 * @param {string} id
 * @param {string[]} coveredArray
 * @return {string} JSON-stringified coveredArray
 */
function updateCoveredTopics(id, coveredArray) {
  const ss = getIBTRSpreadsheet();
  const sh = ss.getSheetByName(COACHING_SHEET);
  const data = sh.getDataRange().getValues();
  const headers = data.shift();
  const idCol = headers.indexOf('ID');
  let covCol = headers.indexOf('CoveredTopics');
  const updCol = headers.indexOf('UpdatedAt');

  // If header missing, add it
  if (covCol < 0) {
    sh.getRange(1, headers.length+1).setValue('CoveredTopics');
    covCol = headers.length;
  }

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][idCol]) === id) {
      const rowIdx  = i + 2;
      const jsonStr = JSON.stringify(coveredArray);
      sh.getRange(rowIdx, covCol+1).setValue(jsonStr);
      sh.getRange(rowIdx, updCol+1).setValue(new Date().toISOString());
      return jsonStr;
    }
  }
  throw new Error('Coaching record not found: ' + id);
}

/**
 * Marks the session complete, emails the coachee, and includes the ack-form link.
 */
function completeCoachingAndNotify(id) {
  setupSheets();

  // Build your Web App’s base URL and ack‐form link
  const baseUrl = ScriptApp.getService().getUrl();
  const ackUrl  = `${baseUrl}&page=ackform&id=${encodeURIComponent(id)}`;

  const ss      = getIBTRSpreadsheet();
  const sh      = ss.getSheetByName(COACHING_SHEET);
  const data    = sh.getDataRange().getValues();
  const headers = data.shift();

  // Find column indexes
  const idCol    = headers.indexOf('ID');
  const emailCol = headers.indexOf('CoacheeEmail');
  const nameCol  = headers.indexOf('CoacheeName');
  const coachCol = headers.indexOf('AgentName');
  const dateCol  = headers.indexOf('SessionDate');
  const updCol   = headers.indexOf('UpdatedAt');
  const compCol  = headers.indexOf('Completed');  // optional

  if (emailCol < 0) {
    throw new Error('Cannot send notification — no “CoacheeEmail” column in sheet.');
  }

  // Locate the right row
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][idCol]) === id) {
      const rowIdx    = i + 2;
      const recipient = data[i][emailCol];
      const coachee   = data[i][nameCol];
      const coachName = data[i][coachCol];
      const sessionOn = data[i][dateCol];

      // Optionally mark “Completed”
      if (compCol >= 0) {
        sh.getRange(rowIdx, compCol + 1).setValue('Yes');
      }
      // Update timestamp
      sh.getRange(rowIdx, updCol + 1)
        .setValue(new Date().toISOString());

      // Send the styled email, signing off with coachName
      MailApp.sendEmail({
        to: recipient,
        subject: `Coaching Session on ${sessionOn} Completed`,
        htmlBody: `
          <div style="
            font-family: Arial, sans-serif;
            background-color: #f4f6f8;
            padding: 20px;">
            <table width="100%" cellpadding="0" cellspacing="0" style="
                max-width:600px;
                margin:auto;
                background:#ffffff;
                border-radius:8px;
                overflow:hidden;
                box-shadow:0 2px 8px rgba(0,0,0,0.1);">
              <tr>
                <td style="background-color:#4e73df; padding:16px; text-align:center;">
                  <h1 style="color:#ffffff; font-size:24px; margin:0;">
                    Session Complete
                  </h1>
                </td>
              </tr>
              <tr>
                <td style="padding:24px; color:#333333; line-height:1.5;">
                  <p style="margin-top:0;">Hi <strong>${coachee}</strong>,</p>
                  <p>Your coaching session held on <strong>${sessionOn}</strong> has been marked as complete.</p>
                  <p style="margin:24px 0;">
                    <a href="${ackUrl}" target="_blank" style="
                      display:inline-block;
                      padding:10px 20px;
                      background:#4e73df;
                      color:#fff;
                      text-decoration:none;
                      border-radius:4px;
                      font-weight:bold;">
                      Acknowledge Coaching
                    </a>
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
          </div>
        `
      });

      return;  // finished
    }
  }

  throw new Error('Coaching record not found: ' + id);
}

/**
 * Saves the acknowledgement into both CoachingRecords and Quality sheets.
 *
 * @param {string} id       Coaching-record ID
 * @param {string} ackHtml  The rich-text HTML acknowledgement
 * @return {string} ISO timestamp when saved
 */
function acknowledgeCoaching(id, ackHtml) {
  setupSheets();
  const ss = getIBTRSpreadsheet();
  const crSh = ss.getSheetByName(COACHING_SHEET);
  const crData = crSh.getDataRange().getValues();
  const crHdr = crData.shift();

  // ── find CoachingRecords columns ──
  const crIdCol = crHdr.indexOf('ID');
  const crAckTextCol = crHdr.indexOf('AcknowledgementText');
  const crAckOnCol = crHdr.indexOf('AcknowledgedOn');
  const crUpdCol = crHdr.indexOf('UpdatedAt');
  const crQaIdCol = crHdr.indexOf('QAId');  // ensure you saved QAId when session was created

  if (crIdCol<0 || crAckTextCol<0 || crAckOnCol<0 || crQaIdCol<0) {
    throw new Error('Missing one of: ID, AcknowledgementText, AcknowledgedOn or QAId header in CoachingRecords');
  }

  const now = new Date().toISOString();
  let qaId;

  // ── update CoachingRecords and grab QAId ──
  for (let i = 0; i < crData.length; i++) {
    if (String(crData[i][crIdCol]) === id) {
      const rowIdx = i + 2;  // account for header
      crSh.getRange(rowIdx, crAckTextCol+1).setValue(ackHtml);
      crSh.getRange(rowIdx, crAckOnCol+1).setValue(now);
      crSh.getRange(rowIdx, crUpdCol+1).setValue(now);
      qaId = String(crData[i][crQaIdCol]);
      break;
    }
  }
  if (!qaId) {
    throw new Error('Coaching record not found or missing QAId: ' + id);
  }

  // ── now update Quality sheet ──
  const qaSh = ss.getSheetByName(QA_RECORDS);
  const qaData = qaSh.getDataRange().getValues();
  const qaHdr = qaData.shift();

  const qaIdCol = qaHdr.indexOf('ID');
  const feedbackOnCol = qaHdr.indexOf('FeedbackShared');
  const agentFbCol = qaHdr.indexOf('AgentFeedback');

  if (qaIdCol<0 || feedbackOnCol<0 || agentFbCol<0) {
    throw new Error('Missing one of: ID, FeedbackShared or AgentFeedback header in Quality sheet');
  }

  for (let j = 0; j < qaData.length; j++) {
    if (String(qaData[j][qaIdCol]) === qaId) {
      const qaRow = j + 2;
      qaSh.getRange(qaRow, feedbackOnCol+1).setValue(now);
      qaSh.getRange(qaRow, agentFbCol+1).setValue(ackHtml);
      break;
    }
  }

  return now;
}

// ═══════════════════════════════════════════════════════════════════════════════
// CLIENT-ACCESSIBLE FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Client-accessible function to send QA results email
 */
function clientSendQAResultsEmail(emailData) {
  return sendQAResultsEmail(emailData);
}

/**
 * Enhanced coaching session save function that links to QA record
 */
function clientSaveCoachingSession(sessionData) {
  try {
    // Save the coaching session
    const result = saveCoachingSession(sessionData);
    
    // Also update the QA record to mark that coaching was provided
    if (sessionData.qaId) {
      updateQACoachingStatus(sessionData.qaId, true);
    }
    
    return result;
  } catch (error) {
    console.error('Error saving coaching session:', error);
    throw error;
  }
}

/**
 * Update QA record to indicate coaching was provided
 */
function updateQACoachingStatus(qaId, coachingProvided) {
  try {
    const ss = getIBTRSpreadsheet();
    const sh = ss.getSheetByName(QA_RECORDS);
    const data = sh.getDataRange().getValues();
    const headers = data.shift();
    
    const idCol = headers.indexOf('ID');
    let coachingCol = headers.indexOf('CoachingProvided');
    
    // Add column if it doesn't exist
    if (coachingCol === -1) {
      coachingCol = headers.length;
      sh.getRange(1, coachingCol + 1).setValue('CoachingProvided');
    }
    
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][idCol]) === String(qaId)) {
        const rowIdx = i + 2;
        sh.getRange(rowIdx, coachingCol + 1).setValue(coachingProvided ? 'Yes' : 'No');
        break;
      }
    }
  } catch (error) {
    console.warn('Could not update coaching status:', error);
  }
}
