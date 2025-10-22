

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
    const normalizedPct = pct > 1 ? pct / 100 : pct;

    return {
      id: row[idIdx],
      date: dateStr,
      percentage: normalizedPct,
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
  const safeTopics = sanitizeTopicsPayload_(sess.topicsPlanned);
  const safeSummary = sanitizeRichHtml_(sess.summary);
  const safePlan = sanitizeRichHtml_(sess.plan);
  const safeNotes = sanitizeRichHtml_(sess.followUpNotes);
  const safeCoachName = sanitizeAiText_(sess.coachName);
  const safeEmployeeName = sanitizeAiText_(sess.employeeName);
  const safeCoacheeEmail = sanitizeAiText_(sess.coacheeEmail);
  const safeQaId = sanitizeAiText_(sess.qaId);

  // build a map that matches COACHING_HEADERS
  const rowMap = {
    ID: id,
    QAId: safeQaId,
    SessionDate: sess.coachingDate,
    AgentName: safeCoachName,
    CoacheeName: safeEmployeeName,
    CoacheeEmail: safeCoacheeEmail,      // ← new
    TopicsPlanned: safeTopics,
    CoveredTopics: JSON.stringify([]),
    Summary: safeSummary,
    ActionPlan: safePlan,
    FollowUpDate: sess.followUpDate,
    Notes: safeNotes,
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
  const safeCovered = Array.isArray(coveredArray)
    ? coveredArray.map(function (topic) { return sanitizeAiText_(topic); })
    : [];

  // If header missing, add it
  if (covCol < 0) {
    sh.getRange(1, headers.length+1).setValue('CoveredTopics');
    covCol = headers.length;
  }

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][idCol]) === id) {
      const rowIdx  = i + 2;
      const jsonStr = JSON.stringify(safeCovered);
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
        const value = coachingProvided ? 'Yes' : 'No';
        sh.getRange(rowIdx, coachingCol + 1).setValue(value);
        return value;
      }
    }
    return null;
  } catch (error) {
    console.warn('Could not update coaching status:', error);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// COACHING HUB INTELLIGENCE
// ═══════════════════════════════════════════════════════════════════════════════

const COACHING_AI_MOTIVATORS = [
  'Celebrate the wins, then turn them into repeatable habits for the next call.',
  'Confidence comes from clarity—highlight what went right before refining the next play.',
  'Every QA insight is a map. Coaching turns that map into a guided tour for the agent.',
  'Lumina Coaching Hub is ready—pair recognition with a focused drill to keep momentum.',
  'Quality signals spotted. Now transform them into customer-obsessed moments.',
  'Keep the conversation human. Blend empathy, precision, and policy fluency every time.'
];

const COACHING_AI_ETIQUETTE_TIPS = [
  'Use the customer’s name naturally at least twice to reinforce personal connection.',
  'Mirror the customer’s pace while staying calm—tempo control keeps conversations confident.',
  'Summarize the resolution path before closing to confirm alignment and next steps.',
  'Invite questions before ending the call to ensure no hidden concerns remain.',
  'Document commitments in real time so wrap-up notes reflect the live conversation.',
  'Reinforce policy moments with “because” statements—clarity reduces escalations.',
  'Thank the customer for their patience when hold time or research was required.',
  'Transition between topics using “first…next” language to keep structure crisp.'
];

const COACHING_AI_CELEBRATIONS = [
  'Spotlight this win in the next huddle—agents repeat what gets recognized.',
  'Quality Command Center flags this as a best practice worth sharing team-wide.',
  'Capture this moment in your playbook; it reinforces what outstanding sounds like.',
  '👏 Lumina AI tagged this interaction as a customer delight moment—keep the momentum!',
  'Performance Command Center logged this as a consistency milestone. Reinforce it today.'
];

const COACHING_AI_FOCUS_OPENERS = [
  'Coach with intent—focus here to close the experience gap quickly.',
  'Drill this scenario to turn risk into readiness before the next QA sample.',
  'Lean on role play and job aids to reinforce the muscle memory needed here.',
  'Pair the note below with a micro-learning clip to accelerate improvement.'
];

const COACHING_AI_KEYWORD_TOPICS = [
  {
    keywords: ['empathy', 'tone', 'courteous', 'courtesy', 'rapport'],
    name: 'Empathy Calibration',
    detail: 'Practice reflective language, empathetic acknowledgements, and confidence statements.'
  },
  {
    keywords: ['hold', 'follow-up', 'callback', 'delay'],
    name: 'Expectation Setting',
    detail: 'Coach on setting timelines, providing status updates, and confirming next steps before closing.'
  },
  {
    keywords: ['compliance', 'policy', 'verification', 'authentication'],
    name: 'Policy & Compliance Precision',
    detail: 'Refresh verification scripts and ensure regulatory disclosures happen without friction.'
  },
  {
    keywords: ['documentation', 'notes', 'wrap', 'after call'],
    name: 'Documentation Excellence',
    detail: 'Review live note-taking structure and reinforce the must-have disposition elements.'
  },
  {
    keywords: ['process', 'procedure', 'steps', 'workflow'],
    name: 'Process Adherence',
    detail: 'Map the journey step-by-step and rehearse decision points where agents hesitate.'
  }
];

const QA_COACHING_BLUEPRINT = {
  q1: {
    label: 'Warm Greeting & Identity Confirmation',
    category: 'Customer Courtesy',
    weight: 3,
    celebrate: 'Opened with a confident greeting that set a positive tone.',
    coach: 'Rehearse the opening script so brand, name, and assistance offer flow naturally.',
    etiquette: 'Smile through your voice and reference the customer name quickly.',
    topicName: 'Greeting Refresh',
    topicDetail: 'Role-play a 30-second opening that includes a warm greeting and verification prompt.'
  },
  q2: {
    label: 'Empathy & Ownership',
    category: 'Customer Courtesy',
    weight: 5,
    celebrate: 'Demonstrated empathy and took ownership for the customer experience.',
    coach: 'Use acknowledgement statements when customers share frustration to reinforce trust.',
    etiquette: 'Pair empathy with action—state what you will do right after you acknowledge feelings.',
    topicName: 'Empathy Ladder',
    topicDetail: 'Practice acknowledgement phrases that lead into confident problem statements.'
  },
  q3: {
    label: 'Active Listening & Probing',
    category: 'Customer Courtesy',
    weight: 7,
    celebrate: 'Kept the conversation customer-led with smart clarifying questions.',
    coach: 'Coach on layered probing so the agent captures root cause before solving.',
    etiquette: 'Use “just to make sure I’ve got this right” as a transition into the customer summary.',
    topicName: 'Discovery Skills',
    topicDetail: 'Drill layered probing and summarizing to confirm needs before solutioning.'
  },
  q4: {
    label: 'Call Control & Confidence',
    category: 'Customer Courtesy',
    weight: 10,
    celebrate: 'Maintained confident call control while keeping rapport intact.',
    coach: 'Reinforce signposting so tough conversations stay structured and efficient.',
    etiquette: 'Preview next steps before placing a customer on hold or changing topics.',
    topicName: 'Call Flow Mastery',
    topicDetail: 'Walk through the ideal call structure and practice confident transitions.'
  },
  q5: {
    label: 'Professional Tone & Language',
    category: 'Customer Courtesy',
    weight: 5,
    celebrate: 'Tone and language reflected polished, brand-aligned professionalism.',
    coach: 'Review filler words and ensure tone stays confident even under pressure.',
    etiquette: 'Swap casual filler for purposeful reassurance (“Absolutely, I can help with that…”).',
    topicName: 'Tone Calibration',
    topicDetail: 'Listen to call snippets and coach on tone shifts that reinforce credibility.'
  },
  q6: {
    label: 'Issue Diagnosis',
    category: 'Resolution',
    weight: 8,
    celebrate: 'Quickly pinpointed the core issue without the customer repeating themselves.',
    coach: 'Coach on diagnostic checklists to avoid missing prerequisite questions.',
    etiquette: 'Verbalize what you are checking so the customer knows progress is happening.',
    topicName: 'Root Cause Playbook',
    topicDetail: 'Review troubleshooting flows and practice pacing so discovery stays efficient.'
  },
  q7: {
    label: 'Solution Accuracy',
    category: 'Resolution',
    weight: 8,
    celebrate: 'Delivered the right solution path on the first attempt.',
    coach: 'Rebuild confidence with scenario-based practice on complex resolutions.',
    etiquette: 'Confirm the plan and recap why it solves the customer’s stated issue.',
    topicName: 'Solution Mapping',
    topicDetail: 'Map common scenarios and verify the agent can articulate the “why” behind each fix.'
  },
  q8: {
    label: 'Authentication & Security',
    category: 'Resolution',
    weight: 15,
    celebrate: 'Validated security perfectly before advancing the conversation.',
    coach: 'Reinforce authentication scripts and escalate paths for failed verification.',
    etiquette: 'Explain the “why” behind security steps to preserve trust during verification.',
    topicName: 'Security Protocol Drill',
    topicDetail: 'Rehearse verification workflows and contingency paths when data mismatches occur.'
  },
  q9: {
    label: 'Effective Follow-through',
    category: 'Resolution',
    weight: 9,
    celebrate: 'Outlined next steps clearly and confirmed customer agreement.',
    coach: 'Coach on setting expectations and capturing timelines to avoid repeat contacts.',
    etiquette: 'Confirm understanding by asking the customer to restate the agreed next step.',
    topicName: 'Expectation Setting',
    topicDetail: 'Practice closing scripts that cover recap, timeline, and ownership statements.'
  },
  q10: {
    label: 'Documentation Accuracy',
    category: 'Documentation',
    weight: 8,
    celebrate: 'Documented key actions with accuracy that supports downstream teams.',
    coach: 'Review note templates to ensure critical data points appear every time.',
    etiquette: 'Capture commitments as bullet-style notes while they are fresh.',
    topicName: 'Note-Taking Framework',
    topicDetail: 'Walk through disposition templates and audit for missing compliance cues.'
  },
  q11: {
    label: 'System Navigation',
    category: 'Documentation',
    weight: 6,
    celebrate: 'Navigated systems confidently without dead air or uncertainty.',
    coach: 'Run guided simulations to speed up navigation through tricky workflows.',
    etiquette: 'Narrate quietly what is happening during longer system loads.',
    topicName: 'System Drill-Down',
    topicDetail: 'Shadow navigation clicks and capture shortcuts that keep calls moving.'
  },
  q12: {
    label: 'Knowledge Base Utilization',
    category: 'Documentation',
    weight: 6,
    celebrate: 'Leveraged the knowledge base to validate policies on the fly.',
    coach: 'Coach on searching smarter and bookmarking the best articles for quick recall.',
    etiquette: 'Tell the customer when you are double-checking policy so they trust the answer.',
    topicName: 'Knowledge Search Mastery',
    topicDetail: 'Practice search queries and highlight go-to resources for complex questions.'
  },
  q13: {
    label: 'Compliance Notations',
    category: 'Documentation',
    weight: 7,
    celebrate: 'Captured compliance statements accurately in the case record.',
    coach: 'Revisit regulatory checklist items and ensure nothing is left implied.',
    etiquette: 'State required disclosures confidently and note them verbatim when needed.',
    topicName: 'Compliance Deep Dive',
    topicDetail: 'Review compliance scripts and document how to capture mandatory language.'
  },
  q14: {
    label: 'Knowledge Transfer',
    category: 'Documentation',
    weight: 3,
    celebrate: 'Provided context so the next teammate can pick up without friction.',
    coach: 'Coach on summarizing the conversation in two crisp sentences for hand-offs.',
    etiquette: 'Use consistent tags or categories that downstream teams expect.',
    topicName: 'Handoff Clarity',
    topicDetail: 'Practice writing closing summaries focused on hand-off readiness.'
  },
  q15: {
    label: 'Critical Process Adherence',
    category: 'Process',
    weight: 10,
    celebrate: 'Followed critical process steps flawlessly—zero risk flags.',
    coach: 'Run through the process map and highlight must-not-miss checkpoints.',
    etiquette: 'Announce when you are following a required step so customers understand the pause.',
    topicName: 'Process Calibration',
    topicDetail: 'Simulate the full process and identify where to slow down for accuracy.'
  },
  q16: {
    label: 'Tool & Resource Usage',
    category: 'Process',
    weight: 5,
    celebrate: 'Used the right tools without assistance, keeping the call efficient.',
    coach: 'Refresh where to find each resource so the agent is never searching live.',
    etiquette: 'Narrate briefly when jumping between systems so customers feel guided.',
    topicName: 'Tool Navigation Refresh',
    topicDetail: 'Create a quick-reference map of tools and when to use each.'
  },
  q17: {
    label: 'Escalation Judgment',
    category: 'Process',
    weight: 4,
    celebrate: 'Made a smart judgment on when to escalate versus own the solution.',
    coach: 'Clarify escalation triggers and self-service thresholds.',
    etiquette: 'Explain escalation paths to the customer to avoid surprises.',
    topicName: 'Escalation Playbook',
    topicDetail: 'Review decision trees for escalations and what to communicate at each stage.'
  },
  q18: {
    label: 'Policy Risk Avoidance',
    category: 'Process',
    weight: 5,
    celebrate: 'Protected the brand by respecting high-risk policy steps.',
    coach: 'Reinforce red-line policies and the wording required when exceptions appear.',
    etiquette: 'Anchor policy statements in customer value to keep trust intact.',
    topicName: 'Risk Guardrails',
    topicDetail: 'Revisit policy guardrails and practice messaging when denying requests.'
  },
  q19: {
    label: 'Customer Commitment Check',
    category: 'Customer Courtesy',
    weight: 5,
    celebrate: 'Confirmed satisfaction before ending, reinforcing trust.',
    coach: 'Coach on final check questions that invite lingering concerns.',
    etiquette: 'Ask “What else can I take off your plate today?” before closing.',
    topicName: 'Closing Confidence',
    topicDetail: 'Practice closing scripts that blend gratitude, recap, and next-step confirmation.'
  }
};

const QA_QUESTION_ORDER = Object.keys(QA_COACHING_BLUEPRINT);

function buildEmptyCoachingIntel_(qaId) {
  return {
    qaId: qaId || '',
    summary: 'Select a QA record to generate Lumina Coaching Hub intelligence.',
    agentName: '',
    percentage: null,
    scoreText: '—',
    motivation: COACHING_AI_MOTIVATORS[Math.floor(Math.random() * COACHING_AI_MOTIVATORS.length)],
    celebrations: [],
    focusAreas: [],
    etiquetteTips: COACHING_AI_ETIQUETTE_TIPS.slice(0, 3),
    acknowledgementPrompts: [],
    feedbackSignals: [],
    recommendedTopics: [],
    metadata: {}
  };
}

function getQARecordById_(qaId) {
  if (!qaId) return null;
  setupSheets();
  const ss = getIBTRSpreadsheet();
  const sh = ss.getSheetByName(QA_RECORDS);
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  if (!data.length) return null;
  const headers = data.shift();
  const idIdx = headers.indexOf('ID');
  if (idIdx === -1) return null;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(qaId)) {
      const row = {};
      headers.forEach((header, index) => {
        row[header] = data[i][index];
      });
      return row;
    }
  }
  return null;
}

function sanitizeAiText_(value) {
  if (!value) return '';
  return String(value)
    .replace(/\s+/g, ' ')
    .replace(/<[^>]+>/g, '')
    .trim();
}

function sanitizeRichHtml_(html) {
  if (!html) return '';
  return String(html)
    .replace(/<script[\s\S]*?>[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?>[\s\S]*?<\/style>/gi, '')
    .replace(/on[a-z]+\s*=\s*"[^"]*"/gi, '')
    .replace(/on[a-z]+\s*=\s*'[^']*'/gi, '')
    .replace(/on[a-z]+\s*=\s*[^\s>]+/gi, '')
    .replace(/javascript:/gi, '');
}

function sanitizeTopicsPayload_(topicsJson) {
  if (!topicsJson) return '[]';
  try {
    const parsed = JSON.parse(topicsJson);
    if (!Array.isArray(parsed)) {
      return '[]';
    }
    const cleaned = parsed.map(function (topic) {
      const name = topic && topic.name ? sanitizeAiText_(topic.name) : '';
      const detail = topic && topic.detail ? sanitizeRichHtml_(topic.detail) : '';
      return { name: name, detail: detail };
    });
    return JSON.stringify(cleaned);
  } catch (err) {
    console.warn('Unable to sanitize topics payload:', err);
    return '[]';
  }
}

function extractCoachingSignalsFromText_(text) {
  const clean = sanitizeAiText_(text);
  if (!clean) return [];
  const sentences = clean.split(/[.!?]+/).map(function (s) { return s.trim(); }).filter(Boolean);
  const keywords = ['coach', 'coaching', 'improve', 'improvement', 'focus', 'train', 'training', 'recommend', 'recommendation', 'follow up', 'follow-up'];
  return sentences.filter(function (sentence) {
    const lower = sentence.toLowerCase();
    return keywords.some(function (kw) { return lower.indexOf(kw) !== -1; });
  });
}

function findKeywordTopics_(text) {
  const clean = sanitizeAiText_(text).toLowerCase();
  if (!clean) return [];
  const matches = [];
  COACHING_AI_KEYWORD_TOPICS.forEach(function (topic) {
    if (topic.keywords.some(function (kw) { return clean.indexOf(kw) !== -1; })) {
      matches.push({ name: topic.name, detail: topic.detail });
    }
  });
  return matches;
}

function generateCoachingHubInsights(qaId) {
  const base = buildEmptyCoachingIntel_(qaId);
  const qaRecord = getQARecordById_(qaId);
  if (!qaRecord) {
    base.summary = 'No QA record found. Capture a QA evaluation to unlock AI coaching guidance.';
    return base;
  }

  const agentName = sanitizeAiText_(qaRecord.AgentName || qaRecord.Agent || '');
  const percentageRaw = qaRecord.Percentage || qaRecord.FinalScore || qaRecord['QA Score'];
  const percentage = typeof percentageRaw === 'number'
    ? percentageRaw
    : parseFloat(percentageRaw) || 0;
  const normalizedPct = percentage > 1 ? percentage / 100 : percentage;
  const scoreText = Math.round(normalizedPct * 100) + '%';

  base.agentName = agentName;
  base.percentage = normalizedPct;
  base.scoreText = scoreText;
  base.metadata = {
    client: sanitizeAiText_(qaRecord.ClientName || qaRecord.Client || ''),
    callDate: qaRecord.CallDate instanceof Date
      ? Utilities.formatDate(qaRecord.CallDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : sanitizeAiText_(qaRecord.CallDate),
    qaId: qaId
  };

  const performanceTone = normalizedPct >= 0.9
    ? 'is exceeding expectations.'
    : normalizedPct >= 0.8
      ? 'is meeting key expectations with a few targeted refinements needed.'
      : 'needs a focused coaching huddle to lift critical skills quickly.';

  base.summary = [
    'Lumina Coaching Hub analyzed QA Performance Command Center data for ',
    agentName || 'the selected agent',
    ' and found the latest evaluation scored at ',
    scoreText,
    ' which ',
    performanceTone
  ].join('');

  base.motivation = COACHING_AI_MOTIVATORS[Math.floor(Math.random() * COACHING_AI_MOTIVATORS.length)];

  const strengths = [];
  const focusAreas = [];
  const recommendedTopics = [];

  QA_QUESTION_ORDER.forEach(function (key) {
    const meta = QA_COACHING_BLUEPRINT[key];
    const col = key.replace('q', 'Q');
    const answerRaw = qaRecord[col];
    const answer = sanitizeAiText_(answerRaw).toLowerCase();
    const note = sanitizeAiText_(qaRecord[col + ' Note'] || qaRecord['C' + key.replace('q', '')]);

    if (!answer || answer === 'n/a' || answer === 'na') {
      return;
    }

    if (answer === 'yes' || (key === 'q17' && answer === 'no')) {
      strengths.push({
        title: meta.label,
        detail: note || meta.celebrate,
        category: meta.category,
        weight: meta.weight
      });
    } else {
      focusAreas.push({
        title: meta.label,
        detail: note || meta.coach,
        category: meta.category,
        weight: meta.weight
      });
      recommendedTopics.push({ name: meta.topicName, detail: meta.topicDetail });
    }
  });

  strengths.sort(function (a, b) { return b.weight - a.weight; });
  focusAreas.sort(function (a, b) { return b.weight - a.weight; });

  base.celebrations = strengths.slice(0, 4).map(function (item) {
    const meta = QA_QUESTION_ORDER
      .map(function (key) { return QA_COACHING_BLUEPRINT[key]; })
      .find(function (entry) { return entry && entry.label === item.title; });
    return {
      title: item.title,
      detail: item.detail || (meta ? meta.celebrate : ''),
      category: item.category,
      callout: COACHING_AI_CELEBRATIONS[Math.floor(Math.random() * COACHING_AI_CELEBRATIONS.length)]
    };
  });

  base.focusAreas = focusAreas.slice(0, 6).map(function (item) {
    const meta = QA_QUESTION_ORDER
      .map(function (key) { return QA_COACHING_BLUEPRINT[key]; })
      .find(function (entry) { return entry && entry.label === item.title; });
    return {
      title: item.title,
      detail: item.detail || (meta ? meta.coach : ''),
      category: item.category,
      callout: COACHING_AI_FOCUS_OPENERS[Math.floor(Math.random() * COACHING_AI_FOCUS_OPENERS.length)]
    };
  });

  const dedupedTopics = [];
  const seenTopicNames = {};
  recommendedTopics.forEach(function (topic) {
    if (!topic.name) return;
    const key = topic.name.toLowerCase();
    if (!seenTopicNames[key]) {
      seenTopicNames[key] = true;
      dedupedTopics.push(topic);
    }
  });
  base.recommendedTopics = dedupedTopics.slice(0, 6);

  const etiquetteSet = new Set();
  base.focusAreas.forEach(function (area) {
    const meta = QA_COACHING_BLUEPRINT[QA_QUESTION_ORDER.find(function (q) {
      return QA_COACHING_BLUEPRINT[q].label === area.title;
    })];
    if (meta && meta.etiquette) {
      etiquetteSet.add(meta.etiquette);
    }
  });
  while (etiquetteSet.size < 3) {
    etiquetteSet.add(COACHING_AI_ETIQUETTE_TIPS[Math.floor(Math.random() * COACHING_AI_ETIQUETTE_TIPS.length)]);
  }
  base.etiquetteTips = Array.from(etiquetteSet).slice(0, 5);

  const feedbackSignals = [];
  const overallFeedback = qaRecord.OverallFeedback || qaRecord.Recommendations;
  const qaNotes = qaRecord.Notes || qaRecord.AgentFeedback;
  [overallFeedback, qaNotes].forEach(function (text) {
    extractCoachingSignalsFromText_(text).forEach(function (sentence) {
      feedbackSignals.push(sentence);
    });
  });
  base.feedbackSignals = feedbackSignals.slice(0, 5);

  const keywordTopics = [];
  [overallFeedback, qaNotes].forEach(function (text) {
    findKeywordTopics_(text).forEach(function (topic) {
      keywordTopics.push(topic);
    });
  });
  keywordTopics.forEach(function (topic) {
    const key = topic.name.toLowerCase();
    if (!seenTopicNames[key]) {
      seenTopicNames[key] = true;
      dedupedTopics.push(topic);
    }
  });
  base.recommendedTopics = dedupedTopics.slice(0, 8);

  const ackPrompts = [];
  if (base.focusAreas.length) {
    ackPrompts.push('Highlight one celebration, then coach through the first focus area using the plan below.');
    ackPrompts.push('Ask the agent to summarize their next action in their own words before ending the session.');
  } else {
    ackPrompts.push('Celebrate the consistency and capture the playbook steps in a quick Loom or knowledge base update.');
  }
  ackPrompts.push('Schedule the follow-up in the Coaching Hub now so cadence stays predictable.');
  base.acknowledgementPrompts = ackPrompts;

  return base;
}

function clientGetCoachingHubInsights(qaId) {
  return generateCoachingHubInsights(qaId);
}
