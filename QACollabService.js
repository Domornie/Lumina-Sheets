// Code.gs

// ─────────── Submit a new QA audit (standalone) ────────────────────────────────
function submitQA(auditObj) {
  const ss = getIBTRSpreadsheet();
  const sh = ss.getSheetByName(QA_COLLAB_RECORDS);
  const id = Utilities.getUuid();
  const ts = new Date().toISOString();

  const weights = {
    q1:3, q2:5, q3:7,  q4:10, q5:5, q19:5,
    q6:8, q7:8, q8:15, q9:9,
    q10:8, q11:6, q12:6, q13:7, q14:3,
    q15:10, q16:5, q17:4, q18:5
  };

  let earned = 0, applicable = 0;
  Object.keys(weights).forEach(k => {
    const ans = auditObj[k];
    if (ans && ans !== 'na') {
      applicable += weights[k];
      if (ans === 'yes') earned += weights[k];
    }
  });

  if (auditObj.q8 === 'no') {
    earned = 0;
  } else if (auditObj.q15 === 'no') {
    earned = Math.max(earned - (weights.q15/2), 0);
  }

  const percentage = applicable ? earned / applicable : 0;

  const row = QA_COLLAB_HEADERS.map(col => {
    switch(col) {
      case 'ID': return id;
      case 'Timestamp': return ts;
      case 'CallerName': return auditObj.callerName || '';
      case 'AgentName': return auditObj.agentName || '';
      case 'AgentEmail': return auditObj.agentEmail || '';
      case 'ClientName': return auditObj.clientName || '';
      case 'CallDate': return auditObj.callDate || '';
      case 'CaseNumber': return auditObj.caseNumber || '';
      case 'CallLink': return '';  // removed audio URL
      case 'AuditorName': return auditObj.auditorName || '';
      case 'AuditDate': return auditObj.auditDate || '';
      case 'FeedbackShared':return auditObj.feedbackShared ? 'Yes' : 'No';
      case 'TotalScore': return earned;
      case 'Percentage': return percentage;
      case 'Notes': return auditObj.notes || '';
      case 'AgentFeedback': return '';
      default:
        if (/^Q\d+$/.test(col)) {
          const v = auditObj[col.toLowerCase()];
          return !v || v==='na' ? 'N/A' : (v==='yes' ? 'Yes' : 'No');
        }
        return '';
    }
  });

  sh.appendRow(row);
  return { id, ts };
}

// ─────────── SaveQA(form): now just calls submitQA ───────────────────────────────
function SaveQA(form) {
  return submitQA(form);
}

/**
 * Returns all QA records as an array of objects.
 * Each object’s keys are the header names from the QA sheet.
 * Date cells (Timestamp, AuditDate, etc.) are formatted as ISO strings.
 */
function getAllCollabQA() {
  const ss = getIBTRSpreadsheet();
  const sh = ss.getSheetByName(QA_COLLAB_RECORDS);
  if (!sh) return [];

  // Read everything
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];  // no rows

  const headers = data.shift();    // first row = headers
  const timeZone = Session.getScriptTimeZone();

  return data.map(row => {
    const obj = {};
    row.forEach((cell, colIdx) => {
      const key = headers[colIdx];
      if (cell instanceof Date) {
        // ISO‐style date/time
        obj[key] = Utilities.formatDate(cell, timeZone, "yyyy-MM-dd'T'HH:mm:ss");
      } else {
        obj[key] = cell;
      }
    });
    return obj;
  });
}

/**
 * Deletes a QA record by ID.
 * @param {string} id
 */
function deleteCollabQARecord(id) {
  const ss = getIBTRSpreadsheet();
  const sh = ss.getSheetByName(QA_COLLAB_RECORDS);
  const vals = sh.getDataRange().getValues();
  const idCol = vals[0].indexOf('ID');
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][idCol]) === id) {
      sh.deleteRow(i + 1);
      return true;
    }
  }
  throw new Error('QA record not found: ' + id);
}

/**
 * Returns one QA record matching the given ID.
 * @param {string} id
 * @return {Object}
 */
function getCollabQARecordById(id) {
  const ss = getIBTRSpreadsheet();
  const sh = ss.getSheetByName(QA_COLLAB_RECORDS);
  const data = sh.getDataRange().getValues();
  const headers = data.shift();
  const idCol = headers.indexOf('ID');
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][idCol]) === id) {
      const row = data[i];
      const obj = {};
      row.forEach((cell, j) => {
        const key = headers[j];
        if (cell instanceof Date) {
          obj[key] = Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
          obj[key] = cell;
        }
      });
      return obj;
    }
  }
  throw new Error('QA record not found: ' + id);
}

/**
 * Updates a QA record matching the given ID with the provided data object.
 * @param {string} id
 * @param {Object} data  // keys: callerName, agentName, … q1…q19, overallFeedback, etc.
 */
function updateQARecord(id, data) {
  const ss = getIBTRSpreadsheet();
  const sh = ss.getSheetByName(QA_COLLAB_RECORDS);
  const vals = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol = headers.indexOf('ID');
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][idCol]) === id) {
      const newRow = headers.map(col => {
        if (col === 'ID') return id;
        if (col === 'Timestamp') return vals[i][headers.indexOf('Timestamp')];
        const key = col.charAt(0).toLowerCase() + col.slice(1);
        const v = data[key];
        if (v === 'na' || v == null) return 'N/A';
        if (typeof v === 'boolean') return v ? 'Yes' : 'No';
        return v;
      });
      sh.getRange(i + 1, 1, 1, newRow.length).setValues([newRow]);
      return true;
    }
  }
  throw new Error('QA record not found: ' + id);
}

/**
 * filterQAByPeriodAndAgent(granularity, period, agentFilter)
 *   • granularity “Week”|“Month”|“Quarter”|“Year”
 *   • period:   e.g. “2025-W24”, “2025-06”, “Q2-2025”, “2025”
 *   • agentFilter: exact AgentName or empty for all
 * Returns only those QA rows whose CallDate falls in the period, and matches agent.
 */
function filterQAByPeriodAndAgent(granularity, period, agentFilter) {
  // 1) grab everything
  const all = getAllQA();            // you already have this
  // 2) convert period → date range
  const { startDate, endDate } = isoPeriodToDateRange(granularity, period);
  // 3) filter by date + agent
  return all.filter(r => {
    const d = new Date(r.CallDate);
    const inRange = d >= startDate && d <= endDate;
    const matchesAgent = !agentFilter || r.AgentName === agentFilter;
    return inRange && matchesAgent;
  });
}

/**
 * computeCategoryMetrics(records)
 *   Given an array of QA record objects,
 *   returns an object mapping each category to { avgScore, passPct }.
 */
function computeCategoryMetrics(records) {
  const weights = {
    q1:3, q2:5, q3:7,  q4:10, q5:5, q19:5,
    q6:8, q7:8, q8:15, q9:9,
    q10:8, q11:6, q12:6, q13:7, q14:3,
    q15:10, q16:5, q17:4, q18:5
  };
  const categories = {
    'Courtesy & Communication': ['q1','q2','q3','q4','q5','q19'],
    'Resolution':               ['q6','q7','q8','q9'],
    'Documentation':            ['q10','q11','q12','q13','q14'],
    'Compliance':               ['q15','q16','q17','q18']
  };

  const out = {};
  Object.entries(categories).forEach(([cat, keys]) => {
    const scores = [], passes = [];
    records.forEach(r => {
      let earned = 0, total = 0, allYes = true;
      keys.forEach(k => {
        total += weights[k];
        if ((r[k]||'').toString().toLowerCase() === 'yes') {
          earned += weights[k];
        } else {
          allYes = false;
        }
      });
      scores.push(total ? Math.round(earned/total*100) : 0);
      passes.push(allYes ? 1 : 0);
    });
    out[cat] = {
      avgScore: scores.length
        ? Math.round(scores.reduce((a,b)=>a+b,0)/scores.length)
        : 0,
      passPct: passes.length
        ? Math.round(passes.reduce((a,b)=>a+b,0)/passes.length*100)
        : 0
    };
  });
  return out;
}
