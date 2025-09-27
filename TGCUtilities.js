/***************************************************************
 * TGCUtilities.gs  (CHAT EDITION)
 * Infra-only: config, health checks, Drive/Sheet helpers, users
 ***************************************************************/
var TGC = this.TGC || {};
TGC.Util = (function () {
  /* --------------------- CONFIG --------------------- */
  var CFG = {
    ROOT_FOLDER_ID: '1filZWMVJdMPATBtq5wabAfTgxAVr4I4X', // optional for PDFs etc.
    QA_SPREADSHEET_ID: '1_bH68Z6Oj54rwdDxKvjCFGG1QUsp4X4DI53zxwGBmm0',               // separate spreadsheet for CHAT QA
    QA_SHEET_NAME: 'Chat_QA_Records',
    QA_HEADERS: [
      'ID','Timestamp',
      'CustomerName','AgentName','AgentEmail',
      'Channel','ConversationId','ChatDate',
      'AuditorName','AuditDate','ZeroTolerance',
      'Q1','Q2','Q3','Q4','Q5','Q6','Q7','Q8','Q9',
      'TotalEarned','TotalPossible','Percentage','PassStatus',
      'Cat_Professionalism','Cat_Comprehension','Cat_Resolution','Cat_FollowUp','Cat_Writing','Cat_Process',
      // Notes
      'NotesHtml'
    ]
  };

  function setConfig(partial) { Object.assign(CFG, partial || {}); }
  function getConfig() { return Object.assign({}, CFG); }

  /* --------------------- CORE HELPERS --------------------- */
  function tz() { return Session.getScriptTimeZone() || 'America/Jamaica'; }
  function todayYMD() { return Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd'); }
  function safe(v) { return v == null ? '' : String(v); }

  function getSpreadsheet() {
    if (!CFG.QA_SPREADSHEET_ID || CFG.QA_SPREADSHEET_ID.startsWith('REPLACE')) {
      throw new Error('TGC.Util: QA_SPREADSHEET_ID is not configured.');
    }
    return SpreadsheetApp.openById(CFG.QA_SPREADSHEET_ID);
  }

  function ensureSheetWithHeaders(ss, sheetName, headers) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh) sh = ss.insertSheet(sheetName);

    var lastRow = sh.getLastRow();
    if (lastRow === 0) {
      sh.appendRow(headers);
      return sh;
    }

    // verify header row; if mismatch, rewrite row 1 to headers
    var lastCol = sh.getLastColumn() || headers.length;
    var existing = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var sameLen = (existing.length === headers.length);
    var same = sameLen && existing.every(function (c, i) { return c === headers[i]; });

    if (!same) {
      sh.clear();
      sh.appendRow(headers);
    }
    return sh;
  }

  function writeError(where, err) {
    try {
      var ss = getSpreadsheet();
      var sh = ss.getSheetByName('_Logs') || ss.insertSheet('_Logs');
      sh.appendRow([
        new Date(),
        where,
        (err && err.message) ? err.message : String(err),
        (err && err.stack) ? err.stack : ''
      ]);
    } catch (e) {
      console.error('TGC.Util.writeError failed:', e);
    }
  }

  /* --------------------- USER PANEL --------------------- */
  /**
   * Reads the Users sheet (Name, Email, Role ...).
   * Returns [{name, email, role?, ...}]
   */
  function getUsersFromPanel() {
    try {
      var ss = getSpreadsheet();
      var sh = ss.getSheetByName(CFG.USERS_SHEET_NAME);
      if (!sh) return [];
      var vals = sh.getDataRange().getValues();
      if (vals.length < 2) return [];
      var headers = vals.shift();
      var idxName = headers.indexOf('Name');
      var idxEmail = headers.indexOf('Email');
      var idxRole = headers.indexOf('Role');
      return vals
        .filter(function (r) { return (r[idxName] || r[idxEmail]); })
        .map(function (r) {
          return {
            name: safe(r[idxName]),
            email: safe(r[idxEmail]),
            role: idxRole >= 0 ? safe(r[idxRole]) : ''
          };
        });
    } catch (err) {
      writeError('TGC.Util.getUsersFromPanel', err);
      return [];
    }
  }

  /* --------------------- HEALTH --------------------- */
  function ensureQAInfra() {
    var ok = { spreadsheet: false, sheet: false, users: false };
    var issues = [];

    try {
      var ss = getSpreadsheet();
      ok.spreadsheet = true;
      var sh = ensureSheetWithHeaders(ss, CFG.QA_SHEET_NAME, CFG.QA_HEADERS);
      ok.sheet = !!sh;

      var u = ss.getSheetByName(CFG.USERS_SHEET_NAME);
      ok.users = !!u;
      if (!ok.users) issues.push('Users sheet not found: ' + CFG.USERS_SHEET_NAME);
    } catch (e) {
      issues.push('Spreadsheet/Sheet: ' + (e.message || e));
    }

    return { ok: ok, issues: issues };
  }

  return {
    // config
    setConfig: setConfig,
    getConfig: getConfig,

    // helpers
    tz: tz,
    todayYMD: todayYMD,
    safe: safe,
    getSpreadsheet: getSpreadsheet,
    ensureSheetWithHeaders: ensureSheetWithHeaders,
    writeError: writeError,

    // users + health
    getUsersFromPanel: getUsersFromPanel,
    ensureQAInfra: ensureQAInfra
  };
})();

/* ----------------- Lightweight wrappers (for client) ----------------- */
function TGC_getUsers()      { return TGC.Util.getUsersFromPanel(); }
function TGC_healthCheck()   { return TGC.Util.ensureQAInfra(); }
