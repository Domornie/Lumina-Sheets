/**
 * EscalationService.gs
 *
 * Service functions to read, create, update, and delete escalation records.
 * fetchAllEscalations()
 *
 * Reads the entire Escalations sheet and returns an array of objects:
 *   [
 *     {
 *       id:        string,
 *       timestamp: Date,
 *       user:      string,
 *       type:      string,       // "Client" or "Supervisor"
 *       notes:     string,
 *       createdAt: Date,
 *       updatedAt: Date
 *     },
 *     …
 *   ]
 */
/**
 * EscalationService.gs
 *
 * CRUD operations on the “Escalations” sheet, using
 * the constants and helpers defined in Utilities.gs.
 */

/**
 * Returns all escalations in a form your client-side expects:
 *   { id, user, type, timestamp, notes }
 */
function fetchAllEscalations() {
  const tz = Session.getScriptTimeZone();
  const rows = readSheet(ESCALATIONS_SHEET);   // from Utilities.gs
  return rows.map(r => ({
    id: r.ID,
    user: r.User,
    type: r.Type,
    timestamp: r.Timestamp instanceof Date
      ? Utilities.formatDate(r.Timestamp, tz, 'yyyy-MM-dd HH:mm:ss')
      : r.Timestamp,
    notes: r.Notes
  }));
}

/**
 * Inserts a new escalation record.
 * @param {{user:string, type:string, timestamp:string, notes:string}} rec
 * @return {string} the generated ID
 */
function addEscalation(rec) {
  const sheet = getIBTRSpreadsheet()
    .getSheetByName(ESCALATIONS_SHEET);
  const id = Utilities.getUuid();
  const now = new Date();
  const tsDate = rec.timestamp
    ? new Date(rec.timestamp.replace('T', ' '))
    : now;
  // ESCALATIONS_HEADERS = ["ID","Timestamp","User","Type","Notes","CreatedAt","UpdatedAt"]
  sheet.appendRow([
    id,
    tsDate,
    rec.user,
    rec.type,
    rec.notes,
    now,
    now
  ]);
  return id;
}

/**
 * Updates an existing escalation by ID.
 * @param {string} id
 * @param {{user:string, type:string, timestamp:string, notes:string}} rec
 */
function updateEscalation(id, rec) {
  const sheet = getIBTRSpreadsheet()
    .getSheetByName(ESCALATIONS_SHEET);
  if (!sheet) {
    throw new Error('Escalations sheet not found.');
  }
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const row = i + 1;
      const createdAt = data[i][5];  // preserve original CreatedAt
      const tsDate = rec.timestamp
        ? new Date(rec.timestamp.replace('T', ' '))
        : new Date();
      // overwrite Timestamp, User, Type, Notes, keep CreatedAt, set UpdatedAt
      sheet.getRange(row, 2, 1, 6).setValues([[
        tsDate,
        rec.user,
        rec.type,
        rec.notes,
        createdAt,
        now
      ]]);
      return;
    }
  }
  throw new Error(`Escalation with ID "${id}" not found.`);
}

/**
 * Deletes an escalation row by ID.
 * @param {string} id
 */
function removeEscalation(id) {
  const sheet = getIBTRSpreadsheet()
    .getSheetByName(ESCALATIONS_SHEET);
  if (!sheet) {
    throw new Error('Escalations sheet not found.');
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
  throw new Error(`Escalation with ID "${id}" not found.`);
}
