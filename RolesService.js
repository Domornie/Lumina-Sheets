// ────────────────────────────────────────────────────────────────────────────
// Role & UserRole CRUD with SheetsDB awareness
// ────────────────────────────────────────────────────────────────────────────

const ROLE_SCOPE_DEFAULT = 'global';
const ROLE_SCOPE_ALLOWED = ['global', 'campaign', 'team'];

function hasDatabaseManager() {
  return typeof DatabaseManager !== 'undefined' && DatabaseManager && typeof DatabaseManager.defineTable === 'function';
}

function normalizeStringValue(value) {
  if (value === null || typeof value === 'undefined') return '';
  return String(value).trim();
}

function normalizeScopeValue(scope) {
  const candidate = normalizeStringValue(scope).toLowerCase();
  if (!candidate) return ROLE_SCOPE_DEFAULT;
  return ROLE_SCOPE_ALLOWED.indexOf(candidate) !== -1 ? candidate : ROLE_SCOPE_DEFAULT;
}

function isSoftDeletedRow(row) {
  if (!row) return false;
  const marker = row.DeletedAt || row.deletedAt || row.deleted_at;
  if (marker === null || typeof marker === 'undefined') return false;
  const str = String(marker).trim();
  return !!str;
}

function getRolesTable() {
  return DatabaseManager.defineTable(ROLES_SHEET || 'Roles', {
    headers: Array.isArray(ROLES_HEADER) && ROLES_HEADER.length
      ? ROLES_HEADER
      : ['ID', 'Name', 'NormalizedName', 'Scope', 'Description', 'CreatedAt', 'UpdatedAt', 'DeletedAt'],
    idColumn: 'ID',
    defaults: {
      Scope: ROLE_SCOPE_DEFAULT,
      Description: '',
      DeletedAt: ''
    }
  });
}

function getUserRolesTable() {
  return DatabaseManager.defineTable(USER_ROLES_SHEET || 'UserRoles', {
    headers: Array.isArray(USER_ROLES_HEADER) && USER_ROLES_HEADER.length
      ? USER_ROLES_HEADER
      : ['ID', 'UserId', 'RoleId', 'Scope', 'AssignedBy', 'CreatedAt', 'UpdatedAt', 'DeletedAt'],
    idColumn: 'ID',
    defaults: {
      Scope: ROLE_SCOPE_DEFAULT,
      AssignedBy: '',
      DeletedAt: ''
    }
  });
}

function legacyEnsureSheet(name, headers) {
  if (typeof ensureSheetWithHeaders === 'function') {
    return ensureSheetWithHeaders(name, headers);
  }
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function legacyGetAllRoles() {
  const sheet = legacyEnsureSheet(ROLES_SHEET, ROLES_HEADER);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const idIdx = headers.indexOf('ID');
  const nameIdx = headers.indexOf('Name');
  const normalizedIdx = headers.indexOf('NormalizedName');
  const scopeIdx = headers.indexOf('Scope');
  const descriptionIdx = headers.indexOf('Description');
  const deletedIdx = headers.indexOf('DeletedAt');
  return data.slice(1)
    .filter(row => !normalizeStringValue(deletedIdx !== -1 ? row[deletedIdx] : ''))
    .map(row => ({
      id: normalizeStringValue(idIdx !== -1 ? row[idIdx] : row[0]),
      name: nameIdx !== -1 ? (row[nameIdx] || '') : (row[1] || ''),
      normalizedName: normalizedIdx !== -1 ? (row[normalizedIdx] || '') : (row[2] || ''),
      scope: normalizeScopeValue(scopeIdx !== -1 ? row[scopeIdx] : ROLE_SCOPE_DEFAULT),
      description: descriptionIdx !== -1 ? (row[descriptionIdx] || '') : ''
    }));
}

function getAllRoles() {
  if (hasDatabaseManager()) {
    try {
      const table = getRolesTable();
      const rows = table.find({ filter: function (row) { return !isSoftDeletedRow(row); } }) || [];
      return rows.map(function (row) {
        return {
          id: normalizeStringValue(row.ID || row.Id),
          name: row.Name || '',
          normalizedName: row.NormalizedName || '',
          scope: normalizeScopeValue(row.Scope),
          description: row.Description || ''
        };
      });
    } catch (err) {
      if (typeof safeWriteError === 'function') safeWriteError('getAllRolesDb', err);
    }
  }
  return legacyGetAllRoles();
}

function legacyAddRole(name, options) {
  const sheet = legacyEnsureSheet(ROLES_SHEET, ROLES_HEADER);
  const id = Utilities.getUuid();
  const now = new Date();
  const trimmedName = (name || '').trim();
  const normalizedName = trimmedName.toUpperCase();
  const scope = normalizeScopeValue(options && options.scope);
  const description = (options && options.description) ? String(options.description) : '';
  const row = [
    id,
    trimmedName,
    normalizedName,
    scope,
    description,
    now,
    now,
    ''
  ];
  sheet.appendRow(row);
  return id;
}

function addRole(name, options) {
  const trimmed = (name || '').trim();
  if (!trimmed) {
    throw new Error('Role name is required');
  }
  const settings = options || {};
  if (hasDatabaseManager()) {
    try {
      const table = getRolesTable();
      const record = table.insert({
        Name: trimmed,
        NormalizedName: trimmed.toUpperCase(),
        Scope: normalizeScopeValue(settings.scope),
        Description: settings.description || '',
        DeletedAt: ''
      });
      return record.ID;
    } catch (err) {
      if (typeof safeWriteError === 'function') safeWriteError('addRoleDb', err);
    }
  }
  return legacyAddRole(trimmed, settings);
}

function legacyUpdateRole(id, updates) {
  const sheet = legacyEnsureSheet(ROLES_SHEET, ROLES_HEADER);
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const headers = data[0];
  const nameIdx = headers.indexOf('Name');
  const normalizedIdx = headers.indexOf('NormalizedName');
  const scopeIdx = headers.indexOf('Scope');
  const descriptionIdx = headers.indexOf('Description');
  const updatedIdx = headers.indexOf('UpdatedAt');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      const rowIndex = i + 1;
      if (updates.name) {
        if (nameIdx !== -1) sheet.getRange(rowIndex, nameIdx + 1).setValue(updates.name);
        if (normalizedIdx !== -1) sheet.getRange(rowIndex, normalizedIdx + 1).setValue(updates.name.toUpperCase());
      }
      if (updates.scope && scopeIdx !== -1) sheet.getRange(rowIndex, scopeIdx + 1).setValue(normalizeScopeValue(updates.scope));
      if (Object.prototype.hasOwnProperty.call(updates, 'description') && descriptionIdx !== -1) {
        sheet.getRange(rowIndex, descriptionIdx + 1).setValue(updates.description || '');
      }
      if (updatedIdx !== -1) sheet.getRange(rowIndex, updatedIdx + 1).setValue(now);
      break;
    }
  }
}

function updateRole(id, newName, options) {
  if (!id) return;
  const updates = { name: (newName || '').trim() };
  const settings = options || {};
  if (settings.scope) updates.scope = settings.scope;
  if (Object.prototype.hasOwnProperty.call(settings, 'description')) updates.description = settings.description;

  if (hasDatabaseManager()) {
    try {
      const table = getRolesTable();
      const payload = {};
      if (updates.name) {
        payload.Name = updates.name;
        payload.NormalizedName = updates.name.toUpperCase();
      }
      if (updates.scope) payload.Scope = normalizeScopeValue(updates.scope);
      if (Object.prototype.hasOwnProperty.call(updates, 'description')) payload.Description = updates.description || '';
      if (Object.keys(payload).length) {
        table.update(id, payload);
      }
      return;
    } catch (err) {
      if (typeof safeWriteError === 'function') safeWriteError('updateRoleDb', err);
    }
  }

  legacyUpdateRole(id, updates);
}

function legacyDeleteRole(id) {
  const rolesSheet = legacyEnsureSheet(ROLES_SHEET, ROLES_HEADER);
  const roleData = rolesSheet.getDataRange().getValues();
  const roleHeaders = roleData[0];
  const idIdx = roleHeaders.indexOf('ID');
  for (let i = roleData.length - 1; i >= 1; i--) {
    const rowId = idIdx !== -1 ? roleData[i][idIdx] : roleData[i][0];
    if (String(rowId) === String(id)) {
      rolesSheet.deleteRow(i + 1);
    }
  }

  const userRolesSheet = legacyEnsureSheet(USER_ROLES_SHEET, USER_ROLES_HEADER);
  const urData = userRolesSheet.getDataRange().getValues();
  const urHeaders = urData[0];
  const roleIdx = urHeaders.indexOf('RoleId') !== -1 ? urHeaders.indexOf('RoleId') : urHeaders.indexOf('RoleID');
  for (let j = urData.length - 1; j >= 1; j--) {
    const roleValue = roleIdx !== -1 ? urData[j][roleIdx] : urData[j][2];
    if (String(roleValue) === String(id)) {
      userRolesSheet.deleteRow(j + 1);
    }
  }
}

function deleteRole(id) {
  if (!id) return false;
  if (hasDatabaseManager()) {
    try {
      const rolesTable = getRolesTable();
      const userRolesTable = getUserRolesTable();
      rolesTable.delete(id);
      const assignments = userRolesTable.find({
        filter: function (row) { return normalizeStringValue(row.RoleId || row.RoleID) === normalizeStringValue(id); }
      }) || [];
      assignments.forEach(function (assignment) {
        if (assignment && assignment.ID) {
          userRolesTable.delete(assignment.ID);
        }
      });
      return true;
    } catch (err) {
      if (typeof safeWriteError === 'function') safeWriteError('deleteRoleDb', err);
    }
  }
  legacyDeleteRole(id);
  return true;
}

function legacyAddUserRole(userId, roleId, scope, assignedBy) {
  const sheet = legacyEnsureSheet(USER_ROLES_SHEET, USER_ROLES_HEADER);
  const id = Utilities.getUuid();
  const now = new Date();
  const row = [
    id,
    userId,
    roleId,
    normalizeScopeValue(scope),
    assignedBy || '',
    now,
    now,
    ''
  ];
  sheet.appendRow(row);
  return id;
}

function addUserRole(userId, roleId, scope, assignedBy) {
  if (!userId || !roleId) return null;
  const scopeValue = normalizeScopeValue(scope);
  if (hasDatabaseManager()) {
    try {
      const table = getUserRolesTable();
      const record = table.insert({
        UserId: userId,
        UserID: userId,
        RoleId: roleId,
        RoleID: roleId,
        Scope: scopeValue,
        AssignedBy: assignedBy || '',
        DeletedAt: ''
      });
      return record.ID;
    } catch (err) {
      if (typeof safeWriteError === 'function') safeWriteError('addUserRoleDb', err);
    }
  }
  return legacyAddUserRole(userId, roleId, scopeValue, assignedBy || '');
}

function legacyDeleteUserRoles(userId) {
  const sheet = legacyEnsureSheet(USER_ROLES_SHEET, USER_ROLES_HEADER);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const userIdx = headers.indexOf('UserId') !== -1 ? headers.indexOf('UserId') : headers.indexOf('UserID');
  for (let i = data.length - 1; i >= 1; i--) {
    const rowUser = userIdx !== -1 ? data[i][userIdx] : data[i][1];
    if (String(rowUser) === String(userId)) {
      sheet.deleteRow(i + 1);
    }
  }
}

function deleteUserRoles(userId) {
  if (!userId) return;
  if (hasDatabaseManager()) {
    try {
      const table = getUserRolesTable();
      const rows = table.find({
        filter: function (row) { return normalizeStringValue(row.UserId || row.UserID) === normalizeStringValue(userId); }
      }) || [];
      rows.forEach(function (assignment) {
        if (assignment && assignment.ID) {
          table.delete(assignment.ID);
        }
      });
      return;
    } catch (err) {
      if (typeof safeWriteError === 'function') safeWriteError('deleteUserRolesDb', err);
    }
  }
  legacyDeleteUserRoles(userId);
}

function legacyGetUserRoleIds(userId) {
  const sheet = legacyEnsureSheet(USER_ROLES_SHEET, USER_ROLES_HEADER);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const userIdx = headers.indexOf('UserId') !== -1 ? headers.indexOf('UserId') : headers.indexOf('UserID');
  const roleIdx = headers.indexOf('RoleId') !== -1 ? headers.indexOf('RoleId') : headers.indexOf('RoleID');
  const deletedIdx = headers.indexOf('DeletedAt');
  return data.slice(1)
    .filter(row => String(userIdx !== -1 ? row[userIdx] : row[1]) === String(userId)
      && !normalizeStringValue(deletedIdx !== -1 ? row[deletedIdx] : ''))
    .map(row => roleIdx !== -1 ? row[roleIdx] : row[2]);
}

function getUserRoleIds(userId) {
  if (!userId) return [];
  if (hasDatabaseManager()) {
    try {
      const table = getUserRolesTable();
      const rows = table.find({
        filter: function (row) {
          if (isSoftDeletedRow(row)) return false;
          return normalizeStringValue(row.UserId || row.UserID) === normalizeStringValue(userId);
        }
      }) || [];
      return rows.map(function (row) { return normalizeStringValue(row.RoleId || row.RoleID); });
    } catch (err) {
      if (typeof safeWriteError === 'function') safeWriteError('getUserRoleIdsDb', err);
    }
  }
  return legacyGetUserRoleIds(userId);
}

function getUserRoles(userId) {
  const allRoles = getAllRoles();
  const assigned = getUserRoleIds(userId);
  if (!assigned.length) return [];
  return allRoles.filter(function (role) {
    return assigned.indexOf(role.id) !== -1;
  });
}

function getRolesMapping() {
  try {
    const roles = getAllRoles();
    const mapping = {};
    roles.forEach(role => { mapping[role.id] = role.name; });
    return mapping;
  } catch (error) {
    if (typeof safeWriteError === 'function') safeWriteError('getRolesMapping', error);
    return {};
  }
}

function getUserRoleNames(userRoleIds) {
  try {
    if (!userRoleIds || !Array.isArray(userRoleIds)) {
      return [];
    }
    const rolesMapping = getRolesMapping();
    return userRoleIds.map(roleId => rolesMapping[roleId]).filter(Boolean);
  } catch (error) {
    if (typeof safeWriteError === 'function') safeWriteError('getUserRoleNames', error);
    return [];
  }
}
