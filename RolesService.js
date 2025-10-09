// ────────────────────────────────────────────────────────────────────────────
// Role & UserRole CRUD
// ────────────────────────────────────────────────────────────────────────────

function resolveRoleUserIdentity(userId) {
  const normalized = String(userId || '').trim();
  if (!normalized) {
    return { success: false, error: 'User ID is required', errorCode: 'USER_ID_REQUIRED' };
  }

  function projectRecord(record) {
    if (typeof projectRecordToCanonicalUser === 'function') {
      return projectRecordToCanonicalUser(record, { preferIdentityService: false });
    }
    const source = record || {};
    const clone = {};
    Object.keys(source).forEach(function (key) {
      clone[key] = source[key];
    });
    return clone;
  }

  if (typeof IdentityService !== 'undefined'
    && IdentityService
    && typeof IdentityService.getUserIdentityById === 'function') {
    try {
      const lookup = IdentityService.getUserIdentityById(normalized);
      if (lookup && lookup.success && lookup.identity) {
        const fields = lookup.identity.fields || projectRecord(lookup.identity.raw || {});
        return {
          success: true,
          id: String(fields.ID || normalized).trim(),
          identity: lookup.identity,
          summary: lookup.summary || null,
          fields: fields
        };
      }
      if (lookup && lookup.success === false) {
        return {
          success: false,
          error: lookup.error || 'User not found',
          errorCode: lookup.errorCode || 'USER_NOT_FOUND'
        };
      }
    } catch (identityError) {
      console.warn('RolesService.resolveRoleUserIdentity: IdentityService lookup failed', identityError);
    }
  }

  try {
    if (typeof readSheet === 'function') {
      if (typeof ensureSheetWithHeaders === 'function') {
        try {
          ensureSheetWithHeaders(getUsersSheetName ? getUsersSheetName() : (typeof USERS_SHEET === 'string' && USERS_SHEET) ? USERS_SHEET : 'Users',
            (typeof getCanonicalUserHeaders === 'function') ? getCanonicalUserHeaders() : ((typeof USERS_HEADERS !== 'undefined' && Array.isArray(USERS_HEADERS)) ? USERS_HEADERS : undefined));
        } catch (ensureErr) {
          console.warn('RolesService.resolveRoleUserIdentity: ensure headers failed', ensureErr);
        }
      }
      const users = readSheet((typeof USERS_SHEET === 'string' && USERS_SHEET) ? USERS_SHEET : 'Users') || [];
      const hit = users.find(function (row) {
        return String(row && (row.ID || row.Id || row.UserId)).trim() === normalized;
      }) || null;
      if (!hit) {
        return { success: false, error: 'User not found', errorCode: 'USER_NOT_FOUND' };
      }
      const fields = projectRecord(hit);
      return {
        success: true,
        id: String(fields.ID || normalized).trim(),
        identity: null,
        summary: null,
        fields: fields
      };
    }
  } catch (readError) {
    console.warn('RolesService.resolveRoleUserIdentity: fallback lookup failed', readError);
  }

  return { success: true, id: normalized, identity: null, summary: null, fields: null };
}

/** Returns [{ id, name, normalizedName, scope, description, createdAt, updatedAt, deletedAt }] */
function getAllRoles() {
  try {
    const rows = _readRolesSheet_();
    return rows
      .map(normalizeRoleRecord_)
      .filter(Boolean);
  } catch (error) {
    console.error('Error getting all roles:', error);
    return [];
  }
}

/** Adds a new role */
function addRole(nameOrPayload, scopeOrPayload, description) {
  try {
    const payload = normalizeRoleInput_(nameOrPayload, scopeOrPayload, description);
    if (!payload.name) throw new Error('Role name is required');

    const sheet = _getOrCreateRolesSheet_();
    const headers = _getSheetHeaders_(sheet, ROLES_HEADER);
    const id = Utilities.getUuid();
    const now = new Date();
    const normalizedName = payload.normalizedName || payload.name.toUpperCase();

    const rowMap = {
      ID: id,
      Name: payload.name,
      NormalizedName: normalizedName,
      Scope: payload.scope || '',
      Description: payload.description || '',
      CreatedAt: now,
      UpdatedAt: now,
      DeletedAt: ''
    };

    sheet.appendRow(_mapRowFromHeaders_(headers, rowMap));
    if (typeof invalidateCache === 'function') invalidateCache(ROLES_SHEET);
    return id;
  } catch (error) {
    console.error('Error adding role:', error);
    throw error;
  }
}

/** Updates an existing role */
function updateRole(idOrPayload, name, scope, description) {
  try {
    const payload = normalizeRoleUpdate_(idOrPayload, name, scope, description);
    if (!payload.id) throw new Error('Role ID is required');

    const sheet = _getOrCreateRolesSheet_();
    const data = sheet.getDataRange().getValues();
    if (!data.length) return;

    const headers = data[0].map(h => String(h || '').trim());
    const idx = {
      ID: headers.indexOf('ID'),
      Name: headers.indexOf('Name'),
      NormalizedName: headers.indexOf('NormalizedName'),
      Scope: headers.indexOf('Scope'),
      Description: headers.indexOf('Description'),
      UpdatedAt: headers.indexOf('UpdatedAt')
    };

    const now = new Date();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idx.ID]) === String(payload.id)) {
        const row = i + 1;
        if (idx.Name >= 0) sheet.getRange(row, idx.Name + 1).setValue(payload.name || '');
        if (idx.NormalizedName >= 0) sheet.getRange(row, idx.NormalizedName + 1).setValue((payload.normalizedName || payload.name || '').toUpperCase());
        if (idx.Scope >= 0) sheet.getRange(row, idx.Scope + 1).setValue(payload.scope || '');
        if (idx.Description >= 0) sheet.getRange(row, idx.Description + 1).setValue(payload.description || '');
        if (idx.UpdatedAt >= 0) sheet.getRange(row, idx.UpdatedAt + 1).setValue(now);
        if (typeof invalidateCache === 'function') invalidateCache(ROLES_SHEET);
        return;
      }
    }
  } catch (error) {
    console.error('Error updating role:', error);
    throw error;
  }
}

/** Deletes a role (and any user-role links) */
function deleteRole(id) {
  try {
    const ss = SpreadsheetApp.getActive();
    const rolesSheet = ss.getSheetByName(ROLES_SHEET);
    if (rolesSheet) {
      const data = rolesSheet.getDataRange().getValues();
      if (data.length) {
        const headers = data[0].map(h => String(h || '').trim());
        const idCol = headers.indexOf('ID');
        for (let i = data.length - 1; i >= 1; i--) {
          if (idCol >= 0 && String(data[i][idCol]) === String(id)) {
            rolesSheet.deleteRow(i + 1);
          }
        }
      }
    }

    const userRolesSheet = ss.getSheetByName(USER_ROLES_SHEET);
    if (userRolesSheet) {
      const data = userRolesSheet.getDataRange().getValues();
      if (data.length) {
        const headers = data[0].map(h => String(h || '').trim());
        const roleIdx = headers.indexOf('RoleId');
        for (let i = data.length - 1; i >= 1; i--) {
          if (roleIdx >= 0 && String(data[i][roleIdx]) === String(id)) {
            userRolesSheet.deleteRow(i + 1);
          }
        }
      }
    }

    if (typeof invalidateCache === 'function') {
      invalidateCache(ROLES_SHEET);
      invalidateCache(USER_ROLES_SHEET);
    }
  } catch (error) {
    console.error('Error deleting role:', error);
    throw error;
  }
}

/** Assigns a role to a user */
function addUserRole(userId, roleId, scopeOrOptions, assignedBy) {
  try {
    if (!userId || !roleId) throw new Error('User ID and Role ID are required');

    const identityInfo = resolveRoleUserIdentity(userId);
    if (!identityInfo.success) {
      throw new Error(identityInfo.error || 'Unable to resolve user identity');
    }

    const canonicalUserId = identityInfo.id;

    const sheet = SpreadsheetApp.getActive().getSheetByName(USER_ROLES_SHEET);
    if (!sheet) throw new Error(`Sheet ${USER_ROLES_SHEET} not found`);

    const headers = _getSheetHeaders_(sheet, USER_ROLES_HEADER);
    const now = new Date();
    const id = Utilities.getUuid();

    let scope = '';
    let assigned = '';
    if (scopeOrOptions && typeof scopeOrOptions === 'object' && !Array.isArray(scopeOrOptions)) {
      scope = scopeOrOptions.scope || scopeOrOptions.Scope || '';
      assigned = scopeOrOptions.assignedBy || scopeOrOptions.AssignedBy || scopeOrOptions.assigned_by || '';
    } else {
      scope = scopeOrOptions || '';
      assigned = assignedBy || '';
    }

    const rowMap = {
      ID: id,
      UserId: canonicalUserId,
      RoleId: roleId,
      Scope: scope,
      AssignedBy: assigned,
      CreatedAt: now,
      UpdatedAt: now,
      DeletedAt: ''
    };

    sheet.appendRow(_mapRowFromHeaders_(headers, rowMap));
    if (typeof invalidateCache === 'function') invalidateCache(USER_ROLES_SHEET);
    return id;
  } catch (error) {
    console.error('Error assigning user role:', error);
    throw error;
  }
}

/** Removes roles for a user (optionally target a specific role) */
function deleteUserRoles(userId, roleId) {
  try {
    const identityInfo = resolveRoleUserIdentity(userId);
    if (!identityInfo.success) {
      return;
    }
    const canonicalUserId = identityInfo.id;

    const sheet = SpreadsheetApp.getActive().getSheetByName(USER_ROLES_SHEET);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (!data.length) return;

    const headers = data[0].map(h => String(h || '').trim());
    const userIdx = headers.indexOf('UserId');
    const roleIdx = headers.indexOf('RoleId');
    if (userIdx < 0) return;

    for (let i = data.length - 1; i >= 1; i--) {
      const matchesUser = String(data[i][userIdx]) === String(canonicalUserId);
      const matchesRole = roleId ? String(data[i][roleIdx]) === String(roleId) : true;
      if (matchesUser && matchesRole) {
        sheet.deleteRow(i + 1);
      }
    }

    if (typeof invalidateCache === 'function') invalidateCache(USER_ROLES_SHEET);
  } catch (error) {
    console.error('Error deleting user roles:', error);
    throw error;
  }
}

/** Returns an array of roleIds for the given user */
function getUserRoleIds(userId) {
  try {
    if (!userId) return [];
    const identityInfo = resolveRoleUserIdentity(userId);
    if (!identityInfo.success) {
      return [];
    }
    const canonicalUserId = identityInfo.id;
    const sheet = SpreadsheetApp.getActive().getSheetByName(USER_ROLES_SHEET);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0].map(h => String(h || '').trim());
    const userIdx = headers.indexOf('UserId');
    const roleIdx = headers.indexOf('RoleId');
    if (userIdx < 0 || roleIdx < 0) return [];

    return data
      .slice(1)
      .filter(row => String(row[userIdx]) === String(canonicalUserId))
      .map(row => row[roleIdx])
      .filter(Boolean);
  } catch (error) {
    console.error('Error getting user role IDs:', error);
    return [];
  }
}

/** Returns full role objects for a user */
function getUserRoles(userId) {
  try {
    const identityInfo = resolveRoleUserIdentity(userId);
    if (!identityInfo.success) {
      return [];
    }
    const canonicalUserId = identityInfo.id;
    const allRoles = getAllRoles();
    const assigned = new Set(getUserRoleIds(canonicalUserId).map(String));
    return allRoles.filter(r => assigned.has(String(r.id || r.ID)));
  } catch (error) {
    console.error('Error getting user roles:', error);
    return [];
  }
}

/**
 * getRolesMapping()
 * Returns a mapping object of roleId -> roleName for quick lookups
 * Used in doGet to convert user role IDs to role names
 */
function getRolesMapping() {
  try {
    const roles = getAllRoles();
    const mapping = {};
    roles.forEach(role => {
      const id = role.id || role.ID;
      const name = role.name || role.Name;
      if (id) mapping[id] = name;
    });
    return mapping;
  } catch (error) {
    console.error('Error getting roles mapping:', error);
    return {};
  }
}

/**
 * Alternative helper function to get user role names directly
 */
function getUserRoleNames(userRoleIds) {
  try {
    if (!userRoleIds || !Array.isArray(userRoleIds)) {
      return [];
    }

    const rolesMapping = getRolesMapping();
    return userRoleIds.map(roleId => rolesMapping[roleId]).filter(Boolean);
  } catch (error) {
    console.error('Error getting user role names:', error);
    return [];
  }
}

// ────────────────────────────────────────────────────────────────────────────
// Helpers
// ────────────────────────────────────────────────────────────────────────────

function _getOrCreateRolesSheet_() {
  if (typeof ensureSheetWithHeaders === 'function') {
    return ensureSheetWithHeaders(ROLES_SHEET, ROLES_HEADER);
  }
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(ROLES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(ROLES_SHEET);
    sheet.getRange(1, 1, 1, ROLES_HEADER.length).setValues([ROLES_HEADER]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function _getSheetHeaders_(sheet, fallback) {
  if (!sheet) return Array.isArray(fallback) ? fallback.slice() : [];
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return Array.isArray(fallback) ? fallback.slice() : [];
  const values = sheet.getRange(1, 1, 1, lastColumn).getValues();
  if (!values.length) return Array.isArray(fallback) ? fallback.slice() : [];
  return values[0].map(h => String(h || '').trim());
}

function _mapRowFromHeaders_(headers, valueMap) {
  return headers.map(header => {
    if (!header) return '';
    return Object.prototype.hasOwnProperty.call(valueMap, header) ? valueMap[header] : '';
  });
}

function _readRolesSheet_() {
  if (typeof readSheet === 'function') {
    const rows = readSheet(ROLES_SHEET) || [];
    if (Array.isArray(rows)) return rows;
  }

  const sheet = SpreadsheetApp.getActive().getSheetByName(ROLES_SHEET);
  if (!sheet) return [];
  const range = sheet.getDataRange();
  if (!range) return [];
  const values = range.getValues();
  if (!values.length) return [];
  const headers = values.shift().map(h => String(h || '').trim());
  return values
    .map(row => {
      const obj = {};
      let hasData = false;
      headers.forEach((header, idx) => {
        if (!header) return;
        const value = row[idx];
        if (value !== '' && value != null) hasData = true;
        obj[header] = value;
      });
      return hasData ? obj : null;
    })
    .filter(Boolean);
}

function normalizeRoleRecord_(record) {
  if (!record || typeof record !== 'object') return null;
  const base = Array.isArray(record) ? {} : Object.assign({}, record);

  let id = base.ID || base.Id || base.id || '';
  if (id != null) id = String(id).trim();
  let name = base.Name || base.name || '';
  if (name != null) name = String(name).trim();
  const normalizedName = (base.NormalizedName || base.normalizedName || (name ? name.toUpperCase() : '') || '').toString();
  const scope = base.Scope || base.scope || '';
  const description = base.Description || base.description || '';

  const createdValue = base.CreatedAt || base.createdAt || null;
  const updatedValue = base.UpdatedAt || base.updatedAt || null;
  const deletedValue = base.DeletedAt || base.deletedAt || null;

  return Object.assign({}, base, {
    ID: id || base.ID,
    id: id,
    Name: name || base.Name,
    name: name,
    NormalizedName: normalizedName,
    normalizedName: normalizedName,
    Scope: scope,
    scope: scope,
    Description: description,
    description: description,
    CreatedAt: createdValue,
    createdAt: _toClientDate_(createdValue),
    UpdatedAt: updatedValue,
    updatedAt: _toClientDate_(updatedValue),
    DeletedAt: deletedValue,
    deletedAt: _toClientDate_(deletedValue)
  });
}

function normalizeRoleInput_(nameOrPayload, scopeOrPayload, description) {
  if (nameOrPayload && typeof nameOrPayload === 'object' && !Array.isArray(nameOrPayload)) {
    return {
      name: nameOrPayload.name || nameOrPayload.Name || '',
      normalizedName: nameOrPayload.normalizedName || nameOrPayload.NormalizedName || '',
      scope: nameOrPayload.scope || nameOrPayload.Scope || '',
      description: nameOrPayload.description || nameOrPayload.Description || ''
    };
  }

  return {
    name: nameOrPayload || '',
    scope: scopeOrPayload && typeof scopeOrPayload === 'string' ? scopeOrPayload : '',
    description: description || ''
  };
}

function normalizeRoleUpdate_(idOrPayload, name, scope, description) {
  if (idOrPayload && typeof idOrPayload === 'object' && !Array.isArray(idOrPayload)) {
    const payload = Object.assign({}, idOrPayload);
    payload.id = idOrPayload.id || idOrPayload.ID || '';
    payload.name = idOrPayload.name || idOrPayload.Name || payload.name || '';
    payload.scope = idOrPayload.scope || idOrPayload.Scope || payload.scope || '';
    payload.description = idOrPayload.description || idOrPayload.Description || payload.description || '';
    payload.normalizedName = idOrPayload.normalizedName || idOrPayload.NormalizedName || payload.normalizedName || '';
    return payload;
  }

  return {
    id: idOrPayload,
    name: name || '',
    scope: scope || '',
    description: description || '',
    normalizedName: name ? name.toUpperCase() : ''
  };
}

function _toClientDate_(value) {
  if (!value && value !== 0) return null;
  if (value instanceof Date) return value.toISOString();
  return value;
}