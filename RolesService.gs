// ────────────────────────────────────────────────────────────────────────────
// Role & UserRole CRUD
// ────────────────────────────────────────────────────────────────────────────

/** Returns [{ id, name, normalizedName }] */
function getAllRoles() {
  const rows = SpreadsheetApp
    .getActive()
    .getSheetByName(ROLES_SHEET)
    .getDataRange()
    .getValues()
    .slice(1);
  return rows.map(r => ({
    id: r[0],
    name: r[1],
    normalizedName: r[2]
  }));
}

/** Adds a new role */
function addRole(name) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ROLES_SHEET);
  const id = Utilities.getUuid();
  const now   = new Date();
  sheet.appendRow([id, name, name.toUpperCase(), now, now]);
  return id;
}

/** Updates an existing role */
function updateRole(id, newName) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ROLES_SHEET);
  const data  = sheet.getDataRange().getValues();
  const now   = new Date();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const row = i + 1;
      sheet.getRange(row, 2).setValue(newName);
      sheet.getRange(row, 3).setValue(newName.toUpperCase());
      sheet.getRange(row, 5).setValue(now);
      break;
    }
  }
}

/** Deletes a role (and any user-role links) */
function deleteRole(id) {
  const ss = SpreadsheetApp.getActive();
  // remove from Roles
  const sh = ss.getSheetByName(ROLES_SHEET);
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === id) sh.deleteRow(i+1);
  }
  // remove any UserRoles
  const shUR = ss.getSheetByName(USER_ROLES_SHEET);
  const urData = shUR.getDataRange().getValues();
  for (let i = urData.length - 1; i >= 1; i--) {
    if (urData[i][1] === id) shUR.deleteRow(i+1);
  }
}

/** Assigns a role to a user */
function addUserRole(userId, roleId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(USER_ROLES_SHEET);
  const now   = new Date();
  sheet.appendRow([userId, roleId, now, now]);
}

/** Removes all roles for a user (you can also add targeted removal) */
function deleteUserRoles(userId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(USER_ROLES_SHEET);
  const data  = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === userId) sheet.deleteRow(i+1);
  }
}

/** Returns an array of roleIds for the given user */
function getUserRoleIds(userId) {
  const rows = SpreadsheetApp
    .getActive()
    .getSheetByName(USER_ROLES_SHEET)
    .getDataRange()
    .getValues()
    .slice(1)
    .filter(r => r[0] === userId);
  return rows.map(r => r[1]);
}

/** Returns full role objects for a user */
function getUserRoles(userId) {
  const allRoles = getAllRoles();
  const assigned = getUserRoleIds(userId);
  return allRoles.filter(r => assigned.indexOf(r.id) !== -1);
}

/**
 * getRolesMapping()
 * Returns a mapping object of roleId -> roleName for quick lookups
 * Used in doGet to convert user role IDs to role names
 */
function getRolesMapping() {
  try {
    const roles = getAllRoles(); // This function exists in your system
    const mapping = {};

    roles.forEach(role => {
      mapping[role.id] = role.name;
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