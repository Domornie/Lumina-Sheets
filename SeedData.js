/**
 * SeedData.js
 * -----------------------------------------------------------------------------
 * Lightweight bootstrap seeded entirely through the public service APIs.
 *
 * Run seedDefaultData() once to ensure:
 *   • Core roles exist
 *   • A couple of starter campaigns are provisioned
 *   • The Lumina administrator account is created (or refreshed) with a known login
 *
 * The implementation deliberately delegates to the same helpers used by the
 * production flows (UserService, RolesService, CampaignService, and
 * AuthenticationService) so the seeded data matches real runtime expectations.
 */

const SEED_ROLE_NAMES = [
  'Super Admin',
  'Administrator',
  'Operations Manager',
  'Agent'
];

const SEED_CAMPAIGNS = [
  { name: 'Lumina HQ', description: 'Lumina internal operations workspace' },
];

const SEED_LUMINA_ADMIN_PROFILE = {
  userName: 'lumina.admin',
  fullName: 'Lumina Admin',
  email: 'lumina@vlbpo.com',
  password: 'ChangeMe123!',
  defaultCampaign: 'Lumina HQ',
  roleNames: ['Administrator'],
  seedLabel: 'Lumina Administrator'
};

const PASSWORD_UTILS = (function resolvePasswordUtilities() {
  if (typeof ensurePasswordUtilities === 'function') {
    return ensurePasswordUtilities();
  }

  if (typeof PasswordUtilities !== 'undefined' && PasswordUtilities) {
    return PasswordUtilities;
  }

  if (typeof __createPasswordUtilitiesModule === 'function') {
    const utils = __createPasswordUtilitiesModule();

    if (typeof PasswordUtilities === 'undefined' || !PasswordUtilities) {
      PasswordUtilities = utils;
    }

    if (typeof ensurePasswordUtilities !== 'function') {
      ensurePasswordUtilities = function ensurePasswordUtilities() { return utils; };
    }

    return utils;
  }

  throw new Error('PasswordUtilities module is not available.');

})();

/**
 * Public entry point. Returns a structured summary of what was ensured.
 */
function seedDefaultData() {
  const summary = {
    roles: { created: [], existing: [] },
    campaigns: { created: [], existing: [] },
    luminaAdmin: null
  };

  try {
    // Make sure the identity sheets exist up front.
    if (typeof AuthenticationService !== 'undefined' && AuthenticationService.ensureSheets) {
      AuthenticationService.ensureSheets();
    }

    ensureSheetWithHeaders(ROLES_SHEET, ROLES_HEADER);
    ensureSheetWithHeaders(USER_ROLES_SHEET, USER_ROLES_HEADER);
    ensureSheetWithHeaders(CAMPAIGNS_SHEET, CAMPAIGNS_HEADERS);
    if (typeof USER_CAMPAIGNS_SHEET !== 'undefined' && typeof USER_CAMPAIGNS_HEADERS !== 'undefined') {
      ensureSheetWithHeaders(USER_CAMPAIGNS_SHEET, USER_CAMPAIGNS_HEADERS);
    }

    const roleIdsByName = ensureCoreRoles(summary);
    const campaignIdsByName = ensureCoreCampaigns(summary);

    const luminaAdminInfo = ensureLuminaAdminUser(roleIdsByName, campaignIdsByName);
    summary.luminaAdmin = luminaAdminInfo;

    return {
      success: true,
      message: 'Seed data ensured successfully.',
      details: summary
    };
  } catch (error) {
    console.error('seedDefaultData failed:', error);
    if (typeof writeError === 'function') {
      writeError('seedDefaultData', error);
    }
    return {
      success: false,
      message: 'Seed data failed: ' + (error && error.message ? error.message : error),
      details: summary
    };
  }
}

/**
 * Ensure the baseline roles exist and capture their IDs.
 * @returns {Object} Map of role name -> roleId
 */
function ensureCoreRoles(summary) {
  const existingRoles = (typeof getAllRoles === 'function') ? getAllRoles() : [];
  const roleMap = {};
  existingRoles.forEach(role => {
    if (role && role.name) {
      roleMap[role.name.toLowerCase()] = role.id;
    }
  });

  SEED_ROLE_NAMES.forEach(name => {
    const key = name.toLowerCase();
    if (roleMap[key]) {
      summary.roles.existing.push(name);
      return;
    }

    if (typeof addRole !== 'function') {
      throw new Error('RolesService.addRole is not available');
    }

    const newId = addRole(name);
    roleMap[key] = newId;
    summary.roles.created.push(name);
  });

  // Rebuild the mapping using the authoritative data to capture IDs even if
  // roles already existed or were just created.
  const finalRoles = (typeof getAllRoles === 'function') ? getAllRoles() : [];
  const idsByName = {};
  finalRoles.forEach(role => {
    if (role && role.name) {
      idsByName[role.name] = role.id;
      idsByName[role.name.toLowerCase()] = role.id;
    }
  });

  return idsByName;
}

/**
 * Ensure the baseline campaigns exist and capture their IDs.
 * @returns {Object} Map of campaign name -> campaignId
 */
function ensureCoreCampaigns(summary) {
  const existing = getCampaignsIndex();

  SEED_CAMPAIGNS.forEach(campaign => {
    const key = campaign.name.toLowerCase();
    if (existing[key]) {
      summary.campaigns.existing.push(campaign.name);
      return;
    }

    if (typeof csCreateCampaign !== 'function') {
      throw new Error('CampaignService.csCreateCampaign is not available');
    }

    const result = csCreateCampaign(campaign.name, campaign.description || '');
    if (result && result.success) {
      summary.campaigns.created.push(campaign.name);
    } else {
      // Treat duplicates as existing so re-runs stay idempotent.
      summary.campaigns.existing.push(campaign.name);
    }
  });

  // Refresh to pick up any IDs assigned during creation.
  return getCampaignsIndex(true);
}

/**
 * Build a lookup of campaign name -> id using CampaignService helpers.
 * @param {boolean} forceRefresh Whether to re-read campaigns
 */
function getCampaignsIndex(forceRefresh) {
  let campaigns = [];
  if (typeof csGetAllCampaigns === 'function') {
    campaigns = csGetAllCampaigns();
  }

  if ((forceRefresh || !campaigns || !campaigns.length) && typeof readSheet === 'function') {
    campaigns = (readSheet(CAMPAIGNS_SHEET) || []).map(c => ({
      id: c.ID,
      name: c.Name,
      description: c.Description || ''
    }));
  }

  const index = {};
  (campaigns || []).forEach(c => {
    if (c && c.name) {
      index[c.name.toLowerCase()] = c.id;
      index[c.name] = c.id;
    }
  });
  return index;
}

/**
 * Ensure there is a Lumina admin user with a known password and permissions.
 */
function ensureLuminaAdminUser(roleIdsByName, campaignIdsByName) {
  return ensureSeedAdministrator(SEED_LUMINA_ADMIN_PROFILE, roleIdsByName, campaignIdsByName);
}

function applySeedPasswordForUser(userRecord, profile, label) {
  if (!profile || !profile.password || !userRecord || !userRecord.ID) {
    return userRecord;
  }

  const resolvedLabel = label || profile.seedLabel || profile.fullName || profile.email;

  if (userRecord.EmailConfirmation) {
    const setPasswordResult = setPasswordWithToken(userRecord.EmailConfirmation, profile.password);
    if (!setPasswordResult || !setPasswordResult.success) {
      throw new Error('Failed to set ' + resolvedLabel + ' password: ' + (setPasswordResult && setPasswordResult.message ? setPasswordResult.message : 'Unknown error'));
    }
  } else {
    setUserPasswordDirect(userRecord.ID, profile.password);
  }

  if (typeof AuthenticationService !== 'undefined' && AuthenticationService.getUserByEmail) {
    const refreshed = AuthenticationService.getUserByEmail(profile.email);
    if (!refreshed) {
      throw new Error(resolvedLabel + ' record not found after password update.');
    }
    return refreshed;
  }

  return userRecord;
}

/**
 * Shared implementation for creating or refreshing privileged seed accounts.
 */
function ensureSeedAdministrator(profile, roleIdsByName, campaignIdsByName) {
  if (!profile || !profile.email) {
    throw new Error('Seed administrator profile is not configured correctly.');
  }

  const label = profile.seedLabel || profile.fullName || profile.email;
  const desiredRoleIds = (profile.roleNames || [])
    .map(name => {
      if (!name) return null;
      const key = String(name);
      return roleIdsByName[key] || roleIdsByName[key.toLowerCase()];
    })
    .filter(Boolean);

  const defaultCampaignKey = profile.defaultCampaign
    ? String(profile.defaultCampaign).toLowerCase()
    : '';

  const primaryCampaignId = (defaultCampaignKey && (campaignIdsByName[defaultCampaignKey] || campaignIdsByName[profile.defaultCampaign]))
    || Object.values(campaignIdsByName)[0]
    || '';

  if (!primaryCampaignId) {
    throw new Error('No campaigns exist to assign to the administrator.');
  }

  const accountFlags = Object.assign({
    canLogin: true,
    isAdmin: true,
    permissionLevel: 'ADMIN',
    canManageUsers: true,
    canManagePages: true
  }, profile.accountOverrides || {});

  const payload = Object.assign({
    userName: profile.userName,
    fullName: profile.fullName,
    email: profile.email,
    campaignId: primaryCampaignId,
    roles: desiredRoleIds
  }, accountFlags);

  const existing = (typeof AuthenticationService !== 'undefined' && AuthenticationService.getUserByEmail)
    ? AuthenticationService.getUserByEmail(profile.email)
    : null;

  if (existing) {
    const updateResult = clientUpdateUser(existing.ID, payload);

    if (!updateResult || !updateResult.success) {
      throw new Error('Failed to refresh ' + label + ': ' + (updateResult && updateResult.error ? updateResult.error : 'Unknown error'));
    }

    applySeedPasswordForUser(existing, profile, label);
    syncUserRoleLinks(existing.ID, desiredRoleIds);
    assignAdminCampaignAccess(existing.ID, Object.values(campaignIdsByName));
    ensureCanLoginFlag(existing.ID, true);

    const result = {
      status: 'updated',
      userId: existing.ID,
      email: profile.email,
      message: (updateResult && updateResult.message) || (label + ' refreshed.')
    };

    if (profile.password) {
      result.password = profile.password;
    }

    return result;
  }

  const createResult = clientRegisterUser(payload);

  if (!createResult || !createResult.success) {
    throw new Error('Failed to create ' + label + ': ' + (createResult && createResult.error ? createResult.error : 'Unknown error'));
  }

  let adminRecord = AuthenticationService.getUserByEmail(profile.email);
  if (!adminRecord) {
    throw new Error(label + ' record not found after creation.');
  }

  adminRecord = applySeedPasswordForUser(adminRecord, profile, label);
  syncUserRoleLinks(adminRecord.ID, desiredRoleIds);
  assignAdminCampaignAccess(adminRecord.ID, Object.values(campaignIdsByName));
  ensureCanLoginFlag(adminRecord.ID, true);

  const result = {
    status: 'created',
    userId: adminRecord.ID,
    email: adminRecord.Email,
    message: label + ' account created with default credentials. Please change the password after first login.'
  };

  if (profile.password) {
    result.password = profile.password;
  }

  return result;
}

/**
 * Ensure UserRoles contains links for each desired role without duplicating rows.
 */
function syncUserRoleLinks(userId, roleIds) {
  if (!userId || !Array.isArray(roleIds) || !roleIds.length) {
    return;
  }

  const existingIds = (typeof getUserRoleIds === 'function') ? getUserRoleIds(userId) : [];
  const existingSet = new Set((existingIds || []).map(String));

  roleIds.forEach(roleId => {
    if (!roleId) return;
    const key = String(roleId);
    if (existingSet.has(key)) return;
    if (typeof addUserRole === 'function') {
      addUserRole(userId, roleId);
    }
    existingSet.add(key);
  });
}

/**
 * Give the administrator access to every campaign at the ADMIN level.
 */
function assignAdminCampaignAccess(userId, campaignIds) {
  if (!userId || !Array.isArray(campaignIds)) {
    return;
  }

  const uniqueIds = Array.from(new Set(campaignIds.map(id => String(id || ''))))
    .filter(id => id);

  uniqueIds.forEach(campaignId => {
    if (typeof setCampaignUserPermissions === 'function') {
      setCampaignUserPermissions(campaignId, userId, 'ADMIN', true, true);
    }
    if (typeof addUserToCampaign === 'function') {
      addUserToCampaign(userId, campaignId);
    }
  });
}

/**
 * Toggle the CanLogin flag for a specific user.
 */
function ensureCanLoginFlag(userId, canLogin) {
  if (!userId) return;

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);
  if (!sh) return;

  const data = sh.getDataRange().getValues();
  if (!data || !data.length) return;

  const headers = data[0];
  const idIdx = headers.indexOf('ID');
  const canLoginIdx = headers.indexOf('CanLogin');
  const resetRequiredIdx = headers.indexOf('ResetRequired');

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][idIdx]) === String(userId)) {
      if (canLoginIdx >= 0) {
        sh.getRange(r + 1, canLoginIdx + 1).setValue(canLogin ? 'TRUE' : 'FALSE');
      }
      if (resetRequiredIdx >= 0 && canLogin) {
        sh.getRange(r + 1, resetRequiredIdx + 1).setValue('FALSE');
      }
      break;
    }
  }

  if (typeof invalidateCache === 'function') {
    invalidateCache(USERS_SHEET);
  }
}

/**
 * Directly set a password hash when a setup token is unavailable.
 */
function setUserPasswordDirect(userId, password) {
  if (!userId || !password) return;

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);
  if (!sh) return;

  const data = sh.getDataRange().getValues();
  if (!data || data.length < 2) return;

  const headers = data[0];
  const idIdx = headers.indexOf('ID');
  const resetIdx = headers.indexOf('ResetRequired');
  const updatedIdx = headers.indexOf('UpdatedAt');

  const passwordUpdate = PASSWORD_UTILS.createPasswordUpdate(password);
  const updateColumns = passwordUpdate.columns || { PasswordHash: passwordUpdate.hash };
  const now = new Date();

  const columnIndexMap = {};
  headers.forEach((header, idx) => {
    columnIndexMap[String(header)] = idx;
  });

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][idIdx]) === String(userId)) {
      Object.keys(updateColumns).forEach(columnName => {
        const columnIdx = columnIndexMap[columnName];
        if (typeof columnIdx === 'number' && columnIdx >= 0) {
          sh.getRange(r + 1, columnIdx + 1).setValue(updateColumns[columnName]);
        }
      });

      if (passwordUpdate.algorithm) {
        const algoIdx = columnIndexMap['PasswordHashAlgorithm'];
        if (typeof algoIdx === 'number' && algoIdx >= 0) {
          sh.getRange(r + 1, algoIdx + 1).setValue(passwordUpdate.algorithm);
        }
      }

      if (resetIdx >= 0) sh.getRange(r + 1, resetIdx + 1).setValue('FALSE');
      if (updatedIdx >= 0) sh.getRange(r + 1, updatedIdx + 1).setValue(now);
      break;
    }
  }

  SpreadsheetApp.flush();
  if (typeof invalidateCache === 'function') {
    invalidateCache(USERS_SHEET);
  }
}
