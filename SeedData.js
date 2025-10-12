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
  {
    name: 'Lumina HQ',
    description: 'Lumina internal operations workspace',
    clientName: 'Lumina',
    status: 'Active',
    channel: 'Operations',
    timezone: 'America/New_York',
    slaTier: 'Enterprise'
  },
  {
    name: 'Credit Suite',
    description: 'Credit Suite operations campaign',
    clientName: 'Credit Suite',
    status: 'Active',
    channel: 'Financial Services',
    timezone: 'America/New_York',
    slaTier: 'Gold'
  },
  {
    name: 'HiyaCar',
    description: 'HiyaCar mobility support campaign',
    clientName: 'HiyaCar',
    status: 'Active',
    channel: 'Mobility',
    timezone: 'Europe/London',
    slaTier: 'Standard'
  },
  {
    name: 'Benefits Resource Center (iBTR)',
    description: 'Benefits Resource Center member services',
    clientName: 'Benefits Resource Center',
    status: 'Active',
    channel: 'Benefits Support',
    timezone: 'America/Chicago',
    slaTier: 'Gold'
  },
  {
    name: 'Independence Insurance Agency',
    description: 'Independence Insurance Agency customer care',
    clientName: 'Independence Insurance Agency',
    status: 'Active',
    channel: 'Insurance',
    timezone: 'America/New_York',
    slaTier: 'Gold'
  },
  {
    name: 'JSC',
    description: 'JSC customer success campaign',
    clientName: 'JSC',
    status: 'Active',
    channel: 'Customer Success',
    timezone: 'America/New_York',
    slaTier: 'Standard'
  },
  {
    name: 'Kids in the Game',
    description: 'Kids in the Game coaching support',
    clientName: 'Kids in the Game',
    status: 'Active',
    channel: 'Coaching Support',
    timezone: 'America/New_York',
    slaTier: 'Standard'
  },
  {
    name: 'Kofi Group',
    description: 'Kofi Group recruiting operations',
    clientName: 'Kofi Group',
    status: 'Active',
    channel: 'Recruiting',
    timezone: 'America/Chicago',
    slaTier: 'Standard'
  },
  {
    name: 'PAW Law Firm',
    description: 'PAW Law Firm client services',
    clientName: 'PAW Law Firm',
    status: 'Active',
    channel: 'Legal Services',
    timezone: 'America/New_York',
    slaTier: 'Standard'
  },
  {
    name: 'Pro House Photos',
    description: 'Pro House Photos client success',
    clientName: 'Pro House Photos',
    status: 'Active',
    channel: 'Real Estate',
    timezone: 'America/Chicago',
    slaTier: 'Standard'
  },
  {
    name: 'Independence Agency & Credit Suite',
    description: 'Independence Agency & Credit Suite blended operations',
    clientName: 'Independence Agency & Credit Suite',
    status: 'Active',
    channel: 'Blended Operations',
    timezone: 'America/New_York',
    slaTier: 'Gold'
  },
  {
    name: 'Proozy',
    description: 'Proozy ecommerce support',
    clientName: 'Proozy',
    status: 'Active',
    channel: 'Ecommerce',
    timezone: 'America/Chicago',
    slaTier: 'Standard'
  },
  {
    name: 'The Grounding',
    description: 'The Grounding campaign (TGC)',
    clientName: 'The Grounding Company',
    status: 'Active',
    channel: 'Wellness',
    timezone: 'America/Los_Angeles',
    slaTier: 'Standard'
  },
  {
    name: 'CO',
    description: 'CO services campaign',
    clientName: 'CO',
    status: 'Active',
    channel: 'Operations',
    timezone: 'America/New_York',
    slaTier: 'Standard'
  }
];

const SEED_LUMINA_ADMIN_PROFILE = {
  userName: 'lumina.admin',
  fullName: 'Lumina Admin',
  email: 'lumina@vlbpo.com',
  defaultCampaign: 'Lumina HQ',
  roleNames: ['Super Admin', 'Administrator'],
  claimTypes: ['system.admin', 'lumina.admin', 'manage.users', 'manage.pages'],
  seedLabel: 'Lumina Administrator'
};

const IDENTITY_ROLE_SEED = [
  { role: 'System Admin', description: 'Bootstrap superuser', isGlobal: 'Y', permissions: [
    { capability: 'VIEW_USERS', scope: 'Global' },
    { capability: 'MANAGE_USERS', scope: 'Global' },
    { capability: 'ASSIGN_ROLES', scope: 'Global' },
    { capability: 'TRANSFER_USERS', scope: 'Global' },
    { capability: 'TERMINATE_USERS', scope: 'Global' },
    { capability: 'MANAGE_EQUIPMENT', scope: 'Global' },
    { capability: 'VIEW_AUDIT', scope: 'Global' },
    { capability: 'MANAGE_POLICIES', scope: 'Global' }
  ] },
  { role: 'CEO', description: 'Executive leadership', isGlobal: 'Y', permissions: [
    { capability: 'VIEW_USERS', scope: 'Global' },
    { capability: 'VIEW_AUDIT', scope: 'Global' }
  ] },
  { role: 'COO', description: 'Operations executive', isGlobal: 'Y', permissions: [
    { capability: 'VIEW_USERS', scope: 'Global' },
    { capability: 'MANAGE_USERS', scope: 'Global' },
    { capability: 'TRANSFER_USERS', scope: 'Global' },
    { capability: 'VIEW_AUDIT', scope: 'Global' }
  ] },
  { role: 'CFO', description: 'Finance leadership', isGlobal: 'Y', permissions: [
    { capability: 'VIEW_USERS', scope: 'Global' },
    { capability: 'VIEW_AUDIT', scope: 'Global' }
  ] },
  { role: 'CTO', description: 'Technology leadership', isGlobal: 'Y', permissions: [
    { capability: 'VIEW_USERS', scope: 'Global' },
    { capability: 'MANAGE_POLICIES', scope: 'Global' }
  ] },
  { role: 'Call Center Director', description: 'Multi-campaign operations leader', isGlobal: 'N', permissions: [
    { capability: 'VIEW_USERS', scope: 'Global' },
    { capability: 'MANAGE_USERS', scope: 'Global' },
    { capability: 'ASSIGN_ROLES', scope: 'Global' },
    { capability: 'TRANSFER_USERS', scope: 'Global' },
    { capability: 'TERMINATE_USERS', scope: 'Global' },
    { capability: 'VIEW_AUDIT', scope: 'Global' }
  ] },
  { role: 'Operations Manager', description: 'Multi-campaign operations manager', isGlobal: 'N', permissions: [
    { capability: 'VIEW_USERS', scope: 'Global' },
    { capability: 'MANAGE_USERS', scope: 'Global' },
    { capability: 'ASSIGN_ROLES', scope: 'Global' },
    { capability: 'TRANSFER_USERS', scope: 'Global' },
    { capability: 'TERMINATE_USERS', scope: 'Global' },
    { capability: 'MANAGE_EQUIPMENT', scope: 'Global' },
    { capability: 'VIEW_AUDIT', scope: 'Global' }
  ] },
  { role: 'Account Manager', description: 'Client-facing operations lead', isGlobal: 'N', permissions: [
    { capability: 'VIEW_USERS', scope: 'Global' },
    { capability: 'MANAGE_USERS', scope: 'Campaign' },
    { capability: 'ASSIGN_ROLES', scope: 'Campaign' },
    { capability: 'TRANSFER_USERS', scope: 'Campaign' },
    { capability: 'TERMINATE_USERS', scope: 'Campaign' },
    { capability: 'VIEW_AUDIT', scope: 'Campaign' }
  ] },
  { role: 'Workforce Manager', description: 'Workforce management', isGlobal: 'N', permissions: [
    { capability: 'VIEW_USERS', scope: 'Campaign' }
  ] },
  { role: 'Quality Assurance Manager', description: 'QA manager', isGlobal: 'N', permissions: [
    { capability: 'VIEW_USERS', scope: 'Campaign' }
  ] },
  { role: 'Training Manager', description: 'Training oversight', isGlobal: 'N', permissions: [
    { capability: 'VIEW_USERS', scope: 'Campaign' }
  ] },
  { role: 'Team Supervisor', description: 'Team-level supervisor', isGlobal: 'N', permissions: [
    { capability: 'VIEW_USERS', scope: 'Team' }
  ] },
  { role: 'Floor Supervisor', description: 'Floor supervisor', isGlobal: 'N', permissions: [
    { capability: 'VIEW_USERS', scope: 'Team' }
  ] },
  { role: 'Escalations Manager', description: 'Escalations oversight', isGlobal: 'N', permissions: [
    { capability: 'VIEW_USERS', scope: 'Campaign' },
    { capability: 'TRANSFER_USERS', scope: 'Campaign' }
  ] },
  { role: 'Client Success Manager', description: 'Client delivery partner', isGlobal: 'N', permissions: [
    { capability: 'VIEW_USERS', scope: 'Campaign' },
    { capability: 'VIEW_AUDIT', scope: 'Campaign' }
  ] },
  { role: 'Compliance Manager', description: 'Compliance oversight', isGlobal: 'N', permissions: [
    { capability: 'VIEW_USERS', scope: 'Campaign' },
    { capability: 'MANAGE_POLICIES', scope: 'Campaign' },
    { capability: 'VIEW_AUDIT', scope: 'Global' }
  ] },
  { role: 'IT Support Manager', description: 'IT device support', isGlobal: 'N', permissions: [
    { capability: 'VIEW_USERS', scope: 'Campaign' },
    { capability: 'MANAGE_EQUIPMENT', scope: 'Campaign' }
  ] },
  { role: 'Reporting Analyst / Metrics Lead', description: 'Reporting & analytics', isGlobal: 'N', permissions: [
    { capability: 'VIEW_USERS', scope: 'Campaign' },
    { capability: 'VIEW_AUDIT', scope: 'Campaign' }
  ] },
  { role: 'Campaign Manager', description: 'Primary campaign manager', isGlobal: 'N', permissions: [
    { capability: 'VIEW_USERS', scope: 'Campaign' },
    { capability: 'MANAGE_USERS', scope: 'Campaign' },
    { capability: 'ASSIGN_ROLES', scope: 'Campaign' },
    { capability: 'TRANSFER_USERS', scope: 'Campaign' },
    { capability: 'TERMINATE_USERS', scope: 'Campaign' },
    { capability: 'MANAGE_EQUIPMENT', scope: 'Campaign' },
    { capability: 'VIEW_AUDIT', scope: 'Campaign' }
  ] },
  { role: 'Guest (Client Owner)', description: 'Read-only client access', isGlobal: 'N', permissions: [
    { capability: 'VIEW_USERS', scope: 'Campaign' }
  ] }
];

const IDENTITY_CAMPAIGN_SEED = [
  { CampaignId: 'lumina-hq', Name: 'Lumina HQ', Status: 'Active', ClientOwnerEmail: 'executive@lumina.com' },
  { CampaignId: 'credit-suite', Name: 'Credit Suite', Status: 'Active', ClientOwnerEmail: 'client@creditsuite.com' }
];

function seedLuminaIdentity() {
  if (typeof IdentityRepository === 'undefined' || typeof AuthService === 'undefined') {
    throw new Error('Load IdentityRepository and AuthService before seeding identity data.');
  }
  var utilitiesService = (typeof globalThis !== 'undefined' && globalThis.Utilities) ? globalThis.Utilities
    : (typeof Utilities !== 'undefined' ? Utilities : null);
  if (!utilitiesService) {
    throw new Error('Utilities service unavailable');
  }
  var now = new Date().toISOString();

  IDENTITY_CAMPAIGN_SEED.forEach(function(campaign) {
    IdentityRepository.upsert('Campaigns', 'CampaignId', Object.assign({
      CreatedAt: now,
      SettingsJSON: '{}'
    }, campaign));
  });

  IDENTITY_ROLE_SEED.forEach(function(roleSeed) {
    IdentityRepository.upsert('Roles', 'Role', {
      Role: roleSeed.role,
      Description: roleSeed.description,
      IsGlobal: roleSeed.isGlobal
    });
    roleSeed.permissions.forEach(function(permission) {
      IdentityRepository.upsert('RolePermissions', 'PermissionId', {
        PermissionId: roleSeed.role + '::' + permission.capability + '::' + permission.scope,
        Role: roleSeed.role,
        Capability: permission.capability,
        Scope: permission.scope,
        Allowed: 'Y'
      });
    });
  });

  var adminEmail = 'identity.admin@lumina.com';
  var existingAdmin = IdentityRepository.find('Users', function(row) {
    return row.Email === adminEmail;
  });
  var tempPassword = 'ChangeMe!1!';
  var adminId = existingAdmin ? existingAdmin.UserId : utilitiesService.getUuid();
  var adminRecord = {
    UserId: adminId,
    Email: adminEmail,
    Username: 'lumina.identity',
    PasswordHash: AuthService.hashPassword(tempPassword),
    EmailVerified: 'Y',
    TOTPEnabled: 'N',
    TOTPSecretHash: '',
    Status: 'Active',
    LastLoginAt: '',
    CreatedAt: now
  };
  IdentityRepository.upsert('Users', 'UserId', adminRecord);

  var assignment = {
    AssignmentId: utilitiesService.getUuid(),
    UserId: adminId,
    CampaignId: 'lumina-hq',
    Role: 'System Admin',
    IsPrimary: 'Y',
    AddedBy: 'seed',
    AddedAt: now,
    Watchlist: 'N'
  };
  IdentityRepository.upsert('UserCampaigns', 'AssignmentId', assignment);

  var employmentExists = IdentityRepository.list('EmploymentStatus').some(function(row) {
    return row.UserId === adminId && row.CampaignId === 'lumina-hq' && row.State === 'Active';
  });
  if (!employmentExists) {
    IdentityRepository.append('EmploymentStatus', {
      UserId: adminId,
      CampaignId: 'lumina-hq',
      State: 'Active',
      EffectiveDate: now,
      Reason: 'Seed data',
      Notes: 'Seeded system administrator'
    });
  }

  return {
    adminEmail: adminEmail,
    tempPassword: tempPassword
  };
}


/**
 * Public entry point. Returns a structured summary of what was ensured.
 */
function seedDefaultData() {
  const summary = {
    roles: { created: [], existing: [] },
    campaigns: { created: [], updated: [], existing: [], errors: [] },
    systemPages: { initialized: false, added: 0, updated: 0, total: 0 },
    navigation: {},
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

    summary.systemPages = ensureSystemPageCatalog();
    const pageCatalog = resolveSeedPageCatalog();
    summary.navigation = ensureCampaignNavigationSeeds(campaignIdsByName, pageCatalog);

    const luminaAdminInfo = ensureLuminaAdminUser(roleIdsByName, campaignIdsByName, pageCatalog);
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
  const existingRecords = loadCampaignSeedRecords();
  const existingIndex = {};

  existingRecords.forEach(record => {
    const normalized = normalizeKey(record.name);
    if (normalized && !existingIndex[normalized]) {
      existingIndex[normalized] = record;
    }
  });

  SEED_CAMPAIGNS.forEach(seedDefinition => {
    const desired = normalizeCampaignSeedDefinition(seedDefinition);
    const normalizedName = normalizeKey(desired.name);
    if (!normalizedName) {
      return;
    }

    const existing = existingIndex[normalizedName];
    if (existing && existing.id) {
      summary.campaigns.existing.push(desired.name);
      const syncResult = synchronizeCampaignSeed(existing, desired);
      if (syncResult.updated) {
        summary.campaigns.updated.push(desired.name);
      }
      if (syncResult.error) {
        if (!summary.campaigns.errors) summary.campaigns.errors = [];
        summary.campaigns.errors.push({ name: desired.name, error: syncResult.error });
      }
      return;
    }

    if (typeof csCreateCampaign !== 'function') {
      throw new Error('CampaignService.csCreateCampaign is not available');
    }

    const creationOptions = {
      clientName: desired.clientName,
      status: desired.status,
      channel: desired.channel,
      timezone: desired.timezone,
      slaTier: desired.slaTier,
      deletedAt: desired.deletedAt
    };

    const result = csCreateCampaign(desired.name, desired.description || '', creationOptions);
    if (result && result.success) {
      summary.campaigns.created.push(desired.name);
    } else {
      const errorMessage = result && result.error ? result.error : 'Unknown campaign creation error';
      if (/already exists/i.test(errorMessage)) {
        summary.campaigns.existing.push(desired.name);
      } else {
        summary.campaigns.existing.push(desired.name);
        if (!summary.campaigns.errors) summary.campaigns.errors = [];
        summary.campaigns.errors.push({ name: desired.name, error: errorMessage });
      }
    }
  });

  // Refresh to pick up any IDs assigned during creation.
  return getCampaignsIndex(true);
}

function normalizeCampaignSeedDefinition(definition) {
  const name = definition && definition.name ? String(definition.name).trim() : '';
  const description = definition && definition.description ? String(definition.description).trim() : '';
  const clientName = definition && definition.clientName ? String(definition.clientName).trim() : name;
  const status = definition && definition.status ? String(definition.status).trim() : 'Active';
  const channel = definition && definition.channel ? String(definition.channel).trim() : 'Operations';
  const timezone = definition && definition.timezone ? String(definition.timezone).trim() : 'UTC';
  const slaTier = definition && definition.slaTier ? String(definition.slaTier).trim() : 'Standard';
  const deletedAt = definition && Object.prototype.hasOwnProperty.call(definition, 'deletedAt')
    ? (definition.deletedAt || '')
    : '';

  return { name, description, clientName, status, channel, timezone, slaTier, deletedAt };
}

function loadCampaignSeedRecords() {
  if (typeof readSheet !== 'function') {
    return [];
  }

  const rows = readSheet(CAMPAIGNS_SHEET) || [];
  if (!Array.isArray(rows)) {
    return [];
  }

  return rows.map(projectCampaignRow).filter(record => record.name);
}

function projectCampaignRow(row) {
  const record = row || {};
  return {
    id: record.ID || record.Id || record.id || '',
    name: record.Name || record.name || '',
    description: record.Description || record.description || '',
    clientName: record.ClientName || record.clientName || '',
    status: record.Status || record.status || '',
    channel: record.Channel || record.channel || '',
    timezone: record.Timezone || record.timezone || '',
    slaTier: record.SlaTier || record.slaTier || '',
    createdAt: record.CreatedAt || record.createdAt || record.Created || '',
    updatedAt: record.UpdatedAt || record.updatedAt || record.Updated || '',
    deletedAt: record.DeletedAt || record.deletedAt || ''
  };
}

function synchronizeCampaignSeed(existing, desired) {
  const updates = {};

  if (!valuesEqual(existing.description, desired.description)) {
    updates.Description = desired.description || '';
  }
  if (!valuesEqual(existing.clientName, desired.clientName)) {
    updates.ClientName = desired.clientName || '';
  }
  if (!valuesEqual(existing.status, desired.status)) {
    updates.Status = desired.status || '';
  }
  if (!valuesEqual(existing.channel, desired.channel)) {
    updates.Channel = desired.channel || '';
  }
  if (!valuesEqual(existing.timezone, desired.timezone)) {
    updates.Timezone = desired.timezone || '';
  }
  if (!valuesEqual(existing.slaTier, desired.slaTier)) {
    updates.SlaTier = desired.slaTier || '';
  }
  if (!valuesEqual(existing.deletedAt, desired.deletedAt)) {
    updates.DeletedAt = desired.deletedAt || '';
  }

  if (!existing.createdAt) {
    updates.CreatedAt = existing.updatedAt || new Date();
  }

  if (!Object.keys(updates).length) {
    return { updated: false };
  }

  updates.UpdatedAt = new Date();

  const applyResult = applyCampaignUpdates(existing.id, updates);
  if (!applyResult.success) {
    return { updated: false, error: applyResult.error || 'Failed to update campaign metadata.' };
  }

  return { updated: true };
}

function valuesEqual(a, b) {
  const normalizedA = normalizeCampaignValue(a);
  const normalizedB = normalizeCampaignValue(b);
  return normalizedA === normalizedB;
}

function normalizeCampaignValue(value) {
  if (value === null || typeof value === 'undefined') {
    return '';
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  const stringValue = String(value).trim();
  return stringValue;
}

function applyCampaignUpdates(campaignId, updates) {
  if (!campaignId) {
    return { success: false, error: 'Campaign ID is required for updates.' };
  }

  if (!updates || !Object.keys(updates).length) {
    return { success: true, updated: false };
  }

  if (typeof SpreadsheetApp === 'undefined' || !SpreadsheetApp || !SpreadsheetApp.getActiveSpreadsheet) {
    return { success: false, error: 'SpreadsheetApp is not available.' };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CAMPAIGNS_SHEET);
  if (!sheet) {
    return { success: false, error: 'Campaigns sheet not found.' };
  }

  const data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) {
    return { success: false, error: 'Campaigns data is empty.' };
  }

  const headers = data[0];
  const idIndex = headers.indexOf('ID');
  if (idIndex === -1) {
    return { success: false, error: 'Campaigns sheet is missing an ID column.' };
  }

  for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
    if (String(data[rowIndex][idIndex]) !== String(campaignId)) {
      continue;
    }

    Object.keys(updates).forEach(columnName => {
      const columnIndex = headers.indexOf(columnName);
      if (columnIndex === -1) {
        return;
      }
      sheet.getRange(rowIndex + 1, columnIndex + 1).setValue(updates[columnName]);
    });

    if (!Object.prototype.hasOwnProperty.call(updates, 'UpdatedAt')) {
      const updatedIndex = headers.indexOf('UpdatedAt');
      if (updatedIndex !== -1) {
        sheet.getRange(rowIndex + 1, updatedIndex + 1).setValue(new Date());
      }
    }

    if (typeof commitChanges === 'function') {
      commitChanges();
    }

    if (typeof clearCampaignCaches === 'function') {
      clearCampaignCaches(campaignId);
    }

    return { success: true, updated: true };
  }

  return { success: false, error: 'Campaign row not found for updates.' };
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
 * Ensure there is a Lumina admin user with administrative permissions.
 */
function ensureLuminaAdminUser(roleIdsByName, campaignIdsByName, pageCatalog) {
  return ensureSeedAdministrator(SEED_LUMINA_ADMIN_PROFILE, roleIdsByName, campaignIdsByName, pageCatalog);
}

/**
 * Ensure system pages are synchronized so campaign navigation can be seeded accurately.
 */
function ensureSystemPageCatalog() {
  const result = { initialized: false, added: 0, updated: 0, total: 0 };

  try {
    ensureSheetWithHeaders(PAGES_SHEET, PAGES_HEADERS);

    if (typeof initializeEnhancedSystemPages === 'function') {
      try {
        initializeEnhancedSystemPages();
        result.initialized = true;
      } catch (initError) {
        console.warn('initializeEnhancedSystemPages during seeding failed:', initError);
      }
    }

    if (typeof enhancedAutoDiscoverAndSavePages === 'function') {
      try {
        const discovery = enhancedAutoDiscoverAndSavePages({ force: true, minIntervalSec: 0 });
        if (discovery) {
          if (discovery.skipped) {
            result.skipped = true;
          }
          if (discovery.success === false) {
            result.error = discovery.error || 'Unknown discovery error';
          } else {
            result.added = discovery.added || 0;
            result.updated = discovery.updated || 0;
            result.total = discovery.total || 0;
          }
        }
      } catch (discoveryError) {
        console.warn('enhancedAutoDiscoverAndSavePages during seeding failed:', discoveryError);
        result.error = discoveryError && discoveryError.message ? discoveryError.message : String(discoveryError);
      }
    }

    if (!result.total) {
      const rows = (typeof readSheet === 'function') ? (readSheet(PAGES_SHEET) || []) : [];
      result.total = Array.isArray(rows) ? rows.length : 0;
    }
  } catch (error) {
    console.error('ensureSystemPageCatalog error:', error);
    if (typeof writeError === 'function') {
      writeError('ensureSystemPageCatalog', error);
    }
    result.error = error && error.message ? error.message : String(error);
  }

  return result;
}

/**
 * Ensure every seeded campaign receives default categories and page assignments.
 */
function ensureCampaignNavigationSeeds(campaignIdsByName, providedPageCatalog) {
  const result = {};

  try {
    if (!campaignIdsByName || !Object.keys(campaignIdsByName).length) {
      return result;
    }

    const pageCatalog = Array.isArray(providedPageCatalog) && providedPageCatalog.length
      ? providedPageCatalog
      : resolveSeedPageCatalog();
    const categoryDefinitions = resolveSeedCategoryDefinitions(pageCatalog);
    const seededCampaignNames = Array.isArray(SEED_CAMPAIGNS)
      ? SEED_CAMPAIGNS.map(c => c && c.name).filter(Boolean)
      : [];
    const processedNames = new Set();

    seededCampaignNames.forEach(campaignName => {
      const normalizedName = normalizeKey(campaignName);
      if (!normalizedName || processedNames.has(normalizedName)) {
        return;
      }

      processedNames.add(normalizedName);
      const campaignId = resolveCampaignIdByName(campaignIdsByName, campaignName);

      if (!campaignId) {
        result[campaignName] = { error: 'Campaign not found during navigation seeding.' };
        return;
      }

      result[campaignName] = ensureCampaignNavigationForCampaign(
        campaignId,
        campaignName,
        pageCatalog,
        categoryDefinitions
      );
    });
  } catch (error) {
    console.error('ensureCampaignNavigationSeeds error:', error);
    if (typeof writeError === 'function') {
      writeError('ensureCampaignNavigationSeeds', error);
    }
  }

  return result;
}

function ensureCampaignNavigationForCampaign(campaignId, campaignName, pageCatalog, categoryDefinitions) {
  const summary = {
    categories: { created: [], existing: [] },
    pages: { created: [], updated: [], existing: [] }
  };

  try {
    const categoryResult = ensureCampaignCategoryRecords(campaignId, campaignName, categoryDefinitions);
    summary.categories.created = categoryResult.created || [];
    summary.categories.existing = categoryResult.existing || [];
    if (categoryResult.errors && categoryResult.errors.length) {
      summary.categories.errors = categoryResult.errors;
    }

    const pageResult = ensureCampaignPageRecords(
      campaignId,
      campaignName,
      pageCatalog,
      categoryResult.idsByName || {},
      categoryDefinitions
    );
    summary.pages.created = pageResult.created || [];
    summary.pages.updated = pageResult.updated || [];
    summary.pages.existing = pageResult.existing || [];
    if (pageResult.errors && pageResult.errors.length) {
      summary.pages.errors = pageResult.errors;
    }

    if (categoryResult.changed || pageResult.changed) {
      if (typeof clearCampaignCaches === 'function') {
        try { clearCampaignCaches(campaignId); } catch (cacheError) { console.warn('clearCampaignCaches during seeding failed:', cacheError); }
      }
    }

    if (typeof csRefreshNavigation === 'function') {
      try { csRefreshNavigation(campaignId); } catch (navError) { console.warn('csRefreshNavigation during seeding failed:', navError); }
    } else if (typeof forceRefreshCampaignNavigation === 'function') {
      try { forceRefreshCampaignNavigation(campaignId); } catch (navError) { console.warn('forceRefreshCampaignNavigation during seeding failed:', navError); }
    }
  } catch (error) {
    console.error('ensureCampaignNavigationForCampaign error:', error);
    if (typeof writeError === 'function') {
      writeError('ensureCampaignNavigationForCampaign', error);
    }
    summary.error = error && error.message ? error.message : String(error);
  }

  return summary;
}

function ensureCampaignCategoryRecords(campaignId, campaignName, categoryDefinitions) {
  const result = { created: [], existing: [], errors: [], idsByName: {}, changed: false };

  try {
    ensureSheetWithHeaders(PAGE_CATEGORIES_SHEET, PAGE_CATEGORIES_HEADERS);
    const existingCategories = loadCampaignCategoryRows(campaignId);
    const existingMap = {};

    existingCategories.forEach(cat => {
      const name = cat.categoryName || cat.CategoryName;
      const id = cat.id || cat.ID;
      const normalized = normalizeKey(name);
      if (normalized) {
        existingMap[normalized] = id;
        result.idsByName[normalized] = id;
      }
      if (name && id) {
        result.idsByName[name] = id;
      }
    });

    const entries = Object.entries(categoryDefinitions || {}).sort((a, b) => {
      const sortA = parseInt(a[1] && a[1].sortOrder, 10) || 999;
      const sortB = parseInt(b[1] && b[1].sortOrder, 10) || 999;
      if (sortA === sortB) return a[0].localeCompare(b[0]);
      return sortA - sortB;
    });

    entries.forEach(([name, meta], index) => {
      const normalized = normalizeKey(name);
      if (!normalized) return;

      if (existingMap[normalized]) {
        result.existing.push(name);
        return;
      }

      const icon = normalizeCategoryIcon(meta && meta.icon);
      const sortOrder = parseInt(meta && meta.sortOrder, 10) || ((index + 1) * 10);
      let createResult = null;

      if (typeof csCreateCategory === 'function') {
        createResult = csCreateCategory(campaignId, name, icon, sortOrder);
      }

      if (!createResult || createResult.success === false) {
        createResult = addCategoryRowDirect(campaignId, name, icon, sortOrder);
      }

      if (createResult && createResult.success) {
        result.created.push(name);
        result.changed = true;
      } else if (createResult && /exists/i.test(createResult.error || '')) {
        result.existing.push(name);
      } else {
        result.errors.push({ category: name, error: (createResult && createResult.error) || 'Unknown error' });
      }
    });

    const refreshed = loadCampaignCategoryRows(campaignId);
    refreshed.forEach(cat => {
      const name = cat.categoryName || cat.CategoryName;
      const id = cat.id || cat.ID;
      const normalized = normalizeKey(name);
      if (name && id) {
        result.idsByName[name] = id;
      }
      if (normalized) {
        result.idsByName[normalized] = id;
      }
    });
  } catch (error) {
    console.error('ensureCampaignCategoryRecords error:', error);
    if (typeof writeError === 'function') {
      writeError('ensureCampaignCategoryRecords', error);
    }
    result.errors.push({ error: error && error.message ? error.message : String(error) });
  }

  if (!result.errors.length) delete result.errors;
  return result;
}

function ensureCampaignPageRecords(campaignId, campaignName, pageCatalog, categoryIdsByName, categoryDefinitions) {
  const result = { created: [], updated: [], existing: [], errors: [], changed: false };

  try {
    ensureSheetWithHeaders(CAMPAIGN_PAGES_SHEET, CAMPAIGN_PAGES_HEADERS);
    const existingPages = loadCampaignPageRows(campaignId);
    const existingMap = {};

    existingPages.forEach(page => {
      const key = page.pageKey || page.PageKey;
      const normalized = normalizeKey(key);
      if (normalized && !existingMap[normalized]) {
        existingMap[normalized] = page;
      }
    });

    const systemPages = (typeof readSheet === 'function') ? (readSheet(PAGES_SHEET) || []) : [];
    const systemPageMap = {};
    (systemPages || []).forEach(page => {
      const normalized = normalizeKey(page.PageKey);
      if (normalized) {
        systemPageMap[normalized] = page;
      }
    });

    const categoryCounters = {};
    const processedKeys = new Set();

    (pageCatalog || []).forEach(page => {
      const key = page && page.key ? String(page.key).trim() : '';
      if (!key) return;

      const normalizedKey = key.toLowerCase();
      if (processedKeys.has(normalizedKey)) {
        return;
      }
      processedKeys.add(normalizedKey);

      const categoryName = page.category || 'General';
      const categoryKey = normalizeKey(categoryName);
      const categoryMeta = categoryDefinitions[categoryName] || categoryDefinitions[categoryKey] || categoryDefinitions['General'] || {};
      const categoryId = categoryIdsByName[categoryKey] || categoryIdsByName[categoryName] || categoryIdsByName['general'] || categoryIdsByName['General'] || '';

      const indexWithinCategory = (categoryCounters[categoryKey] || 0) + 1;
      categoryCounters[categoryKey] = indexWithinCategory;
      const desiredSortOrder = computeCategorySortOrder(categoryMeta, indexWithinCategory);

      const systemPage = systemPageMap[normalizedKey] || {};
      const title = page.title || systemPage.PageTitle || inferPageTitleFromKey(key);
      let icon = normalizePageIcon(page.icon, key);
      if (!icon && systemPage.PageIcon) {
        icon = normalizePageIcon(systemPage.PageIcon, key);
      }

      const existing = existingMap[normalizedKey];
      if (existing) {
        const updates = {};
        const existingTitle = existing.pageTitle || existing.PageTitle || '';
        const existingIcon = existing.pageIcon || existing.PageIcon || '';
        const existingCategory = existing.categoryId || existing.CategoryID || '';
        const existingSortOrder = parseInt(existing.sortOrder || existing.SortOrder, 10) || 0;

        if (existingTitle !== title) {
          updates.PageTitle = title;
        }
        if (existingIcon !== icon) {
          updates.PageIcon = icon;
        }
        const normalizedExistingCategory = existingCategory ? String(existingCategory) : '';
        const normalizedDesiredCategory = categoryId ? String(categoryId) : '';
        if (normalizedExistingCategory !== normalizedDesiredCategory) {
          updates.CategoryID = categoryId || '';
        }
        if (existingSortOrder !== desiredSortOrder) {
          updates.SortOrder = desiredSortOrder;
        }

        if (Object.keys(updates).length > 0) {
          let updateResult = null;
          if (typeof csUpdateCampaignPage === 'function') {
            updateResult = csUpdateCampaignPage(existing.id || existing.ID, updates);
          }

          if (!updateResult || updateResult.success === false) {
            updateResult = updateCampaignPageRowDirect(existing.id || existing.ID || existing.Id, updates);
          }

          if (updateResult && updateResult.success) {
            result.updated.push(key);
            result.changed = true;
          } else if (updateResult && updateResult.error) {
            result.errors.push({ pageKey: key, error: updateResult.error });
          }
        } else {
          result.existing.push(key);
        }

        return;
      }

      let createResult = null;
      if (typeof csAddPageToCampaign === 'function') {
        createResult = csAddPageToCampaign(campaignId, key, title, icon, categoryId || null, desiredSortOrder);
      }

      if (!createResult || createResult.success === false) {
        createResult = addCampaignPageRowDirect(campaignId, {
          pageKey: key,
          pageTitle: title,
          pageIcon: icon,
          categoryId: categoryId || '',
          sortOrder: desiredSortOrder
        });
      }

      if (createResult && createResult.success) {
        result.created.push(key);
        result.changed = true;
      } else if (createResult && /assigned/i.test(createResult.error || '')) {
        result.existing.push(key);
      } else {
        result.errors.push({ pageKey: key, error: (createResult && createResult.error) || 'Unknown error' });
      }
    });
  } catch (error) {
    console.error('ensureCampaignPageRecords error:', error);
    if (typeof writeError === 'function') {
      writeError('ensureCampaignPageRecords', error);
    }
    result.errors.push({ error: error && error.message ? error.message : String(error) });
  }

  if (!result.errors.length) delete result.errors;
  return result;
}

function resolveSeedPageCatalog() {
  let pages = [];

  try {
    if (typeof getAllPagesFromActualRouting === 'function') {
      pages = getAllPagesFromActualRouting() || [];
    }
  } catch (error) {
    console.warn('getAllPagesFromActualRouting during seeding failed:', error);
  }

  if (!Array.isArray(pages) || !pages.length) {
    pages = [{
      key: 'dashboard',
      title: 'Dashboard',
      icon: 'fas fa-tachometer-alt',
      description: 'Primary dashboard overview',
      category: 'Dashboard & Analytics'
    }];
  }

  const seen = new Set();
  const normalized = [];

  pages.forEach(page => {
    if (!page || !page.key) return;
    const key = String(page.key).trim();
    if (!key) return;

    const normalizedKey = key.toLowerCase();
    if (seen.has(normalizedKey)) {
      return;
    }
    seen.add(normalizedKey);

    normalized.push({
      key,
      title: page.title || inferPageTitleFromKey(key),
      icon: normalizePageIcon(page.icon, key),
      description: page.description || '',
      category: page.category || 'General',
      requiresAdmin: page.requiresAdmin === true,
      isPublic: page.isPublic === true
    });
  });

  normalized.sort((a, b) => {
    const categoryCompare = a.category.localeCompare(b.category);
    if (categoryCompare !== 0) return categoryCompare;
    return a.title.localeCompare(b.title);
  });

  return normalized;
}

function resolveSeedCategoryDefinitions(pageCatalog) {
  const definitions = {};

  try {
    if (typeof getEnhancedPageCategories === 'function') {
      const base = getEnhancedPageCategories();
      Object.keys(base || {}).forEach(name => {
        if (!name) return;
        const meta = base[name] || {};
        definitions[name] = {
          icon: normalizeCategoryIcon(meta.icon),
          description: meta.description || '',
          sortOrder: meta.sortOrder || meta.order || meta.position || 999
        };
      });
    }
  } catch (error) {
    console.warn('getEnhancedPageCategories during seeding failed:', error);
  }

  (pageCatalog || []).forEach(page => {
    if (!page || !page.category) return;
    if (!definitions[page.category]) {
      definitions[page.category] = {
        icon: normalizeCategoryIcon('fas fa-folder'),
        description: '',
        sortOrder: 999
      };
    }
  });

  if (!definitions.General) {
    definitions.General = {
      icon: normalizeCategoryIcon('fas fa-folder-open'),
      description: 'General purpose pages and utilities',
      sortOrder: 999
    };
  }

  const names = Object.keys(definitions).sort((a, b) => {
    const sortA = parseInt(definitions[a].sortOrder, 10) || 999;
    const sortB = parseInt(definitions[b].sortOrder, 10) || 999;
    if (sortA === sortB) return a.localeCompare(b);
    return sortA - sortB;
  });

  names.forEach((name, index) => {
    const meta = definitions[name];
    meta.icon = normalizeCategoryIcon(meta.icon);
    const parsedSort = parseInt(meta.sortOrder, 10);
    meta.sortOrder = Number.isFinite(parsedSort) ? parsedSort : ((index + 1) * 10);
  });

  return definitions;
}

function resolveCampaignIdByName(map, name) {
  if (!map || !name) return '';
  if (map[name]) return map[name];
  const normalized = normalizeKey(name);
  if (normalized && map[normalized]) return map[normalized];
  return '';
}

function normalizeKey(value) {
  if (value === null || typeof value === 'undefined') {
    return '';
  }
  return String(value).trim().toLowerCase();
}

function normalizePageIcon(icon, key) {
  let resolved = icon;
  if (!resolved && typeof suggestIconForPageKey === 'function') {
    try { resolved = suggestIconForPageKey(key); } catch (error) { console.warn('suggestIconForPageKey during seeding failed:', error); }
  }
  if (!resolved) {
    return 'fas fa-file';
  }
  resolved = String(resolved).trim();
  if (/^(fas|far|fal|fad|fab)\s+fa-/.test(resolved)) {
    return resolved;
  }
  if (/^fa-/.test(resolved)) {
    return 'fas ' + resolved;
  }
  return resolved;
}

function normalizeCategoryIcon(icon) {
  if (!icon) {
    return 'fas fa-folder';
  }
  const value = String(icon).trim();
  if (/^(fas|far|fal|fad|fab)\s+fa-/.test(value)) {
    return value;
  }
  if (/^fa-/.test(value)) {
    return 'fas ' + value;
  }
  return value;
}

function inferPageTitleFromKey(key) {
  const value = String(key || '')
    .replace(/[-_]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!value) {
    return 'Untitled Page';
  }
  return value.replace(/\b([a-z])/gi, (_, ch) => ch.toUpperCase());
}

function computeCategorySortOrder(meta, indexWithinCategory) {
  const base = parseInt(meta && meta.sortOrder, 10);
  const normalizedBase = Number.isFinite(base) ? base : 999;
  const offset = Number(indexWithinCategory) || 1;
  return (normalizedBase * 100) + offset;
}

function loadCampaignCategoryRows(campaignId) {
  try {
    if (typeof csGetCampaignCategories === 'function') {
      return (csGetCampaignCategories(campaignId) || [])
        .map(cat => ({
          id: cat.id || cat.ID,
          categoryName: cat.categoryName || cat.CategoryName,
          sortOrder: cat.sortOrder || cat.SortOrder
        }))
        .filter(cat => cat.id && cat.categoryName);
    }

    const rows = (typeof readSheet === 'function') ? (readSheet(PAGE_CATEGORIES_SHEET) || []) : [];
    return rows
      .filter(row => row && String(row.CampaignID) === String(campaignId) && isRowActive(row.IsActive))
      .map(row => ({ id: row.ID, categoryName: row.CategoryName, sortOrder: row.SortOrder }))
      .filter(cat => cat.id && cat.categoryName);
  } catch (error) {
    console.warn('loadCampaignCategoryRows error:', error);
    return [];
  }
}

function loadCampaignPageRows(campaignId) {
  try {
    if (typeof csGetCampaignPages === 'function') {
      return (csGetCampaignPages(campaignId) || []).map(page => ({
        id: page.id || page.ID,
        pageKey: page.pageKey || page.PageKey,
        pageTitle: page.pageTitle || page.PageTitle,
        pageIcon: page.pageIcon || page.PageIcon,
        categoryId: page.categoryId || page.CategoryID,
        sortOrder: page.sortOrder || page.SortOrder
      }));
    }

    const rows = (typeof readSheet === 'function') ? (readSheet(CAMPAIGN_PAGES_SHEET) || []) : [];
    return rows
      .filter(row => row && String(row.CampaignID) === String(campaignId) && isRowActive(row.IsActive))
      .map(row => ({
        id: row.ID,
        pageKey: row.PageKey,
        pageTitle: row.PageTitle,
        pageIcon: row.PageIcon,
        categoryId: row.CategoryID,
        sortOrder: row.SortOrder
      }));
  } catch (error) {
    console.warn('loadCampaignPageRows error:', error);
    return [];
  }
}

function addCategoryRowDirect(campaignId, name, icon, sortOrder) {
  try {
    const sheet = ensureSheetWithHeaders(PAGE_CATEGORIES_SHEET, PAGE_CATEGORIES_HEADERS);
    const id = Utilities.getUuid();
    const now = new Date();
    sheet.appendRow([id, campaignId, name, icon, parseInt(sortOrder, 10) || 999, true, now, now]);
    if (typeof commitChanges === 'function') {
      commitChanges();
    }
    return { success: true, data: { id } };
  } catch (error) {
    console.error('addCategoryRowDirect error:', error);
    if (typeof writeError === 'function') {
      writeError('addCategoryRowDirect', error);
    }
    return { success: false, error: error && error.message ? error.message : String(error) };
  }
}

function addCampaignPageRowDirect(campaignId, details) {
  try {
    const sheet = ensureSheetWithHeaders(CAMPAIGN_PAGES_SHEET, CAMPAIGN_PAGES_HEADERS);
    const id = Utilities.getUuid();
    const now = new Date();

    sheet.appendRow([
      id,
      campaignId,
      details.pageKey,
      details.pageTitle,
      details.pageIcon,
      details.categoryId || '',
      parseInt(details.sortOrder, 10) || 999,
      true,
      now,
      now
    ]);

    if (typeof commitChanges === 'function') {
      commitChanges();
    }

    return { success: true, data: { id } };
  } catch (error) {
    console.error('addCampaignPageRowDirect error:', error);
    if (typeof writeError === 'function') {
      writeError('addCampaignPageRowDirect', error);
    }
    return { success: false, error: error && error.message ? error.message : String(error) };
  }
}

function updateCampaignPageRowDirect(pageId, updates) {
  try {
    if (!pageId) {
      return { success: false, error: 'Page ID is required for updates.' };
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CAMPAIGN_PAGES_SHEET);
    if (!sheet) {
      return { success: false, error: 'Campaign pages sheet not found.' };
    }

    const data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) {
      return { success: false, error: 'Campaign pages data is empty.' };
    }

    const headers = data[0];
    const idIndex = headers.indexOf('ID');
    if (idIndex === -1) {
      return { success: false, error: 'Campaign pages sheet is missing an ID column.' };
    }

    for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
      if (String(data[rowIndex][idIndex]) === String(pageId)) {
        Object.keys(updates || {}).forEach(field => {
          const colIndex = headers.indexOf(field);
          if (colIndex === -1) return;

          let value = updates[field];
          if (field === 'SortOrder') {
            value = parseInt(value, 10) || 999;
          }
          if (field === 'CategoryID' && (value === null || typeof value === 'undefined')) {
            value = '';
          }

          sheet.getRange(rowIndex + 1, colIndex + 1).setValue(value);
        });

        const updatedAtIndex = headers.indexOf('UpdatedAt');
        if (updatedAtIndex !== -1) {
          sheet.getRange(rowIndex + 1, updatedAtIndex + 1).setValue(new Date());
        }

        if (typeof commitChanges === 'function') {
          commitChanges();
        }

        return { success: true };
      }
    }

    return { success: false, error: 'Campaign page row not found for update.' };
  } catch (error) {
    console.error('updateCampaignPageRowDirect error:', error);
    if (typeof writeError === 'function') {
      writeError('updateCampaignPageRowDirect', error);
    }
    return { success: false, error: error && error.message ? error.message : String(error) };
  }
}

function isRowActive(value) {
  if (typeof isActive === 'function') {
    return isActive(value);
  }
  if (value === true) return true;
  if (value === false || value === null || typeof value === 'undefined') return false;
  const normalized = String(value).trim().toUpperCase();
  if (!normalized) return false;
  return normalized === 'TRUE' || normalized === 'YES' || normalized === 'Y' || normalized === '1' || normalized === 'ON';
}

function ensureUserClaims(userId, claimTypes) {
  const result = { requested: [], created: [], existing: [] };
  if (!userId) {
    return result;
  }

  const requested = Array.from(new Set((Array.isArray(claimTypes) ? claimTypes : [])
    .map(type => String(type || '').trim())
    .filter(Boolean)));

  result.requested = requested.slice();

  if (!requested.length) {
    return result;
  }

  ensureSheetWithHeaders(USER_CLAIMS_SHEET, CLAIMS_HEADERS);

  const existingRows = (typeof readSheet === 'function') ? (readSheet(USER_CLAIMS_SHEET) || []) : [];
  const existingSet = new Set();

  existingRows.forEach(row => {
    if (!row) return;
    const rowUserId = row.UserId || row.UserID || row.User || row.userId || row.userid;
    if (String(rowUserId) !== String(userId)) return;
    const type = row.ClaimType || row.Type || row.Name || row.claimType;
    if (!type) return;
    existingSet.add(String(type).trim().toLowerCase());
  });

  if (typeof SpreadsheetApp === 'undefined' || !SpreadsheetApp || !SpreadsheetApp.getActiveSpreadsheet) {
    result.existing = requested.filter(type => existingSet.has(type.toLowerCase()));
    return result;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(USER_CLAIMS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(USER_CLAIMS_SHEET);
    sh.getRange(1, 1, 1, CLAIMS_HEADERS.length).setValues([CLAIMS_HEADERS]);
  }

  const lastColumn = sh.getLastColumn() || CLAIMS_HEADERS.length;
  const headers = sh.getRange(1, 1, 1, lastColumn).getValues()[0];
  const lookup = {};
  headers.forEach((header, idx) => {
    const key = String(header || '').trim();
    if (key && !Object.prototype.hasOwnProperty.call(lookup, key)) {
      lookup[key] = idx;
    }
  });

  const idxId = Object.prototype.hasOwnProperty.call(lookup, 'ID') ? lookup.ID : -1;
  const idxUserId = Object.prototype.hasOwnProperty.call(lookup, 'UserId')
    ? lookup.UserId
    : (Object.prototype.hasOwnProperty.call(lookup, 'UserID') ? lookup.UserID : -1);
  const idxClaimType = Object.prototype.hasOwnProperty.call(lookup, 'ClaimType') ? lookup.ClaimType : -1;
  const idxCreatedAt = Object.prototype.hasOwnProperty.call(lookup, 'CreatedAt') ? lookup.CreatedAt : -1;
  const idxUpdatedAt = Object.prototype.hasOwnProperty.call(lookup, 'UpdatedAt') ? lookup.UpdatedAt : -1;

  requested.forEach(type => {
    const normalized = type.toLowerCase();
    if (existingSet.has(normalized)) {
      result.existing.push(type);
      return;
    }

    const row = new Array(headers.length).fill('');
    const nowIso = new Date().toISOString();
    if (idxId >= 0) row[idxId] = Utilities.getUuid();
    if (idxUserId >= 0) row[idxUserId] = userId;
    if (idxClaimType >= 0) row[idxClaimType] = type;
    if (idxCreatedAt >= 0) row[idxCreatedAt] = nowIso;
    if (idxUpdatedAt >= 0) row[idxUpdatedAt] = nowIso;
    sh.appendRow(row);
    existingSet.add(normalized);
    result.created.push(type);
  });

  if (result.created.length && typeof invalidateCache === 'function') {
    invalidateCache(USER_CLAIMS_SHEET);
  }

  return result;
}

function summarizeCampaignAssignments(campaignIdsByName, assignedIds) {
  if (!Array.isArray(assignedIds) || !assignedIds.length) {
    return [];
  }

  const nameMap = buildCampaignIdNameMap(campaignIdsByName);
  const seen = new Set();
  const summary = [];

  assignedIds.forEach(id => {
    const key = String(id || '').trim();
    if (!key || seen.has(key)) {
      return;
    }
    seen.add(key);
    summary.push({ id: key, name: nameMap[key] || key });
  });

  return summary;
}

function buildCampaignIdNameMap(campaignIdsByName) {
  const map = {};

  if (Array.isArray(SEED_CAMPAIGNS)) {
    SEED_CAMPAIGNS.forEach(campaign => {
      if (!campaign || !campaign.name) return;
      const id = resolveCampaignIdByName(campaignIdsByName, campaign.name);
      if (id) {
        map[String(id)] = campaign.name;
      }
    });
  }

  if (campaignIdsByName) {
    Object.keys(campaignIdsByName).forEach(key => {
      const id = campaignIdsByName[key];
      if (!id) return;
      const normalizedKey = String(key || '').trim();
      if (!normalizedKey) return;
      const idKey = String(id);
      if (!map[idKey]) {
        map[idKey] = normalizedKey;
      }
    });
  }

  return map;
}

function ensureCanLoginFlag(userId, canLogin) {
  return;
}



