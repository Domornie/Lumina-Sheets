/**
 * SeedData.js
 * -----------------------------------------------------------------------------
 * Lightweight bootstrap seeded entirely through the public service APIs.
 *
 * Run seedDefaultData() once to ensure:
 *   • Core roles exist
 *   • A couple of starter campaigns are provisioned
 *   • A super administrator account is created (or refreshed) with a known login
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
  { name: 'Credit Suite', description: 'Credit Suite client workspace' }
];

const SEED_ADMIN_PROFILE = {
  userName: 'admin',
  fullName: 'Lumina Administrator',
  email: 'admin@vlbpo.com',
  password: 'ChangeMe123!',
  defaultCampaign: 'Lumina HQ',
  roleNames: ['Super Admin', 'Administrator'],
  seedLabel: 'Super Administrator'
};

const SEED_LUMINA_ADMIN_PROFILE = {
  userName: 'lumina.admin',
  fullName: 'Lumina Admin',
  email: 'lumina@vlbpo.com',
  password: 'ChangeMe123!',
  defaultCampaign: 'Lumina HQ',
  roleNames: ['Administrator'],
  seedLabel: 'Lumina Administrator'
};

const DECEMBER_QA_AGENTS = [
  'Sonny Cuerdo',
  'Yonique Byfield',
  'Gabrielle Howell',
  'Dwayne Gordan',
  'Alexi Martin',
  'Romaun Grant',
  'Terry Ann Green',
  'Erzulye Parkes',
  'Ayeishia Rotti',
  'Sherika Williams',
  'Mishaunna Morrison',
  'Shanara Thompson',
  'Neako Powell',
  'Ayesha Reddie',
  'Sashana Dixon',
  'Sashagay Wanchope',
  'Briana Lindo-Dixon',
  'Vinnelle Fyine',
  'Shanique Buchanan',
  'Shaenelle Francis'
];

const DECEMBER_QA_CLIENTS = [
  'CCBCU',
  'Protective',
  'Premier Transportation Resource Center',
  'CCBCU',
  'WellStreet Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU',
  'CCBCU',
  'CCBCU',
  'CCBCU',
  'Progress Rail Service Center',
  'Protective Life Corporation',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'MSHS',
  'Progress Rail Service Center',
  'U.S. Silica',
  'Progress Rail Service Center',
  'CCBCU Resource Center',
  'WellStreet Resource Center',
  'CCBCU',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Progress Rail Service Center',
  'Protective Life Corporation',
  'CCBCU Resource Center',
  'Protective Life Corporation',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Progress Rail Service Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Mayville Engineering Company',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Protective Life Corporation',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Protective Life Corporation',
  'Progress Rail Service Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Protective Life Corporation',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Protective Life Corporation',
  'Progress Rail Service Center',
  'Premier Transportation',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Protective Life Corporation',
  'Mayville Engineering Company',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Progress Rail Service Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Progress Rail Service Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Progress Rail Service Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Progress Rail Service Center',
  'CCBCU Resource Center',
  'Protective Life Corporation',
  'CCBCU Resource Center',
  'Premier Transportation',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Mayville Engineering Company',
  'Progress Rail Service Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Progress Rail Service Center',
  'Protective Life Corporation',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Premier Transportation',
  'CCBCU Resource Center',
  'Protective Life Corporation',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Progress Rail Service Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Protective Life Corporation',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Columbus School District',
  'Columbus City Schools',
  'Columbus City Schools',
  'Columbus City Schools',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Premier Transportation',
  'Protective Life Corporation',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Protective Life Corporation',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Parma City School District',
  'Columbus School District',
  'Premier Transportation',
  'CCBCU Resource Center',
  'Columbus City Schools',
  'Premier Transportation',
  'Columbus City Schools',
  'Columbus City Schools',
  'CCBCU Resource Center',
  'Premier Transportation',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Protective Life Corporation',
  'Columbus City Schools',
  'Protective Life Corporation',
  'Protective Life Corporation',
  'Columbus School District',
  'Columbus City Schools',
  'Columbus City Schools',
  'CCBCU Resource Center',
  'Progress Rail Service Center',
  'Columbus City Schools',
  'Wabash',
  'CCBCU Resource Center',
  'Parma City School District',
  'Protective Life Corporation',
  'Premier Transportation',
  'Wabash',
  'Columbus City Schools',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Columbus School District',
  'Columbus City Schools',
  'Columbus City Schools',
  'Wabash',
  'Columbus City Schools',
  'Columbus City Schools',
  'Columbus City Schools',
  'CCBCU Resource Center',
  'Premier Transportation',
  'Parma City School District',
  'Premier Transportation',
  'Columbus City Schools',
  'Columbus City Schools',
  'Columbus City Schools',
  'Columbus City Schools',
  'Columbus City Schools',
  'Premier Transportation',
  'Premier Transportation',
  'Premier Transportation',
  'Columbus City Schools',
  'Columbus City Schools',
  'Columbus City Schools',
  'Premier Transportation',
  'Premier Transportation',
  'Columbus City Schools',
  'Columbus City Schools',
  'Columbus City Schools',
  'Columbus City Schools',
  'Columbus City Schools',
  'Columbus City Schools',
  'Columbus City Schools',
  'Columbus City Schools',
  'Wabash',
  'City of North Las Vegas',
  'Premier Transportation',
  'Premier Transportation',
  'Enstructure',
  'Columbus City Schools',
  'Premier Transportation',
  'CCBCU Resource Center',
  'Columbus City Schools',
  'Columbus City Schools',
  'Wabash',
  'Columbus City Schools',
  'Columbus City Schools',
  'Columbus City Schools',
  'Columbus City Schools',
  'Columbus City Schools',
  'Columbus City Schools',
  'Premier Transportation',
  'Columbus City Schools',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Protective Life Corporation',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Premier Transportation',
  'CCBCU Resource Center',
  'Expro',
  'CCBCU Resource Center',
  'Columbus City Schools',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Progress Rail Service Center',
  'Protective Life Corporation',
  'Premier Transportation',
  'Audacy',
  'CCBCU Resource Center',
  'Expro',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Audacy',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Protective Life Corporation',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Wabash',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Wabash',
  'Wabash',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU',
  'Premier Transportation',
  'Premier Transportation',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Enstructure',
  'Columbus City Schools',
  'Enstructure',
  'CCBCU',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Wabash',
  'CCBCU Resource Center',
  'Protective Life Corporation',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Audacy',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Progress Rail Service Center',
  'Wabash',
  'CCBCU Resource Center',
  'Protective Life Corporation',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'City of North Las Vegas',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Wabash',
  'Protective Life Corporation',
  'CCBCU Resource Center',
  'Audacy',
  'Wabash',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Progress Rail Service Center',
  'Mayville Engineering Company',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Wabash',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Wabash',
  'Wabash',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Audacy',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Audacy',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Wabash',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Expro Americas LLC',
  'Wabash',
  'Expro Americas LLC',
  'Protective Life Corporation',
  'Protective Life Corporation',
  'Wabash',
  'Wabash',
  'Wabash',
  'Protective Life Corporation',
  'Areas',
  'Protective Life Corporation',
  'Wabash',
  'Wabash',
  'CCBCU Resource Center',
  'Wabash',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Protective Life Corporation',
  'Audacy',
  'Audacy',
  'Rogers Electrical',
  'Wabash',
  'Wabash',
  'Areas',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Expro Americas LLC',
  'CCBCU Resource Center',
  'Wabash',
  'CCBCU Resource Center',
  'Peachtree Orthopaedic Clinic',
  'Sisecam Chemical Resources',
  'Audacy',
  'CCBCU Resource Center',
  'Wabash',
  'Rogers Electrical',
  'Rogers Electrical',
  'Areas',
  'CCBCU Resource Center',
  'Kidde Global Solutions',
  'Wabash',
  'CCBCU Resource Center',
  'Kidde Global Solutions',
  'Protective Life Corporation',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Areas',
  'Wabash',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Audacy',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Rogers Electrical',
  'CCBCU Resource Center',
  'Sisecam Chemical Resources LLC',
  'Wabash',
  'Wabash',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Protective Life Corporation',
  'Premier Transportation',
  'Wabash',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'CCBCU Resource Center',
  'Protective Life Corporation',
  'Protective Life Corporation',
  'Protective Life Corporation',
  'CCBCU Resource Center',
  'Wabash',
  'Wabash',
  'Areas',
  'Areas',
  'Premier Transportation',
  'Audacy',
  'CCBCU Resource Center',
  'Protective Life Corporation',
  'CCBCU Resource Center',
  'Wabash'
];

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
    admin: null,
    luminaAdmin: null,
    decemberQa: null
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

    const adminInfo = ensureSuperAdminUser(roleIdsByName, campaignIdsByName);
    summary.admin = adminInfo;

    const luminaAdminInfo = ensureLuminaAdminUser(roleIdsByName, campaignIdsByName);
    summary.luminaAdmin = luminaAdminInfo;

    const decemberQaInfo = seedDecemberQualityData();
    summary.decemberQa = decemberQaInfo;

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

function seedDecemberQualityData(options) {
  const config = options || {};
  const year = typeof config.year === 'number' ? config.year : new Date().getFullYear();
  const targetMonth = 11;
  const sheetName = (typeof QA_RECORDS !== 'undefined' && QA_RECORDS) ? QA_RECORDS : 'Quality';
  const headers = (typeof QA_HEADERS !== 'undefined' && Array.isArray(QA_HEADERS) && QA_HEADERS.length)
    ? QA_HEADERS.slice()
    : [];

  if (!headers.length) {
    throw new Error('QA_HEADERS is not available for seeding.');
  }

  const sheet = ensureSheetWithHeaders(sheetName, headers);
  const existingCount = countDecemberQaEntries_(sheet, headers, year, targetMonth);
  if (existingCount > 0 && !config.force) {
    return {
      inserted: 0,
      skipped: existingCount,
      message: 'December QA data already exists. Pass force=true to seed again.'
    };
  }

  const rows = buildDecemberQaRows_(headers, year, targetMonth);
  if (!rows.length) {
    return { inserted: 0, skipped: existingCount, message: 'No QA rows generated.' };
  }

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);

  return { inserted: rows.length, skipped: existingCount, message: 'December QA seed data inserted.' };
}

function countDecemberQaEntries_(sheet, headers, year, targetMonth) {
  const dateIndex = headers.indexOf('CallDate');
  if (dateIndex === -1 || sheet.getLastRow() < 2) {
    return 0;
  }

  const data = sheet.getRange(2, dateIndex + 1, sheet.getLastRow() - 1, 1).getValues();
  let count = 0;
  for (let i = 0; i < data.length; i += 1) {
    const value = data[i][0];
    if (Object.prototype.toString.call(value) === '[object Date]') {
      if (value.getFullYear() === year && value.getMonth() === targetMonth) {
        count += 1;
      }
    }
  }
  return count;
}

function buildDecemberQaRows_(headers, year, targetMonth) {
  const rows = [];
  const questionCount = 19;
  const auditors = ['Quality Auditor', 'QA Team Lead', 'Quality Coach'];
  const feedbackOptions = ['Yes', 'No'];
  const overallFeedback = [
    'Great call handling with clear documentation.',
    'Solid compliance with a few improvement areas.',
    'Needs additional attention to policy reminders.',
    'Strong empathy and problem resolution.',
    'Follow-up steps were well communicated.'
  ];
  const agentFeedback = [
    'Keep up the strong client rapport.',
    'Remember to confirm next steps before ending calls.',
    'Great job keeping the call concise.',
    'Focus on summarizing the resolution.',
    'Nice consistency across checks.'
  ];

  for (let day = 1; day <= 31; day += 1) {
    const callDate = new Date(year, targetMonth, day);
    DECEMBER_QA_AGENTS.forEach((agentName, agentIndex) => {
      const randomSeed = (day * 1000) + agentIndex;
      const random = createSeededRandom_(randomSeed);
      const clientName = pickRandomFrom_(DECEMBER_QA_CLIENTS, random) || 'CCBCU Resource Center';
      const callerName = pickRandomFrom_(DECEMBER_QA_AGENTS, random);
      const auditorName = pickRandomFrom_(auditors, random);
      const feedbackShared = pickRandomFrom_(feedbackOptions, random);
      const questionResults = buildQuestionResults_(questionCount, random);
      const totalScore = questionResults.totalYes;
      const percentage = Math.round((totalScore / questionCount) * 100);
      const callTimestamp = new Date(year, targetMonth, day, 9 + Math.floor(random() * 7), Math.floor(random() * 60));
      const auditDate = new Date(year, targetMonth, day, 14 + Math.floor(random() * 4), Math.floor(random() * 60));

      const row = headers.map(header => buildQaCell_({
        header: header,
        agentName: agentName,
        callerName: callerName,
        clientName: clientName,
        auditorName: auditorName,
        feedbackShared: feedbackShared,
        questionResults: questionResults,
        totalScore: totalScore,
        percentage: percentage,
        callDate: callDate,
        callTimestamp: callTimestamp,
        auditDate: auditDate,
        random: random,
        overallFeedback: overallFeedback,
        agentFeedback: agentFeedback
      }));

      rows.push(row);
    });
  }

  return rows;
}

function buildQuestionResults_(questionCount, random) {
  const results = [];
  let totalYes = 0;
  for (let i = 1; i <= questionCount; i += 1) {
    const isYes = random() > 0.2;
    results.push({
      label: 'Q' + i,
      value: isYes ? 'Yes' : 'No',
      note: isYes ? '' : 'Needs improvement on Q' + i + '.'
    });
    if (isYes) {
      totalYes += 1;
    }
  }
  return { results: results, totalYes: totalYes };
}

function buildQaCell_(context) {
  const header = context.header;
  if (header === 'ID') return Utilities.getUuid();
  if (header === 'Timestamp') return context.callTimestamp;
  if (header === 'CallerName') return context.callerName;
  if (header === 'AgentName') return context.agentName;
  if (header === 'AgentEmail') return buildAgentEmail_(context.agentName);
  if (header === 'ClientName') return context.clientName;
  if (header === 'CallDate') return context.callDate;
  if (header === 'CaseNumber') return 'CASE-' + String(Math.floor(context.random() * 900000) + 100000);
  if (header === 'CallLink') return 'https://call.example.com/' + Utilities.getUuid();
  if (header === 'AuditorName') return context.auditorName;
  if (header === 'AuditDate') return context.auditDate;
  if (header === 'FeedbackShared') return context.feedbackShared;
  if (header.indexOf('Q') === 0 && header.indexOf('Note') === -1) {
    return getQuestionValue_(header, context.questionResults);
  }
  if (header.indexOf('Q') === 0 && header.indexOf('Note') > -1) {
    return getQuestionNote_(header, context.questionResults);
  }
  if (header === 'OverallFeedback') return pickRandomFrom_(context.overallFeedback, context.random);
  if (header === 'TotalScore') return context.totalScore;
  if (header === 'Percentage') return context.percentage;
  if (header === 'Notes') return 'Seeded December QA record.';
  if (header === 'AgentFeedback') return pickRandomFrom_(context.agentFeedback, context.random);
  if (header === 'CoachingProvided') return context.totalScore < 15 ? 'Yes' : 'No';
  return '';
}

function getQuestionValue_(header, questionResults) {
  const index = parseInt(header.replace('Q', ''), 10) - 1;
  const entry = questionResults.results[index];
  return entry ? entry.value : '';
}

function getQuestionNote_(header, questionResults) {
  const number = header.replace('Q', '').replace(' Note', '');
  const index = parseInt(number, 10) - 1;
  const entry = questionResults.results[index];
  return entry ? entry.note : '';
}

function pickRandomFrom_(list, random) {
  if (!list || !list.length) return '';
  const index = Math.floor(random() * list.length);
  return list[index];
}

function buildAgentEmail_(agentName) {
  const normalized = String(agentName || '')
    .toLowerCase()
    .replace(/[^a-z\s-]/g, '')
    .replace(/\s+/g, '.');
  return normalized ? normalized + '@vlbpo.com' : 'agent@vlbpo.com';
}

function createSeededRandom_(seedValue) {
  let seed = 0;
  if (typeof seedValue === 'number') {
    seed = seedValue;
  } else if (seedValue) {
    const text = String(seedValue);
    for (let i = 0; i < text.length; i += 1) {
      seed = (seed << 5) - seed + text.charCodeAt(i);
      seed |= 0;
    }
  } else {
    seed = Date.now();
  }

  return function seededRandom() {
    seed = (seed * 9301 + 49297) % 233280;
    return seed / 233280;
  };
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
 * Ensure there is a super admin user with a known password and permissions.
 */
function ensureSuperAdminUser(roleIdsByName, campaignIdsByName) {
  return ensureSeedAdministrator(SEED_ADMIN_PROFILE, roleIdsByName, campaignIdsByName);
}

/**
 * Ensure there is a Lumina admin user with a known password and permissions.
 */
function ensureLuminaAdminUser(roleIdsByName, campaignIdsByName) {
  return ensureSeedAdministrator(SEED_LUMINA_ADMIN_PROFILE, roleIdsByName, campaignIdsByName);
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

    syncUserRoleLinks(existing.ID, desiredRoleIds);
    assignAdminCampaignAccess(existing.ID, Object.values(campaignIdsByName));
    ensureCanLoginFlag(existing.ID, true);

    return {
      status: 'updated',
      userId: existing.ID,
      email: profile.email,
      message: (updateResult && updateResult.message) || (label + ' refreshed.')
    };
  }

  const createResult = clientRegisterUser(payload);

  if (!createResult || !createResult.success) {
    throw new Error('Failed to create ' + label + ': ' + (createResult && createResult.error ? createResult.error : 'Unknown error'));
  }

  let adminRecord = AuthenticationService.getUserByEmail(profile.email);
  if (!adminRecord) {
    throw new Error(label + ' record not found after creation.');
  }

  if (profile.password) {
    if (adminRecord.EmailConfirmation) {
      const setPasswordResult = setPasswordWithToken(adminRecord.EmailConfirmation, profile.password);
      if (!setPasswordResult || !setPasswordResult.success) {
        throw new Error('Failed to set ' + label + ' password: ' + (setPasswordResult && setPasswordResult.message ? setPasswordResult.message : 'Unknown error'));
      }
    } else {
      setUserPasswordDirect(adminRecord.ID, profile.password);
    }
  }

  adminRecord = AuthenticationService.getUserByEmail(profile.email);
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
  const pwdIdx = headers.indexOf('PasswordHash');
  const resetIdx = headers.indexOf('ResetRequired');
  const updatedIdx = headers.indexOf('UpdatedAt');

  const hash = (typeof PASSWORD_UTILS.createPasswordHash === 'function')
    ? PASSWORD_UTILS.createPasswordHash(password)
    : PASSWORD_UTILS.hashPassword(password);
  const now = new Date();

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][idIdx]) === String(userId)) {
      if (pwdIdx >= 0) sh.getRange(r + 1, pwdIdx + 1).setValue(hash);
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
