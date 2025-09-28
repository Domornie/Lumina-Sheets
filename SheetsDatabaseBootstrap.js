function initializeSheetsDatabase() {
  if (typeof SheetsDB === 'undefined') {
    console.warn('SheetsDB is not available; skipping database bootstrap');
    return;
  }

  const usersTableName = (typeof USERS_SHEET === 'string' && USERS_SHEET) ? USERS_SHEET : 'Users';
  const sessionsTableName = (typeof SESSIONS_SHEET === 'string' && SESSIONS_SHEET) ? SESSIONS_SHEET : 'Sessions';

  function buildUserHeaders() {
    if (typeof USER_HEADERS !== 'undefined' && Array.isArray(USER_HEADERS) && USER_HEADERS.length) {
      return USER_HEADERS.slice();
    }
    if (typeof USERS_HEADERS !== 'undefined' && Array.isArray(USERS_HEADERS) && USERS_HEADERS.length) {
      return USERS_HEADERS.slice();
    }
    return ['ID', 'UserName', 'FullName', 'Email', 'CampaignID', 'PasswordHash', 'ResetRequired',
      'EmailConfirmation', 'EmailConfirmed', 'PhoneNumber', 'EmploymentStatus', 'HireDate', 'Country',
      'LockoutEnd', 'TwoFactorEnabled', 'CanLogin', 'Roles', 'Pages', 'CreatedAt', 'UpdatedAt', 'IsAdmin'];
  }

  function buildSessionHeaders() {
    if (typeof SESSION_HEADERS !== 'undefined' && Array.isArray(SESSION_HEADERS) && SESSION_HEADERS.length) {
      return SESSION_HEADERS.slice();
    }
    if (typeof SESSIONS_HEADERS !== 'undefined' && Array.isArray(SESSIONS_HEADERS) && SESSIONS_HEADERS.length) {
      return SESSIONS_HEADERS.slice();
    }
    return ['Token', 'UserId', 'CreatedAt', 'ExpiresAt', 'RememberMe', 'CampaignScope', 'UserAgent', 'IpAddress'];
  }

  function mapUserColumn(header) {
    switch (header) {
      case 'ID':
        return { name: header, type: 'string', primaryKey: true };
      case 'Email':
        return { name: header, type: 'string', required: true, unique: true, maxLength: 320 };
      case 'CanLogin':
        return { name: header, type: 'boolean', defaultValue: true };
      case 'ResetRequired':
      case 'EmailConfirmed':
      case 'TwoFactorEnabled':
      case 'IsAdmin':
        return { name: header, type: 'boolean', defaultValue: false };
      case 'HireDate':
        return { name: header, type: 'date', nullable: true };
      case 'LockoutEnd':
      case 'LastLogin':
        return { name: header, type: 'timestamp', nullable: true };
      case 'CreatedAt':
      case 'UpdatedAt':
        return { name: header, type: 'timestamp', required: true };
      case 'DeletedAt':
        return { name: header, type: 'timestamp', nullable: true };
      default:
        return { name: header, type: 'string', nullable: true };
    }
  }

  function mapSessionColumn(header) {
    switch (header) {
      case 'Token':
        return { name: header, type: 'string', primaryKey: true };
      case 'UserId':
        return { name: header, type: 'string', required: true };
      case 'CreatedAt':
      case 'UpdatedAt':
      case 'ExpiresAt':
        return { name: header, type: 'timestamp', required: true };
      case 'RememberMe':
        return { name: header, type: 'boolean', defaultValue: false };
      case 'CampaignScope':
      case 'ClientContext':
        return { name: header, type: 'json', nullable: true };
      case 'UserAgent':
        return { name: header, type: 'string', nullable: true, maxLength: 512 };
      case 'IpAddress':
        return { name: header, type: 'string', nullable: true, maxLength: 64 };
      case 'DeletedAt':
        return { name: header, type: 'timestamp', nullable: true };
      default:
        return { name: header, type: 'string', nullable: true };
    }
  }

  function buildFallbackAuthSchemas() {
    const userHeaders = buildUserHeaders();
    if (userHeaders.indexOf('LastLogin') === -1) {
      userHeaders.push('LastLogin');
    }
    if (userHeaders.indexOf('DeletedAt') === -1) {
      userHeaders.push('DeletedAt');
    }

    const sessionHeaders = buildSessionHeaders();
    if (sessionHeaders.indexOf('UpdatedAt') === -1) {
      const createdIdx = sessionHeaders.indexOf('CreatedAt');
      const insertIdx = createdIdx !== -1 ? createdIdx + 1 : sessionHeaders.length;
      sessionHeaders.splice(insertIdx, 0, 'UpdatedAt');
    }
    if (sessionHeaders.indexOf('ClientContext') === -1) {
      const userAgentIdx = sessionHeaders.indexOf('UserAgent');
      const insertIdx = userAgentIdx !== -1 ? userAgentIdx + 1 : sessionHeaders.length;
      sessionHeaders.splice(insertIdx, 0, 'ClientContext');
    }
    if (sessionHeaders.indexOf('DeletedAt') === -1) {
      sessionHeaders.push('DeletedAt');
    }

    return [
      {
        name: usersTableName,
        version: 2,
        primaryKey: 'ID',
        idPrefix: 'USR_',
        columns: userHeaders.map(mapUserColumn),
        indexes: [
          { name: usersTableName + '_Email_idx', field: 'Email', unique: true },
          { name: usersTableName + '_Campaign_idx', field: 'CampaignID' }
        ]
      },
      {
        name: sessionsTableName,
        version: 2,
        primaryKey: 'Token',
        idPrefix: 'SES_',
        columns: sessionHeaders.map(mapSessionColumn),
        indexes: [
          { name: sessionsTableName + '_User_idx', field: 'UserId' },
          { name: sessionsTableName + '_Expiry_idx', field: 'ExpiresAt' }
        ],
        retentionDays: 45
      }
    ];
  }

  try {
    let authSchemas = [];
    if (typeof AuthenticationService !== 'undefined' && AuthenticationService && typeof AuthenticationService.getTableSchemas === 'function') {
      try {
        authSchemas = AuthenticationService.getTableSchemas() || [];
      } catch (authErr) {
        console.warn('initializeSheetsDatabase: Unable to load auth schemas from AuthenticationService:', authErr);
      }
    }

    if (!authSchemas || !authSchemas.length) {
      authSchemas = buildFallbackAuthSchemas();
    }

    authSchemas.forEach(function (schema) {
      SheetsDB.defineTable(schema);
    });

    SheetsDB.defineTable({
      name: 'AgentProfiles',
      version: 1,
      primaryKey: 'id',
      idPrefix: 'AGENT_',
      columns: [
        { name: 'id', type: 'string', primaryKey: true },
        { name: 'tenantId', type: 'string', required: true, index: true },
        { name: 'userId', type: 'string', required: true, references: { table: usersTableName, column: 'ID', allowNull: false } },
        { name: 'teamId', type: 'string', nullable: true },
        { name: 'supervisorId', type: 'string', nullable: true, references: { table: usersTableName, column: 'ID', allowNull: true } },
        { name: 'employmentType', type: 'enum', required: true, allowedValues: ['full-time', 'part-time', 'contract'], defaultValue: 'full-time' },
        { name: 'primarySkillGroup', type: 'string', nullable: true },
        { name: 'status', type: 'enum', required: true, allowedValues: ['active', 'onboarding', 'inactive', 'terminated'], defaultValue: 'active' },
        { name: 'hireDate', type: 'date', nullable: true },
        { name: 'lastActiveAt', type: 'timestamp', nullable: true },
        { name: 'notes', type: 'string', nullable: true, maxLength: 4000 }
      ],
      indexes: [
        { name: 'AgentProfiles_user', field: 'userId', unique: true },
        { name: 'AgentProfiles_team', field: 'teamId' },
        { name: 'AgentProfiles_status', field: 'status' }
      ]
    });

    SheetsDB.defineTable({
      name: 'AgentSkillAssignments',
      version: 1,
      primaryKey: 'id',
      idPrefix: 'SKILL_',
      columns: [
        { name: 'id', type: 'string', primaryKey: true },
        { name: 'tenantId', type: 'string', required: true, index: true },
        { name: 'agentId', type: 'string', required: true, references: { table: 'AgentProfiles', column: 'id', allowNull: false } },
        { name: 'skillName', type: 'string', required: true, maxLength: 128 },
        { name: 'proficiency', type: 'enum', required: true, allowedValues: ['novice', 'intermediate', 'advanced', 'expert'], defaultValue: 'intermediate' },
        { name: 'certifiedAt', type: 'timestamp', nullable: true },
        { name: 'expiresAt', type: 'timestamp', nullable: true },
        { name: 'notes', type: 'string', nullable: true, maxLength: 1024 }
      ],
      indexes: [
        { name: 'AgentSkillAssignments_agent', field: 'agentId' },
        { name: 'AgentSkillAssignments_skill', field: 'skillName' }
      ]
    });

    SheetsDB.defineTable({
      name: 'AgentStatusEvents',
      version: 1,
      primaryKey: 'id',
      idPrefix: 'STAT_',
      columns: [
        { name: 'id', type: 'string', primaryKey: true },
        { name: 'tenantId', type: 'string', required: true, index: true },
        { name: 'agentId', type: 'string', required: true, references: { table: 'AgentProfiles', column: 'id', allowNull: false } },
        { name: 'status', type: 'enum', required: true, allowedValues: ['available', 'after-call', 'break', 'offline', 'training', 'meeting'] },
        { name: 'reason', type: 'string', nullable: true, maxLength: 512 },
        { name: 'occurredAt', type: 'timestamp', required: true },
        { name: 'expectedEndAt', type: 'timestamp', nullable: true },
        { name: 'durationSeconds', type: 'number', nullable: true, min: 0 },
        { name: 'metadata', type: 'json', nullable: true }
      ],
      indexes: [
        { name: 'AgentStatusEvents_agent', field: 'agentId' },
        { name: 'AgentStatusEvents_status', field: 'status' }
      ],
      retentionDays: 120
    });

    SheetsDB.defineTable({
      name: 'AgentPerformanceSummaries',
      version: 1,
      primaryKey: 'id',
      idPrefix: 'PERF_',
      columns: [
        { name: 'id', type: 'string', primaryKey: true },
        { name: 'tenantId', type: 'string', required: true, index: true },
        { name: 'agentId', type: 'string', required: true, references: { table: 'AgentProfiles', column: 'id', allowNull: false } },
        { name: 'periodStart', type: 'timestamp', required: true },
        { name: 'periodEnd', type: 'timestamp', required: true },
        { name: 'contactsHandled', type: 'number', required: true, min: 0 },
        { name: 'talkTimeSeconds', type: 'number', required: true, min: 0 },
        { name: 'afterCallWorkSeconds', type: 'number', required: true, min: 0 },
        { name: 'handleTimeSeconds', type: 'number', required: true, min: 0 },
        { name: 'serviceLevel', type: 'number', nullable: true, min: 0, max: 100 },
        { name: 'firstContactResolution', type: 'number', nullable: true, min: 0, max: 100 },
        { name: 'qualityScore', type: 'number', nullable: true, min: 0, max: 100 },
        { name: 'coachingNotes', type: 'string', nullable: true, maxLength: 4000 }
      ],
      indexes: [
        { name: 'AgentPerformanceSummaries_agent', field: 'agentId' },
        { name: 'AgentPerformanceSummaries_period', field: 'periodStart' }
      ],
      retentionDays: 730
    });
  } catch (err) {
    console.error('Failed to initialize Sheets database schemas', err);
    throw err;
  }
}
