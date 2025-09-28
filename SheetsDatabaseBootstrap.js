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

    const campaignsTableName = (typeof CAMPAIGNS_SHEET === 'string' && CAMPAIGNS_SHEET) ? CAMPAIGNS_SHEET : 'Campaigns';
    const userCampaignsTableName = (typeof USER_CAMPAIGNS_SHEET === 'string' && USER_CAMPAIGNS_SHEET) ? USER_CAMPAIGNS_SHEET : 'UserCampaigns';
    const rolesTableName = (typeof ROLES_SHEET === 'string' && ROLES_SHEET) ? ROLES_SHEET : 'Roles';
    const userRolesTableName = (typeof USER_ROLES_SHEET === 'string' && USER_ROLES_SHEET) ? USER_ROLES_SHEET : 'UserRoles';
    const campaignPermissionsTableName = (typeof CAMPAIGN_USER_PERMISSIONS_SHEET === 'string' && CAMPAIGN_USER_PERMISSIONS_SHEET)
      ? CAMPAIGN_USER_PERMISSIONS_SHEET
      : 'CampaignUserPermissions';

    function resolveHeaders(preferred, fallback) {
      if (Array.isArray(preferred) && preferred.length) {
        return preferred.slice();
      }
      return fallback.slice();
    }

    function ensureHeader(headers, name) {
      if (headers.indexOf(name) === -1) {
        headers.push(name);
      }
    }

    const campaignHeaders = resolveHeaders(typeof CAMPAIGNS_HEADERS !== 'undefined' ? CAMPAIGNS_HEADERS : null, [
      'ID', 'Name', 'Description', 'ClientName', 'Status', 'Channel', 'Timezone', 'SlaTier', 'CreatedAt', 'UpdatedAt', 'DeletedAt'
    ]);
    ensureHeader(campaignHeaders, 'DeletedAt');

    function mapCampaignColumn(header) {
      switch (header) {
        case 'ID':
          return { name: header, type: 'string', primaryKey: true };
        case 'Name':
          return { name: header, type: 'string', required: true, maxLength: 160 };
        case 'Description':
          return { name: header, type: 'string', nullable: true, maxLength: 1024 };
        case 'ClientName':
          return { name: header, type: 'string', nullable: true, maxLength: 160 };
        case 'Status':
          return { name: header, type: 'enum', required: true, allowedValues: ['draft', 'active', 'paused', 'retired'], defaultValue: 'active' };
        case 'Channel':
          return { name: header, type: 'enum', nullable: true, allowedValues: ['voice', 'email', 'chat', 'back-office', 'omnichannel'] };
        case 'Timezone':
          return { name: header, type: 'string', nullable: true, maxLength: 64 };
        case 'SlaTier':
          return { name: header, type: 'string', nullable: true, maxLength: 64 };
        case 'CreatedAt':
        case 'UpdatedAt':
          return { name: header, type: 'timestamp', required: true };
        case 'DeletedAt':
          return { name: header, type: 'timestamp', nullable: true };
        default:
          return { name: header, type: 'string', nullable: true };
      }
    }

    SheetsDB.defineTable({
      name: campaignsTableName,
      version: 1,
      primaryKey: 'ID',
      idPrefix: 'CMP_',
      columns: campaignHeaders.map(mapCampaignColumn),
      indexes: [
        { name: campaignsTableName + '_Status_idx', field: 'Status' },
        { name: campaignsTableName + '_Client_idx', field: 'ClientName' }
      ]
    });

    const userCampaignHeaders = resolveHeaders(typeof USER_CAMPAIGNS_HEADERS !== 'undefined' ? USER_CAMPAIGNS_HEADERS : null, [
      'ID', 'UserId', 'CampaignId', 'Role', 'IsPrimary', 'CreatedAt', 'UpdatedAt', 'DeletedAt'
    ]);
    ensureHeader(userCampaignHeaders, 'DeletedAt');

    function mapUserCampaignColumn(header) {
      switch (header) {
        case 'ID':
          return { name: header, type: 'string', primaryKey: true };
        case 'UserId':
        case 'UserID':
          return { name: header, type: 'string', required: true, references: { table: usersTableName, column: 'ID', allowNull: false } };
        case 'CampaignId':
        case 'CampaignID':
          return { name: header, type: 'string', required: true, references: { table: campaignsTableName, column: 'ID', allowNull: false } };
        case 'Role':
          return { name: header, type: 'enum', nullable: true, allowedValues: ['agent', 'lead', 'qa', 'supervisor', 'trainer', 'support'], defaultValue: 'agent' };
        case 'IsPrimary':
          return { name: header, type: 'boolean', defaultValue: false };
        case 'CreatedAt':
        case 'UpdatedAt':
          return { name: header, type: 'timestamp', required: true };
        case 'DeletedAt':
          return { name: header, type: 'timestamp', nullable: true };
        default:
          return { name: header, type: 'string', nullable: true };
      }
    }

    SheetsDB.defineTable({
      name: userCampaignsTableName,
      version: 1,
      primaryKey: 'ID',
      idPrefix: 'UCAMP_',
      columns: userCampaignHeaders.map(mapUserCampaignColumn),
      indexes: [
        { name: userCampaignsTableName + '_User_idx', field: 'UserId' },
        { name: userCampaignsTableName + '_Campaign_idx', field: 'CampaignId' }
      ]
    });

    const roleHeaders = resolveHeaders(typeof ROLES_HEADER !== 'undefined' ? ROLES_HEADER : null, [
      'ID', 'Name', 'NormalizedName', 'Scope', 'Description', 'CreatedAt', 'UpdatedAt', 'DeletedAt'
    ]);
    ensureHeader(roleHeaders, 'DeletedAt');

    function mapRoleColumn(header) {
      switch (header) {
        case 'ID':
          return { name: header, type: 'string', primaryKey: true };
        case 'Name':
          return { name: header, type: 'string', required: true, maxLength: 80 };
        case 'NormalizedName':
          return { name: header, type: 'string', required: true, maxLength: 80 };
        case 'Scope':
          return { name: header, type: 'enum', required: true, allowedValues: ['global', 'campaign', 'team'], defaultValue: 'global' };
        case 'Description':
          return { name: header, type: 'string', nullable: true, maxLength: 512 };
        case 'CreatedAt':
        case 'UpdatedAt':
          return { name: header, type: 'timestamp', required: true };
        case 'DeletedAt':
          return { name: header, type: 'timestamp', nullable: true };
        default:
          return { name: header, type: 'string', nullable: true };
      }
    }

    SheetsDB.defineTable({
      name: rolesTableName,
      version: 1,
      primaryKey: 'ID',
      idPrefix: 'ROLE_',
      columns: roleHeaders.map(mapRoleColumn),
      indexes: [
        { name: rolesTableName + '_Name_idx', field: 'NormalizedName', unique: true },
        { name: rolesTableName + '_Scope_idx', field: 'Scope' }
      ]
    });

    const userRoleHeaders = resolveHeaders(typeof USER_ROLES_HEADER !== 'undefined' ? USER_ROLES_HEADER : null, [
      'ID', 'UserId', 'RoleId', 'Scope', 'AssignedBy', 'CreatedAt', 'UpdatedAt', 'DeletedAt'
    ]);
    ensureHeader(userRoleHeaders, 'DeletedAt');

    function mapUserRoleColumn(header) {
      switch (header) {
        case 'ID':
          return { name: header, type: 'string', primaryKey: true };
        case 'UserId':
        case 'UserID':
          return { name: header, type: 'string', required: true, references: { table: usersTableName, column: 'ID', allowNull: false } };
        case 'RoleId':
        case 'RoleID':
          return { name: header, type: 'string', required: true, references: { table: rolesTableName, column: 'ID', allowNull: false } };
        case 'Scope':
          return { name: header, type: 'enum', nullable: true, allowedValues: ['global', 'campaign', 'team'] };
        case 'AssignedBy':
          return { name: header, type: 'string', nullable: true, references: { table: usersTableName, column: 'ID', allowNull: true } };
        case 'CreatedAt':
        case 'UpdatedAt':
          return { name: header, type: 'timestamp', required: true };
        case 'DeletedAt':
          return { name: header, type: 'timestamp', nullable: true };
        default:
          return { name: header, type: 'string', nullable: true };
      }
    }

    SheetsDB.defineTable({
      name: userRolesTableName,
      version: 1,
      primaryKey: 'ID',
      idPrefix: 'UROLE_',
      columns: userRoleHeaders.map(mapUserRoleColumn),
      indexes: [
        { name: userRolesTableName + '_User_idx', field: 'UserId' },
        { name: userRolesTableName + '_Role_idx', field: 'RoleId' }
      ]
    });

    const campaignPermissionHeaders = resolveHeaders(
      typeof CAMPAIGN_USER_PERMISSIONS_HEADERS !== 'undefined' ? CAMPAIGN_USER_PERMISSIONS_HEADERS : null,
      ['ID', 'CampaignID', 'UserID', 'PermissionLevel', 'Role', 'CanManageUsers', 'CanManagePages', 'Notes', 'CreatedAt', 'UpdatedAt', 'DeletedAt']
    );
    ensureHeader(campaignPermissionHeaders, 'DeletedAt');

    function mapCampaignPermissionColumn(header) {
      switch (header) {
        case 'ID':
          return { name: header, type: 'string', primaryKey: true };
        case 'CampaignID':
        case 'CampaignId':
          return { name: header, type: 'string', required: true, references: { table: campaignsTableName, column: 'ID', allowNull: false } };
        case 'UserID':
        case 'UserId':
          return { name: header, type: 'string', required: true, references: { table: usersTableName, column: 'ID', allowNull: false } };
        case 'PermissionLevel':
          return { name: header, type: 'enum', required: true, allowedValues: ['viewer', 'editor', 'manager', 'owner'], defaultValue: 'viewer' };
        case 'Role':
          return { name: header, type: 'string', nullable: true, maxLength: 120 };
        case 'CanManageUsers':
        case 'CanManagePages':
          return { name: header, type: 'boolean', defaultValue: false };
        case 'Notes':
          return { name: header, type: 'string', nullable: true, maxLength: 1024 };
        case 'CreatedAt':
        case 'UpdatedAt':
          return { name: header, type: 'timestamp', required: true };
        case 'DeletedAt':
          return { name: header, type: 'timestamp', nullable: true };
        default:
          return { name: header, type: 'string', nullable: true };
      }
    }

    SheetsDB.defineTable({
      name: campaignPermissionsTableName,
      version: 1,
      primaryKey: 'ID',
      idPrefix: 'CPERM_',
      columns: campaignPermissionHeaders.map(mapCampaignPermissionColumn),
      indexes: [
        { name: campaignPermissionsTableName + '_Campaign_idx', field: 'CampaignID' },
        { name: campaignPermissionsTableName + '_User_idx', field: 'UserID' }
      ]
    });
  } catch (err) {
    console.error('Failed to initialize Sheets database schemas', err);
    throw err;
  }
}
