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
      name: 'Customers',
      version: 1,
      primaryKey: 'id',
      idPrefix: 'CUS_',
      columns: [
        { name: 'id', type: 'string', primaryKey: true },
        { name: 'tenantId', type: 'string', required: true, index: true },
        { name: 'email', type: 'string', required: true, unique: true, index: true, pattern: '^[^@\s]+@[^@\s]+\\.[^@\s]+$' },
        { name: 'firstName', type: 'string', required: true, minLength: 1, maxLength: 100 },
        { name: 'lastName', type: 'string', required: true, minLength: 1, maxLength: 100 },
        { name: 'phone', type: 'string', nullable: true, maxLength: 32 },
        { name: 'status', type: 'enum', required: true, allowedValues: ['prospect', 'active', 'inactive'], defaultValue: 'prospect' },
        { name: 'notes', type: 'json', nullable: true }
      ],
      indexes: [
        { name: 'Customers_email', field: 'email' },
        { name: 'Customers_tenant', field: 'tenantId' }
      ],
      retentionDays: 365
    });

    SheetsDB.defineTable({
      name: 'Orders',
      version: 1,
      primaryKey: 'id',
      idPrefix: 'ORD_',
      columns: [
        { name: 'id', type: 'string', primaryKey: true },
        { name: 'tenantId', type: 'string', required: true, index: true },
        { name: 'customerId', type: 'string', required: true, references: { table: 'Customers', column: 'id', allowNull: false } },
        { name: 'orderTotal', type: 'number', required: true, min: 0 },
        { name: 'currency', type: 'enum', required: true, allowedValues: ['USD', 'EUR', 'GBP', 'PHP'], defaultValue: 'USD' },
        { name: 'status', type: 'enum', required: true, allowedValues: ['open', 'processing', 'closed', 'cancelled'], defaultValue: 'open' },
        { name: 'orderDate', type: 'timestamp', required: true },
        { name: 'fulfilledAt', type: 'timestamp', nullable: true }
      ],
      indexes: [
        { name: 'Orders_customer', field: 'customerId' },
        { name: 'Orders_status', field: 'status' }
      ],
      retentionDays: 730
    });

    SheetsDB.defineTable({
      name: 'WebhooksOutbox',
      version: 1,
      primaryKey: 'id',
      idPrefix: 'OUT_',
      columns: [
        { name: 'id', type: 'string', primaryKey: true },
        { name: 'eventType', type: 'string', required: true },
        { name: 'payload', type: 'json', required: true },
        { name: 'targetUrl', type: 'string', required: true },
        { name: 'deliveryStatus', type: 'enum', required: true, allowedValues: ['pending', 'sent', 'failed'], defaultValue: 'pending' },
        { name: 'lastError', type: 'string', nullable: true },
        { name: 'retryCount', type: 'number', required: true, defaultValue: 0, min: 0 },
        { name: 'nextAttemptAt', type: 'timestamp', nullable: true }
      ],
      indexes: [
        { name: 'Outbox_status', field: 'deliveryStatus' }
      ],
      retentionDays: 90
    });
  } catch (err) {
    console.error('Failed to initialize Sheets database schemas', err);
    throw err;
  }
}
