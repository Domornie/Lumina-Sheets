/**
 * Lumina Sheets v2 - Rebuilt core runtime
 * ------------------------------------------------------------
 * This implementation replaces the ad-hoc collection of global
 * functions with a modular, testable, and cache aware runtime
 * that treats Google Sheets as the primary database. The design
 * follows a clean architecture style so each concern (config,
 * sheet access, domain stores, application services, and HTTP
 * rendering) is isolated and easy to reason about.
 *
 * Key goals
 * ---------
 * - Deterministic bootstrap that guarantees every sheet exists
 *   with the expected headers before serving a request.
 * - Token based authentication and lightweight session store
 *   backed by the Sessions sheet with cache acceleration.
 * - Query optimisations that avoid repeatedly scanning sheets
 *   by memoising read models through CacheService/Properties.
 * - Simple REST style APIs for the HTML front-end to consume
 *   (`google.script.run`) without fragile global state.
 * - Minimal but extensible routing layer that can render the
 *   rebuilt V2 client while leaving legacy pages untouched.
 */

var LuminaAppV2 = (function () {
  'use strict';

  // ---------------------------------------------------------------------------
  // Configuration & Constants
  // ---------------------------------------------------------------------------

  /** Sheet names reused from the original MainUtilities module */
  const SHEETS = {
    USERS: 'Users',
    ROLES: 'Roles',
    USER_ROLES: 'UserRoles',
    USER_CLAIMS: 'UserClaims',
    SESSIONS: 'Sessions',
    CAMPAIGNS: 'Campaigns',
    PAGES: 'Pages',
    CAMPAIGN_PAGES: 'CampaignPages',
    PAGE_CATEGORIES: 'PageCategories',
    CAMPAIGN_USER_PERMISSIONS: 'CampaignUserPermissions',
    USER_MANAGERS: 'UserManagers',
    USER_CAMPAIGNS: 'UserCampaigns',
    DEBUG_LOGS: 'DebugLogs',
    ERROR_LOGS: 'ErrorLogs',
    NOTIFICATIONS: 'Notifications',
    CHAT_GROUPS: 'ChatGroups',
    CHAT_CHANNELS: 'ChatChannels',
    CHAT_MESSAGES: 'ChatMessages',
    CHAT_GROUP_MEMBERS: 'ChatGroupMembers',
    CHAT_MESSAGE_REACTIONS: 'ChatMessageReactions',
    CHAT_USER_PREFERENCES: 'ChatUserPreferences',
    CHAT_ANALYTICS: 'ChatAnalytics',
    CHAT_CHANNEL_MEMBERS: 'ChatChannelMembers'
  };

  /** Headers copied verbatim from MainUtilities so the Sheets schema matches */
  const HEADERS = {
    USERS: [
      'ID', 'UserName', 'FullName', 'Email', 'CampaignID', 'PasswordHash', 'ResetRequired',
      'EmailConfirmation', 'EmailConfirmed', 'PhoneNumber', 'EmploymentStatus', 'HireDate', 'Country',
      'LockoutEnd', 'TwoFactorEnabled', 'CanLogin', 'Roles', 'Pages', 'CreatedAt', 'UpdatedAt', 'IsAdmin'
    ],
    ROLES: ['ID', 'Name', 'NormalizedName', 'Scope', 'Description', 'CreatedAt', 'UpdatedAt', 'DeletedAt'],
    USER_ROLES: ['ID', 'UserId', 'RoleId', 'Scope', 'AssignedBy', 'CreatedAt', 'UpdatedAt', 'DeletedAt'],
    USER_CLAIMS: ['ID', 'UserId', 'ClaimType', 'CreatedAt', 'UpdatedAt'],
    SESSIONS: ['Token', 'UserId', 'CreatedAt', 'ExpiresAt', 'RememberMe', 'CampaignScope', 'UserAgent', 'IpAddress'],
    CAMPAIGNS: ['ID', 'Name', 'Description', 'ClientName', 'Status', 'Channel', 'Timezone', 'SlaTier', 'CreatedAt', 'UpdatedAt', 'DeletedAt'],
    PAGES: ['PageKey', 'PageTitle', 'PageIcon', 'Description', 'IsSystemPage', 'RequiresAdmin', 'CreatedAt', 'UpdatedAt'],
    CAMPAIGN_PAGES: ['ID', 'CampaignID', 'PageKey', 'PageTitle', 'PageIcon', 'CategoryID', 'SortOrder', 'IsActive', 'CreatedAt', 'UpdatedAt'],
    PAGE_CATEGORIES: ['ID', 'CampaignID', 'CategoryName', 'CategoryIcon', 'SortOrder', 'IsActive', 'CreatedAt', 'UpdatedAt'],
    CAMPAIGN_USER_PERMISSIONS: ['ID', 'CampaignID', 'UserID', 'PermissionLevel', 'Role', 'CanManageUsers', 'CanManagePages', 'Notes', 'CreatedAt', 'UpdatedAt', 'DeletedAt'],
    USER_MANAGERS: ['ID', 'ManagerUserID', 'ManagedUserID', 'CampaignID', 'CreatedAt', 'UpdatedAt'],
    USER_CAMPAIGNS: ['ID', 'UserId', 'CampaignId', 'Role', 'IsPrimary', 'CreatedAt', 'UpdatedAt', 'DeletedAt'],
    DEBUG_LOGS: ['Timestamp', 'Message'],
    ERROR_LOGS: ['Timestamp', 'Error', 'Stack'],
    NOTIFICATIONS: ['ID', 'UserId', 'Type', 'Severity', 'Title', 'Message', 'Data', 'Read', 'ActionTaken', 'CreatedAt', 'ReadAt', 'ExpiresAt'],
    CHAT_GROUPS: ['ID', 'Name', 'Description', 'CreatedBy', 'CreatedAt', 'UpdatedAt'],
    CHAT_CHANNELS: ['ID', 'GroupId', 'Name', 'Description', 'IsPrivate', 'CreatedBy', 'CreatedAt', 'UpdatedAt'],
    CHAT_MESSAGES: ['ID', 'ChannelId', 'UserId', 'Message', 'Timestamp', 'EditedAt', 'ParentMessageId', 'IsDeleted'],
    CHAT_GROUP_MEMBERS: ['ID', 'GroupId', 'UserId', 'JoinedAt', 'Role', 'IsActive'],
    CHAT_MESSAGE_REACTIONS: ['ID', 'MessageId', 'UserId', 'Reaction', 'Timestamp'],
    CHAT_USER_PREFERENCES: ['UserId', 'NotificationSettings', 'Theme', 'LastSeen', 'Status'],
    CHAT_ANALYTICS: ['Timestamp', 'UserId', 'Action', 'Details', 'SessionId'],
    CHAT_CHANNEL_MEMBERS: ['ID', 'ChannelId', 'UserId', 'JoinedAt', 'Role', 'IsActive']
  };

  const CACHE_TTL = {
    SHORT: 60,
    MEDIUM: 300,
    LONG: 600
  };

  const PASSWORD_DIGEST = {
    ALGORITHM: Utilities.DigestAlgorithm.SHA_256,
    ENCODING: Utilities.Charset.UTF_8
  };

  // ---------------------------------------------------------------------------
  // Low level helpers
  // ---------------------------------------------------------------------------

  const scriptCache = CacheService.getScriptCache();
  const properties = PropertiesService.getScriptProperties();

  function cacheKey(parts) {
    return ['lumina', 'v2'].concat(parts).join('::');
  }

  function withCache(keyParts, ttl, producer) {
    const key = cacheKey(keyParts);
    try {
      const cached = scriptCache.get(key);
      if (cached) {
        return JSON.parse(cached);
      }
    } catch (err) {
      logError('CacheRead', err);
    }

    const value = producer();
    try {
      scriptCache.put(key, JSON.stringify(value), ttl);
    } catch (err) {
      logError('CacheWrite', err);
    }
    return value;
  }

  function nowIso() {
    return new Date().toISOString();
  }

  function toBool(value) {
    if (typeof value === 'boolean') return value;
    if (typeof value === 'number') return value !== 0;
    if (!value) return false;
    const normalized = String(value).trim().toLowerCase();
    return normalized === 'true' || normalized === '1' || normalized === 'yes';
  }

  function safeJsonParse(value, fallback) {
    try {
      return JSON.parse(value);
    } catch (err) {
      return fallback;
    }
  }

  function generateId(prefix) {
    const raw = Utilities.getUuid();
    return prefix ? prefix + '_' + raw.replace(/-/g, '') : raw.replace(/-/g, '');
  }

  function hashPassword(raw) {
    if (!raw) return '';
    const digest = Utilities.computeDigest(PASSWORD_DIGEST.ALGORITHM, raw, PASSWORD_DIGEST.ENCODING);
    return Utilities.base64Encode(digest);
  }

  function verifyPassword(candidate, storedHash) {
    if (!storedHash) return false;
    return hashPassword(candidate) === storedHash;
  }

  function getSpreadsheet() {
    return SpreadsheetApp.getActive();
  }

  function ensureSheetWithHeaders(name, headers) {
    const spreadsheet = getSpreadsheet();
    let sheet = spreadsheet.getSheetByName(name);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(name);
    }
    const firstRow = sheet.getRange(1, 1, 1, headers.length);
    const existing = firstRow.getValues()[0].map(String);
    if (existing.join('||') !== headers.join('||')) {
      firstRow.setValues([headers]);
      firstRow.setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    return sheet;
  }

  function getSheet(name, headers) {
    return ensureSheetWithHeaders(name, headers || []);
  }

  function readTable(sheetName, headers) {
    return withCache(['table', sheetName], CACHE_TTL.MEDIUM, function () {
      const sheet = getSheet(sheetName, headers);
      const range = sheet.getDataRange();
      const values = range.getValues();
      if (values.length <= 1) return [];
      const headerRow = values[0];
      const rows = [];
      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        if (row.join('').trim() === '') continue;
        const obj = {};
        for (let col = 0; col < headerRow.length; col++) {
          const key = headerRow[col];
          obj[key] = row[col];
        }
        obj.__rowIndex = i + 1;
        rows.push(obj);
      }
      return rows;
    });
  }

  function clearTableCache(sheetName) {
    try {
      scriptCache.remove(cacheKey(['table', sheetName]));
    } catch (err) {
      logError('CacheRemove', err);
    }
  }

  function writeRow(sheetName, headers, rowIndex, values) {
    const sheet = getSheet(sheetName, headers);
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([values]);
    clearTableCache(sheetName);
  }

  function appendRow(sheetName, headers, values) {
    const sheet = getSheet(sheetName, headers);
    sheet.appendRow(values);
    clearTableCache(sheetName);
  }

  function findRowByColumn(sheetName, headers, columnName, value) {
    const sheet = getSheet(sheetName, headers);
    const columnIndex = headers.indexOf(columnName) + 1;
    if (columnIndex <= 0) {
      throw new Error('Column not found: ' + columnName + ' in sheet ' + sheetName);
    }
    const finder = sheet.createTextFinder(String(value));
    finder.matchCase(true);
    const match = finder.findNext();
    if (!match) return null;
    if (match.getColumn() !== columnIndex) return null;
    const rowValues = sheet.getRange(match.getRow(), 1, 1, headers.length).getValues()[0];
    const rowObj = {};
    headers.forEach(function (header, idx) {
      rowObj[header] = rowValues[idx];
    });
    rowObj.__rowIndex = match.getRow();
    return rowObj;
  }

  function logDebug(context, message) {
    console.log('[Lumina][Debug][' + context + '] ' + message);
    try {
      const sheet = getSheet(SHEETS.DEBUG_LOGS, HEADERS.DEBUG_LOGS);
      sheet.appendRow([nowIso(), '[' + context + '] ' + message]);
    } catch (err) {
      console.warn('Unable to persist debug log', err);
    }
  }

  function logError(context, error) {
    const err = error instanceof Error ? error : new Error(String(error));
    console.error('[Lumina][Error][' + context + ']', err);
    try {
      const sheet = getSheet(SHEETS.ERROR_LOGS, HEADERS.ERROR_LOGS);
      sheet.appendRow([nowIso(), err.message, err.stack || '']);
    } catch (sheetErr) {
      console.warn('Unable to persist error log', sheetErr);
    }
  }

  // ---------------------------------------------------------------------------
  // Domain stores (Users, Sessions, Campaigns, Pages, Notifications)
  // ---------------------------------------------------------------------------

  function ensureBootstrap() {
    const bootKey = cacheKey(['bootstrap', 'completed']);
    const state = properties.getProperty(bootKey);
    if (state === '1') {
      return;
    }
    const specs = [
      { sheet: SHEETS.USERS, headers: HEADERS.USERS },
      { sheet: SHEETS.ROLES, headers: HEADERS.ROLES },
      { sheet: SHEETS.USER_ROLES, headers: HEADERS.USER_ROLES },
      { sheet: SHEETS.USER_CLAIMS, headers: HEADERS.USER_CLAIMS },
      { sheet: SHEETS.SESSIONS, headers: HEADERS.SESSIONS },
      { sheet: SHEETS.CAMPAIGNS, headers: HEADERS.CAMPAIGNS },
      { sheet: SHEETS.PAGES, headers: HEADERS.PAGES },
      { sheet: SHEETS.CAMPAIGN_PAGES, headers: HEADERS.CAMPAIGN_PAGES },
      { sheet: SHEETS.PAGE_CATEGORIES, headers: HEADERS.PAGE_CATEGORIES },
      { sheet: SHEETS.CAMPAIGN_USER_PERMISSIONS, headers: HEADERS.CAMPAIGN_USER_PERMISSIONS },
      { sheet: SHEETS.USER_MANAGERS, headers: HEADERS.USER_MANAGERS },
      { sheet: SHEETS.USER_CAMPAIGNS, headers: HEADERS.USER_CAMPAIGNS },
      { sheet: SHEETS.DEBUG_LOGS, headers: HEADERS.DEBUG_LOGS },
      { sheet: SHEETS.ERROR_LOGS, headers: HEADERS.ERROR_LOGS },
      { sheet: SHEETS.NOTIFICATIONS, headers: HEADERS.NOTIFICATIONS },
      { sheet: SHEETS.CHAT_GROUPS, headers: HEADERS.CHAT_GROUPS },
      { sheet: SHEETS.CHAT_CHANNELS, headers: HEADERS.CHAT_CHANNELS },
      { sheet: SHEETS.CHAT_MESSAGES, headers: HEADERS.CHAT_MESSAGES },
      { sheet: SHEETS.CHAT_GROUP_MEMBERS, headers: HEADERS.CHAT_GROUP_MEMBERS },
      { sheet: SHEETS.CHAT_MESSAGE_REACTIONS, headers: HEADERS.CHAT_MESSAGE_REACTIONS },
      { sheet: SHEETS.CHAT_USER_PREFERENCES, headers: HEADERS.CHAT_USER_PREFERENCES },
      { sheet: SHEETS.CHAT_ANALYTICS, headers: HEADERS.CHAT_ANALYTICS },
      { sheet: SHEETS.CHAT_CHANNEL_MEMBERS, headers: HEADERS.CHAT_CHANNEL_MEMBERS }
    ];
    specs.forEach(function (spec) {
      ensureSheetWithHeaders(spec.sheet, spec.headers);
    });
    properties.setProperty(bootKey, '1');
    logDebug('Bootstrap', 'Core sheets ensured');
  }

  function sanitizeUser(user) {
    if (!user) return null;
    const clone = {};
    Object.keys(user).forEach(function (key) {
      if (key === 'PasswordHash') return;
      if (key.indexOf('__') === 0) return;
      clone[key] = user[key];
    });
    clone.CanLogin = toBool(user.CanLogin);
    clone.EmailConfirmed = toBool(user.EmailConfirmed);
    clone.ResetRequired = toBool(user.ResetRequired);
    clone.IsAdmin = toBool(user.IsAdmin);
    return clone;
  }

  function findUserByIdentity(identity) {
    if (!identity) return null;
    const normalized = String(identity).trim().toLowerCase();
    const users = readTable(SHEETS.USERS, HEADERS.USERS);
    for (let i = 0; i < users.length; i++) {
      const user = users[i];
      const username = String(user.UserName || '').trim().toLowerCase();
      const email = String(user.Email || '').trim().toLowerCase();
      if (username === normalized || email === normalized) {
        return user;
      }
    }
    return null;
  }

  function findUserById(id) {
    if (!id) return null;
    const users = readTable(SHEETS.USERS, HEADERS.USERS);
    for (let i = 0; i < users.length; i++) {
      if (String(users[i].ID) === String(id)) {
        return users[i];
      }
    }
    return null;
  }

  function getUserCampaigns(userId) {
    const memberships = readTable(SHEETS.USER_CAMPAIGNS, HEADERS.USER_CAMPAIGNS)
      .filter(function (row) {
        return String(row.UserId) === String(userId) && !row.DeletedAt;
      });
    if (!memberships.length) return [];
    const campaigns = readTable(SHEETS.CAMPAIGNS, HEADERS.CAMPAIGNS);
    const byId = {};
    campaigns.forEach(function (campaign) {
      byId[String(campaign.ID)] = campaign;
    });
    return memberships.map(function (membership) {
      const campaign = byId[String(membership.CampaignId)] || {};
      return {
        id: membership.CampaignId,
        name: campaign.Name || 'Campaign ' + membership.CampaignId,
        role: membership.Role || '',
        isPrimary: toBool(membership.IsPrimary)
      };
    });
  }

  function getUserNotifications(userId) {
    return readTable(SHEETS.NOTIFICATIONS, HEADERS.NOTIFICATIONS)
      .filter(function (row) {
        return String(row.UserId) === String(userId) && !toBool(row.Read);
      })
      .map(function (row) {
        return {
          id: row.ID,
          severity: row.Severity || 'info',
          title: row.Title || '',
          message: row.Message || '',
          createdAt: row.CreatedAt || null
        };
      });
  }

  function getNavigationForUser(user, campaignId) {
    const pages = readTable(SHEETS.PAGES, HEADERS.PAGES);
    const categories = readTable(SHEETS.PAGE_CATEGORIES, HEADERS.PAGE_CATEGORIES)
      .filter(function (cat) {
        return !cat.CampaignID || String(cat.CampaignID) === String(campaignId);
      });

    const byCategory = {};
    categories.forEach(function (cat) {
      const key = String(cat.ID || cat.CategoryName || 'default');
      byCategory[key] = {
        id: key,
        name: cat.CategoryName || 'General',
        icon: cat.CategoryIcon || 'fas fa-stream',
        pages: []
      };
    });

    pages.forEach(function (page) {
      const requiresAdmin = toBool(page.RequiresAdmin);
      if (requiresAdmin && !toBool(user.IsAdmin)) return;
      const categoryKey = String(page.CategoryID || 'default');
      if (!byCategory[categoryKey]) {
        byCategory[categoryKey] = {
          id: categoryKey,
          name: 'General',
          icon: 'fas fa-stream',
          pages: []
        };
      }
      byCategory[categoryKey].pages.push({
        key: page.PageKey,
        title: page.PageTitle,
        icon: page.PageIcon || 'far fa-circle'
      });
    });

    if (!byCategory['default']) {
      byCategory['default'] = {
        id: 'default',
        name: 'Workspace',
        icon: 'fas fa-stream',
        pages: []
      };
    }
    const defaultPages = byCategory['default'].pages;
    if (!defaultPages.some(function (p) { return String(p.key).toLowerCase() === 'dashboard'; })) {
      defaultPages.unshift({ key: 'dashboard', title: 'Dashboard', icon: 'fas fa-chart-line' });
    }

    return Object.keys(byCategory).map(function (key) {
      const cat = byCategory[key];
      cat.pages.sort(function (a, b) {
        return a.title.localeCompare(b.title);
      });
      return cat;
    }).sort(function (a, b) {
      return a.name.localeCompare(b.name);
    });
  }

  // ---------------------------------------------------------------------------
  // Session management
  // ---------------------------------------------------------------------------

  function createSession(userId, rememberMe, context) {
    const token = generateId('sess');
    const now = new Date();
    const expires = new Date(now.getTime() + (rememberMe ? 1000 * 60 * 60 * 24 * 30 : 1000 * 60 * 60 * 12));
    appendRow(SHEETS.SESSIONS, HEADERS.SESSIONS, [
      token,
      userId,
      now.toISOString(),
      expires.toISOString(),
      rememberMe ? 1 : 0,
      context && context.campaignId ? context.campaignId : '',
      context && context.userAgent ? context.userAgent : '',
      context && context.ip ? context.ip : ''
    ]);
    return {
      token: token,
      expiresAt: expires.toISOString()
    };
  }

  function findSession(token) {
    if (!token) return null;
    return withCache(['session', token], CACHE_TTL.SHORT, function () {
      const row = findRowByColumn(SHEETS.SESSIONS, HEADERS.SESSIONS, 'Token', token);
      if (!row) return null;
      return row;
    });
  }

  function removeSession(token) {
    if (!token) return;
    const row = findRowByColumn(SHEETS.SESSIONS, HEADERS.SESSIONS, 'Token', token);
    if (row && row.__rowIndex) {
      getSheet(SHEETS.SESSIONS, HEADERS.SESSIONS).deleteRow(row.__rowIndex);
      clearTableCache(SHEETS.SESSIONS);
      try {
        scriptCache.remove(cacheKey(['session', token]));
      } catch (err) {
        logError('SessionCacheRemove', err);
      }
    }
  }

  function requireSession(token) {
    if (!token) {
      throw new Error('Session token missing');
    }
    const session = findSession(token);
    if (!session) {
      throw new Error('Session not found or expired');
    }
    const expires = session.ExpiresAt ? new Date(session.ExpiresAt) : null;
    if (expires && expires < new Date()) {
      removeSession(token);
      throw new Error('Session expired');
    }
    const user = findUserById(session.UserId);
    if (!user) {
      throw new Error('User not found for session');
    }
    if (!toBool(user.CanLogin)) {
      throw new Error('User account disabled');
    }
    return {
      token: token,
      user: user,
      context: {
        campaignId: session.CampaignScope || ''
      }
    };
  }

  // ---------------------------------------------------------------------------
  // Application services
  // ---------------------------------------------------------------------------

  function login(payload, context) {
    ensureBootstrap();
    const identity = payload && payload.identity ? String(payload.identity).trim() : '';
    const password = payload && payload.password ? String(payload.password) : '';
    const remember = !!(payload && payload.remember);

    if (!identity || !password) {
      throw new Error('Username/email and password are required');
    }

    const user = findUserByIdentity(identity);
    if (!user) {
      throw new Error('Account not found');
    }
    if (!toBool(user.CanLogin)) {
      throw new Error('Account disabled');
    }

    if (!verifyPassword(password, user.PasswordHash)) {
      throw new Error('Invalid credentials');
    }

    const session = createSession(user.ID, remember, context || {});

    return {
      token: session.token,
      expiresAt: session.expiresAt,
      user: sanitizeUser(user),
      campaigns: getUserCampaigns(user.ID)
    };
  }

  function logout(payload) {
    ensureBootstrap();
    const token = payload && payload.token;
    removeSession(token);
    return { success: true };
  }

  function getBootstrap(payload) {
    ensureBootstrap();
    const session = requireSession(payload && payload.token);
    const user = sanitizeUser(session.user);
    const campaigns = getUserCampaigns(user.ID || user.Id);
    const primary = campaigns.find(function (c) { return c.isPrimary; }) || campaigns[0] || null;
    const navigation = getNavigationForUser(session.user, primary ? primary.id : '');
    return {
      user: user,
      token: session.token,
      campaigns: campaigns,
      navigation: navigation
    };
  }

  function getDashboard(payload) {
    ensureBootstrap();
    const session = requireSession(payload && payload.token);
    const user = sanitizeUser(session.user);
    const campaigns = getUserCampaigns(user.ID || user.Id);
    const notifications = getUserNotifications(user.ID || user.Id);

    return {
      user: user,
      campaigns: campaigns,
      notifications: notifications,
      generatedAt: nowIso()
    };
  }

  function getNotifications(payload) {
    ensureBootstrap();
    const session = requireSession(payload && payload.token);
    return getUserNotifications(session.user.ID || session.user.Id);
  }

  // ---------------------------------------------------------------------------
  // Rendering (HTML)
  // ---------------------------------------------------------------------------

  function renderTemplate(name, data) {
    const template = HtmlService.createTemplateFromFile(name);
    template.app = {
      baseUrl: ScriptApp.getService().getUrl(),
      data: data || {}
    };
    return template.evaluate()
      .setTitle('Lumina Sheets')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  function doGet(e) {
    try {
      ensureBootstrap();
      const page = (e && e.parameter && e.parameter.page) ? String(e.parameter.page).toLowerCase() : '';
      if (page === 'login') {
        return renderTemplate('V2_Login', {});
      }
      return renderTemplate('V2_AppShell', {
        initialView: page || 'dashboard'
      });
    } catch (err) {
      logError('doGet', err);
      const tpl = HtmlService.createHtmlOutput('<h1>System Error</h1><p>' + err.message + '</p>');
      tpl.setTitle('System Error');
      tpl.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      return tpl;
    }
  }

  function doPost(e) {
    try {
      ensureBootstrap();
      const body = e && e.postData && e.postData.contents ? safeJsonParse(e.postData.contents, {}) : {};
      const action = (body && body.action) ? String(body.action) : '';
      switch (action) {
        case 'login':
          return ContentService.createTextOutput(JSON.stringify(login(body.payload || {}, {})))
            .setMimeType(ContentService.MimeType.JSON);
        case 'logout':
          return ContentService.createTextOutput(JSON.stringify(logout(body.payload || {})))
            .setMimeType(ContentService.MimeType.JSON);
        default:
          throw new Error('Unsupported action: ' + action);
      }
    } catch (err) {
      logError('doPost', err);
      return ContentService.createTextOutput(JSON.stringify({ error: err.message || 'Unexpected error' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // expose API for google.script.run bindings
  return {
    doGet: doGet,
    doPost: doPost,
    login: login,
    logout: logout,
    getBootstrap: getBootstrap,
    getDashboard: getDashboard,
    getNotifications: getNotifications
  };
})();

// -----------------------------------------------------------------------------
// Global entry points for Apps Script runtime & client bindings
// -----------------------------------------------------------------------------

function doGet(e) {
  return LuminaAppV2.doGet(e || {});
}

function doPost(e) {
  return LuminaAppV2.doPost(e || {});
}

function appLogin(payload) {
  return LuminaAppV2.login(payload || {}, {});
}

function appLogout(payload) {
  return LuminaAppV2.logout(payload || {});
}

function appBootstrap(payload) {
  return LuminaAppV2.getBootstrap(payload || {});
}

function appDashboard(payload) {
  return LuminaAppV2.getDashboard(payload || {});
}

function appNotifications(payload) {
  return LuminaAppV2.getNotifications(payload || {});
}

