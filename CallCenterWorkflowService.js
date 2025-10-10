/**
 * CallCenterWorkflowService.gs
 * Unifies all major call center workflows (authentication, scheduling, performance,
 * coaching, reporting, collaboration) on top of the centralized DatabaseManager.
 *
 * Responsibilities
 *  - Register sheet schemas with DatabaseManager/DatabaseBindings so every table is
 *    addressable through tenant-aware CRUD helpers
 *  - Expose tenant-scoped orchestration helpers that hydrate dashboards with
 *    schedules, performance metrics, coaching queues, and collaboration feeds
 *  - Provide write helpers for common manager/agent flows (scheduling, attendance,
 *    QA reviews, coaching acknowledgements, chat messages)
 */
(function (global) {
  if (global.CallCenterWorkflowService) return;

  var WorkflowService = {};
  var initialized = false;
  var tableRegistry = {};
  var CAMPAIGN_BROADCAST_TYPE = 'CampaignBroadcast';
  var DEFAULT_COMMUNICATION_SEVERITY = 'Info';

  // ───────────────────────────────────────────────────────────────────────────────
  // Table registration helpers
  // ───────────────────────────────────────────────────────────────────────────────

  function toStr(value) {
    if (value === null || typeof value === 'undefined') return '';
    return String(value).trim();
  }

  function toNumber(value) {
    if (value === null || typeof value === 'undefined' || value === '') return null;
    if (value instanceof Date) return value.getTime();
    if (typeof value === 'number') return isNaN(value) ? null : value;
    var str = String(value).replace(/[^0-9.\-]/g, '');
    if (!str) return null;
    var num = parseFloat(str);
    return isNaN(num) ? null : num;
  }

  function toDate(value) {
    if (!value && value !== 0) return null;
    if (value instanceof Date) return new Date(value.getTime());
    var num = toNumber(value);
    if (typeof value === 'string') {
      var parsed = new Date(value);
      if (!isNaN(parsed.getTime())) return parsed;
    }
    if (num === null) return null;
    var d = new Date(num);
    return isNaN(d.getTime()) ? null : d;
  }

  function nowIso() {
    if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.formatDate === 'function') {
      var tz = 'UTC';
      if (typeof Session !== 'undefined' && Session && typeof Session.getScriptTimeZone === 'function') {
        tz = Session.getScriptTimeZone() || tz;
      }
      return Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ssXXX");
    }
    return new Date().toISOString();
  }

  function newUuid() {
    if (typeof Utilities !== 'undefined' && Utilities && Utilities.getUuid) {
      return Utilities.getUuid();
    }
    return 'uuid-' + Date.now() + '-' + Math.floor(Math.random() * 1000000);
  }

  function buildValueSet(values) {
    var set = {};
    (values || []).forEach(function (value) {
      var key = toStr(value);
      if (key) set[key] = true;
    });
    return set;
  }

  function dedupeList(values) {
    var seen = {};
    var out = [];
    (values || []).forEach(function (value) {
      var key = toStr(value);
      if (!key) return;
      if (!seen[key]) {
        seen[key] = true;
        out.push(key);
      }
    });
    return out;
  }

  function pickFirst(values) {
    if (!values || !values.length) return null;
    for (var i = 0; i < values.length; i++) {
      var key = toStr(values[i]);
      if (key) return key;
    }
    return null;
  }

  function getHeadersFromGlobal(name) {
    var value = global[name];
    if (Array.isArray(value)) return value.slice();
    return null;
  }

  function preferTenantColumn(headers, preference) {
    var candidates = [];
    if (Array.isArray(preference)) {
      candidates = candidates.concat(preference);
    } else if (preference) {
      candidates.push(preference);
    }
    candidates = candidates.concat(['CampaignID', 'CampaignId', 'campaignId', 'TenantID', 'TenantId', 'tenantId']);
    if (!Array.isArray(headers)) headers = [];
    var normalized = headers.map(function (h) { return toStr(h); });
    for (var i = 0; i < candidates.length; i++) {
      var idx = normalized.indexOf(toStr(candidates[i]));
      if (idx !== -1) return headers[idx];
    }
    return null;
  }

  function clone(obj) {
    if (!obj || typeof obj !== 'object') return obj;
    var copy = {};
    Object.keys(obj).forEach(function (key) {
      copy[key] = obj[key];
    });
    return copy;
  }

  function tableDefinitions() {
    var defs = [
      { key: 'users', name: global.USERS_SHEET || 'Users', headerVar: 'USERS_HEADERS', idColumn: 'ID', preferTenantColumns: ['CampaignID', 'CampaignId'], requireTenant: false, cacheTTL: 1800 },
      { key: 'roles', name: global.ROLES_SHEET || 'Roles', headerVar: 'ROLES_HEADER', idColumn: 'ID', cacheTTL: 3600 },
      { key: 'userRoles', name: global.USER_ROLES_SHEET || 'UserRoles', headerVar: 'USER_ROLES_HEADER', idColumn: null, cacheTTL: 1800 },
      { key: 'userClaims', name: global.USER_CLAIMS_SHEET || 'UserClaims', headerVar: 'CLAIMS_HEADERS', idColumn: 'ID', cacheTTL: 1800 },
      { key: 'sessions', name: global.SESSIONS_SHEET || 'Sessions', headerVar: 'SESSIONS_HEADERS', idColumn: 'TokenHash' },
      { key: 'campaigns', name: global.CAMPAIGNS_SHEET || 'Campaigns', headerVar: 'CAMPAIGNS_HEADERS', idColumn: 'ID', requireTenant: false, cacheTTL: 3600 },
      { key: 'campaignPermissions', name: global.CAMPAIGN_USER_PERMISSIONS_SHEET || 'CampaignUserPermissions', headerVar: 'CAMPAIGN_USER_PERMISSIONS_HEADERS', idColumn: 'ID', preferTenantColumns: ['CampaignID', 'CampaignId'], requireTenant: true, cacheTTL: 2700 },
      { key: 'userCampaigns', name: global.USER_CAMPAIGNS_SHEET || 'UserCampaigns', headerVar: 'USER_CAMPAIGNS_HEADERS', idColumn: 'ID', preferTenantColumns: ['CampaignId', 'CampaignID'], requireTenant: true, cacheTTL: 1800 },
      { key: 'userManagers', name: global.USER_MANAGERS_SHEET || 'UserManagers', headerVar: 'USER_MANAGERS_HEADERS', idColumn: 'ID', preferTenantColumns: ['CampaignID', 'CampaignId'], requireTenant: true, cacheTTL: 2700 },
      { key: 'notifications', name: global.NOTIFICATIONS_SHEET || 'Notifications', headerVar: 'NOTIFICATIONS_HEADERS', idColumn: 'ID', preferTenantColumns: ['CampaignID', 'CampaignId'], requireTenant: false, cacheTTL: 1800 },
      { key: 'schedules', name: typeof global.SCHEDULE_GENERATION_SHEET === 'string' ? global.SCHEDULE_GENERATION_SHEET : 'GeneratedSchedules', headerVar: 'SCHEDULE_GENERATION_HEADERS', idColumn: 'ID', preferTenantColumns: ['CampaignID', 'CampaignId'], requireTenant: false },
      { key: 'shiftSlots', name: typeof global.SHIFT_SLOTS_SHEET === 'string' ? global.SHIFT_SLOTS_SHEET : 'ShiftSlots', headerVar: 'SHIFT_SLOTS_HEADERS', idColumn: 'ID', preferTenantColumns: ['CampaignID', 'CampaignId'], requireTenant: false },
      { key: 'adherence', name: typeof global.SCHEDULE_ADHERENCE_SHEET === 'string' ? global.SCHEDULE_ADHERENCE_SHEET : 'ScheduleAdherence', headerVar: 'SCHEDULE_ADHERENCE_HEADERS', idColumn: 'ID', preferTenantColumns: ['CampaignID', 'CampaignId'], requireTenant: false },
      { key: 'attendanceLog', name: typeof global.ATTENDANCE_LOG_SHEET === 'string' ? global.ATTENDANCE_LOG_SHEET : (global.ATTENDANCE_SHEET || 'AttendanceLog'), headerVar: 'ATTENDANCE_LOG_HEADERS', idColumn: 'ID', preferTenantColumns: ['CampaignID', 'CampaignId'], requireTenant: false },
      { key: 'attendanceStatus', name: typeof global.ATTENDANCE_STATUS_SHEET === 'string' ? global.ATTENDANCE_STATUS_SHEET : 'AttendanceStatus', headerVar: 'ATTENDANCE_STATUS_HEADERS', idColumn: 'ID', preferTenantColumns: ['CampaignID', 'CampaignId'], requireTenant: false },
      { key: 'qaRecords', name: typeof global.QA_RECORDS === 'string' ? global.QA_RECORDS : 'QA Records', headerVar: 'QA_HEADERS', idColumn: 'ID', preferTenantColumns: ['CampaignID', 'CampaignId'], requireTenant: false, cacheTTL: 2700 },
      { key: 'coaching', name: typeof global.COACHING_SHEET === 'string' ? global.COACHING_SHEET : 'CoachingRecords', headerVar: 'COACHING_HEADERS', idColumn: 'ID', preferTenantColumns: ['CampaignID', 'CampaignId'], requireTenant: false, cacheTTL: 2700 },
      { key: 'chatMessages', name: global.CHAT_MESSAGES_SHEET || 'ChatMessages', headerVar: 'CHAT_MESSAGES_HEADERS', idColumn: 'ID', preferTenantColumns: ['CampaignID', 'CampaignId'], requireTenant: false },
      { key: 'chatChannels', name: global.CHAT_CHANNELS_SHEET || 'ChatChannels', headerVar: 'CHAT_CHANNELS_HEADERS', idColumn: 'ID', preferTenantColumns: ['CampaignID', 'CampaignId'], requireTenant: false, cacheTTL: 1800 },
      { key: 'chatGroups', name: global.CHAT_GROUPS_SHEET || 'ChatGroups', headerVar: 'CHAT_GROUPS_HEADERS', idColumn: 'ID', preferTenantColumns: ['CampaignID', 'CampaignId'], requireTenant: false, cacheTTL: 1800 },
      { key: 'chatMemberships', name: global.CHAT_CHANNEL_MEMBERS_SHEET || 'ChatChannelMembers', headerVar: 'CHAT_CHANNEL_MEMBERS_HEADERS', idColumn: 'ID', preferTenantColumns: ['CampaignID', 'CampaignId'], requireTenant: false, cacheTTL: 1800 }
    ];
    return defs;
  }

  function ensureInitialized() {
    if (initialized) return tableRegistry;
    tableRegistry = {};
    var defs = tableDefinitions();
    for (var i = 0; i < defs.length; i++) {
      registerTable(defs[i]);
    }
    initialized = true;
    return tableRegistry;
  }

  function registerTable(def) {
    if (!def || !def.name) return;
    var headers = def.headers ? def.headers.slice() : getHeadersFromGlobal(def.headerVar) || [];
    var tenantColumn = def.tenantColumn || preferTenantColumn(headers, def.preferTenantColumns);
    var config = { headers: Array.isArray(headers) ? headers.slice() : undefined };
    if (Object.prototype.hasOwnProperty.call(def, 'idColumn')) {
      config.idColumn = def.idColumn;
    }
    if (tenantColumn) {
      config.tenantColumn = tenantColumn;
      if (typeof def.requireTenant !== 'undefined') {
        config.requireTenant = def.requireTenant;
      } else {
        config.requireTenant = true;
      }
    } else if (typeof def.requireTenant !== 'undefined') {
      config.requireTenant = def.requireTenant;
    }
    if (typeof def.cacheTTL === 'number' && isFinite(def.cacheTTL) && def.cacheTTL > 0) {
      config.cacheTTL = def.cacheTTL;
    }
    try {
      if (typeof global.registerTableSchema === 'function') {
        var schema = global.registerTableSchema(def.name, config);
        tableRegistry[def.key] = {
          name: def.name,
          idColumn: schema && Object.prototype.hasOwnProperty.call(schema, 'idColumn') ? schema.idColumn : config.idColumn,
          tenantColumn: schema && schema.tenantColumn ? schema.tenantColumn : tenantColumn,
          headers: schema && schema.headers ? schema.headers.slice() : headers.slice(),
          requireTenant: schema && Object.prototype.hasOwnProperty.call(schema, 'requireTenant') ? schema.requireTenant : config.requireTenant,
          cacheTTL: schema && Object.prototype.hasOwnProperty.call(schema, 'cacheTTL') ? schema.cacheTTL : config.cacheTTL
        };
        return;
      }
    } catch (err) {
      logError('registerTable:' + def.name, err);
    }
    tableRegistry[def.key] = {
      name: def.name,
      idColumn: config.idColumn,
      tenantColumn: tenantColumn,
      headers: headers.slice(),
      requireTenant: config.requireTenant,
      cacheTTL: config.cacheTTL
    };
  }

  function getTable(key) {
    ensureInitialized();
    return tableRegistry[key] || null;
  }

  // ───────────────────────────────────────────────────────────────────────────────
  // CRUD wrappers
  // ───────────────────────────────────────────────────────────────────────────────

  function crudFunctions() {
    return {
      select: (typeof global.dbTenantSelect === 'function') ? global.dbTenantSelect : function (name, context, options) {
        if (typeof global.dbSelect === 'function') {
          return global.dbSelect(name, options || {}, context);
        }
        return [];
      },
      create: (typeof global.dbTenantCreate === 'function') ? global.dbTenantCreate : function (name, context, record) {
        if (typeof global.dbCreate === 'function') {
          return global.dbCreate(name, record, context);
        }
        return record;
      },
      update: (typeof global.dbTenantUpdate === 'function') ? global.dbTenantUpdate : function (name, context, identifier, updates) {
        if (typeof global.dbUpdate === 'function') {
          return global.dbUpdate(name, identifier, updates, context);
        }
        return updates;
      }
    };
  }

  function safeSelect(key, context, options) {
    ensureInitialized();
    var table = getTable(key);
    if (!table) return [];
    var funcs = crudFunctions();
    try {
      return funcs.select(table.name, context || null, options || {}) || [];
    } catch (err) {
      logError('select:' + key, err);
      return [];
    }
  }

  function safeCreate(key, context, record) {
    ensureInitialized();
    var table = getTable(key);
    if (!table) throw new Error('Table not registered: ' + key);
    var funcs = crudFunctions();
    try {
      return funcs.create(table.name, context || null, record);
    } catch (err) {
      logError('create:' + key, err);
      throw err;
    }
  }

  function safeUpdate(key, context, identifier, updates) {
    ensureInitialized();
    var table = getTable(key);
    if (!table) throw new Error('Table not registered: ' + key);
    var funcs = crudFunctions();
    try {
      return funcs.update(table.name, context || null, identifier, updates);
    } catch (err) {
      logError('update:' + key, err);
      throw err;
    }
  }

  // ───────────────────────────────────────────────────────────────────────────────
  // Tenant helpers
  // ───────────────────────────────────────────────────────────────────────────────

  function getTenantTools(userId, campaignId, options) {
    ensureInitialized();
    options = options || {};
    if (global.TenantSecurity && typeof global.TenantSecurity.getTenantContext === 'function') {
      if (campaignId && typeof global.TenantSecurity.assertCampaignAccess === 'function') {
        try {
          global.TenantSecurity.assertCampaignAccess(userId, campaignId, options);
        } catch (err) {
          logError('assertCampaignAccess', err);
          throw err;
        }
      }
      try {
        var ctxInfo = global.TenantSecurity.getTenantContext(userId, campaignId, options);
        return {
          profile: ctxInfo.profile,
          context: ctxInfo.context || {}
        };
      } catch (err) {
        logError('getTenantContext', err);
        throw err;
      }
    }
    var context = {};
    if (campaignId) context.tenantId = toStr(campaignId);
    return { profile: null, context: context };
  }

  function assignTenant(record, table, context, fallbackCampaignId) {
    if (!table || !table.tenantColumn) return;
    var tenantColumn = table.tenantColumn;
    if (!record[tenantColumn]) {
      var candidate = fallbackCampaignId || null;
      if (!candidate && context) {
        if (context.tenantId) candidate = context.tenantId;
        else if (Array.isArray(context.tenantIds) && context.tenantIds.length === 1) {
          candidate = context.tenantIds[0];
        }
      }
      if (candidate) {
        record[tenantColumn] = candidate;
      }
    }
  }

  function loadUserById(userId, context) {
    if (!userId) return null;
    var options = { where: { ID: userId } };
    var rows = safeSelect('users', context, options);
    if (rows && rows.length) return rows[0];
    var fallbackCtx = clone(context || {});
    fallbackCtx.allowAllTenants = true;
    rows = safeSelect('users', fallbackCtx, options);
    return rows && rows.length ? rows[0] : null;
  }

  function normalizeUserIdentifiers(record, userId) {
    var id = toStr(userId);
    if (!id) return;
    if (typeof record.UserID === 'undefined') record.UserID = id;
    if (typeof record.UserId === 'undefined') record.UserId = id;
    if (typeof record.AgentID === 'undefined') record.AgentID = id;
    if (typeof record.AgentId === 'undefined') record.AgentId = id;
  }

  function inferCampaignFromRow(row) {
    if (!row || typeof row !== 'object') return '';
    return toStr(row.CampaignID || row.CampaignId || row.campaignId || row.TenantID || row.TenantId);
  }

  function logError(scope, err) {
    if (typeof console !== 'undefined' && console && console.error) {
      console.error('[CallCenterWorkflowService:' + scope + ']', err);
    }
    if (typeof global.safeWriteError === 'function') {
      try { global.safeWriteError('CallCenterWorkflowService.' + scope, err); } catch (_) { }
    }
  }

  // ───────────────────────────────────────────────────────────────────────────────
  // Data aggregation helpers
  // ───────────────────────────────────────────────────────────────────────────────

  function buildUserIndex(users) {
    var byId = {};
    var byCampaign = {};
    var ids = [];
    (users || []).forEach(function (user) {
      var id = toStr(user.ID || user.Id || user.UserId || user.UserID);
      if (!id) return;
      byId[id] = user;
      ids.push(id);
      var campaign = inferCampaignFromRow(user);
      if (campaign) {
        if (!byCampaign[campaign]) byCampaign[campaign] = [];
        byCampaign[campaign].push(id);
      }
    });
    return {
      byId: byId,
      byCampaign: byCampaign,
      allowedIds: ids,
      allowedSet: buildValueSet(ids)
    };
  }

  function matchesAllowedUser(row, allowedSet) {
    if (!row || !allowedSet) return false;
    var keys = ['UserID', 'UserId', 'AgentID', 'AgentId', 'EmployeeID', 'EmployeeId', 'AssigneeId'];
    for (var i = 0; i < keys.length; i++) {
      var key = toStr(row[keys[i]]);
      if (key && allowedSet[key]) return true;
    }
    return false;
  }

  function filterByProfile(rows, profile, userIndex) {
    rows = rows || [];
    if (!profile) return rows;
    if (profile.isGlobalAdmin) return rows;
    var allowedCampaigns = buildValueSet(profile.allowedCampaignIds || []);
    return rows.filter(function (row) {
      var campaign = inferCampaignFromRow(row);
      if (campaign && allowedCampaigns[campaign]) return true;
      return matchesAllowedUser(row, userIndex.allowedSet);
    });
  }

  function summarizeQa(rows) {
    var total = rows.length;
    var scoreSum = 0;
    var scoreCount = 0;
    var failing = 0;
    rows.forEach(function (row) {
      var val = toNumber(row.FinalScore || row.Score || row.TotalScore || row.Percentage || row.Percent);
      if (val !== null) {
        scoreSum += val;
        scoreCount += 1;
        if (val < 80) failing += 1;
      }
    });
    return {
      totalEvaluations: total,
      averageScore: scoreCount ? Math.round((scoreSum / scoreCount) * 100) / 100 : null,
      failingEvaluations: failing
    };
  }

  function summarizeAttendance(rows) {
    var total = rows.length;
    var statusCounts = {};
    var absent = 0;
    var late = 0;
    rows.forEach(function (row) {
      var status = toStr(row.Status || row.State || row.AttendanceStatus || row.Result);
      if (!status) status = 'Unspecified';
      statusCounts[status] = (statusCounts[status] || 0) + 1;
      if (/absent|no show|noshow|callout/i.test(status)) absent += 1;
      if (/late/i.test(status) || toNumber(row.MinutesLate || row.LateMinutes) > 0) late += 1;
    });
    var present = total - absent;
    return {
      totalEvents: total,
      statusCounts: statusCounts,
      absentCount: absent,
      lateCount: late,
      attendanceRate: total ? Math.round((present / Math.max(total, 1)) * 10000) / 100 : null
    };
  }

  function summarizeAdherence(rows) {
    var total = rows.length;
    var scoreSum = 0;
    var scoreCount = 0;
    rows.forEach(function (row) {
      var score = toNumber(row.AdherenceScore || row.Score || row.Percent || row.Percentage);
      if (score !== null) {
        scoreSum += score;
        scoreCount += 1;
      }
    });
    return {
      totalEvents: total,
      averageScore: scoreCount ? Math.round((scoreSum / scoreCount) * 100) / 100 : null
    };
  }

  function summarizeCoaching(rows) {
    var summary = {
      totalSessions: rows.length,
      pendingCount: 0,
      acknowledgedCount: 0,
      overdueCount: 0
    };
    var today = new Date();
    rows.forEach(function (row) {
      var status = toStr(row.Status || row.AcknowledgementStatus || row.State || row.Stage);
      if (!status) status = 'Pending';
      if (/ack/i.test(status)) summary.acknowledgedCount += 1;
      else summary.pendingCount += 1;
      var due = toDate(row.DueDate || row.FollowUpDate || row.AcknowledgeBy);
      var ack = toDate(row.AcknowledgedAt || row.AcknowledgementDate);
      if (!ack && due && due < today) summary.overdueCount += 1;
    });
    return summary;
  }

  function loadAuthentication(profile, context, options) {
    options = options || {};
    var authUser = profile ? profile.user : null;
    var roles = safeSelect('roles', { allowAllTenants: true }, {});
    var userRoles = [];
    var userClaims = [];
    var sessions = [];
    if (authUser) {
      userRoles = safeSelect('userRoles', null, { where: { UserId: authUser.ID || authUser.Id } });
      userClaims = safeSelect('userClaims', null, { where: { UserId: authUser.ID || authUser.Id } });
      sessions = safeSelect('sessions', null, { where: { UserId: authUser.ID || authUser.Id } });
    }
    return {
      user: authUser,
      roles: roles,
      userRoles: userRoles,
      claims: userClaims,
      activeSessions: sessions
    };
  }

  function loadUsers(profile, context, options) {
    options = options || {};
    var rows = safeSelect('users', context, {});
    rows = filterByProfile(rows, profile, buildUserIndex(rows));
    if (options.activeOnly) {
      rows = rows.filter(function (row) {
        var status = toStr(row.EmploymentStatus || row.Status || row.State);
        return !status || /active/i.test(status);
      });
    }
    return rows;
  }

  function loadScheduling(profile, context, userIndex, options) {
    var rows = safeSelect('schedules', context, {});
    rows = filterByProfile(rows, profile, userIndex);
    var today = [];
    var upcoming = [];
    var now = new Date();
    var startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    var startOfTomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
    rows.forEach(function (row) {
      var date = toDate(row.Date || row.ScheduleDate || row.ShiftDate || row.StartDate);
      if (!date) return;
      if (date >= startOfToday && date < startOfTomorrow) {
        today.push(row);
      } else if (date >= startOfTomorrow) {
        upcoming.push(row);
      }
    });
    upcoming.sort(function (a, b) {
      var da = toDate(a.Date || a.ScheduleDate || a.ShiftDate || a.StartDate) || new Date(0);
      var db = toDate(b.Date || b.ScheduleDate || b.ShiftDate || b.StartDate) || new Date(0);
      return da.getTime() - db.getTime();
    });
    return {
      entries: rows,
      today: today,
      upcoming: upcoming,
      totalShifts: rows.length
    };
  }

  function loadPerformance(profile, context, userIndex) {
    var qaRows = filterByProfile(safeSelect('qaRecords', context, {}), profile, userIndex);
    var attendanceRows = filterByProfile(safeSelect('attendanceLog', context, {}), profile, userIndex);
    var adherenceRows = filterByProfile(safeSelect('adherence', context, {}), profile, userIndex);
    return {
      qa: { rows: qaRows, summary: summarizeQa(qaRows) },
      attendance: { rows: attendanceRows, summary: summarizeAttendance(attendanceRows) },
      adherence: { rows: adherenceRows, summary: summarizeAdherence(adherenceRows) }
    };
  }

  function loadCoaching(profile, context, userIndex) {
    var rows = filterByProfile(safeSelect('coaching', context, {}), profile, userIndex);
    var summary = summarizeCoaching(rows);
    var pending = rows.filter(function (row) {
      var status = toStr(row.Status || row.AcknowledgementStatus || row.State || row.Stage);
      if (!status) status = 'Pending';
      return !/ack/i.test(status);
    });
    return {
      rows: rows,
      pending: pending,
      summary: summary
    };
  }

  function loadCollaboration(profile, context, userIndex, options) {
    options = options || {};
    var rows = filterByProfile(safeSelect('chatMessages', context, { filter: function (row) {
      var deleted = row.IsDeleted;
      if (typeof deleted === 'boolean') return !deleted;
      var str = toStr(deleted);
      return !(str && (str === 'TRUE' || str === 'true' || str === '1'));
    } }), profile, userIndex);
    rows.sort(function (a, b) {
      var ta = toDate(a.Timestamp || a.CreatedAt || a.SentAt) || new Date(0);
      var tb = toDate(b.Timestamp || b.CreatedAt || b.SentAt) || new Date(0);
      return tb.getTime() - ta.getTime();
    });
    var limit = options.maxMessages || 25;
    var latest = rows.slice(0, limit);
    return {
      messages: latest,
      totalMessages: rows.length
    };
  }

  function loadCampaignCommunications(profile, context, userIndex, options) {
    options = options || {};
    var typeFilter = toStr(options.type) || CAMPAIGN_BROADCAST_TYPE;
    var preferredLimit = typeof options.limit === 'number' ? options.limit : options.maxMessages;
    var limit = typeof preferredLimit === 'number' && preferredLimit > 0 ? preferredLimit : 25;
    var rows = filterByProfile(safeSelect('notifications', context, {}), profile, userIndex);
    rows = rows.filter(function (row) {
      var type = toStr(row.Type || row.NotificationType);
      return !typeFilter || type === typeFilter;
    });
    rows.sort(function (a, b) {
      var ta = toDate(a.CreatedAt || a.Timestamp || a.SentAt) || new Date(0);
      var tb = toDate(b.CreatedAt || b.Timestamp || b.SentAt) || new Date(0);
      return tb.getTime() - ta.getTime();
    });
    return {
      records: rows.slice(0, limit),
      total: rows.length
    };
  }

  function buildCampaignSnapshot(profile, context, options) {
    options = options || {};
    var users = loadUsers(profile, context, { activeOnly: !!options.activeOnly });
    var userIndex = buildUserIndex(users);
    var scheduling = loadScheduling(profile, context, userIndex, options);
    var performance = loadPerformance(profile, context, userIndex);
    var coaching = loadCoaching(profile, context, userIndex);
    var communications = loadCampaignCommunications(profile, context, userIndex, options);
    var reporting = buildReporting(profile, context, userIndex, {
      scheduling: scheduling,
      performance: performance,
      coaching: coaching
    });
    return {
      roster: {
        total: users.length,
        records: options.includeRoster ? users : undefined
      },
      scheduling: scheduling,
      performance: performance,
      coaching: coaching,
      communications: communications,
      reporting: reporting
    };
  }

  function buildReporting(profile, context, userIndex, aggregates) {
    var scheduling = aggregates.scheduling;
    var performance = aggregates.performance;
    var coaching = aggregates.coaching;
    var totalAgents = userIndex.allowedIds.length;
    var qaAvg = performance.qa.summary.averageScore;
    var attendanceRate = performance.attendance.summary.attendanceRate;
    var adherenceAvg = performance.adherence.summary.averageScore;
    var pendingCoaching = coaching.summary.pendingCount;
    return {
      totals: {
        agents: totalAgents,
        todayShifts: scheduling.today.length,
        upcomingShifts: scheduling.upcoming.length,
        qaEvaluations: performance.qa.summary.totalEvaluations,
        attendanceEvents: performance.attendance.summary.totalEvents,
        coachingSessions: coaching.summary.totalSessions
      },
      metrics: [
        { key: 'qaAverage', label: 'QA Average', value: qaAvg, format: 'percentage' },
        { key: 'attendanceRate', label: 'Attendance Rate', value: attendanceRate, format: 'percentage' },
        { key: 'adherenceAverage', label: 'Adherence Average', value: adherenceAvg, format: 'percentage' },
        { key: 'pendingCoaching', label: 'Pending Coaching', value: pendingCoaching, format: 'count' }
      ]
    };
  }

  function loadCampaigns(profile) {
    var campaigns = safeSelect('campaigns', { allowAllTenants: true }, {});
    if (!profile || profile.isGlobalAdmin) return campaigns;
    var allowed = buildValueSet(profile.allowedCampaignIds || []);
    return campaigns.filter(function (campaign) {
      var cid = toStr(campaign.ID || campaign.Id || campaign.id);
      return cid && allowed[cid];
    });
  }

  function buildCampaignAccess(profile) {
    var campaigns = safeSelect('campaigns', { allowAllTenants: true }, {});
    if (!profile) {
      return campaigns.map(function (campaign) {
        return {
          id: toStr(campaign.ID || campaign.Id || campaign.id),
          name: campaign.Name || campaign.name || '',
          description: campaign.Description || campaign.description || '',
          isManaged: false,
          isAdmin: false,
          isDefault: false
        };
      });
    }

    var managed = buildValueSet(profile.managedCampaignIds || []);
    var admin = buildValueSet(profile.adminCampaignIds || []);
    var allowed = buildValueSet(profile.allowedCampaignIds || []);
    var defaultCampaign = toStr(profile.defaultCampaignId);

    return campaigns
      .filter(function (campaign) {
        if (profile.isGlobalAdmin) return true;
        var cid = toStr(campaign.ID || campaign.Id || campaign.id);
        return cid && allowed[cid];
      })
      .map(function (campaign) {
        var cid = toStr(campaign.ID || campaign.Id || campaign.id);
        return {
          id: cid,
          name: campaign.Name || campaign.name || '',
          description: campaign.Description || campaign.description || '',
          isManaged: !!managed[cid],
          isAdmin: profile.isGlobalAdmin || !!admin[cid],
          isDefault: !!(cid && defaultCampaign && cid === defaultCampaign)
        };
      });
  }

  // ───────────────────────────────────────────────────────────────────────────────
  // Public API: Aggregations
  // ───────────────────────────────────────────────────────────────────────────────

  WorkflowService.initialize = function () {
    ensureInitialized();
    return clone(tableRegistry);
  };

  WorkflowService.getTenantContext = function (userId, campaignId, options) {
    return getTenantTools(userId, campaignId, options);
  };

  WorkflowService.getWorkspace = function (userId, options) {
    ensureInitialized();
    options = options || {};
    var campaignId = options.campaignId || null;
    var ctxInfo = getTenantTools(userId, campaignId, {});
    var profile = ctxInfo.profile;
    var context = ctxInfo.context || {};

    var users = loadUsers(profile, context, { activeOnly: !!options.activeOnly });
    var userIndex = buildUserIndex(users);
    userIndex.isGlobalAdmin = profile && profile.isGlobalAdmin;

    var scheduling = loadScheduling(profile, context, userIndex, options);
    var performance = loadPerformance(profile, context, userIndex);
    var coaching = loadCoaching(profile, context, userIndex);
    var collaboration = loadCollaboration(profile, context, userIndex, options);
    var reporting = buildReporting(profile, context, userIndex, {
      scheduling: scheduling,
      performance: performance,
      coaching: coaching
    });
    var authentication = loadAuthentication(profile, context, options);
    var campaigns = loadCampaigns(profile);

    return {
      profile: profile,
      context: context,
      campaigns: campaigns,
      users: { records: users, total: users.length },
      authentication: authentication,
      scheduling: scheduling,
      performance: performance,
      coaching: coaching,
      collaboration: collaboration,
      reporting: reporting
    };
  };

  WorkflowService.listCampaignAccess = function (userId) {
    ensureInitialized();
    var ctxInfo = getTenantTools(userId, null, {});
    return buildCampaignAccess(ctxInfo.profile || null);
  };

  WorkflowService.getManagerCampaignDashboard = function (managerId, campaignId, options) {
    ensureInitialized();
    if (!campaignId) throw new Error('getManagerCampaignDashboard requires a campaignId');
    var ctxInfo = getTenantTools(managerId, campaignId, { requireManager: true });
    var profile = ctxInfo.profile;
    var context = ctxInfo.context;
    var snapshot = buildCampaignSnapshot(profile, context, Object.assign({}, options, { includeRoster: true }));
    snapshot.campaign = (typeof csGetCampaignById === 'function') ? csGetCampaignById(campaignId) : { id: campaignId };
    return snapshot;
  };

  WorkflowService.sendCampaignCommunication = function (managerId, campaignId, payload) {
    ensureInitialized();
    if (!campaignId) throw new Error('sendCampaignCommunication requires a campaignId');
    payload = payload || {};
    var message = toStr(payload.message || payload.Message);
    if (!message) throw new Error('sendCampaignCommunication requires a message');
    var title = toStr(payload.title || payload.Title) || 'Campaign Update';
    var severity = toStr(payload.severity || payload.Severity) || DEFAULT_COMMUNICATION_SEVERITY;
    var ctxInfo = getTenantTools(managerId, campaignId, { requireManager: true });
    var profile = ctxInfo.profile;
    var context = ctxInfo.context;
    var table = getTable('notifications');
    if (!table) throw new Error('Notifications table is not registered');

    var roster = loadUsers(profile, context, {});
    var rosterIndex = buildUserIndex(roster);
    var allowedSet = rosterIndex.allowedSet;

    var recipients = [];
    var explicitIds = payload.userIds || payload.UserIds;
    if (Array.isArray(explicitIds) && explicitIds.length) {
      explicitIds.forEach(function (id) {
        var key = toStr(id);
        if (key && allowedSet[key]) recipients.push(key);
      });
    } else {
      recipients = rosterIndex.allowedIds.slice();
    }

    recipients = dedupeList(recipients);
    if (!recipients.length) {
      return { success: false, error: 'No eligible recipients for campaign communication' };
    }

    var now = nowIso();
    var createdIds = [];
    var metadata = payload.metadata || payload.data || null;

    recipients.forEach(function (userId) {
      var record = {
        ID: newUuid(),
        UserId: userId,
        Type: payload.type || payload.Type || CAMPAIGN_BROADCAST_TYPE,
        Severity: severity,
        Title: title,
        Message: message,
        Read: false,
        ActionTaken: '',
        CreatedAt: now,
        ReadAt: '',
        ExpiresAt: payload.expiresAt || payload.ExpiresAt || ''
      };
      if (metadata) {
        try {
          record.Data = typeof metadata === 'string' ? metadata : JSON.stringify(metadata);
        } catch (_) {
          record.Data = '';
        }
      } else {
        record.Data = '';
      }
      assignTenant(record, table, context, campaignId);
      var result = safeCreate('notifications', context, record);
      createdIds.push(result && result.ID ? result.ID : record.ID);
    });

    return {
      success: true,
      recipients: createdIds.length,
      notificationIds: createdIds
    };
  };

  WorkflowService.getExecutiveAnalytics = function (executiveId, options) {
    ensureInitialized();
    options = options || {};
    var ctxInfo = getTenantTools(executiveId, null, {});
    var profile = ctxInfo.profile;
    if (!profile) throw new Error('User not found for executive analytics');

    var campaigns = buildCampaignAccess(profile);
    if (!campaigns.length) {
      return { campaigns: [], summary: { totalCampaigns: 0, totalAgents: 0 } };
    }

    var snapshots = [];
    campaigns.forEach(function (campaign) {
      var cid = campaign.id;
      if (!cid) return;
      try {
        var ctx = getTenantTools(executiveId, cid, {});
        var snapshot = buildCampaignSnapshot(ctx.profile || profile, ctx.context, options);
        snapshots.push({
          campaign: campaign,
          snapshot: snapshot
        });
      } catch (err) {
        logError('getExecutiveAnalytics:' + cid, err);
      }
    });

    var totals = {
      totalCampaigns: snapshots.length,
      totalAgents: 0,
      qaEvaluations: 0,
      qaScoreSum: 0,
      qaScoreCount: 0,
      attendanceEvents: 0,
      attendanceRateSum: 0,
      attendanceRateCount: 0
    };

    snapshots.forEach(function (entry) {
      var snapshot = entry.snapshot;
      totals.totalAgents += snapshot.roster.total;
      var qaSummary = snapshot.performance.qa.summary;
      totals.qaEvaluations += qaSummary.totalEvaluations;
      if (qaSummary.averageScore !== null && !isNaN(qaSummary.averageScore)) {
        totals.qaScoreSum += qaSummary.averageScore;
        totals.qaScoreCount += 1;
      }
      var attendanceSummary = snapshot.performance.attendance.summary;
      totals.attendanceEvents += attendanceSummary.totalEvents;
      if (attendanceSummary.attendanceRate !== null && !isNaN(attendanceSummary.attendanceRate)) {
        totals.attendanceRateSum += attendanceSummary.attendanceRate;
        totals.attendanceRateCount += 1;
      }
    });

    var summary = {
      totalCampaigns: totals.totalCampaigns,
      totalAgents: totals.totalAgents,
      averageQaScore: totals.qaScoreCount ? Math.round((totals.qaScoreSum / totals.qaScoreCount) * 100) / 100 : null,
      averageAttendanceRate: totals.attendanceRateCount ? Math.round((totals.attendanceRateSum / totals.attendanceRateCount) * 100) / 100 : null,
      totalQaEvaluations: totals.qaEvaluations,
      totalAttendanceEvents: totals.attendanceEvents
    };

    return {
      campaigns: snapshots,
      summary: summary
    };
  };

  // ───────────────────────────────────────────────────────────────────────────────
  // Public API: Authentication helpers
  // ───────────────────────────────────────────────────────────────────────────────

  WorkflowService.login = function () {
    return { success: false, error: 'Authentication has been removed from this deployment.' };
  };

  WorkflowService.issueSessionForUser = function () {
    return { success: false, error: 'Session issuance is unavailable because authentication has been removed.' };
  };

  // ───────────────────────────────────────────────────────────────────────────────
  // Public API: Scheduling + Attendance
  // ───────────────────────────────────────────────────────────────────────────────

  WorkflowService.scheduleAgentShift = function (managerId, payload) {
    ensureInitialized();
    payload = payload || {};
    var userId = toStr(payload.UserID || payload.UserId || payload.AgentId || payload.AgentID);
    if (!userId) throw new Error('scheduleAgentShift requires a User ID');

    var campaignId = toStr(payload.CampaignID || payload.CampaignId || payload.campaignId);
    var ctxInfo = getTenantTools(managerId, campaignId || null, { requireManager: !!campaignId });
    var context = ctxInfo.context;
    var table = getTable('schedules');
    if (!table) throw new Error('Schedules table is not registered');

    var agent = loadUserById(userId, context);
    if (!agent) {
      throw new Error('Agent not accessible for scheduling: ' + userId);
    }
    if (!campaignId) {
      campaignId = inferCampaignFromRow(agent);
    }

    var record = clone(payload);
    record.ID = record.ID || newUuid();
    record.CreatedAt = record.CreatedAt || nowIso();
    record.UpdatedAt = record.UpdatedAt || record.CreatedAt;
    normalizeUserIdentifiers(record, userId);
    assignTenant(record, table, context, campaignId);
    if (!record.Status) record.Status = 'Scheduled';
    if (!record.StartTime && record.Start) record.StartTime = record.Start;
    if (!record.EndTime && record.End) record.EndTime = record.End;

    return safeCreate('schedules', context, record);
  };

  WorkflowService.recordAttendanceEvent = function (actorId, payload) {
    ensureInitialized();
    payload = payload || {};
    var userId = toStr(payload.UserID || payload.UserId || payload.AgentId || payload.AgentID);
    if (!userId) throw new Error('recordAttendanceEvent requires a User ID');
    var campaignId = toStr(payload.CampaignID || payload.CampaignId || payload.campaignId);
    var ctxInfo = getTenantTools(actorId, campaignId || null, {});
    var context = ctxInfo.context;
    var table = getTable('attendanceLog');
    if (!table) throw new Error('Attendance log table is not registered');

    var agent = loadUserById(userId, context);
    if (!agent) {
      throw new Error('Agent not accessible for attendance logging: ' + userId);
    }
    if (!campaignId) {
      campaignId = inferCampaignFromRow(agent);
    }

    var record = clone(payload);
    record.ID = record.ID || newUuid();
    record.Timestamp = record.Timestamp || nowIso();
    record.Date = record.Date || record.AttendanceDate || record.ShiftDate;
    record.CreatedAt = record.CreatedAt || nowIso();
    record.UpdatedAt = record.UpdatedAt || record.CreatedAt;
    normalizeUserIdentifiers(record, userId);
    assignTenant(record, table, context, campaignId);
    return safeCreate('attendanceLog', context, record);
  };

  // ───────────────────────────────────────────────────────────────────────────────
  // Public API: Performance (QA)
  // ───────────────────────────────────────────────────────────────────────────────

  WorkflowService.logPerformanceReview = function (reviewerId, payload) {
    ensureInitialized();
    payload = payload || {};
    var userId = toStr(payload.UserID || payload.UserId || payload.AgentId || payload.AgentID);
    if (!userId) throw new Error('logPerformanceReview requires an evaluated User ID');
    var campaignId = toStr(payload.CampaignID || payload.CampaignId || payload.campaignId);
    var ctxInfo = getTenantTools(reviewerId, campaignId || null, { requireManager: !!campaignId });
    var context = ctxInfo.context;
    var table = getTable('qaRecords');
    if (!table) throw new Error('QA records table is not registered');

    var agent = loadUserById(userId, context);
    if (!agent) {
      throw new Error('Agent not accessible for QA logging: ' + userId);
    }
    if (!campaignId) campaignId = inferCampaignFromRow(agent);

    var record = clone(payload);
    record.ID = record.ID || newUuid();
    record.CreatedAt = record.CreatedAt || nowIso();
    record.UpdatedAt = record.UpdatedAt || record.CreatedAt;
    record.ReviewedBy = record.ReviewedBy || reviewerId;
    normalizeUserIdentifiers(record, userId);
    assignTenant(record, table, context, campaignId);
    if (!record.Status) record.Status = 'Completed';
    return safeCreate('qaRecords', context, record);
  };

  // ───────────────────────────────────────────────────────────────────────────────
  // Public API: Coaching
  // ───────────────────────────────────────────────────────────────────────────────

  WorkflowService.createCoachingSession = function (managerId, payload) {
    ensureInitialized();
    payload = payload || {};
    var userId = toStr(payload.UserID || payload.UserId || payload.AgentId || payload.AgentID);
    if (!userId) throw new Error('createCoachingSession requires a User ID');
    var campaignId = toStr(payload.CampaignID || payload.CampaignId || payload.campaignId);
    var ctxInfo = getTenantTools(managerId, campaignId || null, { requireManager: !!campaignId });
    var context = ctxInfo.context;
    var table = getTable('coaching');
    if (!table) throw new Error('Coaching table is not registered');

    var agent = loadUserById(userId, context);
    if (!agent) {
      throw new Error('Agent not accessible for coaching: ' + userId);
    }
    if (!campaignId) campaignId = inferCampaignFromRow(agent);

    var record = clone(payload);
    record.ID = record.ID || newUuid();
    record.CreatedAt = record.CreatedAt || nowIso();
    record.UpdatedAt = record.UpdatedAt || record.CreatedAt;
    record.CoachId = record.CoachId || managerId;
    normalizeUserIdentifiers(record, userId);
    assignTenant(record, table, context, campaignId);
    if (!record.Status) record.Status = 'Pending';
    return safeCreate('coaching', context, record);
  };

  WorkflowService.updateCoachingSession = function (managerId, coachingId, updates, campaignId) {
    ensureInitialized();
    if (!coachingId) throw new Error('updateCoachingSession requires a coaching session ID');
    updates = updates || {};
    var resolvedCampaign = toStr(campaignId || updates.CampaignID || updates.CampaignId || updates.campaignId);
    var ctxInfo = getTenantTools(managerId, resolvedCampaign || null, { requireManager: !!resolvedCampaign });
    var context = ctxInfo.context;
    var table = getTable('coaching');
    if (!table) throw new Error('Coaching table is not registered');

    if (resolvedCampaign) {
      updates.CampaignID = resolvedCampaign;
      updates.CampaignId = resolvedCampaign;
    }
    updates.UpdatedAt = nowIso();
    return safeUpdate('coaching', context, coachingId, updates);
  };

  WorkflowService.acknowledgeCoaching = function (agentId, coachingId, ackData) {
    ensureInitialized();
    if (!coachingId) throw new Error('acknowledgeCoaching requires a coaching session ID');
    var ctxInfo = getTenantTools(agentId, toStr(ackData && (ackData.CampaignID || ackData.CampaignId || ackData.campaignId)) || null, {});
    var context = ctxInfo.context;
    var table = getTable('coaching');
    if (!table) throw new Error('Coaching table is not registered');
    var updates = clone(ackData || {});
    updates.AcknowledgedAt = updates.AcknowledgedAt || nowIso();
    updates.AcknowledgedBy = updates.AcknowledgedBy || agentId;
    updates.Status = updates.Status || 'Acknowledged';
    updates.UpdatedAt = nowIso();
    return safeUpdate('coaching', context, coachingId, updates);
  };

  // ───────────────────────────────────────────────────────────────────────────────
  // Public API: Collaboration
  // ───────────────────────────────────────────────────────────────────────────────

  WorkflowService.postCollaborationMessage = function (userId, channelId, message, options) {
    ensureInitialized();
    if (!userId) throw new Error('postCollaborationMessage requires a user ID');
    if (!channelId) throw new Error('postCollaborationMessage requires a channel ID');
    var text = toStr(message);
    if (!text) throw new Error('Message body cannot be empty');
    options = options || {};
    var campaignId = toStr(options.CampaignID || options.CampaignId || options.campaignId);
    var ctxInfo = getTenantTools(userId, campaignId || null, {});
    var context = ctxInfo.context;
    var table = getTable('chatMessages');
    if (!table) throw new Error('Chat messages table is not registered');
    var record = {
      ID: newUuid(),
      ChannelId: channelId,
      UserId: userId,
      Message: text,
      Timestamp: nowIso(),
      IsDeleted: false
    };
    assignTenant(record, table, context, campaignId || (ctxInfo.profile && ctxInfo.profile.defaultCampaignId));
    return safeCreate('chatMessages', context, record);
  };

  WorkflowService.getCollaborationDigest = function (userId, options) {
    options = options || {};
    var workspace = WorkflowService.getWorkspace(userId, { campaignId: options.campaignId, maxMessages: options.maxMessages || 25 });
    return workspace.collaboration;
  };

  global.CallCenterWorkflowService = WorkflowService;

})(typeof globalThis !== 'undefined' ? globalThis : this);
