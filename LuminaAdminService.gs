/**
 * LuminaAdminService.gs
 * -----------------------------------------------------------------------------
 * System-owner automation, telemetry, and policy enforcement for Lumina Sheets.
 * Provides bootstrap seeding, RBAC/ABAC helpers, background job orchestration,
 * and global navigation registry for the lumina-admin@system superuser.
 */
(function bootstrapLuminaAdmin(global) {
  if (!global) return;
  if (global.LuminaAdmin && typeof global.LuminaAdmin === 'object') {
    return;
  }

  var Utilities = global.Utilities;
  var ScriptApp = global.ScriptApp;
  var MailApp = global.MailApp;
  var LockService = global.LockService;

  function getRepository() {
    var repo = global.IdentityRepository;
    if (!repo || typeof repo.list !== 'function') {
      throw new Error('IdentityRepository not initialized');
    }
    return repo;
  }

  function getAuditService() {
    var audit = global.AuditService;
    if (!audit || typeof audit.log !== 'function') {
      throw new Error('AuditService not initialized');
    }
    return audit;
  }

  var SYSTEM_OWNER = {
    USER_ID: 'LUMINA-SYSTEM-OWNER',
    USERNAME: 'lumina-admin',
    EMAIL: 'lumina-admin@system',
    ROLE: 'SystemOwner'
  };

  var DEFAULT_FEATURE_FLAGS = [
    { Key: 'autonomous_mode', Value: 'true', Env: 'prod', Description: 'Enable Lumina Admin autonomous agent', UpdatedAt: new Date().toISOString(), Flag: 'autonomous_mode', Notes: 'SystemOwner bootstrap default' },
    { Key: 'notify_email', Value: 'true', Env: 'prod', Description: 'Send email notifications for Lumina Admin messages', UpdatedAt: new Date().toISOString(), Flag: 'notify_email', Notes: 'SystemOwner bootstrap default' },
    { Key: 'notify_inapp', Value: 'true', Env: 'prod', Description: 'Enable in-app SystemMessages banner', UpdatedAt: new Date().toISOString(), Flag: 'notify_inapp', Notes: 'SystemOwner bootstrap default' },
    { Key: 'risky_auto_actions', Value: 'false', Env: 'prod', Description: 'Allow autonomous corrective actions without approval', UpdatedAt: new Date().toISOString(), Flag: 'risky_auto_actions', Notes: 'Disabled by default' }
  ];

  var DEFAULT_JOBS = [
    { JobId: 'integrity_scan', Name: 'integrity_scan', Schedule: 'every5m', Enabled: 'true' },
    { JobId: 'rollup_attendance', Name: 'rollup_attendance', Schedule: 'hourly', Enabled: 'true' },
    { JobId: 'qa_monitor', Name: 'qa_monitor', Schedule: 'hourly', Enabled: 'true' },
    { JobId: 'policy_enforcer', Name: 'policy_enforcer', Schedule: 'hourly', Enabled: 'true' },
    { JobId: 'benefits_eligibility', Name: 'benefits_eligibility', Schedule: 'daily-0200', Enabled: 'true' },
    { JobId: 'integration_health', Name: 'integration_health', Schedule: 'hourly', Enabled: 'true' },
    { JobId: 'featureflag_audit', Name: 'featureflag_audit', Schedule: 'daily-0205', Enabled: 'true' }
  ];

  var CATEGORY_REGISTRY = {
    Overview: ['Admin Dashboard', 'System Health', 'Audit Trail'],
    People: ['User Directory', 'Roles & Policies', 'Equipment Ledger', 'Lifecycle Actions'],
    Scheduling: ['Shifts', 'Adherence', 'Attendance'],
    Quality: ['QA Audits', 'Coaching', 'Scorecards'],
    Operations: ['Tickets', 'Tasks', 'Knowledge Base'],
    Finance: ['Payroll Hooks', 'Deductions', 'Benefits Eligibility'],
    Integrations: ['Google', 'Mail', 'Cloudinary', 'iCoPAY'],
    System: ['Config', 'Feature Flags', 'Backups', 'Jobs', 'Logs']
  };

  var BACKGROUND_SEVERITIES = ['INFO', 'NOTICE', 'WARNING', 'CRITICAL', 'KUDOS'];

  var JOB_HANDLERS = {
    integrity_scan: runIntegrityScan,
    policy_enforcer: runPolicyEnforcer,
    rollup_attendance: runAttendanceRollup,
    qa_monitor: runQaMonitor,
    benefits_eligibility: runBenefitsEligibility,
    integration_health: runIntegrationHealth,
    featureflag_audit: runFeatureFlagAudit
  };

  function clone(value) {
    if (value === null || typeof value === 'undefined') return value;
    try {
      return JSON.parse(JSON.stringify(value));
    } catch (err) {
      return value;
    }
  }

  function nowIso() {
    return new Date().toISOString();
  }

  function ensureSystemOwnerUser() {
    var repo = getRepository();
    var users = repo.list('Users');
    var existing = users.find(function(row) { return row.UserId === SYSTEM_OWNER.USER_ID || String(row.Email || '').toLowerCase() === SYSTEM_OWNER.EMAIL; });
    var flagsJson = JSON.stringify({ BackgroundAgent: true, PII_RedactionDefault: true });
    var payload = {
      UserId: SYSTEM_OWNER.USER_ID,
      Email: SYSTEM_OWNER.EMAIL,
      Username: SYSTEM_OWNER.USERNAME,
      Status: 'Active',
      CreatedAt: existing ? (existing.CreatedAt || nowIso()) : nowIso(),
      UpdatedAt: nowIso(),
      Role: SYSTEM_OWNER.ROLE,
      CampaignId: '*',
      FlagsJson: flagsJson,
      Watchlist: 'false'
    };
    repo.upsert('Users', 'UserId', payload);
    repo.upsert('UserCampaigns', 'AssignmentId', {
      AssignmentId: SYSTEM_OWNER.USER_ID + '::GLOBAL',
      UserId: SYSTEM_OWNER.USER_ID,
      CampaignId: '*',
      Role: SYSTEM_OWNER.ROLE,
      IsPrimary: 'true',
      AddedBy: SYSTEM_OWNER.USER_ID,
      AddedAt: nowIso(),
      Watchlist: 'false'
    });
    return payload;
  }

  function ensureRoles() {
    var repo = getRepository();
    var roles = repo.list('Roles');
    var roleIndex = {};
    roles.forEach(function(role) {
      roleIndex[String(role.Name || role.Role || '').toLowerCase()] = role;
    });
    var desired = [
      { Name: 'SystemOwner', Permissions: ['*'], DefaultForCampaignManager: false, IsGlobal: 'Y' },
      { Name: 'CampaignManager', Permissions: ['manage_campaign', 'manage_people', 'view_reports'], DefaultForCampaignManager: true, IsGlobal: 'N' },
      { Name: 'Agent', Permissions: ['view_tasks', 'submit_updates'], DefaultForCampaignManager: false, IsGlobal: 'N' },
      { Name: 'GuestClient', Permissions: ['view_reports_readonly'], DefaultForCampaignManager: false, IsGlobal: 'N' }
    ];

    desired.forEach(function(role) {
      var key = role.Name.toLowerCase();
      var existing = roleIndex[key];
      var payload = {
        RoleId: role.Name.toUpperCase(),
        Role: role.Name,
        Name: role.Name,
        Description: role.Name + ' role',
        PermissionsJson: JSON.stringify(role.Permissions),
        DefaultForCampaignManager: String(role.DefaultForCampaignManager),
        IsGlobal: role.IsGlobal || 'N'
      };
      repo.upsert('Roles', 'RoleId', payload);
      if (role.Name === 'SystemOwner') {
        repo.upsert('RolePermissions', 'PermissionId', {
          PermissionId: 'SystemOwner::*',
          Role: 'SystemOwner',
          Capability: '*',
          Scope: 'Global',
          Allowed: 'Y'
        });
      }
    });
  }

  function ensureFeatureFlags() {
    var repo = getRepository();
    var existing = repo.list('FeatureFlags');
    var index = {};
    existing.forEach(function(flag) {
      index[String(flag.Key || flag.Flag || '').toLowerCase()] = flag;
    });
    DEFAULT_FEATURE_FLAGS.forEach(function(flag) {
      var key = String(flag.Key).toLowerCase();
      if (index[key]) {
        return;
      }
      repo.upsert('FeatureFlags', 'Key', flag);
    });
  }

  function ensureJobs() {
    var repo = getRepository();
    var existing = repo.list('Jobs');
    var index = {};
    existing.forEach(function(job) {
      index[String(job.JobId || job.Name || '').toLowerCase()] = job;
    });
    DEFAULT_JOBS.forEach(function(job) {
      var key = String(job.JobId).toLowerCase();
      if (index[key]) {
        return;
      }
      var payload = Object.assign({
        LastRunAt: '',
        LastStatus: 'NEVER',
        ConfigJson: '{}',
        RunHash: ''
      }, job);
      repo.upsert('Jobs', 'JobId', payload);
    });
  }

  function ensureSeeded() {
    var lock = LockService ? LockService.getScriptLock() : null;
    if (lock) {
      lock.waitLock(30000);
    }
    try {
      ensureSystemOwnerUser();
      ensureRoles();
      ensureFeatureFlags();
      ensureJobs();
    } finally {
      if (lock) {
        lock.releaseLock();
      }
    }
  }

  function categoryRegistry() {
    return clone(CATEGORY_REGISTRY);
  }

  function isSystemOwnerUser(user) {
    if (!user) return false;
    if (String(user.UserId || user.userId || '').toUpperCase() === SYSTEM_OWNER.USER_ID) return true;
    if (String(user.Email || user.email || '').toLowerCase() === SYSTEM_OWNER.EMAIL) return true;
    if (String(user.Role || user.role || '') === 'SystemOwner') return true;
    return false;
  }

  function buildSystemOwnerNavigation() {
    var nav = { categories: [], uncategorizedPages: [] };
    Object.keys(CATEGORY_REGISTRY).forEach(function(category) {
      var pages = CATEGORY_REGISTRY[category] || [];
      nav.categories.push({
        CategoryName: category,
        CategoryIcon: inferCategoryIcon(category),
        pages: pages.map(function(page) {
          return {
            PageTitle: page,
            PageKey: slugify(page),
            PageIcon: inferPageIcon(page)
          };
        })
      });
    });
    return nav;
  }

  function inferCategoryIcon(category) {
    var lower = String(category || '').toLowerCase();
    if (lower === 'overview') return 'fas fa-chart-pie';
    if (lower === 'people') return 'fas fa-users';
    if (lower === 'scheduling') return 'fas fa-calendar-check';
    if (lower === 'quality') return 'fas fa-star';
    if (lower === 'operations') return 'fas fa-tasks';
    if (lower === 'finance') return 'fas fa-coins';
    if (lower === 'integrations') return 'fas fa-plug';
    if (lower === 'system') return 'fas fa-cog';
    return 'fas fa-folder-open';
  }

  function inferPageIcon(page) {
    var lower = String(page || '').toLowerCase();
    if (/dashboard/.test(lower)) return 'fas fa-chart-line';
    if (/health/.test(lower)) return 'fas fa-heartbeat';
    if (/audit/.test(lower)) return 'fas fa-clipboard-check';
    if (/user/.test(lower)) return 'fas fa-user-cog';
    if (/roles/.test(lower)) return 'fas fa-user-shield';
    if (/equipment/.test(lower)) return 'fas fa-laptop';
    if (/lifecycle/.test(lower)) return 'fas fa-random';
    if (/shift|schedule/.test(lower)) return 'fas fa-clock';
    if (/qa|quality/.test(lower)) return 'fas fa-check-circle';
    if (/coaching/.test(lower)) return 'fas fa-chalkboard-teacher';
    if (/ticket|task/.test(lower)) return 'fas fa-inbox';
    if (/knowledge/.test(lower)) return 'fas fa-book';
    if (/payroll|benefit/.test(lower)) return 'fas fa-money-check';
    if (/integration/.test(lower)) return 'fas fa-project-diagram';
    if (/config|flag/.test(lower)) return 'fas fa-sliders-h';
    if (/backup/.test(lower)) return 'fas fa-database';
    if (/job/.test(lower)) return 'fas fa-robot';
    if (/log/.test(lower)) return 'fas fa-scroll';
    return 'fas fa-file-alt';
  }

  function slugify(text) {
    return String(text || '')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '');
  }

  function publishMessage(severity, title, body, target) {
    if (BACKGROUND_SEVERITIES.indexOf(severity) === -1) {
      throw new Error('Invalid severity: ' + severity);
    }
    var repo = getRepository();
    var metadata = (target && target.metadata) || {};
    var sanitizedBody = redactSensitiveText(body, target && target.allowPII);
    var payload = {
      MessageId: Utilities.getUuid(),
      Severity: severity,
      Title: title,
      Body: sanitizedBody,
      TargetRole: target && target.role ? target.role : '',
      TargetCampaignId: target && target.campaignId ? target.campaignId : '',
      Status: 'Open',
      CreatedAt: nowIso(),
      ResolvedAt: '',
      CreatedBy: target && target.createdBy ? target.createdBy : SYSTEM_OWNER.USER_ID,
      MetadataJson: JSON.stringify(metadata)
    };
    repo.append('SystemMessages', payload);
    if (shouldSendEmail()) {
      trySendEmailNotification(payload);
    }
    return payload;
  }

  function shouldSendEmail() {
    var repo = getRepository();
    var flags = repo.list('FeatureFlags');
    return flags.some(function(flag) {
      var key = String(flag.Key || flag.Flag || '').toLowerCase();
      var value = String(flag.Value || flag.value || '').toLowerCase();
      return key === 'notify_email' && (value === 'true' || value === '1');
    });
  }

  function trySendEmailNotification(message) {
    if (!MailApp) {
      return;
    }
    var subject = '[Lumina Admin] ' + message.Severity + ' • ' + message.Title;
    var htmlBody = '<p>' + escapeHtml(message.Body) + '</p>';
    MailApp.sendEmail({ to: SYSTEM_OWNER.EMAIL, subject: subject, htmlBody: htmlBody });
  }

  function escapeHtml(text) {
    return String(text || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function redactSensitiveText(text, allowPII) {
    if (allowPII) {
      return text;
    }
    return String(text || '').replace(/([A-Z][a-z]+)\s+([A-Z])[a-z]*/g, '$1 $2.');
  }

  function logAudit(action, resourceType, resourceId, before, after, mode) {
    getAuditService().log({
      ActorUserId: SYSTEM_OWNER.USER_ID,
      ActorRole: SYSTEM_OWNER.ROLE,
      CampaignId: resourceType === 'Campaign' ? resourceId : '',
      Target: resourceType + '::' + resourceId,
      Action: action,
      Before: before,
      After: after,
      Timestamp: nowIso(),
      Mode: mode || 'Autonomous'
    });
  }

  function runJob(jobName, context) {
    ensureSeeded();
    var repo = getRepository();
    var job = repo.list('Jobs').find(function(row) {
      return String(row.JobId || row.Name || '').toLowerCase() === String(jobName).toLowerCase();
    });
    if (!job) {
      throw new Error('Unknown job: ' + jobName);
    }
    if (String(job.Enabled).toLowerCase() === 'false') {
      return { skipped: true, reason: 'Job disabled' };
    }
    var handler = JOB_HANDLERS[job.Name];
    if (!handler) {
      throw new Error('No handler registered for job ' + job.Name);
    }
    var result = handler(context || {}) || {};
    var hashSource = result.metrics || result;
    var newHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(hashSource || {})));
    if (job.RunHash && job.RunHash === newHash) {
      return { skipped: true, reason: 'No changes detected', hash: newHash };
    }
    var messages = Array.isArray(result.messages) ? result.messages : [];
    messages.forEach(function(message) {
      publishMessage(message.severity || 'INFO', message.title || 'Lumina Admin Notice', message.body || '', message.target || {});
    });
    repo.upsert('Jobs', 'JobId', Object.assign({}, job, {
      LastRunAt: nowIso(),
      LastStatus: (result && result.status) || 'SUCCESS',
      RunHash: newHash
    }));
    return Object.assign({}, result, {
      hash: newHash,
      messagesPublished: messages.length
    });
  }

  function registerTriggers() {
    if (!ScriptApp) {
      return { installed: false, reason: 'ScriptApp unavailable in this environment' };
    }
    var existing = ScriptApp.getProjectTriggers();
    var required = {
      every5m: { type: 'timeBased', minutes: 5, handler: 'LuminaAdmin_runIntegrity' },
      hourly: { type: 'timeBased', minutes: 60, handler: 'LuminaAdmin_runHourly' },
      daily: { type: 'timeBased', hours: 24, handler: 'LuminaAdmin_runDaily' }
    };
    Object.keys(required).forEach(function(key) {
      var config = required[key];
      var present = existing.some(function(trigger) {
        return trigger.getHandlerFunction && trigger.getHandlerFunction() === config.handler;
      });
      if (present) {
        return;
      }
      var builder = ScriptApp.newTrigger(config.handler).timeBased();
      if (config.minutes === 5) builder.everyMinutes(5);
      else if (config.minutes === 60) builder.everyHours(1);
      else builder.atHour(2).everyDays(1);
      builder.create();
    });
    return { installed: true };
  }

  function runIntegrityScan() {
    var repo = getRepository();
    var orphanedEquipment = repo.list('Equipment').filter(function(eq) {
      return !eq.UserId;
    });
    var messages = [];
    if (orphanedEquipment.length) {
      messages.push({
        severity: 'WARNING',
        title: 'Equipment without custodian',
        body: orphanedEquipment.length + ' assets require review.',
        target: { role: 'SystemOwner' }
      });
    }
    return {
      status: 'SUCCESS',
      metrics: { orphanedEquipment: orphanedEquipment.length },
      messages: messages
    };
  }

  function runPolicyEnforcer() {
    var repo = getRepository();
    var shifts = repo.list('Shifts');
    var uncovered = shifts.filter(function(shift) {
      return !shift.UserId || String(shift.Status || '').toLowerCase() === 'uncovered';
    });
    var messages = [];
    if (uncovered.length) {
      messages.push({
        severity: 'CRITICAL',
        title: 'Uncovered shifts detected',
        body: uncovered.length + ' shifts are missing coverage.',
        target: { role: 'CampaignManager' }
      });
    }
    return {
      status: 'SUCCESS',
      metrics: { uncoveredShifts: uncovered.length },
      messages: messages
    };
  }

  function runAttendanceRollup() {
    var repo = getRepository();
    var attendance = repo.list('Attendance');
    var overtime = attendance.filter(function(record) {
      return Number(record.Minutes || 0) > 40 * 60;
    });
    var messages = [];
    if (overtime.length) {
      messages.push({
        severity: 'NOTICE',
        title: 'Weekly overtime review',
        body: overtime.length + ' agents exceeded 40 hours.',
        target: { role: 'SystemOwner' }
      });
    }
    return {
      status: 'SUCCESS',
      metrics: { overtimeRecords: overtime.length },
      messages: messages
    };
  }

  function runQaMonitor() {
    var repo = getRepository();
    var audits = repo.list('QAAudits');
    var recent = audits.slice(-20);
    var belowThreshold = recent.filter(function(audit) {
      return Number(audit.Score || 0) < 80;
    });
    var messages = [];
    if (belowThreshold.length) {
      messages.push({
        severity: 'NOTICE',
        title: 'QA score dip detected',
        body: belowThreshold.length + ' recent audits fell below 80%.',
        target: { role: 'Quality' }
      });
    }
    return {
      status: 'SUCCESS',
      metrics: { lowScores: belowThreshold.length },
      messages: messages
    };
  }

  function runBenefitsEligibility() {
    var repo = getRepository();
    var benefits = repo.list('Benefits');
    var needingReview = benefits.filter(function(row) {
      return String(row.Eligible || '').toLowerCase() === 'false';
    });
    var messages = [];
    if (needingReview.length) {
      messages.push({
        severity: 'NOTICE',
        title: 'Benefits eligibility review',
        body: needingReview.length + ' team members need eligibility review.',
        target: { role: 'CampaignManager' }
      });
    }
    return {
      status: 'SUCCESS',
      metrics: { pending: needingReview.length },
      messages: messages
    };
  }

  function runIntegrationHealth() {
    var repo = getRepository();
    var jobs = repo.list('Jobs');
    var failures = jobs.filter(function(job) {
      return String(job.LastStatus || '').toUpperCase() === 'FAILED';
    });
    var messages = [];
    if (failures.length) {
      messages.push({
        severity: 'WARNING',
        title: 'Integration job failures',
        body: failures.length + ' background jobs reported failures.',
        target: { role: 'SystemOwner' }
      });
    }
    return {
      status: 'SUCCESS',
      metrics: { failingJobs: failures.length },
      messages: messages
    };
  }

  function runFeatureFlagAudit() {
    var repo = getRepository();
    var flags = repo.list('FeatureFlags');
    var risky = flags.filter(function(flag) {
      var key = String(flag.Key || flag.Flag || '').toLowerCase();
      var value = String(flag.Value || '').toLowerCase();
      return key === 'risky_auto_actions' && value === 'true';
    });
    var messages = [];
    if (risky.length) {
      messages.push({
        severity: 'WARNING',
        title: 'Risky automation enabled',
        body: 'Risky autonomous actions are enabled. Confirm this is intentional.',
        target: { role: 'SystemOwner' }
      });
    }
    return {
      status: 'SUCCESS',
      metrics: { riskyFlags: risky.length },
      messages: messages
    };
  }

  function lifecycleAction(action, payload, actor) {
    ensureSeeded();
    var repo = getRepository();
    var before;
    var after;
    if (action === 'hire') {
      var userId = payload.UserId || Utilities.getUuid();
      before = null;
      after = Object.assign({
        UserId: userId,
        Status: 'Active',
        CreatedAt: nowIso(),
        UpdatedAt: nowIso()
      }, payload);
      repo.upsert('Users', 'UserId', after);
      publishMessage('INFO', 'New hire onboarded', redactSensitiveText('User ' + (payload.DisplayName || payload.Username || 'New Hire') + ' created for campaign ' + (payload.CampaignId || ''), false), { role: 'SystemOwner' });
    } else if (action === 'transfer') {
      before = repo.list('Users').find(function(row) { return row.UserId === payload.UserId; });
      after = Object.assign({}, before, { CampaignId: payload.TargetCampaignId, UpdatedAt: nowIso() });
      repo.upsert('Users', 'UserId', after);
      publishMessage('NOTICE', 'Transfer completed', 'User transferred to campaign ' + payload.TargetCampaignId + '.', { role: 'SystemOwner' });
    } else if (action === 'promote') {
      before = repo.list('Users').find(function(row) { return row.UserId === payload.UserId; });
      after = Object.assign({}, before, { Role: payload.Role, UpdatedAt: nowIso() });
      repo.upsert('Users', 'UserId', after);
      publishMessage('KUDOS', 'Promotion recorded', 'User promoted to ' + payload.Role + '.', { role: 'SystemOwner' });
    } else if (action === 'terminate') {
      before = repo.list('Users').find(function(row) { return row.UserId === payload.UserId; });
      after = Object.assign({}, before, { Status: 'Inactive', UpdatedAt: nowIso() });
      repo.upsert('Users', 'UserId', after);
      publishMessage('CRITICAL', 'Access revoked', 'User access revoked for security compliance.', { role: 'SystemOwner' });
    } else {
      throw new Error('Unsupported lifecycle action: ' + action);
    }
    logAudit(action.toUpperCase(), 'User', payload.UserId || (after && after.UserId) || '', before, after, actor && actor.mode);
    return { success: true };
  }

  function getMessages(filters) {
    var rows = getRepository().list('SystemMessages');
    if (!filters) return rows;
    return rows.filter(function(row) {
      if (filters.status && String(row.Status).toLowerCase() !== String(filters.status).toLowerCase()) return false;
      if (filters.severity && String(row.Severity).toUpperCase() !== String(filters.severity).toUpperCase()) return false;
      return true;
    });
  }

  function updateMessageStatus(messageId, status) {
    var repo = getRepository();
    var rows = repo.list('SystemMessages');
    var existing = rows.find(function(row) { return row.MessageId === messageId; });
    if (!existing) {
      throw new Error('Message not found: ' + messageId);
    }
    var updated = Object.assign({}, existing, {
      Status: status,
      ResolvedAt: status === 'Resolved' ? nowIso() : existing.ResolvedAt
    });
    repo.upsert('SystemMessages', 'MessageId', updated);
    logAudit('MESSAGE_STATUS', 'SystemMessage', messageId, existing, updated, 'Interactive');
    return updated;
  }

  function ensureTotpSecret(userId) {
    var repo = getRepository();
    var users = repo.list('Users');
    var user = users.find(function(row) { return row.UserId === userId; });
    if (!user) {
      throw new Error('User not found for TOTP enrollment');
    }
    if (user.TOTPSecretHash) {
      return user.TOTPSecretHash;
    }
    var secret = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid()));
    repo.upsert('Users', 'UserId', Object.assign({}, user, {
      TOTPSecretHash: secret,
      TOTPEnabled: 'true',
      UpdatedAt: nowIso()
    }));
    logAudit('TOTP_ENROLL', 'User', userId, user, Object.assign({}, user, { TOTPSecretHash: secret, TOTPEnabled: 'true' }), 'Interactive');
    return secret;
  }

  function isAutonomousModeEnabled() {
    var repo = getRepository();
    var flags = repo.list('FeatureFlags');
    return flags.some(function(flag) {
      var key = String(flag.Key || flag.Flag || '').toLowerCase();
      var value = String(flag.Value || '').toLowerCase();
      return key === 'autonomous_mode' && (value === 'true' || value === '1');
    });
  }

  function ensureAutonomousRun(jobName) {
    if (!isAutonomousModeEnabled()) {
      return { skipped: true, reason: 'Autonomous mode disabled' };
    }
    return runJob(jobName || 'integrity_scan', { mode: 'autonomous' });
  }

  global.LuminaAdmin = {
    ensureSeeded: ensureSeeded,
    ensureTriggers: registerTriggers,
    isSystemOwner: isSystemOwnerUser,
    getCategoryRegistry: categoryRegistry,
    buildSystemOwnerNavigation: buildSystemOwnerNavigation,
    publishMessage: publishMessage,
    runJob: runJob,
    ensureAutonomousRun: ensureAutonomousRun,
    lifecycleAction: lifecycleAction,
    getMessages: getMessages,
    updateMessageStatus: updateMessageStatus,
    ensureTotpSecret: ensureTotpSecret
  };

  if (typeof global.LuminaAdmin_runIntegrity !== 'function') {
    global.LuminaAdmin_runIntegrity = function() {
      return ensureAutonomousRun('integrity_scan');
    };
  }

  if (typeof global.LuminaAdmin_runHourly !== 'function') {
    global.LuminaAdmin_runHourly = function() {
      ensureAutonomousRun('rollup_attendance');
      ensureAutonomousRun('qa_monitor');
      ensureAutonomousRun('policy_enforcer');
      ensureAutonomousRun('integration_health');
    };
  }

  if (typeof global.LuminaAdmin_runDaily !== 'function') {
    global.LuminaAdmin_runDaily = function() {
      ensureAutonomousRun('benefits_eligibility');
      ensureAutonomousRun('featureflag_audit');
    };
  }

})(GLOBAL_SCOPE);
