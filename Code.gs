/** Enhanced Multi-Campaign Google Apps Script - Code.gs
 * Simplified Token-Based Authentication System
 * 
 * Features:
 * - Token-based authentication
 * - Campaign-specific routing (e.g., CreditSuite.QAForm, IBTR.QACollabList)
 * - Enhanced authentication and access control
 * - Clean URL structure
 * - Secure session management via tokens
 */

// ───────────────────────────────────────────────────────────────────────────────
// GLOBAL CONSTANTS AND CONFIGURATION
// ───────────────────────────────────────────────────────────────────────────────

const SCRIPT_URL = 'https://script.google.com/a/macros/vlbpo.com/s/AKfycbxeQ0AnupBHM71M6co3LVc5NPrxTblRXLd6AuTOpxMs2rMehF9dBSkGykIcLGHROywQ/exec';
const FAVICON_URL = 'https://res.cloudinary.com/dr8qd3xfc/image/upload/v1754763514/vlbpo/lumina/3_dgitcx.png';

/** Toggle for debug traces */
const ACCESS_DEBUG = true;

/** Canonical page access definitions */
const ACCESS = {
  ADMIN_ONLY_PAGES: new Set(['admin.users', 'admin.roles', 'admin.campaigns']),
  PUBLIC_PAGES: new Set([
    'login', 'setpassword', 'resetpassword', 'forgotpassword', 'forgot-password',
    'resendverification', 'resend-verification', 'emailconfirmed', 'email-confirmed'
  ]),
  DEFAULT_PAGE: 'dashboard',
  PRIVS: { SYSTEM_ADMIN: 'SYSTEM_ADMIN', MANAGE_USERS: 'MANAGE_USERS', MANAGE_PAGES: 'MANAGE_PAGES' }
};

// ───────────────────────────────────────────────────────────────────────────────
// CAMPAIGN-SCOPED ROUTING SYSTEM
// ───────────────────────────────────────────────────────────────────────────────

function __case_slug__(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/[^\w\s-]/g, '')
    .trim()
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-');
}

function __case_norm__(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

/** Campaign definitions with their specific templates */
const CASE_DEFS = [
  // Credit Suite
  {
    case: 'CreditSuite',
    aliases: ['CS', 'Credit Suite', 'creditsuite', 'credit-suite'],
    idHint: 'credit-suite',
    pages: {
      QAForm: 'CreditSuiteQAForm',
      QADashboard: 'CreditSuiteQADashboard',
      QAList: 'CreditSuiteQAList',
      QAView: 'CreditSuiteQAView',
      AttendanceReports: 'CreditSuiteAttendanceReports',
      CallReports: 'CreditSuiteCallReports',
      Dashboard: 'CreditSuiteDashboard'
    }
  },

  // HiyaCar
  {
    case: 'HiyaCar',
    aliases: ['HYC', 'hiya car', 'hiya-car', 'hiyacar'],
    idHint: 'hiya-car',
    pages: {
      QAForm: 'HiyaCarQAForm',
      QADashboard: 'HiyaCarQADashboard',
      QAList: 'HiyaCarQAList',
      QAView: 'HiyaCarQAView',
      AttendanceReports: 'HiyaCarAttendanceReports',
      CallReports: 'HiyaCarCallReports',
      Dashboard: 'HiyaCarDashboard'
    }
  },

  // Benefits Resource Center (iBTR)
  {
    case: 'IBTR',
    aliases: ['Benefits Resource Center (iBTR)', 'Benefits Resource Center', 'iBTR', 'IBTR', 'benefits-resource-center-ibtr', 'BRC'],
    idHint: 'ibtr',
    pages: {
      QAForm: 'IBTRQAForm',
      QADashboard: 'IBTRQADashboard',
      QAList: 'IBTRQAList',
      QAView: 'IBTRQAView',
      QACollabList: 'IBTRQACollabList', // special page
      AttendanceReports: 'IBTRAttendanceReports',
      CallReports: 'IBTRCallReports',
      Dashboard: 'IBTRDashboard'
    }
  },

  // Independence Insurance Agency
  {
    case: 'IndependenceInsuranceAgency',
    aliases: ['Independence Insurance Agency', 'Independence', 'IIA', 'independence-insurance-agency'],
    idHint: 'independence-insurance-agency',
    pages: {
      QAForm: 'IndependenceQAForm',
      QADashboard: 'IndependenceQADashboard',
      QAList: 'IndependenceQAList',
      QAView: 'IndependenceQAView',
      AttendanceReports: 'IndependenceAttendanceReports',
      CallReports: 'IndependenceCallReports',
      Dashboard: 'IndependenceDashboard',
      CoachingForm: 'IndependenceCoachingForm'
    }
  },

  // JSC
  {
    case: 'JSC',
    aliases: ['JSC'],
    idHint: 'jsc',
    pages: {
      QAForm: 'JSCQAForm',
      QADashboard: 'JSCQADashboard',
      QAList: 'JSCQAList',
      QAView: 'JSCQAView',
      AttendanceReports: 'JSCAttendanceReports',
      CallReports: 'JSCCallReports',
      Dashboard: 'JSCDashboard'
    }
  },

  // Kids in the Game
  {
    case: 'KidsInTheGame',
    aliases: ['Kids in the Game', 'KITG', 'kids-in-the-game'],
    idHint: 'kids-in-the-game',
    pages: {
      QAForm: 'KidsInTheGameQAForm',
      QADashboard: 'KidsInTheGameQADashboard',
      QAList: 'KidsInTheGameQAList',
      QAView: 'KidsInTheGameQAView',
      AttendanceReports: 'KidsInTheGameAttendanceReports',
      CallReports: 'KidsInTheGameCallReports',
      Dashboard: 'KidsInTheGameDashboard'
    }
  },

  // Kofi Group
  {
    case: 'KofiGroup',
    aliases: ['Kofi Group', 'KOFI', 'kofi-group'],
    idHint: 'kofi-group',
    pages: {
      QAForm: 'KofiGroupQAForm',
      QADashboard: 'KofiGroupQADashboard',
      QAList: 'KofiGroupQAList',
      QAView: 'KofiGroupQAView',
      AttendanceReports: 'KofiGroupAttendanceReports',
      CallReports: 'KofiGroupCallReports',
      Dashboard: 'KofiGroupDashboard'
    }
  },

  // PAW LAW FIRM
  {
    case: 'PAWLawFirm',
    aliases: ['PAW LAW FIRM', 'PAW', 'paw-law-firm'],
    idHint: 'paw-law-firm',
    pages: {
      QAForm: 'PAWLawFirmQAForm',
      QADashboard: 'PAWLawFirmQADashboard',
      QAList: 'PAWLawFirmQAList',
      QAView: 'PAWLawFirmQAView',
      AttendanceReports: 'PAWLawFirmAttendanceReports',
      CallReports: 'PAWLawFirmCallReports',
      Dashboard: 'PAWLawFirmDashboard'
    }
  },

  // Pro House Photos
  {
    case: 'ProHousePhotos',
    aliases: ['Pro House Photos', 'PHP', 'pro-house-photos'],
    idHint: 'pro-house-photos',
    pages: {
      QAForm: 'ProHousePhotosQAForm',
      QADashboard: 'ProHousePhotosQADashboard',
      QAList: 'ProHousePhotosQAList',
      QAView: 'ProHousePhotosQAView',
      AttendanceReports: 'ProHousePhotosAttendanceReports',
      CallReports: 'ProHousePhotosCallReports',
      Dashboard: 'ProHousePhotosDashboard'
    }
  },

  // Independence Agency & Credit Suite
  {
    case: 'IndependenceAgencyCreditSuite',
    aliases: ['Independence Agency & Credit Suite', 'IACS', 'independence-agency-credit-suite'],
    idHint: 'independence-agency-credit-suite',
    pages: {
      QAForm: 'IACSQAForm',
      QADashboard: 'IACSQADashboard',
      QAList: 'IACSQAList',
      QAView: 'IACSQAView',
      AttendanceReports: 'IACSAttendanceReports',
      CallReports: 'IACSCallReports',
      Dashboard: 'IACSQADashboard'
    }
  },

  // Proozy
  {
    case: 'Proozy',
    aliases: ['Proozy', 'PRZ', 'proozy'],
    idHint: 'proozy',
    pages: {
      QAForm: 'ProozyQAForm',
      QADashboard: 'ProozyQADashboard',
      QAList: 'ProozyQAList',
      QAView: 'ProozyQAView',
      AttendanceReports: 'ProozyAttendanceReports',
      CallReports: 'ProozyCallReports',
      Dashboard: 'ProozyDashboard'
    }
  },

  // The Grounding (TGC)
  {
    case: 'TGC',
    aliases: ['The Grounding', 'TG', 'TGC', 'the-grounding'],
    idHint: 'the-grounding',
    pages: {
      QAForm: 'TheGroundingQAForm',
      QADashboard: 'TheGroundingQADashboard',
      QAList: 'TheGroundingQAList',
      QAView: 'TheGroundingQAView',
      AttendanceReports: 'TGCAttendanceReports',
      ChatReport: 'TGCChatReport', // Special: chat instead of calls
      Dashboard: 'TheGroundingDashboard'
    }
  },

  // CO
  {
    case: 'CO',
    aliases: ['CO', 'CO ()', 'co'],
    idHint: 'co',
    pages: {
      QAForm: 'COQAForm',
      QADashboard: 'COQADashboard',
      QAList: 'COQAList',
      QAView: 'COQAView',
      AttendanceReports: 'COAttendanceReports',
      CallReports: 'COCallReports',
      Dashboard: 'CODashboard'
    }
  }
];

/** Generic fallback templates */
const GENERIC_FALLBACKS = {
  QAForm: ['QualityForm'],
  QADashboard: ['QADashboard', 'UnifiedQADashboard'],
  QAList: ['QAList'],
  QAView: ['QualityView'],
  AttendanceReports: ['AttendanceReports'],
  CallReports: ['CallReports'],
  ChatReport: ['ChatReports', 'Chat'],
  Dashboard: ['Dashboard'],
  CoachingForm: ['CoachingForm'],
  TaskForm: ['TaskForm'],
  TaskList: ['TaskList'],
  TaskBoard: ['TaskBoard']
};

// ───────────────────────────────────────────────────────────────────────────────
// UTILITY FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

function _truthy(v) {
  return v === true || String(v).toUpperCase() === 'TRUE';
}

function _normalizePageKey(k) {
  return (k || '').toString().trim().toLowerCase();
}

function _normalizeId(x) {
  return (x == null ? '' : String(x)).trim();
}

function _now() {
  return new Date();
}

function _safeDate(d) {
  try {
    return new Date(d);
  } catch (e) {
    return null;
  }
}

function _isFuture(d) {
  try {
    return d && d.getTime && d.getTime() > Date.now();
  } catch (_) {
    return false;
  }
}

function _csvToSet(csv) {
  const set = new Set();
  (String(csv || '').split(',') || []).forEach(s => {
    const k = _normalizePageKey(s);
    if (k) set.add(k);
  });
  return set;
}

function _stringifyForTemplate_(obj) {
  try {
    return JSON.stringify(obj || {}).replace(/<\/script>/g, '<\\/script>');
  } catch (_) {
    return '{}';
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// CACHING HELPERS
// ───────────────────────────────────────────────────────────────────────────────

function _cacheGet(key) {
  try {
    const c = CacheService.getScriptCache().get(key);
    if (c) return JSON.parse(c);
  } catch (_) { }
  try {
    const p = PropertiesService.getScriptProperties().getProperty(key);
    if (p) return JSON.parse(p);
  } catch (_) { }
  return null;
}

function _cachePut(key, obj, ttlSec) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(obj), Math.min(21600, Math.max(5, ttlSec || 60)));
  } catch (_) { }
  try {
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(obj));
  } catch (_) { }
}


// ───────────────────────────────────────────────────────────────────────────────
// SIMPLIFIED AUTHENTICATION FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Get current user from session using Google Apps Script built-in authentication
 */
function getCurrentUser() {
  try {
    const email = String(
      (Session.getActiveUser() && Session.getActiveUser().getEmail()) ||
      (Session.getEffectiveUser() && Session.getEffectiveUser().getEmail()) ||
      ''
    ).trim().toLowerCase();

    const row = _findUserByEmail_(email);
    const client = _toClientUser_(row, email);

    // Hydrate campaign context
    try {
      if (client.CampaignID) {
        if (typeof getCampaignById === 'function') {
          const c = getCampaignById(client.CampaignID);
          client.campaignName = c ? (c.Name || c.name || '') : '';
        }
        if (typeof getUserCampaignPermissions === 'function') {
          client.campaignPermissions = getUserCampaignPermissions(client.ID);
        }
        if (typeof getCampaignNavigation === 'function') {
          client.campaignNavigation = getCampaignNavigation(client.CampaignID);
        }
      }
    } catch (ctxErr) {
      console.warn('getCurrentUser: campaign context hydrate failed:', ctxErr);
    }

    return client;
  } catch (e) {
    writeError && writeError('getCurrentUser', e);
    return _toClientUser_(null, '');
  }
}

/**
 * Simple authentication using token parameter
 */
function authenticateUser(e) {
  try {
    // First check if there's a token parameter
    const token = e.parameter.token;
    if (token && typeof AuthenticationService !== 'undefined' && AuthenticationService.getSessionUser) {
      const user = AuthenticationService.getSessionUser(token);
      if (user) {
        return user;
      }
    }

    // Fall back to current user via Google session
    const user = getCurrentUser();
    if (!user || !user.ID) {
      return null;
    }

    // Check if user can login
    if (!_truthy(user.CanLogin)) {
      return null;
    }

    return user;
  } catch (error) {
    console.error('Authentication error:', error);
    writeError('authenticateUser', error);
    return null;
  }
}

/**
 * Generate URL with token if needed
 */
function getAuthenticatedUrl(page, campaignId, additionalParams = {}) {
  let url = SCRIPT_URL;
  const params = new URLSearchParams();
  
  if (page) {
    params.set('page', page);
  }
  
  if (campaignId) {
    params.set('campaign', campaignId);
  }
  
  // Add additional parameters
  Object.entries(additionalParams).forEach(([key, value]) => {
    if (value !== null && value !== undefined) {
      params.set(key, value);
    }
  });
  
  const queryString = params.toString();
  return queryString ? `${url}?${queryString}` : url;
}

function getBaseUrl() {
  return SCRIPT_URL;
}

// ───────────────────────────────────────────────────────────────────────────────
// USER MANAGEMENT FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

function _findUserByEmail_(email) {
  if (!email) return null;
  try {
    const CK = 'USR_BY_EMAIL_' + email.toLowerCase();
    const cached = _cacheGet(CK);
    if (cached) return cached;

    const rows = (typeof readSheet === 'function') ? (readSheet('Users') || []) : [];
    const hit = rows.find(r => String(r.Email || '').trim().toLowerCase() === email.toLowerCase()) || null;
    if (hit) _cachePut(CK, hit, 120);
    return hit;
  } catch (e) {
    writeError && writeError('_findUserByEmail_', e);
    return null;
  }
}

function _toClientUser_(row, fallbackEmail) {
  const rolesMap = (typeof getRolesMapping === 'function') ? getRolesMapping() : {};
  const roleIds = String(row && row.Roles || '')
    .split(',').map(s => s.trim()).filter(Boolean);
  const roleNames = roleIds.map(id => rolesMap[id]).filter(Boolean);

  const isAdminFlag =
    (String(row && row.IsAdmin).toLowerCase() === 'true') ||
    roleNames.some(n => /system\s*admin|administrator|super\s*admin/i.test(String(n || '')));

  const client = {
    ID: String(row && row.ID || ''),
    Email: String(row && row.Email || fallbackEmail || '').trim(),
    FullName: String(row && (row.FullName || row.UserName || '') || '').trim() ||
      String(fallbackEmail || '').split('@')[0],
    CampaignID: String(row && (row.CampaignID || row.CampaignId || '') || ''),
    roleNames: roleNames,
    IsAdmin: !!isAdminFlag,
    CanLogin: (row && row.CanLogin !== undefined) ? _truthy(row.CanLogin) : true,
    EmailConfirmed: (row && row.EmailConfirmed !== undefined) ? _truthy(row.EmailConfirmed) : true,
    ResetRequired: !!(row && _truthy(row.ResetRequired)),
    Pages: String(row && (row.Pages || row.pages || '') || '')
  };
  return client;
}

function clientGetCurrentUser() {
  return getCurrentUser();
}

function __injectCurrentUser_(tpl, explicitUser) {
  try {
    const u = explicitUser && (explicitUser.Email || explicitUser.email || explicitUser.ID)
      ? explicitUser
      : getCurrentUser();
    tpl.user = u;
    tpl.currentUserJson = _stringifyForTemplate_(u);
  } catch (e) {
    writeError && writeError('__injectCurrentUser_', e);
    tpl.user = null;
    tpl.currentUserJson = '{}';
  }
}


// ───────────────────────────────────────────────────────────────────────────────
// ACCESS CONTROL SYSTEM
// ───────────────────────────────────────────────────────────────────────────────

function _normalizeUser(user) {
  const out = Object.assign({}, user || {});
  out.IsAdmin = _truthy(out.IsAdmin) || String((out.roleNames || [])).toLowerCase().includes('system admin');
  out.CanLogin = _truthy(out.CanLogin !== undefined ? out.CanLogin : true);
  out.EmailConfirmed = _truthy(out.EmailConfirmed !== undefined ? out.EmailConfirmed : true);
  out.ResetRequired = _truthy(out.ResetRequired);
  out.LockoutEnd = out.LockoutEnd ? _safeDate(out.LockoutEnd) : null;
  out.ID = _normalizeId(out.ID || out.id);
  out.CampaignID = _normalizeId(out.CampaignID || out.campaignId);
  out.PagesCsv = String(out.Pages || out.pages || '');
  return out;
}

function isSystemAdmin(user) {
  try {
    const u = _normalizeUser(user);
    return !!u.IsAdmin;
  } catch (e) {
    writeError && writeError('isSystemAdmin', e);
    return false;
  }
}

function evaluatePageAccess(user, pageKey, campaignId) {
  const trace = [];
  try {
    const u = _normalizeUser(user);
    const page = _normalizePageKey(pageKey || '');
    const cid = _normalizeId(campaignId || '');

    // Basic account checks
    if (!u || !u.ID) return { allow: false, reason: 'No session', trace };
    if (!_truthy(u.CanLogin)) return { allow: false, reason: 'Account disabled', trace };
    if (u.LockoutEnd && _isFuture(u.LockoutEnd)) return { allow: false, reason: 'Account locked', trace };
    if (!ACCESS.PUBLIC_PAGES.has(page)) {
      if (!_truthy(u.EmailConfirmed)) return { allow: false, reason: 'Email not confirmed', trace };
      if (_truthy(u.ResetRequired)) return { allow: false, reason: 'Password reset required', trace };
    }

    // Public pages
    if (ACCESS.PUBLIC_PAGES.has(page)) {
      trace.push('PUBLIC page');
      return { allow: true, reason: 'public', trace };
    }

    // Admin-only pages
    if (ACCESS.ADMIN_ONLY_PAGES.has(page)) {
      if (isSystemAdmin(u)) {
        trace.push('admin-only: allowed');
        return { allow: true, reason: 'admin', trace };
      }
      return { allow: false, reason: 'System Admin required', trace };
    }

    // System admin has access to everything
    if (isSystemAdmin(u)) {
      trace.push('System Admin: allow');
      return { allow: true, reason: 'admin', trace };
    }

    // For regular users, check if they have campaign access
    if (cid && !hasCampaignAccess(u, cid)) {
      return { allow: false, reason: 'No campaign access', trace };
    }

    // Default allow for authenticated users
    trace.push('Authenticated user access');
    return { allow: true, reason: 'authenticated', trace };

  } catch (e) {
    writeError && writeError('evaluatePageAccess', e);
    trace.push('exception:' + e.message);
    return { allow: false, reason: 'Evaluator error', trace };
  }
}

function hasCampaignAccess(user, campaignId) {
  try {
    const u = _normalizeUser(user);
    const cid = _normalizeId(campaignId);
    if (!cid) return true;
    if (isSystemAdmin(u)) return true;
    if (_normalizeId(u.CampaignID) === cid) return true;
    return false;
  } catch (e) {
    writeError && writeError('hasCampaignAccess', e);
    return false;
  }
}

function renderAccessDenied(message) {
  const baseUrl = getBaseUrl();
  const tpl = HtmlService.createTemplateFromFile('AccessDenied');
  tpl.baseUrl = baseUrl;
  tpl.message = message || 'You do not have permission to view this page.';
  return tpl.evaluate()
    .setTitle('Access Denied')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Simplified authentication requirement function
 */
function requireAuth(e) {
  try {
    const user = authenticateUser(e);

    if (!user) {
      return renderLoginPage();
    }

    const pageParam = String(e?.parameter?.page || '').toLowerCase();
    const page = canonicalizePageKey(pageParam);
    const campaignId = String(e?.parameter?.campaign || user.CampaignID || '');

    const decision = evaluatePageAccess(user, page, campaignId);
    _debugAccess && _debugAccess('route', decision, user, page, campaignId);

    if (!decision || decision.allow !== true) {
      return renderAccessDenied((decision && decision.reason) || 'You do not have permission to view this page.');
    }

    // Hydrate campaign context
    if (campaignId) {
      try {
        if (typeof getCampaignNavigation === 'function') user.campaignNavigation = getCampaignNavigation(campaignId);
        if (typeof getUserCampaignPermissions === 'function') user.campaignPermissions = getUserCampaignPermissions(user.ID);
        if (typeof getCampaignById === 'function') {
          const c = getCampaignById(campaignId);
          user.campaignName = c ? (c.Name || c.name || '') : '';
        }
      } catch (ctxErr) {
        console.warn('Campaign context hydrate failed:', ctxErr);
      }
    }

    return user;

  } catch (error) {
    writeError('requireAuth', error);
    return renderAccessDenied('Authentication error occurred');
  }
}

function renderLoginPage() {
  const tpl = HtmlService.createTemplateFromFile('Login');
  tpl.baseUrl = getBaseUrl();
  tpl.scriptUrl = SCRIPT_URL;
  return tpl.evaluate()
    .setTitle('Login - VLBPO LuminaHQ')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function canonicalizePageKey(k) {
  const key = String(k || '').trim().toLowerCase();
  if (!key) return key;

  // Map legacy slugs/aliases → canonical keys used by the Access Engine
  switch (key) {
    // Admin Pages
    case 'manageuser':
    case 'users':
      return 'admin.users';
    case 'manageroles':
    case 'roles':
      return 'admin.roles';
    case 'managecampaign':
    case 'campaigns':
      return 'admin.campaigns';

    // Task Management (Default Pages)
    case 'tasklist':
    case 'task-list':
      return 'tasks.list';
    case 'taskboard':
    case 'task-board':
      return 'tasks.board';
    case 'taskform':
    case 'task-form':
    case 'newtask':
    case 'edittask':
      return 'tasks.form';
    case 'taskview':
    case 'task-view':
      return 'tasks.view';

    // Communication & Collaboration (Default Pages)
    case 'chat':
      return 'global.chat';
    case 'search':
      return 'global.search';
    case 'bookmarks':
      return 'global.bookmarks';

    // Schedule Management (Default Page)
    case 'schedule':
    case 'schedulemanagement':
      return 'global.schedule';

    // Other Global Pages
    case 'notifications':
      return 'global.notifications';
    case 'settings':
      return 'global.settings';

    // QA System Pages
    case 'qa-dashboard':
    case 'ibtrqualityreports':
      return 'qa.dashboard';
    case 'qualityform':
    case 'independencequality':
    case 'creditsuiteqa':
      return 'qa.form';
    case 'qualityview':
      return 'qa.view';
    case 'qualitylist':
      return 'qa.list';

    // Report Pages
    case 'callreports':
      return 'reports.calls';
    case 'attendancereports':
      return 'reports.attendance';

    // Coaching Pages
    case 'coachingdashboard':
      return 'coaching.dashboard';
    case 'coachinglist':
    case 'coachings':
      return 'coaching.list';
    case 'coachingview':
      return 'coaching.view';
    case 'coachingsheet':
    case 'coaching':
      return 'coaching.form';

    // Calendar and Schedule
    case 'attendancecalendar':
      return 'calendar.attendance';
    case 'slotmanagement':
      return 'schedule.slots';

    // Dashboard
    case 'dashboard':
      return 'dashboard';

    // Proxy and Special Pages
    case 'proxy':
      return 'global.proxy';

    default:
      return key; // unknowns fall through; let the engine decide
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// ENHANCED doGet WITH SIMPLIFIED ROUTING
// ───────────────────────────────────────────────────────────────────────────────

function doGet(e) {
  try {
    const baseUrl = getBaseUrl();

    // Initialize system
    initializeSystem();

    // Handle special actions
    if (e.parameter.page === 'proxy') {
      console.log('doGet: Handling proxy request');
      return serveEnhancedProxy(e);
    }

    if (e.parameter.action === 'logout') {
      console.log('doGet: Handling logout action');
      return handleLogoutRequest(e);
    }

    if (e.parameter.action === 'confirmEmail' && e.parameter.token) {
      const confirmationToken = e.parameter.token;
      const success = confirmEmail(confirmationToken);
      if (success) {
        const redirectUrl = `${baseUrl}?page=setpassword&token=${encodeURIComponent(confirmationToken)}`;
        return HtmlService
          .createHtmlOutput(`<script>window.location.href = "${redirectUrl}";</script>`)
          .setTitle('Redirecting...');
      } else {
        const tpl = HtmlService.createTemplateFromFile('EmailConfirmed');
        tpl.baseUrl = baseUrl;
        tpl.success = false;
        tpl.token = confirmationToken;
        return tpl.evaluate()
          .setTitle('Email Confirmation')
          .addMetaTag('viewport', 'width=device-width,initial-scale=1');
      }
    }

    // Handle public pages
    const page = (e.parameter.page || "").toLowerCase();

    if (page === 'login' || (!page)) {
      return renderLoginPage();
    }

    // Handle other public pages
    const publicPages = ['setpassword', 'resetpassword', 'resend-verification', 'resendverification',
      'forgotpassword', 'forgot-password', 'emailconfirmed', 'email-confirmed'];

    if (publicPages.includes(page)) {
      return handlePublicPage(page, e, baseUrl);
    }

    // Protected pages - require authentication
    const auth = requireAuth(e);
    if (auth.getContent) {
      return auth; // Login or access denied page
    }

    const user = auth;

    // Handle password reset requirement
    if (_truthy(user.ResetRequired)) {
      const tpl = HtmlService.createTemplateFromFile('ChangePassword');
      tpl.baseUrl = baseUrl;
      tpl.scriptUrl = SCRIPT_URL;
      return tpl.evaluate()
        .setTitle('Change Password')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // Default route: redirect to dashboard
    if (!page) {
      const userCampaignId = user.CampaignID || '';
      const redirectUrl = getAuthenticatedUrl('dashboard', userCampaignId);
      return HtmlService
        .createHtmlOutput(`<script>window.location.href = "${redirectUrl}";</script>`)
        .setTitle('Redirecting to Dashboard...');
    }

    const campaignId = e.parameter.campaign || user.CampaignID || '';

    // Handle CSV exports
    if (page === "callreports" && e.parameter.action === "exportCallCsv") {
      const gran = e.parameter.granularity || "Week";
      const period = e.parameter.period || weekStringFromDate(new Date());
      const agent = e.parameter.agent || "";
      const csv = exportCallAnalyticsCsv(gran, period, agent);
      return ContentService.createTextOutput(csv)
        .setMimeType(ContentService.MimeType.CSV)
        .downloadAsFile(`callAnalytics_${period}.csv`);
    }

    if (page === 'attendancereports' && e.parameter.action === 'exportCsv') {
      const gran = e.parameter.granularity || 'Week';
      const period = e.parameter.period || weekStringFromDate(new Date());
      const agent = e.parameter.agent || '';
      const csv = exportAttendanceCsv(gran, period, agent);
      return ContentService.createTextOutput(csv)
        .setMimeType(ContentService.MimeType.CSV)
        .downloadAsFile(`attendance_${period}.csv`);
    }

    // Route to appropriate page
    return routeToPage(page, e, baseUrl, user, campaignId);

  } catch (error) {
    console.error('Error in doGet:', error);
    writeError('doGet', error);
    return createErrorPage('System Error', `An error occurred: ${error.message}`);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// ENHANCED ROUTING WITH CAMPAIGN SUPPORT (Updated for Clean URLs)
// ───────────────────────────────────────────────────────────────────────────────

function routeToPage(page, e, baseUrl, user, campaignIdFromCaller) {
  try {
    const raw = String(page || '').trim();

    // ═══════════════════════════════════════════════════════════════════════════
    // DEFAULT/GLOBAL PAGES (Always available, campaign-independent)
    // ═══════════════════════════════════════════════════════════════════════════

    // Task Management (Default Pages)
    if (page === "tasklist" || page === "task-list") {
      return serveGlobalPage('TaskList', e, baseUrl, user);
    }

    if (page === "taskboard" || page === "task-board") {
      return serveGlobalPage('TaskBoard', e, baseUrl, user);
    }

    if (page === "taskform" || page === "task-form" || page === "newtask" || page === "edittask") {
      return serveGlobalPage('TaskForm', e, baseUrl, user);
    }

    if (page === "taskview" || page === "task-view") {
      return serveGlobalPage('TaskView', e, baseUrl, user);
    }

    // Communication & Collaboration (Default Pages)
    if (page === 'chat') {
      return serveGlobalPage('Chat', e, baseUrl, user);
    }

    if (page === 'search') {
      return serveGlobalPage('Search', e, baseUrl, user);
    }

    if (page === 'bookmarks') {
      return serveGlobalPage('BookmarkManager', e, baseUrl, user);
    }

    // Administration (Default Pages)
    if (page === 'users' || page === 'manageuser') {
      return serveAdminPage('Users', e, baseUrl, user);
    }

    if (page === 'roles' || page === 'manageroles') {
      return serveAdminPage('RoleManagement', e, baseUrl, user);
    }

    if (page === 'campaigns' || page === 'managecampaign') {
      return serveAdminPage('CampaignManagement', e, baseUrl, user);
    }

    // Schedule Management (Default Page)
    if (page === 'schedule' || page === 'schedulemanagement') {
      return serveGlobalPage('ScheduleManagement', e, baseUrl, user);
    }

    // Other Default Global Pages
    if (page === 'notifications') {
      return serveGlobalPage('Notifications', e, baseUrl, user);
    }

    if (page === 'settings') {
      return serveGlobalPage('Settings', e, baseUrl, user);
    }

    if (page === 'proxy') {
      return serveProxy(e);
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // CAMPAIGN-SPECIFIC ROUTING
    // ═══════════════════════════════════════════════════════════════════════════

    // Campaign-specific routing pattern: Case.PageKind
    const PATTERN = /^([^._:\-][^._:\-]*(?:[ _:\-][^._:\-]+)*)[._:\-](QAForm|QADashboard|QAList|QAView|QACollabList|AttendanceReports|CallReports|ChatReport|Dashboard|CoachingForm|TaskForm|TaskList|TaskBoard)$/i;
    const match = PATTERN.exec(raw);

    if (match) {
      const caseToken = match[1].trim();
      const pageKind = match[2];
      const def = __case_resolve__(caseToken);

      if (!def) {
        return createErrorPage('Unknown Campaign', 'No case mapping found for: ' + caseToken);
      }

      // Determine campaign ID
      const explicitCid = String(e?.parameter?.campaign || '').trim();
      const sheetCid = __case_findCampaignId__(def);
      const cid = explicitCid || campaignIdFromCaller || sheetCid || (user && user.CampaignID) || '';

      const candidates = __case_templateCandidates__(def, pageKind);
      return __case_serve__(candidates, e, baseUrl, user, cid);
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // LEGACY/BACKWARD COMPATIBILITY ROUTES
    // ═══════════════════════════════════════════════════════════════════════════

    switch (page) {
      case "dashboard":
        return serveCampaignPage('Dashboard', e, baseUrl, user, campaignIdFromCaller);

      case "qualityform":
        return routeQAForm(e, baseUrl, user, campaignIdFromCaller);

      case "ibtrqualityreports":
        return routeQADashboard(e, baseUrl, user, campaignIdFromCaller);

      case 'qualityview':
        return routeQAView(e, baseUrl, user, campaignIdFromCaller);

      case 'qualitylist':
        return routeQAList(e, baseUrl, user, campaignIdFromCaller);

      case "independencequality":
        return serveCampaignPage('IndependenceQAForm', e, baseUrl, user, campaignIdFromCaller);

      case "independenceqadashboard":
        return serveCampaignPage('UnifiedQADashboard', e, baseUrl, user, campaignIdFromCaller);

      case "creditsuiteqa":
        return serveCampaignPage('CreditSuiteQAForm', e, baseUrl, user, campaignIdFromCaller);

      case "qa-dashboard":
        return serveCampaignPage('CreditSuiteQADashboard', e, baseUrl, user, campaignIdFromCaller);

      case "callreports":
        return serveCampaignPage('CallReports', e, baseUrl, user, campaignIdFromCaller);

      case "attendancereports":
        return serveCampaignPage('AttendanceReports', e, baseUrl, user, campaignIdFromCaller);

      case "coachingdashboard":
        return serveCampaignPage('CoachingDashboard', e, baseUrl, user, campaignIdFromCaller);

      case "coachingview":
        return serveCampaignPage('CoachingView', e, baseUrl, user, campaignIdFromCaller);

      case "coachinglist":
      case "coachings":
        return serveCampaignPage('CoachingList', e, baseUrl, user, campaignIdFromCaller);

      case "coachingsheet":
      case "coaching":
        return serveCampaignPage('CoachingForm', e, baseUrl, user, campaignIdFromCaller);

      case "eodreport":
        return serveGlobalPage('EODReport', e, baseUrl, user);

      case "attendancecalendar":
        return serveCampaignPage('Calendar', e, baseUrl, user, campaignIdFromCaller);

      case "escalations":
        return serveCampaignPage('Escalations', e, baseUrl, user, campaignIdFromCaller);

      case "incentives":
        return serveCampaignPage('Incentives', e, baseUrl, user, campaignIdFromCaller);

      case 'import':
        return serveAdminPage('ImportCsv', e, baseUrl, user);

      case 'importattendance':
        return serveAdminPage('ImportAttendance', e, baseUrl, user);

      case 'slotmanagement':
        return serveShiftSlotManagement(e, baseUrl, user, campaignIdFromCaller);

      case 'ackform':
        return serveAckForm(e, baseUrl, user);

      case "agent-schedule":
        return serveAgentSchedulePage(e, baseUrl, e.parameter.token);

      case 'import':
        // Allow both admin and campaign-level access for general imports
        if (isSystemAdmin(user)) {
          return serveAdminPage('ImportCsv', e, baseUrl, user);
        } else {
          return serveCampaignPage('ImportCsv', e, baseUrl, user, campaignIdFromCaller);
        }

      case 'importattendance':
        // Allow managers and supervisors for attendance imports
        if (isSystemAdmin(user) || hasManagerRole(user) || hasSupervisorRole(user)) {
          return serveAdminPage('ImportAttendance', e, baseUrl, user);
        } else {
          return renderAccessDenied('You need manager or supervisor privileges to import attendance data.');
        }

      default:
        // Unknown page - redirect to dashboard
        const defaultCampaignId = user.CampaignID || '';
        const redirectUrl = getCampaignUrl('dashboard', defaultCampaignId);

        return HtmlService
          .createHtmlOutput(`<script>window.location.href = "${redirectUrl}";</script>`)
          .setTitle('Redirecting to Dashboard...');
    }
  } catch (error) {
    console.error(`Error routing to page ${page}:`, error);
    writeError('routeToPage', error);
    return createErrorPage('Page Error', `Error loading page: ${error.message}`);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// PAGE SERVING FUNCTIONS (Updated for Clean URLs)
// ───────────────────────────────────────────────────────────────────────────────

function serveCampaignPage(templateName, e, baseUrl, user, campaignId) {
  try {
    const tpl = HtmlService.createTemplateFromFile(templateName);
    tpl.baseUrl = baseUrl;
    tpl.scriptUrl = SCRIPT_URL;
    tpl.currentPage = templateName.replace(/([A-Z])/g, ' $1').trim();
    tpl.user = user;
    tpl.campaignId = campaignId;

    __injectCurrentUser_(tpl, user);

    if (campaignId) {
      tpl.campaignName = (tpl.user && tpl.user.campaignName) || '';
      tpl.campaignNavigation = (tpl.user && tpl.user.campaignNavigation) || { categories: [], uncategorizedPages: [] };
    }

    addTemplateSpecificData(tpl, templateName, e, tpl.user, campaignId);

    return tpl.evaluate()
      .setTitle(tpl.currentPage + ' - VLBPO LuminaHQ')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    writeError(`serveCampaignPage(${templateName})`, error);
    return createErrorPage('Failed to load page', error.message);
  }
}

function serveGlobalPage(templateName, e, baseUrl, user) {
  try {
    const tpl = HtmlService.createTemplateFromFile(templateName);
    tpl.baseUrl = baseUrl;
    tpl.scriptUrl = SCRIPT_URL;
    tpl.currentPage = templateName.replace(/([A-Z])/g, ' $1').trim();
    tpl.user = user;

    __injectCurrentUser_(tpl, user);

    addTemplateSpecificData(tpl, templateName, e, tpl.user, null);

    return tpl.evaluate()
      .setTitle(tpl.currentPage + ' - VLBPO LuminaHQ')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    writeError(`serveGlobalPage(${templateName})`, error);
    return createErrorPage('Failed to load page', error.message);
  }
}

function serveAdminPage(templateName, e, baseUrl, user) {
  // Single admin check - no need for redundancy
  if (!isSystemAdmin(user)) {
    return renderAccessDenied('This page requires System Admin privileges.');
  }

  try {
    const tpl = HtmlService.createTemplateFromFile(templateName);
    tpl.baseUrl = baseUrl;
    tpl.scriptUrl = SCRIPT_URL;
    tpl.currentPage = templateName.replace(/([A-Z])/g, ' $1').trim();
    tpl.user = user;

    // Inject current user data consistently with other serve functions
    __injectCurrentUser_(tpl, user);

    // Add admin page specific data
    addTemplateSpecificData(tpl, templateName, e, user, null);

    return tpl.evaluate()
      .setTitle(tpl.currentPage + ' - VLBPO LuminaHQ')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    console.error(`Error serving admin page ${templateName}:`, error);
    writeError(`serveAdminPage(${templateName})`, error);
    return createErrorPage('Failed to load page', error.message);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// PUBLIC PAGE HANDLERS (Updated for Clean URLs)
// ───────────────────────────────────────────────────────────────────────────────

function handlePublicPage(page, e, baseUrl) {
  const scriptUrl = SCRIPT_URL;

  switch (page) {
    case 'setpassword':
    case 'resetpassword':
      const resetToken = e.parameter.token || '';
      const tpl = HtmlService.createTemplateFromFile('ChangePassword');
      tpl.baseUrl = baseUrl;
      tpl.scriptUrl = scriptUrl;
      tpl.token = resetToken;
      tpl.isReset = page === 'resetpassword';

      return tpl.evaluate()
        .setTitle((page === 'resetpassword' ? 'Reset' : 'Set') + ' Your Password - VLBPO LuminaHQ')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    case 'resend-verification':
    case 'resendverification':
      const verifyTpl = HtmlService.createTemplateFromFile('ResendVerification');
      verifyTpl.baseUrl = baseUrl;
      verifyTpl.scriptUrl = scriptUrl;

      return verifyTpl.evaluate()
        .setTitle('Resend Email Verification - VLBPO LuminaHQ')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    case 'forgotpassword':
    case 'forgot-password':
      const forgotTpl = HtmlService.createTemplateFromFile('ForgotPassword');
      forgotTpl.baseUrl = baseUrl;
      forgotTpl.scriptUrl = scriptUrl;

      return forgotTpl.evaluate()
        .setTitle('Forgot Password - VLBPO LuminaHQ')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    case 'emailconfirmed':
    case 'email-confirmed':
      const confirmTpl = HtmlService.createTemplateFromFile('EmailConfirmed');
      confirmTpl.baseUrl = baseUrl;
      confirmTpl.success = e.parameter.success === 'true';
      confirmTpl.token = e.parameter.token || '';

      return confirmTpl.evaluate()
        .setTitle('Email Confirmation - VLBPO LuminaHQ')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    default:
      return createErrorPage('Page Not Found', `The page "${page}" was not found.`);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// REST OF THE CODE (INCLUDES ALL EXISTING FUNCTIONS)
// These functions remain largely unchanged but with clean URL updates where needed
// ───────────────────────────────────────────────────────────────────────────────

// [Include all existing functions from the original Code.gs with URL updates]
// Campaign resolution helpers, template handling, user management, etc.
// The implementation continues with all existing functionality...

// Campaign resolution helpers
function __case_allCampaignsRaw__() {
  try {
    return (typeof readSheet === 'function') ? (readSheet('CAMPAIGNS') || []) : [];
  } catch (_) {
    return [];
  }
}

function __case_findCampaignId__(def) {
  try {
    const rows = __case_allCampaignsRaw__();
    const aliasSet = new Set([def.case].concat(def.aliases || []));
    const aliasSlugs = new Set(Array.from(aliasSet).map(__case_slug__));

    const hit = rows.find(r => aliasSet.has(String(r.Name || r.name || '').trim())) ||
      rows.find(r => aliasSlugs.has(__case_slug__(String(r.Name || r.name || '').trim())));
    if (hit) return String(hit.ID || hit.Id || hit.id || '').trim() || def.idHint || '';
  } catch (_) { }
  return def.idHint || '';
}

function __case_registry__() {
  const byCase = new Map();
  const byAlias = new Map();
  CASE_DEFS.forEach(def => {
    byCase.set(def.case, def);
    (def.aliases || []).forEach(a => byAlias.set(__case_norm__(a), def));
    byAlias.set(__case_norm__(def.case), def);
  });
  return { byCase, byAlias };
}

function __case_resolve__(token) {
  const reg = __case_registry__();
  const key = __case_norm__(token);
  return reg.byAlias.get(key) || null;
}

function __case_templateCandidates__(def, pageKind) {
  const chosen = (def.pages || {})[pageKind];
  const candidates = [];
  if (chosen) candidates.push(chosen);

  // Special handling for Independence dashboard
  if (pageKind === 'QADashboard' && def.case === 'IndependenceInsuranceAgency') {
    candidates.push('UnifiedQADashboard');
  }

  // Add generic fallbacks
  (GENERIC_FALLBACKS[pageKind] || []).forEach(t => candidates.push(t));

  // Remove duplicates
  return candidates.filter((v, i, a) => a.indexOf(v) === i);
}

function __case_serve__(candidates, e, baseUrl, user, campaignId) {
  for (var i = 0; i < candidates.length; i++) {
    try {
      HtmlService.createTemplateFromFile(candidates[i]); // Test existence
      return serveCampaignPage(candidates[i], e, baseUrl, user, campaignId);
    } catch (_) {
      // Template doesn't exist, try next
    }
  }
  return createErrorPage('Missing Page', `Template not found for this campaign/page. Tried: ${candidates.join(', ')}`);
}

// QA System routing
function routeQAForm(e, baseUrl, user, campaignId) {
  const campaign = getCampaignById(campaignId);
  const campaignName = campaign?.Name?.toLowerCase() || '';

  if (campaignName.includes('independence') || campaignName.includes('insurance')) {
    return serveCampaignPage('IndependenceQAForm', e, baseUrl, user, campaignId);
  } else if (campaignName.includes('credit') || campaignName.includes('suite')) {
    return serveCampaignPage('CreditSuiteQAForm', e, baseUrl, user, campaignId);
  } else {
    return serveCampaignPage('QualityForm', e, baseUrl, user, campaignId);
  }
}

function routeQADashboard(e, baseUrl, user, campaignId) {
  const campaign = getCampaignById(campaignId);
  const campaignName = campaign?.Name?.toLowerCase() || '';

  if (campaignName.includes('independence') || campaignName.includes('insurance')) {
    return serveCampaignPage('UnifiedQADashboard', e, baseUrl, user, campaignId);
  } else if (campaignName.includes('credit') || campaignName.includes('suite')) {
    return serveCampaignPage('CreditSuiteQADashboard', e, baseUrl, user, campaignId);
  } else {
    return serveCampaignPage('QADashboard', e, baseUrl, user, campaignId);
  }
}

function routeQAView(e, baseUrl, user, campaignId) {
  const campaign = getCampaignById(campaignId);
  const campaignName = campaign?.Name?.toLowerCase() || '';

  if (campaignName.includes('independence') || campaignName.includes('insurance')) {
    return serveCampaignPage('IndependenceQAView', e, baseUrl, user, campaignId);
  } else if (campaignName.includes('credit') || campaignName.includes('suite')) {
    return serveCampaignPage('CreditSuiteQAView', e, baseUrl, user, campaignId);
  } else {
    return serveCampaignPage('QualityView', e, baseUrl, user, campaignId);
  }
}

function routeQAList(e, baseUrl, user, campaignId) {
  const campaign = getCampaignById(campaignId);
  const campaignName = campaign?.Name?.toLowerCase() || '';

  if (campaignName.includes('independence') || campaignName.includes('insurance')) {
    return serveCampaignPage('IndependenceQAList', e, baseUrl, user, campaignId);
  } else if (campaignName.includes('credit') || campaignName.includes('suite')) {
    return serveCampaignPage('CreditSuiteQAList', e, baseUrl, user, campaignId);
  } else {
    return serveCampaignPage('QAList', e, baseUrl, user, campaignId);
  }
}

// [Continue with all remaining functions from the original Code.gs...]
// This includes all the template data handling, user management, utility functions, etc.
// The functions remain the same but with updated URL generation

// Template data handling functions
function addTemplateSpecificData(tpl, templateName, e, user, campaignId) {
  try {
    const timeZone = Session.getScriptTimeZone();
    const granularity = (e.parameter.granularity || "Week").toString();
    const periodValue = (e.parameter.period || weekStringFromDate(new Date())).toString();
    const selectedAgent = (e.parameter.agent || "").toString();
    const selectedCampaign = (e.parameter.campaign || campaignId || "").toString();

    // Set common template properties
    tpl.granularity = granularity;
    tpl.period = periodValue;
    tpl.periodValue = periodValue;
    tpl.selectedPeriod = periodValue;
    tpl.selectedAgent = selectedAgent;
    tpl.selectedCampaign = selectedCampaign;
    tpl.timeZone = timeZone;

    handleTemplateSpecificData(tpl, templateName, e, user, campaignId);
  } catch (error) {
    console.error(`Error adding template data for ${templateName}:`, error);
    writeError(`addTemplateSpecificData(${templateName})`, error);
  }
}

function handleTemplateSpecificData(tpl, templateName, e, user, campaignId) {
  try {
    switch (templateName) {
      case 'Dashboard':
        handleDashboardData(tpl, e, user, campaignId);
        break;

      case 'AttendanceReports':
        handleAttendanceReportsData(tpl, e, user, campaignId);
        break;

      case 'CallReports':
        handleCallReportsData(tpl, e, user, campaignId);
        break;

      // Campaign-specific QA forms
      case 'IndependenceQAForm':
      case 'CreditSuiteQAForm':
      case 'HiyaCarQAForm':
      case 'IBTRQAForm':
      case 'JSCQAForm':
      case 'KidsInTheGameQAForm':
      case 'KofiGroupQAForm':
      case 'PAWLawFirmQAForm':
      case 'ProHousePhotosQAForm':
      case 'IACSQAForm':
      case 'ProozyQAForm':
      case 'TheGroundingQAForm':
      case 'COQAForm':
        handleQAFormData(tpl, e, user, templateName, campaignId);
        break;

      case 'QualityForm':
        handleQAFormData(tpl, e, user, templateName, campaignId);
        break;

      // Campaign-specific QA views
      case 'IndependenceQAView':
      case 'CreditSuiteQAView':
      case 'HiyaCarQAView':
      case 'IBTRQAView':
      case 'JSCQAView':
      case 'KidsInTheGameQAView':
      case 'KofiGroupQAView':
      case 'PAWLawFirmQAView':
      case 'ProHousePhotosQAView':
      case 'IACSQAView':
      case 'ProozyQAView':
      case 'TheGroundingQAView':
      case 'COQAView':
        handleQAViewData(tpl, e, user, templateName);
        break;

      case 'QualityView':
        handleQAViewData(tpl, e, user, templateName);
        break;

      // Campaign-specific QA dashboards
      case 'IndependenceQADashboard':
      case 'UnifiedQADashboard':
      case 'CreditSuiteQADashboard':
      case 'HiyaCarQADashboard':
      case 'IBTRQADashboard':
      case 'JSCQADashboard':
      case 'KidsInTheGameQADashboard':
      case 'KofiGroupQADashboard':
      case 'PAWLawFirmQADashboard':
      case 'ProHousePhotosQADashboard':
      case 'IACSQADashboard':
      case 'ProozyQADashboard':
      case 'TheGroundingQADashboard':
      case 'COQADashboard':
        handleQADashboardData(tpl, e, user, templateName, campaignId);
        break;

      case 'QADashboard':
        handleQADashboardData(tpl, e, user, templateName, campaignId);
        break;

      // Campaign-specific QA lists
      case 'IndependenceQAList':
      case 'CreditSuiteQAList':
      case 'HiyaCarQAList':
      case 'IBTRQAList':
      case 'JSCQAList':
      case 'KidsInTheGameQAList':
      case 'KofiGroupQAList':
      case 'PAWLawFirmQAList':
      case 'ProHousePhotosQAList':
      case 'IACSQAList':
      case 'ProozyQAList':
      case 'TheGroundingQAList':
      case 'COQAList':
      case 'IBTRQACollabList':
        handleQAListData(tpl, e, user, templateName);
        break;

      case 'QAList':
        handleQAListData(tpl, e, user, templateName);
        break;

      case 'Users':
        handleUsersData(tpl, e, user);
        break;

      case 'TaskForm':
      case 'TaskList':
      case 'TaskBoard':
      case 'TaskView':
        handleTaskData(tpl, e, user, templateName);
        break;

      case 'CoachingForm':
      case 'CoachingList':
      case 'CoachingView':
      case 'CoachingDashboard':
        handleCoachingData(tpl, e, user, templateName, campaignId);
        break;

      case 'EODReport':
        handleEODReportData(tpl, e, user);
        break;

      case 'Calendar':
        handleCalendarData(tpl, e, user, campaignId);
        break;

      case 'Search':
        handleSearchData(tpl, e);
        break;

      case 'Chat':
        handleChatData(tpl, e, user);
        break;

      case 'Notifications':
        handleNotificationsData(tpl, e, user);
        break;

      default:
        console.log(`No specific data handler for template: ${templateName}`);
        break;
    }
  } catch (error) {
    console.error(`Error handling data for ${templateName}:`, error);
    writeError(`handleTemplateSpecificData(${templateName})`, error);
  }
}

// Individual template data handlers (these remain largely the same)
function handleDashboardData(tpl, e, user, campaignId) {
  try {
    const granularity = e.parameter.granularity || "Week";
    const periodValue = e.parameter.period || weekStringFromDate(new Date());
    const selectedAgent = e.parameter.agent || "";

    if (typeof getDashboardOkrs === 'function') {
      const okrData = getDashboardOkrs(granularity, periodValue, selectedAgent);
      tpl.okrData = JSON.stringify(okrData).replace(/<\/script>/g, '<\\/script>');
    } else {
      tpl.okrData = JSON.stringify({});
    }

    // Use manager-filtered user list
    const requestingUserId = user && user.ID ? user.ID : null;
    tpl.userList = clientGetAssignedAgentNames(campaignId || '', requestingUserId);
  } catch (error) {
    console.error('Error handling dashboard data:', error);
    tpl.okrData = JSON.stringify({});
    tpl.userList = [];
  }
}

function handleAttendanceReportsData(tpl, e, user, campaignId) {
  try {
    const granularity = e.parameter.granularity || "Week";
    const periodValue = e.parameter.period || weekStringFromDate(new Date());
    const selectedAgent = e.parameter.agent || "";

    if (typeof clientGetEnhancedAttendanceAnalytics === 'function') {
      const enhancedAnalytics = clientGetEnhancedAttendanceAnalytics(granularity, periodValue, selectedAgent);

      const allRows = enhancedAnalytics.filteredRows || [];
      const rowPage = parseInt(e.parameter.rowPage, 10) || 1;
      const PAGE_SIZE = 50;
      const startRowIdx = (rowPage - 1) * PAGE_SIZE;
      enhancedAnalytics.filteredRows = allRows.slice(startRowIdx, startRowIdx + PAGE_SIZE);

      tpl.attendanceData = JSON.stringify(enhancedAnalytics).replace(/<\/script>/g, '<\\/script>');
      tpl.currentRowPage = rowPage;
      tpl.totalRows = allRows.length;
      tpl.PAGE_SIZE = PAGE_SIZE;
      tpl.executiveMetrics = JSON.stringify(enhancedAnalytics.executiveMetrics || {}).replace(/<\/script>/g, '<\\/script>');
    } else {
      const attendanceAnalytics = getAttendanceAnalyticsByPeriod(granularity, periodValue, selectedAgent);
      const allRows = attendanceAnalytics.filteredRows || [];
      const rowPage = parseInt(e.parameter.rowPage, 10) || 1;
      const PAGE_SIZE = 50;
      const startRowIdx = (rowPage - 1) * PAGE_SIZE;
      attendanceAnalytics.filteredRows = allRows.slice(startRowIdx, startRowIdx + PAGE_SIZE);

      tpl.attendanceData = JSON.stringify(attendanceAnalytics).replace(/<\/script>/g, '<\\/script>');
      tpl.currentRowPage = rowPage;
      tpl.totalRows = allRows.length;
      tpl.PAGE_SIZE = PAGE_SIZE;
      tpl.executiveMetrics = JSON.stringify({});
    }

    // Use manager-filtered user list
    const requestingUserId = user && user.ID ? user.ID : null;
    tpl.userList = clientGetAssignedAgentNames(campaignId || user.CampaignID || '', requestingUserId);

  } catch (error) {
    console.error('Error handling attendance reports data:', error);
    writeError('handleAttendanceReportsData', error);
    tpl.attendanceData = JSON.stringify({ filteredRows: [], summary: {} });
    tpl.executiveMetrics = JSON.stringify({});
    tpl.userList = [];
  }
}

function handleCallReportsData(tpl, e, user, campaignId) {
  try {
    const granularity = e.parameter.granularity || "Week";
    const periodValue = e.parameter.period || weekStringFromDate(new Date());
    const selectedAgent = e.parameter.agent || "";

    if (typeof getAnalyticsByPeriod === 'function') {
      const analytics = getAnalyticsByPeriod("Week", periodValue, "");
      const rawReps = analytics.repMetrics || [];
      const pageNum = parseInt(e.parameter.page, 10) || 1;
      const PAGE_SIZE = 50;
      const startPageIdx = (pageNum - 1) * PAGE_SIZE;
      const pageSlice = rawReps.slice(startPageIdx, startPageIdx + PAGE_SIZE);

      const formattedReps = pageSlice.map((r) => {
        const totalCalls = r.totalCalls || 0;
        const totalTalkDecimal = parseFloat(r.totalTalk) || 0;
        const totalTalkFormatted = formatDuration ? formatDuration(totalTalkDecimal) : totalTalkDecimal;
        const avgTalkDecimal = totalCalls > 0 ? totalTalkDecimal / totalCalls : 0;
        const avgTalkFormatted = formatDuration ? formatDuration(avgTalkDecimal) : avgTalkDecimal;
        return {
          agent: r.agent,
          totalCalls: totalCalls,
          totalTalkFormatted: totalTalkFormatted,
          avgTalkFormatted: avgTalkFormatted,
        };
      });

      tpl.PAGE_SIZE = PAGE_SIZE;
      tpl.data = formattedReps;

      // Use manager-filtered user list
      const requestingUserId = user && user.ID ? user.ID : null;
      tpl.userList = clientGetAssignedAgentNames(campaignId || user.CampaignID || '', requestingUserId);

      // Chart data
      tpl.callVolumeLast7 = JSON.stringify(analytics.callTrend || []);
      tpl.hourlyHeatmapLast7 = JSON.stringify(analytics.hourlyHeatmap || []);
      tpl.avgIntervalByAgentLast7 = JSON.stringify(analytics.avgInterval || []);
      tpl.talkTimeByAgentLast7 = JSON.stringify(analytics.talkTrend || []);
      tpl.wrapupCountsLast7 = JSON.stringify(analytics.wrapDist || []);
      tpl.csatDistLast7 = JSON.stringify(analytics.csatDist || []);
      tpl.policyCountsLast7 = JSON.stringify(analytics.policyDist || []);
      tpl.agentLeaderboardLast7 = JSON.stringify(analytics.repMetrics || []);
    } else {
      tpl.PAGE_SIZE = 50;
      tpl.data = [];

      // Use manager-filtered user list
      const requestingUserId = user && user.ID ? user.ID : null;
      tpl.userList = clientGetAssignedAgentNames(campaignId || user.CampaignID || '', requestingUserId);

      tpl.callVolumeLast7 = JSON.stringify([]);
      tpl.talkTimeByAgentLast7 = JSON.stringify([]);
      tpl.wrapupCountsLast7 = JSON.stringify([]);
      tpl.csatDistLast7 = JSON.stringify([]);
      tpl.policyCountsLast7 = JSON.stringify([]);
      tpl.agentLeaderboardLast7 = JSON.stringify([]);
    }
  } catch (error) {
    console.error('Error handling call reports data:', error);
    writeError('handleCallReportsData', error);
    tpl.PAGE_SIZE = 50;
    tpl.data = [];
    tpl.userList = [];
  }
}

// Updated QA Form Data Handlers for Cookie-Based Auth
function handleQAFormData(tpl, e, user, templateName, campaignId) {
  try {
    const qaId = (e.parameter.id || "").toString();
    tpl.recordId = qaId;

    // Get manager-filtered user list based on template and user
    let userList = [];
    const requestingUserId = user && user.ID ? user.ID : null;

    console.log('handleQAFormData - Getting users for template:', templateName, 'campaignId:', campaignId, 'requestingUserId:', requestingUserId);

    if (templateName.includes('Independence')) {
      userList = clientGetIndependenceUsers(requestingUserId);
      tpl.campaignName = 'Independence Insurance';
    } else if (templateName.includes('CreditSuite')) {
      userList = clientGetCreditSuiteUsers(campaignId, requestingUserId);
      tpl.campaignName = 'Credit Suite';
    } else if (templateName.includes('IBTR')) {
      userList = clientGetIBTRUsers(requestingUserId);
      tpl.campaignName = 'Benefits Resource Center (iBTR)';
    } else {
      // For generic QA forms, use the new unified function
      userList = getQAFormUsers(campaignId, requestingUserId);
    }

    console.log('handleQAFormData - Final user list count:', userList.length);

    // Set both users and userList for backward compatibility
    tpl.users = userList;
    tpl.userList = userList;

    // Get existing record if editing
    if (qaId) {
      if (templateName.includes('Independence') && typeof clientGetIndependenceQAById === 'function') {
        const record = clientGetIndependenceQAById(qaId);
        tpl.record = record ? JSON.stringify(record).replace(/<\/script>/g, '<\\/script>') : "{}";
      } else if (templateName.includes('CreditSuite') && typeof clientGetCreditSuiteQAById === 'function') {
        const record = clientGetCreditSuiteQAById(qaId);
        tpl.record = record ? JSON.stringify(record).replace(/<\/script>/g, '<\\/script>') : "{}";
      } else if (typeof getQARecordById === 'function') {
        const record = getQARecordById(qaId);
        tpl.record = record ? JSON.stringify(record).replace(/<\/script>/g, '<\\/script>') : "{}";
      } else {
        tpl.record = "{}";
      }
    } else {
      tpl.record = "{}";
    }
  } catch (error) {
    console.error('Error handling QA form data:', error);
    tpl.users = [];
    tpl.userList = [];
    tpl.recordId = "";
    tpl.record = "{}";
  }
}

function handleQAViewData(tpl, e, user, templateName) {
  try {
    const qaViewId = (e.parameter.id || '').toString();
    tpl.recordId = qaViewId;

    if (!qaViewId) {
      tpl.record = "{}";
      return;
    }

    if (templateName.includes('Independence') && typeof clientGetIndependenceQAById === 'function') {
      const qaRec = clientGetIndependenceQAById(qaViewId);
      tpl.record = JSON.stringify(qaRec || {}).replace(/<\/script>/g, '<\\/script>');
      tpl.campaignName = 'Independence Insurance';
    } else if (templateName.includes('CreditSuite') && typeof clientGetCreditSuiteQAById === 'function') {
      const qaRec = clientGetCreditSuiteQAById(qaViewId);
      tpl.record = JSON.stringify(qaRec || {}).replace(/<\/script>/g, '<\\/script>');
      tpl.campaignName = 'Credit Suite';
    } else if (typeof getQARecordById === 'function') {
      const qaRec = getQARecordById(qaViewId);
      tpl.record = JSON.stringify(qaRec || {}).replace(/<\/script>/g, '<\\/script>');
    } else {
      tpl.record = "{}";
    }
  } catch (error) {
    console.error('Error handling QA view data:', error);
    tpl.record = "{}";
    tpl.recordId = "";
  }
}

function handleQADashboardData(tpl, e, user, templateName, campaignId) {
  try {
    const granularity = e.parameter.granularity || "Week";
    const periodValue = e.parameter.period || weekStringFromDate(new Date());
    const selectedAgent = e.parameter.agent || "";

    tpl.userList = clientGetAssignedAgentNames(campaignId || '');

    if (templateName === 'UnifiedQADashboard') {
      // For unified dashboard, get both campaign analytics
      try {
        if (typeof clientGetIndependenceQAAnalytics === 'function') {
          const independenceAnalytics = clientGetIndependenceQAAnalytics(granularity, periodValue, selectedAgent, "");
          tpl.independenceAnalytics = JSON.stringify(independenceAnalytics).replace(/<\/script>/g, '<\\/script>');
        } else {
          tpl.independenceAnalytics = JSON.stringify(getEmptyQAAnalytics());
        }
      } catch (error) {
        console.error('Error getting Independence analytics:', error);
        tpl.independenceAnalytics = JSON.stringify(getEmptyQAAnalytics());
      }

      try {
        if (typeof clientGetCreditSuiteQAAnalytics === 'function') {
          const creditSuiteAnalytics = clientGetCreditSuiteQAAnalytics(granularity, periodValue, selectedAgent, "");
          tpl.creditSuiteAnalytics = JSON.stringify(creditSuiteAnalytics).replace(/<\/script>/g, '<\\/script>');
        } else {
          tpl.creditSuiteAnalytics = JSON.stringify(getEmptyQAAnalytics());
        }
      } catch (error) {
        console.error('Error getting Credit Suite analytics:', error);
        tpl.creditSuiteAnalytics = JSON.stringify(getEmptyQAAnalytics());
      }

      tpl.qaAnalytics = tpl.independenceAnalytics;

    } else if (templateName.includes('CreditSuite')) {
      tpl.campaignName = 'Credit Suite';
      try {
        if (typeof clientGetCreditSuiteQAAnalytics === 'function') {
          const analytics = clientGetCreditSuiteQAAnalytics(granularity, periodValue, selectedAgent);
          tpl.qaAnalytics = JSON.stringify(analytics).replace(/<\/script>/g, '<\\/script>');
        } else {
          tpl.qaAnalytics = JSON.stringify(getEmptyQAAnalytics());
        }
      } catch (error) {
        console.error('Error getting Credit Suite QA analytics:', error);
        tpl.qaAnalytics = JSON.stringify(getEmptyQAAnalytics());
      }

    } else if (templateName.includes('Independence')) {
      tpl.campaignName = 'Independence Insurance';
      try {
        if (typeof clientGetIndependenceQAAnalytics === 'function') {
          const analytics = clientGetIndependenceQAAnalytics(granularity, periodValue, selectedAgent);
          tpl.qaAnalytics = JSON.stringify(analytics).replace(/<\/script>/g, '<\\/script>');
        } else {
          tpl.qaAnalytics = JSON.stringify(getEmptyQAAnalytics());
        }
      } catch (error) {
        console.error('Error getting Independence QA analytics:', error);
        tpl.qaAnalytics = JSON.stringify(getEmptyQAAnalytics());
      }

    } else {
      // Default QA dashboard
      try {
        if (typeof getAllQA === 'function') {
          const qaRecords = getAllQA();
          tpl.qaRecords = JSON.stringify(qaRecords).replace(/<\/script>/g, '<\\/script>');
        } else {
          tpl.qaRecords = JSON.stringify([]);
        }
      } catch (error) {
        console.error('Error getting QA dashboard data:', error);
        tpl.qaRecords = JSON.stringify([]);
      }
    }

  } catch (error) {
    console.error('Error handling QA dashboard data:', error);
    tpl.userList = [];
    tpl.qaAnalytics = JSON.stringify(getEmptyQAAnalytics());
  }
}

function handleQAListData(tpl, e, user, templateName) {
  try {
    if (templateName.includes('Independence')) {
      let qaRecords = [];
      if (typeof clientGetIndependenceQARecords === 'function') {
        qaRecords = clientGetIndependenceQARecords();
      }
      tpl.qaRecords = JSON.stringify(qaRecords).replace(/<\/script>/g, '<\\/script>');
      tpl.campaignName = 'Independence Insurance';

    } else if (templateName.includes('CreditSuite')) {
      let qaRecords = [];
      if (typeof clientGetCreditSuiteQARecords === 'function') {
        qaRecords = clientGetCreditSuiteQARecords();
      }
      tpl.qaRecords = JSON.stringify(qaRecords).replace(/<\/script>/g, '<\\/script>');
      tpl.campaignName = 'Credit Suite';

    } else {
      // Default or other campaign QA list
      if (typeof getAllQA === 'function') {
        tpl.qaRecords = JSON.stringify(getAllQA()).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.qaRecords = JSON.stringify([]);
      }
    }
  } catch (error) {
    console.error('Error handling QA list data:', error);
    tpl.qaRecords = JSON.stringify([]);
  }
}

function handleUsersData(tpl, e, user) {
  try {
    tpl.campaignList = getAllCampaigns();
    const roleList = getAllRoles() || [];
    tpl.roleList = roleList.filter(r => {
      const n = String(r.name || r.Name || '').toLowerCase();
      return n !== 'super admin' && n !== 'administrator';
    });
    const knownPages = Object.keys(_loadAllPagesMeta());
    tpl.pagesList = knownPages;
  } catch (error) {
    writeError('handleUsersData', error);
    tpl.campaignList = [];
    tpl.roleList = [];
    tpl.pagesList = [];
  }
}

function handleTaskData(tpl, e, user, templateName) {
  try {
    const currentUserEmail = user.email || user.Email || '';

    if (templateName === 'TaskForm') {
      tpl.currentUser = currentUserEmail;
      tpl.users = getUsers().map(u => u.email || u.Email).filter(e => e);

      if (typeof getTasksFor === 'function') {
        const visible = getTasksFor(currentUserEmail);
        tpl.tasksJson = JSON.stringify(visible).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.tasksJson = JSON.stringify([]);
      }

      const taskId = (e.parameter.id || '').toString();
      tpl.recordId = taskId;

      if (taskId && typeof getTaskById === 'function') {
        tpl.record = JSON.stringify(getTaskById(taskId)).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.record = '{}';
      }

    } else if (templateName === 'TaskList') {
      tpl.currentUser = currentUserEmail;

      if (typeof getTasksFor === 'function') {
        tpl.tasks = JSON.stringify(getTasksFor(currentUserEmail)).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.tasks = JSON.stringify([]);
      }

      tpl.users = JSON.stringify(getUsers().map(u => u.email || u.Email).filter(e => e));
      tpl.recordId = '';
      tpl.record = '{}';

    } else if (templateName === 'TaskBoard') {
      tpl.currentUser = currentUserEmail;

      if (typeof getTasksFor === 'function') {
        const visibleTasks = getTasksFor(currentUserEmail);
        tpl.tasksJson = JSON.stringify(visibleTasks).replace(/<\/script>/g, '<\\/script>');

        const ownersArr = [...new Set(visibleTasks.map(t => t.Owner).filter(x => x))].sort();
        tpl.ownersJson = JSON.stringify(ownersArr);
        tpl.selectedOwner = e.parameter.owner || '';
      } else {
        tpl.tasksJson = JSON.stringify([]);
        tpl.ownersJson = JSON.stringify([]);
        tpl.selectedOwner = '';
      }
    }

  } catch (error) {
    console.error('Error handling task data:', error);
    tpl.currentUser = user.email || user.Email || '';
    tpl.users = [];
    tpl.tasksJson = JSON.stringify([]);
    tpl.recordId = '';
    tpl.record = '{}';
  }
}

function handleCoachingData(tpl, e, user, templateName) {
  try {
    const granularity = e.parameter.granularity || "Week";
    const periodValue = e.parameter.period || weekStringFromDate(new Date());
    const selectedAgent = e.parameter.agent || "";

    if (templateName === 'CoachingForm') {
      // No additional data needed beyond base template data

    } else if (templateName === 'CoachingList') {
      if (typeof getAllCoaching === 'function') {
        tpl.coachingRecords = JSON.stringify(getAllCoaching()).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.coachingRecords = JSON.stringify([]);
      }

    } else if (templateName === 'CoachingView') {
      const coachingId = (e.parameter.id || '').toString();
      tpl.recordId = coachingId;

      if (coachingId && typeof getCoachingRecordById === 'function') {
        const coachingRec = getCoachingRecordById(coachingId);
        tpl.record = JSON.stringify(coachingRec || {}).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.record = '{}';
      }

    } else if (templateName === 'CoachingDashboard') {
      if (typeof getDashboardCoaching === 'function') {
        const coachingMetrics = getDashboardCoaching(granularity, periodValue, selectedAgent);
        const coachingUserList = getUsers().map(u => u.name || u.FullName || u.UserName).filter(n => n).sort();

        tpl.userList = coachingUserList;
        tpl.sessionsTrend = JSON.stringify(coachingMetrics.sessionsTrend || []);
        tpl.topicsDist = JSON.stringify(coachingMetrics.topicsDist || []);
        tpl.upcomingFollow = JSON.stringify(coachingMetrics.upcoming || []);
      } else {
        tpl.userList = [];
        tpl.sessionsTrend = JSON.stringify([]);
        tpl.topicsDist = JSON.stringify([]);
        tpl.upcomingFollow = JSON.stringify([]);
      }
    }

  } catch (error) {
    console.error('Error handling coaching data:', error);
    tpl.userList = [];
    tpl.sessionsTrend = JSON.stringify([]);
    tpl.topicsDist = JSON.stringify([]);
    tpl.upcomingFollow = JSON.stringify([]);
  }
}

function handleEODReportData(tpl, e, user) {
  try {
    const today = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'
    );

    if (typeof getEODTasksByDate === 'function') {
      const allComplete = getEODTasksByDate(today);
      const me = (user.email || user.Email || '').toLowerCase();
      const visibleEOD = allComplete.filter(t =>
        String(t.Owner || '').toLowerCase() === me ||
        String(t.Delegations || '')
          .split(',')
          .map(x => x.trim().toLowerCase())
          .includes(me)
      );

      tpl.reportDate = today;
      tpl.tasks = JSON.stringify(visibleEOD).replace(/<\/script>/g, '<\\/script>');
    } else {
      tpl.reportDate = today;
      tpl.tasks = JSON.stringify([]);
    }
  } catch (error) {
    console.error('Error handling EOD report data:', error);
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    tpl.reportDate = today;
    tpl.tasks = JSON.stringify([]);
  }
}

function handleCalendarData(tpl, e, user, campaignId) {
  try {
    const selectedAgent = e.parameter.agent || "";

    if (e.parameter.page === 'attendancecalendar') {
      if (typeof getAttendanceAnalyticsByPeriod === 'function') {
        const today = new Date();
        const isoWeek = weekStringFromDate(today);
        const attendanceAnalytics = getAttendanceAnalyticsByPeriod("Week", isoWeek, selectedAgent);

        tpl.userList = clientGetAssignedAgentNames(campaignId || '');
        tpl.attendanceData = JSON.stringify(attendanceAnalytics).replace(/<\/script>/g, '<\\/script>');
      } else {
        tpl.userList = clientGetAssignedAgentNames(campaignId || '');
        tpl.attendanceData = JSON.stringify({});
      }
    } else {
      tpl.userList = clientGetAssignedAgentNames(campaignId || '');
    }
  } catch (error) {
    console.error('Error handling calendar data:', error);
    tpl.userList = [];
    tpl.attendanceData = JSON.stringify({});
  }
}

function handleSearchData(tpl, e) {
  try {
    const query = (e.parameter.query || '').trim();
    const pageIndex = parseInt(e.parameter.pageIndex || '1', 10);
    const startIdx = (pageIndex - 1) * 10 + 1;
    let items = [], totalResults = 0, error = null;

    if (query) {
      try {
        if (typeof searchWeb === 'function') {
          const resp = searchWeb(query, startIdx);
          items = resp.items || [];
          totalResults = parseInt(resp.searchInformation.totalResults || '0', 10);

          items = items.map(item => ({
            ...item,
            proxyUrl: `${SCRIPT_URL}?page=proxy&url=${encodeURIComponent(item.link)}`,
            displayTitle: item.htmlTitle || item.title || 'Untitled',
            displaySnippet: item.htmlSnippet || item.snippet || 'No description available'
          }));
        }
      } catch (err) {
        error = err.message;
      }
    }

    tpl.query = query;
    tpl.results = items;
    tpl.totalResults = totalResults;
    tpl.pageIndex = pageIndex;
    tpl.error = error;
    tpl.scriptUrl = SCRIPT_URL;

  } catch (error) {
    console.error('Error handling search data:', error);
    tpl.query = '';
    tpl.results = [];
    tpl.totalResults = 0;
    tpl.pageIndex = 1;
    tpl.error = error.message;
    tpl.scriptUrl = SCRIPT_URL;
  }
}

function handleChatData(tpl, e, user) {
  try {
    const groupId = e.parameter.groupId || '';
    const channelId = e.parameter.channelId || '';
    tpl.groupId = groupId;
    tpl.channelId = channelId;

    let userGroups = [];
    let groupChannels = [];
    let messages = [];

    try {
      if (typeof getUserChatGroups === 'function') {
        userGroups = getUserChatGroups(user.ID) || [];
      }
    } catch (e) {
      console.warn('Error getting user groups:', e);
    }

    try {
      if (groupId && typeof getChatChannels === 'function') {
        groupChannels = getChatChannels(groupId);
      }
    } catch (e) {
      console.warn('Error getting group channels:', e);
    }

    try {
      if (channelId && typeof getChatMessages === 'function') {
        messages = getChatMessages(channelId);
      }
    } catch (e) {
      console.warn('Error getting messages:', e);
    }

    tpl.userGroups = JSON.stringify(userGroups).replace(/<\/script>/g, '<\\/script>');
    tpl.groupChannels = JSON.stringify(groupChannels).replace(/<\/script>/g, '<\\/script>');
    tpl.messages = JSON.stringify(messages).replace(/<\/script>/g, '<\\/script>');
  } catch (error) {
    console.error('Error handling chat data:', error);
    tpl.groupId = '';
    tpl.channelId = '';
    tpl.userGroups = JSON.stringify([]);
    tpl.groupChannels = JSON.stringify([]);
    tpl.messages = JSON.stringify([]);
  }
}

function handleNotificationsData(tpl, e, user) {
  try {
    if (typeof getAllTasks === 'function') {
      const notifToday = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const dueTasks = getAllTasks()
        .filter(t => (t.Status !== 'Done' && t.Status !== 'Completed') && t.DueDate === notifToday);
      tpl.tasksJson = JSON.stringify(dueTasks).replace(/<\/script>/g, '<\\/script>');
    } else {
      tpl.tasksJson = JSON.stringify([]);
    }
  } catch (error) {
    console.error('Error handling notifications data:', error);
    tpl.tasksJson = JSON.stringify([]);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// USER MANAGEMENT FUNCTIONS (Updated for Cookie Auth)
// ───────────────────────────────────────────────────────────────────────────────

function _normStr_(s) { return String(s || '').trim(); }
function _normEmail_(s) { return _normStr_(s).toLowerCase(); }
function _toBool_(v) { return v === true || String(v || '').trim().toUpperCase() === 'TRUE'; }

function _tryGetRoleNames_(userId) {
  try {
    if (typeof getUserRoleNamesSafe === 'function') {
      var list = getUserRoleNamesSafe(userId) || [];
      return Array.isArray(list) ? list : [];
    }
  } catch (_) { }
  return [];
}

function _campaignNameMap_() {
  try {
    const rows = (typeof readSheet === 'function') ? (readSheet('CAMPAIGNS') || []) : [];
    const map = {};
    rows.forEach(function (r) {
      const id = _normStr_(r.ID || r.Id || r.id);
      const nm = _normStr_(r.Name || r.name);
      if (id) map[id] = nm;
    });
    return map;
  } catch (_) {
    return {};
  }
}

function _uiUserShape_(u, cmap) {
  const cid = _normStr_(u.CampaignID || u.campaignId);
  return {
    ID: u.ID,
    UserName: u.UserName,
    FullName: u.FullName,
    Email: u.Email,
    CampaignID: u.CampaignID,
    campaignName: cmap[cid] || _normStr_(u.CampaignName || ''),
    CanLogin: u.CanLogin,
    IsAdmin: u.IsAdmin,
    canLoginBool: _toBool_(u.CanLogin),
    isAdminBool: _toBool_(u.IsAdmin),
    roleNames: _tryGetRoleNames_(u.ID),
    pages: []
  };
}

function getUsers() {
  try {
    console.log('getUsers() called');

    const meEmail = _normEmail_((Session.getActiveUser() && Session.getActiveUser().getEmail()) || '');
    console.log('Current user email:', meEmail);

    if (!meEmail) {
      console.warn('No current user email found');
      return [];
    }

    const allUsers = (typeof readSheet === 'function') ? (readSheet('Users') || []) : [];
    console.log('All users from sheet:', allUsers.length);

    if (allUsers.length === 0) {
      console.warn('No users found in Users sheet');
      return [];
    }

    const me = allUsers.find(function (u) {
      return _normEmail_(u.Email || u.email) === meEmail;
    });

    if (!me) {
      console.warn('Current user not found in Users sheet');
      return allUsers.map(u => _uiUserShape_(u, _campaignNameMap_()));
    }

    console.log('Found current user:', me.FullName || me.UserName);

    let muRows = [];
    try {
      muRows = (typeof readSheet === 'function') ? (readSheet('MANAGER_USERS') || []) : [];
      console.log('Manager-user relationships found:', muRows.length);
    } catch (error) {
      console.warn('Could not read manager-user relationships, will return all users:', error);
      return allUsers.map(u => _uiUserShape_(u, _campaignNameMap_()));
    }

    const assignedIds = new Set(
      muRows.filter(function (a) { return String(a.ManagerUserID) === String(me.ID); })
        .map(function (a) { return String(a.UserID); })
        .filter(Boolean)
    );

    console.log('Assigned user IDs:', Array.from(assignedIds));

    if (assignedIds.size === 0) {
      console.log('No assigned users found, using campaign-based filtering');
      const sameCampaignUsers = allUsers.filter(u =>
        (u.CampaignID || u.campaignId) === (me.CampaignID || me.campaignId)
      );
      return sameCampaignUsers.map(u => _uiUserShape_(u, _campaignNameMap_()));
    }

    const byId = new Map(allUsers.filter(function (u) { return u && u.ID; }).map(function (u) { return [String(u.ID), u]; }));
    const cmap = _campaignNameMap_();

    const out = [_uiUserShape_(me, cmap)];
    assignedIds.forEach(function (id) {
      const u = byId.get(id);
      if (u) out.push(_uiUserShape_(u, cmap));
    });

    const seen = new Set();
    const dedup = out.filter(function (u) {
      const k = String(u.ID || '');
      if (!k || seen.has(k)) return false;
      seen.add(k);
      return true;
    });

    dedup.sort(function (a, b) {
      return String(a.FullName || '').localeCompare(String(b.FullName || '')) ||
        String(a.UserName || '').localeCompare(String(b.UserName || ''));
    });

    console.log('Final user list:', dedup.length, 'users');
    return dedup;

  } catch (e) {
    console.error('Error in getUsers:', e);
    writeError && writeError('getUsers (enhanced)', e);

    try {
      const currentUser = getCurrentUser();
      if (currentUser) {
        return [currentUser];
      }
    } catch (fallbackError) {
      console.error('Fallback getCurrentUser also failed:', fallbackError);
    }

    return [];
  }
}

function clientGetAssignedAgentNames(campaignId) {
  try {
    console.log('clientGetAssignedAgentNames called with campaignId:', campaignId);

    let users = [];
    try {
      users = getUsers();
      console.log('getUsers() returned:', users.length, 'users');
    } catch (error) {
      console.warn('getUsers() failed, trying fallback approaches:', error);
    }

    if (!users || users.length === 0) {
      try {
        console.log('Trying direct Users sheet read...');
        const allUsers = readSheet('Users') || [];
        console.log('Direct Users sheet read returned:', allUsers.length, 'users');

        if (campaignId && campaignId.trim() !== '') {
          users = allUsers.filter(u =>
            String(u.CampaignID || u.campaignId || '').trim() === String(campaignId).trim()
          );
          console.log('Filtered by campaignId:', users.length, 'users');
        } else {
          users = allUsers;
        }

        users = users.map(u => ({
          ID: u.ID || u.id,
          FullName: u.FullName || u.UserName || u.name,
          UserName: u.UserName || u.name,
          Email: u.Email || u.email,
          CampaignID: u.CampaignID || u.campaignId
        }));
      } catch (error) {
        console.error('Direct Users sheet read failed:', error);
      }
    }

    if (!users || users.length === 0) {
      try {
        console.log('No users found, trying to get current user...');
        const currentUser = getCurrentUser();
        if (currentUser && currentUser.FullName) {
          users = [currentUser];
          console.log('Using current user as fallback:', currentUser.FullName);
        }
      } catch (error) {
        console.error('getCurrentUser failed:', error);
      }
    }

    if (!users || users.length === 0) {
      console.warn('All user retrieval strategies failed, using empty array');
      users = [];
    }

    const displayNames = users
      .map(u => {
        const name = u.FullName || u.UserName || u.Email || u.name || '';
        return name.trim();
      })
      .filter(name => name !== '')
      .filter((name, index, arr) => arr.indexOf(name) === index)
      .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));

    console.log('Final userList:', displayNames);
    return displayNames;

  } catch (error) {
    console.error('Error in clientGetAssignedAgentNames:', error);
    writeError('clientGetAssignedAgentNames', error);
    return [];
  }
}

function clientGetUserList(campaignId) {
  try {
    let userList = clientGetAssignedAgentNames(campaignId);

    if (!userList || userList.length === 0) {
      console.log('Assigned approach failed, trying simple approach');
      userList = getSimpleUserList(campaignId);
    }

    if (!userList || userList.length === 0) {
      console.log('Campaign-filtered approach failed, trying all users');
      userList = getSimpleUserList('');
    }

    return userList;

  } catch (error) {
    console.error('Error in clientGetUserList:', error);
    return [];
  }
}

function getSimpleUserList(campaignId) {
  try {
    console.log('getSimpleUserList called with campaignId:', campaignId);

    const users = readSheet('Users') || [];
    console.log('Read', users.length, 'users from sheet');

    let filteredUsers = users;

    if (campaignId && campaignId.trim() !== '') {
      filteredUsers = users.filter(u =>
        String(u.CampaignID || u.campaignId || '').trim() === String(campaignId).trim()
      );
      console.log('Filtered to', filteredUsers.length, 'users for campaign:', campaignId);
    }

    const names = filteredUsers
      .map(u => u.FullName || u.UserName || u.Email || u.name || '')
      .filter(name => name.trim() !== '')
      .filter((name, index, arr) => arr.indexOf(name) === index)
      .sort();

    console.log('Returning', names.length, 'user names');
    return names;

  } catch (error) {
    console.error('Error in getSimpleUserList:', error);
    return [];
  }
}

function getUsersByCampaign(campaignId) {
  try {
    var rows = (typeof readSheet === 'function') ? (readSheet('Users') || []) : [];
    return rows.filter(function (r) { return String(r.CampaignID || r.campaignId) === String(campaignId); });
  } catch (_) {
    return [];
  }
}

function getAllUsers() {
  try {
    if (typeof readSheet === 'function') {
      return readSheet('Users') || [];
    }
    return [];
  } catch (error) {
    console.error('Error getting all users:', error);
    return [];
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// CAMPAIGN-SPECIFIC USER FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

function clientGetIndependenceUsers() {
  try {
    return getIndependenceInsuranceUsers();
  } catch (error) {
    console.error('Error in clientGetIndependenceUsers:', error);
    writeError('clientGetIndependenceUsers', error);
    return [];
  }
}

function getIndependenceInsuranceUsers() {
  try {
    const allUsers = getAllUsers();
    const users = allUsers.filter(user => {
      const campaignName = user.CampaignName || user.campaignName || '';
      const userCampaignId = user.CampaignID || user.campaignId || '';

      return campaignName.toLowerCase().includes('independence') ||
        campaignName.toLowerCase().includes('insurance') ||
        userCampaignId === 'independence-insurance-agency' ||
        userCampaignId.toLowerCase().includes('independence');
    });

    return users
      .map(user => ({
        name: user.FullName || user.UserName || user.name,
        email: user.Email || user.email,
        id: user.ID || user.id
      }))
      .filter(user => user.name && user.email)
      .sort((a, b) => a.name.localeCompare(b.name));

  } catch (error) {
    console.error('Error getting Independence users:', error);
    writeError('getIndependenceInsuranceUsers', error);
    return [];
  }
}

function clientGetCreditSuiteUsers(campaignId = null) {
  try {
    return getCreditSuiteUsers(campaignId);
  } catch (error) {
    console.error('Error in clientGetCreditSuiteUsers:', error);
    writeError('clientGetCreditSuiteUsers', error);
    return [];
  }
}

function getCreditSuiteUsers(campaignId = null) {
  try {
    let users;

    if (campaignId) {
      users = getUsersByCampaign(campaignId);
    } else {
      const allUsers = getAllUsers();
      users = allUsers.filter(user => {
        const campaignName = user.CampaignName || user.campaignName || '';
        const userCampaignId = user.CampaignID || user.campaignId || '';

        return campaignName.toLowerCase().includes('credit') ||
          campaignName.toLowerCase().includes('suite') ||
          userCampaignId === 'credit-suite' ||
          userCampaignId.toLowerCase().includes('credit');
      });
    }

    return users
      .map(user => ({
        name: user.FullName || user.UserName || user.name,
        email: user.Email || user.email,
        id: user.ID || user.id
      }))
      .filter(user => user.name && user.email)
      .sort((a, b) => a.name.localeCompare(b.name));

  } catch (error) {
    console.error('Error getting Credit Suite users:', error);
    writeError('getCreditSuiteUsers', error);
    return [];
  }
}

function clientGetIBTRUsers() {
  try {
    const allUsers = getAllUsers();
    const users = allUsers.filter(user => {
      const campaignName = user.CampaignName || user.campaignName || '';
      const userCampaignId = user.CampaignID || user.campaignId || '';

      return campaignName.toLowerCase().includes('benefits resource center') ||
        campaignName.toLowerCase().includes('ibtr') ||
        userCampaignId === 'ibtr' ||
        userCampaignId.toLowerCase().includes('benefits');
    });

    return users
      .map(user => ({
        name: user.FullName || user.UserName || user.name,
        email: user.Email || user.email,
        id: user.ID || user.id
      }))
      .filter(user => user.name && user.email)
      .sort((a, b) => a.name.localeCompare(b.name));

  } catch (error) {
    console.error('Error getting IBTR users:', error);
    writeError('clientGetIBTRUsers', error);
    return [];
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// CAMPAIGN AND NAVIGATION FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

function getAllCampaigns() {
  try {
    var campaigns;
    if (typeof clientGetAllCampaigns === 'function') {
      campaigns = clientGetAllCampaigns();
    } else if (typeof readSheet === 'function') {
      campaigns = readSheet('CAMPAIGNS') || [];
    } else {
      campaigns = [];
    }

    var currentUser = null;
    try {
      if (typeof getCurrentUser === 'function') {
        currentUser = getCurrentUser();
      }
    } catch (err) {
      console.warn('getAllCampaigns: failed to hydrate current user', err);
    }

    if (currentUser && currentUser.ID && typeof TenantSecurity !== 'undefined' && TenantSecurity) {
      try {
        return TenantSecurity.filterCampaignList(currentUser.ID, campaigns);
      } catch (filterErr) {
        console.warn('getAllCampaigns: tenant filter error', filterErr);
      }
    }

    return campaigns;
  } catch (error) {
    console.error('Error getting all campaigns:', error);
    return [];
  }
}

function getCampaignById(campaignId) {
  try {
    const campaigns = getAllCampaigns();
    return campaigns.find(campaign =>
      String(campaign.ID || campaign.id) === String(campaignId)
    );
  } catch (error) {
    console.error('Error getting campaign by ID:', error);
    writeError('getCampaignById', error);
    return null;
  }
}

function getCampaignNavigation(campaignId) {
  try {
    if (typeof readSheet === 'function') {
      const navigation = readSheet('CampaignNavigation') || [];
      const campaignNav = navigation.filter(nav =>
        String(nav.CampaignID || nav.campaignId) === String(campaignId)
      );

      const categories = {};
      const uncategorizedPages = [];

      campaignNav.forEach(nav => {
        const category = nav.Category || nav.category || 'Uncategorized';
        if (category === 'Uncategorized') {
          uncategorizedPages.push(nav);
        } else {
          if (!categories[category]) {
            categories[category] = [];
          }
          categories[category].push(nav);
        }
      });

      return {
        categories: Object.entries(categories).map(([name, pages]) => ({
          name,
          pages
        })),
        uncategorizedPages
      };
    }

    return {
      categories: [
        {
          name: 'Main',
          pages: [
            { PageKey: 'dashboard', Name: 'Dashboard', Icon: 'dashboard' },
            { PageKey: 'reports', Name: 'Reports', Icon: 'bar_chart' }
          ]
        }
      ],
      uncategorizedPages: []
    };
  } catch (error) {
    console.error('Error getting campaign navigation:', error);
    writeError('getCampaignNavigation', error);
    return { categories: [], uncategorizedPages: [] };
  }
}

function getUserCampaignPermissions(userId) {
  try {
    if (typeof readSheet === 'function') {
      const permissions = readSheet('CAMPAIGN_USER_PERMISSIONS') || [];
      return permissions.filter(perm =>
        String(perm.UserID || perm.userId) === String(userId)
      );
    }
    return [];
  } catch (error) {
    console.error('Error getting user campaign permissions:', error);
    writeError('getUserCampaignPermissions', error);
    return [];
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// ROLES AND PERMISSIONS
// ───────────────────────────────────────────────────────────────────────────────

function getAllRoles() {
  try {
    if (typeof readSheet === 'function') {
      return readSheet('Roles') || [];
    } else {
      return [];
    }
  } catch (error) {
    console.error('Error getting all roles:', error);
    return [];
  }
}

function getRolesMapping() {
  try {
    let roles = [];
    if (typeof getAllRoles === 'function') {
      roles = getAllRoles();
    } else if (typeof readSheet === 'function') {
      roles = readSheet('Roles') || [];
    }

    const mapping = {};
    roles.forEach(role => {
      mapping[role.ID || role.id] = role.Name || role.name;
    });
    return mapping;
  } catch (error) {
    console.error('Error getting roles mapping:', error);
    return {};
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// PAGE ACCESS AND METADATA
// ───────────────────────────────────────────────────────────────────────────────

function _loadAllPagesMeta() {
  const CK = 'ACCESS_PAGES_META_V1';
  const cached = _cacheGet(CK);
  if (cached) return cached;

  let rows = [];
  try {
    rows = readSheet('Pages') || [];
  } catch (_) {
  }

  const map = {};
  rows.forEach(r => {
    const key = _normalizePageKey(r.PageKey || r.key);
    if (!key) return;
    map[key] = {
      key,
      title: r.PageTitle || r.Name || '',
      requiresAdmin: _truthy(r.RequiresAdmin),
      campaignSpecific: _truthy(r.CampaignSpecific) || _truthy(r.IsCampaignSpecific) || false,
      active: (r.Active === undefined) ? true : _truthy(r.Active)
    };
  });
  _cachePut(CK, map, 180);
  return map;
}

function _getPageMeta(pageKey) {
  const key = _normalizePageKey(pageKey);
  const meta = _loadAllPagesMeta()[key];
  return meta || { key, title: pageKey, requiresAdmin: false, campaignSpecific: false, active: true };
}

// ───────────────────────────────────────────────────────────────────────────────
// SPECIAL PAGE HANDLERS (Updated for Clean URLs)
// ───────────────────────────────────────────────────────────────────────────────

function serveAckForm(e, baseUrl, user) {
  try {
    const id = (e.parameter.id || "").toString();
    const tpl = HtmlService.createTemplateFromFile('CoachingAckForm');

    tpl.baseUrl = baseUrl;
    tpl.scriptUrl = SCRIPT_URL;
    tpl.currentPage = 'Acknowledge Coaching';
    tpl.recordId = id;

    if (id && typeof getCoachingRecordById === 'function') {
      tpl.record = JSON.stringify(getCoachingRecordById(id)).replace(/<\/script>/g, '<\\/script>');
    } else {
      tpl.record = '{}';
    }

    tpl.user = user;
    __injectCurrentUser_(tpl, user);

    return tpl.evaluate()
      .setTitle('Acknowledge Coaching')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('Error serving ack form:', error);
    return createErrorPage('Form Error', error.message);
  }
}

function serveAgentSchedulePage(e, baseUrl, token) {
  try {
    const agentData = validateAgentToken(token);

    if (!agentData.success) {
      return createErrorPage('Invalid Token', 'Your schedule link has expired or is invalid.');
    }

    const tpl = HtmlService.createTemplateFromFile('AgentSchedule');
    tpl.baseUrl = baseUrl;
    tpl.token = token;
    tpl.agentData = JSON.stringify(agentData).replace(/<\/script>/g, '<\\/script>');

    return tpl.evaluate()
      .setTitle('My Schedule - VLBPO')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('Error serving agent schedule page:', error);
    return createErrorPage('Schedule Error', error.message);
  }
}

function serveShiftSlotManagement(e, baseUrl, user, campaignId) {
  try {
    if (!isUserAdmin(user) && !hasPermission(user, 'manage_shifts')) {
      return renderAccessDenied('You do not have permission to manage shift slots.');
    }

    const tpl = HtmlService.createTemplateFromFile('SlotManagementInterface');
    tpl.baseUrl = baseUrl;
    tpl.scriptUrl = SCRIPT_URL;
    tpl.currentPage = 'Shift Slot Management';
    tpl.user = user;
    tpl.campaignId = campaignId;

    __injectCurrentUser_(tpl, user);

    if (campaignId) {
      tpl.campaignName = user.campaignName || '';
    }

    return tpl.evaluate()
      .setTitle('Shift Slot Management - VLBPO LuminaHQ')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    console.error('Error serving shift slot management page:', error);
    writeError('serveShiftSlotManagement', error);
    return createErrorPage('Failed to load page', error.message);
  }
}

function hasPermission(user, permission) {
  try {
    const userRoles = user.Roles || user.roles || '';

    if (isUserAdmin(user)) return true;

    switch (permission) {
      case 'manage_shifts':
        return userRoles.toLowerCase().includes('manager') ||
          userRoles.toLowerCase().includes('supervisor') ||
          userRoles.toLowerCase().includes('scheduler');
      default:
        return false;
    }
  } catch (error) {
    console.error('Error checking permission:', error);
    return false;
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// ENHANCED PROXY SERVICES
// ───────────────────────────────────────────────────────────────────────────────

function serveProxy(e) {
  try {
    return serveEnhancedProxy(e);
  } catch (error) {
    console.error('Proxy service error:', error);
    writeError('serveProxy', error);
    return serveBasicProxy(e);
  }
}

function serveEnhancedProxy(e) {
  try {
    var target = e.parameter.url || '';
    if (!target) {
      return HtmlService.createHtmlOutput('<h3>Proxy Error</h3><p>Missing url</p>')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    if (!/^https?:\/\//i.test(target)) target = 'https://' + target;

    var resp = UrlFetchApp.fetch(target, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
    });
    var ct = String(resp.getHeaders()['Content-Type'] || '').toLowerCase();
    var body = resp.getContentText();

    if (ct.indexOf('text/html') === -1) {
      return ContentService.createTextOutput(body);
    }

    body = body
      .replace(/<meta[^>]+http-equiv=["']?content-security-policy["']?[^>]*>/gi, '')
      .replace(/<script[^>]*>[\s\S]*?top\.location(?:.|\s)*?<\/script>/gi, '')
      .replace(/<script[^>]*>[\s\S]*?if\s*\(\s*top\s*[!=]=?\s*self\s*\)[\s\S]*?<\/script>/gi, '');

    var base = target.replace(/([?#].*)$/, '');
    if (body.indexOf('<base ') === -1) {
      body = body.replace(/<head([^>]*)>/i, function (m, attrs) {
        return '<head' + attrs + '><base href="' + base + '">';
      });
    }

    body = body.replace('</head>', '<style>*,*:before,*:after{box-sizing:border-box}</style></head>');

    return HtmlService.createHtmlOutput(body)
      .setTitle('LuminaHQ Browser')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    writeError && writeError('serveEnhancedProxy', err);
    return HtmlService.createHtmlOutput('<h3>Proxy Error</h3><pre>' + String(err) + '</pre>')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function serveBasicProxy(e) {
  try {
    const target = e.parameter.url;
    if (!target) {
      return HtmlService.createHtmlOutput(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>Proxy Error</title>
            <style>
                body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
                .error { color: #e74c3c; }
            </style>
        </head>
        <body>
            <h1 class="error">Proxy Error</h1>
            <p>Missing URL parameter</p>
            <button onclick="history.back()">Go Back</button>
        </body>
        </html>
      `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    let normalizedUrl = target;
    if (!normalizedUrl.startsWith('http')) {
      normalizedUrl = 'https://' + normalizedUrl;
    }

    const resp = UrlFetchApp.fetch(normalizedUrl, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });

    const content = resp.getContentText();
    const contentType = resp.getHeaders()['Content-Type'] || '';

    if (contentType.toLowerCase().includes('text/html')) {
      let processedContent = content;

      processedContent = processedContent.replace(
        /<script[^>]*>[\s\S]*?if\s*\(\s*top\s*[!=]=?\s*self\s*\)[\s\S]*?<\/script>/gi,
        ''
      );
      processedContent = processedContent.replace(
        /<script[^>]*>[\s\S]*?top\.location[\s\S]*?<\/script>/gi,
        ''
      );

      const enhancedContent = processedContent.replace(
        '</head>',
        `<style>
          body { margin-top: 0 !important; }
          * { box-sizing: border-box; }
        </style></head>`
      );

      return HtmlService.createHtmlOutput(enhancedContent)
        .setTitle('LuminaHQ Browser')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } else {
      return ContentService.createTextOutput(content);
    }

  } catch (error) {
    console.error('Basic proxy error:', error);
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
      <head>
          <title>Proxy Error</title>
          <style>
              body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
              .error { color: #e74c3c; }
          </style>
      </head>
      <body>
          <h1 class="error">Unable to Load Page</h1>
          <p>Error: ${error.message}</p>
          <button onclick="history.back()">Go Back</button>
          <button onclick="location.reload()">Retry</button>
      </body>
      </html>
    `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// UTILITY FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

function createErrorPage(title, message) {
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <title>${title}</title>
      <meta name="viewport" content="width=device-width,initial-scale=1">
      <style>
        body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
        .error { color: #e74c3c; }
        .message { margin: 20px 0; }
        .back-link { color: #3498db; text-decoration: none; }
      </style>
    </head>
    <body>
      <h1 class="error">${title}</h1>
      <p class="message">${message}</p>
      <a href="#" onclick="history.back()" class="back-link">← Go Back</a>
    </body>
    </html>
  `;

  return HtmlService.createHtmlOutput(html)
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function writeError(functionName, error) {
  try {
    const errorMessage = error && error.message ? error.message : error.toString();
    console.error(`[${functionName}] ${errorMessage}`);

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let errorSheet = ss.getSheetByName('ErrorLog');

      if (!errorSheet) {
        errorSheet = ss.insertSheet('ErrorLog');
        errorSheet.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Function', 'Error', 'Stack']]);
      }

      const timestamp = new Date();
      const stack = error && error.stack ? error.stack : 'No stack trace';

      errorSheet.appendRow([timestamp, functionName, errorMessage, stack]);
    } catch (logError) {
      console.error('Failed to log error to sheet:', logError);
    }
  } catch (e) {
    console.error('Error in writeError function:', e);
  }
}

function writeDebug(message) {
  console.log(`[DEBUG] ${message}`);
}

function confirmEmail(token) {
  try {
    if (!token) {
      console.warn('confirmEmail called with empty token');
      return false;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Users');
    if (!sh) {
      console.error('Users sheet not found');
      return false;
    }

    const data = sh.getDataRange().getValues();
    if (data.length < 2) {
      console.warn('No users found in Users sheet');
      return false;
    }

    const headers = data.shift();
    const colTok = headers.indexOf('EmailConfirmation');
    const colConf = headers.indexOf('EmailConfirmed');
    const colUpd = headers.indexOf('UpdatedAt');
    const colEmail = headers.indexOf('Email');

    if (colTok < 0 || colConf < 0) {
      console.error('Required columns not found in Users sheet');
      return false;
    }

    let found = false;
    let userEmail = null;

    data.forEach((row, i) => {
      if (String(row[colTok]) === String(token)) {
        found = true;
        userEmail = row[colEmail];
        const rowNum = i + 2;

        sh.getRange(rowNum, colConf + 1).setValue(true);

        if (colUpd >= 0) {
          sh.getRange(rowNum, colUpd + 1).setValue(new Date());
        }

        console.log(`Email confirmed for user: ${userEmail}`);
      }
    });

    return found;
  } catch (error) {
    console.error('Error confirming email:', error);
    writeError('confirmEmail', error);
    return false;
  }
}

function weekStringFromDate(date) {
  try {
    if (!date || !(date instanceof Date)) {
      date = new Date();
    }

    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    const weekNum = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    return `${d.getUTCFullYear()}-W${weekNum.toString().padStart(2, '0')}`;
  } catch (error) {
    console.error('Error generating week string:', error);
    return new Date().getFullYear() + '-W01';
  }
}

function searchWeb(query, startIndex) {
  try {
    if (!query || typeof query !== 'string') {
      throw new Error('Invalid search query.');
    }
    const CSE_ID = '130aba31c8a2d439c';
    const API_KEY = 'AIzaSyAg-puM5l9iQpjz_NplMJaKbUNRH7ld7sY';
    const baseUrl = 'https://www.googleapis.com/customsearch/v1';
    const params = [
      `key=${API_KEY}`,
      `cx=${CSE_ID}`,
      `q=${encodeURIComponent(query)}`,
      startIndex ? `start=${startIndex}` : ''
    ].filter(Boolean).join('&');
    const url = `${baseUrl}?${params}`;

    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = resp.getResponseCode();
    const text = resp.getContentText();
    if (code !== 200) {
      throw new Error(`Search API error [${code}]: ${text}`);
    }
    return JSON.parse(text);
  } catch (error) {
    console.error('Error in searchWeb:', error);
    writeError('searchWeb', error);
    throw error;
  }
}

function getEmptyQAAnalytics() {
  return {
    avgScore: 0,
    passRate: 0,
    totalEvaluations: 0,
    agentsEvaluated: 0,
    avgScoreChange: 0,
    passRateChange: 0,
    evaluationsChange: 0,
    agentsChange: 0,
    categories: { labels: [], values: [] },
    trends: { labels: [], values: [] },
    agents: { labels: [], values: [] }
  };
}

// ───────────────────────────────────────────────────────────────────────────────
// SYSTEM INITIALIZATION
// ───────────────────────────────────────────────────────────────────────────────

function initializeSystem() {
  try {
    console.log('Initializing system...');

    if (typeof AuthenticationService !== 'undefined' && AuthenticationService.ensureSheets) {
      AuthenticationService.ensureSheets();
    }

    initializeMainSheets();
    initializeCampaignSystems();

    if (typeof CallCenterWorkflowService !== 'undefined' && CallCenterWorkflowService.initialize) {
      try {
        CallCenterWorkflowService.initialize();
      } catch (err) {
        console.warn('CallCenterWorkflowService initialization failed', err);
      }
    }

    console.log('System initialization completed successfully');
    return { success: true, message: 'System initialized' };

  } catch (error) {
    console.error('System initialization failed:', error);
    writeError('initializeSystem', error);
    return { success: false, error: error.message };
  }
}

function initializeMainSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requiredSheets = ['Users', 'Roles', 'Pages', 'CAMPAIGNS'];

    requiredSheets.forEach(sheetName => {
      try {
        let sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
          console.log(`Creating missing sheet: ${sheetName}`);
          sheet = ss.insertSheet(sheetName);
          initializeSheetHeaders(sheet, sheetName);
        }
      } catch (sheetError) {
        console.warn(`Could not initialize sheet ${sheetName}:`, sheetError);
      }
    });

    console.log('Main sheets initialized');

  } catch (error) {
    console.error('Error initializing main sheets:', error);
    writeError('initializeMainSheets', error);
  }
}

function initializeSheetHeaders(sheet, sheetName) {
  try {
    let headers = [];

    switch (sheetName) {
      case 'Users':
        headers = ['ID', 'Email', 'FullName', 'Password', 'Roles', 'CampaignID', 'EmailConfirmed', 'EmailConfirmation', 'CreatedAt', 'UpdatedAt'];
        break;
      case 'Roles':
        headers = ['ID', 'Name', 'Description', 'Permissions', 'CreatedAt'];
        break;
      case 'Pages':
        headers = ['PageKey', 'Name', 'Description', 'RequiredRole', 'CampaignSpecific', 'Active'];
        break;
      case 'CAMPAIGNS':
        headers = ['ID', 'Name', 'Description', 'Active', 'Settings', 'CreatedAt'];
        break;
      default:
        console.warn(`No headers defined for sheet: ${sheetName}`);
        return;
    }

    if (headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      console.log(`Headers set for ${sheetName}:`, headers);
    }

  } catch (error) {
    console.error(`Error setting headers for ${sheetName}:`, error);
  }
}

function initializeCampaignSystems() {
  try {
    if (typeof initializeIndependenceQASystem === 'function') {
      try {
        initializeIndependenceQASystem();
        console.log('Independence QA system initialized');
      } catch (error) {
        console.warn('Independence QA system initialization failed:', error);
      }
    }

    if (typeof initializeCreditSuiteQASystem === 'function') {
      try {
        initializeCreditSuiteQASystem();
        console.log('Credit Suite QA system initialized');
      } catch (error) {
        console.warn('Credit Suite QA system initialization failed:', error);
      }
    }

    if (typeof readSheet === 'function') {
      const pages = readSheet('Pages');
      if (pages.length === 0) {
        initializeSystemPages();
      }
    }

    console.log('Campaign systems initialization completed');

  } catch (error) {
    console.error('Error initializing campaign systems:', error);
    writeError('initializeCampaignSystems', error);
  }
}

function initializeSystemPages() {
  try {
    const systemPages = [
      {
        PageKey: 'dashboard',
        Name: 'Dashboard',
        Description: 'Main dashboard page',
        RequiredRole: '',
        CampaignSpecific: false,
        Active: true
      },
      {
        PageKey: 'users',
        Name: 'User Management',
        Description: 'Manage system users',
        RequiredRole: 'admin',
        CampaignSpecific: false,
        Active: true
      },
      {
        PageKey: 'roles',
        Name: 'Role Management',
        Description: 'Manage user roles',
        RequiredRole: 'admin',
        CampaignSpecific: false,
        Active: true
      },
      {
        PageKey: 'campaigns',
        Name: 'Campaign Management',
        Description: 'Manage campaigns',
        RequiredRole: 'admin',
        CampaignSpecific: false,
        Active: true
      },
      {
        PageKey: 'tasks',
        Name: 'Task Management',
        Description: 'Task board and management',
        RequiredRole: '',
        CampaignSpecific: true,
        Active: true
      },
      {
        PageKey: 'search',
        Name: 'Search',
        Description: 'Global search functionality',
        RequiredRole: '',
        CampaignSpecific: false,
        Active: true
      },
      {
        PageKey: 'chat',
        Name: 'Chat',
        Description: 'Team communication',
        RequiredRole: '',
        CampaignSpecific: false,
        Active: true
      },
      {
        PageKey: 'notifications',
        Name: 'Notifications',
        Description: 'System notifications',
        RequiredRole: '',
        CampaignSpecific: false,
        Active: true
      }
    ];

    if (typeof writeSheet === 'function') {
      writeSheet('Pages', systemPages);
      console.log('System pages initialized');
    }

  } catch (error) {
    console.error('Error initializing system pages:', error);
    writeError('initializeSystemPages', error);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// GLOBAL FAVICON INJECTOR AND TEMPLATE HELPERS
// ───────────────────────────────────────────────────────────────────────────────

(function () {
  if (HtmlService.__faviconPatched === true) return;
  HtmlService.__faviconPatched = true;

  const _origCreate = HtmlService.createTemplateFromFile;

  HtmlService.createTemplateFromFile = function (file) {
    const tpl = _origCreate.call(HtmlService, file);

    const _origEval = tpl.evaluate;
    tpl.evaluate = function () {
      const out = _origEval.apply(tpl, arguments);
      return (typeof out.setFaviconUrl === 'function') ? out.setFaviconUrl(FAVICON_URL) : out;
    };

    return tpl;
  };
})();

const __INCLUDE_STACK = [];
const __INCLUDED_ONCE = new Set();

function include(file, params) {
  if (__INCLUDE_STACK.indexOf(file) !== -1) {
    throw new RangeError('Cyclic include detected: ' + __INCLUDE_STACK.concat(file).join(' → '));
  }
  if (__INCLUDE_STACK.length > 20) {
    throw new RangeError('Include depth > 20; aborting to avoid stack overflow.');
  }

  __INCLUDE_STACK.push(file);
  try {
    const tpl = HtmlService.createTemplateFromFile(file);
    if (params) Object.keys(params).forEach(k => (tpl[k] = params[k]));

    try {
      if (!tpl.user) tpl.user = getCurrentUser();
      tpl.currentUserJson = _stringifyForTemplate_(tpl.user);
    } catch (_) {
      tpl.user = null;
      tpl.currentUserJson = '{}';
    }

    return tpl.evaluate().getContent();
  } finally {
    __INCLUDE_STACK.pop();
  }
}

function includeOnce(file, params) {
  if (__INCLUDED_ONCE.has(file)) return '';
  __INCLUDED_ONCE.add(file);
  return include(file, params);
}

// ───────────────────────────────────────────────────────────────────────────────
// CAMPAIGN HELPERS FOR CLIENT ACCESS
// ───────────────────────────────────────────────────────────────────────────────

function clientListCasesAndPages() {
  return CASE_DEFS.map(def => ({
    case: def.case,
    aliases: def.aliases,
    idHint: def.idHint,
    pages: Object.keys(def.pages || {})
  }));
}

function clientBuildCaseHref(caseNameOrAlias, pageKind) {
  const base = getBaseUrl();
  const page = String(caseNameOrAlias || '').trim() + '.' + String(pageKind || 'QAForm').trim();
  return base + '?page=' + encodeURIComponent(page);
}

function __routeCase__(caseName, pageKind, e, baseUrl, user, cid) {
  const def = __case_resolve__(caseName);
  if (!def) return createErrorPage('Unknown Campaign', 'No case mapping for: ' + caseName);
  const targetCid = cid || __case_findCampaignId__(def) || (user && user.CampaignID) || '';
  const candidates = __case_templateCandidates__(def, pageKind);
  return __case_serve__(candidates, e, baseUrl, user, targetCid);
}

// Named route wrappers for specific campaigns
function routeCreditSuiteQAForm(e, baseUrl, user, cid) {
  return __routeCase__('CreditSuite', 'QAForm', e, baseUrl, user, cid);
}

function routeIBTRQACollabList(e, baseUrl, user, cid) {
  return __routeCase__('IBTR', 'QACollabList', e, baseUrl, user, cid);
}

function routeTGCChatReport(e, baseUrl, user, cid) {
  return __routeCase__('TGC', 'ChatReport', e, baseUrl, user, cid);
}

function routeIBTRAttendanceReports(e, baseUrl, user, cid) {
  return __routeCase__('IBTR', 'AttendanceReports', e, baseUrl, user, cid);
}

function routeIndependenceQAForm(e, baseUrl, user, cid) {
  return __routeCase__('IndependenceInsuranceAgency', 'QAForm', e, baseUrl, user, cid);
}

function _debugAccess(label, result, user, pageKey, campaignId) {
  try {
    if (!ACCESS_DEBUG) return;
    console.log('[ACCESS] ' + label + ' → allow=' + result.allow + ' reason=' + result.reason +
      ' page=' + pageKey + ' campaign=' + (campaignId || '-') +
      ' user=' + (user && (user.Email || user.UserName || user.ID)));
    (result.trace || []).forEach((t, i) => console.log('   [' + i + '] ' + t));
  } catch (_) {
  }
}

function isUserAdmin(user) {
  return isSystemAdmin(user);
}

function validateAgentToken(token) {
  // Placeholder function - implement based on your token validation system
  try {
    // Add your token validation logic here
    return {
      success: true,
      agentId: 'sample-agent',
      agentName: 'Sample Agent'
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

// Helper function to get QA form users (unified approach)
function getQAFormUsers(campaignId, requestingUserId) {
  try {
    console.log('getQAFormUsers called with campaignId:', campaignId, 'requestingUserId:', requestingUserId);

    // Use the existing user management system
    const users = getUsers();

    return users
      .filter(user => {
        // Filter by campaign if specified
        if (campaignId && user.CampaignID && user.CampaignID !== campaignId) {
          return false;
        }
        return true;
      })
      .map(user => ({
        name: user.FullName || user.UserName,
        email: user.Email,
        id: user.ID
      }))
      .filter(user => user.name && user.email)
      .sort((a, b) => a.name.localeCompare(b.name));
  } catch (error) {
    console.error('Error in getQAFormUsers:', error);
    return [];
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// FINAL INITIALIZATION LOG
// ───────────────────────────────────────────────────────────────────────────────

console.log('Enhanced Multi-Campaign Code.gs with Simplified Authentication loaded successfully');
console.log('Features: Token-based authentication, Campaign-aware routing, Enhanced access control');
console.log('Base URL:', SCRIPT_URL);
console.log('Supported Campaigns:', CASE_DEFS.map(def => def.case).join(', '));
