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
    summary.navigation = ensureCampaignNavigationSeeds(campaignIdsByName);

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
function ensureCampaignNavigationSeeds(campaignIdsByName) {
  const result = {};

  try {
    if (!campaignIdsByName || !Object.keys(campaignIdsByName).length) {
      return result;
    }

    const pageCatalog = resolveSeedPageCatalog();
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

function ensurePasswordWithToken(token, password, options) {
  const label = options && options.label ? String(options.label) : 'user';
  const userId = options && options.userId ? String(options.userId) : '';

  if (!token) {
    return { success: false, message: 'No setup token provided for ' + label };
  }

  const directHandlers = [];

  if (typeof setPasswordWithToken === 'function') {
    directHandlers.push({
      name: 'global setPasswordWithToken',
      fn: setPasswordWithToken
    });
  }

  if (typeof AuthenticationService !== 'undefined'
    && AuthenticationService
    && typeof AuthenticationService.setPasswordWithToken === 'function') {
    directHandlers.push({
      name: 'AuthenticationService.setPasswordWithToken',
      fn: function invokeAuthService(tokenValue, passwordValue) {
        return AuthenticationService.setPasswordWithToken(tokenValue, passwordValue);
      }
    });
  }

  let lastError = null;

  for (let i = 0; i < directHandlers.length; i++) {
    const handler = directHandlers[i];
    try {
      const result = handler.fn(token, password);
      if (result && result.success) {
        return Object.assign({ via: handler.name }, result);
      }
      lastError = result || { message: handler.name + ' returned an unexpected response' };
    } catch (handlerErr) {
      lastError = { message: handlerErr && handlerErr.message ? handlerErr.message : String(handlerErr) };
      if (typeof console !== 'undefined' && console && typeof console.warn === 'function') {
        console.warn('ensurePasswordWithToken: ' + handler.name + ' failed', handlerErr);
      }
    }
  }

  const fallbackResult = setPasswordWithTokenViaSheet(token, password, { label, userId });
  if (fallbackResult && fallbackResult.success) {
    return fallbackResult;
  }

  return fallbackResult || lastError || { success: false, message: 'Unable to set password for ' + label };
}

function setPasswordWithTokenViaSheet(token, password, options) {
  const label = options && options.label ? String(options.label) : 'user';
  const normalizedToken = token ? String(token).trim() : '';
  const normalizedUserId = options && options.userId ? String(options.userId).trim() : '';

  if (!normalizedToken && !normalizedUserId) {
    return { success: false, message: 'No token or user ID available to set password for ' + label };
  }

  if (typeof SpreadsheetApp === 'undefined'
    || !SpreadsheetApp
    || typeof SpreadsheetApp.getActiveSpreadsheet !== 'function') {
    return { success: false, message: 'Spreadsheet access unavailable to set password for ' + label };
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    return { success: false, message: 'Spreadsheet not available for password update' };
  }

  const usersSheetName = (typeof USERS_SHEET !== 'undefined' && USERS_SHEET) ? USERS_SHEET : 'Users';
  const sheet = spreadsheet.getSheetByName(usersSheetName);
  if (!sheet) {
    return { success: false, message: 'Users sheet not found for password update' };
  }

  const range = sheet.getDataRange();
  if (!range) {
    return { success: false, message: 'Users sheet range unavailable for password update' };
  }

  const data = range.getValues();
  if (!data || data.length < 2) {
    return { success: false, message: 'Users sheet does not contain any data to update passwords' };
  }

  const headers = data[0].map(value => (value == null ? '' : String(value)));
  const columnIndex = {};
  headers.forEach((header, idx) => {
    const normalized = String(header || '').trim();
    if (!normalized) return;
    columnIndex[normalized] = idx;
    columnIndex[normalized.toLowerCase()] = idx;
  });

  const idIdx = columnIndex.ID != null ? columnIndex.ID : columnIndex.id;
  const tokenIdx = columnIndex.EmailConfirmation != null ? columnIndex.EmailConfirmation : columnIndex.emailconfirmation;
  const tokenHashIdx = columnIndex.EmailConfirmationTokenHash != null ? columnIndex.EmailConfirmationTokenHash : columnIndex.emailconfirmationtokenhash;
  const confirmedIdx = columnIndex.EmailConfirmed != null ? columnIndex.EmailConfirmed : columnIndex.emailconfirmed;
  const resetIdx = columnIndex.ResetRequired != null ? columnIndex.ResetRequired : columnIndex.resetrequired;
  const updatedIdx = columnIndex.UpdatedAt != null ? columnIndex.UpdatedAt : columnIndex.updatedat;

  const tokenHashes = [];
  if (tokenHashIdx != null && tokenHashIdx >= 0 && PASSWORD_UTILS && typeof PASSWORD_UTILS.createPasswordUpdate === 'function' && normalizedToken) {
    try {
      const tokenRecord = PASSWORD_UTILS.createPasswordUpdate(normalizedToken);
      if (tokenRecord) {
        if (tokenRecord.hash) {
          tokenHashes.push(String(tokenRecord.hash).trim());
        }
        if (tokenRecord.variants) {
          Object.keys(tokenRecord.variants).forEach(key => {
            const variantValue = tokenRecord.variants[key];
            if (variantValue) {
              tokenHashes.push(String(variantValue).trim());
            }
          });
        }
      }
    } catch (hashErr) {
      if (typeof console !== 'undefined' && console && typeof console.warn === 'function') {
        console.warn('setPasswordWithTokenViaSheet: failed to compute token hash', hashErr);
      }
    }
  }

  for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
    const row = data[rowIndex] || [];
    const rowToken = tokenIdx != null && tokenIdx >= 0 ? String(row[tokenIdx] || '').trim() : '';
    const rowId = idIdx != null && idIdx >= 0 ? String(row[idIdx] || '').trim() : '';
    const rowTokenHash = tokenHashIdx != null && tokenHashIdx >= 0 ? String(row[tokenHashIdx] || '').trim() : '';

    const matchesToken = normalizedToken && rowToken && rowToken === normalizedToken;
    const matchesHash = rowTokenHash && tokenHashes.some(candidate => candidate && candidate === rowTokenHash);
    const matchesUserId = normalizedUserId && rowId && rowId === normalizedUserId;

    if (!matchesToken && !matchesHash && !matchesUserId) {
      continue;
    }

    const targetUserId = rowId || normalizedUserId || '';
    if (targetUserId) {
      setUserPasswordDirect(targetUserId, password);
    } else if (typeof console !== 'undefined' && console && typeof console.warn === 'function') {
      console.warn('setPasswordWithTokenViaSheet: unable to resolve user ID for password update', { rowIndex });
    }

    if (tokenIdx != null && tokenIdx >= 0) {
      sheet.getRange(rowIndex + 1, tokenIdx + 1).setValue('');
    }
    if (tokenHashIdx != null && tokenHashIdx >= 0) {
      sheet.getRange(rowIndex + 1, tokenHashIdx + 1).setValue('');
    }
    if (confirmedIdx != null && confirmedIdx >= 0) {
      sheet.getRange(rowIndex + 1, confirmedIdx + 1).setValue('TRUE');
    }
    if (resetIdx != null && resetIdx >= 0) {
      sheet.getRange(rowIndex + 1, resetIdx + 1).setValue('FALSE');
    }
    if (updatedIdx != null && updatedIdx >= 0) {
      sheet.getRange(rowIndex + 1, updatedIdx + 1).setValue(new Date());
    }

    if (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp && typeof SpreadsheetApp.flush === 'function') {
      SpreadsheetApp.flush();
    }

    const invalidateTarget = (typeof USERS_SHEET !== 'undefined' && USERS_SHEET) ? USERS_SHEET : usersSheetName;
    if (typeof invalidateCache === 'function') {
      try { invalidateCache(invalidateTarget); } catch (_) { }
    }

    return { success: true, via: 'sheet-direct' };
  }

  return { success: false, message: 'No matching user row found for password update' };
}

function applySeedPasswordForUser(userRecord, profile, label) {
  if (!profile || !profile.password || !userRecord || !userRecord.ID) {
    return userRecord;
  }

  const resolvedLabel = label || profile.seedLabel || profile.fullName || profile.email;

  if (userRecord.EmailConfirmation) {
    const setPasswordResult = ensurePasswordWithToken(
      userRecord.EmailConfirmation,
      profile.password,
      {
        label: resolvedLabel,
        userId: userRecord.ID || userRecord.Id || userRecord.id
      }
    );

    if (!setPasswordResult || !setPasswordResult.success) {
      throw new Error('Failed to set ' + resolvedLabel + ' password: ' + (setPasswordResult && setPasswordResult.message ? setPasswordResult.message : 'Unknown error'));
    }
  } else {
    setUserPasswordDirect(userRecord.ID, profile.password);
  }

  if (typeof AuthenticationService !== 'undefined' && AuthenticationService.getUserByEmail) {
    const refreshed = AuthenticationService.getUserByEmail(profile.email);
    if (refreshed && userExistsInSheet(refreshed)) {
      return refreshed;
    }
  }

  const fallbackRow = findUserSheetRow(
    profile.email,
    userRecord.ID || userRecord.Id || userRecord.id,
    { forceFresh: true }
  );
  if (fallbackRow) {
    return buildUserRecordFromRow(
      fallbackRow,
      profile.email,
      userRecord && (userRecord.ID || userRecord.Id || userRecord.id)
    );
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

  let existing = (typeof AuthenticationService !== 'undefined' && AuthenticationService.getUserByEmail)
    ? AuthenticationService.getUserByEmail(profile.email)
    : null;

  if (existing && !userExistsInSheet(existing)) {
    existing = null;
  }

  if (existing) {
    const updateResult = clientUpdateUser(existing.ID, payload);

    if (!updateResult || !updateResult.success) {
      throw new Error('Failed to refresh ' + label + ': ' + (updateResult && updateResult.error ? updateResult.error : 'Unknown error'));
    }

    existing = applySeedPasswordForUser(existing, profile, label);
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

  const adminRecord = loadSeedUserRecord(profile, label, createResult.userId || createResult.userID || createResult.id);

  if (!adminRecord) {
    throw new Error(label + ' record not found after creation.');
  }

  const persistedRecord = applySeedPasswordForUser(adminRecord, profile, label) || adminRecord;
  const adminId = persistedRecord.ID || adminRecord.ID;

  syncUserRoleLinks(adminId, desiredRoleIds);
  assignAdminCampaignAccess(adminId, Object.values(campaignIdsByName));
  ensureCanLoginFlag(adminId, true);

  const result = {
    status: 'created',
    userId: adminId,
    email: persistedRecord.Email || adminRecord.Email,
    message: label + ' account created with default credentials. Please change the password after first login.'
  };

  if (profile.password) {
    result.password = profile.password;
  }

  return result;
}

function userExistsInSheet(user) {
  if (!user) return false;
  if (typeof readSheet !== 'function' && typeof dbSelect !== 'function') {
    return true;
  }
  return Boolean(findUserSheetRow(
    user.Email || user.email,
    user.ID || user.Id || user.id,
    { forceFresh: false }
  ));
}

function loadSeedUserRecord(profile, label, expectedUserId) {
  const attempts = 5;
  const delayMs = 500;
  let candidateId = expectedUserId ? String(expectedUserId) : '';

  if (typeof invalidateCache === 'function') {
    try { invalidateCache(USERS_SHEET); } catch (_) { }
  }

  for (let attempt = 0; attempt < attempts; attempt++) {
    const sheetRow = findUserSheetRow(profile.email, candidateId, { forceFresh: attempt > 0 });
    if (sheetRow) {
      return buildUserRecordFromRow(sheetRow, profile.email, candidateId);
    }

    let record = null;
    if (typeof AuthenticationService !== 'undefined' && AuthenticationService.getUserByEmail) {
      try {
        record = AuthenticationService.getUserByEmail(profile.email);
      } catch (lookupErr) {
        console.warn('loadSeedUserRecord: getUserByEmail failed', lookupErr);
      }
    }

    if (record) {
      const recordId = record.ID || record.Id || record.id || candidateId;
      if (userExistsInSheet(record)) {
        return record;
      }

      const refreshedRow = findUserSheetRow(profile.email, recordId, { forceFresh: true });
      if (refreshedRow) {
        return buildUserRecordFromRow(refreshedRow, profile.email, recordId);
      }

      candidateId = recordId ? String(recordId) : candidateId;
    }

    if (attempt < attempts - 1) {
      if (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp && typeof SpreadsheetApp.flush === 'function') {
        SpreadsheetApp.flush();
      }
      if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.sleep === 'function') {
        Utilities.sleep(delayMs);
      }
    }
  }

  throw new Error(label + ' record not found after creation.');
}

function findUserSheetRow(email, userId, options) {
  const normalizedEmail = email ? String(email).trim().toLowerCase() : '';
  const normalizedId = userId ? String(userId).trim() : '';
  const forceFresh = options && options.forceFresh;

  function matchRows(rows) {
    if (!Array.isArray(rows)) return null;
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const rowId = row.ID || row.Id || row.id;
      if (normalizedId && rowId && String(rowId).trim() === normalizedId) {
        return row;
      }

      if (normalizedEmail) {
        const candidates = [
          row.Email,
          row.email,
          row.NormalizedEmail,
          row.normalizedEmail
        ];
        for (let j = 0; j < candidates.length; j++) {
          const candidate = candidates[j];
          if (candidate && String(candidate).trim().toLowerCase() === normalizedEmail) {
            return row;
          }
        }
      }
    }
    return null;
  }

  if (typeof dbSelect === 'function') {
    const queries = [];
    if (normalizedId) {
      queries.push({ where: { ID: normalizedId }, cache: false, limit: 1 });
    }
    if (normalizedEmail) {
      queries.push({ where: { Email: normalizedEmail }, cache: false, limit: 1 });
      queries.push({ where: { NormalizedEmail: normalizedEmail }, cache: false, limit: 1 });
    }
    if (!queries.length) {
      queries.push({ cache: false });
    }

    for (let q = 0; q < queries.length; q++) {
      try {
        const dbRows = dbSelect(USERS_SHEET, queries[q]) || [];
        const match = matchRows(dbRows);
        if (match) {
          return match;
        }
      } catch (dbErr) {
        console.warn('findUserSheetRow: dbSelect failed', dbErr);
      }
    }
  }

  if (forceFresh && typeof invalidateCache === 'function') {
    try { invalidateCache(USERS_SHEET); } catch (_) { }
  }

  if (typeof readSheet === 'function') {
    const primaryOptions = forceFresh
      ? { allowScriptCache: false, useCache: false }
      : { allowScriptCache: false };

    try {
      const sheetRows = readSheet(USERS_SHEET, primaryOptions) || [];
      const match = matchRows(sheetRows);
      if (match) {
        return match;
      }

      if (!forceFresh) {
        const refreshedRows = readSheet(USERS_SHEET, { allowScriptCache: false, useCache: false }) || [];
        if (refreshedRows !== sheetRows) {
          const refreshedMatch = matchRows(refreshedRows);
          if (refreshedMatch) {
            return refreshedMatch;
          }
        }
      }
    } catch (sheetErr) {
      console.warn('findUserSheetRow: sheet read failed', sheetErr);
    }
  }

  return null;
}

function buildUserRecordFromRow(row, fallbackEmail, fallbackId) {
  if (!row) return null;
  const hydrated = Object.assign({}, row);

  const resolvedEmail = hydrated.Email || hydrated.email || hydrated.NormalizedEmail || hydrated.normalizedEmail || fallbackEmail;
  if (resolvedEmail) {
    hydrated.Email = resolvedEmail;
    hydrated.NormalizedEmail = String(resolvedEmail).trim().toLowerCase();
  }

  const resolvedId = hydrated.ID || hydrated.Id || hydrated.id || fallbackId;
  if (resolvedId) {
    hydrated.ID = resolvedId;
  }

  if (!hydrated.EmailConfirmation) {
    hydrated.EmailConfirmation = hydrated.emailConfirmation || hydrated.emailconfirmation || hydrated.EmailConfirmationToken || hydrated.emailConfirmationToken;
  }

  if (!hydrated.UserName && hydrated.username) {
    hydrated.UserName = hydrated.username;
  }

  if (!hydrated.NormalizedUserName && hydrated.UserName) {
    hydrated.NormalizedUserName = String(hydrated.UserName).trim().toLowerCase();
  }

  return hydrated;
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
