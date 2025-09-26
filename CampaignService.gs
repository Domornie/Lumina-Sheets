/**
 * REVAMPED CampaignService.gs - Clean, Simple, Reliable
 * Version 2.0 - Eliminates conflicts and simplifies architecture
 */

// ────────────────────────────────────────────────────────────────────────────
// CONSTANTS AND CONFIGURATION
// ────────────────────────────────────────────────────────────────────────────
const CACHE_TTL_MINUTES = 5;
const MAX_RETRIES = 3;
const RETRY_DELAY_MS = 500;

// ────────────────────────────────────────────────────────────────────────────
// UTILITY FUNCTIONS
// ────────────────────────────────────────────────────────────────────────────
// REPLACE your current addUserToCampaign with this:
function addUserToCampaign(userId, campaignId) {
  try {
    if (!userId || !campaignId) return false;

    const sh = ensureSheetWithHeaders(USER_CAMPAIGNS_SHEET, USER_CAMPAIGNS_HEADERS);
    const data = sh.getDataRange().getValues();
    const headers = data[0] || [];
    const idx = {
      ID: headers.indexOf('ID'),
      UserID: headers.indexOf('UserID'),
      CampaignID: headers.indexOf('CampaignID'),
      CreatedAt: headers.indexOf('CreatedAt'),
      UpdatedAt: headers.indexOf('UpdatedAt')
    };

    // If columns missing, just append safely
    let exists = false;
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][idx.UserID]) === String(userId) &&
        String(data[r][idx.CampaignID]) === String(campaignId)) {
        exists = true;
        // Touch UpdatedAt for existing relationship
        if (idx.UpdatedAt >= 0) sh.getRange(r + 1, idx.UpdatedAt + 1).setValue(new Date());
        commitChanges();
        return true;
      }
    }

    // Append new relation
    const now = new Date();
    const row = [];
    if (idx.ID >= 0) row[idx.ID] = Utilities.getUuid();
    if (idx.UserID >= 0) row[idx.UserID] = userId;
    if (idx.CampaignID >= 0) row[idx.CampaignID] = campaignId;
    if (idx.CreatedAt >= 0) row[idx.CreatedAt] = now;
    if (idx.UpdatedAt >= 0) row[idx.UpdatedAt] = now;

    // Fill any undefined cells to preserve width
    for (let c = 0; c < headers.length; c++) if (typeof row[c] === 'undefined') row[c] = '';
    sh.appendRow(row);

    commitChanges();
    invalidateSheetCache(USER_CAMPAIGNS_SHEET);
    return true;
  } catch (e) {
    safeWriteError('addUserToCampaign', e);
    return false;
  }
}

function removeUserFromCampaign(userId, campaignId) {
  try {
    if (!userId || !campaignId) return { success: false, error: 'userId and campaignId required' };
    const sh = ensureSheetWithHeaders(USER_CAMPAIGNS_SHEET, USER_CAMPAIGNS_HEADERS);
    const data = sh.getDataRange().getValues();
    const headers = data[0] || [];
    const idx = {
      UserID: headers.indexOf('UserID'),
      CampaignID: headers.indexOf('CampaignID')
    };
    let removed = 0;
    for (let r = data.length - 1; r >= 1; r--) {
      if (String(data[r][idx.UserID]) === String(userId) &&
        String(data[r][idx.CampaignID]) === String(campaignId)) {
        sh.deleteRow(r + 1);
        removed++;
      }
    }
    if (removed) { commitChanges(); invalidateSheetCache(USER_CAMPAIGNS_SHEET); }
    return { success: true, removed };
  } catch (e) {
    safeWriteError('removeUserFromCampaign', e);
    return { success: false, error: e.message };
  }
}

function csGetUserCampaignIds(userId) {
  try {
    if (!userId) return [];
    const sh = ensureSheetWithHeaders(USER_CAMPAIGNS_SHEET, USER_CAMPAIGNS_HEADERS);
    const data = sh.getDataRange().getValues();
    const headers = data[0] || [];
    const uIdx = headers.indexOf('UserID');
    const cIdx = headers.indexOf('CampaignID');
    const out = new Set();
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][uIdx]) === String(userId)) out.add(String(data[r][cIdx]));
    }
    return Array.from(out);
  } catch (e) {
    safeWriteError('csGetUserCampaignIds', e);
    return [];
  }
}

function csGetUserCampaigns(userId) {
  try {
    const ids = csGetUserCampaignIds(userId);
    if (!ids.length) return [];
    const campaigns = readSheet(CAMPAIGNS_SHEET) || [];
    const want = new Set(ids.map(String));
    return campaigns.filter(c => want.has(String(c.ID)))
      .map(c => ({ id: c.ID, name: c.Name, description: c.Description || '' }));
  } catch (e) {
    safeWriteError('csGetUserCampaigns', e);
    return [];
  }
}

/**
 * Safe sheet operations with retry logic
 */
function safeSheetOperation(operation, maxRetries = MAX_RETRIES) {
  let lastError;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const result = operation();
      if (attempt > 1) {
        console.log(`Operation succeeded on attempt ${attempt}`);
      }
      return result;
    } catch (error) {
      lastError = error;
      console.warn(`Attempt ${attempt} failed:`, error.message);

      if (attempt < maxRetries) {
        Utilities.sleep(RETRY_DELAY_MS * attempt);
      }
    }
  }

  throw lastError;
}

/**
 * Force commit sheet changes
 */
function commitChanges() {
  try {
    SpreadsheetApp.flush();
    Utilities.sleep(100); // Give Google Sheets time to process
  } catch (error) {
    console.warn('Failed to commit changes:', error);
  }
}

/**
 * Clear all caches for a campaign
 */
function clearCampaignCaches(campaignId) {
  const cacheKeys = [
    `CAMPAIGN_DATA_${campaignId}`,
    `NAVIGATION_${campaignId}`,
    `CATEGORIES_${campaignId}`,
    `PAGES_${campaignId}`,
    `DATA_${CAMPAIGNS_SHEET}`,
    `DATA_${PAGE_CATEGORIES_SHEET}`,
    `DATA_${CAMPAIGN_PAGES_SHEET}`
  ];

  cacheKeys.forEach(key => {
    try {
      scriptCache.remove(key);
      console.log(`Cleared cache: ${key}`);
    } catch (error) {
      console.warn(`Failed to clear cache ${key}:`, error);
    }
  });

  // Also clear using the existing invalidation functions
  if (typeof invalidateCache === 'function') {
    invalidateCache(CAMPAIGNS_SHEET);
    invalidateCache(PAGE_CATEGORIES_SHEET);
    invalidateCache(CAMPAIGN_PAGES_SHEET);
  }

  if (typeof invalidateNavigationCache === 'function') {
    invalidateNavigationCache(campaignId);
  }
}

// ────────────────────────────────────────────────────────────────────────────
// CORE CAMPAIGN OPERATIONS
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get all campaigns
 */
function csGetAllCampaigns() {
  try {
    console.log('csGetAllCampaigns: Starting...');

    const campaigns = safeSheetOperation(() => {
      return readSheet(CAMPAIGNS_SHEET) || [];
    });

    const result = campaigns
      .filter(c => c && c.ID && c.Name && String(c.Name).trim() !== '')
      .map(c => ({
        id: c.ID,
        name: c.Name,
        description: c.Description || '',
        createdAt: c.CreatedAt ? new Date(c.CreatedAt).toISOString() : null,
        updatedAt: c.UpdatedAt ? new Date(c.UpdatedAt).toISOString() : null
      }))
      .sort((a, b) => {
        const da = a.createdAt ? new Date(a.createdAt).getTime() : 0;
        const db = b.createdAt ? new Date(b.createdAt).getTime() : 0;
        return db - da;
      });

    console.log(`csGetAllCampaigns: Returning ${result.length} campaigns`);
    return result;

  } catch (error) {
    console.error('csGetAllCampaigns error:', error);
    safeWriteError('csGetAllCampaigns', error);
    return [];
  }
}

/**
 * Get campaign statistics
 */
function csGetCampaignStats() {
  try {
    const campaigns = readSheet(CAMPAIGNS_SHEET) || [];
    const users = readSheet(USERS_SHEET) || [];
    const campaignPages = readSheet(CAMPAIGN_PAGES_SHEET) || [];

    return campaigns.map(c => ({
      id: c.ID,
      userCount: users.filter(u => u.CampaignID === c.ID).length,
      pageCount: campaignPages.filter(p => p.CampaignID === c.ID && isActive(p.IsActive)).length
    }));

  } catch (error) {
    console.error('csGetCampaignStats error:', error);
    safeWriteError('csGetCampaignStats', error);
    return [];
  }
}

/**
 * Create new campaign
 */
function csCreateCampaign(name, description = '') {
  try {
    if (!name || !name.trim()) {
      return { success: false, error: 'Campaign name is required' };
    }

    name = name.trim();
    description = (description || '').trim();

    // Check for duplicates
    const existing = readSheet(CAMPAIGNS_SHEET) || [];
    const duplicate = existing.find(c =>
      c.Name && c.Name.toLowerCase() === name.toLowerCase()
    );

    if (duplicate) {
      return { success: false, error: 'A campaign with this name already exists' };
    }

    return safeSheetOperation(() => {
      const campaignId = Utilities.getUuid();
      const now = new Date();
      const sheet = ensureSheetWithHeaders(CAMPAIGNS_SHEET, CAMPAIGNS_HEADERS);

      sheet.appendRow([campaignId, name, description, now, now]);
      commitChanges();

      // Clear caches
      clearCampaignCaches(campaignId);

      console.log(`Created campaign: ${name} with ID: ${campaignId}`);
      return {
        success: true,
        message: `Campaign "${name}" created successfully`,
        data: {
          id: campaignId,
          name: name,
          description: description,
          createdAt: now.toISOString(),
          updatedAt: now.toISOString()
        }
      };
    });

  } catch (error) {
    console.error('csCreateCampaign error:', error);
    safeWriteError('csCreateCampaign', error);
    return { success: false, error: 'Failed to create campaign: ' + error.message };
  }
}

/**
 * Update campaign
 */
function csUpdateCampaign(campaignId, name, description = '') {
  try {
    if (!name || !name.trim()) {
      return { success: false, error: 'Campaign name is required' };
    }

    return safeSheetOperation(() => {
      name = name.trim();
      description = (description || '').trim();

      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CAMPAIGNS_SHEET);
      if (!sheet) {
        return { success: false, error: 'Campaigns sheet not found' };
      }

      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const rows = data.slice(1);

      const rowIndex = rows.findIndex(r => r[0] === campaignId);
      if (rowIndex === -1) {
        return { success: false, error: 'Campaign not found' };
      }

      // Check for duplicate names
      const duplicateIndex = rows.findIndex((r, i) =>
        i !== rowIndex && (r[1] || '').toLowerCase() === name.toLowerCase()
      );
      if (duplicateIndex !== -1) {
        return { success: false, error: 'A campaign with this name already exists' };
      }

      // Update the row
      const actualRowIndex = rowIndex + 2;
      sheet.getRange(actualRowIndex, 2).setValue(name);
      sheet.getRange(actualRowIndex, 3).setValue(description);
      sheet.getRange(actualRowIndex, 5).setValue(new Date());

      commitChanges();
      clearCampaignCaches(campaignId);

      console.log(`Updated campaign: ${campaignId} to name: ${name}`);
      return { success: true, message: 'Campaign updated successfully' };
    });

  } catch (error) {
    console.error('csUpdateCampaign error:', error);
    safeWriteError('csUpdateCampaign', error);
    return { success: false, error: 'Failed to update campaign: ' + error.message };
  }
}

/**
 * Delete campaign
 */
function csDeleteCampaign(campaignId) {
  try {
    return safeSheetOperation(() => {
      // Check for assigned users
      const users = readSheet(USERS_SHEET) || [];
      const assignedUsers = users.filter(u => u.CampaignID === campaignId);

      if (assignedUsers.length > 0) {
        const userNames = assignedUsers.map(u => u.FullName || u.UserName).join(', ');
        return {
          success: false,
          error: `Cannot delete campaign. ${assignedUsers.length} user${assignedUsers.length !== 1 ? 's are' : ' is'} still assigned: ${userNames}`
        };
      }

      // PATCH: drop this block into csDeleteCampaign() right after you compute `assignedUsers`
      const ucSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_CAMPAIGNS_SHEET);
      if (ucSheet) {
        const data = ucSheet.getDataRange().getValues();
        const headers = data[0] || [];
        const cIdx = headers.indexOf('CampaignID');
        const uIdx = headers.indexOf('UserID');
        let linkCount = 0;
        const linkedUserIds = new Set();
        for (let r = 1; r < data.length; r++) {
          if (String(data[r][cIdx]) === String(campaignId)) {
            linkCount++;
            if (uIdx >= 0 && data[r][uIdx]) linkedUserIds.add(String(data[r][uIdx]));
          }
        }
        if (linkCount > 0) {
          return {
            success: false,
            error: `Cannot delete campaign. ${linkCount} user-campaign link${linkCount !== 1 ? 's' : ''} exist` +
              (linkedUserIds.size ? ` (users: ${Array.from(linkedUserIds).join(', ')})` : '')
          };
        }
      }

      // Get campaign name
      const campaigns = readSheet(CAMPAIGNS_SHEET) || [];
      const campaign = campaigns.find(c => c.ID === campaignId);
      const campaignName = campaign ? campaign.Name : 'Campaign';

      // Delete the campaign
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CAMPAIGNS_SHEET);
      if (!sheet) {
        return { success: false, error: 'Campaigns sheet not found' };
      }

      const data = sheet.getDataRange().getValues();
      const rowIndex = data.findIndex((row, index) => index > 0 && row[0] === campaignId);
      if (rowIndex === -1) {
        return { success: false, error: 'Campaign not found' };
      }

      sheet.deleteRow(rowIndex + 1);

      // Clean up related data
      cleanupCampaignData(campaignId);
      commitChanges();
      clearCampaignCaches(campaignId);

      console.log(`Deleted campaign: ${campaignName} with ID: ${campaignId}`);
      return {
        success: true,
        message: `Campaign "${campaignName}" and all associated data deleted successfully`
      };
    });

  } catch (error) {
    console.error('csDeleteCampaign error:', error);
    safeWriteError('csDeleteCampaign', error);
    return { success: false, error: 'Failed to delete campaign: ' + error.message };
  }
}

function csGetCampaignStatsV2() {
  try {
    const campaigns = readSheet(CAMPAIGNS_SHEET) || [];
    const users = readSheet(USERS_SHEET) || [];
    const campaignPages = readSheet(CAMPAIGN_PAGES_SHEET) || [];
    let uc = [];
    try {
      const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_CAMPAIGNS_SHEET);
      uc = sh ? sh.getDataRange().getValues().slice(1) : [];
    } catch (_) {}

    // Build a quick multimap of campaignId -> linked user count
    const ucHeaders = (uc.length ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_CAMPAIGNS_SHEET).getDataRange().getValues()[0] : []) || [];
    const ucCIdx = ucHeaders.indexOf('CampaignID');
    const ucUCount = {};
    uc.forEach(r => {
      const cid = String(r[ucCIdx]);
      ucUCount[cid] = (ucUCount[cid] || 0) + 1;
    });

    return campaigns.map(c => {
      const cid = String(c.ID);
      const oneToOne = users.filter(u => String(u.CampaignID) === cid).length;
      const manyToMany = ucUCount[cid] || 0;
      return {
        id: cid,
        userCount: oneToOne + manyToMany,
        pageCount: campaignPages.filter(p => String(p.CampaignID) === cid && isActive(p.IsActive)).length
      };
    });
  } catch (e) {
    safeWriteError('csGetCampaignStatsV2', e);
    return [];
  }
}

function csGetCampaignById(campaignId) {
  try {
    const rows = readSheet(CAMPAIGNS_SHEET) || [];
    const c = rows.find(r => String(r.ID) === String(campaignId));
    if (!c) return null;
    return { id: c.ID, name: c.Name, description: c.Description || '', createdAt: c.CreatedAt, updatedAt: c.UpdatedAt };
  } catch (e) {
    safeWriteError('csGetCampaignById', e);
    return null;
  }
}

function csDeactivateCategory(categoryId, active) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PAGE_CATEGORIES_SHEET);
    if (!sh) return { success: false, error: 'Categories sheet not found' };
    const data = sh.getDataRange().getValues();
    const headers = data[0] || [];
    const idCol = 1 + headers.indexOf('ID');
    const activeCol = 1 + headers.indexOf('IsActive');
    const updatedCol = 1 + headers.indexOf('UpdatedAt');

    let row = -1, campaignId = null;
    for (let r = 2; r <= data.length; r++) {
      if (String(sh.getRange(r, idCol).getValue()) === String(categoryId)) {
        row = r;
        const cIdCol = 1 + headers.indexOf('CampaignID');
        campaignId = cIdCol > 0 ? sh.getRange(r, cIdCol).getValue() : null;
        break;
      }
    }
    if (row === -1) return { success: false, error: 'Category not found' };
    if (activeCol > 0) sh.getRange(row, activeCol).setValue(!!active);
    if (updatedCol > 0) sh.getRange(row, updatedCol).setValue(new Date());
    commitChanges();
    if (campaignId) clearCampaignCaches(campaignId);
    return { success: true, message: `Category ${active ? 'activated' : 'deactivated'}` };
  } catch (e) {
    safeWriteError('csDeactivateCategory', e);
    return { success: false, error: e.message };
  }
}

function csDeactivateCampaignPage(pageId, active) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CAMPAIGN_PAGES_SHEET);
    if (!sh) return { success: false, error: 'Campaign pages sheet not found' };
    const data = sh.getDataRange().getValues();
    const headers = data[0] || [];
    const idCol = 1 + headers.indexOf('ID');
    const activeCol = 1 + headers.indexOf('IsActive');
    const updatedCol = 1 + headers.indexOf('UpdatedAt');
    const campaignCol = 1 + headers.indexOf('CampaignID');

    let row = -1;
    for (let r = 2; r <= data.length; r++) {
      if (String(sh.getRange(r, idCol).getValue()) === String(pageId)) { row = r; break; }
    }
    if (row === -1) return { success: false, error: 'Campaign page not found' };

    const campaignId = campaignCol > 0 ? sh.getRange(row, campaignCol).getValue() : null;
    if (activeCol > 0) sh.getRange(row, activeCol).setValue(!!active);
    if (updatedCol > 0) sh.getRange(row, updatedCol).setValue(new Date());

    commitChanges();
    if (campaignId) clearCampaignCaches(campaignId);
    return { success: true, message: `Page ${active ? 'activated' : 'deactivated'}` };
  } catch (e) {
    safeWriteError('csDeactivateCampaignPage', e);
    return { success: false, error: e.message };
  }
}

function csRebuildAllNavigations() {
  try {
    const all = csGetAllCampaigns();
    const results = [];
    all.forEach(c => {
      try {
        clearCampaignCaches(c.id);
        const nav = csGetCampaignNavigation(c.id);
        results.push({ id: c.id, name: c.name, categories: nav.categories.length, uncategorized: nav.uncategorizedPages.length });
      } catch (e) {
        results.push({ id: c.id, name: c.name, error: e.message });
      }
    });
    return { success: true, results };
  } catch (e) {
    safeWriteError('csRebuildAllNavigations', e);
    return { success: false, error: e.message };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// PAGE CATEGORY OPERATIONS
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get categories for a campaign
 */
function csGetCampaignCategories(campaignId) {
  try {
    if (!campaignId) {
      console.warn('csGetCampaignCategories called without campaign ID');
      return [];
    }

    console.log(`Getting categories for campaign: ${campaignId}`);

    const allCategories = safeSheetOperation(() => {
      return readSheet(PAGE_CATEGORIES_SHEET) || [];
    });

    const categories = allCategories
      .filter(c => {
        if (!c || !c.CampaignID) return false;
        const campaignMatch = String(c.CampaignID) === String(campaignId);
        const active = isActive(c.IsActive);
        return campaignMatch && active;
      })
      .map(c => ({
        id: c.ID,
        campaignId: c.CampaignID,
        categoryName: c.CategoryName,
        categoryIcon: c.CategoryIcon || 'fas fa-folder',
        sortOrder: parseInt(c.SortOrder) || 999,
        isActive: c.IsActive,
        createdAt: c.CreatedAt,
        updatedAt: c.UpdatedAt
      }))
      .sort((a, b) => (a.sortOrder || 999) - (b.sortOrder || 999));

    console.log(`Found ${categories.length} categories for campaign ${campaignId}`);
    return categories;

  } catch (error) {
    console.error('csGetCampaignCategories error:', error);
    safeWriteError('csGetCampaignCategories', error);
    return [];
  }
}

/**
 * Create category for campaign
 */
function csCreateCategory(campaignId, categoryName, categoryIcon, sortOrder = 999) {
  try {
    console.log('Creating category:', { campaignId, categoryName, categoryIcon, sortOrder });

    if (!campaignId || !categoryName || !categoryIcon) {
      return { success: false, error: 'Campaign ID, category name, and icon are required' };
    }

    return safeSheetOperation(() => {
      // Verify campaign exists
      const campaigns = readSheet(CAMPAIGNS_SHEET) || [];
      if (!campaigns.find(c => c.ID === campaignId)) {
        return { success: false, error: 'Campaign not found' };
      }

      // Check for duplicates
      const existing = readSheet(PAGE_CATEGORIES_SHEET) || [];
      const duplicate = existing.find(c =>
        c.CampaignID === campaignId &&
        isActive(c.IsActive) &&
        (c.CategoryName || '').toLowerCase() === categoryName.trim().toLowerCase()
      );

      if (duplicate) {
        return { success: false, error: 'A category with this name already exists in this campaign' };
      }

      // Create the category
      const id = Utilities.getUuid();
      const now = new Date();
      const sheet = ensureSheetWithHeaders(PAGE_CATEGORIES_SHEET, PAGE_CATEGORIES_HEADERS);

      sheet.appendRow([
        id,
        campaignId,
        categoryName.trim(),
        categoryIcon.trim(),
        parseInt(sortOrder) || 999,
        true, // IsActive - use boolean
        now,
        now
      ]);

      commitChanges();
      clearCampaignCaches(campaignId);

      console.log('Successfully created category:', categoryName);

      // Verify creation
      const verification = readSheet(PAGE_CATEGORIES_SHEET)
        .find(c => c.ID === id && c.CampaignID === campaignId);

      if (!verification) {
        console.error('Category verification failed');
        return { success: false, error: 'Failed to verify category creation' };
      }

      return {
        success: true,
        message: 'Category created successfully',
        data: {
          id: id,
          campaignId: campaignId,
          categoryName: categoryName,
          categoryIcon: categoryIcon,
          sortOrder: sortOrder
        }
      };
    });

  } catch (error) {
    console.error('csCreateCategory error:', error);
    safeWriteError('csCreateCategory', error);
    return { success: false, error: 'Failed to create category: ' + error.message };
  }
}

/**
 * Delete category
 */
function csDeleteCategory(categoryId) {
  try {
    console.log('Deleting category:', categoryId);

    return safeSheetOperation(() => {
      // Get category info first
      const categories = readSheet(PAGE_CATEGORIES_SHEET) || [];
      const category = categories.find(c => c.ID === categoryId);

      if (!category) {
        return { success: false, error: 'Category not found' };
      }

      // Move pages to uncategorized
      const campaignPages = readSheet(CAMPAIGN_PAGES_SHEET) || [];
      const pagesInCategory = campaignPages.filter(p => p.CategoryID === categoryId);

      if (pagesInCategory.length > 0) {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CAMPAIGN_PAGES_SHEET);
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const categoryIdCol = headers.indexOf('CategoryID') + 1;
        const updatedAtCol = headers.indexOf('UpdatedAt') + 1;

        pagesInCategory.forEach((page, index) => {
          const rowIndex = data.findIndex((row, i) => i > 0 && row[0] === page.ID);
          if (rowIndex > 0) {
            sheet.getRange(rowIndex + 1, categoryIdCol).setValue('');
            if (updatedAtCol > 0) {
              sheet.getRange(rowIndex + 1, updatedAtCol).setValue(new Date());
            }
          }
        });

        console.log(`Moved ${pagesInCategory.length} pages to uncategorized`);
      }

      // Delete the category
      const categorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PAGE_CATEGORIES_SHEET);
      if (!categorySheet) {
        return { success: false, error: 'Categories sheet not found' };
      }

      const categoryData = categorySheet.getDataRange().getValues();
      const categoryRowIndex = categoryData.findIndex((row, index) => index > 0 && row[0] === categoryId);
      if (categoryRowIndex === -1) {
        return { success: false, error: 'Category not found in sheet' };
      }

      categorySheet.deleteRow(categoryRowIndex + 1);
      commitChanges();
      clearCampaignCaches(category.CampaignID);

      console.log('Successfully deleted category');
      return { success: true, message: 'Category deleted successfully' };
    });

  } catch (error) {
    console.error('csDeleteCategory error:', error);
    safeWriteError('csDeleteCategory', error);
    return { success: false, error: 'Failed to delete category: ' + error.message };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// CAMPAIGN PAGE OPERATIONS
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get pages for a campaign
 */
function csGetCampaignPages(campaignId) {
  try {
    const campaignPages = readSheet(CAMPAIGN_PAGES_SHEET) || [];
    return campaignPages
      .filter(p => p && p.CampaignID === campaignId && isActive(p.IsActive))
      .map(p => ({
        id: p.ID,
        campaignId: p.CampaignID,
        pageKey: p.PageKey,
        pageTitle: p.PageTitle,
        pageIcon: p.PageIcon,
        categoryId: p.CategoryID || null,
        sortOrder: p.SortOrder || 999,
        isActive: p.IsActive
      }))
      .sort((a, b) => (a.sortOrder || 999) - (b.sortOrder || 999));

  } catch (error) {
    console.error('csGetCampaignPages error:', error);
    safeWriteError('csGetCampaignPages', error);
    return [];
  }
}

/**
 * Add page to campaign
 */
function csAddPageToCampaign(campaignId, pageKey, pageTitle, pageIcon, categoryId = null, sortOrder = 999) {
  try {
    console.log('Adding page to campaign:', { campaignId, pageKey, pageTitle, pageIcon, categoryId, sortOrder });

    if (!campaignId || !pageKey || !pageTitle || !pageIcon) {
      return { success: false, error: 'Campaign ID, page key, title, and icon are required' };
    }

    return safeSheetOperation(() => {
      // Verify campaign exists
      const campaigns = readSheet(CAMPAIGNS_SHEET) || [];
      if (!campaigns.find(c => c.ID === campaignId)) {
        return { success: false, error: 'Campaign not found' };
      }

      // Check if page already assigned
      const existing = readSheet(CAMPAIGN_PAGES_SHEET) || [];
      const existingPage = existing.find(cp =>
        cp.CampaignID === campaignId &&
        cp.PageKey === pageKey &&
        isActive(cp.IsActive)
      );

      if (existingPage) {
        return { success: false, error: 'Page is already assigned to this campaign' };
      }

      // Verify system page exists
      const systemPages = readSheet(PAGES_SHEET) || [];
      if (!systemPages.find(p => p.PageKey === pageKey)) {
        return { success: false, error: 'System page not found' };
      }

      // Validate category if provided
      if (categoryId && categoryId !== '') {
        const categories = readSheet(PAGE_CATEGORIES_SHEET) || [];
        const category = categories.find(c =>
          c.ID === categoryId &&
          c.CampaignID === campaignId &&
          isActive(c.IsActive)
        );
        if (!category) {
          console.warn('Invalid categoryId provided:', categoryId, 'setting to null');
          categoryId = null;
        }
      } else {
        categoryId = null;
      }

      const id = Utilities.getUuid();
      const now = new Date();
      const sheet = ensureSheetWithHeaders(CAMPAIGN_PAGES_SHEET, CAMPAIGN_PAGES_HEADERS);

      sheet.appendRow([
        id,
        campaignId,
        pageKey,
        pageTitle.trim(),
        pageIcon.trim(),
        categoryId,
        parseInt(sortOrder) || 999,
        true,
        now,
        now
      ]);

      commitChanges();
      clearCampaignCaches(campaignId);

      console.log('Successfully added page to campaign');
      return {
        success: true,
        message: 'Page added to campaign successfully',
        data: {
          id: id,
          campaignId: campaignId,
          pageKey: pageKey,
          pageTitle: pageTitle,
          pageIcon: pageIcon,
          categoryId: categoryId,
          sortOrder: sortOrder
        }
      };
    });

  } catch (error) {
    console.error('csAddPageToCampaign error:', error);
    safeWriteError('csAddPageToCampaign', error);
    return { success: false, error: 'Failed to add page to campaign: ' + error.message };
  }
}

/**
 * Update campaign page
 */
function csUpdateCampaignPage(pageId, updates) {
  try {
    console.log('Updating campaign page:', pageId, 'with updates:', updates);

    return safeSheetOperation(() => {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CAMPAIGN_PAGES_SHEET);
      if (!sheet) return { success: false, error: 'Campaign pages sheet not found' };

      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const rows = data.slice(1);

      const rowIndex = rows.findIndex(row => row[0] === pageId);
      if (rowIndex === -1) return { success: false, error: 'Campaign page not found' };

      const campaignId = rows[rowIndex][1];
      const actualRowIndex = rowIndex + 2;

      const updateMap = {
        'PageTitle': headers.indexOf('PageTitle') + 1,
        'PageIcon': headers.indexOf('PageIcon') + 1,
        'CategoryID': headers.indexOf('CategoryID') + 1,
        'SortOrder': headers.indexOf('SortOrder') + 1,
        'IsActive': headers.indexOf('IsActive') + 1
      };

      // Validate CategoryID if being updated
      if (updates.CategoryID !== undefined) {
        if (updates.CategoryID === null || updates.CategoryID === '') {
          updates.CategoryID = null;
        } else {
          const categories = readSheet(PAGE_CATEGORIES_SHEET) || [];
          const validCategory = categories.find(c =>
            c.ID === updates.CategoryID &&
            c.CampaignID === campaignId &&
            isActive(c.IsActive)
          );

          if (!validCategory) {
            console.warn('Invalid CategoryID provided:', updates.CategoryID, 'setting to null');
            updates.CategoryID = null;
          }
        }
      }

      // Apply updates
      Object.keys(updates).forEach(field => {
        const colIndex = updateMap[field];
        if (colIndex) {
          let value = updates[field];

          if (field === 'SortOrder') {
            value = parseInt(value) || 999;
          } else if (field === 'IsActive') {
            value = Boolean(value);
          } else if (field === 'CategoryID') {
            value = value || '';
          }

          console.log(`Updating ${field} to:`, value, 'at column:', colIndex);
          sheet.getRange(actualRowIndex, colIndex).setValue(value);
        }
      });

      // Update timestamp
      const updatedAtCol = headers.indexOf('UpdatedAt') + 1;
      if (updatedAtCol > 0) {
        sheet.getRange(actualRowIndex, updatedAtCol).setValue(new Date());
      }

      commitChanges();
      clearCampaignCaches(campaignId);

      console.log('Successfully updated campaign page');
      return { success: true, message: 'Campaign page updated successfully' };
    });

  } catch (error) {
    console.error('csUpdateCampaignPage error:', error);
    safeWriteError('csUpdateCampaignPage', error);
    return { success: false, error: 'Failed to update campaign page: ' + error.message };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// NAVIGATION OPERATIONS
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get complete navigation for a campaign
 */
function csGetCampaignNavigation(campaignId) {
  try {
    if (!campaignId) {
      console.warn('csGetCampaignNavigation called without campaign ID');
      return { categories: [], uncategorizedPages: [] };
    }

    console.log(`Building navigation for campaign ${campaignId}`);

    const pages = csGetCampaignPages(campaignId);
    const categories = csGetCampaignCategories(campaignId);
    const navigation = { categories: [], uncategorizedPages: [] };

    console.log(`Found ${pages.length} pages and ${categories.length} categories`);

    if (categories.length === 0) {
      console.warn(`No categories found for campaign ${campaignId}. All pages will be uncategorized.`);
      navigation.uncategorizedPages = pages.map(page => ({
        ...page,
        pageIcon: page.pageIcon || 'fas fa-file'
      }));
      return navigation;
    }

    // Create category map
    const categoryMap = {};
    categories.forEach(cat => {
      categoryMap[cat.id] = {
        id: cat.id,
        campaignId: cat.campaignId,
        categoryName: cat.categoryName,
        categoryIcon: cat.categoryIcon || 'fas fa-folder',
        sortOrder: cat.sortOrder || 999,
        isActive: cat.isActive,
        pages: []
      };
    });

    // Assign pages to categories
    pages.forEach(page => {
      page.pageIcon = page.pageIcon || 'fas fa-file';

      if (page.categoryId && categoryMap[page.categoryId]) {
        categoryMap[page.categoryId].pages.push(page);
      } else {
        navigation.uncategorizedPages.push(page);
      }
    });

    // Only include categories with pages
    navigation.categories = Object.values(categoryMap).filter(cat => cat.pages.length > 0);

    console.log(`Final navigation: ${navigation.categories.length} categories with pages, ${navigation.uncategorizedPages.length} uncategorized pages`);

    return navigation;

  } catch (error) {
    console.error('csGetCampaignNavigation error:', error);
    safeWriteError('csGetCampaignNavigation', error);
    return { categories: [], uncategorizedPages: [] };
  }
}

/**
 * Force refresh navigation cache
 */
function csRefreshNavigation(campaignId) {
  try {
    console.log('Force refreshing navigation for campaign:', campaignId);

    clearCampaignCaches(campaignId);
    const navigation = csGetCampaignNavigation(campaignId);

    return {
      success: true,
      message: 'Navigation refreshed successfully',
      data: navigation
    };

  } catch (error) {
    console.error('csRefreshNavigation error:', error);
    safeWriteError('csRefreshNavigation', error);
    return { success: false, error: error.message };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// UTILITY FUNCTIONS
// ────────────────────────────────────────────────────────────────────────────

/**
 * Check if a value represents "active" status
 */
function isActive(value) {
  if (value === true || value === 'TRUE' || value === 'true') return true;
  if (String(value).toLowerCase() === 'true') return true;
  return false;
}

/**
 * Get all system pages
 */
function csGetAllPages() {
  try {
    const pages = readSheet(PAGES_SHEET) || [];
    return pages
      .filter(p => p && p.PageKey && p.PageTitle)
      .map(p => ({
        pageKey: p.PageKey,
        pageTitle: p.PageTitle,
        pageIcon: p.PageIcon || 'fas fa-file',
        requiresAdmin: p.RequiresAdmin || false
      }))
      .sort((a, b) => (a.pageTitle || '').localeCompare(b.pageTitle || ''));

  } catch (error) {
    console.error('csGetAllPages error:', error);
    safeWriteError('csGetAllPages', error);
    return [];
  }
}

/**
 * Clean up campaign-related data
 */
function cleanupCampaignData(campaignId) {
  try {
    console.log('Cleaning up data for campaign:', campaignId);

    const sheetsToClean = [
      { sheet: CAMPAIGN_PAGES_SHEET, column: 'CampaignID' },
      { sheet: PAGE_CATEGORIES_SHEET, column: 'CampaignID' },
      { sheet: CAMPAIGN_USER_PERMISSIONS_SHEET, column: 'CampaignID' }
    ];

    sheetsToClean.forEach(({ sheet, column }) => {
      try {
        deleteRowsByColumnValue(sheet, column, campaignId);
      } catch (error) {
        console.warn(`Failed to clean ${sheet}:`, error);
      }
    });

    console.log('Cleanup completed for campaign:', campaignId);
  } catch (error) {
    console.error('Error during cleanup:', error);
    safeWriteError('cleanupCampaignData', error);
  }
}

/**
 * Delete rows where column matches value
 */
function deleteRowsByColumnValue(sheetName, columnName, value) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    const headers = data[0];
    const colIndex = headers.indexOf(columnName);
    if (colIndex === -1) return;

    let deletedCount = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][colIndex] === value) {
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
    }

    if (deletedCount > 0) {
      console.log(`Deleted ${deletedCount} rows from ${sheetName} where ${columnName} = ${value}`);
    }
  } catch (error) {
    console.error(`Error deleting from ${sheetName}:`, error);
  }
}

// ────────────────────────────────────────────────────────────────────────────
// CLIENT-FACING FUNCTIONS (Legacy Compatibility)
// ────────────────────────────────────────────────────────────────────────────

// Maintain compatibility with existing frontend code
function clientGetAllCampaigns() { return csGetAllCampaigns(); }
function clientGetCampaignStats() { return csGetCampaignStats(); }
function clientAddCampaign(name, description) { return csCreateCampaign(name, description); }
function clientUpdateCampaign(id, name, description) { return csUpdateCampaign(id, name, description); }
function clientDeleteCampaign(id) { return csDeleteCampaign(id); }
function clientGetAllPages() { return csGetAllPages(); }
function clientGetCampaignPages(id) { return csGetCampaignPages(id); }
function clientGetCampaignNavigation(id) { return csGetCampaignNavigation(id); }
function clientAddCategoryToCampaign(cId, name, icon, sort) { return csCreateCategory(cId, name, icon, sort); }
function clientDeleteCampaignCategory(id) { return csDeleteCategory(id); }
function clientAddPageToCampaign(cId, key, title, icon, catId, sort) { return csAddPageToCampaign(cId, key, title, icon, catId, sort); }
function clientUpdateCampaignPage(id, updates) { return csUpdateCampaignPage(id, updates); }
function forceRefreshCampaignNavigation(id) { return csRefreshNavigation(id); }

// ────────────────────────────────────────────────────────────────────────────
// TESTING AND DEBUG FUNCTIONS
// ────────────────────────────────────────────────────────────────────────────

/**
 * Test connection
 */
function testFrontendConnection() {
  try {
    console.log('testFrontendConnection: Starting...');

    const campaigns = csGetAllCampaigns();
    const sampleCampaigns = campaigns.slice(0, 3);

    return {
      success: true,
      timestamp: new Date(),
      message: 'Frontend connection working',
      sampleCampaignCount: campaigns.length,
      sampleCampaigns: sampleCampaigns,
      functionsAvailable: {
        readSheet: typeof readSheet === 'function',
        ensureSheetWithHeaders: typeof ensureSheetWithHeaders === 'function',
        safeWriteError: typeof safeWriteError === 'function'
      }
    };

  } catch (error) {
    console.error('testFrontendConnection error:', error);
    safeWriteError('testFrontendConnection', error);
    return { success: false, error: error.message || String(error) };
  }
}

/**
 * Debug campaign issues
 */
function debugCampaignIssues(campaignId) {
  try {
    console.log('=== DEBUGGING CAMPAIGN ISSUES ===');

    // Auto-select first campaign if none provided
    if (!campaignId) {
      const campaigns = csGetAllCampaigns();
      if (campaigns.length > 0) {
        campaignId = campaigns[0].id;
        console.log('Using first campaign:', campaignId);
      } else {
        return { error: 'No campaigns found' };
      }
    }

    const campaigns = csGetAllCampaigns();
    const campaign = campaigns.find(c => c.id === campaignId);

    if (!campaign) {
      return {
        error: 'Campaign not found',
        availableCampaigns: campaigns.map(c => ({ id: c.id, name: c.name }))
      };
    }

    const categories = csGetCampaignCategories(campaignId);
    const pages = csGetCampaignPages(campaignId);
    const navigation = csGetCampaignNavigation(campaignId);

    const result = {
      campaignId,
      campaignName: campaign.name,
      categoriesCount: categories.length,
      pagesCount: pages.length,
      navigationCategoriesCount: navigation.categories.length,
      uncategorizedPagesCount: navigation.uncategorizedPages.length,
      categories: categories.slice(0, 3), // Sample data
      pages: pages.slice(0, 3) // Sample data
    };

    console.log('Debug result:', result);
    return result;

  } catch (error) {
    console.error('Debug error:', error);
    return { error: error.message, stack: error.stack };
  }
}

/**
 * Clear all caches
 */
function forceClearAllCaches(campaignId) {
  try {
    if (!campaignId) {
      const campaigns = csGetAllCampaigns();
      if (campaigns.length > 0) {
        campaignId = campaigns[0].id;
      } else {
        return { success: false, error: 'No campaigns found' };
      }
    }

    clearCampaignCaches(campaignId);

    return {
      success: true,
      message: 'All caches cleared successfully',
      campaignId
    };

  } catch (error) {
    console.error('Error clearing caches:', error);
    return { success: false, error: error.message };
  }
}