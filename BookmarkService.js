/**
 * BookmarkService.gs
 *
 * Handles bookmark CRUD operations for users
 */

const __BOOKMARK_GLOBAL__ = (typeof globalThis !== 'undefined')
  ? globalThis
  : (typeof this !== 'undefined' ? this : {});
const __BOOKMARK_G__ = (__BOOKMARK_GLOBAL__ && typeof __BOOKMARK_GLOBAL__.G === 'object')
  ? __BOOKMARK_GLOBAL__.G
  : __BOOKMARK_GLOBAL__;

const BOOKMARKS_DEFAULT_HEADERS = [
  'ID', 'UserID', 'UserEmail', 'Title', 'URL', 'Description', 'Tags',
  'Folder', 'Created', 'LastAccessed', 'AccessCount'
];

// eslint-disable-next-line no-var
var BOOKMARKS_SHEET = (typeof BOOKMARKS_SHEET !== 'undefined')
  ? BOOKMARKS_SHEET
  : ((typeof __BOOKMARK_G__.BOOKMARKS_SHEET === 'string' && __BOOKMARK_G__.BOOKMARKS_SHEET)
    || 'Bookmarks');

// eslint-disable-next-line no-var
var BOOKMARKS_HEADERS = (typeof BOOKMARKS_HEADERS !== 'undefined' && Array.isArray(BOOKMARKS_HEADERS))
  ? BOOKMARKS_HEADERS.slice()
  : (Array.isArray(__BOOKMARK_G__.BOOKMARKS_HEADERS)
    ? __BOOKMARK_G__.BOOKMARKS_HEADERS.slice()
    : BOOKMARKS_DEFAULT_HEADERS.slice());

if (!Array.isArray(BOOKMARKS_HEADERS) || !BOOKMARKS_HEADERS.length) {
  BOOKMARKS_HEADERS = BOOKMARKS_DEFAULT_HEADERS.slice();
}

if (__BOOKMARK_G__ && (!Array.isArray(__BOOKMARK_G__.BOOKMARKS_HEADERS) || __BOOKMARK_G__.BOOKMARKS_HEADERS.length !== BOOKMARKS_HEADERS.length)) {
  __BOOKMARK_G__.BOOKMARKS_HEADERS = BOOKMARKS_HEADERS.slice();
}

function getBookmarkHeaders_() {
  return Array.isArray(BOOKMARKS_HEADERS) && BOOKMARKS_HEADERS.length
    ? BOOKMARKS_HEADERS.slice()
    : BOOKMARKS_DEFAULT_HEADERS.slice();
}

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function normalizeId_(value) {
  const str = String(value || '').trim();
  return str || '';
}

function normalizeHeaderKey_(header) {
  return String(header || '').trim().toLowerCase();
}

function safeWriteError_(label, error) {
  if (typeof writeError === 'function') {
    try {
      writeError(label, error);
    } catch (_) { /* ignore logging failures */ }
  }
}

function getActiveUserEmail_() {
  if (typeof Session === 'undefined' || !Session || typeof Session.getActiveUser !== 'function') {
    return '';
  }
  try {
    const activeUser = Session.getActiveUser();
    if (activeUser && typeof activeUser.getEmail === 'function') {
      return activeUser.getEmail();
    }
  } catch (_) { /* ignore session lookup issues */ }
  return '';
}

function getBookmarkSheet_(createIfMissing) {
  const ss = (typeof getIBTRSpreadsheet === 'function') ? getIBTRSpreadsheet() : null;
  if (!ss) {
    throw new Error('Spreadsheet not available');
  }

  let sheet = ss.getSheetByName(BOOKMARKS_SHEET);
  if (!sheet && createIfMissing !== false) {
    if (typeof ensureSheetWithHeaders === 'function') {
      sheet = ensureSheetWithHeaders(BOOKMARKS_SHEET, getBookmarkHeaders_());
    } else {
      sheet = ss.insertSheet(BOOKMARKS_SHEET);
      const headers = getBookmarkHeaders_();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  }

  return sheet;
}

function getBookmarkSheetHeaders_(sheet) {
  if (!sheet) return getBookmarkHeaders_();
  const lastColumn = sheet.getLastColumn();
  if (!lastColumn) return getBookmarkHeaders_();
  const values = sheet.getRange(1, 1, 1, lastColumn).getValues();
  if (!values || !values[0]) return getBookmarkHeaders_();
  return values[0].map(value => (value === null || typeof value === 'undefined') ? '' : String(value));
}

function buildBookmarkRow_(headers, record) {
  const headerList = Array.isArray(headers) && headers.length ? headers : getBookmarkHeaders_();
  const row = new Array(headerList.length);

  for (let i = 0; i < headerList.length; i++) {
    const rawHeader = headerList[i];
    const key = normalizeHeaderKey_(rawHeader);
    switch (key) {
      case 'id':
        row[i] = record.id;
        break;
      case 'userid':
        row[i] = record.userId;
        break;
      case 'useremail':
      case 'email':
        row[i] = record.userEmail;
        break;
      case 'title':
        row[i] = record.title;
        break;
      case 'url':
        row[i] = record.url;
        break;
      case 'description':
        row[i] = record.description;
        break;
      case 'tags':
        row[i] = record.tags;
        break;
      case 'folder':
        row[i] = record.folder;
        break;
      case 'created':
      case 'createdat':
        row[i] = record.created;
        break;
      case 'lastaccessed':
      case 'lastaccessedat':
        row[i] = record.lastAccessed;
        break;
      case 'accesscount':
      case 'views':
        row[i] = record.accessCount;
        break;
      default:
        row[i] = (record.extra && Object.prototype.hasOwnProperty.call(record.extra, rawHeader))
          ? record.extra[rawHeader]
          : '';
        break;
    }
  }

  return row;
}

function resolveUserIdentity_(identityOrEmail) {
  if (identityOrEmail && typeof identityOrEmail === 'object' &&
    Object.prototype.hasOwnProperty.call(identityOrEmail, 'email') &&
    Object.prototype.hasOwnProperty.call(identityOrEmail, 'normalizedEmail') &&
    Object.prototype.hasOwnProperty.call(identityOrEmail, 'userId')) {
    return {
      email: identityOrEmail.email || '',
      normalizedEmail: identityOrEmail.normalizedEmail || normalizeEmail_(identityOrEmail.email || ''),
      userId: normalizeId_(identityOrEmail.userId || identityOrEmail.UserID || identityOrEmail.ID || ''),
      user: identityOrEmail.user || null
    };
  }

  let email = '';
  let userId = '';
  let possibleNames = [];

  if (identityOrEmail && typeof identityOrEmail === 'object') {
    email = identityOrEmail.email || identityOrEmail.userEmail || identityOrEmail.UserEmail || '';
    userId = identityOrEmail.userId || identityOrEmail.UserID || identityOrEmail.id || identityOrEmail.ID || '';
    possibleNames = [
      identityOrEmail.userName,
      identityOrEmail.UserName,
      identityOrEmail.fullName,
      identityOrEmail.FullName,
      identityOrEmail.displayName,
      identityOrEmail.DisplayName
    ].filter(Boolean);
  } else if (typeof identityOrEmail === 'string') {
    email = identityOrEmail;
  }

  if (!email) {
    email = getActiveUserEmail_();
  }

  const normalizedEmail = normalizeEmail_(email);
  let normalizedUserId = normalizeId_(userId);
  let matchedUser = null;

  if (typeof UserService !== 'undefined' && UserService && typeof UserService.buildUserIdentifierLookup === 'function') {
    try {
      const lookup = UserService.buildUserIdentifierLookup();
      if (!matchedUser && normalizedUserId && lookup.byId && lookup.byId[normalizedUserId]) {
        matchedUser = lookup.byId[normalizedUserId];
      }
      if (!matchedUser && normalizedEmail && lookup.byEmail && lookup.byEmail[normalizedEmail]) {
        matchedUser = lookup.byEmail[normalizedEmail];
      }
      if (!matchedUser && Array.isArray(possibleNames) && possibleNames.length && lookup.byUser) {
        for (let i = 0; i < possibleNames.length; i++) {
          const candidate = String(possibleNames[i] || '').trim().toLowerCase();
          if (candidate && lookup.byUser[candidate]) {
            matchedUser = lookup.byUser[candidate];
            break;
          }
        }
      }
    } catch (lookupError) {
      safeWriteError_('BookmarkService.resolveUserIdentity', lookupError);
    }
  }

  if (matchedUser) {
    if (!normalizedUserId) {
      normalizedUserId = normalizeId_(matchedUser.ID || matchedUser.Id || matchedUser.UserID || matchedUser.userId);
    }
    if (!email) {
      email = matchedUser.Email || matchedUser.email || '';
    }
  }

  return {
    email: email || '',
    normalizedEmail,
    userId: normalizedUserId,
    user: matchedUser || null
  };
}

function doesRowBelongToIdentity_(row, userIdCol, userEmailCol, identity) {
  if (!identity) return false;
  const rowUserId = userIdCol >= 0 ? normalizeId_(row[userIdCol]) : '';
  if (identity.userId && rowUserId && rowUserId === identity.userId) {
    return true;
  }

  const rowEmail = userEmailCol >= 0 ? normalizeEmail_(row[userEmailCol]) : '';
  if (identity.normalizedEmail && rowEmail && rowEmail === identity.normalizedEmail) {
    return true;
  }

  if (!rowUserId && !rowEmail && (identity.userId || identity.normalizedEmail)) {
    return true;
  }

  return false;
}

function maybeBackfillBookmarkIdentityRow_(sheet, rowIndex, headers, row, identity, userIdCol, userEmailCol) {
  if (!sheet || !identity) return;
  const updates = [];

  if (userIdCol >= 0 && identity.userId && !normalizeId_(row[userIdCol])) {
    updates.push({ column: userIdCol, value: identity.userId });
  }

  if (userEmailCol >= 0 && identity.email && !normalizeEmail_(row[userEmailCol])) {
    updates.push({ column: userEmailCol, value: identity.email });
  }

  if (!updates.length) return;

  updates.forEach(update => {
    sheet.getRange(rowIndex + 2, update.column + 1).setValue(update.value);
  });
}

function cloneAndEnhanceBookmark_(bookmark, identity) {
  if (!bookmark || typeof bookmark !== 'object') return bookmark;
  const clone = Object.assign({}, bookmark);

  if (identity) {
    if (identity.userId && !normalizeId_(clone.UserID || clone.UserId || clone.userId)) {
      clone.UserID = identity.userId;
    }
    if (identity.email && !normalizeEmail_(clone.UserEmail || clone.Email || clone.userEmail)) {
      clone.UserEmail = identity.email;
    }
  }

  return clone;
}

function toLowerSafe_(value) {
  return (value === null || typeof value === 'undefined')
    ? ''
    : String(value).toLowerCase();
}

/**
 * Add a new bookmark for a user
 */
function addBookmark(identityOrEmail, title, url, description = '', tags = '', folder = 'General') {
  const identity = resolveUserIdentity_(identityOrEmail);
  if (!identity.userId && !identity.normalizedEmail) {
    const error = new Error('Unable to resolve current user identity for bookmark creation');
    safeWriteError_('addBookmark', error);
    return { success: false, error: error.message };
  }

  try {
    const sheet = getBookmarkSheet_(true);
    if (!sheet) {
      throw new Error('Bookmarks sheet not found');
    }

    const headers = getBookmarkSheetHeaders_(sheet);
    const id = Utilities.getUuid();
    const now = new Date();

    const record = {
      id,
      userId: identity.userId,
      userEmail: identity.email,
      title,
      url,
      description,
      tags,
      folder: folder || 'General',
      created: now,
      lastAccessed: null,
      accessCount: 0,
      extra: {}
    };

    const row = buildBookmarkRow_(headers, record);
    sheet.appendRow(row);

    if (typeof invalidateCache === 'function') {
      invalidateCache(BOOKMARKS_SHEET);
    }

    return {
      success: true,
      id,
      message: 'Bookmark added successfully'
    };
  } catch (error) {
    safeWriteError_('addBookmark', error);
    return {
      success: false,
      error: error && error.message ? error.message : String(error)
    };
  }
}

/**
 * Get all bookmarks for a user
 */
function getUserBookmarks(identityOrEmail) {
  const identity = resolveUserIdentity_(identityOrEmail);
  if (!identity.userId && !identity.normalizedEmail) {
    return [];
  }

  try {
    const bookmarks = (typeof readSheet === 'function') ? readSheet(BOOKMARKS_SHEET) : [];
    if (!Array.isArray(bookmarks) || !bookmarks.length) {
      return [];
    }

    return bookmarks
      .filter(bookmark => {
        if (!bookmark || typeof bookmark !== 'object') return false;
        const rowUserId = normalizeId_(bookmark.UserID || bookmark.UserId || bookmark.userId);
        if (identity.userId && rowUserId && rowUserId === identity.userId) {
          return true;
        }
        const rowEmail = normalizeEmail_(bookmark.UserEmail || bookmark.Email || bookmark.userEmail);
        if (identity.normalizedEmail && rowEmail && rowEmail === identity.normalizedEmail) {
          return true;
        }
        return false;
      })
      .map(bookmark => cloneAndEnhanceBookmark_(bookmark, identity));
  } catch (error) {
    safeWriteError_('getUserBookmarks', error);
    return [];
  }
}

/**
 * Get bookmarks by folder for a user
 */
function getUserBookmarksByFolder(identityOrEmail, folder = null) {
  try {
    const bookmarks = getUserBookmarks(identityOrEmail);

    if (folder) {
      return bookmarks.filter(bookmark => (bookmark.Folder || 'General') === folder);
    }

    const grouped = {};
    bookmarks.forEach(bookmark => {
      const folderName = bookmark.Folder || 'General';
      if (!grouped[folderName]) {
        grouped[folderName] = [];
      }
      grouped[folderName].push(bookmark);
    });

    return grouped;
  } catch (error) {
    safeWriteError_('getUserBookmarksByFolder', error);
    return folder ? [] : {};
  }
}

/**
 * Update bookmark access count and last accessed time
 */
function updateBookmarkAccess(bookmarkId, identityOrEmail) {
  const identity = resolveUserIdentity_(identityOrEmail);
  if (!bookmarkId || (!identity.userId && !identity.normalizedEmail)) {
    return false;
  }

  try {
    const sheet = getBookmarkSheet_(false);
    if (!sheet) return false;

    const headers = getBookmarkSheetHeaders_(sheet).map(String);
    const idCol = headers.findIndex(h => normalizeHeaderKey_(h) === 'id');
    const lastAccessedCol = headers.findIndex(h => normalizeHeaderKey_(h) === 'lastaccessed' || normalizeHeaderKey_(h) === 'lastaccessedat');
    const accessCountCol = headers.findIndex(h => normalizeHeaderKey_(h) === 'accesscount' || normalizeHeaderKey_(h) === 'views');
    const userIdCol = headers.findIndex(h => normalizeHeaderKey_(h) === 'userid');
    const userEmailCol = headers.findIndex(h => normalizeHeaderKey_(h) === 'useremail' || normalizeHeaderKey_(h) === 'email');

    if (idCol === -1 || lastAccessedCol === -1 || accessCountCol === -1) {
      return false;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;

    const rows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    const targetId = normalizeId_(bookmarkId);

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (normalizeId_(row[idCol]) !== targetId) {
        continue;
      }
      if (!doesRowBelongToIdentity_(row, userIdCol, userEmailCol, identity)) {
        continue;
      }

      const currentCount = Number(row[accessCountCol] || 0);
      sheet.getRange(i + 2, lastAccessedCol + 1).setValue(new Date());
      sheet.getRange(i + 2, accessCountCol + 1).setValue(isNaN(currentCount) ? 1 : currentCount + 1);
      maybeBackfillBookmarkIdentityRow_(sheet, i, headers, row, identity, userIdCol, userEmailCol);

      if (typeof invalidateCache === 'function') {
        invalidateCache(BOOKMARKS_SHEET);
      }

      return true;
    }

    return false;
  } catch (error) {
    safeWriteError_('updateBookmarkAccess', error);
    return false;
  }
}

/**
 * Delete a bookmark
 */
function deleteBookmark(bookmarkId, identityOrEmail) {
  const identity = resolveUserIdentity_(identityOrEmail);
  if (!bookmarkId || (!identity.userId && !identity.normalizedEmail)) {
    return false;
  }

  try {
    const sheet = getBookmarkSheet_(false);
    if (!sheet) return false;

    const headers = getBookmarkSheetHeaders_(sheet).map(String);
    const idCol = headers.findIndex(h => normalizeHeaderKey_(h) === 'id');
    const userIdCol = headers.findIndex(h => normalizeHeaderKey_(h) === 'userid');
    const userEmailCol = headers.findIndex(h => normalizeHeaderKey_(h) === 'useremail' || normalizeHeaderKey_(h) === 'email');

    if (idCol === -1) return false;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;

    const rows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    const targetId = normalizeId_(bookmarkId);

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (normalizeId_(row[idCol]) !== targetId) {
        continue;
      }
      if (!doesRowBelongToIdentity_(row, userIdCol, userEmailCol, identity)) {
        continue;
      }

      sheet.deleteRow(i + 2);
      if (typeof invalidateCache === 'function') {
        invalidateCache(BOOKMARKS_SHEET);
      }
      return true;
    }

    return false;
  } catch (error) {
    safeWriteError_('deleteBookmark', error);
    return false;
  }
}

/**
 * Update a bookmark
 */
function updateBookmark(bookmarkId, identityOrEmail, updates) {
  const identity = resolveUserIdentity_(identityOrEmail);
  if (!bookmarkId || (!identity.userId && !identity.normalizedEmail)) {
    return false;
  }

  try {
    const sheet = getBookmarkSheet_(false);
    if (!sheet) return false;

    const headers = getBookmarkSheetHeaders_(sheet).map(String);
    const idCol = headers.findIndex(h => normalizeHeaderKey_(h) === 'id');
    const userIdCol = headers.findIndex(h => normalizeHeaderKey_(h) === 'userid');
    const userEmailCol = headers.findIndex(h => normalizeHeaderKey_(h) === 'useremail' || normalizeHeaderKey_(h) === 'email');

    const titleCol = headers.findIndex(h => normalizeHeaderKey_(h) === 'title');
    const descCol = headers.findIndex(h => normalizeHeaderKey_(h) === 'description');
    const tagsCol = headers.findIndex(h => normalizeHeaderKey_(h) === 'tags');
    const folderCol = headers.findIndex(h => normalizeHeaderKey_(h) === 'folder');

    if (idCol === -1) return false;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;

    const rows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    const targetId = normalizeId_(bookmarkId);

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (normalizeId_(row[idCol]) !== targetId) {
        continue;
      }
      if (!doesRowBelongToIdentity_(row, userIdCol, userEmailCol, identity)) {
        continue;
      }

      let changed = false;
      if (updates && Object.prototype.hasOwnProperty.call(updates, 'title') && titleCol !== -1) {
        sheet.getRange(i + 2, titleCol + 1).setValue(updates.title);
        changed = true;
      }
      if (updates && Object.prototype.hasOwnProperty.call(updates, 'description') && descCol !== -1) {
        sheet.getRange(i + 2, descCol + 1).setValue(updates.description);
        changed = true;
      }
      if (updates && Object.prototype.hasOwnProperty.call(updates, 'tags') && tagsCol !== -1) {
        sheet.getRange(i + 2, tagsCol + 1).setValue(updates.tags);
        changed = true;
      }
      if (updates && Object.prototype.hasOwnProperty.call(updates, 'folder') && folderCol !== -1) {
        sheet.getRange(i + 2, folderCol + 1).setValue(updates.folder);
        changed = true;
      }

      if (changed) {
        maybeBackfillBookmarkIdentityRow_(sheet, i, headers, row, identity, userIdCol, userEmailCol);
        if (typeof invalidateCache === 'function') {
          invalidateCache(BOOKMARKS_SHEET);
        }
      }

      return changed;
    }

    return false;
  } catch (error) {
    safeWriteError_('updateBookmark', error);
    return false;
  }
}

/**
 * Search bookmarks for a user
 */
function searchUserBookmarks(identityOrEmail, searchTerm) {
  try {
    const bookmarks = getUserBookmarks(identityOrEmail);
    const term = String(searchTerm || '').trim().toLowerCase();
    if (!term) {
      return bookmarks;
    }

    return bookmarks.filter(bookmark => {
      return (
        toLowerSafe_(bookmark.Title).includes(term) ||
        toLowerSafe_(bookmark.URL).includes(term) ||
        toLowerSafe_(bookmark.Description).includes(term) ||
        toLowerSafe_(bookmark.Tags).includes(term)
      );
    });
  } catch (error) {
    safeWriteError_('searchUserBookmarks', error);
    return [];
  }
}

/**
 * Get user's bookmark folders
 */
function getUserBookmarkFolders(identityOrEmail) {
  try {
    const bookmarks = getUserBookmarks(identityOrEmail);
    const folders = [...new Set(bookmarks.map(b => b.Folder || 'General'))];
    return folders.sort();
  } catch (error) {
    safeWriteError_('getUserBookmarkFolders', error);
    return ['General'];
  }
}

/**
 * Check if URL is already bookmarked by user
 */
function isBookmarked(identityOrEmail, url) {
  try {
    const bookmarks = getUserBookmarks(identityOrEmail);
    const target = String(url || '').trim();
    if (!target) return false;
    return bookmarks.some(bookmark => String(bookmark.URL || '').trim() === target);
  } catch (error) {
    safeWriteError_('isBookmarked', error);
    return false;
  }
}

/**
 * Get most accessed bookmarks for a user
 */
function getMostAccessedBookmarks(identityOrEmail, limit = 10) {
  try {
    const bookmarks = getUserBookmarks(identityOrEmail);
    return bookmarks
      .slice()
      .sort((a, b) => (Number(b.AccessCount || 0) - Number(a.AccessCount || 0)))
      .slice(0, limit);
  } catch (error) {
    safeWriteError_('getMostAccessedBookmarks', error);
    return [];
  }
}

/**
 * Get recent bookmarks for a user
 */
function getRecentBookmarks(identityOrEmail, limit = 10) {
  try {
    const bookmarks = getUserBookmarks(identityOrEmail);
    return bookmarks
      .slice()
      .sort((a, b) => new Date(b.Created) - new Date(a.Created))
      .slice(0, limit);
  } catch (error) {
    safeWriteError_('getRecentBookmarks', error);
    return [];
  }
}

// Export functions for client access
function clientAddBookmark(title, url, description, tags, folder) {
  return addBookmark(resolveUserIdentity_(), title, url, description, tags, folder);
}

function clientGetUserBookmarks() {
  return getUserBookmarks(resolveUserIdentity_());
}

function clientGetUserBookmarksByFolder(folder) {
  return getUserBookmarksByFolder(resolveUserIdentity_(), folder);
}

function clientUpdateBookmarkAccess(bookmarkId) {
  return updateBookmarkAccess(bookmarkId, resolveUserIdentity_());
}

function clientDeleteBookmark(bookmarkId) {
  return deleteBookmark(bookmarkId, resolveUserIdentity_());
}

function clientUpdateBookmark(bookmarkId, updates) {
  return updateBookmark(bookmarkId, resolveUserIdentity_(), updates);
}

function clientSearchUserBookmarks(searchTerm) {
  return searchUserBookmarks(resolveUserIdentity_(), searchTerm);
}

function clientGetUserBookmarkFolders() {
  return getUserBookmarkFolders(resolveUserIdentity_());
}

function clientIsBookmarked(url) {
  return isBookmarked(resolveUserIdentity_(), url);
}

function clientGetMostAccessedBookmarks(limit) {
  return getMostAccessedBookmarks(resolveUserIdentity_(), limit);
}

function clientGetRecentBookmarks(limit) {
  return getRecentBookmarks(resolveUserIdentity_(), limit);
}
