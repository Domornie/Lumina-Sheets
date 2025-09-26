/**
 * BookmarkService.gs
 * 
 * Handles bookmark CRUD operations for users
 */

/**
 * Add a new bookmark for a user
 */
function addBookmark(userEmail, title, url, description = '', tags = '', folder = 'General') {
  try {
    const ss = getIBTRSpreadsheet();
    const sheet = ss.getSheetByName(BOOKMARKS_SHEET);
    
    if (!sheet) {
      throw new Error('Bookmarks sheet not found');
    }
    
    const id = Utilities.getUuid();
    const now = new Date();
    
    const newRow = [
      id,
      userEmail,
      title,
      url,
      description,
      tags,
      now,
      null, // LastAccessed
      0,    // AccessCount
      folder
    ];
    
    sheet.appendRow(newRow);
    invalidateCache(BOOKMARKS_SHEET);
    
    return {
      success: true,
      id: id,
      message: 'Bookmark added successfully'
    };
  } catch (error) {
    writeError('addBookmark', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Get all bookmarks for a user
 */
function getUserBookmarks(userEmail) {
  try {
    const bookmarks = readSheet(BOOKMARKS_SHEET);
    return bookmarks.filter(bookmark => bookmark.UserEmail === userEmail);
  } catch (error) {
    writeError('getUserBookmarks', error);
    return [];
  }
}

/**
 * Get bookmarks by folder for a user
 */
function getUserBookmarksByFolder(userEmail, folder = null) {
  try {
    const bookmarks = getUserBookmarks(userEmail);
    
    if (folder) {
      return bookmarks.filter(bookmark => bookmark.Folder === folder);
    }
    
    // Group by folder
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
    writeError('getUserBookmarksByFolder', error);
    return folder ? [] : {};
  }
}

/**
 * Update bookmark access count and last accessed time
 */
function updateBookmarkAccess(bookmarkId) {
  try {
    const ss = getIBTRSpreadsheet();
    const sheet = ss.getSheetByName(BOOKMARKS_SHEET);
    
    if (!sheet) return false;
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const idCol = headers.indexOf('ID');
    const lastAccessedCol = headers.indexOf('LastAccessed');
    const accessCountCol = headers.indexOf('AccessCount');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === bookmarkId) {
        const currentCount = data[i][accessCountCol] || 0;
        
        sheet.getRange(i + 1, lastAccessedCol + 1).setValue(new Date());
        sheet.getRange(i + 1, accessCountCol + 1).setValue(currentCount + 1);
        
        invalidateCache(BOOKMARKS_SHEET);
        return true;
      }
    }
    
    return false;
  } catch (error) {
    writeError('updateBookmarkAccess', error);
    return false;
  }
}

/**
 * Delete a bookmark
 */
function deleteBookmark(bookmarkId, userEmail) {
  try {
    const ss = getIBTRSpreadsheet();
    const sheet = ss.getSheetByName(BOOKMARKS_SHEET);
    
    if (!sheet) return false;
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const idCol = headers.indexOf('ID');
    const userEmailCol = headers.indexOf('UserEmail');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === bookmarkId && data[i][userEmailCol] === userEmail) {
        sheet.deleteRow(i + 1);
        invalidateCache(BOOKMARKS_SHEET);
        return true;
      }
    }
    
    return false;
  } catch (error) {
    writeError('deleteBookmark', error);
    return false;
  }
}

/**
 * Update a bookmark
 */
function updateBookmark(bookmarkId, userEmail, updates) {
  try {
    const ss = getIBTRSpreadsheet();
    const sheet = ss.getSheetByName(BOOKMARKS_SHEET);
    
    if (!sheet) return false;
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const idCol = headers.indexOf('ID');
    const userEmailCol = headers.indexOf('UserEmail');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === bookmarkId && data[i][userEmailCol] === userEmail) {
        // Update allowed fields
        if (updates.title !== undefined) {
          const titleCol = headers.indexOf('Title');
          sheet.getRange(i + 1, titleCol + 1).setValue(updates.title);
        }
        
        if (updates.description !== undefined) {
          const descCol = headers.indexOf('Description');
          sheet.getRange(i + 1, descCol + 1).setValue(updates.description);
        }
        
        if (updates.tags !== undefined) {
          const tagsCol = headers.indexOf('Tags');
          sheet.getRange(i + 1, tagsCol + 1).setValue(updates.tags);
        }
        
        if (updates.folder !== undefined) {
          const folderCol = headers.indexOf('Folder');
          sheet.getRange(i + 1, folderCol + 1).setValue(updates.folder);
        }
        
        invalidateCache(BOOKMARKS_SHEET);
        return true;
      }
    }
    
    return false;
  } catch (error) {
    writeError('updateBookmark', error);
    return false;
  }
}

/**
 * Search bookmarks for a user
 */
function searchUserBookmarks(userEmail, searchTerm) {
  try {
    const bookmarks = getUserBookmarks(userEmail);
    const term = searchTerm.toLowerCase();
    
    return bookmarks.filter(bookmark => 
      bookmark.Title.toLowerCase().includes(term) ||
      bookmark.URL.toLowerCase().includes(term) ||
      bookmark.Description.toLowerCase().includes(term) ||
      bookmark.Tags.toLowerCase().includes(term)
    );
  } catch (error) {
    writeError('searchUserBookmarks', error);
    return [];
  }
}

/**
 * Get user's bookmark folders
 */
function getUserBookmarkFolders(userEmail) {
  try {
    const bookmarks = getUserBookmarks(userEmail);
    const folders = [...new Set(bookmarks.map(b => b.Folder || 'General'))];
    return folders.sort();
  } catch (error) {
    writeError('getUserBookmarkFolders', error);
    return ['General'];
  }
}

/**
 * Check if URL is already bookmarked by user
 */
function isBookmarked(userEmail, url) {
  try {
    const bookmarks = getUserBookmarks(userEmail);
    return bookmarks.some(bookmark => bookmark.URL === url);
  } catch (error) {
    writeError('isBookmarked', error);
    return false;
  }
}

/**
 * Get most accessed bookmarks for a user
 */
function getMostAccessedBookmarks(userEmail, limit = 10) {
  try {
    const bookmarks = getUserBookmarks(userEmail);
    return bookmarks
      .sort((a, b) => (b.AccessCount || 0) - (a.AccessCount || 0))
      .slice(0, limit);
  } catch (error) {
    writeError('getMostAccessedBookmarks', error);
    return [];
  }
}

/**
 * Get recent bookmarks for a user
 */
function getRecentBookmarks(userEmail, limit = 10) {
  try {
    const bookmarks = getUserBookmarks(userEmail);
    return bookmarks
      .sort((a, b) => new Date(b.Created) - new Date(a.Created))
      .slice(0, limit);
  } catch (error) {
    writeError('getRecentBookmarks', error);
    return [];
  }
}

// Export functions for client access
function clientAddBookmark(title, url, description, tags, folder) {
  // Get current user from session
  const userEmail = Session.getActiveUser().getEmail();
  return addBookmark(userEmail, title, url, description, tags, folder);
}

function clientGetUserBookmarks() {
  const userEmail = Session.getActiveUser().getEmail();
  return getUserBookmarks(userEmail);
}

function clientGetUserBookmarksByFolder(folder) {
  const userEmail = Session.getActiveUser().getEmail();
  return getUserBookmarksByFolder(userEmail, folder);
}

function clientUpdateBookmarkAccess(bookmarkId) {
  return updateBookmarkAccess(bookmarkId);
}

function clientDeleteBookmark(bookmarkId) {
  const userEmail = Session.getActiveUser().getEmail();
  return deleteBookmark(bookmarkId, userEmail);
}

function clientUpdateBookmark(bookmarkId, updates) {
  const userEmail = Session.getActiveUser().getEmail();
  return updateBookmark(bookmarkId, userEmail, updates);
}

function clientSearchUserBookmarks(searchTerm) {
  const userEmail = Session.getActiveUser().getEmail();
  return searchUserBookmarks(userEmail, searchTerm);
}

function clientGetUserBookmarkFolders() {
  const userEmail = Session.getActiveUser().getEmail();
  return getUserBookmarkFolders(userEmail);
}

function clientIsBookmarked(url) {
  const userEmail = Session.getActiveUser().getEmail();
  return isBookmarked(userEmail, url);
}

function clientGetMostAccessedBookmarks(limit) {
  const userEmail = Session.getActiveUser().getEmail();
  return getMostAccessedBookmarks(userEmail, limit);
}

function clientGetRecentBookmarks(limit) {
  const userEmail = Session.getActiveUser().getEmail();
  return getRecentBookmarks(userEmail, limit);
}