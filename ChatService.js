/**
 * ChatService.gs
 *
 * Chat backend that relies on Utilities.gs for:
 *   • setupSheets()
 *   • readSheet(sheetName)
 *   • writeDebug(msg)
 *   • writeError(context, err)
 *   • ensureSheetWithHeaders(name, headers)
 *   • readSheetPaged(sheetName, pageIndex, pageSize)
 */

// ────────────────────────────────────────────────────────────────────────────
// INITIALIZATION
// ────────────────────────────────────────────────────────────────────────────
/**
 * Ensure all sheets and headers exist before using any chat functions.
 */
function initializeChatService() {
  setupMainSheets();  // from Utilities.gs
}

// ────────────────────────────────────────────────────────────────────────────
// UTILITY HELPERS
// ────────────────────────────────────────────────────────────────────────────
/** Generate a UUID for new records */
function generateId() {
  return Utilities.getUuid();
}

/** Get the current user's email */
function getCurrentUser() {
  return Session.getActiveUser().getEmail();
}

/**
 * Returns true if the given (or current) user has IsAdmin === true in the Users sheet.
 */
function isUserAdmin(userEmail) {
  const email = userEmail || getCurrentUser();
  const user = readSheet(USERS_SHEET).find(u => u.Email === email);
  return !!(user && user.IsAdmin);
}

// ────────────────────────────────────────────────────────────────────────────
// GROUP MANAGEMENT
// ────────────────────────────────────────────────────────────────────────────
/** List groups visible to the current user */
function listMyGroups() {
  initializeChatService();
  const me = getCurrentUser();
  let groups = readSheet(CHAT_GROUPS_SHEET);
  if (isUserAdmin(me)) return groups;
  const memberRows = readSheet(CHAT_GROUP_MEMBERS_SHEET)
    .filter(m => m.UserId === me && m.IsActive);
  const ids = [...new Set(memberRows.map(m => m.GroupId))];
  return groups.filter(g => ids.includes(g.ID));
}

/** Create a new chat group */
function createChatGroup(name, description = '') {
  initializeChatService();
  if (!name.trim()) return { success: false, error: 'Group name required' };
  const now = new Date(), user = getCurrentUser(), id = generateId();
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHAT_GROUPS_SHEET);
    sh.appendRow([ id, name, description, user, now, now ]);
    writeDebug(`group_created:${id}`);
    // auto‑add creator as admin member & default channel
    addUserToChatGroup(id, user, 'admin');
    createChatChannel(id, 'general', 'Default channel');
    return { success: true, groupId: id };
  } catch (e) {
    writeError('createChatGroup', e);
    return { success: false, error: e.message };
  }
}

/** Update an existing chat group */
function updateChatGroup(groupId, updates) {
  initializeChatService();
  const rows = readSheet(CHAT_GROUPS_SHEET);
  const idx  = rows.findIndex(r => r.ID === groupId);
  if (idx < 0) return { success: false, error: 'Group not found' };
  try {
    const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHAT_GROUPS_SHEET);
    const data = sh.getDataRange().getValues();
    const hdr  = data.shift();
    if (updates.name)        data[idx][1] = updates.name;
    if (updates.description) data[idx][2] = updates.description;
    data[idx][5] = new Date();  // UpdatedAt
    sh.clear(); sh.appendRow(hdr);
    sh.getRange(2,1,data.length,data[0].length).setValues(data);
    writeDebug(`group_updated:${groupId}`);
    return { success: true };
  } catch (e) {
    writeError('updateChatGroup', e);
    return { success: false, error: e.message };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// CHANNEL MANAGEMENT
// ────────────────────────────────────────────────────────────────────────────
/** List channels for a given group */
function getChatChannels(groupId) {
  initializeChatService();
  return readSheet(CHAT_CHANNELS_SHEET).filter(c => c.GroupId === groupId);
}

/** Create a new channel */
function createChatChannel(groupId, name, description = '', isPrivate = false) {
  initializeChatService();
  if (!name.trim()) return { success: false, error: 'Channel name required' };
  const now = new Date(), user = getCurrentUser(), id = generateId();
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHAT_CHANNELS_SHEET);
    sh.appendRow([ id, groupId, name, description, isPrivate, user, now, now ]);
    writeDebug(`channel_created:${id}`);
    return { success: true, channelId: id };
  } catch (e) {
    writeError('createChatChannel', e);
    return { success: false, error: e.message };
  }
}

/** Update an existing channel */
function updateChatChannel(channelId, updates) {
  initializeChatService();
  const rows = readSheet(CHAT_CHANNELS_SHEET);
  const idx  = rows.findIndex(r => r.ID === channelId);
  if (idx < 0) return { success: false, error: 'Channel not found' };
  try {
    const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHAT_CHANNELS_SHEET);
    const data = sh.getDataRange().getValues();
    const hdr  = data.shift();
    if (updates.name)        data[idx][2] = updates.name;
    if (updates.description) data[idx][3] = updates.description;
    if (typeof updates.isPrivate === 'boolean') data[idx][4] = updates.isPrivate;
    data[idx][7] = new Date();  // UpdatedAt
    sh.clear(); sh.appendRow(hdr);
    sh.getRange(2,1,data.length,data[0].length).setValues(data);
    writeDebug(`channel_updated:${channelId}`);
    return { success: true };
  } catch (e) {
    writeError('updateChatChannel', e);
    return { success: false, error: e.message };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// MESSAGE MANAGEMENT
// ────────────────────────────────────────────────────────────────────────────
/** Get all messages in a channel (chronological) */
function getChatMessages(channelId) {
  initializeChatService();
  return readSheet(CHAT_MESSAGES_SHEET)
    .filter(m => m.ChannelId === channelId && !m.IsDeleted)
    .sort((a,b) => new Date(a.Timestamp) - new Date(b.Timestamp));
}

/** Paginated message fetch (newest first) */
function getChatMessagesPaginated(channelId, limit = MESSAGE_BATCH_SIZE, offset = 0) {
  initializeChatService();
  const all = readSheet(CHAT_MESSAGES_SHEET)
    .filter(m => m.ChannelId === channelId && !m.IsDeleted)
    .sort((a,b) => new Date(b.Timestamp) - new Date(a.Timestamp));
  const slice = all.slice(offset, offset + limit);
  return {
    messages: slice.reverse(),
    hasMore: all.length > offset + limit,
    total: all.length
  };
}

/** Post a new message */
function postChatMessage(channelId, userId, text, parentMessageId = null) {
  initializeChatService();
  if (!text.trim()) return { success: false, error: 'Message cannot be empty' };
  if (text.length > 2000) return { success: false, error: 'Message too long' };
  const now = new Date(), id = generateId();
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHAT_MESSAGES_SHEET);
    sh.appendRow([ id, channelId, userId, text, now, null, parentMessageId, false ]);
    writeDebug(`message_posted:${id}`);
    updateUserLastSeen(userId);
    return { success: true, messageId: id };
  } catch (e) {
    writeError('postChatMessage', e);
    return { success: false, error: e.message };
  }
}

/** Soft-delete a message */
function deleteChatMessage(messageId, userId) {
  initializeChatService();
  const rows = readSheet(CHAT_MESSAGES_SHEET);
  const idx  = rows.findIndex(r => r.ID === messageId && (r.UserId === userId || isUserAdmin(userId)));
  if (idx < 0) return { success: false, error: 'Not found or unauthorized' };
  try {
    const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHAT_MESSAGES_SHEET);
    const data = sh.getDataRange().getValues();
    const hdr  = data.shift();
    data[idx][7] = true;        // IsDeleted
    data[idx][5] = new Date();  // EditedAt
    sh.clear(); sh.appendRow(hdr);
    sh.getRange(2,1,data.length,data[0].length).setValues(data);
    writeDebug(`message_deleted:${messageId}`);
    return { success: true };
  } catch (e) {
    writeError('deleteChatMessage', e);
    return { success: false, error: e.message };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// GROUP MEMBER MANAGEMENT
// ────────────────────────────────────────────────────────────────────────────
/** Add a user to a group */
function addUserToChatGroup(groupId, userId, role = 'member') {
  initializeChatService();
  const exists = readSheet(CHAT_GROUP_MEMBERS_SHEET)
    .some(m => m.GroupId === groupId && m.UserId === userId && m.IsActive);
  if (exists) return { success: false, error: 'Already a member' };
  const now = new Date(), id = generateId();
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHAT_GROUP_MEMBERS_SHEET);
    sh.appendRow([ id, groupId, userId, now, role, true ]);
    writeDebug(`user_added:${groupId}|${userId}`);
    return { success: true, memberId: id };
  } catch (e) {
    writeError('addUserToChatGroup', e);
    return { success: false, error: e.message };
  }
}

/** Remove a user from a group */
function removeUserFromChatGroup(groupId, userId) {
  initializeChatService();
  const rows = readSheet(CHAT_GROUP_MEMBERS_SHEET);
  const idx  = rows.findIndex(r => r.GroupId === groupId && r.UserId === userId);
  if (idx < 0) return { success: false, error: 'Membership not found' };
  try {
    const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHAT_GROUP_MEMBERS_SHEET);
    const data = sh.getDataRange().getValues();
    const hdr  = data.shift();
    data[idx][5] = false; // IsActive
    sh.clear(); sh.appendRow(hdr);
    sh.getRange(2,1,data.length,data[0].length).setValues(data);
    writeDebug(`user_removed:${groupId}|${userId}`);
    return { success: true };
  } catch (e) {
    writeError('removeUserFromChatGroup', e);
    return { success: false, error: e.message };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// MESSAGE REACTIONS
// ────────────────────────────────────────────────────────────────────────────
/** Add a reaction to a message */
function addMessageReaction(messageId, reaction, userId) {
  initializeChatService();
  const user = userId || getCurrentUser();
  const exists = readSheet(CHAT_MESSAGE_REACTIONS_SHEET)
    .some(r => r.MessageId === messageId && r.UserId === user && r.Reaction === reaction);
  if (exists) return { success: false, error: 'Already reacted' };
  const now = new Date(), id = generateId();
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHAT_MESSAGE_REACTIONS_SHEET);
    sh.appendRow([ id, messageId, user, reaction, now ]);
    writeDebug(`reaction_added:${messageId}|${reaction}`);
    return { success: true, reactionId: id };
  } catch (e) {
    writeError('addMessageReaction', e);
    return { success: false, error: e.message };
  }
}

/** Remove a reaction from a message */
function removeMessageReaction(messageId, reaction, userId) {
  initializeChatService();
  const user = userId || getCurrentUser();
  const rows = readSheet(CHAT_MESSAGE_REACTIONS_SHEET);
  const filtered = rows.filter(r => !(r.MessageId === messageId && r.UserId === user && r.Reaction === reaction));
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHAT_MESSAGE_REACTIONS_SHEET);
    const data = sh.getDataRange().getValues();
    const hdr  = data.shift();
    sh.clear(); sh.appendRow(hdr);
    if (filtered.length) {
      sh.getRange(2,1,filtered.length,filtered[0].length)
        .setValues(filtered.map(r => hdr.map(h => r[h])));
    }
    writeDebug(`reaction_removed:${messageId}|${reaction}`);
    return { success: true };
  } catch (e) {
    writeError('removeMessageReaction', e);
    return { success: false, error: e.message };
  }
}

// ────────────────────────────────────────────────────────────────────────────
// SEARCH
// ────────────────────────────────────────────────────────────────────────────
/** Search messages optionally within a group */
function searchChatMessages(query, groupId, limit = 50) {
  initializeChatService();
  const term = (query||'').trim().toLowerCase();
  if (term.length < 2) return { success: false, error: 'Query too short' };
  let msgs = readSheet(CHAT_MESSAGES_SHEET).filter(m => !m.IsDeleted);
  if (groupId) {
    const channelIds = getChatChannels(groupId).map(c => c.ID);
    msgs = msgs.filter(m => channelIds.includes(m.ChannelId));
  }
  const results = msgs
    .filter(m => m.Message.toLowerCase().includes(term))
    .sort((a,b)=> new Date(b.Timestamp) - new Date(a.Timestamp))
    .slice(0, limit)
    .map(m => {
      const ch = readSheet(CHAT_CHANNELS_SHEET).find(c=>c.ID===m.ChannelId) || {};
      const gr = readSheet(CHAT_GROUPS_SHEET).find(g=>g.ID===ch.GroupId) || {};
      return Object.assign({}, m, {
        ChannelName: ch.Name||'',
        GroupName:   gr.Name||''
      });
    });
  writeDebug(`search:${term}|found:${results.length}`);
  return { success: true, results, total: results.length };
}

// ────────────────────────────────────────────────────────────────────────────
// ANALYTICS & PREFERENCES
// ────────────────────────────────────────────────────────────────────────────
/** Log an analytics event */
function logChatActivity(action, details, userId) {
  try {
    const user = userId || getCurrentUser();
    const sess = Session.getTemporaryActiveUserKey() || '';
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHAT_ANALYTICS_SHEET);
    sh.appendRow([ new Date(), user, action, JSON.stringify(details), sess ]);
  } catch (e) {
    writeError('logChatActivity', e);
  }
}

/** Fetch or initialize user preferences */
function getUserPreferences(userId) {
  initializeChatService();
  const user = userId || getCurrentUser();
  return readSheet(CHAT_USER_PREFERENCES_SHEET)
    .find(p => p.UserId === user)
  || { UserId:user, NotificationSettings:'{}', Theme:'light', LastSeen:new Date(), Status:'active' };
}

/** Update user preferences */
function updateUserPreferences(updates, userId) {
  initializeChatService();
  const user = userId || getCurrentUser();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHAT_USER_PREFERENCES_SHEET);
  const data = sh.getDataRange().getValues(), hdr = data.shift();
  const idx = data.findIndex(r=>r[0]===user);
  const row = [
    user,
    updates.NotificationSettings || JSON.stringify({}),
    updates.Theme               || 'light',
    updates.LastSeen            || new Date(),
    updates.Status              || 'active'
  ];
  if (idx >= 0) data[idx] = row; else data.push(row);
  sh.clear(); sh.appendRow(hdr);
  sh.getRange(2,1,data.length,data[0].length).setValues(data);
  return { success: true };
}

/** Update last seen timestamp */
function updateUserLastSeen(userId) {
  const prefs = getUserPreferences(userId);
  return updateUserPreferences(Object.assign({}, prefs, { LastSeen: new Date() }), userId);
}

// ────────────────────────────────────────────────────────────────────────────
// CHAT SERVICE HELPER
// ────────────────────────────────────────────────────────────────────────────
/**
 * Returns an array of full group objects that the given userId is
 * actively a member of.  Relies on Utilities.gs:
 *   • initializeChatService()
 *   • readSheet(sheetName)
 */
function getUserChatGroups(userId) {
  // make sure all sheets exist
  initializeChatService();
  // get active memberships
  const memberships = readSheet(CHAT_GROUP_MEMBERS_SHEET)
    .filter(m => m.UserId === userId && m.IsActive);
  const groupIds = memberships.map(m => m.GroupId);
  if (!groupIds.length) return [];
  // fetch all groups once
  const allGroups = readSheet(CHAT_GROUPS_SHEET);
  // return only those matching
  return allGroups.filter(g => groupIds.includes(g.ID));
}
