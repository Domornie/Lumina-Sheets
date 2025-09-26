/**
 * AuthenticationService.gs - Simplified Token-Based Authentication Service
 * Updated to use standard token-based session management
 * 
 * Features:
 * - Token-based session management
 * - Enhanced security with session tokens
 * - Email integration for password resets and confirmations
 * - Simple URL structure
 * - Improved error handling and logging
 */

// ───────────────────────────────────────────────────────────────────────────────
// AUTHENTICATION CONFIGURATION
// ───────────────────────────────────────────────────────────────────────────────

// Session TTL: 1 hour for regular sessions, 24 hours for remember me
const SESSION_TTL_MS = 60 * 60 * 1000; // 1 hour
const REMEMBER_ME_TTL_MS = 24 * 60 * 60 * 1000; // 24 hours

// ───────────────────────────────────────────────────────────────────────────────
// AUTHENTICATION SERVICE IMPLEMENTATION
// ───────────────────────────────────────────────────────────────────────────────

var AuthenticationService = (function () {
  
  // ─── Internal helpers ─────────────────────────────────────────────────────

  function getSS() {
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  function getOrCreateSheet(name, headers) {
    let sh = getSS().getSheetByName(name);
    if (!sh) {
      sh = getSS().insertSheet(name);
      sh.clear();
      if (headers && headers.length > 0) {
        sh.appendRow(headers);
      }
    }
    return sh;
  }

  function readTable(sheetName) {
    const sh = getOrCreateSheet(sheetName, []);
    const vals = sh.getDataRange().getValues();
    if (vals.length < 2) return [];
    const hdrs = vals.shift();
    return vals.map(row => {
      const obj = {};
      hdrs.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
  }

  function writeRow(sheetName, rowArr) {
    getOrCreateSheet(sheetName, []).appendRow(rowArr);
  }

  function updateRowByToken(sheetName, token, columnIndex, newValue) {
    const sh = getOrCreateSheet(sheetName, []);
    const data = sh.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(token)) {
        sh.getRange(i + 1, columnIndex + 1).setValue(newValue);
        return true;
      }
    }
    return false;
  }

  function deleteSessionRow(idx) {
    getOrCreateSheet(SESSIONS_SHEET, SESSIONS_HEADERS).deleteRow(idx + 2);
  }

  function cleanExpiredSessions() {
    const now = Date.now();
    const sessions = readTable(SESSIONS_SHEET);
    sessions.forEach((s, i) => {
      try {
        const expiry = new Date(s.ExpiresAt).getTime();
        if (expiry < now) {
          deleteSessionRow(i);
        }
      } catch (e) {
        console.warn('Error cleaning expired session:', e);
        deleteSessionRow(i);
      }
    });
  }

  function hashPwd(raw) {
    return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw)
      .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2))
      .join('');
  }

  function generateSecureToken() {
    return Utilities.getUuid() + '_' + Date.now();
  }

  // ─── Public API ────────────────────────────────────────────────────────────

  /** Ensure all identity & session sheets exist with proper headers */
  function ensureSheets() {
    try {
      getOrCreateSheet(USERS_SHEET, USERS_HEADERS);
      getOrCreateSheet(ROLES_SHEET, ROLES_HEADER);
      getOrCreateSheet(USER_ROLES_SHEET, USER_ROLES_HEADER);
      getOrCreateSheet(USER_CLAIMS_SHEET, CLAIMS_HEADERS);
      getOrCreateSheet(SESSIONS_SHEET, SESSIONS_HEADERS);
      console.log('Authentication sheets initialized successfully');
    } catch (error) {
      console.error('Error ensuring authentication sheets:', error);
      throw error;
    }
  }

  /** Find a user row by email (case-insensitive) */
  function getUserByEmail(email) {
    try {
      return readTable(USERS_SHEET)
        .find(u => String(u.Email).toLowerCase() === String(email || '').toLowerCase())
        || null;
    } catch (error) {
      console.error('Error getting user by email:', error);
      return null;
    }
  }

  /** Create a new session for a user ID, returning the session token */
  function createSessionFor(userId, existingToken = null, rememberMe = false) {
    try {
      const token = existingToken || generateSecureToken();
      const now = new Date();
      const ttl = rememberMe ? REMEMBER_ME_TTL_MS : SESSION_TTL_MS;
      const expiresAt = new Date(now.getTime() + ttl);

      // Get user agent and IP (limited in Apps Script environment)
      const userAgent = 'Google Apps Script';
      const ipAddress = 'N/A';

      writeRow(SESSIONS_SHEET, [
        token,
        userId,
        now.toISOString(),
        expiresAt.toISOString(),
        rememberMe,
        userAgent,
        ipAddress
      ]);

      console.log(`Session created for user ${userId}, expires: ${expiresAt.toISOString()}`);
      return token;
    } catch (error) {
      console.error('Error creating session:', error);
      return null;
    }
  }

  /**
   * Simplified login function 
   * @param {string} email - User email
   * @param {string} rawPwd - Plain text password
   * @param {boolean} rememberMe - Whether to create long-term session
   * @return {Object} Login result with detailed status information
   */
  function login(email, rawPwd, rememberMe = false) {
    try {
      if (!email || !rawPwd) {
        return {
          success: false,
          error: 'Email and password are required',
          errorCode: 'MISSING_CREDENTIALS'
        };
      }

      cleanExpiredSessions();
      const user = getUserByEmail(email);

      if (!user) {
        return {
          success: false,
          error: 'Invalid email or password',
          errorCode: 'INVALID_CREDENTIALS'
        };
      }

      // Check if user can login
      const canLogin = String(user.CanLogin).toUpperCase() === 'TRUE';
      if (!canLogin) {
        return {
          success: false,
          error: 'Your account has been disabled. Please contact support.',
          errorCode: 'ACCOUNT_DISABLED'
        };
      }

      // Check email confirmation
      const emailConfirmed = String(user.EmailConfirmed).toUpperCase() === 'TRUE';
      if (!emailConfirmed) {
        return {
          success: false,
          error: 'Please confirm your email address before logging in.',
          errorCode: 'EMAIL_NOT_CONFIRMED',
          needsEmailConfirmation: true
        };
      }

      // Check if password has been set
      const hasPassword = String(user.PasswordHash || '').length > 0;
      if (!hasPassword) {
        return {
          success: false,
          error: 'Please set up your password using the link from your welcome email.',
          errorCode: 'PASSWORD_NOT_SET',
          needsPasswordSetup: true
        };
      }

      // Verify password
      const storedHash = String(user.PasswordHash || '').toLowerCase();
      const providedHash = hashPwd(rawPwd).toLowerCase();

      if (storedHash !== providedHash) {
        return {
          success: false,
          error: 'Invalid email or password',
          errorCode: 'INVALID_CREDENTIALS'
        };
      }

      // Check if password reset is required
      const resetRequired = String(user.ResetRequired).toUpperCase() === 'TRUE';
      if (resetRequired) {
        // Create a temporary session for password reset
        const resetToken = createSessionFor(user.ID, null, false);
        return {
          success: false,
          error: 'You must change your password before continuing.',
          errorCode: 'PASSWORD_RESET_REQUIRED',
          resetToken: resetToken,
          needsPasswordReset: true
        };
      }

      // Create session
      const token = createSessionFor(user.ID, null, rememberMe);
      if (!token) {
        return {
          success: false,
          error: 'Failed to create session. Please try again.',
          errorCode: 'SESSION_CREATION_FAILED'
        };
      }

      // Update last login
      updateLastLogin(user.ID);

      // Get user roles and permissions
      const userRoles = getUserRoles(user.ID);
      const userClaims = getUserClaims(user.ID);

      return {
        success: true,
        sessionToken: token,
        user: {
          ID: user.ID,
          UserName: user.UserName,
          FullName: user.FullName,
          Email: user.Email,
          IsAdmin: user.IsAdmin,
          CampaignID: user.CampaignID,
          roles: userRoles,
          claims: userClaims,
          pages: getUserCampaignPages(user.ID, user.CampaignID)
        },
        message: 'Login successful'
      };

    } catch (error) {
      console.error('Login error:', error);
      writeError('AuthenticationService.login', error);
      return {
        success: false,
        error: 'An error occurred during login. Please try again.',
        errorCode: 'SYSTEM_ERROR'
      };
    }
  }

  function getUserCampaignPages(userId, campaignId) {
    try {
      if (!campaignId) return [];

      // Get campaign pages from campaign configuration
      const campaignPages = getCampaignPages(campaignId);
      return campaignPages
        .filter(cp => cp.IsActive !== false)
        .map(cp => cp.PageKey);
    } catch (e) {
      console.warn('Error getting user campaign pages:', e);
      return [];
    }
  }

  function getCampaignPages(campaignId) {
    try {
      // This would typically come from a CAMPAIGN_PAGES sheet
      // For now, return default pages
      return [
        { PageKey: 'dashboard', IsActive: true },
        { PageKey: 'reports', IsActive: true },
        { PageKey: 'qa', IsActive: true }
      ];
    } catch (error) {
      console.warn('Error getting campaign pages:', error);
      return [];
    }
  }

  /**
   * Simplified logout function that handles session cleanup
   * @param {string} sessionToken - Session token to invalidate
   * @return {Object} Logout result
   */
  function logout(sessionToken) {
    try {
      console.log('Logout initiated for token:', sessionToken ? sessionToken.substring(0, 8) + '...' : 'no token');
      
      let sessionRemoved = false;
      
      if (sessionToken) {
        sessionRemoved = _invalidateSessionByToken(sessionToken);
        console.log('Session invalidation result:', sessionRemoved);
      }

      return {
        success: true,
        sessionRemoved: sessionRemoved,
        message: 'Logged out successfully'
      };
    } catch (error) {
      console.error('Error during logout:', error);
      writeError('logout', error);
      return {
        success: false,
        error: error.message,
        message: 'Logout completed with errors'
      };
    }
  }

  /**
   * Validate & extend a session token (sliding expiration).
   * Returns the full user object (with roles/pages arrays) or null.
   */
  function getSessionUser(sessionToken) {
    try {
      if (!sessionToken) {
        return null;
      }

      cleanExpiredSessions();
      const sessions = readTable(SESSIONS_SHEET);
      const sessionRec = sessions.find(s => s.SessionToken === sessionToken);
      
      if (!sessionRec) {
        console.log('Session not found for token:', sessionToken.substring(0, 8) + '...');
        return null;
      }

      // Check if session is expired
      const expiryTime = new Date(sessionRec.ExpiresAt).getTime();
      const now = Date.now();
      
      if (expiryTime < now) {
        console.log('Session expired for token:', sessionToken.substring(0, 8) + '...');
        _invalidateSessionByToken(sessionToken);
        return null;
      }

      // Extend session expiry (sliding expiration)
      const isRememberMe = String(sessionRec.RememberMe).toUpperCase() === 'TRUE';
      const ttl = isRememberMe ? REMEMBER_ME_TTL_MS : SESSION_TTL_MS;
      const newExpiry = new Date(now + ttl);
      
      // Update session expiry
      updateRowByToken(SESSIONS_SHEET, sessionToken, 3, newExpiry.toISOString());

      // Load user record
      const userRow = readTable(USERS_SHEET).find(u => u.ID === sessionRec.UserId);
      if (!userRow) {
        console.warn('User not found for session:', sessionRec.UserId);
        _invalidateSessionByToken(sessionToken);
        return null;
      }

      // Parse roles and pages (handle both JSON and CSV formats)
      try {
        userRow.roles = userRow.Roles ?
          (userRow.Roles.startsWith('[') ?
            JSON.parse(userRow.Roles) :
            userRow.Roles.split(',').map(r => r.trim()).filter(Boolean)
          ) : [];
      } catch {
        userRow.roles = [];
      }

      try {
        userRow.pages = userRow.Pages ?
          (userRow.Pages.startsWith('[') ?
            JSON.parse(userRow.Pages) :
            userRow.Pages.split(',').map(p => p.trim()).filter(Boolean)
          ) : [];
      } catch {
        userRow.pages = [];
      }

      // Check if password reset is required
      userRow.needsReset = String(userRow.ResetRequired).toUpperCase() === 'TRUE';

      // Add session info
      userRow.sessionToken = sessionToken;
      userRow.sessionExpiry = newExpiry.toISOString();

      console.log('Session validated and extended for user:', userRow.Email);
      return userRow;

    } catch (error) {
      console.error('Error validating session:', error);
      writeError('getSessionUser', error);
      return null;
    }
  }

  function updateLastLogin(userId) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(USERS_SHEET);
      const data = sh.getDataRange().getValues();
      const headers = data[0];

      // Add LastLogin column if it doesn't exist
      let lastLoginIndex = headers.indexOf('LastLogin');
      if (lastLoginIndex === -1) {
        // Add the column
        sh.getRange(1, headers.length + 1).setValue('LastLogin');
        lastLoginIndex = headers.length;
      }

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === userId) {
          const row = i + 1;
          const col = lastLoginIndex + 1;

          const now = new Date();
          sh.getRange(row, col).setValue(now);
          sh.getRange(row, col).setNumberFormat('yyyy-mm-dd hh:mm:ss');
          break;
        }
      }
    } catch (error) {
      console.warn('Failed to update last login:', error);
    }
  }

  /**
   * Simplified keepAlive function 
   * @param {string} sessionToken - Session token to validate
   */
  function keepAlive(sessionToken) {
    try {
      if (!sessionToken) {
        return {
          success: false,
          expired: true,
          message: 'No session token provided'
        };
      }

      const user = getSessionUser(sessionToken);
      if (!user) {
        return {
          success: false,
          expired: true,
          message: 'Session expired or invalid'
        };
      }

      return {
        success: true,
        message: 'Session active',
        user: {
          ID: user.ID,
          FullName: user.FullName,
          Email: user.Email,
          CampaignID: user.CampaignID,
          IsAdmin: user.IsAdmin
        }
      };
    } catch (error) {
      console.error('Error in keepAlive:', error);
      writeError('keepAlive', error);
      return {
        success: false,
        error: error.message,
        message: 'Session check failed'
      };
    }
  }

  /**
   * Get user claims by user ID
   */
  function getUserClaims(userId) {
    try {
      const claims = readTable(USER_CLAIMS_SHEET);
      return claims.filter(claim => claim.UserId === userId);
    } catch (error) {
      console.error('Error getting user claims:', error);
      return [];
    }
  }

  /**
   * Get user roles by user ID
   */
  function getUserRoles(userId) {
    try {
      const userRoles = readTable(USER_ROLES_SHEET);
      const allRoles = readTable(ROLES_SHEET);

      const userRoleIds = userRoles
        .filter(ur => ur.UserId === userId)
        .map(ur => ur.RoleId);

      return allRoles.filter(role => userRoleIds.includes(role.ID));
    } catch (error) {
      console.error('Error getting user roles:', error);
      return [];
    }
  }

  /**
   * Simplified authentication requirement function
   */
  function requireAuth(e) {
    try {
      // Check for token parameter
      const token = e.parameter.token;
      if (token) {
        const user = getSessionUser(token);
        if (user) {
          return user;
        }
      }

      // Fall back to current user via Google session
      const user = getCurrentUser();
      if (!user || !user.ID) {
        return HtmlService
          .createTemplateFromFile('Login')
          .evaluate()
          .setTitle('Please Log In')
          .addMetaTag('viewport', 'width=device-width,initial-scale=1')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }

      const pageParam = (e.parameter.page || 'dashboard').toLowerCase();
      const allowed = (user.pages || []).map(p => p.toLowerCase());

      if (e.parameter.page
        && allowed.length > 0
        && allowed.indexOf(pageParam) < 0) {
        const tpl = HtmlService.createTemplateFromFile('AccessDenied');
        tpl.baseUrl = ScriptApp.getService().getUrl();
        return tpl
          .evaluate()
          .setTitle('Access Denied')
          .addMetaTag('viewport', 'width=device-width,initial-scale=1')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }

      return user;
    } catch (error) {
      console.error('Error in requireAuth:', error);
      writeError('requireAuth', error);
      throw error;
    }
  }

  /**
   * Enforce both auth & campaign match.
   * Returns user or throws.
   */
  function requireCampaignAuth(e) {
    const maybe = requireAuth(e);
    if (maybe.getContent) return maybe;
    const user = maybe;
    const cid = String(e.parameter.campaignId || '');
    if (!cid || cid !== String(user.CampaignID)) {
      throw new Error('Not authorized for campaign: ' + cid);
    }
    return user;
  }

  /**
   * Change password function
   * @param {string} sessionToken - Session token
   * @param {string} oldPassword - Current password
   * @param {string} newPassword - New password
   * @return {Object} Change result
   */
  function changePassword(sessionToken, oldPassword, newPassword) {
    try {
      const user = getSessionUser(sessionToken);
      if (!user) {
        return { success: false, message: 'Not authenticated.' };
      }

      // Validate new password strength
      if (!newPassword || newPassword.length < 8) {
        return { success: false, message: 'Password must be at least 8 characters long.' };
      }

      // Check password complexity
      const hasUpper = /[A-Z]/.test(newPassword);
      const hasLower = /[a-z]/.test(newPassword);
      const hasNumber = /[0-9]/.test(newPassword);
      const hasSpecial = /[^A-Za-z0-9]/.test(newPassword);

      if (!hasUpper || !hasLower || !hasNumber || !hasSpecial) {
        return {
          success: false,
          message: 'Password must contain at least one uppercase letter, one lowercase letter, one number, and one special character.'
        };
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(USERS_SHEET);
      const data = sh.getDataRange().getValues();
      const headers = data[0];

      const oldHash = hashPwd(oldPassword);
      const newHash = hashPwd(newPassword);
      const now = new Date();

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === user.ID) {
          const pwdColIndex = headers.indexOf('PasswordHash');
          const resetColIndex = headers.indexOf('ResetRequired');
          const updatedAtIndex = headers.indexOf('UpdatedAt');

          if (String(data[i][pwdColIndex]) !== oldHash) {
            return { success: false, message: 'Current password is incorrect.' };
          }

          const rowNum = i + 1;
          sh.getRange(rowNum, pwdColIndex + 1).setValue(newHash);
          sh.getRange(rowNum, resetColIndex + 1).setValue(false);
          if (updatedAtIndex >= 0) {
            sh.getRange(rowNum, updatedAtIndex + 1).setValue(now);
          }

          // Send password change confirmation email
          try {
            if (typeof sendPasswordChangeConfirmation === 'function') {
              sendPasswordChangeConfirmation(user.Email, { timestamp: now });
            }
          } catch (emailError) {
            console.warn('Failed to send password change email:', emailError);
          }

          return { success: true, message: 'Password changed successfully.' };
        }
      }

      return { success: false, message: 'User record not found.' };
    } catch (error) {
      console.error('Error changing password:', error);
      writeError('changePassword', error);
      return { success: false, message: 'An error occurred while changing password.' };
    }
  }

  /**
   * Set password with token function (for new users and password resets)
   * @param {string} token - Email confirmation or reset token
   * @param {string} newPassword - New password
   * @return {Object} Set password result
   */
  function setPasswordWithToken(token, newPassword) {
    try {
      // Validate password strength
      if (!newPassword || newPassword.length < 8) {
        return { success: false, message: 'Password must be at least 8 characters long.' };
      }

      const hasUpper = /[A-Z]/.test(newPassword);
      const hasLower = /[a-z]/.test(newPassword);
      const hasNumber = /[0-9]/.test(newPassword);
      const hasSpecial = /[^A-Za-z0-9]/.test(newPassword);

      if (!hasUpper || !hasLower || !hasNumber || !hasSpecial) {
        return {
          success: false,
          message: 'Password must contain uppercase, lowercase, number, and special character.'
        };
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(USERS_SHEET);
      const data = sh.getDataRange().getValues();
      const headers = data[0];

      const newHash = hashPwd(newPassword);
      const now = new Date();

      for (let i = 1; i < data.length; i++) {
        const tokenColIndex = headers.indexOf('EmailConfirmation'); // Password setup token

        if (String(data[i][tokenColIndex]) === token) {
          const rowNum = i + 1;
          const email = data[i][headers.indexOf('Email')];

          // Set password and update status
          sh.getRange(rowNum, headers.indexOf('PasswordHash') + 1).setValue(newHash);
          sh.getRange(rowNum, headers.indexOf('ResetRequired') + 1).setValue('FALSE');
          sh.getRange(rowNum, headers.indexOf('EmailConfirmation') + 1).setValue(''); // Clear token
          sh.getRange(rowNum, headers.indexOf('UpdatedAt') + 1).setValue(now);

          // Send confirmation email
          try {
            if (typeof sendPasswordChangeConfirmation === 'function') {
              sendPasswordChangeConfirmation(email, { timestamp: now });
            }
          } catch (emailError) {
            console.warn('Failed to send password confirmation email:', emailError);
          }

          return {
            success: true,
            message: 'Password set successfully. You can now log in.'
          };
        }
      }

      return { success: false, message: 'Invalid or expired setup link.' };
    } catch (error) {
      writeError('setPasswordWithToken', error);
      return { success: false, message: 'An error occurred while setting password.' };
    }
  }

  /**
   * Generate and send password reset token
   * @param {string} email - User email
   * @return {Object} Reset result
   */
  function requestPasswordReset(email) {
    try {
      const user = getUserByEmail(email);
      if (!user) {
        return {
          success: true,
          message: 'If an account with this email exists, a password reset link has been sent.'
        };
      }

      // Check if user has a password set
      if (!user.PasswordHash) {
        return {
          success: false,
          error: 'This account needs initial password setup. Check your welcome email.'
        };
      }

      // Generate reset token and send email
      const resetToken = Utilities.getUuid();

      // Update user record
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(USERS_SHEET);
      const data = sh.getDataRange().getValues();
      const headers = data[0];

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][headers.indexOf('Email')]).toLowerCase() === email.toLowerCase()) {
          const rowNum = i + 1;
          sh.getRange(rowNum, headers.indexOf('EmailConfirmation') + 1).setValue(resetToken);
          sh.getRange(rowNum, headers.indexOf('UpdatedAt') + 1).setValue(new Date());
          break;
        }
      }

      // Send reset email
      let emailSent = false;
      try {
        if (typeof sendPasswordResetEmail === 'function') {
          emailSent = sendPasswordResetEmail(email, resetToken);
        }
      } catch (emailError) {
        console.error('Error sending password reset email:', emailError);
      }

      return emailSent ?
        { success: true, message: 'Password reset email sent.' } :
        { success: false, error: 'Failed to send email. Please try again.' };

    } catch (error) {
      writeError('requestPasswordReset', error);
      return { success: false, error: 'An error occurred. Please try again later.' };
    }
  }

  /**
   * Resend password setup email for new users
   */
  function resendPasswordSetupEmail(email) {
    try {
      const user = getUserByEmail(email);
      if (!user) {
        return {
          success: true,
          message: 'If an account with this email exists, a password setup link has been sent.'
        };
      }

      // Check if user needs password setup
      const needsSetup = !user.PasswordHash || String(user.ResetRequired).toUpperCase() === 'TRUE';
      if (!needsSetup) {
        return {
          success: false,
          error: 'This account already has a password set. Use "Forgot Password" instead.'
        };
      }

      // Generate new setup token
      const setupToken = Utilities.getUuid();

      // Update user record with new token
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(USERS_SHEET);
      const data = sh.getDataRange().getValues();
      const headers = data[0];

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][headers.indexOf('Email')]).toLowerCase() === email.toLowerCase()) {
          const rowNum = i + 1;
          sh.getRange(rowNum, headers.indexOf('EmailConfirmation') + 1).setValue(setupToken);
          sh.getRange(rowNum, headers.indexOf('UpdatedAt') + 1).setValue(new Date());
          break;
        }
      }

      // Send email
      let emailSent = false;
      try {
        if (typeof sendPasswordSetupEmail === 'function') {
          emailSent = sendPasswordSetupEmail(email, {
            userName: user.UserName,
            fullName: user.FullName || user.UserName,
            passwordSetupToken: setupToken
          });
        }
      } catch (emailError) {
        console.error('Error sending password setup email:', emailError);
      }

      return emailSent ?
        { success: true, message: 'Password setup email sent successfully.' } :
        { success: false, error: 'Failed to send email. Please try again.' };

    } catch (error) {
      writeError('resendPasswordSetupEmail', error);
      return { success: false, error: 'An error occurred. Please try again later.' };
    }
  }

  /**
   * Log user activity for audit purposes
   * @param {string} sessionToken - Session token
   * @param {Object} activity - Activity data to log
   */
  function logUserActivity(sessionToken, activity) {
    try {
      const user = getSessionUser(sessionToken);
      if (!user) {
        throw new Error('User not authenticated');
      }

      // Ensure "UserActivityLog" sheet exists
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sh = ss.getSheetByName('UserActivityLog');
      if (!sh) {
        sh = ss.insertSheet('UserActivityLog');
        sh.appendRow(['Timestamp', 'UserEmail', 'ActivityPayload']);
      }

      // Append the audit entry
      sh.appendRow([
        new Date(),
        user.Email,
        JSON.stringify(activity, null, 2)
      ]);

      return { success: true };
    } catch (error) {
      console.error('Error logging user activity:', error);
      writeError('logUserActivity', error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Private helper to invalidate a session by token
   */
  function _invalidateSessionByToken(sessionToken) {
    try {
      if (!sessionToken) return false;
      
      const sh = getOrCreateSheet(SESSIONS_SHEET, SESSIONS_HEADERS);
      const data = sh.getDataRange().getValues();
      if (data.length < 2) return false;

      const headers = data[0];
      const tokenIdx = headers.indexOf('SessionToken');
      if (tokenIdx === -1) {
        console.error('SessionToken column not found in SESSIONS sheet');
        return false;
      }

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][tokenIdx]) === String(sessionToken)) {
          sh.deleteRow(i + 1);
          console.log('Session invalidated:', sessionToken.substring(0, 8) + '...');
          return true;
        }
      }
      
      console.log('Session not found for invalidation:', sessionToken.substring(0, 8) + '...');
      return false;
    } catch (error) {
      console.error('Error invalidating session:', error);
      return false;
    }
  }

  // Return public API
  return {
    ensureSheets,
    login,
    logout,
    getSessionUser,
    keepAlive,
    requireAuth,
    requireCampaignAuth,
    getUserByEmail,
    getUserClaims,
    getUserRoles,
    changePassword,
    setPasswordWithToken,
    requestPasswordReset,
    resendPasswordSetupEmail,
    logUserActivity,
    createSessionFor,
    _invalidateSessionByToken
  };
})();

// ───────────────────────────────────────────────────────────────────────────────
// CLIENT-ACCESSIBLE FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Simplified login function for standard authentication
 * Called by the client-side code via google.script.run
 */
function loginUser(email, password, rememberMe = false) {
  try {
    console.log('Server-side login for:', email);
    
    // Use AuthenticationService
    const result = AuthenticationService.login(email, password, rememberMe);
    
    if (!result.success) {
      return result;
    }

    // Create response with redirect URL including token
    const response = {
      success: true,
      message: 'Login successful',
      redirectUrl: SCRIPT_URL + '?page=dashboard&token=' + encodeURIComponent(result.sessionToken),
      sessionToken: result.sessionToken,
      user: result.user
    };

    return response;
  } catch (error) {
    console.error('Server login error:', error);
    writeError('loginUser', error);
    return {
      success: false,
      error: 'Login failed. Please try again.',
      errorCode: 'SYSTEM_ERROR'
    };
  }
}

/**
 * Simplified logout function
 */
function logoutUser(sessionToken) {
  try {
    console.log('Server-side logout');
    
    const result = AuthenticationService.logout(sessionToken);
    
    return {
      success: true,
      message: 'Logout successful',
      redirectUrl: SCRIPT_URL
    };
  } catch (error) {
    console.error('Server logout error:', error);
    writeError('logoutUser', error);
    return {
      success: false,
      error: error.message,
      redirectUrl: SCRIPT_URL
    };
  }
}

/**
 * Keep alive function for sessions
 */
function keepAliveSession(sessionToken) {
  try {
    return AuthenticationService.keepAlive(sessionToken);
  } catch (error) {
    console.error('Keep alive error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

// Expose individual functions for google.script.run
function getUserClaims(userId) {
  return AuthenticationService.getUserClaims(userId);
}

function getUserRoles(userId) {
  return AuthenticationService.getUserRoles(userId);
}

function setPasswordWithToken(token, newPassword) {
  return AuthenticationService.setPasswordWithToken(token, newPassword);
}

function changePassword(sessionToken, oldPassword, newPassword) {
  return AuthenticationService.changePassword(sessionToken, oldPassword, newPassword);
}

function requestPasswordReset(email) {
  return AuthenticationService.requestPasswordReset(email);
}

function resendPasswordSetupEmail(email) {
  return AuthenticationService.resendPasswordSetupEmail(email);
}

function login(email, password) {
  return AuthenticationService.login(email, password);
}

function logout(sessionToken) {
  return AuthenticationService.logout(sessionToken);
}

function keepAlive(sessionToken) {
  return AuthenticationService.keepAlive(sessionToken);
}

function clientLogUserActivity(sessionToken, activity) {
  return AuthenticationService.logUserActivity(sessionToken, activity);
}

function ensureAuthenticationSheets() {
  return AuthenticationService.ensureSheets();
}

// ───────────────────────────────────────────────────────────────────────────────
// UTILITY FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

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

/**
 * Get current user without requiring token
 * Uses Google Apps Script's built-in session management
 */
function getCurrentUser() {
  try {
    // Use Google's built-in user session
    const email = String(
      (Session.getActiveUser() && Session.getActiveUser().getEmail()) ||
      (Session.getEffectiveUser() && Session.getEffectiveUser().getEmail()) ||
      ''
    ).trim().toLowerCase();

    if (!email) {
      return null;
    }

    // Get user from your existing user management system
    const user = AuthenticationService.getUserByEmail(email);
    if (!user) {
      return null;
    }

    // Return user in expected format
    return {
      ID: user.ID,
      Email: user.Email,
      FullName: user.FullName,
      UserName: user.UserName,
      CampaignID: user.CampaignID,
      IsAdmin: String(user.IsAdmin).toUpperCase() === 'TRUE',
      CanLogin: String(user.CanLogin).toUpperCase() === 'TRUE',
      EmailConfirmed: String(user.EmailConfirmed).toUpperCase() === 'TRUE',
      ResetRequired: String(user.ResetRequired).toUpperCase() === 'TRUE'
    };
  } catch (error) {
    console.error('Error getting current user:', error);
    return null;
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// INITIALIZATION LOG
// ───────────────────────────────────────────────────────────────────────────────

console.log('Simplified AuthenticationService.gs loaded successfully');
console.log('Features: Token-based sessions, Enhanced security, Simple URLs, Comprehensive error handling');