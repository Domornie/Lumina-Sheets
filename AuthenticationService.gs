/**
 * AuthenticationService.gs - Enhanced Authentication Service with Email Integration
 * Updated to work with comprehensive email system and improved error handling
 */

// 1-hour session TTL:
const SESSION_TTL_MS = 60 * 60 * 1000;

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
      sh.appendRow(headers);
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

  function deleteSessionRow(idx) {
    getOrCreateSheet(SESSIONS_SHEET, []).deleteRow(idx + 2);
  }

  function cleanExpiredSessions() {
    const now = Date.now();
    const sessions = readTable(SESSIONS_SHEET);
    sessions.forEach((s, i) => {
      if (new Date(s.ExpiresAt).getTime() < now) {
        deleteSessionRow(i);
      }
    });
  }

  function hashPwd(raw) {
    return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw)
      .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2))
      .join('');
  }

  // ─── Public API ────────────────────────────────────────────────────────────

  /** Ensure all identity & session sheets exist with proper headers */
  function ensureSheets() {
    getOrCreateSheet(USERS_SHEET, USERS_HEADERS);
    getOrCreateSheet(ROLES_SHEET, ROLES_HEADER);
    getOrCreateSheet(USER_ROLES_SHEET, USER_ROLES_HEADER);
    getOrCreateSheet(USER_CLAIMS_SHEET, CLAIMS_HEADERS);
    getOrCreateSheet(SESSIONS_SHEET, SESSIONS_HEADERS);
  }

  /** Find a user row by email (case-insensitive) */
  function getUserByEmail(email) {
    return readTable(USERS_SHEET)
      .find(u => String(u.Email).toLowerCase() === String(email || '').toLowerCase())
      || null;
  }

  /** Create a new session for a user ID, returning the session token */
  function createSessionFor(userId) {
    const token = Utilities.getUuid();
    const now = new Date();
    const expiresAt = new Date(now.getTime() + SESSION_TTL_MS);
    writeRow(SESSIONS_SHEET, [
      token,
      userId,
      now.toISOString(),
      expiresAt.toISOString()
    ]);
    return token;
  }

  /**
   * Enhanced login function with comprehensive error handling and email integration
   * @param {string} email - User email
   * @param {string} rawPwd - Plain text password
   * @return {Object} Login result with detailed status information
   */
  function login(email, rawPwd) {
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

      // Create session
      const token = createSessionFor(user.ID);

      // Update last login
      updateLastLogin(user.ID);

      return {
        success: true,
        token: token,
        user: {
          ID: user.ID,
          UserName: user.UserName,
          FullName: user.FullName,
          Email: user.Email,
          IsAdmin: user.IsAdmin,
          CampaignID: user.CampaignID,
          roles: getUserRoleIds(user.ID),
          pages: getUserCampaignPages(user.ID, user.CampaignID)
        },
        message: 'Login successful'
      };

    } catch (error) {
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

      const campaignPages = getCampaignPages(campaignId);
      return campaignPages
        .filter(cp => cp.IsActive)
        .map(cp => cp.PageKey);
    } catch (e) {
      writeError('getUserCampaignPages', e);
      return [];
    }
  }

  /**
   * Enhanced logout function that handles session cleanup and redirect
   * @param {string} token - Session token to invalidate
   * @return {Object} Logout result with redirect information
   */
  // REPLACE the existing logout(token) with this
  function logout(token) {
    try {
      const removed = _invalidateSessionByToken(token);
      const baseUrl = ScriptApp.getService().getUrl(); // no token => login page
      return {
        success: true,
        sessionRemoved: removed,
        redirectUrl: baseUrl,
        message: 'Logged out successfully'
      };
    } catch (error) {
      console.error('Error during logout:', error);
      writeError('logout', error);
      const baseUrl = ScriptApp.getService().getUrl();
      return {
        success: false,
        error: error.message,
        redirectUrl: baseUrl,
        message: 'Logout completed with errors'
      };
    }
  }


  /**
   * Handle logout via doGet (for direct logout links)
   * @param {Object} e - Event object from doGet
   */
  function handleLogoutRequest(e) {
    const token = e.parameter.token || '';
    const SCRIPT_URL = 'https://script.google.com/a/macros/vlbpo.com/s/AKfycbxeQ0AnupBHM71M6co3LVc5NPrxTblRXLd6AuTOpxMs2rMehF9dBSkGykIcLGHROywQ/exec';

    try {
      // Invalidate the session
      if (token) {
        _invalidateSessionByToken(token);
      }

      // Create logout confirmation page with auto-redirect
      const html = `
            <!DOCTYPE html>
            <html>
            <head>
                <title>Logging Out...</title>
                <meta name="viewport" content="width=device-width,initial-scale=1">
                <link rel="icon" type="image/png" href="https://res.cloudinary.com/dr8qd3xfc/image/upload/v1754763514/vlbpo/lumina/3_dgitcx.png">
                <style>
                    body {
                        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                        min-height: 100vh;
                        display: flex;
                        align-items: center;
                        justify-content: center;
                        margin: 0;
                        color: #333;
                    }
                    .logout-container {
                        background: white;
                        border-radius: 20px;
                        box-shadow: 0 20px 60px rgba(0, 0, 0, 0.1);
                        padding: 3rem;
                        text-align: center;
                        max-width: 400px;
                        position: relative;
                        overflow: hidden;
                    }
                    .logout-container::before {
                        content: '';
                        position: absolute;
                        top: 0;
                        left: 0;
                        right: 0;
                        height: 4px;
                        background: linear-gradient(90deg, #ff6b6b, #feca57, #48dbfb, #ff9ff3);
                    }
                    .logout-icon {
                        font-size: 3rem;
                        color: #27ae60;
                        margin-bottom: 1rem;
                        animation: pulse 1.5s infinite;
                    }
                    @keyframes pulse {
                        0% { transform: scale(1); opacity: 1; }
                        50% { transform: scale(1.1); opacity: 0.7; }
                        100% { transform: scale(1); opacity: 1; }
                    }
                    h1 {
                        color: #2c3e50;
                        margin-bottom: 1rem;
                        font-weight: 300;
                    }
                    .message {
                        color: #7f8c8d;
                        margin-bottom: 2rem;
                        line-height: 1.6;
                    }
                    .spinner {
                        border: 3px solid #f3f3f3;
                        border-top: 3px solid #667eea;
                        border-radius: 50%;
                        width: 30px;
                        height: 30px;
                        animation: spin 1s linear infinite;
                        margin: 0 auto 1rem;
                    }
                    @keyframes spin {
                        0% { transform: rotate(0deg); }
                        100% { transform: rotate(360deg); }
                    }
                    .redirect-link {
                        color: #667eea;
                        text-decoration: none;
                        font-weight: 500;
                    }
                    .redirect-link:hover {
                        text-decoration: underline;
                    }
                </style>
            </head>
            <body>
                <div class="logout-container">
                    <div class="logout-icon">👋</div>
                    <h1>Logging Out...</h1>
                    <div class="spinner"></div>
                    <p class="message">
                        You have been successfully logged out.<br>
                        Redirecting you to the login page...
                    </p>
                    <p>
                        <a href="${SCRIPT_URL}" class="redirect-link">
                            Click here if you're not redirected automatically
                        </a>
                    </p>
                </div>
                
                <script>
                    // Clear any cached data
                    if (typeof(Storage) !== "undefined") {
                        localStorage.clear();
                        sessionStorage.clear();
                    }
                    
                    // Redirect after 2 seconds
                    setTimeout(function() {
                        window.location.href = "${SCRIPT_URL}";
                    }, 2000);
                </script>
            </body>
            </html>
            `;

      return HtmlService.createHtmlOutput(html)
        .setTitle('Logging Out - VLBPO LuminaHQ')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    } catch (error) {
      console.error('Error in logout request:', error);
      writeError('handleLogoutRequest', error);

      // Even on error, redirect to login
      return HtmlService.createHtmlOutput(`
                <script>window.location.href = "${SCRIPT_URL}";</script>
            `).setTitle('Redirecting...');
    }
  }

  /**
   * Client-accessible logout function for google.script.run
   * @param {string} token - Session token to invalidate
   */
  function clientLogout(token) {
    try {
      const result = logout(token);

      // Always ensure we have a redirect URL
      if (!result.redirectUrl) {
        const SCRIPT_URL = 'https://script.google.com/a/macros/vlbpo.com/s/AKfycbxeQ0AnupBHM71M6co3LVc5NPrxTblRXLd6AuTOpxMs2rMehF9dBSkGykIcLGHROywQ/exec';
        result.redirectUrl = SCRIPT_URL;
      }

      return result;

    } catch (error) {
      console.error('Error in client logout:', error);
      writeError('clientLogout', error);

      const SCRIPT_URL = 'https://script.google.com/a/macros/vlbpo.com/s/AKfycbxeQ0AnupBHM71M6co3LVc5NPrxTblRXLd6AuTOpxMs2rMehF9dBSkGykIcLGHROywQ/exec';

      return {
        success: false,
        error: error.message,
        redirectUrl: SCRIPT_URL,
        message: 'Logout failed but redirecting to login'
      };
    }
  }

  /**
   * Validate & extend a session token (sliding expiration).
   * Returns the full user object (with roles/pages arrays) or null.
   */
  function getSessionUser(token) {
    cleanExpiredSessions();
    const sessions = readTable(SESSIONS_SHEET);
    const rec = sessions.find(s => s.Token === token);
    if (!rec) return null;

    // extend expiry
    const newExp = new Date(Date.now() + SESSION_TTL_MS).toISOString();
    const sheet = getOrCreateSheet(SESSIONS_SHEET, SESSIONS_HEADERS);
    const rowIdx = sessions.indexOf(rec) + 2;
    sheet.getRange(rowIdx, 4).setValue(newExp);

    // load user record
    const userRow = readTable(USERS_SHEET).find(u => u.ID === rec.UserId);
    if (!userRow) return null;

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

    return userRow;
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

          const now = new Date(); // date + time
          sh.getRange(row, col).setValue(now);
          sh.getRange(row, col).setNumberFormat('yyyy-mm-dd hh:mm:ss'); // show time too
          break;
        }
      }
    } catch (error) {
      console.warn('Failed to update last login:', error);
    }
  }

  /**
   * Enhanced keepAlive function that handles session expiration
   * @param {string} token - Session token to validate
   */
  // REPLACE the existing keepAlive(token) with this
  function keepAlive(token) {
    try {
      // Use local getSessionUser (extends expiry) instead of calling AuthenticationService.keepAlive
      const user = getSessionUser(token);
      if (!user) {
        return {
          success: false,
          expired: true,
          redirectUrl: ScriptApp.getService().getUrl(),
          message: 'Session expired'
        };
      }
      return { success: true, message: 'Session active' };
    } catch (error) {
      console.error('Error in keepAlive:', error);
      writeError('keepAlive', error);
      return {
        success: false,
        error: error.message,
        redirectUrl: ScriptApp.getService().getUrl(),
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
   * Enforce authentication & page-access.
   * Returns a user object or an HtmlOutput for Login/AccessDenied.
   */
  function requireAuth(e) {
    const token = e.parameter.token || '';
    const user = token && getSessionUser(token);
    if (!user) {
      return HtmlService
        .createTemplateFromFile('Login')
        .evaluate()
        .setTitle('Please Log In')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    const pageParam = (e.parameter.page || 'Search').toLowerCase();
    const allowed = (user.pages || []).map(p => p.toLowerCase());

    if (e.parameter.page
      && allowed.length > 0
      && allowed.indexOf(pageParam) < 0) {
      const tpl = HtmlService.createTemplateFromFile('AccessDenied');
      tpl.baseUrl = ScriptApp.getService().getUrl()
        + '?token=' + encodeURIComponent(token);
      return tpl
        .evaluate()
        .setTitle('Access Denied')
        .addMetaTag('viewport', 'width=device-width,initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    return user;
  }

  /**
   * Enforce both auth & campaign match.
   * Returns user or throws.
   */
  function requireCampaignAuth(e) {
    const maybe = requireAuth(e);
    if (maybe.getContent) return maybe;
    const user = /** @type {!Object} */(maybe);
    const cid = String(e.parameter.campaignId || '');
    if (!cid || cid !== String(user.CampaignID)) {
      throw new Error('Not authorized for campaign: ' + cid);
    }
    return user;
  }

  /**
   * Enhanced change password function with email notification
   * @param {string} token - Session token
   * @param {string} oldPassword - Current password
   * @param {string} newPassword - New password
   * @return {Object} Change result
   */
  function changePassword(token, oldPassword, newPassword) {
    try {
      const user = getSessionUser(token);
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
            if (typeof onPasswordChanged === 'function') {
              onPasswordChanged(user.Email, { timestamp: now });
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

  function resendPasswordSetupEmail(email) {
    try {
      const user = getUserByEmail(email);
      if (!user) {
        // Don't reveal if email exists for security
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
      const emailSent = sendPasswordSetupEmail(email, {
        userName: user.UserName,
        fullName: user.FullName || user.UserName,
        passwordSetupToken: setupToken
      });

      invalidateCache(USERS_SHEET);

      return emailSent ?
        { success: true, message: 'Password setup email sent successfully.' } :
        { success: false, error: 'Failed to send email. Please try again.' };

    } catch (error) {
      writeError('resendPasswordSetupEmail', error);
      return { success: false, error: 'An error occurred. Please try again later.' };
    }
  }

  /**
   * Enhanced set password with token function (for new users and password resets)
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
          sh.getRange(rowNum, headers.indexOf('ResetRequired') + 1).setValue('FALSE'); // No longer needs setup
          sh.getRange(rowNum, headers.indexOf('EmailConfirmation') + 1).setValue(''); // Clear token
          sh.getRange(rowNum, headers.indexOf('UpdatedAt') + 1).setValue(now);

          invalidateCache(USERS_SHEET);

          // Send confirmation email
          try {
            sendPasswordChangeConfirmation(email, { timestamp: now });
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
      const emailSent = sendPasswordResetEmail(email, resetToken);

      invalidateCache(USERS_SHEET);

      return emailSent ?
        { success: true, message: 'Password reset email sent.' } :
        { success: false, error: 'Failed to send email. Please try again.' };

    } catch (error) {
      writeError('requestPasswordReset', error);
      return { success: false, error: 'An error occurred. Please try again later.' };
    }
  }

  /**
   * logUserActivity(token, activity)
   *
   * Records a timestamped audit of the authenticated user’s activity payload
   * into the “UserActivityLog” sheet.
   *
   * @param {string} token    Session token
   * @param {Object} activity Any serializable data describing the user’s action
   * @return {Object}         { success: true } on success
   */
  function logUserActivity(token, activity) {
    // Authenticate & fetch user
    const user = AuthenticationService.getSessionUser(token);
    if (!user) {
      throw new Error('User not authenticated');
    }

    // Ensure “UserActivityLog” sheet exists
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
  }

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
    logUserActivity
  };
})();

// Add this private helper inside the IIFE (near the other internals)
function _invalidateSessionByToken(token) {
  if (!token) return false;
  const sh = getOrCreateSheet(SESSIONS_SHEET, SESSIONS_HEADERS);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return false;

  const headers = data[0];
  const tokenIdx = headers.indexOf('Token');         // expected header
  if (tokenIdx === -1) throw new Error('SESSIONS_HEADERS must include "Token".');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][tokenIdx]) === String(token)) {
      sh.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

// Expose functions for google.script.run
function getUserClaims(userId) {
  return AuthenticationService.getUserClaims(userId);
}

function getUserRoles(userId) {
  return AuthenticationService.getUserRoles(userId);
}

function setPasswordWithToken(token, newPassword) {
  return AuthenticationService.setPasswordWithToken(token, newPassword);
}

function changePassword(token, oldPassword, newPassword) {
  return AuthenticationService.changePassword(token, oldPassword, newPassword);
}

function requestPasswordReset(email) {
  return AuthenticationService.requestPasswordReset(email);
}

// Enhanced login function for client access
function login(email, password) {
  return AuthenticationService.login(email, password);
}

/**
 * clientLogUserActivity(token, activity)
 *
 * Exposed to client via google.script.run to record user activity.
 */
function clientLogUserActivity(token, activity) {
  return AuthenticationService.logUserActivity(token, activity);
}