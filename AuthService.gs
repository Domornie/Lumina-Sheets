/**
 * AuthService.gs
 * -----------------------------------------------------------------------------
 * Comprehensive authentication and verification service for Lumina Identity.
 */
(function bootstrapAuthService(global) {
  if (!global) return;
  if (global.AuthService && typeof global.AuthService === 'object') {
    return;
  }

  function getIdentityRepository() {
    var repo = global.IdentityRepository;
    if (!repo || typeof repo.list !== 'function') {
      throw new Error('IdentityRepository not initialized');
    }
    return repo;
  }

  function getSessionService() {
    var sessionService = global.SessionService;
    if (!sessionService || typeof sessionService.issueSession !== 'function') {
      throw new Error('SessionService not initialized');
    }
    return sessionService;
  }

  function getAuditService() {
    var auditService = global.AuditService;
    if (!auditService || typeof auditService.log !== 'function') {
      throw new Error('AuditService not initialized');
    }
    return auditService;
  }
  var Utilities = global.Utilities;
  var MailApp = global.MailApp;
  var PropertiesService = global.PropertiesService;
  var EnterpriseSecurity = global.EnterpriseSecurity;

  var PASSWORD_REGEX = /^(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*[^A-Za-z0-9]).{8,}$/;
  var OTP_TTL_MS = 5 * 60 * 1000;
  var OTP_MAX_RESENDS = 3;
  var OTP_MAX_ATTEMPTS = 5;
  var LOGIN_RATE_LIMIT_1M = 5;
  var LOGIN_RATE_LIMIT_15M = 20;

  function getSecuritySalt() {
    var scriptProperties = PropertiesService && PropertiesService.getScriptProperties();
    if (!scriptProperties) {
      throw new Error('Script properties unavailable');
    }
    var salt = scriptProperties.getProperty('IDENTITY_PASSWORD_SALT');
    if (!salt) {
      salt = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid()));
      scriptProperties.setProperty('IDENTITY_PASSWORD_SALT', salt);
    }
    return salt;
  }

  function hashPassword(password) {
    if (!password) {
      throw new Error('Password required');
    }
    var salt = getSecuritySalt();
    var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_512, salt + '|' + password);
    return Utilities.base64Encode(digest);
  }

  function validatePassword(password) {
    if (!PASSWORD_REGEX.test(password)) {
      throw new Error('Password does not meet complexity requirements');
    }
  }

  function constantTimeEquals(a, b) {
    if (a === b) return true;
    if (!a || !b) return false;
    var strA = String(a);
    var strB = String(b);
    if (strA.length !== strB.length) {
      return false;
    }
    var diff = 0;
    for (var i = 0; i < strA.length; i++) {
      diff |= strA.charCodeAt(i) ^ strB.charCodeAt(i);
    }
    return diff === 0;
  }

  function findUser(emailOrUsername) {
    if (!emailOrUsername) return null;
    var normalized = String(emailOrUsername).toLowerCase();
    var users = getIdentityRepository().list('Users');
    return users.find(function(user) {
      return String(user.Email || '').toLowerCase() === normalized || String(user.Username || '').toLowerCase() === normalized;
    }) || null;
  }

  function generateOtpCode() {
    var code = Math.floor(100000 + Math.random() * 900000);
    return String(code);
  }

  function otpCacheKey(email, purpose) {
    return String(email).toLowerCase() + '::' + purpose;
  }

  function createOtp(email, purpose) {
    var now = Date.now();
    var key = otpCacheKey(email, purpose);
    var record = {
      Key: key,
      Email: email,
      Code: generateOtpCode(),
      Purpose: purpose,
      ExpiresAt: new Date(now + OTP_TTL_MS).toISOString(),
      Attempts: 0,
      LastSentAt: new Date(now).toISOString(),
      ResendCount: 0
    };
    getIdentityRepository().upsert('OTP', 'Key', record);
    return record;
  }

  function sendOtpEmail(email, code, purpose) {
    if (!MailApp) {
      console.log('MailApp unavailable – OTP code: ' + code);
      return;
    }
    MailApp.sendEmail({
      to: email,
      subject: 'Your Lumina Identity verification code',
      htmlBody: '<p>Your verification code is <strong>' + code + '</strong> (valid for 5 minutes) for ' + purpose + '.</p>'
    });
  }

  function enforceOtpRateLimit(email, purpose) {
    var existing = getIdentityRepository().find('OTP', function(row) {
      return row.Key === otpCacheKey(email, purpose);
    });
    if (!existing) {
      return;
    }
    var lastSent = new Date(existing.LastSentAt).getTime();
    if (Date.now() - lastSent < 60 * 1000) {
      throw new Error('OTP recently sent. Please wait before requesting another code.');
    }
    if (Number(existing.ResendCount || 0) >= OTP_MAX_RESENDS) {
      throw new Error('Maximum OTP resends reached. Contact support.');
    }
  }

  function requestOtp(emailOrUsername, purpose, context) {
    purpose = purpose || 'login';
    var user = findUser(emailOrUsername);
    if (!user) {
      throw new Error('User not found');
    }
    enforceLoginRateLimit(emailOrUsername, context && context.ip);
    enforceOtpRateLimit(user.Email, purpose);
    var otp = createOtp(user.Email, purpose);
    sendOtpEmail(user.Email, otp.Code, purpose);
    getIdentityRepository().upsert('OTP', 'Key', Object.assign({}, otp, {
      ResendCount: Number(otp.ResendCount || 0) + 1
    }));
    getAuditService().log({
      ActorUserId: user.UserId,
      ActorRole: 'SYSTEM',
      CampaignId: '',
      Target: user.Email,
      Action: 'OTP_SENT',
      IP: context && context.ip,
      UA: context && context.ua
    });
    return { success: true };
  }

  function verifyOtp(email, code, purpose, context) {
    var key = otpCacheKey(email, purpose);
    var record = getIdentityRepository().find('OTP', function(row) {
      return row.Key === key;
    });
    if (!record) {
      throw new Error('OTP not requested.');
    }
    var now = Date.now();
    if (new Date(record.ExpiresAt).getTime() < now) {
      getIdentityRepository().remove('OTP', 'Key', key);
      throw new Error('OTP expired.');
    }
    if (Number(record.Attempts || 0) >= OTP_MAX_ATTEMPTS) {
      throw new Error('Maximum OTP attempts exceeded.');
    }
    if (!constantTimeEquals(record.Code, code)) {
      getIdentityRepository().upsert('OTP', 'Key', Object.assign({}, record, {
        Attempts: Number(record.Attempts || 0) + 1
      }));
      getAuditService().log({
        ActorUserId: '',
        ActorRole: 'SYSTEM',
        CampaignId: '',
        Target: email,
        Action: 'OTP_INVALID',
        IP: context && context.ip,
        UA: context && context.ua
      });
      throw new Error('Invalid OTP code.');
    }
    getIdentityRepository().remove('OTP', 'Key', key);
    return true;
  }

  function enforceLoginRateLimit(identifier, ip) {
    var row = getIdentityRepository().find('LoginAttempts', function(item) {
      return item.EmailOrUsername === identifier;
    }) || {
      EmailOrUsername: identifier,
      Count1m: 0,
      Count15m: 0,
      LastAttemptAt: new Date(0).toISOString(),
      LockedUntil: ''
    };
    var now = Date.now();
    var lockedUntil = row.LockedUntil ? new Date(row.LockedUntil).getTime() : 0;
    if (lockedUntil && lockedUntil > now) {
      throw new Error('Temporarily locked — try again later.');
    }
    var lastAttempt = new Date(row.LastAttemptAt).getTime();
    var within1m = (now - lastAttempt) < 60 * 1000;
    var within15m = (now - lastAttempt) < 15 * 60 * 1000;
    row.Count1m = within1m ? Number(row.Count1m || 0) + 1 : 1;
    row.Count15m = within15m ? Number(row.Count15m || 0) + 1 : 1;
    row.LastAttemptAt = new Date(now).toISOString();
    if (row.Count1m > LOGIN_RATE_LIMIT_1M || row.Count15m > LOGIN_RATE_LIMIT_15M) {
      row.LockedUntil = new Date(now + 5 * 60 * 1000).toISOString();
      getIdentityRepository().upsert('LoginAttempts', 'EmailOrUsername', row);
      getAuditService().log({
        ActorUserId: '',
        ActorRole: 'SYSTEM',
        CampaignId: '',
        Target: identifier,
        Action: 'LOGIN_RATE_LIMITED',
        IP: ip || ''
      });
      throw new Error('Temporarily locked — try again at ' + new Date(row.LockedUntil).toLocaleTimeString());
    }
    getIdentityRepository().upsert('LoginAttempts', 'EmailOrUsername', row);
  }

  function resetLoginAttempts(identifier) {
    getIdentityRepository().remove('LoginAttempts', 'EmailOrUsername', identifier);
  }

  function login(payload, context) {
    context = context || {};
    var identifier = payload.emailOrUsername;
    enforceLoginRateLimit(identifier, context.ip);
    var user = findUser(identifier);
    if (!user) {
      throw new Error('Invalid credentials');
    }
    if (user.Status === 'Locked') {
      throw new Error('Account locked. Contact administrator.');
    }
    if (!payload.password) {
      throw new Error('Password required');
    }
    var hashed = hashPassword(payload.password);
    if (!constantTimeEquals(user.PasswordHash, hashed)) {
      getAuditService().log({
        ActorUserId: user.UserId,
        ActorRole: 'SYSTEM',
        CampaignId: '',
        Target: user.UserId,
        Action: 'LOGIN_FAILED',
        IP: context.ip,
        UA: context.ua
      });
      throw new Error('Invalid credentials');
    }
    if (payload.otp) {
      verifyOtp(user.Email, payload.otp, 'login', context);
    }
    if (user.TOTPEnabled === 'Y') {
      if (!payload.totp) {
        throw new Error('TOTP required');
      }
      if (!verifyTotp(user, payload.totp)) {
        throw new Error('Invalid TOTP');
      }
    }
    resetLoginAttempts(identifier);
    getIdentityRepository().upsert('Users', 'UserId', Object.assign({}, user, {
      LastLoginAt: new Date().toISOString()
    }));
    var primaryCampaign = getPrimaryCampaign(user.UserId);
    var session = getSessionService().issueSession(user, primaryCampaign ? primaryCampaign.CampaignId : '', context.ip, context.ua);
    getAuditService().log({
      ActorUserId: user.UserId,
      ActorRole: primaryCampaign ? primaryCampaign.Role : '',
      CampaignId: primaryCampaign ? primaryCampaign.CampaignId : '',
      Target: user.UserId,
      Action: 'LOGIN_SUCCESS',
      IP: context.ip,
      UA: context.ua,
      After: { sessionId: session.SessionId }
    });
    return session;
  }

  function getPrimaryCampaign(userId) {
    var assignments = getIdentityRepository().list('UserCampaigns').filter(function(row) {
      return row.UserId === userId;
    });
    if (!assignments.length) {
      return null;
    }
    var primary = assignments.find(function(row) { return row.IsPrimary === 'Y' || row.IsPrimary === true; });
    return primary || assignments[0];
  }

  function enableTotp(user, secret, code) {
    if (!secret) {
      throw new Error('TOTP secret required');
    }
    if (!code) {
      throw new Error('Verification code required');
    }
    if (!verifyTotpCode(secret, code)) {
      throw new Error('Unable to verify TOTP code');
    }
    var record = Object.assign({}, user, {
      TOTPEnabled: 'Y',
      TOTPSecretHash: encryptTotpSecret(user.UserId, secret)
    });
    getIdentityRepository().upsert('Users', 'UserId', record);
    getAuditService().log({
      ActorUserId: user.UserId,
      ActorRole: 'SYSTEM',
      CampaignId: '',
      Target: user.UserId,
      Action: 'TOTP_ENABLED'
    });
    return true;
  }

  function disableTotp(user) {
    var record = Object.assign({}, user, {
      TOTPEnabled: 'N',
      TOTPSecretHash: ''
    });
    getIdentityRepository().upsert('Users', 'UserId', record);
    getAuditService().log({
      ActorUserId: user.UserId,
      ActorRole: 'SYSTEM',
      CampaignId: '',
      Target: user.UserId,
      Action: 'TOTP_DISABLED'
    });
  }

  function base32Decode(secret) {
    var alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567';
    var padding = '=';
    secret = String(secret || '').replace(new RegExp(padding, 'g'), '').toUpperCase();
    var bits = '';
    for (var i = 0; i < secret.length; i++) {
      var val = alphabet.indexOf(secret.charAt(i));
      if (val === -1) {
        continue;
      }
      bits += ('00000' + val.toString(2)).slice(-5);
    }
    var result = [];
    for (var j = 0; j + 8 <= bits.length; j += 8) {
      result.push(parseInt(bits.slice(j, j + 8), 2));
    }
    return result;
  }

  function decryptTotpSecret(user) {
    if (!user.TOTPSecretHash) {
      return null;
    }
    if (EnterpriseSecurity && EnterpriseSecurity.isEncryptedValue(user.TOTPSecretHash)) {
      return EnterpriseSecurity.decryptValue(user.TOTPSecretHash, { tenant: user.UserId, purpose: 'TOTP' });
    }
    return user.TOTPSecretHash;
  }

  function encryptTotpSecret(userId, secret) {
    if (!EnterpriseSecurity) {
      return secret;
    }
    return EnterpriseSecurity.encryptValue(secret, { tenant: userId, purpose: 'TOTP' });
  }

  function verifyTotp(user, token) {
    var secret = decryptTotpSecret(user);
    if (!secret) {
      return false;
    }
    return verifyTotpCode(secret, token);
  }

  function verifyTotpCode(secret, token) {
    if (!secret || !token) {
      return false;
    }
    var secretBytes = base32Decode(secret);
    var timestep = Math.floor(Date.now() / 30000);
    for (var offset = -1; offset <= 1; offset++) {
      var counter = timestep + offset;
      var bytes = new Array(8);
      for (var i = 7; i >= 0; i--) {
        bytes[i] = counter & 0xff;
        counter = counter >> 8;
      }
      var hmac = Utilities.computeHmacSha1Signature(bytes, secretBytes);
      var offsetBits = hmac[hmac.length - 1] & 0x0f;
      var binary = ((hmac[offsetBits] & 0x7f) << 24) |
        ((hmac[offsetBits + 1] & 0xff) << 16) |
        ((hmac[offsetBits + 2] & 0xff) << 8) |
        (hmac[offsetBits + 3] & 0xff);
      var generated = (binary % 1000000).toString();
      while (generated.length < 6) {
        generated = '0' + generated;
      }
      if (constantTimeEquals(generated, token)) {
        return true;
      }
    }
    return false;
  }

  function logout(sessionId, context) {
    getSessionService().invalidateSession(sessionId);
    getAuditService().log({
      ActorUserId: context && context.userId,
      ActorRole: context && context.role,
      CampaignId: context && context.campaignId,
      Target: context && context.userId,
      Action: 'LOGOUT',
      IP: context && context.ip,
      UA: context && context.ua
    });
  }

  global.AuthService = {
    validatePassword: validatePassword,
    hashPassword: hashPassword,
    requestOtp: requestOtp,
    verifyOtp: verifyOtp,
    login: login,
    enableTotp: enableTotp,
    disableTotp: disableTotp,
    logout: logout,
    verifyTotpCode: verifyTotpCode
  };
})(GLOBAL_SCOPE);
