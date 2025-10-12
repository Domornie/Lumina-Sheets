/**
 * tests/IdentityTests.gs
 * -----------------------------------------------------------------------------
 * Lightweight unit tests for critical Lumina Identity flows. Execute
 * `runIdentityTests()` from the Apps Script editor to validate OTP/TOTP,
 * RBAC, and lifecycle safeguards after deployment.
 */
function runIdentityTests() {
  var results = [];
  results.push(testTotpVerification());
  results.push(testRbacEnforcement());
  results.push(testLifecycleTerminationGuard());
  return results;
}

function testTotpVerification() {
  var secret = 'JBSWY3DPEHPK3PXP'; // Base32 for "Hello!"
  var originalNow = Date.now;
  try {
    Date.now = function() { return 0; };
    var expectedCode = generateTotp(secret, 0);
    var passed = AuthService.verifyTotpCode(secret, expectedCode);
    return formatTestResult('TOTP verification accepts valid codes', passed);
  } finally {
    Date.now = originalNow;
  }
}

function testRbacEnforcement() {
  var originalList = IdentityRepository.list;
  try {
    IdentityRepository.list = function(name) {
      if (name === 'RolePermissions') {
        return [
          { Role: 'Campaign Manager', Capability: 'VIEW_USERS', Scope: 'Campaign', Allowed: 'Y' },
          { Role: 'Guest (Client Owner)', Capability: 'VIEW_USERS', Scope: 'Campaign', Allowed: 'Y' }
        ];
      }
      if (name === 'UserCampaigns') {
        return [
          { AssignmentId: '1', UserId: 'user-1', CampaignId: 'camp-1', Role: 'Campaign Manager', IsPrimary: 'Y', Watchlist: 'N' }
        ];
      }
      return [];
    };
    var managerAllowed = RBACService.isAllowed('Campaign Manager', RBACService.CAPABILITIES.MANAGE_USERS, { scope: 'campaign' });
    var guestAllowed = RBACService.isAllowed('Guest (Client Owner)', RBACService.CAPABILITIES.MANAGE_USERS, { scope: 'campaign' });
    var passed = managerAllowed === false && guestAllowed === false;
    return formatTestResult('RBAC denies MANAGE_USERS for guests', passed);
  } finally {
    IdentityRepository.list = originalList;
  }
}

function testLifecycleTerminationGuard() {
  var originalAssert = RBACService.assertPermission;
  var originalEquipment = EquipmentService.hasOutstandingEquipment;
  var originalAppend = IdentityRepository.append;
  var originalAudit = AuditService.log;
  try {
    RBACService.assertPermission = function() { return true; };
    EquipmentService.hasOutstandingEquipment = function() { return true; };
    IdentityRepository.append = function() { throw new Error('Should not append when blocked'); };
    AuditService.log = function() {};
    var threw = false;
    try {
      UserService.updateLifecycle({ UserId: 'actor-1', Roles: ['Campaign Manager'], PrimaryRole: 'Campaign Manager' }, 'user-1', {
        CampaignId: 'camp-1',
        State: 'Terminated',
        EffectiveDate: new Date().toISOString()
      });
    } catch (err) {
      threw = err && err.message && err.message.indexOf('Outstanding equipment') >= 0;
    }
    return formatTestResult('Lifecycle termination blocked when equipment outstanding', threw);
  } finally {
    RBACService.assertPermission = originalAssert;
    EquipmentService.hasOutstandingEquipment = originalEquipment;
    IdentityRepository.append = originalAppend;
    AuditService.log = originalAudit;
  }
}

function generateTotp(secret, timestamp) {
  var counter = Math.floor(timestamp / 30000);
  var bytes = [];
  for (var i = 7; i >= 0; i--) {
    bytes[i] = counter & 0xff;
    counter = Math.floor(counter / 256);
  }
  var keyBytes = (function base32Decode(value) {
    var alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567';
    value = value.replace(/=+$/, '').toUpperCase();
    var bits = '';
    for (var i = 0; i < value.length; i++) {
      var idx = alphabet.indexOf(value.charAt(i));
      if (idx < 0) continue;
      bits += ('00000' + idx.toString(2)).slice(-5);
    }
    var out = [];
    for (var j = 0; j + 8 <= bits.length; j += 8) {
      out.push(parseInt(bits.slice(j, j + 8), 2));
    }
    return out;
  })(secret);
  var hmac = Utilities.computeHmacSha1Signature(bytes, keyBytes);
  var offset = hmac[hmac.length - 1] & 0x0f;
  var binary = ((hmac[offset] & 0x7f) << 24) |
    ((hmac[offset + 1] & 0xff) << 16) |
    ((hmac[offset + 2] & 0xff) << 8) |
    (hmac[offset + 3] & 0xff);
  var otp = (binary % 1000000).toString();
  while (otp.length < 6) {
    otp = '0' + otp;
  }
  return otp;
}

function formatTestResult(message, passed) {
  return { message: message, passed: !!passed };
}
