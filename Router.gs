/**
 * Router.gs
 * -----------------------------------------------------------------------------
 * Minimal API router for Lumina Identity actions. The router expects an action
 * parameter and JSON body on POST requests. Responses are always JSON.
 */
var IdentityRouter = (function createIdentityRouter(global) {
  var SessionService = global.SessionService;
  var AuthService = global.AuthService;
  var UserService = global.UserService;
  var EquipmentService = global.EquipmentService;
  var PolicyService = global.PolicyService;
  var AuditService = global.AuditService;
  var IdentityRepository = global.IdentityRepository;

  function parseRequest(e) {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      try {
        body = JSON.parse(e.postData.contents);
      } catch (err) {
        throw new Error('Invalid JSON body');
      }
    }
    return {
      action: (e.parameter && e.parameter.action) || body.action,
      body: body,
      sessionId: (e.parameter && e.parameter.sessionId) || body.sessionId,
      csrf: (e.parameter && e.parameter.csrf) || body.csrf,
      campaignId: (e.parameter && e.parameter.campaignId) || body.campaignId,
      query: e.parameter || {},
      ip: (e && e.context && e.context.clientIp) || '',
      ua: (e && e.headers && e.headers['User-Agent']) || ''
    };
  }

  function jsonResponse(payload, status) {
    var output = ContentService.createTextOutput(JSON.stringify(payload));
    output.setMimeType(ContentService.MimeType.JSON);
    if (status) {
      output.setContent(JSON.stringify(Object.assign({ status: status }, payload)));
    }
    return output;
  }

  function handlePost(e) {
    try {
      var request = parseRequest(e);
      var result = dispatch(request, e);
      return jsonResponse({ ok: true, result: result });
    } catch (err) {
      console.error('IdentityRouter error', err);
      return jsonResponse({ ok: false, error: err.message || 'Request failed' });
    }
  }

  function handleGet(e) {
    try {
      var action = e.parameter && e.parameter.action;
      if (action === 'health') {
        return jsonResponse({ ok: true, version: '1.0.0', timestamp: new Date().toISOString() });
      }
      return jsonResponse({ ok: false, error: 'Unsupported GET action' });
    } catch (err) {
      return jsonResponse({ ok: false, error: err.message || 'Request failed' });
    }
  }

  function dispatch(request, rawEvent) {
    var action = (request.action || '').toLowerCase();
    switch (action) {
      case 'auth/request-otp':
        return AuthService.requestOtp(request.body.emailOrUsername, request.body.purpose, { ip: request.ip, ua: request.ua });
      case 'auth/login':
        return AuthService.login(request.body, { ip: request.ip, ua: request.ua });
      case 'auth/verify-otp':
        return AuthService.verifyOtp(request.body.email, request.body.code, request.body.purpose || 'login', { ip: request.ip, ua: request.ua });
      case 'auth/enable-totp':
        return AuthService.enableTotp(requireUserContext(request, rawEvent), request.body.secret, request.body.code);
      case 'auth/disable-totp':
        return AuthService.disableTotp(requireUserContext(request, rawEvent));
      case 'auth/logout':
        var context = requireSession(request);
        AuthService.logout(request.sessionId, { userId: context.user.UserId, role: '', campaignId: context.session.CampaignId, ip: request.ip, ua: request.ua });
        return true;
      case 'users/list':
        var actor = requireActor(request, rawEvent);
        return UserService.listUsers(actor, request.body.campaignId || actor.CampaignId);
      case 'users/create':
        return UserService.createUser(requireActor(request, rawEvent), request.body);
      case 'users/update':
        return UserService.updateUser(requireActor(request, rawEvent), request.body.userId, request.body);
      case 'users/transfer':
        return UserService.transferUser(requireActor(request, rawEvent), request.body.userId, request.body.toCampaignId);
      case 'users/lifecycle':
        return UserService.updateLifecycle(requireActor(request, rawEvent), request.body.userId, request.body);
      case 'equipment/assign':
        return EquipmentService.assignEquipment(requireActor(request, rawEvent), request.body);
      case 'equipment/update':
        return EquipmentService.updateEquipment(requireActor(request, rawEvent), request.body.equipmentId, request.body);
      case 'equipment/list':
        return EquipmentService.listEquipment(requireActor(request, rawEvent), request.body);
      case 'policies/list':
        return PolicyService.listPolicies(request.body.scope || 'Global');
      case 'audit/list':
        return AuditService.list(request.body);
      case 'health':
        return { ok: true, version: '1.0.0' };
      default:
        throw new Error('Unsupported action: ' + action);
    }
  }

  function requireActor(request, rawEvent) {
    var context = requireSession(request);
    var assignments = IdentityRepository.list('UserCampaigns').filter(function(row) {
      return row.UserId === context.user.UserId;
    });
    var primary = assignments.find(function(row) { return row.IsPrimary === 'Y' || row.IsPrimary === true; });
    return {
      UserId: context.user.UserId,
      Roles: assignments.map(function(row) { return row.Role; }),
      PrimaryRole: primary ? primary.Role : '',
      CampaignId: primary ? primary.CampaignId : '',
      session: context.session
    };
  }

  function requireUserContext(request) {
    var context = requireSession(request);
    return context.user;
  }

  function requireSession(request) {
    if (!request.sessionId) {
      throw new Error('Session required');
    }
    if (!SessionService.validateCsrf(request.sessionId, request.csrf)) {
      throw new Error('Invalid CSRF token');
    }
    var session = SessionService.readSession(request.sessionId);
    if (!session) {
      throw new Error('Session expired');
    }
    var user = IdentityRepository.find('Users', function(row) { return row.UserId === session.UserId; });
    if (!user) {
      throw new Error('User not found');
    }
    var renewed = SessionService.renewSession(session.SessionId);
    return { session: renewed || session, user: user };
  }

  return {
    handlePost: handlePost,
    handleGet: handleGet
  };
})(GLOBAL_SCOPE);

function doPost(e) {
  return IdentityRouter.handlePost(e);
}
