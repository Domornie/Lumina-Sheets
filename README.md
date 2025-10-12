# Lumina-Sheets

Call center management system built on Google Apps Script + Google Sheets.

## Lumina Identity platform

The repository now bundles **Lumina Identity**, a tenant-aware security layer that
implements authentication, OTP/TOTP verification, RBAC, campaign isolation, and
full employment lifecycle tracking. The identity stack lives in the new `*.gs`
services (`AuthService`, `SessionService`, `RBACService`, etc.) plus a set of
HTML front-ends inside `Html/`.

### Prerequisites

1. Create a dedicated Google Sheet to store the identity tables. Add tabs with
   the exact headers defined in `IdentityRepository.TABLE_HEADERS`. The
   bootstrap helpers will auto-create sheets that are missing.
2. (Optional for seeding) Set the following script properties in the Apps Script
   project:
   - `IDENTITY_SPREADSHEET_ID`: ID of the sheet created in step 1. When omitted
     the seeding helpers fall back to the active spreadsheet, but production
     deployments should still configure this value explicitly.
   - `IDENTITY_PASSWORD_SALT`: optional, if omitted a random salt is generated
     on first login.
3. Deploy the web app with `Execute as: User accessing the web app` and
   `Who has access: Anyone`. The router enforces per-request permissions.

### Bootstrap & seed data

- Run `seedDefaultData()` to bootstrap the workspace. The helper now also calls
  `seedLuminaIdentity()` under the hood so the identity roles, permissions, and
  `System Admin` account are provisioned automatically. Update the default
  password immediately after the first login.
- The bootstrap flow now mirrors the seeded administrator into the legacy
  `Users` directory tab so the default Apps Script container sheet immediately
  shows the account and its assigned metadata.
- Use the `/auth/request-otp` endpoint (exposed via the router) to verify email
  delivery through Apps Script `MailApp` or your preferred SMTP relay.

### API overview

`Router.gs` exposes a JSON API over `doPost`/`action` for
authentication (`auth/login`, `auth/request-otp`, `auth/enable-totp`), user
administration (`users/list`, `users/create`, `users/transfer`, `users/lifecycle`),
equipment control, policies, and audit retrieval. All state-changing requests
require a valid session plus CSRF token (returned with every login response).

For front-end consumption, the repository ships new HtmlService templates:

- `Html/LuminaIdentityLanding.html` – public landing & marketing hero.
- `Html/LuminaIdentityLogin.html` – password + OTP/TOTP login interface.
- `Html/LuminaIdentityApp.html` – authenticated workspace with campaign
  directory, equipment tracker, policy viewer, and audit timeline.

Integrate the views via `HtmlService.createTemplateFromFile` as needed for your
deployment. The login page expects the API to be served from the same Apps Script
endpoint so fetch requests can POST directly.

## Schema change guidelines

- When adding functionality that requires a new column—either in the web app UI
  definitions or in the backing Google Sheets—add the new column explicitly
  instead of renaming an existing one. Renaming breaks historical data bindings
  and cached lookups, so always create a fresh column with the new object name
  and migrate data deliberately if needed.

## Frontend lazy-loading harness

- Every layout now includes `EntityLoader.html`, which exposes a global
  `LuminaEntityLoader` helper for registering async entities with shared
  skeletons and a promise-based `google.script.run` wrapper.
- Pilot implementation lives in `QualityForm.html`; additional feature pages can
  follow the same pattern to hydrate UI fragments once data returns.
- Refer to [`docs/lazy-loading.md`](docs/lazy-loading.md) for usage examples,
  skeleton registration, and helper APIs.

## Trigger maintenance

- `checkRealtimeUpdatesJob()` now self-throttles by default, limiting each run to
  one minute and pausing at least five minutes between executions. Adjust the
  cadence by setting script properties such as
  `REALTIME_JOB_MAX_RUNTIME_MS`, `REALTIME_JOB_MIN_INTERVAL_MS`, or
  `REALTIME_JOB_SLEEP_MS`.
- Run `listProjectTriggers()` from the Apps Script editor if you suspect a
  legacy time-driven job is still active. The helper logs every trigger and the
  handler function it tries to invoke.
- Use `removeLegacyRealtimeTrigger()` only when you intentionally want to delete
  the realtime job trigger (for example, before replacing it with a new
  schedule).

## Google Sheets database manager

The `DatabaseManager.gs` module turns any worksheet into a CRUD-ready table. Define
schema defaults once and re-use the same interface across every client campaign.

> **Compatibility note:** shared helpers such as `ensureSheetWithHeaders`,
> `readSheet`, and `invalidateCache` are wrapped so they prefer the centralized
> database logic while falling back to any legacy implementations already loaded
> in your project. This eliminates Apps Script "function already defined"
> conflicts when combining the database layer with existing utility files.

### Quick start

```javascript
const users = DatabaseManager.defineTable('Users', {
  headers: ['ID', 'UserName', 'Email', 'CampaignID'],
  defaults: {},
});

// Create
const user = users.insert({
  UserName: 'jsmith',
  Email: 'jsmith@example.com',
  CampaignID: 'credit-suite'
});

// Read
const perCampaign = users.find({ where: { CampaignID: 'credit-suite' } });

// Update
users.update(user.ID, { IsActive: false });

// Delete
users.delete(user.ID);
```

The manager automatically ensures each sheet exists, appends missing headers, and
stores `CreatedAt`/`UpdatedAt` timestamps for every record. Cached reads keep lookups
fast while still honoring sheet edits from other services.

### Global CRUD helpers

`DatabaseBindings.gs` registers the common sheets used across the call center platform
and exposes lightweight helpers so existing Apps Script functions can switch to the
database abstraction without rewriting business logic:

```javascript
// Read data with optional filters/sorting/pagination
const activeUsers = dbSelect(USERS_SHEET, {
  where: { CampaignID: campaignId },
  sortBy: 'FullName'
});

// Create/update/delete
const created = dbCreate(USERS_SHEET, payload);
const updated = dbUpdate(USERS_SHEET, created.ID, { ResetRequired: false });
dbDelete(USERS_SHEET, created.ID);

// Upsert by any condition (automatically creates IDs when needed)
dbUpsert(CAMPAIGN_USER_PERMISSIONS_SHEET, { UserID, CampaignID }, {
  PermissionLevel: 'Manager',
  CanManageUsers: true
});
```

The existing `readSheet`/`ensureSheetWithHeaders` helpers now rely on these bindings,
so all services automatically share cached queries and schema registration through
`DatabaseManager`.

### Multi-tenant SaaS safeguards

Every campaign operates as an isolated tenant. Tables that contain campaign data are
now registered with a `tenantColumn`, so `DatabaseManager` automatically blocks reads
and writes that target campaigns outside the active tenant context.

```javascript
// Create a tenant-scoped CRUD context for the signed-in manager
const { context } = getTenantContextForUser(currentUserId, 'credit-suite');
const campaignPages = DatabaseManager.table(CAMPAIGN_PAGES_SHEET, context);

// All queries will be transparently filtered to the campaigns in `context`
const navigation = campaignPages.find({ sortBy: 'SortOrder' });
```

Use `TenantSecurityService.gs` to calculate per-user access profiles, enforce
manager/admin permissions, and obtain reusable CRUD contexts:

```javascript
const { profile, context } = TenantSecurity.getTenantContext(userId, campaignId);
// Throws if the user cannot access the campaign

// One-line helpers for Apps Script services
const db = dbWithContext(context);
const notifications = db.select(NOTIFICATIONS_SHEET, { where: { UserId: userId } });

// When you need a DatabaseManager table instance:
const scopedUsers = TenantSecurity.getScopedTable(userId, campaignId, USERS_SHEET);
```

Legacy helpers such as `dbSelect` and `dbCreate` accept an optional tenant context as
their final argument, while dedicated helpers (`dbTenantSelect`, `dbTenantCreate`,
etc.) provide a more explicit API.

### Authentication & campaign-scoped sessions

Lumina Identity replaces the legacy authentication layer. Passwords are hashed
with per-project salts, OTPs expire within five minutes, TOTP secrets are stored
encrypted, and all sessions are short-lived (10–30 minute sliding window)
backed by CacheService/PropertiesService. Every login, OTP issuance, lifecycle
change, transfer, and equipment update writes an immutable entry to the
`AuditLog` sheet for forensic review.


## End-to-end call center workflows

`CallCenterWorkflowService.gs` stitches together authentication, scheduling,
performance, coaching, reporting, and collaboration into a single tenant-aware
facade. It automatically registers the critical sheets with `DatabaseManager`,
hydrates dashboard-ready aggregates, and exposes the most common CRUD flows used by
managers and agents.

```javascript
// Initialize during startup (handled automatically inside initializeSystem)
CallCenterWorkflowService.initialize();

// Hydrate the authenticated user's workspace
const workspace = CallCenterWorkflowService.getWorkspace(currentUserId, {
  campaignId: activeCampaignId,
  activeOnly: true,
});

// Schedule a shift for an agent (campaign enforcement handled internally)
CallCenterWorkflowService.scheduleAgentShift(managerId, {
  UserID: agentId,
  Date: '2024-04-01',
  StartTime: '09:00',
  EndTime: '17:00',
  CampaignID: activeCampaignId,
});

// Log QA/performance results
CallCenterWorkflowService.logPerformanceReview(qaLeadId, {
  UserID: agentId,
  CampaignID: activeCampaignId,
  EvaluationDate: new Date(),
  FinalScore: 94.2,
  Notes: 'Excellent rapport building',
});

// Create and acknowledge coaching engagements
const coaching = CallCenterWorkflowService.createCoachingSession(managerId, {
  UserID: agentId,
  CampaignID: activeCampaignId,
  FocusArea: 'Call control',
  DueDate: '2024-04-05',
});
CallCenterWorkflowService.acknowledgeCoaching(agentId, coaching.ID, {
  Notes: 'Reviewed and will implement feedback',
});

// Campaign-specific manager dashboard and communications
const managerDashboard = CallCenterWorkflowService.getManagerCampaignDashboard(managerId, activeCampaignId, {
  includeRoster: true,
  maxMessages: 10,
});

CallCenterWorkflowService.sendCampaignCommunication(managerId, activeCampaignId, {
  title: 'Script refresh',
  message: 'Please review the updated talk track before tomorrow\'s shift.',
  userIds: [agentId],
});

// Executive view across campaigns
const executiveOverview = CallCenterWorkflowService.getExecutiveAnalytics(executiveId);

// Post tenant-scoped collaboration updates
CallCenterWorkflowService.postCollaborationMessage(agentId, channelId, 'QA review completed', {
  campaignId: activeCampaignId,
});
```

The workspace payload returned by `getWorkspace` bundles:

* **Authentication context** – the signed-in user, their roles, claims, and active
  sessions.
* **Scheduling + attendance** – upcoming shifts, today's coverage, and shift counts.
* **Performance** – QA averages, attendance distribution, and adherence trends.
* **Coaching** – pending acknowledgements, overdue sessions, and coaching summaries.
* **Collaboration** – recent chat activity scoped to the user's accessible campaigns.
* **Reporting** – aggregated metrics (agents, shifts, QA scores, attendance rates,
  coaching backlog) ready for dashboards.

Each write helper (`scheduleAgentShift`, `recordAttendanceEvent`,
`logPerformanceReview`, `createCoachingSession`, `acknowledgeCoaching`,
`postCollaborationMessage`) automatically asserts campaign permissions through
`TenantSecurityService`, assigns the correct tenant column, and preserves the sheet
schemas registered with `DatabaseManager`.

## Extensibility considerations

- **Modular HTML views.** Dashboard panels such as Top Performers, QA rollups, and
  Coaching lists should be implemented as self-contained widgets that communicate
  through shared events and data loaders instead of directly depending on each
  other. This keeps the UI flexible so new KPIs or workflows can be introduced
  without breaking existing modules.
- **Service-level CRUD contracts.** Google Apps Script services (for example
  `UserService.gs`, `ScheduleService.gs`, `QAService.gs`, and `CoachingService.gs`)
  should expose consistent create/read/update/delete helpers for their
  corresponding sheets. Keeping these interfaces explicit enables automation,
  scheduled jobs, and future integrations to reuse the same endpoints instead of
  duplicating business logic.

## Deployment & security checklist

- **Manifest alignment.** The repository now ships the canonical
  [`appsscript.json`](./appsscript.json) manifest so deployments inherit the
  correct V8 runtime, advanced service enablement, and OAuth scopes. Ensure this
  file is committed with any future scope or dependency changes so the Apps
  Script project stays in sync across environments.
- **Web app execution context.** Web deployments execute as the signed-in user
  (`executeAs: USER_ACCESSING`) and require authentication (`access: ANYONE`).
  If perimeter authentication is handled outside the script, ensure the Apps
  Script deployment remains restricted to the intended audience because the
  in-app authentication module has been removed.
- **Operational documentation.** See [`docs/implementation-map.md`](./docs/implementation-map.md)
  for a cross-reference between the requirements, HTML front-end modules, and
  Apps Script services that fulfill them. Update this document whenever new
  modules are introduced so auditors can validate coverage quickly.

### Enterprise security controls

- **EnterpriseSecurityService.** Sensitive data written through `DatabaseManager`
  is now protected by `EnterpriseSecurityService.js`. The module derives a
  tenant-scoped encryption key from a master secret stored in Apps Script
  properties, encrypts flagged columns (such as session tokens), and attaches
  tamper-evident signatures to each record.
- **Automated audit trail.** Every insert, update, and delete routed through
  `DatabaseManager` emits a redacted audit event to the `SecurityAuditTrail`
  sheet so investigators can trace the actor, tenant, and change history
  without exposing the underlying secrets.
- **Schema-bound protections.** `DatabaseBindings.js` registers security
  metadata for the `Users`, `Sessions`, and `UserClaims` tables, ensuring that
  the new encryption and audit guarantees are enforced automatically whenever
  those sheets are accessed.

