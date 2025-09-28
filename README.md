# Lumina-Sheets

Call center management system built on Google Apps Script + Google Sheets.

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
  defaults: { CanLogin: true },
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
users.update(user.ID, { CanLogin: false });

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
  where: { CampaignID: campaignId, CanLogin: true },
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

## Structured Google Sheets datastore (SheetsDB)

For APIs that require stronger guarantees than the legacy helpers provide, use the
`SheetsDatabase.js` datastore. It layers a typed schema, validation, optimistic
locking, audit logging, and REST-style access on top of standard worksheets.

### Key capabilities

- **Typed schema & validation** – declare column types (`string`, `number`,
  `timestamp`, `enum`, `json`), required fields, allowed values, numeric ranges,
  and regex-based constraints. Unknown fields are rejected automatically.
- **Primary keys & timestamps** – string IDs (e.g., `COACH_000123`) are generated on
  insert alongside `createdAt`, `updatedAt`, and `deletedAt` columns.
- **Soft delete & retention** – `deletedAt` powers soft deletes, while automatic
  archive helpers move aged rows into partitioned archive sheets.
- **Uniqueness & foreign keys** – mark columns as `unique` or attach
  `{ table, column }` references. Writes fail if the constraints are violated.
- **Indexes & derived sheets** – opt-in indexes (`__idx_*` helper sheets) keep
  frequent lookups fast. Audit logs and an outbox table support downstream
  integrations.
- **Idempotent writes** – pass an `idempotencyKey` to `create` requests to guard
  against duplicate inserts when clients retry.
- **Optimistic concurrency** – updates may include `expectedUpdatedAt` to ensure
  records have not changed between reads and writes.
- **Pagination & filtering** – cursor-based pagination (`updatedAt` + ID), offset
  pagination, and whitelisted operators (`=`, `contains`, `<`, `>`, `<=`, `>=`) are
  available through the API layer.
- **Backups & maintenance** – `SheetsDB.runMaintenance()` snapshots active tables
  and purges archived data according to retention policies.

### Bootstrapping tables

During `initializeSystem()` the project calls `initializeSheetsDatabase()` to define
sample tables (`CoachingSessions`, `QualityReviews`, and `WebhooksOutbox`). Use this as a template
for your own schemas:

```javascript
SheetsDB.defineTable({
  name: 'CoachingSessions',
  primaryKey: 'id',
  idPrefix: 'COACH_',
  columns: [
    { name: 'id', type: 'string', primaryKey: true },
    { name: 'tenantId', type: 'string', required: true },
    { name: 'agentId', type: 'string', required: true,
      references: { table: 'Users', column: 'ID', allowNull: false } },
    { name: 'coachId', type: 'string', required: true,
      references: { table: 'Users', column: 'ID', allowNull: false } },
    { name: 'sessionDate', type: 'timestamp', required: true },
    { name: 'status', type: 'enum', required: true,
      allowedValues: ['scheduled', 'completed', 'cancelled'], defaultValue: 'scheduled' }
  ],
  indexes: [{ name: 'CoachingSessions_agent', field: 'agentId' }],
  retentionDays: 365
});
```

### REST-style access

The web app exposes `/exec?api=db` for programmatic access:

```
GET  ?api=db&table=CoachingSessions&limit=50
GET  ?api=db&table=CoachingSessions&id=COACH_000123
POST ?api=db (body: { "action": "create", "table": "CoachingSessions", ... })
```

Protect the endpoint with script properties:

```javascript
PropertiesService.getScriptProperties().setProperty('SHEETS_DB_API_KEYS', JSON.stringify({
  'prod-reader-key': 'reader',
  'prod-writer-key': 'writer',
  'prod-admin-key': 'admin'
}));
```

Roles map to permissions (`reader`, `writer`, `admin`). Requests without a matching
API key are rejected with standardized JSON errors. Responses include `records`,
`total`, and optional `nextCursor` tokens for pagination.

### Authentication & campaign-scoped sessions

`AuthenticationService.gs` now persists sessions through `DatabaseManager`, upgrading
the `Sessions` sheet to track remember-me flags, user agents, IP addresses, and a
serialized tenant scope. Every login computes a tenant access profile via
`TenantSecurityService`, blocks accounts that are not assigned to at least one
campaign (unless they are global administrators), and returns the full list of
allowed/managed/admin campaigns to the client. Session renewals automatically refresh
the campaign scope so managers cannot switch to unauthorized tenants mid-session.


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
  This configuration enforces the tenant-aware permission checks implemented in
  `AuthenticationService.gs` and `TenantSecurityService.gs` while still allowing
  external client stakeholders to authenticate with Google accounts.
- **Operational documentation.** See [`docs/implementation-map.md`](./docs/implementation-map.md)
  for a cross-reference between the requirements, HTML front-end modules, and
  Apps Script services that fulfill them. Update this document whenever new
  modules are introduced so auditors can validate coverage quickly.

