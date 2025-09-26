# Lumina-Sheets

Call center management system built on Google Apps Script + Google Sheets.

## Google Sheets database manager

The `DatabaseManager.gs` module turns any worksheet into a CRUD-ready table. Define
schema defaults once and re-use the same interface across every client campaign.

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
