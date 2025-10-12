# Lumina Admin System Owner

The Lumina Admin module provisions a system-owned superuser and a fully automated governance framework for Lumina Sheets deployments. This guide explains the core concepts, how to enable or disable autonomous mode, and how to extend categories, jobs, and policies.

## System Owner Identity

* **User ID:** `LUMINA-SYSTEM-OWNER`
* **Username:** `lumina-admin`
* **Email:** `lumina-admin@system`
* **Role:** `SystemOwner`

The account is immutable and bypasses tenant scoping. Seeding occurs automatically through `LuminaAdmin.ensureSeeded()` which is invoked during `seedDefaultData()` and exposed as `seedLuminaSystemOwner()`.

## Access Control Model

The SystemOwner role receives the wildcard capability (`*`) which grants read/write to every campaign. Other roles include `CampaignManager`, `Agent`, and `GuestClient`. Attribute-based checks rely on the `CampaignId` field: non-SystemOwner users must have the campaign in their assignment list.

## Sidebar Registry

`LuminaAdmin.buildSystemOwnerNavigation()` produces a global sidebar grouped by category. Update the registry in `LuminaAdminService.gs` (the `CATEGORY_REGISTRY` constant) to add new categories or pages.

## Background Jobs

Autonomous operations rely on the Jobs sheet. Default jobs include integrity scans, policy enforcement, QA monitoring, benefits eligibility checks, integration health, and feature flag audits. Each job is idempotent by persisting a `RunHash` of the last results.

* Call `LuminaAdmin.ensureTriggers()` after deployment to register Apps Script time-based triggers.
* Manually run jobs with `LuminaAdmin.runJob('integrity_scan')` or via the HTML client helper `LuminaAdminClient_runJob`.

## Notifications and Messages

Alerts persist to the `SystemMessages` sheet. Severities include `INFO`, `NOTICE`, `WARNING`, and `CRITICAL`. The Messages Center UI (`LuminaAdminMessages.html`) allows acknowledge/resolve actions which append to the audit log.

Email delivery honors the `notify_email` feature flag. In-app banners should read from the same sheet for realtime dashboards.

## Lifecycle Actions

Use `LuminaAdmin.lifecycleAction(action, payload, actor)` for hires, transfers, promotions, and terminations. Each action logs to `AuditLog` with before/after snapshots.

## TOTP and Security

`LuminaAdmin.ensureTotpSecret(userId)` provisions a TOTP shared secret and enables multifactor authentication for the specified user. Use this when onboarding privileged administrators.

## Extensibility

* **Add categories/pages:** edit `CATEGORY_REGISTRY`.
* **Add jobs:** push new definitions into `DEFAULT_JOBS` and add handlers in `JOB_HANDLERS`.
* **Add policies:** extend job handlers to evaluate additional constraints and publish messages.

Disable autonomous mode by setting the `autonomous_mode` feature flag to `false`. This skips scheduled jobs while still allowing interactive runs.

## .NET Adapter Hooks

If the optional ASP.NET Core services are deployed, expose API endpoints that forward to Apps Script webhooks:

* `POST /api/admin/jobs/run` → `LuminaAdmin.runJob`
* `GET /api/system/messages` → `LuminaAdmin.getMessages`
* `GET /api/audit` → `AuditService.list`

Use service accounts with read-only scopes for telemetry scans and limit write access to auditable flows.
