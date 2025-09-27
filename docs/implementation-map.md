# Implementation Coverage Map

This document cross-references the core requirements with the Apps Script
services and HTML front-end modules found in this repository. Use it alongside
[`requirements.md`](./requirements.md) and [`next-steps.md`](./next-steps.md) to
validate that each workflow remains up to date as additional campaigns and
features are introduced.

## 1. Authentication & Authorization

| Requirement focus | Primary services | Supporting assets |
| --- | --- | --- |
| Secure login, password resets, verification | `AuthenticationService.js`, `EmailService.js`, `ForgotPassword.html`, `EmailConfirmed.html`, `ResendVerification.html`, `ChangePassword.html` |
| Role & claim management | `RolesService.js`, `RoleManagement.html`, `Users.html` |
| Session persistence & renewal | `AuthenticationService.js`, `MainUtilities.js`, `ScheduleUtilities.js` |
| Tenant-aware access enforcement | `TenantSecurityService.js`, `CallCenterWorkflowService.js`, `TenantSecurityService` helpers consumed by `CallCenterWorkflowService.js` and individual feature services |

## 2. Multi-Campaign Operations

| Requirement focus | Primary services | Supporting assets |
| --- | --- | --- |
| Campaign CRUD & permissions | `CallCenterWorkflowService.js`, `CampaignService.js`, `CampaignManagement.html`, `TenantSecurityService.js` |
| Executive views across campaigns | `ManagerExecutiveService.js`, `ManagerExecutiveExperience.html`, `Dashboard.html`, `UnifiedQADashboard.html` |
| Campaign navigation for agents | `AgentExperienceService.js`, `AgentExperience.html`, `ScheduleManagement.html` |

## 3. Scheduling & Attendance

| Requirement focus | Primary services | Supporting assets |
| --- | --- | --- |
| Shift generation & management | `ScheduleService.js`, `ScheduleManagement.html`, `ScheduleUtilities.js`, `Schedule.html` |
| Attendance logging & reporting | `AttendanceService.js`, `AttendanceReports.html`, `ImportAttendance.html` |
| Slot management & workforce planning | `SlotManagementInterface.html`, `ScheduleUtilities.js`, `TaskBoard.html` |

## 4. Quality Assurance & Coaching

| Requirement focus | Primary services | Supporting assets |
| --- | --- | --- |
| QA forms, dashboards, and collaboration | `QAService.js`, `QAPdfService.js`, `QACollabService.js`, `QACollabList.html`, `QADashboard.html`, `QualityCollabView.html` |
| Campaign-specific QA variants | `CreditSuiteQAServices.js`, `CreditSuiteQAForm.html`, `IndependenceQAForm.html`, `IBTRUtilities.js`, `IndependenceQAUtilities.js` |
| Coaching workflows & acknowledgements | `CoachingService.js`, `CoachingDashboard.html`, `CoachingAckForm.html`, `CoachingList.html`, `CoachingView.html` |

## 5. Collaboration & Messaging

| Requirement focus | Primary services | Supporting assets |
| --- | --- | --- |
| In-app chat threads | `ChatService.js`, `Chat.html`, `ChatBubble.html`, `chatHeader.html` |
| Notifications & broadcasts | `CallCenterWorkflowService.js`, `Notifications.html`, `BookmarkService.js`, `BookmarkManager.html` |
| Searchable knowledge base | `SearchService.js`, `SearchDeploymentService.js`, `SearchSecurityService.js`, `Search.html`, `SearchSecurityDashboard.html` |

## 6. Shared Utilities & Database Access

- **DatabaseManager.js** and **DatabaseBindings.js** register schema metadata,
  provide CRUD helpers, and attach timestamp/tenant hooks used across all
  services.
- **MainUtilities.js**, **ScheduleUtilities.js**, **IBTRUtilities.js**, and
  **TGCUtilities.js** encapsulate cross-cutting helpers for formatting,
  calculations, and service orchestration.
- **SeedData.js** bootstraps sheet headers and sample records for new
  deployments.

## 7. HTML Module Inventory

The following HTML bundles correspond to the major personas described in the
requirements:

- **Agents:** `AgentExperience.html`, `AgentSchedule.html`, `Chat.html`,
  `CoachingAckForm.html`, `QualityView.html`.
- **Managers:** `ManagerExecutiveExperience.html`, `ScheduleManagement.html`,
  `CoachingDashboard.html`, `AttendanceReports.html`, `CampaignManagement.html`,
  `CallReports.html`.
- **Executives:** `Dashboard.html`, `UnifiedQADashboard.html`,
  `IndependenceCoachingDashboard.html`, `CollaborationReporting.html`.
- **QA & Compliance:** `QualityCollabForm.html`, `QualityCollabView.html`,
  `QADashboard.html`, `GroundingQAForm.html`, `ComplianceReportingService.js`.
- **Administration:** `Users.html`, `RoleManagement.html`, `TasksConfig.js`,
  `TasksService.js`, `UserService.js`.

Keep these mappings updated so auditors can quickly validate coverage against
[`requirements.md`](./requirements.md).

## 8. Audit & Next Steps Alignment

The `docs/next-steps.md` plan focuses on security reviews and UI audits. The
following artifacts support that effort today:

- **Service catalogue:** The tables above summarize CRUD-oriented services and
  their respective modules, addressing the "Catalogue all services" task.
- **Authorization checkpoints:** `TenantSecurityService.js`,
  `AuthenticationService.js`, and `CallCenterWorkflowService.js` centralize
  campaign and role validation as required by the audit plan.
- **UI inventory:** Section 7 aligns HTML modules to the personas defined in the
  requirements, providing the baseline requested in the next-steps document.

Update this implementation map whenever new campaigns, modules, or services are
added so the Apps Script project stays compliant with the documented
requirements.
