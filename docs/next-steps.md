# Security and UI Audit Next Steps

This plan elaborates on the immediate follow-up actions required to secure and
validate the current Lumina Sheets implementation before adding new
functionality.

## 1. Audit Apps Script Services

1. **Catalogue all services**
   - List every `.gs` file that exposes CRUD operations.
   - Document the campaigns, roles, and data entities each service touches.
2. **Evaluate authorization gates**
   - Confirm `Session.getActiveUser()` or equivalent identity checks prevent
     anonymous access.
   - Verify each entry point enforces campaign scoping, rejecting requests that
     lack a campaign identifier or reference an unauthorized campaign.
   - Ensure role validation occurs prior to mutating operations (create, update,
     delete) and that read access is restricted accordingly.
3. **Identify remediation items**
   - Note missing checks, ambiguous role definitions, or shared state across
     campaigns.
   - File remediation tasks that include the affected service, risk level, and
     proposed fix.

## 2. Inventory HTML/JavaScript Modules

1. **Map UI modules to requirements**
   - Associate each `.html` asset with its functional requirement(s) from
     `requirements.md`.
   - Highlight requirements without a current UI implementation.
2. **Evaluate data exposure**
   - Confirm client-side scripts only request campaign-scoped data.
   - Review embedded scripts for assumptions about user roles or hard-coded
     identifiers.
3. **Document gaps**
   - Capture missing modules, broken interactions, or modules that bypass role
     checks via direct service calls.

## 3. Prioritize Security Enhancements

1. **Assess perimeter security**
   - Authentication has been removed from the Apps Script layer. Ensure the
     deployment URL is protected by your hosting environment or network
     controls.
   - Document who should have access to the open dashboards and implement
     protections outside of the script when needed.
2. **Harden campaign scoping**
   - Centralize campaign resolution (e.g., a shared utility) to remove duplicate
     logic across services.
   - Implement guard clauses that terminate requests when campaign mismatches
     are detected.
3. **Refine role management**
   - Normalize role definitions and inheritance in `RolesService.gs` or related
     utilities.
   - Provide administrative tooling to adjust roles within the constraints of
     campaign assignments.
4. **Produce a remediation roadmap**
   - Sequence the identified fixes by risk and dependency.
   - Reserve capacity for regression testing before resuming feature work.

## Deliverables

- A consolidated audit report covering both service-side and UI findings.
- A prioritized backlog with estimated effort, owners, and target release
  windows for the remediation items.
- Updated documentation reflecting the enforced security model and any new
  utilities introduced during hardening.
