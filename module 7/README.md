# Module 7 — Apps and Customization in SharePoint Online

**Course:** Modern SharePoint Online for Administrators (3-Day, 2026 aligned)  \
**Module duration (suggested):** 1.5–2.0 hours instruction + 60–90 min lab  \
**Level:** Intermediate

## Scenario continuity (shared tenant)
This module continues the course scenario:
- Scenario overview: ../scenario/Lab-Scenario-Overview.md

Shared-tenant note (trainer + 10 admin participants):
- App deployment and API permission approvals can be **tenant-impacting**. Tenant-wide actions are **Trainer-only** unless the trainer explicitly assigns a safe, isolated exercise.
- Participants should only create/modify objects inside their own practice site using strict naming:
  - Site: `NW-Pxx-ProjectSite`
  - Prefix: `NW-Pxx-...`

## Learning objectives
By the end of this module, learners will be able to:
1. Describe modern SharePoint customization options: out-of-box configuration, declarative JSON formatting, and SharePoint Framework (SPFx).
2. Use **column formatting** and **view formatting** to customize a list without custom code.
3. Explain what the SharePoint **Apps** management surface does and what “enabling” an app means.
4. Explain what the **API access** page is used for in SharePoint Online and why approvals are governance-sensitive.
5. Identify key governance decisions for app deployment (scope, permissions, and rollout).

## Deliverables in this folder
- Student guide chapter: Module-07-Student-Guide.md
- Lab guide: Lab-07-Managing-Apps-and-Customization.md
- Slides outline: Module-07-Slides.md

## Validation checkpoint (end of module)
Learners can:
- Create a list `NW-Pxx-AppRequests` and apply column + view formatting using JSON.
- Explain (at a high level) how SPFx solutions are deployed (Apps / App Catalog) and why deployment scope matters.
- Explain what API access approvals do, and why they require the right admin role and a review process.
