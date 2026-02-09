# Module 10 — Administration and Automation with PowerShell

**Course:** Modern SharePoint Online for Administrators (3-Day, 2026 aligned)  \
**Module duration (suggested):** 1.5–2.0 hours instruction + 75–105 min lab  \
**Level:** Intermediate

## Scenario continuity (shared tenant)
This module continues the course scenario:
- Scenario overview: ../scenario/Lab-Scenario-Overview.md

Shared-tenant note (trainer + 10 admin participants):
- PowerShell can change tenant-wide settings quickly. In this course delivery:
  - Participants run **read-only reporting** commands safely.
  - Any **bulk change** commands must be scoped only to `NW-Pxx-...` artifacts and explicitly approved.

## Learning objectives
By the end of this module, learners will be able to:
1. Connect to SharePoint Online using the SharePoint Online Management Shell (`Connect-SPOService`).
2. Produce basic tenant and site reports (sites list, storage info, user/group membership).
3. Use `Export-Csv`, `Out-File`, and CSV-driven loops (`Import-Csv` + `ForEach-Object`) for repeatable admin tasks.
4. Connect to Microsoft Graph using Microsoft Graph PowerShell (`Connect-MgGraph`) and understand scope-based authentication.
5. Apply safe automation practices in a shared tenant (idempotence, scoping, dry-run mindset).

## Deliverables in this folder
- Student guide chapter: Module-10-Student-Guide.md
- Lab guide: Lab-10-Automating-SharePoint-Administration.md
- Slides outline: Module-10-Slides.md

## Validation checkpoint (end of module)
Learners can:
- Connect to SharePoint Online PowerShell and list their `NW-Pxx` sites.
- Generate and export a permissions/user membership report for their own site.
- Explain when to use SharePoint Online PowerShell vs Microsoft Graph PowerShell.
