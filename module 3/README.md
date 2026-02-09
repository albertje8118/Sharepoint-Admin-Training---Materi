# Module 3 — Working with Site Collections (Modern Sites)

**Course:** Modern SharePoint Online for Administrators (3-Day, 2026 aligned)  
**Module duration (suggested):** 2.0–2.5 hours instruction + 90–120 min lab  
**Level:** Intermediate

## Scenario continuity (shared tenant)
This module continues the course scenario:
- Scenario overview: ../scenario/Lab-Scenario-Overview.md

Shared-tenant note (trainer + 10 admin participants):
- Treat tenant-wide changes as **Trainer-only** unless explicitly isolated.
- Participant hands-on changes must be limited to participant-owned sites using `NW-Pxx-...` naming.

In this module, participants create and manage:
- Primary practice site: `NW-Pxx-ProjectSite` (persist across later modules)
- Temporary test site: `NW-Pxx-RestoreTest` (used for delete/restore drill)

## Learning objectives
By the end of this module, learners will be able to:
1. Create modern SharePoint sites (site collections) from the SharePoint admin center.
2. Explain key differences between common modern site types at a practical admin level.
3. Locate and use the site details panel (membership/admins, settings) in the SharePoint admin center.
4. Perform common site admin operational tasks (membership, access requests, recycle bin recovery) within a site.
5. Describe how SharePoint storage is managed (automatic vs manual) and what can be set per site.
6. Safely delete and restore a SharePoint site using the SharePoint admin center.
7. Use SharePoint Online PowerShell at a basic level to connect and retrieve site properties.

## Deliverables in this folder
- Student guide chapter: Module-03-Student-Guide.md
- Lab guide: Lab-03-Working-with-Site-Collections.md
- Slides outline: Module-03-Slides.md

## Validation checkpoint (end of module)
Learners can:
- Create (or confirm) their `NW-Pxx-ProjectSite` and record the URL.
- Identify where to manage site admins/membership and sharing settings for a site.
- Demonstrate the delete/restore lifecycle using a test site without impacting others.
- Run a basic PowerShell connection and retrieve site properties for their site.
