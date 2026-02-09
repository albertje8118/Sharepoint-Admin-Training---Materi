# Module 4 — Permissions and Collaboration Model

**Course:** Modern SharePoint Online for Administrators (3-Day, 2026 aligned)  
**Module duration (suggested):** 2.0–2.5 hours instruction + 75–105 min lab  
**Level:** Intermediate

## Scenario continuity (shared tenant)
This module continues the course scenario:
- Scenario overview: ../scenario/Lab-Scenario-Overview.md

Shared-tenant note (trainer + 10 admin participants):
- Treat tenant-wide changes (sharing settings, guest invites, org policies) as **Trainer-only** unless explicitly isolated.
- Participant hands-on changes must be limited to participant-owned sites using `NW-Pxx-...` naming.

In this module, participants work inside their existing practice site:
- `NW-Pxx-ProjectSite` (created earlier; persists across later modules)

## Learning objectives
By the end of this module, learners will be able to:
1. Explain permission inheritance and why unique permissions (permission scopes) must be used sparingly.
2. Distinguish site-level permissions from library/folder/item permissions and know when to break inheritance.
3. Explain how permissions work differently on group-connected team sites vs non-group-connected sites.
4. Design a practical permission model for a real collaboration scenario (Northwind contracts workflow).
5. Implement a permission model using SharePoint groups and standard permission levels (Full Control, Edit, Read).
6. Validate and troubleshoot access using “Check Permissions” and modern “Manage access” link views.
7. Describe how sharing links work (organization links vs specific people) and how to remove/revoke access.

## Deliverables in this folder
- Student guide chapter: Module-04-Student-Guide.md
- Lab guide: Lab-04-Designing-a-Permission-Model.md
- Slides outline: Module-04-Slides.md

## Validation checkpoint (end of module)
Learners can:
- Describe (and diagram) inheritance from site → library → folder → item.
- Implement unique permissions at a library and folder with minimal scope count.
- Demonstrate how to verify access using “Check Permissions”.
- Demonstrate how to review/remove sharing links (“Manage access”).
