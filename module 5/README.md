# Module 5 — Managing Metadata and the Term Store

**Course:** Modern SharePoint Online for Administrators (3-Day, 2026 aligned)  
**Module duration (suggested):** 2.0–2.5 hours instruction + 75–105 min lab  
**Level:** Intermediate

## Scenario continuity (shared tenant)
This module continues the course scenario:
- Scenario overview: ../scenario/Lab-Scenario-Overview.md

Shared-tenant note (trainer + 10 admin participants):
- The **Term store is tenant-wide**, so participants must only create items with their own `NW-Pxx-...` names.
- Do not edit or delete other participants’ term groups/sets.
- If you do not have permission to create a term group, use the lab’s **Local term set fallback** (site-scoped) and proceed.

In this module, participants work inside their existing practice site and library:
- Site: `NW-Pxx-ProjectSite`
- Library (from Module 4): `NW-Pxx-Contracts`

## Learning objectives
By the end of this module, learners will be able to:
1. Explain information architecture basics and why metadata matters for governance and search.
2. Distinguish **managed metadata** (term store) from **site columns** and when to use each.
3. Navigate the Term store in the SharePoint admin center and explain the hierarchy: term groups → term sets → terms.
4. Explain delegated term management roles (term store admin, group manager, contributor).
5. Create a term group and term set safely using participant-specific naming.
6. Add a managed metadata column to a library and map it to a term set.
7. Validate metadata entry behavior and troubleshoot common taxonomy issues.

## Deliverables in this folder
- Student guide chapter: Module-05-Student-Guide.md
- Lab guide: Lab-05-Creating-and-Managing-Metadata.md
- Slides outline: Module-05-Slides.md

## Validation checkpoint (end of module)
Learners can:
- Point to their term set in the Term store (or a local term set if using fallback).
- Add a managed metadata column to `NW-Pxx-Contracts`.
- Tag at least 3 documents with consistent terms.
- Explain how term store permissions enable delegated management.
