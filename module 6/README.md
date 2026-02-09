# Module 6 — Search in SharePoint Online and Microsoft Search

**Course:** Modern SharePoint Online for Administrators (3-Day, 2026 aligned)  \
**Module duration (suggested):** 1.5–2.0 hours instruction + 60–90 min lab  \
**Level:** Intermediate

## Scenario continuity (shared tenant)
This module continues the course scenario:
- Scenario overview: ../scenario/Lab-Scenario-Overview.md

Shared-tenant note (trainer + 10 admin participants):
- Microsoft Search “answers” (Bookmarks/Acronyms) are **organization-level** content. Use **strict `NW-Pxx` naming** and keywords so you don’t impact other participants.
- Avoid “reserved keywords” for common terms (for example, “helpdesk”, “benefits”). Use only Northwind training keywords.
- Treat tenant-wide Search Schema changes as **Trainer-only** unless the trainer explicitly assigns a safe, isolated exercise.

In this module, participants work inside their existing practice site and library:
- Site: `NW-Pxx-ProjectSite`
- Library: `NW-Pxx-Contracts`

## Learning objectives
By the end of this module, learners will be able to:
1. Explain the difference between SharePoint search and Microsoft Search entry points, including security trimming.
2. Describe how Microsoft Search “answers” work (Bookmarks, Acronyms) and which admin roles manage them.
3. Create a **training-scoped** Bookmark and Acronym using `NW-Pxx` naming conventions.
4. Validate that SharePoint content is searchable and troubleshoot common indexing issues.
5. Explain (at a high level) crawled vs managed properties and why search schema changes often require reindexing.
6. Use basic query operators (phrases, Boolean operators, simple property restrictions) for faster troubleshooting.

## Deliverables in this folder
- Student guide chapter: Module-06-Student-Guide.md
- Lab guide: Lab-06-Configuring-Search-Experience.md
- Slides outline: Module-06-Slides.md

## Validation checkpoint (end of module)
Learners can:
- Upload two “seed” documents to `NW-Pxx-Contracts` and find them via search.
- Reindex a library (when appropriate) and explain the impact.
- Create a training-scoped Bookmark that is visible immediately after publishing.
- Create a training-scoped Acronym and explain why it can take time to appear.
