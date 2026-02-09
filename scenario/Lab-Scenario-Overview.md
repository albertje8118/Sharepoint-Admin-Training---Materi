# Lab Scenario Overview — Shared Tenant (Trainer + 10 Admin Participants)

## Scenario title
**Project Northwind Intranet Modernization**

## Scenario summary
You are the Microsoft 365/SharePoint administration team for a fictional organization, **Northwind Traders**, which is modernizing its intranet and collaboration platform in SharePoint Online.

Across the 3 days, you will design and implement:
- A modern site and permission model
- Secure collaboration (including external collaboration controls)
- Information architecture (metadata/term store)
- Search experience configuration
- App governance basics
- Compliance controls (Purview)
- Operational monitoring and repeatable admin tasks (PowerShell)

This scenario is designed for a **single shared Microsoft 365 tenant** used by **10 participants**, all with administrator roles, plus a trainer.

---

## Shared-tenant operating rules (critical)
Because everyone works in the same tenant at the same time:
- Prefer **read-only verification** tasks when the learning goal is understanding an admin surface.
- When configuration changes are required:
  - The **trainer performs tenant-wide changes** once for the class, OR
  - Each participant performs changes only inside **their own isolated practice site**, using unique naming.
- Never ask participants to change settings that can disrupt the whole class (e.g., tenant-wide sharing lockdown) unless explicitly marked as **Trainer-only**.

---

## Participant identity and naming convention
Each participant gets a unique identifier:
- **Participant ID:** P01–P10 (assigned by trainer)

Use this ID to prevent naming collisions.

### Standard naming pattern
Use this format for anything you create:
- Sites: `NW-Pxx-<Purpose>`
- Microsoft 365 Groups (if used later): `NW-Pxx-<Purpose>`
- SharePoint lists/libraries: `NW-Pxx-<Name>`
- Term sets (if created by participants later): `NW-Pxx-<Name>`

Examples:
- `NW-P03-ProjectSite`
- `NW-P07-Policies`

---

## Environment baseline (trainer prepares once)
Trainer should prepare the tenant so labs are predictable:
- Confirm all participants can sign in and have the intended admin roles
- Create (or designate) a **class hub** location for shared reference materials

> Note: In this course delivery, **participants will create their own practice sites** using the naming convention below.

---

## Continuity artifacts (used across modules)
Participants should maintain a small record of what they did. For this course, keep a simple table (in your notes) with:
- Participant ID (Pxx)
- Practice site URL(s)
- Key settings observed/changed (module-by-module)
- Any issues encountered and how they were resolved

---

## What starts in Module 1
Module 1 focuses on **orientation and verification**:
- Confirm access to Microsoft 365 admin center and SharePoint admin center
- Review service health and message center
- Identify where tenant-level SharePoint settings live

No disruptive tenant-wide changes are made in Module 1.

---

## What changes in Module 2 (Identity, Access, External Sharing)
Module 2 introduces secure collaboration concepts and the difference between:
- Tenant-level policy (applies to everyone)
- Site-level configuration (can be isolated per participant practice site)

### External partner for the story
Northwind will collaborate with an external partner:
- Partner name: **Fabrikam (external)**

### Guest account approach (shared tenant)
To avoid sending 10x invitations and creating excess guest objects:
- **Trainer-only:** invites (or prepares) one or a small number of guest accounts to use for demonstrations.
- **Participants:** do not invite new guests unless explicitly instructed by the trainer.

### Practice site expectation
Starting Module 2 labs, each participant should have an isolated practice site:
- Site name pattern: `NW-Pxx-ProjectSite`
- Created by the participant as part of the course flow (when instructed).

---

## What starts in Module 3 (Working with Site Collections)
Module 3 is where each participant creates and manages their own modern site collection(s) as part of the Northwind story.

### Primary practice site (persist across later modules)
Each participant creates (or confirms) a primary practice site:
- Site name pattern: `NW-Pxx-ProjectSite`

This site is used as the “safe sandbox” for later labs (permissions, information architecture, search, governance, and compliance topics).

### Site lifecycle skills introduced
In Module 3, learners practice safe, admin-focused site lifecycle tasks:
- Create a modern site from the SharePoint admin center
- Review and manage site ownership/admins and membership surfaces
- Observe how site storage limits work (and whether the tenant is using automatic vs manual storage management)
- Practice delete/restore using a **temporary test site** (see below)
- Introduce SharePoint Online PowerShell cmdlets in a safe way (targeting only participant-owned sites)

### Temporary test site (delete/restore drill)
To learn deletion and recovery without risking the primary practice site:
- Test site name pattern: `NW-Pxx-RestoreTest`
- This site is created, deleted, and restored during Module 3.

Shared-tenant safety rule:
- Participants only delete/restore sites that match their own `NW-Pxx-...` naming.
- Trainer may optionally handle any permanent deletion/cleanup after the course.

---

## What starts in Module 4 (Permissions and Collaboration Model)
Module 4 introduces a practical permissions design exercise using a Northwind contracts workflow.

### Module 4 continuity artifacts (inside `NW-Pxx-ProjectSite`)
Each participant creates and maintains the following objects inside their own practice site:
- Document library: `NW-Pxx-Contracts`
- Folders: `01-Drafts`, `02-InReview`, `03-Final`
- SharePoint groups:
  - `NW-Pxx-Contracts Owners`
  - `NW-Pxx-Contracts Editors`
  - `NW-Pxx-Contracts Readers`

### Shared-tenant safety notes (permissions work)
- Permissions changes are limited to participant-owned objects inside `NW-Pxx-ProjectSite`.
- The lab demonstrates breaking inheritance at the **library** and at **one folder only** to keep unique permission scopes low.
- Sharing link exercises should be internal-only unless the trainer explicitly provides test accounts and approves external sharing tests.

---

## What starts in Module 5 (Managing Metadata and the Term Store)
Module 5 introduces information architecture (IA) concepts and managed metadata for consistent classification.

### Module 5 continuity artifacts
Participants extend their contracts library with managed metadata.

Tenant-wide (Term store) artifacts (only create items with your own prefix):
- Term group: `NW-Pxx-TermGroup`
- Term set: `NW-Pxx-ContractType`
- Terms (minimum): NDA, MSA, SOW, Renewal

Site/library artifacts (inside `NW-Pxx-ProjectSite`):
- Library column (managed metadata): `NW-Pxx-ContractType` added to `NW-Pxx-Contracts`

### Shared-tenant safety notes (term store)
- The Term store is tenant-wide; do not edit or delete other participants’ term groups/sets.
- If a participant cannot create a term group (insufficient term store role), use a **local term set** created from the site/library column experience.

---

## What starts in Module 6 (Search in SharePoint Online and Microsoft Search)
Module 6 focuses on information discovery, search validation, and safe, training-scoped search “answers”.

### Module 6 continuity artifacts
Participants create “seed” documents and validate indexing:
- Seed documents uploaded to: `NW-Pxx-Contracts/02-InReview`
  - `NW-Pxx-Search-Seed-A.docx`
  - `NW-Pxx-Search-Seed-B.docx`

Optional (only if the participant has Search admin/editor permissions):
- Microsoft Search Bookmark (training-scoped)
  - Title pattern: `NW-Pxx Project Site`
  - Keywords: must include `Pxx` and must be unique/training-only
- Microsoft Search Acronym (training-scoped)
  - Acronym pattern: `NWPxx` (no spaces)
  - Meaning: “Northwind Participant Pxx”

### Shared-tenant safety notes (search)
- Microsoft Search answers (Bookmarks/Acronyms) are organization-level curated content.
- Avoid creating keywords that overlap across participants; include `Pxx` in keywords.
- Do not create “reserved keywords” for common org terms (for example, “helpdesk”, “benefits”).
- Reindexing can create load; only request reindex after a search visibility or schema change.

### Note on Q&A answers
Some Microsoft Search answer types (including Q&As) have changed availability over time. Treat Q&As as a concept; design hands-on labs primarily around Bookmarks and Acronyms.

---

## What starts in Module 7 (Apps and Customization in SharePoint Online)
Module 7 introduces safe customization patterns and the governance surfaces for modern SharePoint solutions.

### Module 7 continuity artifacts
Participants create a small “app request intake” list and apply formatting (no code, no app deployment required):
- List: `NW-Pxx-AppRequests` (inside `NW-Pxx-ProjectSite`)
- Customization:
  - Column formatting applied to `Status`
  - View formatting applied to improve scanability

Optional (trainer-led / read-only tours):
- SharePoint admin center → **More features** → **Apps** (manage solutions, including SPFx `.sppkg`)
- SharePoint admin center → **API access** (review pending/approved API permission requests)

### Shared-tenant safety notes (apps/customization)
- Treat tenant-wide app deployment as **Trainer-only** unless explicitly assigned.
- Do not approve API access requests in a shared training tenant.
- Prefer low-risk customization first (lists, JSON formatting) and keep governance discussions grounded in least-privilege.

### 2026 alignment note (add-ins)
Microsoft documentation states SharePoint add-ins are being retired for SharePoint in Microsoft 365. In this course delivery, treat add-ins as legacy and focus on SPFx + declarative customization + governance.

---

## What starts in Module 8 (Content Governance and Compliance with Microsoft Purview)
Module 8 introduces compliance controls in Microsoft Purview with a shared-tenant-safe split:
- **Trainer-led:** create/publish labels, DLP policies, and any tenant-wide compliance configuration
- **Participant hands-on:** apply labels to content inside their own `NW-Pxx-ProjectSite`

### Module 8 continuity artifacts (inside `NW-Pxx-ProjectSite`)
Participants create and maintain:
- Folder: `NW-Pxx-Contracts/04-Compliance`
- Document: `NW-Pxx-Confidential-Statement.docx` (FAKE content, training only)
- Applied label outcomes recorded in notes:
  - Sensitivity label name (if published)
  - Retention label name (if published)

Optional (only if trainer assigns permissions and approves scope):
- eDiscovery case name: `NW-Pxx-eDiscovery-Case`
- Search keyword: `CONFIDENTIAL`

### Shared-tenant safety notes (Purview)
- Labels and policies can be tenant-wide in effect; do not create/publish policies as a participant unless explicitly assigned.
- eDiscovery holds can preserve content and may affect deletion workflows; treat holds as trainer-led by default.
- Never use real personal data for DLP/label testing.

---

## What starts in Module 9 (OneDrive for Business Administration)
Module 9 focuses on OneDrive administration surfaces and operational behaviors, with an emphasis on avoiding tenant-wide disruption in a shared training environment.

### Module 9 continuity artifacts
Participants create and maintain a small OneDrive test file (internal-only sharing) to observe policy behavior:
- OneDrive file: `NW-Pxx-OneDrive-Sharing-Test.txt` (or Word equivalent)

### Shared-tenant safety notes (OneDrive)
- Most OneDrive settings are configured in the **SharePoint admin center** and are **tenant-wide**.
- **Trainer-only:** any change to tenant-wide OneDrive/SharePoint settings (Sharing, Sync, Storage limit, Retention, Access control).
- **Participants:** do read-only verification in admin center and perform safe internal sharing tests inside their own OneDrive.
- Unmanaged device access controls can take time to apply and may not affect users already signed in; plan demonstrations accordingly.
- Deleted-user OneDrive retention and restore are administrative lifecycle concepts; treat restore drills as trainer-led unless explicitly assigned.

---

## What starts in Module 10 (Administration and Automation with PowerShell)
Module 10 introduces repeatable administration and reporting using PowerShell, with safe scoping practices for a shared training tenant.

### Module 10 continuity artifacts
Participants generate and keep local outputs from the lab (for review and comparison):
- CSV report: `NW-Pxx-Sites.csv` (site inventory scoped to `NW-Pxx`)
- CSV report: `NW-Pxx-Users.csv` (site user membership for `NW-Pxx-ProjectSite`)
- CSV report: `Targets-Users.csv` (CSV-driven loop output)

Optional (trainer-approved):
- Graph connection confirmation notes (scopes used, any consent prompts)

### Shared-tenant safety notes (PowerShell/automation)
- PowerShell can change tenant-wide settings quickly; treat tenant-wide changes as **Trainer-only**.
- Participants should run **read-only reporting** commands, and scope any queries/filters to their own `NW-Pxx` artifacts.
- Prefer an export-first workflow: capture baseline (CSV), then make changes only if explicitly assigned.

---

## What starts in Module 11 (Operations at Scale: Monitoring, Lifecycle, External Sharing Governance)
Module 11 consolidates operational admin behaviors:
- monitoring and incident awareness
- troubleshooting common SharePoint/OneDrive issues
- site lifecycle guardrails (delete/restore)
- external sharing governance and default link strategy

### Module 11 continuity artifacts
Participants record operational notes (read-only verification) and keep local report outputs:
- Incident triage worksheet notes (Service health / Message center)
- External sharing governance worksheet notes (org-level + site-level constraints)
- Optional PowerShell exports (scoped to `NW-Pxx`):
  - `NW-Pxx-Sites.csv`
  - `NW-Pxx-Users.csv`

### Shared-tenant safety notes (ops/governance)
- Do not change tenant-wide policies as a participant.
- Treat governance as: **observe → document → request change**.
- Any demonstrations that can affect the whole tenant (sharing policies, lock state, deletions/restores) are **Trainer-only** unless explicitly assigned.

---

## What starts in Module 12 (Optional) — Workflow Automation with Power Automate and Power Apps
Module 12 turns the “app request intake” list into a basic workflow and form experience:
- Power Automate approval flow (SharePoint list item trigger → approval → update item)
- Power Apps custom form for a SharePoint list
- Governance discussion: environments + DLP (Data Loss Prevention) policies

### Module 12 continuity artifacts (inside `NW-Pxx-ProjectSite`)
Participants create and maintain:
- Flow (name pattern): `NW-Pxx-AppRequests-Approval`
- Power Apps custom form published to list: `NW-Pxx-AppRequests`

Recommended list behavior for training:
- Use a trigger pattern that avoids loops (prefer **When an item is created** for workflows that update the same item).

### Shared-tenant safety notes (Power Platform)
- Build flows/apps only for your own `NW-Pxx-...` list/site.
- Keep the approver set to **yourself** (or a trainer-designated account) to avoid spamming other participants.
- Treat environments and DLP policy changes as **Trainer-only** unless explicitly assigned.
