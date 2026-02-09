# Lab 04 — Designing a Permission Model (Northwind Contracts)

**Module:** 4 — Permissions and Collaboration Model  
**Estimated time:** 75–105 minutes  
**Lab type:** Individual (shared-tenant safe)

## Lab goal
Design and implement a permission model that:
- keeps most collaboration simple (site-level access),
- isolates sensitive content in a dedicated library,
- uses groups instead of individual permissions,
- minimizes unique permissions (permission scopes),
- and includes a repeatable validation/troubleshooting approach.

## Prerequisites
- You have a participant ID `Pxx` (P01–P10).
- You have a modern SharePoint site named `NW-Pxx-ProjectSite`.
- You can access:
  - SharePoint admin center (read-only verification is OK)
  - Your `NW-Pxx-ProjectSite` as a site owner

## Shared-tenant safety rules
- Do not change tenant-wide sharing policies unless the step is explicitly marked **Trainer-only**.
- Only modify content inside your own `NW-Pxx-ProjectSite`.
- Do not add other participants to your site unless the trainer explicitly approves a buddy test.

---

## Exercise 1 — Design worksheet (15–20 minutes)

### 1A. Define the business requirement (Northwind)
Northwind has a contracts workflow where drafts are edited by a small team, while final contracts must be broadly readable but not editable.

Define the roles:
- **Contracts Owners**: can manage permissions and content
- **Contracts Editors**: can create/edit drafts
- **Contracts Readers**: can read final contracts

### 1B. Decide the permission boundaries
In the table below, decide where to apply permissions.

| Object | Inherit from site? | Unique permissions? | Who should have access? |
|---|---:|---:|---|
| Site (`NW-Pxx-ProjectSite`) | Yes | No | Owners/Members/Visitors as normal |
| Library (`NW-Pxx-Contracts`) |  |  |  |
| Folder `01-Drafts` |  |  |  |
| Folder `03-Final` |  |  |  |

Rule of thumb for this lab:
- Break inheritance at the **library**.
- Break inheritance at **one folder** only (to demonstrate and keep scope count low).

Validation check:
- You can explain *why* you chose site vs library vs folder as the boundary.

---

## Exercise 2 — Baseline: understand the current site permissions (10–15 minutes)

### 2A. Confirm whether the site is group-connected
1. Open your site: `NW-Pxx-ProjectSite`.
2. Select **Settings (gear)** > **Site permissions**.
3. Observe whether permissions are tied to a Microsoft 365 group (group-connected team site behavior).

Record:
- Site type observation (group-connected vs not): ________

### 2B. Identify the default groups
From the site permissions panel (or Advanced permissions if available):
- Identify Owners / Members / Visitors.

Validation check:
- You can locate where Owners/Members/Visitors are managed for this site.

---

## Exercise 3 — Create the Contracts library and sample content (10–15 minutes)

### 3A. Create a dedicated library
1. In `NW-Pxx-ProjectSite`, create a new **document library** named:
   - `NW-Pxx-Contracts`
2. Create folders:
   - `01-Drafts`
   - `02-InReview`
   - `03-Final`
3. Upload or create 2–3 sample files (simple placeholders are fine), for example:
   - `Contract-Draft-001.docx`
   - `Contract-Final-001.docx`

Validation check:
- The library exists and contains the three folders and at least one file.

---

## Exercise 4 — Implement the permission model (25–35 minutes)

### 4A. Break inheritance on the library
Goal: isolate contracts access from broad site membership.

1. Open the library `NW-Pxx-Contracts`.
2. Open **Manage access** (or the equivalent permissions UI for the library).
3. Stop inheriting permissions (break inheritance) so the library has **unique permissions**.

Important:
- Ensure you (as the site owner) keep Full Control so you cannot lock yourself out.

### 4B. Create SharePoint groups for the library
Create these SharePoint groups (site-scoped groups are fine):
- `NW-Pxx-Contracts Owners`
- `NW-Pxx-Contracts Editors`
- `NW-Pxx-Contracts Readers`

Add members (shared-tenant safe default):
- Add **your own account** to all three groups for now (this is only to allow you to validate and troubleshoot without involving other participants).

### 4C. Grant permissions to the library
Grant these permission levels at the library level:
- `NW-Pxx-Contracts Owners` → **Full Control**
- `NW-Pxx-Contracts Editors` → **Edit**
- `NW-Pxx-Contracts Readers` → **Read**

Remove broad access if present:
- If your site Members/Visitors were carried into the library after breaking inheritance, remove them so access is only via the Contracts groups (and any required owners).

Validation check:
- The library permissions list shows only the Contracts groups (and any required owners).

---

## Exercise 5 — Folder-level drill (unique permissions with minimal scope count) (15–20 minutes)

Goal: demonstrate when folder-level unique permissions are appropriate.

### 5A. Break inheritance on one folder only
Choose one folder:
- `03-Final`

Then:
1. Break inheritance on `03-Final`.
2. Set permissions so that:
   - `NW-Pxx-Contracts Readers` has **Read**
   - `NW-Pxx-Contracts Editors` has **Read** (not Edit) for the final folder
   - `NW-Pxx-Contracts Owners` has **Full Control**

Explain in one sentence:
- Why are editors read-only in `03-Final`? ________

Validation check:
- Permissions differ between `01-Drafts` (Edit for editors) and `03-Final` (Read for editors).

---

## Exercise 6 — Validate and troubleshoot (10–15 minutes)

### 6A. Validate with “Check Permissions”
Use the **Check Permissions** feature to validate:
- Your own account
- The three Contracts groups

Reference path (may vary slightly by experience):
- Settings (gear) > Site permissions > Advanced permissions > **Check Permissions**

If you can’t find it in the modern UI, use the official troubleshooting steps:
- https://learn.microsoft.com/en-us/troubleshoot/sharepoint/administration/access-denied-or-need-permission-error-sharepoint-online-or-onedrive-for-business

Record results:
- Your account effective permissions on library: ________
- `NW-Pxx-Contracts Editors` effective permissions on `03-Final`: ________

### 6B. Validate via “Manage access” links view
1. On a sample file in `03-Final`, open **Manage access**.
2. Confirm access is coming from:
   - group permissions and/or
   - sharing links

Validation check:
- You can explain whether access is granted by group membership or by a link.

---

## Exercise 7 — Sharing link drill (internal-only) (10–15 minutes)

Goal: practice safe collaboration without changing tenant-wide policy.

1. Pick a sample file (preferably in `01-Drafts`).
2. Share it using a safe link type:
   - **People in your organization** (if available) OR
   - **Specific people** (select yourself, or a trainer-provided internal test account)
3. Open **Manage access** and locate the link.
4. Remove/revoke the link.

Validation check:
- The link no longer appears in Manage access.

---

## Cleanup / end state
Leave the following in place for later modules:
- Library: `NW-Pxx-Contracts`
- Folders: `01-Drafts`, `02-InReview`, `03-Final`
- SharePoint groups: `NW-Pxx-Contracts Owners/Editors/Readers`

Remove:
- Any ad-hoc sharing links you created during Exercise 7.

---

## Troubleshooting (common issues)

1) **I can’t find “Check Permissions”.**
- Use the Advanced permissions view if available.
- Confirm you are a site owner.
- Use the official troubleshooting article path as guidance:
  - https://learn.microsoft.com/en-us/troubleshoot/sharepoint/administration/access-denied-or-need-permission-error-sharepoint-online-or-onedrive-for-business

2) **I locked myself out of the library or folder.**
- As a site owner, regain access by re-adding yourself (or your admin account) with Full Control at the library level.
- Avoid removing Owners until you confirm your access.

3) **Permissions keep changing when I edit the parent site.**
- That’s inheritance. Confirm you broke inheritance at the intended boundary.

4) **Too many unique permissions (hard to manage).**
- Prefer library-level boundaries over many folder/item exceptions.
- Use groups to reduce individual assignments.
- Review permission scope guidance:
  - https://learn.microsoft.com/en-us/sharepoint/manage-permission-scope
