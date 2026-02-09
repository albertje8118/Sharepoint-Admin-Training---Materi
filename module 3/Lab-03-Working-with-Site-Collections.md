# Lab 3: Working with Site Collections (Create, Manage, Restore)

**Estimated time:** 90–120 minutes  
**Lab type:** UI + light PowerShell (SharePoint admin center + SharePoint Online Management Shell)  
**Goal:** Create and manage participant-owned sites safely in a shared tenant; practice the delete/restore lifecycle; and use PowerShell to retrieve site properties.

## Scenario context (continuity)
This lab continues the scenario:
- Scenario overview: ../scenario/Lab-Scenario-Overview.md
- Project: **Project Northwind Intranet Modernization**

## Shared-tenant operating rules (critical)
- Only create, delete, restore, and configure sites that match your own naming prefix: `NW-Pxx-...`.
- **Never** delete or modify other participants’ sites.
- Tenant-wide settings changes are **Trainer-only**.

---

## Deliverables for this lab (what you must produce)
Create a short “Site Lifecycle Worksheet” in your notes named: `Module3-SiteLifecycle-Pxx`.

Minimum fields:
- Participant ID (Pxx)
- SharePoint admin URL you used (the `-admin` URL)
- Primary practice site:
  - Site name: `NW-Pxx-ProjectSite`
  - Site URL
  - Site type you chose (Team or Communication)
  - Listed owner/admins (from admin center)
- Day-to-day site admin checks (record what you see):
   - Membership/admin surfaces you can access (1–2 notes)
   - Site Activity signals you can view (Yes/No + 1 note)
   - Access requests enabled? (Yes/No/Not visible)
   - Access requests recipient (site owners / specific email / unknown)
- Storage observations:
  - Tenant storage management mode (Automatic vs Manual), if visible
  - Whether you can edit the site’s storage limit (Yes/No)
  - Whether storage notifications are enabled (Yes/No)
- Delete/restore drill:
  - Test site name: `NW-Pxx-RestoreTest`
  - Deleted? (Yes/No)
  - Restored? (Yes/No)
  - Any warning/notes you observed
- Access review practice (monthly/regular):
   - Site admin review notes (3–5 bullets)
   - SharePoint admin review notes (3–5 bullets)
- Recycle Bin drill:
   - Item restored from site Recycle Bin? (Yes/No)
- PowerShell verification:
  - Command(s) you ran (copy/paste)
  - Key properties you observed from `Get-SPOSite` (any 3)
- Troubleshooting drill answers (Exercise 6)

---

## Prerequisites
- Participant ID assigned by trainer: **P01–P10**
- SharePoint admin center access
- Permissions: SharePoint Administrator (or higher)
- Optional (for Exercise 5): SharePoint Online Management Shell installed and working

If you can’t run PowerShell in your environment:
- Complete the UI parts of the lab.
- For Exercise 5, do a “paper run” by writing the commands you *would* run and what you expect each command to return.

---

## Exercise 0 — Setup your worksheet (required)
1) Start a new note named: `Module3-SiteLifecycle-Pxx`.
2) Add the Minimum fields list (above) as headings.

### Expected result
- You have a structured place to record evidence and decisions.

### Validation check
- Your worksheet includes Participant ID (Pxx).

---

## Exercise 1 — Create your primary practice site (required)
Your primary practice site is the persistent sandbox for later modules.

Target name:
- `NW-Pxx-ProjectSite`

1) Go to **Active sites** in the SharePoint admin center.
2) Select **Create**.
3) Choose **Communication site** (preferred for this course) or **Team site** (only if your trainer asks you to use a group-connected site).
4) Provide:
   - Site name: `NW-Pxx-ProjectSite`
   - Owner: yourself (your admin account)
   - Any other required fields shown by your tenant
5) Create the site.

If the site address is not available (for example, name collision or redirect):
- Choose a slightly modified name, record the final URL, and notify the trainer.

### Expected result
- Your `NW-Pxx-ProjectSite` exists and appears in Active sites.

### Validation check
- Record the site URL from the Active sites list.

Reference:
- Create a site: https://learn.microsoft.com/en-us/sharepoint/create-site-collection

---

## Exercise 2 — Review site details (admins/membership/settings) (required)
1) In **Active sites**, select your `NW-Pxx-ProjectSite`.
2) Open the site details panel.
3) Record in your worksheet:
   - Listed owner
   - Any additional admins shown
   - Any visible sharing/setting summary

### Expected result
- You can locate the site’s admin surfaces and interpret what you’re allowed to manage.

### Validation check
- Your worksheet includes the listed owner/admins for the site.

---

## Exercise 2B — Day-to-day membership/admin management (admin center) (required)
Goal: practice the common “someone needs access/admin rights” workflow safely.

1) In SharePoint admin center **Active sites**, select your `NW-Pxx-ProjectSite`.
2) Select **Membership**.
3) Record in your worksheet what you can manage from here (examples: owners, members, site admins).

Optional (only if your trainer provides a test account to use):
- Add the provided account as an additional **site admin** for your site, select **Save**, then remove it again and **Save**.

### Expected result
- You can find the Membership surface and explain what can be managed there.

### Validation check
- Your worksheet includes 1–2 notes about Membership/admins.

Reference:
- Add/remove site admins (new SharePoint admin center): https://learn.microsoft.com/en-us/sharepoint/manage-site-collection-administrators#add-or-remove-site-admins-in-the-new-sharepoint-admin-center

---

## Exercise 2C — Review site activity (operational check) (required)
Goal: learn where admins can view site activity/usage signals.

1) In SharePoint admin center, open the site details panel for `NW-Pxx-ProjectSite`.
2) Locate the **Activity** tab (if present).
3) Record one thing you can observe.

If the Activity tab isn’t available:
- Record “Activity tab not available in this tenant” and continue.

### Expected result
- You can locate (or confirm absence of) the Activity view for a site.

### Validation check
- Your worksheet includes 1 activity observation (or a note that it isn’t available).

---

## Exercise 2D — Monthly access review (site admin + SharePoint admin) (required)
Goal: practice a lightweight, repeatable access review that you can run monthly (or on a schedule) to reduce “stale access” risk.

Important shared-tenant safety rule:
- In this course, you **review and document** access first. Only remove access if your trainer approves.

### Part A — Site admin access review (site-scoped)
Perform this review inside your `NW-Pxx-ProjectSite`.

1) Open `NW-Pxx-ProjectSite`.
2) Go to **Settings** (gear) > **Site permissions**.
3) Record in your worksheet:
   - Who are the site Owners? (names or count)
   - Who are the site Members? (names or count)
   - Who are the site Visitors? (names or count)
4) Open **Access Request Settings** and record whether requests are enabled and where they go.

Site admin “monthly checklist” (write answers as short bullets):
- Are Owners limited to people who truly administer the site?
- Are Members appropriate for who needs edit access?
- Are Visitors appropriate for read-only access?
- Is the access request recipient correct (not a departed employee / wrong mailbox)?

Northwind example findings (use these as realistic patterns to look for):
- “Too many Owners” for a project site (Owners should be a small set).
- A generic or personal mailbox is receiving access requests (should be a monitored owner/admin mailbox).
- Members include people who only need read access (move to Visitors).
- Visitors include accounts that no longer work on the Northwind intranet project (stale access).

References:
- Sharing and permissions (modern experience): https://learn.microsoft.com/en-us/sharepoint/modern-experience-sharing-permissions
- Access request settings (navigation + troubleshooting context): https://learn.microsoft.com/en-us/troubleshoot/sharepoint/sharing-and-permissions/request-access-to-resource

### Part B — SharePoint Administrator access review (admin center)
Perform this review from SharePoint admin center for your `NW-Pxx-ProjectSite`.

1) Go to **Sites** > **Active sites**.
2) Find and select your `NW-Pxx-ProjectSite`.
3) Review these tabs/areas (what you see depends on the tenant):
   - **Membership**: site admins/owners/members/visitors
   - **Activity**: activity signals (if available)
   - **Settings**: note any high-risk settings visible (observe-first)
4) Record 3–5 bullets in your worksheet describing what you would flag for follow-up (examples: too many admins, unexpected owners, unusual activity, external sharing posture).

Northwind example findings (SharePoint admin perspective):
- Additional site admins include accounts outside the Northwind admin team (unnecessary elevated access).
- Owner is unexpected (for example, not the designated `Pxx` owner) or ownership looks “orphaned”.
- Activity signals are inconsistent with the project phase (example: sudden spike in sharing/external collaboration activity).
- Access requests are enabled but route to the wrong recipient.

Reference:
- Manage sites (Membership/Activity in new admin center): https://learn.microsoft.com/en-us/sharepoint/manage-sites-in-new-admin-center

### Optional (advanced, tenant feature): Site access reviews for oversharing
If (and only if) your tenant has **Data access governance** reports available, SharePoint admins can initiate **site access reviews** for site owners of overshared sites.

Reference:
- Initiate site access reviews (Data access governance): https://learn.microsoft.com/en-us/sharepoint/site-access-review

### Expected result
- You can perform a repeatable review and produce a short set of findings.

### Validation check
- Your worksheet includes both: site admin review notes + SharePoint admin review notes.

---

## Exercise 3 — Storage observation and (safe) configuration (required)
Goal: learn what storage controls are available without changing tenant-wide settings.

1) In SharePoint admin center, open your site details panel (still on `NW-Pxx-ProjectSite`).
2) Locate:
   - Current storage usage
   - Storage limit field (if shown)
3) Record:
   - Whether you can edit the storage limit
   - Whether notifications are enabled (if shown)

If your tenant is using **manual** storage limits and you can edit the limit:
- Do **not** reduce the limit below current usage.
- Prefer enabling notifications (if not already enabled).

### Expected result
- You can explain whether your tenant uses pooled automatic storage or manual per-site quotas.

### Validation check
- Your worksheet includes “Can edit site storage limit: Yes/No”.

Reference:
- Manage site storage limits: https://learn.microsoft.com/en-us/sharepoint/manage-site-collection-storage-limits

---

## Exercise 3B — Access requests (day-to-day site admin operation) (required, observe-first)
Access requests are a common operational topic: users get “Access denied”, then ask “how do I request access?”

Default approach in this course:
- Observe and document first.
- Only change access request settings if your trainer approves.

1) Open your `NW-Pxx-ProjectSite` in the browser.
2) Go to **Settings** (gear icon) > **Site settings**.
3) Under **Users and Permissions**, select **Site permissions**.
4) Select **Access Request Settings**.
5) Record in your worksheet:
   - Is **Allow Access Requests** enabled?
   - Who receives requests (site owners vs a specific email), if shown?

Optional (only if your trainer approves):
- Enable access requests and set the recipient to your own email.

### Expected result
- You can locate and interpret access request settings for your site.

### Validation check
- Your worksheet includes the access request status and recipient.

Reference:
- Access requests (navigation + troubleshooting context): https://learn.microsoft.com/en-us/troubleshoot/sharepoint/sharing-and-permissions/request-access-to-resource

---

## Exercise 4 — Delete + restore drill using a test site (required)
Do **not** use your primary practice site for this drill.

Target test site:
- `NW-Pxx-RestoreTest`

### Part A — Create the test site
1) In **Active sites**, select **Create**.
2) Create a **Communication site** named: `NW-Pxx-RestoreTest`.
3) Record the URL in your worksheet.

### Part B — Delete the test site
1) In **Active sites**, select `NW-Pxx-RestoreTest`.
2) Select **Delete**, confirm **Delete**.

### Part C — Restore the test site
1) Go to **Deleted sites** in the SharePoint admin center.
2) Select `NW-Pxx-RestoreTest`.
3) Select **Restore**.

### Expected result
- You can complete the full lifecycle: create → delete → restore.

### Validation check
- Confirm the site returns to **Active sites** (it may take a short time).

References:
- Delete a site: https://learn.microsoft.com/en-us/sharepoint/delete-site-collection
- Restore deleted sites: https://learn.microsoft.com/en-us/sharepoint/restore-deleted-site-collection

---

## Exercise 4B — Recycle Bin recovery drill (site admin) (required)
Goal: practice the day-to-day “I deleted a file/page by accident” recovery workflow.

1) In your `NW-Pxx-ProjectSite`, open the default **Documents** library (or another library available).
2) Upload a small test file named: `NW-Pxx-RecycleBin-Test.txt`.
3) Delete the file.
4) Open the site **Recycle Bin**.
5) Restore the file.
6) Confirm the file returns to the library.

Optional (if you want to understand second-stage behavior):
- Delete the file again.
- Remove it from the first-stage recycle bin.
- Check the second-stage recycle bin (if available) and restore.

### Expected result
- You can recover a deleted item using the Recycle Bin.

### Validation check
- Your worksheet includes “Item restored from site Recycle Bin: Yes”.

References:
- Restore items from Recycle Bin: https://support.microsoft.com/office/restore-items-in-the-recycle-bin-that-were-deleted-from-sharepoint-or-teams-6df466b6-55f2-4898-8d6e-c0dff851a0be
- Manage the Recycle Bin: https://support.microsoft.com/office/manage-the-recycle-bin-of-a-sharepoint-site-8a6c2198-910e-42dc-9a9c-bc5bc4f327da

---

## Exercise 5 — PowerShell verification (light) (required, with fallback)
Goal: connect and retrieve properties for your site.

### Part A — Connect
1) Open PowerShell.
2) Connect to your tenant admin URL:

```powershell
Connect-SPOService -Url https://<tenant>-admin.sharepoint.com -UseSystemBrowser $true
```

Record the admin URL you used.

Reference:
- Connect-SPOService: https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/connect-sposervice?view=sharepoint-ps

### Part B — Retrieve properties for your site
Run `Get-SPOSite` for your **own** site only:

```powershell
Get-SPOSite -Identity "https://<tenant>.sharepoint.com/sites/NW-Pxx-ProjectSite"
```

In your worksheet, record any three properties you can see (examples: `Url`, `Owner`, `StorageQuota`, `StorageUsageCurrent`, `SharingCapability`).

Reference:
- Get-SPOSite: https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/get-sposite?view=sharepoint-ps

If you can’t run PowerShell:
- Write down the two commands above and a one-sentence “expected outcome” for each.

### Expected result
- You can connect and retrieve site properties without changing tenant-wide settings.

### Validation check
- Your worksheet includes the command output observations (at least 3 properties).

### Optional (advanced): set a backup site collection admin with PowerShell
Only do this if your trainer provides a test account to use.

```powershell
Set-SPOUser -Site "https://<tenant>.sharepoint.com/sites/NW-Pxx-ProjectSite" -LoginName "<userUPN>" -IsSiteCollectionAdmin $true
```

Reference:
- Set-SPOUser: https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/set-spouser?view=sharepoint-ps

---

## Trainer-only demo script (5–10 minutes): assign and remove a backup site admin
Use this as a short live-demo so participants see the workflow without risking cross-participant changes.

### Prerequisites
- A **trainer-owned** target site (recommended) such as `NW-TRN-ProjectSite`, or a volunteer participant’s `NW-Pxx-ProjectSite` (with consent).
- A test user account to temporarily grant admin rights to (do not use a real participant unless intended).

### Safety rules
- Demo must target **one** known site URL and be reversed immediately after validation.
- Do not grant permanent admin rights broadly in a shared-tenant class.

### Steps
1) Connect:

```powershell
Connect-SPOService -Url https://<tenant>-admin.sharepoint.com -UseSystemBrowser $true
```

2) Assign backup admin:

```powershell
Set-SPOUser -Site "https://<tenant>.sharepoint.com/sites/NW-TRN-ProjectSite" -LoginName "<testUserUPN>" -IsSiteCollectionAdmin $true
```

3) Validate (one of the following):
- In SharePoint admin center > Active sites > select the site > **Membership**, confirm the test user appears as a site admin (UI may take a moment), OR
- Re-run `Get-SPOSite -Identity <siteUrl>` and confirm expected properties are returned.

4) Cleanup (remove the assignment):

```powershell
Set-SPOUser -Site "https://<tenant>.sharepoint.com/sites/NW-TRN-ProjectSite" -LoginName "<testUserUPN>" -IsSiteCollectionAdmin $false
```

---

## Exercise 6 — Troubleshooting drill (required)
For each scenario below, write the **first three checks** you would perform, in order, and what you expect to find.

Scenario 1: “I can’t create a site — the Create button is missing or greyed out.”

Scenario 2: “My site exists, but I can’t see it in Active sites.”

Scenario 3: “I deleted the test site, but it does not appear under Deleted sites.”

Scenario 4: “Users see Access denied, and the request access link is missing (or requests go to the wrong person).”

Rules:
- At least one of your checks must consider **permissions/roles**.
- At least one of your checks must consider **scope** (tenant setting vs site setting).

### Expected result
- You troubleshoot systematically instead of guessing.

### Validation check
- Your worksheet contains 3 scenarios x 3 checks.

---

## Cleanup
- No cleanup required for `NW-Pxx-ProjectSite`.
- For `NW-Pxx-RestoreTest`:
  - Leave it restored for the trainer to confirm outcomes, OR
  - If the trainer asks, delete it again (do not permanently delete unless instructed).

---

## Troubleshooting (common issues)
1) **Site creation options aren’t available**
   - Verify you have SharePoint Administrator role or equivalent.
   - Ask the trainer whether site creation is restricted in this tenant.

2) **PowerShell connect errors**
   - Verify you are using the `-admin` URL.
   - Ensure you can authenticate using the system browser method.

3) **Restore button missing**
   - Ensure only one deleted site is selected (some admin center actions are disabled for multi-select).
