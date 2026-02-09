# Module 3: Working with Site Collections (Modern Sites)

## Course lab scenario (continuity)
This course uses one shared Microsoft 365 tenant (trainer + 10 admin participants) and one scenario story.

- Scenario overview: ../scenario/Lab-Scenario-Overview.md

In this module:
- Participants create and manage their own practice sites using `NW-Pxx-...` naming.
- Tenant-wide configuration changes are treated as **Trainer-only**.

---

## Module objectives
After completing this module, you will be able to:
- Create a modern SharePoint site from the SharePoint admin center.
- Identify where site ownership/admins and site settings are managed.
- Perform common day-to-day site admin tasks (membership, access requests, recycle bin recovery).
- Explain and observe how storage limits work at the tenant and site level.
- Delete and restore a site safely (using a test site).
- Connect to SharePoint Online using PowerShell and retrieve site properties.

---

## 1. Modern “site collections” in SharePoint Online
In SharePoint Online, modern sites are administered as **sites** (formerly commonly called “site collections” in older terminology). The SharePoint admin center treats these as top-level manageable entities with their own:
- URL
- Owner/admins
- Sharing configuration (site-level sharing posture)
- Storage usage and (sometimes) storage limits

Key admin takeaway:
- Many governance and security controls are scoped to the **site** level, which is why each participant uses an isolated `NW-Pxx-ProjectSite`.

Reference:
- Create a site: https://learn.microsoft.com/en-us/sharepoint/create-site-collection

---

## 2. Create and manage sites in the SharePoint admin center
### Where admins work
In the SharePoint admin center, the **Active sites** list is the operational “control panel” for site lifecycle tasks:
- Create a site
- Find a site by name/URL
- Open the site details panel (General, Membership, Settings)
- Delete a site (and later restore it from Deleted sites)

### Day-to-day site admin operations (site-scoped)
In real operations, admins frequently handle “small but urgent” site issues. In this course, these are practiced only within your own `NW-Pxx-...` site:

- **Membership / site admins**: add or remove additional site admins (site-scoped, common for break-glass access).
- **Access requests**: verify whether requests are enabled and who receives them.
- **Recycle Bin recovery**: restore deleted files/pages when users make mistakes.

References:
- Manage site collection administrators (Membership panel): https://learn.microsoft.com/en-us/sharepoint/manage-site-collection-administrators#add-or-remove-site-admins-in-the-new-sharepoint-admin-center
- Access requests troubleshooting (includes navigation): https://learn.microsoft.com/en-us/troubleshoot/sharepoint/sharing-and-permissions/request-access-to-resource
- Restore items from the Recycle Bin: https://support.microsoft.com/office/restore-items-in-the-recycle-bin-that-were-deleted-from-sharepoint-or-teams-6df466b6-55f2-4898-8d6e-c0dff851a0be

### Monthly/regular access review (best practice)
As a best practice, review who has access on a regular cadence (for example monthly), especially for sites with sensitive content or frequent membership changes.

In this course, the “access review” practice is done in two layers:
- **Site admin layer** (inside the site): review Owners/Members/Visitors and ensure access requests route to the correct recipient.
- **SharePoint admin layer** (admin center): review Membership and Activity signals to spot unusual access patterns or unexpected admins.

Optional advanced (tenant feature):
- SharePoint admins can initiate **site access reviews** for overshared sites using Data access governance reports (if available in your tenant).

References:
- Manage sites in the new SharePoint admin center (Membership/Activity): https://learn.microsoft.com/en-us/sharepoint/manage-sites-in-new-admin-center
- Initiate site access reviews (Data access governance): https://learn.microsoft.com/en-us/sharepoint/site-access-review

Reference:
- Create a site (admin center): https://learn.microsoft.com/en-us/sharepoint/create-site-collection#create-a-team-site-or-communication-site

---

## 3. Storage: what you can (and can’t) control
SharePoint storage is managed at two levels:

1) **Tenant storage management mode**
- Many tenants use pooled storage and automatic management.
- Some tenants use manual per-site limits.

2) **Per-site storage settings**
- If the tenant is set to manual, admins can set a maximum storage limit (GB) per site and configure notifications.
- If the tenant is set to automatic, you may not be able to set a per-site limit.

Important admin nuance:
- In a shared tenant, changing the tenant storage management mode is **Trainer-only**.
- Viewing storage usage and reviewing the site’s storage settings is safe for participants.

Reference:
- Manage site storage limits: https://learn.microsoft.com/en-us/sharepoint/manage-site-collection-storage-limits

---

## 4. Delete and restore: the site lifecycle safety model
### What deletion means
Deleting a site removes access to the site and all its contents. In the SharePoint admin center, admins can delete modern and classic sites.

Important warning for shared tenant labs:
- Do not delete the organization’s root site.
- Do not delete any site you do not own.
- Use a dedicated test site (`NW-Pxx-RestoreTest`) for the delete/restore drill.

Reference:
- Delete a site: https://learn.microsoft.com/en-us/sharepoint/delete-site-collection
- Restore deleted sites: https://learn.microsoft.com/en-us/sharepoint/restore-deleted-site-collection

---

## 5. PowerShell (SharePoint Online Management Shell) — admin basics
PowerShell is useful for repeatable administrative tasks, bulk inspection, and scripted changes. In this module we focus on:
- Connecting safely
- Retrieving properties for a single site

Core workflow:
1) Connect to the SharePoint admin service
2) Run a read-only command (or target only your own site)

References:
- Connect-SPOService: https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/connect-sposervice?view=sharepoint-ps
- Get-SPOSite: https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/get-sposite?view=sharepoint-ps
- Set-SPOSite: https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/set-sposite?view=sharepoint-ps
- Remove-SPOSite: https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/remove-sposite?view=sharepoint-ps
- Restore-SPODeletedSite: https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/restore-spodeletedsite?view=sharepoint-ps

---

## Summary
In this module you established your own site collection “sandbox” in a shared tenant, practiced safe site lifecycle tasks (create, inspect, delete/restore), and introduced PowerShell connections and basic site inspection.

---

## Knowledge check (self-assessment)
1) Why do we use a dedicated `NW-Pxx-RestoreTest` site for the delete/restore drill?
2) When can you set a per-site storage limit in the SharePoint admin center?
3) What must you do before running `Get-SPOSite`?
4) Why is changing tenant-wide storage management mode considered Trainer-only in this course?

### Suggested answers
1) To avoid damaging the persistent practice site and to avoid impacting other participants.
2) Typically only when the tenant storage management option is set to manual.
3) Connect to the SharePoint Online admin service with `Connect-SPOService`.
4) Because it affects how storage is allocated across the tenant and can impact the entire class.
