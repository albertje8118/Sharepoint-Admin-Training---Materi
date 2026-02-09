# Lab 11 — Operations at Scale (Monitoring, Lifecycle, External Sharing Governance)

## Goal
Practice an operations workflow that combines:
- incident awareness (Service health + Message center)
- troubleshooting triage (permissions / lock state)
- lifecycle governance understanding (deleted sites restore)
- external sharing governance review (link types, defaults)
- safe reporting (PowerShell exports)

## Estimated time
75–105 minutes

## Prerequisites
### Roles and access
- Participants
  - Read-only access to Microsoft 365 admin center and SharePoint admin center is sufficient for observation tasks.
  - Site Owner/Admin access to your `NW-Pxx-ProjectSite`.
- Trainer
  - SharePoint Administrator (or higher) for any tenant-wide demonstrations.

### Lab worksheets (from your participant pack)
Use these templates in your generated pack:
- `TXT-Templates/M11-Incident-Triage-Worksheet.txt`
- `TXT-Templates/M11-Site-Lifecycle-Checklist.txt`
- `TXT-Templates/M11-External-Sharing-Governance-Worksheet.txt`
- `TXT-Templates/M11-Change-Request-Template.txt`

---

## Exercise 1 — Incident awareness workflow (read-only)
1. In Microsoft 365 admin center, open **Service health** and record:
   - any active advisories/incidents for SharePoint/OneDrive
   - last update time and user impact summary
2. Open **Message center** and record:
   - any relevant “Plan for change” messages that could affect SharePoint/OneDrive

References:
- Service health: https://learn.microsoft.com/en-us/microsoft-365/enterprise/view-service-health
- Message center: https://learn.microsoft.com/en-us/microsoft-365/admin/manage/message-center

Validation check:
- Worksheet has at least 1 entry (or “No active issues observed”).

---

## Exercise 2 — Triage an access issue (safe, training-scoped)
Scenario (training): a user reports “Access Denied” to a file/site.

1. In your own `NW-Pxx-ProjectSite`, identify a library/item to test.
2. Use **Check Permissions** to confirm the expected access.
3. Record what you checked and the result.

References:
- Access denied troubleshooting: https://learn.microsoft.com/en-us/troubleshoot/sharepoint/administration/access-denied-or-need-permission-error-sharepoint-online-or-onedrive-for-business
- Diagnostics overview (Check User Access): https://learn.microsoft.com/en-us/troubleshoot/sharepoint/diagnostics/sharepoint-and-onedrive-diagnostics

Validation check:
- You can explain whether the issue is identity mismatch vs permission vs link type.

---

## Exercise 3 — Site lifecycle governance (delete/restore concept)
This is primarily an observation-and-planning exercise in a shared tenant.

1. Review the retention/restore model for deleted sites.
2. Record:
   - who can restore sites
   - retention period and why it matters
   - what is risky about the root site

References:
- Restore deleted sites: https://learn.microsoft.com/en-us/sharepoint/restore-deleted-site-collection
- Root site impact context: https://learn.microsoft.com/en-us/troubleshoot/sharepoint/sites/url-that-resides-under-root-site-collection-is-broken

Validation check:
- Your checklist includes a “don’t delete root site” guardrail.

---

## Exercise 4 — External sharing governance review (read-only)
1. In SharePoint admin center, review organization-level sharing setting.
2. Record:
   - external sharing level
   - default link type
   - Anyone-link expiration/permissions (if applicable)
3. (Optional) For your own `NW-Pxx-ProjectSite`, record what site-level sharing can be set to (do not change).

References:
- Manage sharing settings: https://learn.microsoft.com/en-us/sharepoint/turn-external-sharing-on-or-off
- Shareable links explained: https://learn.microsoft.com/en-us/sharepoint/shareable-links-anyone-specific-people-organization

Validation check:
- You can state which link types are allowed and what the default is.

---

## Exercise 5 — Safe reporting (PowerShell) (participant-safe)
> Run only read-only commands.

1. Connect to SharePoint Online PowerShell:
   - `Connect-SPOService -Url https://<tenant>-admin.sharepoint.com`
2. Export a `NW-Pxx` site inventory:
   - `Get-SPOSite -Limit All | Where-Object { $_.Url -match "NW-P" } | Select-Object Url, Title, StorageUsageCurrent | Export-Csv .\NW-Pxx-Sites.csv -NoTypeInformation`
3. Export user membership for your own site:
   - `$siteUrl = "https://<tenant>.sharepoint.com/sites/NW-Pxx-ProjectSite"`
   - `Get-SPOUser -Site $siteUrl | Select-Object LoginName, DisplayName, IsSiteAdmin | Export-Csv .\NW-Pxx-Users.csv -NoTypeInformation`

References:
- Connect to SharePoint Online PowerShell: https://learn.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online
- Manage users/groups with PowerShell: https://learn.microsoft.com/en-us/microsoft-365/enterprise/manage-sharepoint-users-and-groups-with-powershell

Validation check:
- Your exports are limited to your own `NW-Pxx` scope.

---

## Wrap-up checkpoint
You should be able to:
- explain where to look first during incidents
- describe a safe triage flow for access errors
- explain deleted site restore model and root-site risk
- describe link types and default link governance
- produce a basic inventory + membership export

## Deliverables
- Completed Module 11 worksheets:
   - `M11-Incident-Triage-Worksheet.txt`
   - `M11-Site-Lifecycle-Checklist.txt`
   - `M11-External-Sharing-Governance-Worksheet.txt`
   - `M11-Change-Request-Template.txt`
- Optional exports (if completed):
   - `NW-Pxx-Sites.csv`
   - `NW-Pxx-Users.csv`

## Troubleshooting (common)
- If users report “Access Denied”: validate identity vs the link target account and use **Check Permissions** before changing access.
- If you suspect a tenant issue: check **Service health** and **Message center** first; don’t burn time on local troubleshooting if it’s a known incident.
- If something looks read-only/locked: confirm site lock state/maintenance context; document what you observe and escalate to trainer for changes.
