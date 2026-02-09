# Module 11 — Operations at Scale: Monitoring, Lifecycle, and External Sharing Governance (Student Guide)

## Learning objectives
You will learn to:
1. Use Microsoft 365 Service health and Message center for incident awareness.
2. Triage common SharePoint/OneDrive access failures and apply safe fixes.
3. Explain site lifecycle controls (delete/restore) and operational guardrails.
4. Understand external sharing link types and governance options.

---

## 1) Monitoring and incident readiness
### Service health and Message center
During an incident, start by checking whether Microsoft has an active advisory/incident.

Official references:
- How to check Microsoft 365 service health: https://learn.microsoft.com/en-us/microsoft-365/enterprise/view-service-health
- Message center: https://learn.microsoft.com/en-us/microsoft-365/admin/manage/message-center

Operational guidance (training)
- If it’s in Service health, avoid duplicate troubleshooting—track updates.
- If it’s not in Service health, use “report an issue” or open a support request.

---

## 2) Troubleshooting common SharePoint/OneDrive symptoms
### Access denied / need permission
Typical admin checks:
- verify permission level expectation
- use **Check Permissions** on the site
- ensure the user is the intended identity for the shared link

Official references:
- Access denied / need permission errors: https://learn.microsoft.com/en-us/troubleshoot/sharepoint/administration/access-denied-or-need-permission-error-sharepoint-online-or-onedrive-for-business
- SharePoint/OneDrive self-help diagnostics (Check User Access): https://learn.microsoft.com/en-us/troubleshoot/sharepoint/diagnostics/sharepoint-and-onedrive-diagnostics

### Read-only / locked site messages
Official reference:
- SharePoint or OneDrive read-only error messages: https://learn.microsoft.com/en-us/troubleshoot/sharepoint/sites/site-is-read-only

Shared-tenant note
- Prefer read-only verification in admin centers; document findings in the worksheet.

---

## 3) Site lifecycle governance (delete/restore)
Deleted sites are retained for a limited time and can be restored by admins.

Official reference:
- Restore deleted sites: https://learn.microsoft.com/en-us/sharepoint/restore-deleted-site-collection

Important operational note
- Deleting the root site can make other sites inaccessible until restored.

Official reference:
- Root site impact and recovery context: https://learn.microsoft.com/en-us/troubleshoot/sharepoint/sites/url-that-resides-under-root-site-collection-is-broken

---

## 4) External sharing governance (SharePoint + OneDrive)
Key concepts:
- organization-level sharing setting provides a ceiling
- sites can be further restricted, but not made more permissive than org-level
- link types: **Anyone**, **People in your organization**, **Specific people**

Official references:
- Manage sharing settings (org-level + link defaults): https://learn.microsoft.com/en-us/sharepoint/turn-external-sharing-on-or-off
- How shareable links work: https://learn.microsoft.com/en-us/sharepoint/shareable-links-anyone-specific-people-organization
- Plan sharing and collaboration options: https://learn.microsoft.com/en-us/sharepoint/collaboration-options

Shared-tenant rule
- Any tenant-wide sharing changes are trainer-led.

---

## 5) Reporting and change control (the admin habit)
Even when you don’t change settings, you can provide value with:
- inventory exports
- “what changed / when / who approved” notes
- clear rollback steps

PowerShell reporting reference pattern:
- Manage SharePoint users and groups with PowerShell: https://learn.microsoft.com/en-us/microsoft-365/enterprise/manage-sharepoint-users-and-groups-with-powershell
