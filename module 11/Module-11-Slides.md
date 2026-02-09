# Module 11 — Operations at Scale (Slides Outline)

## Slide 1 — Module title + outcomes
- Monitoring, Lifecycle, External Sharing Governance
- Outcomes: triage, document, report, govern safely

## Slide 2 — Incident workflow: first 5 minutes
- Check Service health (incident/advisory)
- Check Message center (planned maintenance / changes)
- Decide: wait + communicate vs troubleshoot locally

References:
- Service health: https://learn.microsoft.com/en-us/microsoft-365/enterprise/view-service-health
- Message center: https://learn.microsoft.com/en-us/microsoft-365/admin/manage/message-center

## Slide 3 — Access issues: common causes
- Wrong identity for the link
- Permission misconfiguration
- Guest lifecycle / directory mismatch
- Use Check Permissions + diagnostics

Reference:
- Access denied troubleshooting: https://learn.microsoft.com/en-us/troubleshoot/sharepoint/administration/access-denied-or-need-permission-error-sharepoint-online-or-onedrive-for-business

## Slide 4 — Locked/read-only sites
- Lock state vs maintenance
- Check Message center + Service health

Reference:
- Read-only error messages: https://learn.microsoft.com/en-us/troubleshoot/sharepoint/sites/site-is-read-only

## Slide 5 — Site lifecycle: delete/restore model
- Deleted sites retention window
- Restore from SharePoint admin center
- Root site mistakes are high impact

Reference:
- Restore deleted sites: https://learn.microsoft.com/en-us/sharepoint/restore-deleted-site-collection

## Slide 6 — External sharing governance
- Org-level setting is the ceiling
- Link types: Anyone / Org / Specific people
- Defaults and Anyone-link restrictions

References:
- Manage sharing settings: https://learn.microsoft.com/en-us/sharepoint/turn-external-sharing-on-or-off
- Shareable links: https://learn.microsoft.com/en-us/sharepoint/shareable-links-anyone-specific-people-organization

## Slide 7 — Operational reports (safe)
- Site inventory export (Get-SPOSite → Export-Csv)
- Membership export (Get-SPOUser)

Reference:
- Manage users/groups with PowerShell: https://learn.microsoft.com/en-us/microsoft-365/enterprise/manage-sharepoint-users-and-groups-with-powershell

## Slide 8 — Lab briefing
- Incident worksheet (what you checked)
- Lifecycle checklist
- External sharing governance worksheet
- Safe reporting exports

## Slide 9 — Wrap-up
- “Observe → document → request changes”
- Shared tenant: tenant-wide changes are trainer-led

---

## Trainer demo script (minimal talk-track)
Use this as a short talk-track while showing the admin centers and setting expectations.

1) Set expectations (30–45 seconds)
- “Module 11 is about operations: observe, document, and raise change requests.”
- “In a shared tenant, we avoid policy changes; trainer does tenant-wide changes.”

2) Incident triage muscle memory (60–90 seconds)
- Open Service health and show how to read: affected services, user impact, last updated.
- Open Message center and show how to spot ‘Plan for change’ items.

3) Access issue triage (60–90 seconds)
- Emphasize identity mismatch vs permission vs link type.
- Show Check Permissions on a site/item and document findings.

4) Lifecycle guardrail (30–45 seconds)
- “Restore deleted sites is an admin workflow; root site mistakes are high impact.”
- Keep delete/restore actions trainer-led unless explicitly assigned.

5) Sharing governance (60–90 seconds)
- Explain org-level is the ceiling; sites can only be more restrictive.
- Link types: Anyone / People in your org / Specific people; default link type matters.

6) Close the loop (30 seconds)
- “Turn your findings into a change request: what/why/risk/rollback.”
