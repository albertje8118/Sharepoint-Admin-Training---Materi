# Module 10 — Administration and Automation with PowerShell (Slides Outline)

## Slide 1 — Module title + outcomes
- Administration and Automation with PowerShell
- Outcomes:
  - Connect to SharePoint Online PowerShell
  - Generate reports and export outputs
  - Understand safe bulk operations
  - Connect to Graph PowerShell (scopes)

## Slide 2 — Admin principle: automation safely
- Start with read-only reporting
- Scope first (URL prefix, sites list)
- Export baseline before changes
- Shared tenant: avoid tenant-wide changes

## Slide 3 — SharePoint Online PowerShell: connect
- `Connect-SPOService -Url https://<tenant>-admin.sharepoint.com`
- One connection per session

Reference:
- https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/connect-sposervice

## Slide 4 — Reporting patterns
- `Get-SPOSite` → filter → `Select-Object` → `Export-Csv`
- `Get-SPOUser -Site <url>` → membership report

Reference:
- https://learn.microsoft.com/en-us/microsoft-365/enterprise/manage-sharepoint-users-and-groups-with-powershell

## Slide 5 — Bulk operations via CSV
- `Import-Csv` + `ForEach-Object` + cmdlet
- Run on one safe target first

## Slide 6 — Microsoft Graph PowerShell: connect + scopes
- `Connect-MgGraph -Scopes "User.Read.All"`
- Scopes govern what you can do
- Admin consent may be required

References:
- https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.authentication/connect-mggraph
- https://learn.microsoft.com/en-us/microsoft-365/enterprise/connect-to-microsoft-365-powershell

## Slide 7 — Lab briefing
- Export `NW-Pxx` sites report
- Export site users report
- CSV-driven loop
- Optional Graph connect

## Slide 8 — Wrap-up
- When to use SPO PowerShell vs Graph PowerShell
- Common troubleshooting: connection/auth errors, scoping mistakes

---

## Trainer demo script (minimal talk-track)
Use this as a short talk-track while showing the lab setup and first commands.

1) Safety framing (30–45 seconds)
- “In a shared tenant, we start with reporting only.”
- “Anything tenant-wide is trainer-led; your scope is your own `NW-Pxx` sites.”

2) Connect to SPO (45–60 seconds)
- Show the tenant admin URL format.
- Run `Connect-SPOService` and explain that it establishes session context for SPO cmdlets.

3) Show a safe filter first (60–90 seconds)
- Run `Get-SPOSite -Limit All`.
- Immediately filter to `NW-Pxx` URLs before exporting.
- Export to CSV and open the CSV to validate it contains only training-scoped sites.

4) Membership report (60–90 seconds)
- Run `Get-SPOUser -Site <NW-Pxx site>`.
- Export CSV and validate your account appears.

5) Optional Graph note (30 seconds)
- “Graph uses scopes; you might see consent prompts.”
- “We’ll avoid write scopes unless explicitly assigned.”
