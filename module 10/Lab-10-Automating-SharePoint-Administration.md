# Lab 10 — Automating SharePoint Administration

## Goal
Practice safe, repeatable admin automation:
- connect to SharePoint Online PowerShell
- generate site and permissions reports (export to CSV/text)
- perform a small, controlled bulk task scoped to `NW-Pxx` artifacts
- (optional) connect to Microsoft Graph PowerShell and validate identity data

## Estimated time
75–105 minutes

## Prerequisites
### Access and roles
This lab is designed for a **single shared tenant**.

- **Participant prerequisites (required)**
  - You can run PowerShell (Windows PowerShell or PowerShell 7).
  - You have permissions to manage your own `NW-Pxx` sites (Site Admin/Owner).

- **Trainer prerequisites (recommended)**
  - SharePoint Administrator (or higher) account for any tenant-wide demonstrations.
  - Admin consent capability if Graph scopes require approval.

References (official):
- `Connect-SPOService`: https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/connect-sposervice
- `Connect-MgGraph`: https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.authentication/connect-mggraph

### PowerShell modules (required)
Install these before starting Exercise 1.

#### Module A — SharePoint Online Management Shell (`Connect-SPOService`, `Get-SPOSite`, etc.)
Option 1 (most common): install the SharePoint Online Management Shell package.
1. Download and install: https://www.microsoft.com/download/details.aspx?id=35588

Validation check:
- Run: `Get-Command Connect-SPOService`
- Expected result: command details are returned (no “not recognized” error).

#### Module B — Microsoft Graph PowerShell (used in Exercise 5)
> Install this even if Exercise 5 is optional. It’s also used by the trainer in tenant setup scripts.

1. Install:
   - `Install-Module Microsoft.Graph -Scope CurrentUser -Repository PSGallery -Force`

Validation check:
- Run: `Get-Command Connect-MgGraph`

Reference:
- Install the Microsoft Graph PowerShell SDK: https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0

Troubleshooting (common)
- If `Install-Module` prompts about an untrusted repository (PSGallery), choose `Yes`.
- If you’re on Windows PowerShell and module install fails due to prerequisites, use PowerShell 7 (`pwsh`) for this lab (preferred).

### Shared-tenant safety rules (Module 10)
- Start with **read-only reporting**.
- Do not run tenant-wide changes.
- Any change commands must be scoped only to your own `NW-Pxx-...` artifacts.

### Lab worksheets (from your participant pack)
In your generated participant pack folder, use these TXT templates to capture outputs:
- `TXT-Templates/M10-PowerShell-Setup-Notes.txt`
- `TXT-Templates/M10-Reporting-Worksheet.txt`

---

## Exercise 1 — Connect to SharePoint Online PowerShell (participant)
### Task A — Connect
1. Open PowerShell.
2. Connect to your tenant admin endpoint:
   - `Connect-SPOService -Url https://<tenant>-admin.sharepoint.com`

Validation check:
- Your session connects successfully (no authentication error).

Reference:
- https://learn.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online

Troubleshooting (common)
- If you cannot connect, confirm you are using an up-to-date module and TLS 1.2.
  - https://learn.microsoft.com/en-us/troubleshoot/sharepoint/administration/errors-connecting-to-management-shell

---

## Exercise 2 — Generate a site inventory report (participant-safe)
### Task A — Export your `NW-Pxx` sites
1. Run a tenant sites query (read-only):
   - `Get-SPOSite -Limit All`
2. Filter to only your own sites by prefix (example):
   - `... | Where-Object { $_.Url -match "NW-Pxx" }`
3. Export the result to CSV:
   - `... | Select-Object Url, Title, StorageUsageCurrent, Template | Export-Csv .\NW-Pxx-Sites.csv -NoTypeInformation`

Validation check:
- You have a CSV file created locally with only `NW-Pxx`-scoped sites.

---

## Exercise 3 — Generate a permissions/user membership report for your own site
### Task A — Choose your target site
1. Identify your primary practice site URL (`NW-Pxx-ProjectSite`).
2. Set a variable:
   - `$siteUrl = "https://<tenant>.sharepoint.com/sites/NW-Pxx-ProjectSite"`

### Task B — Export users from the site
1. Run:
   - `Get-SPOUser -Site $siteUrl | Select-Object LoginName, DisplayName, IsSiteAdmin | Export-Csv .\NW-Pxx-Users.csv -NoTypeInformation`

Validation check:
- The CSV contains your own account and any expected site admins.

Reference:
- `Get-SPOUser`: https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/get-spouser

---

## Exercise 4 — Controlled bulk task using CSV (safe pattern)
This exercise practices the pattern without touching tenant-wide settings.

### Task A — Create a CSV of targets
Create a CSV file `Targets.csv` with one row and your own site URL:
- Column: `SiteUrl`
- Value: your `NW-Pxx-ProjectSite` URL

### Task B — Loop across the CSV and collect a report
Run:
- `Import-Csv .\Targets.csv | ForEach-Object { Get-SPOUser -Site $_.SiteUrl | Select-Object LoginName, DisplayName, IsSiteAdmin } | Export-Csv .\Targets-Users.csv -NoTypeInformation`

Validation check:
- You can re-run the command and get the same output (repeatable).

Reference pattern:
- Microsoft example of CSV-driven bulk operations: https://learn.microsoft.com/en-us/microsoft-365/enterprise/manage-sharepoint-users-and-groups-with-powershell

---

## Exercise 5 — Optional: connect to Microsoft Graph PowerShell
> Do this only if Microsoft Graph PowerShell is available and the trainer approves the scope.

### Task A — Connect (delegated)
1. Connect with a minimal scope for identity read:
   - `Connect-MgGraph -Scopes "User.Read.All"`

Validation check:
- You see a successful welcome/connection and can run Graph cmdlets.

Reference:
- `Connect-MgGraph`: https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.authentication/connect-mggraph

---

## Wrap-up checkpoint
You should be able to:
- connect to SharePoint Online PowerShell
- export a `NW-Pxx`-scoped site inventory report
- export a user membership report for your own site
- use CSV-driven loops safely

## Deliverables
- `NW-Pxx-Sites.csv`
- `NW-Pxx-Users.csv`
- `Targets.csv`
- `Targets-Users.csv`
