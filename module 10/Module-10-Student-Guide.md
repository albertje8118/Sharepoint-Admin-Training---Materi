# Module 10 — Administration and Automation with PowerShell (Student Guide)

## Learning objectives
By the end of this module, you will be able to:
1. Connect to SharePoint Online using SharePoint Online PowerShell.
2. Generate and export basic admin reports (sites, users/groups).
3. Run safe bulk operations using CSV-driven scripts.
4. Connect to Microsoft Graph using Microsoft Graph PowerShell and understand scopes.

---

## 1) Why PowerShell matters for SharePoint/OneDrive admins
The admin centers are great for interactive work, but PowerShell is better when you need:
- repeatability (same steps across many sites)
- reporting (export to CSV/text)
- controlled bulk actions (carefully scoped)

Shared-tenant mindset:
- Scripts can impact many users quickly.
- Always start with **read-only reporting**.
- Scope by URL prefix (`NW-Pxx-...`) before making any changes.

---

## 2) SharePoint Online Management Shell: connect and context
In SharePoint Online PowerShell, you typically connect to your tenant admin endpoint and then run SharePoint Online cmdlets.

Core cmdlet:
- `Connect-SPOService` connects your PowerShell session to the SharePoint Online admin center.

Reference:
- `Connect-SPOService`: https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/connect-sposervice

Operational note:
- Only one connection is maintained per PowerShell session (connecting again replaces the previous connection).

---

## 3) Reporting patterns (the admin toolbox)
Useful patterns you’ll use repeatedly:
- Filter early (`Where-Object`) to scope down to your target
- Select the columns you need (`Select-Object`)
- Export cleanly (`Export-Csv -NoTypeInformation`)
- Save human-readable outputs (`Out-File`)

Example report concept (users in a site):
- `Get-SPOUser -Site <siteUrl>`

References:
- `Get-SPOUser`: https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/get-spouser
- User/group management + reporting examples: https://learn.microsoft.com/en-us/microsoft-365/enterprise/manage-sharepoint-users-and-groups-with-powershell

---

## 4) Bulk operations (CSV-driven)
Bulk operations are powerful but risky.
Safe approach:
1. Export the current state (baseline)
2. Prepare a small CSV
3. Run the change on a single safe target
4. Expand only when validated

Microsoft shows a pattern using `Import-Csv` piped into a `ForEach` block to apply a cmdlet across many rows.

Reference example:
- https://learn.microsoft.com/en-us/microsoft-365/enterprise/manage-sharepoint-users-and-groups-with-powershell

---

## 5) Microsoft Graph PowerShell (tenant-wide API access)
Microsoft Graph PowerShell is used when you need Graph endpoints and permissions.

Core cmdlet:
- `Connect-MgGraph` authenticates to Microsoft Graph.

Key concept: scopes
- In delegated mode, you request scopes like `User.Read.All`.

References:
- `Connect-MgGraph`: https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.authentication/connect-mggraph
- Microsoft 365 + Graph PowerShell connection guidance: https://learn.microsoft.com/en-us/microsoft-365/enterprise/connect-to-microsoft-365-powershell

Shared-tenant note:
- Graph scopes can require admin consent. Treat Graph write operations as trainer-led unless explicitly assigned.

---

## References (official)
- SharePoint Online PowerShell: https://learn.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online
- `Connect-SPOService`: https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/connect-sposervice
- `Get-SPOUser`: https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/get-spouser
- Manage SharePoint users and groups with PowerShell: https://learn.microsoft.com/en-us/microsoft-365/enterprise/manage-sharepoint-users-and-groups-with-powershell
- `Connect-MgGraph`: https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.authentication/connect-mggraph
- Connect to Microsoft 365 with Graph PowerShell: https://learn.microsoft.com/en-us/microsoft-365/enterprise/connect-to-microsoft-365-powershell
