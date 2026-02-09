# Lab 09 — Configuring OneDrive Settings

## Goal
Practice OneDrive administration in a **shared-tenant-safe** way:
- locate key OneDrive settings in the SharePoint admin center
- document current tenant configuration (sharing, sync, storage, retention, access control)
- perform a safe internal OneDrive sharing test
- understand operational implications (replication delays, conditional access dependencies, user lifecycle)

## Estimated time
60–90 minutes

## Prerequisites
### Access and roles
This lab is designed for a **single shared tenant**.

- **Participant prerequisites (required)**
  - You can sign in to Microsoft 365.
  - You can access your OneDrive for Business.
  - You can access SharePoint admin center (read-only verification is fine).

- **Trainer prerequisites (recommended, tenant-wide)**
  - A trainer account with permissions to manage tenant-wide SharePoint/OneDrive settings.
  - If demonstrating PowerShell: SharePoint Online Management Shell and SharePoint Administrator role.

References (official):
- OneDrive settings are managed from SharePoint admin center: https://learn.microsoft.com/en-us/sharepoint/onedrive-overview#manage-onedrive
- Unmanaged devices access controls: https://learn.microsoft.com/en-us/sharepoint/control-access-from-unmanaged-devices
- Retention for deleted users: https://learn.microsoft.com/en-us/sharepoint/set-retention

### Shared-tenant safety rules (Module 9)
- **Trainer-only:** any change to tenant-wide Sharing/Sync/Storage/Retention/Access control settings.
- **Participants:** read-only verification + safe internal user-level tasks inside your own OneDrive.
- Do not share externally (no guests) unless the trainer explicitly authorizes a controlled test.

---

## Exercise 1 — Identify where OneDrive settings live (participant)
### Task A — Open the relevant admin pages
1. Open the SharePoint admin center.
2. Locate the OneDrive-related settings areas:
   - **Sharing**
   - **Settings** → (OneDrive sections such as Sync / Storage limit / Retention / Notifications)
   - **Access control**

Record what you see in your worksheet:
- Use `M09-OneDrive-Settings-Observation.txt` from your participant pack.

Validation check:
- You can find at least the following pages in the admin center: Sharing, Settings (OneDrive sections), Access control.

Reference:
- https://learn.microsoft.com/en-us/sharepoint/onedrive-overview#manage-onedrive

---

## Exercise 2 — Sharing settings (participant: observe + record)
### Task A — Observe sharing controls
1. In SharePoint admin center, open **Sharing**.
2. Record (do not change) the organization-level sharing state relevant to OneDrive/SharePoint.
3. Capture any defaults shown for link type or permissions (UI varies by tenant).

Validation check:
- Your worksheet contains the current sharing state and any visible defaults.

---

## Exercise 3 — Sync settings (participant: observe + record; trainer demo optional)
### Task A — Observe Sync settings
1. In SharePoint admin center, open **Settings**.
2. Select **Sync**.
3. Record any relevant settings shown (for example: whether Sync button is shown on OneDrive website; any restrictions; blocked file types).

Reference:
- Prevent users from installing the OneDrive sync app (Sync settings UI): https://learn.microsoft.com/en-us/sharepoint/prevent-installation

### Optional trainer demo — Hide the Sync button via PowerShell
This is **trainer-only** in a shared tenant.
- Microsoft notes you can hide the Sync button using SharePoint Online Management Shell:
  - `Set-SPOTenant -HideSyncButtonOnTeamSite $true`

Reference:
- https://learn.microsoft.com/en-us/sharepoint/sharepoint-sync

Validation check:
- Participants can explain the difference between “hide sync button” and “disable existing syncs” (hiding blocks new sync starts; existing syncs continue).

---

## Exercise 4 — Storage and retention (participant: observe + record)
### Task A — Observe OneDrive storage limit settings
1. In SharePoint admin center **Settings**, open the **Storage limit** area (if shown).
2. Record the default storage limit setting.

Reference:
- Set default OneDrive storage space: https://learn.microsoft.com/en-us/sharepoint/set-default-storage-space

### Task B — Observe retention for deleted users
1. In SharePoint admin center **Settings**, open **Retention**.
2. Record the current setting for “Days to retain files a deleted user’s OneDrive”.

Reference:
- Set retention for deleted users (30–3650 days): https://learn.microsoft.com/en-us/sharepoint/set-retention

Validation check:
- Your worksheet includes both storage-limit notes and the current deleted-user retention days.

---

## Exercise 5 — Access control (unmanaged devices) (participant: observe; trainer demo optional)
### Task A — Observe the unmanaged devices policy
1. In SharePoint admin center, open **Access control**.
2. Select **Unmanaged devices**.
3. Record which option is currently selected:
   - Allow full access
   - Allow limited, web-only access
   - Block access

Operational notes to record:
- Microsoft states changes can take up to **24 hours** to take effect and won’t affect users already signed in.
- Microsoft recommends also blocking access from apps that don’t use modern authentication to reduce bypass risk.

Reference:
- Control access from unmanaged devices: https://learn.microsoft.com/en-us/sharepoint/control-access-from-unmanaged-devices

Validation check:
- You can explain what a user can/can’t do on an unmanaged device under the selected policy.

---

## Exercise 6 — Safe OneDrive sharing test (participant)
### Task A — Create a test file in your OneDrive
1. Open your OneDrive for Business.
2. Create a new file named: `NW-Pxx-OneDrive-Sharing-Test.txt` (or a Word document).
3. Paste the content from your template: `M09-OneDrive-Sharing-Test.txt`.

### Task B — Share internally (no external sharing)
1. Share the file with **one internal class account** (another participant or the trainer).
2. Record:
   - Sharing method (specific person vs link)
   - Permissions granted (view/edit)
   - Whether you saw any restrictions or prompts

Validation check:
- The target internal user can access the file (as permitted).

Troubleshooting (common)
- If a policy was just changed by the trainer, allow time for it to apply.
- If unmanaged-device controls were changed, Microsoft notes it can take up to 24 hours and won’t impact already signed-in sessions.

---

## Wrap-up checkpoint
You should be able to:
- point to where OneDrive tenant settings live in SharePoint admin center
- explain the intent of each major settings area (Sharing, Sync, Storage, Retention, Access control)
- explain the deleted-user retention + 93-day deleted-state restore concept
- demonstrate safe internal sharing using OneDrive

## References (official)
- Manage OneDrive settings: https://learn.microsoft.com/en-us/sharepoint/onedrive-overview#manage-onedrive
- Sync guidance and hiding sync button: https://learn.microsoft.com/en-us/sharepoint/sharepoint-sync
- Default storage limit setting: https://learn.microsoft.com/en-us/sharepoint/set-default-storage-space
- Deleted user retention setting: https://learn.microsoft.com/en-us/sharepoint/set-retention
- Retention/deletion lifecycle: https://learn.microsoft.com/en-us/sharepoint/retention-and-deletion
- Restore deleted OneDrive: https://learn.microsoft.com/en-us/sharepoint/restore-deleted-onedrive
- Unmanaged devices access control: https://learn.microsoft.com/en-us/sharepoint/control-access-from-unmanaged-devices
