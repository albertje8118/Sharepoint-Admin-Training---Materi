# Module 9 — OneDrive for Business Administration (Student Guide)

## Learning objectives
By the end of this module, you will be able to:
1. Explain how OneDrive for Business relates to SharePoint Online.
2. Locate the OneDrive admin settings in the SharePoint admin center.
3. Describe key policy areas: sharing, sync, storage, retention, access control.
4. Explain the OneDrive lifecycle for deleted users and basic restore options.

---

## 1) OneDrive architecture (admin mental model)
OneDrive for Business is built on SharePoint Online.
- Each user’s OneDrive is a **personal site** (a SharePoint site collection) hosted under your tenant’s `-my.sharepoint.com` domain.
- This is why many “OneDrive settings” appear in the **SharePoint admin center**, and why SharePoint permissions and compliance features (retention, eDiscovery, etc.) can apply.

Operational implications:
- OneDrive is “personal” from a user perspective, but many settings are **tenant-wide policies**.
- In a shared training tenant, policy changes can impact everyone immediately or within hours.

---

## 2) Where OneDrive settings live
Microsoft’s OneDrive admin guidance points to the SharePoint admin center as the place to manage key OneDrive settings.
Common categories include:
- **Sharing** (organization-wide)
- **Sync** (tenant sync controls)
- **Storage limit** (default quota behavior)
- **Retention** (deleted-user OneDrive retention)
- **Access control** (unmanaged devices / network location)
- **Notifications**

Reference:
- OneDrive overview (Manage OneDrive section): https://learn.microsoft.com/en-us/sharepoint/onedrive-overview#manage-onedrive

---

## 3) Sharing controls (admin view)
Sharing controls are typically set at the organization level and influence both SharePoint and OneDrive behavior.
Admin checklist:
- Confirm external sharing level and default sharing link behaviors.
- For training: keep sharing **internal-only** unless the trainer explicitly authorizes external tests.

Concept note (from earlier modules):
- Sharing is a blend of tenant-level policy + site-level configuration + user action.

---

## 4) Sync controls (and shortcuts vs sync)
Users have multiple ways to bring SharePoint/OneDrive content into File Explorer/Finder.
Microsoft guidance recommends **OneDrive shortcuts** (Add shortcut to OneDrive) as the more versatile option versus library Sync, because shortcuts are associated with the user and follow them across devices.

Admin-related takeaways:
- You can hide/remove the **Sync** button to prevent new syncs from being started (existing syncs are not affected).

References:
- Sync in SharePoint and OneDrive: https://learn.microsoft.com/en-us/sharepoint/sharepoint-sync
- Recommended sync configuration (shortcuts): https://learn.microsoft.com/en-us/sharepoint/ideal-state-configuration#shortcuts-to-shared-folders

---

## 5) Storage policies
Storage management usually has layers:
- license-based maximums
- tenant defaults
- per-user overrides

Admin takeaways:
- Changing the default can affect many users.
- If you reduce storage below current usage, the OneDrive may become read-only (per Microsoft warning).

Reference:
- Set default storage space for OneDrive users: https://learn.microsoft.com/en-us/sharepoint/set-default-storage-space

---

## 6) Device access controls (unmanaged devices)
SharePoint admin center includes controls to block or limit access to SharePoint and OneDrive from **unmanaged devices**.

Key behavior:
- **Block access**: users cannot access content from an unmanaged device.
- **Allow limited, web-only access**: browser-only access with restrictions (no download/print/sync); optional controls exist for editing and file types.

Operational notes:
- These controls rely on Microsoft Entra Conditional Access.
- Changes can take up to **24 hours** to take effect, and won’t impact users already signed in.
- Microsoft recommends also blocking access from apps that don’t use modern authentication to prevent bypass.

Reference:
- Control access from unmanaged devices (SharePoint + OneDrive): https://learn.microsoft.com/en-us/sharepoint/control-access-from-unmanaged-devices

---

## 7) User lifecycle: retention and restore
When a user leaves, OneDrive retention and cleanup behavior matters.

Key concepts from Microsoft guidance:
- Deleted user OneDrive content is retained for a configured number of days (default 30; configurable 30–3650).
- After the retention period, the OneDrive remains in a deleted state for **93 days** and can be restored by a SharePoint Administrator.

References:
- Set OneDrive retention for deleted users: https://learn.microsoft.com/en-us/sharepoint/set-retention
- OneDrive retention and deletion: https://learn.microsoft.com/en-us/sharepoint/retention-and-deletion
- Restore a deleted OneDrive: https://learn.microsoft.com/en-us/sharepoint/restore-deleted-onedrive

---

## References (official)
- OneDrive overview (Manage OneDrive): https://learn.microsoft.com/en-us/sharepoint/onedrive-overview
- Sync in SharePoint and OneDrive: https://learn.microsoft.com/en-us/sharepoint/sharepoint-sync
- Set default OneDrive storage: https://learn.microsoft.com/en-us/sharepoint/set-default-storage-space
- Set retention for deleted users: https://learn.microsoft.com/en-us/sharepoint/set-retention
- Control access from unmanaged devices: https://learn.microsoft.com/en-us/sharepoint/control-access-from-unmanaged-devices
- Retention and deletion lifecycle: https://learn.microsoft.com/en-us/sharepoint/retention-and-deletion
- Restore deleted OneDrive: https://learn.microsoft.com/en-us/sharepoint/restore-deleted-onedrive
