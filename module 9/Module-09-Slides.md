# Module 9 — OneDrive for Business Administration (Slides Outline)

## Slide 1 — Module title + outcomes
- OneDrive for Business Administration (SharePoint admin perspective)
- Outcomes:
  - Locate OneDrive tenant settings
  - Explain sharing/sync/storage/retention/access-control controls
  - Understand deleted-user lifecycle and restore options

## Slide 2 — OneDrive architecture (why it’s “SharePoint admin work”)
- OneDrive = personal site hosted on SharePoint Online (`-my.sharepoint.com`)
- Policies are often tenant-wide
- Shared-tenant training rule: observe first; trainer changes only

## Slide 3 — Where settings live
- SharePoint admin center:
  - Sharing
  - Settings → Sync / Storage limit / Retention / Notifications
  - Access control

Reference:
- https://learn.microsoft.com/en-us/sharepoint/onedrive-overview#manage-onedrive

## Slide 4 — Sharing controls (policy vs user action)
- Tenant policy sets boundaries
- Users share within boundaries
- Training: internal sharing only unless trainer authorizes external test

## Slide 5 — Sync controls (shortcuts vs sync)
- Two user approaches:
  - Add shortcut to OneDrive (recommended)
  - Library Sync (device-based)
- Admin: can hide Sync button to prevent new sync starts

References:
- https://learn.microsoft.com/en-us/sharepoint/ideal-state-configuration#shortcuts-to-shared-folders
- https://learn.microsoft.com/en-us/sharepoint/sharepoint-sync

## Slide 6 — Storage policies
- Default storage limit vs per-user overrides
- Risk: lowering storage can cause read-only if user exceeds quota

Reference:
- https://learn.microsoft.com/en-us/sharepoint/set-default-storage-space

## Slide 7 — Device access controls (unmanaged devices)
- Allow full / Allow limited (web-only) / Block
- Uses Entra Conditional Access under the hood
- Operational reality:
  - can take up to 24 hours to apply
  - doesn’t affect already signed-in sessions

Reference:
- https://learn.microsoft.com/en-us/sharepoint/control-access-from-unmanaged-devices

## Slide 8 — User lifecycle (departures)
- Deleted user OneDrive retention (default 30, configurable 30–3650)
- After retention, deleted state for 93 days; restore via admin
- Purview retention/eDiscovery holds can override deletion timing

References:
- https://learn.microsoft.com/en-us/sharepoint/set-retention
- https://learn.microsoft.com/en-us/sharepoint/retention-and-deletion
- https://learn.microsoft.com/en-us/sharepoint/restore-deleted-onedrive

## Slide 9 — Lab briefing (what you will do)
- Observe and document tenant settings
- Perform a safe internal OneDrive sharing test
- Complete worksheets (M09 templates)

## Slide 10 — Wrap-up
- Review: where settings live + what each controls
- Q&A: common support scenarios (sync missing, blocked download, leaver retention)
