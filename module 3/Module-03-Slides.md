# Module 3 Slides — Working with Site Collections (Modern Sites)

> Format: slide title + bullets + speaker notes

## Slide 1 — Module goals
Bullets:
- Create modern sites from SharePoint admin center
- Manage site details (owner/admins, settings)
- Day-to-day site admin operations (membership, access requests, recycle bin)
- Monthly/regular access review (site admin + SharePoint admin)
- Understand storage controls (tenant vs site)
- Practice delete/restore safely
- Intro to PowerShell for site inspection

Speaker notes:
- Reinforce the shared-tenant model: participant changes only in `NW-Pxx-...` sites.

## Slide 2 — Scenario continuity and shared-tenant rules
Bullets:
- Scenario: Project Northwind Intranet Modernization
- One tenant shared by trainer + 10 admin participants
- Use Participant ID P01–P10 and `NW-Pxx-...` naming
- Persistent practice site: `NW-Pxx-ProjectSite`

Speaker notes:
- Module 3 is the foundation for later labs.

## Slide 3 — “Site collections” in SharePoint Online (modern view)
Bullets:
- Admin center manages sites as top-level units
- Site has URL + owner/admins + settings
- Many governance decisions are site-scoped

Speaker notes:
- This is why we isolate work per participant.

## Slide 4 — Create a site (admin center workflow)
Bullets:
- SharePoint admin center > Sites > Active sites
- Create > choose site type
- Set site name, owner, language (as prompted)

Speaker notes:
- UI can vary; teach learners how to find the “Create” entry point.

## Slide 5 — Site details panel: what admins look for
Bullets:
- General: URL, storage usage, primary owner
- Membership/admins: who can administer
- Activity signals (if available in your tenant)
- Settings: sharing settings entry points

Speaker notes:
- Learners should always confirm they’re working on the correct site.

## Slide 6 — Storage controls: tenant vs site
Bullets:
- Tenant may be Automatic (pooled) or Manual
- Manual enables per-site storage limits + notifications
- Tenant mode changes are Trainer-only

Speaker notes:
- Participants focus on observing and documenting in shared tenant.

## Slide 7 — Delete and restore: safe lifecycle drill
Bullets:
- Use test site: `NW-Pxx-RestoreTest`
- Create → Delete (Active sites) → Restore (Deleted sites)
- Don’t delete the root site or other people’s sites

Speaker notes:
- Emphasize “practice on a disposable site.”

## Slide 8 — PowerShell intro: connect + inspect
Bullets:
- Connect-SPOService (admin URL)
- Get-SPOSite (retrieve properties)
- Target only your own `NW-Pxx-...` sites

Speaker notes:
- Use PowerShell to learn repeatable admin tasks.
- Optional advanced demo (trainer-only): Set-SPOUser to assign a temporary backup site collection admin, then immediately remove it as part of cleanup.

## Slide 9 — Lab 3 preview
Bullets:
- Create `NW-Pxx-ProjectSite`
- Capture owner/admins, membership notes, activity signal, and storage observations
- Check access request settings (observe-first)
- Practice a lightweight monthly access review (document findings)
- Recycle Bin restore drill (restore a deleted file)
- Delete/restore `NW-Pxx-RestoreTest`
- Run `Connect-SPOService` + `Get-SPOSite`

Speaker notes:
- Focus on evidence capture (worksheet).

## Slide 10 — Knowledge check
Bullets:
- Why isolate work to `NW-Pxx-...` sites?
- When can you set per-site storage limits?
- Why use a dedicated restore test site?
- What must happen before `Get-SPOSite`?

Speaker notes:
- Keep answers short; prioritize scope and safety.
