# Module 4 — Permissions and Collaboration Model (Slides)

> Format: slide title → bullets → speaker notes.

---

## Slide 1 — Module 4: Permissions and Collaboration Model
- Inheritance vs unique permissions
- SharePoint groups vs Microsoft 365 group-connected permissions
- Sharing links: internal collaboration vs governance
- Lab: Design a permission model (Northwind Contracts)

Speaker notes:
- Reinforce shared-tenant rule: no tenant-wide changes by participants.
- Remind learners: permissions are a “daily admin” topic because most incidents are access-related.

---

## Slide 2 — Outcomes (what you will be able to do)
- Choose the right boundary: site vs library vs folder
- Implement a minimal unique-permissions design
- Validate access with “Check Permissions” and “Manage access”

Speaker notes:
- Set expectations: we will not build a complex, item-by-item permission maze.

---

## Slide 3 — Permission basics (fast refresher)
- Principal: user or group
- Permission level: Full Control / Edit / Read
- Inheritance: parent → child
- Unique permissions: break inheritance (new permission scope)

Speaker notes:
- Use a simple diagram in words: Site → Library → Folder → File.

---

## Slide 4 — Why “unique permissions” is a risk
- Harder to audit
- Harder to troubleshoot
- Easier to misconfigure
- Scale guidance: keep unique scopes under control

Speaker notes:
- Reference Microsoft guidance on permission scopes and best practice.
- Teaching point: boundary choice is the core skill.

---

## Slide 5 — Site type matters
- Communication sites: SharePoint groups are the main model
- Team sites: often Microsoft 365 group-connected
- Teams-connected sites: manage permissions in Teams (especially channel sites)

Speaker notes:
- Emphasize: group-connected doesn’t mean you can’t use SharePoint groups, but governance needs clarity.

---

## Slide 6 — Sharing links vs permissions
- Permissions are durable access design
- Sharing links are collaboration actions
- Always ask: “Is access coming from a group or a link?”

Speaker notes:
- In incidents, a “mystery access” often comes from a link.

---

## Slide 7 — Northwind scenario: Contracts workflow
- Drafts: editors collaborate
- Final: broad read, restricted edit
- Goal: isolate sensitive content in a dedicated library

Speaker notes:
- Map to roles: Owners / Editors / Readers.

---

## Slide 8 — Lab preview (shared-tenant safe)
- Work only inside `NW-Pxx-ProjectSite`
- Create `NW-Pxx-Contracts` library
- Break inheritance at library
- Break inheritance at one folder only
- Validate using “Check Permissions”

Speaker notes:
- Stress “one folder only” drill: demonstrates the concept while keeping scope count low.

---

## Slide 9 — Troubleshooting mindset
- Start with: where is the boundary?
- Check: group membership and direct permissions
- Check: sharing links (“Manage access”)
- Use “Check Permissions” to confirm effective access

Speaker notes:
- Encourage learners to take screenshots/notes of where they checked.

---

## Slide 10 — Wrap-up and validation
- You can explain inheritance + unique scopes
- You can implement library/folder permissions safely
- You can confirm access source (groups vs links)

Speaker notes:
- Point learners to references and remind them to remove ad-hoc links.

---

## References
- Manage Permission Scopes: https://learn.microsoft.com/en-us/sharepoint/manage-permission-scope
- Sharing & permissions (modern): https://learn.microsoft.com/en-us/sharepoint/modern-experience-sharing-permissions
- Shareable links: https://learn.microsoft.com/en-us/sharepoint/shareable-links-anyone-specific-people-organization
- Manage sharing settings: https://learn.microsoft.com/en-us/sharepoint/turn-external-sharing-on-or-off
- External sharing overview: https://learn.microsoft.com/en-us/sharepoint/external-sharing-overview
- Troubleshooting Access Denied (includes Check Permissions steps): https://learn.microsoft.com/en-us/troubleshoot/sharepoint/administration/access-denied-or-need-permission-error-sharepoint-online-or-onedrive-for-business
