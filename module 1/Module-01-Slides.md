# Module 1 Slides — Introduction to Microsoft 365 and SharePoint Online (Admin)

> Format: slide title + bullets + speaker notes

## Slide 1 — Course context (where we are)
Bullets:
- 3-day “Modern SharePoint Online for Administrators”
- Day 1 focus: tenant foundations and site management
- Module 1: admin mental model + portals + operational checks

Speaker notes:
- Set expectations: we’ll avoid memorizing numbers and instead learn where to verify current service behavior in official docs and portals.

## Slide 2 — Lab scenario and shared-tenant rules
Bullets:
- Scenario: Project Northwind Intranet Modernization
- One tenant shared by trainer + 10 admin participants
- Use Participant ID: P01–P10
- Avoid tenant-wide changes unless Trainer-only

Speaker notes:
- Explain why: shared tenant means we must prevent collisions and disruptive changes.
- Set expectation: most “hands-on” changes will happen in participant-isolated practice sites later.

## Slide 3 — Microsoft 365: admin mental model
Bullets:
- Identity: Microsoft Entra ID
- Workloads: SharePoint, OneDrive, Teams, Exchange, Purview, Search
- Admin surfaces: M365 admin center + workload admin centers
- Automation: Graph + PowerShell (later modules)

Speaker notes:
- Emphasize that many “SharePoint issues” are identity/policy issues.

## Slide 4 — Where SharePoint Online fits
Bullets:
- Sites for collaboration and publishing
- Document libraries for controlled content
- Powers Teams-connected collaboration sites
- Underpins OneDrive storage model

Speaker notes:
- Highlight: SharePoint is the content layer across Microsoft 365.

## Slide 5 — Common SharePoint admin responsibilities
Bullets:
- Tenant policies (sharing/access)
- Site lifecycle & ownership
- Storage oversight & reporting
- Governance alignment (Purview later)
- Service awareness (health + message center)

Speaker notes:
- Tie responsibilities to what learners will do in labs.

## Slide 6 — SharePoint Online vs SharePoint Server (conceptual)
Bullets:
- Cloud-operated vs self-operated infrastructure
- Continuous updates vs scheduled upgrades
- Integration-first (M365) vs on-prem integration patterns
- Governance and policy center-of-gravity shifts

Speaker notes:
- Keep it conceptual; avoid deep feature comparisons.

## Slide 7 — Microsoft 365 admin center: what to use it for
Bullets:
- Role and tenant baseline checks
- Service health triage
- Message center change awareness
- Licensing checks when troubleshooting

Speaker notes:
- “Service health first” habit reduces wasted time.

## Slide 8 — SharePoint admin center: what to use it for
Bullets:
- Sites management at scale
- Tenant-wide SharePoint/OneDrive policies (where available)
- Storage & usage views
- Governance levers (within SharePoint scope)

Speaker notes:
- Acknowledge UI differs by tenant; teach finding categories rather than memorizing clicks.

## Slide 9 — Limits, quotas, and boundaries: how to approach
Bullets:
- Limits change; documentation updates
- Don’t memorize numbers—verify current values
- Identify which limit impacts your scenario
- Document assumptions and sources

Speaker notes:
- Set up the pattern: official docs as the source of truth.

## Slide 10 — Lab 1 preview
Bullets:
- Access admin centers
- Review SharePoint tenant settings areas
- Check Service health
- Review Message center

Speaker notes:
- Explain what learners should capture in validation checkpoints.

## Slide 11 — Knowledge check (discussion)
Bullets:
- Entra ID role in access control
- Why Message center matters
- 2 differences: Online vs Server
- Safe method for service limits

Speaker notes:
- Encourage short answers; use this to gauge baseline.
