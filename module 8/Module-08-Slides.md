# Module 8 — Content Governance and Compliance with Microsoft Purview (Slides)

## Slide 1 — Title
- Module 8: Content governance and compliance with Microsoft Purview
- Scenario: Project Northwind Intranet Modernization

## Slide 2 — Why SharePoint admins care
- Compliance settings are often tenant-wide
- Mis-scoping policies can impact everyone
- Shared-tenant lab rule: trainer leads policy creation

## Slide 3 — Purview building blocks (Module 8)
- Retention (data lifecycle/records management)
- Sensitivity labels (classification + protection)
- eDiscovery (cases, searches, holds, export)
- DLP (simulation-first rollout)

## Slide 4 — Retention: policy vs label
- Retention policies: broad (locations)
- Retention labels: item-level (documents)
- Labels must be published via label policies before users can apply

## Slide 5 — Retention label publishing paths (official)
- Purview portal
  - Solutions → Records Management → Policies → Label policies
  - or Solutions → Data Lifecycle Management → Policies → Label policies

## Slide 6 — Retention timing (plan for it)
- Published retention labels to SharePoint/OneDrive typically appear within one day
- Allow up to seven days (replication)

## Slide 7 — Sensitivity labels: create + publish
- Create label: Solutions → Information Protection → Sensitivity labels
- Publish via publishing policy: Solutions → Information Protection → Publishing policies
- Publish to users/groups (not to sites)

## Slide 8 — Sensitivity timing (plan for it)
- Changes replicate across apps and services
- Best practice: pilot with a small group, validate, then expand

## Slide 9 — eDiscovery (modern)
- Classic eDiscovery experiences retired (per Microsoft guidance)
- Workflow
  - Case → Search → Review/Export
  - Optional: Hold to preserve content

## Slide 10 — eDiscovery: shared-tenant safety
- Default delivery: trainer-led demo
- Hands-on only if permissions are assigned
- Always scope searches/holds to NW-Pxx locations only

## Slide 11 — DLP deployment mindset
- Don’t “turn on and pray”
- Use simulation mode first
- Then simulation with policy tips
- Then enforcement

## Slide 12 — Lab overview
- Upload FAKE content to NW-Pxx site
- Apply sensitivity label (if published)
- Apply retention label (if published)
- Walk through eDiscovery case workflow (demo by default)

## Slide 13 — End-of-module checkpoint
- You can explain the difference between retention vs sensitivity
- You can describe replication delays and pilot-first rollout
- You can outline a safe eDiscovery workflow

## Slide 14 — What’s next
- Module 9: OneDrive administration + operational controls
