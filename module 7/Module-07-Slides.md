# Module 7 — Apps and Customization in SharePoint Online (Slides)

## Slide 1 — Title
- Module 7: Apps and customization in SharePoint Online
- Scenario: Project Northwind Intranet Modernization

## Slide 2 — Why admins care
- Customization can improve productivity
- But it must be:
  - supportable,
  - secure,
  - and governed.

## Slide 3 — Customization spectrum
- Out-of-box configuration
- Declarative customization (JSON formatting)
- SharePoint Framework (SPFx) solutions

## Slide 4 — Declarative customization: what it is
- JSON-based formatting for lists/libraries
- Changes rendering, not data
- Fast, low-risk compared to custom code

## Slide 5 — Column formatting
- Applies to a single field
- Where to find it:
  - Column settings → Format this column
- Preview → Save

## Slide 6 — View formatting
- Applies to the current view layout
- Where to find it:
  - View dropdown → Format current view
- Supports List/Compact List, Gallery, Board

## Slide 7 — When formatting is “enough”
- Highlight status
- Improve scannability
- Add small actions (depending on design)

## Slide 8 — SPFx (SharePoint Framework) overview
- Modern extensibility model
- Used for SharePoint + Teams + Viva Connections
- Packaged as `.sppkg`

## Slide 9 — Apps management surface
- SharePoint admin center → More features → Apps
- Admin can upload/enable solutions
- Governance: test first, then roll out

## Slide 10 — 2026 alignment note (add-ins)
- Microsoft documentation states SharePoint add-ins are being retired for SharePoint in Microsoft 365
- Focus admin training on:
  - SPFx,
  - declarative customization,
  - and governance.

## Slide 11 — Permissions and governance (core admin message)
- Scope matters:
  - site-only vs tenant-wide
- Don’t surprise the tenant

## Slide 12 — API access (why it exists)
- SPFx/custom scripts can request Entra ID-secured API permissions
- Admins manage requests in SharePoint admin center → API access

## Slide 13 — Roles and approvals
- Role needed depends on API
- Third-party APIs may be approved by Application Administrator
- Graph/Microsoft APIs typically require Global admin

## Slide 14 — Lab overview
- Create `NW-Pxx-AppRequests`
- Apply column formatting (Status)
- Apply view formatting (row shading)
- Trainer-led tours: Apps + API access

## Slide 15 — End-of-module checkpoint
- You can build a “lightweight app experience” using lists + formatting
- You can explain how app deployment and API approval are governed

## Slide 16 — What’s next
- Module 8: Compliance and governance with Microsoft Purview
