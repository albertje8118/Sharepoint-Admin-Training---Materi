# Module 7 — Apps and Customization in SharePoint Online (Student Guide)

## Learning objectives
By the end of this module, you will be able to:
1. Explain modern customization models in SharePoint Online.
2. Customize a list safely using JSON-based column and view formatting.
3. Describe how SPFx solutions are deployed and governed.
4. Describe what the SharePoint admin center **API access** page is for and why it matters.

---

## 1) Customization models in SharePoint Online (admin view)

As an administrator, you often need to answer two questions:
1) *How can we customize this experience?*  
2) *How do we keep it secure and supportable?*

A practical way to think about customization is a spectrum:

### 1.1 Out-of-box configuration (lowest risk)
Examples:
- list/library settings
- pages and web parts configuration
- permissions, sharing settings, and governance controls

Admin take-away:
- Prefer out-of-box features first. They are easiest to support.

### 1.2 Declarative customization (JSON formatting)
SharePoint supports declarative JSON formatting to change how lists and libraries are displayed:
- **Column formatting** customizes how a field is *rendered in a view*.
- **View formatting** customizes how items/rows/cards are *rendered in the current view*.

Key property:
- Formatting does **not** change the underlying list data; it changes how users see it.

Admin take-away:
- JSON formatting is powerful for “apps-like” experiences without code or app catalogs.

### 1.3 SharePoint Framework (SPFx) (highest flexibility)
SPFx is the modern extensibility model for SharePoint Online and is also used to extend Microsoft Teams and Viva Connections.

Admin take-aways:
- SPFx is the recommended extensibility model in Microsoft 365.
- SPFx packages are commonly deployed as `.sppkg` solutions.
- Deployment scope and permissions must be governed.

Important 2026 note:
- Microsoft documentation states **SharePoint add-ins** are being retired for SharePoint in Microsoft 365 (timeline and details vary by feature area). In 2026-aligned admin training, treat add-ins as legacy and focus on SPFx + declarative customization.

---

## 2) Column formatting (what it is and how admins govern it)

Column formatting lets you render a field value with conditional styles/icons.

What you should remember:
- You open it from a column menu: **Column settings** → **Format this column**.
- You paste JSON, **Preview**, then **Save**.
- It applies to everyone who uses that view/column.

Admin governance guidance:
- Use consistent naming for custom views/columns.
- Keep JSON formatting readable and version-controlled (store snippets in a central library or repo).

---

## 3) View formatting (when you need more than a single column)

View formatting lets you change how rows/cards are rendered in a view:
- You open it from the view menu: **Format current view**.
- It supports different layouts (List/Compact List, Gallery, Board).

Admin governance guidance:
- Prefer minimal formatting that improves scannability.
- Validate accessibility and readability for your organization.

---

## 4) Managing apps: what admins actually manage

In SharePoint Online, admins can manage apps/solutions, including SharePoint Framework packages, using the SharePoint admin center.

From Microsoft’s admin guidance:
- Admins can acquire solutions from the SharePoint Store or distribute custom apps.
- For custom apps, admins upload solutions and can **Enable app** and optionally add it broadly.

Operational guidance (shared tenant training):
- Treat tenant-wide app deployment as **Trainer-only**.
- In real organizations, test first, then deploy gradually.

---

## 5) API access (permissions governance for solutions)

SPFx solutions and custom scripts can request permissions to Microsoft Entra ID-secured APIs.
Admins manage these requests via **API access** in the SharePoint admin center.

Key points from Microsoft guidance:
- The required admin role depends on the API.
  - For many third-party APIs, the **Application Administrator** role can be sufficient.
  - For Microsoft Graph or other Microsoft APIs, **Global Administrator** approval is required.
- Approvals affect the tenant and must be reviewed carefully.

Admin take-away:
- “API access” is a governance and security surface, not a routine day-to-day click-through.

---

## References (official)
- SPFx overview: https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview
- Manage apps (Apps site): https://learn.microsoft.com/en-us/sharepoint/use-app-catalog
- Manage API access (Entra ID-secured APIs): https://learn.microsoft.com/en-us/sharepoint/api-access
- Column formatting: https://learn.microsoft.com/en-us/sharepoint/dev/declarative-customization/column-formatting
- View formatting: https://learn.microsoft.com/en-us/sharepoint/dev/declarative-customization/view-formatting
