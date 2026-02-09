# Lab 07 — Managing Apps and Customization (SharePoint Online)

**Module:** 7 — Apps and Customization in SharePoint Online  \
**Estimated time:** 60–90 minutes  \
**Lab type:** Individual + optional Trainer-led (shared-tenant safe)

## Lab goal
Practice a modern, admin-friendly customization approach that doesn’t require app deployment:
- create a small “app requests” list,
- apply **column formatting** and **view formatting** using JSON,
- review the admin governance surfaces for **Apps** and **API access** (trainer-led).

## Prerequisites
- You have a participant ID `Pxx` (P01–P10).
- You have a site: `NW-Pxx-ProjectSite`.
- You can create lists and views on `NW-Pxx-ProjectSite` (typically **Site Owner / Full Control**).

### Permissions and access (what you need)
- Exercises 1–3 (list + formatting):
  - At minimum, **Full Control** (Site Owner) on `NW-Pxx-ProjectSite`.
- Exercise 4 (Apps management surface):
  - Typically requires **SharePoint Administrator** (trainer-led in shared tenant).
- Exercise 5 (API access):
  - Depends on the API:
    - Third-party API approvals can be possible with **Application Administrator**.
    - Microsoft Graph / Microsoft APIs require **Global Administrator**.

## Shared-tenant safety rules (critical)
- Only create objects prefixed with your own `NW-Pxx-...`.
- Do not upload or deploy tenant-wide app packages unless the trainer explicitly instructs you.
- Do not approve any API access requests in a shared training tenant.

---

## Exercise 1 — Create the Northwind app requests list (15–25 minutes)

### 1A. Create a list
1. Go to your site `NW-Pxx-ProjectSite`.
2. Create a new list named: `NW-Pxx-AppRequests`.

### 1B. Add columns
Add the following columns:
- `RequestType` (Choice)
  - Suggested choices: `SPFx`, `Power App`, `List formatting`, `Other`
- `Status` (Choice)
  - Choices: `Submitted`, `In review`, `Approved`, `Rejected`
  - Default: `Submitted`
- `DueDate` (Date and time) — Date only
- `Owner` (Person)
- `Notes` (Multiple lines of text)

### 1C. Add 3 sample items
Create at least three items using your own participant prefix in the Title:
- Title: `NW-Pxx - Contracts dashboard request`
- Title: `NW-Pxx - New supplier onboarding form`
- Title: `NW-Pxx - Metadata cleanup helper`

Validation check:
- The list exists and has all columns.
- You can see all three items.

---

## Exercise 2 — Apply column formatting to the Status column (15–25 minutes)

Column formatting changes how a single field is rendered in the view.

### What the JSON is (and what it is not)
- The JSON in this exercise is **SharePoint list formatting** (declarative customization).
- It changes **how the column is displayed** (icons, colors, layout) but does **not** change the underlying list data.
- SharePoint stores column/view formatting as JSON. Some tenants expose a **visual** formatter that generates JSON for you.

### 2A. Open the column formatting panel
1. In `NW-Pxx-AppRequests`, open the drop-down menu for the `Status` column.
2. Select **Column settings** → **Format this column**.

### 2A (Alternative) — Create formatting without typing JSON (no code)
If your formatting pane includes a visual/rules-based editor, you can create a similar result without pasting JSON.

1. In the **Format this column** pane, look for options such as **Rules**, **Conditional formatting**, or **Style**.
2. Create rules for the `Status` choices:
   - `Approved` → choose a “success/green” style.
   - `In review` → choose a “warning/yellow” style.
   - `Submitted` → choose a “neutral/low” style.
   - `Rejected` → choose a “blocked/red” style.
3. If the pane offers an **icon** picker per rule, choose icons that match each status.
4. Select **Preview**, then **Save**.

Note:
- If you do not see a rules/visual option, your tenant/UI may require using **Advanced mode** (JSON) for conditional + icon formatting.

### 2B. Paste JSON and save
1. Paste the JSON below.
2. Select **Preview**.
3. Select **Save**.

### What this JSON does (quick explanation)
- It renders the cell as a container (`elmType: div`).
- It uses `@currentField` (the Status value) to:
  - assign a severity-style CSS class (good/warning/low/blocked), and
  - pick an icon name (check mark, warning, etc.).
- It then prints the Status text next to the icon.

Use this JSON (training):
```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json",
  "elmType": "div",
  "attributes": {
    "class": "=if(@currentField == 'Approved', 'sp-field-severity--good', if(@currentField == 'In review', 'sp-field-severity--warning', if(@currentField == 'Submitted', 'sp-field-severity--low', 'sp-field-severity--blocked'))) + ' ms-fontColor-neutralSecondary'"
  },
  "children": [
    {
      "elmType": "span",
      "style": { "display": "inline-block", "padding": "0 4px" },
      "attributes": {
        "iconName": "=if(@currentField == 'Approved', 'CheckMark', if(@currentField == 'In review', 'Error', if(@currentField == 'Submitted', 'Forward', 'ErrorBadge')))"
      }
    },
    {
      "elmType": "span",
      "txtContent": "@currentField"
    }
  ]
}
```

Validation check:
- Status now displays with a visual indicator (style/icon).
- Changing an item’s Status changes the formatting.

---

## Exercise 3 — Apply view formatting (10–20 minutes)

View formatting changes how items/rows are rendered in the current view.

### What this JSON does (quick explanation)
- It applies an extra CSS class to every other row using `@rowIndex`.
- Result: alternating (zebra) row shading to make the view easier to scan.

### 3A. Open the view formatting panel
1. In the list, open the view dropdown.
2. Select **Format current view**.
3. In the formatting pane, choose layout **List** (or **Compact list** if that’s what you use).

### 3A (Alternative) — Try to create view styling without typing JSON (no code)
Depending on your tenant/UI, the view formatting pane may provide visual layout/styling toggles.

1. In **Format current view**, look for options such as **Row styling**, **Alternating rows**, or a preset that changes row background.
2. If available, enable alternating rows (or choose a preset with alternating shading).
3. **Save**.

Note:
- If you don’t see an “alternating rows” option, zebra striping is often only available through **view formatting JSON** in Advanced mode.

### 3B. Paste JSON and save
Paste the JSON below and save:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/sp/view-formatting.schema.json",
  "additionalRowClass": "=if(@rowIndex % 2 == 0,'ms-bgColor-themeLighter ms-bgColor-themeLight--hover','')"
}
```

Validation check:
- Rows display with alternating shading, improving scanability.

---

## Exercise 4 — Optional (Trainer-led): Apps management surface (10–15 minutes)

Goal: identify where admins manage app packages and what “Enable app” means.

1. In the SharePoint admin center, go to **More features**.
2. Under **Apps**, select **Open**.
3. Observe:
   - Where admins upload custom apps (SPFx packages are `.sppkg`).
   - The option to **Enable this app** and (optionally) add it broadly.

2026 alignment note:
- Microsoft documentation states SharePoint add-ins are being retired for SharePoint in Microsoft 365; treat add-ins as legacy and focus on modern approaches.

Validation check:
- You can explain the difference between:
  - uploading/enabling an app (admin), and
  - adding/using an app on a site (site owner).

---

## Exercise 5 — Optional (Trainer-led): API access governance (10–15 minutes)

Goal: identify where admins approve API permission requests and why this requires governance.

1. In the SharePoint admin center, open **API access**.
2. Review the page sections:
   - pending requests,
   - approved requests.
3. Discuss:
   - which admin role is needed depends on the API,
   - Graph/Microsoft API approvals typically require Global admin,
   - approvals should be reviewed; do not approve requests in a shared tenant.

Validation check:
- You can explain what “API access” is used for and why it’s sensitive.

---

## End state (leave in place for later modules)
- List: `NW-Pxx-AppRequests` with sample items
- Status column formatted
- One formatted view

If the trainer requests cleanup:
- Remove view formatting (reset to default)
- Optionally delete `NW-Pxx-AppRequests`

---

## Troubleshooting

1) **I don’t see “Format this column” or “Format current view”.**
- Confirm you’re using a modern list experience.
- Confirm you have permission to create/manage views.

2) **My JSON won’t save.**
- Ensure the JSON is valid and includes the `$schema` line.
- Start with a known-good example and edit incrementally.

3) **I can’t access Apps / API access in the admin center.**
- Those are admin surfaces; in this course, treat them as trainer-led.

---

## References (official)
- Manage apps (Apps site): https://learn.microsoft.com/en-us/sharepoint/use-app-catalog
- Manage API access: https://learn.microsoft.com/en-us/sharepoint/api-access
- Column formatting: https://learn.microsoft.com/en-us/sharepoint/dev/declarative-customization/column-formatting
- View formatting: https://learn.microsoft.com/en-us/sharepoint/dev/declarative-customization/view-formatting
- SPFx overview: https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview
