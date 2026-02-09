# Lab 12 (Optional) — Workflow Automation with Power Automate and Power Apps

## Goal
Create a simple, training-safe workflow that:
- starts from your `NW-Pxx-AppRequests` list,
- routes an approval to a designated approver,
- writes the approval outcome back to the same list item,
- and improves the list form experience using Power Apps.

## Estimated time
60–90 minutes

## Prerequisites
### Roles and access
- Participants
  - Site Owner (Full Control) on your `NW-Pxx-ProjectSite`.
  - Permission to create/edit flows in Power Automate (maker capability depends on tenant settings).
  - Permission to customize list forms with Power Apps.
- Trainer
  - Power Platform Administrator (or equivalent) for governance demonstrations (environments, DLP policies).

### Required continuity artifact
This lab assumes you already have (from Module 7):
- List: `NW-Pxx-AppRequests` (inside your `NW-Pxx-ProjectSite`)

If you don’t have the list, complete Lab 07 first (recommended).

### Lab worksheets (from your participant pack)
Use these templates in your generated pack:
- `TXT-Templates/M12-Flow-Design-Worksheet.txt`
- `TXT-Templates/M12-PowerApps-Form-Customization-Checklist.txt`
- `TXT-Templates/M12-Governance-DLP-Notes.txt`

---

## Shared-tenant safety rules (critical)
- Create flows/apps only for your own `NW-Pxx-...` list/site.
- Keep the approver set to **yourself** (or trainer-designated account) to avoid spamming other participants.
- Prefer **When an item is created** for workflows that update the same item.
- If something requires tenant-wide changes (environments/DLP), treat as Trainer-only.

---

## Exercise 1 — Prepare the request item (10–15 minutes)
1. Open your list `NW-Pxx-AppRequests`.
2. Create a new item:
   - Title: `NW-Pxx - Approval test request`
   - RequestType: choose any value
   - Status: `Submitted`
   - Owner: set to **yourself** (this will be the approver)
   - Notes: (leave blank)

Validation check:
- The item exists and Status is `Submitted`.

---

## Exercise 2 — Build an approval flow (Power Automate) (25–40 minutes)
You will build an **Automated cloud flow** that triggers on item creation.

Reference (official):
- SharePoint connector triggers/actions: https://learn.microsoft.com/en-us/sharepoint/dev/business-apps/power-automate/sharepoint-connector-actions-triggers
- Modern approvals: https://learn.microsoft.com/en-us/power-automate/modern-approvals
- Get started with approvals: https://learn.microsoft.com/en-us/power-automate/get-started-approvals

### 2A. Create the flow
1. Go to Power Automate: https://make.powerautomate.com
2. Create a new **Automated cloud flow**.
3. Choose trigger: **SharePoint — When an item is created**.
4. Configure trigger:
   - Site Address: your `NW-Pxx-ProjectSite`
   - List Name: `NW-Pxx-AppRequests`
5. Save the flow once (so you can return to it).

### 2B. Add the approval step
1. Add action: **Approvals — Start and wait for an approval**.
2. Configure:
   - Approval type: **Approve/Reject — First to respond**
   - Title: `NW-Pxx App Request Approval: <Title>`
   - Assigned to: the item **Owner** email (or set it to your own email)
   - Details: include key fields (Title, RequestType, DueDate if used)
   - Item link: link back to the SharePoint list item (optional but recommended)

Record your choices in `M12-Flow-Design-Worksheet.txt`.

### 2C. Update the item based on outcome
1. Add a **Condition** that checks the approval **Outcome**.
2. If outcome is **Approve**:
   - Action: **SharePoint — Update item**
   - Set Status = `Approved`
   - Optionally append approval comments to Notes
3. If outcome is **Reject**:
   - Action: **SharePoint — Update item**
   - Set Status = `Rejected`
   - Optionally append rejection comments to Notes

Validation check:
- Flow saves successfully.
- Flow has: trigger → Start and wait for approval → condition → update item.

---

## Exercise 3 — Test the flow (10–15 minutes)
1. Create a **new** item in `NW-Pxx-AppRequests` (Status = Submitted, Owner = you).
2. Wait for the approval notification (email or approvals center).
3. Approve or reject.
4. Return to the list and verify the item updates:
   - Status becomes `Approved` or `Rejected`
   - Notes updated (if you configured comments)

Validation check:
- The list item reflects the approval outcome.

Troubleshooting (common):
- If the flow triggers repeatedly, confirm you used **When an item is created** (not created/modified).
- If approvals don’t arrive, confirm Assigned to is a valid user email.

Reference (official):
- Customize triggers with conditions: https://learn.microsoft.com/en-us/power-automate/customize-triggers

---

## Exercise 4 — Customize the list form (Power Apps) (15–25 minutes)
Goal: improve the request intake experience without deploying code.

Reference (official):
- Customize a form for a SharePoint list: https://learn.microsoft.com/en-us/sharepoint/dev/business-apps/power-apps/get-started/create-your-first-custom-form

1. Open `NW-Pxx-AppRequests`.
2. From the command bar, select **Integrate** → **Power Apps** → **Customize forms**.
3. Make a minimal change (choose one):
   - Re-order fields so Status and Owner are near the top, OR
   - Add a label at the top: “Northwind App Request Intake (Training)”, OR
   - Hide the Notes field unless Status = Rejected (advanced).
4. Save and **Publish to SharePoint**.
5. Validate by opening a list item in SharePoint and confirming the form change is visible.

Record completion in `M12-PowerApps-Form-Customization-Checklist.txt`.

Validation check:
- Your customization is visible when you open/edit an item.

---

## Exercise 5 — Governance notes (Trainer-led / read-only) (optional) (10–15 minutes)
Goal: connect workflow automation to admin governance.

1. In the Power Platform admin center (trainer-led), review:
   - environments (where apps/flows live)
   - data policies (DLP) and connector grouping
2. Discuss what happens when a connector is blocked or connector combinations are restricted.

References (official):
- Manage data policies (DLP): https://learn.microsoft.com/en-us/power-platform/admin/prevent-data-loss
- DLP strategy guidance: https://learn.microsoft.com/en-us/power-platform/guidance/adoption/dlp-strategy

Record key takeaways in `M12-Governance-DLP-Notes.txt`.

---

## Cleanup (optional)
- Turn off your training flow after the lab if the class is not continuing with workflow exercises.
