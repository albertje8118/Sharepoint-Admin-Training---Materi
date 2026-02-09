# Module 12 (Optional) — Workflow Automation with Power Automate and Power Apps (Student Guide)

## Learning objectives
By the end of this module, you will be able to:
1. Describe how SharePoint lists integrate with Power Automate triggers/actions.
2. Implement a basic approval pattern using **Start and wait for an approval**.
3. Explain common automation pitfalls (trigger scope, loop avoidance, ownership).
4. Customize a SharePoint list form using Power Apps.
5. Describe admin governance controls: environments and Data Loss Prevention (DLP) policies.

---

## 1) The admin mindset: automate safely, govern intentionally
For administrators, workflow automation is rarely “just a flow.” It is also:
- identity and ownership (who owns the flow/app, who can edit it)
- operational impact (notifications, runs, failures)
- data movement (connectors and data boundaries)

Shared-tenant training rule:
- Build only against your own `NW-Pxx-...` artifacts.
- Treat tenant-wide governance as Trainer-only unless explicitly assigned.

---

## 2) SharePoint + Power Automate: triggers and actions
Power Automate’s SharePoint connector includes common list triggers such as:
- **When an item is created**
- **When an item is created or modified**

Admin take-away:
- Prefer **When an item is created** for approval workflows that update the same item. It reduces the risk of update-trigger loops.

Reference (official):
- SharePoint connector triggers/actions: https://learn.microsoft.com/en-us/sharepoint/dev/business-apps/power-automate/sharepoint-connector-actions-triggers

---

## 3) Approvals: the “Start and wait” pattern
The most common admin-friendly approval pattern is:
1. Trigger from a SharePoint list item creation.
2. Create an approval request and wait for the result.
3. Update the SharePoint item with outcome and comments.

Key concept:
- **Start and wait for an approval** blocks the flow run until the approver responds.

References (official):
- Get started with approvals: https://learn.microsoft.com/en-us/power-automate/get-started-approvals
- Create and test an approval workflow (modern approvals): https://learn.microsoft.com/en-us/power-automate/modern-approvals
- Wait for approval tutorial: https://learn.microsoft.com/en-us/power-automate/wait-for-approvals

---

## 4) Common pitfalls (and how admins prevent them)

### 4.1 Infinite loops (created/modified triggers)
If you use **When an item is created or modified** and the flow updates that same item, it can re-trigger.

Mitigations:
- Prefer **When an item is created** for workflows that update the original item.
- If you must use created/modified, apply trigger conditions to prevent unnecessary runs.

Reference (official):
- Customize triggers with conditions: https://learn.microsoft.com/en-us/power-automate/customize-triggers

### 4.2 Ownership and continuity
Operationally, flows/apps must have:
- a clear owner
- a backup owner
- documentation (what list, what fields, what outcome)

In this course, you will record these in the Module 12 worksheets in your participant pack.

---

## 5) Power Apps: customizing a SharePoint list form
SharePoint/Microsoft Lists can integrate with Power Apps to customize the list form.
Typical admin use cases:
- re-order fields
- show/hide fields based on conditions
- add helper text

References (official):
- Customize a form for a SharePoint list: https://learn.microsoft.com/en-us/sharepoint/dev/business-apps/power-apps/get-started/create-your-first-custom-form
- Understand SharePoint forms integration: https://learn.microsoft.com/en-us/power-apps/maker/canvas-apps/sharepoint-form-integration

---

## 6) Governance: environments and DLP policies (admin view)
As an admin, your key governance controls include:
- **Environments** (where apps/flows live)
- **DLP policies** (which connectors can be used, and which can be combined)

DLP concept (simplified):
- Connectors are classified and constrained so data doesn’t move to risky places.

References (official):
- Manage data policies (DLP): https://learn.microsoft.com/en-us/power-platform/admin/prevent-data-loss
- DLP strategy guidance: https://learn.microsoft.com/en-us/power-platform/guidance/adoption/dlp-strategy

Shared-tenant note:
- DLP/environment changes can impact many apps/flows. In this course delivery, treat these as Trainer-only demonstrations.
