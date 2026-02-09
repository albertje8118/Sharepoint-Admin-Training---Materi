# Module 12 (Optional) — Workflow Automation with Power Automate and Power Apps (Slides Outline)

## Slide 1 — Module title + outcomes
- Workflow Automation with Power Automate and Power Apps
- Outcomes:
  - Build a SharePoint-list-triggered approval flow
  - Update item status based on approval outcome
  - Customize a list form using Power Apps
  - Explain environments + DLP governance at a high level

## Slide 2 — Admin framing: why workflows matter
- Workflows are operational assets
- Ownership, failure handling, notifications
- Shared tenant: scope to `NW-Pxx` only

## Slide 3 — SharePoint + Power Automate integration
- SharePoint triggers: item created / created or modified
- Prefer “item created” for approval workflows that update the same item

Reference:
- https://learn.microsoft.com/en-us/sharepoint/dev/business-apps/power-automate/sharepoint-connector-actions-triggers

## Slide 4 — Approval pattern
- Trigger: item created
- Action: Start and wait for an approval
- Condition: outcome
- Update item: Status and Notes

References:
- https://learn.microsoft.com/en-us/power-automate/get-started-approvals
- https://learn.microsoft.com/en-us/power-automate/modern-approvals

## Slide 5 — Pitfall: loops and unnecessary runs
- Created/modified triggers + update item = re-trigger risk
- Use trigger conditions when needed

Reference:
- https://learn.microsoft.com/en-us/power-automate/customize-triggers

## Slide 6 — Power Apps custom form
- Integrate → Power Apps → Customize forms
- Re-order fields, add helper text, conditional visibility

Reference:
- https://learn.microsoft.com/en-us/sharepoint/dev/business-apps/power-apps/get-started/create-your-first-custom-form

## Slide 7 — Governance (admin view)
- Environments: where solutions live
- DLP policies: connector boundaries and allowed combinations

References:
- https://learn.microsoft.com/en-us/power-platform/admin/prevent-data-loss
- https://learn.microsoft.com/en-us/power-platform/guidance/adoption/dlp-strategy

## Slide 8 — Lab briefing
- Use your `NW-Pxx-AppRequests` list
- Approver = you (avoid spamming)
- Build flow + test + update item
- Customize form + publish

---

## Trainer demo script (minimal talk-track)
1) Safety framing (30 seconds)
- “Build only against your own `NW-Pxx` list.”
- “Approver = you or trainer-designated account.”

2) Approval flow concept (60 seconds)
- Trigger → Approval → Condition → Update item

3) Governance link (60 seconds)
- “DLP controls connectors; environments isolate workloads.”
- “In shared tenant training, governance changes are trainer-led only.”
