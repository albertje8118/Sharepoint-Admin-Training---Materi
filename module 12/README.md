# Module 12 (Optional) — Workflow Automation with Power Automate and Power Apps

**Course:** Modern SharePoint Online for Administrators (3-Day, 2026 aligned)  \
**Module duration (suggested):** 1.25–1.75 hours instruction + 60–90 min lab  \
**Level:** Intermediate

## Scenario continuity (shared tenant)
This optional module continues the Northwind scenario by turning the Module 7 intake list into a simple workflow:
- Scenario overview: ../scenario/Lab-Scenario-Overview.md

Shared-tenant note:
- Participants create flows/apps **only** against their own `NW-Pxx-...` artifacts.
- Tenant-wide governance (environments, DLP policies, tenant settings) is **Trainer-only** by default.

## Learning objectives
By the end of this module, learners will be able to:
1. Explain how SharePoint lists act as a “system of record” for low-code automation.
2. Build a basic approval workflow in Power Automate using SharePoint triggers and **Start and wait for an approval**.
3. Update a SharePoint list item based on approval outcome (Approved/Rejected) safely.
4. Customize a SharePoint list form using Power Apps (embedded form experience).
5. Describe governance controls for Power Platform in an admin context (environments + DLP policies) and how they constrain connectors.

## Deliverables in this folder
- Student guide chapter: Module-12-Student-Guide.md
- Lab guide: Lab-12-Workflow-Automation-PowerPlatform.md
- Slides outline: Module-12-Slides.md

## Validation checkpoint (end of module)
Learners can:
- demonstrate an approval flow that updates their `NW-Pxx-AppRequests` item status
- explain how to avoid infinite loops in list-triggered flows
- publish a simple Power Apps form customization for the list
