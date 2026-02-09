# Lab 1: Exploring the Microsoft 365 Environment (Admin Baseline)

**Estimated time:** 60–75 minutes  
**Lab type:** UI (admin portals)  
**Goal:** Build muscle memory for tenant baseline checks: admin centers, tenant settings areas, service health, and message center.

## Scenario context (shared tenant)
This lab is part of the course-wide scenario:
- Scenario overview: ../scenario/Lab-Scenario-Overview.md

You are working in **one shared Microsoft 365 tenant** with **10 participants**, all with admin roles. This lab is intentionally **verification-only** and should not disrupt other participants.

### Participant ID
Ask your trainer for your Participant ID: **P01–P10**.

> Tip: Use your Participant ID when capturing notes (e.g., “P03: Service health shows…”) so your results are easy to track.

## Deliverable for this lab (what you must produce)
Create a simple **Tenant Baseline Worksheet** in your notes (or a text file) with the fields below. You will fill it in during the lab.

Minimum fields:
- Participant ID (Pxx)
- Tenant name
- Admin roles you have (as shown in the portal)
- SharePoint admin center sections you can access
- SharePoint org-level sharing setting (record only; no changes)
- OneDrive org-level sharing setting (record only; no changes)
- One active (or most recent) Service health item for SharePoint/OneDrive (title + status)
- Two Message center items relevant to SharePoint/OneDrive (title + why it matters)
- One “limit/boundary” you looked up on Microsoft Learn (topic + URL)
- One risk you spotted + one mitigation you would propose

## Prerequisites
- A Microsoft 365 tenant (demo tenant is fine)
- Accounts:
  - 1x admin account with **SharePoint Administrator** role (or Global Administrator)
  - (Optional) 1x non-admin user for comparison
- Browser access to admin portals

## Required roles
- SharePoint Administrator (minimum)  
Some items may require Global Administrator depending on tenant configuration.

---

## Exercise 0 — Create your baseline worksheet (setup)
1) Start a new note named: `Module1-Baseline-Pxx`.
2) Add the Minimum fields list (above) as headings.

### Expected result
- You have a place to capture evidence and findings.

### Validation check
- Your worksheet includes Participant ID (Pxx).

---

## Exercise 1 — Access Microsoft 365 admin center
1) Sign in to the Microsoft 365 admin portal: `https://admin.microsoft.com`.
2) Confirm you are in the correct tenant (check tenant name/branding in the portal header or organization profile area).
3) Locate the left navigation and identify:
   - Where admin roles are managed (you will configure roles in Module 2)
   - Where service notifications are shown (Service health / Message center areas)

4) Open your profile/account details and find where your **roles** are displayed.

### Expected result
- You can access the Microsoft 365 admin center without authorization errors.

### Validation check
- Record (write down) the tenant name and your current admin role(s) as displayed in your account/profile info.

Optional (for continuity):
- Create a one-line note for your course log: `Pxx — Tenant: <name> — Roles: <roles>`

---

## Exercise 1B — Map the admin centers (navigation mastery)
1) In the Microsoft 365 admin center, find the list of **admin centers**.
2) Identify which admin centers are likely relevant to SharePoint administration in this course (minimum: SharePoint; optionally: Entra, Purview, Teams).
3) For each one you list, write a one-sentence description of what you expect to manage there.

### Expected result
- You can quickly choose the correct admin center for a task.

### Validation check
- Your worksheet includes at least three admin centers and their purpose.

---

## Exercise 2 — Open SharePoint admin center and review tenant settings areas
> Note: UI labels can vary by tenant and update cadence. If the exact wording differs, focus on locating the **tenant-level settings categories**.

1) From the Microsoft 365 admin center, open the **SharePoint admin center** (via “Admin centers” / “All admin centers”).
2) In SharePoint admin center, locate the areas for:
   - **Sites** management (active sites list)
   - **Policies** or tenant-wide settings (sharing/access-related configuration)
   - **Settings** areas that impact SharePoint/OneDrive behavior (where available)
3) Open the active sites list and pick one site you have access to. Confirm you can view:
   - Site URL
   - Site owner
   - Storage usage (if visible)
   - Sharing settings summary (if visible)

4) Still in SharePoint admin center, locate **Policies > Sharing** and record the current org-level sharing settings for:
   - SharePoint
   - OneDrive

> Important: In Module 1, do **not** change any org-level settings.

### Expected result
- You can open SharePoint admin center and view at least one site’s admin summary.

### Validation check
- Capture (write down) the URL of one site and its listed owner.

Optional (for continuity):
- Record one site URL you will revisit later (if your trainer has assigned you a practice site, use that URL).

Additional validation (required):
- Record the current **org-level** sharing setting for SharePoint and OneDrive (read-only).

---

## Exercise 2B — Identify “tenant-wide vs site-specific” controls
1) In SharePoint admin center, pick one control that is clearly tenant-wide (example: org-level sharing).
2) Pick one control that is clearly site-scoped (example: a site’s sharing setting).
3) Write two bullet points explaining why confusing these scopes can cause incidents in a shared tenant.

### Expected result
- You can distinguish scope before making changes.

### Validation check
- Your worksheet contains one tenant-wide control and one site-level control.

---

## Exercise 3 — Check Service health
1) Return to Microsoft 365 admin center.
2) Open **Health** and then **Service health** (wording may be “Service health” under Health).
3) Filter or locate entries for:
   - SharePoint Online
   - OneDrive for Business (if listed separately)
4) Review:
   - Current status (healthy/advisory/incident)
   - Any active communications that might affect users/admin tasks

5) Open one SharePoint/OneDrive item (even if it’s informational) and answer:
   - What user impact is described?
   - What admin action is recommended (if any)?

### Expected result
- You can confirm whether SharePoint/OneDrive has any current advisories/incidents.

### Validation check
- Write down either:
  - “No active issues for SharePoint/OneDrive”, or
  - The title/ID of one active advisory/incident.

Additional validation (required):
- Record one item’s status + a one-sentence impact summary.

---

## Exercise 4 — Review Message center for upcoming changes
1) In Microsoft 365 admin center, open **Message center**.
2) Locate at least **two** messages that could impact SharePoint administration (examples: sharing changes, OneDrive changes, search changes, compliance/audit changes).
3) For each message, record:
   - Message title
   - Category (if shown)
   - Published date / rollout timing (if shown)
   - Any recommended admin actions

4) For each message, write:
   - “Who needs to know?” (helpdesk, site owners, security/compliance, end users)
   - “What will we do about it?” (monitor, update policy, update training, communicate)

### Expected result
- You can find and interpret at least one upcoming change relevant to SharePoint.

### Validation check
- Write down two message titles and one potential action item for each.

Optional (for continuity):
- Note whether the message could impact any of the following later modules: sharing, search, compliance/audit, OneDrive.

---

## Exercise 5 — Verify one limit/boundary using Microsoft Learn (research task)
1) Open Microsoft Learn in a browser.
2) Find an official SharePoint Online “limits/boundaries” reference relevant to one scenario:
   - Large lists/libraries
   - External sharing
   - Site/storage constraints
3) Capture:
   - The page URL
   - The specific limit/boundary topic you researched
   - One sentence on why it matters to the Northwind scenario

### Expected result
- You can find and cite the authoritative source rather than guessing.

### Validation check
- Your worksheet includes 1 Microsoft Learn URL and a scenario tie-in.

---

## Cleanup
No cleanup required.

---

## Troubleshooting (common issues)
1) **I can’t see SharePoint admin center**
   - Confirm your account has **SharePoint Administrator** or **Global Administrator** role.
   - Some tenants require time for role assignment to take effect (try sign out/in).

2) **Service health or Message center is missing**
   - Verify you’re in the Microsoft 365 admin center (not a workload admin center).
   - Confirm your role includes permission to view health/messages (higher roles may be required).

3) **I can open SharePoint admin center but can’t view site details**
   - You may lack permission to manage that site, or the tenant restricts visibility.
   - Try selecting a site you own or ask the instructor for a test site.
