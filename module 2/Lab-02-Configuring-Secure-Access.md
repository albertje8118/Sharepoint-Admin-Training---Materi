# Lab 2: Configuring Secure Access (Shared Tenant)

**Estimated time:** 60–90 minutes  
**Lab type:** UI (SharePoint admin center + Microsoft 365/Entra portals)  
**Goal:** Understand where to configure secure access controls (Entra roles/collaboration + SharePoint sharing), and safely apply site-level sharing settings in a shared tenant.

## Scenario context (continuity)
This lab continues the scenario:
- Scenario overview: ../scenario/Lab-Scenario-Overview.md
- Partner for the story: **Fabrikam (external)**

## Shared-tenant operating rules (critical)
- **Trainer-only:** tenant-wide sharing changes and guest invitations (unless the trainer instructs otherwise).
- **Participants:** changes should be limited to your own practice site (`NW-Pxx-ProjectSite`) to avoid impacting other participants.

## Deliverables for this lab (what you must produce)
Create a short “Secure Access Worksheet” in your notes named: `Module2-SecureAccess-Pxx`.

Minimum fields:
- Participant ID (Pxx)
- The minimum admin role required for 5 tasks (you choose the tasks)
- Org-level sharing setting observed (SharePoint + OneDrive)
- Proposed site-level sharing setting for `NW-Pxx-ProjectSite` (with justification)
- Proposed guest collaboration approach for Fabrikam (guest vs link-based) and why
- One Entra external collaboration setting you would review before inviting guests
- Conditional Access: list the names of any existing policies you can see (or “none visible”)
- Troubleshooting drill answers (Exercise 6)

## Prerequisites
- Participant ID assigned by trainer: **P01–P10**
- Access to Microsoft 365 admin center: https://admin.microsoft.com
- Access to SharePoint admin center
- A participant-isolated practice site (you will create it during the course when instructed)

If your practice site is not yet available:
- Complete the verification portions of this lab.
- Observe the configuration demo performed by the trainer.

## Required roles
- To *view* settings: SharePoint Administrator is typically sufficient.
- To *assign admin roles* and *invite guest users*: higher roles (for example, Global Administrator or User Administrator / Privileged Role Administrator) may be required.

> In this course, assume role assignment and guest invitations are **Trainer-only**.

---

## Exercise 0 — Build your “control scope map” (required)
Create a 3-column table in your worksheet with these headers:
- Control area
- Where it lives (Entra / SharePoint tenant / SharePoint site)
- Who should change it (Trainer-only / Participant)

Add at least 8 rows. Include:
- Org-level external sharing
- Site-level external sharing
- Guest invitation policy
- Admin role assignment
- Conditional Access policy
- Site permissions

### Expected result
- You can identify the correct layer before making changes.

### Validation check
- Your table includes at least 8 controls.

## Exercise 1 — Identify the correct control scope (warm-up)
For each request below, write down where you would make the change:
- A) “Allow external vendor access to only one project site.”
- B) “Block anonymous (Anyone) links across the tenant.”
- C) “Restrict who can invite guest users.”
- D) “Give an IT staff member the ability to manage SharePoint admin center.”

### Expected result
- You can map the request to **Entra**, **SharePoint tenant**, or **SharePoint site** scope.

### Validation check
- Compare answers as a class.

---

## Exercise 2 — Review organization-level external sharing (Trainer-only)
> This exercise is **Trainer-only** because it affects the entire tenant.

1) Open SharePoint admin center.
2) Go to **Policies** > **Sharing**.
3) Review the organization-level external sharing level for:
   - SharePoint
   - OneDrive
4) Record the current level.

Optional (trainer decision):
- If the course is intended to support guest collaboration, set SharePoint (and therefore OneDrive) to a level that supports guests (commonly **New and existing guests**).

### Expected result
- The class understands the tenant baseline for external sharing.

### Validation check
- Each participant records the org-level sharing setting in their course log.

---

## Exercise 2B — Review Entra external collaboration settings (read-only)
This exercise is read-only and safe in a shared tenant.

1) Open Microsoft Entra admin center: https://entra.microsoft.com
2) Locate the area for external collaboration / external identities settings.
3) Identify at least one setting that influences guest invitations (for example, who can invite guests).

### Expected result
- You can identify Entra as a dependency for guest collaboration.

### Validation check
- Record one setting you would review before enabling vendor collaboration.

---

## Exercise 3 — Configure site-level external sharing for your practice site (Participant)
This exercise is safe in a shared tenant **if** each participant works on their own practice site.

1) In SharePoint admin center, go to **Sites** > **Active sites**.
2) Locate your practice site (created by you when instructed), typically named: `NW-Pxx-ProjectSite`.
3) Open the site details.
4) Go to the **Settings** tab and select **More sharing settings**.
5) Choose an external sharing option consistent with the scenario.

Suggested scenario choice:
- Use **New and existing guests** if the tenant baseline allows it.
- If your tenant baseline is more restrictive, select the most permissive option that is still allowed.

6) Select **Save**.

### Expected result
- Your practice site has a sharing posture suitable for a vendor collaboration project without changing tenant-wide settings.

### Validation check
- In SharePoint admin center, confirm the site’s external sharing setting reflects your chosen value.

If your practice site does not exist yet:
- Write down your proposed site-level sharing setting and justification in your worksheet.
- You will implement it later when `NW-Pxx-ProjectSite` exists.

---

## Exercise 4 — Guest users: invitation and verification (Trainer-only + Participant verification)
Inviting guests can create directory objects and send emails. To avoid disruption:
- **Trainer-only:** invites (or prepares) a guest account for the class.
- **Participants:** verify the guest exists and understand the invitation status.

### Trainer-only steps (demo)
1) Open Microsoft Entra admin center: https://entra.microsoft.com
2) Go to **Entra ID** > **Users**.
3) Select **New user** > **Invite external user**.
4) Invite the guest user using an email address provided by the trainer.
5) Confirm the guest appears in the directory (user type: Guest) and note invitation status.

### Participant verification steps
1) Open Entra admin center and locate the guest user created by the trainer.
2) Verify:
   - User type shows Guest
   - Invitation state (Pending acceptance vs Accepted), if visible

Optional (only if trainer provides a test guest and approves):
- Add the guest to your practice site (e.g., visitors) or share one test document using “Specific people”, then validate the guest can access.

### Expected result
- Participants understand how guest users enter the directory and how to check status.

### Validation check
- Each participant records the guest display name and invitation state.

---

## Exercise 4B — Optional: Create a participant security group (safe, reusable)
This creates a directory object but does not change tenant-wide policy.

If you have permission to create groups:
1) In Entra admin center, create a **Security group** named: `NW-Pxx-ProjectOwners`.
2) Add yourself as a member.

If you do not have permission:
- Record “Group creation restricted in this tenant” and continue.

### Expected result
- You have a reusable group for later permission labs.

### Validation check
- Record whether the group exists and whether you are a member.

---

## Exercise 5 — Admin roles vs site permissions (Trainer demo + participant self-check)

### Trainer demo: least-privilege role assignment (do not mass-assign)
1) In Microsoft 365 admin center (https://admin.microsoft.com), go to **Users** > **Active users**.
2) Select a user and open **Manage roles**.
3) Show the difference between:
   - SharePoint Administrator
   - Global Administrator
4) Explain that role changes can take time to apply.

### Participant self-check
1) Open your user account details in Microsoft 365 admin center.
2) Review what admin roles you currently have.
3) Compare that to your permissions inside your practice site (site owner/member/visitor).

### Expected result
- Participants can articulate the difference between tenant roles and site permissions.

### Validation check
- Write a one-sentence statement: “I am a(n) <role> in the tenant and a(n) <role> in my practice site.”

---

## Exercise 6 — Troubleshooting drill (required)
For each scenario below, write the **first three checks** you would perform, in order, and what you expect to find.

Scenario 1: “Sharing options are greyed out for a site owner.”

Scenario 2: “A guest user receives ‘Your organization’s policies do not allow you to share with these users’.”

Scenario 3: “External sharing was working yesterday; today, shared links stopped working.”

Rules:
- Your checks must include at least one item from each layer: **Entra**, **SharePoint tenant**, **SharePoint site**.

### Expected result
- You can troubleshoot by scope instead of guessing.

### Validation check
- Your worksheet contains 3 scenarios x 3 checks.

---

## Cleanup
- No cleanup required for site sharing changes.
- If any guest account was created for training, the trainer may remove it after the course.

---

## Troubleshooting (common issues)
1) **I can’t change site sharing settings**
   - You must be at least SharePoint Administrator to change site sharing settings (site owners can’t change this).
   - Confirm you are editing your own practice site.

2) **Sharing options are greyed out / missing**
   - The org-level sharing setting might be too restrictive.
   - Ask the trainer to verify **Policies > Sharing** in SharePoint admin center.

3) **Guest can’t access even though site sharing allows it**
   - Guest invitation not redeemed.
   - Entra external collaboration settings or Conditional Access might block access.
   - Group/Teams-connected settings can affect access if applicable.

---

## Reference links (Microsoft Learn)
- Organization-level sharing settings (SharePoint admin center): https://learn.microsoft.com/en-us/sharepoint/turn-external-sharing-on-or-off
- Site-level sharing settings: https://learn.microsoft.com/en-us/sharepoint/change-external-sharing-site
- Restrict sharing by domain (advanced): https://learn.microsoft.com/en-us/sharepoint/restricted-domains-sharing
- Invite guest users (Entra): https://learn.microsoft.com/en-us/entra/external-id/add-users-administrator
- Assign admin roles (Microsoft 365 admin center): https://learn.microsoft.com/en-us/microsoft-365/admin/add-users/assign-admin-roles
