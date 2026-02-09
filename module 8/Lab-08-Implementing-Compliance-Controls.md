# Lab 08 — Content Governance and Compliance with Microsoft Purview

## Goal
Practice a shared-tenant-safe compliance workflow:
- apply **sensitivity** and **retention** labels to training content in SharePoint Online
- understand replication timing and common admin checks
- walk through an **eDiscovery** case workflow (case → search → optional hold/export) in a controlled way

## Estimated time
60–90 minutes (depending on whether eDiscovery is hands-on or demo-only)

## Prerequisites
### Access and roles
This lab is designed for a **single shared tenant**.

- **Participant prerequisites (required)**
  - You have access to your practice site: `NW-Pxx-ProjectSite`
  - You have edit permissions in your own site (Site Owner / Full Control recommended)

- **Trainer prerequisites (recommended, tenant-wide)**
  - A trainer account with permissions to configure Purview solutions.
  - Retention configuration permissions: Microsoft recommends adding admins to the **Compliance Administrator** role group, or using a custom role group with **Retention Management** (and optionally read-only equivalents).
  - Sensitivity label configuration permissions: use Information Protection role groups, or a custom role group with **Sensitivity Label Administrator**.
  - eDiscovery permissions for demos/hands-on: follow Microsoft guidance for eDiscovery permissions.

References (official):
- Retention permissions: https://learn.microsoft.com/en-us/purview/get-started-with-data-lifecycle-management#permissions
- Sensitivity permissions: https://learn.microsoft.com/en-us/purview/get-started-with-sensitivity-labels#permissions-required-to-create-and-manage-sensitivity-labels
- eDiscovery get started: https://learn.microsoft.com/en-us/purview/edisc-get-started

### Training content (use FAKE data only)
Use the participant pack TXT templates:
- `M08-Confidential-Statement-FAKE.txt`
- `M08-DLP-Test-Data-FAKE.txt` (only if trainer enables a DLP test)

---

## Shared-tenant safety rules (Module 8)
- **Trainer-only:** creating/publishing labels, DLP policies, and tenant-wide compliance settings.
- **Participants:** apply labels only to content inside **your own** `NW-Pxx-...` sites/libraries.
- Never use real personal data.

---

## Exercise 1 — Prepare labeled content (participant)
### Task A — Create a “Compliance” folder and upload FAKE content
1. Go to your site: `NW-Pxx-ProjectSite`.
2. Open your existing library `NW-Pxx-Contracts`.
3. Create a folder: `04-Compliance`.
4. Create a new document (Word online is fine) named: `NW-Pxx-Confidential-Statement.docx`.
5. Paste in the content from `M08-Confidential-Statement-FAKE.txt`.

Validation check:
- The document exists in `NW-Pxx-Contracts/04-Compliance`.

---

## Exercise 2 — Apply sensitivity label (participant)
> This exercise assumes the trainer has already published one or more sensitivity labels to participants.

### Task A — Apply a sensitivity label to the document
Apply one label to `NW-Pxx-Confidential-Statement.docx` using any method available in your environment:
- In SharePoint document library UI (file details/properties), OR
- In Office for the web/desktop (File → Info → Sensitivity), if available

Record your results:
- Label applied: ____________________
- Where you applied it (SharePoint / Office): ____________________

Validation check:
- The document shows a sensitivity label indicator (UI varies by client).

Troubleshooting (common)
- If labels don’t appear yet, this is usually replication timing. Microsoft recommends piloting labels and allowing time for changes to replicate.
- If you can’t see labels in any client, confirm with the trainer that a **publishing policy** includes you.

Reference (official publishing path):
- Sensitivity label publishing policy: Purview portal → **Solutions** → **Information Protection** → **Publishing policies** (then Publish label)
  - https://learn.microsoft.com/en-us/purview/create-sensitivity-labels#publish-sensitivity-labels-by-creating-a-label-policy

---

## Exercise 3 — Apply retention label (participant)
> This exercise assumes the trainer has already published one or more retention labels to SharePoint for the class.

### Task A — Apply a retention label to the same document
1. In SharePoint, select `NW-Pxx-Confidential-Statement.docx`.
2. Apply a retention label using the available UI for retention labels (details/properties experience varies).
3. Record the label name here:
   - Retention label applied: ____________________

Validation check:
- The file shows a retention label field/value in its properties (if your tenant UI exposes it).

Troubleshooting (timing)
- If retention labels are newly published to SharePoint/OneDrive, Microsoft notes they typically appear within one day, but you should allow up to seven days.
  - https://learn.microsoft.com/en-us/purview/create-apply-retention-labels#when-retention-labels-become-available-to-apply

Reference (official publishing paths):
- Purview portal →
  - Records management: **Solutions** → **Records Management** → **Policies** → **Label policies** (Publish labels)
  - Data lifecycle management: **Solutions** → **Data Lifecycle Management** → **Policies** → **Label policies** (Publish labels)
  - https://learn.microsoft.com/en-us/purview/create-apply-retention-labels#how-to-publish-retention-labels

---

## Exercise 4 — eDiscovery workflow (trainer demo by default)
In a shared tenant, eDiscovery can have broader impact if scoped incorrectly. Default delivery is:
- **Trainer performs** the demo steps.
- **Participants observe** and complete the worksheet.

### Optional: participant hands-on (only if trainer assigns eDiscovery permissions)
If (and only if) you have eDiscovery permissions assigned by the trainer, you may perform the steps below **scoped only to your own NW-Pxx site**.

### Task A — Create a case
1. Go to the Microsoft Purview portal: https://purview.microsoft.com/
2. Open the **eDiscovery** solution.
3. Go to **Cases (preview)** and create a new case.
   - Case name: `NW-Pxx-eDiscovery-Case`

Reference:
- https://learn.microsoft.com/en-us/purview/edisc-get-started

### Task B — Create a search scoped to your site
1. In your case, create a search.
2. Scope to your SharePoint site URL for `NW-Pxx-ProjectSite` (do not include other participants’ sites).
3. Add a keyword such as `CONFIDENTIAL` (from your FAKE document).
4. Run the search and review results.

Validation check:
- The search returns your `NW-Pxx-Confidential-Statement.docx`.

### Task C — (Optional) Create a hold
Only do this step if the trainer explicitly approves.
- Follow the official hold creation workflow in your case and scope the hold to **your own** locations.
Reference:
- https://learn.microsoft.com/en-us/purview/edisc-hold-create#create-a-hold

---

## Wrap-up checkpoint
You should be able to explain:
- why labels must be **published** before users can apply them
- why replication timing affects labs and rollouts
- the modern eDiscovery workflow (case → search → actions like hold/export)

What you should have in your site:
- `NW-Pxx-Contracts/04-Compliance/NW-Pxx-Confidential-Statement.docx` with an applied sensitivity label (if published)
- the same document with an applied retention label (if published)
