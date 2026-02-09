# Lab 05 — Creating and Managing Metadata (Term Store + Managed Metadata Columns)

**Module:** 5 — Managing Metadata and the Term Store  
**Estimated time:** 75–105 minutes  
**Lab type:** Individual (shared-tenant safe)

## Lab goal
Create a controlled vocabulary for Northwind contracts and apply it to the `NW-Pxx-Contracts` library using managed metadata.

You will:
- create a term group + term set + terms (or use a local term set fallback),
- apply a managed metadata column to a document library,
- tag documents consistently,
- and validate delegated management concepts.

## Prerequisites
- You have a participant ID `Pxx` (P01–P10).
- You have a site: `NW-Pxx-ProjectSite`.
- You have a library from Module 4: `NW-Pxx-Contracts` with a few documents.
- You can access SharePoint admin center (Term store) OR you can at least create managed metadata columns at the site/library level.

## Shared-tenant safety rules (critical)
- The Term store is tenant-wide: create only items prefixed with your `NW-Pxx-...` naming.
- Never edit or delete term groups/sets that are not yours.
- Do not change working languages or taxonomy global settings unless trainer-led.

---

## Exercise 1 — Design worksheet (10–15 minutes)

### 1A. Define the classification you want
Northwind wants consistent tagging for contracts.

Decide your term set(s). Use this baseline:
- Term set: `NW-Pxx-ContractType`
- Terms:
  - NDA
  - MSA
  - SOW
  - Renewal

Optional second term set (if time permits):
- Term set: `NW-Pxx-Department`
- Terms: Legal, Finance, Sales, Procurement

Record your design:
- Term group name: _____________________
- Term set name(s): _____________________
- Terms list confirmed: Yes / No

Validation check:
- You can explain why these terms should be controlled (not free text).

---

## Exercise 2 — Create term group + term set + terms (20–30 minutes)

### 2A. Create a participant-specific term group (global term store)
1. Open **SharePoint admin center**.
2. In left navigation, under **Content services**, open **Term store**.
3. Create a term group named:
   - `NW-Pxx-TermGroup`
4. (Optional but recommended) Add a description such as:
   - “Northwind training taxonomy for participant Pxx.”

Note:
- Creating a term group requires **term store admin** role.

### 2B. Set delegated management roles (within your group)
1. Select your `NW-Pxx-TermGroup`.
2. Set:
   - Group managers: add yourself
   - Contributors: add yourself

Optional buddy drill (only if trainer approves):
- Add one other participant as a Contributor to demonstrate delegated term management.

### 2C. Create a term set
1. Inside `NW-Pxx-TermGroup`, create a term set named:
   - `NW-Pxx-ContractType`
2. Configure (if available):
   - Submission policy: **Closed** (recommended for controlled taxonomy in this lab)
   - Available for tagging: **Enabled**

### 2D. Add terms
Add at least 4 terms:
- NDA
- MSA
- SOW
- Renewal

Validation check:
- You can locate your term set in Term store and see the terms.

---

## Exercise 3 — Local term set fallback (if you can’t create a term group) (10–20 minutes)

If you do not have permission to create a term group, use a site-scoped term set.

Goal: create a managed metadata column in your site/library and choose **Customize your term set** to create a local term set.

Guidance (official troubleshooting pattern):
- https://learn.microsoft.com/en-us/troubleshoot/sharepoint/administration/create-default-site-term-set

Validation check:
- You can select your new local term set when configuring the managed metadata column.

---

## Exercise 4 — Add a managed metadata column to `NW-Pxx-Contracts` (15–20 minutes)

### 4A. Create the column in the library
1. Go to your site `NW-Pxx-ProjectSite`.
2. Open the library `NW-Pxx-Contracts`.
3. Add a new column:
   - Name: `NW-Pxx-ContractType`
   - Type: **Managed metadata**
4. In term set settings, select:
   - your global term set `NW-Pxx-ContractType` (preferred), OR
   - your local term set (fallback)

### 4B. Tag documents
Tag at least 3 documents using the managed metadata column.

Validation check:
- Documents show consistent term values (type-ahead suggestions appear).

---

## Exercise 5 — Operational drills (15–20 minutes)

### 5A. Controlled vocabulary behavior
1. Attempt to type a value that is not in the term set.
2. Observe what happens.

Discuss:
- If your term set is **Closed**, should users be able to add new terms from tagging?

### 5B. Delegated management concept check
Answer (in 1–2 sentences each):
1. What actions require a term store admin?
2. What actions can a contributor do?
3. What’s the purpose of term groups from a security perspective?

### 5C. Validation checklist
Confirm all are true:
- You created `NW-Pxx-TermGroup` (or used local fallback)
- You created `NW-Pxx-ContractType` term set with ≥4 terms
- `NW-Pxx-Contracts` library has a managed metadata column
- ≥3 documents are tagged

---

## Cleanup / end state
Leave these in place for later modules:
- Term set (global or local): `NW-Pxx-ContractType`
- Library column: `NW-Pxx-ContractType`
- Tagged documents in `NW-Pxx-Contracts`

Do not delete other participants’ taxonomy objects.

---

## Troubleshooting (common issues)

1) **I can’t create a term group.**
- That usually means you are not a term store admin.
- Use Exercise 3 (local term set fallback) and continue.

2) **I can’t find Term store in SharePoint admin center.**
- Confirm you can access SharePoint admin center and have the correct role.
- Use the fallback path.

3) **The term set doesn’t appear when creating the column.**
- Confirm the term set is **Available for tagging**.
- Confirm you are selecting the correct term store group and term set.

4) **Users can type new values (unexpected).**
- Check whether the term set is **Open**.
- Closed term sets restrict adding new terms during tagging.

---

## References (official)
- Open Term store management tool: https://learn.microsoft.com/en-us/sharepoint/open-term-store-management-tool
- Set up a new group for term sets: https://learn.microsoft.com/en-us/sharepoint/set-up-new-group-for-term-sets
- Set up a new term set: https://learn.microsoft.com/en-us/sharepoint/set-up-new-term-set
- Create and manage terms: https://learn.microsoft.com/en-us/sharepoint/create-and-manage-terms
- Term store roles and permissions: https://learn.microsoft.com/en-us/sharepoint/assign-roles-and-permissions-to-manage-term-sets
- Introduction to managed metadata: https://learn.microsoft.com/en-us/sharepoint/managed-metadata
- Local term set troubleshooting pattern: https://learn.microsoft.com/en-us/troubleshoot/sharepoint/administration/create-default-site-term-set
