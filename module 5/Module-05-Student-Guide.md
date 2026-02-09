# Module 5 — Managing Metadata and the Term Store

## Module objectives
By the end of this module, you will be able to:
- Explain basic information architecture (IA) principles and why metadata improves findability and governance
- Describe the Term store hierarchy (groups → term sets → terms) and what “managed metadata” means
- Describe term store roles (term store admin, group manager, contributor) and how delegated term management works
- Apply managed metadata to a document library via a managed metadata column

---

## 5.1 Information architecture (IA) principles (practical)

### Why admins care
In real environments, users don’t search for “the file name” — they search for concepts:
- contract type,
- department,
- customer,
- confidentiality,
- status.

Metadata enables consistent tagging so:
- search refiners work better,
- views can group/filter reliably,
- governance decisions are easier to implement.

### Common IA anti-patterns
- Too many folders that encode meaning only a few people understand.
- Free-text columns for critical classification (results in spelling variants and inconsistent values).

---

## 5.2 Managed metadata vs site columns (what to choose)

### Site columns (non-taxonomy)
A site column is a column definition that can be reused (for example across multiple libraries on a site).
- Great for: dates, numbers, yes/no, choice columns with a small stable list.

### Managed metadata (term store)
A managed metadata column lets users pick a term from a managed term set.
- Great for: organization-wide classification, controlled vocabulary, synonyms, and re-use across many sites.

Key terms (Microsoft definition level):
- **Managed metadata column**: a column that lets users select terms from a term set.
- **Term group**: a container for term sets that share common security requirements.
- **Term set**: a set of related terms.
- **Term**: a single label/value in a term set.

---

## 5.3 Term store hierarchy and roles (delegated management)

### Where term store lives
Term store is accessed in the **SharePoint admin center** under **Content services**.

### Roles you will hear in operations
- **Term store admin**: can create/delete term groups, assign managers/contributors, manage languages.
- **Group manager**: can manage contributors; can manage term sets within the group.
- **Contributor**: can create/change term sets and terms (within scope).

Operational note:
- You can label a term set with “Owner / Contact / Stakeholders”, but those labels don’t automatically grant term store permissions.

---

## 5.4 Applying metadata to libraries (Northwind contracts)

In the course scenario, Northwind wants consistent classification of contracts.
In your `NW-Pxx-Contracts` library you will implement:
- A managed metadata column for “Contract Type” (tag documents as NDA / MSA / SOW, etc.).

Why this design is safe in a shared tenant:
- Everyone uses `NW-Pxx-...` names.
- Terms live in separate participant term groups/sets (or local term sets as fallback).

---

## Summary
You should now be able to:
- explain why metadata improves consistent classification,
- explain term store hierarchy and roles,
- create a term group and term set safely (or use a site-scoped fallback),
- and apply a managed metadata column to a library.

---

## Knowledge check
1. What is the difference between a “term set” and a “term group”?
2. When would you prefer a managed metadata column over a standard Choice column?
3. Who can create a new term group in the term store?
4. What does “delegated term management” mean in practice?
5. What shared-tenant behaviors should you avoid when working with the term store?

---

## References (official)
- Introduction to managed metadata: https://learn.microsoft.com/en-us/sharepoint/managed-metadata
- Open the Term store management tool: https://learn.microsoft.com/en-us/sharepoint/open-term-store-management-tool
- Set up a new group for term sets: https://learn.microsoft.com/en-us/sharepoint/set-up-new-group-for-term-sets
- Set up a new term set: https://learn.microsoft.com/en-us/sharepoint/set-up-new-term-set
- Create and manage terms in a term set: https://learn.microsoft.com/en-us/sharepoint/create-and-manage-terms
- Assign roles and permissions to manage term sets: https://learn.microsoft.com/en-us/sharepoint/assign-roles-and-permissions-to-manage-term-sets
- Troubleshoot creating a default site term set (local fallback pattern): https://learn.microsoft.com/en-us/troubleshoot/sharepoint/administration/create-default-site-term-set
