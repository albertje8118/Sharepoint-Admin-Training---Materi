# Module 5 — Managing Metadata and the Term Store (Slides)

> Format: slide title → bullets → speaker notes.

---

## Slide 1 — Module 5: Metadata and Term Store
- Information architecture (IA) basics
- Managed metadata and the Term store
- Delegated term management roles
- Apply metadata to libraries
- Lab: Create taxonomy + managed metadata column

Speaker notes:
- Set the expectation: term store is tenant-wide. We will use `NW-Pxx-...` naming to avoid collisions.

---

## Slide 2 — Why metadata matters
- Consistent classification
- Better search refiners
- Better views (group/filter)
- Less dependency on “folder tribal knowledge”

Speaker notes:
- Use the Northwind contracts example to make it concrete.

---

## Slide 3 — Managed metadata vs site columns
- Site columns: reusable field definitions (dates, choices, numbers)
- Managed metadata: terms from a controlled term set
- Choose based on governance needs and re-use scope

Speaker notes:
- Highlight controlled vocabulary and synonyms as key value.

---

## Slide 4 — Term store hierarchy
- Term group → term sets → terms
- Groups provide security boundaries (who can manage what)

Speaker notes:
- Emphasize: groups exist to match security requirements.

---

## Slide 5 — Delegated term management roles
- Term store admin
- Group manager
- Contributor

Speaker notes:
- Clarify: Owner/Contact/Stakeholders labels do not grant permissions.

---

## Slide 6 — Shared-tenant safety approach
- Only create `NW-Pxx-...` term groups/sets
- Never edit others’ taxonomy
- If you can’t create a group, use local term set fallback

Speaker notes:
- This keeps the tenant clean and avoids accidental cross-impact.

---

## Slide 7 — Lab preview: Northwind Contracts taxonomy
- Create `NW-Pxx-TermGroup`
- Create term set `NW-Pxx-ContractType`
- Add terms (NDA, MSA, SOW, Renewal)
- Add managed metadata column to `NW-Pxx-Contracts`
- Tag documents

Speaker notes:
- Reinforce that the library was created in Module 4.

---

## Slide 8 — Validation & troubleshooting
- Can’t create term group → use local term set fallback
- Term set not visible → check “Available for tagging”
- Unexpected free-text terms → check open vs closed term set

Speaker notes:
- Encourage documenting exact symptom and where you checked.

---

## Slide 9 — Wrap-up
- You can explain term store concepts
- You can apply managed metadata to libraries
- You can describe delegated management operations

Speaker notes:
- Point learners to references.

---

## References
- Managed metadata overview: https://learn.microsoft.com/en-us/sharepoint/managed-metadata
- Open Term store tool: https://learn.microsoft.com/en-us/sharepoint/open-term-store-management-tool
- Set up term group: https://learn.microsoft.com/en-us/sharepoint/set-up-new-group-for-term-sets
- Set up term set: https://learn.microsoft.com/en-us/sharepoint/set-up-new-term-set
- Create/manage terms: https://learn.microsoft.com/en-us/sharepoint/create-and-manage-terms
- Term store roles: https://learn.microsoft.com/en-us/sharepoint/assign-roles-and-permissions-to-manage-term-sets
