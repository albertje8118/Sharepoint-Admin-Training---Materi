# Module 6 — Search in SharePoint Online and Microsoft Search (Student Guide)

## Learning objectives
By the end of this module, you will be able to:
1. Explain how SharePoint Online search and Microsoft Search relate, and what “security trimming” means.
2. Describe Microsoft Search “answers” (Bookmarks and Acronyms) and who manages them.
3. Validate whether a site/library is searchable and request reindexing when appropriate.
4. Use basic query syntax (phrases, AND/OR/NOT, simple property restrictions) to troubleshoot.

## 1) Search concepts you need as an admin

### 1.1 Security trimming (the #1 rule)
Search results are always filtered by permissions:
- Users only see what they already have access to.
- If content “is missing” for one user but not another, start with permissions and sharing.

### 1.2 SharePoint search vs Microsoft Search
In modern Microsoft 365, users search from multiple entry points (for example SharePoint, Microsoft 365, and other apps). The underlying experience is increasingly unified, but **administration surfaces differ**.

Admin take-away:
- Some configuration is **site/library-scoped** (safe for individual participant labs).
- Some configuration is **organization-level** (must be isolated with naming, audience scoping where possible, and trainer governance).

## 2) Microsoft Search “answers” (Bookmarks and Acronyms)

### 2.1 Where admins manage them
Microsoft documents the primary admin entry point as:
- Microsoft 365 admin center → **Settings** → **Search & intelligence** (Microsoft Search)

Roles:
- A Global admin can assign “Search admin” and “Search editor” roles.
- Search admins/editors can curate organizational content such as Bookmarks and Acronyms.

### 2.2 Bookmarks (admin-curated)
Bookmarks are curated links that can be triggered by keywords.

Operational notes:
- A published bookmark is available **immediately** after publishing.
- Avoid “reserved keywords” for generic terms in a shared tenant; use **training-only** keywords.

### 2.3 Acronyms (admin-curated + system-curated)
Acronyms can be curated by admins, and can also be discovered automatically (system-curated).

Operational notes:
- Admin-curated acronyms can be Draft or Published.
- Microsoft notes it can take **up to a day** for published acronyms to become available.

### 2.4 Q&A answers (status note)
Historically, Microsoft Search included Q&A-style “answers”. As of the retirement of Microsoft Search in Bing (March 2025), Microsoft documentation indicates some answer types (including Q&As) are no longer available in that context.

Practical training guidance (2026):
- Teach Q&As as a **concept**, but design hands-on labs around **Bookmarks + Acronyms**, which remain commonly available.
- If your tenant still exposes Q&As in the admin center, treat them as optional.

## 3) Search schema: crawled vs managed properties (overview)

### 3.1 Why admins care
When content is crawled, metadata and content are discovered as properties.
- **Crawled properties** are what the crawler finds.
- **Managed properties** are what’s kept in the index and can be queried.

Admin warning:
- Changing mappings/managed properties can affect other Microsoft 365 experiences.
- Search schema changes typically require **reindexing** the affected site/library/list.

In this course delivery:
- Treat schema changes as trainer-led unless explicitly assigned.
- Most learners should focus on: “How do I validate indexing?” and “How do I troubleshoot missing results?”

## 4) Indexing and reindexing: what to do (and what not to do)

### 4.1 When reindexing is appropriate
Reindex is appropriate after changes that affect what should be in the index (for example, schema changes, or library/site search visibility changes).

Caution:
- Reindexing a site can create a high load; avoid reindexing unless needed.

### 4.2 Typical troubleshooting flow
When a document doesn’t appear in search:
1. Confirm the user has access to the item (security trimming).
2. Confirm the site/library is allowed to appear in search results.
3. If you changed a relevant setting recently, request a **reindex**.
4. Allow time for indexing; validate again.

## 5) Query basics for admins (KQL/KeyQL concepts)
SharePoint supports Keyword Query Language (KQL/KeyQL) concepts that help with targeted troubleshooting.

Useful patterns:
- Phrase search: `"Northwind Search Drill Alpha"`
- Boolean operators (operators should be uppercase): `Alpha AND Harborlight`
- Exclusion: `Alpha -Beta`
- Simple property restrictions (examples depend on what is indexed/queryable):
  - `author:"First Last"`
  - `filetype:docx`
  - `title:(Contract OR Agreement)`

Training note:
- Not every property is queryable in every experience. Use property restrictions as a troubleshooting tool, not as a guarantee.

## References (official)
- Set up Microsoft Search (roles, admin center path): https://learn.microsoft.com/en-us/microsoftsearch/setup-microsoft-search
- Manage bookmarks: https://learn.microsoft.com/en-us/microsoftsearch/manage-bookmarks
- Manage acronyms: https://learn.microsoft.com/en-us/microsoftsearch/manage-acronyms
- Reindex site/library/list: https://learn.microsoft.com/en-us/sharepoint/crawl-site-content
- Manage the search schema in SharePoint (overview): https://learn.microsoft.com/en-us/sharepoint/manage-search-schema
- KQL/KeyQL syntax reference: https://learn.microsoft.com/en-us/sharepoint/dev/general-development/keyword-query-language-kql-syntax-reference
- Retirement note (context on answer types in Bing experience): https://learn.microsoft.com/en-us/microsoftsearch/retirement-microsoft-search-bing
