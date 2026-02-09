# Lab 06 — Configuring Search Experience (SharePoint + Microsoft Search)

**Module:** 6 — Search in SharePoint Online and Microsoft Search  \
**Estimated time:** 60–90 minutes  \
**Lab type:** Individual (shared-tenant safe)

## Lab goal
Validate that your Northwind content is discoverable and learn how admins curate “high-confidence” search answers.

You will:
- create two searchable “seed” documents,
- validate search results and indexing behavior,
- practice safe, end-user-facing KQL query patterns (scoping + property restrictions),
- reindex a library if required,
- create a training-scoped Bookmark,
- create a training-scoped Acronym (visibility may be delayed).

## Prerequisites
- You have a participant ID `Pxx` (P01–P10).
- You have a site: `NW-Pxx-ProjectSite`.
- You have a library: `NW-Pxx-Contracts` (from Module 4) with folders.
- You have access to SharePoint site and library settings (typically **Site Owner / Full Control** on `NW-Pxx-ProjectSite`).

### Permissions and access (what you need)
Because this is a shared tenant, some actions are tenant-wide and may be restricted.

- Exercises 1–3 (seed docs, indexing checks, KQL drills):
   - At minimum, **Edit/Contribute** permission to upload documents to `NW-Pxx-Contracts/02-InReview`.
   - To verify and change search visibility settings and to request reindexing, you generally need **Full Control** (Site Owner) on `NW-Pxx-ProjectSite`.
- Exercise 4–5 (Bookmarks/Acronyms in Microsoft Search):
   - One of these Microsoft 365 roles, assigned by a **Global admin**:
      - **Search admin**, or
      - **Search editor**.
- Exercise 6 (Search schema tour):
   - Tenant-level schema in SharePoint admin center is typically **Trainer-only** (SharePoint admin / Global admin).
   - Site collection schema view requires **Site Collection Admin** (or equivalent elevated permissions) for the site.

## Shared-tenant safety rules (critical)
- Bookmarks/Acronyms are organization-level curated content: use **only** `NW-Pxx` training names and keywords.
- Do not create reserved keywords for common words (for example “help”, “benefits”, “password”).
- Do not modify Search Schema mappings or managed properties unless trainer-led.

---

## Exercise 1 — Create and upload searchable seed documents (10–15 minutes)

### 1A. Create two documents from your TXT templates
1. Open your participant pack folder:
   - `participant-packs/Pxx/TXT-Templates/`
2. Use these templates:
   - `M06-Search-Seed-Document-A.txt`
   - `M06-Search-Seed-Document-B.txt`
3. Create two new Word documents (or plain text files) and paste the template content.
4. Save them as:
   - `NW-Pxx-Search-Seed-A.docx`
   - `NW-Pxx-Search-Seed-B.docx`

### 1B. Upload to SharePoint
1. Go to your site `NW-Pxx-ProjectSite`.
2. Open `NW-Pxx-Contracts`.
3. Upload both seed documents into folder `02-InReview`.

Validation check:
- Both documents are visible in the library and have correct names.

---

## Exercise 2 — Validate search and indexing (15–25 minutes)

### 2A. Search within the library
1. In `NW-Pxx-Contracts`, use the library search box.
2. Search for the unique phrases from the templates (they include your participant ID).

Expected result:
- You eventually see both documents in results.

### 2B. Validate search scope (library vs site)
This step helps you recognize the *scope* you’re searching.

1. From the library results page, switch the scope (if available) from **This library** to **This site**.
2. Search again for the same unique phrase.

Expected result:
- In **This library**, results should be limited to `NW-Pxx-Contracts`.
- In **This site**, results can include pages and content across `NW-Pxx-ProjectSite` (depending on what exists and has been indexed).

### 2C. If results are missing: check search visibility settings
If the documents don’t appear after a reasonable wait:

1) Confirm the library is searchable
1. In the library, open **Settings**.
2. Choose **Library settings** (or **More library settings**).
3. Open **Advanced settings**.
4. Confirm **Allow items from this document library to appear in search results?** is set to **Yes**.

2) Confirm the site is searchable
1. In the site, open **Settings**.
2. Open **Site settings** (or **Site information** → **View all site settings**).
3. Under **Search**, open **Search and offline availability**.
4. Confirm **Allow this site to appear in Search results** is set to **Yes**.

### 2D. Request reindexing (only if needed)
Reindexing is useful after a visibility/schema change (and can create load), so don’t do it “just because”.

Reindex the library:
1. Open **Library settings**.
2. Under **General Settings**, open **Advanced settings**.
3. Scroll to **Reindex Document Library**, and select the button.
4. Confirm the warning.

Optional: Reindex the site (use sparingly)
1. Site settings → **Search and offline availability**.
2. Use **Reindex site**.

Validation check:
- You can explain why you reindexed (and why you shouldn’t do it “just because”).

---

## Exercise 3 — KQL query drills (safe, end-user focused) (10–20 minutes)

SharePoint search supports Keyword Query Language (KQL). In KQL, queries are case-insensitive, but **operators are case-sensitive (uppercase)**.

### 3A. Capture the “Path” you will scope to
1. Open your library `NW-Pxx-Contracts`.
2. Copy the URL of the library (or the `02-InReview` folder).

Tip:
- Scoping with `Path:"<url>"` is a fast way to focus search to a specific library/folder.

### 3B. Run these KQL searches
Use the SharePoint search box while you’re in your site (or library). Try each query and observe how results change.

1) Phrase search (baseline)
- `"Northwind Search Drill Alpha - Pxx"`

2) Scope to your library (replace with your copied URL)
- `Path:"<YOUR-NW-Pxx-CONTRACTS-URL>" AND "Northwind Search Drill"`

3) Scope + file type
- `Path:"<YOUR-NW-Pxx-CONTRACTS-URL>" AND Filetype:docx`

4) Author restriction (use your display name as shown in M365)
- `Author:"<YOUR DISPLAY NAME>" AND "Northwind Search Drill"`

5) Combine conditions with parentheses
- `(Filetype:docx OR Filetype:txt) AND ("Northwind Search Drill" OR "Northwind Contract")`

Validation check:
- You can explain which part of each query is (a) free text and (b) a property restriction.
- You can explain why `AND`/`OR` must be uppercase.

---

## Exercise 4 — Create a training-scoped Bookmark (15–25 minutes)

If you don’t have Microsoft Search permissions, skip to Exercise 4.

### 4A. Open Microsoft Search admin area
1. Open Microsoft 365 admin center.
2. Go to **Settings** → **Search & intelligence**.
3. Open **Bookmarks**.

Reference: Microsoft documentation for managing bookmarks.

### 4B. Create the bookmark
Create a bookmark with training-only naming:
- Title: `NW-Pxx Project Site`
- URL: your `NW-Pxx-ProjectSite` home page URL
- Keywords (examples):
  - `NW Pxx ProjectSite`
  - `Northwind Pxx site`
  - `NW-Pxx Project Site`
- State: **Published**

Important:
- Do not use keywords that other participants may also use.

### 4C. Validate bookmark behavior
1. In SharePoint, use the main search box.
2. Search using one of your bookmark keywords.

Expected result:
- A published bookmark becomes available immediately after publishing.

---

## Exercise 5 — Create a training-scoped Acronym (10–15 minutes)

### 5A. Create the acronym
1. In Microsoft Search admin area, open **Acronyms**.
2. Create an acronym that is unique to you (no spaces):
   - Acronym: `NWPxx`
   - Stands for: `Northwind Participant Pxx`
   - Source/URL: your `NW-Pxx-ProjectSite` URL
   - State: Draft or Published

Operational reality:
- Microsoft notes published acronyms can take **up to a day** to become visible.

Validation check:
- You can explain why acronyms aren’t always a “same-day” training validation.

---

## Exercise 6 — Optional: Search schema tour (read-only) (5–10 minutes)

Only do this if the trainer asks.

Option A (tenant-level, trainer-led):
- SharePoint admin center → **More features** → **Search** → **Open** → **Manage Search Schema**.

Option B (site collection level):
- Site settings → **Site Collection Administration** → **Search Schema**.

Look for:
- Managed properties (for example, common properties like author/title)
- The idea of mapping crawled → managed properties

Do not change mappings in a shared training tenant unless explicitly assigned.

---

## End state (leave in place for later modules)
- Two seed documents stored in `NW-Pxx-Contracts/02-InReview`
- One published training bookmark (optional, if you had permissions)
- One draft/published training acronym (optional)

If the trainer requests cleanup:
- Set your bookmark to Draft or delete it.
- Set your acronym to Draft or delete it.

---

## Troubleshooting (common issues)

1) **My documents don’t appear in search.**
- Confirm you can find them by browsing the library (upload succeeded).
- Confirm library and site search visibility settings are enabled.
- Reindex the library only if you changed a search-related setting.
- Allow time for indexing.

2) **I uploaded content, but it still doesn’t show up.**
- Confirm you’re searching the correct scope (This library vs This site).
- If you’re testing pages/news, be aware draft items may not be crawled.

3) **I can’t access Bookmarks/Acronyms in the admin center.**
- You likely don’t have Search admin/editor permissions.
- Complete Exercises 1–2 and discuss Exercises 3–4 conceptually.

4) **My bookmark keyword triggers someone else’s result.**
- Use more specific keywords that include your `Pxx`.
- Avoid reserved keywords.

---

## References (official)
- Set up Microsoft Search (roles, admin center path): https://learn.microsoft.com/en-us/microsoftsearch/setup-microsoft-search
- Manage bookmarks: https://learn.microsoft.com/en-us/microsoftsearch/manage-bookmarks
- Manage acronyms: https://learn.microsoft.com/en-us/microsoftsearch/manage-acronyms
- Reindex site/library/list: https://learn.microsoft.com/en-us/sharepoint/crawl-site-content
- Keyword Query Language (KQL) syntax reference: https://learn.microsoft.com/en-us/sharepoint/dev/general-development/keyword-query-language-kql-syntax-reference
- Manage the search schema in SharePoint: https://learn.microsoft.com/en-us/sharepoint/manage-search-schema
