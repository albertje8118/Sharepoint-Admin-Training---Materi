# Module 6 — Search in SharePoint Online and Microsoft Search (Slides)

## Slide 1 — Title
- Module 6: Search in SharePoint Online and Microsoft Search
- Scenario: Project Northwind Intranet Modernization

## Slide 2 — Why admins care about search
- Information discovery = productivity
- Search is only as good as:
  - permissions,
  - content quality,
  - metadata,
  - and governance.

## Slide 3 — Security trimming (non-negotiable)
- Search never “overrides” permissions
- If a user can’t see a file, search won’t show it

## Slide 4 — Entry points (user experience)
- SharePoint search box
- Microsoft 365 search experiences
- (Concept) multiple entry points, consistent governance goals

## Slide 5 — What admins can tune (big buckets)
- Content quality (IA, metadata, naming)
- Search visibility settings (site/library)
- Curated “answers” (Bookmarks, Acronyms)
- (Advanced) Search schema (managed properties)

## Slide 6 — Microsoft Search admin surface
- Microsoft 365 admin center → Settings → Search & intelligence
- Roles: Search admin, Search editor

## Slide 7 — Bookmarks
- What they are: curated links triggered by keywords
- Publishing: visible immediately after publish
- Shared-tenant rule: training-only `NW-Pxx` keywords

## Slide 8 — Acronyms
- Admin-curated vs system-curated
- Draft vs Published
- Published acronyms can take up to a day to appear

## Slide 9 — Q&As (status note)
- Historically part of “answers”
- Availability can vary; design labs around Bookmarks/Acronyms

## Slide 10 — Crawled vs managed properties (overview)
- Crawled: discovered during crawl
- Managed: used in index; queryable/retrievable settings
- Changes often require reindex

## Slide 11 — Reindexing (when and why)
- Useful after search visibility/schema changes
- Caution: avoid unnecessary reindexing (load)

## Slide 12 — Query basics (admin troubleshooting)
- Phrase queries with quotes
- AND/OR/NOT (uppercase)
- File type / author / title patterns (when supported)

## Slide 13 — Lab overview
- Upload seed docs
- Validate search results
- Reindex library if needed
- Create Bookmark (immediate)
- Create Acronym (may take time)

## Slide 14 — End-of-module checkpoint
- Seed docs discoverable
- You can explain indexing + reindexing
- You can curate at least one safe “answer” (Bookmark)

## Slide 15 — What’s next
- Module 7: Apps and customization (governance + deployment basics)
