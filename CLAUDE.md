# CLAUDE.md

Microsoft Word add-in (web + native) for pentest report writing. Two commands, both
operating on finding sections in a security report:

1. **Sort findings by severity** — reorder the subsections of a user-chosen section
   Critical → High → Medium → Low → Informational (score descending within a severity).
2. **Insert rainbow findings table** — build a severity-coloured summary table of a
   chosen section's findings, with columns `#`, `Severity`, `Score`, `Title`, where `#`
   and `Title` are clickable cross-references to the finding.

`instructions.md` is the source of truth for requirements. Read it before changing behaviour.

## Design constraints

**KISS and YAGNI, aggressively.** This ships to a handful of colleagues via
OneDrive/SharePoint, not to AppSource. No telemetry, no settings UI, no plugin
architecture, no state management library. If a feature isn't in `instructions.md`,
don't build it.

## Stack

- **TypeScript + React** taskpane, scaffolded from the Yeoman `office` generator
  (`yo office`, projectType `react`, host `word`, TypeScript).
- **Office.js**, requirement set **WordApi 1.3** — the floor for the paragraph and OOXML
  APIs this add-in needs, and widely available in Word on the web and desktop. Bookmark
  APIs want 1.4; the OOXML path covers 1.3 clients, so don't raise the manifest floor
  without a reason.
- **XML manifest** (not the unified JSON manifest), sideloaded from a shared
  SharePoint/OneDrive folder.
- **Assets hosted in SharePoint.** Consequences for the build:
  - output must be plain static files with **relative paths** and **stable filenames** —
    no content-hashed bundles, since deployment is a manual upload over existing files;
  - `SourceLocation` and `AppDomains` in the manifest point at the SharePoint library URL;
  - nothing may be fetched from a CDN at runtime — bundle everything.

## The document model

The structure is **relative, not tied to absolute heading levels.** The user selects a
section at some heading level `N`; its findings are the headings at level `N+1` directly
beneath it, and each finding's severity lives in a heading at level `N+2`. Different
reports place findings at different depths — never hardcode "findings are Heading2".

```
Heading(N)    <selected section>
Heading(N+1)    <finding title>          <- sorted / tabulated
Heading(N+2)      Risk: <Sev> (<score>)  <- severity lives here
Heading(N+2)      Status: ...
Heading(N+2)      Description
Heading(N+2)      Recommendation
```

As it appears in `Security Review Template V3.1.docx` (the real template this targets),
with `N = 1`:

```
Heading1  Vulnerabilities
Heading2    <finding title>
Heading3      Risk: [TODO]
Heading3      Status / Description / Technical details / PoC / Recommendation
Heading1  Weaknesses               <- second findings section, same shape
```

Facts that constrain the implementation:

- The template uses **built-in** heading styles. Detect them with
  `Paragraph.styleBuiltIn` (`Word.Style.heading1`…), **never** `Paragraph.style` — the
  document's language is `sv-SE` and `style` returns the localized display name.
- **Heading1 and Heading2 are auto-numbered** by a style-linked list (`numId 23`);
  Heading3 is not. Never write a literal section number into the document, and never
  reorder in a way that assumes numbers are text — Word renumbers on its own. Note that
  at other depths the numbered/unnumbered split falls elsewhere, which is another reason
  the `#` column is a field rather than text (see below).
- A finding is its heading paragraph **plus everything up to the next heading at level
  `N+1` or shallower**, including
  tables, code blocks (`Codeblock` style) and images. Reordering must move the whole
  block. Prefer round-tripping via `getOoxml()` / `insertOoxml()` over rebuilding content
  paragraph by paragraph — it preserves formatting the naive path destroys.
- The template also defines custom **unnumbered** heading variants (`Heading_1 No` …
  `Heading_5 No`) with matching outline levels. They are unused in the template body but
  may appear in real reports; they are *not* built-in, so they will not be detected.

## Severity parsing — strict

The severity heading must match `Risk: <Severity> (<score>)`, where `<Severity>` is one of
Informational, Low, Medium, High, Critical, and the `(<score>)` part is **optional**.

- Matching is case-insensitive on the severity word; everything else is literal.
- A finding whose Risk heading is missing, unparseable, or still `[TODO]` is **not
  guessed at**. Leave it in place, and report it to the user in the taskpane as
  unprocessed, naming the finding. Silent mis-parses are worse than a visible skip.
- Sort order: Critical, High, Medium, Low, Informational; within a severity, higher score
  first; findings with equal or absent scores keep their original relative order (stable
  sort).

## Cross-references in the rainbow table

The `#` and `Title` cells must be real Word cross-references — clickable in Word **and
surviving PDF export as internal links**. That rules out plain text and rules out
`insertHyperlink` to a heading. It means bookmarks + `REF` fields:

- Bookmark each finding heading (a hidden `_Ref`-style name of our own, e.g.
  `_pt_finding_<n>`; do not reuse or renumber Word's own `_Ref…` bookmarks).
  `Range.insertBookmark` needs **WordApi 1.4**; where that isn't available, emit
  `w:bookmarkStart`/`w:bookmarkEnd` in the OOXML round-trip we already perform.
- `#` cell → `REF <bookmark> \w \h`. `\w` is full-context paragraph numbering, which yields
  `3.1` from anywhere in the document; `\r` (relative context) can collapse to `1` when the
  table sits in a different branch. `\h` makes it a hyperlink — **without `\h` the PDF
  export has nothing to link.**
- `Title` cell → `REF <bookmark> \h`.
- Emit fields with a **cached result** — a `w:fldSimple` (or `fldChar` run pair) whose
  inner run holds the computed text. A field with an empty result renders blank until
  someone presses F9, and PDF export does not necessarily update fields first. Compute the
  cached number by counting heading ordinals along the path to the finding; Office.js does
  not reliably expose the computed string for style-linked numbering, so
  `paragraph.listItemOrNullObject.listString` is a hint, not a source of truth.
- Sorting a section **after** a table exists must leave the fields intact: bookmarks travel
  with the OOXML of the moved block, so the references keep pointing at the right findings
  and the numbers correct themselves on update. Verify this — it is the most likely thing
  to break, and it must be checked in the PDF, not just on screen.

## Behaviour details

- The user picks a section from a list of the document's headings, populated by scanning
  the document — do not hardcode "Vulnerabilities"/"Weaknesses", and do not restrict the
  list to Heading1.
- The rainbow table is inserted at the current selection, with each row's severity cell
  shaded by severity colour.
- Both commands are single undo-able user actions as far as is practical; do not leave the
  document half-modified if a step throws.

## Layout

```
manifest.xml              add-in manifest; localhost URLs must become SharePoint URLs to ship
src/taskpane/             React entry point + taskpane shell
src/taskpane/components/App.tsx   section picker and the two command buttons
src/word/headings.ts      heading scan and the relative section/finding model
src/word/sortFindings.ts       feature 1 — not implemented
src/word/findingsTable.ts      feature 2 — not implemented
```

Word-facing logic lives in `src/word/` and stays free of React; components call into it.

## Commands

| | |
|---|---|
| `npm run dev-server` | webpack dev server on https://localhost:3000 |
| `npm run build` | production bundle into `dist/` |
| `npm run validate` | validate `manifest.xml` |
| `npm run lint` / `lint:fix` | office-addin-lint |
| `npx tsc --noEmit` | typecheck — **the build uses babel-loader and does not typecheck**, so run this explicitly |

`npm start` (office-addin-debugging) sideloads into Word **desktop** and does not work on
Linux. Develop with `npm run dev-server` plus manual sideloading in Word on the web.

## Working in this repo

- Everything runs inside `Word.run(async context => …)`. Batch property loads; call
  `context.sync()` as few times as the logic allows — sync count is the main performance
  lever in Office.js.
- Test against `Security Review Template V3.1.docx` in the repo root. Changes to sorting or
  parsing must be checked against it before being called done.
- Verify in **Word on the web** as well as the desktop client; OOXML, bookmark and list
  APIs behave differently between them.
- Any change touching the table or sorting needs a **PDF export check**: cross-references
  that work in Word can still export as dead text.
