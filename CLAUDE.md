# CLAUDE.md

Microsoft Word add-in (web + native) for pentest report writing. Two commands, both
operating on finding sections in a security report:

1. **Sort findings by severity** — reorder the subsections of a user-chosen section
   Critical → High → Medium → Low → Informational (score descending within a severity).
2. **Insert rainbow findings table** — build a severity-coloured summary table of a
   chosen section's findings, with columns `#`, `Severity`, `Score`, `Title`, where `#`
   and `Title` are clickable cross-references to the finding.

3. **CVSS 3.1 calculator** — pick base metrics, write `Risk: <Severity> (<score>)` onto the
   Risk line at the cursor, and put the vector beneath it as a link to the NVD
   calculator.

4. **HTTP code block** — paste a Burp request or response into the pane and insert it at
   the cursor as a syntax-highlighted `Codeblock`.

`instructions.md` is the source of truth for requirements. Read it before changing behaviour.

## Design constraints

**KISS and YAGNI, aggressively.** This ships to a handful of colleagues via
OneDrive/SharePoint, not to AppSource. No telemetry, no settings UI, no plugin
architecture, no state management library. If a feature isn't in `instructions.md`,
don't build it.

## Stack

- **TypeScript, no framework.** The task pane is plain DOM: `taskpane.html` holds the
  markup and styles, and one module per panel wires it up. React and Fluent were removed
  once it was clear three panels of buttons did not need them. Do not reintroduce a
  framework without a reason the DOM cannot meet.
- **No polyfills.** `babel.config.json` targets `esmodules: true`. Requiring WordApi 1.4
  already restricts the add-in to Office 2021+/M365, whose webviews are Chromium or
  WKWebView, so core-js and regenerator were transpiling for browsers that cannot run this
  add-in anyway. Between dropping the framework and the polyfill the pane went from
  620 KB of JavaScript to 23 KB. Raising the babel target back to ES5 undoes that.
- **Office.js**, requirement set **WordApi 1.4** — the floor for the paragraph and OOXML
  APIs, plus `Range.insertBookmark` and `Document.getBookmarkRange`, which the findings
  table needs. This excludes Word 2019 and older; the add-in targets M365, where 1.4 has
  been available for years. It was raised from 1.3 deliberately: writing bookmarks through
  the API means the findings region is never rewritten just to add them, and it is what
  makes the inserted table removable again.
- **XML manifest** (not the unified JSON manifest).
- **Assets hosted on GitHub Pages** at `https://pontushanssen.github.io/rainbowtable/`,
  deployed by `.github/workflows/deploy.yml` on every push to master. Consequences:
  - `urlProd` in `webpack.config.js` is the production origin; a production build rewrites
    every `https://localhost:3000` in the manifest to it, with or without a trailing slash;
  - **the published site is world-readable**, even were the repo private. Nothing secret
    may enter `dist/` — no client names, no findings, no internal URLs;
  - filenames stay stable (no content hashes) so an asset update needs no manifest change;
  - nothing may be fetched from a CDN at runtime except Office.js itself, which must come
    from Microsoft.

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
- **Never move content with a series of separate `insertOoxml` calls.** Every seam where
  two inserted blocks meet merges: the tail of one block lands inside the next block's
  heading paragraph, which then renders as a numbered heading full of body text. This was
  found in live testing, and expanding each block to include its trailing paragraph mark
  did *not* fix it — the merge is at the join between inserts, not at the block boundary.
- **The shape that works**: capture the whole region with one `getOoxml()`, rearrange the
  paragraphs inside that package, and put it back with one
  `insertOoxml(..., "Replace")`. `sortFindings` and `restoreSection` both do this, and
  `src/word/ooxml.ts` holds the surgery. A single Replace has no seams, and it preserves
  each finding's XML verbatim — styles, numbering, tables and images included.
- `spanRange()` builds region spans that include the trailing paragraph mark, by expanding
  to the start of the following paragraph. `getRange("Whole")` alone stops short of the
  mark.
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
- **Report unreadable findings before editing anything.** Any command that modifies the
  document runs a read-only pass first (`previewSort` and its equivalents); if a finding
  in the selected section cannot be parsed, the taskpane names it and waits for the user
  to confirm. The user then either fixes the document or accepts that those findings stay
  put — but never discovers it after the edit has already happened.
- Sort order: Critical, High, Medium, Low, Informational; within a severity, higher score
  first; findings with equal or absent scores keep their original relative order (stable
  sort).

## Cross-references in the rainbow table

The `#` and `Title` cells must be real Word cross-references — clickable in Word **and
surviving PDF export as internal links**. That rules out plain text and rules out
`insertHyperlink` to a heading. It means bookmarks + `REF` fields:

- Bookmark each finding heading with `Range.insertBookmark`. **Reuse the bookmark already
  on the heading** (`existingBookmark()` over `Range.getBookmarks(true, false)`) before
  minting a new one — a name is derived from the title, so without reuse a rename would
  add a second bookmark to the same heading and leave the old one behind, one per title
  the finding has ever had.
- New names come from `bookmarkName()`: a hash of the section and finding titles, never
  the position, so re-running after a sort cannot repoint an older table's references at
  whatever now occupies that slot. The leading underscore makes them hidden; Word allows
  letters, digits and underscores only, up to 40 chars.
- Renaming a finding leaves an existing table's links working, but its cached Title text
  stays stale until fields update. Deleting a heading outright destroys its bookmark, and
  that table row then shows Word's "Error! Bookmark not defined.".
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
- **Ctrl+Z cannot undo what this add-in does.** Office.js edits do not enter Word's undo
  stack in Word on the web — confirmed by testing, not assumed. Every command that
  modifies the document therefore captures the OOXML of the region it is about to rewrite
  and returns it, and the task pane offers its own "Undo" that puts it back
  (`sortFindings` → `SortResult.snapshot` → `restoreSection`). A command without that is
  a command whose effects the user cannot reverse.
- The snapshot is held in task pane state only: it survives until the next action, a
  rescan, or the pane closing. That is deliberate — it is an undo, not a version history.
- Do not leave the document half-modified if a step throws.

## Layout

```
manifest.xml              add-in manifest; localhost URLs must become SharePoint URLs to ship
src/taskpane/taskpane.html      markup and styles for the whole pane
src/taskpane/taskpane.ts        entry point: Office.onReady, tab switching
src/taskpane/dom.ts             byId/show/make plus the per-panel feedback lines
src/taskpane/findingsPanel.ts   section picker, sorting, findings table
src/taskpane/cvssPanel.ts       CVSS metric buttons and the live score
src/taskpane/httpPanel.ts       the Burp paste box
src/word/headings.ts      heading scan and the relative section/finding model
src/word/severity.ts      strict `Risk:` parsing and the sort comparator
src/word/ooxml.ts         reordering paragraphs inside a captured OOXML package
src/word/sortFindings.ts  feature 1 — previewSort, sortFindings, restoreSection (working)
src/word/section.ts       scanning a section into findings, shared by both features
src/word/cvss.ts          CVSS 3.1 base score arithmetic, pure
src/word/insertRisk.ts    feature 3 — insertRisk, undoRisk
src/word/newFinding.ts    insertFinding, removeFinding — the empty finding skeleton
src/word/http.ts          HTTP message tokenising, pure
src/word/httpBlock.ts     feature 4 — insertHttpBlock, removeHttpBlock
src/word/findingsTable.ts feature 2 — previewTable, insertFindingsTable, removeFindingsTable
```

Word-facing logic lives in `src/word/` and never touches the DOM; panels call into it.
That split is what keeps the whole test suite runnable in node, and it is why replacing
the UI framework did not disturb a single test.

## Commands

| | |
|---|---|
| `npm test` | compile `src/word` + `test` and run the node:test suite |
| `npm run dev-server` | webpack dev server on https://localhost:3000 |
| `npm run build` | production bundle into `dist/` |
| `npm run validate` | validate `manifest.xml` |
| `npm run lint` / `lint:fix` | office-addin-lint |
| `npx tsc --noEmit` | typecheck — **the build uses babel-loader and does not typecheck**, so run this explicitly |

`npm start` (office-addin-debugging) sideloads into Word **desktop** and does not work on
Linux. Develop with `npm run dev-server` plus manual sideloading in Word on the web:
**Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in**, pointing at
`manifest.xml`. The document has to be open from OneDrive/SharePoint.

HTTPS certificates: `npx office-addin-dev-certs install` writes them to
`~/.office-addin-dev-certs`, but its CA install step fails on Fedora (it targets Arch's
trust path) — the certificates themselves are still generated, and `webpack.config.js`
reads them from disk. The browser must trust `ca.crt` or the task pane iframe will not
load; for Chromium that is
`certutil -d sql:$HOME/.pki/nssdb -A -t "C,," -n "Office Add-in Dev CA" -i ~/.office-addin-dev-certs/ca.crt`.
Firefox keeps its own store and needs a separate import.

## Working in this repo

- Everything runs inside `Word.run(async context => …)`. Batch property loads; call
  `context.sync()` as few times as the logic allows — sync count is the main performance
  lever in Office.js.
- Test against `Security Review Template V3.1.docx` (kept locally, not in the repo, since
  `*.docx` is ignored). Changes to sorting or parsing must be checked against it before
  being called done.
- `npm test` covers the pure logic only — parsing, ordering, block boundaries — using
  node:test with no extra dependencies. Keep Word-facing code thin enough that the part
  worth testing stays pure. **Nothing that touches Office.js is covered**, so the OOXML
  round trip and range handling still need a real document.
- `test/fixtures/template.ts` mirrors the real template's heading outline. Note it repeats
  the title "Vulnerabilities" at two depths, which is why sections are identified by
  paragraph index and level rather than by title.
- Verify in **Word on the web** as well as the desktop client; OOXML, bookmark and list
  APIs behave differently between them.
- Any change touching the table or sorting needs a **PDF export check**: cross-references
  that work in Word can still export as dead text.

## The CVSS calculator

- `src/word/cvss.ts` implements the v3.1 base score formulas verbatim, including the
  specification's integer-arithmetic `roundUp`. The tests check it against published CVE
  vectors (9.8 BlueKeep, 10.0 Zerologon, 7.5 Heartbleed, 6.1 reflected XSS); if you touch
  the arithmetic, those are the guard rail.
- CVSS's 0.0 band is called "None"; it is mapped to **Informational**, the severity the
  rest of the add-in and the reports use.
- The point of the feature is that what it writes is exactly what `severity.ts` parses.
  It rewrites the whole `Risk:` line rather than inserting at the cursor, so the result
  cannot come out malformed. The cursor must be on a `Risk:` line; anything else is an
  error rather than a guess.
- The vector links to NVD's calculator, which takes `?vector=…&version=3.1`. Keep it a
  query parameter, not a fragment: Word splits a hyperlink on `#` into address and
  location. The report template already cites this calculator, so the link matches how
  the reports are written.
- This feature deliberately uses **no OOXML**: `Paragraph.insertText` preserves each
  paragraph's own style, and `Range.hyperlink` (WordApi 1.3) links the vector without
  needing a relationship part, which a hand-built package cannot carry.

## The HTTP code block

- **Input comes through the task pane only**, never from a selection in the document.
  Text pasted into Word gets autocorrect, smart quotes and autocapitalisation applied, and
  a request in a report is evidence — the bytes have to be the ones Burp produced.
- **Never reformat the message.** No pretty-printing, no re-indenting, no re-encoding;
  only colour is added. `test/http.test.ts` reassembles the tokens and asserts they equal
  the input for every fixture, which is the invariant to preserve.
- Detection degrades instead of failing: an unparseable body is still highlighted as far
  as it goes (truncated JSON is normal in reports), and text that is not HTTP at all
  becomes a plain, uncoloured code block.
- Output is one `Codeblock` paragraph per line. That style has `contextualSpacing` and
  four-sided borders, so consecutive paragraphs merge into a single continuous box —
  which is why lines are paragraphs rather than `w:br` breaks.
- **A package must define every style it names.** `insertOoxml` discards a `w:pStyle`
  pointing at a style the package does not define, so the first version of this arrived as
  ordinary body text at body spacing. `wrapInPackage(body, styles)` adds the styles part
  and its relationship; where the document already defines the style, that definition wins
  and ours is ignored. Applies to any future OOXML that names a style.
- Line spacing is stated directly on each paragraph as well as in the style. Code set at
  the document's body spacing reads as a list of stray lines rather than a block.
- Colours are picked for the template's beige `D7D2CB` shading rather than white, and
  bold carries the emphasis on the start line and header names so the block still reads
  in greyscale print.
- Every run repeats Courier New 10pt. That duplicates what `Codeblock` already sets, so it
  costs nothing in this template, but keeps the block monospaced in a document that has no
  such style.

## Deploying

`npm run build` produces `dist/`, which is the whole site: task pane, commands page,
icons, and a `manifest.xml` rewritten to the production origin. Pushing to master runs
`.github/workflows/deploy.yml`, which tests, builds and publishes `dist/` to GitHub Pages.

To install the add-in in Word, the manifest from `dist/manifest.xml` has to reach users:
either through **Microsoft 365 admin centre → Integrated apps** (assigns it to people
automatically, needs an admin), or the tenant **SharePoint App Catalog**, where it appears
under Add-ins → My Organization for each person to add once.

Bump `<Version>` in `manifest.xml` whenever the manifest itself changes — new URLs, ribbon
entries, or a raised requirement set — or Word will keep the copy it already has. Asset-only
changes need no version bump and no redeployment of the manifest.

## Icons

`assets/icon-{16,32,64,80,128}.png` are generated from `logo.png` (a rainbow over a picnic
table — the repo's name). Two things to know before regenerating them:

- **`logo.png` has a checkerboard painted into it.** It looks transparent but every pixel
  is opaque: the grey-and-white squares are real. They must be keyed out, and because the
  sticker's interior is white too, a colour key is not enough — flood fill inwards from the
  four corners, which stops at the blue outline. PIL's `floodfill` threshold compares the
  *sum* across channels, so it needs to exceed ~117 (grey to white) and stay under ~275
  (grey to the blue outline); 150 works.
- The full sticker is kept at every size rather than cropping tighter for the small ones.
  The white disc is what keeps the icon legible against a dark ribbon, and one crop for all
  sizes means the icon does not change shape as display scaling changes.

## Inserting a new finding

`insertFinding` writes an empty finding at the cursor: title, `Risk: [TODO]`, Status,
Description, Technical details, Proof of concept, Recommendation, with `[TODO]` bodies.

- Heading levels come from the **section chosen in the task pane**, not from whatever the
  cursor happens to sit on. A finding belongs to a section, and reading the level off the
  selection guesses wrong the moment the cursor is on a body paragraph.
- **It must work with no section selected.** `findSections` only reports a section that
  already contains findings, so the first finding in a section exists before the section is
  detectable — requiring a selection made the button do nothing precisely when it was most
  needed. With nothing selected it falls back to `DEFAULT_FINDING_LEVEL`, the template's
  depth, and says so in the status line.
- It uses `insertParagraph` and `styleBuiltIn`, not OOXML. Setting the built-in style picks
  up the document's own heading styles including their numbering, where a hand-built
  package would have to define those styles and risk overriding the template's.
- `Risk: [TODO]` deliberately does not parse, so a half-written finding shows up in the
  sort and table pre-flight warnings until it is scored.

## A note on the `hidden` class

`.hidden` in `taskpane.html` carries `!important`. Component rules like `.confirm` also set
`display`, and at equal specificity the later rule wins — which once left the confirmation
block on screen with an empty button. Any new utility class that hides things needs the
same treatment, or it must come last.

## Never fail silently

A task pane button that returns early without saying why is indistinguishable from a
broken one — this happened once already, when all three findings buttons returned quietly
if no section was selected. Every action either does something or reports why it cannot.
Where an action genuinely needs a selection, its button is disabled (`updateAvailability`
in `findingsPanel.ts`), because a disabled button explains itself and a dead one does not.
Note that `guard()` re-enables every button in its `finally`, so availability has to be
reapplied afterwards; `act()` exists to do that.
