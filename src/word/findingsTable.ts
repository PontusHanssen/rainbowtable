import { Heading, Section, headingNumber } from "./headings";
import { run, wrapInPackage } from "./ooxml";
import {
  Block,
  FindingSummary,
  SkippedFinding,
  findingSummaries,
  scanSection,
  skippedFindings,
} from "./section";
import { Severity } from "./severity";

/* global Word */

export interface FindingRow {
  /** Bookmark on the finding's heading; the # and Title cells reference it. */
  bookmark: string;
  /** Cached field result, e.g. "4.2". Word recomputes it when fields update. */
  number: string;
  severity?: Severity;
  score?: number;
  title: string;
}

export interface TablePreview {
  rows: number;
  /** Findings with no readable risk rating. They are still listed, without a severity. */
  skipped: SkippedFinding[];
  /** The rows the table would carry, in document order, for the pane to show first. */
  findings: FindingSummary[];
}

export interface TableResult extends TablePreview {
  /** Bookmark wrapping the inserted table, so it can be removed again. */
  bookmark: string;
}

/** The report template's table style, which paints its own header row. */
export const TABLE_STYLE = "Omegapointtabellbl";

/**
 * The template's own character styles for severities — they colour the text rather than
 * shading the cell, which is the house style. The template calls the lowest band "Info",
 * where the rest of the add-in says "Informational"; the style id is what matters here.
 */
const SEVERITY_STYLES: Record<Severity, string> = {
  Critical: "Critical",
  High: "High",
  Medium: "Medium",
  Low: "Low",
  Informational: "Info",
};

/**
 * Fallbacks, used only where the document does not define these itself. A package must
 * define every style it names or Word discards the reference, but where the report
 * template is in play its own definitions win and these are ignored.
 */
const STYLE_DEFINITIONS =
  `<w:style w:type="table" w:customStyle="1" w:styleId="${TABLE_STYLE}">` +
  '<w:name w:val="Omegapoint tabell (blå)"/><w:uiPriority w:val="99"/>' +
  "<w:tblPr><w:tblBorders>" +
  ["top", "left", "bottom", "right", "insideH", "insideV"]
    .map((edge) => `<w:${edge} w:val="single" w:sz="4" w:space="0" w:color="D7D2CB"/>`)
    .join("") +
  "</w:tblBorders></w:tblPr>" +
  '<w:tblStylePr w:type="firstRow"><w:rPr><w:b w:val="0"/><w:color w:val="FFC000"/></w:rPr>' +
  '<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="003349"/></w:tcPr></w:tblStylePr>' +
  "</w:style>" +
  (
    [
      ["Critical", "A50021"],
      ["High", "FF0000"],
      ["Medium", "FFC000"],
      ["Low", "00B050"],
      ["Info", "00B0F0"],
    ] as const
  )
    .map(
      ([id, colour]) =>
        `<w:style w:type="character" w:customStyle="1" w:styleId="${id}">` +
        `<w:name w:val="${id}"/><w:uiPriority w:val="1"/><w:qFormat/>` +
        `<w:rPr><w:color w:val="${colour}"/></w:rPr></w:style>`
    )
    .join("");

const COLUMNS = [900, 1600, 900, 5960];

/** 32-bit FNV-1a, used to give each finding a stable bookmark name. */
function hash(text: string): string {
  let value = 0x811c9dc5;

  for (let i = 0; i < text.length; i++) {
    value ^= text.charCodeAt(i);
    value = Math.imul(value, 0x01000193) >>> 0;
  }
  return value.toString(16).padStart(8, "0");
}

/**
 * A bookmark name for a finding.
 *
 * Derived from the titles rather than from position, so re-running after a sort reuses
 * the same name for the same finding instead of repointing an older table's references
 * at whatever now sits in that slot. The leading underscore hides it from Word's bookmark
 * list; Word allows letters, digits and underscores only, up to 40 characters.
 */
export function bookmarkName(sectionTitle: string, findingTitle: string, occurrence = 0): string {
  return `_ptf${hash(`${sectionTitle}|${findingTitle}`)}${occurrence}`;
}

/** The bookmark wrapping a section's inserted table. */
export function tableBookmarkName(sectionTitle: string): string {
  return `_ptt${hash(sectionTitle)}`;
}

/**
 * A REF field with its result cached inside.
 *
 * `\h` is what makes it a hyperlink, and without it PDF export has nothing to link. `\w`
 * gives the full-context paragraph number ("4.2" rather than "2"). The cached run matters
 * because a field with an empty result renders blank until someone presses F9, and
 * exporting to PDF does not necessarily update fields first.
 */
function referenceField(
  bookmark: string,
  switches: string,
  cached: string,
  properties = ""
): string {
  return `<w:fldSimple w:instr=" REF ${bookmark} ${switches} ">${run(cached, properties)}</w:fldSimple>`;
}

function cell(width: number, content: string, fill?: string): string {
  const shading = fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>` : "";
  return (
    `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${shading}<w:vAlign w:val="center"/></w:tcPr>` +
    `<w:p><w:pPr><w:spacing w:before="20" w:after="20"/></w:pPr>${content}</w:p></w:tc>`
  );
}

function headerRow(): string {
  // No shading or bold here: the table style's firstRow formatting supplies both, and
  // hard-coding them would override the template's look.
  const cells = ["#", "Severity", "Score", "Title"]
    .map((label, i) => cell(COLUMNS[i], run(label)))
    .join("");

  return `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cells}</w:tr>`;
}

function findingRow(row: FindingRow): string {
  const severityRun = row.severity
    ? run(row.severity, `<w:rStyle w:val="${SEVERITY_STYLES[row.severity]}"/>`)
    : run("—");

  return (
    "<w:tr>" +
    cell(COLUMNS[0], referenceField(row.bookmark, "\\w \\h", row.number)) +
    cell(COLUMNS[1], severityRun) +
    cell(COLUMNS[2], run(row.score === undefined ? "" : row.score.toFixed(1))) +
    cell(COLUMNS[3], referenceField(row.bookmark, "\\h", row.title)) +
    "</w:tr>"
  );
}

/** The findings table as an OOXML package, ready to insert. Exported for testing. */
export function buildFindingsTable(rows: FindingRow[]): string {
  const grid = COLUMNS.map((width) => `<w:gridCol w:w="${width}"/>`).join("");

  // tblLook is what switches the style's conditional formatting on; without firstRow set,
  // the header row is painted like any other and the style's navy banner never appears.
  const table =
    "<w:tbl>" +
    "<w:tblPr>" +
    `<w:tblStyle w:val="${TABLE_STYLE}"/>` +
    '<w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/>' +
    '<w:tblLook w:val="0420" w:firstRow="1" w:lastRow="0" w:firstColumn="0"' +
    ' w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>' +
    "</w:tblPr>" +
    `<w:tblGrid>${grid}</w:tblGrid>` +
    headerRow() +
    rows.map(findingRow).join("") +
    "</w:tbl>";

  // The trailing paragraph keeps the table from fusing with whatever follows it.
  return wrapInPackage(`${table}<w:p/>`, STYLE_DEFINITIONS);
}

/** Our bookmarks on a heading, if any. Names are hidden, so ask for hidden ones. */
export const FINDING_BOOKMARK_PREFIX = "_ptf";

/**
 * Exactly what `bookmarkName` produces, and nothing else.
 *
 * A reused name is document input, not ours: `getBookmarks` returns whatever the file
 * carries, and the name goes straight into a field instruction — `w:instr=" REF <name>
 * \w \h "` — where a space adds switches and a quote escapes the attribute. Word's own
 * UI cannot create such a name, but the file format allows one and a report may have
 * passed through other tools. Anything that does not look like ours is ignored, which
 * costs only a fresh bookmark alongside it.
 */
const FINDING_BOOKMARK = /^_ptf[0-9a-f]{8}\d+$/;

/** The bookmark this add-in already put on a heading, if there is one. */
export function existingBookmark(names: string[]): string | undefined {
  return [...names].filter((name) => FINDING_BOOKMARK.test(name)).sort()[0];
}

/**
 * Turn a section's findings into table rows, in the order they appear in the document.
 *
 * `existing[i]` is the bookmark already on that finding's heading. Reusing it is what
 * makes renaming a finding safe: the name is derived from the title, so a rename would
 * otherwise mint a second bookmark on the same heading and leave the old one behind.
 */
export function buildRows(
  headings: Heading[],
  sectionTitle: string,
  blocks: Block[],
  existing: (string | undefined)[] = []
): FindingRow[] {
  const seen = new Map<string, number>();

  return blocks.map((block, i) => {
    const occurrence = seen.get(block.heading.text) ?? 0;
    seen.set(block.heading.text, occurrence + 1);

    return {
      bookmark: existing[i] ?? bookmarkName(sectionTitle, block.heading.text, occurrence),
      number: headingNumber(headings, headings.indexOf(block.heading)),
      severity: block.risk?.severity,
      score: block.risk?.score,
      title: block.heading.text,
    };
  });
}

/** What the table would contain, without touching the document. */
export async function previewTable(section: Section): Promise<TablePreview> {
  return Word.run(async (context) => {
    const { blocks } = await scanSection(context, section);
    return {
      rows: blocks.length,
      skipped: skippedFindings(blocks),
      findings: findingSummaries(blocks),
    };
  });
}

/**
 * Insert a severity-coloured findings table at the current selection.
 *
 * Every finding gets a row, including ones whose risk could not be read — leaving a
 * finding out of a summary table would misrepresent the report. Those rows carry no
 * severity and no colour.
 */
export async function insertFindingsTable(section: Section): Promise<TableResult> {
  return Word.run(async (context) => {
    const { paragraphs, headings, blocks } = await scanSection(context, section);
    if (blocks.length === 0) {
      throw new Error(`"${section.heading.text}" has no findings to tabulate.`);
    }

    // Reuse the bookmark already on each heading where there is one, so a finding keeps
    // its identity across renames instead of collecting a bookmark per title it has had.
    const found = blocks.map((block) =>
      paragraphs.items[block.start].getRange("Whole").getBookmarks(true, false)
    );
    await context.sync();

    const rows = buildRows(
      headings,
      section.heading.text,
      blocks,
      found.map((names) => existingBookmark(names.value))
    );
    rows.forEach((row, i) =>
      paragraphs.items[blocks[i].start].getRange("Whole").insertBookmark(row.bookmark)
    );

    const bookmark = tableBookmarkName(section.heading.text);
    const inserted = context.document
      .getSelection()
      .insertOoxml(buildFindingsTable(rows), Word.InsertLocation.replace);
    inserted.insertBookmark(bookmark);
    await context.sync();

    return {
      rows: rows.length,
      skipped: skippedFindings(blocks),
      findings: findingSummaries(blocks),
      bookmark,
    };
  });
}

/** Remove a table inserted by `insertFindingsTable`. */
export async function removeFindingsTable(bookmark: string): Promise<void> {
  return Word.run(async (context) => {
    const range = context.document.getBookmarkRangeOrNullObject(bookmark);
    range.load("text");
    await context.sync();

    if (range.isNullObject) {
      throw new Error("That table is no longer in the document.");
    }

    range.delete();
    context.document.deleteBookmark(bookmark);
    await context.sync();
  });
}
