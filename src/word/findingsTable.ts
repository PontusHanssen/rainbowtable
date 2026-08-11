import { Heading, Section, headingNumber } from "./headings";
import { Block, SkippedFinding, scanSection, skippedFindings } from "./section";
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
}

export interface TableResult extends TablePreview {
  /** Bookmark wrapping the inserted table, so it can be removed again. */
  bookmark: string;
}

/** Fill and text colour per severity — the "rainbow" part. */
const SEVERITY_COLOURS: Record<Severity, { fill: string; text: string }> = {
  Critical: { fill: "C00000", text: "FFFFFF" },
  High: { fill: "E36C0A", text: "FFFFFF" },
  Medium: { fill: "FFC000", text: "000000" },
  Low: { fill: "92D050", text: "000000" },
  Informational: { fill: "BFBFBF", text: "000000" },
};

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

function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function run(text: string, properties = ""): string {
  const rPr = properties ? `<w:rPr>${properties}</w:rPr>` : "";
  return `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
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
  const cells = ["#", "Severity", "Score", "Title"]
    .map((label, i) => cell(COLUMNS[i], run(label, "<w:b/>"), "D9D9D9"))
    .join("");

  return `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cells}</w:tr>`;
}

function findingRow(row: FindingRow): string {
  const colours = row.severity ? SEVERITY_COLOURS[row.severity] : undefined;
  const severityRun = colours
    ? run(row.severity as string, `<w:b/><w:color w:val="${colours.text}"/>`)
    : run("—");

  return (
    "<w:tr>" +
    cell(COLUMNS[0], referenceField(row.bookmark, "\\w \\h", row.number)) +
    cell(COLUMNS[1], severityRun, colours?.fill) +
    cell(COLUMNS[2], run(row.score === undefined ? "" : row.score.toFixed(1))) +
    cell(COLUMNS[3], referenceField(row.bookmark, "\\h", row.title)) +
    "</w:tr>"
  );
}

/** Wrap document body content in the minimal flat OPC package insertOoxml expects. */
function wrapInPackage(body: string): string {
  return (
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
    '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">' +
    '<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/word/document.xml"/>' +
    "</Relationships></pkg:xmlData></pkg:part>" +
    '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">' +
    '<pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${body}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`
  );
}

/** The findings table as an OOXML package, ready to insert. Exported for testing. */
export function buildFindingsTable(rows: FindingRow[]): string {
  const border = (edge: string) =>
    `<w:${edge} w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>`;
  const borders = ["top", "left", "bottom", "right", "insideH", "insideV"].map(border).join("");
  const grid = COLUMNS.map((width) => `<w:gridCol w:w="${width}"/>`).join("");

  const table =
    "<w:tbl>" +
    `<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${borders}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>` +
    `<w:tblGrid>${grid}</w:tblGrid>` +
    headerRow() +
    rows.map(findingRow).join("") +
    "</w:tbl>";

  // The trailing paragraph keeps the table from fusing with whatever follows it.
  return wrapInPackage(`${table}<w:p/>`);
}

/** Our bookmarks on a heading, if any. Names are hidden, so ask for hidden ones. */
export const FINDING_BOOKMARK_PREFIX = "_ptf";

/** The bookmark this add-in already put on a heading, if there is one. */
export function existingBookmark(names: string[]): string | undefined {
  return [...names]
    .filter((name) => name.toLowerCase().startsWith(FINDING_BOOKMARK_PREFIX))
    .sort()[0];
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
    return { rows: blocks.length, skipped: skippedFindings(blocks) };
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

    return { rows: rows.length, skipped: skippedFindings(blocks), bookmark };
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
