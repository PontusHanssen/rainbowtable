import { Section } from "./headings";

/* global Word */

export interface NewFinding {
  /** The heading level the finding's title was written at. */
  level: number;
  paragraphs: number;
  /** Bookmark wrapping the inserted finding, so it can be removed again. */
  bookmark: string;
}

const NEW_FINDING_BOOKMARK = "_ptfinding";

/** The style names Word accepts, indexed by level; there is no Heading10. */
const HEADING_STYLES = [
  "Heading1",
  "Heading2",
  "Heading3",
  "Heading4",
  "Heading5",
  "Heading6",
  "Heading7",
  "Heading8",
  "Heading9",
] as const;

type BuiltInStyle = (typeof HEADING_STYLES)[number] | "Normal";

/** A finding needs a level for its own sections, so it cannot start deeper than 8. */
const DEEPEST_FINDING = HEADING_STYLES.length - 1;

/**
 * The skeleton of a finding, mirroring the report template: a title, the risk rating the
 * CVSS tab fills in, then the prose sections. `level` is the finding's own heading level;
 * its sections sit one below.
 */
function skeleton(level: number): { style: BuiltInStyle; text: string }[] {
  const heading = HEADING_STYLES[level - 1];
  const section = HEADING_STYLES[level];

  return [
    { style: heading, text: "[TODO] Finding title" },
    { style: section, text: "Risk: [TODO]" },
    { style: "Normal", text: "[TODO] Score this from the CVSS 3.1 tab." },
    { style: section, text: "Status: [TODO]" },
    { style: section, text: "Description" },
    { style: "Normal", text: "[TODO]" },
    { style: section, text: "Technical details" },
    { style: "Normal", text: "[TODO]" },
    { style: section, text: "Proof of concept" },
    { style: "Normal", text: "[TODO]" },
    { style: section, text: "Recommendation" },
    { style: "Normal", text: "[TODO]" },
  ];
}

/**
 * Insert an empty finding at the cursor, ready to fill in.
 *
 * The heading levels come from the section chosen in the task pane rather than from
 * whatever surrounds the cursor: a finding belongs to a section, and reading the level off
 * the selection would guess wrong whenever the cursor sat on a body paragraph.
 *
 * Built from `insertParagraph` rather than OOXML on purpose. Setting `styleBuiltIn` picks
 * up the document's own heading styles, numbering included, where a hand-built package
 * would have to define those styles and would risk overriding them.
 */
export async function insertFinding(section: Section): Promise<NewFinding> {
  const level = Math.min(section.heading.level + 1, DEEPEST_FINDING);
  const blocks = skeleton(level);

  return Word.run(async (context) => {
    const cursor = context.document.getSelection().paragraphs.getFirst();

    let previous: Word.Paragraph = cursor;
    const inserted: Word.Paragraph[] = [];
    blocks.forEach((block) => {
      previous = previous.insertParagraph(block.text, Word.InsertLocation.after);
      previous.styleBuiltIn = block.style;
      inserted.push(previous);
    });

    const span = inserted[0]
      .getRange("Whole")
      .expandTo(inserted[inserted.length - 1].getRange("Whole"));
    span.insertBookmark(NEW_FINDING_BOOKMARK);
    await context.sync();

    return { level, paragraphs: blocks.length, bookmark: NEW_FINDING_BOOKMARK };
  });
}

/** Remove a finding inserted by `insertFinding`. */
export async function removeFinding(bookmark: string): Promise<void> {
  return Word.run(async (context) => {
    const range = context.document.getBookmarkRangeOrNullObject(bookmark);
    range.load("text");
    await context.sync();

    if (range.isNullObject) {
      throw new Error("That finding is no longer in the document.");
    }

    range.delete();
    context.document.deleteBookmark(bookmark);
    await context.sync();
  });
}
