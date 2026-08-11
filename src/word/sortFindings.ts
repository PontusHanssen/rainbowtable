import { Heading, Section, childHeadings, toHeadings } from "./headings";
import { Risk, compareRisk, isRiskHeading, parseRisk } from "./severity";

/* global Word, OfficeExtension */

export interface SkippedFinding {
  title: string;
  reason: string;
}

export interface SortPreview {
  /** True when sorting would actually move something. */
  changed: boolean;
  /** How many findings would take part in the sort. */
  sorted: number;
  /** Findings whose risk rating could not be read; they stay where they are. */
  skipped: SkippedFinding[];
}

export interface SortResult extends SortPreview {
  /** OOXML of the findings as they were before the sort; feed to `restoreSection`. */
  snapshot?: string;
}

/** A finding and the span of paragraphs it owns, inclusive. */
export interface Block {
  heading: Heading;
  start: number;
  end: number;
  risk?: Risk;
  skipReason?: string;
}

interface Scan {
  paragraphs: Word.ParagraphCollection;
  sectionHeading: Heading;
  blocks: Block[];
}

/**
 * The order the findings should end up in.
 *
 * Findings whose risk could not be read keep their slot: the sortable ones are
 * rearranged among the positions they already occupy, and everything else stays put.
 * Exported for testing.
 */
export function planOrder<T extends { risk?: Risk }>(findings: T[]): T[] {
  const sortable = findings.filter((finding) => finding.risk !== undefined);
  const ordered = [...sortable].sort((a, b) => compareRisk(a.risk as Risk, b.risk as Risk));

  let next = 0;
  return findings.map((finding) => (finding.risk === undefined ? finding : ordered[next++]));
}

/**
 * What sorting this section would do, without touching the document.
 *
 * Always run this first: findings with an unreadable risk rating must be shown to the
 * user before anything is edited, so they can fix the document instead of discovering
 * afterwards that some findings were silently left behind.
 */
export async function previewSort(section: Section): Promise<SortPreview> {
  return Word.run(async (context) => {
    const { blocks } = await scanSection(context, section);
    return summarize(blocks);
  });
}

/**
 * A range covering paragraphs `start` through `end` *including the trailing paragraph
 * mark*.
 *
 * `getRange("Whole")` stops just short of a paragraph's mark, so OOXML captured from it
 * ends mid-paragraph: re-inserting two such blocks back to back merges the tail of one
 * into the heading of the next. Expanding to the start of the following paragraph is
 * what pulls the mark in. Only a span reaching the last paragraph of the document has
 * no following paragraph to expand to, and nothing follows it to merge with.
 */
function spanRange(paragraphs: Word.ParagraphCollection, start: number, end: number): Word.Range {
  const from = paragraphs.items[start].getRange("Whole");
  const following = end + 1;

  return following < paragraphs.items.length
    ? from.expandTo(paragraphs.items[following].getRange("Start"))
    : from.expandTo(paragraphs.items[end].getRange("Whole"));
}

/**
 * Reorder a section's findings by severity, moving each finding's whole block of
 * content with it.
 *
 * Returns a snapshot of the region as it was before the edit. Office.js edits do not
 * enter Word's own undo stack on the web, so Ctrl+Z cannot reach them — `restoreSection`
 * plus this snapshot is the add-in's own undo.
 */
export async function sortFindings(section: Section): Promise<SortResult> {
  return Word.run(async (context) => {
    const { paragraphs, sectionHeading, blocks } = await scanSection(context, section);
    const summary = summarize(blocks);
    if (!summary.changed) {
      return summary;
    }

    // Capture each block's OOXML before touching the document; a round trip through
    // OOXML is what preserves tables, images and code blocks inside a finding.
    const ranges = new Map<Block, Word.Range>();
    const ooxml = new Map<Block, OfficeExtension.ClientResult<string>>();
    blocks.forEach((block) => {
      const range = spanRange(paragraphs, block.start, block.end);
      ranges.set(block, range);
      ooxml.set(block, range.getOoxml());
    });
    const snapshot = spanRange(
      paragraphs,
      blocks[0].start,
      blocks[blocks.length - 1].end
    ).getOoxml();
    await context.sync();

    // Remove every block, then re-insert them immediately after the section heading in
    // reverse order: each insert pushes the previous one down, so they land in order.
    const anchor = paragraphs.items[sectionHeading.index].getRange("Whole");
    blocks.forEach((block) => (ranges.get(block) as Word.Range).delete());
    [...planOrder(blocks)]
      .reverse()
      .forEach((block) =>
        anchor.insertOoxml(
          (ooxml.get(block) as OfficeExtension.ClientResult<string>).value,
          Word.InsertLocation.after
        )
      );
    await context.sync();

    // getOoxml() hands back a ClientResult, which is filled in by the sync above and
    // needs no load() — the office-addins lint rule cannot tell the two apart.
    // eslint-disable-next-line office-addins/load-object-before-read
    return { ...summary, snapshot: snapshot.value };
  });
}

/**
 * Put a section's findings back as they were, from a snapshot taken by `sortFindings`.
 *
 * The snapshot covers the findings only, not the section heading, so this replaces the
 * same span the sort rewrote. It assumes the section has not been edited since.
 */
export async function restoreSection(section: Section, snapshot: string): Promise<void> {
  return Word.run(async (context) => {
    const { paragraphs, blocks } = await scanSection(context, section);
    if (blocks.length === 0) {
      throw new Error(`"${section.heading.text}" no longer has any findings to restore.`);
    }

    const region = spanRange(paragraphs, blocks[0].start, blocks[blocks.length - 1].end);
    region.insertOoxml(snapshot, Word.InsertLocation.replace);
    await context.sync();
  });
}

/** Read the document and work out what each finding of the section covers. */
async function scanSection(context: Word.RequestContext, section: Section): Promise<Scan> {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items/styleBuiltIn,items/text");
  await context.sync();

  const headings = toHeadings(paragraphs.items);
  const position = locateSection(headings, section);
  const findings = childHeadings(headings, position);

  return {
    paragraphs,
    sectionHeading: headings[position],
    blocks: buildBlocks(headings, position, findings, paragraphs.items.length),
  };
}

function summarize(blocks: Block[]): SortPreview {
  const skipped = blocks
    .filter((block) => block.risk === undefined)
    .map((block) => ({ title: block.heading.text, reason: block.skipReason as string }));
  const order = planOrder(blocks);

  return {
    changed: order.some((block, i) => block !== blocks[i]),
    sorted: blocks.length - skipped.length,
    skipped,
  };
}

/**
 * Re-find the section in a freshly scanned document: the user may have edited since the
 * task pane last scanned, which invalidates paragraph indexes.
 */
function locateSection(headings: Heading[], section: Section): number {
  const exact = headings.findIndex(
    (heading) => heading.index === section.heading.index && heading.text === section.heading.text
  );
  if (exact >= 0) {
    return exact;
  }

  const moved = headings.findIndex(
    (heading) => heading.text === section.heading.text && heading.level === section.heading.level
  );
  if (moved >= 0) {
    return moved;
  }

  throw new Error(
    `Could not find the section "${section.heading.text}". Rescan the document and try again.`
  );
}

/**
 * The paragraph span each finding owns: from its heading up to the next finding, or to
 * the end of the section for the last one. Exported for testing.
 */
export function buildBlocks(
  headings: Heading[],
  position: number,
  findings: Heading[],
  paragraphCount: number
): Block[] {
  const sectionHeading = headings[position];
  const after = headings.find(
    (heading, i) => i > position && heading.level <= sectionHeading.level
  );
  const sectionEnd = after ? after.index : paragraphCount;

  return findings.map((finding, i) => {
    const block: Block = {
      heading: finding,
      start: finding.index,
      end: (i + 1 < findings.length ? findings[i + 1].index : sectionEnd) - 1,
    };

    const rating = readRisk(headings, finding);
    if ("risk" in rating) {
      block.risk = rating.risk;
    } else {
      block.skipReason = rating.reason;
    }
    return block;
  });
}

/** The risk of a finding, read from its own child headings. */
function readRisk(headings: Heading[], finding: Heading): { risk: Risk } | { reason: string } {
  const children = childHeadings(headings, headings.indexOf(finding));

  for (const child of children) {
    const risk = parseRisk(child.text);
    if (risk) {
      return { risk };
    }
  }

  const malformed = children.find((child) => isRiskHeading(child.text));
  return {
    reason: malformed
      ? `"${malformed.text}" is not of the form "Risk: <Severity> (<score>)"`
      : `no "Risk:" heading`,
  };
}
