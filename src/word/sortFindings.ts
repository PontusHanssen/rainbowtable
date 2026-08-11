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

/** A finding and the span of paragraphs it owns, inclusive. */
interface Block {
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
 * Reorder a section's findings by severity, moving each finding's whole block of
 * content with it.
 *
 * The edit is a single Word.run batch — one round trip that deletes the old blocks and
 * re-inserts them in order — so it lands as one undo step rather than leaving the
 * document half-sorted.
 */
export async function sortFindings(section: Section): Promise<SortPreview> {
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
      const range = paragraphs.items[block.start]
        .getRange("Whole")
        .expandTo(paragraphs.items[block.end].getRange("Whole"));
      ranges.set(block, range);
      ooxml.set(block, range.getOoxml());
    });
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

    return summary;
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
 * the end of the section for the last one.
 */
function buildBlocks(
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
