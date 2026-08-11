import { Section } from "./headings";
import { reorderFindings } from "./ooxml";
import { Block, SkippedFinding, scanSection, skippedFindings, spanRange } from "./section";
import { Risk, compareRisk } from "./severity";

/* global Word */

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

    // Capture the whole findings region in one piece. Reordering happens inside this
    // package and goes back as a single Replace: inserting the findings one at a time
    // merges every seam, putting the tail of one finding into the next one's heading.
    const region = spanRange(paragraphs, blocks[0].start, blocks[blocks.length - 1].end);
    const captured = region.getOoxml();
    await context.sync();

    // getOoxml() hands back a ClientResult, which is filled in by the sync above and
    // needs no load() — the office-addins lint rule cannot tell the two apart.
    // eslint-disable-next-line office-addins/load-object-before-read
    const snapshot = captured.value;
    const order = planOrder(blocks).map((block) => blocks.indexOf(block));

    region.insertOoxml(
      reorderFindings(snapshot, sectionHeading.level + 1, order),
      Word.InsertLocation.replace
    );
    await context.sync();

    return { ...summary, snapshot };
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

function summarize(blocks: Block[]): SortPreview {
  const skipped = skippedFindings(blocks);
  const order = planOrder(blocks);

  return {
    changed: order.some((block, i) => block !== blocks[i]),
    sorted: blocks.length - skipped.length,
    skipped,
  };
}
