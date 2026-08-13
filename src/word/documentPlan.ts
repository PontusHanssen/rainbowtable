import { Block, Inline } from "./markdown";

/**
 * What to write into the document, as data.
 *
 * `insertOoxml` costs about six seconds per call whatever it carries, so content is built
 * through the Office.js API instead — see the measurements in CLAUDE.md. That would make
 * rendering untestable, since it becomes a sequence of Word calls rather than a string, so
 * the decisions live here as a plain plan and the executor in `markdownDoc.ts` does as it
 * is told.
 */

export interface RunPlan {
  text: string;
  bold?: boolean;
  italic?: boolean;
  /** Inline code, which takes the template's character style. */
  code?: boolean;
  /** Makes the run a hyperlink to this address. */
  link?: string;
}

export type ParagraphPlan =
  | { kind: "heading"; level: number; runs: RunPlan[] }
  | { kind: "body"; runs: RunPlan[] }
  | { kind: "code"; text: string }
  | { kind: "listItem"; ordered: boolean; group: number; runs: RunPlan[] };

/** Markdown goes to six #s, and Word has nine heading levels. */
const DEEPEST_HEADING = 6;

function runsFrom(spans: Inline[]): RunPlan[] {
  return spans.map((span) => {
    switch (span.kind) {
      case "bold":
        return { text: span.text, bold: true };
      case "italic":
        return { text: span.text, italic: true };
      case "code":
        return { text: span.text, code: true };
      case "link":
        return { text: span.text, link: span.url };
      default:
        return { text: span.text };
    }
  });
}

/**
 * Turn parsed markdown into paragraphs to write.
 *
 * Consecutive list items of the same kind share a `group`, because Word numbers a list by
 * attaching its paragraphs to one list object: without the grouping, every item would
 * start again at 1.
 */
export function planFromBlocks(blocks: Block[]): ParagraphPlan[] {
  const plans: ParagraphPlan[] = [];
  let group = 0;
  let previous: "bullet" | "number" | undefined;

  for (const block of blocks) {
    if (block.kind === "bullet" || block.kind === "number") {
      if (block.kind !== previous) {
        group += 1;
      }
      previous = block.kind;
      plans.push({
        kind: "listItem",
        ordered: block.kind === "number",
        group,
        runs: runsFrom(block.spans),
      });
      continue;
    }

    previous = undefined;

    switch (block.kind) {
      case "heading":
        plans.push({
          kind: "heading",
          level: Math.min(block.level, DEEPEST_HEADING),
          runs: runsFrom(block.spans),
        });
        break;
      case "code":
        // One paragraph per line: the code style's contextual spacing and matching
        // borders are what join them into a single box.
        block.lines.forEach((line) => plans.push({ kind: "code", text: line }));
        break;
      default:
        plans.push({ kind: "body", runs: runsFrom(block.spans) });
    }
  }

  return plans;
}

/** The contiguous list runs in a plan, as `group` to the items belonging to it. */
export function listGroups(plans: ParagraphPlan[]): Map<number, number[]> {
  const groups = new Map<number, number[]>();

  plans.forEach((plan, index) => {
    if (plan.kind === "listItem") {
      const members = groups.get(plan.group) ?? [];
      members.push(index);
      groups.set(plan.group, members);
    }
  });

  return groups;
}
