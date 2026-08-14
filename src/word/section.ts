import { Heading, Section, childHeadings, toHeadings } from "./headings";
import { Risk, Severity, isRiskHeading, parseRisk } from "./severity";

/* global Word */

export interface SkippedFinding {
  title: string;
  reason: string;
}

/** A finding and the span of paragraphs it owns, inclusive. */
export interface Block {
  heading: Heading;
  start: number;
  end: number;
  risk?: Risk;
  skipReason?: string;
}

export interface Scan {
  paragraphs: Word.ParagraphCollection;
  headings: Heading[];
  sectionHeading: Heading;
  blocks: Block[];
}

/** Read the document and work out what each finding of the section covers. */
export async function scanSection(context: Word.RequestContext, section: Section): Promise<Scan> {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items/styleBuiltIn,items/text");
  await context.sync();

  const headings = toHeadings(paragraphs.items);
  const position = locateSection(headings, section);
  const findings = childHeadings(headings, position);

  return {
    paragraphs,
    headings,
    sectionHeading: headings[position],
    blocks: buildBlocks(headings, position, findings, paragraphs.items.length),
  };
}

/**
 * A range covering paragraphs `start` through `end` *including the trailing paragraph
 * mark*, which `getRange("Whole")` stops short of. Expanding to the start of the
 * following paragraph is what pulls the mark in.
 */
export function spanRange(
  paragraphs: Word.ParagraphCollection,
  start: number,
  end: number
): Word.Range {
  const from = paragraphs.items[start].getRange("Whole");
  const following = end + 1;

  return following < paragraphs.items.length
    ? from.expandTo(paragraphs.items[following].getRange("Start"))
    : from.expandTo(paragraphs.items[end].getRange("Whole"));
}

/**
 * A finding as the task pane lists it back to the user, before anything is written.
 *
 * The pane cannot show `Block`s: they carry paragraph indexes and Word ranges, and the
 * point of the preview is to be readable.
 */
export interface FindingSummary {
  title: string;
  /** Absent when the risk heading could not be read; `reason` then says why. */
  severity?: Severity;
  score?: number;
  reason?: string;
}

/**
 * Findings as they would be listed, in the order given.
 *
 * Every findings action already computes this and throws it away — sorting to decide
 * whether anything moved, tabulating to count the rows — so showing the user what is
 * about to happen costs no extra round trip to Word.
 */
export function findingSummaries(blocks: Block[]): FindingSummary[] {
  return blocks.map((block) => ({
    title: block.heading.text,
    severity: block.risk?.severity,
    score: block.risk?.score,
    reason: block.risk ? undefined : block.skipReason,
  }));
}

/** The findings whose risk rating could not be read. */
export function skippedFindings(blocks: Block[]): SkippedFinding[] {
  return blocks
    .filter((block) => block.risk === undefined)
    .map((block) => ({ title: block.heading.text, reason: block.skipReason as string }));
}

/**
 * Re-find the section in a freshly scanned document: the user may have edited since the
 * task pane last scanned, which invalidates paragraph indexes.
 */
export function locateSection(headings: Heading[], section: Section): number {
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
