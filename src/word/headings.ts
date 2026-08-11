/* global Word */

/**
 * A heading paragraph in the document.
 *
 * `index` is the position in `body.paragraphs`, which is how every other module
 * addresses the paragraph afterwards. It is only valid until the document is
 * modified — re-scan after any edit.
 */
export interface Heading {
  index: number;
  /** 1-9, taken from the built-in Heading1..Heading9 styles. */
  level: number;
  text: string;
}

/**
 * A section the user can act on: a heading whose children look like findings.
 * The findings are the direct child headings at `heading.level + 1`.
 */
export interface Section {
  heading: Heading;
  findings: Heading[];
}

const HEADING_STYLE = /^Heading([1-9])$/;

/**
 * The heading level of a paragraph style, or 0 if it is not a heading.
 *
 * Only built-in heading styles are recognised. `styleBuiltIn` is locale-independent;
 * `style` is not, and the report template is a Swedish-language document.
 */
function headingLevel(styleBuiltIn: Word.Style | string): number {
  const match = HEADING_STYLE.exec(String(styleBuiltIn));
  return match ? Number(match[1]) : 0;
}

/** The headings among a loaded paragraph collection. Requires styleBuiltIn and text. */
export function toHeadings(
  paragraphs: { styleBuiltIn: Word.Style | string; text: string }[]
): Heading[] {
  const headings: Heading[] = [];
  paragraphs.forEach((paragraph, index) => {
    const level = headingLevel(paragraph.styleBuiltIn);
    if (level > 0) {
      headings.push({ index, level, text: paragraph.text.trim() });
    }
  });
  return headings;
}

/** Every heading in the document, in document order. */
export async function getHeadings(): Promise<Heading[]> {
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/styleBuiltIn,items/text");
    await context.sync();
    return toHeadings(paragraphs.items);
  });
}

/**
 * The dotted number Word displays for a heading, e.g. "4.2", worked out by counting
 * heading ordinals.
 *
 * Office.js does not reliably expose the computed string for style-linked numbering, and
 * this only has to be close enough: it is the cached result baked into a REF field, which
 * Word recomputes when fields update. Headings styled with something other than the
 * built-in Heading styles are invisible here, so a document that mixes numbering schemes
 * can drift until the fields refresh.
 */
export function headingNumber(headings: Heading[], position: number): string {
  if (position < 0 || position >= headings.length) {
    throw new Error(`No heading at position ${position}.`);
  }

  const ordinals: number[] = [];

  for (let i = 0; i <= position; i++) {
    const level = headings[i].level;
    ordinals.length = level;
    ordinals[level - 1] = (ordinals[level - 1] ?? 0) + 1;
  }

  return Array.from(ordinals, (ordinal) => ordinal ?? 1).join(".");
}

/** The direct child headings of `headings[position]`, i.e. its findings. */
export function childHeadings(headings: Heading[], position: number): Heading[] {
  const parent = headings[position];
  const children: Heading[] = [];

  for (let i = position + 1; i < headings.length; i++) {
    const heading = headings[i];
    if (heading.level <= parent.level) {
      break;
    }
    if (heading.level === parent.level + 1) {
      children.push(heading);
    }
  }
  return children;
}

/**
 * Sections that have the shape the add-in operates on: a heading with child
 * headings that themselves have children (where `Risk:` lives). Findings sit at
 * different depths in different reports, so this is derived, never hardcoded.
 */
export function findSections(headings: Heading[]): Section[] {
  const sections: Section[] = [];

  headings.forEach((heading, position) => {
    const findings = childHeadings(headings, position);
    const hasGrandchildren = findings.some(
      (finding) => childHeadings(headings, headings.indexOf(finding)).length > 0
    );
    if (findings.length > 0 && hasGrandchildren) {
      sections.push({ heading, findings });
    }
  });
  return sections;
}
