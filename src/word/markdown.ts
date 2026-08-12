/**
 * A small markdown parser, covering exactly what a finding needs: headings, paragraphs,
 * bold, italic, inline code, fenced code blocks, bullet and numbered lists, and autolinks
 * written `<https://example.com>`.
 *
 * Deliberately not CommonMark. Tables, images, blockquotes, reference links and nested
 * lists are out of scope; anything unrecognised stays literal text rather than being
 * silently dropped.
 */

export type Inline =
  | { kind: "text"; text: string }
  | { kind: "bold"; text: string }
  | { kind: "italic"; text: string }
  | { kind: "code"; text: string }
  | { kind: "link"; text: string; url: string };

export type Block =
  | { kind: "heading"; level: number; spans: Inline[] }
  | { kind: "paragraph"; spans: Inline[] }
  | { kind: "bullet"; spans: Inline[] }
  | { kind: "number"; spans: Inline[] }
  | { kind: "code"; lines: string[]; language?: string };

/**
 * One pass over the line, trying each construct at each position.
 *
 * Order matters: code first, so backticked text is never re-read as emphasis; then
 * autolinks; then bold before italic, or `**bold**` would be taken as two italics.
 */
const INLINE = /`([^`]+)`|<((?:https?:\/\/|mailto:)[^>\s]+)>|\*\*([^*]+)\*\*|\*([^*\n]+)\*/g;

export function parseInline(text: string): Inline[] {
  const spans: Inline[] = [];
  let last = 0;
  let match: RegExpExecArray | null;

  INLINE.lastIndex = 0;
  while ((match = INLINE.exec(text)) !== null) {
    if (match.index > last) {
      spans.push({ kind: "text", text: text.slice(last, match.index) });
    }

    if (match[1] !== undefined) {
      spans.push({ kind: "code", text: match[1] });
    } else if (match[2] !== undefined) {
      spans.push({ kind: "link", text: match[2], url: match[2] });
    } else if (match[3] !== undefined) {
      spans.push({ kind: "bold", text: match[3] });
    } else {
      spans.push({ kind: "italic", text: match[4] });
    }

    last = match.index + match[0].length;
  }

  if (last < text.length) {
    spans.push({ kind: "text", text: text.slice(last) });
  }
  return spans;
}

const FENCE = /^\s*```\s*(\w+)?\s*$/;
const HEADING = /^(#{1,6})\s+(.*)$/;
const BULLET = /^\s*[-*+]\s+(.*)$/;
const NUMBERED = /^\s*\d+[.)]\s+(.*)$/;

/** Parse markdown into the blocks the document builder knows how to write. */
export function parseMarkdown(source: string): Block[] {
  const lines = source.replace(/\r\n?/g, "\n").split("\n");
  const blocks: Block[] = [];
  let paragraph: string[] = [];

  // Consecutive lines are one paragraph, as in markdown: a single newline is a soft wrap.
  const flush = () => {
    if (paragraph.length > 0) {
      blocks.push({ kind: "paragraph", spans: parseInline(paragraph.join(" ")) });
      paragraph = [];
    }
  };

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    const fence = FENCE.exec(line);
    if (fence) {
      flush();
      const body: string[] = [];
      i += 1;
      while (i < lines.length && !FENCE.test(lines[i])) {
        body.push(lines[i]);
        i += 1;
      }
      blocks.push({ kind: "code", lines: body, language: fence[1] });
      continue;
    }

    if (line.trim() === "") {
      flush();
      continue;
    }

    const heading = HEADING.exec(line);
    if (heading) {
      flush();
      blocks.push({
        kind: "heading",
        level: heading[1].length,
        spans: parseInline(heading[2].trim()),
      });
      continue;
    }

    const bullet = BULLET.exec(line);
    if (bullet) {
      flush();
      blocks.push({ kind: "bullet", spans: parseInline(bullet[1]) });
      continue;
    }

    const numbered = NUMBERED.exec(line);
    if (numbered) {
      flush();
      blocks.push({ kind: "number", spans: parseInline(numbered[1]) });
      continue;
    }

    paragraph.push(line.trim());
  }

  flush();
  return blocks;
}

export interface Outline {
  headings: { level: number; text: string }[];
  paragraphs: number;
  listItems: number;
  codeBlocks: number;
  links: number;
}

/** The plain text of a run of spans, for previews and headings. */
export function spanText(spans: Inline[]): string {
  return spans.map((span) => span.text).join("");
}

/** What the pane shows before anything is written to the document. */
export function outline(blocks: Block[]): Outline {
  const spansOf = (block: Block) => ("spans" in block ? block.spans : []);

  return {
    headings: blocks
      .filter((block): block is Extract<Block, { kind: "heading" }> => block.kind === "heading")
      .map((block) => ({ level: block.level, text: spanText(block.spans) })),
    paragraphs: blocks.filter((block) => block.kind === "paragraph").length,
    listItems: blocks.filter((block) => block.kind === "bullet" || block.kind === "number").length,
    codeBlocks: blocks.filter((block) => block.kind === "code").length,
    links: blocks.reduce(
      (total, block) => total + spansOf(block).filter((span) => span.kind === "link").length,
      0
    ),
  };
}
