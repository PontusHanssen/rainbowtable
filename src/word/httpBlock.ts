import { HttpMessage, Token, TokenKind, highlightHttp } from "./http";
import { escapeXml, wrapInPackage } from "./ooxml";

/* global Word */

/**
 * Colours are chosen to read on the template's Codeblock shading (D7D2CB, a warm grey),
 * not on white, and to stay legible in greyscale print — hue alone never carries meaning.
 */
const COLOURS: Record<TokenKind, { colour: string; bold?: boolean }> = {
  method: { colour: "0A3069", bold: true },
  target: { colour: "24292F" },
  version: { colour: "4C4C4C" },
  status2xx: { colour: "1A7F37", bold: true },
  status3xx: { colour: "0A3069", bold: true },
  status4xx: { colour: "9A3412", bold: true },
  status5xx: { colour: "A40E26", bold: true },
  reason: { colour: "4C4C4C" },
  headerName: { colour: "5A2D8C", bold: true },
  punctuation: { colour: "4C4C4C" },
  headerValue: { colour: "24292F" },
  key: { colour: "0A3069" },
  string: { colour: "032F62" },
  number: { colour: "953800" },
  literal: { colour: "A40E26" },
  tag: { colour: "1A7F37" },
  attrName: { colour: "5A2D8C" },
  attrValue: { colour: "032F62" },
  text: { colour: "24292F" },
};

/** The paragraph style the report template defines for code. */
export const CODE_STYLE = "Codeblock";

/**
 * Courier New at 10pt, matching what the template's Codeblock style sets.
 *
 * Repeated on every run on purpose: a document without that style would otherwise render
 * the message in a proportional font, which makes an HTTP message much harder to read.
 */
const FONT =
  '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/><w:sz w:val="20"/><w:szCs w:val="20"/>';

function tokenRun(token: Token): string {
  const { colour, bold } = COLOURS[token.kind];
  const properties = `${FONT}${bold ? "<w:b/>" : ""}<w:color w:val="${colour}"/>`;

  return `<w:r><w:rPr>${properties}</w:rPr><w:t xml:space="preserve">${escapeXml(token.text)}</w:t></w:r>`;
}

/**
 * Single-spaced with nothing above or below, stated directly rather than left to the
 * style: a code block set at the document's body spacing reads as a list of stray lines.
 * `contextualSpacing` is what lets consecutive lines merge into one continuous box.
 */
const CODE_SPACING =
  '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:contextualSpacing/>';

function codeParagraph(tokens: Token[]): string {
  return (
    `<w:p><w:pPr><w:pStyle w:val="${CODE_STYLE}"/>${CODE_SPACING}</w:pPr>` +
    tokens.map(tokenRun).join("") +
    "</w:p>"
  );
}

/**
 * A definition of the Codeblock style, shipped with the block.
 *
 * A `w:pStyle` naming a style the package does not define is discarded, which left the
 * message as ordinary body text. Where the document already defines Codeblock — as the
 * report template does — that definition wins and this one is ignored; elsewhere it
 * supplies the same look.
 */
const CODE_STYLE_DEFINITION =
  `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${CODE_STYLE}">` +
  '<w:name w:val="Code block"/><w:basedOn w:val="Normal"/><w:qFormat/>' +
  "<w:pPr>" +
  '<w:pBdr><w:top w:val="single" w:sz="24" w:space="1" w:color="D7D2CB"/>' +
  '<w:left w:val="single" w:sz="24" w:space="4" w:color="D7D2CB"/>' +
  '<w:bottom w:val="single" w:sz="24" w:space="1" w:color="D7D2CB"/>' +
  '<w:right w:val="single" w:sz="24" w:space="4" w:color="D7D2CB"/></w:pBdr>' +
  '<w:shd w:val="clear" w:color="auto" w:fill="D7D2CB"/>' +
  CODE_SPACING +
  "</w:pPr>" +
  `<w:rPr>${FONT}</w:rPr></w:style>`;

/** The highlighted message as an OOXML package, ready to insert. Exported for testing. */
export function buildHttpBlock(message: HttpMessage): string {
  // The trailing paragraph keeps the block from fusing with whatever follows the cursor,
  // and gives somewhere to carry on typing after it.
  return wrapInPackage(message.lines.map(codeParagraph).join("") + "<w:p/>", CODE_STYLE_DEFINITION);
}

export interface HttpBlockResult {
  kind: HttpMessage["kind"];
  body: HttpMessage["body"];
  lines: number;
  /** Bookmark wrapping the inserted block, so it can be removed again. */
  bookmark: string;
}

const BLOCK_BOOKMARK = "_pthttp";

/** Insert a highlighted HTTP message at the cursor. */
export async function insertHttpBlock(raw: string): Promise<HttpBlockResult> {
  if (!raw.trim()) {
    throw new Error("Paste an HTTP request or response first.");
  }

  const message = highlightHttp(raw);

  return Word.run(async (context) => {
    const inserted = context.document
      .getSelection()
      .insertOoxml(buildHttpBlock(message), Word.InsertLocation.replace);
    inserted.insertBookmark(BLOCK_BOOKMARK);
    await context.sync();

    return {
      kind: message.kind,
      body: message.body,
      lines: message.lines.length,
      bookmark: BLOCK_BOOKMARK,
    };
  });
}

/** Remove a block inserted by `insertHttpBlock`. */
export async function removeHttpBlock(bookmark: string): Promise<void> {
  return Word.run(async (context) => {
    const range = context.document.getBookmarkRangeOrNullObject(bookmark);
    range.load("text");
    await context.sync();

    if (range.isNullObject) {
      throw new Error("That block is no longer in the document.");
    }

    range.delete();
    context.document.deleteBookmark(bookmark);
    await context.sync();
  });
}
