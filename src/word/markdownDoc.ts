import { Block, Inline, parseMarkdown } from "./markdown";
import { escapeXml, run, wrapInPackage } from "./ooxml";

/* global Word */

/** The template's styles. Shipped as fallbacks too, since a package must define what it names. */
const CODE_BLOCK_STYLE = "Codeblock";
const INLINE_CODE_STYLE = "Codeinline20";
const LIST_STYLE = "ListParagraph";

/** High ids, to stay clear of the numbering the destination document already uses. */
const BULLET_NUM = 880;
const NUMBER_NUM = 881;

/** Markdown goes to six #s, which is as deep as this needs to style. */
const DEEPEST_HEADING = 6;

export interface MarkdownInsertion {
  blocks: number;
  /** Bookmark wrapping what was written, which is how it is removed again. */
  bookmark: string;
}

const MARKDOWN_BOOKMARK = "_ptmd";

function inlineRun(span: Inline): string {
  switch (span.kind) {
    case "bold":
      return run(span.text, "<w:b/>");
    case "italic":
      return run(span.text, "<w:i/>");
    case "code":
      return run(span.text, `<w:rStyle w:val="${INLINE_CODE_STYLE}"/>`);
    case "link":
      // A HYPERLINK field needs no relationship part, unlike a w:hyperlink element.
      return (
        `<w:fldSimple w:instr=" HYPERLINK &quot;${escapeXml(span.url)}&quot; ">` +
        run(span.text, '<w:rStyle w:val="Hyperlink"/>') +
        "</w:fldSimple>"
      );
    default:
      return run(span.text);
  }
}

function paragraph(spans: Inline[], properties = ""): string {
  const pPr = properties ? `<w:pPr>${properties}</w:pPr>` : "";
  return `<w:p>${pPr}${spans.map(inlineRun).join("")}</w:p>`;
}

function listProperties(numId: number): string {
  return (
    `<w:pStyle w:val="${LIST_STYLE}"/>` +
    `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>` +
    "<w:contextualSpacing/>"
  );
}

/** Render one markdown block. `#` is Heading1, `##` Heading2, and so on. */
function blockXml(block: Block): string {
  switch (block.kind) {
    case "heading": {
      const level = Math.min(block.level, DEEPEST_HEADING);
      return paragraph(block.spans, `<w:pStyle w:val="Heading${level}"/>`);
    }
    case "bullet":
      return paragraph(block.spans, listProperties(BULLET_NUM));
    case "number":
      return paragraph(block.spans, listProperties(NUMBER_NUM));
    case "code":
      // One paragraph per line, as in the HTTP block: the style's contextual spacing and
      // matching borders are what merge them into a single box.
      return block.lines
        .map((line) =>
          paragraph([{ kind: "text", text: line }], `<w:pStyle w:val="${CODE_BLOCK_STYLE}"/>`)
        )
        .join("");
    default:
      return paragraph(block.spans);
  }
}

/**
 * Word's built-in headings, declared so `w:pStyle` references to them resolve.
 *
 * A package must define every style it names — without these the headings arrived as
 * ordinary body text. `w:name` has to be the built-in name ("heading 2"), which is what
 * ties these to Word's own heading styles; a document that defines them, as the report
 * template does, still wins.
 */
const HEADING_DEFINITIONS = Array.from(
  { length: DEEPEST_HEADING },
  (_unused, index) =>
    `<w:style w:type="paragraph" w:styleId="Heading${index + 1}">` +
    `<w:name w:val="heading ${index + 1}"/><w:basedOn w:val="Normal"/>` +
    `<w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>` +
    `<w:pPr><w:outlineLvl w:val="${index}"/></w:pPr>` +
    "<w:rPr><w:b/></w:rPr></w:style>"
).join("");

const STYLE_DEFINITIONS =
  HEADING_DEFINITIONS +
  `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${CODE_BLOCK_STYLE}">` +
  '<w:name w:val="Code block"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:shd w:val="clear" w:color="auto" w:fill="D7D2CB"/>' +
  '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr>' +
  '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="20"/></w:rPr></w:style>' +
  `<w:style w:type="character" w:customStyle="1" w:styleId="${INLINE_CODE_STYLE}">` +
  '<w:name w:val="Code inline 2.0"/>' +
  '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="18"/>' +
  '<w:shd w:val="clear" w:color="auto" w:fill="EDEAE7"/></w:rPr></w:style>' +
  `<w:style w:type="paragraph" w:styleId="${LIST_STYLE}">` +
  '<w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>' +
  '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>' +
  '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>';

/** Bullets use Symbol's F0B7, the character Word itself uses for a round bullet. */
const NUMBERING_DEFINITIONS =
  `<w:abstractNum w:abstractNumId="${BULLET_NUM}"><w:multiLevelType w:val="hybridMultilevel"/>` +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/>' +
  '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>' +
  '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>' +
  `<w:abstractNum w:abstractNumId="${NUMBER_NUM}"><w:multiLevelType w:val="hybridMultilevel"/>` +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
  '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
  `<w:num w:numId="${BULLET_NUM}"><w:abstractNumId w:val="${BULLET_NUM}"/></w:num>` +
  `<w:num w:numId="${NUMBER_NUM}"><w:abstractNumId w:val="${NUMBER_NUM}"/></w:num>`;

/** The markdown as an OOXML package, ready to insert. Exported for testing. */
export function buildMarkdown(blocks: Block[]): string {
  const body = blocks.map(blockXml).join("");

  // The trailing paragraph keeps the last block from fusing with what follows the cursor.
  return wrapInPackage(`${body}<w:p/>`, STYLE_DEFINITIONS, NUMBERING_DEFINITIONS);
}

/**
 * Insert markdown at the cursor.
 *
 * Headings map straight across: `#` is Heading1, `##` is Heading2. Nothing is inferred
 * from elsewhere in the pane, so what is written is what appears.
 */
export async function insertMarkdown(source: string): Promise<MarkdownInsertion> {
  const blocks = parseMarkdown(source);
  if (blocks.length === 0) {
    throw new Error("There is no markdown to insert.");
  }

  return Word.run(async (context) => {
    const inserted = context.document
      .getSelection()
      .insertOoxml(buildMarkdown(blocks), Word.InsertLocation.replace);
    inserted.insertBookmark(MARKDOWN_BOOKMARK);
    await context.sync();

    return { blocks: blocks.length, bookmark: MARKDOWN_BOOKMARK };
  });
}

/** Remove what `insertMarkdown` wrote. */
export async function removeMarkdown(bookmark: string): Promise<void> {
  return Word.run(async (context) => {
    const range = context.document.getBookmarkRangeOrNullObject(bookmark);
    range.load("text");
    await context.sync();

    if (range.isNullObject) {
      throw new Error("That content is no longer in the document.");
    }

    range.delete();
    context.document.deleteBookmark(bookmark);
    await context.sync();
  });
}
