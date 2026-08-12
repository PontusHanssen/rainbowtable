import { Section } from "./headings";
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

/** Word has no Heading10, so markdown depth is clamped rather than silently mis-styled. */
const DEEPEST_HEADING = 9;

export interface MarkdownInsertion {
  blocks: number;
  /** The heading level the top-level `#` was written at. */
  baseLevel: number;
  bookmark: string;
  snapshot: string;
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

/** Render one markdown block, with headings offset so `#` lands at `baseLevel`. */
function blockXml(block: Block, baseLevel: number): string {
  switch (block.kind) {
    case "heading": {
      const level = Math.min(baseLevel + block.level - 1, DEEPEST_HEADING);
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

const STYLE_DEFINITIONS =
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
export function buildMarkdown(blocks: Block[], baseLevel: number): string {
  const body = blocks.map((block) => blockXml(block, baseLevel)).join("");

  // The trailing paragraph keeps the last block from fusing with what follows the cursor.
  return wrapInPackage(`${body}<w:p/>`, STYLE_DEFINITIONS, NUMBERING_DEFINITIONS);
}

/**
 * Insert markdown at the cursor as a finding.
 *
 * `#` becomes a finding title under the selected section, `##` its sections, and so on —
 * the same rule the New finding button uses, so the two produce the same shape. With no
 * section selected the template's depth is assumed.
 */
export async function insertMarkdown(
  source: string,
  section?: Section
): Promise<MarkdownInsertion> {
  const blocks = parseMarkdown(source);
  if (blocks.length === 0) {
    throw new Error("There is no markdown to insert.");
  }

  const baseLevel = section ? Math.min(section.heading.level + 1, DEEPEST_HEADING) : 2;

  return Word.run(async (context) => {
    const captured = context.document.body.getRange("Whole").getOoxml();
    await context.sync();

    const inserted = context.document
      .getSelection()
      .insertOoxml(buildMarkdown(blocks, baseLevel), Word.InsertLocation.replace);
    inserted.insertBookmark(MARKDOWN_BOOKMARK);
    await context.sync();

    return {
      blocks: blocks.length,
      baseLevel,
      bookmark: MARKDOWN_BOOKMARK,
      // eslint-disable-next-line office-addins/load-object-before-read
      snapshot: captured.value,
    };
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
