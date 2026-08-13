import { ParagraphPlan, RunPlan, listGroups, planFromBlocks } from "./documentPlan";
import { parseMarkdown } from "./markdown";

/* global Word */

/**
 * Write markdown into the document through the Office.js API.
 *
 * Not through `insertOoxml`, which costs about six seconds per call whatever it carries —
 * see the measurements in CLAUDE.md. Everything here happens inside one `Word.run`, and a
 * round trip is tens of milliseconds.
 */

/** The report template's styles. A document without them is handled by the retry below. */
const CODE_STYLE = "Code block";
const INLINE_CODE_STYLE = "Code inline 2.0";

const MARKDOWN_BOOKMARK = "_ptmd";

/** Heading style names, indexed by level, so no cast is needed to assign one. */
const HEADING_STYLES = [
  "Heading1",
  "Heading2",
  "Heading3",
  "Heading4",
  "Heading5",
  "Heading6",
] as const;

export interface MarkdownInsertion {
  blocks: number;
  /** Bookmark wrapping what was written, which is how it is removed again. */
  bookmark: string;
  /** True when the template's styles were unavailable and plain ones were used. */
  plainStyles: boolean;
}

/** Write one paragraph's runs, styling each as it goes. */
function writeRuns(paragraph: Word.Paragraph, runs: RunPlan[], templateStyles: boolean): void {
  runs.forEach((run) => {
    if (run.text === "") {
      return;
    }
    const range = paragraph.insertText(run.text, Word.InsertLocation.end);

    if (run.bold) {
      range.font.bold = true;
    }
    if (run.italic) {
      range.font.italic = true;
    }
    if (run.code) {
      if (templateStyles) {
        range.style = INLINE_CODE_STYLE;
      } else {
        range.font.name = "Courier New";
      }
    }
    if (run.link) {
      range.hyperlink = run.link;
    }
  });
}

/** Create the paragraphs, in order, each chained after the last. */
function writeParagraphs(
  after: Word.Paragraph,
  plans: ParagraphPlan[],
  templateStyles: boolean
): Word.Paragraph[] {
  const written: Word.Paragraph[] = [];
  let previous = after;

  for (const plan of plans) {
    const paragraph = previous.insertParagraph("", Word.InsertLocation.after);

    switch (plan.kind) {
      case "heading":
        paragraph.styleBuiltIn = HEADING_STYLES[plan.level - 1];
        writeRuns(paragraph, plan.runs, templateStyles);
        break;

      case "code":
        if (templateStyles) {
          paragraph.style = CODE_STYLE;
        } else {
          paragraph.styleBuiltIn = "Normal";
          paragraph.font.name = "Courier New";
        }
        // One unformatted run: the style carries the monospace and the shading.
        writeRuns(paragraph, [{ text: plan.text }], templateStyles);
        break;

      case "listItem":
        paragraph.styleBuiltIn = "ListParagraph";
        writeRuns(paragraph, plan.runs, templateStyles);
        break;

      default:
        paragraph.styleBuiltIn = "Normal";
        writeRuns(paragraph, plan.runs, templateStyles);
    }

    written.push(paragraph);
    previous = paragraph;
  }

  return written;
}

/**
 * Turn each run of list items into a real Word list.
 *
 * The first item starts a list and the rest attach to it by id, which needs its own round
 * trip: the id does not exist until Word has made the list. Grouping matters — without it
 * every item would begin again at 1.
 */
async function attachLists(
  context: Word.RequestContext,
  plans: ParagraphPlan[],
  written: Word.Paragraph[]
): Promise<void> {
  const groups = listGroups(plans);
  if (groups.size === 0) {
    return;
  }

  const lists = new Map<number, { list: Word.List; members: number[]; ordered: boolean }>();
  groups.forEach((members, group) => {
    const plan = plans[members[0]] as Extract<ParagraphPlan, { kind: "listItem" }>;
    const list = written[members[0]].startNewList();
    list.load("id");
    lists.set(group, { list, members, ordered: plan.ordered });
  });
  await context.sync();

  lists.forEach(({ list, members, ordered }) => {
    // Numbering and bullets are set by different calls; there is no bullet numbering.
    if (ordered) {
      list.setLevelNumbering(0, Word.ListNumbering.arabic);
    } else {
      list.setLevelBullet(0, Word.ListBullet.solid);
    }
    members.slice(1).forEach((index) => written[index].attachToList(list.id, 0));
  });
  await context.sync();
}

async function write(source: string, templateStyles: boolean): Promise<MarkdownInsertion> {
  const plans = planFromBlocks(parseMarkdown(source));

  return Word.run(async (context) => {
    const cursor = context.document.getSelection().paragraphs.getFirst();
    const written = writeParagraphs(cursor, plans, templateStyles);

    written[0]
      .getRange("Whole")
      .expandTo(written[written.length - 1].getRange("Whole"))
      .insertBookmark(MARKDOWN_BOOKMARK);
    await context.sync();

    await attachLists(context, plans, written);

    return { blocks: plans.length, bookmark: MARKDOWN_BOOKMARK, plainStyles: !templateStyles };
  });
}

/**
 * Insert markdown at the cursor. `#` is Heading1, `##` is Heading2.
 *
 * A document without the report template's styles rejects them by name, and the failure
 * lands on `sync` once the whole batch is queued. Rather than checking each style up
 * front, this writes once and, if that fails, writes again with built-in styles and direct
 * formatting; the caller is told which happened.
 */
export async function insertMarkdown(source: string): Promise<MarkdownInsertion> {
  if (planFromBlocks(parseMarkdown(source)).length === 0) {
    throw new Error("There is no markdown to insert.");
  }

  try {
    return await write(source, true);
  } catch {
    // Clear anything the failed attempt left behind before trying again.
    await removeMarkdown(MARKDOWN_BOOKMARK).catch(() => undefined);
    return write(source, false);
  }
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
