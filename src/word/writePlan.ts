import { ParagraphPlan, RunPlan, listGroups } from "./documentPlan";

/* global Word */

/**
 * Carry out a plan through the Office.js API.
 *
 * `insertOoxml` costs about six seconds per call whatever it carries, while the API writes
 * hundreds of paragraphs in one — see the measurements in CLAUDE.md. Both the markdown and
 * HTTP paths funnel through here so there is one place that knows how to talk to Word.
 */

/** The report template's styles. A document without them is handled by the retry below. */
const CODE_STYLE = "Code block";
const INLINE_CODE_STYLE = "Code inline 2.0";

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

/**
 * The schemes a run may link to.
 *
 * The pane is the only side that can touch the document, and a plan reaches it from the
 * dialog over `messageParent` as well as from its own panels, so what gets written is
 * checked here rather than trusted from whoever built the plan. `file:` is the one that
 * matters: a UNC link in a report leaks the reader's credentials to whatever host it
 * names, on a single click, in a document that circulates outside the team.
 */
const LINKABLE = /^(?:https?|mailto):/i;

/** The address a run may be linked to, or undefined to write it as plain text. */
export function safeLink(link: string): string | undefined {
  return LINKABLE.test(link.trim()) ? link.trim() : undefined;
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
    if (run.colour) {
      range.font.color = `#${run.colour}`;
    }
    if (run.link) {
      // An address we will not link to still leaves its text in place, which keeps the
      // evidence intact and visible rather than quietly dropping part of a finding.
      const address = safeLink(run.link);
      if (address) {
        range.hyperlink = address;
      }
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
        writeRuns(paragraph, plan.runs, templateStyles);
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

export interface WriteResult {
  paragraphs: number;
  bookmark: string;
  /** True when the template's styles were unavailable and plain ones were used. */
  plainStyles: boolean;
}

async function attempt(
  plans: ParagraphPlan[],
  bookmark: string,
  templateStyles: boolean
): Promise<WriteResult> {
  return Word.run(async (context) => {
    const cursor = context.document.getSelection().paragraphs.getFirst();
    const written = writeParagraphs(cursor, plans, templateStyles);

    written[0]
      .getRange("Whole")
      .expandTo(written[written.length - 1].getRange("Whole"))
      .insertBookmark(bookmark);
    await context.sync();

    await attachLists(context, plans, written);

    return { paragraphs: written.length, bookmark, plainStyles: !templateStyles };
  });
}

/**
 * Write a plan at the cursor, bookmarked so it can be removed again.
 *
 * A document without the report template's styles rejects them by name, and the failure
 * lands on `sync` once the whole batch is queued. Rather than checking each style up
 * front, this writes once and, if that fails, writes again with built-in styles and direct
 * formatting; the caller is told which happened.
 */
export async function writePlan(plans: ParagraphPlan[], bookmark: string): Promise<WriteResult> {
  if (plans.length === 0) {
    throw new Error("There is nothing to insert.");
  }

  try {
    return await attempt(plans, bookmark, true);
  } catch (first) {
    // The first attempt may have written part of its paragraphs before the sync failed,
    // so it has to come out before the second one goes in. Nothing to remove is the
    // ordinary case — the failure usually lands before anything reaches the document.
    // Anything else means content is in there that we could not take back, and writing
    // again would leave the user with two copies and no way to tell them apart.
    try {
      await removeWritten(bookmark);
    } catch (cleanup) {
      if (String(cleanup).indexOf(NOTHING_WRITTEN) < 0) {
        throw new Error(
          `${first} Some of it was written and could not be removed (${cleanup}), so nothing ` +
            "was inserted a second time. Check the document at the cursor before retrying."
        );
      }
    }
    return attempt(plans, bookmark, false);
  }
}

/**
 * What `removeWritten` says when the bookmark is gone. `writePlan` tells that apart from
 * a removal that genuinely failed, which is the difference between retrying safely and
 * duplicating content.
 */
export const NOTHING_WRITTEN = "That content is no longer in the document.";

/** Remove what `writePlan` wrote. */
export async function removeWritten(bookmark: string): Promise<void> {
  return Word.run(async (context) => {
    const range = context.document.getBookmarkRangeOrNullObject(bookmark);
    range.load("text");
    await context.sync();

    if (range.isNullObject) {
      throw new Error(NOTHING_WRITTEN);
    }

    range.delete();
    context.document.deleteBookmark(bookmark);
    await context.sync();
  });
}
