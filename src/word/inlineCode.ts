/* global Word, Office */

/**
 * Markdown-style `inline code` in Word: find backtick-delimited spans, drop the
 * backticks, and apply the template's inline code character style.
 */

/** The character style the report template defines for code inside a sentence. */
export const INLINE_CODE_STYLE = "Code inline 2.0";

/**
 * Word's wildcard syntax, not a regular expression: a backtick, one or more characters
 * that are not backticks (`@` means "one or more of the preceding"), then a backtick.
 *
 * The exclusion is what keeps it from being greedy. A plain `` `*` `` would match from the
 * first backtick on the line to the last, swallowing everything between two separate
 * spans. It also means empty `` `` `` is not a span, which is what we want.
 */
export const CODE_SPAN_PATTERN = "`[!`]@`";

/** The text of a span without its backticks. */
export function stripDelimiters(match: string): string {
  return match.startsWith("`") && match.endsWith("`") && match.length >= 2
    ? match.slice(1, -1)
    : match;
}

export interface InlineCodeResult {
  converted: number;
  /** Where the conversion ran, for the pane to report back. */
  scope: "selection" | "document";
  /** The whole body as it was, so the change can be taken back. */
  snapshot: string;
}

/**
 * Convert every `code` span to the inline code style.
 *
 * Runs over the selection when there is one and the whole document otherwise, so it can
 * be used both to fix up a paragraph and to sweep a finished report.
 *
 * The snapshot is always the whole body even when only a selection was converted: undo
 * has to work after the cursor has moved on, and a selection cannot be relied on to still
 * be there.
 */
export async function convertInlineCode(): Promise<InlineCodeResult> {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    const captured = context.document.body.getRange("Whole").getOoxml();
    await context.sync();

    const selected = selection.text;
    const inSelection = selected.trim() !== "";
    const scope = inSelection ? selection : context.document.body.getRange("Whole");

    const found = scope.search(CODE_SPAN_PATTERN, { matchWildcards: true });
    found.load("items/text");
    await context.sync();

    found.items.forEach((span) => {
      // Replacing the text drops the backticks; the returned range is what gets styled.
      const replaced = span.insertText(stripDelimiters(span.text), Word.InsertLocation.replace);
      replaced.style = INLINE_CODE_STYLE;
    });
    await context.sync();

    return {
      converted: found.items.length,
      scope: inSelection ? "selection" : "document",
      // eslint-disable-next-line office-addins/load-object-before-read
      snapshot: captured.value,
    };
  });
}

/** Put the document back as it was before the last conversion. */
export async function undoInlineCode(snapshot: string): Promise<void> {
  return Word.run(async (context) => {
    context.document.body.getRange("Whole").insertOoxml(snapshot, Word.InsertLocation.replace);
    await context.sync();
  });
}

/** Whether this Word build can report paragraph edits as they happen. */
export function canWatch(): boolean {
  return Office.context.requirements.isSetSupported("WordApi", "1.6");
}

/**
 * Convert the spans in a single paragraph, without a snapshot.
 *
 * This is the live path: it touches only the paragraph Word says changed, so it stays
 * cheap enough to run while someone is typing.
 */
async function convertInParagraph(context: Word.RequestContext, id: string): Promise<number> {
  const paragraph = context.document.getParagraphByUniqueLocalId(id);
  const found = paragraph.search(CODE_SPAN_PATTERN, { matchWildcards: true });
  found.load("items/text");
  await context.sync();

  found.items.forEach((span) => {
    const replaced = span.insertText(stripDelimiters(span.text), Word.InsertLocation.replace);
    replaced.style = INLINE_CODE_STYLE;
  });
  await context.sync();

  return found.items.length;
}

/**
 * Convert spans as they are typed, by reacting to the paragraphs Word reports as changed.
 *
 * Only the changed paragraphs are examined, so this stays cheap. It does not recurse:
 * converting a span removes its backticks, so the edit it triggers finds nothing left to
 * convert. `running` guards against overlapping passes while one is still syncing.
 *
 * The task pane must stay open — the handler lives in the pane's runtime. There is no undo
 * for conversions made this way; they remove backticks the user has only just typed.
 *
 * Returns a function that stops watching. Needs WordApi 1.6 — check `canWatch` first.
 */
export async function watchInlineCode(
  onConverted: (count: number) => void
): Promise<() => Promise<void>> {
  let running = false;
  // Checked first thing in the handler. Deregistration below is the real mechanism, but
  // this makes switching off take effect immediately and survives it failing.
  let watching = true;

  const handler = async (args: Word.ParagraphChangedEventArgs) => {
    if (!watching || running) {
      return;
    }
    running = true;
    try {
      const converted = await Word.run(async (context) => {
        let total = 0;
        for (const id of args.uniqueLocalIds) {
          total += await convertInParagraph(context, id);
        }
        return total;
      });
      if (converted > 0) {
        onConverted(converted);
      }
    } finally {
      running = false;
    }
  };

  // The registration has to be kept: a handler can only be removed through the same
  // RequestContext it was added in, and every Word.run creates a new one. Calling
  // remove() on a fresh context silently does nothing and the handler keeps firing.
  const registration = await Word.run(async (context) => {
    const added = context.document.onParagraphChanged.add(handler);
    await context.sync();
    return added;
  });

  return async () => {
    watching = false;
    // Removed through the context it was added in, then synced on that same context.
    registration.remove();
    await registration.context.sync();
  };
}
