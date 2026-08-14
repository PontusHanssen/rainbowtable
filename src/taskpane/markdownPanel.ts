import { ToDialog, ToPane, decode, encode } from "../shared/protocol";
import { insertMarkdown, removeMarkdown } from "../word/markdownDoc";
import { removeWritten, writePlan } from "../word/writePlan";
import { byId, feedbackFor, guard, messageOf, show } from "./dom";

/* global Office, URL, HTMLButtonElement, HTMLTextAreaElement, location */

/**
 * The pane's half of the authoring dialog.
 *
 * The dialog has no document access, so it stays open and sends what it wants written;
 * this serves those requests and answers over the same channel. The pane has to stay open
 * too — it is the only side that can touch the document.
 */
export function setUpMarkdownPanel(): void {
  const open = byId<HTMLButtonElement>("markdown-open");
  const undo = byId<HTMLButtonElement>("markdown-undo");
  const fallback = byId("markdown-fallback");
  const input = byId<HTMLTextAreaElement>("markdown-input");
  const insert = byId<HTMLButtonElement>("markdown-insert");

  const feedback = feedbackFor("markdown");
  const buttons = [open, undo, insert];
  let lastBookmark: string | undefined;

  const describe = (paragraphs: number, plainStyles: boolean) =>
    `Inserted ${paragraphs} paragraphs.` +
    (plainStyles ? " This document has no template styles, so plain ones were used." : "");

  const write = (source: string) =>
    guard(
      buttons,
      feedback,
      async () => {
        const result = await insertMarkdown(source);
        feedback.status(describe(result.paragraphs, result.plainStyles));
        lastBookmark = result.bookmark;
        show(undo, true);
      },
      "Inserting the finding…"
    );

  /** Show the in-pane editor when the dialog cannot be used. */
  const useFallback = (why: string) => {
    show(fallback, true);
    show(open, false);
    feedback.status(`${why} Write the finding here instead.`);
  };

  /** Serve one request from the dialog, replying on the same channel either way. */
  const serve = async (dialog: Office.Dialog, request: ToPane): Promise<void> => {
    const reply = (message: ToDialog) => dialog.messageChild(encode(message));

    /*
     * The heartbeat is answered without touching the feedback region. It arrives every
     * five seconds, so clearing the busy state here would wipe "Inserting the finding…"
     * part-way through the insert it exists to explain.
     */
    if (request.kind === "ping") {
      reply({ kind: "pong", requestId: request.requestId });
      return;
    }

    try {
      if (request.kind === "insert") {
        // The dialog is where the user is looking, but the writing happens here and takes
        // seconds — say so, or the pane looks idle while it works.
        feedback.busy("Inserting the finding…");
        // Already planned and coloured by the dialog; the pane only writes.
        const result = await writePlan(request.plans, "_ptmd");
        lastBookmark = result.bookmark;
        show(undo, true);
        feedback.status(describe(result.paragraphs, result.plainStyles));
        reply({
          kind: "inserted",
          requestId: request.requestId,
          bookmark: result.bookmark,
          paragraphs: result.paragraphs,
          plainStyles: result.plainStyles,
        });
      } else if (request.kind === "remove") {
        feedback.busy("Removing the finding…");
        await removeWritten(request.bookmark);
        feedback.status("Removed what the editor last inserted.");
        reply({ kind: "removed", requestId: request.requestId });
      } else {
        dialog.close();
        feedback.status("Editor closed.");
      }
    } catch (err) {
      feedback.error(messageOf(err));
      reply({ kind: "failed", requestId: request.requestId, reason: messageOf(err) });
    } finally {
      feedback.busy();
    }
  };

  open.onclick = () => {
    if (!Office.context.requirements.isSetSupported("DialogApi", "1.1")) {
      useFallback("This Word build cannot open add-in dialogs.");
      return;
    }

    // Same origin as the task pane, which is what a dialog is restricted to.
    const url = new URL("dialog.html", location.href).href;

    Office.context.ui.displayDialogAsync(url, { height: 80, width: 70 }, (opened) => {
      if (opened.status !== Office.AsyncResultStatus.Succeeded) {
        useFallback(`The dialog could not be opened (${opened.error.message}).`);
        return;
      }

      const dialog = opened.value;
      feedback.status("Editor open. Leave this pane open — it does the writing.");

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        const raw = (arg as { message?: string }).message ?? "";

        const request = decode<ToPane>(raw);
        if (request) {
          serve(dialog, request);
        }
      });

      dialog.addEventHandler(Office.EventType.DialogEventReceived, () =>
        feedback.status("Editor closed.")
      );
    });
  };

  insert.onclick = () => write(input.value);

  undo.onclick = () =>
    guard(
      buttons,
      feedback,
      async () => {
        if (!lastBookmark) {
          return;
        }
        await removeMarkdown(lastBookmark);
        lastBookmark = undefined;
        show(undo, false);
        feedback.status("Removed what was last inserted.");
      },
      "Removing the finding…"
    );
}
