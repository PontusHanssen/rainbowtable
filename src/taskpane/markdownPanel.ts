import { insertMarkdown, removeMarkdown } from "../word/markdownDoc";
import { byId, feedbackFor, guard, show } from "./dom";

/* global Office, URL, HTMLButtonElement, HTMLTextAreaElement, location */

/**
 * The markdown editor lives in an Office dialog, which has room for it. The dialog has no
 * document access of its own, so it sends the markdown back here and this writes it.
 *
 * Where dialogs are unavailable the pane falls back to its own editor rather than losing
 * the feature.
 */
export function setUpMarkdownPanel(): void {
  const open = byId<HTMLButtonElement>("markdown-open");
  const undo = byId<HTMLButtonElement>("markdown-undo");
  const fallback = byId("markdown-fallback");
  const input = byId<HTMLTextAreaElement>("markdown-input");
  const insert = byId<HTMLButtonElement>("markdown-insert");

  const feedback = feedbackFor("markdown");
  const buttons = [open, undo, insert];
  let inserted: string | undefined;

  const write = (source: string) =>
    guard(buttons, feedback, async () => {
      const result = await insertMarkdown(source);
      feedback.status(`Inserted ${result.blocks} blocks at the cursor.`);
      inserted = result.bookmark;
      show(undo, true);
    });

  /** Show the in-pane editor when the dialog cannot be used. */
  const useFallback = (why: string) => {
    show(fallback, true);
    show(open, false);
    feedback.status(`${why} Write the finding here instead.`);
  };

  open.onclick = () => {
    if (!Office.context.requirements.isSetSupported("DialogApi", "1.1")) {
      useFallback("This Word build cannot open add-in dialogs.");
      return;
    }

    // Same origin as the task pane, which is what a dialog is restricted to.
    const url = new URL("dialog.html", location.href).href;

    Office.context.ui.displayDialogAsync(url, { height: 70, width: 55 }, (opened) => {
      if (opened.status !== Office.AsyncResultStatus.Succeeded) {
        useFallback(`The dialog could not be opened (${opened.error.message}).`);
        return;
      }

      const dialog = opened.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        const source = (arg as { message?: string }).message ?? "";

        // Development-only probe: the dialog sends a payload of a known size and this
        // reports back what actually arrived, which is how the message channel's
        // undocumented limit gets measured. Real markdown is never JSON.
        if (source.startsWith('{"kind":"probe"')) {
          const probe = JSON.parse(source) as { bytes: number; payload: string };
          dialog.messageChild(
            JSON.stringify({ kind: "probeResult", bytes: probe.bytes, got: probe.payload.length })
          );
          return;
        }

        dialog.close();

        if (source.trim() === "") {
          feedback.status("Cancelled — nothing was inserted.");
          return;
        }
        write(source);
      });
    });
  };

  insert.onclick = () => write(input.value);

  undo.onclick = () =>
    guard(buttons, feedback, async () => {
      if (!inserted) {
        return;
      }
      await removeMarkdown(inserted);
      inserted = undefined;
      show(undo, false);
      feedback.status("Removed the inserted content.");
    });
}
