import { canWatch, convertInlineCode, undoInlineCode, watchInlineCode } from "../word/inlineCode";
import { Feedback, byId, guard, show } from "./dom";

/* global HTMLButtonElement, HTMLInputElement */

/**
 * The markdown-style inline code controls. They share the Code tab's feedback lines with
 * the HTTP block, so the pane never shows two competing status messages.
 */
export function setUpInlineCodePanel(feedback: Feedback): void {
  const convert = byId<HTMLButtonElement>("inline-convert");
  const undo = byId<HTMLButtonElement>("inline-undo");
  const watchRow = byId("inline-watch-row");
  const watch = byId<HTMLInputElement>("inline-watch");

  const buttons = [convert, undo];
  let snapshot: string | undefined;
  let stopWatching: (() => void) | undefined;

  const report = (converted: number, scope: string) =>
    feedback.status(
      converted === 0
        ? "No `code` spans found."
        : `Styled ${converted} span${converted === 1 ? "" : "s"} across the ${scope}.`
    );

  convert.onclick = () =>
    guard(buttons, feedback, async () => {
      const result = await convertInlineCode();
      report(result.converted, result.scope);
      if (result.converted > 0) {
        snapshot = result.snapshot;
        show(undo, true);
      }
    });

  undo.onclick = () =>
    guard(buttons, feedback, async () => {
      if (!snapshot) {
        return;
      }
      await undoInlineCode(snapshot);
      snapshot = undefined;
      show(undo, false);
      feedback.status("Undone — the backticks are back.");
    });

  // Live conversion needs WordApi 1.6, above the manifest's floor, so it is offered only
  // where Word actually supports it rather than failing when switched on.
  show(watchRow, canWatch());

  watch.onchange = async () => {
    feedback.error("");
    try {
      if (watch.checked) {
        stopWatching = await watchInlineCode((converted) =>
          feedback.status(`Styled ${converted} span${converted === 1 ? "" : "s"} as you typed.`)
        );
        feedback.status("Watching for `code` spans. Leave the pane open.");
      } else {
        stopWatching?.();
        stopWatching = undefined;
        feedback.status("No longer watching.");
      }
    } catch (err) {
      watch.checked = false;
      feedback.error(String(err));
    }
  };
}
