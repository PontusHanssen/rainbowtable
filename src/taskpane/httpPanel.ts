import { highlightHttp } from "../word/http";
import { insertHttpBlock, removeHttpBlock } from "../word/httpBlock";
import { byId, feedbackFor, guard, show } from "./dom";
import { setUpInlineCodePanel } from "./inlineCodePanel";

/* global HTMLButtonElement, HTMLTextAreaElement */

const DESCRIBE_BODY: Record<string, string> = {
  json: "JSON body",
  xml: "XML body",
  form: "form-encoded body",
  none: "no structured body",
};

export function setUpHttpPanel(): void {
  const input = byId<HTMLTextAreaElement>("http-input");
  const detected = byId("http-detected");
  const insert = byId<HTMLButtonElement>("http-insert");
  const undo = byId<HTMLButtonElement>("http-undo");

  const feedback = feedbackFor("http");
  const buttons = [insert, undo];

  let inserted: string | undefined;

  /** Parsed on every keystroke so the pane says what it sees before anything is written. */
  const describe = () => {
    const raw = input.value;
    insert.disabled = raw.trim() === "";

    if (!raw.trim()) {
      detected.textContent = "";
      return;
    }

    const message = highlightHttp(raw);
    detected.textContent =
      message.kind === "unknown"
        ? "Not recognised as HTTP — it will be inserted as a plain code block."
        : `Recognised as an HTTP ${message.kind} with ${DESCRIBE_BODY[message.body]}.`;
  };

  input.oninput = () => {
    feedback.status("");
    describe();
  };

  insert.onclick = () =>
    guard(buttons, feedback, async () => {
      const result = await insertHttpBlock(input.value);
      feedback.status(`Inserted ${result.lines} lines as a ${result.kind}.`);
      inserted = result.bookmark;
      show(undo, true);
    });

  undo.onclick = () =>
    guard(buttons, feedback, async () => {
      if (!inserted) {
        return;
      }
      await removeHttpBlock(inserted);
      inserted = undefined;
      show(undo, false);
      feedback.status("Removed the inserted block.");
    });

  setUpInlineCodePanel(feedback);
  describe();
}
