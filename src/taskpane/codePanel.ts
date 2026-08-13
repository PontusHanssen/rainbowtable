import { LANGUAGE_ALIASES } from "../shared/codeColours";
import { detectLanguage } from "../shared/detectLanguage";
import { planCode } from "../shared/planCode";
import { highlightHttp } from "../word/http";
import { removeWritten, writePlan } from "../word/writePlan";
import { byId, feedbackFor, guard, make, show } from "./dom";

/* global HTMLButtonElement, HTMLInputElement, HTMLSelectElement, HTMLTextAreaElement */

const CODE_BOOKMARK = "_ptcode";

/** The languages offered by name, deduplicated from the aliases. */
const LANGUAGES = [...new Set(Object.values(LANGUAGE_ALIASES))].sort();

/**
 * Pasting code or a captured HTTP message into the document.
 *
 * Pasting here rather than into Word keeps the text exactly as it was captured —
 * autocorrect and smart quotes would otherwise alter what is evidence.
 */
export function setUpCodePanel(): void {
  const input = byId<HTMLTextAreaElement>("code-input");
  const detected = byId("code-detected");
  const language = byId<HTMLSelectElement>("code-language");
  const highlight = byId<HTMLInputElement>("code-highlight");
  const insert = byId<HTMLButtonElement>("code-insert");
  const undo = byId<HTMLButtonElement>("code-undo");

  const feedback = feedbackFor("code");
  const buttons = [insert, undo];
  let inserted: string | undefined;

  LANGUAGES.forEach((name) => {
    const option = make("option", undefined, name);
    option.value = name;
    language.appendChild(option);
  });

  /** Say what will happen before it happens, as the markdown preview does. */
  const describe = () => {
    const code = input.value;
    insert.disabled = code.trim() === "";

    if (!code.trim()) {
      detected.textContent = "";
      return;
    }
    if (!highlight.checked) {
      detected.textContent = "Highlighting off — it will go in as plain code.";
      return;
    }
    if (language.value !== "auto") {
      detected.textContent = `Will be highlighted as ${language.value}.`;
      return;
    }

    const message = highlightHttp(code);
    if (message.kind !== "unknown") {
      detected.textContent = `Recognised as an HTTP ${message.kind}.`;
      return;
    }

    const guess = detectLanguage(code);
    detected.textContent = guess
      ? `Looks like ${guess}.`
      : "Not recognised — it will go in as plain code.";
  };

  input.oninput = describe;
  language.onchange = describe;
  highlight.onchange = describe;

  insert.onclick = () =>
    guard(buttons, feedback, async () => {
      const planned = await planCode(input.value, {
        language: language.value === "auto" ? undefined : language.value,
        highlight: highlight.checked,
      });

      const written = await writePlan(planned.plans, CODE_BOOKMARK);
      feedback.status(
        `Inserted ${written.paragraphs} lines` +
          (planned.language ? ` as ${planned.language}.` : " as plain code.") +
          (written.plainStyles ? " This document has no code style, so plain text was used." : "")
      );
      inserted = written.bookmark;
      show(undo, true);
    });

  undo.onclick = () =>
    guard(buttons, feedback, async () => {
      if (!inserted) {
        return;
      }
      await removeWritten(inserted);
      inserted = undefined;
      show(undo, false);
      feedback.status("Removed the inserted block.");
    });

  describe();
}
