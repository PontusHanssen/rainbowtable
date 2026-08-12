import { outline, parseMarkdown } from "../word/markdown";
import { insertMarkdown, removeMarkdown } from "../word/markdownDoc";
import { byId, feedbackFor, guard, show } from "./dom";

/* global HTMLButtonElement, HTMLTextAreaElement */

export function setUpMarkdownPanel(): void {
  const input = byId<HTMLTextAreaElement>("markdown-input");
  const preview = byId("markdown-preview");
  const insert = byId<HTMLButtonElement>("markdown-insert");
  const undo = byId<HTMLButtonElement>("markdown-undo");

  const feedback = feedbackFor("markdown");
  const buttons = [insert, undo];
  let inserted: string | undefined;

  /** Show what will be written before anything is written. */
  const describe = () => {
    const source = input.value;
    insert.disabled = source.trim() === "";

    if (!source.trim()) {
      preview.textContent = "";
      return;
    }

    const summary = outline(parseMarkdown(source));

    const counts = [
      `${summary.headings.length} heading${summary.headings.length === 1 ? "" : "s"}`,
      `${summary.paragraphs} paragraph${summary.paragraphs === 1 ? "" : "s"}`,
      summary.listItems > 0 ? `${summary.listItems} list items` : "",
      summary.codeBlocks > 0 ? `${summary.codeBlocks} code blocks` : "",
      summary.links > 0 ? `${summary.links} links` : "",
    ].filter(Boolean);

    // The outline, indented as it will appear, with the heading style each line becomes.
    const tree = summary.headings
      .map((heading) => `${"  ".repeat(heading.level - 1)}H${heading.level}  ${heading.text}`)
      .join("\n");

    preview.textContent = `${counts.join(", ")}.` + (tree ? `\n\n${tree}` : "");
  };

  input.oninput = () => {
    feedback.status("");
    describe();
  };

  insert.onclick = () =>
    guard(buttons, feedback, async () => {
      const result = await insertMarkdown(input.value);
      feedback.status(`Inserted ${result.blocks} blocks at the cursor.`);
      inserted = result.bookmark;
      show(undo, true);
    });

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

  describe();
}
