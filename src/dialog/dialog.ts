import { outline, parseMarkdown } from "../word/markdown";

/* global document, Office, HTMLElement, HTMLButtonElement, HTMLTextAreaElement */

/**
 * The markdown editor, running in an Office dialog.
 *
 * A dialog has its own runtime and no access to the document, so this cannot insert
 * anything itself. It hands the markdown to the task pane with `messageParent`, and the
 * pane — which does have document access — writes it.
 */
Office.onReady(() => {
  const input = document.getElementById("markdown") as HTMLTextAreaElement;
  const preview = document.getElementById("preview") as HTMLElement;
  const insert = document.getElementById("insert") as HTMLButtonElement;
  const cancel = document.getElementById("cancel") as HTMLButtonElement;

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

    const tree = summary.headings
      .map((heading) => `${"  ".repeat(heading.level - 1)}H${heading.level}  ${heading.text}`)
      .join("\n");

    preview.textContent = `${counts.join(", ")}.` + (tree ? `\n\n${tree}` : "");
  };

  input.oninput = describe;

  insert.onclick = () => Office.context.ui.messageParent(input.value);
  cancel.onclick = () => Office.context.ui.messageParent("");

  describe();
  input.focus();
});
