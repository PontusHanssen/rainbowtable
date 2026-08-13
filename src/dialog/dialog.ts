import { outline, parseMarkdown } from "../word/markdown";
import { Measurement, formatMeasurements, sizeLadder, syntheticMessage } from "../word/limits";

/* global document, Office, location, setTimeout, clearTimeout, HTMLElement, HTMLButtonElement, HTMLTextAreaElement */

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

  setUpMessageProbe();

  describe();
  input.focus();
});

/**
 * Development-only: measure how large a message `messageParent` will actually carry.
 *
 * Office documents no limit for it, and it is the channel every finding crosses on its way
 * to the document. The pane echoes back the length it received, so a payload that is
 * truncated or dropped shows up as a mismatch rather than as silence.
 */
function setUpMessageProbe(): void {
  if (location.hostname !== "localhost") {
    return;
  }

  const probe = document.getElementById("probe") as HTMLElement;
  const run = document.getElementById("probe-run") as HTMLButtonElement;
  const output = document.getElementById("probe-output") as HTMLElement;
  probe.hidden = false;

  let awaiting: ((got: number) => void) | undefined;

  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
    const reply = JSON.parse((arg as { message: string }).message) as { got: number };
    awaiting?.(reply.got);
    awaiting = undefined;
  });

  run.onclick = async () => {
    run.disabled = true;
    const results: Measurement[] = [];

    for (const bytes of sizeLadder(16 * 1024, 8 * 1024 * 1024)) {
      const payload = syntheticMessage(bytes);
      const started = Date.now();

      // No reply within ten seconds counts as a failure: a message too large to carry
      // tends to vanish rather than raise anything.
      const got = await new Promise<number>((resolve) => {
        const timer = setTimeout(() => resolve(-1), 10000);
        awaiting = (value) => {
          clearTimeout(timer);
          resolve(value);
        };
        Office.context.ui.messageParent(JSON.stringify({ kind: "probe", bytes, payload }));
      });

      const ms = Date.now() - started;
      if (got === payload.length) {
        results.push({ bytes, ok: true, ms });
      } else {
        results.push({
          bytes,
          ok: false,
          ms,
          detail: got < 0 ? "no reply within 10s" : `arrived as ${got} bytes`,
        });
        output.textContent = formatMeasurements("messageParent", results);
        break;
      }
      output.textContent = formatMeasurements("messageParent", results);
    }

    run.disabled = false;
  };
}
