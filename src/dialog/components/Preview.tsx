import { ReactElement } from "react";
import { Block, Inline, parseMarkdown } from "../../word/markdown";

/* global window */

/**
 * The preview renders from our own parser, never a markdown library.
 *
 * That way it shows what Word will actually receive, including the deliberate limits of
 * the subset: syntax we do not support stays literal here exactly as it will in the
 * document. A general-purpose renderer would show CommonMark behaviour and mislead.
 */
function spans(runs: Inline[]): ReactElement[] {
  return runs.map((span, index) => {
    switch (span.kind) {
      case "bold":
        return <strong key={index}>{span.text}</strong>;
      case "italic":
        return <em key={index}>{span.text}</em>;
      case "code":
        return (
          <code key={index} className="inline">
            {span.text}
          </code>
        );
      case "link":
        /*
         * The click must never reach the frame's own navigation. A finding routinely
         * carries links — the CVSS bar writes one itself — and following one in here
         * would replace the dialog, taking the unwritten finding with it: nothing is
         * persisted, and a dialog that has lost its pane cannot be reconnected to.
         * preventDefault is what guarantees that, whatever the host makes of `target`.
         */
        return (
          <a
            key={index}
            href={span.url}
            rel="noopener noreferrer"
            target="_blank"
            onClick={(event) => {
              event.preventDefault();
              window.open(span.url, "_blank", "noopener");
            }}
          >
            {span.text}
          </a>
        );
      default:
        return <span key={index}>{span.text}</span>;
    }
  });
}

function leaf(item: Block, index: number): ReactElement {
  switch (item.kind) {
    case "heading": {
      const Tag = `h${item.level}` as "h1";
      return <Tag key={index}>{spans(item.spans)}</Tag>;
    }
    case "bullet":
    case "number":
      return <li key={index}>{spans(item.spans)}</li>;
    case "code":
      return <pre key={index}>{item.lines.join("\n")}</pre>;
    default:
      return <p key={index}>{spans(item.spans)}</p>;
  }
}

/**
 * Wrap runs of list items in a real list.
 *
 * They used to be emitted as bare `li` elements into a `div`, with `list-style` and a
 * left margin faking the look — invalid markup, no list semantics for a screen reader,
 * and ordered numbering left to whatever the browser recovered.
 */
function render(blocks: Block[]): ReactElement[] {
  const out: ReactElement[] = [];

  for (let i = 0; i < blocks.length; ) {
    const kind = blocks[i].kind;
    if (kind !== "bullet" && kind !== "number") {
      out.push(leaf(blocks[i], i));
      i += 1;
      continue;
    }

    const start = i;
    while (i < blocks.length && blocks[i].kind === kind) {
      i += 1;
    }

    const items = blocks.slice(start, i).map((item, offset) => leaf(item, start + offset));
    out.push(
      kind === "number" ? <ol key={start}>{items}</ol> : <ul key={start}>{items}</ul>
    );
  }

  return out;
}

export function Preview({ markdown }: { markdown: string }): ReactElement {
  return <div className="preview">{render(parseMarkdown(markdown))}</div>;
}
