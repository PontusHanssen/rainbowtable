import { ReactElement } from "react";
import { Block, Inline, parseMarkdown } from "../../word/markdown";

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
        return (
          <a key={index} href={span.url}>
            {span.text}
          </a>
        );
      default:
        return <span key={index}>{span.text}</span>;
    }
  });
}

function block(item: Block, index: number): ReactElement {
  switch (item.kind) {
    case "heading": {
      const Tag = `h${item.level}` as "h1";
      return <Tag key={index}>{spans(item.spans)}</Tag>;
    }
    case "bullet":
      return <li key={index}>{spans(item.spans)}</li>;
    case "number":
      return (
        <li key={index} className="ordered">
          {spans(item.spans)}
        </li>
      );
    case "code":
      return <pre key={index}>{item.lines.join("\n")}</pre>;
    default:
      return <p key={index}>{spans(item.spans)}</p>;
  }
}

export function Preview({ markdown }: { markdown: string }): ReactElement {
  return <div className="preview">{parseMarkdown(markdown).map(block)}</div>;
}
