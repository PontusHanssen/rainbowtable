import { CodeHighlighter, ParagraphPlan, planFromBlocks, planFromHttp } from "../word/documentPlan";
import { highlightHttp } from "../word/http";
import { colourOf } from "../word/httpColours";
import { parseMarkdown } from "../word/markdown";
import { canonicalLanguage } from "./codeColours";
import { grammarFor, highlightWith } from "./highlightCode";

/**
 * Turn markdown into a plan, colouring fenced code on the way.
 *
 * This happens in the dialog rather than the pane because the grammars live here: the pane
 * writes to the document and is kept small, and a plan crosses the message channel just as
 * easily as text.
 */

/** A fence tagged `http`, or one that simply looks like a request or response. */
function httpLines(block: { lines: string[]; language?: string }) {
  const looksHttp = block.language?.toLowerCase() === "http";
  if (!looksHttp && block.language) {
    return undefined;
  }

  const message = highlightHttp(block.lines.join("\n"));
  if (message.kind === "unknown") {
    return undefined;
  }

  return planFromHttp(message, colourOf).map((plan) => plan.runs);
}

export async function planMarkdown(markdown: string): Promise<ParagraphPlan[]> {
  const blocks = parseMarkdown(markdown);

  // Load every grammar the document names before planning, since planning is synchronous.
  const languages = new Set(
    blocks
      .filter((block) => block.kind === "code")
      .map((block) => canonicalLanguage((block as { language?: string }).language))
      .filter((language): language is string => language !== undefined)
  );
  const grammars = new Map(
    await Promise.all(
      [...languages].map(async (language) => [language, await grammarFor(language)] as const)
    )
  );

  const highlight: CodeHighlighter = (block) => {
    const asHttp = httpLines(block);
    if (asHttp) {
      return asHttp;
    }

    const grammar = grammars.get(canonicalLanguage(block.language) ?? "");
    return grammar ? highlightWith(grammar, block.lines.join("\n")) : undefined;
  };

  return planFromBlocks(blocks, highlight);
}
