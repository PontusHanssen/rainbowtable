import { ParagraphPlan, planFromHttp } from "../word/documentPlan";
import { highlightHttp } from "../word/http";
import { colourOf } from "../word/httpColours";
import { canonicalLanguage } from "./codeColours";
import { detectLanguage } from "./detectLanguage";

/**
 * Turning a pasted snippet into paragraphs, coloured or plain.
 *
 * The grammars are hundreds of KB and the task pane is deliberately tiny, so the
 * highlighter is reached through a dynamic `import()`: nothing is loaded until someone
 * pastes code with highlighting switched on.
 */

export interface CodePlan {
  plans: ParagraphPlan[];
  /** What it was treated as, for the pane to report back. */
  language: string | undefined;
  /** True when it was recognised as an HTTP message rather than a language. */
  http: boolean;
}

/** Plain paragraphs, one per line, with no colour at all. */
function plainLines(code: string): ParagraphPlan[] {
  return code.split("\n").map((line) => ({ kind: "code" as const, runs: [{ text: line }] }));
}

/**
 * Plan a snippet.
 *
 * `language` overrides detection; `"none"` means the reader asked for no highlighting at
 * all, which is honoured even for HTTP.
 */
export async function planCode(
  code: string,
  options: { language?: string; highlight: boolean }
): Promise<CodePlan> {
  if (!options.highlight) {
    return { plans: plainLines(code), language: undefined, http: false };
  }

  // HTTP first: it is recognised by shape, and no general grammar understands a status line.
  const asked = options.language?.toLowerCase();
  if (!asked || asked === "http") {
    const message = highlightHttp(code);
    if (message.kind !== "unknown") {
      return { plans: planFromHttp(message, colourOf), language: "http", http: true };
    }
    if (asked === "http") {
      return { plans: plainLines(code), language: undefined, http: false };
    }
  }

  const language = canonicalLanguage(asked) ?? detectLanguage(code);
  if (!language) {
    return { plans: plainLines(code), language: undefined, http: false };
  }

  const { grammarFor, highlightWith } = await import("./highlightCode");
  const grammar = await grammarFor(language);
  if (!grammar) {
    return { plans: plainLines(code), language: undefined, http: false };
  }

  const lines = highlightWith(grammar, code);
  return {
    plans: lines.map((runs) => ({ kind: "code" as const, runs })),
    language,
    http: false,
  };
}
