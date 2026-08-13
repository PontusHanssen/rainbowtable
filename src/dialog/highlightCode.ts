import { Language, StreamLanguage } from "@codemirror/language";
import { Highlighter, highlightTree, tagHighlighter, tags } from "@lezer/highlight";
import { RunPlan } from "../word/documentPlan";
import { CodeToken, canonicalLanguage, colourFor } from "./codeColours";

/**
 * Colouring fenced code for the document.
 *
 * The same lezer grammars CodeMirror uses in the editor, walked with `highlightTree`, so
 * the editor, the preview and the document agree about what a token is. Two separate
 * highlighters would drift apart.
 *
 * Grammars are loaded on demand: together they are far larger than everything else here,
 * and most findings name at most one language.
 */

/** Lezer tags mapped to the handful of classes the document actually distinguishes. */
const HIGHLIGHTER: Highlighter = tagHighlighter([
  { tag: tags.keyword, class: "keyword" },
  { tag: tags.controlKeyword, class: "keyword" },
  { tag: tags.moduleKeyword, class: "keyword" },
  { tag: tags.operatorKeyword, class: "keyword" },
  { tag: tags.definitionKeyword, class: "keyword" },
  { tag: tags.string, class: "string" },
  { tag: tags.special(tags.string), class: "string" },
  { tag: tags.number, class: "number" },
  { tag: tags.bool, class: "number" },
  { tag: tags.null, class: "number" },
  { tag: tags.comment, class: "comment" },
  { tag: tags.lineComment, class: "comment" },
  { tag: tags.blockComment, class: "comment" },
  { tag: tags.variableName, class: "name" },
  { tag: tags.propertyName, class: "name" },
  { tag: tags.attributeName, class: "name" },
  { tag: tags.function(tags.variableName), class: "name" },
  { tag: tags.typeName, class: "type" },
  { tag: tags.tagName, class: "type" },
  { tag: tags.className, class: "type" },
  { tag: tags.operator, class: "operator" },
  { tag: tags.punctuation, class: "punctuation" },
  { tag: tags.bracket, class: "punctuation" },
]);

const loaders: Record<string, () => Promise<Language>> = {
  javascript: async () => (await import("@codemirror/lang-javascript")).javascriptLanguage,
  html: async () => (await import("@codemirror/lang-html")).htmlLanguage,
  css: async () => (await import("@codemirror/lang-css")).cssLanguage,
  json: async () => (await import("@codemirror/lang-json")).jsonLanguage,
  xml: async () => (await import("@codemirror/lang-xml")).xmlLanguage,
  sql: async () => (await import("@codemirror/lang-sql")).StandardSQL.language,
  php: async () => (await import("@codemirror/lang-php")).phpLanguage,
  python: async () => (await import("@codemirror/lang-python")).pythonLanguage,
  yaml: async () => (await import("@codemirror/lang-yaml")).yamlLanguage,
  // Shell has no lezer grammar of its own; the legacy stream mode is what CodeMirror uses.
  shell: async () =>
    StreamLanguage.define((await import("@codemirror/legacy-modes/mode/shell")).shell),
};

const loaded = new Map<string, Language>();

/** Load a grammar once, or undefined for a language we do not carry. */
export async function grammarFor(fence: string | undefined): Promise<Language | undefined> {
  const language = canonicalLanguage(fence);
  if (!language || !loaders[language]) {
    return undefined;
  }

  const already = loaded.get(language);
  if (already) {
    return already;
  }

  const grammar = await loaders[language]();
  loaded.set(language, grammar);
  return grammar;
}

/**
 * Split `code` into one array of runs per line, coloured by the grammar.
 *
 * `highlightTree` reports spans over the whole block, so they are cut at newlines: the
 * document holds one paragraph per line, and a span may cover several.
 */
export function highlightWith(grammar: Language, code: string): RunPlan[][] {
  const lines: RunPlan[][] = [[]];
  let at = 0;

  const emit = (from: number, to: number, token: CodeToken) => {
    const text = code.slice(from, to);
    if (!text) {
      return;
    }

    const { colour, bold } = colourFor(token);
    text.split("\n").forEach((part, index) => {
      if (index > 0) {
        lines.push([]);
      }
      if (part) {
        lines[lines.length - 1].push({ text: part, colour, bold });
      }
    });
  };

  highlightTree(grammar.parser.parse(code), HIGHLIGHTER, (from, to, classes) => {
    // Anything the grammar did not claim still has to reach the document.
    emit(at, from, "plain");
    emit(from, to, (classes as CodeToken) || "plain");
    at = to;
  });
  emit(at, code.length, "plain");

  return lines;
}
