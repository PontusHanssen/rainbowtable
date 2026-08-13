/**
 * One colour table for syntax highlighting, shared by the preview and the document.
 *
 * The editor colours itself with CodeMirror's own theme; these are what the *document*
 * gets, so they are chosen to read on the template's Codeblock shading (D7D2CB, a warm
 * grey) rather than on white, and to survive greyscale print.
 */

/** Token classes, kept small on purpose: a report is not an IDE. */
export type CodeToken =
  | "keyword"
  | "string"
  | "number"
  | "comment"
  | "name"
  | "type"
  | "operator"
  | "punctuation"
  | "plain";

export const CODE_COLOURS: Record<CodeToken, { colour: string; bold?: boolean }> = {
  keyword: { colour: "0A3069", bold: true },
  string: { colour: "032F62" },
  number: { colour: "953800" },
  comment: { colour: "5A5A5A" },
  name: { colour: "5A2D8C" },
  type: { colour: "1A7F37" },
  operator: { colour: "4C4C4C" },
  punctuation: { colour: "4C4C4C" },
  plain: { colour: "24292F" },
};

export function colourFor(token: CodeToken): { colour: string; bold?: boolean } {
  return CODE_COLOURS[token] ?? CODE_COLOURS.plain;
}

/**
 * The languages a ```fence can name, and the aliases people actually type.
 *
 * `http` is deliberately absent: it is handled by `highlightHttp`, which understands
 * request and status lines in a way no general grammar does.
 */
export const LANGUAGE_ALIASES: Record<string, string> = {
  js: "javascript",
  javascript: "javascript",
  jsx: "javascript",
  ts: "javascript",
  typescript: "javascript",
  tsx: "javascript",
  html: "html",
  xhtml: "html",
  css: "css",
  json: "json",
  xml: "xml",
  sql: "sql",
  mysql: "sql",
  postgres: "sql",
  psql: "sql",
  php: "php",
  py: "python",
  python: "python",
  yaml: "yaml",
  yml: "yaml",
  sh: "shell",
  bash: "shell",
  shell: "shell",
  console: "shell",
};

/** The canonical language a fence names, or undefined to leave it plain. */
export function canonicalLanguage(fence: string | undefined): string | undefined {
  return fence ? LANGUAGE_ALIASES[fence.toLowerCase()] : undefined;
}
