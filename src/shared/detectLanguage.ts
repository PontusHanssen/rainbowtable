/**
 * Guessing what a pasted snippet is.
 *
 * Deliberately conservative: a wrong guess colours code as something it is not, which is
 * worse than leaving it plain. Every rule looks for something that would be odd in any
 * other language, and anything unrecognised returns undefined.
 */

/** Ordered: the first rule that matches wins, so the specific ones come first. */
const RULES: { language: string; matches: RegExp }[] = [
  // Unambiguous openers.
  { language: "php", matches: /^\s*<\?php\b/ },
  { language: "xml", matches: /^\s*<\?xml\b/ },
  { language: "html", matches: /^\s*<(!doctype html|html\b)/i },

  // Shebangs say it outright.
  { language: "shell", matches: /^#!.*\b(sh|bash|zsh)\b/ },
  { language: "python", matches: /^#!.*\bpython/ },

  // Declarations that would be strange anywhere else.
  { language: "go", matches: /^\s*package\s+\w+\s*$[\s\S]*^\s*func\b/m },
  { language: "csharp", matches: /^\s*using\s+System\b|\bConsole\.(Write|WriteLine)\s*\(/m },
  {
    language: "java",
    matches: /\b(public|private)\s+(static\s+)?(final\s+)?class\b|\bSystem\.out\.print/,
  },
  { language: "python", matches: /^\s*(def|class)\s+\w+.*:\s*$|^\s*(from\s+\S+\s+)?import\s+\w+/m },
  { language: "php", matches: /\$\w+\s*=|\becho\s+["'$]/ },

  {
    language: "sql",
    matches:
      /^\s*(SELECT|INSERT\s+INTO|UPDATE|DELETE\s+FROM|CREATE\s+(TABLE|DATABASE)|ALTER\s+TABLE|DROP\s+TABLE)\b/i,
  },

  {
    language: "javascript",
    matches: /\b(function\s*\w*\s*\(|const\s+\w+\s*=|let\s+\w+\s*=|=>|console\.log\s*\()/,
  },

  // A shell session, as pasted from a terminal.
  { language: "shell", matches: /^\s*[$#]\s+\S/m },

  { language: "css", matches: /^[.#]?[\w-]+(\s*[,>][\w\s.#-]+)*\s*\{[^}]*:[^}]*;/m },
  { language: "yaml", matches: /^---\s*$|^[\w-]+:\s+\S+$/m },
  { language: "xml", matches: /^\s*<[a-zA-Z][\w:-]*(\s+[\w:-]+\s*=|\s*\/?>)/ },
];

/** True when the whole snippet parses as JSON, which is a stronger signal than any regex. */
function looksLikeJson(code: string): boolean {
  const trimmed = code.trim();
  if (!/^[[{]/.test(trimmed)) {
    return false;
  }

  try {
    JSON.parse(trimmed);
    return true;
  } catch {
    // Reports quote truncated bodies constantly; a JSON-ish opener is still worth calling
    // JSON, but only when it carries the punctuation to match.
    return /"\s*:\s*/.test(trimmed);
  }
}

/**
 * The language a snippet appears to be, or undefined to leave it plain.
 *
 * HTTP is not detected here — `highlightHttp` recognises requests and responses itself,
 * and the callers try it first.
 */
export function detectLanguage(code: string): string | undefined {
  if (code.trim() === "") {
    return undefined;
  }

  if (looksLikeJson(code)) {
    return "json";
  }

  return RULES.find((rule) => rule.matches.test(code))?.language;
}
