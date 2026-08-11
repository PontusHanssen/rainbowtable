/**
 * Syntax highlighting for HTTP requests and responses pasted from Burp.
 *
 * The message is never reformatted — not re-indented, not pretty-printed, not
 * re-encoded. A request in a pentest report is evidence, so the bytes that go in are the
 * bytes that come out; only colour is added.
 */

export type TokenKind =
  | "method"
  | "target"
  | "version"
  | "status2xx"
  | "status3xx"
  | "status4xx"
  | "status5xx"
  | "reason"
  | "headerName"
  | "punctuation"
  | "headerValue"
  | "key"
  | "string"
  | "number"
  | "literal"
  | "tag"
  | "attrName"
  | "attrValue"
  | "text";

export interface Token {
  text: string;
  kind: TokenKind;
}

export type MessageKind = "request" | "response" | "unknown";

export interface HttpMessage {
  kind: MessageKind;
  /** Detected body syntax, for the pane to report back. */
  body: "json" | "xml" | "form" | "none";
  lines: Token[][];
}

const REQUEST_LINE = /^([A-Z]+)(\s+)(\S+)(\s+)(HTTP\/\d(?:\.\d)?)\s*$/;
const STATUS_LINE = /^(HTTP\/\d(?:\.\d)?)(\s+)(\d{3})(\s*)(.*)$/;

/** Split on either line ending; Burp pastes CRLF, Word and browsers vary. */
function toLines(raw: string): string[] {
  return raw.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
}

function statusKind(code: string): TokenKind {
  switch (code[0]) {
    case "2":
      return "status2xx";
    case "3":
      return "status3xx";
    case "4":
      return "status4xx";
    default:
      return code[0] === "5" ? "status5xx" : "reason";
  }
}

/** The request or status line, or undefined when the first line is neither. */
function startLine(line: string): { kind: MessageKind; tokens: Token[] } | undefined {
  const request = REQUEST_LINE.exec(line);
  if (request) {
    return {
      kind: "request",
      tokens: [
        { text: request[1], kind: "method" },
        { text: request[2], kind: "text" },
        { text: request[3], kind: "target" },
        { text: request[4], kind: "text" },
        { text: request[5], kind: "version" },
      ],
    };
  }

  const response = STATUS_LINE.exec(line);
  if (response) {
    const tokens: Token[] = [
      { text: response[1], kind: "version" },
      { text: response[2], kind: "text" },
      { text: response[3], kind: statusKind(response[3]) },
    ];
    if (response[4]) {
      tokens.push({ text: response[4], kind: "text" });
    }
    if (response[5]) {
      tokens.push({ text: response[5], kind: "reason" });
    }
    return { kind: "response", tokens };
  }

  return undefined;
}

/** `Name: value`, split at the first colon so values keep their own colons. */
function headerLine(line: string): Token[] {
  const colon = line.indexOf(":");
  if (colon <= 0) {
    return [{ text: line, kind: "text" }];
  }

  const value = line.slice(colon + 1);
  const leading = /^\s*/.exec(value)?.[0] ?? "";

  return [
    { text: line.slice(0, colon), kind: "headerName" },
    { text: ":", kind: "punctuation" },
    ...(leading ? [{ text: leading, kind: "text" as TokenKind }] : []),
    ...(value.slice(leading.length)
      ? [{ text: value.slice(leading.length), kind: "headerValue" as TokenKind }]
      : []),
  ];
}

function scanJson(body: string): Token[] {
  const tokens: Token[] = [];
  let plain = "";
  let i = 0;

  const flush = () => {
    if (plain) {
      tokens.push({ text: plain, kind: "text" });
      plain = "";
    }
  };

  while (i < body.length) {
    const character = body[i];

    if (character === '"') {
      let end = i + 1;
      while (end < body.length && body[end] !== '"') {
        end += body[end] === "\\" ? 2 : 1;
      }
      const text = body.slice(i, Math.min(end + 1, body.length));

      // A string followed by a colon is a key, whatever it contains.
      const after = /^\s*:/.test(body.slice(end + 1));
      flush();
      tokens.push({ text, kind: after ? "key" : "string" });
      i = end + 1;
      continue;
    }

    const rest = body.slice(i);
    const number = /^-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?/.exec(rest);
    if (number && !/[\w.]/.test(body[i - 1] ?? "")) {
      flush();
      tokens.push({ text: number[0], kind: "number" });
      i += number[0].length;
      continue;
    }

    const literal = /^(?:true|false|null)\b/.exec(rest);
    if (literal) {
      flush();
      tokens.push({ text: literal[0], kind: "literal" });
      i += literal[0].length;
      continue;
    }

    plain += character;
    i += 1;
  }

  flush();
  return tokens;
}

function scanXml(body: string): Token[] {
  const tokens: Token[] = [];
  const pattern = /<[^>]*>/g;
  let last = 0;
  let match: RegExpExecArray | null;

  while ((match = pattern.exec(body)) !== null) {
    if (match.index > last) {
      tokens.push({ text: body.slice(last, match.index), kind: "text" });
    }

    const tag = match[0];
    const attributes = /([\w:.-]+)(\s*=\s*)("[^"]*"|'[^']*')/g;
    let cursor = 0;
    let attribute: RegExpExecArray | null;

    while ((attribute = attributes.exec(tag)) !== null) {
      tokens.push({ text: tag.slice(cursor, attribute.index), kind: "tag" });
      tokens.push({ text: attribute[1], kind: "attrName" });
      tokens.push({ text: attribute[2], kind: "punctuation" });
      tokens.push({ text: attribute[3], kind: "attrValue" });
      cursor = attribute.index + attribute[0].length;
    }

    tokens.push({ text: tag.slice(cursor), kind: "tag" });
    last = match.index + tag.length;
  }

  if (last < body.length) {
    tokens.push({ text: body.slice(last), kind: "text" });
  }
  return tokens;
}

function scanForm(body: string): Token[] {
  const tokens: Token[] = [];

  body.split(/(&)/).forEach((part) => {
    if (part === "&") {
      tokens.push({ text: part, kind: "punctuation" });
      return;
    }
    const equals = part.indexOf("=");
    if (equals < 0) {
      tokens.push({ text: part, kind: "text" });
      return;
    }
    tokens.push({ text: part.slice(0, equals), kind: "key" });
    tokens.push({ text: "=", kind: "punctuation" });
    tokens.push({ text: part.slice(equals + 1), kind: "string" });
  });

  return tokens;
}

/** What syntax the body appears to be. Detection never throws; it falls back to none. */
export function detectBody(body: string): HttpMessage["body"] {
  const trimmed = body.trim();
  if (!trimmed) {
    return "none";
  }

  if (/^[[{]/.test(trimmed)) {
    try {
      JSON.parse(trimmed);
      return "json";
    } catch {
      // Truncated bodies are normal in a report; highlight them as JSON anyway.
      return /["}\]]/.test(trimmed) ? "json" : "none";
    }
  }
  if (/^<[a-zA-Z?!/]/.test(trimmed)) {
    return "xml";
  }
  if (/^[^=&\s]+=[^&]*(?:&[^=&\s]*=[^&]*)*$/.test(trimmed)) {
    return "form";
  }
  return "none";
}

/** Break a flat token list back into lines, keeping each token's kind. */
function intoLines(tokens: Token[]): Token[][] {
  const lines: Token[][] = [[]];

  for (const token of tokens) {
    const parts = token.text.split("\n");
    parts.forEach((part, i) => {
      if (i > 0) {
        lines.push([]);
      }
      if (part) {
        lines[lines.length - 1].push({ text: part, kind: token.kind });
      }
    });
  }
  return lines;
}

/**
 * Tokenise a pasted HTTP message: start line, headers, then the body in whatever syntax
 * it turns out to be. Anything unrecognised stays plain text rather than failing.
 */
export function highlightHttp(raw: string): HttpMessage {
  const lines = toLines(raw.replace(/\s+$/, ""));
  const first = lines.length > 0 ? startLine(lines[0]) : undefined;
  const result: Token[][] = [];

  if (!first) {
    // Not an HTTP message we recognise: still a code block, just uncoloured.
    return {
      kind: "unknown",
      body: "none",
      lines: lines.map((line) => (line ? [{ text: line, kind: "text" as TokenKind }] : [])),
    };
  }

  result.push(first.tokens);

  let i = 1;
  for (; i < lines.length && lines[i].trim() !== ""; i++) {
    result.push(headerLine(lines[i]));
  }

  const body = lines.slice(i + 1).join("\n");
  const syntax = detectBody(body);

  if (i < lines.length) {
    result.push([]);
  }

  if (body) {
    const scan =
      syntax === "json"
        ? scanJson
        : syntax === "xml"
          ? scanXml
          : syntax === "form"
            ? scanForm
            : null;
    const tokens = scan ? scan(body) : [{ text: body, kind: "text" as TokenKind }];
    result.push(...intoLines(tokens));
  }

  return { kind: first.kind, body: syntax, lines: result };
}
