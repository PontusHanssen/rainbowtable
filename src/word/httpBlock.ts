import { HttpMessage, TokenKind, highlightHttp } from "./http";
import { planFromHttp } from "./documentPlan";
import { WriteResult, removeWritten, writePlan } from "./writePlan";

/**
 * Colours are chosen to read on the template's Codeblock shading (D7D2CB, a warm grey),
 * not on white, and to stay legible in greyscale print — hue alone never carries meaning.
 */
const COLOURS: Record<TokenKind, { colour: string; bold?: boolean }> = {
  method: { colour: "0A3069", bold: true },
  target: { colour: "24292F" },
  version: { colour: "4C4C4C" },
  status2xx: { colour: "1A7F37", bold: true },
  status3xx: { colour: "0A3069", bold: true },
  status4xx: { colour: "9A3412", bold: true },
  status5xx: { colour: "A40E26", bold: true },
  reason: { colour: "4C4C4C" },
  headerName: { colour: "5A2D8C", bold: true },
  punctuation: { colour: "4C4C4C" },
  headerValue: { colour: "24292F" },
  key: { colour: "0A3069" },
  string: { colour: "032F62" },
  number: { colour: "953800" },
  literal: { colour: "A40E26" },
  tag: { colour: "1A7F37" },
  attrName: { colour: "5A2D8C" },
  attrValue: { colour: "032F62" },
  text: { colour: "24292F" },
};

/** The colour and weight a token kind is written in. Exported for the preview to share. */
export function colourOf(kind: string): { colour: string; bold?: boolean } {
  return COLOURS[kind as TokenKind] ?? COLOURS.text;
}

const BLOCK_BOOKMARK = "_pthttp";

export interface HttpBlockResult extends WriteResult {
  kind: HttpMessage["kind"];
  body: HttpMessage["body"];
  lines: number;
}

/**
 * Insert a highlighted HTTP message at the cursor, through the API rather than OOXML.
 *
 * Cost here scales with the number of runs rather than being a flat six seconds: a typical
 * response is far quicker this way, though a very large and heavily tokenised one has more
 * runs to write. See the measurements in CLAUDE.md.
 */
export async function insertHttpBlock(raw: string): Promise<HttpBlockResult> {
  if (!raw.trim()) {
    throw new Error("Paste an HTTP request or response first.");
  }

  const message = highlightHttp(raw);
  const written = await writePlan(planFromHttp(message, colourOf), BLOCK_BOOKMARK);

  return { ...written, kind: message.kind, body: message.body, lines: message.lines.length };
}

/** Remove a block inserted by `insertHttpBlock`. */
export async function removeHttpBlock(bookmark: string): Promise<void> {
  return removeWritten(bookmark);
}
