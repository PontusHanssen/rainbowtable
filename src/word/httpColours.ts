import { TokenKind } from "./http";

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
