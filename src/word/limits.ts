import { escapeXml, wrapInPackage } from "./ooxml";

/**
 * Measuring where large insertions stop working.
 *
 * Two hops carry a finding from the dialog into the document, and neither has a documented
 * size limit: `messageParent` from the dialog to the task pane, and `insertOoxml` from the
 * pane into Word. This module generates the material to push both until they break, and
 * the pane and dialog drive it. Everything here is pure so the ladder and the payloads can
 * be tested without Word.
 */

/** Sizes to try, doubling, so a run costs a handful of attempts rather than a search. */
export function sizeLadder(from: number, to: number): number[] {
  const sizes: number[] = [];
  for (let size = from; size <= to; size *= 2) {
    sizes.push(size);
  }
  return sizes;
}

/**
 * A payload of roughly `bytes`, shaped like the content that actually gets sent: a finding
 * whose evidence is a long fenced block. Padding with random-ish text rather than one
 * repeated character keeps any compression along the way from flattering the result.
 */
export function syntheticFinding(bytes: number): string {
  const header = ["## Size test", "", "### Proof of Concept", "", "```"].join("\n");
  const footer = "\n```\n";
  const line = (n: number) =>
    `    {"id": ${n}, "token": "${((n * 2654435761) >>> 0).toString(36)}", "note": "line ${n}"},`;

  const body: string[] = [];
  let length = header.length + footer.length;
  for (let n = 0; length < bytes; n++) {
    const next = line(n);
    body.push(next);
    length += next.length + 1;
  }

  return `${header}\n${body.join("\n")}${footer}`;
}

/** A plain string of about `bytes`, for measuring the message channel on its own. */
export function syntheticMessage(bytes: number): string {
  const chunk = "0123456789abcdef";
  return chunk.repeat(Math.ceil(bytes / chunk.length)).slice(0, bytes);
}

export interface Measurement {
  bytes: number;
  ok: boolean;
  /** Milliseconds the attempt took, whether it succeeded or not. */
  ms: number;
  /** Why it failed, or what came back wrong. */
  detail?: string;
}

/** A run's results as a table, for pasting into the notes. */
export function formatMeasurements(title: string, results: Measurement[]): string {
  const rows = results.map((result) => {
    const size =
      result.bytes >= 1024 ? `${(result.bytes / 1024).toFixed(0)} KB` : `${result.bytes} B`;
    const outcome = result.ok ? "ok" : `FAILED — ${result.detail ?? "no detail"}`;
    return `  ${size.padStart(8)}  ${String(result.ms).padStart(6)} ms  ${outcome}`;
  });

  const lastOk = [...results].reverse().find((result) => result.ok);
  const summary = lastOk
    ? `largest that worked: ${(lastOk.bytes / 1024).toFixed(0)} KB`
    : "nothing succeeded";

  return [`${title} — ${summary}`, ...rows].join("\n");
}

/**
 * How a code block is laid out in OOXML. Inserting is slow in Word on the web, and the
 * question these answer is what it is slow *at* — the bytes, the paragraphs, or resolving
 * a style for each of them.
 */
export type Shape = "paragraphs" | "breaks" | "unstyled";

const CODE_STYLE_DEFINITION =
  '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Codeblock">' +
  '<w:name w:val="Code block"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:shd w:val="clear" w:color="auto" w:fill="D7D2CB"/><w:contextualSpacing/></w:pPr>' +
  '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="20"/></w:rPr></w:style>';

/**
 * The same lines, three ways:
 *
 * - `paragraphs` — one styled paragraph per line, what the add-in does today.
 * - `breaks` — a single styled paragraph, lines separated by `w:br`. Same bytes, one
 *   paragraph instead of hundreds, and the same single shaded box on the page.
 * - `unstyled` — one paragraph per line with no style, to separate the cost of creating
 *   paragraphs from the cost of resolving a style for each.
 */
export function codeBlockPackage(lines: string[], shape: Shape): string {
  const text = (line: string) => `<w:t xml:space="preserve">${escapeXml(line)}</w:t>`;

  if (shape === "breaks") {
    const runs = lines.map(text).join("<w:br/>");
    return wrapInPackage(
      `<w:p><w:pPr><w:pStyle w:val="Codeblock"/></w:pPr><w:r>${runs}</w:r></w:p><w:p/>`,
      CODE_STYLE_DEFINITION
    );
  }

  const pPr = shape === "paragraphs" ? '<w:pPr><w:pStyle w:val="Codeblock"/></w:pPr>' : "";
  const body = lines.map((line) => `<w:p>${pPr}<w:r>${text(line)}</w:r></w:p>`).join("");

  return wrapInPackage(`${body}<w:p/>`, shape === "paragraphs" ? CODE_STYLE_DEFINITION : undefined);
}

/** The lines a synthetic finding of `bytes` would contain, for the shape comparison. */
export function syntheticLines(bytes: number): string[] {
  return syntheticFinding(bytes)
    .split("\n")
    .filter((line) => !line.startsWith("```") && !line.startsWith("#") && line.trim() !== "");
}
