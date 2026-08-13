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
