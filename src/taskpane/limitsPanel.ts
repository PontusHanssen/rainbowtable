import { Measurement, formatMeasurements, sizeLadder, syntheticFinding } from "../word/limits";
import { insertMarkdown, removeMarkdown } from "../word/markdownDoc";
import { byId, show } from "./dom";

/* global HTMLButtonElement, location */

/**
 * The development-only limits harness.
 *
 * Answers a question the Office documentation does not: how large a finding can actually
 * be inserted. Shown only on localhost, alongside the dev badge — it writes test content
 * into the document and removes it again, which is not something to offer in a report.
 */
export function setUpLimitsPanel(): void {
  const panel = byId("limits");
  const run = byId<HTMLButtonElement>("limits-run");
  const output = byId("limits-output");

  show(panel, location.hostname === "localhost");

  run.onclick = async () => {
    run.disabled = true;
    const results: Measurement[] = [];

    try {
      // From 64 KB, where a real finding with evidence sits, upwards until it breaks.
      for (const bytes of sizeLadder(64 * 1024, 8 * 1024 * 1024)) {
        const finding = syntheticFinding(bytes);
        const started = Date.now();

        try {
          const inserted = await insertMarkdown(finding);
          const ms = Date.now() - started;
          results.push({ bytes: finding.length, ok: true, ms });

          // Take it straight back out: this is a measurement, not an edit.
          await removeMarkdown(inserted.bookmark);
        } catch (err) {
          results.push({
            bytes: finding.length,
            ok: false,
            ms: Date.now() - started,
            detail: String(err),
          });
          break;
        }

        output.textContent = formatMeasurements("insertOoxml", results);
      }
    } finally {
      output.textContent = formatMeasurements("insertOoxml", results);
      run.disabled = false;
    }
  };
}
