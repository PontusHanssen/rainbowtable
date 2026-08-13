import {
  Measurement,
  Shape,
  codeBlockPackage,
  formatMeasurements,
  sizeLadder,
  syntheticFinding,
  syntheticLines,
} from "../word/limits";
import { insertMarkdown, removeMarkdown } from "../word/markdownDoc";
import { insertProbePackage } from "../word/limitsProbe";
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

  const shapes = byId<HTMLButtonElement>("limits-shapes");
  const shapeOutput = byId("limits-shape-output");

  /**
   * The same content three ways at one size. Inserting is slow; this says what it is slow
   * at, which decides whether the fix is fewer paragraphs or fewer bytes.
   */
  shapes.onclick = async () => {
    shapes.disabled = true;
    const lines = syntheticLines(32 * 1024);
    const report: string[] = [`32 KB of code, ${lines.length} lines`];

    try {
      for (const shape of ["paragraphs", "breaks", "unstyled"] as Shape[]) {
        const started = Date.now();
        try {
          const bookmark = await insertProbePackage(codeBlockPackage(lines, shape));
          const ms = Date.now() - started;
          report.push(`  ${shape.padEnd(12)} ${String(ms).padStart(6)} ms`);
          shapeOutput.textContent = report.join("\n");
          await removeMarkdown(bookmark);
        } catch (err) {
          report.push(`  ${shape.padEnd(12)} FAILED — ${String(err)}`);
          shapeOutput.textContent = report.join("\n");
        }
      }
    } finally {
      shapeOutput.textContent = report.join("\n");
      shapes.disabled = false;
    }
  };

  run.onclick = async () => {
    run.disabled = true;
    const results: Measurement[] = [];

    try {
      // From 4 KB: a finding with a screen of evidence. 64 KB already took 15 seconds,
      // so the interesting part of the curve is below where the first run started.
      for (const bytes of sizeLadder(4 * 1024, 512 * 1024)) {
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
