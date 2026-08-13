import test from "node:test";
import assert from "node:assert/strict";
import {
  Measurement,
  formatMeasurements,
  sizeLadder,
  syntheticFinding,
  syntheticMessage,
} from "../src/word/limits";
import { parseMarkdown } from "../src/word/markdown";

test("the ladder doubles from the start up to the ceiling", () => {
  assert.deepEqual(sizeLadder(16, 256), [16, 32, 64, 128, 256]);
  assert.deepEqual(sizeLadder(1024, 1024), [1024]);
  assert.deepEqual(sizeLadder(2048, 1024), [], "nothing to try below the start");
});

test("a synthetic message is exactly the size asked for", () => {
  for (const bytes of [1, 15, 16, 1000, 65536]) {
    assert.equal(syntheticMessage(bytes).length, bytes);
  }
});

test("a synthetic finding is about the size asked for and is real markdown", () => {
  for (const bytes of [4096, 65536]) {
    const finding = syntheticFinding(bytes);
    assert.ok(finding.length >= bytes, "at least the requested size");
    assert.ok(finding.length < bytes * 1.2, "and not wildly over");

    // It has to exercise the real path, not just be a blob of text.
    const blocks = parseMarkdown(finding);
    assert.ok(
      blocks.some((block) => block.kind === "code"),
      "the bulk is a fenced block, as captured evidence would be"
    );
    assert.ok(blocks.some((block) => block.kind === "heading"));
  }
});

test("the payload does not compress to nothing, which would flatter the result", () => {
  const finding = syntheticFinding(16384);
  const distinct = new Set(finding.split("\n")).size;

  assert.ok(distinct > 100, `only ${distinct} distinct lines`);
});

test("results format as a table naming the largest size that worked", () => {
  const results: Measurement[] = [
    { bytes: 16384, ok: true, ms: 12 },
    { bytes: 32768, ok: true, ms: 40 },
    { bytes: 65536, ok: false, ms: 5, detail: "message truncated to 65535" },
  ];

  const table = formatMeasurements("messageParent", results);
  assert.match(table, /largest that worked: 32 KB/);
  assert.match(table, /FAILED — message truncated/);
});

test("a run where nothing worked says so rather than claiming a limit", () => {
  const table = formatMeasurements("insertOoxml", [
    { bytes: 1024, ok: false, ms: 3, detail: "boom" },
  ]);

  assert.match(table, /nothing succeeded/);
});
