import test from "node:test";
import assert from "node:assert/strict";
import { CODE_COLOURS, canonicalLanguage, colourFor } from "../src/dialog/codeColours";

test("the aliases people actually type resolve to a grammar", () => {
  assert.equal(canonicalLanguage("js"), "javascript");
  assert.equal(canonicalLanguage("TypeScript"), "javascript");
  assert.equal(canonicalLanguage("bash"), "shell");
  assert.equal(canonicalLanguage("yml"), "yaml");
  assert.equal(canonicalLanguage("psql"), "sql");
});

test("an unknown or absent language stays plain rather than guessing", () => {
  assert.equal(canonicalLanguage("brainfuck"), undefined);
  assert.equal(canonicalLanguage(""), undefined);
  assert.equal(canonicalLanguage(undefined), undefined);
});

test("http is not a grammar here — highlightHttp handles it", () => {
  // A general grammar has no idea what a status line is.
  assert.equal(canonicalLanguage("http"), undefined);
});

test("every token class has a colour, and they are distinguishable", () => {
  const colours = Object.values(CODE_COLOURS).map((entry) => entry.colour);

  assert.ok(colours.every((colour) => /^[0-9A-F]{6}$/.test(colour)));
  assert.ok(new Set(colours).size >= 6, "not all the same shade");
  assert.equal(colourFor("plain").colour, CODE_COLOURS.plain.colour);
});

test("keywords carry weight, so the block still reads in greyscale", () => {
  assert.ok(CODE_COLOURS.keyword.bold, "hue alone must not carry the meaning");
});

test("highlighting never alters the code, only colours it", async () => {
  // The same invariant the HTTP block has: what is written is evidence.
  const { StandardSQL } = await import("@codemirror/lang-sql");
  const { highlightWith } = await import("../src/dialog/highlightCode");

  const sql = "SELECT id, name\nFROM users\nWHERE name = 'admin' -- injected\n";
  const rebuilt = highlightWith(StandardSQL.language, sql)
    .map((line) => line.map((run) => run.text).join(""))
    .join("\n");

  assert.equal(rebuilt, sql);
});

test("a grammar colours what it recognises", async () => {
  const { StandardSQL } = await import("@codemirror/lang-sql");
  const { highlightWith } = await import("../src/dialog/highlightCode");
  const { CODE_COLOURS: colours } = await import("../src/dialog/codeColours");

  const runs = highlightWith(StandardSQL.language, "SELECT * FROM users WHERE id = 1").flat();
  const find = (text: string) => runs.find((run) => run.text.trim() === text);

  assert.equal(find("SELECT")?.colour, colours.keyword.colour, "keywords stand out");
  assert.equal(find("1")?.colour, colours.number.colour, "numbers do too");
  assert.ok(
    runs.every((run) => /^[0-9A-F]{6}$/.test(run.colour ?? "")),
    "no run reaches the document without a colour"
  );
});

test("a span crossing a newline is cut at the line, since lines are paragraphs", async () => {
  const { StandardSQL } = await import("@codemirror/lang-sql");
  const { highlightWith } = await import("../src/dialog/highlightCode");

  const lines = highlightWith(StandardSQL.language, "SELECT 'a\nmultiline'\nFROM t");

  assert.equal(lines.length, 3, "one array of runs per line");
  assert.ok(lines.every((line) => line.every((run) => !run.text.includes("\n"))));
});
