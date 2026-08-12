import test from "node:test";
import assert from "node:assert/strict";
import { Block, outline, parseInline, parseMarkdown, spanText } from "../src/word/markdown";

const kinds = (source: string) => parseMarkdown(source).map((block) => block.kind);
const first = (source: string): Block => parseMarkdown(source)[0];

test("headings keep their level and text", () => {
  const blocks = parseMarkdown("# Title\n\n## Description\n\n###### Deep");

  assert.deepEqual(
    blocks.map((block) => (block.kind === "heading" ? [block.level, spanText(block.spans)] : null)),
    [
      [1, "Title"],
      [2, "Description"],
      [6, "Deep"],
    ]
  );
});

test("a run of lines is one paragraph, a blank line starts another", () => {
  const blocks = parseMarkdown("one\ntwo\n\nthree");

  assert.deepEqual(kinds("one\ntwo\n\nthree"), ["paragraph", "paragraph"]);
  assert.equal(spanText((blocks[0] as Extract<Block, { kind: "paragraph" }>).spans), "one two");
});

test("bullet and numbered lists are recognised in their usual spellings", () => {
  assert.deepEqual(kinds("- a\n* b\n+ c"), ["bullet", "bullet", "bullet"]);
  assert.deepEqual(kinds("1. a\n2) b"), ["number", "number"]);
});

test("emphasis at the start of a line is not mistaken for a bullet", () => {
  assert.deepEqual(kinds("*italic* text"), ["paragraph"]);
  assert.deepEqual(kinds("**bold** text"), ["paragraph"]);
});

test("fenced code keeps its lines verbatim, including blanks and markdown", () => {
  const block = first("```bash\n$ curl -s http://x\n\n# not a heading\n```");

  assert.equal(block.kind, "code");
  assert.deepEqual((block as Extract<Block, { kind: "code" }>).lines, [
    "$ curl -s http://x",
    "",
    "# not a heading",
  ]);
  assert.equal((block as Extract<Block, { kind: "code" }>).language, "bash");
});

test("an unclosed fence still yields a code block rather than swallowing the input", () => {
  const block = first("```\nstill code");
  assert.equal(block.kind, "code");
  assert.deepEqual((block as Extract<Block, { kind: "code" }>).lines, ["still code"]);
});

test("inline spans are told apart", () => {
  assert.deepEqual(parseInline("plain **bold** *italic* `code` end"), [
    { kind: "text", text: "plain " },
    { kind: "bold", text: "bold" },
    { kind: "text", text: " " },
    { kind: "italic", text: "italic" },
    { kind: "text", text: " " },
    { kind: "code", text: "code" },
    { kind: "text", text: " end" },
  ]);
});

test("bold wins over italic, so **x** is not two emphases", () => {
  assert.deepEqual(parseInline("**x**"), [{ kind: "bold", text: "x" }]);
});

test("code spans are never re-read as emphasis", () => {
  // A shell glob inside code must not turn into italics.
  assert.deepEqual(parseInline("`ls *.txt *.md`"), [{ kind: "code", text: "ls *.txt *.md" }]);
});

test("autolinks become links carrying their own text", () => {
  assert.deepEqual(parseInline("see <https://example.com/a?b=1>"), [
    { kind: "text", text: "see " },
    { kind: "link", text: "https://example.com/a?b=1", url: "https://example.com/a?b=1" },
  ]);
  assert.deepEqual(parseInline("<mailto:a@b.c>"), [
    { kind: "link", text: "mailto:a@b.c", url: "mailto:a@b.c" },
  ]);
});

test("angle brackets that are not links stay literal", () => {
  assert.deepEqual(parseInline("<not a link>"), [{ kind: "text", text: "<not a link>" }]);
  assert.deepEqual(parseInline("a < b"), [{ kind: "text", text: "a < b" }]);
});

test("no markdown means one plain paragraph, unchanged", () => {
  assert.deepEqual(parseInline("just words"), [{ kind: "text", text: "just words" }]);
});

test("the outline counts what the preview reports", () => {
  const summary = outline(
    parseMarkdown(
      [
        "# SQL injection",
        "Risk: High (7.8)",
        "",
        "## Description",
        "The `q` parameter, see <https://owasp.org>.",
        "",
        "- one",
        "- two",
        "",
        "```",
        "SELECT 1",
        "```",
      ].join("\n")
    )
  );

  assert.deepEqual(
    summary.headings.map((h) => [h.level, h.text]),
    [
      [1, "SQL injection"],
      [2, "Description"],
    ]
  );
  assert.equal(summary.paragraphs, 2);
  assert.equal(summary.listItems, 2);
  assert.equal(summary.codeBlocks, 1);
  assert.equal(summary.links, 1);
});
