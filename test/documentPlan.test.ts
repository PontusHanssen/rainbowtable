import test from "node:test";
import assert from "node:assert/strict";
import { ParagraphPlan, listGroups, planFromBlocks } from "../src/word/documentPlan";
import { parseMarkdown } from "../src/word/markdown";

const plan = (source: string) => planFromBlocks(parseMarkdown(source));
const kinds = (source: string) => plan(source).map((step) => step.kind);

test("headings map straight across, # to level 1", () => {
  const steps = plan("# One\n\n## Two\n\n### Three");

  assert.deepEqual(
    steps.map((step) => (step.kind === "heading" ? step.level : null)),
    [1, 2, 3]
  );
});

test("headings stop at the six levels markdown has", () => {
  const [step] = plan("###### Six");
  assert.equal(step.kind === "heading" && step.level, 6);
});

test("emphasis, code and links become formatted runs", () => {
  const [step] = plan("plain **bold** *italic* `code` <https://x.test>");
  const runs = step.kind === "body" ? step.runs : [];

  assert.deepEqual(runs, [
    { text: "plain " },
    { text: "bold", bold: true },
    { text: " " },
    { text: "italic", italic: true },
    { text: " " },
    { text: "code", code: true },
    { text: " " },
    { text: "https://x.test", link: "https://x.test" },
  ]);
});

test("a fenced block becomes one code paragraph per line, verbatim", () => {
  const steps = plan("```\n<script>a & b</script>\n\n# not a heading\n```");

  assert.deepEqual(
    steps.map((step) => (step.kind === "code" ? step.text : step.kind)),
    ["<script>a & b</script>", "", "# not a heading"]
  );
});

test("consecutive list items of one kind share a group", () => {
  const steps = plan("- a\n- b\n- c");
  const groups = steps.map((step) => (step.kind === "listItem" ? step.group : null));

  assert.deepEqual(groups, [1, 1, 1], "one list, or each item restarts at 1");
});

test("switching between bullets and numbers starts a new list", () => {
  const steps = plan("- a\n- b\n\n1. one\n2. two");
  const listed = steps.filter(
    (step): step is Extract<ParagraphPlan, { kind: "listItem" }> => step.kind === "listItem"
  );

  assert.deepEqual(
    listed.map((step) => [step.group, step.ordered]),
    [
      [1, false],
      [1, false],
      [2, true],
      [2, true],
    ]
  );
});

test("a paragraph between two lists separates them", () => {
  const steps = plan("- a\n\ntext\n\n- b");
  const listed = steps.filter(
    (step): step is Extract<ParagraphPlan, { kind: "listItem" }> => step.kind === "listItem"
  );

  assert.notEqual(listed[0].group, listed[1].group, "the second list numbers from 1 again");
});

test("listGroups gathers each list's paragraph positions", () => {
  const steps = plan("# T\n\n- a\n- b\n\ntext\n\n1. one");
  const groups = listGroups(steps);

  assert.equal(groups.size, 2);
  assert.deepEqual([...groups.values()], [[1, 2], [4]]);
});

test("ordinary prose produces no lists to attach", () => {
  assert.equal(listGroups(plan("# T\n\njust text")).size, 0);
});

test("the plan covers every block kind markdown can produce", () => {
  assert.deepEqual(kinds("# h\n\ntext\n\n- b\n\n1. n\n\n```\nc\n```"), [
    "heading",
    "body",
    "listItem",
    "listItem",
    "code",
  ]);
});
