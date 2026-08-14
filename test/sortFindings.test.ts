import test from "node:test";
import assert from "node:assert/strict";
import { childHeadings, findSections, toHeadings } from "../src/word/headings";
import { parseRisk } from "../src/word/severity";
import { buildBlocks } from "../src/word/section";
import { planOrder, unchanged } from "../src/word/sortFindings";
import { FixtureParagraph, filledParagraphs, templateParagraphs } from "./fixtures/template";

/**
 * The blocks of the section named `name`, as sortFindings computes them. The level is
 * part of the lookup because a report can repeat a heading title at different depths —
 * the template has "Vulnerabilities" as both a summary subsection and a findings section.
 */
function blocksOf(paragraphs: FixtureParagraph[], name: string, level = 1) {
  const headings = toHeadings(paragraphs);
  const position = headings.findIndex(
    (heading) => heading.text === name && heading.level === level
  );
  return buildBlocks(headings, position, childHeadings(headings, position), paragraphs.length);
}

test("the worked example from instructions.md", () => {
  const blocks = blocksOf(filledParagraphs, "Findings");

  assert.deepEqual(
    planOrder(blocks).map((block) => block.heading.text),
    ["SQLi", "XSS in Y", "SSRF A"]
  );
});

test("planOrder is a no-op when the findings are already sorted", () => {
  const items = [
    { id: "a", risk: parseRisk("Risk: Critical (9.0)") },
    { id: "b", risk: parseRisk("Risk: Medium (5.4)") },
  ];
  assert.deepEqual(planOrder(items), items);
});

test("scored findings precede unscored ones of the same severity, stably", () => {
  const items = [
    { id: "a", risk: parseRisk("Risk: High") },
    { id: "b", risk: parseRisk("Risk: High (8.0)") },
    { id: "c", risk: parseRisk("Risk: High") },
  ];
  assert.deepEqual(
    planOrder(items).map((item) => item.id),
    ["b", "a", "c"]
  );
});

test("unreadable findings hold their slot while the rest sort around them", () => {
  const items = [
    { id: "low", risk: parseRisk("Risk: Low (1.0)") },
    { id: "todo", risk: undefined },
    { id: "crit", risk: parseRisk("Risk: Critical (9.9)") },
  ];
  assert.deepEqual(
    planOrder(items).map((item) => item.id),
    ["crit", "todo", "low"]
  );
});

test("a block covers its finding's heading through to the next finding", () => {
  const blocks = blocksOf(filledParagraphs, "Findings");

  // Each finding in the fixture is a heading plus seven paragraphs of content.
  assert.deepEqual(
    blocks.map(({ heading, start, end }) => [heading.text, start, end]),
    [
      ["XSS in Y", 1, 8],
      ["SQLi", 9, 16],
      ["SSRF A", 17, 24],
    ]
  );
});

test("the last block stops at the end of its section, not the end of the document", () => {
  const blocks = blocksOf(templateParagraphs, "Vulnerabilities");
  const headings = toHeadings(templateParagraphs);
  const weaknesses = headings.find((heading) => heading.text === "Weaknesses");

  assert.equal(blocks.length, 1);
  assert.equal(blocks[0].end, (weaknesses as { index: number }).index - 1);
});

test("the last block of a trailing section reaches the end of the document", () => {
  const paragraphs: FixtureParagraph[] = [
    { styleBuiltIn: "Heading1", text: "Findings" },
    { styleBuiltIn: "Heading2", text: "SQLi" },
    { styleBuiltIn: "Heading3", text: "Risk: Critical (9.0)" },
    { styleBuiltIn: "Normal", text: "last paragraph" },
  ];

  assert.equal(blocksOf(paragraphs, "Findings")[0].end, paragraphs.length - 1);
});

test("the template's placeholder risk headings are reported, not guessed at", () => {
  const blocks = blocksOf(templateParagraphs, "Weaknesses");

  assert.equal(blocks.length, 4);
  for (const block of blocks) {
    assert.equal(block.risk, undefined);
    assert.match(block.skipReason as string, /"Risk: \[TODO\]" is not of the form/);
  }
});

test("a finding with no risk heading at all says so", () => {
  const paragraphs: FixtureParagraph[] = [
    { styleBuiltIn: "Heading1", text: "Findings" },
    { styleBuiltIn: "Heading2", text: "Undocumented" },
    { styleBuiltIn: "Heading3", text: "Description" },
  ];

  assert.equal(blocksOf(paragraphs, "Findings")[0].skipReason, 'no "Risk:" heading');
});

test("a section whose findings all parse produces no skips", () => {
  const blocks = blocksOf(filledParagraphs, "Findings");

  assert.deepEqual(
    blocks.map((block) => block.skipReason),
    [undefined, undefined, undefined]
  );
  assert.deepEqual(
    blocks.map((block) => block.risk?.severity),
    ["Medium", "Critical", "Medium"]
  );
});

test("findSections and buildBlocks agree on what a section's findings are", () => {
  const sections = findSections(toHeadings(filledParagraphs));
  const section = sections.find((candidate) => candidate.heading.text === "Findings");

  assert.deepEqual(
    section?.findings.map((finding) => finding.text),
    blocksOf(filledParagraphs, "Findings").map((block) => block.heading.text)
  );
});

/**
 * Undo replaces the whole findings region with the snapshot, so anything written since
 * would go with it. The sort records how it left the region and the undo compares.
 */
test("unchanged recognises a region nobody has touched", () => {
  const written = ["Finding A", "Risk: High (7.5)", "Description", "Finding B"];

  assert.equal(unchanged(written, [...written]), true);
  assert.equal(unchanged(written, written.slice(0, 3)), false, "a paragraph removed");
  assert.equal(unchanged(written, [...written, "New evidence"]), false, "a paragraph added");
  assert.equal(
    unchanged(written, ["Finding A", "Risk: High (7.5)", "Description, rewritten", "Finding B"]),
    false,
    "a paragraph edited"
  );
  assert.equal(unchanged([], []), true);
});
