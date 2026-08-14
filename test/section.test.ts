import test from "node:test";
import assert from "node:assert/strict";
import { childHeadings, toHeadings } from "../src/word/headings";
import { buildBlocks, findingSummaries } from "../src/word/section";
import { FixtureParagraph, filledParagraphs } from "./fixtures/template";

/** The blocks of a named section, as the findings features compute them. */
function blocksOf(paragraphs: FixtureParagraph[], name: string, level = 1) {
  const headings = toHeadings(paragraphs);
  const position = headings.findIndex(
    (heading) => heading.text === name && heading.level === level
  );
  return buildBlocks(headings, position, childHeadings(headings, position), paragraphs.length);
}

test("findingSummaries reports each finding's title, severity and score", () => {
  const summaries = findingSummaries(blocksOf(filledParagraphs, "Findings"));

  assert.deepEqual(summaries, [
    { title: "XSS in Y", severity: "Medium", score: 5.4, reason: undefined },
    { title: "SQLi", severity: "Critical", score: 9.0, reason: undefined },
    { title: "SSRF A", severity: "Medium", score: 4.0, reason: undefined },
  ]);
});

test("findingSummaries keeps the order it is given, so it can show a planned sort", () => {
  const blocks = blocksOf(filledParagraphs, "Findings");
  const reversed = [...blocks].reverse();

  assert.deepEqual(
    findingSummaries(reversed).map((summary) => summary.title),
    ["SSRF A", "SQLi", "XSS in Y"]
  );
});

/**
 * A finding nobody can rate is the case the preview exists for: it has to appear in the
 * list, named, with the reason showing, rather than being counted and dropped.
 */
test("findingSummaries carries the reason a finding could not be read", () => {
  const paragraphs: FixtureParagraph[] = [
    { text: "Findings", styleBuiltIn: "Heading1" },
    { text: "Half-written finding", styleBuiltIn: "Heading2" },
    { text: "Risk: [TODO]", styleBuiltIn: "Heading3" },
  ];

  const [summary] = findingSummaries(blocksOf(paragraphs, "Findings"));

  assert.equal(summary.title, "Half-written finding");
  assert.equal(summary.severity, undefined);
  assert.equal(summary.score, undefined);
  assert.ok(summary.reason, "an unreadable finding must say why");
});

test("findingSummaries leaves the score out when the heading carries none", () => {
  const paragraphs: FixtureParagraph[] = [
    { text: "Findings", styleBuiltIn: "Heading1" },
    { text: "Unscored", styleBuiltIn: "Heading2" },
    { text: "Risk: High", styleBuiltIn: "Heading3" },
  ];

  assert.deepEqual(findingSummaries(blocksOf(paragraphs, "Findings")), [
    { title: "Unscored", severity: "High", score: undefined, reason: undefined },
  ]);
});
