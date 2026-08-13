import test from "node:test";
import assert from "node:assert/strict";
import { applyScore } from "../src/dialog/applyScore";
import { parseRisk } from "../src/word/severity";
import { parseMarkdown } from "../src/word/markdown";

const SKELETON = ["## Title", "", "### Risk:", "", "### Technical Details", "", "Some prose."].join(
  "\n"
);

const LINK =
  "<https://nvd.nist.gov/vuln-metrics/cvss/v3-calculator?vector=CVSS:3.1/AV:N&version=3.1>";

test("the score fills in the existing Risk heading, keeping its level", () => {
  const result = applyScore(SKELETON, "High (7.8)", LINK);

  assert.ok(result.includes("### Risk: High (7.8)"));
  assert.ok(!result.includes("### Risk:\n"), "the empty heading is gone");
});

test("what it writes is what the report parser reads back", () => {
  // The whole point: the calculator has to produce a heading severity.ts accepts.
  const result = applyScore(SKELETON, "Critical (9.8)", LINK);
  const heading = result.split("\n").find((line) => line.includes("Risk:")) ?? "";

  assert.deepEqual(parseRisk(heading.replace(/^#+\s*/, "")), {
    severity: "Critical",
    score: 9.8,
  });
});

test("the vector goes on the line below, as an autolink", () => {
  const blocks = parseMarkdown(applyScore(SKELETON, "Low (2.0)", LINK));
  const links = blocks.flatMap((block) =>
    "spans" in block ? block.spans.filter((span) => span.kind === "link") : []
  );

  assert.equal(links.length, 1, "one clickable vector");
  assert.match(links[0].text, /^https:\/\/nvd\.nist\.gov/);
});

test("rescoring replaces the vector rather than stacking another", () => {
  const once = applyScore(SKELETON, "Low (2.0)", LINK);
  const twice = applyScore(once, "High (7.5)", LINK);

  assert.equal(twice.match(/nvd\.nist\.gov/g)?.length, 1);
  assert.ok(twice.includes("### Risk: High (7.5)"));
  assert.ok(!twice.includes("Low (2.0)"));
});

test("a finding with no Risk heading gains one rather than losing the score", () => {
  const result = applyScore("## Title\n\nJust prose.", "Medium (5.4)", LINK);

  assert.ok(result.includes("### Risk: Medium (5.4)"));
  assert.ok(result.includes("nvd.nist.gov"));
  assert.ok(result.startsWith("## Title"), "the finding is left intact");
});

test("a heading at another depth keeps that depth", () => {
  const result = applyScore("# Finding\n\n## Risk:\n", "High (8.1)", LINK);
  assert.ok(result.includes("## Risk: High (8.1)"), "not forced to ###");
});
