import test from "node:test";
import assert from "node:assert/strict";
import { childHeadings, findSections, toHeadings } from "../src/word/headings";
import { templateParagraphs } from "./fixtures/template";

test("toHeadings keeps only built-in heading styles, with their paragraph index", () => {
  const headings = toHeadings([
    { styleBuiltIn: "Normal", text: "prose" },
    { styleBuiltIn: "Heading1", text: "Findings" },
    { styleBuiltIn: "Heading2", text: "SQLi" },
    { styleBuiltIn: "Codeblock", text: "curl ..." },
    { styleBuiltIn: "Heading9", text: "Deep" },
  ]);

  assert.deepEqual(headings, [
    { index: 1, level: 1, text: "Findings" },
    { index: 2, level: 2, text: "SQLi" },
    { index: 4, level: 9, text: "Deep" },
  ]);
});

test("toHeadings ignores custom heading-like styles", () => {
  // The template defines "Heading_1 No" and similar unnumbered variants. They are not
  // built-in, so they are deliberately not recognised.
  const headings = toHeadings([
    { styleBuiltIn: "Heading1No", text: "Unnumbered" },
    { styleBuiltIn: "Huvudrubrik", text: "Security Review" },
    { styleBuiltIn: "Heading1", text: "Real" },
  ]);

  assert.deepEqual(
    headings.map((heading) => heading.text),
    ["Real"]
  );
});

test("childHeadings returns direct children only, and stops at the next sibling", () => {
  const headings = toHeadings([
    { styleBuiltIn: "Heading1", text: "Findings" },
    { styleBuiltIn: "Heading2", text: "SQLi" },
    { styleBuiltIn: "Heading3", text: "Risk: Critical (9.0)" },
    { styleBuiltIn: "Heading2", text: "XSS" },
    { styleBuiltIn: "Heading1", text: "Appendix" },
    { styleBuiltIn: "Heading2", text: "Not a finding" },
  ]);

  assert.deepEqual(
    childHeadings(headings, 0).map((heading) => heading.text),
    ["SQLi", "XSS"]
  );
});

test("the template's outline yields the findings sections", () => {
  const sections = findSections(toHeadings(templateParagraphs));
  const byName = new Map(sections.map((section) => [section.heading.text, section]));

  const vulnerabilities = byName.get("Vulnerabilities");
  assert.ok(vulnerabilities, "Vulnerabilities is offered as a section");
  assert.equal(vulnerabilities.heading.level, 1);
  assert.equal(vulnerabilities.findings.length, 1);

  const weaknesses = byName.get("Weaknesses");
  assert.equal(weaknesses?.findings.length, 4);
});

test("a heading with no grandchildren is not offered as a section", () => {
  // "Results" has subsections, but they hold prose rather than findings.
  const sections = findSections(toHeadings(templateParagraphs));
  assert.ok(!sections.some((section) => section.heading.text === "Results"));
});

test("findings are not tied to a fixed heading depth", () => {
  const headings = toHeadings([
    { styleBuiltIn: "Heading1", text: "Report" },
    { styleBuiltIn: "Heading2", text: "Findings" },
    { styleBuiltIn: "Heading3", text: "SSRF" },
    { styleBuiltIn: "Heading4", text: "Risk: Low (3.0)" },
    { styleBuiltIn: "Heading3", text: "SQLi" },
    { styleBuiltIn: "Heading4", text: "Risk: Critical (9.0)" },
  ]);

  const section = findSections(headings).find((candidate) => candidate.heading.text === "Findings");
  assert.equal(section?.heading.level, 2);
  assert.deepEqual(
    section?.findings.map((f) => f.text),
    ["SSRF", "SQLi"]
  );
});
