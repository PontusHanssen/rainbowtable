import test from "node:test";
import assert from "node:assert/strict";
import { Section } from "../src/word/headings";
import { DEFAULT_FINDING_LEVEL, findingLevelFor } from "../src/word/newFinding";

const section = (level: number): Section => ({
  heading: { index: 0, level, text: "Weaknesses" },
  findings: [],
});

test("a finding sits one level below its section", () => {
  assert.equal(findingLevelFor(section(1)), 2);
  assert.equal(findingLevelFor(section(3)), 4);
});

test("a finding can be created before any section is detectable", () => {
  // findSections only reports a section that already contains findings, so the first
  // finding in a section has to be creatable with nothing selected at all.
  assert.equal(findingLevelFor(undefined), DEFAULT_FINDING_LEVEL);
  assert.equal(DEFAULT_FINDING_LEVEL, 2, "the depth the report template uses");
});

test("the level never runs past the headings Word has", () => {
  // A finding needs a level below it for Risk, Description and the rest.
  assert.equal(findingLevelFor(section(8)), 8);
  assert.equal(findingLevelFor(section(9)), 8);
});
