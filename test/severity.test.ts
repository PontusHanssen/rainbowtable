import test from "node:test";
import assert from "node:assert/strict";
import { compareRisk, isRiskHeading, parseRisk } from "../src/word/severity";

test("parseRisk accepts the documented form", () => {
  assert.deepEqual(parseRisk("Risk: Medium (5.4)"), { severity: "Medium", score: 5.4 });
  assert.deepEqual(parseRisk("Risk: Critical (9.0)"), { severity: "Critical", score: 9 });
  assert.deepEqual(parseRisk("Risk: Low (2.1)"), { severity: "Low", score: 2.1 });
  assert.deepEqual(parseRisk("Risk: Informational"), { severity: "Informational" });
});

test("parseRisk tolerates case, padding, nbsp and a decimal comma", () => {
  assert.deepEqual(parseRisk("risk:  high  ( 7,5 )"), { severity: "High", score: 7.5 });
  assert.deepEqual(parseRisk(" Risk: Medium (5.4) "), {
    severity: "Medium",
    score: 5.4,
  });
});

test("parseRisk rejects anything else, including the template placeholder", () => {
  const rejected = [
    "Risk: [TODO]",
    "Risk",
    "Risk:",
    "Severity: High (7.0)",
    "Risk: Sever (9)",
    "Risk: Medium (5.4) - see below",
    "Risk: Medium 5.4",
    "Status: Open",
    "Description",
  ];
  for (const text of rejected) {
    assert.equal(parseRisk(text), undefined, `should not parse: ${text}`);
  }
});

test("isRiskHeading spots a risk heading even when it is malformed", () => {
  assert.ok(isRiskHeading("Risk: [TODO]"));
  assert.ok(isRiskHeading("risk :  whatever"));
  assert.ok(!isRiskHeading("Status: [TODO]"));
  assert.ok(!isRiskHeading("Description"));
});

test("compareRisk orders by severity first", () => {
  const order = ["Critical", "High", "Medium", "Low", "Informational"] as const;
  for (let i = 0; i < order.length - 1; i++) {
    assert.ok(
      compareRisk({ severity: order[i] }, { severity: order[i + 1] }) < 0,
      `${order[i]} before ${order[i + 1]}`
    );
  }
});

test("compareRisk orders by score descending within a severity", () => {
  assert.ok(compareRisk({ severity: "Medium", score: 5.4 }, { severity: "Medium", score: 4 }) < 0);
  assert.ok(compareRisk({ severity: "Medium", score: 4 }, { severity: "Medium", score: 5.4 }) > 0);
  assert.equal(compareRisk({ severity: "Medium", score: 4 }, { severity: "Medium", score: 4 }), 0);
});

test("compareRisk puts scored findings ahead of unscored ones of the same severity", () => {
  assert.ok(compareRisk({ severity: "High", score: 1 }, { severity: "High" }) < 0);
  assert.ok(compareRisk({ severity: "High" }, { severity: "High", score: 1 }) > 0);
  assert.equal(compareRisk({ severity: "High" }, { severity: "High" }), 0);
});

test('the template\'s "Info" is accepted for Informational', () => {
  // The report template's own risk scale writes Info, not Informational.
  assert.deepEqual(parseRisk("Risk: Info (2.0)"), { severity: "Informational", score: 2 });
  assert.deepEqual(parseRisk("Risk: Info"), { severity: "Informational" });
  assert.deepEqual(parseRisk("risk: INFO (0.0)"), { severity: "Informational", score: 0 });

  // Both spellings mean the same thing.
  assert.deepEqual(parseRisk("Risk: Informational (2.0)"), parseRisk("Risk: Info (2.0)"));
});

test("a word merely starting with Info is not a severity", () => {
  assert.equal(parseRisk("Risk: Infomercial"), undefined);
  assert.equal(parseRisk("Risk: Informative (2.0)"), undefined);
});
