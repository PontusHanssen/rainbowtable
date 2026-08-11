import test from "node:test";
import assert from "node:assert/strict";
import {
  CvssVector,
  DEFAULT_VECTOR,
  baseScore,
  calculatorUrl,
  formatVector,
  isVectorLine,
  roundUp,
  severityFor,
} from "../src/word/cvss";

/** Build a vector from its string form, so tests read like published CVSS vectors. */
function parse(vector: string): CvssVector {
  const entries = vector
    .replace(/^CVSS:3\.[01]\//, "")
    .split("/")
    .map((part) => part.split(":") as [string, string]);
  return Object.fromEntries(entries) as CvssVector;
}

const score = (vector: string) => baseScore(parse(vector));

test("published vectors score as documented", () => {
  // Well-known CVEs, whose scores are published by NVD and first.org.
  const cases: [string, number][] = [
    ["AV:N/AC:L/PR:N/UI:N/S:U/C:H/I:H/A:H", 9.8], // BlueKeep, CVE-2019-0708
    ["AV:N/AC:L/PR:N/UI:N/S:C/C:H/I:H/A:H", 10.0], // Zerologon, CVE-2020-1472
    ["AV:N/AC:L/PR:N/UI:N/S:U/C:H/I:N/A:N", 7.5], // Heartbleed, CVE-2014-0160
    ["AV:N/AC:L/PR:N/UI:N/S:U/C:N/I:N/A:H", 7.5], // remote denial of service
    ["AV:L/AC:L/PR:L/UI:N/S:U/C:H/I:H/A:H", 7.8], // local privilege escalation
    ["AV:N/AC:L/PR:N/UI:R/S:C/C:L/I:L/A:N", 6.1], // reflected XSS
  ];

  for (const [vector, expected] of cases) {
    assert.equal(score(vector), expected, vector);
  }
});

test("no impact scores zero whatever the exploitability", () => {
  assert.equal(score("AV:N/AC:L/PR:N/UI:N/S:U/C:N/I:N/A:N"), 0);
  assert.equal(score("AV:P/AC:H/PR:H/UI:R/S:C/C:N/I:N/A:N"), 0);
  assert.equal(baseScore(DEFAULT_VECTOR), 0, "the pane starts at a harmless vector");
});

test("changing scope raises the score", () => {
  const unchanged = score("AV:N/AC:L/PR:L/UI:N/S:U/C:L/I:L/A:L");
  const changed = score("AV:N/AC:L/PR:L/UI:N/S:C/C:L/I:L/A:L");

  assert.ok(changed > unchanged, `${changed} > ${unchanged}`);
});

test("scores stay within the scale", () => {
  for (const av of ["N", "A", "L", "P"]) {
    for (const ac of ["L", "H"]) {
      for (const pr of ["N", "L", "H"]) {
        for (const s of ["U", "C"]) {
          const value = score(`AV:${av}/AC:${ac}/PR:${pr}/UI:N/S:${s}/C:H/I:H/A:H`);
          assert.ok(value >= 0 && value <= 10, `${value} out of range`);
          assert.equal(value, Math.round(value * 10) / 10, "one decimal place");
        }
      }
    }
  }
});

test("roundUp rounds up to one decimal, exactly as the spec defines", () => {
  assert.equal(roundUp(4.02), 4.1);
  assert.equal(roundUp(4.0), 4.0);
  assert.equal(roundUp(0.0), 0.0);
  // The case CVSS 3.1 introduced integer arithmetic to get right.
  assert.equal(roundUp(6.1 - 0.0000000000000005), 6.1);
});

test("severity bands follow the qualitative scale", () => {
  assert.equal(severityFor(0), "Informational");
  assert.equal(severityFor(0.1), "Low");
  assert.equal(severityFor(3.9), "Low");
  assert.equal(severityFor(4.0), "Medium");
  assert.equal(severityFor(6.9), "Medium");
  assert.equal(severityFor(7.0), "High");
  assert.equal(severityFor(8.9), "High");
  assert.equal(severityFor(9.0), "Critical");
  assert.equal(severityFor(10.0), "Critical");
});

test("the vector string keeps the specification's metric order", () => {
  assert.equal(
    formatVector(parse("AV:N/AC:L/PR:N/UI:N/S:U/C:H/I:H/A:H")),
    "CVSS:3.1/AV:N/AC:L/PR:N/UI:N/S:U/C:H/I:H/A:H"
  );
});

test("the calculator link carries the vector as its fragment", () => {
  assert.equal(
    calculatorUrl(parse("AV:N/AC:L/PR:N/UI:N/S:U/C:H/I:H/A:H")),
    "https://www.first.org/cvss/calculator/3.1#CVSS:3.1/AV:N/AC:L/PR:N/UI:N/S:U/C:H/I:H/A:H"
  );
});

test("isVectorLine spots a vector paragraph so it gets replaced, not repeated", () => {
  assert.ok(isVectorLine("CVSS:3.1/AV:N/AC:L/PR:N/UI:N/S:U/C:H/I:H/A:H"));
  assert.ok(isVectorLine("  CVSS:3.0/AV:N/AC:L/PR:N/UI:N/S:U/C:H/I:H/A:H  "));
  assert.ok(!isVectorLine("Risk: Critical (9.8)"));
  assert.ok(!isVectorLine("The CVSS:3.1 vector is given below"));
  assert.ok(!isVectorLine(""));
});

test("the severity the calculator produces is one the report parser accepts", () => {
  // The whole point of the feature: what it writes must satisfy severity.ts.
  const rating = `Risk: ${severityFor(score("AV:N/AC:L/PR:N/UI:N/S:U/C:H/I:H/A:H"))} (${score(
    "AV:N/AC:L/PR:N/UI:N/S:U/C:H/I:H/A:H"
  ).toFixed(1)})`;

  assert.equal(rating, "Risk: Critical (9.8)");
});
