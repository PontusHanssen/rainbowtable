import { Severity } from "./severity";

/**
 * CVSS v3.1 base score, per the specification's formulas (section 7.1).
 *
 * Base metrics only. Temporal and environmental metrics are not modelled: reports quote
 * a base score, and adding the rest would be a much larger UI for something unused.
 */

export type MetricId = "AV" | "AC" | "PR" | "UI" | "S" | "C" | "I" | "A";

export type CvssVector = Record<MetricId, string>;

export interface Metric {
  id: MetricId;
  name: string;
  /** Value code paired with its display name, in the specification's order. */
  options: [string, string][];
}

export const METRICS: Metric[] = [
  {
    id: "AV",
    name: "Attack Vector",
    options: [
      ["N", "Network"],
      ["A", "Adjacent"],
      ["L", "Local"],
      ["P", "Physical"],
    ],
  },
  {
    id: "AC",
    name: "Attack Complexity",
    options: [
      ["L", "Low"],
      ["H", "High"],
    ],
  },
  {
    id: "PR",
    name: "Privileges Required",
    options: [
      ["N", "None"],
      ["L", "Low"],
      ["H", "High"],
    ],
  },
  {
    id: "UI",
    name: "User Interaction",
    options: [
      ["N", "None"],
      ["R", "Required"],
    ],
  },
  {
    id: "S",
    name: "Scope",
    options: [
      ["U", "Unchanged"],
      ["C", "Changed"],
    ],
  },
  {
    id: "C",
    name: "Confidentiality",
    options: [
      ["H", "High"],
      ["L", "Low"],
      ["N", "None"],
    ],
  },
  {
    id: "I",
    name: "Integrity",
    options: [
      ["H", "High"],
      ["L", "Low"],
      ["N", "None"],
    ],
  },
  {
    id: "A",
    name: "Availability",
    options: [
      ["H", "High"],
      ["L", "Low"],
      ["N", "None"],
    ],
  },
];

/** The least severe vector, and the starting point in the task pane. */
export const DEFAULT_VECTOR: CvssVector = {
  AV: "N",
  AC: "L",
  PR: "N",
  UI: "N",
  S: "U",
  C: "N",
  I: "N",
  A: "N",
};

const ATTACK_VECTOR: Record<string, number> = { N: 0.85, A: 0.62, L: 0.55, P: 0.2 };
const ATTACK_COMPLEXITY: Record<string, number> = { L: 0.77, H: 0.44 };
const USER_INTERACTION: Record<string, number> = { N: 0.85, R: 0.62 };
const IMPACT: Record<string, number> = { H: 0.56, L: 0.22, N: 0 };

/** Privileges Required is worth more when the attack changes scope. */
const PRIVILEGES: Record<"U" | "C", Record<string, number>> = {
  U: { N: 0.85, L: 0.62, H: 0.27 },
  C: { N: 0.85, L: 0.68, H: 0.5 },
};

/**
 * The specification's Roundup: round up to one decimal, defined over integer arithmetic
 * to avoid the floating-point edge cases that made CVSS 3.0 implementations disagree.
 */
export function roundUp(value: number): number {
  const scaled = Math.round(value * 100000);
  return scaled % 10000 === 0 ? scaled / 100000 : (Math.floor(scaled / 10000) + 1) / 10;
}

/** The CVSS v3.1 base score, 0.0 to 10.0. */
export function baseScore(vector: CvssVector): number {
  const scopeChanged = vector.S === "C";

  const iss = 1 - (1 - IMPACT[vector.C]) * (1 - IMPACT[vector.I]) * (1 - IMPACT[vector.A]);
  const impact = scopeChanged ? 7.52 * (iss - 0.029) - 3.25 * Math.pow(iss - 0.02, 15) : 6.42 * iss;

  if (impact <= 0) {
    return 0;
  }

  const exploitability =
    8.22 *
    ATTACK_VECTOR[vector.AV] *
    ATTACK_COMPLEXITY[vector.AC] *
    PRIVILEGES[scopeChanged ? "C" : "U"][vector.PR] *
    USER_INTERACTION[vector.UI];

  const combined = scopeChanged ? 1.08 * (impact + exploitability) : impact + exploitability;
  return roundUp(Math.min(combined, 10));
}

/**
 * The qualitative rating for a score.
 *
 * CVSS calls the 0.0 band "None"; reports call it Informational, which is also the
 * severity the rest of the add-in knows about.
 */
export function severityFor(score: number): Severity {
  if (score === 0) {
    return "Informational";
  }
  if (score < 4) {
    return "Low";
  }
  if (score < 7) {
    return "Medium";
  }
  if (score < 9) {
    return "High";
  }
  return "Critical";
}

/** The vector string, e.g. `CVSS:3.1/AV:N/AC:L/PR:N/UI:N/S:U/C:H/I:H/A:H`. */
export function formatVector(vector: CvssVector): string {
  return `CVSS:3.1/${METRICS.map((metric) => `${metric.id}:${vector[metric.id]}`).join("/")}`;
}

/** The first.org calculator, opened on this vector so the scoring can be checked. */
export function calculatorUrl(vector: CvssVector): string {
  return `https://www.first.org/cvss/calculator/3.1#${formatVector(vector)}`;
}

/** True when a line of text is a CVSS vector, so it can be replaced rather than repeated. */
export function isVectorLine(text: string): boolean {
  return /^CVSS:3\.[01]\//.test(text.trim());
}
