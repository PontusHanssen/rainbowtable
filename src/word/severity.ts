/** Severities in report order: most severe first. */
export const SEVERITIES = ["Critical", "High", "Medium", "Low", "Informational"] as const;

export type Severity = (typeof SEVERITIES)[number];

export interface Risk {
  severity: Severity;
  /** Absent when the heading carries no "(score)" part. */
  score?: number;
}

/**
 * The one accepted form: `Risk: <Severity> (<score>)`, score optional.
 *
 * Deliberately strict — a heading that does not match is reported to the user rather
 * than guessed at, because a silently mis-sorted report is worse than a visible skip.
 * The only latitude is case, surrounding whitespace, and a decimal comma, since the
 * template is a Swedish-language document.
 */
const RISK_HEADING =
  /^Risk:\s*(Critical|High|Medium|Low|Informational)\s*(?:\(\s*(\d+(?:[.,]\d+)?)\s*\))?\s*$/i;

/** Recognises a heading that is *meant* to be a risk rating, well-formed or not. */
export function isRiskHeading(text: string): boolean {
  return /^Risk\s*:/i.test(normalize(text));
}

/** The risk rating in a heading, or undefined if it is not in the accepted form. */
export function parseRisk(text: string): Risk | undefined {
  const match = RISK_HEADING.exec(normalize(text));
  if (!match) {
    return undefined;
  }

  const severity = SEVERITIES.find(
    (candidate) => candidate.toLowerCase() === match[1].toLowerCase()
  );
  if (!severity) {
    return undefined;
  }

  return match[2] ? { severity, score: Number(match[2].replace(",", ".")) } : { severity };
}

/**
 * Sort comparator: severity first, then score descending.
 *
 * Findings whose heading carries no score sort after the scored ones of the same
 * severity, and hold their original relative order there (callers rely on
 * `Array.prototype.sort` being stable).
 */
export function compareRisk(a: Risk, b: Risk): number {
  const bySeverity = SEVERITIES.indexOf(a.severity) - SEVERITIES.indexOf(b.severity);
  if (bySeverity !== 0) {
    return bySeverity;
  }
  if (a.score === undefined || b.score === undefined) {
    return a.score === b.score ? 0 : a.score === undefined ? 1 : -1;
  }
  return b.score - a.score;
}

/** Word uses non-breaking spaces in numbered headings often enough to matter. */
function normalize(text: string): string {
  return text.replace(/\u00a0/g, " ").trim();
}
