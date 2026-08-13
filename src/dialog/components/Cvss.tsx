import { ReactElement, useState } from "react";
import {
  CvssVector,
  DEFAULT_VECTOR,
  METRICS,
  baseScore,
  calculatorUrl,
  formatVector,
  severityFor,
} from "../../word/cvss";
import { Severity } from "../../word/severity";

/** The colours the report template's own severity styles use. */
const SEVERITY_COLOURS: Record<Severity, string> = {
  Critical: "#A50021",
  High: "#FF0000",
  Medium: "#FFC000",
  Low: "#00B050",
  Informational: "#00B0F0",
};

/**
 * Scoring, as a bar above the editor rather than a tab beside it: it is consulted while
 * writing, so hiding it behind a tab makes it a detour. Metrics are shown by their
 * specification letters with the full name on hover, which is what keeps eight of them on
 * one or two rows.
 *
 * It needs no document access — `cvss.ts` is pure. Applying a score rewrites the `Risk:`
 * line and the vector in the markdown, which the existing pipeline turns into a heading
 * `severity.ts` can read and a clickable link.
 */
export function Cvss({
  onApply,
}: {
  onApply: (risk: string, vector: string) => void;
}): ReactElement {
  const [vector, setVector] = useState<CvssVector>({ ...DEFAULT_VECTOR });

  const score = baseScore(vector);
  const severity = severityFor(score);

  return (
    <div className="cvss-bar">
      <div className="metrics">
        {METRICS.map((metric) => (
          <div className="metric" key={metric.id} title={metric.name}>
            <span className="metric-id">{metric.id}</span>
            {metric.options.map(([code, name]) => (
              <button
                key={code}
                type="button"
                className="choice"
                aria-pressed={vector[metric.id] === code}
                title={`${metric.name}: ${name}`}
                onClick={() => setVector((current) => ({ ...current, [metric.id]: code }))}
              >
                {code}
              </button>
            ))}
          </div>
        ))}
      </div>

      <div className="score-row">
        <span className="score-number" style={{ color: SEVERITY_COLOURS[severity] }}>
          {score.toFixed(1)}
        </span>
        <span className="score-severity" style={{ color: SEVERITY_COLOURS[severity] }}>
          {severity}
        </span>
        <span className="vector">{formatVector(vector)}</span>
        <button
          type="button"
          onClick={() => onApply(`${severity} (${score.toFixed(1)})`, `<${calculatorUrl(vector)}>`)}
        >
          Put in finding
        </button>
      </div>
    </div>
  );
}
