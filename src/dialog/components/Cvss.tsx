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
 * Scoring, which needs no document access at all: `cvss.ts` is pure, so it runs here
 * unchanged. Applying a score rewrites the finding's `Risk:` line and the vector beneath
 * it *in the markdown*, which the existing pipeline already turns into a real heading and
 * a clickable link — so this feature adds no Word code.
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
    <div className="cvss">
      <div className="score">
        <span className="score-number">{score.toFixed(1)}</span>
        <span className="score-severity" style={{ color: SEVERITY_COLOURS[severity] }}>
          {severity}
        </span>
      </div>

      {METRICS.map((metric) => (
        <div className="metric" key={metric.id}>
          <span className="label">{metric.name}</span>
          <div className="choices">
            {metric.options.map(([code, name]) => (
              <button
                key={code}
                type="button"
                className="choice"
                aria-pressed={vector[metric.id] === code}
                title={`${metric.id}:${code} — ${name}`}
                onClick={() => setVector((current) => ({ ...current, [metric.id]: code }))}
              >
                {name}
              </button>
            ))}
          </div>
        </div>
      ))}

      <div className="vector">{formatVector(vector)}</div>

      <button
        type="button"
        className="primary"
        onClick={() => onApply(`${severity} (${score.toFixed(1)})`, `<${calculatorUrl(vector)}>`)}
      >
        Put this score in the finding
      </button>
    </div>
  );
}
