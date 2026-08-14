import { ReactElement } from "react";
import {
  CvssVector,
  METRICS,
  baseScore,
  calculatorUrl,
  formatVector,
  severityFor,
} from "../../word/cvss";

/**
 * Scoring, as a bar above the editor rather than a tab beside it: it is consulted while
 * writing, so hiding it behind a tab makes it a detour. Metrics are shown by their
 * specification letters with the full name on hover, which is what keeps eight of them on
 * one or two rows.
 *
 * It needs no document access — `cvss.ts` is pure. Applying a score rewrites the `Risk:`
 * line and the vector in the markdown, which the existing pipeline turns into a heading
 * `severity.ts` can read and a clickable link.
 *
 * The vector is owned by `App`, not by this component: it has to be cleared alongside the
 * editor when a finding is inserted, or the next finding silently starts out carrying the
 * previous one's score.
 */
export function Cvss({
  vector,
  onChange,
  onApply,
}: {
  vector: CvssVector;
  onChange: (vector: CvssVector) => void;
  onApply: (risk: string, vector: string) => void;
}): ReactElement {
  const score = baseScore(vector);
  const severity = severityFor(score);

  return (
    <div className="cvss-bar">
      <div className="metrics">
        {METRICS.map((metric) => (
          <div
            className="metric"
            key={metric.id}
            role="group"
            aria-label={metric.name}
            title={metric.name}
          >
            <span className="metric-id">{metric.id}</span>
            {metric.options.map(([code, name]) => (
              <button
                key={code}
                type="button"
                className="choice"
                aria-pressed={vector[metric.id] === code}
                aria-label={`${metric.name}: ${name}`}
                title={`${metric.name}: ${name}`}
                onClick={() => onChange({ ...vector, [metric.id]: code })}
              >
                {code}
              </button>
            ))}
          </div>
        ))}
      </div>

      <div className="score-row">
        <span className="score-number">{score.toFixed(1)}</span>
        {/* The template's colours as a pill: as text they were unreadable, and the
            default state opens on the worst of them. */}
        <span className={`pill sev-${severity.toLowerCase()}`}>{severity}</span>
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
