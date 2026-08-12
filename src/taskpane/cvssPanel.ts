import {
  CvssVector,
  DEFAULT_VECTOR,
  METRICS,
  baseScore,
  formatVector,
  severityFor,
} from "../word/cvss";
import { RiskUndo, insertRisk, undoRisk } from "../word/insertRisk";
import { Severity } from "../word/severity";
import { byId, feedbackFor, guard, make, show } from "./dom";

/* global HTMLButtonElement */

/** The same palette the findings table uses, so the pane and the report agree. */
const SEVERITY_COLOURS: Record<Severity, string> = {
  Critical: "#C00000",
  High: "#E36C0A",
  Medium: "#FFC000",
  Low: "#92D050",
  Informational: "#8A8A8A",
};

export function setUpCvssPanel(): void {
  const scoreLabel = byId("cvss-score");
  const severityLabel = byId("cvss-severity");
  const metrics = byId("cvss-metrics");
  const vectorLabel = byId("cvss-vector");
  const insert = byId<HTMLButtonElement>("cvss-insert");
  const undo = byId<HTMLButtonElement>("cvss-undo");

  const feedback = feedbackFor("cvss");
  const buttons = [insert, undo];

  const vector: CvssVector = { ...DEFAULT_VECTOR };
  let lastWritten: RiskUndo | undefined;

  /** Every choice button, so the pressed state can be refreshed after a change. */
  const choices: { id: keyof CvssVector; code: string; button: HTMLButtonElement }[] = [];

  const render = () => {
    const score = baseScore(vector);
    const severity = severityFor(score);

    scoreLabel.textContent = score.toFixed(1);
    severityLabel.textContent = severity;
    severityLabel.style.color = SEVERITY_COLOURS[severity];
    vectorLabel.textContent = formatVector(vector);

    choices.forEach(({ id, code, button }) =>
      button.setAttribute("aria-pressed", String(vector[id] === code))
    );
  };

  METRICS.forEach((metric) => {
    const group = make("div", "metric");
    group.appendChild(make("label", undefined, metric.name));

    const row = make("div", "choices");
    metric.options.forEach(([code, name]) => {
      const button = make("button", "choice", name);
      button.title = `${metric.id}:${code} — ${name}`;
      button.onclick = () => {
        vector[metric.id] = code;
        feedback.status("");
        render();
      };
      choices.push({ id: metric.id, code, button });
      row.appendChild(button);
    });

    group.appendChild(row);
    metrics.appendChild(group);
  });

  insert.onclick = () =>
    guard(buttons, feedback, async () => {
      const written = await insertRisk(vector);
      feedback.status(`Wrote "Risk: ${written.rating}" and the vector below it.`);
      lastWritten = written.undo;
      show(undo, true);
    });

  undo.onclick = () =>
    guard(buttons, feedback, async () => {
      if (!lastWritten) {
        return;
      }
      await undoRisk(lastWritten);
      lastWritten = undefined;
      show(undo, false);
      feedback.status("Undone — the Risk line is back as it was.");
    });

  render();
}
