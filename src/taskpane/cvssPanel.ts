import {
  CvssVector,
  DEFAULT_VECTOR,
  METRICS,
  baseScore,
  formatVector,
  severityFor,
} from "../word/cvss";
import { RiskUndo, insertRisk, undoRisk } from "../word/insertRisk";
import { byId, feedbackFor, guard, make, show } from "./dom";

/* global HTMLButtonElement */

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
    // The template's own severity colours, worn as a pill. As text colour they were
    // unreadable — its amber is about 1.9:1 against white.
    severityLabel.className = `pill sev-${severity.toLowerCase()}`;
    vectorLabel.textContent = formatVector(vector);

    choices.forEach(({ id, code, button }) =>
      button.setAttribute("aria-pressed", String(vector[id] === code))
    );
  };

  METRICS.forEach((metric) => {
    const group = make("div", "metric");
    const label = make("label", undefined, metric.name);
    label.id = `cvss-metric-${metric.id}`;
    group.appendChild(label);

    // Named as a group, so the choices are not read out as bare words with no context.
    const row = make("div", "choices");
    row.setAttribute("role", "group");
    row.setAttribute("aria-labelledby", label.id);

    metric.options.forEach(([code, name]) => {
      const button = make("button", "choice", name);
      button.title = `${metric.id}:${code} — ${name}`;
      button.onclick = () => {
        // The last result described the vector as it was; the undo it left behind still
        // points at real text in the document, so that stays.
        feedback.status("");
        vector[metric.id] = code;
        render();
      };
      choices.push({ id: metric.id, code, button });
      row.appendChild(button);
    });

    group.appendChild(row);
    metrics.appendChild(group);
  });

  insert.onclick = () =>
    guard(
      buttons,
      feedback,
      async () => {
        const written = await insertRisk(vector);
        feedback.status(`Wrote "Risk: ${written.rating}" and the vector below it.`);
        lastWritten = written.undo;
        show(undo, true);
      },
      "Writing the risk…"
    );

  undo.onclick = () =>
    guard(
      buttons,
      feedback,
      async () => {
        if (!lastWritten) {
          return;
        }
        await undoRisk(lastWritten);
        lastWritten = undefined;
        show(undo, false);
        feedback.status("Undone — the Risk line is back as it was.");
      },
      "Putting the Risk line back…"
    );

  render();
}
