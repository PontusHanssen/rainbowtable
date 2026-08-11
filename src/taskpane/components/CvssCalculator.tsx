import * as React from "react";
import { Button, ToggleButton, makeStyles, tokens } from "@fluentui/react-components";
import {
  CvssVector,
  DEFAULT_VECTOR,
  METRICS,
  baseScore,
  formatVector,
  severityFor,
} from "../../word/cvss";
import { RiskUndo, insertRisk, undoRisk } from "../../word/insertRisk";
import { Severity } from "../../word/severity";

/** The same palette the findings table uses, so the pane and the report agree. */
const SEVERITY_COLOURS: Record<Severity, string> = {
  Critical: "#C00000",
  High: "#E36C0A",
  Medium: "#FFC000",
  Low: "#92D050",
  Informational: "#BFBFBF",
};

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  metric: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
  label: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  choices: {
    display: "flex",
    flexWrap: "wrap",
    gap: "4px",
  },
  score: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "8px 12px",
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  number: {
    fontSize: tokens.fontSizeHero700,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: "1",
  },
  severity: {
    fontWeight: tokens.fontWeightSemibold,
  },
  vector: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    wordBreak: "break-all",
    color: tokens.colorNeutralForeground3,
  },
  status: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  error: {
    color: tokens.colorPaletteRedForeground1,
    fontSize: tokens.fontSizeBase200,
  },
});

const CvssCalculator: React.FC = () => {
  const styles = useStyles();
  const [vector, setVector] = React.useState<CvssVector>(DEFAULT_VECTOR);
  const [busy, setBusy] = React.useState(false);
  const [status, setStatus] = React.useState("");
  const [error, setError] = React.useState("");
  const [undo, setUndo] = React.useState<RiskUndo | undefined>();

  const score = baseScore(vector);
  const severity = severityFor(score);

  const choose = (id: keyof CvssVector, value: string) => {
    setVector((current) => ({ ...current, [id]: value }));
    setStatus("");
  };

  const onInsert = async () => {
    setBusy(true);
    setError("");
    try {
      const written = await insertRisk(vector);
      setStatus(`Wrote "Risk: ${written.rating}" and the vector below it.`);
      setUndo(written.undo);
    } catch (err) {
      setError(String(err));
    } finally {
      setBusy(false);
    }
  };

  const onUndo = async () => {
    if (!undo) {
      return;
    }
    setBusy(true);
    setError("");
    try {
      await undoRisk(undo);
      setUndo(undefined);
      setStatus("Undone — the Risk line is back as it was.");
    } catch (err) {
      setError(String(err));
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.score}>
        <span className={styles.number}>{score.toFixed(1)}</span>
        <span className={styles.severity} style={{ color: SEVERITY_COLOURS[severity] }}>
          {severity}
        </span>
      </div>

      {METRICS.map((metric) => (
        <div key={metric.id} className={styles.metric}>
          <span className={styles.label}>{metric.name}</span>
          <div className={styles.choices}>
            {metric.options.map(([code, name]) => (
              <ToggleButton
                key={code}
                size="small"
                checked={vector[metric.id] === code}
                onClick={() => choose(metric.id, code)}
                title={`${metric.id}:${code} — ${name}`}
              >
                {name}
              </ToggleButton>
            ))}
          </div>
        </div>
      ))}

      <div className={styles.vector}>{formatVector(vector)}</div>

      <Button appearance="primary" disabled={busy} onClick={onInsert}>
        Write the risk at the cursor
      </Button>

      {undo && !busy && (
        <Button appearance="outline" onClick={onUndo}>
          Undo the last risk
        </Button>
      )}

      <div className={styles.status}>
        Put the cursor on the finding&apos;s &quot;Risk:&quot; line. The vector goes on the line
        below, linked to the NVD calculator.
      </div>

      {status && !error && <div className={styles.status}>{status}</div>}
      {error && <div className={styles.error}>{error}</div>}
    </div>
  );
};

export default CvssCalculator;
