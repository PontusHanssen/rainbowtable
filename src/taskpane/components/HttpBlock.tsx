import * as React from "react";
import { Button, Field, Textarea, makeStyles, tokens } from "@fluentui/react-components";
import { highlightHttp } from "../../word/http";
import { insertHttpBlock, removeHttpBlock } from "../../word/httpBlock";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  input: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
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

const DESCRIBE_BODY: Record<string, string> = {
  json: "JSON body",
  xml: "XML body",
  form: "form-encoded body",
  none: "no structured body",
};

const HttpBlock: React.FC = () => {
  const styles = useStyles();
  const [raw, setRaw] = React.useState("");
  const [busy, setBusy] = React.useState(false);
  const [status, setStatus] = React.useState("");
  const [error, setError] = React.useState("");
  const [inserted, setInserted] = React.useState<string | undefined>();

  /** Parsed on every keystroke so the pane can say what it sees before anything is written. */
  const preview = React.useMemo(() => (raw.trim() ? highlightHttp(raw) : undefined), [raw]);

  const onInsert = async () => {
    setBusy(true);
    setError("");
    try {
      const result = await insertHttpBlock(raw);
      setStatus(`Inserted ${result.lines} lines as a ${result.kind}.`);
      setInserted(result.bookmark);
    } catch (err) {
      setError(String(err));
    } finally {
      setBusy(false);
    }
  };

  const onUndo = async () => {
    if (!inserted) {
      return;
    }
    setBusy(true);
    setError("");
    try {
      await removeHttpBlock(inserted);
      setInserted(undefined);
      setStatus("Removed the inserted block.");
    } catch (err) {
      setError(String(err));
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className={styles.root}>
      <Field label="Paste the request or response from Burp">
        <Textarea
          className={styles.input}
          resize="vertical"
          rows={12}
          value={raw}
          placeholder={"GET /admin HTTP/1.1\nHost: app.example.com"}
          onChange={(_event, data) => {
            setRaw(data.value);
            setStatus("");
          }}
        />
      </Field>

      {preview && (
        <div className={styles.status}>
          {preview.kind === "unknown"
            ? "Not recognised as HTTP — it will be inserted as a plain code block."
            : `Recognised as an HTTP ${preview.kind} with ${DESCRIBE_BODY[preview.body]}.`}
        </div>
      )}

      <Button appearance="primary" disabled={busy || !raw.trim()} onClick={onInsert}>
        Insert at the cursor
      </Button>

      {inserted && !busy && (
        <Button appearance="outline" onClick={onUndo}>
          Remove the inserted block
        </Button>
      )}

      <div className={styles.status}>
        Pasting here rather than into Word keeps the message exactly as Burp gave it — autocorrect
        and smart quotes would otherwise alter your evidence.
      </div>

      {status && !error && <div className={styles.status}>{status}</div>}
      {error && <div className={styles.error}>{error}</div>}
    </div>
  );
};

export default HttpBlock;
