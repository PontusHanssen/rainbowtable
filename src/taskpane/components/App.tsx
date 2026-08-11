import * as React from "react";
import { Button, Dropdown, Field, Option, Spinner, makeStyles, tokens } from "@fluentui/react-components";
import { Section, findSections, getHeadings } from "../../word/headings";
import { sortFindings } from "../../word/sortFindings";
import { insertFindingsTable } from "../../word/findingsTable";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    padding: "16px",
  },
  actions: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
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

const App: React.FC = () => {
  const styles = useStyles();
  const [sections, setSections] = React.useState<Section[]>([]);
  const [selected, setSelected] = React.useState<Section | undefined>();
  const [busy, setBusy] = React.useState(false);
  const [status, setStatus] = React.useState("");
  const [error, setError] = React.useState("");

  const refresh = React.useCallback(async () => {
    setBusy(true);
    setError("");
    try {
      const found = findSections(await getHeadings());
      setSections(found);
      setSelected((current) => found.find((section) => section.heading.text === current?.heading.text));
      setStatus(`${found.length} section${found.length === 1 ? "" : "s"} with findings.`);
    } catch (err) {
      setError(String(err));
    } finally {
      setBusy(false);
    }
  }, []);

  React.useEffect(() => {
    refresh();
  }, [refresh]);

  const run = async (action: (section: Section) => Promise<void>) => {
    if (!selected) {
      return;
    }
    setBusy(true);
    setError("");
    try {
      await action(selected);
      await refresh();
    } catch (err) {
      setError(String(err));
      setBusy(false);
    }
  };

  return (
    <div className={styles.root}>
      <Field label="Section">
        <Dropdown
          placeholder="Select a section"
          disabled={busy || sections.length === 0}
          value={selected ? selected.heading.text : ""}
          selectedOptions={selected ? [selected.heading.text] : []}
          onOptionSelect={(_event, data) =>
            setSelected(sections.find((section) => section.heading.text === data.optionValue))
          }
        >
          {sections.map((section) => (
            <Option key={section.heading.index} value={section.heading.text}>
              {`${section.heading.text} (${section.findings.length})`}
            </Option>
          ))}
        </Dropdown>
      </Field>

      <div className={styles.actions}>
        <Button appearance="primary" disabled={!selected || busy} onClick={() => run(sortFindings)}>
          Sort findings by severity
        </Button>
        <Button disabled={!selected || busy} onClick={() => run(insertFindingsTable)}>
          Insert findings table
        </Button>
        <Button appearance="subtle" disabled={busy} onClick={refresh}>
          Rescan document
        </Button>
      </div>

      {busy && <Spinner size="tiny" labelPosition="after" label="Working…" />}
      {status && !error && <div className={styles.status}>{status}</div>}
      {error && <div className={styles.error}>{error}</div>}
    </div>
  );
};

export default App;
