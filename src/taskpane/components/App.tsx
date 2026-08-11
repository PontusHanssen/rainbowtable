import * as React from "react";
import {
  Button,
  Dropdown,
  Field,
  Option,
  Spinner,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { Section, findSections, getHeadings } from "../../word/headings";
import { SkippedFinding, previewSort, sortFindings } from "../../word/sortFindings";
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
  confirm: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    padding: "12px",
    borderLeft: `3px solid ${tokens.colorPaletteDarkOrangeBorderActive}`,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  status: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  warnings: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    color: tokens.colorPaletteDarkOrangeForeground1,
    fontSize: tokens.fontSizeBase200,
  },
  error: {
    color: tokens.colorPaletteRedForeground1,
    fontSize: tokens.fontSizeBase200,
  },
});

const describe = ({ title, reason }: SkippedFinding) => `"${title}" — ${reason}`;

/**
 * Identifies a section in the dropdown. Titles repeat in real reports — the template
 * has "Vulnerabilities" as both a summary subsection and a findings section — so the
 * paragraph index is what keeps them apart.
 */
const key = (section: Section) => String(section.heading.index);

/** Re-find the selected section after a rescan, whose paragraph indexes may have moved. */
function reselect(sections: Section[], previous: Section): Section | undefined {
  return (
    sections.find(
      (section) =>
        section.heading.index === previous.heading.index &&
        section.heading.text === previous.heading.text
    ) ??
    sections.find(
      (section) =>
        section.heading.text === previous.heading.text &&
        section.heading.level === previous.heading.level
    )
  );
}

const App: React.FC = () => {
  const styles = useStyles();
  const [sections, setSections] = React.useState<Section[]>([]);
  const [selected, setSelected] = React.useState<Section | undefined>();
  const [busy, setBusy] = React.useState(false);
  const [status, setStatus] = React.useState("");
  const [warnings, setWarnings] = React.useState<string[]>([]);
  const [error, setError] = React.useState("");
  /** Set when sorting is waiting for the user to accept unreadable findings. */
  const [unconfirmed, setUnconfirmed] = React.useState<Section | undefined>();

  /** Re-read the document. `announce` is off when refreshing after an action, so the
      action's own result stays on screen. */
  const scan = React.useCallback(async (announce: boolean) => {
    const found = findSections(await getHeadings());
    setSections(found);
    setSelected((current) => (current ? reselect(found, current) : undefined));
    if (announce) {
      setStatus(`${found.length} section${found.length === 1 ? "" : "s"} with findings.`);
    }
  }, []);

  const refresh = React.useCallback(async () => {
    setBusy(true);
    setError("");
    setWarnings([]);
    setUnconfirmed(undefined);
    try {
      await scan(true);
    } catch (err) {
      setError(String(err));
    } finally {
      setBusy(false);
    }
  }, [scan]);

  React.useEffect(() => {
    refresh();
  }, [refresh]);

  const run = async (action: (section: Section) => Promise<void>) => {
    if (!selected) {
      return;
    }
    setBusy(true);
    setError("");
    setWarnings([]);
    setUnconfirmed(undefined);
    try {
      await action(selected);
      await scan(false);
    } catch (err) {
      setError(String(err));
    } finally {
      setBusy(false);
    }
  };

  const applySort = (section: Section) =>
    run(async () => {
      const result = await sortFindings(section);
      setStatus(
        result.changed
          ? `Sorted ${result.sorted} finding${result.sorted === 1 ? "" : "s"} by severity.`
          : "Already in severity order — nothing changed."
      );
      setWarnings(result.skipped.map((finding) => `Left in place: ${describe(finding)}`));
    });

  // Nothing is edited until the user has seen any findings we cannot read.
  const onSort = () =>
    run(async (section) => {
      const preview = await previewSort(section);

      if (preview.skipped.length > 0) {
        setUnconfirmed(section);
        setWarnings(preview.skipped.map(describe));
        setStatus(
          `${preview.skipped.length} of ${preview.skipped.length + preview.sorted} findings have no readable risk rating. ` +
            "The document has not been changed."
        );
        return;
      }

      if (!preview.changed) {
        setStatus("Already in severity order — nothing changed.");
        return;
      }

      const result = await sortFindings(section);
      setStatus(`Sorted ${result.sorted} finding${result.sorted === 1 ? "" : "s"} by severity.`);
    });

  const onInsertTable = () => run((section) => insertFindingsTable(section));

  return (
    <div className={styles.root}>
      <Field label="Section">
        <Dropdown
          placeholder="Select a section"
          disabled={busy || sections.length === 0}
          value={selected ? selected.heading.text : ""}
          selectedOptions={selected ? [selected.heading.text] : []}
          onOptionSelect={(_event, data) => {
            setUnconfirmed(undefined);
            setSelected(sections.find((section) => key(section) === data.optionValue));
          }}
        >
          {sections.map((section) => (
            <Option key={key(section)} value={key(section)} text={section.heading.text}>
              {`${section.heading.text} (${section.findings.length})`}
            </Option>
          ))}
        </Dropdown>
      </Field>

      <div className={styles.actions}>
        <Button appearance="primary" disabled={!selected || busy} onClick={onSort}>
          Sort findings by severity
        </Button>
        <Button disabled={!selected || busy} onClick={onInsertTable}>
          Insert findings table
        </Button>
        <Button appearance="subtle" disabled={busy} onClick={refresh}>
          Rescan document
        </Button>
      </div>

      {busy && <Spinner size="tiny" labelPosition="after" label="Working…" />}
      {status && !error && <div className={styles.status}>{status}</div>}

      {warnings.length > 0 && (
        <div className={styles.warnings}>
          {warnings.map((warning) => (
            <div key={warning}>{warning}</div>
          ))}
        </div>
      )}

      {unconfirmed && !busy && (
        <div className={styles.confirm}>
          <div className={styles.status}>
            Fix the risk headings above, or sort the rest and leave those findings where they are.
          </div>
          <Button appearance="primary" onClick={() => applySort(unconfirmed)}>
            Sort the rest anyway
          </Button>
          <Button appearance="subtle" onClick={() => setUnconfirmed(undefined)}>
            Cancel
          </Button>
        </div>
      )}

      {error && <div className={styles.error}>{error}</div>}
    </div>
  );
};

export default App;
