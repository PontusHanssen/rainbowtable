import * as React from "react";
import {
  Button,
  Dropdown,
  Field,
  Option,
  Spinner,
  Tab,
  TabList,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { Section, findSections, getHeadings } from "../../word/headings";
import { SkippedFinding } from "../../word/section";
import { previewSort, restoreSection, sortFindings } from "../../word/sortFindings";
import { insertFindingsTable, previewTable, removeFindingsTable } from "../../word/findingsTable";
import CvssCalculator from "./CvssCalculator";

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

/** A completed action that can still be taken back. Word's own undo cannot reach ours. */
type Undoable =
  { kind: "sort"; section: Section; snapshot: string } | { kind: "table"; bookmark: string };

/** An action held back until the user has seen the findings we cannot read. */
type Pending = { section: Section; action: "sort" | "table" };

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
  const [tab, setTab] = React.useState<"findings" | "cvss">("findings");
  const [pending, setPending] = React.useState<Pending | undefined>();
  const [undoable, setUndoable] = React.useState<Undoable | undefined>();

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
    setPending(undefined);
    setUndoable(undefined);
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

  /** Run a document action, keeping the pane's state consistent whatever happens. */
  const run = async (action: () => Promise<void>) => {
    setBusy(true);
    setError("");
    setWarnings([]);
    setPending(undefined);
    setUndoable(undefined);
    try {
      await action();
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
      if (result.snapshot) {
        setUndoable({ kind: "sort", section, snapshot: result.snapshot });
      }
    });

  const applyTable = (section: Section) =>
    run(async () => {
      const result = await insertFindingsTable(section);
      setStatus(`Inserted a table of ${result.rows} finding${result.rows === 1 ? "" : "s"}.`);
      setWarnings(
        result.skipped.map((finding) => `Listed without a severity: ${describe(finding)}`)
      );
      setUndoable({ kind: "table", bookmark: result.bookmark });
    });

  // Neither command edits anything until the user has seen the findings we cannot read.
  const onSort = () => {
    const section = selected;
    if (!section) {
      return;
    }
    run(async () => {
      const preview = await previewSort(section);

      if (preview.skipped.length > 0) {
        setPending({ section, action: "sort" });
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

      await applySort(section);
    });
  };

  const onInsertTable = () => {
    const section = selected;
    if (!section) {
      return;
    }
    run(async () => {
      const preview = await previewTable(section);

      if (preview.skipped.length > 0) {
        setPending({ section, action: "table" });
        setWarnings(preview.skipped.map(describe));
        setStatus(
          `${preview.skipped.length} of ${preview.rows} findings have no readable risk rating and would be ` +
            "listed without a severity. Nothing has been inserted."
        );
        return;
      }

      await applyTable(section);
    });
  };

  const onUndo = () => {
    const action = undoable;
    if (!action) {
      return;
    }
    run(async () => {
      if (action.kind === "sort") {
        await restoreSection(action.section, action.snapshot);
        setStatus("Undone — the findings are back in their original order.");
      } else {
        await removeFindingsTable(action.bookmark);
        setStatus("Removed the inserted table.");
      }
    });
  };

  return (
    <div className={styles.root}>
      <TabList
        selectedValue={tab}
        onTabSelect={(_event, data) => setTab(data.value as "findings" | "cvss")}
      >
        <Tab value="findings">Findings</Tab>
        <Tab value="cvss">CVSS 3.1</Tab>
      </TabList>

      {tab === "cvss" && <CvssCalculator />}

      {tab === "findings" && (
        <>
          <Field label="Section">
            <Dropdown
              placeholder="Select a section"
              disabled={busy || sections.length === 0}
              value={selected ? selected.heading.text : ""}
              selectedOptions={selected ? [key(selected)] : []}
              onOptionSelect={(_event, data) => {
                setPending(undefined);
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

          {undoable && !busy && (
            <Button appearance="outline" onClick={onUndo}>
              {undoable.kind === "sort" ? "Undo the last sort" : "Remove the inserted table"}
            </Button>
          )}

          {busy && <Spinner size="tiny" labelPosition="after" label="Working…" />}
          {status && !error && <div className={styles.status}>{status}</div>}

          {warnings.length > 0 && (
            <div className={styles.warnings}>
              {warnings.map((warning) => (
                <div key={warning}>{warning}</div>
              ))}
            </div>
          )}

          {pending && !busy && (
            <div className={styles.confirm}>
              <div className={styles.status}>
                {pending.action === "sort"
                  ? "Fix the risk headings above, or sort the rest and leave those findings where they are."
                  : "Fix the risk headings above, or insert the table with those rows left blank."}
              </div>
              <Button
                appearance="primary"
                onClick={() =>
                  pending.action === "sort"
                    ? applySort(pending.section)
                    : applyTable(pending.section)
                }
              >
                {pending.action === "sort" ? "Sort the rest anyway" : "Insert the table anyway"}
              </Button>
              <Button appearance="subtle" onClick={() => setPending(undefined)}>
                Cancel
              </Button>
            </div>
          )}

          {error && <div className={styles.error}>{error}</div>}
        </>
      )}
    </div>
  );
};

export default App;
