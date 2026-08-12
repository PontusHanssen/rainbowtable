import { Section, findSections, getHeadings } from "../word/headings";
import { SkippedFinding } from "../word/section";
import { previewSort, restoreSection, sortFindings } from "../word/sortFindings";
import { insertFindingsTable, previewTable, removeFindingsTable } from "../word/findingsTable";
import { insertFinding, removeFinding } from "../word/newFinding";
import { Feedback, byId, feedbackFor, guard, make, show } from "./dom";

/* global HTMLButtonElement, HTMLSelectElement */

/** A completed action that can still be taken back. Word's own undo cannot reach ours. */
type Undoable =
  | { kind: "sort"; section: Section; snapshot: string }
  | { kind: "table"; bookmark: string }
  | { kind: "finding"; bookmark: string };

const describe = ({ title, reason }: SkippedFinding) => `"${title}" — ${reason}`;

/**
 * Sections are identified by the paragraph index of their heading. Titles repeat in real
 * reports — the template has "Vulnerabilities" as both a summary subsection and a
 * findings section — so the title alone cannot tell them apart.
 */
const key = (section: Section) => String(section.heading.index);

/**
 * The section currently chosen, for panels that need the same heading levels — the
 * markdown tab nests its headings under whatever the Findings tab is pointing at.
 */
let chosen: Section | undefined;

export const selectedSection = (): Section | undefined => chosen;

export function setUpFindingsPanel(): void {
  const select = byId<HTMLSelectElement>("findings-section");
  const sort = byId<HTMLButtonElement>("findings-sort");
  const table = byId<HTMLButtonElement>("findings-table");
  const create = byId<HTMLButtonElement>("findings-new");
  const rescan = byId<HTMLButtonElement>("findings-rescan");
  const undo = byId<HTMLButtonElement>("findings-undo");
  const confirm = byId("findings-confirm");
  const confirmText = byId("findings-confirm-text");
  const confirmYes = byId<HTMLButtonElement>("findings-confirm-yes");
  const confirmNo = byId<HTMLButtonElement>("findings-confirm-no");

  const feedback: Feedback = feedbackFor("findings");
  const buttons = [sort, table, create, rescan, undo, confirmYes, confirmNo];

  let sections: Section[] = [];
  let undoable: Undoable | undefined;

  const selected = () => {
    chosen = sections.find((section) => key(section) === select.value);
    return chosen;
  };

  /** Sorting and tabulating need an existing section; creating a finding does not. */
  const updateAvailability = () => {
    const available = sections.length > 0;
    sort.disabled = !available;
    table.disabled = !available;
    select.disabled = !available;
  };

  /** Run an action, then restore which buttons make sense — guard() enables them all. */
  const act = (action: () => Promise<void>) =>
    guard(buttons, feedback, action).then(updateAvailability);

  const setUndoable = (next: Undoable | undefined) => {
    undoable = next;
    show(undo, next !== undefined);
    const labels = {
      sort: "Undo the last sort",
      table: "Remove the inserted table",
      finding: "Remove the new finding",
    };
    undo.textContent = next ? labels[next.kind] : "";
  };

  const hideConfirm = () => show(confirm, false);

  /** Re-read the document. `announce` is off after an action so its result stays on screen. */
  const scan = async (announce: boolean) => {
    const previous = select.value;
    sections = findSections(await getHeadings());

    select.textContent = "";
    sections.forEach((section) => {
      const option = make(
        "option",
        undefined,
        `${section.heading.text} (${section.findings.length})`
      );
      option.value = key(section);
      select.appendChild(option);
    });

    if (sections.some((section) => key(section) === previous)) {
      select.value = previous;
    }
    if (announce) {
      feedback.status(
        sections.length === 0
          ? 'No sections with findings yet — use "New finding at cursor" to start one.'
          : `${sections.length} section${sections.length === 1 ? "" : "s"} with findings.`
      );
    }
    selected();
    updateAvailability();
  };

  const refresh = () =>
    act(async () => {
      feedback.reset();
      hideConfirm();
      setUndoable(undefined);
      await scan(true);
    });

  const applySort = (section: Section) =>
    act(async () => {
      hideConfirm();
      const result = await sortFindings(section);
      feedback.status(
        result.changed
          ? `Sorted ${result.sorted} finding${result.sorted === 1 ? "" : "s"} by severity.`
          : "Already in severity order — nothing changed."
      );
      feedback.warnings(result.skipped.map((finding) => `Left in place: ${describe(finding)}`));
      setUndoable(
        result.snapshot ? { kind: "sort", section, snapshot: result.snapshot } : undefined
      );
      await scan(false);
    });

  const applyTable = (section: Section) =>
    act(async () => {
      hideConfirm();
      const result = await insertFindingsTable(section);
      feedback.status(`Inserted a table of ${result.rows} finding${result.rows === 1 ? "" : "s"}.`);
      feedback.warnings(
        result.skipped.map((finding) => `Listed without a severity: ${describe(finding)}`)
      );
      setUndoable({ kind: "table", bookmark: result.bookmark });
      await scan(false);
    });

  create.onclick = () =>
    act(async () => {
      const section = selected();
      feedback.reset();
      hideConfirm();
      setUndoable(undefined);

      const result = await insertFinding(section);
      feedback.status(
        `Inserted an empty finding at heading level ${result.level}` +
          (section ? `, matching "${section.heading.text}".` : ", the template's depth.") +
          " Fill in the [TODO] parts."
      );
      setUndoable({ kind: "finding", bookmark: result.bookmark });
      await scan(false);
    });

  /** Show what cannot be read and wait, rather than editing and reporting afterwards. */
  const askFirst = (
    section: Section,
    skipped: SkippedFinding[],
    message: string,
    label: string
  ) => {
    feedback.warnings(skipped.map(describe));
    feedback.status(message);
    confirmText.textContent =
      label === "Sort the rest anyway"
        ? "Fix the risk headings above, or sort the rest and leave those findings where they are."
        : "Fix the risk headings above, or insert the table with those rows left blank.";
    confirmYes.textContent = label;
    confirmYes.onclick = () =>
      label === "Sort the rest anyway" ? applySort(section) : applyTable(section);
    show(confirm, true);
  };

  sort.onclick = () =>
    act(async () => {
      const section = selected();
      feedback.reset();
      if (!section) {
        feedback.status("Select a section first — this document has none with findings yet.");
        return;
      }
      hideConfirm();
      setUndoable(undefined);

      const preview = await previewSort(section);
      if (preview.skipped.length > 0) {
        askFirst(
          section,
          preview.skipped,
          `${preview.skipped.length} of ${preview.skipped.length + preview.sorted} findings have no ` +
            "readable risk rating. The document has not been changed.",
          "Sort the rest anyway"
        );
        return;
      }
      if (!preview.changed) {
        feedback.status("Already in severity order — nothing changed.");
        return;
      }
      await applySort(section);
    });

  table.onclick = () =>
    act(async () => {
      const section = selected();
      feedback.reset();
      if (!section) {
        feedback.status("Select a section first — this document has none with findings yet.");
        return;
      }
      hideConfirm();
      setUndoable(undefined);

      const preview = await previewTable(section);
      if (preview.skipped.length > 0) {
        askFirst(
          section,
          preview.skipped,
          `${preview.skipped.length} of ${preview.rows} findings have no readable risk rating and ` +
            "would be listed without a severity. Nothing has been inserted.",
          "Insert the table anyway"
        );
        return;
      }
      await applyTable(section);
    });

  undo.onclick = () =>
    act(async () => {
      if (!undoable) {
        return;
      }
      feedback.reset();
      if (undoable.kind === "sort") {
        await restoreSection(undoable.section, undoable.snapshot);
        feedback.status("Undone — the findings are back in their original order.");
      } else if (undoable.kind === "table") {
        await removeFindingsTable(undoable.bookmark);
        feedback.status("Removed the inserted table.");
      } else {
        await removeFinding(undoable.bookmark);
        feedback.status("Removed the new finding.");
      }
      setUndoable(undefined);
      await scan(false);
    });

  rescan.onclick = refresh;
  confirmNo.onclick = hideConfirm;
  select.onchange = () => {
    selected();
    hideConfirm();
  };

  refresh();
}
