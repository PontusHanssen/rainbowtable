import { Section, findSections, getHeadings } from "../word/headings";
import { FindingSummary, SkippedFinding } from "../word/section";
import { previewSort, restoreSection, sortFindings } from "../word/sortFindings";
import { insertFindingsTable, previewTable, removeFindingsTable } from "../word/findingsTable";
import { insertFinding, removeFinding } from "../word/newFinding";
import { Feedback, byId, clear, feedbackFor, guard, make, show } from "./dom";

/* global HTMLButtonElement, HTMLOListElement, HTMLSelectElement */

/** A completed action that can still be taken back. Word's own undo cannot reach ours. */
type Undoable =
  | { kind: "sort"; section: Section; snapshot: string; written?: string[] }
  | { kind: "table"; bookmark: string }
  | { kind: "finding"; bookmark: string };

/**
 * An action the user has been shown and not yet agreed to.
 *
 * The kind is carried explicitly rather than inferred from the confirm button's label,
 * which is what this used to do — a string comparison against "Sort the rest anyway"
 * decided both the explanation and which function ran.
 */
type Pending = { kind: "sort" | "table"; section: Section };

const describe = ({ title, reason }: SkippedFinding) => `"${title}" — ${reason}`;

const plural = (count: number, noun: string) => `${count} ${noun}${count === 1 ? "" : "s"}`;

/**
 * Sections are identified by the paragraph index of their heading. Titles repeat in real
 * reports — the template has "Vulnerabilities" as both a summary subsection and a
 * findings section — so the title alone cannot tell them apart.
 */
const key = (section: Section) => String(section.heading.index);

/** One finding as a row of the preview: severity, score, title, and why it was skipped. */
function previewRow(finding: FindingSummary) {
  const row = make("li", "preview-row");

  const pill = make(
    "span",
    `pill sev-${finding.severity ? finding.severity.toLowerCase() : "none"}`,
    finding.severity ?? "No rating"
  );
  row.appendChild(pill);

  const title = make("span", "preview-title", finding.title);
  if (finding.score !== undefined) {
    title.textContent = `${finding.score.toFixed(1)} — ${finding.title}`;
  }
  row.appendChild(title);

  if (finding.reason) {
    row.appendChild(make("span", "preview-note", finding.reason));
  }
  return row;
}

export function setUpFindingsPanel(): void {
  const select = byId<HTMLSelectElement>("findings-section");
  const sort = byId<HTMLButtonElement>("findings-sort");
  const table = byId<HTMLButtonElement>("findings-table");
  const create = byId<HTMLButtonElement>("findings-new");
  const createEmpty = byId<HTMLButtonElement>("findings-new-empty");
  const rescan = byId<HTMLButtonElement>("findings-rescan");
  const undo = byId<HTMLButtonElement>("findings-undo");
  const empty = byId("findings-empty");
  const controls = byId("findings-controls");
  const pending = byId("findings-pending");
  const pendingTitle = byId("findings-pending-title");
  const pendingNote = byId("findings-pending-note");
  const previewList = byId<HTMLOListElement>("findings-preview");
  const confirmYes = byId<HTMLButtonElement>("findings-confirm-yes");
  const confirmNo = byId<HTMLButtonElement>("findings-confirm-no");

  const feedback: Feedback = feedbackFor("findings");
  const buttons = [sort, table, create, createEmpty, rescan, undo, confirmYes, confirmNo];

  let sections: Section[] = [];
  let undoable: Undoable | undefined;
  let pendingAction: Pending | undefined;

  const selected = () => sections.find((section) => key(section) === select.value);

  /** Sorting and tabulating need an existing section; creating a finding does not. */
  const updateAvailability = () => {
    const available = sections.length > 0;
    sort.disabled = !available;
    table.disabled = !available;
    select.disabled = !available;
    // A document with no findings yet gets an explanation and the one button that works,
    // rather than a row of dead controls above an empty dropdown.
    show(empty, !available);
    show(controls, available);
  };

  /** Run an action, then restore which buttons make sense — guard() enables them all. */
  const act = (busyLabel: string, action: () => Promise<void>) =>
    guard(buttons, feedback, action, busyLabel).then(updateAvailability);

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

  const hidePending = () => {
    pendingAction = undefined;
    show(pending, false);
  };

  /**
   * Show what an action would do and wait to be told to do it.
   *
   * Every finding appears in the list, including the ones whose risk could not be read —
   * naming them here is what lets the user fix the document instead of discovering
   * afterwards that some were left behind.
   */
  const askFirst = (
    action: Pending,
    findings: FindingSummary[],
    heading: string,
    note: string,
    confirmLabel: string
  ) => {
    pendingAction = action;
    pendingTitle.textContent = heading;
    pendingNote.textContent = note;
    confirmYes.textContent = confirmLabel;

    clear(previewList);
    findings.forEach((finding) => previewList.appendChild(previewRow(finding)));

    show(pending, true);
  };

  const applySort = (section: Section) =>
    act("Sorting the findings…", async () => {
      hidePending();
      const result = await sortFindings(section);
      feedback.status(
        result.changed
          ? `Sorted ${plural(result.sorted, "finding")} by severity.`
          : "Already in severity order — nothing changed."
      );
      feedback.warnings(result.skipped.map((finding) => `Left in place: ${describe(finding)}`));
      setUndoable(
        result.snapshot
          ? { kind: "sort", section, snapshot: result.snapshot, written: result.written }
          : undefined
      );
      await scan(false);
    });

  const applyTable = (section: Section) =>
    act("Building the table…", async () => {
      hidePending();
      const result = await insertFindingsTable(section);
      feedback.status(`Inserted a table of ${plural(result.rows, "finding")}.`);
      feedback.warnings(
        result.skipped.map((finding) => `Listed without a severity: ${describe(finding)}`)
      );
      setUndoable({ kind: "table", bookmark: result.bookmark });
      await scan(false);
    });

  /** Re-read the document. `announce` is off after an action so its result stays on screen. */
  const scan = async (announce: boolean) => {
    const previous = select.value;
    sections = findSections(await getHeadings());

    select.textContent = "";
    sections.forEach((section) => {
      const option = make(
        "option",
        undefined,
        `${section.heading.text} (${plural(section.findings.length, "finding")})`
      );
      option.value = key(section);
      select.appendChild(option);
    });

    if (sections.some((section) => key(section) === previous)) {
      select.value = previous;
    }
    if (announce && sections.length > 0) {
      feedback.status(`${plural(sections.length, "section")} with findings.`);
    }
    updateAvailability();
  };

  const refresh = () =>
    act("Reading the document…", async () => {
      feedback.reset();
      hidePending();
      setUndoable(undefined);
      await scan(true);
    });

  const newFinding = () =>
    act("Writing the finding…", async () => {
      const section = selected();
      feedback.reset();
      hidePending();
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

  sort.onclick = () =>
    act("Checking the findings…", async () => {
      const section = selected();
      feedback.reset();
      hidePending();
      setUndoable(undefined);
      if (!section) {
        feedback.status("Choose a section first.");
        return;
      }

      const preview = await previewSort(section);
      if (!preview.changed) {
        feedback.status("Already in severity order — nothing would move.");
        feedback.warnings(preview.skipped.map((finding) => `Cannot be read: ${describe(finding)}`));
        return;
      }

      askFirst(
        { kind: "sort", section },
        preview.order,
        "The order this would leave",
        "Nothing has been changed yet. Sorting rewrites the whole section, and Word's own " +
          "undo cannot reach it — use the Undo button here instead." +
          (preview.skipped.length > 0
            ? ` ${plural(preview.skipped.length, "finding")} above cannot be rated and will ` +
              "stay where they are."
            : ""),
        "Sort the findings"
      );
    });

  table.onclick = () =>
    act("Checking the findings…", async () => {
      const section = selected();
      feedback.reset();
      hidePending();
      setUndoable(undefined);
      if (!section) {
        feedback.status("Choose a section first.");
        return;
      }

      const preview = await previewTable(section);
      askFirst(
        { kind: "table", section },
        preview.findings,
        "The rows this table would carry",
        "Nothing has been inserted yet. The table goes in at the cursor." +
          (preview.skipped.length > 0
            ? ` ${plural(preview.skipped.length, "row")} above would be listed without a ` +
              "severity."
            : ""),
        "Insert the table"
      );
    });

  undo.onclick = () =>
    act("Putting it back…", async () => {
      if (!undoable) {
        return;
      }
      feedback.reset();
      if (undoable.kind === "sort") {
        await restoreSection(undoable.section, undoable.snapshot, undoable.written);
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

  confirmYes.onclick = () => {
    if (pendingAction) {
      const { kind, section } = pendingAction;
      void (kind === "sort" ? applySort(section) : applyTable(section));
    }
  };

  create.onclick = newFinding;
  createEmpty.onclick = newFinding;
  rescan.onclick = refresh;
  confirmNo.onclick = hidePending;
  select.onchange = hidePending;

  refresh();
}
