import { ReactElement, useState } from "react";
import { Editor } from "./Editor";
import { Cvss } from "./Cvss";
import { Preview } from "./Preview";
import { applyScore } from "../applyScore";
import { usePane } from "../usePane";
import { CvssVector, DEFAULT_VECTOR } from "../../word/cvss";

/* global navigator, window */
import { planMarkdown } from "../planMarkdown";

const SKELETON = [
  "## Title",
  "",
  "### Risk:",
  "",
  "### Technical Details",
  "",
  "### Proof of Concept",
  "",
  "### Recommendation",
  "",
].join("\n");

interface Written {
  bookmark: string;
  title: string;
  paragraphs: number;
}

/** The first heading, which is what the finding is called. */
function titleOf(markdown: string): string {
  const heading = markdown.split("\n").find((line) => /^#{1,6}\s+\S/.test(line));
  return heading?.replace(/^#+\s*/, "").trim() || "Untitled finding";
}

/**
 * What to show for a thrown value. `String(err)` glues "Error: " to the front, and the
 * messages the pane sends back are written as sentences for the person reading them.
 */
function messageOf(err: unknown): string {
  const text = err instanceof Error ? err.message : String(err);
  return text.replace(/^Error:\s*/, "");
}

export function App(): ReactElement {
  const pane = usePane();
  const [markdown, setMarkdown] = useState(SKELETON);
  const [vector, setVector] = useState<CvssVector>({ ...DEFAULT_VECTOR });
  const [preview, setPreview] = useState(false);
  const [highlight, setHighlight] = useState(true);
  const [written, setWritten] = useState<Written[]>([]);
  const [busy, setBusy] = useState("");
  const [status, setStatus] = useState("");
  const [error, setError] = useState("");
  const [focusToken, setFocusToken] = useState(0);

  /** Text worth warning about before it is thrown away. */
  const unsaved = markdown.trim() !== "" && markdown.trim() !== SKELETON.trim();

  const insert = async () => {
    setBusy("Inserting…");
    setError("");
    try {
      const result = await pane.insert(await planMarkdown(markdown, { highlight }));
      setWritten((all) => [
        ...all,
        { bookmark: result.bookmark, title: titleOf(markdown), paragraphs: result.paragraphs },
      ]);
      setStatus(
        `Inserted ${result.paragraphs} paragraphs.` +
          (result.plainStyles
            ? " The document has no template styles, so plain ones were used."
            : "")
      );
      // Ready for the next one: the editor stays open across a whole report. The score
      // goes with it — carrying the last finding's rating into the next one silently is
      // how a finding ends up mis-scored.
      setMarkdown(SKELETON);
      setVector({ ...DEFAULT_VECTOR });
      setFocusToken((token) => token + 1);
    } catch (err) {
      setError(messageOf(err));
    } finally {
      setBusy("");
    }
  };

  const remove = async (entry: Written) => {
    setBusy("Removing…");
    setError("");
    try {
      await pane.remove(entry.bookmark);
      setWritten((all) => all.filter((other) => other !== entry));
      setStatus(`Removed “${entry.title}”.`);
    } catch (err) {
      setError(messageOf(err));
    } finally {
      setBusy("");
    }
  };

  /*
   * The recovery path for unsaved work, so it has to say whether it worked. It used to
   * no-op silently where `navigator.clipboard` is absent and swallow a rejected
   * permission — the two cases where knowing matters most.
   */
  const copy = async () => {
    try {
      await navigator.clipboard.writeText(markdown);
      setStatus("Copied the markdown to the clipboard.");
    } catch {
      setError("Could not reach the clipboard. Select the text in the editor and copy it.");
    }
  };

  const close = () => {
    if (unsaved && !window.confirm("Close the editor? The finding you are writing is not saved anywhere.")) {
      return;
    }
    pane.close();
  };

  return (
    <div className="app">
      {!pane.connected && (
        <div className="disconnected" role="alert">
          <strong>The task pane is not answering.</strong> Nothing can be inserted, and this window
          cannot reconnect to a pane that was closed.{" "}
          <strong>Copy your text out before closing this window</strong> — nothing here is stored
          anywhere — then reopen the editor from the pane in Word.
          <button type="button" onClick={copy}>
            Copy the markdown
          </button>
        </div>
      )}

      <Cvss
        vector={vector}
        onChange={setVector}
        onApply={(risk, applied) => {
          setMarkdown((current) => applyScore(current, risk, applied));
          setStatus(`Scored ${risk}.`);
        }}
      />

      <div className="toolbar">
        <label className="check" title="Colour ```language fences in the document">
          <input
            type="checkbox"
            checked={highlight}
            onChange={(event) => setHighlight(event.target.checked)}
          />
          Syntax highlighting
        </label>
        <span className="spacer" />
        <button
          type="button"
          className="choice"
          aria-pressed={preview}
          onClick={() => setPreview((on) => !on)}
        >
          Preview
        </button>
      </div>

      <div className={preview ? "split" : "single"}>
        <Editor value={markdown} onChange={setMarkdown} focusToken={focusToken} />
        {preview && <Preview markdown={markdown} />}
      </div>

      {written.length > 0 && (
        <div>
          <div className="written-head">Written this session</div>
          <div className="written">
            {written.map((entry) => (
              <div className="entry" key={entry.bookmark}>
                <span>
                  {entry.title} <span className="muted">({entry.paragraphs} paragraphs)</span>
                </span>
                <button type="button" disabled={busy !== ""} onClick={() => remove(entry)}>
                  Remove
                </button>
              </div>
            ))}
          </div>
        </div>
      )}

      <div className="dialog-actions">
        {/* Status and error are shown together: a stale error used to hide every later
            result until the next insert cleared it. */}
        <div className="feedback" role="status" aria-live="polite">
          {busy && (
            <p className="busy">
              <span className="spinner" aria-hidden="true" />
              {busy}
            </p>
          )}
          {status && !busy && <p className="status">{status}</p>}
          {error && <p className="error">{error}</p>}
        </div>

        <button type="button" onClick={close}>
          Close
        </button>
        {/* Disabled for three different reasons, so it says which one. */}
        <button
          type="button"
          className="primary"
          disabled={busy !== "" || !pane.ready || !pane.connected}
          title={
            !pane.connected
              ? "The task pane is not answering"
              : !pane.ready
                ? "Still connecting to the task pane"
                : undefined
          }
          onClick={insert}
        >
          Insert at the cursor
        </button>
      </div>
    </div>
  );
}
