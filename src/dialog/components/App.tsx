import { ReactElement, useEffect, useMemo, useState } from "react";
import { Editor } from "./Editor";
import { Cvss } from "./Cvss";
import { Preview } from "./Preview";
import { applyScore } from "../applyScore";
import { usePane } from "../usePane";
import { createDraftStore } from "../draft";

/* global navigator, localStorage */
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

export function App(): ReactElement {
  const pane = usePane();
  const drafts = useMemo(
    () =>
      createDraftStore(typeof localStorage === "undefined" ? undefined : localStorage, SKELETON),
    []
  );

  // The draft outlives this window, so a pane that closed mid-finding costs no writing.
  const [markdown, setMarkdown] = useState(() => drafts.load() ?? SKELETON);
  const [restored] = useState(() => drafts.load() !== undefined);
  const [preview, setPreview] = useState(false);
  const [written, setWritten] = useState<Written[]>([]);
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("");
  const [error, setError] = useState("");

  useEffect(() => drafts.save(markdown), [drafts, markdown]);

  const insert = async () => {
    setBusy(true);
    setError("");
    try {
      const result = await pane.insert(await planMarkdown(markdown));
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
      // Ready for the next one: the editor stays open across a whole report.
      drafts.clear();
      setMarkdown(SKELETON);
    } catch (err) {
      setError(String(err));
    } finally {
      setBusy(false);
    }
  };

  const remove = async (entry: Written) => {
    setBusy(true);
    setError("");
    try {
      await pane.remove(entry.bookmark);
      setWritten((all) => all.filter((other) => other !== entry));
      setStatus(`Removed “${entry.title}”.`);
    } catch (err) {
      setError(String(err));
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="app">
      {!pane.connected && (
        <div className="disconnected" role="alert">
          <strong>The task pane is not answering.</strong> Nothing can be inserted. If the pane was
          closed, this window cannot reconnect to a new one — close it and open the editor again
          from the pane in Word. <strong>Your draft is saved</strong> and will come back when you
          do.
          <button type="button" onClick={() => navigator.clipboard?.writeText(markdown)}>
            Copy the markdown
          </button>
        </div>
      )}

      {restored && <p className="status">Restored the draft you were writing.</p>}

      <Cvss
        onApply={(risk, vector) => {
          setMarkdown((current) => applyScore(current, risk, vector));
          setStatus(`Scored ${risk}.`);
        }}
      />

      <nav className="tabs">
        <span className="spacer" />
        <button
          type="button"
          className="tab"
          aria-selected={preview}
          onClick={() => setPreview((on) => !on)}
        >
          Preview
        </button>
      </nav>

      <div className={preview ? "split" : "single"}>
        <Editor value={markdown} onChange={setMarkdown} />
        {preview && <Preview markdown={markdown} />}
      </div>

      {written.length > 0 && (
        <div className="written">
          <span className="label">Written this session</span>
          {written.map((entry) => (
            <div className="entry" key={entry.bookmark}>
              <span>
                {entry.title} <span className="muted">({entry.paragraphs} paragraphs)</span>
              </span>
              <button type="button" disabled={busy} onClick={() => remove(entry)}>
                Remove
              </button>
            </div>
          ))}
        </div>
      )}

      {status && !error && <p className="status">{status}</p>}
      {error && <p className="error">{error}</p>}

      <div className="actions">
        <button type="button" onClick={() => pane.close()}>
          Close
        </button>
        <button
          type="button"
          className="primary"
          disabled={busy || !pane.ready || !pane.connected}
          title={pane.connected ? undefined : "The task pane is not answering"}
          onClick={insert}
        >
          Insert at the cursor
        </button>
      </div>
    </div>
  );
}
