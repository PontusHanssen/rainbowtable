import { ReactElement, useEffect, useRef } from "react";
import { EditorState } from "@codemirror/state";
import { EditorView, keymap, lineNumbers, highlightActiveLine } from "@codemirror/view";
import { defaultKeymap, history, historyKeymap, indentWithTab } from "@codemirror/commands";
import { markdown } from "@codemirror/lang-markdown";
import { HighlightStyle, syntaxHighlighting } from "@codemirror/language";
import { tags } from "@lezer/highlight";

const editorHighlightStyle = HighlightStyle.define([
  { tag: [tags.keyword, tags.controlKeyword, tags.moduleKeyword], color: "var(--editor-keyword)" },
  { tag: [tags.string, tags.special(tags.string)], color: "var(--editor-string)" },
  { tag: [tags.number, tags.bool, tags.null], color: "var(--editor-number)" },
  { tag: [tags.comment, tags.lineComment, tags.blockComment], color: "var(--editor-comment)" },
  { tag: [tags.variableName, tags.propertyName, tags.attributeName], color: "var(--editor-name)" },
  { tag: [tags.typeName, tags.tagName, tags.className], color: "var(--editor-type)" },
  { tag: [tags.operator, tags.punctuation, tags.bracket], color: "var(--editor-operator)" },
  { tag: tags.heading, color: "var(--editor-keyword)", fontWeight: "bold" },
  { tag: tags.strong, fontWeight: "bold" },
  { tag: tags.emphasis, fontStyle: "italic" },
]);

/**
 * The markdown editor.
 *
 * CodeMirror rather than a textarea: it brings the shortcuts, list continuation and
 * syntax highlighting that were otherwise going to be hand-rolled against
 * `selectionStart`, and it is the same machinery that will colour fenced code.
 */
export function Editor({
  value,
  onChange,
  focusToken,
}: {
  value: string;
  onChange: (next: string) => void;
  /**
   * Bumped by the parent to ask for focus back — after an insert, which empties the
   * editor to the skeleton and would otherwise leave you clicking into it before you can
   * type the next finding.
   */
  focusToken?: number;
}): ReactElement {
  const host = useRef<HTMLDivElement>(null);
  const view = useRef<EditorView | undefined>(undefined);
  const latest = useRef(onChange);
  latest.current = onChange;

  useEffect(() => {
    if (!host.current) {
      return undefined;
    }

    const editor = new EditorView({
      parent: host.current,
      state: EditorState.create({
        doc: value,
        extensions: [
          lineNumbers(),
          highlightActiveLine(),
          history(),
          markdown(),
          syntaxHighlighting(editorHighlightStyle, { fallback: true }),
          // indentWithTab last: Tab should indent here rather than leave the editor.
          keymap.of([...defaultKeymap, ...historyKeymap, indentWithTab]),
          EditorView.lineWrapping,
          EditorView.updateListener.of((update) => {
            if (update.docChanged) {
              latest.current(update.state.doc.toString());
            }
          }),
        ],
      }),
    });

    view.current = editor;
    editor.focus();
    return () => editor.destroy();
    // Created once: the document is owned by CodeMirror from here on.
  }, []);

  /** Only push back in when something other than typing changed the text. */
  useEffect(() => {
    const editor = view.current;
    if (editor && value !== editor.state.doc.toString()) {
      editor.dispatch({
        changes: { from: 0, to: editor.state.doc.length, insert: value },
      });
    }
  }, [value]);

  useEffect(() => {
    if (focusToken) {
      view.current?.focus();
    }
  }, [focusToken]);

  return <div className="editor" ref={host} />;
}
