/**
 * Keeping the finding being written, so nothing is lost when the channel breaks.
 *
 * A dialog whose task pane has closed cannot be reconnected: `messageParent` targets the
 * pane instance that opened the dialog, and a reopened pane never registered a handler for
 * it. The only way back is to close this window and open the editor again — which would
 * throw away the draft unless it is kept somewhere. `localStorage` is same-origin and
 * survives the window, so it is exactly the right place.
 */

/** The part of `Storage` used here, so the store can be tested without a browser. */
export interface DraftStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
}

export interface DraftStore {
  save(markdown: string): void;
  /** The saved draft, or undefined when there is nothing worth restoring. */
  load(): string | undefined;
  clear(): void;
}

const KEY = "rainbowtable.draft";

export function createDraftStore(storage: DraftStorage | undefined, blank: string): DraftStore {
  return {
    save(markdown) {
      try {
        // An untouched skeleton is not a draft; restoring it would be noise.
        if (markdown.trim() === "" || markdown === blank) {
          storage?.removeItem(KEY);
        } else {
          storage?.setItem(KEY, markdown);
        }
      } catch {
        // Storage can be unavailable or full. Losing the backup must not stop the editor.
      }
    },

    load() {
      try {
        const saved = storage?.getItem(KEY);
        return saved && saved.trim() !== "" && saved !== blank ? saved : undefined;
      } catch {
        return undefined;
      }
    },

    clear() {
      try {
        storage?.removeItem(KEY);
      } catch {
        // As above.
      }
    },
  };
}
