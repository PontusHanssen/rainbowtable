/* global document, HTMLElement, HTMLElementTagNameMap, HTMLButtonElement */

/** Small helpers so the panels read as intent rather than as DOM plumbing. */

export function byId<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) {
    throw new Error(`The task pane markup has no element with id "${id}".`);
  }
  return element as T;
}

export function show(element: HTMLElement, visible: boolean): void {
  element.classList.toggle("hidden", !visible);
}

export function clear(element: HTMLElement): void {
  element.textContent = "";
}

/** Create an element, optionally with a class and text. */
export function make<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
  text?: string
): HTMLElementTagNameMap[K] {
  const element = document.createElement(tag);
  if (className) {
    element.className = className;
  }
  if (text !== undefined) {
    element.textContent = text;
  }
  return element;
}

/**
 * The status, warning and error lines a panel reports through.
 *
 * Each panel owns three elements named `<panel>-status`, `-warnings` and `-error`, so
 * this wires all three from the panel's name.
 */
export interface Feedback {
  status(text: string): void;
  warnings(items: string[]): void;
  error(text: string): void;
  /** Say what is running, or pass nothing to say that nothing is. */
  busy(label?: string): void;
  reset(): void;
}

export function feedbackFor(panel: string): Feedback {
  const status = byId(`${panel}-status`);
  const warnings = byId(`${panel}-warnings`);
  const error = byId(`${panel}-error`);
  const busy = byId(`${panel}-busy`);
  const busyText = byId(`${panel}-busy-text`);

  const setBusy = (label?: string) => {
    busyText.textContent = label ?? "";
    show(busy, label !== undefined);
  };

  return {
    status(text) {
      status.textContent = text;
    },
    warnings(items) {
      clear(warnings);
      items.forEach((item) => warnings.appendChild(make("div", undefined, item)));
    },
    error(text) {
      error.textContent = text;
    },
    busy: setBusy,
    reset() {
      clear(status);
      clear(warnings);
      clear(error);
      setBusy();
    },
  };
}

/**
 * What to show the user for a thrown value.
 *
 * `String(err)` on an Error glues "Error: " to the front of it, and the messages this
 * code throws are written as sentences for the person reading them — an undo explaining
 * why it refused, a write explaining what it could not take back. The prefix makes them
 * read like a crash.
 */
export function messageOf(err: unknown): string {
  const text = err instanceof Error ? err.message : String(err);
  return text.replace(/^Error:\s*/, "");
}

/**
 * Run a document action with the panel's buttons disabled, reporting failures rather
 * than letting them reach the console where nobody will see them.
 *
 * `busyLabel` matters more than it looks: `insertOoxml` costs about six seconds whatever
 * it carries, and every findings action reads the whole document first. Dimmed buttons on
 * their own are indistinguishable from a pane that has stopped working.
 */
export async function guard(
  buttons: HTMLButtonElement[],
  feedback: Feedback,
  action: () => Promise<void>,
  busyLabel?: string
): Promise<void> {
  const progress = byId("progress");

  buttons.forEach((button) => (button.disabled = true));
  feedback.error("");
  if (busyLabel) {
    feedback.busy(busyLabel);
    show(progress, true);
  }

  try {
    await action();
  } catch (err) {
    feedback.error(messageOf(err));
  } finally {
    buttons.forEach((button) => (button.disabled = false));
    feedback.busy();
    show(progress, false);
  }
}
