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
  reset(): void;
}

export function feedbackFor(panel: string): Feedback {
  const status = byId(`${panel}-status`);
  const warnings = byId(`${panel}-warnings`);
  const error = byId(`${panel}-error`);

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
    reset() {
      clear(status);
      clear(warnings);
      clear(error);
    },
  };
}

/**
 * Run a document action with the panel's buttons disabled, reporting failures rather
 * than letting them reach the console where nobody will see them.
 */
export async function guard(
  buttons: HTMLButtonElement[],
  feedback: Feedback,
  action: () => Promise<void>
): Promise<void> {
  buttons.forEach((button) => (button.disabled = true));
  feedback.error("");

  try {
    await action();
  } catch (err) {
    feedback.error(String(err));
  } finally {
    buttons.forEach((button) => (button.disabled = false));
  }
}
