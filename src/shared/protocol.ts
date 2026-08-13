import { ParagraphPlan } from "../word/documentPlan";

/**
 * The messages between the dialog and the task pane.
 *
 * A dialog runs in its own runtime with no access to the document, so everything it wants
 * written has to cross this channel and be carried out by the pane. The two sides share no
 * other code at runtime, so the shapes live here and both import them: an untyped string
 * channel between two bundles rots quietly.
 *
 * Size is not a concern — 8 MB crossed `messageParent` in 41 ms when measured, so nothing
 * here needs chunking.
 */

export type ToPane =
  /** Already planned: the dialog owns the parser and the grammars, the pane just writes. */
  | { kind: "insert"; requestId: string; plans: ParagraphPlan[] }
  | { kind: "remove"; requestId: string; bookmark: string }
  | { kind: "close"; requestId: string }
  /** Heartbeat: the dialog cannot otherwise tell that the pane has gone. */
  | { kind: "ping"; requestId: string };

export type ToDialog =
  | {
      kind: "inserted";
      requestId: string;
      bookmark: string;
      paragraphs: number;
      plainStyles: boolean;
    }
  | { kind: "removed"; requestId: string }
  | { kind: "failed"; requestId: string; reason: string }
  | { kind: "pong"; requestId: string };

let counter = 0;

/**
 * Ties a reply to the request that asked for it, since the dialog stays open across many.
 *
 * Counted rather than random: randomness alone collides often enough to matter — four
 * base-36 characters collide roughly once in every few hundred ids — and a collision here
 * means a reply applied to the wrong request. The random part only separates one dialog
 * session from another.
 */
export function nextRequestId(): string {
  counter += 1;
  return `r${counter}-${Math.random().toString(36).slice(2, 8)}`;
}

export function encode(message: ToPane | ToDialog): string {
  return JSON.stringify(message);
}

/**
 * Decode a message, or return undefined for anything that is not one of ours.
 *
 * The channel carries whatever a page sends, and the development probes put other things
 * on it, so this refuses rather than assuming.
 */
export function decode<T extends ToPane | ToDialog>(raw: string): T | undefined {
  if (!raw.startsWith("{")) {
    return undefined;
  }

  try {
    const message = JSON.parse(raw) as T;
    return typeof message.kind === "string" && typeof message.requestId === "string"
      ? message
      : undefined;
  } catch {
    return undefined;
  }
}
