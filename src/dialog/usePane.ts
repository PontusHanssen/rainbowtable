import { useEffect, useRef, useState } from "react";
import { ToDialog, ToPane, decode, encode, nextRequestId } from "../shared/protocol";

/**
 * `Omit` over a union collapses it to the keys every member shares, which would erase the
 * protocol's variants. Distributing keeps each one intact.
 */
type WithoutId<T> = T extends { requestId: string } ? Omit<T, "requestId"> : never;

/* global Office */

/**
 * Talking to the task pane, which is the only side that can touch the document.
 *
 * Requests are promises resolved when the matching reply arrives, so callers can await an
 * insert without knowing a message channel is involved. The dialog stays open across many
 * findings, which is why replies carry the id of the request that asked for them.
 */
export interface PaneChannel {
  insert(markdown: string): Promise<{ bookmark: string; paragraphs: number; plainStyles: boolean }>;
  remove(bookmark: string): Promise<void>;
  close(): void;
  /** False until Office is ready; the editor is usable before then, insertion is not. */
  ready: boolean;
}

export function usePane(): PaneChannel {
  const [ready, setReady] = useState(false);
  const pending = useRef(
    new Map<string, { resolve: (m: ToDialog) => void; reject: (e: Error) => void }>()
  );

  useEffect(() => {
    Office.onReady(() => {
      Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
        const reply = decode<ToDialog>((arg as { message: string }).message);
        if (!reply) {
          return;
        }

        const waiting = pending.current.get(reply.requestId);
        pending.current.delete(reply.requestId);

        if (reply.kind === "failed") {
          waiting?.reject(new Error(reply.reason));
        } else {
          waiting?.resolve(reply);
        }
      });
      setReady(true);
    });
  }, []);

  const send = (message: WithoutId<ToPane>): Promise<ToDialog> =>
    new Promise((resolve, reject) => {
      const requestId = nextRequestId();
      pending.current.set(requestId, { resolve, reject });
      Office.context.ui.messageParent(encode({ ...message, requestId } as ToPane));
    });

  return {
    ready,
    async insert(markdown) {
      const reply = await send({ kind: "insert", markdown });
      if (reply.kind !== "inserted") {
        throw new Error("Unexpected reply to an insert.");
      }
      return {
        bookmark: reply.bookmark,
        paragraphs: reply.paragraphs,
        plainStyles: reply.plainStyles,
      };
    },
    async remove(bookmark) {
      await send({ kind: "remove", bookmark });
    },
    close() {
      Office.context.ui.messageParent(encode({ kind: "close", requestId: nextRequestId() }));
    },
  };
}
