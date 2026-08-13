import { useCallback, useEffect, useRef, useState } from "react";
import { ParagraphPlan } from "../word/documentPlan";
import { ToDialog, ToPane, decode, encode, nextRequestId } from "../shared/protocol";

/* global Office, setTimeout, clearTimeout, setInterval, clearInterval */

/**
 * `Omit` over a union collapses it to the keys every member shares, which would erase the
 * protocol's variants. Distributing keeps each one intact.
 */
type WithoutId<T> = T extends { requestId: string } ? Omit<T, "requestId"> : never;

/** A reply must arrive within this long, or the pane is treated as gone. */
const REPLY_TIMEOUT = 15000;
const PING_TIMEOUT = 4000;
const PING_EVERY = 5000;

/**
 * Talking to the task pane, which is the only side that can touch the document.
 *
 * The pane can disappear — closed, reloaded by Word, or navigated away from — and nothing
 * announces it: `messageParent` simply goes nowhere and no reply ever arrives. Discovering
 * that only when a finished finding fails to insert would mean risking the writing, so a
 * heartbeat watches the channel and every request times out rather than hanging.
 */
export interface PaneChannel {
  insert(plans: ParagraphPlan[]): Promise<{
    bookmark: string;
    paragraphs: number;
    plainStyles: boolean;
  }>;
  remove(bookmark: string): Promise<void>;
  close(): void;
  /** False until Office is ready; the editor works before then, inserting does not. */
  ready: boolean;
  /** False when the pane has stopped answering. Nothing can be written until it returns. */
  connected: boolean;
}

export function usePane(): PaneChannel {
  const [ready, setReady] = useState(false);
  const [connected, setConnected] = useState(true);
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

        // Anything arriving at all means the pane is answering again.
        setConnected(true);

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

  const send = useCallback(
    (message: WithoutId<ToPane>, timeout = REPLY_TIMEOUT): Promise<ToDialog> =>
      new Promise((resolve, reject) => {
        const requestId = nextRequestId();

        const timer = setTimeout(() => {
          pending.current.delete(requestId);
          setConnected(false);
          reject(new Error("The task pane did not answer. It may have been closed."));
        }, timeout);

        pending.current.set(requestId, {
          resolve: (reply) => {
            clearTimeout(timer);
            resolve(reply);
          },
          reject: (error) => {
            clearTimeout(timer);
            reject(error);
          },
        });

        Office.context.ui.messageParent(encode({ ...message, requestId } as ToPane));
      }),
    []
  );

  /** Watch the channel, so a lost pane is noticed before a finding is written into it. */
  useEffect(() => {
    if (!ready) {
      return undefined;
    }

    const beat = () => {
      send({ kind: "ping" }, PING_TIMEOUT).catch(() => setConnected(false));
    };

    beat();
    const timer = setInterval(beat, PING_EVERY);
    return () => clearInterval(timer);
  }, [ready, send]);

  return {
    ready,
    connected,
    async insert(plans) {
      const reply = await send({ kind: "insert", plans });
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
