import { HttpMessage, highlightHttp } from "./http";
import { colourOf } from "./httpColours";
import { planFromHttp } from "./documentPlan";
import { WriteResult, removeWritten, writePlan } from "./writePlan";

const BLOCK_BOOKMARK = "_pthttp";

export interface HttpBlockResult extends WriteResult {
  kind: HttpMessage["kind"];
  body: HttpMessage["body"];
  lines: number;
}

/**
 * Insert a highlighted HTTP message at the cursor, through the API rather than OOXML.
 *
 * Cost here scales with the number of runs rather than being a flat six seconds: a typical
 * response is far quicker this way, though a very large and heavily tokenised one has more
 * runs to write. See the measurements in CLAUDE.md.
 */
export async function insertHttpBlock(raw: string): Promise<HttpBlockResult> {
  if (!raw.trim()) {
    throw new Error("Paste an HTTP request or response first.");
  }

  const message = highlightHttp(raw);
  const written = await writePlan(planFromHttp(message, colourOf), BLOCK_BOOKMARK);

  return { ...written, kind: message.kind, body: message.body, lines: message.lines.length };
}

/** Remove a block inserted by `insertHttpBlock`. */
export async function removeHttpBlock(bookmark: string): Promise<void> {
  return removeWritten(bookmark);
}
