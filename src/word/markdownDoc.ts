import { planFromBlocks } from "./documentPlan";
import { parseMarkdown } from "./markdown";
import { WriteResult, removeWritten, writePlan } from "./writePlan";

const MARKDOWN_BOOKMARK = "_ptmd";

export interface MarkdownInsertion extends WriteResult {
  blocks: number;
}

/** Insert markdown at the cursor. `#` is Heading1, `##` is Heading2. */
export async function insertMarkdown(source: string): Promise<MarkdownInsertion> {
  const plans = planFromBlocks(parseMarkdown(source));
  if (plans.length === 0) {
    throw new Error("There is no markdown to insert.");
  }

  const written = await writePlan(plans, MARKDOWN_BOOKMARK);
  return { ...written, blocks: plans.length };
}

/** Remove what `insertMarkdown` wrote. */
export async function removeMarkdown(bookmark: string): Promise<void> {
  return removeWritten(bookmark);
}
