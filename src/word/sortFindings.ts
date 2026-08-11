import { Section } from "./headings";

/**
 * Reorder a section's findings by severity: Critical, High, Medium, Low,
 * Informational; higher score first within a severity; stable otherwise.
 *
 * Not implemented yet. See CLAUDE.md for the constraints that shape it — whole-block
 * moves via OOXML, strict `Risk: <Severity> (<score>)` parsing, and unparseable
 * findings reported rather than guessed at.
 */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
export async function sortFindings(_section: Section): Promise<void> {
  throw new Error("Sorting findings is not implemented yet.");
}
