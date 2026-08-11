import { Section } from "./headings";

/**
 * Insert a severity-coloured findings table for a section at the current selection,
 * with columns #, Severity, Score, Title.
 *
 * Not implemented yet. See CLAUDE.md — the # and Title cells must be bookmark + REF
 * field cross-references carrying a cached result, so they stay clickable in Word and
 * survive PDF export.
 */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
export async function insertFindingsTable(_section: Section): Promise<void> {
  throw new Error("Inserting the findings table is not implemented yet.");
}
