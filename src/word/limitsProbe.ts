/* global Word */

/**
 * Insert a package as-is and bookmark it, for the shape comparison in the limits harness.
 * Development only: it takes prebuilt OOXML rather than markdown, which nothing else does.
 */
const PROBE_BOOKMARK = "_ptprobe";

export async function insertProbePackage(ooxml: string): Promise<string> {
  return Word.run(async (context) => {
    const inserted = context.document
      .getSelection()
      .insertOoxml(ooxml, Word.InsertLocation.replace);
    inserted.insertBookmark(PROBE_BOOKMARK);
    await context.sync();
    return PROBE_BOOKMARK;
  });
}
