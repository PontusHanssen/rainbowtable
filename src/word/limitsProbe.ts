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

/**
 * The fixed cost of talking to Word at all, with no content involved: one round trip that
 * reads a single number. If this is already seconds, no payload change can help.
 */
export async function timeEmptyRoundTrip(): Promise<number> {
  const started = Date.now();
  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/style");
    await context.sync();
  });
  return Date.now() - started;
}

/** One trivial paragraph through the OOXML path, to separate fixed cost from payload. */
export async function timeTinyInsert(): Promise<{ ms: number; bookmark: string }> {
  const started = Date.now();
  const bookmark = await insertProbePackage(
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
      '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">' +
      '<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/word/document.xml"/>' +
      "</Relationships></pkg:xmlData></pkg:part>" +
      '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">' +
      '<pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      "<w:body><w:p><w:r><w:t>probe</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>"
  );
  return { ms: Date.now() - started, bookmark };
}

/**
 * The same lines through the API instead of OOXML: `insertParagraph` chained, one sync.
 *
 * If this beats the OOXML path the workaround is to build content the way `newFinding`
 * does, rather than to shrink or reshape the package.
 */
export async function timeApiInsert(lines: string[]): Promise<{ ms: number; bookmark: string }> {
  const started = Date.now();
  const bookmark = "_ptprobeapi";

  await Word.run(async (context) => {
    let previous: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
    const inserted: Word.Paragraph[] = [];

    lines.forEach((line) => {
      previous = previous.insertParagraph(line, Word.InsertLocation.after);
      previous.styleBuiltIn = "Normal";
      inserted.push(previous);
    });

    inserted[0]
      .getRange("Whole")
      .expandTo(inserted[inserted.length - 1].getRange("Whole"))
      .insertBookmark(bookmark);
    await context.sync();
  });

  return { ms: Date.now() - started, bookmark };
}
