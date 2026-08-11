/**
 * Reordering findings inside a captured OOXML package.
 *
 * Moving blocks by deleting them and re-inserting each one separately does not work:
 * every seam between two inserted blocks merges, so the tail of one finding lands in the
 * next finding's heading. Capturing the whole findings region, reordering the paragraphs
 * inside that package and putting it back with a single Replace has no seams at all.
 *
 * The functions here are plain string surgery so they can be tested without a DOM.
 */

/** The end of the tag starting at `from`, ignoring `>` inside attribute values. */
function findTagEnd(xml: string, from: number): number {
  let quote = "";

  for (let i = from; i < xml.length; i++) {
    const character = xml[i];
    if (quote) {
      if (character === quote) {
        quote = "";
      }
    } else if (character === '"' || character === "'") {
      quote = character;
    } else if (character === ">") {
      return i;
    }
  }
  return xml.length - 1;
}

/**
 * The top-level elements of an XML fragment, each as its full source text.
 *
 * Text and whitespace between elements is dropped. That is safe here: between block
 * elements in a Word document body, whitespace carries no meaning.
 */
export function splitTopLevel(xml: string): string[] {
  const elements: string[] = [];
  let depth = 0;
  let start = -1;
  let i = 0;

  while (i < xml.length) {
    const open = xml.indexOf("<", i);
    if (open < 0) {
      break;
    }

    if (xml.startsWith("<!--", open)) {
      i = xml.indexOf("-->", open) + 3;
      continue;
    }
    if (xml.startsWith("<![CDATA[", open)) {
      i = xml.indexOf("]]>", open) + 3;
      continue;
    }
    if (xml.startsWith("<?", open) || xml.startsWith("<!", open)) {
      i = findTagEnd(xml, open) + 1;
      continue;
    }

    const end = findTagEnd(xml, open);

    if (xml[open + 1] === "/") {
      depth--;
      if (depth === 0 && start >= 0) {
        elements.push(xml.slice(start, end + 1));
        start = -1;
      }
    } else if (xml[end - 1] === "/") {
      if (depth === 0) {
        elements.push(xml.slice(open, end + 1));
      }
    } else {
      if (depth === 0) {
        start = open;
      }
      depth++;
    }

    i = end + 1;
  }

  return elements;
}

/** True when `element` is a paragraph carrying the built-in heading style for `level`. */
export function isHeadingParagraph(element: string, level: number): boolean {
  return (
    /^<w:p[\s>]/.test(element) &&
    new RegExp(`<w:pStyle\\s+w:val="Heading${level}"\\s*/?>`).test(element)
  );
}

/**
 * Split a document body into one segment per finding: a heading paragraph at
 * `headingLevel` plus everything after it until the next such heading. Content before the
 * first heading, and the trailing sectPr, are kept aside so they stay put.
 */
function splitIntoFindings(inner: string, headingLevel: number) {
  const segments: string[] = [];
  let prefix = "";
  let suffix = "";

  for (const element of splitTopLevel(inner)) {
    if (/^<w:sectPr[\s>]/.test(element)) {
      suffix += element;
    } else if (isHeadingParagraph(element, headingLevel)) {
      segments.push(element);
    } else if (segments.length === 0) {
      prefix += element;
    } else {
      segments[segments.length - 1] += element;
    }
  }

  return { prefix, segments, suffix };
}

/**
 * Reorder the findings inside an OOXML package captured from a findings region.
 *
 * `order` holds the original index of each finding in its new position, so `[2, 0, 1]`
 * means "the third finding comes first". The package is otherwise left untouched, which
 * is what keeps styles, numbering and embedded content intact.
 */
export function reorderFindings(pkg: string, headingLevel: number, order: number[]): string {
  // A package holds several parts, and more than one can contain a w:body — a glossary
  // document has its own. Start from the main document part so the right one is found.
  const documentPart = pkg.indexOf('pkg:name="/word/document.xml"');
  const searchFrom = documentPart < 0 ? 0 : documentPart;

  const bodyOpen = /<w:body(?:\s[^>]*)?>/.exec(pkg.slice(searchFrom));
  if (!bodyOpen) {
    throw new Error("The captured OOXML has no document body to reorder.");
  }

  const from = searchFrom + bodyOpen.index + bodyOpen[0].length;
  const bodyClose = pkg.indexOf("</w:body>", from);
  if (bodyClose < 0) {
    throw new Error("The captured OOXML has no document body to reorder.");
  }
  const { prefix, segments, suffix } = splitIntoFindings(pkg.slice(from, bodyClose), headingLevel);

  if (segments.length !== order.length) {
    throw new Error(
      `Expected ${order.length} findings in the captured content but found ${segments.length}.`
    );
  }

  const reordered = order.map((original) => segments[original]).join("");
  return pkg.slice(0, from) + prefix + reordered + suffix + pkg.slice(bodyClose);
}

/** XML-escape text destined for a w:t element or an attribute. */
export function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** A run of text, optionally with run properties (already-formed XML). */
export function run(text: string, properties = ""): string {
  const rPr = properties ? `<w:rPr>${properties}</w:rPr>` : "";
  return `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
}

/**
 * Wrap document body content in the minimal flat OPC package insertOoxml expects.
 *
 * Only the document part is included: anything needing a relationship (an external
 * hyperlink, an image) has to be expressed as a field instead, which is why links here
 * are HYPERLINK fields rather than w:hyperlink elements.
 */
export function wrapInPackage(body: string): string {
  return (
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
    '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">' +
    '<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/word/document.xml"/>' +
    "</Relationships></pkg:xmlData></pkg:part>" +
    '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">' +
    '<pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${body}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`
  );
}
