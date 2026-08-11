import test from "node:test";
import assert from "node:assert/strict";
import { isHeadingParagraph, reorderFindings, splitTopLevel } from "../src/word/ooxml";

/** A finding as it appears in a captured package: heading, risk, prose, sometimes a table. */
function finding(title: string, risk: string, extra = ""): string {
  return (
    `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>${title}</w:t></w:r></w:p>` +
    `<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>${risk}</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t>Recommendation for ${title}.</w:t></w:r></w:p>` +
    extra
  );
}

/** The shape getOoxml() returns: a flat OPC package wrapping a document part. */
function packaged(body: string): string {
  return (
    `<?xml version="1.0" standalone="yes"?><pkg:package xmlns:pkg="http://x">` +
    `<pkg:part pkg:name="/word/styles.xml"><w:styles><w:body>decoy</w:body></w:styles></pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml"><w:document><w:body>${body}</w:body></w:document></pkg:part>` +
    `</pkg:package>`
  );
}

const titles = (xml: string) => [...xml.matchAll(/<w:t>([^<]*)<\/w:t>/g)].map((m) => m[1]);

test("splitTopLevel returns whole elements, ignoring nesting", () => {
  const elements = splitTopLevel("<a><b><c/></b></a><d/><e>text</e>");
  assert.deepEqual(elements, ["<a><b><c/></b></a>", "<d/>", "<e>text</e>"]);
});

test("splitTopLevel is not fooled by > inside attribute values", () => {
  const elements = splitTopLevel('<w:p w:rsidR="a>b"><w:r/></w:p><w:tbl/>');
  assert.deepEqual(elements, ['<w:p w:rsidR="a>b"><w:r/></w:p>', "<w:tbl/>"]);
});

test("splitTopLevel skips comments and processing instructions", () => {
  assert.deepEqual(splitTopLevel("<!-- <x/> --><?pi ?><a/>"), ["<a/>"]);
});

test("splitTopLevel handles repeated sibling elements of the same name", () => {
  assert.equal(splitTopLevel("<w:p><w:r/></w:p><w:p><w:r/></w:p><w:p/>").length, 3);
});

test("isHeadingParagraph matches only paragraphs at the given level", () => {
  const heading2 = '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr></w:p>';
  assert.ok(isHeadingParagraph(heading2, 2));
  assert.ok(!isHeadingParagraph(heading2, 3));
  assert.ok(!isHeadingParagraph('<w:tbl><w:pStyle w:val="Heading2"/></w:tbl>', 2));
  assert.ok(!isHeadingParagraph("<w:p><w:r><w:t>plain</w:t></w:r></w:p>", 2));
});

test("reorderFindings moves each finding with all of its content", () => {
  const pkg = packaged(
    finding("XSS in Y", "Risk: Medium (5.4)") +
      finding("SQLi", "Risk: Critical (9.0)") +
      finding("SSRF A", "Risk: Medium (4.0)")
  );

  const sorted = reorderFindings(pkg, 2, [1, 0, 2]);

  assert.deepEqual(titles(sorted), [
    "SQLi",
    "Risk: Critical (9.0)",
    "Recommendation for SQLi.",
    "XSS in Y",
    "Risk: Medium (5.4)",
    "Recommendation for XSS in Y.",
    "SSRF A",
    "Risk: Medium (4.0)",
    "Recommendation for SSRF A.",
  ]);
});

test("reorderFindings keeps tables attached to the finding they belong to", () => {
  const table =
    "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>evidence</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
  const pkg = packaged(
    finding("First", "Risk: Low (1.0)", table) + finding("Second", "Risk: Critical (9.0)")
  );

  const sorted = reorderFindings(pkg, 2, [1, 0]);

  assert.ok(sorted.indexOf("evidence") > sorted.indexOf("Second"), "table moved with First");
  assert.equal(sorted.match(/evidence/g)?.length, 1, "table is not duplicated");
});

test("reorderFindings leaves the rest of the package untouched", () => {
  const pkg = packaged(finding("A", "Risk: Low (1.0)") + finding("B", "Risk: High (8.0)"));
  const sorted = reorderFindings(pkg, 2, [1, 0]);

  assert.ok(sorted.startsWith('<?xml version="1.0" standalone="yes"?>'));
  assert.ok(sorted.includes('<pkg:part pkg:name="/word/styles.xml">'));
  assert.ok(sorted.endsWith("</pkg:package>"));
});

test("reorderFindings keeps a trailing sectPr at the end of the body", () => {
  const pkg = packaged(
    finding("A", "Risk: Low (1.0)") +
      finding("B", "Risk: High (8.0)") +
      "<w:sectPr><w:pgSz/></w:sectPr>"
  );

  const sorted = reorderFindings(pkg, 2, [1, 0]);
  assert.ok(sorted.indexOf("<w:sectPr>") > sorted.indexOf("Recommendation for A."));
  assert.ok(sorted.indexOf("<w:sectPr>") < sorted.lastIndexOf("</w:body>"));
});

test("reorderFindings is a no-op for the identity order", () => {
  const pkg = packaged(finding("A", "Risk: Critical (9.0)") + finding("B", "Risk: Low (1.0)"));
  assert.equal(reorderFindings(pkg, 2, [0, 1]), pkg);
});

test("reorderFindings refuses to guess when the content does not match the plan", () => {
  const pkg = packaged(finding("Only one", "Risk: Low (1.0)"));
  assert.throws(() => reorderFindings(pkg, 2, [1, 0]), /Expected 2 findings .* found 1/);
});

test("reorderFindings rejects a package with no document body", () => {
  assert.throws(() => reorderFindings("<pkg:package/>", 2, []), /no document body/);
});
