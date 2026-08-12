import test from "node:test";
import assert from "node:assert/strict";
import { parseMarkdown } from "../src/word/markdown";
import { buildMarkdown } from "../src/word/markdownDoc";

const build = (source: string) => buildMarkdown(parseMarkdown(source));

test("headings map straight across, # to Heading1", () => {
  const xml = build("# Finding\n\n## Description\n\n### Detail");

  assert.ok(xml.includes('<w:pStyle w:val="Heading1"/>'));
  assert.ok(xml.includes('<w:pStyle w:val="Heading2"/>'));
  assert.ok(xml.includes('<w:pStyle w:val="Heading3"/>'));
});

test("the package defines the heading styles it names", () => {
  // Without these the pStyle references are discarded and headings arrive as body text.
  const xml = build("# One\n\n###### Six");

  for (let level = 1; level <= 6; level++) {
    assert.ok(xml.includes(`w:styleId="Heading${level}"`), `Heading${level} is defined`);
    assert.ok(xml.includes(`<w:name w:val="heading ${level}"/>`), "under Word's built-in name");
  }
});

test("headings stop at the six levels markdown has", () => {
  const xml = build("###### Six");
  assert.ok(xml.includes('<w:pStyle w:val="Heading6"/>'));
  assert.ok(!/Heading[789]/.test(xml));
});

test("emphasis and inline code become runs, not literal markdown", () => {
  const xml = build("**bold** *italic* `code`");

  assert.ok(xml.includes("<w:b/>"), "bold");
  assert.ok(xml.includes("<w:i/>"), "italic");
  assert.ok(xml.includes('<w:rStyle w:val="Codeinline20"/>'), "the template's inline code style");
  assert.ok(!xml.includes("**"), "no markdown syntax survives into the document");
});

test("autolinks become clickable HYPERLINK fields", () => {
  const xml = build("see <https://example.com/a?b=1>");

  assert.ok(xml.includes("HYPERLINK &quot;https://example.com/a?b=1&quot;"));
  assert.ok(xml.includes('<w:rStyle w:val="Hyperlink"/>'), "and look like links");
  assert.ok(!xml.includes("&lt;https"), "the angle brackets are consumed");
});

test("lists carry numbering, and bullets differ from numbers", () => {
  const bullets = build("- one\n- two");
  const numbers = build("1. one\n2. two");

  assert.ok(bullets.includes('<w:numId w:val="880"/>'));
  assert.ok(numbers.includes('<w:numId w:val="881"/>'));
  assert.ok(bullets.includes('<w:pStyle w:val="ListParagraph"/>'));
});

test("the package defines its numbering, or Word drops the bullets", () => {
  const xml = build("- one");

  assert.ok(xml.includes('pkg:name="/word/numbering.xml"'), "a numbering part is present");
  assert.ok(xml.includes('w:abstractNumId="880"'));
  assert.ok(
    xml.includes(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"'
    )
  );
});

test("the package defines every style it names", () => {
  const xml = build("# T\n\n`c`\n\n- l\n\n```\nx\n```");

  assert.ok(xml.includes('pkg:name="/word/styles.xml"'));
  for (const style of ["Codeblock", "Codeinline20", "ListParagraph", "Hyperlink"]) {
    assert.ok(xml.includes(`w:styleId="${style}"`), `${style} is defined`);
  }
});

test("fenced code becomes one Codeblock paragraph per line, escaped", () => {
  const xml = build("```\n<script>a & b</script>\nsecond\n```");

  assert.equal(xml.match(/<w:pStyle w:val="Codeblock"\/>/g)?.length, 2, "one per line");
  assert.ok(xml.includes("&lt;script&gt;a &amp; b&lt;/script&gt;"));
});

test("markdown inside a code block is not interpreted", () => {
  const xml = build("```\n# not a heading\n**not bold**\n```");
  // Only the body: the styles part legitimately contains bold, in the heading definitions.
  const body = xml.slice(xml.indexOf("<w:body>"), xml.indexOf("</w:body>"));

  assert.ok(body.includes("# not a heading"), "kept verbatim");
  assert.ok(body.includes("**not bold**"), "asterisks kept as written");
  assert.ok(!body.includes("<w:b/>"), "no bold run");
  assert.ok(!body.includes('w:val="Heading1"'), "and no heading");
});

test("the result is a complete package ending in a spare paragraph", () => {
  const xml = build("# T");

  assert.ok(xml.startsWith("<pkg:package"));
  assert.ok(xml.includes("<w:p/></w:body>"), "so the last block cannot fuse with what follows");
  assert.ok(xml.trimEnd().endsWith("</pkg:package>"));
});
