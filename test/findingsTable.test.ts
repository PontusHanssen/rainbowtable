import test from "node:test";
import assert from "node:assert/strict";
import { childHeadings, headingNumber, toHeadings } from "../src/word/headings";
import { buildBlocks } from "../src/word/section";
import {
  FindingRow,
  bookmarkName,
  buildFindingsTable,
  buildRows,
  existingBookmark,
  tableBookmarkName,
} from "../src/word/findingsTable";
import { filledParagraphs, templateParagraphs } from "./fixtures/template";

const row = (overrides: Partial<FindingRow> = {}): FindingRow => ({
  bookmark: "_ptfdeadbeef0",
  number: "4.1",
  severity: "Critical",
  score: 9.1,
  title: "Weak transport layer security",
  ...overrides,
});

test("headingNumber counts ordinals into a dotted number", () => {
  const headings = toHeadings(templateParagraphs);
  const number = (text: string, level: number) =>
    headingNumber(
      headings,
      headings.findIndex((heading) => heading.text === text && heading.level === level)
    );

  assert.equal(number("Background", 1), "1");
  assert.equal(number("Results", 1), "2");
  assert.equal(number("Vulnerabilities", 1), "3");
  assert.equal(number("Weaknesses", 1), "4");
  assert.equal(number("[TODO Web1: Missing security headers]", 2), "4.3");
});

test("headingNumber restarts the deeper level under each new parent", () => {
  const headings = toHeadings([
    { styleBuiltIn: "Heading1", text: "One" },
    { styleBuiltIn: "Heading2", text: "One.One" },
    { styleBuiltIn: "Heading3", text: "One.One.One" },
    { styleBuiltIn: "Heading1", text: "Two" },
    { styleBuiltIn: "Heading2", text: "Two.One" },
  ]);

  assert.deepEqual(
    headings.map((_, i) => headingNumber(headings, i)),
    ["1", "1.1", "1.1.1", "2", "2.1"]
  );
});

test("bookmark names are stable per finding and independent of position", () => {
  const first = bookmarkName("Weaknesses", "Weak transport layer security");
  const again = bookmarkName("Weaknesses", "Weak transport layer security");

  assert.equal(first, again, "same finding, same name after a sort moves it");
  assert.notEqual(first, bookmarkName("Weaknesses", "Missing security headers"));
  assert.notEqual(first, bookmarkName("Vulnerabilities", "Weak transport layer security"));
});

test("bookmark names satisfy Word's rules", () => {
  const names = [
    bookmarkName("Weaknesses", "Outdated & vulnerable dependencies (npm)"),
    tableBookmarkName("Appendix A – Automated analysis"),
  ];

  for (const name of names) {
    assert.match(name, /^_[A-Za-z0-9_]+$/, `${name} is alphanumeric and hidden`);
    assert.ok(name.length <= 40, `${name} is within 40 characters`);
  }
});

test("duplicate finding titles in a section get distinct bookmarks", () => {
  const headings = toHeadings(templateParagraphs);
  const position = headings.findIndex((h) => h.text === "Weaknesses" && h.level === 1);
  const blocks = buildBlocks(
    headings,
    position,
    childHeadings(headings, position),
    templateParagraphs.length
  );

  const rows = buildRows(headings, "Weaknesses", blocks);
  assert.equal(new Set(rows.map((r) => r.bookmark)).size, rows.length);
});

test("rows carry the severity, score and number of each finding", () => {
  const headings = toHeadings(filledParagraphs);
  const position = headings.findIndex((h) => h.text === "Findings");
  const blocks = buildBlocks(
    headings,
    position,
    childHeadings(headings, position),
    filledParagraphs.length
  );

  assert.deepEqual(
    buildRows(headings, "Findings", blocks).map((r) => [r.number, r.severity, r.score, r.title]),
    [
      ["1.1", "Medium", 5.4, "XSS in Y"],
      ["1.2", "Critical", 9, "SQLi"],
      ["1.3", "Medium", 4, "SSRF A"],
    ]
  );
});

test("the # and Title cells are hyperlinked REF fields with cached results", () => {
  const xml = buildFindingsTable([row()]);

  assert.ok(
    xml.includes('<w:fldSimple w:instr=" REF _ptfdeadbeef0 \\w \\h ">'),
    "# cell references the bookmark with full-context numbering"
  );
  assert.ok(
    xml.includes('<w:fldSimple w:instr=" REF _ptfdeadbeef0 \\h ">'),
    "title cell references the bookmark"
  );
  assert.ok(xml.includes(">4.1<"), "the number is cached inside the field");
  assert.ok(
    xml.includes(">Weak transport layer security<"),
    "the title is cached inside the field"
  );
});

test("each severity carries the template's own character style", () => {
  const styles = {
    Critical: "Critical",
    High: "High",
    Medium: "Medium",
    Low: "Low",
    Informational: "Info",
  } as const;

  for (const [severity, style] of Object.entries(styles)) {
    const xml = buildFindingsTable([row({ severity: severity as FindingRow["severity"] })]);
    assert.ok(
      xml.includes(`<w:rStyle w:val="${style}"/>`),
      `${severity} should use the ${style} character style`
    );
  }
});

test("the severity cell is coloured by style, not by shading", () => {
  // The template colours the text; shading the cell would not be the house style.
  const xml = buildFindingsTable([row({ severity: "Critical" })]);
  const severityCell = xml.slice(xml.indexOf("<w:tr><w:tc"), xml.indexOf("</w:tbl>"));

  assert.ok(!severityCell.includes('w:fill="A50021"'), "no fill on the severity cell");
});

test("the table uses the template's table style and enables its header row", () => {
  const xml = buildFindingsTable([row()]);

  assert.ok(xml.includes('<w:tblStyle w:val="Omegapointtabellbl"/>'));
  // Without firstRow in tblLook the style's header banner never appears.
  assert.ok(xml.includes('w:firstRow="1"'));
  assert.ok(!xml.includes('w:fill="D9D9D9"'), "the style paints the header, not us");
});

test("the package defines the styles it names, or Word discards them", () => {
  const xml = buildFindingsTable([row()]);

  assert.ok(xml.includes('pkg:name="/word/styles.xml"'), "a styles part is present");
  assert.ok(xml.includes('w:styleId="Omegapointtabellbl"'));
  for (const style of ["Critical", "High", "Medium", "Low", "Info"]) {
    assert.ok(xml.includes(`w:styleId="${style}"`), `${style} is defined as a fallback`);
  }
});

test("a finding with no readable risk is still listed, without severity or score", () => {
  const xml = buildFindingsTable([row({ severity: undefined, score: undefined })]);

  assert.ok(xml.includes(">—<"), "severity shows a dash");
  assert.ok(xml.includes(">Weak transport layer security<"), "the finding is still in the table");
});

test("scores render with one decimal, the CVSS convention", () => {
  assert.ok(buildFindingsTable([row({ score: 9 })]).includes(">9.0<"));
  assert.ok(buildFindingsTable([row({ score: 5.4 })]).includes(">5.4<"));
});

test("titles are XML-escaped", () => {
  const xml = buildFindingsTable([row({ title: 'Ampersand & "quotes" <tag>' })]);

  assert.ok(xml.includes("Ampersand &amp; &quot;quotes&quot; &lt;tag&gt;"));
  assert.ok(!xml.includes("<tag>"), "no raw markup leaks into the document");
});

test("the table is a complete package with a header row per column", () => {
  const xml = buildFindingsTable([row(), row({ severity: "Low", score: 2.1, title: "Second" })]);

  assert.ok(xml.startsWith("<pkg:package"), "insertOoxml needs a flat OPC package");
  assert.ok(xml.includes("<w:tblHeader/>"), "the header row repeats across pages");
  assert.equal(xml.match(/<w:tr>/g)?.length, 3, "one header row plus two findings");
  assert.equal(xml.match(/<w:gridCol/g)?.length, 4, "four columns");
  assert.ok(xml.trimEnd().endsWith("</pkg:package>"));
});

test("a bookmark already on the heading wins over a freshly derived name", () => {
  const headings = toHeadings(filledParagraphs);
  const position = headings.findIndex((h) => h.text === "Findings");
  const blocks = buildBlocks(
    headings,
    position,
    childHeadings(headings, position),
    filledParagraphs.length
  );

  const rows = buildRows(headings, "Findings", blocks, ["_ptfaaaaaaaa0", undefined, undefined]);

  assert.equal(rows[0].bookmark, "_ptfaaaaaaaa0", "the existing bookmark survives a rename");
  assert.equal(rows[1].bookmark, bookmarkName("Findings", "SQLi"), "others are derived as usual");
});

test("existingBookmark picks only our own bookmarks, deterministically", () => {
  assert.equal(existingBookmark(["_Ref12345", "_ptfbbbb0", "_GoBack"]), "_ptfbbbb0");
  assert.equal(existingBookmark(["_ptfcccc0", "_ptfaaaa0"]), "_ptfaaaa0", "stable choice");
  assert.equal(existingBookmark(["_Ref12345", "_GoBack"]), undefined);
  assert.equal(existingBookmark([]), undefined);
});
