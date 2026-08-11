/**
 * The heading outline of "Security Review Template V3.1.docx", the template the add-in
 * targets, with a body paragraph under each heading to give the blocks some width.
 *
 * Kept as a fixture rather than read from the .docx so the tests need no Word, no zip
 * reader and no copy of the template. If the real template's outline changes, update
 * this and the tests will tell you what the change breaks.
 */
export interface FixtureParagraph {
  styleBuiltIn: string;
  text: string;
}

const body = (text: string): FixtureParagraph => ({ styleBuiltIn: "Normal", text });
const heading = (level: number, text: string): FixtureParagraph => ({
  styleBuiltIn: `Heading${level}`,
  text,
});

/** A finding as the template lays one out: title, risk, status, then prose sections. */
function finding(title: string, risk: string): FixtureParagraph[] {
  return [
    heading(2, title),
    heading(3, risk),
    body("TODO https://nvd.nist.gov/vuln-metrics/cvss/v3-calculator"),
    heading(3, "Status: [TODO]"),
    heading(3, "Description"),
    body("[TODO]"),
    heading(3, "Recommendation"),
    body("[TODO]"),
  ];
}

export const templateParagraphs: FixtureParagraph[] = [
  heading(1, "Background"),
  body("lorem ipsum"),
  heading(2, "Level of security and compliance"),
  body("[TODO]"),
  heading(1, "Results"),
  heading(2, "Vulnerabilities"),
  body("[TODO]"),
  heading(1, "Vulnerabilities"),
  ...finding("TODO", "Risk: [TODO]"),
  heading(1, "Weaknesses"),
  ...finding("TODO", "Risk: [TODO]"),
  ...finding("[TODO Web1: Weak transport layer security]", "Risk: [TODO]"),
  ...finding("[TODO Web1: Missing security headers]", "Risk: [TODO]"),
  ...finding("[TODO Usage of outdated or vulnerable dependencies]", "Risk: [TODO]"),
  heading(1, "Appendix A – Automated analysis"),
  heading(2, "Analysis of TLS configuration"),
  heading(3, "[TODO System name]"),
  body("[TODO]"),
];

/** The same shape as the template, but with the risk headings filled in. */
export const filledParagraphs: FixtureParagraph[] = [
  heading(1, "Findings"),
  ...finding("XSS in Y", "Risk: Medium (5.4)"),
  ...finding("SQLi", "Risk: Critical (9.0)"),
  ...finding("SSRF A", "Risk: Medium (4.0)"),
];
