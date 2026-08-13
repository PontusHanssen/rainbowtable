/**
 * Writing a CVSS score into the markdown being edited.
 *
 * The finding's `Risk:` heading is rewritten in place and the vector put on the line below
 * it, which is exactly what the existing markdown pipeline turns into a heading Word can
 * parse and a clickable link. Pure, so it is tested.
 */

const RISK_HEADING = /^(#{1,6})\s*Risk:.*$/i;
const VECTOR_LINE = /^\s*<?https:\/\/nvd\.nist\.gov\/vuln-metrics/i;

export function applyScore(markdown: string, risk: string, vectorLink: string): string {
  const lines = markdown.replace(/\r\n?/g, "\n").split("\n");
  const at = lines.findIndex((line) => RISK_HEADING.test(line));

  if (at < 0) {
    // No Risk heading to fill in: add one rather than silently dropping the score.
    return [markdown.replace(/\s*$/, ""), "", `### Risk: ${risk}`, "", vectorLink, ""].join("\n");
  }

  const hashes = RISK_HEADING.exec(lines[at])?.[1] ?? "###";
  const rewritten = [...lines];
  rewritten[at] = `${hashes} Risk: ${risk}`;

  // Replace a vector already sitting under the heading rather than stacking another.
  const next = rewritten.findIndex((line, index) => index > at && line.trim() !== "");
  if (next > 0 && VECTOR_LINE.test(rewritten[next])) {
    rewritten[next] = vectorLink;
  } else {
    rewritten.splice(at + 1, 0, "", vectorLink);
  }

  return rewritten.join("\n");
}
