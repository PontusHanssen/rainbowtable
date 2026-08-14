import test from "node:test";
import assert from "node:assert/strict";
import { safeLink } from "../src/word/writePlan";

/**
 * A plan reaches the pane from the dialog as well as from its own panels, and the pane is
 * the only side that can write to the document, so the address is checked where it is
 * written rather than where it was built.
 */
test("safeLink allows the schemes a finding legitimately links to", () => {
  assert.equal(safeLink("https://example.com/a?b=c#d"), "https://example.com/a?b=c#d");
  assert.equal(safeLink("http://10.0.0.1:8080/"), "http://10.0.0.1:8080/");
  assert.equal(safeLink("mailto:security@example.com"), "mailto:security@example.com");
  assert.equal(safeLink("HTTPS://Example.com"), "HTTPS://Example.com", "scheme is case-blind");
  assert.equal(safeLink("  https://example.com  "), "https://example.com", "trimmed");
});

test("safeLink refuses anything that could reach the reader's machine or credentials", () => {
  [
    "file://attacker.example.com/share/x",
    "\\\\attacker.example.com\\share\\x",
    "javascript:alert(1)",
    "vbscript:msgbox",
    "data:text/html,<script>",
    "ms-word:ofe|u|https://example.com",
    "",
    " ",
  ].forEach((link) => assert.equal(safeLink(link), undefined, link));
});
