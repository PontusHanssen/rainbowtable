import test from "node:test";
import assert from "node:assert/strict";
import { CODE_SPAN_PATTERN, stripDelimiters } from "../src/word/inlineCode";

test("stripDelimiters removes the backticks and nothing else", () => {
  assert.equal(stripDelimiters("`code`"), "code");
  assert.equal(stripDelimiters("`a b c`"), "a b c");
  assert.equal(stripDelimiters("`/etc/passwd`"), "/etc/passwd");
});

test("stripDelimiters leaves text that is not a span alone", () => {
  // Word's matcher should never hand these over, but the function must not corrupt them.
  assert.equal(stripDelimiters("plain"), "plain");
  assert.equal(stripDelimiters("`unclosed"), "`unclosed");
  assert.equal(stripDelimiters("unopened`"), "unopened`");
});

test("the search pattern excludes backticks so it cannot span two spans", () => {
  // A greedy `*` would match from the first backtick to the last, swallowing the text
  // between two separate spans. This is Word wildcard syntax, not a regular expression.
  assert.equal(CODE_SPAN_PATTERN, "`[!`]@`");
  assert.ok(CODE_SPAN_PATTERN.includes("[!`]"), "non-backtick characters only");
  assert.ok(!CODE_SPAN_PATTERN.includes("*"), "no greedy wildcard");
});
