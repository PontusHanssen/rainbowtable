import test from "node:test";
import assert from "node:assert/strict";
import { ToDialog, ToPane, decode, encode, nextRequestId } from "../src/shared/protocol";

test("a message survives the round trip unchanged", () => {
  const sent: ToPane = { kind: "insert", requestId: "r1", markdown: "# Title\n\n- item" };
  assert.deepEqual(decode<ToPane>(encode(sent)), sent);

  const reply: ToDialog = {
    kind: "inserted",
    requestId: "r1",
    bookmark: "_ptmd",
    paragraphs: 4,
    plainStyles: false,
  };
  assert.deepEqual(decode<ToDialog>(encode(reply)), reply);
});

test("markdown with quotes and newlines crosses intact", () => {
  // The channel is a string, so anything that breaks JSON breaks the feature.
  const markdown = '## "Quoted"\n\n```\n{"a": 1}\n```\n\n- <https://x.test>';
  const decoded = decode<ToPane>(encode({ kind: "insert", requestId: "r2", markdown }));

  assert.equal(decoded?.kind === "insert" && decoded.markdown, markdown);
});

test("anything that is not one of our messages is refused, not guessed at", () => {
  // The development probes put other JSON on this channel, and a page can send anything.
  assert.equal(decode("plain markdown, not JSON"), undefined);
  assert.equal(decode(""), undefined);
  assert.equal(decode("{ broken"), undefined);
  assert.equal(decode('{"kind":"probe","bytes":1024}'), undefined, "no requestId");
  assert.equal(decode('{"requestId":"r1"}'), undefined, "no kind");
});

test("request ids are unique, so replies cannot be matched to the wrong request", () => {
  const ids = new Set(Array.from({ length: 500 }, nextRequestId));
  assert.equal(ids.size, 500);
});
