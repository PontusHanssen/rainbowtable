import test from "node:test";
import assert from "node:assert/strict";
import { ToDialog, ToPane, decode, encode, nextRequestId } from "../src/shared/protocol";
import { planFromBlocks } from "../src/word/documentPlan";
import { parseMarkdown } from "../src/word/markdown";

test("a message survives the round trip unchanged", () => {
  const plans = planFromBlocks(parseMarkdown("# Title\n\n- item"));
  const sent: ToPane = { kind: "insert", requestId: "r1", plans };
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

test("a plan with quotes, colours and links crosses intact", () => {
  // The channel is a string, so anything that breaks JSON breaks the feature.
  const plans = planFromBlocks(
    parseMarkdown('## "Quoted"\n\n```\n{"a": 1}\n```\n\n- <https://x.test>')
  );
  const decoded = decode<ToPane>(encode({ kind: "insert", requestId: "r2", plans }));

  assert.deepEqual(decoded?.kind === "insert" && decoded.plans, plans);
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
