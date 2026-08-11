import test from "node:test";
import assert from "node:assert/strict";
import { Token, detectBody, highlightHttp } from "../src/word/http";
import { buildHttpBlock } from "../src/word/httpBlock";

const REQUEST = [
  "POST /api/v1/login?next=/admin HTTP/1.1",
  "Host: app.example.com",
  "User-Agent: Mozilla/5.0",
  "Content-Type: application/json",
  "",
  '{"username":"admin","password":"hunter2","remember":true,"attempts":3}',
].join("\r\n");

const RESPONSE = [
  "HTTP/1.1 401 Unauthorized",
  "Server: nginx",
  "Content-Type: application/json",
  "",
  '{"error":"invalid credentials"}',
].join("\r\n");

/** All tokens of a message, flattened, for asserting on kinds. */
const flatten = (raw: string): Token[] => highlightHttp(raw).lines.flat();
const kindOf = (raw: string, text: string) =>
  flatten(raw).find((token) => token.text === text)?.kind;

/** The text of a message, reassembled — it must equal what went in. */
const rebuild = (raw: string) =>
  highlightHttp(raw)
    .lines.map((line) => line.map((token) => token.text).join(""))
    .join("\n");

test("highlighting never alters the message", () => {
  // The bytes are evidence: no reformatting, no re-indenting, no re-encoding.
  assert.equal(rebuild(REQUEST), REQUEST.replace(/\r\n/g, "\n"));
  assert.equal(rebuild(RESPONSE), RESPONSE.replace(/\r\n/g, "\n"));

  const awkward = "GET /a?b=1&c=%20%22 HTTP/1.1\r\nHost: x\r\n\r\n<not json & not xml>";
  assert.equal(rebuild(awkward), awkward.replace(/\r\n/g, "\n"));
});

test("a request line is split into method, target and version", () => {
  assert.equal(highlightHttp(REQUEST).kind, "request");
  assert.equal(kindOf(REQUEST, "POST"), "method");
  assert.equal(kindOf(REQUEST, "/api/v1/login?next=/admin"), "target");
  assert.equal(kindOf(REQUEST, "HTTP/1.1"), "version");
});

test("status codes are coloured by class", () => {
  const codeKind = (line: string) => highlightHttp(line).lines[0][2].kind;

  assert.equal(codeKind("HTTP/1.1 200 OK"), "status2xx");
  assert.equal(codeKind("HTTP/1.1 302 Found"), "status3xx");
  assert.equal(codeKind("HTTP/1.1 401 Unauthorized"), "status4xx");
  assert.equal(codeKind("HTTP/1.1 500 Internal Server Error"), "status5xx");
});

test("header names and values are told apart, and values keep their colons", () => {
  const tokens = highlightHttp(RESPONSE).lines[1];

  assert.deepEqual(
    tokens.map((token) => [token.kind, token.text]),
    [
      ["headerName", "Server"],
      ["punctuation", ":"],
      ["text", " "],
      ["headerValue", "nginx"],
    ]
  );

  const dated = highlightHttp("HTTP/1.1 200 OK\r\nDate: Mon, 11 Aug 2026 09:15:00 GMT");
  const value = dated.lines[1][dated.lines[1].length - 1];
  assert.equal(value.text, "Mon, 11 Aug 2026 09:15:00 GMT");
});

test("JSON keys, strings, numbers and literals are distinguished", () => {
  assert.equal(kindOf(REQUEST, '"username"'), "key");
  assert.equal(kindOf(REQUEST, '"admin"'), "string");
  assert.equal(kindOf(REQUEST, "3"), "number");
  assert.equal(kindOf(REQUEST, "true"), "literal");
  assert.equal(highlightHttp(REQUEST).body, "json");
});

test("body syntax detection covers the shapes Burp produces", () => {
  assert.equal(detectBody('{"a":1}'), "json");
  assert.equal(detectBody("[1,2,3]"), "json");
  assert.equal(detectBody("<html><body>hi</body></html>"), "xml");
  assert.equal(detectBody('<?xml version="1.0"?><a/>'), "xml");
  assert.equal(detectBody("user=admin&pass=hunter2"), "form");
  assert.equal(detectBody("just some text"), "none");
  assert.equal(detectBody(""), "none");
});

test("a truncated JSON body is still highlighted rather than rejected", () => {
  // Reports quote partial responses constantly; failing to parse must not lose colour.
  const truncated =
    'HTTP/1.1 200 OK\r\nContent-Type: application/json\r\n\r\n{"items":[{"id":1,"na';
  const message = highlightHttp(truncated);

  assert.equal(message.body, "json");
  assert.equal(rebuild(truncated), truncated.replace(/\r\n/g, "\n"));
});

test("XML tags, attributes and text are separated", () => {
  const xml = 'HTTP/1.1 200 OK\r\nContent-Type: text/xml\r\n\r\n<user id="7">admin</user>';

  assert.equal(kindOf(xml, "admin"), "text");
  assert.equal(kindOf(xml, "id"), "attrName");
  assert.equal(kindOf(xml, '"7"'), "attrValue");
  assert.equal(rebuild(xml), xml.replace(/\r\n/g, "\n"));
});

test("form-encoded bodies separate names from values", () => {
  const form = "POST /login HTTP/1.1\r\nHost: x\r\n\r\nuser=admin&pass=hunter2";

  assert.equal(highlightHttp(form).body, "form");
  assert.equal(kindOf(form, "user"), "key");
  assert.equal(kindOf(form, "admin"), "string");
});

test("text that is not an HTTP message is kept verbatim as plain code", () => {
  const noise = "not a request\nat all";
  const message = highlightHttp(noise);

  assert.equal(message.kind, "unknown");
  assert.ok(message.lines.flat().every((token) => token.kind === "text"));
  assert.equal(rebuild(noise), noise);
});

test("the blank line between headers and body is preserved", () => {
  const message = highlightHttp(REQUEST);
  const blank = message.lines.findIndex((line) => line.length === 0);

  assert.equal(blank, 4, "after three headers");
  assert.ok(message.lines[blank + 1].length > 0, "the body follows it");
});

test("the block is a package of Codeblock paragraphs, escaped", () => {
  const xml = buildHttpBlock(highlightHttp("GET /?q=<b>&x=1 HTTP/1.1\r\nHost: x"));

  assert.ok(xml.startsWith("<pkg:package"));
  assert.ok(xml.includes('<w:pStyle w:val="Codeblock"/>'));
  assert.ok(xml.includes("&lt;b&gt;"), "markup in the message is escaped");
  assert.ok(xml.includes('w:ascii="Courier New"'), "monospace even without the style");
  assert.ok(xml.includes('xml:space="preserve"'), "spacing in the message is kept");
});

test("every token kind has a colour", () => {
  // A missing entry would render that token with no run properties at all.
  const kinds = new Set(
    [
      ...flatten(REQUEST),
      ...flatten(RESPONSE),
      ...flatten('HTTP/1.1 200 OK\r\n\r\n<a b="c">d</a>'),
      ...flatten("POST / HTTP/1.1\r\n\r\nu=1&v=2"),
    ].map((token) => token.kind)
  );

  const xml = buildHttpBlock(highlightHttp(REQUEST));
  assert.ok(kinds.size >= 8, "the fixtures exercise a decent spread of kinds");
  assert.ok(!xml.includes('<w:color w:val="undefined"/>'));
});
