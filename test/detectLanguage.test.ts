import test from "node:test";
import assert from "node:assert/strict";
import { detectLanguage } from "../src/shared/detectLanguage";
import { loadableLanguages } from "../src/shared/highlightCode";

test("the languages a pentest report quotes are recognised", () => {
  const samples: [string, string][] = [
    ["SELECT * FROM users WHERE id = 1", "sql"],
    ["  delete from sessions where expired = 1", "sql"],
    ['{"user": "admin", "role": 1}', "json"],
    ["<?php echo $_GET['q']; ?>", "php"],
    ["<!DOCTYPE html>\n<html><body>hi</body></html>", "html"],
    ['<?xml version="1.0"?><root/>', "xml"],
    ["def handler(request):\n    return request.GET['q']", "python"],
    ["import os\nos.system(cmd)", "python"],
    ["package main\n\nfunc main() {}", "go"],
    ["using System;\nConsole.WriteLine(x);", "csharp"],
    ["public class Main { System.out.println(x); }", "java"],
    ["const token = req.query.q;\nconsole.log(token);", "javascript"],
    ["#!/bin/bash\nrm -rf /tmp/x", "shell"],
    ["$ curl -s https://x.test/api\n$ echo done", "shell"],
    ["body { color: #fff; margin: 0; }", "css"],
    ["---\nname: deploy\nreplicas: 3", "yaml"],
  ];

  for (const [code, language] of samples) {
    assert.equal(detectLanguage(code), language, code.split("\n")[0]);
  }
});

test("prose and unrecognised text stay plain rather than being guessed at", () => {
  // A wrong guess colours code as something it is not, which is worse than no colour.
  assert.equal(detectLanguage("The parameter is not validated before use."), undefined);
  assert.equal(detectLanguage(""), undefined);
  assert.equal(detectLanguage("   \n  "), undefined);
  assert.equal(detectLanguage("lorem ipsum dolor sit amet"), undefined);
});

test("a truncated JSON body is still JSON", () => {
  // Reports quote partial responses constantly.
  assert.equal(detectLanguage('{"items":[{"id":1,"na'), "json");
});

test("something that merely starts with a brace is not JSON", () => {
  assert.equal(detectLanguage("{ this is not json at all }"), undefined);
});

test("every language detected has a grammar to colour it with", () => {
  // Detecting a language nothing can highlight would promise colour and deliver none.
  const samples = [
    "SELECT 1",
    '{"a":1}',
    "<?php $x = 1;",
    "<!DOCTYPE html>",
    "def f():\n    pass",
    "package main\n\nfunc f() {}",
    "using System;",
    "public class X {}",
    "const a = 1;",
    "#!/bin/bash",
    "body { color: red; }",
    "---\nkey: value",
  ];

  const detected = new Set(samples.map(detectLanguage).filter(Boolean) as string[]);
  const unsupported = [...detected].filter((language) => !loadableLanguages().includes(language));

  assert.deepEqual(unsupported, [], "detected but not highlightable");
});
