import test from "node:test";
import assert from "node:assert/strict";
import { DraftStorage, createDraftStore } from "../src/dialog/draft";

const SKELETON = "## Title\n\n### Risk:\n";

function memory(): DraftStorage & { entries: Map<string, string> } {
  const entries = new Map<string, string>();
  return {
    entries,
    getItem: (key) => entries.get(key) ?? null,
    setItem: (key, value) => void entries.set(key, value),
    removeItem: (key) => void entries.delete(key),
  };
}

test("a draft survives to be loaded again", () => {
  const storage = memory();
  createDraftStore(storage, SKELETON).save("## Real finding\n\nprose");

  assert.equal(createDraftStore(storage, SKELETON).load(), "## Real finding\n\nprose");
});

test("an untouched skeleton is not a draft", () => {
  // Restoring one would be noise, and would hide that nothing was written.
  const storage = memory();
  const drafts = createDraftStore(storage, SKELETON);

  drafts.save(SKELETON);
  assert.equal(drafts.load(), undefined);

  drafts.save("   \n  ");
  assert.equal(drafts.load(), undefined);
});

test("saving the skeleton again clears a previous draft", () => {
  const storage = memory();
  const drafts = createDraftStore(storage, SKELETON);

  drafts.save("## Something");
  drafts.save(SKELETON);

  assert.equal(drafts.load(), undefined, "stale drafts must not linger");
});

test("inserting clears the draft", () => {
  const storage = memory();
  const drafts = createDraftStore(storage, SKELETON);

  drafts.save("## Written and inserted");
  drafts.clear();

  assert.equal(drafts.load(), undefined);
});

test("storage that throws does not break the editor", () => {
  // Storage can be unavailable or full; losing the backup must not stop someone writing.
  const hostile: DraftStorage = {
    getItem: () => {
      throw new Error("denied");
    },
    setItem: () => {
      throw new Error("full");
    },
    removeItem: () => {
      throw new Error("denied");
    },
  };
  const drafts = createDraftStore(hostile, SKELETON);

  assert.doesNotThrow(() => drafts.save("## Finding"));
  assert.equal(drafts.load(), undefined);
  assert.doesNotThrow(() => drafts.clear());
});

test("no storage at all is tolerated", () => {
  const drafts = createDraftStore(undefined, SKELETON);

  assert.doesNotThrow(() => drafts.save("## Finding"));
  assert.equal(drafts.load(), undefined);
});
