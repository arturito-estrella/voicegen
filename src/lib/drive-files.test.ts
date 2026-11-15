import assert from "node:assert/strict";
import test from "node:test";
import { deckQuery, readDriveEntries, readFileBytes } from "./drive-files.ts";

test("reads drive files from common payload shapes", () => {
  const entries = readDriveEntries({
    files: [
      { id: "1", name: "Q3.pptx", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
      { fileId: "1", title: "duplicate" },
      { id: "2", title: "Talk", mime_type: "application/vnd.google-apps.presentation" },
    ],
  });
  assert.equal(entries.length, 2);
  assert.equal(entries[1]?.name, "Talk");
});

test("reads a single created folder and base64 bytes", () => {
  const entries = readDriveEntries({ id: "folder-1", name: "Spanish", mimeType: "application/vnd.google-apps.folder" });
  assert.equal(entries[0]?.id, "folder-1");
  const bytes = readFileBytes({
    content_base64: Buffer.from("this is a long enough slide payload for a test file, with extra words so the base64 clears the length check.").toString("base64"),
  });
  assert.ok(bytes && bytes.length > 32);
});

test("deck query stays inside Drive search syntax", () => {
  const query = deckQuery("Q3 'board'");
  assert.match(query, /trashed = false/);
  assert.match(query, /name contains 'Q3 board'/);
  assert.equal(query.includes("''"), false);
});
