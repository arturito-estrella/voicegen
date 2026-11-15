import assert from "node:assert/strict";
import test from "node:test";
import { utils, write } from "cfb";
import { buildSamplePptx } from "./deck-build.ts";
import { parsePptRecords, readPpt } from "./deck-ppt.ts";
import { readPptx } from "./deck-pptx.ts";
import { googleExportUrl } from "./google-slides.ts";

test("google slides links become an export url", () => {
  assert.equal(
    googleExportUrl("https://docs.google.com/presentation/d/abc_123/edit?usp=sharing"),
    "https://docs.google.com/presentation/d/abc_123/export/pptx",
  );
  assert.equal(googleExportUrl("https://example.com/presentation/d/abc"), null);
  assert.equal(googleExportUrl("not a url"), null);
});

test("pptx text round-trips through a Spanish rewrite", async () => {
  const sample = await buildSamplePptx();
  const deck = await readPptx(sample);
  assert.equal(deck.slides.length, 3);
  const title = deck.slides[0]?.blocks.find((block) => block.text.includes("Harbor"));
  assert.ok(title, "sample title was extracted");
  const notes = deck.slides[1]?.blocks.find((block) => block.kind === "notes");
  assert.ok(notes?.text.includes("soup"), "speaker notes were extracted");

  const spanish = new Map<string, string>();
  for (const slide of deck.slides) {
    for (const block of slide.blocks) {
      spanish.set(block.id, `ES:${block.text}`);
    }
  }
  const rewritten = await deck.apply(spanish);
  const again = await readPptx(rewritten);
  const rewrittenTitle = again.slides[0]?.blocks.find((block) => block.text.startsWith("ES:"));
  assert.ok(rewrittenTitle, "translated title was written back");
  const rewrittenNotes = again.slides[1]?.blocks.find((block) => block.kind === "notes" && block.text.startsWith("ES:"));
  assert.ok(rewrittenNotes, "translated notes were written back");
});

function atom(type: number, payload: Uint8Array): Uint8Array {
  const out = new Uint8Array(8 + payload.length);
  const view = new DataView(out.buffer);
  view.setUint16(0, 0, true);
  view.setUint16(2, type, true);
  view.setUint32(4, payload.length, true);
  out.set(payload, 8);
  return out;
}

function container(type: number, children: Uint8Array[]): Uint8Array {
  const length = children.reduce((sum, child) => sum + child.length, 0);
  const out = new Uint8Array(8 + length);
  const view = new DataView(out.buffer);
  view.setUint16(0, 0x000f, true);
  view.setUint16(2, type, true);
  view.setUint32(4, length, true);
  let offset = 8;
  for (const child of children) {
    out.set(child, offset);
    offset += child.length;
  }
  return out;
}

function u32(value: number): Uint8Array {
  const bytes = new Uint8Array(4);
  new DataView(bytes.buffer).setUint32(0, value, true);
  return bytes;
}

function utf16(value: string): Uint8Array {
  const bytes = new Uint8Array(value.length * 2);
  const view = new DataView(bytes.buffer);
  for (let index = 0; index < value.length; index += 1) view.setUint16(index * 2, value.charCodeAt(index), true);
  return bytes;
}

test("legacy ppt records keep slide boundaries", () => {
  const slideList = container(0x0ff0, [
    atom(0x03f3, new Uint8Array(20)),
    atom(0x0f9f, u32(0)),
    atom(0x0fa0, utf16("Harbor Market")),
    atom(0x0f9f, u32(1)),
    atom(0x0fa0, utf16("Opens Friday\rBring the crates")),
    atom(0x03f3, new Uint8Array(20)),
    atom(0x0f9f, u32(0)),
    atom(0x0fa0, utf16("Second slide")),
  ]);
  const slides = parsePptRecords(slideList);
  assert.equal(slides.length, 2);
  assert.equal(slides[0]?.blocks[0]?.text, "Harbor Market");
  assert.equal(slides[0]?.blocks[0]?.kind, "title");
  assert.deepEqual(
    slides[0]?.blocks.slice(1).map((block) => block.text),
    ["Opens Friday", "Bring the crates"],
  );
  assert.equal(slides[1]?.blocks[0]?.text, "Second slide");
});

test("legacy ppt inside a compound file is readable", () => {
  const slideList = container(0x0ff0, [
    atom(0x03f3, new Uint8Array(20)),
    atom(0x0f9f, u32(0)),
    atom(0x0fa0, utf16("Staff briefing")),
  ]);
  const book = utils.cfb_new();
  utils.cfb_add(book, "/PowerPoint Document", slideList);
  const bytes = Buffer.from(write(book, { type: "buffer" }));
  const slides = readPpt(bytes);
  assert.equal(slides[0]?.blocks[0]?.text, "Staff briefing");
});
