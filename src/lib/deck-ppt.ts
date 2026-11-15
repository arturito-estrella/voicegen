import { find, read, type CFB$Blob } from "cfb";
import type { ExtractedSlide } from "./deck-pptx.ts";

const RT_SLIDE_LIST = 0x0ff0;
const RT_PERSIST = 0x03f3;
const RT_TEXT_HEADER = 0x0f9f;
const RT_TEXT_CHARS = 0x0fa0;
const RT_TEXT_BYTES = 0x0fa8;

type TextKind = "title" | "body" | "notes";

type RawSlide = { texts: { kind: TextKind; text: string }[] };

function asBytes(content: CFB$Blob | undefined): Uint8Array {
  if (!content) return new Uint8Array();
  if (content instanceof Uint8Array) return content;
  return Uint8Array.from(content);
}

function textKind(code: number): TextKind {
  if (code === 0 || code === 6) return "title";
  if (code === 2) return "notes";
  return "body";
}

function cleanPieces(raw: string, kind: TextKind): { kind: TextKind; text: string }[] {
  const pieces = raw
    .replace(/\u000b/g, "\n")
    .split(/\r/)
    .map((part) => part.replace(/[\u0000-\u0008\u000e-\u001f]/g, "").trim())
    .filter((part) => /\p{L}|\p{N}/u.test(part));
  return pieces.map((text, index) => ({
    kind: kind === "title" && index > 0 ? "body" : kind,
    text,
  }));
}

function decodeChars(bytes: Uint8Array): string {
  if (bytes.byteLength < 2) return "";
  const even = bytes.byteLength - (bytes.byteLength % 2);
  return new TextDecoder("utf-16le").decode(bytes.subarray(0, even));
}

function decodeLatin(bytes: Uint8Array): string {
  let text = "";
  for (const byte of bytes) text += String.fromCharCode(byte);
  return text;
}

export function parsePptRecords(data: Uint8Array): ExtractedSlide[] {
  const lists: RawSlide[][] = [];
  let listIndex = -1;
  let pending: TextKind = "body";

  const ensureSlide = () => {
    if (listIndex < 0) {
      lists.push([]);
      listIndex = 0;
    }
    const list = lists[listIndex]!;
    if (list.length === 0) list.push({ texts: [] });
    return list[list.length - 1]!;
  };

  const walk = (start: number, end: number) => {
    let offset = start;
    while (offset + 8 <= end) {
      const view = new DataView(data.buffer, data.byteOffset + offset, 8);
      const version = view.getUint16(0, true) & 0xf;
      const type = view.getUint16(2, true);
      const length = view.getUint32(4, true);
      const body = offset + 8;
      const next = body + length;
      if (length < 0 || next > end || next <= offset) break;

      if (version === 0xf) {
        const openedList = type === RT_SLIDE_LIST;
        const previous = listIndex;
        if (openedList) {
          lists.push([]);
          listIndex = lists.length - 1;
        }
        walk(body, next);
        if (openedList) listIndex = previous;
      } else if (type === RT_PERSIST && listIndex >= 0) {
        lists[listIndex]!.push({ texts: [] });
      } else if (type === RT_TEXT_HEADER && length >= 4) {
        const header = new DataView(data.buffer, data.byteOffset + body, 4);
        pending = textKind(header.getUint32(0, true));
      } else if (type === RT_TEXT_CHARS || type === RT_TEXT_BYTES) {
        const payload = data.subarray(body, next);
        const raw = type === RT_TEXT_CHARS ? decodeChars(payload) : decodeLatin(payload);
        const pieces = cleanPieces(raw, pending);
        if (pieces.length > 0) ensureSlide().texts.push(...pieces);
      }

      offset = next;
    }
  };

  walk(0, data.length);

  const scored = lists
    .map((slides, index) => ({
      index,
      slides,
      textCount: slides.reduce((sum, slide) => sum + slide.texts.length, 0),
      noteCount: slides.reduce(
        (sum, slide) => sum + slide.texts.filter((item) => item.kind === "notes").length,
        0,
      ),
    }))
    .filter((list) => list.textCount > 0);

  if (scored.length === 0) return [];

  scored.sort((a, b) => b.textCount - b.noteCount - (a.textCount - a.noteCount));
  const main = scored[0]!;
  const notesList = scored.find(
    (list) => list.index !== main.index && list.slides.length === main.slides.length,
  );

  return main.slides.map((slide, index) => {
    const blocks = slide.texts.map((item, blockIndex) => ({
      id: `p${index + 1}b${blockIndex}`,
      kind: item.kind,
      text: item.text,
    }));
    const noteSlide = notesList?.slides[index];
    if (noteSlide) {
      const seen = new Set(blocks.map((block) => block.text));
      for (const item of noteSlide.texts) {
        if (seen.has(item.text)) continue;
        blocks.push({
          id: `p${index + 1}n${blocks.length}`,
          kind: "notes",
          text: item.text,
        });
      }
    }
    return { index: index + 1, blocks };
  });
}

export function readPpt(bytes: Buffer): ExtractedSlide[] {
  const container = read(bytes, { type: "buffer" });
  const encrypted = container.FileIndex.some((entry) => /encrypt/i.test(entry.name));
  if (encrypted) {
    throw new Error("This file is password-protected. Remove the password and try again.");
  }

  const entry =
    find(container, "/PowerPoint Document") ??
    find(container, "PowerPoint Document") ??
    container.FileIndex.find((item) => /PowerPoint Document$/i.test(item.name)) ??
    null;
  if (!entry?.content?.length) {
    throw new Error("That .ppt file did not contain readable slides.");
  }

  const slides = parsePptRecords(asBytes(entry.content)).filter((slide) =>
    slide.blocks.some((block) => block.text.trim()),
  );
  if (slides.length === 0) {
    throw new Error("No readable text was found in that .ppt file.");
  }
  if (slides.length > 80) {
    throw new Error("That deck has more than 80 slides. Split it and translate each part.");
  }
  return slides.map((slide, index) => ({ ...slide, index: index + 1 }));
}
