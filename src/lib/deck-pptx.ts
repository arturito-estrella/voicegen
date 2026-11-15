import JSZip from "jszip";
import { applyHits, extractParagraphs, maskFields, type ParaHit } from "./xml-text.ts";
import type { BlockKind } from "./voicegen-types.ts";

export type ExtractedBlock = {
  id: string;
  kind: BlockKind;
  text: string;
};

export type ExtractedSlide = {
  index: number;
  blocks: ExtractedBlock[];
};

type XmlJob = {
  path: string;
  xml: string;
  hits: ParaHit[];
  restore: (value: string) => string;
};

export type PptxDeck = {
  slides: ExtractedSlide[];
  apply: (spanishById: Map<string, string>) => Promise<Buffer>;
};

function resolveRel(fromFile: string, target: string): string {
  const dir = fromFile.split("/").slice(0, -1);
  for (const part of target.split("/")) {
    if (part === "..") dir.pop();
    else if (part !== "." && part !== "") dir.push(part);
  }
  return dir.join("/");
}

function slideOrder(presentationXml: string, relsXml: string): string[] {
  const targets = new Map<string, string>();
  for (const tag of relsXml.match(/<Relationship\b[^>]*>/g) ?? []) {
    const id = /Id="([^"]+)"/.exec(tag)?.[1];
    const target = /Target="([^"]+)"/.exec(tag)?.[1];
    const type = /Type="([^"]+)"/.exec(tag)?.[1] ?? "";
    if (!id || !target || !type.endsWith("/slide")) continue;
    const clean = target.replace(/^\//, "");
    targets.set(id, clean.startsWith("ppt/") ? clean : `ppt/${clean}`);
  }
  const ordered: string[] = [];
  for (const tag of presentationXml.match(/<p:sldId\b[^>]*\/>/g) ?? []) {
    const id = /r:id="([^"]+)"/.exec(tag)?.[1];
    const path = id ? targets.get(id) : undefined;
    if (path) ordered.push(path);
  }
  return ordered;
}

function fallbackSlides(zip: JSZip): string[] {
  return Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => {
      const na = Number(/slide(\d+)\.xml$/.exec(a)?.[1] ?? 0);
      const nb = Number(/slide(\d+)\.xml$/.exec(b)?.[1] ?? 0);
      return na - nb;
    });
}

async function notesFor(zip: JSZip, slidePath: string): Promise<string | null> {
  const file = slidePath.split("/").pop();
  if (!file) return null;
  const relsPath = `ppt/slides/_rels/${file}.rels`;
  const rels = await zip.file(relsPath)?.async("string");
  if (!rels) return null;
  for (const tag of rels.match(/<Relationship\b[^>]*>/g) ?? []) {
    const type = /Type="([^"]+)"/.exec(tag)?.[1] ?? "";
    const target = /Target="([^"]+)"/.exec(tag)?.[1];
    if (!target || !type.endsWith("/notesSlide")) continue;
    return resolveRel(slidePath, target);
  }
  return null;
}

export async function readPptx(bytes: Buffer): Promise<PptxDeck> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch {
    throw new Error("That PowerPoint file could not be opened.");
  }

  if (zip.file("EncryptedPackage") || zip.file("encryption.xml")) {
    throw new Error("This file is password-protected. Remove the password and try again.");
  }

  const presentation = await zip.file("ppt/presentation.xml")?.async("string");
  const rels = await zip.file("ppt/_rels/presentation.xml.rels")?.async("string");
  const ordered =
    presentation && rels ? slideOrder(presentation, rels) : fallbackSlides(zip);
  const paths = ordered.length > 0 ? ordered : fallbackSlides(zip);
  if (paths.length === 0) {
    throw new Error("No slides were found in that file.");
  }
  if (paths.length > 80) {
    throw new Error("That deck has more than 80 slides. Split it and translate each part.");
  }

  const jobs: XmlJob[] = [];
  const slides: ExtractedSlide[] = [];

  for (let index = 0; index < paths.length; index += 1) {
    const path = paths[index]!;
    const raw = await zip.file(path)?.async("string");
    if (!raw) continue;
    const masked = maskFields(raw);
    const hits = extractParagraphs(masked.xml, `s${index + 1}`, false);
    jobs.push({ path, xml: masked.xml, hits, restore: masked.restore });

    const blocks: ExtractedBlock[] = hits.map((hit) => ({
      id: hit.id,
      kind: hit.kind,
      text: hit.text,
    }));

    const notesPath = await notesFor(zip, path);
    if (notesPath) {
      const notesRaw = await zip.file(notesPath)?.async("string");
      if (notesRaw) {
        const notesMasked = maskFields(notesRaw);
        const noteHits = extractParagraphs(notesMasked.xml, `s${index + 1}n`, true);
        jobs.push({
          path: notesPath,
          xml: notesMasked.xml,
          hits: noteHits,
          restore: notesMasked.restore,
        });
        for (const hit of noteHits) {
          blocks.push({ id: hit.id, kind: "notes", text: hit.text });
        }
      }
    }

    slides.push({ index: index + 1, blocks });
  }

  if (slides.length === 0) throw new Error("No slides were found in that file.");

  return {
    slides,
    apply: async (spanishById) => {
      for (const job of jobs) {
        const next = job.restore(applyHits(job.xml, job.hits, spanishById));
        zip.file(job.path, next);
      }
      const out = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
      return Buffer.from(out);
    },
  };
}

export function isZip(bytes: Buffer): boolean {
  return bytes.length > 4 && bytes[0] === 0x50 && bytes[1] === 0x4b;
}

export function isOle(bytes: Buffer): boolean {
  return (
    bytes.length > 8 &&
    bytes[0] === 0xd0 &&
    bytes[1] === 0xcf &&
    bytes[2] === 0x11 &&
    bytes[3] === 0xe0
  );
}
