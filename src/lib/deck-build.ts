import PptxGenJS from "pptxgenjs";
import type { ViewSlide } from "./voicegen-types.ts";

function asBuffer(value: unknown): Buffer {
  if (Buffer.isBuffer(value)) return value;
  if (value instanceof Uint8Array) return Buffer.from(value);
  if (value instanceof ArrayBuffer) return Buffer.from(value);
  throw new Error("Could not build the PowerPoint file.");
}

export async function buildSamplePptx(): Promise<Buffer> {
  const pptx = new PptxGenJS();
  pptx.title = "Harbor Market Opening";
  pptx.author = "VoiceGen";
  pptx.layout = "LAYOUT_WIDE";

  const first = pptx.addSlide();
  first.addText("Harbor Market Opening", {
    x: 0.7,
    y: 1.8,
    w: 11.5,
    h: 1.3,
    fontFace: "Calibri",
    fontSize: 40,
    color: "1A1C1E",
    margin: 0,
  });
  first.addText("A short briefing for the Tuesday staff meeting.", {
    x: 0.7,
    y: 3.3,
    w: 10,
    h: 0.6,
    fontFace: "Calibri",
    fontSize: 20,
    color: "3A3E42",
    margin: 0,
  });

  const second = pptx.addSlide();
  second.addText("What guests will notice", {
    x: 0.7,
    y: 0.5,
    w: 11.5,
    h: 0.8,
    fontFace: "Calibri",
    fontSize: 32,
    color: "1A1C1E",
    margin: 0,
  });
  second.addText(
    [
      {
        text: "The east awning is repaired, and the lights are warm after six.",
        options: { bullet: true, breakLine: true },
      },
      {
        text: "We added a children's menu with grilled cheese and fruit.",
        options: { bullet: true, breakLine: true },
      },
      {
        text: "Please greet regulars by name when you know it.",
        options: { bullet: true },
      },
    ],
    {
      x: 0.8,
      y: 1.7,
      w: 11.2,
      h: 4.2,
      fontFace: "Calibri",
      fontSize: 22,
      color: "2A2E32",
      paraSpaceAfter: 14,
      valign: "top",
    },
  );
  second.addNotes("Mention the soup special only if someone asks. It is nine dollars.");

  const third = pptx.addSlide();
  third.addText("Ask for this week", {
    x: 0.7,
    y: 0.5,
    w: 11.5,
    h: 0.8,
    fontFace: "Calibri",
    fontSize: 32,
    color: "1A1C1E",
    margin: 0,
  });
  third.addText(
    "Arrive ten minutes early on Friday. Tell Maya if a station is short before the doors open.",
    {
      x: 0.7,
      y: 1.8,
      w: 11.2,
      h: 2.4,
      fontFace: "Calibri",
      fontSize: 24,
      color: "2A2E32",
      margin: 0,
    },
  );

  return asBuffer(await pptx.write({ outputType: "nodebuffer" }));
}

export async function buildTranslatedPptx(slides: ViewSlide[]): Promise<Buffer> {
  const pptx = new PptxGenJS();
  pptx.title = "Spanish translation";
  pptx.author = "VoiceGen";
  pptx.layout = "LAYOUT_WIDE";
  pptx.theme = { headFontFace: "Calibri", bodyFontFace: "Calibri" };

  for (const slide of slides) {
    const page = pptx.addSlide();
    page.background = { color: "F3F1EE" };
    page.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: 13.33,
      h: 0.1,
      fill: { color: "1A1C1E" },
    });

    const titles = slide.blocks.filter((block) => block.kind === "title").map((block) => block.spanish);
    const body = slide.blocks.filter((block) => block.kind === "body").map((block) => block.spanish);
    const notes = slide.blocks.filter((block) => block.kind === "notes").map((block) => block.spanish);

    let y = 0.45;
    if (titles.length > 0) {
      page.addText(titles.join("\n"), {
        x: 0.65,
        y,
        w: 12,
        h: 1.35,
        fontFace: "Calibri",
        fontSize: 32,
        color: "1A1C1E",
        margin: 0,
        valign: "top",
      });
      y = 1.95;
    }
    if (body.length > 0) {
      page.addText(
        body.map((text, index) => ({
          text,
          options: { bullet: body.length > 1, breakLine: index < body.length - 1 },
        })),
        {
          x: 0.7,
          y,
          w: 11.9,
          h: 4.8,
          fontFace: "Calibri",
          fontSize: 20,
          color: "2A2E32",
          paraSpaceAfter: 12,
          valign: "top",
        },
      );
    }
    if (notes.length > 0) page.addNotes(notes.join("\n\n"));
  }

  return asBuffer(await pptx.write({ outputType: "nodebuffer" }));
}
