export type BlockKind = "title" | "body" | "notes";

export type ViewBlock = {
  id: string;
  kind: BlockKind;
  original: string;
  spanish: string;
};

export type ViewSlide = {
  index: number;
  blocks: ViewBlock[];
};

export type DeckResult = {
  sourceName: string;
  downloadName: string;
  preservedLayout: boolean;
  slides: ViewSlide[];
  pptxBase64: string;
  warning?: string;
};

export type TranslateRequest =
  | { kind: "sample" }
  | { kind: "file"; name: string; dataBase64: string }
  | { kind: "google"; url: string };
