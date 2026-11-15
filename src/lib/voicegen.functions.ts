import { createServerFn } from "@tanstack/react-start";
import type { DeckResult, TranslateRequest } from "./voicegen-types.ts";

function parseTranslateRequest(input: unknown): TranslateRequest {
  if (!input || typeof input !== "object") throw new Error("Missing deck.");
  const kind = (input as { kind?: unknown }).kind;
  if (kind === "sample") return { kind: "sample" };
  if (kind === "google") {
    const url = (input as { url?: unknown }).url;
    if (typeof url !== "string" || url.trim().length < 12 || url.length > 2000) {
      throw new Error("Paste a Google Slides link.");
    }
    return { kind: "google", url: url.trim() };
  }
  if (kind === "file") {
    const name = (input as { name?: unknown }).name;
    const dataBase64 = (input as { dataBase64?: unknown }).dataBase64;
    if (typeof name !== "string" || typeof dataBase64 !== "string") {
      throw new Error("Choose a PowerPoint file.");
    }
    if (dataBase64.length > 18_000_000) {
      throw new Error("That file is over 12 MB. Try a smaller deck.");
    }
    return { kind: "file", name: name.slice(0, 180), dataBase64 };
  }
  throw new Error("Choose a PowerPoint file or a Google Slides link.");
}

export const translateDeck = createServerFn({ method: "POST" })
  .validator(parseTranslateRequest)
  .handler(async ({ data }): Promise<DeckResult> => {
    const { runTranslation } = await import("./translate-deck.ts");
    return runTranslation(data);
  });
