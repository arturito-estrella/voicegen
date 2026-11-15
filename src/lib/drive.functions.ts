import { createServerFn } from "@tanstack/react-start";
import type { DeckResult } from "./voicegen-types.ts";
import type { DriveEntry } from "./drive-files.ts";

type DriveErrorBody = {
  ok: false;
  kind: string;
  message: string;
  loginUrl?: string;
  pending: boolean;
};

function cleanId(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const id = value.trim();
  if (!id || id.length > 200) return null;
  return id;
}

export const listDriveDecks = createServerFn({ method: "POST" })
  .validator((input: unknown) => {
    const query =
      input && typeof input === "object" && typeof (input as { query?: unknown }).query === "string"
        ? (input as { query: string }).query.slice(0, 80)
        : "";
    return { query };
  })
  .handler(async ({ data }) => {
    const { listDecks } = await import("./drive.server.ts");
    return listDecks(data.query);
  });

export const listDriveFolders = createServerFn({ method: "POST" })
  .validator((input: unknown) => ({
    parentId: cleanId(input && typeof input === "object" ? (input as { parentId?: unknown }).parentId : null),
  }))
  .handler(async ({ data }) => {
    const { listFolders } = await import("./drive.server.ts");
    return listFolders(data.parentId);
  });

export const createDriveFolder = createServerFn({ method: "POST" })
  .validator((input: unknown) => {
    if (!input || typeof input !== "object") throw new Error("Name the folder.");
    const name = (input as { name?: unknown }).name;
    if (typeof name !== "string" || name.trim().length < 1) throw new Error("Name the folder.");
    return {
      name: name.trim().slice(0, 80),
      parentId: cleanId((input as { parentId?: unknown }).parentId),
    };
  })
  .handler(async ({ data }) => {
    const { createDriveFolder: create } = await import("./drive.server.ts");
    return create(data.name, data.parentId);
  });

export const translateDriveFile = createServerFn({ method: "POST" })
  .validator((input: unknown) => {
    const fileId = cleanId(input && typeof input === "object" ? (input as { fileId?: unknown }).fileId : null);
    if (!fileId) throw new Error("Choose a file from Google Drive.");
    return { fileId };
  })
  .handler(async ({ data }): Promise<DeckResult | DriveErrorBody> => {
    const { readDeckFile } = await import("./drive.server.ts");
    const file = await readDeckFile(data.fileId);
    if (!file.ok) return file;
    if (file.bytes.length > 12 * 1024 * 1024) {
      return {
        ok: false,
        kind: "error",
        pending: false,
        message: "That Drive file is over 12 MB. Download a smaller copy and drop it here.",
      };
    }
    const { runTranslation } = await import("./translate-deck.ts");
    return runTranslation({
      kind: "file",
      name: file.name,
      dataBase64: file.bytes.toString("base64"),
    });
  });

export const saveDeckToDrive = createServerFn({ method: "POST" })
  .validator((input: unknown) => {
    if (!input || typeof input !== "object") throw new Error("Nothing to save.");
    const name = (input as { name?: unknown }).name;
    const dataBase64 = (input as { dataBase64?: unknown }).dataBase64;
    if (typeof name !== "string" || typeof dataBase64 !== "string" || dataBase64.length < 32) {
      throw new Error("Nothing to save.");
    }
    if (dataBase64.length > 12_000_000) {
      throw new Error("That deck is too large to send back to Drive. Download it instead.");
    }
    return {
      name: name.slice(0, 180),
      dataBase64,
      folderId: cleanId((input as { folderId?: unknown }).folderId),
    };
  })
  .handler(async ({ data }) => {
    const { saveDeckFile } = await import("./drive.server.ts");
    return saveDeckFile({
      folderId: data.folderId,
      name: data.name,
      bytes: Buffer.from(data.dataBase64, "base64"),
    });
  });

export type { DriveEntry };
