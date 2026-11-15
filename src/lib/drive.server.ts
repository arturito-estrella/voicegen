import { classifyCallToolError } from "@/lib/app-data/errors.ts";
import type { CallToolResult } from "@/lib/app-data/types.ts";
import {
  deckQuery,
  folderQuery,
  PPTX_MIME,
  readDriveEntries,
  readFileBytes,
  type DriveEntry,
} from "./drive-files.ts";

export type DriveFailure = {
  ok: false;
  kind: string;
  message: string;
  loginUrl?: string;
  pending: boolean;
};

export type DriveListResult = { ok: true; entries: DriveEntry[] } | DriveFailure;

const UPLOAD_TOOLS = ["google_drive_upload_file", "google_drive_create_file"];

function pack(result: CallToolResult): DriveFailure {
  const state = classifyCallToolError(result);
  return {
    ok: false,
    kind: state?.kind ?? "error",
    message: state?.message ?? "Couldn't reach Google Drive.",
    pending: result.pending === true,
    ...(result.loginUrl ? { loginUrl: result.loginUrl } : {}),
  };
}

async function driveCall(toolName: string, args: Record<string, unknown>): Promise<CallToolResult> {
  const { callTool } = await import("@/lib/app-data/client.server.ts");
  const { ConnectorType } = await import("@/lib/app-data/types.ts");
  return callTool(toolName, args, { connectorType: ConnectorType.GoogleDrive });
}

export async function listDecks(query: string): Promise<DriveListResult> {
  const { GoogleDriveTools } = await import("@/lib/app-data/types.ts");
  const result = await driveCall(GoogleDriveTools.search, {
    q: deckQuery(query),
    page_size: 30,
  });
  if (!result.ok) return pack(result);
  const entries = readDriveEntries(result.data).filter(
    (entry) => !entry.mimeType || entry.mimeType.includes("presentation") || entry.mimeType.includes("powerpoint"),
  );
  return { ok: true, entries };
}

export async function listFolders(parentId: string | null): Promise<DriveListResult> {
  const { GoogleDriveTools } = await import("@/lib/app-data/types.ts");
  const folderId = parentId && parentId !== "root" ? parentId : "root";
  const listed = await driveCall(GoogleDriveTools.listFolder, { folder_id: folderId, folderId });
  if (listed.ok) {
    const entries = readDriveEntries(listed.data).filter(
      (entry) => !entry.mimeType || entry.mimeType === "application/vnd.google-apps.folder",
    );
    return { ok: true, entries };
  }
  if (listed.pending || listed.loginRequired) return pack(listed);
  const searched = await driveCall(GoogleDriveTools.search, { q: folderQuery(parentId), page_size: 40 });
  if (!searched.ok) return pack(searched);
  return {
    ok: true,
    entries: readDriveEntries(searched.data).filter(
      (entry) => entry.mimeType === "application/vnd.google-apps.folder",
    ),
  };
}

export async function readDeckFile(fileId: string): Promise<{ ok: true; name: string; bytes: Buffer } | DriveFailure> {
  const { GoogleDriveTools } = await import("@/lib/app-data/types.ts");
  const attempts: Record<string, unknown>[] = [
    { fileId, exportMimeType: PPTX_MIME },
    { file_id: fileId, export_mime_type: PPTX_MIME },
    { fileId },
    { file_id: fileId },
  ];
  let last: CallToolResult | null = null;
  for (const args of attempts) {
    const result = await driveCall(GoogleDriveTools.readFile, args);
    last = result;
    if (!result.ok) {
      if (result.pending || result.loginRequired) return pack(result);
      continue;
    }
    const bytes = readFileBytes(result.data);
    if (!bytes) continue;
    const record =
      result.data && typeof result.data === "object" ? (result.data as Record<string, unknown>) : {};
    const name =
      typeof record.name === "string"
        ? record.name
        : typeof record.title === "string"
          ? record.title
          : "Google Drive.pptx";
    return { ok: true, name, bytes };
  }
  if (last && !last.ok) return pack(last);
  return {
    ok: false,
    kind: "error",
    pending: false,
    message: "That Drive file didn't include the slides. Download it as PowerPoint and drop it here.",
  };
}

export async function saveDeckFile(input: {
  folderId: string | null;
  name: string;
  bytes: Buffer;
}): Promise<{ ok: true; name: string } | DriveFailure> {
  const parent = input.folderId && input.folderId !== "root" ? input.folderId : undefined;
  const payload = {
    name: input.name,
    mime_type: PPTX_MIME,
    mimeType: PPTX_MIME,
    content_base64: input.bytes.toString("base64"),
    ...(parent ? { parent_id: parent, folder_id: parent, parentId: parent } : {}),
  };
  let last: CallToolResult | null = null;
  for (const toolName of UPLOAD_TOOLS) {
    const result = await driveCall(toolName, payload);
    last = result;
    if (result.ok) return { ok: true, name: input.name };
    if (result.pending || result.loginRequired) return pack(result);
    const message = (result.errorMessage ?? "").toLowerCase();
    if (!/unknown|not found|not_found|unexpected|unrecognized|invalid|required/.test(message)) {
      return pack(result);
    }
  }
  return last
    ? pack(last)
    : { ok: false, kind: "error", pending: false, message: "Couldn't save the deck to Google Drive." };
}

export async function createDriveFolder(name: string, parentId: string | null): Promise<DriveListResult> {
  const { GoogleDriveTools } = await import("@/lib/app-data/types.ts");
  const parent = parentId && parentId !== "root" ? parentId : undefined;
  const result = await driveCall(GoogleDriveTools.createFolder, {
    name,
    ...(parent ? { parent_id: parent, folder_id: parent, parentId: parent } : {}),
  });
  if (!result.ok) return pack(result);
  return { ok: true, entries: readDriveEntries(result.data) };
}
