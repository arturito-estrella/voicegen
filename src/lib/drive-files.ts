export type DriveEntry = {
  id: string;
  name: string;
  mimeType: string;
};

const DECK_MIMES = [
  "application/vnd.google-apps.presentation",
  "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  "application/vnd.ms-powerpoint",
  "application/vnd.ms-powerpoint.presentation.macroenabled.12",
] as const;

export const FOLDER_MIME = "application/vnd.google-apps.folder";
export const PPTX_MIME =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation";

export function isDeckMime(mimeType: string): boolean {
  return (DECK_MIMES as readonly string[]).includes(mimeType) || mimeType.includes("presentation") || mimeType.includes("powerpoint");
}

function asRecord(value: unknown): Record<string, unknown> | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  return value as Record<string, unknown>;
}

function stringField(record: Record<string, unknown>, keys: string[]): string {
  for (const key of keys) {
    const value = record[key];
    if (typeof value === "string" && value.trim()) return value.trim();
  }
  return "";
}

export function readDriveEntries(data: unknown): DriveEntry[] {
  const root = asRecord(data);
  const lists: unknown[] = [];
  if (Array.isArray(data)) lists.push(...data);
  if (root) {
    if (stringField(root, ["id", "file_id", "fileId"])) lists.push(root);
    for (const key of ["files", "items", "folders", "results", "entries"]) {
      if (Array.isArray(root[key])) lists.push(...(root[key] as unknown[]));
    }
  }
  const entries: DriveEntry[] = [];
  const seen = new Set<string>();
  for (const item of lists) {
    const record = asRecord(item);
    if (!record) continue;
    const id = stringField(record, ["id", "file_id", "fileId"]);
    const name = stringField(record, ["name", "title", "file_name", "fileName"]) || "Untitled";
    const mimeType = stringField(record, ["mimeType", "mime_type", "type"]);
    if (!id || seen.has(id)) continue;
    seen.add(id);
    entries.push({ id, name, mimeType });
  }
  return entries;
}

function looksLikeBase64(value: string): boolean {
  const compact = value.replace(/\s/g, "");
  return compact.length > 80 && /^[A-Za-z0-9+/=]+$/.test(compact.slice(0, 200));
}

export function readFileBytes(data: unknown): Buffer | null {
  const found = findBase64(data, 0);
  if (!found) return null;
  const bytes = Buffer.from(found, "base64");
  return bytes.length > 32 ? bytes : null;
}

function findBase64(value: unknown, depth: number): string | null {
  if (depth > 4 || value == null) return null;
  if (typeof value === "string") return looksLikeBase64(value) ? value.replace(/\s/g, "") : null;
  if (Array.isArray(value)) {
    for (const item of value) {
      const found = findBase64(item, depth + 1);
      if (found) return found;
    }
    return null;
  }
  const record = asRecord(value);
  if (!record) return null;
  for (const key of [
    "content_base64",
    "base64",
    "bytes",
    "data",
    "content",
    "file_content",
    "fileContent",
    "body",
  ]) {
    const found = findBase64(record[key], depth + 1);
    if (found) return found;
  }
  return null;
}

export function deckQuery(name: string): string {
  const mime = DECK_MIMES.map((item) => `mimeType = '${item}'`).join(" or ");
  const safe = name.replace(/'/g, "").trim();
  const nameClause = safe ? ` and name contains '${safe}'` : "";
  return `trashed = false and (${mime})${nameClause}`;
}

export function folderQuery(parentId: string | null): string {
  const parent = parentId && parentId !== "root" ? `'${parentId.replace(/'/g, "")}' in parents and ` : "";
  return `${parent}trashed = false and mimeType = '${FOLDER_MIME}'`;
}
