import { useCallback, useEffect, useState } from "react";
import { ChevronLeft, Folder, HardDrive, Loader2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { redirectToLoginIfRequired, useRefetchWhenConnectorReady } from "@/lib/app-data";
import {
  createDriveFolder,
  listDriveDecks,
  listDriveFolders,
  saveDeckToDrive,
} from "@/lib/drive.functions";
import type { DriveEntry } from "@/lib/drive-files";

type DriveError = {
  kind: string;
  message: string;
  loginUrl?: string;
  pending: boolean;
};

function isFailure(value: { ok: boolean }): value is DriveError & { ok: false } {
  return value.ok === false;
}

export function DriveBrowser({
  mode,
  disabled,
  saveName,
  saveBase64,
  onClose,
  onPick,
  onSaved,
}: {
  mode: "pick" | "save";
  disabled?: boolean;
  saveName?: string;
  saveBase64?: string;
  onClose: () => void;
  onPick?: (file: DriveEntry) => void;
  onSaved?: (folderName: string) => void;
}) {
  const [query, setQuery] = useState("");
  const [entries, setEntries] = useState<DriveEntry[]>([]);
  const [stack, setStack] = useState<{ id: string | null; name: string }[]>([{ id: null, name: "My Drive" }]);
  const [error, setError] = useState<DriveError | null>(null);
  const [loading, setLoading] = useState(true);
  const [folderName, setFolderName] = useState("");
  const [saving, setSaving] = useState(false);
  const current = stack[stack.length - 1] ?? { id: null, name: "My Drive" };

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const result =
        mode === "pick"
          ? await listDriveDecks({ data: { query } })
          : await listDriveFolders({ data: { parentId: current.id } });
      if (isFailure(result)) {
        setEntries([]);
        setError(result);
        return;
      }
      setEntries(result.entries);
    } catch (caught) {
      setEntries([]);
      setError({
        kind: "error",
        pending: false,
        message: caught instanceof Error ? caught.message : "Couldn't reach Google Drive.",
      });
    } finally {
      setLoading(false);
    }
  }, [mode, query, current.id]);

  useEffect(() => {
    const timer = window.setTimeout(() => {
      void load();
    }, mode === "pick" ? 250 : 0);
    return () => window.clearTimeout(timer);
  }, [load, mode]);

  useRefetchWhenConnectorReady(error?.pending === true, load);

  function continueWithGrok() {
    if (!error?.loginUrl) return;
    redirectToLoginIfRequired({
      ok: false,
      data: null,
      loginRequired: true,
      loginUrl: error.loginUrl,
    });
  }

  async function makeFolder() {
    const name = folderName.trim();
    if (!name) return;
    setSaving(true);
    setError(null);
    try {
      const result = await createDriveFolder({ data: { name, parentId: current.id } });
      if (isFailure(result)) {
        setError(result);
        return;
      }
      setFolderName("");
      const created = result.entries[0];
      if (created) setStack((prev) => [...prev, { id: created.id, name: created.name }]);
      else await load();
    } catch (caught) {
      setError({
        kind: "error",
        pending: false,
        message: caught instanceof Error ? caught.message : "Couldn't create that folder.",
      });
    } finally {
      setSaving(false);
    }
  }

  async function saveHere() {
    if (!saveBase64 || !saveName) return;
    setSaving(true);
    setError(null);
    try {
      const result = await saveDeckToDrive({
        data: { name: saveName, dataBase64: saveBase64, folderId: current.id },
      });
      if (isFailure(result)) {
        setError(result);
        return;
      }
      onSaved?.(current.name);
    } catch (caught) {
      setError({
        kind: "error",
        pending: false,
        message: caught instanceof Error ? caught.message : "Couldn't save to Google Drive.",
      });
    } finally {
      setSaving(false);
    }
  }

  return (
    <div className="space-y-4 rounded-xl border border-border bg-card p-4">
      <div className="flex items-start justify-between gap-3">
        <div className="space-y-1">
          <p className="flex items-center gap-2 text-sm font-medium">
            <HardDrive className="size-4" aria-hidden="true" />
            {mode === "pick" ? "Google Drive" : current.name}
          </p>
          <p className="text-sm text-muted-foreground">
            {mode === "pick"
              ? "PowerPoint and Google Slides files from your Drive."
              : "Choose the folder for the Spanish deck."}
          </p>
        </div>
        <Button type="button" variant="ghost" onClick={onClose}>
          Close
        </Button>
      </div>

      {mode === "pick" ? (
        <Input
          value={query}
          placeholder="Search by name"
          disabled={disabled || saving}
          onChange={(event) => setQuery(event.target.value)}
        />
      ) : (
        <div className="flex items-center gap-2">
          <Button
            type="button"
            variant="outline"
            size="icon"
            aria-label="Back"
            disabled={stack.length < 2 || saving}
            onClick={() => setStack((prev) => prev.slice(0, -1))}
          >
            <ChevronLeft className="size-5" />
          </Button>
          <p className="min-w-0 flex-1 truncate text-sm text-muted-foreground">
            {stack.map((item) => item.name).join(" / ")}
          </p>
        </div>
      )}

      {error?.kind === "login" ? (
        <div className="space-y-3">
          <p className="text-sm">{error.message}</p>
          {error.loginUrl ? (
            <Button type="button" onClick={continueWithGrok}>
              Continue with Grok
            </Button>
          ) : null}
        </div>
      ) : null}

      {error && error.kind !== "login" ? (
        <p className="text-sm" role="alert">
          {error.pending ? "Connecting to your data…" : error.message}
        </p>
      ) : null}

      {loading ? (
        <p className="flex items-center gap-2 text-sm text-muted-foreground" role="status">
          <Loader2 className="size-4 animate-spin" aria-hidden="true" />
          Looking in Drive
        </p>
      ) : null}

      {!loading && !error && entries.length === 0 ? (
        <p className="text-sm text-muted-foreground">
          {mode === "pick" ? "No PowerPoint or Slides files matched." : "No folders here."}
        </p>
      ) : null}

      <ul className="max-h-72 space-y-2 overflow-y-auto">
        {entries.map((entry) => (
          <li key={entry.id}>
            <button
              type="button"
              disabled={disabled || saving}
              className="flex w-full items-center gap-3 rounded-md border border-border px-3 py-3 text-left text-sm hover:bg-background disabled:opacity-60"
              onClick={() => {
                if (mode === "pick") onPick?.(entry);
                else setStack((prev) => [...prev, { id: entry.id, name: entry.name }]);
              }}
            >
              {mode === "save" ? <Folder className="size-4 shrink-0" aria-hidden="true" /> : null}
              <span className="min-w-0 flex-1 truncate">{entry.name}</span>
            </button>
          </li>
        ))}
      </ul>

      {mode === "save" ? (
        <div className="space-y-3 border-t border-border pt-3">
          <Button type="button" className="w-full" disabled={saving || disabled || !saveBase64} onClick={() => void saveHere()}>
            {saving ? "Saving" : `Save in ${current.name}`}
          </Button>
          <div className="flex gap-2">
            <Input
              value={folderName}
              placeholder="New folder name"
              disabled={saving}
              onChange={(event) => setFolderName(event.target.value)}
            />
            <Button type="button" variant="secondary" disabled={saving || folderName.trim().length === 0} onClick={() => void makeFolder()}>
              Create
            </Button>
          </div>
        </div>
      ) : null}
    </div>
  );
}
