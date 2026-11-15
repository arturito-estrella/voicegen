import { useEffect, useMemo, useState } from "react";
import { ChevronLeft, ChevronRight, Download, FileUp, HardDrive, Square, Volume2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { DriveBrowser } from "@/components/drive-browser";
import { translateDriveFile } from "@/lib/drive.functions";
import { translateDeck } from "@/lib/voicegen.functions";
import type { DeckResult, TranslateRequest, ViewBlock, ViewSlide } from "@/lib/voicegen-types";
import { cn } from "@/lib/cn";

const ACCEPT =
  ".ppt,.pptx,.pptm,.pps,.ppsx,application/vnd.ms-powerpoint,application/vnd.openxmlformats-officedocument.presentationml.presentation";

function errorText(error: unknown): string {
  if (error instanceof Error && error.message) return error.message.replace(/^Error:\s*/, "");
  return "Something went wrong. Try again.";
}

function fileToBase64(file: File): Promise<string> {
  return file.arrayBuffer().then((buffer) => {
    const bytes = new Uint8Array(buffer);
    let binary = "";
    const size = 0x8000;
    for (let index = 0; index < bytes.length; index += size) {
      binary += String.fromCharCode(...bytes.subarray(index, index + size));
    }
    return btoa(binary);
  });
}

function downloadPptx(base64: string, filename: string) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  const url = URL.createObjectURL(
    new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    }),
  );
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function onSlideText(blocks: ViewBlock[]): ViewBlock[] {
  return blocks.filter((block) => block.kind !== "notes");
}

function noteBlocks(slide: ViewSlide): ViewBlock[] {
  const shown = new Set(onSlideText(slide.blocks).map((block) => block.original.trim()));
  return slide.blocks.filter((block) => block.kind === "notes" && !shown.has(block.original.trim()));
}

function scriptFor(blocks: ViewBlock[]): string {
  return blocks
    .map((block) => block.spanish.trim())
    .filter(Boolean)
    .join("\n");
}

function BlockCopy({ blocks, spanish }: { blocks: ViewBlock[]; spanish: boolean }) {
  const titles = blocks.filter((block) => block.kind === "title");
  const body = blocks.filter((block) => block.kind !== "title");
  const heading = titles.length > 0 ? titles : body.length > 1 && (spanish ? body[0]!.spanish : body[0]!.original).length < 90 ? [body[0]!] : [];
  const rest = titles.length > 0 ? body : heading.length > 0 ? body.slice(1) : body;
  const read = (block: ViewBlock) => (spanish ? block.spanish : block.original);

  if (heading.length === 0 && rest.length === 0) {
    return <p className="text-muted-foreground">This slide has no text to translate.</p>;
  }

  return (
    <div className="space-y-4">
      {heading.length > 0 ? (
        <h2 className={cn("font-display text-3xl leading-tight text-balance", spanish ? "text-foreground" : "text-muted-foreground")} lang={spanish ? "es" : "en"}>
          {heading.map((block) => read(block)).join(" ")}
        </h2>
      ) : null}
      {rest.length > 0 ? (
        <div className={cn("space-y-3 text-pretty", spanish ? "text-lg leading-normal" : "text-sm leading-normal text-muted-foreground")} lang={spanish ? "es" : "en"}>
          {rest.map((block) => (
            <p key={`${block.id}-${spanish ? "es" : "en"}`} className="whitespace-pre-wrap">
              {read(block)}
            </p>
          ))}
        </div>
      ) : null}
    </div>
  );
}

export function VoiceGenApp() {
  const [link, setLink] = useState("");
  const [dragging, setDragging] = useState(false);
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("");
  const [error, setError] = useState("");
  const [deck, setDeck] = useState<DeckResult | null>(null);
  const [slideIndex, setSlideIndex] = useState(0);
  const [listenState, setListenState] = useState<"idle" | "playing">("idle");
  const [spoken, setSpoken] = useState<"slide" | "notes" | null>(null);
  const [listenNote, setListenNote] = useState("");
  const [driveOpen, setDriveOpen] = useState(false);
  const [saveOpen, setSaveOpen] = useState(false);
  const [saveNote, setSaveNote] = useState("");

  useEffect(() => {
    return () => {
      window.speechSynthesis?.cancel();
    };
  }, []);

  const slide = deck?.slides[slideIndex] ?? null;
  const face = useMemo(() => (slide ? onSlideText(slide.blocks) : []), [slide]);
  const notes = useMemo(() => (slide ? noteBlocks(slide) : []), [slide]);
  const textSlides = deck?.slides.filter((item) => onSlideText(item.blocks).some((block) => block.original.trim())).length ?? 0;

  function stopAudio() {
    window.speechSynthesis?.cancel();
    setListenState("idle");
    setSpoken(null);
  }

  function reset() {
    stopAudio();
    setDeck(null);
    setSlideIndex(0);
    setError("");
    setStatus("");
    setListenNote("");
    setSaveNote("");
    setSaveOpen(false);
    setDriveOpen(false);
  }

  async function translate(input: TranslateRequest) {
    stopAudio();
    setBusy(true);
    setError("");
    setDeck(null);
    setSaveNote("");
    setStatus("Reading the slides and translating them into Spanish.");
    try {
      const result = await translateDeck({ data: input });
      setDeck(result);
      setSlideIndex(0);
      setStatus("");
      setDriveOpen(false);
    } catch (caught) {
      setError(errorText(caught));
      setStatus("");
    } finally {
      setBusy(false);
    }
  }

  async function onDriveFile(file: { id: string; name: string }) {
    if (busy) return;
    stopAudio();
    setBusy(true);
    setError("");
    setDeck(null);
    setSaveNote("");
    setStatus(`Opening ${file.name} from Google Drive.`);
    try {
      const result = await translateDriveFile({ data: { fileId: file.id } });
      if (!("slides" in result)) {
        setError(result.message);
        setStatus("");
        return;
      }
      setDeck(result);
      setSlideIndex(0);
      setStatus("");
      setDriveOpen(false);
    } catch (caught) {
      setError(errorText(caught));
      setStatus("");
    } finally {
      setBusy(false);
    }
  }

  async function onFile(file: File | undefined) {
    if (!file || busy) return;
    const lower = file.name.toLowerCase();
    if (!/\.(ppt|pptx|pptm|pps|ppsx)$/.test(lower)) {
      setError("Use a PowerPoint file (.ppt or .pptx).");
      return;
    }
    if (file.size > 12 * 1024 * 1024) {
      setError("That file is over 12 MB. Try a smaller deck.");
      return;
    }
    const dataBase64 = await fileToBase64(file);
    await translate({ kind: "file", name: file.name, dataBase64 });
  }

  function chooseSlide(next: number) {
    if (!deck) return;
    stopAudio();
    setListenNote("");
    setSlideIndex(Math.min(deck.slides.length - 1, Math.max(0, next)));
  }

  function listen(blocks: ViewBlock[], target: "slide" | "notes") {
    if (!slide || typeof window === "undefined" || !window.speechSynthesis) {
      setListenNote("This browser can't play speech.");
      return;
    }
    if (spoken === target && listenState === "playing") {
      stopAudio();
      return;
    }
    stopAudio();
    const full = scriptFor(blocks);
    if (!full) return;
    const text = full.slice(0, 3500);
    const utterance = new SpeechSynthesisUtterance(text);
    utterance.lang = "es-ES";
    utterance.rate = 0.95;
    const spanish = window.speechSynthesis.getVoices().find((item) => item.lang.toLowerCase().startsWith("es"));
    if (spanish) utterance.voice = spanish;
    utterance.onend = () => {
      setListenState("idle");
      setSpoken(null);
    };
    utterance.onerror = () => {
      setListenState("idle");
      setSpoken(null);
      setListenNote("Spanish speech didn't start. Try again.");
    };
    setSpoken(target);
    setListenState("playing");
    setListenNote(full.length > 3500 ? "Playing the first part of a long slide." : "");
    window.speechSynthesis.speak(utterance);
  }

  function controlLabel(target: "slide" | "notes", idle: string) {
    if (spoken === target && listenState === "playing") return "Stop";
    return idle;
  }

  return (
    <main className="safe-x mx-auto min-h-dvh w-full max-w-2xl">
      <header className="safe-top sticky top-0 z-10 mb-6 flex flex-wrap items-center justify-between gap-3 border-b border-border bg-background py-3">
        <button type="button" onClick={reset} className="font-display text-2xl leading-none text-foreground">
          VoiceGen
        </button>
        {deck ? (
          <div className="flex items-center gap-2">
            <Button variant="secondary" onClick={() => setSaveOpen((open) => !open)}>
              <HardDrive className="size-4" aria-hidden="true" />
              Save to Drive
            </Button>
            <Button onClick={() => downloadPptx(deck.pptxBase64, deck.downloadName)}>
              <Download className="size-4" aria-hidden="true" />
              Download
            </Button>
          </div>
        ) : null}
      </header>

      {!deck ? (
        <section className="space-y-6 pb-10">
          <div className="space-y-3">
            <p className="text-sm font-medium text-muted-foreground">English to Spanish</p>
            <h1 className="font-display text-5xl leading-none text-balance">Every slide, in Spanish.</h1>
            <p className="max-w-xl text-pretty text-lg leading-normal text-muted-foreground">
              Drop a PowerPoint file, pick one from Google Drive, or paste a Slides link. The words on each slide are translated. Nothing is read aloud unless you ask.
            </p>
          </div>

          <label
            className={cn(
              "flex min-h-48 cursor-pointer flex-col items-center justify-center gap-3 rounded-xl border border-dashed border-border bg-card px-6 py-8 text-center",
              dragging && "border-foreground",
              busy && "pointer-events-none opacity-60",
            )}
            onDragOver={(event) => {
              event.preventDefault();
              setDragging(true);
            }}
            onDragLeave={() => setDragging(false)}
            onDrop={(event) => {
              event.preventDefault();
              setDragging(false);
              void onFile(event.dataTransfer.files[0]);
            }}
          >
            <FileUp className="size-6 text-muted-foreground" aria-hidden="true" />
            <span className="text-lg font-medium">Drop a .ppt or .pptx file</span>
            <span className="text-sm text-muted-foreground">PowerPoint, or a Google Slides file saved as PowerPoint</span>
            <input
              className="sr-only"
              type="file"
              accept={ACCEPT}
              disabled={busy}
              onChange={(event) => {
                void onFile(event.target.files?.[0]);
                event.currentTarget.value = "";
              }}
            />
          </label>

          <Button variant="secondary" className="w-full" disabled={busy} onClick={() => setDriveOpen((open) => !open)}>
            <HardDrive className="size-4" aria-hidden="true" />
            Choose from Google Drive
          </Button>
          {driveOpen ? (
            <DriveBrowser mode="pick" disabled={busy} onClose={() => setDriveOpen(false)} onPick={(file) => void onDriveFile(file)} />
          ) : null}

          <form
            className="space-y-3"
            onSubmit={(event) => {
              event.preventDefault();
              if (!busy) void translate({ kind: "google", url: link });
            }}
          >
            <label className="block space-y-2" htmlFor="slides-link">
              <span className="text-sm font-medium">Google Slides link</span>
              <Input
                id="slides-link"
                value={link}
                disabled={busy}
                placeholder="https://docs.google.com/presentation/d/…"
                inputMode="url"
                autoCapitalize="none"
                autoCorrect="off"
                onChange={(event) => setLink(event.target.value)}
              />
            </label>
            <Button type="submit" variant="secondary" disabled={busy || link.trim().length < 12} className="w-full sm:w-auto">
              Translate this link
            </Button>
            <p className="text-sm text-muted-foreground">
              The deck needs to be shared so anyone with the link can view it. Private decks can be downloaded as PowerPoint and dropped above.
            </p>
          </form>

          <Button variant="ghost" disabled={busy} onClick={() => void translate({ kind: "sample" })}>
            Try a sample deck
          </Button>

          {busy ? (
            <p className="text-sm text-muted-foreground" role="status">
              {status}
            </p>
          ) : null}
          {error ? (
            <p className="rounded-md border border-border bg-card px-4 py-3 text-sm" role="alert">
              {error}
            </p>
          ) : null}
        </section>
      ) : (
        <section className="space-y-5 pb-48">
          <div className="space-y-2">
            <p className="text-sm text-muted-foreground">
              {deck.sourceName} · {deck.slides.length} slides
              {textSlides !== deck.slides.length ? ` · text on ${textSlides}` : ""}
            </p>
            <p className="text-pretty text-sm leading-normal text-muted-foreground">
              {deck.preservedLayout
                ? "Pictures and layout stay as they were. The words are now Spanish."
                : "This older .ppt file was rebuilt as a new PowerPoint file with the Spanish text."}
            </p>
            {deck.warning ? <p className="text-sm text-muted-foreground">{deck.warning}</p> : null}
            {saveNote ? <p className="text-sm text-muted-foreground">{saveNote}</p> : null}
            {saveOpen ? (
              <DriveBrowser
                mode="save"
                saveName={deck.downloadName}
                saveBase64={deck.pptxBase64}
                onClose={() => setSaveOpen(false)}
                onSaved={(folderName) => {
                  setSaveOpen(false);
                  setSaveNote(`Saved in ${folderName} on Google Drive.`);
                }}
              />
            ) : null}
          </div>

          <div className="flex gap-2 overflow-x-auto pb-1">
            {deck.slides.map((item, index) => (
              <button
                key={item.index}
                type="button"
                aria-label={`Slide ${item.index}`}
                aria-current={index === slideIndex}
                onClick={() => chooseSlide(index)}
                className={cn(
                  "size-11 shrink-0 rounded-md border text-sm tabular-nums",
                  index === slideIndex
                    ? "border-primary bg-primary text-primary-foreground"
                    : "border-border bg-card text-foreground",
                )}
              >
                {item.index}
              </button>
            ))}
          </div>

          {slide ? (
            <article className="space-y-6 rounded-xl border border-border bg-card p-5">
              <p className="text-sm tabular-nums text-muted-foreground">
                Slide {slide.index} of {deck.slides.length}
              </p>
              <BlockCopy blocks={face} spanish />
              <div className="border-t border-border pt-4">
                <p className="mb-3 text-sm font-medium text-muted-foreground">Original English</p>
                <BlockCopy blocks={face} spanish={false} />
              </div>
              {notes.length > 0 ? (
                <div className="border-t border-border pt-4">
                  <p className="mb-3 text-sm font-medium text-muted-foreground">Speaker notes</p>
                  <div className="space-y-3" lang="es">
                    {notes.map((block) => (
                      <p key={block.id} className="whitespace-pre-wrap text-pretty leading-normal">
                        {block.spanish}
                      </p>
                    ))}
                  </div>
                  <p className="mt-3 text-sm whitespace-pre-wrap text-muted-foreground" lang="en">
                    {notes.map((block) => block.original).join("\n")}
                  </p>
                  <Button
                    variant="outline"
                    className="mt-4"
                    onClick={() => void listen(notes, "notes")}
                    disabled={scriptFor(notes).length === 0}
                  >
                    {spoken === "notes" && listenState === "playing" ? (
                      <Square className="size-4" />
                    ) : (
                      <Volume2 className="size-4" />
                    )}
                    {controlLabel("notes", "Listen to notes")}
                  </Button>
                </div>
              ) : null}
            </article>
          ) : null}
        </section>
      )}

      {deck && slide ? (
        <div className="fixed inset-x-0 bottom-0 z-10 border-t border-border bg-background">
          <div className="safe-x safe-bottom mx-auto flex max-w-2xl flex-col gap-3 px-0 py-3">
            <Button
              variant="secondary"
              className="w-full"
              disabled={scriptFor(face).length === 0}
              onClick={() => listen(face, "slide")}
            >
              {spoken === "slide" && listenState === "playing" ? (
                <Square className="size-4" aria-hidden="true" />
              ) : (
                <Volume2 className="size-4" aria-hidden="true" />
              )}
              {controlLabel("slide", "Listen")}
            </Button>
            {listenNote ? <p className="text-sm text-muted-foreground">{listenNote}</p> : null}
            <div className="flex items-center justify-between gap-3">
              <Button variant="outline" size="icon" aria-label="Previous slide" disabled={slideIndex === 0} onClick={() => chooseSlide(slideIndex - 1)}>
                <ChevronLeft className="size-5" />
              </Button>
              <p className="text-sm tabular-nums text-muted-foreground">
                {slide.index} / {deck.slides.length}
              </p>
              <Button
                variant="outline"
                size="icon"
                aria-label="Next slide"
                disabled={slideIndex >= deck.slides.length - 1}
                onClick={() => chooseSlide(slideIndex + 1)}
              >
                <ChevronRight className="size-5" />
              </Button>
            </div>
          </div>
        </div>
      ) : null}
    </main>
  );
}
