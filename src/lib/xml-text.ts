export function decodeXml(value: string): string {
  return value.replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]+);/g, (all, ent: string) => {
    if (ent === "amp") return "&";
    if (ent === "lt") return "<";
    if (ent === "gt") return ">";
    if (ent === "quot") return '"';
    if (ent === "apos") return "'";
    if (ent.startsWith("#x")) {
      const code = Number.parseInt(ent.slice(2), 16);
      return Number.isFinite(code) ? String.fromCodePoint(code) : all;
    }
    if (ent.startsWith("#")) {
      const code = Number.parseInt(ent.slice(1), 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : all;
    }
    return all;
  });
}

export function escapeXml(value: string): string {
  return value
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
    .replace(/&/g, "&")
    .replace(/</g, "<")
    .replace(/>/g, ">");
}

export function needsTranslation(text: string): boolean {
  const letters = text.match(/\p{L}/gu);
  return !!letters && letters.length >= 2;
}

export function maskFields(xml: string): { xml: string; restore: (value: string) => string } {
  const fields: string[] = [];
  const masked = xml.replace(/<a:fld\b[\s\S]*?<\/a:fld>/g, (match) => {
    const token = `<!--FLD${fields.length}-->`;
    fields.push(match);
    return token;
  });
  return {
    xml: masked,
    restore: (value) => value.replace(/<!--FLD(\d+)-->/g, (_, index) => fields[Number(index)] ?? ""),
  };
}

function withSpanish(runProps: string): string {
  if (!runProps) return `<a:rPr lang="es-ES" dirty="0"/>`;
  if (/lang="/.test(runProps)) return runProps.replace(/lang="[^"]*"/, 'lang="es-ES"');
  return runProps.replace(/<a:rPr\b/, '<a:rPr lang="es-ES"');
}

export function rewriteParagraph(para: string, spanish: string): string {
  const openEnd = para.indexOf(">");
  if (openEnd === -1) return para;
  const open = para.slice(0, openEnd + 1);
  const props =
    /<a:pPr\b[\s\S]*?<\/a:pPr>|<a:pPr\b[^>]*\/>/.exec(para)?.[0] ?? "";
  const endProps =
    /<a:endParaRPr\b[\s\S]*?<\/a:endParaRPr>|<a:endParaRPr\b[^>]*\/>/.exec(para)?.[0] ??
    "";
  const runProps =
    /<a:rPr\b[\s\S]*?<\/a:rPr>|<a:rPr\b[^>]*\/>/.exec(para)?.[0] ?? "";
  const styled = withSpanish(runProps);
  const lines = spanish.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  const runs = lines
    .map((line, index) => {
      const breakMark = index === 0 ? "" : "<a:br/>";
      return `${breakMark}<a:r>${styled}<a:t>${escapeXml(line)}</a:t></a:r>`;
    })
    .join("");
  return `${open}${props}${runs}${endProps}</a:p>`;
}

const PARA_RE = /<a:p\b[^>]*>[\s\S]*?<\/a:p>/g;
const RUN_RE = /<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g;

export type ParaHit = {
  id: string;
  kind: "title" | "body" | "notes";
  text: string;
  start: number;
  end: number;
};

export function paragraphKind(
  xml: string,
  start: number,
  forceNotes: boolean,
): "title" | "body" | "notes" {
  if (forceNotes) return "notes";
  const before = xml.slice(Math.max(0, start - 4000), start);
  const shape = before.lastIndexOf("<p:sp");
  const shapeEnd = before.lastIndexOf("</p:sp>");
  const scope = shape > shapeEnd ? before.slice(shape) : before;
  const placeholder = /<p:ph\b[^>]*\/?>/.exec(scope)?.[0] ?? "";
  if (/type="(?:title|ctrTitle)"/.test(placeholder)) return "title";
  return "body";
}

export function extractParagraphs(
  xml: string,
  idPrefix: string,
  forceNotes: boolean,
): ParaHit[] {
  const hits: ParaHit[] = [];
  PARA_RE.lastIndex = 0;
  let match: RegExpExecArray | null;
  let ordinal = 0;
  while ((match = PARA_RE.exec(xml))) {
    const para = match[0];
    const chunks: string[] = [];
    RUN_RE.lastIndex = 0;
    let run: RegExpExecArray | null;
    let cursor = 0;
    while ((run = RUN_RE.exec(para))) {
      const between = para.slice(cursor, run.index);
      if (/<a:br\b/.test(between)) chunks.push("\n");
      chunks.push(decodeXml(run[1]));
      cursor = run.index + run[0].length;
    }
    const text = chunks.join("").replace(/\u00a0/g, " ").trim();
    if (!text) continue;
    hits.push({
      id: `${idPrefix}p${ordinal}`,
      kind: paragraphKind(xml, match.index, forceNotes),
      text,
      start: match.index,
      end: match.index + para.length,
    });
    ordinal += 1;
  }
  return hits;
}

export function applyHits(xml: string, hits: ParaHit[], spanishById: Map<string, string>): string {
  const ordered = [...hits].sort((a, b) => b.start - a.start);
  let next = xml;
  for (const hit of ordered) {
    const spanish = spanishById.get(hit.id);
    if (!spanish) continue;
    const rewritten = rewriteParagraph(next.slice(hit.start, hit.end), spanish);
    next = next.slice(0, hit.start) + rewritten + next.slice(hit.end);
  }
  return next;
}
