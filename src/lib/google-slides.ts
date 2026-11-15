const GOOGLE_HOSTS = new Set(["docs.google.com", "drive.google.com"]);

export function googleExportUrl(input: string): string | null {
  let url: URL;
  try {
    url = new URL(input.trim());
  } catch {
    return null;
  }
  if (url.protocol !== "https:") return null;
  const host = url.hostname.toLowerCase();
  if (!GOOGLE_HOSTS.has(host)) return null;

  const presentation = url.pathname.match(
    /\/presentation\/(?:u\/\d+\/)?d\/([a-zA-Z0-9_-]+)/,
  );
  if (presentation && presentation[1] !== "e") {
    return `https://docs.google.com/presentation/d/${presentation[1]}/export/pptx`;
  }

  const published = url.pathname.match(
    /\/presentation\/d\/e\/(2PACX-[a-zA-Z0-9_-]+)/,
  );
  if (published) {
    return `https://docs.google.com/presentation/d/e/${published[1]}/export/pptx`;
  }

  const driveFile = url.pathname.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (driveFile) {
    return `https://docs.google.com/presentation/d/${driveFile[1]}/export/pptx`;
  }

  const id = url.searchParams.get("id");
  if (id && /^[a-zA-Z0-9_-]+$/.test(id)) {
    return `https://docs.google.com/presentation/d/${id}/export/pptx`;
  }

  return null;
}

export function isAllowedSlideHost(hostname: string): boolean {
  const host = hostname.toLowerCase();
  return (
    host === "docs.google.com" ||
    host === "drive.google.com" ||
    host.endsWith(".googleusercontent.com") ||
    host.endsWith(".google.com")
  );
}

export function filenameFromDisposition(header: string | null): string | null {
  if (!header) return null;
  const star = /filename\*=UTF-8''([^;]+)/i.exec(header);
  if (star?.[1]) {
    try {
      return decodeURIComponent(star[1].replace(/["']/g, ""));
    } catch {
      return star[1];
    }
  }
  const plain = /filename="?([^";]+)"?/i.exec(header);
  return plain?.[1]?.trim() || null;
}
