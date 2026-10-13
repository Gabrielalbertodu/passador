import { createId } from "./local-db";
import type { PresentationItem, PresentationKind } from "./presentation-types";

export interface ParsedImport {
  assetId: string;
  kind: PresentationKind;
  title: string;
  mimeType: string;
  size: number;
  blob: Blob;
  items: PresentationItem[];
}

const extensions: Record<string, { kind: PresentationKind; mime: string }> = {
  png: { kind: "image", mime: "image/png" },
  jpg: { kind: "image", mime: "image/jpeg" },
  jpeg: { kind: "image", mime: "image/jpeg" },
  webp: { kind: "image", mime: "image/webp" },
  gif: { kind: "image", mime: "image/gif" },
  svg: { kind: "image", mime: "image/svg+xml" },
  mp4: { kind: "video", mime: "video/mp4" },
  webm: { kind: "video", mime: "video/webm" },
  mp3: { kind: "audio", mime: "audio/mpeg" },
  wav: { kind: "audio", mime: "audio/wav" },
  m4a: { kind: "audio", mime: "audio/mp4" },
  aac: { kind: "audio", mime: "audio/aac" },
  ogg: { kind: "audio", mime: "audio/ogg" },
  opus: { kind: "audio", mime: "audio/ogg" },
  flac: { kind: "audio", mime: "audio/flac" },
  pdf: { kind: "pdf", mime: "application/pdf" },
  pptx: { kind: "pptx", mime: "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
};

function fileKind(file: File): { kind: PresentationKind; mime: string } | undefined {
  const extension = file.name.split(".").pop()?.toLowerCase();
  return extension ? extensions[extension] : undefined;
}

function createSlideItem(
  assetId: string,
  kind: PresentationKind,
  fileName: string,
  options: Partial<PresentationItem> = {},
): PresentationItem {
  return {
    id: createId(),
    kind,
    title: options.title ?? fileName.replace(/\.[^.]+$/, ""),
    fileName,
    assetId,
    volume: 0.8,
    startAt: kind === "video" ? 0 : undefined,
    ...options,
  };
}

function configurePdfWorker(pdfjs: typeof import("pdfjs-dist")): void {
  if (!pdfjs.GlobalWorkerOptions.workerSrc) {
    pdfjs.GlobalWorkerOptions.workerSrc = new URL("pdfjs-dist/build/pdf.worker.min.mjs", import.meta.url).toString();
  }
}

export async function prepareFile(file: File): Promise<ParsedImport> {
  const definition = fileKind(file);
  if (!definition) {
    throw new Error(`O formato de “${file.name}” não é suportado. Use PPTX, PDF, PNG, JPG, MP4, WEBM ou áudio.`);
  }

  const assetId = createId();
  const baseName = file.name.replace(/\.[^.]+$/, "");
  const common = { assetId, fileName: file.name };
  let items: PresentationItem[];

  if (definition.kind === "pptx") {
    const [{ default: JSZip }, buffer] = await Promise.all([
      import("jszip"),
      file.arrayBuffer(),
    ]);
    const archive = await JSZip.loadAsync(buffer);
    const presentationXml = await archive.file("ppt/presentation.xml")?.async("text");
    if (!presentationXml) throw new Error(`Não foi possível ler a estrutura de “${file.name}”.`);
    const parsedXml = new DOMParser().parseFromString(presentationXml, "application/xml");
    if (parsedXml.getElementsByTagName("parsererror").length) throw new Error(`A estrutura de “${file.name}” está inválida.`);
    const namespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const slideCount = parsedXml.getElementsByTagNameNS(namespace, "sldId").length;
    if (!slideCount) throw new Error(`“${file.name}” não contém slides para importar.`);
    const slideSize = parsedXml.getElementsByTagNameNS(namespace, "sldSz")[0];
    const width = Number(slideSize?.getAttribute("cx")) / 9525;
    const height = Number(slideSize?.getAttribute("cy")) / 9525;
    items = Array.from({ length: slideCount }, (_, index) =>
      createSlideItem(assetId, "pptx", file.name, {
        title: `${baseName} · slide ${index + 1}`,
        slideNumber: index + 1,
        slideCount,
        slideWidth: Number.isFinite(width) ? width : undefined,
        slideHeight: Number.isFinite(height) ? height : undefined,
      }),
    );
  } else if (definition.kind === "pdf") {
    const pdfjs = await import("pdfjs-dist");
    configurePdfWorker(pdfjs);
    const loadingTask = pdfjs.getDocument({ data: await file.arrayBuffer() });
    try {
      const document = await loadingTask.promise;
      const count = document.numPages;
      items = Array.from({ length: count }, (_, index) =>
        createSlideItem(assetId, "pdf", file.name, {
          title: `${baseName} · página ${index + 1}`,
          slideNumber: index + 1,
          slideCount: count,
        }),
      );
    } finally {
      await loadingTask.destroy();
    }
  } else {
    items = [
      createSlideItem(assetId, definition.kind, file.name, {
        ...common,
        mimeType: file.type || definition.mime,
      }),
    ];
  }

  return {
    assetId,
    kind: definition.kind,
    title: file.name,
    mimeType: file.type || definition.mime,
    size: file.size,
    blob: file,
    items,
  };
}

export function createLinkItem(url: string, title?: string): PresentationItem {
  return {
    id: createId(),
    kind: "link",
    title: title?.trim() || new URL(url).hostname,
    linkUrl: url,
  };
}

export function fileAccept(): string {
  return ".pptx,.pdf,.png,.jpg,.jpeg,.webp,.gif,.svg,.mp4,.webm,.mp3,.wav,.m4a,.aac,.ogg,.opus,.flac";
}
