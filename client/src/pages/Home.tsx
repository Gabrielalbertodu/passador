import { lazy, Suspense, useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  ArrowLeft,
  ArrowRight,
  Check,
  Copy,
  ExternalLink,
  FileText,
  Film,
  GripVertical,
  Image as ImageIcon,
  Layers3,
  Link2,
  Loader2,
  Maximize2,
  Music2,
  Play,
  Plus,
  Presentation,
  Trash2,
  Upload,
  Volume2,
  X,
} from "lucide-react";
import MediaSlide from "@/components/MediaSlide";
const PptxViewerStage = lazy(() => import("@/components/PptxViewerStage"));
import { createLinkItem, fileAccept, prepareFile, type ParsedImport } from "@/lib/importers";
import { createId, estimateLocalStorage, readDeck, readMedia, writeDeck, writeMedia } from "@/lib/local-db";
import type { PresentationItem, PresentationKind } from "@/lib/presentation-types";

const EMPTY_DECK_TITLE = "Minha apresentação";

function kindLabel(kind: PresentationKind): string {
  return ({ pptx: "PowerPoint", pdf: "PDF", image: "Imagem", video: "Vídeo", audio: "Áudio", link: "Link" })[kind];
}

function KindIcon({ kind, size = 16 }: { kind: PresentationKind; size?: number }) {
  const common = { size, strokeWidth: 1.9 };
  switch (kind) {
    case "pptx": return <Presentation {...common} />;
    case "pdf": return <FileText {...common} />;
    case "image": return <ImageIcon {...common} />;
    case "video": return <Film {...common} />;
    case "audio": return <Music2 {...common} />;
    case "link": return <Link2 {...common} />;
  }
}

function formatTime(seconds?: number): string {
  if (!Number.isFinite(seconds) || seconds == null || seconds < 0) return "00:00";
  const value = Math.floor(seconds);
  const hh = Math.floor(value / 3600);
  const mm = Math.floor((value % 3600) / 60);
  const ss = value % 60;
  return hh > 0
    ? `${String(hh).padStart(2, "0")}:${String(mm).padStart(2, "0")}:${String(ss).padStart(2, "0")}`
    : `${String(mm).padStart(2, "0")}:${String(ss).padStart(2, "0")}`;
}

function formatSize(bytes: number): string {
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1).replace(".", ",")} MB`;
}

function createCopy(item: PresentationItem): PresentationItem {
  return { ...item, id: createId(), title: `${item.title} · cópia`, volume: item.volume ?? 0.8 };
}

function cacheProjectionPptx(current: Record<string, { item: PresentationItem; blob: Blob }>, key: string, item: PresentationItem, blob: Blob) {
  const cached = current[key];
  if (cached?.item.id === item.id) return current;
  return { ...current, [key]: { item, blob: cached?.blob ?? blob } };
}

function getFileStartIndices(items: PresentationItem[]): number[] {
  const seenFiles = new Set<string>();
  return items.reduce<number[]>((indices, item, index) => {
    const fileKey = item.assetId ?? item.id;
    if (!seenFiles.has(fileKey)) {
      seenFiles.add(fileKey);
      indices.push(index);
    }
    return indices;
  }, []);
}

function normalizeLink(value: string): string {
  const input = value.trim();
  const withScheme = /^https?:\/\//i.test(input) ? input : `https://${input}`;
  const parsed = new URL(withScheme);
  if (parsed.protocol !== "https:" && parsed.protocol !== "http:") {
    throw new Error("Insira um endereço da web que comece com http:// ou https://.");
  }
  return parsed.href;
}

export default function Home() {
  const [items, setItems] = useState<PresentationItem[]>([]);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [title, setTitle] = useState(EMPTY_DECK_TITLE);
  const [isLoaded, setIsLoaded] = useState(false);
  const [saveState, setSaveState] = useState<"saved" | "saving" | "error">("saved");
  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState("");
  const [presenting, setPresenting] = useState(false);
  const [projectionControlsVisible, setProjectionControlsVisible] = useState(false);
  const [mediaRestartToken, setMediaRestartToken] = useState(0);
  const [linkDialog, setLinkDialog] = useState(false);
  const [linkValue, setLinkValue] = useState("");
  const [linkName, setLinkName] = useState("");
  const [activeMedia, setActiveMedia] = useState<Blob | null>(null);
  const [activeMediaKey, setActiveMediaKey] = useState<string | null>(null);
  const [projectionPptx, setProjectionPptx] = useState<Record<string, { item: PresentationItem; blob: Blob }>>({});
  const [projectionPptxKey, setProjectionPptxKey] = useState<string | null>(null);
  const [loadedProjectionPptxKey, setLoadedProjectionPptxKey] = useState<string | null>(null);
  const [localSpace, setLocalSpace] = useState<string | null>(null);
  const [draggedId, setDraggedId] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const projectionRef = useRef<HTMLDivElement>(null);
  const copiedItem = useRef<PresentationItem | null>(null);
  const pptxKeyboardForwardedRef = useRef(false);
  const pptxReadyKeyRef = useRef<string | null>(null);
  const loadedPptxKeyRef = useRef<string | null>(null);
  const pptxSlideIndexRef = useRef<Record<string, number>>({});
  const pendingPptxNavigationRef = useRef<{ key: string; fromIndex: number; targetIndex: number } | null>(null);
  const projectionControlsTimer = useRef<number | undefined>(undefined);
  const noticeTimer = useRef<number | undefined>(undefined);

  const selectedItem = items.find((item) => item.id === selectedId) ?? items[0] ?? null;
  const activeIndex = Math.max(items.findIndex((item) => item.id === selectedItem?.id), 0);
  const selectedPosition = selectedItem ? activeIndex + 1 : 0;
  const itemTotal = items.length;
  const currentPptxKey = selectedItem?.kind === "pptx" ? (selectedItem.assetId ?? selectedItem.id) : null;
  const pptxStageEntry = useMemo(() => {
    if (!presenting) {
      return selectedItem?.kind === "pptx" && currentPptxKey && activeMedia && activeMediaKey === currentPptxKey
        ? { key: currentPptxKey, item: selectedItem, blob: activeMedia }
        : null;
    }
    if (selectedItem?.kind === "pptx" && currentPptxKey) {
      const selectedDeck = projectionPptx[currentPptxKey];
      if (selectedDeck) return { key: currentPptxKey, item: selectedItem, blob: selectedDeck.blob };
      if (activeMedia && activeMediaKey === currentPptxKey) return { key: currentPptxKey, item: selectedItem, blob: activeMedia };
    }
    for (const key of [projectionPptxKey, loadedProjectionPptxKey]) {
      const parkedDeck = key ? projectionPptx[key] : null;
      if (key && parkedDeck) return { key, ...parkedDeck };
    }
    return null;
  }, [activeMedia, activeMediaKey, currentPptxKey, loadedProjectionPptxKey, presenting, projectionPptx, projectionPptxKey, selectedItem]);
  const showNotice = useCallback((message: string) => {
    setNotice(message);
    if (noticeTimer.current != null) window.clearTimeout(noticeTimer.current);
    noticeTimer.current = window.setTimeout(() => setNotice(""), 3600);
  }, []);

  const revealProjectionControls = useCallback(() => {
    setProjectionControlsVisible(true);
    if (projectionControlsTimer.current != null) window.clearTimeout(projectionControlsTimer.current);
    projectionControlsTimer.current = window.setTimeout(() => setProjectionControlsVisible(false), 1800);
  }, []);

  useEffect(() => {
    void readDeck().then((deck) => {
      if (deck) {
        setTitle(deck.title || EMPTY_DECK_TITLE);
        setItems((deck.items || []).map((item) => {
          const cleanItem = { ...item };
          delete cleanItem.previewSvg;
          return cleanItem;
        }));
        setSelectedId(deck.items?.[0]?.id ?? null);
      }
    }).catch((error: unknown) => {
      console.error("Falha ao abrir os dados locais:", error);
      showNotice("O armazenamento local não pôde ser aberto neste navegador.");
    }).finally(() => setIsLoaded(true));
    void estimateLocalStorage().then(({ used, quota }) => {
      if (quota && used != null) setLocalSpace(`${formatSize(used)} de ${formatSize(quota)} usados neste navegador`);
    });
    return () => {
      if (noticeTimer.current != null) window.clearTimeout(noticeTimer.current);
      if (projectionControlsTimer.current != null) window.clearTimeout(projectionControlsTimer.current);
    };
  }, [showNotice]);

  useEffect(() => {
    if (!isLoaded) return;
    setSaveState("saving");
    const timeout = window.setTimeout(() => {
      void writeDeck(title, items).then(() => setSaveState("saved")).catch((error: unknown) => {
        console.error("Falha ao salvar a apresentação local:", error);
        setSaveState("error");
      });
    }, 220);
    return () => window.clearTimeout(timeout);
  }, [title, items, isLoaded]);

  useEffect(() => {
    let cancelled = false;
    if (!selectedItem?.assetId) {
      setActiveMedia(null);
      setActiveMediaKey(null);
      return;
    }
    const assetKey = selectedItem.assetId;
    setActiveMedia(null);
    setActiveMediaKey(null);
    void readMedia(assetKey).then((media) => {
      if (cancelled) return;
      if (media) {
        setActiveMedia(media.blob);
        setActiveMediaKey(assetKey);
      }
      else showNotice("O arquivo vinculado a este item não está mais no armazenamento local.");
    }).catch((error: unknown) => {
      console.error("Falha ao abrir o arquivo:", error);
      if (!cancelled) showNotice("Não foi possível abrir este arquivo local.");
    });
    return () => { cancelled = true; };
  }, [selectedItem?.assetId, showNotice]);

  useEffect(() => {
    if (!presenting || selectedItem?.kind !== "pptx" || !activeMedia || activeMediaKey !== selectedItem.assetId) return;
    const key = selectedItem.assetId ?? selectedItem.id;
    setProjectionPptx((current) => cacheProjectionPptx(current, key, selectedItem, activeMedia));
    setProjectionPptxKey(key);
  }, [activeMedia, activeMediaKey, presenting, selectedItem]);

  useEffect(() => {
    const onFullscreenChange = () => {
      if (!document.fullscreenElement) {
        setPresenting(false);
        setProjectionPptx({});
        setProjectionPptxKey(null);
        setLoadedProjectionPptxKey(null);
        setProjectionControlsVisible(false);
        pptxReadyKeyRef.current = null;
        loadedPptxKeyRef.current = null;
        pendingPptxNavigationRef.current = null;
        if (projectionControlsTimer.current != null) window.clearTimeout(projectionControlsTimer.current);
      }
    };
    document.addEventListener("fullscreenchange", onFullscreenChange);
    return () => document.removeEventListener("fullscreenchange", onFullscreenChange);
  }, []);

  const updateItem = useCallback((id: string, patch: Partial<PresentationItem>) => {
    setItems((current) => current.map((item) => item.id === id ? { ...item, ...patch } : item));
  }, []);

  const dispatchPptxKeys = useCallback((key: string, keys: string[]) => {
    const layer = Array.from(document.querySelectorAll<HTMLElement>(".stage-pptx-layer"))
      .find((element) => element.dataset.pptxKey === key);
    const target = layer?.querySelector<HTMLElement>(".pptx-viewer-frame");
    if (!target) return false;
    pptxKeyboardForwardedRef.current = true;
    try {
      for (const pressedKey of keys) {
        target.dispatchEvent(new KeyboardEvent("keydown", { key: pressedKey, bubbles: true, cancelable: true }));
      }
    } finally {
      pptxKeyboardForwardedRef.current = false;
    }
    return true;
  }, []);

  const flushPendingPptxNavigation = useCallback((key: string) => {
    const request = pendingPptxNavigationRef.current;
    if (!request || request.key !== key || pptxReadyKeyRef.current !== key) return;
    window.requestAnimationFrame(() => window.requestAnimationFrame(() => {
      if (pendingPptxNavigationRef.current !== request || pptxReadyKeyRef.current !== key) return;
      const delta = request.targetIndex - request.fromIndex;
      const keys = delta === 1 ? ["ArrowRight"]
        : delta === -1 ? ["ArrowLeft"]
          : [...String(request.targetIndex + 1), "Enter"];
      if (dispatchPptxKeys(key, keys)) pendingPptxNavigationRef.current = null;
    }));
  }, [dispatchPptxKeys]);

  const onPptxViewerReady = useCallback((key: string) => {
    loadedPptxKeyRef.current = key;
    setLoadedProjectionPptxKey(key);
    if (presenting) {
      pptxReadyKeyRef.current = key;
      flushPendingPptxNavigation(key);
    } else {
      pptxReadyKeyRef.current = null;
    }
  }, [flushPendingPptxNavigation, presenting]);

  const onPptxActiveSlideChange = useCallback((key: string, slideIndex: number) => {
    pptxSlideIndexRef.current[key] = slideIndex;
  }, []);

  const queuePptxNavigation = useCallback((key: string, fromIndex: number, targetIndex: number) => {
    if (fromIndex === targetIndex) {
      pendingPptxNavigationRef.current = null;
      return;
    }
    pendingPptxNavigationRef.current = { key, fromIndex, targetIndex };
    flushPendingPptxNavigation(key);
  }, [flushPendingPptxNavigation]);

  const navigateTo = useCallback((target: number) => {
    if (!items.length) return;
    const bounded = Math.max(0, Math.min(target, items.length - 1));
    const nextItem = items[bounded]!;
    if (presenting && selectedItem?.kind === "pptx" && activeMedia && activeMediaKey === selectedItem.assetId) {
      const key = selectedItem.assetId ?? selectedItem.id;
      setProjectionPptx((current) => cacheProjectionPptx(current, key, selectedItem, activeMedia));
      setProjectionPptxKey(key);
    }
    if (presenting && nextItem.kind === "pptx") {
      const key = nextItem.assetId ?? nextItem.id;
      const targetIndex = Math.max(0, (nextItem.slideNumber ?? 1) - 1);
      const currentIndex = pptxSlideIndexRef.current[key]
        ?? (projectionPptx[key] ? Math.max(0, (projectionPptx[key]!.item.slideNumber ?? 1) - 1) : targetIndex);
      const targetBlob = projectionPptx[key]?.blob ?? (activeMediaKey === key ? activeMedia : null);
      if (targetBlob) setProjectionPptx((current) => cacheProjectionPptx(current, key, nextItem, targetBlob));
      setProjectionPptxKey(key);
      if (loadedPptxKeyRef.current === key) {
        queuePptxNavigation(key, currentIndex, targetIndex);
      } else {
        pendingPptxNavigationRef.current = null;
        pptxReadyKeyRef.current = null;
      }
    }
    setSelectedId(nextItem.id);
  }, [activeMedia, activeMediaKey, items, presenting, projectionPptx, queuePptxNavigation, selectedItem]);

  const addParsedImports = useCallback(async (files: FileList | File[]) => {
    if (!files.length) return;
    const queue = Array.from(files);
    setBusy(true);
    try {
      const prepared: ParsedImport[] = [];
      for (const file of queue) {
        try {
          const imported = await prepareFile(file);
          await writeMedia({
            id: imported.assetId,
            name: imported.title,
            mimeType: imported.mimeType,
            size: imported.size,
            blob: imported.blob,
            storedAt: Date.now(),
          });
          prepared.push(imported);
        } catch (error) {
          const message = error instanceof Error ? error.message : `Não foi possível importar “${file.name}”.`;
          showNotice(message);
        }
      }
      const newItems = prepared.flatMap((entry) => entry.items);
      if (newItems.length) {
        setItems((current) => {
          const safeIndex = selectedId ? Math.max(0, current.findIndex((item) => item.id === selectedId) + 1) : current.length;
          return [...current.slice(0, safeIndex), ...newItems, ...current.slice(safeIndex)];
        });
        setSelectedId(newItems[0]!.id);
        const slides = newItems.length;
        showNotice(`${queue.length > 1 ? `${queue.length} arquivos importados` : `“${newItems[0]!.fileName}” importado`} · ${slides} ${slides === 1 ? "item" : "itens"} na sequência.`);
      }
    } finally {
      setBusy(false);
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  }, [items, selectedId, showNotice]);

  const insertLink = (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    try {
      const url = normalizeLink(linkValue);
      const link = createLinkItem(url, linkName);
      const safeIndex = selectedId ? Math.max(0, items.findIndex((item) => item.id === selectedId) + 1) : items.length;
      const next = [...items.slice(0, safeIndex), link, ...items.slice(safeIndex)];
      setItems(next);
      setSelectedId(link.id);
      setLinkDialog(false);
      setLinkValue("");
      setLinkName("");
      showNotice("Link inserido na sequência. Na projeção, use a seta para continuar.");
    } catch {
      showNotice("Confira o endereço do link e tente novamente.");
    }
  };

  const addDuplicate = useCallback((item: PresentationItem, afterId: string) => {
    const copy = createCopy(item);
    setItems((current) => {
      const at = current.findIndex((entry) => entry.id === afterId);
      const index = at < 0 ? current.length : at + 1;
      return [...current.slice(0, index), copy, ...current.slice(index)];
    });
    setSelectedId(copy.id);
    return copy;
  }, []);

  const moveItem = useCallback((fromId: string, toId: string, after = false) => {
    if (fromId === toId) return;
    setItems((current) => {
      const from = current.findIndex((item) => item.id === fromId);
      const to = current.findIndex((item) => item.id === toId);
      if (from < 0 || to < 0) return current;
      const reordered = [...current];
      const [moving] = reordered.splice(from, 1);
      const adjustedTarget = to > from ? to - 1 : to;
      reordered.splice(adjustedTarget + Number(after), 0, moving!);
      return reordered;
    });
  }, []);

  const startProjection = async () => {
    if (!items.length) {
      showNotice("Adicione pelo menos um arquivo ou link para começar a projeção.");
      return;
    }
    if (!selectedItem && items[0]) setSelectedId(items[0].id);
    try {
      const target = projectionRef.current ?? document.documentElement;
      if (!target.requestFullscreen) throw new Error("Tela cheia indisponível neste navegador.");
      await target.requestFullscreen({ navigationUI: "hide" });
      if (selectedItem?.kind === "pptx" && activeMedia && activeMediaKey === selectedItem.assetId) {
        const key = selectedItem.assetId ?? selectedItem.id;
        setProjectionPptx({ [key]: { item: selectedItem, blob: activeMedia } });
        setProjectionPptxKey(key);
      } else {
        setProjectionPptx({});
        setProjectionPptxKey(null);
      }
      pptxReadyKeyRef.current = null;
      pendingPptxNavigationRef.current = null;
      setProjectionControlsVisible(false);
      setPresenting(true);
    } catch {
      setPresenting(false);
      showNotice("A janela do navegador bloqueou a tela cheia. Tente F11 ou permita a tela cheia para este aplicativo.");
    }
  };

  const stopProjection = () => {
    if (document.fullscreenElement) void document.exitFullscreen();
    setProjectionPptx({});
    setProjectionPptxKey(null);
    setLoadedProjectionPptxKey(null);
    setProjectionControlsVisible(false);
    pptxReadyKeyRef.current = null;
    loadedPptxKeyRef.current = null;
    pendingPptxNavigationRef.current = null;
    if (projectionControlsTimer.current != null) window.clearTimeout(projectionControlsTimer.current);
    setPresenting(false);
  };

  const onKeyDown = useCallback((event: KeyboardEvent) => {
    if (pptxKeyboardForwardedRef.current) return;
    const target = event.target as HTMLElement | null;
    const editingText = !!target?.closest("input, textarea, select, [contenteditable='true']");

    if (presenting && ["ArrowRight", "ArrowDown", "PageDown", "ArrowLeft", "ArrowUp", "PageUp"].includes(event.key)) {
      event.preventDefault();
      event.stopImmediatePropagation();
      if (event.repeat) return;
      const forward = event.key === "ArrowRight" || event.key === "ArrowDown" || event.key === "PageDown";
      navigateTo(activeIndex + (forward ? 1 : -1));
      return;
    }
    if (presenting && /^[1-9]$/.test(event.key)) {
      event.preventDefault();
      event.stopImmediatePropagation();
      if (event.repeat) return;
      const destination = getFileStartIndices(items)[Number(event.key) - 1];
      if (destination != null) {
        navigateTo(destination);
        setMediaRestartToken((current) => current + 1);
      }
      return;
    }

    if (!editingText && (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c" && selectedItem) {
      copiedItem.current = selectedItem;
      return;
    }
    if (!editingText && (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "v" && copiedItem.current) {
      event.preventDefault();
      const anchor = selectedItem?.id ?? items[items.length - 1]?.id;
      if (anchor) addDuplicate(copiedItem.current, anchor);
      return;
    }
    if (!editingText && (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
      event.preventDefault();
      showNotice("A apresentação é salva automaticamente neste navegador.");
      return;
    }
    if (editingText && !presenting) return;

    if (presenting && event.key.toLowerCase() === "f") {
      event.preventDefault();
      event.stopImmediatePropagation();
      if (document.fullscreenElement) void document.exitFullscreen();
      else void projectionRef.current?.requestFullscreen?.({ navigationUI: "hide" });
      return;
    }

    if (!presenting && !editingText && event.key === "ArrowRight") {
      event.preventDefault();
      navigateTo(activeIndex + 1);
    } else if (!presenting && !editingText && event.key === "ArrowLeft") {
      event.preventDefault();
      navigateTo(activeIndex - 1);
    }
  }, [activeIndex, addDuplicate, items, navigateTo, presenting, selectedItem, showNotice]);

  useEffect(() => {
    window.addEventListener("keydown", onKeyDown, true);
    return () => window.removeEventListener("keydown", onKeyDown, true);
  }, [onKeyDown]);

  const updateClipTime = (field: "startAt" | "endAt", value: number) => {
    if (!selectedItem) return;
    const clamped = Math.max(0, value);
    if (field === "startAt") {
      const end = selectedItem.endAt ?? selectedItem.duration;
      const maximumStart = selectedItem.duration == null ? Number.POSITIVE_INFINITY : Math.max(0, selectedItem.duration - 0.001);
      const nextStart = Math.min(clamped, maximumStart);
      updateItem(selectedItem.id, { startAt: end && nextStart >= end ? Math.max(0, end - 0.001) : nextStart });
    } else {
      const start = selectedItem.startAt ?? 0;
      const maximumEnd = selectedItem.duration ?? Number.POSITIVE_INFINITY;
      const minimumEnd = Math.min(start + 0.001, maximumEnd);
      updateItem(selectedItem.id, { endAt: Math.max(minimumEnd, Math.min(clamped, maximumEnd)) });
    }
  };

  const removeSelected = () => {
    if (!selectedItem) return;
    const next = items.filter((item) => item.id !== selectedItem.id);
    setItems(next);
    setSelectedId(next[Math.min(activeIndex, next.length - 1)]?.id ?? null);
  };

  const titleStatus = !isLoaded ? "Abrindo apresentação local" : saveState === "saving" ? "Salvando" : saveState === "error" ? "Erro ao salvar" : "Salvo neste dispositivo";
  const activeMediaType = selectedItem?.kind === "audio" || selectedItem?.kind === "video";
  const inputId = "presentation-file-picker";

  return (
    <div className={`app-shell ${presenting ? "is-presenting" : ""}`} ref={projectionRef}>
      {!presenting && (
        <header className="topbar">
          <a className="brand" href="/" aria-label="Apresenta — início">
            <span className="brand-mark"><Presentation size={19} strokeWidth={2.1} /></span>
            <span className="brand-name">apresenta<span>.</span></span>
          </a>
          <div className="topbar-divider" />
          <label className="deck-name-wrap" htmlFor="deck-title">NOME DA APRESENTAÇÃO</label>
          <input id="deck-title" className="deck-title-input" value={title} maxLength={80} onChange={(event) => setTitle(event.target.value)} aria-label="Nome da apresentação" disabled={!isLoaded} />
          <div className={`save-indicator ${saveState}`}><span className="save-dot" /><span>{titleStatus}</span></div>
          <button className="button button-add" type="button" onClick={() => fileInputRef.current?.click()} disabled={!isLoaded || busy}>
            {busy ? <Loader2 className="spin" size={16} /> : <Plus size={17} />}
            Adicionar arquivos
          </button>
          <button className="button button-present" type="button" onClick={() => void startProjection()} disabled={!isLoaded || !items.length} title="Iniciar projeção em tela cheia">
            <Play size={16} fill="currentColor" /> Apresentar
          </button>
          <input ref={fileInputRef} id={inputId} className="visually-hidden" type="file" accept={fileAccept()} multiple disabled={!isLoaded || busy} onChange={(event) => { if (event.target.files) void addParsedImports(event.target.files); }} />
        </header>
      )}

      <main className={`editor-layout ${presenting ? "projection-layout" : ""}`}>
        {!presenting && (
          <aside className="sequence-panel" aria-label="Ordem da apresentação">
            <div className="sequence-heading">
              <div><span className="eyebrow">SEQUÊNCIA</span><h1>Seus arquivos</h1></div>
              <span className="count-badge">{String(items.length).padStart(2, "0")}</span>
            </div>
            <div className="sequence-actions">
              <button type="button" onClick={() => fileInputRef.current?.click()} disabled={!isLoaded || busy} className="action-add-file"><Plus size={16} /> Adicionar arquivo</button>
              <button type="button" onClick={() => setLinkDialog(true)} disabled={!isLoaded} className="action-add-link"><Link2 size={15} /> Inserir link</button>
            </div>

            <div className="sequence-scroll" aria-label="Itens ordenados">
              {items.length === 0 ? (
                <div className="sequence-empty">
                  <div className="sequence-empty-mark"><Layers3 size={20} /></div>
                  <strong>A sequência começa aqui</strong>
                  <p>Adicione arquivos ou links. Cada slide vira um item independente.</p>
                  <button type="button" onClick={() => fileInputRef.current?.click()} disabled={!isLoaded}><Upload size={15} /> Escolher arquivos</button>
                </div>
              ) : items.map((item, index) => (
                <button
                  key={item.id}
                  type="button"
                  draggable
                  className={`sequence-item ${item.id === selectedItem?.id ? "is-selected" : ""} ${draggedId === item.id ? "is-dragged" : ""}`}
                  onClick={() => setSelectedId(item.id)}
                  onDragStart={(event) => { setDraggedId(item.id); event.dataTransfer.effectAllowed = "move"; event.dataTransfer.setData("text/plain", item.id); }}
                  onDragEnd={() => setDraggedId(null)}
                  onDragOver={(event) => { event.preventDefault(); event.dataTransfer.dropEffect = "move"; }}
                  onDrop={(event) => {
                    event.preventDefault();
                    const from = event.dataTransfer.getData("text/plain") || draggedId;
                    const bounds = event.currentTarget.getBoundingClientRect();
                    const dropAfter = event.clientY >= bounds.top + bounds.height / 2;
                    if (from) moveItem(from, item.id, dropAfter);
                    setDraggedId(null);
                  }}
                  aria-label={`${index + 1}: ${item.title}, ${kindLabel(item.kind)}. Arraste para reorganizar.`}
                >
                  <span className="drag-grip" title="Arraste para reordenar"><GripVertical size={15} /></span>
                  <span className={`sequence-number ${item.id === selectedItem?.id ? "active" : ""}`}>{String(index + 1).padStart(2, "0")}</span>
                  <span className={`sequence-icon kind-${item.kind}`}><KindIcon kind={item.kind} /></span>
                  <span className="sequence-copy"><strong>{item.title}</strong><span>{kindLabel(item.kind)}{item.slideNumber ? ` · ${item.slideNumber}/${item.slideCount}` : ""}</span></span>
                  {item.kind === "video" && <span className="row-audio"><Volume2 size={13} /></span>}
                </button>
              ))}
            </div>

            <div className="local-card">
              <span className="local-card-icon"><Check size={15} strokeWidth={2.4} /></span>
              <span><strong>Seu conteúdo fica aqui</strong><small>Arquivos guardados só neste navegador.</small>{localSpace && <small className="storage-detail">{localSpace}</small>}</span>
            </div>
            <div className="panel-footer">ARQUIVOS NESTE DISPOSITIVO</div>
          </aside>
        )}

        <section className={`stage-panel ${presenting ? "presenter" : ""}`} aria-label={presenting ? "Projeção em tela cheia" : "Prévia da apresentação"}>
          {!presenting && (
            <div className="stage-toolbar">
              <div className="stage-title-group"><span className="eyebrow">PRÉVIA DA PROJEÇÃO</span><span className="preview-hint"><span className="hint-dot" /> O conteúdo mantém as proporções originais</span></div>
              <div className="stage-actions">
                <button className="icon-button" type="button" title="Inserir link depois do item selecionado" onClick={() => setLinkDialog(true)} disabled={!isLoaded} aria-label="Inserir link"><Link2 size={16} /></button>
                <button className="icon-button" type="button" title="Duplicar o item selecionado" onClick={() => selectedItem && addDuplicate(selectedItem, selectedItem.id)} disabled={!selectedItem} aria-label="Duplicar item"><Copy size={16} /></button>
                <button className="icon-button danger-on-hover" type="button" title="Excluir o item selecionado" onClick={removeSelected} disabled={!selectedItem} aria-label="Excluir item"><Trash2 size={16} /></button>
                <span className="toolbar-divider" />
                <button className="icon-button expand-button" type="button" title="Prévia em tela cheia" onClick={() => void startProjection()} disabled={!isLoaded || !items.length} aria-label="Entrar em tela cheia"><Maximize2 size={16} /></button>
              </div>
            </div>
          )}

          <div
            className={`stage-area ${presenting ? "stage-fullscreen" : ""} ${presenting && projectionControlsVisible ? "projection-controls-visible" : ""}`}
            onPointerMove={(event) => { if (presenting && event.pointerType === "mouse") revealProjectionControls(); }}
          >
            {selectedItem ? (
              <>
                <div className="stage-media-layer" aria-hidden={selectedItem.kind === "pptx"}>
                  <MediaSlide
                    item={selectedItem}
                    blob={activeMediaKey === selectedItem.assetId ? activeMedia : null}
                    presenting={presenting}
                    restartToken={mediaRestartToken}
                    onVolumeChange={(volume) => updateItem(selectedItem.id, { volume })}
                    onDuration={(duration) => updateItem(selectedItem.id, { duration })}
                    onVideoEnd={() => { /* O fim do clipe pausa aqui; avançar é sempre manual. */ }}
                  />
                </div>
                {pptxStageEntry && (
                  <div key="presentation-pptx-stage" data-pptx-key={pptxStageEntry.key} className={`stage-pptx-layer ${selectedItem.kind === "pptx" ? "is-current" : "is-parked"}`} aria-hidden={selectedItem.kind !== "pptx"}>
                    <Suspense fallback={<div className="render-message">Carregando o visualizador PowerPoint…</div>}>
                      <PptxViewerStage
                        item={pptxStageEntry.item}
                        blob={pptxStageEntry.blob}
                        presenting={presenting}
                        onViewerReady={onPptxViewerReady}
                        onActiveSlideChange={onPptxActiveSlideChange}
                      />
                    </Suspense>
                  </div>
                )}
              </>
            ) : (
              <div className="stage-welcome">
                <div className="welcome-card">
                  <div className="welcome-visual"><div className="screen-illustration"><span /><span /><span /></div><div className="screen-stand" /></div>
                  <span className="eyebrow">DO SEU JEITO, NA SUA ORDEM</span>
                  <h2>Monte a sequência.<br /><span>Entre e apresente.</span></h2>
                  <p>PDF, PowerPoint, imagens, vídeo, áudio e links — tudo junto, no lugar certo.</p>
                  <div className="welcome-actions">
                    <button className="button button-present" type="button" onClick={() => fileInputRef.current?.click()} disabled={!isLoaded}><Plus size={17} /> Adicionar primeiro arquivo</button>
                    <button className="welcome-link" type="button" onClick={() => setLinkDialog(true)} disabled={!isLoaded}><Link2 size={15} /> Ou começar com um link</button>
                  </div>
                </div>
              </div>
            )}
            {presenting && itemTotal > 0 && (
              <>
                {activeIndex > 0 && <button type="button" className={`projection-arrow projection-prev ${projectionControlsVisible ? "projection-ui-visible" : ""}`} tabIndex={projectionControlsVisible ? 0 : -1} aria-hidden={!projectionControlsVisible} aria-label="Item anterior — seta para esquerda" onClick={() => navigateTo(activeIndex - 1)}><ArrowLeft size={21} /></button>}
                {activeIndex < itemTotal - 1 && <button type="button" className={`projection-arrow projection-next ${projectionControlsVisible ? "projection-ui-visible" : ""}`} tabIndex={projectionControlsVisible ? 0 : -1} aria-hidden={!projectionControlsVisible} aria-label="Próximo item — seta para direita" onClick={() => navigateTo(activeIndex + 1)}><ArrowRight size={21} /></button>}
                <div className={`projection-hud ${projectionControlsVisible ? "projection-ui-visible" : ""}`} aria-hidden={!projectionControlsVisible}>
                  <div className="hud-label"><span>{selectedItem?.title}</span><span className="hud-counter">{selectedPosition} / {itemTotal}</span></div>
                </div>
                <button type="button" className={`projection-exit ${projectionControlsVisible ? "projection-ui-visible" : ""}`} tabIndex={projectionControlsVisible ? 0 : -1} aria-hidden={!projectionControlsVisible} onClick={stopProjection} aria-label="Sair da projeção"><X size={19} /><span>Sair</span></button>
              </>
            )}
          </div>

          {!presenting && (
            <div className="stage-transport">
              <div className="transport-left">
                <button type="button" className="transport-step" onClick={() => navigateTo(activeIndex - 1)} disabled={!items.length || activeIndex === 0} aria-label="Item anterior"><ArrowLeft size={17} /></button>
                <button type="button" className="transport-step" onClick={() => navigateTo(activeIndex + 1)} disabled={!items.length || activeIndex >= items.length - 1} aria-label="Próximo item"><ArrowRight size={17} /></button>
                <span className="transport-position">{items.length ? `${String(selectedPosition).padStart(2, "0")} / ${String(items.length).padStart(2, "0")}` : "00 / 00"}</span>
              </div>
              {selectedItem?.kind === "video" && (
                <div className="clip-inline"><Film size={14} /><span>Trecho</span><strong>{formatTime(selectedItem.startAt)} — {formatTime(selectedItem.endAt ?? selectedItem.duration)}</strong></div>
              )}
              {activeMediaType && <span className="transport-volume"><Volume2 size={15} /> Som deste {selectedItem?.kind === "audio" ? "áudio" : "vídeo"}</span>}
              <button className="button button-present stage-present" type="button" onClick={() => void startProjection()} disabled={!isLoaded || !items.length}><Play size={15} fill="currentColor" /> Projetar agora</button>
            </div>
          )}
        </section>

        {!presenting && (
          <aside className="inspector-panel" aria-label="Configurações do item">
            <div className="inspector-header"><div><span className="eyebrow">AJUSTES</span><h2>Item selecionado</h2></div><span className="inspector-position">{selectedItem ? `${String(selectedPosition).padStart(2, "0")} / ${String(itemTotal).padStart(2, "0")}` : "— / —"}</span></div>
            {selectedItem ? (
              <>
                <div className="inspector-file">
                  <span className={`inspector-file-icon kind-${selectedItem.kind}`}><KindIcon kind={selectedItem.kind} size={18} /></span>
                  <span><strong>{kindLabel(selectedItem.kind)}</strong><small>{selectedItem.fileName || selectedItem.linkUrl || "Conteúdo da sequência"}</small></span>
                  {selectedItem.slideNumber && <span className="slide-badge">{selectedItem.slideNumber}<span>/{selectedItem.slideCount}</span></span>}
                </div>
                <div className="inspector-field"><label htmlFor="item-title">Nome na sequência</label><input id="item-title" className="form-input" value={selectedItem.title} maxLength={100} onChange={(event) => updateItem(selectedItem.id, { title: event.target.value })} /></div>

                {(selectedItem.kind === "video" || selectedItem.kind === "audio") && (
                  <div className="inspector-section">
                    <div className="section-heading"><span className="section-icon sound"><Volume2 size={16} /></span><span><strong>Volume deste arquivo</strong><small>Independente dos outros itens</small></span></div>
                    <div className="volume-slider-line"><span className="volume-low">0</span><input aria-label={`Volume de ${selectedItem.title}`} type="range" min="0.05" max="1" step="0.01" value={selectedItem.volume ?? 0.8} onChange={(event) => updateItem(selectedItem.id, { volume: Number(event.target.value) })} /><strong>{Math.round((selectedItem.volume ?? 0.8) * 100)}%</strong></div>
                    <span className="helper-note">O áudio inicia ativo; nunca silenciamos o vídeo automaticamente.</span>
                  </div>
                )}

                {selectedItem.kind === "video" && (
                  <div className="inspector-section clip-section">
                    <div className="section-heading"><span className="section-icon clip"><Film size={16} /></span><span><strong>Início e fim do trecho</strong><small>O corte vale só para esta ocorrência</small></span></div>
                    <label className="time-field"><span>Começa em</span><span className="time-input-wrap"><input aria-label="Início do vídeo em segundos" type="number" step="0.001" min="0" max={selectedItem.duration} value={selectedItem.startAt ?? 0} disabled={!selectedItem.duration} onChange={(event) => updateClipTime("startAt", Number(event.target.value))} /><span>seg.</span></span></label>
                    <label className="time-field"><span>Termina em</span><span className="time-input-wrap"><input aria-label="Fim do vídeo em segundos" type="number" step="0.001" min={0.001} max={selectedItem.duration} value={selectedItem.endAt ?? selectedItem.duration ?? ""} placeholder="Duração total" disabled={!selectedItem.duration} onChange={(event) => { if (event.target.value === "") updateItem(selectedItem.id, { endAt: undefined }); else updateClipTime("endAt", Number(event.target.value)); }} /><span>seg.</span></span></label>
                    {!selectedItem.duration && <span className="helper-note">Carregando a duração do vídeo…</span>}
                    <div className="cut-callout"><Check size={14} /><span>Ao chegar ao fim do trecho, o vídeo pausa. O próximo item só entra quando você avançar.</span></div>
                  </div>
                )}

                {selectedItem.kind === "link" && (
                  <div className="inspector-section"><div className="section-heading"><span className="section-icon link"><Link2 size={16} /></span><span><strong>Endereço da web</strong><small>Avanço manual, como os outros itens</small></span></div><a className="link-value" href={selectedItem.linkUrl} target="_blank" rel="noreferrer"><span>{selectedItem.linkUrl}</span><ExternalLink size={14} /></a></div>
                )}

                {(selectedItem.kind === "pptx" || selectedItem.kind === "pdf") && (
                  <div className="inspector-section slide-info"><div className="section-heading"><span className="section-icon slide"><Layers3 size={16} /></span><span><strong>Slide independente</strong><small>Posicione e repita sem alterar o original</small></span></div><div className="slide-order-line"><span>Posição atual</span><strong>{String(selectedPosition).padStart(2, "0")} na sequência</strong></div></div>
                )}

                <div className="inspector-actions"><button type="button" onClick={() => addDuplicate(selectedItem, selectedItem.id)}><Copy size={15} /> Duplicar depois deste item</button><button type="button" className="remove-action" onClick={removeSelected}><Trash2 size={15} /> Remover da sequência</button></div>
              </>
            ) : (
              <div className="inspector-empty"><div className="inspector-empty-icon"><Layers3 size={20} /></div><strong>Sem item selecionado</strong><p>Insira um arquivo ou link e ajuste cada ocorrência separadamente.</p></div>
            )}

            <div className="keyboard-card"><span className="eyebrow">ATALHOS</span><div className="shortcut-line"><span>Avançar / voltar</span><span><kbd>←</kbd><kbd>→</kbd></span></div><div className="shortcut-line"><span>Ir para o item</span><span><kbd>1</kbd> a <kbd>9</kbd></span></div><div className="shortcut-line"><span>Copiar / colar item</span><span><kbd>Ctrl</kbd> <kbd>C</kbd> <kbd>V</kbd></span></div></div>
            <div className="inspector-footnote"><Volume2 size={14} /><span>Cada mídia guarda seu próprio volume.</span></div>
          </aside>
        )}
      </main>

      {!presenting && (
        <footer className="bottom-bar"><span><span className={`bottom-status ${saveState}`}><span className="save-dot" /> {!isLoaded ? "Abrindo apresentação local" : saveState === "error" ? "Falha ao salvar" : saveState === "saving" ? "Salvando alterações" : "Salvo neste navegador"}</span><span className="footer-dot">·</span> Sem cadastro <span className="footer-dot">·</span> Seus arquivos não são enviados</span><button type="button" onClick={() => setLinkDialog(true)} disabled={!isLoaded}><Link2 size={13} /> Inserir link na sequência</button></footer>
      )}

      {notice && <div className="toast-note" role="status"><Check size={16} /><span>{notice}</span><button type="button" onClick={() => setNotice("")} aria-label="Fechar aviso"><X size={15} /></button></div>}

      {linkDialog && (
        <div className="modal-backdrop" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget) setLinkDialog(false); }}>
          <section className="link-dialog" role="dialog" aria-modal="true" aria-labelledby="link-title">
            <div className="dialog-head"><span className="dialog-icon"><Link2 size={19} /></span><button type="button" className="dialog-close" onClick={() => setLinkDialog(false)} aria-label="Fechar"><X size={19} /></button></div>
            <span className="eyebrow">NOVO ITEM</span><h2 id="link-title">Insira um link</h2><p>Ele vai para a sequência logo depois do item selecionado.</p>
            <form onSubmit={insertLink}>
              <label htmlFor="link-address">Endereço da web</label><input autoFocus required id="link-address" className="form-input" type="url" placeholder="https://exemplo.com" value={linkValue} onChange={(event) => setLinkValue(event.target.value)} />
              <label htmlFor="link-name">Nome na apresentação <span>opcional</span></label><input id="link-name" className="form-input" maxLength={80} placeholder="Ex.: demonstração do produto" value={linkName} onChange={(event) => setLinkName(event.target.value)} />
              <div className="dialog-note"><Check size={14} /> O link é aberto na sequência; páginas que bloqueiam incorporação podem ser acessadas em outra aba.</div>
              <div className="dialog-actions"><button type="button" className="button button-secondary" onClick={() => setLinkDialog(false)}>Cancelar</button><button className="button button-present" type="submit"><Plus size={16} /> Inserir na sequência</button></div>
            </form>
          </section>
        </div>
      )}
    </div>
  );
}
