import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  ExternalLink,
  Maximize2,
  Pause,
  Play,
  Volume2,
} from "lucide-react";
import type { PresentationItem } from "@/lib/presentation-types";

type PdfLoadingTask = ReturnType<typeof import("pdfjs-dist").getDocument>;
const pdfLoadingTasks = new Map<string, PdfLoadingTask>();

function clock(seconds: number): string {
  if (!Number.isFinite(seconds) || seconds < 0) return "00:00";
  const whole = Math.floor(seconds);
  const hours = Math.floor(whole / 3600);
  const minutes = Math.floor((whole % 3600) / 60);
  const remainder = whole % 60;
  return hours
    ? `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}:${String(remainder).padStart(2, "0")}`
    : `${String(minutes).padStart(2, "0")}:${String(remainder).padStart(2, "0")}`;
}

interface PdfSlideProps {
  assetId: string;
  blob: Blob;
  pageNumber: number;
  alt: string;
}

function PdfSlide({ assetId, blob, pageNumber, alt }: PdfSlideProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const hostRef = useRef<HTMLDivElement>(null);
  const [error, setError] = useState(false);

  useEffect(() => {
    let cancelled = false;
    let cleanupRender: (() => void) | undefined;
    let observer: ResizeObserver | undefined;

    async function renderDocument() {
      try {
        const pdfjs = await import("pdfjs-dist");
        if (!pdfjs.GlobalWorkerOptions.workerSrc) {
          pdfjs.GlobalWorkerOptions.workerSrc = new URL("pdfjs-dist/build/pdf.worker.min.mjs", import.meta.url).toString();
        }
        let loadingTask = pdfLoadingTasks.get(assetId);
        if (!loadingTask) {
          loadingTask = pdfjs.getDocument({ data: await blob.arrayBuffer() });
          pdfLoadingTasks.set(assetId, loadingTask);
        }
        const loaded = await loadingTask.promise;
        if (cancelled) return;
        const page = await loaded.getPage(pageNumber);
        const canvas = canvasRef.current;
        const host = hostRef.current;
        if (!canvas || !host) return;
        const context = canvas.getContext("2d", { alpha: false });
        if (!context) throw new Error("Canvas não disponível.");

        let activeTask: { promise: Promise<void>; cancel: () => void } | undefined;
        let frame = 0;
        const paint = () => {
          if (cancelled) return;
          const bounds = host.getBoundingClientRect();
          if (bounds.width < 20 || bounds.height < 20) return;
          activeTask?.cancel();
          const base = page.getViewport({ scale: 1 });
          const scale = Math.min((bounds.width - 24) / base.width, (bounds.height - 24) / base.height);
          const ratio = Math.min(window.devicePixelRatio || 1, 2);
          const cssWidth = base.width * scale;
          const cssHeight = base.height * scale;
          const viewport = page.getViewport({ scale: scale * ratio });
          canvas.width = Math.ceil(viewport.width);
          canvas.height = Math.ceil(viewport.height);
          canvas.style.width = `${cssWidth}px`;
          canvas.style.height = `${cssHeight}px`;
          activeTask = page.render({
            canvas,
            canvasContext: context,
            viewport,
          });
          activeTask.promise.catch((renderError: { name?: string }) => {
            if (!cancelled && renderError.name !== "RenderingCancelledException") setError(true);
          });
        };
        cleanupRender = () => {
          cancelAnimationFrame(frame);
          activeTask?.cancel();
        };
        observer = new ResizeObserver(() => {
          cancelAnimationFrame(frame);
          frame = requestAnimationFrame(paint);
        });
        observer.observe(host);
        frame = requestAnimationFrame(paint);
      } catch (renderError) {
        console.error("Falha ao abrir a página do PDF:", renderError);
        if (!cancelled) setError(true);
      }
    }

    void renderDocument();
    return () => {
      cancelled = true;
      observer?.disconnect();
      cleanupRender?.();
    };
  }, [assetId, blob, pageNumber]);

  return (
    <div className="pdf-slide pdf-canvas-host" ref={hostRef} aria-label={alt}>
      <canvas ref={canvasRef} aria-label={alt} />
      {error && <p className="render-message">Não foi possível renderizar esta página. Verifique se o PDF não está protegido.</p>}
    </div>
  );
}

interface MediaSlideProps {
  item: PresentationItem;
  blob: Blob | null;
  presenting: boolean;
  restartToken: number;
  onVolumeChange: (volume: number) => void;
  onDuration: (duration: number) => void;
  onVideoEnd: () => void;
}

export default function MediaSlide({ item, blob, presenting, restartToken, onVolumeChange, onDuration, onVideoEnd }: MediaSlideProps) {
  const mediaRef = useRef<HTMLVideoElement | HTMLAudioElement>(null);
  const [playing, setPlaying] = useState(false);
  const [currentTime, setCurrentTime] = useState(0);
  const [duration, setDuration] = useState(item.duration ?? 0);
  const [playDenied, setPlayDenied] = useState(false);
  const [volume, setVolume] = useState(item.volume ?? 0.8);

  const mediaUrl = useMemo(() => {
    if (!blob || item.kind === "pptx" || item.kind === "pdf") return undefined;
    return URL.createObjectURL(blob);
  }, [blob, item.kind]);

  useEffect(() => () => {
    if (mediaUrl) URL.revokeObjectURL(mediaUrl);
  }, [mediaUrl]);

  useEffect(() => {
    setCurrentTime(0);
    setDuration(item.duration ?? 0);
    setPlaying(false);
    setPlayDenied(false);
    setVolume(item.volume ?? 0.8);
  }, [item.id]);

  useEffect(() => {
    const element = mediaRef.current;
    if (!element || item.kind !== "video" || !mediaUrl) return;
    const start = Math.max(item.startAt ?? 0, 0);
    if (Number.isFinite(element.duration) && Math.abs(element.currentTime - start) > 0.08) element.currentTime = start;
  }, [item.id, item.kind, item.startAt, mediaUrl]);

  useEffect(() => {
    const element = mediaRef.current;
    if (!element) return;
    const syncedVolume = Math.max(0.05, Math.min(item.volume ?? 0.8, 1));
    element.volume = syncedVolume;
    element.muted = false;
    setVolume(syncedVolume);
  }, [item.id, item.volume, mediaUrl]);

  const lastRestartToken = useRef(restartToken);
  useEffect(() => {
    if (lastRestartToken.current === restartToken) return;
    lastRestartToken.current = restartToken;
    if (item.kind !== "video" && item.kind !== "audio") return;
    const element = mediaRef.current;
    if (!element) return;
    const requestedStart = item.kind === "video" ? Math.max(item.startAt ?? 0, 0) : 0;
    const start = Number.isFinite(element.duration) && element.duration > 0
      ? Math.min(requestedStart, Math.max(element.duration - 0.01, 0))
      : requestedStart;
    element.currentTime = start;
    setCurrentTime(start);
    if (presenting) {
      void element.play().then(() => setPlayDenied(false)).catch(() => setPlayDenied(true));
    }
  }, [item.id, item.kind, item.startAt, presenting, restartToken]);

  useEffect(() => {
    const element = mediaRef.current;
    if (!element || !presenting || !mediaUrl) return;
    const attempt = element.play();
    if (attempt) void attempt.then(() => setPlayDenied(false)).catch(() => setPlayDenied(true));
  }, [presenting, item.id, mediaUrl]);

  useEffect(() => {
    if (presenting) return;
    const element = mediaRef.current;
    if (element && !element.paused) element.pause();
  }, [presenting, item.id]);

  const togglePlayback = useCallback(() => {
    const element = mediaRef.current;
    if (!element) return;
    if (element.paused) {
      void element.play().then(() => setPlayDenied(false)).catch(() => setPlayDenied(true));
    } else {
      element.pause();
    }
  }, []);

  const updateVolume = (value: number) => {
    const safeVolume = Math.min(1, Math.max(0.05, value));
    setVolume(safeVolume);
    if (mediaRef.current) {
      mediaRef.current.volume = safeVolume;
      mediaRef.current.muted = false;
    }
    onVolumeChange(safeVolume);
  };

  if (item.kind === "pptx") {
    return blob ? null : <div className="render-message">Abrindo o PowerPoint local…</div>;
  }

  if (item.kind === "pdf") {
    return blob ? <PdfSlide assetId={item.assetId ?? item.id} blob={blob} pageNumber={item.slideNumber ?? 1} alt={item.title} /> : <div className="render-message">Abrindo PDF local…</div>;
  }

  if (item.kind === "image") {
    return mediaUrl ? (
      <div className="image-slide"><img className="fit-media" src={mediaUrl} alt={item.title} draggable={false} /></div>
    ) : <div className="render-message">Abrindo imagem local…</div>;
  }

  if (item.kind === "link") {
    return (
      <div className="link-stage">
        <div className="link-browser-bar"><span className="browser-dots"><i /><i /><i /></span><span className="link-address">{item.linkUrl}</span></div>
        {item.linkUrl && (
          <iframe
            className="link-frame"
            src={item.linkUrl}
            title={item.title}
            loading="lazy"
            referrerPolicy="no-referrer"
            allow="autoplay; fullscreen; picture-in-picture"
            sandbox="allow-forms allow-scripts allow-popups allow-presentation"
          />
        )}
        <a className="link-open" href={item.linkUrl} target="_blank" rel="noreferrer">
          <ExternalLink size={15} /> Se a página bloquear a prévia, abra em outra aba
        </a>
      </div>
    );
  }

  if (item.kind !== "video" && item.kind !== "audio") return null;
  if (!mediaUrl || !blob) return <div className="render-message">Abrindo mídia local…</div>;

  const start = item.kind === "video" ? Math.max(item.startAt ?? 0, 0) : 0;
  const end = item.kind === "video" && item.endAt != null && item.endAt > start ? Math.min(item.endAt, duration || item.endAt) : duration;
  const progressMax = Math.max(end, start + 0.1);
  const displayTime = Math.max(start, Math.min(currentTime, progressMax));

  const onLoadedMetadata = (event: React.SyntheticEvent<HTMLMediaElement>) => {
    const element = event.currentTarget;
    const length = Number.isFinite(element.duration) ? element.duration : 0;
    setDuration(length);
    onDuration(length);
    element.volume = Math.max(0.05, Math.min(item.volume ?? 0.8, 1));
    element.muted = false;
    if (item.kind === "video" && start > 0) element.currentTime = Math.min(start, Math.max(length - 0.01, 0));
  };

  const mediaElement = item.kind === "video" ? (
    <video
      key={`${item.id}:${mediaUrl}`}
      ref={mediaRef as React.RefObject<HTMLVideoElement>}
      className="video-surface"
      src={mediaUrl}
      controls={false}
      playsInline
      autoPlay={presenting}
      muted={false}
      preload="auto"
      onLoadedMetadata={onLoadedMetadata}
      onPlay={() => { setPlaying(true); setPlayDenied(false); }}
      onPause={() => setPlaying(false)}
        onTimeUpdate={(event) => {
        const element = event.currentTarget;
        setCurrentTime(element.currentTime);
        if (item.endAt != null && item.endAt > start && element.currentTime >= item.endAt) {
          element.currentTime = item.endAt;
          element.pause();
          onVideoEnd();
        }
      }}
      onEnded={() => { setPlaying(false); onVideoEnd(); }}
      onVolumeChange={(event) => {
        if (event.currentTarget.muted || event.currentTarget.volume < 0.05) {
          event.currentTarget.muted = false;
          event.currentTarget.volume = Math.max(0.05, volume);
        }
      }}
    />
  ) : (
    <audio
      key={`${item.id}:${mediaUrl}`}
      ref={mediaRef as React.RefObject<HTMLAudioElement>}
      src={mediaUrl}
      controls={false}
      autoPlay={presenting}
      muted={false}
      preload="auto"
      onLoadedMetadata={onLoadedMetadata}
      onPlay={() => { setPlaying(true); setPlayDenied(false); }}
      onPause={() => setPlaying(false)}
      onTimeUpdate={(event) => setCurrentTime(event.currentTarget.currentTime)}
      onEnded={() => setPlaying(false)}
      onVolumeChange={(event) => {
        if (event.currentTarget.muted || event.currentTarget.volume < 0.05) {
          event.currentTarget.muted = false;
          event.currentTarget.volume = Math.max(0.05, volume);
        }
      }}
    />
  );

  return (
    <div className={`sound-stage ${item.kind === "audio" ? "audio-stage" : "video-stage"}`}>
      {mediaElement}
      {item.kind === "audio" && (
        <div className="audio-visual">
          <div className="audio-disc"><span /><span /><span /><span /><span /><span /><span /></div>
          <p>Áudio da apresentação</p>
          <span>O próximo item só aparece quando você avançar.</span>
        </div>
      )}
      <div className="media-controls" aria-label="Controles de mídia">
        <button type="button" className="media-play-button" onClick={togglePlayback} aria-label={playing ? "Pausar mídia" : "Reproduzir mídia"} title={playing ? "Pausar" : "Reproduzir"}>
          {playing ? <Pause size={17} fill="currentColor" /> : <Play size={17} fill="currentColor" />}
        </button>
        <span className="media-time">{clock(displayTime)}</span>
        <input
          aria-label="Posição da mídia"
          className="media-seek"
          type="range"
          min={Math.min(start, Math.max(progressMax - 0.1, 0))}
          max={progressMax}
          step="0.05"
          value={Math.min(displayTime, progressMax)}
          onChange={(event) => {
            const element = mediaRef.current;
            if (!element) return;
            element.currentTime = Number(event.target.value);
            setCurrentTime(element.currentTime);
          }}
        />
        <span className="media-time">{clock(end)}</span>
        <label className="volume-control" title="Volume deste arquivo">
          <Volume2 size={16} />
          <input
            aria-label={`Volume: ${item.title}`}
            type="range"
            min="0.05"
            max="1"
            step="0.01"
            value={volume}
            onChange={(event) => updateVolume(Number(event.target.value))}
          />
        </label>
        <button type="button" className="media-expand-button" onClick={() => void document.documentElement.requestFullscreen?.()} aria-label="Tela cheia" title="Tela cheia"><Maximize2 size={16} /></button>
      </div>
      {playDenied && <button type="button" className="play-to-start" onClick={togglePlayback}><Play size={16} fill="currentColor" /> Toque para iniciar com som</button>}
    </div>
  );
}
