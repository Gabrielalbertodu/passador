import { Component, useEffect, useRef, useState, type ReactNode } from "react";
import { PowerPointViewer, type PowerPointViewerHandle } from "pptx-react-viewer";
import { I18nextProvider } from "react-i18next";
import "pptx-react-viewer/styles";
import type { PresentationItem } from "@/lib/presentation-types";
import pptxI18n from "@/lib/pptx-i18n";

const fileByteCache = new WeakMap<Blob, Promise<Uint8Array>>();

function getFileBytes(blob: Blob): Promise<Uint8Array> {
  let cached = fileByteCache.get(blob);
  if (!cached) {
    cached = blob.arrayBuffer().then((buffer) => new Uint8Array(buffer));
    fileByteCache.set(blob, cached);
  }
  return cached;
}

interface PptxViewerStageProps {
  item: PresentationItem;
  blob: Blob;
  presenting: boolean;
  onViewerReady: (assetKey: string) => void;
  onActiveSlideChange: (assetKey: string, slideIndex: number) => void;
}

class ViewerErrorBoundary extends Component<{ children: ReactNode; resetKey: string }, { failed: boolean }> {
  state = { failed: false };

  static getDerivedStateFromError() {
    return { failed: true };
  }

  componentDidCatch(error: Error) {
    console.error("Falha ao renderizar o slide PowerPoint:", error);
  }

  componentDidUpdate(previousProps: Readonly<{ children: ReactNode; resetKey: string }>) {
    if (previousProps.resetKey !== this.props.resetKey && this.state.failed) {
      this.setState({ failed: false });
    }
  }

  render() {
    if (this.state.failed) {
      return <div className="render-message">Não foi possível renderizar este slide PowerPoint.</div>;
    }
    return this.props.children;
  }
}

export default function PptxViewerStage({ item, blob, presenting, onViewerReady, onActiveSlideChange }: PptxViewerStageProps) {
  const viewerRef = useRef<PowerPointViewerHandle>(null);
  const [content, setContent] = useState<{ assetKey: string; bytes: Uint8Array } | null>(null);
  const [loadError, setLoadError] = useState(false);
  const loadedAssetKey = item.assetId ?? item.id;
  const viewerAssetKeyRef = useRef<string | null>(null);
  const slideIndex = Math.max(0, (item.slideNumber ?? 1) - 1);
  useEffect(() => {
    let cancelled = false;
    setLoadError(false);
    void getFileBytes(blob).then((bytes) => {
      if (!cancelled) setContent({ assetKey: loadedAssetKey, bytes });
    }).catch((error: unknown) => {
      console.error("Falha ao ler a apresentação PowerPoint local:", error);
      if (!cancelled) setLoadError(true);
    });
    return () => { cancelled = true; };
  }, [blob, loadedAssetKey]);

  useEffect(() => {
    if (!content || content.assetKey !== loadedAssetKey || !viewerRef.current) return;
    const viewer = viewerRef.current;
    const mode = presenting ? "present" : "preview";
    const changedDeck = viewerAssetKeyRef.current !== loadedAssetKey;
    if (viewer.getMode() !== mode) viewer.setMode(mode);
    if ((!presenting || changedDeck) && viewer.getActiveSlideIndex() !== slideIndex) viewer.goTo(slideIndex);
    viewerAssetKeyRef.current = loadedAssetKey;
    let secondFrame = 0;
    const firstFrame = window.requestAnimationFrame(() => {
      secondFrame = window.requestAnimationFrame(() => onViewerReady(loadedAssetKey));
    });
    return () => {
      window.cancelAnimationFrame(firstFrame);
      if (secondFrame) window.cancelAnimationFrame(secondFrame);
    };
  }, [content, loadedAssetKey, onViewerReady, presenting, slideIndex]);

  return (
    <div className="pptx-viewer-frame" aria-label={`Slide ${item.slideNumber ?? 1}: ${item.title}`}>
      {content ? (
        <ViewerErrorBoundary resetKey={`${loadedAssetKey}:${slideIndex}`}>
          <I18nextProvider i18n={pptxI18n}>
            <PowerPointViewer
              ref={viewerRef}
              content={content.bytes}
              initialSlide={content.assetKey === loadedAssetKey ? slideIndex : 0}
              showToolbar={false}
              showThumbnails={false}
              showCompatibilityToasts={false}
              canEdit={false}
              className="apresenta-pptx-viewer"
              fitPadding={0}
              maxFitScale={null}
              onActiveSlideChange={(index) => onActiveSlideChange(loadedAssetKey, index)}
            />
          </I18nextProvider>
        </ViewerErrorBoundary>
      ) : loadError ? (
        <div className="render-message">Não foi possível abrir esta apresentação PowerPoint.</div>
      ) : (
        <div className="render-message">Abrindo o PowerPoint e preparando o slide…</div>
      )}
      {content && loadError && <div role="alert" className="render-message">Não foi possível abrir este arquivo PowerPoint.</div>}
    </div>
  );
}
