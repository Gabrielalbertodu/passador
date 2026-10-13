export type PresentationKind = "pptx" | "pdf" | "image" | "video" | "audio" | "link";

export interface PresentationItem {
  id: string;
  kind: PresentationKind;
  title: string;
  fileName?: string;
  assetId?: string;
  mimeType?: string;
  slideNumber?: number;
  slideCount?: number;
  slideWidth?: number;
  slideHeight?: number;
  previewSvg?: string;
  linkUrl?: string;
  volume?: number;
  startAt?: number;
  endAt?: number;
  duration?: number;
}
