export interface SvgExportOptions {
  displayMode: boolean;
  fontSizePt: number;
  paddingPx: number;
  background: "transparent" | "white";
  /**
   * Word for Mac can defer resolving inherited/currentColor SVG paint until the
   * drawing is selected. Force every SVG paint carrier used by a Word formula
   * to an explicit black value so the vector and its PNG preview share the same
   * first-frame artwork.
   */
  forceExplicitBlack?: boolean;
}

export interface SvgExportResult {
  svg: string;
  base64: string;
  width: number;
  height: number;
  baseline?: number;
}

export interface PngExportOptions {
  scale?: number;
  background?: "transparent" | "white";
}

export interface PngExportResult {
  blob: Blob;
  base64: string;
  width: number;
  height: number;
}
