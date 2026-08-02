import { createUuid } from "../../runtime/browserCompatibility";
import {
  normalizeFormulaEditorDocument,
  serializeFormulaEditorDocument,
} from "../shared/formulaEditorDocument";
import {
  createFormulaMetadata,
  type VisualTeXFormulaMetadata,
} from "../shared/formulaMetadata";
import { renderOfficeFormulaArtifacts } from "../shared/formulaRenderArtifacts";
import { latexLinesToOmmlArtifacts } from "../omml/latexToOmml";
import type {
  DocumentImportFormulaCommitItem,
} from "../documentImport/documentImportClient";
import type { WordLatexRedrawSpan } from "./wordLatexRedrawParser";

const MAX_WORD_REFERENCE_WIDTH_PT = 500;
const WORD_IMAGE_VISUAL_SCALE = 1.1;

export type WordLatexRedrawOutputKind = "omml" | "image";
export type WordLatexRedrawRenderTarget = WordLatexRedrawSpan & {
  fontSizePt: number;
};

type RenderTemplate = {
  canonicalLatex: string;
  ommlBase64: string;
  ommlDocxBase64: string;
  svgBase64?: string;
  pngBase64?: string;
  width?: number;
  height?: number;
  baseline?: number;
  renderWidthPx?: number;
  renderHeightPx?: number;
  referenceWidthPt?: number;
  referenceHeightPt?: number;
  referenceBaselinePt?: number;
};

function decodeUrlSafeBase64Utf8(value: string) {
  const normalized = value.replace(/-/g, "+").replace(/_/g, "/");
  const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=");
  const binary = atob(padded);
  const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
  return new TextDecoder().decode(bytes);
}

function ommlRetainsLiteralLatexCommand(ommlBase64: string, latex: string) {
  const commands = [...latex.matchAll(/\\([A-Za-z@]+)\b/g)]
    .map((match) => match[1])
    .filter((command, index, values) => values.indexOf(command) === index);
  if (!commands.length) return false;
  const omml = decodeUrlSafeBase64Utf8(ommlBase64);
  return commands.some((command) => omml.includes(`\\${command}`));
}

function calculateReferenceGeometry(widthPx: number, heightPx: number, baselinePx: number) {
  const naturalWidthPt = widthPx * 0.75 * WORD_IMAGE_VISUAL_SCALE;
  const naturalHeightPt = heightPx * 0.75 * WORD_IMAGE_VISUAL_SCALE;
  const scale = Math.min(1, MAX_WORD_REFERENCE_WIDTH_PT / naturalWidthPt);
  const referenceWidthPt = naturalWidthPt * scale;
  const referenceHeightPt = naturalHeightPt * scale;
  const descentRatio = Math.max(0, Math.min(1, (heightPx - baselinePx) / heightPx));
  const referenceBaselinePt = -Math.max(0, referenceHeightPt * descentRatio);
  return { referenceWidthPt, referenceHeightPt, referenceBaselinePt };
}

async function renderTemplate(
  span: WordLatexRedrawRenderTarget,
  outputKind: WordLatexRedrawOutputKind,
): Promise<RenderTemplate> {
  const line = { id: createUuid(), latex: span.latex };
  const editorDocument = normalizeFormulaEditorDocument([line], "raw");
  const canonicalLatex = serializeFormulaEditorDocument(editorDocument);

  if (outputKind === "omml") {
    const omml = latexLinesToOmmlArtifacts(
      editorDocument.lines.map((formulaLine) => formulaLine.latex),
      span.displayMode,
      editorDocument.codeFormat,
    );
    if (ommlRetainsLiteralLatexCommand(omml.ommlBase64, canonicalLatex)) {
      throw new Error("The redraw formula contains a command unsupported by Word OMML.");
    }
    return {
      canonicalLatex,
      ommlBase64: omml.ommlBase64,
      ommlDocxBase64: omml.ommlDocxBase64,
    };
  }

  const artifacts = renderOfficeFormulaArtifacts({
    lines: editorDocument.lines,
    codeFormat: editorDocument.codeFormat,
    displayMode: span.displayMode,
    host: "word",
  });
  if (!artifacts.omml) {
    throw new Error("Unable to generate Word OMML for the LaTeX redraw formula.");
  }
  const { omml, svg } = artifacts;
  if (ommlRetainsLiteralLatexCommand(omml.ommlBase64, canonicalLatex)) {
    throw new Error("The redraw formula contains a command unsupported by Word OMML.");
  }

  const { svgToPng } = await import("../../export/svgToPng");
  const pngBase64 = (
    await svgToPng(svg, { scale: 2, background: "transparent" })
  ).base64;
  const baseline = svg.baseline ?? svg.height;
  return {
    canonicalLatex,
    ommlBase64: omml.ommlBase64,
    ommlDocxBase64: omml.ommlDocxBase64,
    svgBase64: svg.base64,
    pngBase64,
    width: svg.width,
    height: svg.height,
    baseline,
    renderWidthPx: svg.width,
    renderHeightPx: svg.height,
    ...calculateReferenceGeometry(svg.width, svg.height, baseline),
  };
}

function createMetadata(
  formulaId: string,
  span: WordLatexRedrawRenderTarget,
  template: RenderTemplate,
): VisualTeXFormulaMetadata {
  const line = { id: createUuid(), latex: template.canonicalLatex };
  return createFormulaMetadata({
    formulaId,
    title:
      span.displayMode === "inline"
        ? "Redrawn inline Word formula"
        : "Redrawn display Word formula",
    lines: [line],
    codeFormat: "raw",
    sourceLatex: template.canonicalLatex,
    displayMode: span.displayMode,
    numbered: false,
    fontSizePt: span.fontSizePt,
    renderWidthPx: template.renderWidthPx,
    renderHeightPx: template.renderHeightPx,
    referenceWidthPt: template.referenceWidthPt,
    referenceHeightPt: template.referenceHeightPt,
    referenceBaselinePt: template.referenceBaselinePt,
  });
}

/**
 * Mirrors the Windows redraw renderer: scan targets are rendered directly,
 * templates are cached by output/display/font/source, and each Word target gets
 * an independent formula identity. No document-import blocks or preview UI are
 * constructed on this path.
 */
export async function prepareWindowsStyleWordLatexRedrawItems(
  spans: WordLatexRedrawRenderTarget[],
  outputKind: WordLatexRedrawOutputKind,
  onProgress?: (current: number, total: number) => void,
): Promise<DocumentImportFormulaCommitItem[]> {
  const templates = new Map<string, RenderTemplate>();
  const items: DocumentImportFormulaCommitItem[] = [];

  for (let index = 0; index < spans.length; index += 1) {
    const span = spans[index];
    const key = [
      outputKind,
      span.displayMode,
      String(span.fontSizePt),
      span.latex,
    ].join("\x1f");
    let template = templates.get(key);
    if (!template) {
      template = await renderTemplate(span, outputKind);
      templates.set(key, template);
    }

    const formulaId = createUuid();
    const metadata = createMetadata(formulaId, span, template);
    items.push({
      kind: "formula",
      formulaId,
      latex: template.canonicalLatex,
      displayMode: span.displayMode,
      numbered: false,
      fontSizePt: span.fontSizePt,
      metadata,
      ommlBase64: template.ommlBase64,
      ommlDocxBase64: template.ommlDocxBase64,
      svgBase64: template.svgBase64,
      pngBase64: template.pngBase64,
      width: template.width,
      height: template.height,
      baseline: template.baseline,
      sourceStart: span.start,
      sourceEnd: span.end,
      sourceText: span.sourceText,
    });
    onProgress?.(index + 1, spans.length);
  }

  return items;
}
