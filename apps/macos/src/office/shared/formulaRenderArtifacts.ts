import { latexToSvg } from "../../export/latexToSvg";
import type { LatexCodeFormat } from "../../types/formula";
import {
  latexLinesToOmmlArtifacts,
  type OmmlArtifacts,
} from "../omml/latexToOmml";
import {
  normalizeFormulaEditorDocument,
  serializeFormulaEditorDocument,
  type FormulaEditorLine,
} from "./formulaEditorDocument";

export const OFFICE_FORMULA_REFERENCE_FONT_SIZE_PT = 14;

export interface RenderOfficeFormulaArtifactsInput {
  lines: FormulaEditorLine[];
  codeFormat: LatexCodeFormat;
  displayMode: "inline" | "block";
  includeWordOmml?: boolean;
}

export interface OfficeFormulaRenderArtifacts {
  lines: FormulaEditorLine[];
  codeFormat: LatexCodeFormat;
  canonicalLatex: string;
  svg: ReturnType<typeof latexToSvg>;
  omml: OmmlArtifacts | null;
}

/**
 * Builds the canonical source and every vector/native artifact from the same
 * normalized editor document. Document import and edit replacement must use
 * this path so an align/align* formula cannot render differently after edit.
 */
export function renderOfficeFormulaArtifacts({
  lines,
  codeFormat,
  displayMode,
  includeWordOmml = true,
}: RenderOfficeFormulaArtifactsInput): OfficeFormulaRenderArtifacts {
  const document = normalizeFormulaEditorDocument(lines, codeFormat);
  const canonicalLatex = serializeFormulaEditorDocument(document);
  if (!canonicalLatex.trim()) {
    throw new Error("Cannot render an empty Office formula.");
  }
  const svg = latexToSvg(canonicalLatex, {
    displayMode: displayMode === "block",
    fontSizePt: OFFICE_FORMULA_REFERENCE_FONT_SIZE_PT,
    paddingPx: displayMode === "inline" ? 1 : 10,
    background: "transparent",
  });
  const omml = includeWordOmml
    ? latexLinesToOmmlArtifacts(
        document.lines.map((line) => line.latex),
        displayMode,
        document.codeFormat,
      )
    : null;
  return {
    lines: document.lines,
    codeFormat: document.codeFormat,
    canonicalLatex,
    svg,
    omml,
  };
}
