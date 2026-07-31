import {
  formatLatexLines,
  isLatexCodeFormat,
  parseLatexSource,
} from "../../clipboard/LatexCopyService";
import { createUuid } from "../../runtime/browserCompatibility";
import type { LatexCodeFormat } from "../../types/formula";

export interface FormulaEditorLine {
  id: string;
  latex: string;
}

export interface FormulaEditorDocument {
  lines: FormulaEditorLine[];
  codeFormat: LatexCodeFormat;
}

export function serializeFormulaEditorDocument(document: FormulaEditorDocument) {
  return formatLatexLines(
    document.lines.map((line) => line.latex),
    document.codeFormat,
  );
}

interface DetectedFormulaEnvironment {
  codeFormat: LatexCodeFormat;
  source: string;
}

function escapeRegExp(value: string) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function replaceEnvironment(
  source: string,
  original: string,
  replacement: string,
) {
  const escaped = escapeRegExp(original);
  return source
    .replace(
      new RegExp(
        `\\\\begin\\s*\\{${escaped}\\}(?:\\s*\\{[^{}]*\\})?`,
      ),
      `\\begin{${replacement}}`,
    )
    .replace(
      new RegExp(`\\\\end\\s*\\{${escaped}\\}`),
      `\\end{${replacement}}`,
    );
}

function detectFormulaEnvironment(source: string): DetectedFormulaEnvironment | null {
  const normalized = source.replace(/\r\n?/g, "\n").trim();
  if (!normalized) return null;

  const equation = normalized.match(
    /^\\begin\s*\{(equation\*?)\}([\s\S]*)\\end\s*\{\1\}$/,
  );
  if (
    equation &&
    /\\begin\s*\{split\}[\s\S]*\\end\s*\{split\}/.test(equation[2])
  ) {
    return {
      codeFormat:
        equation[1] === "equation*"
          ? "equation-star-split"
          : "equation-split",
      source: normalized,
    };
  }

  const alignedDisplay = normalized.match(
    /^\\\[\s*(\\begin\s*\{aligned\}[\s\S]*\\end\s*\{aligned\})\s*\\\]$/,
  );
  if (alignedDisplay) {
    return { codeFormat: "aligned", source: alignedDisplay[1] };
  }

  const environment = normalized.match(
    /^\\begin\s*\{(align\*?|alignat\*?|aligned|gather\*?|multline\*?|equation\*?|displaymath)\}(?:\s*\{[^{}]*\})?[\s\S]*\\end\s*\{\1\}$/,
  )?.[1];
  if (!environment) return null;

  switch (environment) {
    case "align":
      return { codeFormat: "align", source: normalized };
    case "align*":
      return { codeFormat: "align-star", source: normalized };
    case "alignat":
      return {
        codeFormat: "align",
        source: replaceEnvironment(normalized, "alignat", "align"),
      };
    case "alignat*":
      return {
        codeFormat: "align-star",
        source: replaceEnvironment(normalized, "alignat*", "align*"),
      };
    case "aligned":
      return { codeFormat: "aligned", source: normalized };
    case "gather":
      return { codeFormat: "gather", source: normalized };
    case "gather*":
      return { codeFormat: "gather-star", source: normalized };
    case "multline":
      return { codeFormat: "multline", source: normalized };
    case "multline*":
      return { codeFormat: "multline-star", source: normalized };
    case "equation":
      return { codeFormat: "equation", source: normalized };
    case "equation*":
    case "displaymath":
      return {
        codeFormat: "equation-star",
        source:
          environment === "displaymath"
            ? replaceEnvironment(normalized, "displaymath", "equation*")
            : normalized,
      };
    default:
      return null;
  }
}

export function normalizeFormulaEditorDocument(
  lines: FormulaEditorLine[],
  codeFormat: unknown,
): FormulaEditorDocument {
  const fallbackFormat: LatexCodeFormat = isLatexCodeFormat(codeFormat)
    ? codeFormat
    : "raw";
  const sourceLines = Array.isArray(lines) ? lines : [];
  const safeLines = sourceLines.length
    ? sourceLines.map((line) => ({
        id: line.id || createUuid(),
        latex: typeof line.latex === "string" ? line.latex : "",
      }))
    : [{ id: createUuid(), latex: "" }];

  if (safeLines.length !== 1) {
    return { lines: safeLines, codeFormat: fallbackFormat };
  }

  const detected = detectFormulaEnvironment(safeLines[0].latex);
  if (!detected) {
    return { lines: safeLines, codeFormat: fallbackFormat };
  }

  const parsed = parseLatexSource(detected.source, detected.codeFormat)
    .map((latex) => latex.trim())
    .filter(Boolean);
  if (!parsed.length) {
    return { lines: safeLines, codeFormat: fallbackFormat };
  }

  return {
    codeFormat: detected.codeFormat,
    lines: parsed.map((latex, index) => ({
      id: index === 0 ? safeLines[0].id : createUuid(),
      latex,
    })),
  };
}
