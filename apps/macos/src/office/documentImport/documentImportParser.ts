import { createUuid } from "../../runtime/browserCompatibility";

export type DocumentImportSourceKind = "auto" | "markdown" | "latex";
export type DocumentFormulaDisplayMode = "inline" | "block";
export type DocumentFormulaOutputKind = "omml" | "image";

export interface DocumentTextBlock {
  id: string;
  kind: "text";
  text: string;
}

export interface DocumentFormulaBlock {
  id: string;
  kind: "formula";
  latex: string;
  displayMode: DocumentFormulaDisplayMode;
  numbered: boolean;
  fontSizePt: number;
}

export type DocumentImportBlock = DocumentTextBlock | DocumentFormulaBlock;

const blockEnvironmentNames = new Set([
  "equation",
  "equation*",
  "align",
  "align*",
  "alignat",
  "alignat*",
  "gather",
  "gather*",
  "multline",
  "multline*",
  "displaymath",
]);

function isEscaped(source: string, index: number) {
  let slashCount = 0;
  for (let cursor = index - 1; cursor >= 0 && source[cursor] === "\\"; cursor -= 1) {
    slashCount += 1;
  }
  return slashCount % 2 === 1;
}

function stripLatexComments(source: string) {
  return source
    .split("\n")
    .map((line) => {
      for (let index = 0; index < line.length; index += 1) {
        if (line[index] === "%" && !isEscaped(line, index)) return line.slice(0, index);
      }
      return line;
    })
    .join("\n");
}

function unwrapSimpleLatexCommands(source: string) {
  let value = source;
  const oneArgumentCommands = [
    "textbf",
    "textit",
    "emph",
    "underline",
    "texttt",
    "textrm",
    "textsf",
    "textnormal",
    "mbox",
    "caption",
  ];
  const commandPattern = new RegExp(
    `\\\\(?:${oneArgumentCommands.join("|")})\\s*\\{([^{}]*)\\}`,
    "g",
  );
  for (let pass = 0; pass < 8; pass += 1) {
    const next = value.replace(commandPattern, "$1");
    if (next === value) break;
    value = next;
  }
  value = value.replace(/\\href\s*\{[^{}]*\}\s*\{([^{}]*)\}/g, "$1");
  return value;
}

function plainTextFromMarkup(raw: string, sourceKind: DocumentImportSourceKind) {
  let value = raw.replace(/\r\n?/g, "\n");
  if (sourceKind !== "markdown") {
    value = stripLatexComments(value);
    value = value
      .replace(/\\documentclass(?:\[[^\]]*\])?\s*\{[^{}]*\}/g, "")
      .replace(/\\usepackage(?:\[[^\]]*\])?\s*\{[^{}]*\}/g, "")
      .replace(/\\(?:title|author|date)\s*\{([^{}]*)\}/g, "$1\n")
      .replace(/\\maketitle\b/g, "")
      .replace(/\\begin\s*\{document\}|\\end\s*\{document\}/g, "")
      .replace(
        /\\(?:part|chapter|section|subsection|subsubsection|paragraph|subparagraph)\*?\s*\{([^{}]*)\}/g,
        "\n$1\n",
      )
      .replace(/\\item(?:\[[^\]]*\])?\s*/g, "\n• ")
      .replace(/\\begin\s*\{(?:itemize|enumerate|description|center|flushleft|flushright|quote|quotation)\}/g, "\n")
      .replace(/\\end\s*\{(?:itemize|enumerate|description|center|flushleft|flushright|quote|quotation)\}/g, "\n")
      .replace(/\\(?:newline|linebreak|par)\b/g, "\n")
      .replace(/\\\\(?:\[[^\]]*\])?/g, "\n");
    value = unwrapSimpleLatexCommands(value);
    value = value
      .replace(/\\%/g, "%")
      .replace(/\\&/g, "&")
      .replace(/\\#/g, "#")
      .replace(/\\_/g, "_")
      .replace(/\\\$/g, "$")
      .replace(/\\\{/g, "{")
      .replace(/\\\}/g, "}")
      .replace(/~/g, " ")
      .replace(/\\[a-zA-Z@]+\*?(?:\[[^\]]*\])?/g, "");
  }

  value = value
    .replace(/^\s{0,3}#{1,6}\s+/gm, "")
    .replace(/^\s*>\s?/gm, "")
    .replace(/^\s*[-+*]\s+/gm, "• ")
    .replace(/^\s*(\d+)\.\s+/gm, "$1. ")
    .replace(/!\[([^\]]*)\]\([^)]*\)/g, "$1")
    .replace(/\[([^\]]+)\]\([^)]*\)/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    .replace(/\*\*([^*]+)\*\*/g, "$1")
    .replace(/__([^_]+)__/g, "$1")
    .replace(/(?<!\*)\*([^*\n]+)\*(?!\*)/g, "$1")
    .replace(/(?<!_)_([^_\n]+)_(?!_)/g, "$1")
    .replace(/~~([^~]+)~~/g, "$1")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n[ \t]+/g, "\n")
    .replace(/\n{3,}/g, "\n\n");

  return value;
}

interface MathDelimiter {
  opening: string;
  closing: string;
  displayMode: DocumentFormulaDisplayMode;
  numbered: boolean;
  contentStart: number;
  environment?: string;
}

function delimiterAt(source: string, index: number): MathDelimiter | null {
  if (source.startsWith("$$", index) && !isEscaped(source, index)) {
    return {
      opening: "$$",
      closing: "$$",
      displayMode: "block",
      numbered: false,
      contentStart: index + 2,
    };
  }
  if (source.startsWith("\\[", index)) {
    return {
      opening: "\\[",
      closing: "\\]",
      displayMode: "block",
      numbered: false,
      contentStart: index + 2,
    };
  }
  if (source.startsWith("\\(", index)) {
    return {
      opening: "\\(",
      closing: "\\)",
      displayMode: "inline",
      numbered: false,
      contentStart: index + 2,
    };
  }
  if (source[index] === "$" && !isEscaped(source, index)) {
    return {
      opening: "$",
      closing: "$",
      displayMode: "inline",
      numbered: false,
      contentStart: index + 1,
    };
  }
  if (source.startsWith("\\begin", index)) {
    const match = source.slice(index).match(/^\\begin\s*\{([^{}]+)\}/);
    const environment = match?.[1];
    if (match && environment && blockEnvironmentNames.has(environment)) {
      return {
        opening: match[0],
        closing: `\\end{${environment}}`,
        displayMode: "block",
        numbered: !environment.endsWith("*") && environment !== "displaymath",
        contentStart: index + match[0].length,
        environment,
      };
    }
  }
  return null;
}

function findClosingDelimiter(source: string, delimiter: MathDelimiter) {
  let cursor = delimiter.contentStart;
  if (delimiter.environment) {
    const flexiblePattern = new RegExp(
      `\\\\end\\s*\\{${delimiter.environment.replace(/\*/g, "\\*")}\\}`,
    );
    const remainder = source.slice(cursor);
    const match = remainder.match(flexiblePattern);
    return match?.index === undefined ? -1 : cursor + match.index;
  }
  while (cursor < source.length) {
    const found = source.indexOf(delimiter.closing, cursor);
    if (found < 0) return -1;
    if (delimiter.closing === "$" || delimiter.closing === "$$") {
      if (isEscaped(source, found)) {
        cursor = found + delimiter.closing.length;
        continue;
      }
      if (delimiter.closing === "$" && source.startsWith("$$", found)) {
        cursor = found + 2;
        continue;
      }
    }
    return found;
  }
  return -1;
}

function appendTextBlock(
  blocks: DocumentImportBlock[],
  raw: string,
  sourceKind: DocumentImportSourceKind,
) {
  const text = plainTextFromMarkup(raw, sourceKind);
  if (!text) return;
  const previous = blocks.at(-1);
  if (previous?.kind === "text") {
    previous.text += text;
  } else {
    blocks.push({ id: createUuid(), kind: "text", text });
  }
}

function resolvedSourceKind(source: string, requested: DocumentImportSourceKind) {
  if (requested !== "auto") return requested;
  return /\\(?:documentclass|begin\s*\{document\}|section\*?\s*\{|usepackage)\b/.test(source)
    ? "latex"
    : "markdown";
}

export function parseLatexMarkdownDocument(
  source: string,
  requestedKind: DocumentImportSourceKind = "auto",
  defaultFontSizePt = 12,
): DocumentImportBlock[] {
  const normalized = source.replace(/^\uFEFF/, "").replace(/\r\n?/g, "\n");
  const sourceKind = resolvedSourceKind(normalized, requestedKind);
  const blocks: DocumentImportBlock[] = [];
  let textStart = 0;
  let cursor = 0;

  while (cursor < normalized.length) {
    const delimiter = delimiterAt(normalized, cursor);
    if (!delimiter) {
      cursor += 1;
      continue;
    }
    const closeIndex = findClosingDelimiter(normalized, delimiter);
    if (closeIndex < 0) {
      cursor += delimiter.opening.length;
      continue;
    }

    appendTextBlock(blocks, normalized.slice(textStart, cursor), sourceKind);
    const latex = (delimiter.environment
      ? normalized.slice(cursor, closeIndex + delimiter.closing.length)
      : normalized.slice(delimiter.contentStart, closeIndex)
    )
      .trim()
      .replace(/^\s*\\displaystyle\s*/, "")
      .replace(/\\label\s*\{[^{}]*\}/g, "")
      .replace(/\\tag\*?\s*\{[^{}]*\}/g, "")
      .trim();
    if (latex) {
      blocks.push({
        id: createUuid(),
        kind: "formula",
        latex,
        displayMode: delimiter.displayMode,
        numbered: delimiter.displayMode === "block" && delimiter.numbered,
        fontSizePt: Math.min(512, Math.max(1, defaultFontSizePt)),
      });
    } else {
      appendTextBlock(
        blocks,
        normalized.slice(cursor, closeIndex + delimiter.closing.length),
        sourceKind,
      );
    }
    cursor = closeIndex + delimiter.closing.length;
    textStart = cursor;
  }

  appendTextBlock(blocks, normalized.slice(textStart), sourceKind);
  return blocks.filter((block) => block.kind === "formula" || block.text.length > 0);
}

export function mergeDocumentImportBlocks(
  previous: DocumentImportBlock[],
  next: DocumentImportBlock[],
): DocumentImportBlock[] {
  const previousFormulas = previous.filter(
    (block): block is DocumentFormulaBlock => block.kind === "formula",
  );
  let formulaIndex = 0;
  return next.map((block) => {
    if (block.kind === "text") return block;
    const prior = previousFormulas[formulaIndex];
    formulaIndex += 1;
    if (!prior) return block;
    return {
      ...block,
      id: prior.id,
      fontSizePt: prior.fontSizePt,
      displayMode: prior.displayMode,
      numbered: prior.displayMode === "block" && prior.numbered,
    };
  });
}
