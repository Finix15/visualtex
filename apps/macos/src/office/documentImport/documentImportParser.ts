import { createUuid } from "../../runtime/browserCompatibility";

export type DocumentImportSourceKind = "auto" | "markdown" | "latex";
export type DocumentFormulaDisplayMode = "inline" | "block";
export type DocumentFormulaOutputKind = "omml" | "image";
export type DocumentParagraphStyle =
  | "normal"
  | "heading1"
  | "heading2"
  | "heading3"
  | "heading4"
  | "quote"
  | "code";
export type DocumentListKind = "none" | "bullet" | "number";
export type DocumentParagraphAlignment = "left" | "center" | "right" | "justify";

export interface DocumentParagraphMetadata {
  paragraphId?: string;
  paragraphStyle?: DocumentParagraphStyle;
  paragraphAlignment?: DocumentParagraphAlignment;
  listKind?: DocumentListKind;
  listLevel?: number;
  paragraphStart?: boolean;
  paragraphEnd?: boolean;
}

export interface DocumentTextBlock extends DocumentParagraphMetadata {
  id: string;
  kind: "text";
  text: string;
}

export interface DocumentFormulaBlock extends DocumentParagraphMetadata {
  id: string;
  kind: "formula";
  latex: string;
  displayMode: DocumentFormulaDisplayMode;
  numbered: boolean;
  fontSizePt: number;
}

export type DocumentImportBlock = DocumentTextBlock | DocumentFormulaBlock;

interface ExtractedFormula {
  token: string;
  block: DocumentFormulaBlock;
}

interface ParagraphContext {
  style: DocumentParagraphStyle;
  alignment: DocumentParagraphAlignment;
  listKind: DocumentListKind;
  listLevel: number;
}

const blockEnvironmentNames = new Set([
  "equation",
  "equation*",
  "align",
  "align*",
  "alignat",
  "alignat*",
  "flalign",
  "flalign*",
  "eqnarray",
  "eqnarray*",
  "gather",
  "gather*",
  "multline",
  "multline*",
  "displaymath",
]);
const inlineEnvironmentNames = new Set(["math"]);
const protectedLatexEnvironmentNames = new Set([
  "verbatim",
  "verbatim*",
  "lstlisting",
  "minted",
  "comment",
]);
const formulaTokenPrefix = "\uE000VT_FORMULA_";
const formulaTokenSuffix = "\uE001";
const formulaTokenPattern = /\uE000VT_FORMULA_(\d+)\uE001/g;
const markerPrefix = "\uE100VT_";
const markerSuffix = "\uE101";

function marker(value: string) {
  return `${markerPrefix}${value}${markerSuffix}`;
}

function leadingMarker(line: string) {
  const trimmed = line.trim();
  if (!trimmed.startsWith(markerPrefix)) return null;
  const end = trimmed.indexOf(markerSuffix, markerPrefix.length);
  if (end < 0) return null;
  return {
    value: trimmed.slice(markerPrefix.length, end),
    rest: trimmed.slice(end + markerSuffix.length),
  };
}

function isEscaped(source: string, index: number) {
  let slashCount = 0;
  for (let cursor = index - 1; cursor >= 0 && source[cursor] === "\\"; cursor -= 1) {
    slashCount += 1;
  }
  return slashCount % 2 === 1;
}

interface ProtectedSourceRange {
  start: number;
  end: number;
}

function escapeRegExp(value: string) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function mergeProtectedRanges(ranges: ProtectedSourceRange[]) {
  const sorted = ranges
    .filter((range) => range.end > range.start)
    .sort((left, right) => left.start - right.start || left.end - right.end);
  const merged: ProtectedSourceRange[] = [];
  for (const range of sorted) {
    const previous = merged.at(-1);
    if (!previous || range.start > previous.end) {
      merged.push({ ...range });
    } else {
      previous.end = Math.max(previous.end, range.end);
    }
  }
  return merged;
}

function latexProtectedRanges(source: string) {
  const ranges: ProtectedSourceRange[] = [];
  const beginPattern = /\\begin\s*\{([^{}]+)\}/g;
  for (let match = beginPattern.exec(source); match; match = beginPattern.exec(source)) {
    const environment = match[1].trim();
    if (!protectedLatexEnvironmentNames.has(environment)) continue;
    const endPattern = new RegExp(
      `\\\\end\\s*\\{${escapeRegExp(environment)}\\}`,
      "g",
    );
    endPattern.lastIndex = beginPattern.lastIndex;
    const closing = endPattern.exec(source);
    const end = closing ? endPattern.lastIndex : source.length;
    ranges.push({ start: match.index, end });
    beginPattern.lastIndex = end;
  }

  const verbPattern = /\\verb\*?([^\sA-Za-z0-9])[^\n]*?\1/g;
  for (let match = verbPattern.exec(source); match; match = verbPattern.exec(source)) {
    ranges.push({ start: match.index, end: verbPattern.lastIndex });
  }
  return mergeProtectedRanges(ranges);
}

function markdownProtectedRanges(source: string) {
  const ranges: ProtectedSourceRange[] = [];
  let offset = 0;
  let fenceStart = -1;
  let fenceCharacter = "";
  let fenceLength = 0;

  while (offset < source.length) {
    const newline = source.indexOf("\n", offset);
    const lineEnd = newline < 0 ? source.length : newline;
    const nextOffset = newline < 0 ? source.length : newline + 1;
    const line = source.slice(offset, lineEnd);
    const fence = line.match(/^ {0,3}(`{3,}|~{3,})/);
    if (fenceStart < 0 && fence) {
      fenceStart = offset;
      fenceCharacter = fence[1][0];
      fenceLength = fence[1].length;
    } else if (
      fenceStart >= 0 &&
      new RegExp(`^ {0,3}${escapeRegExp(fenceCharacter)}{${fenceLength},}\\s*$`).test(
        line,
      )
    ) {
      ranges.push({ start: fenceStart, end: nextOffset });
      fenceStart = -1;
      fenceCharacter = "";
      fenceLength = 0;
    } else if (
      fenceStart < 0 &&
      /^(?:\t| {4})/.test(line) &&
      !/^\s*(?:[-+*]|\d+[.)])\s+/.test(line)
    ) {
      ranges.push({ start: offset, end: nextOffset });
    }
    offset = nextOffset;
  }
  if (fenceStart >= 0) ranges.push({ start: fenceStart, end: source.length });

  const fenced = mergeProtectedRanges(ranges);
  const inlineCodePattern = /(`+)([^\n]*?)\1/g;
  for (
    let match = inlineCodePattern.exec(source);
    match;
    match = inlineCodePattern.exec(source)
  ) {
    const insideBlock = fenced.some(
      (range) => match.index >= range.start && match.index < range.end,
    );
    if (!insideBlock) {
      ranges.push({ start: match.index, end: inlineCodePattern.lastIndex });
    }
  }
  return mergeProtectedRanges(ranges);
}

function protectedSourceRanges(
  source: string,
  sourceKind: DocumentImportSourceKind,
) {
  return sourceKind === "latex"
    ? latexProtectedRanges(source)
    : markdownProtectedRanges(source);
}

function stripLatexComments(source: string) {
  const protectedRanges = latexProtectedRanges(source);
  let rangeIndex = 0;
  let lineOffset = 0;
  return source
    .split("\n")
    .map((line) => {
      const lineEnd = lineOffset + line.length;
      while (
        rangeIndex < protectedRanges.length &&
        protectedRanges[rangeIndex].end <= lineOffset
      ) {
        rangeIndex += 1;
      }
      const range = protectedRanges[rangeIndex];
      const protectedLine = Boolean(
        range && range.start <= lineEnd && range.end > lineOffset,
      );
      lineOffset = lineEnd + 1;
      if (protectedLine) return line;
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
    "footnote",
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
  return value.replace(/\\href\s*\{[^{}]*\}\s*\{([^{}]*)\}/g, "$1");
}

function cleanInlineMarkup(raw: string, sourceKind: DocumentImportSourceKind) {
  let value = raw.replace(/\r\n?/g, "\n");
  if (sourceKind === "latex") {
    value = unwrapSimpleLatexCommands(value)
      .replace(/\\(?:newline|linebreak)\b/g, "\n")
      .replace(/\\\\(?:\[[^\]]*\])?/g, "\n")
      .replace(/\\%/g, "%")
      .replace(/\\&/g, "&")
      .replace(/\\#/g, "#")
      .replace(/\\_/g, "_")
      .replace(/\\\$/g, "$")
      .replace(/\\\{/g, "{")
      .replace(/\\\}/g, "}")
      .replace(/~/g, " ")
      .replace(/\\[a-zA-Z@]+\*?(?:\[[^\]]*\])?/g, "");
  } else {
    value = value
      .replace(/!\[([^\]]*)\]\([^)]*\)/g, "$1")
      .replace(/\[([^\]]+)\]\([^)]*\)/g, "$1")
      .replace(/`([^`]+)`/g, "$1")
      .replace(/\*\*([^*]+)\*\*/g, "$1")
      .replace(/__([^_]+)__/g, "$1")
      .replace(/(?<!\*)\*([^*\n]+)\*(?!\*)/g, "$1")
      .replace(/(?<!_)_([^_\n]+)_(?!_)/g, "$1")
      .replace(/~~([^~]+)~~/g, "$1");
  }
  return value
    .replace(/[ \t]*\n[ \t]*/g, " ")
    .replace(/[ \t]{2,}/g, " ");
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
    const environment = match?.[1]?.trim();
    if (
      match &&
      environment &&
      (blockEnvironmentNames.has(environment) || inlineEnvironmentNames.has(environment))
    ) {
      const displayMode = inlineEnvironmentNames.has(environment) ? "inline" : "block";
      return {
        opening: match[0],
        closing: `\\end{${environment}}`,
        displayMode,
        numbered:
          displayMode === "block" &&
          !environment.endsWith("*") &&
          environment !== "displaymath",
        contentStart: index + match[0].length,
        environment,
      };
    }
  }
  return null;
}

interface ClosingDelimiterMatch {
  start: number;
  end: number;
}

function findClosingDelimiter(
  source: string,
  delimiter: MathDelimiter,
): ClosingDelimiterMatch | null {
  let cursor = delimiter.contentStart;
  if (delimiter.environment) {
    const flexiblePattern = new RegExp(
      `\\\\end\\s*\\{${escapeRegExp(delimiter.environment)}\\}`,
      "g",
    );
    flexiblePattern.lastIndex = cursor;
    const match = flexiblePattern.exec(source);
    return match
      ? { start: match.index, end: flexiblePattern.lastIndex }
      : null;
  }
  while (cursor < source.length) {
    const found = source.indexOf(delimiter.closing, cursor);
    if (found < 0) return null;
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
    return { start: found, end: found + delimiter.closing.length };
  }
  return null;
}

function extractFormulas(
  source: string,
  defaultFontSizePt: number,
  sourceKind: DocumentImportSourceKind,
) {
  const formulas: ExtractedFormula[] = [];
  const protectedRanges = protectedSourceRanges(source, sourceKind);
  let protectedIndex = 0;
  let output = "";
  let textStart = 0;
  let cursor = 0;

  while (cursor < source.length) {
    while (
      protectedIndex < protectedRanges.length &&
      protectedRanges[protectedIndex].end <= cursor
    ) {
      protectedIndex += 1;
    }
    const protectedRange = protectedRanges[protectedIndex];
    if (
      protectedRange &&
      cursor >= protectedRange.start &&
      cursor < protectedRange.end
    ) {
      cursor = protectedRange.end;
      continue;
    }

    const delimiter = delimiterAt(source, cursor);
    if (!delimiter) {
      cursor += 1;
      continue;
    }
    const closing = findClosingDelimiter(source, delimiter);
    if (!closing) {
      cursor += delimiter.opening.length;
      continue;
    }
    if (
      protectedRange &&
      closing.start > protectedRange.start &&
      cursor < protectedRange.start
    ) {
      cursor = protectedRange.end;
      continue;
    }

    output += source.slice(textStart, cursor);
    const latex = (delimiter.environment && delimiter.displayMode === "block"
      ? source.slice(cursor, closing.end)
      : source.slice(delimiter.contentStart, closing.start)
    )
      .trim()
      .replace(/^\s*\\displaystyle\s*/, "")
      .replace(/\\label\s*\{[^{}]*\}/g, "")
      .replace(/\\tag\*?\s*\{[^{}]*\}/g, "")
      .trim();

    if (latex) {
      const token = `${formulaTokenPrefix}${formulas.length}${formulaTokenSuffix}`;
      formulas.push({
        token,
        block: {
          id: createUuid(),
          kind: "formula",
          latex,
          displayMode: delimiter.displayMode,
          numbered: delimiter.displayMode === "block" && delimiter.numbered,
          fontSizePt: Math.min(512, Math.max(1, defaultFontSizePt)),
        },
      });
      output += delimiter.displayMode === "block" ? `\n${token}\n` : token;
    } else {
      output += source.slice(cursor, closing.end);
    }
    cursor = closing.end;
    textStart = cursor;
  }
  output += source.slice(textStart);
  return { text: output, formulas };
}

function formulaFromToken(token: string, formulas: ExtractedFormula[]) {
  const match = token.match(/^\uE000VT_FORMULA_(\d+)\uE001$/);
  const index = match ? Number(match[1]) : -1;
  return Number.isInteger(index) ? formulas[index]?.block : undefined;
}

function defaultContext(): ParagraphContext {
  return { style: "normal", alignment: "left", listKind: "none", listLevel: 0 };
}

function headingStyle(level: number): DocumentParagraphStyle {
  if (level <= 1) return "heading1";
  if (level === 2) return "heading2";
  if (level === 3) return "heading3";
  return "heading4";
}

function normalizeLatexStructure(source: string) {
  return source
    .replace(/\\documentclass(?:\[[^\]]*\])?\s*\{[^{}]*\}/g, "")
    .replace(/\\usepackage(?:\[[^\]]*\])?\s*\{[^{}]*\}/g, "")
    .replace(/\\begin\s*\{document\}|\\end\s*\{document\}/g, "")
    .replace(/\\maketitle\b/g, "")
    .replace(/\\title\s*\{([^{}]*)\}/g, `\n${marker("HEADING:1")}$1\n`)
    .replace(/\\author\s*\{([^{}]*)\}/g, `\n${marker("CENTER")}$1\n`)
    .replace(/\\date\s*\{([^{}]*)\}/g, `\n${marker("CENTER")}$1\n`)
    .replace(/\\(?:part|chapter|section)\*?\s*\{([^{}]*)\}/g, `\n${marker("HEADING:1")}$1\n`)
    .replace(/\\subsection\*?\s*\{([^{}]*)\}/g, `\n${marker("HEADING:2")}$1\n`)
    .replace(/\\subsubsection\*?\s*\{([^{}]*)\}/g, `\n${marker("HEADING:3")}$1\n`)
    .replace(/\\(?:paragraph|subparagraph)\*?\s*\{([^{}]*)\}/g, `\n${marker("HEADING:4")}$1\n`)
    .replace(/\\begin\s*\{itemize\}/g, `\n${marker("LIST_START:bullet")}\n`)
    .replace(/\\end\s*\{itemize\}/g, `\n${marker("LIST_END")}\n`)
    .replace(/\\begin\s*\{enumerate\}/g, `\n${marker("LIST_START:number")}\n`)
    .replace(/\\end\s*\{enumerate\}/g, `\n${marker("LIST_END")}\n`)
    .replace(/\\item(?:\[[^\]]*\])?\s*/g, `\n${marker("ITEM")}`)
    .replace(/\\begin\s*\{(?:quote|quotation)\}/g, `\n${marker("QUOTE_START")}\n`)
    .replace(/\\end\s*\{(?:quote|quotation)\}/g, `\n${marker("QUOTE_END")}\n`)
    .replace(/\\begin\s*\{center\}/g, `\n${marker("ALIGN:center")}\n`)
    .replace(/\\end\s*\{center\}/g, `\n${marker("ALIGN:left")}\n`)
    .replace(/\\begin\s*\{flushright\}/g, `\n${marker("ALIGN:right")}\n`)
    .replace(/\\end\s*\{flushright\}/g, `\n${marker("ALIGN:left")}\n`)
    .replace(/\\begin\s*\{flushleft\}/g, `\n${marker("ALIGN:left")}\n`)
    .replace(/\\end\s*\{flushleft\}/g, `\n${marker("ALIGN:left")}\n`)
    .replace(/\\begin\s*\{(?:abstract|description)\}/g, "\n")
    .replace(/\\end\s*\{(?:abstract|description)\}/g, "\n")
    .replace(/\\par\b/g, "\n\n");
}

function emitParagraph(
  blocks: DocumentImportBlock[],
  raw: string,
  context: ParagraphContext,
  formulas: ExtractedFormula[],
  sourceKind: DocumentImportSourceKind,
) {
  if (!raw.trim()) return;
  const paragraphId = createUuid();
  const content: DocumentImportBlock[] = [];
  let cursor = 0;
  formulaTokenPattern.lastIndex = 0;
  for (let match = formulaTokenPattern.exec(raw); match; match = formulaTokenPattern.exec(raw)) {
    const text = cleanInlineMarkup(raw.slice(cursor, match.index), sourceKind);
    if (text) content.push({ id: createUuid(), kind: "text", text });
    const formula = formulaFromToken(match[0], formulas);
    if (formula) content.push({ ...formula });
    cursor = match.index + match[0].length;
  }
  const tail = cleanInlineMarkup(raw.slice(cursor), sourceKind);
  if (tail) content.push({ id: createUuid(), kind: "text", text: tail });
  if (!content.length) return;

  const firstText = content.find((block): block is DocumentTextBlock => block.kind === "text");
  const lastText = [...content]
    .reverse()
    .find((block): block is DocumentTextBlock => block.kind === "text");
  if (firstText) firstText.text = firstText.text.replace(/^\s+/, "");
  if (lastText) lastText.text = lastText.text.replace(/\s+$/, "");

  const visible = content.filter((block) => block.kind === "formula" || block.text.length > 0);
  visible.forEach((block, index) => {
    Object.assign(block, {
      paragraphId,
      paragraphStyle: context.style,
      paragraphAlignment: context.alignment,
      listKind: context.listKind,
      listLevel: context.listLevel,
      paragraphStart: index === 0,
      paragraphEnd: index === visible.length - 1,
    });
  });
  blocks.push(...visible);
}

function emitDisplayFormula(
  blocks: DocumentImportBlock[],
  token: string,
  formulas: ExtractedFormula[],
) {
  const formula = formulaFromToken(token, formulas);
  if (formula) blocks.push({ ...formula });
}

function parseStructuredLines(
  source: string,
  formulas: ExtractedFormula[],
  sourceKind: DocumentImportSourceKind,
) {
  const blocks: DocumentImportBlock[] = [];
  const listStack: DocumentListKind[] = [];
  let alignment: DocumentParagraphAlignment = "left";
  let quoteDepth = 0;
  let current = "";
  let currentContext = defaultContext();
  let codeFence = false;

  const flush = () => {
    emitParagraph(blocks, current, currentContext, formulas, sourceKind);
    current = "";
    currentContext = {
      style: quoteDepth > 0 ? "quote" : codeFence ? "code" : "normal",
      alignment,
      listKind: "none",
      listLevel: 0,
    };
  };

  const beginParagraph = (context: ParagraphContext, text: string) => {
    flush();
    currentContext = context;
    current = text;
  };

  for (const originalLine of source.split("\n")) {
    const line = originalLine.replace(/[ \t]+$/g, "");
    const trimmed = line.trim();
    const leading = leadingMarker(trimmed);
    const control = leading?.value;

    if (control) {
      if (control.startsWith("HEADING:")) {
        const title = leading?.rest ?? "";
        beginParagraph(
          {
            style: headingStyle(Number(control.split(":")[1])),
            alignment: "left",
            listKind: "none",
            listLevel: 0,
          },
          title,
        );
        flush();
      } else if (control.startsWith("LIST_START:")) {
        flush();
        listStack.push(control.endsWith("number") ? "number" : "bullet");
      } else if (control === "LIST_END") {
        flush();
        listStack.pop();
      } else if (control === "ITEM") {
        beginParagraph(
          {
            style: "normal",
            alignment,
            listKind: listStack.at(-1) ?? "bullet",
            listLevel: Math.max(1, listStack.length),
          },
          leading?.rest ?? "",
        );
      } else if (control === "QUOTE_START") {
        flush();
        quoteDepth += 1;
      } else if (control === "QUOTE_END") {
        flush();
        quoteDepth = Math.max(0, quoteDepth - 1);
      } else if (control.startsWith("ALIGN:")) {
        flush();
        const value = control.split(":")[1];
        alignment = value === "center" || value === "right" ? value : "left";
      } else if (control === "CENTER") {
        beginParagraph(
          { style: "normal", alignment: "center", listKind: "none", listLevel: 0 },
          leading?.rest ?? "",
        );
        flush();
      }
      continue;
    }

    const displayFormula = formulaFromToken(trimmed, formulas);
    if (displayFormula?.displayMode === "block") {
      flush();
      emitDisplayFormula(blocks, trimmed, formulas);
      continue;
    }

    if (sourceKind === "markdown") {
      if (/^```/.test(trimmed)) {
        flush();
        codeFence = !codeFence;
        continue;
      }
      const heading = trimmed.match(/^(#{1,6})\s+(.+)$/);
      if (heading) {
        beginParagraph(
          {
            style: headingStyle(heading[1].length),
            alignment: "left",
            listKind: "none",
            listLevel: 0,
          },
          heading[2],
        );
        flush();
        continue;
      }
      const bullet = line.match(/^(\s*)[-+*]\s+(.+)$/);
      if (bullet) {
        beginParagraph(
          {
            style: "normal",
            alignment: "left",
            listKind: "bullet",
            listLevel: Math.max(1, Math.floor(bullet[1].length / 2) + 1),
          },
          bullet[2],
        );
        continue;
      }
      const numbered = line.match(/^(\s*)\d+[.)]\s+(.+)$/);
      if (numbered) {
        beginParagraph(
          {
            style: "normal",
            alignment: "left",
            listKind: "number",
            listLevel: Math.max(1, Math.floor(numbered[1].length / 2) + 1),
          },
          numbered[2],
        );
        continue;
      }
      const quote = trimmed.match(/^>\s?(.*)$/);
      if (quote) {
        beginParagraph(
          { style: "quote", alignment: "left", listKind: "none", listLevel: 0 },
          quote[1],
        );
        continue;
      }
    }

    if (!trimmed) {
      flush();
      continue;
    }
    if (!current) {
      currentContext = {
        style: quoteDepth > 0 ? "quote" : codeFence ? "code" : "normal",
        alignment,
        listKind: "none",
        listLevel: 0,
      };
      current = trimmed;
    } else {
      current += ` ${trimmed}`;
    }
  }
  flush();
  return blocks;
}

function resolvedSourceKind(source: string, requested: DocumentImportSourceKind) {
  if (requested !== "auto") return requested;
  // LaTeX fragments pasted from notes often omit both \documentclass and the
  // document environment. Structural commands and environments must still
  // select the LaTeX parser. Do not end this pattern with \b: environment
  // matches end in `}`, which is not a word character, so that boundary made
  // valid `\begin{itemize}` fragments fail auto-detection.
  return /\\(?:documentclass|usepackage|title\s*\{|author\s*\{|date\s*\{|maketitle\b|(?:part|chapter|section|subsection|subsubsection|paragraph|subparagraph)\*?\s*\{|item(?:\[[^\]]*\])?\s|(?:textbf|textit|emph|underline|texttt|textrm|textsf|textnormal|mbox|caption|footnote|href)\s*\{|(?:newline|linebreak|par)\b|(?:begin|end)\s*\{(?:document|itemize|enumerate|quote|quotation|center|flushleft|flushright|abstract|description|verbatim\*?|lstlisting|minted|comment)\})/.test(
    source,
  )
    ? "latex"
    : "markdown";
}

export function parseLatexMarkdownDocument(
  source: string,
  requestedKind: DocumentImportSourceKind = "auto",
  defaultFontSizePt = 12,
): DocumentImportBlock[] {
  let normalized = source.replace(/^\uFEFF/, "").replace(/\r\n?/g, "\n");
  const sourceKind = resolvedSourceKind(normalized, requestedKind);
  if (sourceKind === "latex") normalized = stripLatexComments(normalized);
  const extracted = extractFormulas(normalized, defaultFontSizePt, sourceKind);
  const structured =
    sourceKind === "latex" ? normalizeLatexStructure(extracted.text) : extracted.text;
  return parseStructuredLines(structured, extracted.formulas, sourceKind);
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
