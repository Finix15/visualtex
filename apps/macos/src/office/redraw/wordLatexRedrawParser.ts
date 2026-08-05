export type WordLatexRedrawDisplayMode = "inline" | "block";

export interface WordLatexRedrawSpan {
  start: number;
  end: number;
  sourceText: string;
  latex: string;
  displayMode: WordLatexRedrawDisplayMode;
}

const MAX_REDRAW_SOURCE_CHARACTERS = 5_000_000;
const MAX_REDRAW_FORMULAS = 1_000;
const DISPLAY_ENVIRONMENT = /\\begin\{(?<name>equation\*?|align\*?|gather\*?|multline\*?|displaymath)\}/iy;

function isEscaped(text: string, index: number) {
  let slashes = 0;
  for (let cursor = index - 1; cursor >= 0 && text[cursor] === "\\"; cursor -= 1) {
    slashes += 1;
  }
  return slashes % 2 === 1;
}

function findUnescaped(text: string, target: string, start: number) {
  for (let index = Math.max(0, start); index < text.length; index += 1) {
    if (text[index] === target && !isEscaped(text, index)) return index;
  }
  return -1;
}

function findUnescapedSequence(text: string, target: string, start: number) {
  if (!target) return -1;
  for (
    let index = Math.max(0, start);
    index <= text.length - target.length;
    index += 1
  ) {
    if (!text.startsWith(target, index)) continue;
    if (!isEscaped(text, index)) return index;
  }
  return -1;
}

function normalizeDisplayEnvironmentLatex(environment: string, body: string) {
  const normalizedBody = body.replace(/\r\n?/g, "\n").trim();
  const baseEnvironment = environment.replace(/\*+$/g, "").toLowerCase();
  if (baseEnvironment === "align") {
    return `\\begin{aligned}${normalizedBody}\\end{aligned}`;
  }
  if (baseEnvironment === "gather" || baseEnvironment === "multline") {
    return `\\begin{gathered}${normalizedBody}\\end{gathered}`;
  }
  return normalizedBody;
}

function normalizeDelimitedDisplayLatex(body: string) {
  return body
    .replace(/\r\n?/g, "\n")
    .replace(/[ \t]*\n+[ \t]*/g, " ")
    .trim();
}

function addFormulaSpan(
  spans: WordLatexRedrawSpan[],
  source: string,
  sourceStart: number,
  sourceEnd: number,
  bodyStart: number,
  bodyEnd: number,
  displayMode: WordLatexRedrawDisplayMode,
  normalizeDisplay: boolean,
) {
  if (bodyEnd <= bodyStart) return;
  const body = source.slice(bodyStart, bodyEnd);
  const latex = normalizeDisplay ? normalizeDelimitedDisplayLatex(body) : body.trim();
  if (!latex) return;
  spans.push({
    start: sourceStart,
    end: sourceEnd,
    sourceText: source.slice(sourceStart, sourceEnd),
    latex,
    displayMode,
  });
}

/**
 * Direct TypeScript port of Windows WordBulkImportParser.FindFormulaSpans.
 * Keep delimiter semantics and UTF-16 offsets aligned with the mature Windows
 * in-place redraw implementation; this parser is intentionally independent of
 * the document-import parser and does not construct document blocks.
 */
export function findWindowsWordLatexRedrawSpans(source: string) {
  if (source.length > MAX_REDRAW_SOURCE_CHARACTERS) {
    throw new Error("LaTeX redraw range cannot exceed 5 MB.");
  }

  const spans: WordLatexRedrawSpan[] = [];
  for (let index = 0; index < source.length; ) {
    if (source[index] === "$" && !isEscaped(source, index)) {
      if (index + 1 < source.length && source[index + 1] === "$") {
        const end = findUnescapedSequence(source, "$$", index + 2);
        if (end >= index + 2) {
          addFormulaSpan(
            spans,
            source,
            index,
            end + 2,
            index + 2,
            end,
            "block",
            true,
          );
          index = end + 2;
          continue;
        }
      } else {
        const end = findUnescaped(source, "$", index + 1);
        if (end > index + 1) {
          addFormulaSpan(
            spans,
            source,
            index,
            end + 1,
            index + 1,
            end,
            "inline",
            false,
          );
          index = end + 1;
          continue;
        }
      }
    }

    if (
      source[index] === "\\" &&
      index + 1 < source.length &&
      (source[index + 1] === "(" || source[index + 1] === "[") &&
      !isEscaped(source, index)
    ) {
      const display = source[index + 1] === "[";
      const endToken = display ? "\\]" : "\\)";
      const end = findUnescapedSequence(source, endToken, index + 2);
      if (end > index + 2) {
        addFormulaSpan(
          spans,
          source,
          index,
          end + endToken.length,
          index + 2,
          end,
          display ? "block" : "inline",
          display,
        );
        index = end + endToken.length;
        continue;
      }
    }

    if (source[index] === "\\" && !isEscaped(source, index)) {
      DISPLAY_ENVIRONMENT.lastIndex = index;
      const environment = DISPLAY_ENVIRONMENT.exec(source);
      if (environment?.index === index) {
        const name = environment.groups?.name ?? "";
        const endToken = `\\end{${name}}`;
        const bodyStart = index + environment[0].length;
        const end = findUnescapedSequence(source, endToken, bodyStart);
        if (end >= bodyStart) {
          const latex = normalizeDisplayEnvironmentLatex(
            name,
            source.slice(bodyStart, end),
          );
          if (latex) {
            spans.push({
              start: index,
              end: end + endToken.length,
              sourceText: source.slice(index, end + endToken.length),
              latex,
              displayMode: "block",
            });
          }
          index = end + endToken.length;
          continue;
        }
      }
    }

    index += 1;
  }

  if (spans.length > MAX_REDRAW_FORMULAS) {
    throw new Error("LaTeX redraw contains too many formulas (maximum 1000).");
  }
  return spans;
}
