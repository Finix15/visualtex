export type DocumentSourceFormat = "auto" | "markdown" | "latex";
export type ResolvedDocumentSourceFormat = Exclude<DocumentSourceFormat, "auto">;
export type DocumentObjectMode = "wordOmml" | "nativeOle";

export type DocumentImportRun =
  | {
      kind: "text";
      text: string;
      bold?: boolean;
      italic?: boolean;
      code?: boolean;
    }
  | {
      kind: "formula";
      latex: string;
      display: boolean;
    };

export type DocumentImportBlockKind =
  | "paragraph"
  | "heading"
  | "bullet"
  | "numbered"
  | "quote"
  | "code"
  | "display";

export interface DocumentImportBlock {
  id: string;
  kind: DocumentImportBlockKind;
  level: number;
  runs: DocumentImportRun[];
}

export interface ParsedDocumentImport {
  format: ResolvedDocumentSourceFormat;
  blocks: DocumentImportBlock[];
  warnings: string[];
  formulaCount: number;
  inlineFormulaCount: number;
  displayFormulaCount: number;
  textCharacterCount: number;
}

const displayEnvironmentPattern = /\\begin\{(equation\*?|align\*?|gather\*?|multline\*?|displaymath)\}/gi;

function id() {
  return crypto.randomUUID();
}

function isEscaped(text: string, index: number) {
  let slashes = 0;
  for (let cursor = index - 1; cursor >= 0 && text[cursor] === "\\"; cursor -= 1) {
    slashes += 1;
  }
  return slashes % 2 === 1;
}

function findUnescaped(text: string, token: string, start = 0) {
  for (let index = Math.max(0, start); index <= text.length - token.length; index += 1) {
    if (text.slice(index, index + token.length) === token && !isEscaped(text, index)) {
      return index;
    }
  }
  return -1;
}

function decodeText(text: string, format: ResolvedDocumentSourceFormat) {
  if (format === "markdown") {
    return text.replace(/\\([\\`*_{}\[\]()#+\-.!$])/g, "$1");
  }
  return text
    .replace(/~/g, "\u00a0")
    .replace(/\\%/g, "%")
    .replace(/\\_/g, "_")
    .replace(/\\&/g, "&")
    .replace(/\\#/g, "#")
    .replace(/\\\$/g, "$")
    .replace(/\\\{/g, "{")
    .replace(/\\\}/g, "}")
    .replace(/\\textbackslash\{\}/g, "\\")
    .replace(/\\newline/g, "\n")
    .replace(/\\\\/g, "\n");
}

function findMatchingBrace(text: string, open: number) {
  let depth = 0;
  for (let index = open; index < text.length; index += 1) {
    if (text[index] === "{" && !isEscaped(text, index)) depth += 1;
    if (text[index] === "}" && !isEscaped(text, index)) {
      depth -= 1;
      if (depth === 0) return index;
    }
  }
  return -1;
}

function mergeTextRuns(runs: DocumentImportRun[]) {
  const merged: DocumentImportRun[] = [];
  for (const run of runs) {
    const previous = merged.at(-1);
    if (
      run.kind === "text" &&
      previous?.kind === "text" &&
      Boolean(previous.bold) === Boolean(run.bold) &&
      Boolean(previous.italic) === Boolean(run.italic) &&
      Boolean(previous.code) === Boolean(run.code)
    ) {
      previous.text += run.text;
    } else {
      merged.push(run);
    }
  }
  return merged;
}

function parseInline(
  text: string,
  format: ResolvedDocumentSourceFormat,
  inherited: { bold?: boolean; italic?: boolean; code?: boolean } = {},
): DocumentImportRun[] {
  const runs: DocumentImportRun[] = [];
  let buffer = "";
  const flush = () => {
    if (!buffer) return;
    runs.push({ kind: "text", text: decodeText(buffer, format), ...inherited });
    buffer = "";
  };

  for (let index = 0; index < text.length; ) {
    if (text[index] === "$" && !isEscaped(text, index) && text[index + 1] !== "$") {
      const end = findUnescaped(text, "$", index + 1);
      if (end > index + 1) {
        flush();
        runs.push({ kind: "formula", latex: text.slice(index + 1, end).trim(), display: false });
        index = end + 1;
        continue;
      }
    }
    if (text.startsWith("\\(", index) && !isEscaped(text, index)) {
      const end = findUnescaped(text, "\\)", index + 2);
      if (end > index + 2) {
        flush();
        runs.push({ kind: "formula", latex: text.slice(index + 2, end).trim(), display: false });
        index = end + 2;
        continue;
      }
    }

    if (format === "markdown") {
      if (text.startsWith("**", index)) {
        const end = text.indexOf("**", index + 2);
        if (end > index + 2) {
          flush();
          runs.push(...parseInline(text.slice(index + 2, end), format, { ...inherited, bold: true }));
          index = end + 2;
          continue;
        }
      }
      if ((text[index] === "*" || text[index] === "_") && !isEscaped(text, index)) {
        const end = findUnescaped(text, text[index], index + 1);
        if (end > index + 1) {
          flush();
          runs.push(...parseInline(text.slice(index + 1, end), format, { ...inherited, italic: true }));
          index = end + 1;
          continue;
        }
      }
      if (text[index] === "`") {
        const end = text.indexOf("`", index + 1);
        if (end > index + 1) {
          flush();
          runs.push({ kind: "text", text: text.slice(index + 1, end), ...inherited, code: true });
          index = end + 1;
          continue;
        }
      }
    } else if (text[index] === "\\") {
      const commands = [
        { name: "\\textbf{", style: { bold: true } },
        { name: "\\textit{", style: { italic: true } },
        { name: "\\emph{", style: { italic: true } },
        { name: "\\texttt{", style: { code: true } },
      ];
      const command = commands.find((candidate) => text.startsWith(candidate.name, index));
      if (command) {
        const open = index + command.name.length - 1;
        const close = findMatchingBrace(text, open);
        if (close > open) {
          flush();
          runs.push(
            ...parseInline(text.slice(open + 1, close), format, {
              ...inherited,
              ...command.style,
            }),
          );
          index = close + 1;
          continue;
        }
      }
    }

    buffer += text[index];
    index += 1;
  }
  flush();
  return mergeTextRuns(runs);
}

function normalizeDisplayEnvironment(environment: string, body: string) {
  const normalized = body.replace(/\r\n?/g, "\n").trim();
  switch (environment.replace(/\*$/, "").toLowerCase()) {
    case "align":
      return `\\begin{aligned}${normalized}\\end{aligned}`;
    case "gather":
    case "multline":
      return `\\begin{gathered}${normalized}\\end{gathered}`;
    default:
      return normalized;
  }
}

function normalizeDelimitedDisplay(body: string) {
  // Newlines inside $$...$$ and \\[...\\] are TeX whitespace, not
  // VisualTeX row boundaries. Collapsing them keeps paired \\left/\\right
  // delimiters in one expression while explicit \\\\ row breaks remain.
  return body
    .replace(/\r\n?/g, "\n")
    .replace(/[ \t]*\n+[ \t]*/g, " ")
    .trim();
}

interface DisplayStart {
  position: number;
  startToken: string;
  endToken: string;
  environment?: string;
}

function findDisplayStart(text: string, format: ResolvedDocumentSourceFormat, from: number): DisplayStart | null {
  const candidates: DisplayStart[] = [];
  const dollars = findUnescaped(text, "$$", from);
  if (dollars >= 0) candidates.push({ position: dollars, startToken: "$$", endToken: "$$" });
  const bracket = findUnescaped(text, "\\[", from);
  if (bracket >= 0) candidates.push({ position: bracket, startToken: "\\[", endToken: "\\]" });
  if (format === "latex") {
    displayEnvironmentPattern.lastIndex = from;
    const match = displayEnvironmentPattern.exec(text);
    if (match && !isEscaped(text, match.index)) {
      candidates.push({
        position: match.index,
        startToken: match[0],
        endToken: `\\end{${match[1]}}`,
        environment: match[1],
      });
    }
  }
  return candidates.sort((left, right) => left.position - right.position)[0] ?? null;
}

function textBlock(
  kind: Exclude<DocumentImportBlockKind, "display" | "code">,
  text: string,
  format: ResolvedDocumentSourceFormat,
  level = 0,
): DocumentImportBlock | null {
  const normalized = text.replace(/\s*\n\s*/g, " ").trim();
  if (!normalized) return null;
  return { id: id(), kind, level, runs: parseInline(normalized, format) };
}

function appendMixedBlocks(
  blocks: DocumentImportBlock[],
  text: string,
  format: ResolvedDocumentSourceFormat,
  warnings: string[],
  textKind: Exclude<DocumentImportBlockKind, "display" | "code"> = "paragraph",
  level = 0,
) {
  let cursor = 0;
  while (cursor < text.length) {
    const start = findDisplayStart(text, format, cursor);
    if (!start) {
      const block = textBlock(textKind, text.slice(cursor), format, level);
      if (block) blocks.push(block);
      return;
    }
    const before = textBlock(textKind, text.slice(cursor, start.position), format, level);
    if (before) blocks.push(before);
    const contentStart = start.position + start.startToken.length;
    const end = findUnescaped(text, start.endToken, contentStart);
    if (end < 0) {
      const body = text.slice(contentStart);
      blocks.push({
        id: id(),
        kind: "display",
        level: 0,
        runs: [
          {
            kind: "formula",
            latex: start.environment
              ? normalizeDisplayEnvironment(start.environment, body)
              : normalizeDelimitedDisplay(body),
            display: true,
          },
        ],
      });
      warnings.push(
        start.environment
          ? `LaTeX 环境 ${start.environment} 未闭合，预览已读取到文末。`
          : `行间公式缺少结束标记 ${start.endToken}，预览已读取到文末。`,
      );
      return;
    }
    const body = text.slice(contentStart, end);
    blocks.push({
      id: id(),
      kind: "display",
      level: 0,
      runs: [
        {
          kind: "formula",
          latex: start.environment
            ? normalizeDisplayEnvironment(start.environment, body)
            : normalizeDelimitedDisplay(body),
          display: true,
        },
      ],
    });
    cursor = end + start.endToken.length;
  }
}

function detectFormat(source: string): ResolvedDocumentSourceFormat {
  return /\\(?:documentclass|begin\{document\}|\[|\(|text(?:bf|it|tt)\{|emph\{|item(?:\s|\[)|(?:part|chapter|section|subsection)\*?\{|begin\{(?:equation|align|itemize|enumerate|quote|quotation|verbatim|lstlisting))/i.test(
    source,
  )
    ? "latex"
    : "markdown";
}

function normalizeLatexSource(source: string, warnings: string[]) {
  let body = source;
  const begin = body.search(/\\begin\{document\}/i);
  if (begin >= 0) {
    const contentStart = begin + body.slice(begin).match(/^\\begin\{document\}/i)![0].length;
    const end = body.slice(contentStart).search(/\\end\{document\}/i);
    body = end >= 0 ? body.slice(contentStart, contentStart + end) : body.slice(contentStart);
    if (end < 0) warnings.push("LaTeX 文档缺少 \\end{document}，预览已读取其余内容。");
  }

  const result: string[] = [];
  let literalEnd = "";
  for (const line of body.split("\n")) {
    const trimmed = line.trim();
    if (literalEnd) {
      result.push(line);
      if (trimmed.toLowerCase() === literalEnd.toLowerCase()) literalEnd = "";
      continue;
    }
    const literal = trimmed.match(/^\\begin\{(verbatim|lstlisting)\}(?:\[[^\]]*\])?\s*$/i);
    if (literal) {
      literalEnd = `\\end{${literal[1]}}`;
      result.push(line);
      continue;
    }
    const comment = findUnescaped(line, "%", 0);
    result.push(comment >= 0 ? line.slice(0, comment) : line);
  }
  return result.join("\n").trim();
}

function listLevel(indentation: string) {
  const columns = [...indentation].reduce((total, character) => total + (character === "\t" ? 4 : 1), 0);
  return Math.min(8, Math.max(0, Math.floor(columns / 2)));
}

export function parseDocumentImport(
  source: string,
  requestedFormat: DocumentSourceFormat,
): ParsedDocumentImport {
  if (!source.trim()) throw new Error("请输入需要导入的 LaTeX 或 Markdown 内容。");
  if (source.length > 5_000_000) throw new Error("批量导入内容不能超过 5 MB。");
  const warnings: string[] = [];
  const format = requestedFormat === "auto" ? detectFormat(source) : requestedFormat;
  const normalized = (format === "latex" ? normalizeLatexSource(source, warnings) : source)
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .trim();
  const blocks: DocumentImportBlock[] = [];
  const paragraph: string[] = [];
  const quote: string[] = [];
  const listModes: Array<"bullet" | "numbered"> = [];
  let inLatexQuote = false;
  let inCode = false;
  let codeEnd = "";
  let codeDescription = "";
  const code: string[] = [];

  const flushParagraph = () => {
    if (!paragraph.length) return;
    appendMixedBlocks(blocks, paragraph.join("\n"), format, warnings);
    paragraph.length = 0;
  };
  const flushQuote = () => {
    if (!quote.length) return;
    appendMixedBlocks(blocks, quote.join("\n"), format, warnings, "quote");
    quote.length = 0;
  };
  const finishCode = (warning?: string) => {
    blocks.push({
      id: id(),
      kind: "code",
      level: 0,
      runs: [{ kind: "text", text: code.join("\n").replace(/[\r\n]+$/, ""), code: true }],
    });
    code.length = 0;
    inCode = false;
    codeEnd = "";
    codeDescription = "";
    if (warning) warnings.push(warning);
  };

  for (const raw of normalized.split("\n")) {
    const trimmed = raw.trim();
    if (format === "markdown" && trimmed.startsWith("```")) {
      flushParagraph();
      flushQuote();
      if (inCode && codeEnd === "```") finishCode();
      else if (!inCode) {
        inCode = true;
        codeEnd = "```";
        codeDescription = "Markdown 代码块";
      } else code.push(raw);
      continue;
    }
    if (format === "latex" && !inCode) {
      const start = trimmed.match(/^\\begin\{(verbatim|lstlisting)\}(?:\[[^\]]*\])?\s*$/i);
      if (start) {
        flushParagraph();
        flushQuote();
        inCode = true;
        codeEnd = `\\end{${start[1]}}`;
        codeDescription = `LaTeX ${start[1]} 环境`;
        continue;
      }
    }
    if (inCode) {
      if (trimmed.toLowerCase() === codeEnd.toLowerCase()) finishCode();
      else code.push(raw);
      continue;
    }

    if (format === "latex") {
      if (/^\\begin\{(?:quote|quotation)\}\s*$/i.test(trimmed)) {
        flushParagraph();
        flushQuote();
        inLatexQuote = true;
        continue;
      }
      if (/^\\end\{(?:quote|quotation)\}\s*$/i.test(trimmed)) {
        flushParagraph();
        flushQuote();
        inLatexQuote = false;
        continue;
      }
      const listStart = trimmed.match(/^\\begin\{(itemize|enumerate)\}\s*$/i);
      if (listStart) {
        flushParagraph();
        flushQuote();
        listModes.push(listStart[1].toLowerCase() === "enumerate" ? "numbered" : "bullet");
        continue;
      }
      if (/^\\end\{(?:itemize|enumerate)\}\s*$/i.test(trimmed)) {
        flushParagraph();
        flushQuote();
        if (listModes.length) listModes.pop();
        else warnings.push(`忽略了没有对应开始标记的 ${trimmed}。`);
        continue;
      }
      const heading = trimmed.match(/^\\(part|chapter|section|subsection|subsubsection|paragraph|subparagraph)\*?\{(.*)\}\s*$/i);
      if (heading) {
        flushParagraph();
        flushQuote();
        const levels: Record<string, number> = {
          part: 1,
          chapter: 1,
          section: 1,
          subsection: 2,
          subsubsection: 3,
          paragraph: 4,
          subparagraph: 5,
        };
        blocks.push({
          id: id(),
          kind: "heading",
          level: levels[heading[1].toLowerCase()] ?? 4,
          runs: parseInline(heading[2], format),
        });
        continue;
      }
      const item = trimmed.match(/^\\item(?:\s*\[[^\]]*\])?\s*(.*)$/);
      if (item) {
        flushParagraph();
        flushQuote();
        appendMixedBlocks(
          blocks,
          item[1],
          format,
          warnings,
          listModes.at(-1) === "numbered" ? "numbered" : "bullet",
          Math.max(0, listModes.length - 1),
        );
        continue;
      }
    } else {
      if (trimmed.startsWith(">")) {
        flushParagraph();
        quote.push(trimmed.replace(/^>+\s?/, ""));
        continue;
      }
      flushQuote();
      const heading = raw.match(/^(#{1,6})\s+(.+?)\s*#*\s*$/);
      if (heading) {
        flushParagraph();
        blocks.push({ id: id(), kind: "heading", level: heading[1].length, runs: parseInline(heading[2], format) });
        continue;
      }
      const bullet = raw.match(/^(\s*)[-+*]\s+(.+)$/);
      if (bullet) {
        flushParagraph();
        appendMixedBlocks(blocks, bullet[2], format, warnings, "bullet", listLevel(bullet[1]));
        continue;
      }
      const numbered = raw.match(/^(\s*)\d+[.)]\s+(.+)$/);
      if (numbered) {
        flushParagraph();
        appendMixedBlocks(blocks, numbered[2], format, warnings, "numbered", listLevel(numbered[1]));
        continue;
      }
    }

    if (!trimmed) {
      flushParagraph();
      flushQuote();
      continue;
    }
    if (inLatexQuote) quote.push(trimmed);
    else paragraph.push(trimmed);
  }

  if (inCode) finishCode(`${codeDescription}未闭合，预览已读取到文末。`);
  flushParagraph();
  flushQuote();
  if (inLatexQuote) warnings.push("LaTeX quote/quotation 环境未闭合，预览已读取到文末。");
  if (listModes.length) warnings.push(`LaTeX 文档有 ${listModes.length} 个列表环境未闭合。`);
  if (!blocks.length) throw new Error("没有找到可以插入 Word 的文字或公式。");

  const runs = blocks.flatMap((block) => block.runs);
  const formulaCount = runs.filter((run) => run.kind === "formula").length;
  const displayFormulaCount = runs.filter((run) => run.kind === "formula" && run.display).length;
  const textCharacterCount = runs.reduce(
    (total, run) => total + (run.kind === "text" ? run.text.length : 0),
    0,
  );
  return {
    format,
    blocks,
    warnings,
    formulaCount,
    inlineFormulaCount: formulaCount - displayFormulaCount,
    displayFormulaCount,
    textCharacterCount,
  };
}
