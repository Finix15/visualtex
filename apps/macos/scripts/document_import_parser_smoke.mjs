import assert from "node:assert/strict";
import { documentImportErrorMessage } from "../src/office/documentImport/documentImportErrors.ts";
import {
  mergeDocumentImportBlocks,
  parseLatexMarkdownDocument,
} from "../src/office/documentImport/documentImportParser.ts";

assert.equal(
  documentImportErrorMessage({ error: { message: "SVG staging failed" } }, "fallback"),
  "SVG staging failed",
);
assert.notEqual(
  documentImportErrorMessage({ code: 4608, details: { stage: "Word" } }, "fallback"),
  "[object Object]",
);

const markdown = [
  "动量满足 $p=mv$，这是行内公式。",
  "",
  "$$E=mc^2$$",
  "",
  "结尾文字。",
].join("\n");
const markdownBlocks = parseLatexMarkdownDocument(markdown, "markdown", 11);
const markdownFormulas = markdownBlocks.filter((block) => block.kind === "formula");
assert.equal(markdownFormulas.length, 2);
assert.deepEqual(
  markdownFormulas.map((block) => [block.displayMode, block.numbered, block.fontSizePt]),
  [
    ["inline", false, 11],
    ["block", false, 11],
  ],
);
assert.equal(markdownFormulas[0].latex, "p=mv");
assert.equal(markdownFormulas[1].latex, "E=mc^2");
assert.match(
  markdownBlocks.filter((block) => block.kind === "text").map((block) => block.text).join(""),
  /动量满足.*这是行内公式。.*结尾文字。/s,
);

const latex = String.raw`\documentclass{article}
\begin{document}
\section{测试}
正文包含 \(a^2+b^2=c^2\)。
\begin{equation}
E=mc^2\label{eq:energy}
\end{equation}
\begin{equation*}
F=ma
\end{equation*}
\begin{align}
a&=b+c\\
d&=e-f
\end{align}
\end{document}`;
const latexBlocks = parseLatexMarkdownDocument(latex, "latex", 12);
const latexFormulas = latexBlocks.filter((block) => block.kind === "formula");
assert.equal(latexFormulas.length, 4);
assert.deepEqual(
  latexFormulas.map((block) => [block.displayMode, block.numbered]),
  [
    ["inline", false],
    ["block", true],
    ["block", false],
    ["block", true],
  ],
);
assert.equal(latexFormulas[0].latex, "a^2+b^2=c^2");
assert.match(latexFormulas[1].latex, /^\\begin\{equation\}/);
assert.doesNotMatch(latexFormulas[1].latex, /\\label/);
assert.match(latexFormulas[3].latex, /^\\begin\{align\}/);

const customized = latexBlocks.map((block) =>
  block.kind === "formula"
    ? { ...block, fontSizePt: 18, numbered: block.displayMode === "block" }
    : block,
);
const reparsed = parseLatexMarkdownDocument(latex, "latex", 10);
const merged = mergeDocumentImportBlocks(customized, reparsed);
const mergedFormulas = merged.filter((block) => block.kind === "formula");
assert.ok(mergedFormulas.every((block) => block.fontSizePt === 18));
assert.equal(mergedFormulas[0].numbered, false);
assert.ok(mergedFormulas.slice(1).every((block) => block.numbered));

const semanticLatex = String.raw`\documentclass{article}
\begin{document}
\section{结构化标题}
正文中的公式 $x=1$ 与文字属于同一段。
\begin{itemize}
\item 第一项
\item 第二项包含 $y=2$
\end{itemize}
\begin{enumerate}
\item 编号一
\item 编号二
\end{enumerate}
$$E=mc^2$$
公式后的正文。
\end{document}`;
const semanticBlocks = parseLatexMarkdownDocument(semanticLatex, "latex", 12);
const heading = semanticBlocks.find(
  (block) => block.kind === "text" && block.text === "结构化标题",
);
assert.equal(heading?.paragraphStyle, "heading1");
assert.equal(heading?.paragraphStart, true);
assert.equal(heading?.paragraphEnd, true);
const bulletParagraphs = new Set(
  semanticBlocks
    .filter((block) => block.listKind === "bullet")
    .map((block) => block.paragraphId),
);
assert.equal(bulletParagraphs.size, 2);
const numberedParagraphs = new Set(
  semanticBlocks
    .filter((block) => block.listKind === "number")
    .map((block) => block.paragraphId),
);
assert.equal(numberedParagraphs.size, 2);
const inlineListFormula = semanticBlocks.find(
  (block) => block.kind === "formula" && block.latex === "y=2",
);
assert.equal(inlineListFormula?.listKind, "bullet");
assert.equal(inlineListFormula?.paragraphEnd, true);
const displayFormula = semanticBlocks.find(
  (block) => block.kind === "formula" && block.latex === "E=mc^2",
);
assert.equal(displayFormula?.displayMode, "block");
assert.equal(displayFormula?.paragraphId, undefined);
assert.ok(
  semanticBlocks
    .filter((block) => block.kind === "text")
    .every((block) => !block.text.includes("\n\n")),
  "Display-formula boundaries must not preserve blank text paragraphs",
);

const semanticMarkdown = parseLatexMarkdownDocument(
  "# 标题\n\n- 项目一\n- 项目二含 $z=3$\n\n1. 编号一\n2. 编号二",
  "markdown",
  12,
);
assert.equal(semanticMarkdown[0]?.paragraphStyle, "heading1");
assert.equal(
  semanticMarkdown.find((block) => block.kind === "formula" && block.latex === "z=3")
    ?.listKind,
  "bullet",
);
assert.equal(
  semanticMarkdown.filter((block) => block.listKind === "number" && block.paragraphStart)
    .length,
  2,
);

console.log("Document import parser smoke test passed");
