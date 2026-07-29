import assert from "node:assert/strict";
import {
  mergeDocumentImportBlocks,
  parseLatexMarkdownDocument,
} from "../src/office/documentImport/documentImportParser.ts";

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

console.log("Document import parser smoke test passed");
