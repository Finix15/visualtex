import assert from "node:assert/strict";
import { DOMParser } from "@xmldom/xmldom";
import { parseLatexMarkdownDocument } from "../src/office/documentImport/documentImportParser.ts";
import { normalizeFormulaEditorDocument } from "../src/office/shared/formulaEditorDocument.ts";
import { renderOfficeFormulaArtifacts } from "../src/office/shared/formulaRenderArtifacts.ts";

globalThis.DOMParser ??= DOMParser;
const domProbe = new DOMParser().parseFromString("<root/>", "application/xml");
const documentPrototype = Object.getPrototypeOf(domProbe);
const elementPrototype = Object.getPrototypeOf(domProbe.documentElement);
if (typeof documentPrototype.querySelector !== "function") {
  documentPrototype.querySelector = function querySelector(name) {
    return this.getElementsByTagName(name)?.item(0) ?? null;
  };
}
if (!("children" in elementPrototype)) {
  Object.defineProperty(elementPrototype, "children", {
    configurable: true,
    get() {
      return Array.from(this.childNodes ?? []).filter((node) => node.nodeType === 1);
    },
  });
}

const allLineEndings = ["lf", "crlf", "cr"];

const fixtures = [
  {
    name: "latex-inline-dollar",
    kind: "latex",
    source: "正文 $a^2+b^2=c^2$ 结尾。",
    expected: [{ mode: "inline", contains: "a^2+b^2" }],
  },
  {
    name: "latex-inline-paren",
    kind: "latex",
    source: String.raw`正文 \(E=mc^2\) 结尾。`,
    expected: [{ mode: "inline", contains: "E=mc^2" }],
  },
  {
    name: "latex-inline-math-environment",
    kind: "latex",
    source: String.raw`正文 \begin{math}\alpha+\beta\end{math} 结尾。`,
    expected: [{ mode: "inline", contains: "\\alpha+\\beta" }],
  },
  {
    name: "latex-display-dollar-single-line",
    kind: "latex",
    source: "前文\n$$E=mc^2$$\n后文",
    lineEndings: allLineEndings,
    expected: [{ mode: "block", numbered: false, contains: "E=mc^2" }],
  },
  {
    name: "latex-display-dollar-multiline",
    kind: "latex",
    source: "$$\n\\frac{1}{2}mv^2\n$$",
    lineEndings: allLineEndings,
    expected: [{ mode: "block", contains: "\\frac{1}{2}" }],
  },
  {
    name: "latex-display-bracket",
    kind: "latex",
    source: String.raw`\[
\int_0^1 x^2\,dx
\]`,
    lineEndings: allLineEndings,
    expected: [{ mode: "block", contains: "\\int_0^1" }],
  },
  {
    name: "latex-equation",
    kind: "latex",
    source: String.raw`\begin{equation}
F=ma
\end{equation}`,
    lineEndings: allLineEndings,
    expected: [{ mode: "block", numbered: true, contains: "\\begin{equation}" }],
  },
  {
    name: "latex-equation-star",
    kind: "latex",
    source: String.raw`\begin{equation*}
E=mc^2
\end{equation*}`,
    expected: [{ mode: "block", numbered: false, contains: "equation*" }],
  },
  {
    name: "latex-align",
    kind: "latex",
    source: String.raw`\begin{align}
a&=b+c\\
d&=e-f
\end{align}`,
    lineEndings: allLineEndings,
    expected: [{ mode: "block", numbered: true, contains: "\\begin{align}" }],
  },
  {
    name: "latex-align-star",
    kind: "latex",
    source: String.raw`\begin{align*}
x&=y\\
y&=z
\end{align*}`,
    expected: [{ mode: "block", numbered: false, contains: "align*" }],
  },
  {
    name: "latex-alignat",
    kind: "latex",
    source: String.raw`\begin{alignat}{2}
a&=b&\qquad c&=d\\
e&=f& g&=h
\end{alignat}`,
    expected: [{ mode: "block", numbered: true, contains: "alignat" }],
  },
  {
    name: "latex-alignat-star",
    kind: "latex",
    source: String.raw`\begin{alignat*}{2}
a&=b& c&=d\\
e&=f& g&=h
\end{alignat*}`,
    expected: [{ mode: "block", numbered: false, contains: "alignat*" }],
  },
  {
    name: "latex-flalign",
    kind: "latex",
    source: String.raw`\begin{flalign}
a&=b+c&&\\
d&=e-f&&
\end{flalign}`,
    expected: [{ mode: "block", numbered: true, contains: "flalign" }],
  },
  {
    name: "latex-flalign-star",
    kind: "latex",
    source: String.raw`\begin{flalign*}
a&=b&&\\
c&=d&&
\end{flalign*}`,
    expected: [{ mode: "block", numbered: false, contains: "flalign*" }],
  },
  {
    name: "latex-eqnarray",
    kind: "latex",
    source: String.raw`\begin{eqnarray}
a&=&b+c\\
d&=&e-f
\end{eqnarray}`,
    expected: [{ mode: "block", numbered: true, contains: "eqnarray" }],
  },
  {
    name: "latex-gather",
    kind: "latex",
    source: String.raw`\begin{gather}
a=b\\
c=d
\end{gather}`,
    expected: [{ mode: "block", numbered: true, contains: "gather" }],
  },
  {
    name: "latex-gather-star",
    kind: "latex",
    source: String.raw`\begin{gather*}
a=b\\
c=d
\end{gather*}`,
    expected: [{ mode: "block", numbered: false, contains: "gather*" }],
  },
  {
    name: "latex-multline",
    kind: "latex",
    source: String.raw`\begin{multline}
a+b+c+d+e\\
=f+g+h
\end{multline}`,
    expected: [{ mode: "block", numbered: true, contains: "multline" }],
  },
  {
    name: "latex-multline-star",
    kind: "latex",
    source: String.raw`\begin{multline*}
a+b+c\\
=d+e
\end{multline*}`,
    expected: [{ mode: "block", numbered: false, contains: "multline*" }],
  },
  {
    name: "latex-equation-split",
    kind: "latex",
    source: String.raw`\begin{equation}
\begin{split}
a&=b+c\\
 &=d+e
\end{split}
\end{equation}`,
    lineEndings: allLineEndings,
    expected: [{ mode: "block", numbered: true, contains: "\\begin{split}" }],
  },
  {
    name: "latex-matrix",
    kind: "latex",
    source: String.raw`\[
A=\begin{matrix}a&b\\c&d\end{matrix}
\]`,
    expected: [{ mode: "block", contains: "\\begin{matrix}" }],
  },
  {
    name: "latex-pmatrix",
    kind: "latex",
    source: String.raw`\[A=\begin{pmatrix}1&2\\3&4\end{pmatrix}\]`,
    expected: [{ mode: "block", contains: "pmatrix" }],
  },
  {
    name: "latex-bmatrix",
    kind: "latex",
    source: String.raw`\[A=\begin{bmatrix}1&0\\0&1\end{bmatrix}\]`,
    expected: [{ mode: "block", contains: "bmatrix" }],
  },
  {
    name: "latex-vmatrix-determinant",
    kind: "latex",
    source: String.raw`\[\det A=\begin{vmatrix}a&b\\c&d\end{vmatrix}\]`,
    expected: [{ mode: "block", contains: "vmatrix" }],
  },
  {
    name: "latex-cases",
    kind: "latex",
    source: String.raw`\[f(x)=\begin{cases}x^2,&x\ge0\\-x,&x<0\end{cases}\]`,
    expected: [{ mode: "block", contains: "cases" }],
  },
  {
    name: "latex-array",
    kind: "latex",
    source: String.raw`\[\begin{array}{c|cc}&A&B\\x&1&2\end{array}\]`,
    expected: [{ mode: "block", contains: "\\begin{array}" }],
  },
  {
    name: "latex-alignedat-inside-brackets",
    kind: "latex",
    source: String.raw`\[\begin{alignedat}{2}a&=b&\quad c&=d\\e&=f&g&=h\end{alignedat}\]`,
    expected: [{ mode: "block", contains: "alignedat" }],
  },
  {
    name: "latex-nested-fractions-roots",
    kind: "latex",
    source: String.raw`\[\frac{1}{1+\frac{x}{\sqrt{1+x^2}}}\]`,
    expected: [{ mode: "block", contains: "\\sqrt" }],
  },
  {
    name: "latex-accents-decorators",
    kind: "latex",
    source: String.raw`\[\hat{x}+\bar{y}+\vec{v}+\overset{!}{=}+\underbrace{a+b}_{c}\]`,
    expected: [{ mode: "block", contains: "\\underbrace" }],
  },
  {
    name: "latex-operators-integrals",
    kind: "latex",
    source: String.raw`\[\sum_{n=1}^{\infty}\int_{-\infty}^{\infty}e^{-x^2}\,dx\]`,
    expected: [{ mode: "block", contains: "\\sum" }],
  },
  {
    name: "latex-label-and-tag-removal",
    kind: "latex",
    source: String.raw`\begin{equation}
E=mc^2\label{eq:e}\tag{A}
\end{equation}`,
    expected: [{ mode: "block", numbered: true, excludes: ["\\label", "\\tag"] }],
  },
  {
    name: "latex-comments-and-escaped-percent",
    kind: "latex",
    source: String.raw`正文 50\% 与 $x=1$。 % remove this
\[
y=2 % remove formula comment
\]`,
    lineEndings: allLineEndings,
    expected: [
      { mode: "inline", contains: "x=1" },
      { mode: "block", contains: "y=2" },
    ],
  },
  {
    name: "latex-escaped-symbols",
    kind: "latex",
    source: String.raw`价格 \$5，比例 50\%，变量 $a\_1+b\&c$。`,
    expected: [{ mode: "inline", contains: "a\\_1" }],
  },
  {
    name: "latex-complete-document",
    kind: "auto",
    source: String.raw`\documentclass[12pt]{article}
\usepackage{amsmath}
\title{测试文档}
\author{VisualTeX}
\begin{document}
\maketitle
\section{公式}
正文 $x=1$。
\begin{equation*}
y=2
\end{equation*}
\end{document}`,
    lineEndings: allLineEndings,
    expected: [
      { mode: "inline", contains: "x=1" },
      { mode: "block", numbered: false, contains: "equation*" },
    ],
    textIncludes: ["测试文档", "VisualTeX", "公式"],
  },
  {
    name: "latex-itemize-inline-formulas",
    kind: "auto",
    source: String.raw`\begin{itemize}
\item 第一项 $a=1$
\item 第二项 \(b=2\)
\end{itemize}`,
    expected: [
      { mode: "inline", contains: "a=1" },
      { mode: "inline", contains: "b=2" },
    ],
  },
  {
    name: "latex-nested-lists",
    kind: "latex",
    source: String.raw`\begin{enumerate}
\item 外层
\begin{itemize}
\item 内层 $x=1$
\end{itemize}
\item 末项
\end{enumerate}`,
    expected: [{ mode: "inline", contains: "x=1" }],
  },
  {
    name: "latex-quote-and-center",
    kind: "latex",
    source: String.raw`\begin{quote}
引用包含 $q=mv$。
\end{quote}
\begin{center}
\[E=mc^2\]
\end{center}`,
    expected: [
      { mode: "inline", contains: "q=mv" },
      { mode: "block", contains: "E=mc^2" },
    ],
  },
  {
    name: "latex-multiple-inline-one-paragraph",
    kind: "latex",
    source: "先 $a=1$，再 $b=2$，最后 \\(c=3\\)。",
    expected: [
      { mode: "inline", contains: "a=1" },
      { mode: "inline", contains: "b=2" },
      { mode: "inline", contains: "c=3" },
    ],
  },
  {
    name: "latex-flexible-environment-spacing",
    kind: "latex",
    source: String.raw`\begin {equation}
E=mc^2
\end {equation}`,
    expected: [{ mode: "block", numbered: true, contains: "E=mc^2" }],
  },
  {
    name: "latex-verbatim-protected",
    kind: "latex",
    source: String.raw`\begin{verbatim}
$not_math$ and \[not_math\]
\end{verbatim}
正文 $x=1$。`,
    expected: [{ mode: "inline", contains: "x=1" }],
    forbiddenFormulaText: ["not_math"],
  },
  {
    name: "latex-verb-protected",
    kind: "latex",
    source: String.raw`命令 \verb|$not_math$| 后面是 $x=1$。`,
    expected: [{ mode: "inline", contains: "x=1" }],
    forbiddenFormulaText: ["not_math"],
  },
  {
    name: "latex-minted-protected",
    kind: "latex",
    source: String.raw`\begin{minted}{tex}
$x$ and \[y\]
\end{minted}
\[z=3\]`,
    expected: [{ mode: "block", contains: "z=3" }],
    forbiddenFormulaText: ["$x$", "y"],
  },
  {
    name: "auto-markdown-with-matrix-environment",
    kind: "auto",
    source: String.raw`**粗体 Markdown** 与矩阵：

\[
A=\begin{matrix}1&2\\3&4\end{matrix}
\]`,
    expected: [{ mode: "block", contains: "\\begin{matrix}" }],
    textIncludes: ["粗体 Markdown"],
    textExcludes: ["**"],
  },
  {
    name: "markdown-heading-inline",
    kind: "markdown",
    source: "# 标题\n\n正文含 $x=1$。",
    lineEndings: allLineEndings,
    expected: [{ mode: "inline", contains: "x=1" }],
    textIncludes: ["标题"],
  },
  {
    name: "markdown-display-dollar",
    kind: "markdown",
    source: "前文\n\n$$\nE=mc^2\n$$\n\n后文",
    lineEndings: allLineEndings,
    expected: [{ mode: "block", contains: "E=mc^2" }],
  },
  {
    name: "markdown-display-bracket",
    kind: "markdown",
    source: String.raw`说明

\[
\frac{a}{b}
\]`,
    expected: [{ mode: "block", contains: "\\frac{a}{b}" }],
  },
  {
    name: "markdown-bullet-list",
    kind: "markdown",
    source: "- 第一项 $a=1$\n- 第二项 $b=2$",
    expected: [
      { mode: "inline", contains: "a=1" },
      { mode: "inline", contains: "b=2" },
    ],
  },
  {
    name: "markdown-nested-list",
    kind: "markdown",
    source: "- 外层\n  - 内层 $x=1$\n    - 更深 $y=2$",
    expected: [
      { mode: "inline", contains: "x=1" },
      { mode: "inline", contains: "y=2" },
    ],
  },
  {
    name: "markdown-numbered-list",
    kind: "markdown",
    source: "1. 第一项 $a=1$\n2) 第二项 $b=2$",
    expected: [
      { mode: "inline", contains: "a=1" },
      { mode: "inline", contains: "b=2" },
    ],
  },
  {
    name: "markdown-quote",
    kind: "markdown",
    source: "> 引用中的公式 $q=mv$\n> 第二行",
    expected: [{ mode: "inline", contains: "q=mv" }],
  },
  {
    name: "markdown-fenced-code-protected",
    kind: "markdown",
    source: "```tex\n$not_math$\n\\[not_math\\]\n```\n\n正文 $x=1$。",
    expected: [{ mode: "inline", contains: "x=1" }],
    forbiddenFormulaText: ["not_math"],
  },
  {
    name: "markdown-tilde-fence-protected",
    kind: "markdown",
    source: "~~~latex\n$$not_math$$\n~~~\n\n正文 $x=1$。",
    expected: [{ mode: "inline", contains: "x=1" }],
    forbiddenFormulaText: ["not_math"],
  },
  {
    name: "markdown-inline-code-protected",
    kind: "markdown",
    source: "代码 `$not_math$`，真实公式 $x=1$。",
    expected: [{ mode: "inline", contains: "x=1" }],
    forbiddenFormulaText: ["not_math"],
  },
  {
    name: "markdown-links-emphasis",
    kind: "markdown",
    source: "**粗体**、*斜体*、[链接](https://example.com) 与 $x=1$。",
    expected: [{ mode: "inline", contains: "x=1" }],
    textIncludes: ["粗体", "斜体", "链接"],
  },
  {
    name: "markdown-multiple-inline",
    kind: "markdown",
    source: "同一段 $a=1$、$b=2$ 和 \\(c=3\\)。",
    expected: [
      { mode: "inline", contains: "a=1" },
      { mode: "inline", contains: "b=2" },
      { mode: "inline", contains: "c=3" },
    ],
  },
  {
    name: "markdown-escaped-dollar",
    kind: "markdown",
    source: "价格是 \\$5，不是公式；真正公式 $x=1$。",
    expected: [{ mode: "inline", contains: "x=1" }],
  },
  {
    name: "markdown-table",
    kind: "markdown",
    source: "| 量 | 值 |\n|---|---|\n| 能量 | $E=mc^2$ |",
    expected: [{ mode: "inline", contains: "E=mc^2" }],
  },
  {
    name: "markdown-mixed-delimiters",
    kind: "markdown",
    source: String.raw`行内 $a=1$ 与 \(b=2\)。

$$c=3$$

\[d=4\]`,
    expected: [
      { mode: "inline", contains: "a=1" },
      { mode: "inline", contains: "b=2" },
      { mode: "block", contains: "c=3" },
      { mode: "block", contains: "d=4" },
    ],
  },
  {
    name: "markdown-multiline-display-body",
    kind: "markdown",
    source: "$$\n\\begin{aligned}\na&=b\\\\\nc&=d\n\\end{aligned}\n$$",
    expected: [{ mode: "block", contains: "\\begin{aligned}" }],
  },
  {
    name: "markdown-indented-code-protected",
    kind: "markdown",
    source: "    $not_math$\n\n正文 $x=1$。",
    expected: [{ mode: "inline", contains: "x=1" }],
    forbiddenFormulaText: ["not_math"],
  },
  {
    name: "markdown-crlf-paragraphs",
    kind: "markdown",
    source: "第一段 $a=1$。\n\n第二段。\n\n$$b=2$$",
    lineEndings: allLineEndings,
    expected: [
      { mode: "inline", contains: "a=1" },
      { mode: "block", contains: "b=2" },
    ],
  },
];

function withLineEnding(source, lineEnding) {
  const normalized = source.replace(/\r\n?/g, "\n");
  if (lineEnding === "crlf") return normalized.replace(/\n/g, "\r\n");
  if (lineEnding === "cr") return normalized.replace(/\n/g, "\r");
  return normalized;
}

let executedCases = 0;
let renderedFormulas = 0;
for (const fixture of fixtures) {
  const variants = fixture.lineEndings ?? ["lf"];
  for (const lineEnding of variants) {
    executedCases += 1;
    const label = `${fixture.name}:${lineEnding}`;
    const blocks = parseLatexMarkdownDocument(
      withLineEnding(fixture.source, lineEnding),
      fixture.kind,
      12,
    );
    const formulas = blocks.filter((block) => block.kind === "formula");
    assert.equal(
      formulas.length,
      fixture.expected.length,
      `${label} formula count`,
    );

    fixture.expected.forEach((expected, index) => {
      const formula = formulas[index];
      assert.equal(formula.displayMode, expected.mode, `${label} formula ${index + 1} mode`);
      if (typeof expected.numbered === "boolean") {
        assert.equal(
          formula.numbered,
          expected.numbered,
          `${label} formula ${index + 1} numbering`,
        );
      }
      if (expected.contains) {
        assert.ok(
          formula.latex.includes(expected.contains),
          `${label} formula ${index + 1} must contain ${expected.contains}`,
        );
      }
      for (const excluded of expected.excludes ?? []) {
        assert.ok(
          !formula.latex.includes(excluded),
          `${label} formula ${index + 1} must remove ${excluded}`,
        );
      }

      const document = normalizeFormulaEditorDocument(
        [{ id: `fixture-${executedCases}-${index}`, latex: formula.latex }],
        "raw",
      );
      let artifacts;
      try {
        artifacts = renderOfficeFormulaArtifacts({
          lines: document.lines,
          codeFormat: document.codeFormat,
          displayMode: formula.displayMode,
          host: "word",
        });
      } catch (error) {
        throw new Error(
          `${label} formula ${index + 1} artifact generation failed: ${error instanceof Error ? error.message : String(error)}`,
          { cause: error },
        );
      }
      assert.ok(artifacts.canonicalLatex.trim(), `${label} canonical LaTeX`);
      assert.ok(artifacts.svg.svg.includes("<svg"), `${label} SVG artifact`);
      assert.match(
        artifacts.svg.svg,
        /(?:fill|stroke)=["']#000000["']/i,
        `${label} explicit Word black paint`,
      );
      assert.ok(artifacts.omml?.ommlBase64, `${label} OMML artifact`);
      assert.ok(artifacts.omml?.ommlDocxBase64, `${label} OMML DOCX artifact`);
      renderedFormulas += 1;
    });

    const allFormulaText = formulas.map((formula) => formula.latex).join("\n");
    for (const forbidden of fixture.forbiddenFormulaText ?? []) {
      assert.ok(
        !allFormulaText.includes(forbidden),
        `${label} protected code must not become a formula: ${forbidden}`,
      );
    }
    const prose = blocks
      .filter((block) => block.kind === "text")
      .map((block) => block.text)
      .join("\n");
    for (const text of fixture.textIncludes ?? []) {
      assert.ok(prose.includes(text), `${label} prose must contain ${text}`);
    }
    for (const text of fixture.textExcludes ?? []) {
      assert.ok(!prose.includes(text), `${label} prose must remove ${text}`);
    }
  }
}

assert.ok(fixtures.length >= 40, "The parser corpus must keep at least 40 syntax fixtures");
assert.ok(executedCases >= 60, "Line-ending expansion must execute at least 60 cases");
console.log(
  `Document import parser corpus passed: ${fixtures.length} syntax fixtures, ${executedCases} line-ending cases, ${renderedFormulas} rendered formulas`,
);
