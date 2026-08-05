import assert from "node:assert/strict";

import {
  findWindowsWordLatexRedrawSpans,
  type WordLatexRedrawSpan,
} from "../src/office/redraw/wordLatexRedrawParser.ts";

function assertExactSpan(source: string, span: WordLatexRedrawSpan) {
  assert.equal(source.slice(span.start, span.end), span.sourceText);
  assert.ok(span.end > span.start);
}

const mixedSource = [
  "前😀缀 $x+1$ 中间 \\$5 不是公式。",
  "显示：\\[ y^2 \\]，行内：\\(z\\)。",
  "被转义的分隔符：\\\\(not-math\\) 与 \\\\[still-not-math\\]。",
].join("\n");
const mixed = findWindowsWordLatexRedrawSpans(mixedSource);
assert.deepEqual(
  mixed.map((span) => [span.sourceText, span.latex, span.displayMode]),
  [
    ["$x+1$", "x+1", "inline"],
    ["\\[ y^2 \\]", "y^2", "block"],
    ["\\(z\\)", "z", "inline"],
  ],
);
for (const span of mixed) assertExactSpan(mixedSource, span);
assert.equal(mixed[0].start, "前😀缀 ".length);

const displayWhitespace = "before $$\n  x +\n y  \n$$ after";
assert.equal(
  findWindowsWordLatexRedrawSpans(displayWhitespace)[0].latex,
  "x + y",
  "Windows redraw treats source newlines inside display delimiters as TeX whitespace",
);

const environmentSource = [
  String.raw`\begin{equation}E=mc^2\end{equation}`,
  String.raw`\begin{align*}a&=b\\c&=d\end{align*}`,
  String.raw`\begin{gather}x\\y\end{gather}`,
  String.raw`\begin{multline*}p\\q\end{multline*}`,
  String.raw`\begin{displaymath}r+s\end{displaymath}`,
].join(" ");
assert.deepEqual(
  findWindowsWordLatexRedrawSpans(environmentSource).map((span) => [
    span.sourceText,
    span.latex,
    span.displayMode,
  ]),
  [
    [String.raw`\begin{equation}E=mc^2\end{equation}`, "E=mc^2", "block"],
    [
      String.raw`\begin{align*}a&=b\\c&=d\end{align*}`,
      String.raw`\begin{aligned}a&=b\\c&=d\end{aligned}`,
      "block",
    ],
    [
      String.raw`\begin{gather}x\\y\end{gather}`,
      String.raw`\begin{gathered}x\\y\end{gathered}`,
      "block",
    ],
    [
      String.raw`\begin{multline*}p\\q\end{multline*}`,
      String.raw`\begin{gathered}p\\q\end{gathered}`,
      "block",
    ],
    [String.raw`\begin{displaymath}r+s\end{displaymath}`, "r+s", "block"],
  ],
);

assert.deepEqual(
  findWindowsWordLatexRedrawSpans(
    String.raw`\begin{alignat}ignored\end{alignat} \begin{math}ignored\end{math}`,
  ),
  [],
  "macOS redraw must keep the exact mature Windows environment allow-list",
);

assert.equal(
  findWindowsWordLatexRedrawSpans("unclosed $x+1").length,
  0,
);
assert.equal(
  findWindowsWordLatexRedrawSpans(String.raw`\begin{align}x&=1`).length,
  0,
);
assert.equal(findWindowsWordLatexRedrawSpans("$$$$").length, 0);

const replacementSource = "甲😀 $a$ 乙 $$b$$ 丙 \\(c\\) 丁";
const replacementSpans = findWindowsWordLatexRedrawSpans(replacementSource);
let replaced = replacementSource;
for (let index = replacementSpans.length - 1; index >= 0; index -= 1) {
  const span = replacementSpans[index];
  replaced =
    replaced.slice(0, span.start) +
    `<FORMULA:${span.latex}>` +
    replaced.slice(span.end);
}
assert.equal(
  replaced,
  "甲😀 <FORMULA:a> 乙 <FORMULA:b> 丙 <FORMULA:c> 丁",
);

const exactlyOneThousand = Array.from(
  { length: 1000 },
  (_value, index) => `$${index}$`,
).join(" ");
assert.equal(findWindowsWordLatexRedrawSpans(exactlyOneThousand).length, 1000);
assert.throws(
  () => findWindowsWordLatexRedrawSpans(`${exactlyOneThousand} $overflow$`),
  /maximum 1000/,
);
assert.throws(
  () => findWindowsWordLatexRedrawSpans("x".repeat(5_000_001)),
  /5 MB/,
);

console.log("VisualTeX Windows-parity Word LaTeX redraw parser regression: PASS");
