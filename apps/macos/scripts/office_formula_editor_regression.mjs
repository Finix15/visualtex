import assert from "node:assert/strict";
import {
  normalizeFormulaEditorDocument,
  serializeFormulaEditorDocument,
} from "../src/office/shared/formulaEditorDocument.ts";
import { createFormulaMetadata } from "../src/office/shared/formulaMetadata.ts";
import { renderOfficeFormulaArtifacts } from "../src/office/shared/formulaRenderArtifacts.ts";
import { latexToSvg } from "../src/export/latexToSvg.ts";
import { errorMessage } from "../src/runtime/errorMessage.ts";

function normalize(source, codeFormat = "raw") {
  return normalizeFormulaEditorDocument(
    [{ id: "original-line-id", latex: source }],
    codeFormat,
  );
}

const multilineCases = [
  {
    name: "align",
    source: String.raw`\begin{align}
a &= b + c \\
d &= e
\end{align}`,
    codeFormat: "align",
    lines: ["a = b + c", "d = e"],
  },
  {
    name: "align-star",
    source: String.raw`\begin{align*}
x &= y \\
y &= z
\end{align*}`,
    codeFormat: "align-star",
    lines: ["x = y", "y = z"],
  },
  {
    name: "aligned",
    source: String.raw`\begin{aligned}
p &= q \\
r &= s
\end{aligned}`,
    codeFormat: "aligned",
    lines: ["p = q", "r = s"],
  },
  {
    name: "gather",
    source: String.raw`\begin{gather}
a=b \\
c=d
\end{gather}`,
    codeFormat: "gather",
    lines: ["a=b", "c=d"],
  },
  {
    name: "gather-star",
    source: String.raw`\begin{gather*}
a=b \\
c=d
\end{gather*}`,
    codeFormat: "gather-star",
    lines: ["a=b", "c=d"],
  },
  {
    name: "multline",
    source: String.raw`\begin{multline}
a+b+c \\
=d+e
\end{multline}`,
    codeFormat: "multline",
    lines: ["a+b+c", "=d+e"],
  },
  {
    name: "multline-star",
    source: String.raw`\begin{multline*}
a+b+c \\
=d+e
\end{multline*}`,
    codeFormat: "multline-star",
    lines: ["a+b+c", "=d+e"],
  },
  {
    name: "equation-split",
    source: String.raw`\begin{equation}
\begin{split}
a &= b \\
c &= d
\end{split}
\end{equation}`,
    codeFormat: "equation-split",
    lines: ["a = b", "c = d"],
  },
  {
    name: "equation-star-split",
    source: String.raw`\begin{equation*}
\begin{split}
a &= b \\
c &= d
\end{split}
\end{equation*}`,
    codeFormat: "equation-star-split",
    lines: ["a = b", "c = d"],
  },
];

for (const testCase of multilineCases) {
  const normalized = normalize(testCase.source);
  assert.equal(
    normalized.codeFormat,
    testCase.codeFormat,
    `${testCase.name} code format`,
  );
  assert.deepEqual(
    normalized.lines.map((line) => line.latex),
    testCase.lines,
    `${testCase.name} rows`,
  );
  assert.equal(
    normalized.lines[0].id,
    "original-line-id",
    `${testCase.name} must preserve the imported first-line UUID`,
  );
  assert.equal(
    new Set(normalized.lines.map((line) => line.id)).size,
    normalized.lines.length,
    `${testCase.name} row UUIDs must remain unique`,
  );
  const canonicalSource = serializeFormulaEditorDocument(normalized);
  const roundTrip = normalize(canonicalSource, normalized.codeFormat);
  assert.equal(
    serializeFormulaEditorDocument(roundTrip),
    canonicalSource,
    `${testCase.name} canonical source must be stable`,
  );
  assert.deepEqual(
    roundTrip.lines.map((line) => line.latex),
    testCase.lines,
    `${testCase.name} canonical source must preserve every row`,
  );
  const metadata = createFormulaMetadata({
    formulaId: "12345678-1234-4234-9234-123456789abc",
    title: testCase.name,
    lines: normalized.lines,
    codeFormat: normalized.codeFormat,
    sourceLatex: canonicalSource,
    displayMode: "block",
  });
  assert.equal(
    metadata.latex,
    canonicalSource,
    `${testCase.name} metadata must store the canonical serialized source`,
  );
  if (testCase.codeFormat === "align" || testCase.codeFormat === "align-star") {
    const rendered = renderOfficeFormulaArtifacts({
      lines: normalized.lines,
      codeFormat: normalized.codeFormat,
      displayMode: "block",
      includeWordOmml: false,
    });
    const firstImportSvg = latexToSvg(canonicalSource, {
      displayMode: true,
      fontSizePt: 14,
      paddingPx: 10,
      background: "transparent",
    });
    assert.equal(
      rendered.canonicalLatex,
      canonicalSource,
      `${testCase.name} edit rendering must rebuild the complete environment`,
    );
    assert.equal(
      rendered.svg.svg.replace(/MJX-\d+-/g, "MJX-N-"),
      firstImportSvg.svg.replace(/MJX-\d+-/g, "MJX-N-"),
      `${testCase.name} first import and edit replacement must share the same SVG`,
    );
    assert.equal(rendered.svg.width, firstImportSvg.width);
    assert.equal(rendered.svg.height, firstImportSvg.height);
    assert.equal(rendered.svg.baseline, firstImportSvg.baseline);
  }
}

const equation = normalize(String.raw`\begin{equation}E=mc^2\end{equation}`);
assert.equal(equation.codeFormat, "equation");
assert.deepEqual(equation.lines, [
  { id: "original-line-id", latex: "E=mc^2" },
]);

const displayMath = normalize(
  String.raw`\begin{displaymath}x^2+y^2=z^2\end{displaymath}`,
);
assert.equal(displayMath.codeFormat, "equation-star");
assert.equal(displayMath.lines[0].latex, "x^2+y^2=z^2");

const alignat = normalize(String.raw`\begin{alignat}{2}
a&=b &\quad c&=d \\
e&=f &\quad g&=h
\end{alignat}`);
assert.equal(alignat.codeFormat, "align");
assert.deepEqual(
  alignat.lines.map((line) => line.latex),
  ["a=b \\quad c=d", "e=f \\quad g=h"],
);

const alreadyNormalized = normalizeFormulaEditorDocument(
  [
    { id: "line-a", latex: "a=b" },
    { id: "line-b", latex: "c=d" },
  ],
  "align-star",
);
assert.equal(alreadyNormalized.codeFormat, "align-star");
assert.deepEqual(alreadyNormalized.lines, [
  { id: "line-a", latex: "a=b" },
  { id: "line-b", latex: "c=d" },
]);

const embeddedEnvironment = normalize(
  String.raw`prefix \begin{align}a&=b\end{align} suffix`,
);
assert.equal(embeddedEnvironment.codeFormat, "raw");
assert.equal(embeddedEnvironment.lines.length, 1);

assert.equal(errorMessage({ message: "direct message" }, "fallback"), "direct message");
assert.equal(
  errorMessage({ error: { description: "nested description" } }, "fallback"),
  "nested description",
);
assert.equal(
  errorMessage({ details: { code: 7400, host: "word" } }, "fallback"),
  JSON.stringify({ code: 7400, host: "word" }),
);
assert.equal(errorMessage({ status: 500 }, "fallback"), '{"status":500}');
assert.equal(errorMessage(42, "fallback"), "42");

const cyclic = {};
cyclic.cause = cyclic;
const cyclicMessage = errorMessage(cyclic, "cyclic fallback");
assert.equal(cyclicMessage, "cyclic fallback");
assert.ok(!cyclicMessage.includes("[object Object]"));

for (const reason of [
  { message: "message" },
  { error: "error" },
  { description: "description" },
  { details: "details" },
  { arbitrary: { value: true } },
]) {
  assert.ok(!errorMessage(reason, "fallback").includes("[object Object]"));
}

console.log("Office formula editor regression passed");
