import assert from "node:assert/strict";
import { latexToSvg, svgToBase64 } from "../src/export/runtime.ts";
import {
  OIINT_SIZE1_OVAL_WIDTH_EM,
  OIINT_SIZE2_OVAL_WIDTH_EM,
  OIIINT_SIZE1_OVAL_WIDTH_EM,
  OIIINT_SIZE2_OVAL_WIDTH_EM,
} from "../src/math/integralGlyphs.ts";
import { RARE_INTEGRAL_GLYPHS } from "../src/math/rareIntegralGlyphs.generated.ts";

const matrixRows = Array.from({ length: 10 }, (_, row) =>
  Array.from({ length: 10 }, (_, column) => `a_{${row + 1}${column + 1}}`).join("&"),
).join("\\\\");

const uncommonIntegralOperators = [
  ["iint", "∬"],
  ["iiint", "∭"],
  ["oint", "∮"],
  ["oiint", "∯"],
  ["oiiint", "∰"],
  ...RARE_INTEGRAL_GLYPHS.flatMap((glyph) =>
    [glyph.command, ...glyph.aliases].map((command) => [
      command,
      glyph.character,
    ]),
  ),
];
const uncommonIntegralByName = new Map(
  uncommonIntegralOperators.map(([command, character]) => [
    `uncommon-integral-${command}`,
    { command, character },
  ]),
);

const cases = [
  ["fraction", String.raw`\frac{a+b}{c+d}`],
  ["root", String.raw`\sqrt[n]{x^2+y^2}`],
  ["integral", String.raw`\int_{-\infty}^{\infty} e^{-x^2}\,\mathrm{d}x`],
  ["sum", String.raw`\sum_{i=1}^{n} i^2`],
  ["matrix", String.raw`\begin{pmatrix}${matrixRows}\end{pmatrix}`],
  ["chinese", String.raw`\text{测试}+\alpha`],
  ["multiline", "a=b+c\nd=e-f\ng=h"],
  [
    "multiline-inner-environment",
    `${String.raw`\begin{pmatrix}a&b\\c&d\end{pmatrix}`}\n${String.raw`\int_0^1 x^2\,\mathrm{d}x`}\n${String.raw`\frac{p}{q}`}`,
  ],
  ["long", Array.from({ length: 25 }, (_, index) => `x_{${index + 1}}`).join("+")],
  ...uncommonIntegralOperators.map(([command]) => [
    `uncommon-integral-${command}`,
    `\\${command}_{a}^{b} f(x)\\,\\mathrm{d}x`,
  ]),
];

for (const [name, latex] of cases) {
  const result = await latexToSvg(latex, {
    displayMode: true,
    fontSizePt: 14,
    paddingPx: 10,
    background: name === "root" ? "white" : "transparent",
  });
  assert.match(result.svg, /^<svg\b/);
  assert.match(result.svg, /\bviewBox=/);
  assert.ok(result.width > 0, `${name} width`);
  assert.ok(result.height > 0, `${name} height`);
  assert.ok((result.baseline ?? -1) >= 0, `${name} baseline`);
  assert.ok(!/<foreignObject\b/i.test(result.svg), `${name} foreignObject`);
  assert.ok(!/<link\b|@import\b/i.test(result.svg), `${name} external CSS`);
  assert.ok(
    !/\b(?:href|xlink:href)=["'](?!#|data:)[^"']+/i.test(result.svg),
    `${name} external href`,
  );
  assert.ok(!/url\(\s*["']?https?:/i.test(result.svg), `${name} remote CSS URL`);
  if (name !== "root") {
    assert.match(
      result.svg,
      /<rect\b[^>]*fill-opacity="0\.001"/,
      `${name} transparent PowerPoint hit target`,
    );
  }
  assert.equal(result.base64, svgToBase64(result.svg));
  const uncommonIntegral = uncommonIntegralByName.get(name);
  if (uncommonIntegral) {
    const codePoint = uncommonIntegral.character
      .codePointAt(0)
      .toString(16)
      .toUpperCase();
    assert.ok(!/merror/i.test(result.svg), `${name} MathJax error`);
    assert.ok(
      !result.svg.includes(uncommonIntegral.command),
      `${name} raw command leakage`,
    );
    assert.ok(
      result.svg.includes(`data-c="${codePoint}"`),
      `${name} operator glyph`,
    );
    assert.doesNotMatch(result.svg, /<text\b/i, `${name} font fallback`);
  }
  const decoded = new TextDecoder().decode(
    Uint8Array.from(atob(result.base64), (character) => character.charCodeAt(0)),
  );
  assert.equal(decoded, result.svg, `${name} UTF-8 base64 round trip`);
}

function italicFTranslateX(svg) {
  const match = svg.match(
    /<g data-mml-node="mi" transform="translate\(([-+\d.eE]+),[-+\d.eE]+\)"><use data-c="1D453"/,
  );
  assert.ok(match, "Expected an italic f operand in the SVG");
  return Number(match[1]);
}

function assertRareIntegralPath(svg, definition, displayMode, label) {
  const codePoint = definition.codePoint.toString(16).toUpperCase();
  const variantName = displayMode ? "LO" : "SO";
  const variant = displayMode ? definition.large : definition.small;
  const path = svg.match(
    new RegExp(
      `<path id="[^"]*-${variantName}-${codePoint}" d="([^"]+)"`,
    ),
  );
  assert.ok(path, `${label} bundled ${variantName} path`);
  assert.match(path[1], /^M.+Z$/, `${label} closed vector outline`);
  assert.ok(
    path[1].includes(variant.mathJaxPath.slice(0, 48)),
    `${label} generated STIX outline`,
  );
  assert.match(svg, new RegExp(`data-c="${codePoint}"`), `${label} data-c`);
  assert.doesNotMatch(svg, /<text\b/i, `${label} system-font fallback`);
  assert.doesNotMatch(svg, /merror/i, `${label} MathJax error`);
}

function assertUsableGeometry(result, label) {
  assert.ok(result.width > 0, `${label} width`);
  assert.ok(result.height > 0, `${label} height`);
  assert.ok(result.baseline >= 0, `${label} non-negative baseline`);
  assert.ok(result.baseline <= result.height, `${label} baseline within image`);
}

const rareIntegralCommands = RARE_INTEGRAL_GLYPHS.flatMap((definition) =>
  [definition.command, ...definition.aliases].map((command) => ({
    command,
    definition,
  })),
);
const rareIntegralFontSizePt = 14;
const rareIntegralFontSizePx = rareIntegralFontSizePt * (96 / 72);

for (const { command, definition } of rareIntegralCommands) {
  const modeResults = new Map();
  for (const displayMode of [false, true]) {
    const mode = displayMode ? "display" : "inline";
    const options = {
      displayMode,
      fontSizePt: rareIntegralFontSizePt,
      paddingPx: 0,
      background: "transparent",
    };
    const bare = latexToSvg(`\\${command} f`, options);
    const limits = latexToSvg(`\\${command}\\limits_{a}^{b}`, options);
    const noLimits = latexToSvg(`\\${command}\\nolimits_{a}^{b}`, options);
    const label = `\\${command} ${mode}`;
    const variant = displayMode ? definition.large : definition.small;

    for (const [layout, result] of [
      ["bare", bare],
      ["limits", limits],
      ["nolimits", noLimits],
    ]) {
      const resultLabel = `${label} ${layout}`;
      assertUsableGeometry(result, resultLabel);
      assertRareIntegralPath(result.svg, definition, displayMode, resultLabel);
    }

    const heightEm = bare.height / rareIntegralFontSizePx;
    assert.ok(
      heightEm > (displayMode ? 1.8 : 0.9) &&
        heightEm < (displayMode ? 2.6 : 1.5),
      `${label} normalized operator height`,
    );
    assert.ok(
      bare.baseline / bare.height > 0.45 && bare.baseline / bare.height < 0.85,
      `${label} plausible baseline`,
    );
    assert.ok(
      limits.height > noLimits.height * 1.25,
      `${label} explicit limits stack instead of becoming side scripts`,
    );

    const operandX = italicFTranslateX(bare.svg);
    const operatorSpacing =
      operandX - variant.advanceWidth - variant.italicCorrection;
    assert.ok(
      operandX > variant.bounds.xMax + 80,
      `${label} following token must not overlap the outline`,
    );
    assert.ok(
      operatorSpacing > 80 && operatorSpacing < 300,
      `${label} following-token spacing`,
    );
    modeResults.set(mode, bare);
  }

  assert.ok(
    modeResults.get("display").height > modeResults.get("inline").height * 1.7,
    `\\${command} display glyph must be substantially taller than inline`,
  );
}

const contourIntegralGeometry = [
  {
    command: "oiint",
    baseCommand: "iint",
    codePoint: "222F",
    size1WidthEm: OIINT_SIZE1_OVAL_WIDTH_EM,
    size2WidthEm: OIINT_SIZE2_OVAL_WIDTH_EM,
  },
  {
    command: "oiiint",
    baseCommand: "iiint",
    codePoint: "2230",
    size1WidthEm: OIIINT_SIZE1_OVAL_WIDTH_EM,
    size2WidthEm: OIIINT_SIZE2_OVAL_WIDTH_EM,
  },
];

for (const displayMode of [false, true]) {
  const options = {
    displayMode,
    fontSizePt: 14,
    paddingPx: 0,
    background: "transparent",
  };
  for (const geometry of contourIntegralGeometry) {
    const target = `\\${geometry.command}`;
    const base = `\\${geometry.baseCommand}`;
    const targetSvg = latexToSvg(target, options);
    const baseSvg = latexToSvg(base, options);
    assert.equal(targetSvg.width, baseSvg.width, `${target} natural width`);
    assert.equal(targetSvg.height, baseSvg.height, `${target} natural height`);
    assert.equal(targetSvg.baseline, baseSvg.baseline, `${target} baseline`);
    assert.doesNotMatch(targetSvg.svg, /<text\b/, `${target} font fallback`);
    assert.match(
      targetSvg.svg,
      new RegExp(`data-c="${geometry.codePoint}"`),
      `${target} registered MathJax path`,
    );

    const targetWithLimits = latexToSvg(`${target}_{a}^{b}`, options);
    const baseWithLimits = latexToSvg(`${base}_{a}^{b}`, options);
    assert.equal(
      targetWithLimits.width,
      baseWithLimits.width,
      `${target} limits width`,
    );
    assert.equal(
      targetWithLimits.height,
      baseWithLimits.height,
      `${target} limits height`,
    );
    assert.equal(
      targetWithLimits.baseline,
      baseWithLimits.baseline,
      `${target} limits baseline`,
    );
    assert.doesNotMatch(
      targetWithLimits.svg,
      /<text\b/,
      `${target} limits font fallback`,
    );

    const targetWithOperand = latexToSvg(`${target} f`, options);
    const baseWithOperand = latexToSvg(`${base} f`, options);
    const targetOperandX = italicFTranslateX(targetWithOperand.svg);
    assert.equal(
      targetOperandX,
      italicFTranslateX(baseWithOperand.svg),
      `${target} following-token position`,
    );
    assert.equal(
      targetWithOperand.width,
      baseWithOperand.width,
      `${target} following-token width`,
    );
    const ovalWidthEm = displayMode
      ? geometry.size2WidthEm
      : geometry.size1WidthEm;
    const operatorGap = targetOperandX - ovalWidthEm * 1000;
    assert.ok(
      operatorGap > 100 && operatorGap < 250,
      `${target} following token must not overlap its oval`,
    );
  }
}

const verticalMatrix = latexToSvg(
  String.raw`\begin{pmatrix}a&b\\c&d\end{pmatrix}`,
);
const verticalIntegral = latexToSvg(String.raw`\int_0^1 x^2\,\mathrm{d}x`);
const verticalDocument = latexToSvg(
  `${String.raw`\begin{pmatrix}a&b\\c&d\end{pmatrix}`}\n${String.raw`\int_0^1 x^2\,\mathrm{d}x`}`,
);
const siblingMatrices = latexToSvg(
  `${String.raw`\begin{pmatrix}a&b\\c&d\end{pmatrix}`}\n${String.raw`\begin{pmatrix}e&f\\g&h\end{pmatrix}`}`,
);
assert.ok(
  verticalDocument.height > Math.max(verticalMatrix.height, verticalIntegral.height) * 1.45,
  "Multiple formula rows, including an inner matrix environment, must stack vertically",
);
assert.ok(
  verticalDocument.width < verticalMatrix.width + verticalIntegral.width,
  "Vertical formula export width must not equal a horizontal concatenation",
);
assert.ok(
  siblingMatrices.height > verticalMatrix.height * 1.65,
  "Two sibling matrix formula rows must remain separate vertical rows",
);

assert.throws(
  () =>
    latexToSvg("", {
      displayMode: true,
      fontSizePt: 12,
      paddingPx: 8,
      background: "transparent",
    }),
  /Cannot export an empty formula/,
);

console.log(`SVG export smoke test passed (${cases.length} formula classes)`);
