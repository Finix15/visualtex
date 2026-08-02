import assert from "node:assert/strict";
import { latexToMathMl, latexToSvg } from "../src/export/runtime.ts";

const formula = String.raw`f_x:V\to\mathbb{R},\quad f_x(y):=\ip{x}{y}`;
const mathMl = latexToMathMl(formula, false);
const svg = latexToSvg(formula, {
  displayMode: false,
  fontSizePt: 11,
  paddingPx: 0,
  background: "transparent",
});

assert.doesNotMatch(mathMl, /\\ip\b/, "MathML must not retain the raw \\ip command");
assert.doesNotMatch(svg.svg, /\\ip\b/, "SVG must not retain the raw \\ip command");
assert.match(mathMl, /(?:⟨|&#x27E8;)/u, "inner product must contain an opening angle bracket");
assert.match(mathMl, /(?:⟩|&#x27E9;)/u, "inner product must contain a closing angle bracket");
assert.ok(svg.width > 0 && svg.height > 0, "SVG dimensions must be positive");

const aliasMathMl = latexToMathMl(String.raw`\innerproduct{u}{v}`, false);
assert.doesNotMatch(
  aliasMathMl,
  /\\innerproduct\b/,
  "MathML must not retain the raw \\innerproduct command",
);

console.log("VisualTeX LaTeX compatibility regression passed for \\ip and \\innerproduct");
