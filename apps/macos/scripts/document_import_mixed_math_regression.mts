import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import {
  findLatexFormulaSpans,
  parseLatexMarkdownDocument,
} from "../src/office/documentImport/documentImportParser";

const source = readFileSync(
  resolve("scripts/fixtures/document-import-mixed-math-20260810.txt"),
  "utf8",
);
const rawSpans = findLatexFormulaSpans(source, 11, "auto");
const blocks = parseLatexMarkdownDocument(source, "auto", 11);
const formulas = blocks.filter((block) => block.kind === "formula");
const textBlocks = blocks.filter((block) => block.kind === "text");

assert.ok(rawSpans.length >= 50, `expected at least 50 formulas, got ${rawSpans.length}`);
assert.equal(
  formulas.length,
  rawSpans.length,
  `document import changed the formula count (${formulas.length} vs ${rawSpans.length})`,
);

for (const span of rawSpans) {
  assert.ok(
    formulas.some(
      (formula) =>
        formula.kind === "formula" &&
        formula.displayMode === span.displayMode &&
        formula.latex.trim() === span.latex.trim(),
    ),
    `formula disappeared or changed during document parsing: ${span.sourceText}`,
  );
}

for (const block of textBlocks) {
  assert.notEqual(block.paragraphStyle, "code", "mixed math sample must not create code paragraphs");
  assert.doesNotMatch(block.text, /#\s*A\^\{-1\}/, "matrix inverse was rewritten as a Setext heading");
  assert.doesNotMatch(block.text, /\\begin\{(?:p|v)?matrix\}/, "matrix source leaked into prose");
  assert.doesNotMatch(block.text, /\\\[|\\\]/, "display-math delimiters leaked into prose");
}

assert.ok(
  formulas.some(
    (formula) => formula.kind === "formula" && formula.latex.includes("A^{-1}") && formula.latex.includes("\\begin{pmatrix}"),
  ),
  "matrix inverse formula must remain a formula block",
);
assert.ok(
  formulas.some(
    (formula) => formula.kind === "formula" && formula.latex.includes("\\mathbf{a}\\cdot\\mathbf{b}"),
  ),
  "vector dot-product formula must remain a formula block",
);

console.log(`VisualTeX mixed-math document import regression: PASS (${formulas.length} formulas)`);
