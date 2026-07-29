import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

const adapter = readFileSync(
  new URL("../office/macos-offline/word/VTWordAdapter.bas", import.meta.url),
  "utf8",
);
const ribbon = readFileSync(
  new URL("../office/macos-offline/word/customUI14.xml", import.meta.url),
  "utf8",
);

const start = adapter.indexOf("Private Sub VTDocumentImportInsertFormula");
const end = adapter.indexOf("Private Sub VTCancelWordDocumentImportDispatch", start);
assert.ok(start >= 0 && end > start, "Document formula insertion procedure is missing");
const insertion = adapter.slice(start, end);

for (const required of [
  "VTPrepareWordCreateInsertionRange",
  "VTInsertNativeEquationAtRange",
  "VTFinalizeInlineNativeEquation",
  "VTPromoteNativeEquationToDisplay",
  "VTEnsureNativeEquationNumber",
  "VTAddWordFormulaPicture",
  "VTNormalizeUnnumberedDisplayParagraph",
  "VTEnsureImageEquationNumber",
  "VTRefreshNumberedImageFormulaFontLayout",
  "VTPlaceCaretAfterDisplayFormula",
  "VTSetWordLatexPayload",
  "VTSetWordOmmlPayload",
  "VTSetWordMetadataPayload",
  "VTSetWordFormulaFormat",
  "VTSetWordImageScaleState",
]) {
  assert.ok(
    insertion.includes(required),
    `Document import must reuse the normal Word layout/state routine ${required}`,
  );
}

assert.match(
  insertion,
  /If displayMode = "inline" Then[\s\S]*candidate\.Range\.Font\.Position = CLng\(baselinePoints\)/,
  "Inline image formulas must preserve the normal VisualTeX baseline",
);
assert.match(
  insertion,
  /If numbered Then[\s\S]*VTEnsureNativeEquationNumber/,
  "Numbered native formulas must use the normal true-centering number layout",
);
assert.match(
  insertion,
  /If numbered Then[\s\S]*VTEnsureImageEquationNumber/,
  "Numbered image formulas must use the normal true-centering number layout",
);
assert.match(
  adapter,
  /insertedRange\.Font\.Reset[\s\S]*insertedRange\.Font\.Italic = False/,
  "Imported prose must not inherit italic formatting from the insertion caret",
);
for (const required of [
  "VTFinalizeDocumentImportParagraph",
  "wdStyleHeading1",
  "wdStyleHeading2",
  "ApplyBulletDefault",
  "ApplyNumberDefault",
  "paragraphRange.ParagraphFormat.SpaceAfter = 0!",
]) {
  assert.ok(
    adapter.includes(required),
    `Structured document import is missing Word formatting behavior ${required}`,
  );
}
assert.match(
  ribbon,
  /id="VisualTeX\.Mac\.Word\.DocumentImport"[\s\S]*onAction="VTWordRibbonDocumentImport"/,
  "Word Ribbon must expose the document import command",
);

console.log("Document import Word layout smoke test passed");
