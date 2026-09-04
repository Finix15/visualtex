import { readFileSync } from "node:fs";
import assert from "node:assert/strict";
import { assertFormulaBatch, defaultOutputPath } from "../src/office/legacyEquations/legacyEquationTypes.ts";
import { legacyEquationError } from "../src/office/legacyEquations/legacyEquationErrors.ts";

const jobId = "12345678-1234-4234-9234-123456789abc";
const formula = { formulaId: "f1", partName: "word/document.xml", relationshipId: "rId1", olePartName: "word/embeddings/ole.bin", progId: "Equation.DSMT4", mtefVersion: 5, displayMode: "inline", status: "extracted", riskLevel: "auto-replace", mathMl: "<math><mi>x</mi></math>", warnings: [], errors: [] };
const batch = assertFormulaBatch({ protocolVersion: 1, jobId, batchIndex: 0, batchCount: 1, formulas: [formula] }, jobId, 0);
assert.equal(batch.formulas.length, 1);
assert.throws(() => assertFormulaBatch({ ...batch, batchIndex: 1 }, jobId, 0));
assert.throws(() => assertFormulaBatch({ ...batch, formulas: [formula, formula] }, jobId, 0));
assert.throws(() => assertFormulaBatch({ ...batch, jobId: "../../escape" }, jobId, 0));
assert.equal(defaultOutputPath("/tmp/Công thức.docx"), "/tmp/Công thức_VisualTeX_OMML.docx");
assert.match(legacyEquationError("Output already exists", "vi"), /đã tồn tại/);
assert.match(legacyEquationError("Output already exists", "en"), /already exists/);
assert.doesNotMatch(legacyEquationError("<img src=x onerror=alert(1)>", "en"), /<img/);

const app = readFileSync("src/office/legacyEquations/LegacyEquationConverterApp.tsx", "utf8");
assert.ok(!app.includes("dangerouslySetInnerHTML"));
assert.ok(!/console\.(log|debug)\(/.test(app));
for (const value of ["auto-replace", "spot-check", "manual-review", "blocked", "role=\"progressbar\"", "aria-live=\"polite\"", "validation-failed", "Duplicate formula across batches", "onCloseRequested", "creating.current"]) assert.ok(app.includes(value), `missing contract ${value}`);
console.log("Legacy equation converter contract passed (schema, safety, locale, accounting, accessibility).")
