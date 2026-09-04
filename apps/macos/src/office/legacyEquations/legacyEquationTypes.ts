export type LegacyJobStatus = "created" | "running" | "awaitingOmml" | "finalizing" | "complete" | "cancelled" | "failed";
export type LegacyRisk = "auto-replace" | "spot-check" | "manual-review" | "blocked";
export type LegacyProgId = "Equation.DSMT4" | "Equation.3" | "Equation.2" | "unknown";

export interface LegacyFormula {
  formulaId: string;
  partName: string;
  relationshipId: string;
  olePartName: string;
  progId: LegacyProgId;
  mtefVersion: 3 | 5 | null;
  displayMode: "inline" | "block";
  status: "detected" | "extracted" | "converted" | "replaced" | "preserved" | "failed";
  riskLevel: LegacyRisk;
  mathMl?: string;
  warnings: string[];
  errors: string[];
}

export interface FormulaBatch {
  protocolVersion: 1;
  jobId: string;
  batchIndex: number;
  batchCount: number;
  formulas: LegacyFormula[];
}

export interface ScanReport {
  protocolVersion: 1;
  jobId: string;
  detected: number;
  batchCount: number;
  batchSize: number;
  byProgId: Record<string, number>;
  byRiskLevel: Record<string, number>;
}

export interface LegacyJobView {
  protocolVersion: number;
  jobId: string;
  inputPath: string;
  outputPath: string;
  inputSha256: string;
  status: LegacyJobStatus;
  createdAtEpochSeconds: number;
  updatedAtEpochSeconds: number;
  error?: string | null;
  scanReport?: ScanReport | null;
  conversionReport?: Record<string, unknown> | null;
}

export interface OmmlFormulaResult {
  formulaId: string;
  status: "replaced" | "preserved";
  ommlBase64?: string;
  warnings: string[];
  errors: string[];
}

export interface OmmlBatch {
  protocolVersion: 1;
  jobId: string;
  batchIndex: number;
  batchCount: number;
  formulas: OmmlFormulaResult[];
}

const UUID_V4 = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;
export function assertFormulaBatch(value: unknown, jobId: string, expectedIndex: number): FormulaBatch {
  if (!value || typeof value !== "object") throw new Error("Malformed formula batch");
  const batch = value as FormulaBatch;
  if (batch.protocolVersion !== 1 || !UUID_V4.test(batch.jobId) || batch.jobId !== jobId ||
      batch.batchIndex !== expectedIndex || !Number.isInteger(batch.batchCount) || batch.batchCount < 0 ||
      !Array.isArray(batch.formulas)) throw new Error("Formula batch does not match the job");
  const ids = new Set<string>();
  for (const formula of batch.formulas) {
    if (!formula || typeof formula.formulaId !== "string" || ids.has(formula.formulaId) ||
        !["auto-replace", "spot-check", "manual-review", "blocked"].includes(formula.riskLevel) ||
        !Array.isArray(formula.warnings) || !Array.isArray(formula.errors)) {
      throw new Error("Formula batch contains an invalid or duplicate formula");
    }
    ids.add(formula.formulaId);
  }
  return batch;
}

export function defaultOutputPath(inputPath: string) {
  return inputPath.replace(/\.docx$/i, "") + "_VisualTeX_OMML.docx";
}
