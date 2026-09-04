import { invokeTauri } from "../shared/tauriTransport";
import type { FormulaBatch, LegacyJobView, OmmlBatch } from "./legacyEquationTypes";

export const legacyEquationClient = {
  create(inputPath: string, outputPath: string) {
    return invokeTauri<LegacyJobView>("create_legacy_equation_job", { request: { inputPath, outputPath } });
  },
  createFromOfficeSession(sessionId: string, outputPath: string) {
    return invokeTauri<LegacyJobView>("create_legacy_equation_job_from_office_session", {
      request: { sessionId, outputPath },
    });
  },
  get(jobId: string) { return invokeTauri<LegacyJobView>("get_legacy_equation_job", { jobId }); },
  async readBatch(jobId: string, batchIndex: number) {
    const payload = await invokeTauri<string>("read_legacy_equation_batch", { jobId, batchIndex });
    return JSON.parse(payload) as FormulaBatch;
  },
  submitBatch(jobId: string, batchIndex: number, batch: OmmlBatch) {
    return invokeTauri<void>("submit_legacy_omml_batch", { jobId, batchIndex, payload: JSON.stringify(batch) });
  },
  finalize(jobId: string) { return invokeTauri<LegacyJobView>("finalize_legacy_equation_job", { jobId }); },
  cancel(jobId: string) { return invokeTauri<LegacyJobView>("cancel_legacy_equation_job", { jobId }); },
  delete(jobId: string) { return invokeTauri<void>("delete_legacy_equation_job", { jobId }); },
  openOutput(jobId: string, target: "word" | "finder") {
    return invokeTauri<void>("open_legacy_equation_output", { request: { jobId, target } });
  },
  openReport(jobId: string) { return invokeTauri<void>("open_legacy_equation_report", { jobId }); },
};
