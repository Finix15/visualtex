use super::{
    jobs::{job_root, parse_job_id, JobRecord, LegacyEquationState},
    worker::resolve_worker,
};
use serde::Deserialize;
use std::path::PathBuf;
use tauri::{AppHandle, Manager, State};

fn root(app: &AppHandle) -> Result<PathBuf, String> {
    app.path()
        .app_data_dir()
        .map(|path| job_root(&path))
        .map_err(|error| error.to_string())
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct CreateLegacyEquationJobRequest {
    pub input_path: PathBuf,
    pub output_path: PathBuf,
}

#[tauri::command]
pub fn create_legacy_equation_job(
    app: AppHandle,
    state: State<'_, LegacyEquationState>,
    request: CreateLegacyEquationJobRequest,
) -> Result<JobRecord, String> {
    // Resolve first so production builds without a bundled worker fail before creating job state.
    let worker = resolve_worker(&app)?;
    let root = root(&app)?;
    let record = state.create(&root, &request.input_path, &request.output_path)?;
    if let Err(error) = state.start(&root, record.job_id, &worker) {
        let _ = state.delete(&root, record.job_id);
        return Err(error);
    }
    state.get(&root, record.job_id)
}

#[tauri::command]
pub fn get_legacy_equation_job(
    app: AppHandle,
    state: State<'_, LegacyEquationState>,
    job_id: String,
) -> Result<JobRecord, String> {
    state.get(&root(&app)?, parse_job_id(&job_id)?)
}

#[tauri::command]
pub fn read_legacy_equation_batch(
    app: AppHandle,
    state: State<'_, LegacyEquationState>,
    job_id: String,
    batch_index: u32,
) -> Result<String, String> {
    String::from_utf8(state.read_batch(&root(&app)?, parse_job_id(&job_id)?, batch_index)?)
        .map_err(|_| "Formula batch is not UTF-8".to_string())
}

#[tauri::command]
pub fn submit_legacy_omml_batch(
    app: AppHandle,
    state: State<'_, LegacyEquationState>,
    job_id: String,
    batch_index: u32,
    payload: String,
) -> Result<(), String> {
    state.submit_batch(
        &root(&app)?,
        parse_job_id(&job_id)?,
        batch_index,
        payload.as_bytes(),
    )
}

#[tauri::command]
pub fn finalize_legacy_equation_job(
    app: AppHandle,
    state: State<'_, LegacyEquationState>,
    job_id: String,
) -> Result<JobRecord, String> {
    let worker = resolve_worker(&app)?;
    let root = root(&app)?;
    let id = parse_job_id(&job_id)?;
    state.run_finalize_worker(&root, id, &worker)?;
    state.finalize(&root, id)
}

#[tauri::command]
pub fn cancel_legacy_equation_job(
    app: AppHandle,
    state: State<'_, LegacyEquationState>,
    job_id: String,
) -> Result<JobRecord, String> {
    state.cancel(&root(&app)?, parse_job_id(&job_id)?)
}

#[tauri::command]
pub fn delete_legacy_equation_job(
    app: AppHandle,
    state: State<'_, LegacyEquationState>,
    job_id: String,
) -> Result<(), String> {
    state.delete(&root(&app)?, parse_job_id(&job_id)?)
}
