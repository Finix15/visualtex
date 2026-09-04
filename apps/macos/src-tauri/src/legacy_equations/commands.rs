use super::{
    jobs::{job_root, parse_job_id, JobRecord, LegacyEquationState},
    worker::resolve_worker,
};
use serde::{Deserialize, Serialize};
use std::path::PathBuf;
use tauri::{AppHandle, Manager, State};
use uuid::Uuid;

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct LegacyEquationJobView {
    #[serde(flatten)]
    pub record: JobRecord,
    pub scan_report: Option<serde_json::Value>,
    pub conversion_report: Option<serde_json::Value>,
}

fn job_view(
    state: &LegacyEquationState,
    root: &std::path::Path,
    id: Uuid,
) -> Result<LegacyEquationJobView, String> {
    let record = state.get(root, id)?;
    Ok(LegacyEquationJobView {
        scan_report: state.read_optional_report(root, id, "scan-report.json")?,
        conversion_report: state.read_optional_report(root, id, "conversion-report.json")?,
        record,
    })
}

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
) -> Result<LegacyEquationJobView, String> {
    // Resolve first so production builds without a bundled worker fail before creating job state.
    let worker = resolve_worker(&app)?;
    let root = root(&app)?;
    let record = state.create(&root, &request.input_path, &request.output_path)?;
    if let Err(error) = state.start(&root, record.job_id, &worker) {
        let _ = state.delete(&root, record.job_id);
        return Err(error);
    }
    job_view(&state, &root, record.job_id)
}

#[tauri::command]
pub fn get_legacy_equation_job(
    app: AppHandle,
    state: State<'_, LegacyEquationState>,
    job_id: String,
) -> Result<LegacyEquationJobView, String> {
    let root = root(&app)?;
    job_view(&state, &root, parse_job_id(&job_id)?)
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
) -> Result<LegacyEquationJobView, String> {
    let worker = resolve_worker(&app)?;
    let root = root(&app)?;
    let id = parse_job_id(&job_id)?;
    state.run_finalize_worker(&root, id, &worker)?;
    state.finalize(&root, id)?;
    job_view(&state, &root, id)
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

#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct OpenLegacyEquationOutputRequest {
    pub job_id: String,
    pub target: String,
}

fn run_open(arguments: &[&str]) -> Result<(), String> {
    let status = std::process::Command::new("/usr/bin/open")
        .args(arguments)
        .status()
        .map_err(|error| format!("Unable to open file: {error}"))?;
    status
        .success()
        .then_some(())
        .ok_or_else(|| "macOS could not open the requested file".to_string())
}

#[tauri::command]
pub fn open_legacy_equation_output(
    app: AppHandle,
    state: State<'_, LegacyEquationState>,
    request: OpenLegacyEquationOutputRequest,
) -> Result<(), String> {
    let record = state.get(&root(&app)?, parse_job_id(&request.job_id)?)?;
    if !record.output_path.is_file() {
        return Err("Legacy-equation output is not available".to_string());
    }
    let path = record.output_path.to_string_lossy();
    match request.target.as_str() {
        "word" => run_open(&["-a", "Microsoft Word", &path]),
        "finder" => run_open(&["-R", &path]),
        _ => Err("Unsupported output open target".to_string()),
    }
}

#[tauri::command]
pub fn open_legacy_equation_report(
    app: AppHandle,
    state: State<'_, LegacyEquationState>,
    job_id: String,
) -> Result<(), String> {
    let root = root(&app)?;
    let id = parse_job_id(&job_id)?;
    state.get(&root, id)?;
    let report = state.report_path(&root, id)?;
    run_open(&[&report.to_string_lossy()])
}
