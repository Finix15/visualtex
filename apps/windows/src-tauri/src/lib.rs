use base64::{engine::general_purpose::STANDARD as BASE64_STANDARD, Engine as _};
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::collections::VecDeque;
use std::env;
use std::fs::{self, OpenOptions};
use std::io::{BufRead, BufReader, BufWriter, Write};
use std::path::{Path, PathBuf};
use std::process::{Child, ChildStdin, ChildStdout, Command, Stdio};
use std::sync::atomic::{AtomicU32, AtomicU64, Ordering};
use std::sync::{Arc, Mutex, TryLockError};
use std::thread;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};
#[cfg(windows)]
use std::os::windows::process::CommandExt;
use tauri::path::BaseDirectory;
use tauri::{AppHandle, Emitter, Manager, State};

mod ocr_install;
mod ocr_offline;
#[cfg(windows)]
mod ocr_python_bundle;
mod office;

use ocr_install::{
    append_install_log, cleanup_runtime_processes, cleanup_stale_process,
    decode_process_output, install_log_path, load_snapshot, run_logged_command, save_snapshot,
    CommandCapture, CommandLimits,
    InstallControl, InstallSnapshot, InstallState,
};

const PADDLE_VERSION: &str = "3.3.1";
const PADDLEOCR_VERSION: &str = "3.7.0";
const DEFAULT_OCR_MODEL: &str = "PP-FormulaNet_plus-M";
pub(crate) const MAX_IMAGE_BYTES: usize = 20 * 1024 * 1024;
const OCR_CANCELLED: &str = "OCR_CANCELLED";
const ALLOWED_MODELS: &[&str] = &[
    "PP-FormulaNet_plus-S",
    "PP-FormulaNet_plus-M",
    "PP-FormulaNet_plus-L",
    "PP-FormulaNet-S",
    "PP-FormulaNet-L",
];
const MAX_OCR_EVENTS: usize = 256;
const OCR_RUNTIME_STATUS_SCHEMA: u32 = 1;
const OCR_RUNTIME_STATUS_FILE: &str = "runtime-status.json";
const OCR_PREFERRED_MODEL_FILE: &str = "preferred-model.txt";
const PRIVATE_PYTHON_SITE_CUSTOMIZE: &str = include_str!("../ocr/private_sitecustomize.py");
const OCR_WORKER_WAIT_POLL_INTERVAL: Duration = Duration::from_millis(100);
const OCR_WORKER_WAIT_NOTICE_INTERVAL: Duration = Duration::from_secs(2);
#[cfg(windows)]
const CREATE_NO_WINDOW: u32 = 0x0800_0000;

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrEventRecord {
    id: u64,
    event: String,
    payload: Value,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrEventEnvelope {
    cursor: u64,
    events: Vec<OcrEventRecord>,
}

#[derive(Clone, Default)]
pub(crate) struct OcrEventBus {
    next_id: Arc<AtomicU64>,
    events: Arc<Mutex<VecDeque<OcrEventRecord>>>,
}

impl OcrEventBus {
    fn publish<T: Serialize>(&self, event: &str, payload: &T) {
        let Ok(payload) = serde_json::to_value(payload) else {
            return;
        };
        let id = self.next_id.fetch_add(1, Ordering::SeqCst) + 1;
        if let Ok(mut events) = self.events.lock() {
            events.push_back(OcrEventRecord {
                id,
                event: event.to_string(),
                payload,
            });
            while events.len() > MAX_OCR_EVENTS {
                events.pop_front();
            }
        }
    }

    pub(crate) fn poll(&self, cursor: u64, event: Option<&str>) -> OcrEventEnvelope {
        let events = self
            .events
            .lock()
            .map(|events| {
                events
                    .iter()
                    .filter(|item| item.id > cursor)
                    .filter(|item| event.is_none_or(|name| item.event == name))
                    .cloned()
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default();
        OcrEventEnvelope {
            cursor: self.next_id.load(Ordering::SeqCst),
            events,
        }
    }
}

#[derive(Clone)]
struct RuntimePaths {
    root: PathBuf,
    venv: PathBuf,
    python: PathBuf,
    input: PathBuf,
    processed: PathBuf,
    logs: PathBuf,
    cache: PathBuf,
    temp: PathBuf,
}

#[derive(Clone)]
pub(crate) struct OcrState {
    worker: Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: Arc<AtomicU32>,
    cancel_generation: Arc<AtomicU64>,
    runtime_status: Arc<Mutex<Option<OcrRuntimeStatus>>>,
    install_control: Arc<InstallControl>,
    desired_warmup_model: Arc<Mutex<Option<String>>>,
    events: OcrEventBus,
}

impl Default for OcrState {
    fn default() -> Self {
        Self {
            worker: Arc::new(Mutex::new(None)),
            worker_pid: Arc::new(AtomicU32::new(0)),
            cancel_generation: Arc::new(AtomicU64::new(0)),
            runtime_status: Arc::new(Mutex::new(None)),
            install_control: Arc::new(InstallControl::default()),
            desired_warmup_model: Arc::new(Mutex::new(None)),
            events: OcrEventBus::default(),
        }
    }
}

fn is_final_ocr_state_owner(worker: &Arc<Mutex<Option<OcrWorker>>>) -> bool {
    Arc::strong_count(worker) == 1
}

impl Drop for OcrState {
    fn drop(&mut self) {
        // OcrState is cloned into Tauri state and the Office companion server.
        // Destroying any temporary clone must not kill the shared OCR worker.
        // Only the final owner is allowed to terminate the process.
        if is_final_ocr_state_owner(&self.worker) {
            self.cancel_generation.fetch_add(1, Ordering::SeqCst);
            let _ = self.install_control.cancel();
            let _ = terminate_worker_process(&self.worker_pid);
        }
    }
}

impl OcrState {
    pub(crate) async fn runtime_status(
        &self,
        app: AppHandle,
        force_refresh: bool,
    ) -> Result<OcrRuntimeStatus, String> {
        let runtime_status = self.runtime_status.clone();
        let install_control = self.install_control.clone();
        let installing = install_control.is_running();
        if !force_refresh || installing {
            if let Some(status) = read_cached_runtime_status(&runtime_status)? {
                if !installing || status.installed {
                    return Ok(status);
                }
            }
        }

        tauri::async_runtime::spawn_blocking(move || {
            // Do not run the full import-tokenizers probe while an install
            // task is active. That would mix a previous partial environment
            // error into the current installation UI.
            let mut status = if should_run_full_runtime_probe(force_refresh, installing) {
                get_runtime_status_inner(&app)?
            } else {
                get_runtime_status_fast(&app)?
            };
            if installing && !status.installed {
                status.message = install_control
                    .snapshot()?
                    .map(|snapshot| snapshot.message)
                    .unwrap_or_else(|| "OCR installation is in progress".to_string());
            }
            write_cached_runtime_status(&runtime_status, Some(status.clone()))?;
            Ok(status)
        })
        .await
        .map_err(|error| format!("OCR runtime status task failed: {error}"))?
    }

    pub(crate) async fn install_runtime(&self, app: AppHandle) -> Result<OcrRuntimeStatus, String> {
        let lease = self.install_control.begin()?;
        let generation = lease.generation();
        let install_control = self.install_control.clone();
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        let runtime_status = self.runtime_status.clone();
        write_cached_runtime_status(&runtime_status, None)?;
        let install_app = app.clone();
        let status = tauri::async_runtime::spawn_blocking(move || {
            let _lease = lease;
            let status = install_runtime_inner(
                &install_app,
                &worker,
                &worker_pid,
                &install_control,
                generation,
            )?;
            write_cached_runtime_status(&runtime_status, Some(status.clone()))?;
            Ok::<OcrRuntimeStatus, String>(status)
        })
        .await
        .map_err(|error| format!("OCR installer task failed: {error}"))??;

        Ok(status)
    }

    pub(crate) async fn install_status(&self, app: AppHandle) -> Result<InstallSnapshot, String> {
        let control = self.install_control.clone();
        tauri::async_runtime::spawn_blocking(move || {
            let paths = runtime_paths(&app)?;
            let mut snapshot = control.snapshot()?.unwrap_or_else(|| {
                load_snapshot(&paths.root)
                    .unwrap_or_else(|| InstallSnapshot::new(&install_log_path(&paths.root)))
            });
            if reconcile_interrupted_install_snapshot(&mut snapshot, control.is_running()) {
                save_snapshot(&paths.root, &snapshot)?;
            }
            control.set_snapshot(snapshot.clone())?;
            Ok(snapshot)
        })
        .await
        .map_err(|error| format!("OCR installation status task failed: {error}"))?
    }

    pub(crate) fn cancel_install(&self) -> Result<(), String> {
        self.install_control.cancel().map(|_| ())
    }

    pub(crate) async fn recognize(
        &self,
        app: AppHandle,
        request: OcrImageRequest,
    ) -> Result<OcrRecognitionResult, String> {
        if ALLOWED_MODELS.contains(&request.model.as_str()) {
            {
                let mut desired = self
                    .desired_warmup_model
                    .lock()
                    .map_err(|_| "OCR warmup selection is unavailable".to_string())?;
                *desired = Some(request.model.clone());
            }
            if let Ok(paths) = runtime_paths(&app) {
                if let Err(error) = write_preferred_ocr_model(&paths, &request.model) {
                    eprintln!("Unable to persist the active OCR model: {error}");
                }
            }
        }
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        let cancel_generation = self.cancel_generation.clone();
        let runtime_status = self.runtime_status.clone();
        tauri::async_runtime::spawn_blocking(move || {
            run_recognition(
                &app,
                &worker,
                &worker_pid,
                &cancel_generation,
                &runtime_status,
                request,
            )
        })
        .await
        .map_err(|error| format!("OCR recognition task failed: {error}"))?
    }

    pub(crate) fn cancel(&self, app: &AppHandle) -> Result<(), String> {
        self.cancel_generation.fetch_add(1, Ordering::SeqCst);
        let terminate_result = terminate_worker_process(&self.worker_pid);

        // Killing the worker is the only reliable way to interrupt a Paddle
        // inference that is blocked inside native code. Rebuild and rewarm it
        // immediately in the background so cancelling one request does not
        // make the next request pay the full model startup cost again.
        schedule_worker_rewarm(
            app.clone(),
            self.worker.clone(),
            self.worker_pid.clone(),
            self.runtime_status.clone(),
            self.desired_warmup_model.clone(),
        );

        terminate_result.map(|_| ())
    }

    pub(crate) async fn restart(&self, app: AppHandle) -> Result<(), String> {
        self.cancel_generation.fetch_add(1, Ordering::SeqCst);
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        tauri::async_runtime::spawn_blocking(move || {
            stop_worker(&worker, &worker_pid)?;
            cleanup_worker_temp(&runtime_paths(&app)?)
        })
        .await
        .map_err(|error| format!("OCR restart task failed: {error}"))?
    }

    pub(crate) fn schedule_startup_warmup(&self, app: AppHandle) {
        let state = self.clone();
        let _ = tauri::async_runtime::spawn(async move {
            // Give the desktop UI enough time to publish the model stored in
            // localStorage. If it already requested a warmup, that task owns
            // model selection and this background fallback must not enqueue a
            // stale second model. Pure background launches still prewarm from
            // the persisted preference after the delay.
            tokio::time::sleep(Duration::from_secs(15)).await;
            let frontend_already_selected = state
                .desired_warmup_model
                .lock()
                .map(|desired| desired.is_some())
                .unwrap_or(false);
            if frontend_already_selected {
                return;
            }
            let status = state.runtime_status(app.clone(), false).await;
            if status.as_ref().is_ok_and(|status| status.installed) {
                let model = runtime_paths(&app)
                    .ok()
                    .and_then(|paths| read_preferred_ocr_model(&paths))
                    .unwrap_or_else(|| DEFAULT_OCR_MODEL.to_string());
                if let Err(error) = state.warmup_model(app, model).await {
                    eprintln!("Unable to prewarm the preferred OCR model: {error}");
                }
            }
        });
    }

    pub(crate) async fn warmup_model(
        &self,
        app: AppHandle,
        model: String,
    ) -> Result<(), String> {
        if !ALLOWED_MODELS.contains(&model.as_str()) {
            return Err(format!("Unsupported PP-FormulaNet model: {model}"));
        }
        {
            let mut desired = self
                .desired_warmup_model
                .lock()
                .map_err(|_| "OCR warmup selection is unavailable".to_string())?;
            *desired = Some(model.clone());
        }
        if let Ok(paths) = runtime_paths(&app) {
            if let Err(error) = write_preferred_ocr_model(&paths, &model) {
                eprintln!("Unable to persist the preferred OCR model: {error}");
            }
        }
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        let runtime_status = self.runtime_status.clone();
        let desired_warmup_model = self.desired_warmup_model.clone();
        tauri::async_runtime::spawn_blocking(move || {
            warmup_worker(
                &app,
                &worker,
                &worker_pid,
                &runtime_status,
                &desired_warmup_model,
                &model,
            )
        })
        .await
        .map_err(|error| format!("OCR warmup task failed: {error}"))?
    }

    pub(crate) async fn reset_runtime(&self, app: AppHandle) -> Result<OcrRuntimeStatus, String> {
        self.cancel_generation.fetch_add(1, Ordering::SeqCst);
        let _ = self.install_control.cancel();
        let install_control = self.install_control.clone();
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        let runtime_status = self.runtime_status.clone();
        tauri::async_runtime::spawn_blocking(move || {
            let deadline = Instant::now() + Duration::from_secs(15);
            while install_control.is_running() && Instant::now() < deadline {
                thread::sleep(Duration::from_millis(100));
            }
            if install_control.is_running() {
                return Err("OCR installation did not stop in time; reset was not started to avoid deleting files that are still in use".to_string());
            }
            write_cached_runtime_status(&runtime_status, None)?;
            stop_worker(&worker, &worker_pid)?;
            let paths = runtime_paths(&app)?;
            let _ = cleanup_stale_process(&paths.root);
            cleanup_runtime_processes(&paths.root)?;
            reset_runtime_contents(&paths.root)?;
            install_control.set_snapshot(InstallSnapshot::new(&install_log_path(&paths.root)))?;
            let status = get_runtime_status_fast(&app)?;
            write_cached_runtime_status(&runtime_status, Some(status.clone()))?;
            Ok(status)
        })
        .await
        .map_err(|error| format!("OCR reset task failed: {error}"))?
    }

    pub(crate) async fn install_optional_model(
        &self,
        app: AppHandle,
        package_path: PathBuf,
    ) -> Result<OcrRuntimeStatus, String> {
        self.cancel_generation.fetch_add(1, Ordering::SeqCst);
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        let runtime_status = self.runtime_status.clone();
        tauri::async_runtime::spawn_blocking(move || {
            stop_worker(&worker, &worker_pid)?;
            let paths = runtime_paths(&app)?;
            ocr_offline::install_optional_model_pack(&package_path, &paths.root)?;
            let status = get_runtime_status_fast(&app)?;
            write_cached_runtime_status(&runtime_status, Some(status.clone()))?;
            Ok(status)
        })
        .await
        .map_err(|error| format!("OCR model pack installation task failed: {error}"))?
    }

    pub(crate) async fn remove_optional_model(
        &self,
        app: AppHandle,
        model: String,
    ) -> Result<OcrRuntimeStatus, String> {
        self.cancel_generation.fetch_add(1, Ordering::SeqCst);
        let worker = self.worker.clone();
        let worker_pid = self.worker_pid.clone();
        let runtime_status = self.runtime_status.clone();
        tauri::async_runtime::spawn_blocking(move || {
            stop_worker(&worker, &worker_pid)?;
            let paths = runtime_paths(&app)?;
            ocr_offline::remove_optional_model(&paths.root, &model)?;
            let status = get_runtime_status_fast(&app)?;
            write_cached_runtime_status(&runtime_status, Some(status.clone()))?;
            Ok(status)
        })
        .await
        .map_err(|error| format!("OCR model removal task failed: {error}"))?
    }

    pub(crate) fn poll_events(&self, cursor: u64, event: Option<&str>) -> OcrEventEnvelope {
        self.events.poll(cursor, event)
    }
}

struct OcrWorker {
    child: Child,
    stdin: BufWriter<ChildStdin>,
    stdout: BufReader<ChildStdout>,
    pid_state: Arc<AtomicU32>,
    log_path: PathBuf,
    loaded_model: Option<String>,
}

impl Drop for OcrWorker {
    fn drop(&mut self) {
        let pid = self.child.id();
        let _ = self.child.kill();
        let _ = self.child.wait();
        let _ = self
            .pid_state
            .compare_exchange(pid, 0, Ordering::SeqCst, Ordering::SeqCst);
    }
}

fn read_worker_json<R: BufRead>(
    reader: &mut R,
    closed_message: &str,
    response_name: &str,
) -> Result<Value, String> {
    loop {
        let mut bytes = Vec::new();
        let count = reader
            .read_until(b'\n', &mut bytes)
            .map_err(|error| format!("Unable to read {response_name}: {error}"))?;
        if count == 0 {
            return Err(closed_message.to_string());
        }

        while bytes
            .last()
            .is_some_and(|byte| matches!(*byte, b'\n' | b'\r'))
        {
            bytes.pop();
        }
        if bytes.iter().all(u8::is_ascii_whitespace) {
            continue;
        }

        let first_non_whitespace = bytes
            .iter()
            .copied()
            .find(|byte| !byte.is_ascii_whitespace());
        if first_non_whitespace != Some(b'{') {
            // Native dependencies occasionally write diagnostics to stdout.
            // Ignore those lines so they cannot corrupt the JSON protocol.
            continue;
        }

        match serde_json::from_slice(&bytes) {
            Ok(value) => return Ok(value),
            Err(error) => {
                // This fallback prevents a legacy Windows code page from
                // crashing the reader. New workers always emit ASCII-safe JSON.
                let output = String::from_utf8_lossy(&bytes);
                return serde_json::from_str(output.trim()).map_err(|lossy_error| {
                    format!(
                        "{response_name} returned invalid JSON: {error}; UTF-8-lossy parse: {lossy_error}; output={output:?}"
                    )
                });
            }
        }
    }
}

impl OcrWorker {
    fn worker_failure(&mut self, message: impl AsRef<str>) -> String {
        let status = match self.child.try_wait() {
            Ok(Some(status)) => status.to_string(),
            Ok(None) => "still running".to_string(),
            Err(error) => format!("status unavailable: {error}"),
        };
        format!(
            "{}; worker_status={status}; log={}",
            message.as_ref(),
            self.log_path.display()
        )
    }

    fn send(&mut self, app: &AppHandle, payload: &Value) -> Result<Value, String> {
        if let Some(status) = self.child.try_wait().map_err(|error| error.to_string())? {
            return Err(format!(
                "OCR worker exited unexpectedly: {status}; log={}",
                self.log_path.display()
            ));
        }

        serde_json::to_writer(&mut self.stdin, payload)
            .map_err(|error| format!("Unable to encode OCR request: {error}"))?;
        self.stdin
            .write_all(b"\n")
            .map_err(|error| self.worker_failure(format!("Unable to send OCR request: {error}")))?;
        self.stdin.flush().map_err(|error| {
            self.worker_failure(format!("Unable to flush OCR request: {error}"))
        })?;

        loop {
            let response = read_worker_json(
                &mut self.stdout,
                "OCR worker closed its output stream",
                "OCR response",
            )
            .map_err(|error| self.worker_failure(error))?;
            if response.get("event").and_then(Value::as_str) == Some("progress") {
                let _ = app.emit("ocr-recognition-progress", &response);
                if let Some(state) = app.try_state::<OcrState>() {
                    state.events.publish("ocr-recognition-progress", &response);
                }
                continue;
            }
            return Ok(response);
        }
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrInstallProgress {
    stage: String,
    state: InstallState,
    percent: u8,
    message: String,
    detail: Option<String>,
    error: Option<String>,
    log_path: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrRuntimeStatus {
    installed: bool,
    python_path: Option<String>,
    python_version: Option<String>,
    paddle_version: Option<String>,
    paddleocr_version: Option<String>,
    runtime_path: String,
    offline_bundle_available: bool,
    installed_models: Vec<String>,
    default_model: String,
    message: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrImageRequest {
    pub(crate) bytes: Vec<u8>,
    pub(crate) extension: String,
    pub(crate) model: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct OcrFormulaResult {
    latex: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrRecognitionResult {
    model: String,
    elapsed_ms: u64,
    processed_width: u32,
    processed_height: u32,
    background_inverted: bool,
    background_luminance: f64,
    formulas: Vec<OcrFormulaResult>,
}

#[derive(Debug, Deserialize)]
struct WorkerRecognitionResponse {
    ok: bool,
    model: Option<String>,
    elapsed_ms: Option<u64>,
    processed_width: Option<u32>,
    processed_height: Option<u32>,
    background_inverted: Option<bool>,
    background_luminance: Option<f64>,
    formulas: Option<Vec<OcrFormulaResult>>,
    error: Option<String>,
    details: Option<String>,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct RuntimeProbe {
    #[serde(rename = "pythonVersion", alias = "python_version")]
    python_version: String,
    #[serde(rename = "paddleVersion", alias = "paddle_version")]
    paddle_version: String,
    #[serde(rename = "paddleocrVersion", alias = "paddleocr_version")]
    paddleocr_version: String,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
struct RuntimeStatusManifest {
    #[serde(rename = "schemaVersion", alias = "schema_version")]
    schema_version: u32,
    #[serde(rename = "pythonPath", alias = "python_path")]
    python_path: String,
    #[serde(rename = "pythonVersion", alias = "python_version")]
    python_version: String,
    #[serde(rename = "paddleVersion", alias = "paddle_version")]
    paddle_version: String,
    #[serde(rename = "paddleocrVersion", alias = "paddleocr_version")]
    paddleocr_version: String,
}

#[derive(Debug, Deserialize)]
struct PythonProbe {
    version: String,
    major: u8,
    minor: u8,
    bits: u8,
    machine: String,
    executable: String,
}

fn runtime_paths(app: &AppHandle) -> Result<RuntimePaths, String> {
    let root = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve application data directory: {error}"))?
        .join("ocr-runtime");
    let offline_python = if cfg!(windows) {
        root.join("python").join("python.exe")
    } else {
        root.join("python").join("bin").join("python3")
    };
    let venv = root.join("venv");
    let legacy_python = if cfg!(windows) {
        venv.join("Scripts").join("python.exe")
    } else {
        venv.join("bin").join("python")
    };
    let python = if offline_python.exists() || !legacy_python.exists() {
        offline_python
    } else {
        legacy_python.clone()
    };
    Ok(RuntimePaths {
        root: root.clone(),
        venv,
        python,
        input: root.join("input"),
        processed: root.join("processed"),
        logs: root.join("logs"),
        cache: root.join("cache"),
        temp: root.join("tmp"),
    })
}

fn read_preferred_ocr_model(paths: &RuntimePaths) -> Option<String> {
    let model = fs::read_to_string(paths.root.join(OCR_PREFERRED_MODEL_FILE))
        .ok()?
        .trim()
        .to_string();
    ALLOWED_MODELS
        .contains(&model.as_str())
        .then_some(model)
}

fn write_preferred_ocr_model(paths: &RuntimePaths, model: &str) -> Result<(), String> {
    if !ALLOWED_MODELS.contains(&model) {
        return Err(format!("Unsupported PP-FormulaNet model: {model}"));
    }
    fs::create_dir_all(&paths.root)
        .map_err(|error| format!("Unable to create OCR runtime directory: {error}"))?;
    fs::write(paths.root.join(OCR_PREFERRED_MODEL_FILE), model.as_bytes())
        .map_err(|error| format!("Unable to save the preferred OCR model: {error}"))
}

fn worker_script_path(app: &AppHandle) -> Result<PathBuf, String> {
    if let Ok(path) = app.path().resolve("ocr/worker.py", BaseDirectory::Resource) {
        if path.exists() {
            return Ok(path);
        }
    }

    #[cfg(debug_assertions)]
    {
        let development_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("ocr")
            .join("worker.py");
        if development_path.exists() {
            return Ok(development_path);
        }
    }

    Err("Unable to locate bundled OCR worker.py".to_string())
}

fn emit_progress(
    app: &AppHandle,
    stage: &str,
    percent: u8,
    message: impl Into<String>,
    detail: Option<String>,
) {
    let state = match stage {
        "complete" => InstallState::Complete,
        "verify" => InstallState::Verifying,
        "dependencies-installed" => InstallState::DependenciesInstalled,
        _ => InstallState::Installing,
    };
    let progress = OcrInstallProgress {
        stage: stage.to_string(),
        state,
        percent,
        message: message.into(),
        detail,
        error: None,
        log_path: runtime_paths(app)
            .ok()
            .map(|paths| install_log_path(&paths.root).display().to_string()),
    };
    let _ = app.emit("ocr-install-progress", &progress);
    if let Some(state) = app.try_state::<OcrState>() {
        state.events.publish("ocr-install-progress", &progress);
    }
}

fn publish_install_snapshot(
    app: &AppHandle,
    control: &InstallControl,
    paths: &RuntimePaths,
    snapshot: &mut InstallSnapshot,
) -> Result<(), String> {
    snapshot.touch();
    save_snapshot(&paths.root, snapshot)?;
    control.set_snapshot(snapshot.clone())?;
    let progress = OcrInstallProgress {
        stage: snapshot
            .current_step
            .clone()
            .unwrap_or_else(|| "status".to_string()),
        state: snapshot.state.clone(),
        percent: snapshot.percent,
        message: snapshot.message.clone(),
        detail: snapshot.detail.clone(),
        error: snapshot.error.clone(),
        log_path: Some(snapshot.log_path.clone()),
    };
    let _ = app.emit("ocr-install-progress", &progress);
    if let Some(state) = app.try_state::<OcrState>() {
        state.events.publish("ocr-install-progress", &progress);
    }
    Ok(())
}

fn set_install_step(
    app: &AppHandle,
    control: &InstallControl,
    paths: &RuntimePaths,
    snapshot: &mut InstallSnapshot,
    state: InstallState,
    step: Option<&str>,
    percent: u8,
    message: impl Into<String>,
    detail: Option<String>,
    error: Option<String>,
) -> Result<(), String> {
    snapshot.state = state;
    snapshot.current_step = step.map(str::to_string);
    snapshot.percent = percent;
    snapshot.message = message.into();
    snapshot.detail = detail;
    snapshot.error = error;
    publish_install_snapshot(app, control, paths, snapshot)
}

fn emit_recognition_progress(
    app: &AppHandle,
    request_id: &str,
    stage: &str,
    message: impl Into<String>,
    model: &str,
) {
    let progress = json!({
        "event": "progress",
        "id": request_id,
        "stage": stage,
        "message": message.into(),
        "model": model,
    });
    let _ = app.emit("ocr-recognition-progress", &progress);
    if let Some(state) = app.try_state::<OcrState>() {
        state.events.publish("ocr-recognition-progress", &progress);
    }
}

fn tail_text(value: &str, max_chars: usize) -> String {
    let total = value.chars().count();
    if total <= max_chars {
        return value.to_string();
    }
    value.chars().skip(total - max_chars).collect()
}

fn hide_windows_console(command: &mut Command) {
    #[cfg(windows)]
    {
        command.creation_flags(CREATE_NO_WINDOW);
    }
    #[cfg(not(windows))]
    {
        let _ = command;
    }
}

fn configure_python_environment(command: &mut Command) {
    command
        .env_remove("PYTHONPATH")
        .env_remove("PYTHONHOME")
        .env_remove("PYTHONUSERBASE")
        .env("PYTHONNOUSERSITE", "1")
        .env("PYTHONSAFEPATH", "1")
        .env("PYTHONUTF8", "1")
        .env("PYTHONIOENCODING", "utf-8")
        .env("PYTHONUNBUFFERED", "1");
}

fn python_command(program: &Path) -> Command {
    let mut command = Command::new(program);
    command.args(["-I", "-X", "utf8"]);
    configure_python_environment(&mut command);
    command
}

fn command_output_detailed(command: &mut Command, label: &str) -> Result<(String, String), String> {
    configure_python_environment(command);
    hide_windows_console(command);
    let output = command
        .output()
        .map_err(|error| format!("Unable to start {label}: {error}"))?;
    let stdout = decode_process_output(&output.stdout);
    let stderr = decode_process_output(&output.stderr);
    if output.status.success() {
        return Ok((stdout, stderr));
    }

    let combined = format!("stdout:\n{stdout}\nstderr:\n{stderr}");
    let tail = tail_text(&combined, 6000);
    Err(format!(
        "{label} failed with {}:\n{}",
        output.status,
        tail.trim()
    ))
}

fn command_output(command: &mut Command, label: &str) -> Result<String, String> {
    command_output_detailed(command, label).map(|(stdout, _)| stdout.trim().to_string())
}

fn probe_python(
    program: &Path,
    prefix_args: &[String],
    label: &str,
) -> Result<PythonProbe, String> {
    let script = r#"import json, platform, struct, sys; print(json.dumps({'version': platform.python_version(), 'major': sys.version_info.major, 'minor': sys.version_info.minor, 'bits': struct.calcsize('P') * 8, 'machine': platform.machine(), 'executable': sys.executable}))"#;
    let mut command = Command::new(program);
    command
        .args(prefix_args)
        .args(["-I", "-X", "utf8"])
        .arg("-c")
        .arg(script);
    configure_python_environment(&mut command);
    let output = command_output(&mut command, &format!("Python version check ({label})"))?;
    serde_json::from_str(&output)
        .map_err(|error| format!("Python returned an invalid version response: {error}"))
}

#[cfg(windows)]
fn find_system_python() -> Result<(PathBuf, PythonProbe), String> {
    let mut candidates: Vec<(String, PathBuf, Vec<String>)> = Vec::new();

    if let Ok(explicit) = env::var("VISUALTEX_PYTHON") {
        let explicit = explicit.trim().trim_matches('"');
        if !explicit.is_empty() {
            candidates.push((
                "VISUALTEX_PYTHON".to_string(),
                PathBuf::from(explicit),
                Vec::new(),
            ));
        }
    }

    // tokenizers 0.19.1 has Windows x64 wheels through Python 3.12, but not
    // Python 3.13. Keep the selection order explicit so a newer incompatible
    // system interpreter can never trigger a Rust source build at 84%.
    for minor in [12, 11, 10, 9] {
        for launcher in ["py.exe", "py"] {
            for selector in [format!("-V:3.{minor}"), format!("-3.{minor}")] {
                candidates.push((
                    format!("{launcher} {selector}"),
                    PathBuf::from(launcher),
                    vec![selector],
                ));
            }
        }
    }
    for name in [
        "python3.12",
        "python3.11",
        "python3.10",
        "python3.9",
        "python.exe",
        "python",
        "py.exe",
        "py",
    ] {
        candidates.push((name.to_string(), PathBuf::from(name), Vec::new()));
    }

    let mut failures = Vec::new();
    for (label, program, prefix_args) in candidates {
        match probe_python(&program, &prefix_args, &label) {
            Ok(probe) => {
                let machine = probe.machine.to_ascii_lowercase();
                let compatible = is_supported_windows_ocr_python(&probe)
                    && matches!(machine.as_str(), "x86_64" | "amd64" | "x64");
                if compatible && !probe.executable.trim().is_empty() {
                    return Ok((PathBuf::from(&probe.executable), probe));
                }
                failures.push(format!(
                    "{label}: Python {}, interpreter {}-bit, platform architecture {}",
                    probe.version, probe.bits, probe.machine
                ));
            }
            Err(error) => failures.push(format!("{label}: {error}")),
        }
    }

    Err(format!(
        "未检测到可用于 OCR 的 64 位 Python 3.9–3.12。建议安装 x64 Python 3.12，并启用 Python Launcher，然后完全重启 VisualTeX。Python 3.13 与本版本固定使用的 tokenizers 0.19.1 不兼容，安装器不会选择它，也不会回退到 Rust 源码编译。\n\nNo compatible 64-bit Python 3.9–3.12 interpreter was found. Install x64 Python 3.12 with the Python Launcher enabled, then restart VisualTeX. Python 3.13 is intentionally rejected because tokenizers 0.19.1 has no compatible Windows wheel and source compilation is disabled.\n\nChecked:\n{}",
        failures.join("\n")
    ))
}

fn is_supported_windows_ocr_python(probe: &PythonProbe) -> bool {
    probe.major == 3 && matches!(probe.minor, 9 | 10 | 11 | 12) && probe.bits == 64
}

fn runtime_status_manifest_path(paths: &RuntimePaths) -> PathBuf {
    paths.root.join(OCR_RUNTIME_STATUS_FILE)
}

fn write_runtime_status_manifest(paths: &RuntimePaths, probe: &RuntimeProbe) -> Result<(), String> {
    fs::create_dir_all(&paths.root)
        .map_err(|error| format!("Unable to create OCR runtime directory: {error}"))?;
    let manifest = RuntimeStatusManifest {
        schema_version: OCR_RUNTIME_STATUS_SCHEMA,
        python_path: paths.python.display().to_string(),
        python_version: probe.python_version.clone(),
        paddle_version: probe.paddle_version.clone(),
        paddleocr_version: probe.paddleocr_version.clone(),
    };
    let content = serde_json::to_vec_pretty(&manifest)
        .map_err(|error| format!("Unable to serialize OCR runtime status: {error}"))?;
    let target = runtime_status_manifest_path(paths);
    let temporary = target.with_extension("json.tmp");
    fs::write(&temporary, content)
        .map_err(|error| format!("Unable to cache OCR runtime status: {error}"))?;
    if target.exists() {
        fs::remove_file(&target)
            .map_err(|error| format!("Unable to replace OCR runtime status: {error}"))?;
    }
    fs::rename(&temporary, &target)
        .map_err(|error| format!("Unable to publish OCR runtime status: {error}"))
}

fn read_runtime_status_manifest(paths: &RuntimePaths) -> Option<RuntimeProbe> {
    let content = fs::read(runtime_status_manifest_path(paths)).ok()?;
    let manifest: RuntimeStatusManifest = serde_json::from_slice(&content).ok()?;
    if manifest.schema_version != OCR_RUNTIME_STATUS_SCHEMA
        || manifest.python_path != paths.python.display().to_string()
        || manifest.paddle_version != PADDLE_VERSION
        || manifest.paddleocr_version != PADDLEOCR_VERSION
    {
        return None;
    }
    Some(RuntimeProbe {
        python_version: manifest.python_version,
        paddle_version: manifest.paddle_version,
        paddleocr_version: manifest.paddleocr_version,
    })
}

fn read_pyvenv_python_version(paths: &RuntimePaths) -> Option<String> {
    let bundled_metadata = paths.root.join("python").join("visualtex-python.json");
    if let Ok(content) = fs::read(&bundled_metadata) {
        if let Ok(value) = serde_json::from_slice::<Value>(&content) {
            if let Some(version) = value.get("pythonVersion").and_then(Value::as_str) {
                if !version.trim().is_empty() {
                    return Some(version.trim().to_string());
                }
            }
        }
    }
    let candidates = [
        paths.venv.join("pyvenv.cfg"),
        paths.root.join("python").join("pyvenv.cfg"),
    ];
    for candidate in candidates {
        let Ok(content) = fs::read_to_string(candidate) else {
            continue;
        };
        for line in content.lines() {
            let Some((key, value)) = line.split_once('=') else {
                continue;
            };
            if matches!(key.trim(), "version" | "version_info") {
                let version = value
                    .trim()
                    .split('.')
                    .take(3)
                    .collect::<Vec<_>>()
                    .join(".");
                if !version.is_empty() {
                    return Some(version);
                }
            }
        }
    }
    None
}

fn active_python_environment_root(paths: &RuntimePaths) -> PathBuf {
    let private_root = paths.root.join("python");
    let python = normalized_windows_path(paths.python.display().to_string());
    let private = normalized_windows_path(private_root.display().to_string());
    if python.starts_with(&private) {
        private_root
    } else {
        paths.venv.clone()
    }
}

fn ensure_private_python_isolation(paths: &RuntimePaths) -> Result<(), String> {
    let private_root = paths.root.join("python");
    let active_root = active_python_environment_root(paths);
    if normalized_windows_path(active_root.display().to_string())
        != normalized_windows_path(private_root.display().to_string())
        || !paths.python.is_file()
    {
        return Ok(());
    }

    let site_packages = private_root.join("Lib").join("site-packages");
    fs::create_dir_all(&site_packages).map_err(|error| {
        format!(
            "Unable to create the private OCR Python site-packages directory: {error}"
        )
    })?;
    let target = site_packages.join("sitecustomize.py");
    if fs::read_to_string(&target).ok().as_deref() == Some(PRIVATE_PYTHON_SITE_CUSTOMIZE) {
        return Ok(());
    }
    let temporary = target.with_extension("py.tmp");
    fs::write(&temporary, PRIVATE_PYTHON_SITE_CUSTOMIZE.as_bytes()).map_err(|error| {
        format!("Unable to stage the private OCR Python isolation guard: {error}")
    })?;
    if target.exists() {
        fs::remove_file(&target).map_err(|error| {
            format!("Unable to replace the private OCR Python isolation guard: {error}")
        })?;
    }
    fs::rename(&temporary, &target).map_err(|error| {
        format!("Unable to activate the private OCR Python isolation guard: {error}")
    })
}

fn site_packages_candidates(paths: &RuntimePaths) -> Vec<PathBuf> {
    let mut candidates = vec![
        paths.venv.join("Lib").join("site-packages"),
        paths.venv.join("lib").join("site-packages"),
        paths.root.join("python").join("Lib").join("site-packages"),
        paths.root.join("python").join("lib").join("site-packages"),
    ];
    let python_lib = paths.root.join("python").join("lib");
    if let Ok(entries) = fs::read_dir(python_lib) {
        for entry in entries.flatten() {
            if entry.file_type().is_ok_and(|kind| kind.is_dir()) {
                candidates.push(entry.path().join("site-packages"));
            }
        }
    }
    candidates
}

fn dist_info_version(site_packages: &Path, distribution: &str) -> Option<String> {
    let prefix = format!("{}-", distribution.to_ascii_lowercase());
    let suffix = ".dist-info";
    for entry in fs::read_dir(site_packages).ok()?.flatten() {
        let name = entry.file_name().to_string_lossy().to_string();
        let lower = name.to_ascii_lowercase();
        if lower.starts_with(&prefix) && lower.ends_with(suffix) {
            let end = name.len().saturating_sub(suffix.len());
            return Some(name[prefix.len()..end].to_string());
        }
    }
    None
}

fn python_import_file_exists(site_packages: &Path, import_name: &str) -> bool {
    site_packages.join(import_name).is_dir()
        || site_packages.join(format!("{import_name}.py")).is_file()
        || fs::read_dir(site_packages).is_ok_and(|entries| {
            entries.flatten().any(|entry| {
                let name = entry.file_name().to_string_lossy().to_ascii_lowercase();
                name.starts_with(&format!("{import_name}.")) && name.ends_with(".pyd")
            })
        })
}

fn required_dependency_version(
    site_packages: &Path,
    distribution: &str,
    import_name: &str,
) -> Result<String, String> {
    if !python_import_file_exists(site_packages, import_name) {
        return Err(format!(
            "OCR dependency '{distribution}' is missing from the VisualTeX virtual environment. Resume installation from the {distribution} step."
        ));
    }
    dist_info_version(site_packages, distribution).ok_or_else(|| {
        format!(
            "OCR dependency '{distribution}' has no package metadata in the VisualTeX virtual environment. Resume installation from the {distribution} step."
        )
    })
}

fn validate_formula_dependency_files(site_packages: &Path) -> Result<(), String> {
    let tokenizers = required_dependency_version(site_packages, "tokenizers", "tokenizers")?;
    if cfg!(windows) && tokenizers != "0.19.1" {
        return Err(format!(
            "OCR dependency tokenizers has version {tokenizers}; VisualTeX 1.2.3 requires the precompiled Windows wheel tokenizers 0.19.1."
        ));
    }
    required_dependency_version(site_packages, "imagesize", "imagesize")?;
    required_dependency_version(site_packages, "ftfy", "ftfy")?;
    required_dependency_version(site_packages, "Wand", "wand")?;
    Ok(())
}

fn probe_runtime_from_files(paths: &RuntimePaths) -> Result<RuntimeProbe, String> {
    if !paths.python.is_file() {
        return Err("OCR Python executable is missing".to_string());
    }

    let python_version = read_pyvenv_python_version(paths)
        .ok_or_else(|| "OCR Python version cache is missing".to_string())?;
    for site_packages in site_packages_candidates(paths) {
        if !site_packages.is_dir()
            || !python_import_file_exists(&site_packages, "paddle")
            || !python_import_file_exists(&site_packages, "paddleocr")
        {
            continue;
        }
        let Some(paddle_version) = dist_info_version(&site_packages, "paddlepaddle") else {
            continue;
        };
        let Some(paddleocr_version) = dist_info_version(&site_packages, "paddleocr") else {
            continue;
        };
        if paddle_version != PADDLE_VERSION || paddleocr_version != PADDLEOCR_VERSION {
            return Err(format!(
                "OCR runtime version mismatch: PaddlePaddle {paddle_version}, PaddleOCR {paddleocr_version}; expected {PADDLE_VERSION} and {PADDLEOCR_VERSION}"
            ));
        }
        validate_formula_dependency_files(&site_packages)?;
        if let Some(probe) = read_runtime_status_manifest(paths) {
            if probe.python_version == python_version {
                return Ok(probe);
            }
        }
        return Ok(RuntimeProbe {
            python_version,
            paddle_version,
            paddleocr_version,
        });
    }
    Err("OCR package metadata is missing; run a forced check or repair the runtime".to_string())
}

fn parse_runtime_probe_output(stdout: &str, stderr: &str) -> Result<RuntimeProbe, String> {
    for line in stdout.lines().rev().map(str::trim).filter(|line| !line.is_empty()) {
        if let Ok(probe) = serde_json::from_str::<RuntimeProbe>(line) {
            return Ok(probe);
        }
    }
    Err(format!(
        "OCR runtime returned invalid version information. Expected pythonVersion/paddleVersion/paddleocrVersion (legacy snake_case is also accepted).\nRaw stdout:\n{}\nRaw stderr:\n{}",
        stdout.trim(),
        stderr.trim()
    ))
}

fn probe_runtime(paths: &RuntimePaths) -> Result<RuntimeProbe, String> {
    if !paths.python.exists() {
        return Err("OCR runtime is not installed".to_string());
    }
    ensure_private_python_isolation(paths)?;
    let script = r#"import json, platform; import paddle; import paddleocr; import tokenizers, imagesize, ftfy, wand; from importlib.metadata import version; from paddleocr import FormulaRecognition; print(json.dumps({'pythonVersion': platform.python_version(), 'paddleVersion': paddle.__version__, 'paddleocrVersion': version('paddleocr')}))"#;
    let (stdout, stderr) = command_output_detailed(
        python_command(&paths.python).arg("-c").arg(script),
        "OCR runtime verification",
    )?;
    let probe = parse_runtime_probe_output(&stdout, &stderr)?;
    if probe.paddle_version != PADDLE_VERSION || probe.paddleocr_version != PADDLEOCR_VERSION {
        return Err(format!(
            "OCR runtime version mismatch: PaddlePaddle {}, PaddleOCR {}; expected {} and {}\nRaw stdout:\n{}\nRaw stderr:\n{}",
            probe.paddle_version, probe.paddleocr_version, PADDLE_VERSION, PADDLEOCR_VERSION,
            stdout.trim(), stderr.trim()
        ));
    }
    write_runtime_status_manifest(paths, &probe)?;
    Ok(probe)
}

fn runtime_status_from_probe(
    app: &AppHandle,
    paths: &RuntimePaths,
    result: Result<RuntimeProbe, String>,
) -> OcrRuntimeStatus {
    let offline_bundle_available = ocr_offline::bundle_available(app);
    let installed_models = ocr_offline::installed_models(&paths.root);
    match result {
        Ok(probe) => {
            let default_cached = installed_models
                .iter()
                .any(|model| model == ocr_offline::OFFLINE_DEFAULT_MODEL);
            OcrRuntimeStatus {
                installed: true,
                python_path: Some(paths.python.display().to_string()),
                python_version: Some(probe.python_version),
                paddle_version: Some(probe.paddle_version),
                paddleocr_version: Some(probe.paddleocr_version),
                runtime_path: paths.root.display().to_string(),
                offline_bundle_available,
                installed_models,
                default_model: ocr_offline::OFFLINE_DEFAULT_MODEL.to_string(),
                message: if default_cached {
                    "PaddleOCR formula runtime and the default M model are ready".to_string()
                } else {
                    "PaddleOCR formula runtime is ready; the default M model will be prepared in the background".to_string()
                },
            }
        }
        Err(error) => OcrRuntimeStatus {
            installed: false,
            python_path: paths
                .python
                .exists()
                .then(|| paths.python.display().to_string()),
            python_version: None,
            paddle_version: None,
            paddleocr_version: None,
            runtime_path: paths.root.display().to_string(),
            offline_bundle_available,
            installed_models,
            default_model: ocr_offline::OFFLINE_DEFAULT_MODEL.to_string(),
            message: if offline_bundle_available {
                format!("Offline OCR package is ready to install. Current runtime: {error}")
            } else {
                error
            },
        },
    }
}

fn get_runtime_status_fast(app: &AppHandle) -> Result<OcrRuntimeStatus, String> {
    let paths = runtime_paths(app)?;
    Ok(runtime_status_from_probe(
        app,
        &paths,
        probe_runtime_from_files(&paths),
    ))
}

fn get_runtime_status_inner(app: &AppHandle) -> Result<OcrRuntimeStatus, String> {
    let paths = runtime_paths(app)?;
    Ok(runtime_status_from_probe(
        app,
        &paths,
        probe_runtime(&paths),
    ))
}

fn should_run_full_runtime_probe(force_refresh: bool, installing: bool) -> bool {
    force_refresh && !installing
}

fn reconcile_interrupted_install_snapshot(
    snapshot: &mut InstallSnapshot,
    installation_is_running: bool,
) -> bool {
    if installation_is_running
        || !matches!(
            snapshot.state,
            InstallState::Installing
                | InstallState::DependenciesInstalled
                | InstallState::Verifying
        )
    {
        return false;
    }

    let verification = matches!(snapshot.state, InstallState::Verifying)
        || snapshot.current_step.as_deref() == Some("verify");
    let step = snapshot
        .current_step
        .clone()
        .unwrap_or_else(|| "unknown".to_string());
    snapshot.state = if verification {
        InstallState::VerificationFailed
    } else {
        InstallState::InstallFailed
    };
    snapshot.message = "上一次 OCR 安装被中断".to_string();
    snapshot.detail = Some("已完成的依赖和步骤已保留，可从当前步骤继续安装".to_string());
    snapshot.error = Some(format!(
        "The previous VisualTeX process exited while OCR installation step '{step}' was active."
    ));
    snapshot.touch();
    true
}

fn read_cached_runtime_status(
    cache: &Arc<Mutex<Option<OcrRuntimeStatus>>>,
) -> Result<Option<OcrRuntimeStatus>, String> {
    cache
        .lock()
        .map(|status| status.clone())
        .map_err(|_| "OCR runtime status cache is unavailable".to_string())
}

fn write_cached_runtime_status(
    cache: &Arc<Mutex<Option<OcrRuntimeStatus>>>,
    status: Option<OcrRuntimeStatus>,
) -> Result<(), String> {
    let mut guard = cache
        .lock()
        .map_err(|_| "OCR runtime status cache is unavailable".to_string())?;
    *guard = status;
    Ok(())
}

fn cleanup_worker_temp(paths: &RuntimePaths) -> Result<(), String> {
    if paths.temp.exists() {
        fs::remove_dir_all(&paths.temp)
            .map_err(|error| format!("Unable to clean OCR temporary files: {error}"))?;
    }
    fs::create_dir_all(&paths.temp)
        .map_err(|error| format!("Unable to create OCR temporary directory: {error}"))
}

fn terminate_worker_process(worker_pid: &AtomicU32) -> Result<bool, String> {
    let pid = worker_pid.swap(0, Ordering::SeqCst);
    if pid == 0 {
        return Ok(false);
    }

    #[cfg(unix)]
    {
        let result = unsafe { libc::kill(pid as i32, libc::SIGKILL) };
        if result == 0 {
            return Ok(true);
        }
        let error = std::io::Error::last_os_error();
        if error.raw_os_error() == Some(libc::ESRCH) {
            return Ok(false);
        }
        return Err(format!("Unable to terminate OCR worker {pid}: {error}"));
    }

    #[cfg(windows)]
    {
        let mut command = Command::new("taskkill");
        command
            .arg("/PID")
            .arg(pid.to_string())
            .arg("/T")
            .arg("/F");
        hide_windows_console(&mut command);
        let status = command
            .status()
            .map_err(|error| format!("Unable to terminate OCR worker {pid}: {error}"))?;
        return Ok(status.success());
    }

    #[allow(unreachable_code)]
    Ok(false)
}

fn stop_worker(
    worker: &Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: &Arc<AtomicU32>,
) -> Result<(), String> {
    let terminate_result = terminate_worker_process(worker_pid);
    if let Ok(mut guard) = worker.lock() {
        guard.take();
    }
    terminate_result.map(|_| ())
}

fn remove_runtime_entry_once(path: &Path) -> std::io::Result<()> {
    let metadata = match fs::symlink_metadata(path) {
        Ok(metadata) => metadata,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(()),
        Err(error) => return Err(error),
    };
    if metadata.is_dir() && !metadata.file_type().is_symlink() {
        fs::remove_dir_all(path)
    } else {
        fs::remove_file(path)
    }
}

fn remove_runtime_entry_with_retry(path: &Path, runtime_root: &Path) -> Result<(), String> {
    let deadline = Instant::now() + Duration::from_secs(20);
    let mut attempt = 0_u32;
    let mut last_error = None;

    loop {
        attempt += 1;
        let _ = cleanup_stale_process(runtime_root);
        let _ = cleanup_runtime_processes(runtime_root);
        match remove_runtime_entry_once(path) {
            Ok(()) => return Ok(()),
            Err(error) => last_error = Some(error),
        }
        if Instant::now() >= deadline {
            let error = last_error
                .map(|error| error.to_string())
                .unwrap_or_else(|| "unknown Windows file access error".to_string());
            return Err(format!(
                "Unable to reset OCR runtime entry {} after {attempt} attempts. Another program may still be using this path. Close OCR-related Python processes or applications viewing files under the OCR runtime, then retry. Last error: {error}",
                path.display(),
            ));
        }
        let delay_ms = (attempt as u64 * 150).min(1_500);
        thread::sleep(Duration::from_millis(delay_ms));
    }
}

fn reset_runtime_contents(runtime_root: &Path) -> Result<(), String> {
    if !runtime_root.exists() {
        return Ok(());
    }
    let entries = fs::read_dir(runtime_root)
        .map_err(|error| format!("Unable to inspect OCR runtime before reset: {error}"))?
        .collect::<Result<Vec<_>, _>>()
        .map_err(|error| format!("Unable to enumerate OCR runtime before reset: {error}"))?;

    for entry in entries {
        if entry.file_name().to_string_lossy().eq_ignore_ascii_case("logs") {
            continue;
        }
        remove_runtime_entry_with_retry(&entry.path(), runtime_root)?;
    }
    fs::create_dir_all(runtime_root.join("logs"))
        .map_err(|error| format!("Unable to preserve OCR reset logs directory: {error}"))
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct InstalledPackageProbe {
    distribution: String,
    version: String,
    module_path: String,
    executable: String,
    python_version: String,
}

fn normalized_windows_path(value: impl AsRef<str>) -> String {
    value
        .as_ref()
        .replace('/', "\\")
        .trim_end_matches('\\')
        .to_ascii_lowercase()
}

fn command_limits_for_step(step: &str) -> CommandLimits {
    match step {
        "paddle" | "paddleocr" => CommandLimits {
            step_timeout: Duration::from_secs(30 * 60),
            idle_timeout: Duration::from_secs(6 * 60),
        },
        "venv" | "pip-bootstrap" => CommandLimits {
            step_timeout: Duration::from_secs(10 * 60),
            idle_timeout: Duration::from_secs(2 * 60),
        },
        _ => CommandLimits {
            step_timeout: Duration::from_secs(15 * 60),
            idle_timeout: Duration::from_secs(4 * 60),
        },
    }
}

fn install_command(
    paths: &RuntimePaths,
    control: &InstallControl,
    generation: u64,
    step: &str,
    command: &mut Command,
    label: &str,
) -> Result<CommandCapture, String> {
    command
        .env("TMP", &paths.temp)
        .env("TEMP", &paths.temp)
        .env("TMPDIR", &paths.temp)
        .env("PIP_DEFAULT_TIMEOUT", "30")
        .env("PIP_RETRIES", "2")
        .env("PIP_CACHE_DIR", paths.cache.join("pip"))
        .env("CARGO_TARGET_DIR", paths.temp.join("cargo-target"));
    run_logged_command(
        command,
        label,
        &paths.root,
        control,
        generation,
        command_limits_for_step(step),
    )
}

fn probe_package_for_install(
    paths: &RuntimePaths,
    control: &InstallControl,
    generation: u64,
    distribution: &str,
    import_name: &str,
    expected_version: Option<&str>,
) -> Result<InstalledPackageProbe, String> {
    let script = r#"import importlib, importlib.metadata as metadata, json, platform, sys
name, import_name = sys.argv[1], sys.argv[2]
module = importlib.import_module(import_name)
print(json.dumps({
  'distribution': name,
  'version': metadata.version(name),
  'modulePath': getattr(module, '__file__', '') or '',
  'executable': sys.executable,
  'pythonVersion': platform.python_version(),
}))"#;
    let mut command = python_command(&paths.python);
    command
        .arg("-c")
        .arg(script)
        .arg(distribution)
        .arg(import_name);
    let capture = install_command(
        paths,
        control,
        generation,
        "package-probe",
        &mut command,
        &format!("Verify Python package {distribution}"),
    )?;
    let json_line = capture
        .stdout
        .lines()
        .rev()
        .map(str::trim)
        .find(|line| line.starts_with('{'))
        .ok_or_else(|| {
            format!(
                "Package {distribution} verification returned no JSON. stdout:\n{}\nstderr:\n{}",
                capture.stdout, capture.stderr
            )
        })?;
    let probe: InstalledPackageProbe = serde_json::from_str(json_line).map_err(|error| {
        format!(
            "Package {distribution} verification returned invalid JSON: {error}\nstdout:\n{}\nstderr:\n{}",
            capture.stdout, capture.stderr
        )
    })?;
    validate_installed_package_probe(paths, distribution, expected_version, &probe)?;
    append_install_log(
        &paths.root,
        &format!(
            "verified package={distribution} version={} module={} executable={} python={}",
            probe.version, probe.module_path, probe.executable, probe.python_version
        ),
    )?;
    Ok(probe)
}

fn validate_installed_package_probe(
    paths: &RuntimePaths,
    distribution: &str,
    expected_version: Option<&str>,
    probe: &InstalledPackageProbe,
) -> Result<(), String> {
    if probe.distribution.to_ascii_lowercase() != distribution.to_ascii_lowercase() {
        return Err(format!(
            "Package verification returned distribution '{}' while '{}' was requested",
            probe.distribution, distribution
        ));
    }
    if let Some(expected) = expected_version {
        if probe.version != expected {
            return Err(format!(
                "Package {distribution} version mismatch: installed {}, expected {expected}",
                probe.version
            ));
        }
    }
    let expected_python = normalized_windows_path(paths.python.display().to_string());
    let actual_python = normalized_windows_path(&probe.executable);
    if actual_python != expected_python {
        return Err(format!(
            "Package {distribution} was verified with the wrong interpreter. Expected {}, actual {}",
            paths.python.display(), probe.executable
        ));
    }
    let environment_root = active_python_environment_root(paths);
    let expected_root = normalized_windows_path(environment_root.display().to_string());
    let module_path = normalized_windows_path(&probe.module_path);
    if module_path.is_empty() || !module_path.starts_with(&expected_root) {
        return Err(format!(
            "Package {distribution} is not installed inside the active VisualTeX OCR Python environment {}. modulePath={}",
            environment_root.display(), probe.module_path
        ));
    }
    Ok(())
}

fn package_probe_requires_install(result: &Result<InstalledPackageProbe, String>) -> bool {
    result.is_err()
}

fn pip_install_arguments(
    requirement: &str,
    only_binary: bool,
    cache_dir: &Path,
) -> Vec<String> {
    let mut arguments = vec![
        "-m".to_string(),
        "pip".to_string(),
        "install".to_string(),
        "--isolated".to_string(),
        "--index-url".to_string(),
        "https://pypi.org/simple".to_string(),
        "--cache-dir".to_string(),
        cache_dir.display().to_string(),
        "--disable-pip-version-check".to_string(),
        "--no-input".to_string(),
        "--progress-bar".to_string(),
        "raw".to_string(),
        "--timeout".to_string(),
        "30".to_string(),
        "--retries".to_string(),
        "2".to_string(),
        "--prefer-binary".to_string(),
    ];
    if only_binary {
        arguments.push("--only-binary=:all:".to_string());
    }
    arguments.push(requirement.to_string());
    arguments
}

fn pip_install_package(
    paths: &RuntimePaths,
    control: &InstallControl,
    generation: u64,
    step: &str,
    requirement: &str,
    only_binary: bool,
) -> Result<(), String> {
    let mut command = python_command(&paths.python);
    command.args(pip_install_arguments(
        requirement,
        only_binary,
        &paths.cache.join("pip"),
    ));
    install_command(
        paths,
        control,
        generation,
        step,
        &mut command,
        &format!("pip install {requirement}"),
    )?;
    Ok(())
}

fn pip_check_runtime(
    paths: &RuntimePaths,
    control: &InstallControl,
    generation: u64,
) -> Result<(), String> {
    let mut command = python_command(&paths.python);
    command.arg("-m").arg("pip").arg("check");
    install_command(
        paths,
        control,
        generation,
        "dependency-check",
        &mut command,
        "OCR private Python dependency check",
    )?;
    Ok(())
}

fn ensure_private_dependency_closure(
    paths: &RuntimePaths,
    control: &InstallControl,
    generation: u64,
) -> Result<(), String> {
    match pip_check_runtime(paths, control, generation) {
        Ok(()) => return Ok(()),
        Err(error) => {
            append_install_log(
                &paths.root,
                &format!(
                    "private OCR dependency closure requires repair; user site-packages remain disabled: {error}"
                ),
            )?;
        }
    }

    pip_install_package(
        paths,
        control,
        generation,
        "dependency-repair",
        &format!("paddlepaddle=={PADDLE_VERSION}"),
        true,
    )?;
    pip_install_package(
        paths,
        control,
        generation,
        "dependency-repair",
        &format!("paddleocr=={PADDLEOCR_VERSION}"),
        false,
    )?;
    pip_check_runtime(paths, control, generation).map_err(|error| {
        format!(
            "The VisualTeX private OCR Python environment still has missing or incompatible dependencies after isolated repair. User site-packages were not used. {error}"
        )
    })
}

fn ensure_package_step(
    app: &AppHandle,
    paths: &RuntimePaths,
    control: &InstallControl,
    generation: u64,
    snapshot: &mut InstallSnapshot,
    step: &str,
    percent: u8,
    message: &str,
    distribution: &str,
    import_name: &str,
    requirement: &str,
    expected_version: Option<&str>,
    only_binary: bool,
) -> Result<(), String> {
    set_install_step(
        app,
        control,
        paths,
        snapshot,
        InstallState::Installing,
        Some(step),
        percent,
        message,
        Some(format!("检查 {distribution} 是否已安装；已完成的步骤不会重复下载")),
        None,
    )?;
    let existing_probe = probe_package_for_install(
        paths,
        control,
        generation,
        distribution,
        import_name,
        expected_version,
    );
    if package_probe_requires_install(&existing_probe) {
        if let Err(error) = &existing_probe {
            append_install_log(
                &paths.root,
                &format!("package {distribution} requires installation or repair: {error}"),
            )?;
        }
        pip_install_package(
            paths,
            control,
            generation,
            step,
            requirement,
            only_binary,
        )?;
    }
    probe_package_for_install(
        paths,
        control,
        generation,
        distribution,
        import_name,
        expected_version,
    )?;
    snapshot.mark_step_complete(step);
    publish_install_snapshot(app, control, paths, snapshot)
}

fn probe_runtime_for_install(
    paths: &RuntimePaths,
    control: &InstallControl,
    generation: u64,
) -> Result<RuntimeProbe, String> {
    let script = r#"import json, platform, sys
import paddle, paddleocr, tokenizers, imagesize, ftfy, wand
from importlib.metadata import version
from paddleocr import FormulaRecognition
print(json.dumps({
  'pythonVersion': platform.python_version(),
  'paddleVersion': paddle.__version__,
  'paddleocrVersion': version('paddleocr'),
  'executable': sys.executable,
}))"#;
    let mut command = python_command(&paths.python);
    command.arg("-c").arg(script);
    let capture = install_command(
        paths,
        control,
        generation,
        "verify",
        &mut command,
        "OCR runtime verification",
    )?;
    let probe = parse_runtime_probe_output(&capture.stdout, &capture.stderr)?;
    if probe.paddle_version != PADDLE_VERSION || probe.paddleocr_version != PADDLEOCR_VERSION {
        return Err(format!(
            "OCR runtime version mismatch: PaddlePaddle {}, PaddleOCR {}; expected {} and {}\nstdout:\n{}\nstderr:\n{}",
            probe.paddle_version,
            probe.paddleocr_version,
            PADDLE_VERSION,
            PADDLEOCR_VERSION,
            capture.stdout,
            capture.stderr
        ));
    }
    write_runtime_status_manifest(paths, &probe)?;
    Ok(probe)
}

fn damaged_venv_python_error(error: &str) -> bool {
    let normalized = error.to_ascii_lowercase();
    !normalized.contains("cancelled")
        && !normalized.contains("timed out")
        && !normalized.contains("no stdout/stderr")
        && (normalized.contains("unable to start ocr virtual environment python check")
            || normalized.contains("ocr virtual environment python check failed with")
            || normalized.contains("returned invalid python information"))
}

fn should_replace_with_bundled_python(error: &str) -> bool {
    let normalized = error.to_ascii_lowercase();
    damaged_venv_python_error(error)
        || normalized.contains("unsupported visualtex ocr python")
        || normalized.contains("32-bit interpreter")
}

fn verify_venv_identity(
    paths: &RuntimePaths,
    control: &InstallControl,
    generation: u64,
) -> Result<PythonProbe, String> {
    ensure_private_python_isolation(paths)?;
    let script = r#"import json, os, platform, site, struct, sys
user_site = site.getusersitepackages()
user_sites = [user_site] if isinstance(user_site, str) else list(user_site)
normalized = lambda value: os.path.normcase(os.path.abspath(value))
user_site_on_path = any(normalized(path) == normalized(candidate) for path in sys.path for candidate in user_sites if candidate)
if site.ENABLE_USER_SITE or user_site_on_path:
    raise RuntimeError(f'VisualTeX OCR Python is not isolated from user site-packages: enabled={site.ENABLE_USER_SITE}, onPath={user_site_on_path}, userSites={user_sites}, sys.path={sys.path}')
print(json.dumps({'version': platform.python_version(), 'major': sys.version_info.major, 'minor': sys.version_info.minor, 'bits': struct.calcsize('P') * 8, 'machine': platform.machine(), 'executable': sys.executable}))"#;
    let mut command = python_command(&paths.python);
    command.arg("-c").arg(script);
    let capture = install_command(
        paths,
        control,
        generation,
        "python",
        &mut command,
        "OCR virtual environment Python check",
    )?;
    let probe: PythonProbe = serde_json::from_str(capture.stdout.trim()).map_err(|error| {
        format!(
            "OCR virtual environment returned invalid Python information: {error}\nstdout:\n{}\nstderr:\n{}",
            capture.stdout, capture.stderr
        )
    })?;
    if probe.bits != 64 {
        return Err(format!(
            "The existing VisualTeX OCR environment uses a 32-bit interpreter: Python {} ({}-bit, platform {}). It must be replaced by the bundled x64 Python 3.12 runtime.",
            probe.version, probe.bits, probe.machine
        ));
    }
    if probe.major != 3 || !matches!(probe.minor, 9 | 10 | 11 | 12) {
        return Err(format!(
            "Unsupported VisualTeX OCR Python {} ({}-bit). It must be replaced by the bundled x64 Python 3.12 runtime.",
            probe.version, probe.bits
        ));
    }
    let expected = normalized_windows_path(paths.python.display().to_string());
    let actual = normalized_windows_path(&probe.executable);
    if expected != actual {
        return Err(format!(
            "OCR virtual environment resolved the wrong interpreter. Expected {}, actual {}",
            paths.python.display(), probe.executable
        ));
    }
    Ok(probe)
}

#[cfg(windows)]
fn install_windows_runtime_inner(
    app: &AppHandle,
    initial_paths: &RuntimePaths,
    control: &InstallControl,
    generation: u64,
) -> Result<OcrRuntimeStatus, String> {
    let mut effective_paths = initial_paths.clone();
    let paths = &mut effective_paths;
    fs::create_dir_all(&paths.root)
        .map_err(|error| format!("Unable to create OCR runtime directory: {error}"))?;
    for directory in [
        &paths.input,
        &paths.processed,
        &paths.logs,
        &paths.cache,
        &paths.temp,
    ] {
        fs::create_dir_all(directory)
            .map_err(|error| format!("Unable to create OCR runtime directory: {error}"))?;
    }
    let residual_processes = cleanup_runtime_processes(&paths.root)?;
    let stale_removed = cleanup_stale_process(&paths.root)?;
    cleanup_worker_temp(paths)?;
    if !residual_processes.is_empty() || stale_removed {
        append_install_log(
            &paths.root,
            &format!(
                "terminated stale OCR installation process trees before starting a new task; residual_pids={residual_processes:?}; recorded_process_removed={stale_removed}"
            ),
        )?;
    }

    let mut snapshot = load_snapshot(&paths.root)
        .unwrap_or_else(|| InstallSnapshot::new(&install_log_path(&paths.root)));
    snapshot.error = None;
    snapshot.detail = None;
    set_install_step(
        app,
        control,
        paths,
        &mut snapshot,
        InstallState::Installing,
        Some("start"),
        1,
        "正在检查可恢复的 OCR 安装环境",
        Some("不会清空已成功安装的 PaddlePaddle、PaddleOCR 或其他依赖".to_string()),
        None,
    )?;

    let installation_result = (|| -> Result<OcrRuntimeStatus, String> {
        if paths.python.is_file() {
            if let Ok(probe) = probe_runtime_for_install(paths, control, generation) {
                snapshot.completed_steps = vec![
                    "venv",
                    "pip-bootstrap",
                    "paddle",
                    "paddleocr",
                    "tokenizers",
                    "imagesize",
                    "ftfy",
                    "wand",
                    "verify",
                ]
                .into_iter()
                .map(str::to_string)
                .collect();
                set_install_step(
                    app,
                    control,
                    paths,
                    &mut snapshot,
                    InstallState::Complete,
                    Some("complete"),
                    100,
                    "已检测到完整 OCR 环境，无需重复安装",
                    Some(format!(
                        "Python {} · Paddle {} · PaddleOCR {}",
                        probe.python_version, probe.paddle_version, probe.paddleocr_version
                    )),
                    None,
                )?;
                return Ok(runtime_status_from_probe(app, paths, Ok(probe)));
            }
        }

        let existing_python_probe = if paths.python.is_file() {
            match verify_venv_identity(paths, control, generation) {
                Ok(probe) => Some(probe),
                Err(error) if should_replace_with_bundled_python(&error) => {
                    append_install_log(
                        &paths.root,
                        &format!(
                            "existing VisualTeX OCR Python is damaged and will be replaced by the bundled x64 Python 3.12 runtime: {error}"
                        ),
                    )?;
                    None
                }
                Err(error) => return Err(error),
            }
        } else {
            append_install_log(
                &paths.root,
                "no usable VisualTeX OCR Python was found; installing the bundled private x64 Python 3.12 runtime",
            )?;
            None
        };

        let python_probe = if let Some(probe) = existing_python_probe {
            probe
        } else {
            let damaged_root = active_python_environment_root(paths);
            if damaged_root.exists() {
                fs::remove_dir_all(&damaged_root).map_err(|error| {
                    format!(
                        "Unable to remove damaged OCR Python environment {}: {error}",
                        damaged_root.display()
                    )
                })?;
            }
            let legacy_python = paths.venv.join("Scripts").join("python.exe");
            if paths.venv.exists() && !legacy_python.is_file() {
                fs::remove_dir_all(&paths.venv).map_err(|error| {
                    format!("Unable to remove incomplete legacy OCR virtual environment: {error}")
                })?;
            }
            set_install_step(
                app,
                control,
                paths,
                &mut snapshot,
                InstallState::Installing,
                Some("python"),
                5,
                "正在校验内置的 64 位 Python 3.12",
                Some("该解释器独立于系统 Python，不写入 PATH，也不受系统 32 位 Python 影响".to_string()),
                None,
            )?;
            set_install_step(
                app,
                control,
                paths,
                &mut snapshot,
                InstallState::Installing,
                Some("python-runtime"),
                12,
                "正在安装 VisualTeX 私有 Python 运行时",
                Some(paths.root.join("python").display().to_string()),
                None,
            )?;
            let private_root = paths.root.join("python");
            let manifest = ocr_python_bundle::install_bundle(app, &private_root)?;
            paths.python = private_root.join("python.exe");
            append_install_log(
                &paths.root,
                &format!(
                    "activated bundled private Python {} ({}) at {}",
                    manifest.python_version,
                    manifest.architecture,
                    paths.python.display()
                ),
            )?;
            snapshot.mark_step_complete("python-runtime");
            verify_venv_identity(paths, control, generation)?
        };
        snapshot.mark_step_complete("venv");
        snapshot.mark_step_complete("python-runtime");
        publish_install_snapshot(app, control, paths, &mut snapshot)?;

        set_install_step(
            app,
            control,
            paths,
            &mut snapshot,
            InstallState::Installing,
            Some("pip-bootstrap"),
            22,
            "正在确认 OCR 虚拟环境与 pip",
            Some(format!(
                "sys.executable={} · Python {}",
                python_probe.executable, python_probe.version
            )),
            None,
        )?;
        let mut pip_version_command = python_command(&paths.python);
        pip_version_command.arg("-m").arg("pip").arg("--version");
        let pip_capture = install_command(
            paths,
            control,
            generation,
            "pip-bootstrap",
            &mut pip_version_command,
            "OCR virtual environment pip check",
        )?;
        let environment_root = active_python_environment_root(paths);
        let expected_environment = normalized_windows_path(environment_root.display().to_string());
        if !normalized_windows_path(&pip_capture.stdout).contains(&expected_environment) {
            return Err(format!(
                "python -m pip resolved outside the active VisualTeX OCR Python environment. Expected path under {}, output: {}",
                environment_root.display(), pip_capture.stdout.trim()
            ));
        }
        append_install_log(
            &paths.root,
            &format!(
                "environment sys.executable={} python={} pip={}",
                python_probe.executable,
                python_probe.version,
                pip_capture.stdout.trim()
            ),
        )?;
        snapshot.mark_step_complete("pip-bootstrap");

        ensure_package_step(
            app,
            paths,
            control,
            generation,
            &mut snapshot,
            "paddle",
            38,
            &format!("正在检查或安装 PaddlePaddle {PADDLE_VERSION}"),
            "paddlepaddle",
            "paddle",
            &format!("paddlepaddle=={PADDLE_VERSION}"),
            Some(PADDLE_VERSION),
            true,
        )?;
        ensure_package_step(
            app,
            paths,
            control,
            generation,
            &mut snapshot,
            "paddleocr",
            66,
            &format!("正在检查或安装 PaddleOCR {PADDLEOCR_VERSION}"),
            "paddleocr",
            "paddleocr",
            &format!("paddleocr=={PADDLEOCR_VERSION}"),
            Some(PADDLEOCR_VERSION),
            false,
        )?;
        ensure_package_step(
            app,
            paths,
            control,
            generation,
            &mut snapshot,
            "tokenizers",
            82,
            "正在安装并验证 tokenizers 预编译 wheel",
            "tokenizers",
            "tokenizers",
            "tokenizers==0.19.1",
            Some("0.19.1"),
            true,
        )?;
        ensure_package_step(
            app,
            paths,
            control,
            generation,
            &mut snapshot,
            "imagesize",
            85,
            "正在安装并验证 imagesize",
            "imagesize",
            "imagesize",
            "imagesize",
            None,
            false,
        )?;
        ensure_package_step(
            app,
            paths,
            control,
            generation,
            &mut snapshot,
            "ftfy",
            88,
            "正在安装并验证 ftfy",
            "ftfy",
            "ftfy",
            "ftfy",
            None,
            false,
        )?;
        ensure_package_step(
            app,
            paths,
            control,
            generation,
            &mut snapshot,
            "wand",
            91,
            "正在安装并单独验证 Wand 导入",
            "Wand",
            "wand",
            "Wand",
            None,
            false,
        )?;

        set_install_step(
            app,
            control,
            paths,
            &mut snapshot,
            InstallState::Installing,
            Some("dependency-check"),
            92,
            "正在检查私有 Python 的完整依赖闭包",
            Some("用户级 site-packages 已禁用；若发现缺失依赖，只在 VisualTeX 私有目录中补装".to_string()),
            None,
        )?;
        ensure_private_dependency_closure(paths, control, generation)?;
        snapshot.mark_step_complete("dependency-check");
        publish_install_snapshot(app, control, paths, &mut snapshot)?;

        set_install_step(
            app,
            control,
            paths,
            &mut snapshot,
            InstallState::DependenciesInstalled,
            Some("dependencies-installed"),
            93,
            "PP-FormulaNet 依赖已全部安装并验证",
            Some("tokenizers、imagesize、ftfy 和 Wand 均来自当前 OCR 虚拟环境".to_string()),
            None,
        )?;
        set_install_step(
            app,
            control,
            paths,
            &mut snapshot,
            InstallState::Verifying,
            Some("verify"),
            94,
            "正在验证 PP-FormulaNet 接口",
            Some("No ccache found 等 Paddle 警告只写入日志，不作为失败原因".to_string()),
            None,
        )?;
        let probe = probe_runtime_for_install(paths, control, generation)?;
        snapshot.mark_step_complete("verify");
        let status = runtime_status_from_probe(app, paths, Ok(probe));
        set_install_step(
            app,
            control,
            paths,
            &mut snapshot,
            InstallState::Complete,
            Some("complete"),
            100,
            "OCR 运行环境安装完成",
            Some("已写入 runtime-status.json；下次启动将直接恢复已安装状态".to_string()),
            None,
        )?;
        Ok(status)
    })();

    match installation_result {
        Ok(status) => Ok(status),
        Err(error) => {
            let cancelled = error.to_ascii_lowercase().contains("cancelled");
            let verification = snapshot.current_step.as_deref() == Some("verify");
            let state = if cancelled {
                InstallState::Cancelled
            } else if verification {
                InstallState::VerificationFailed
            } else {
                InstallState::InstallFailed
            };
            let step = snapshot.current_step.clone();
            let failed_percent = snapshot.percent;
            set_install_step(
                app,
                control,
                paths,
                &mut snapshot,
                state,
                step.as_deref(),
                failed_percent,
                if cancelled {
                    "OCR 安装已取消"
                } else if verification {
                    "OCR 运行时验证失败"
                } else {
                    "OCR 依赖安装失败"
                },
                Some("已成功安装的环境和步骤会保留；可重试当前步骤或查看完整日志".to_string()),
                Some(error.clone()),
            )?;
            Err(error)
        }
    }
}

fn install_runtime_inner(
    app: &AppHandle,
    worker: &Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: &Arc<AtomicU32>,
    install_control: &InstallControl,
    install_generation: u64,
) -> Result<OcrRuntimeStatus, String> {
    stop_worker(worker, worker_pid)?;
    let paths = runtime_paths(app)?;

    #[cfg(windows)]
    {
        return install_windows_runtime_inner(
            app,
            &paths,
            install_control,
            install_generation,
        );
    }

    #[cfg(not(windows))]
    {
        let _ = (install_control, install_generation);
        ocr_offline::install_bundle(app, &paths.root, |stage, percent, message, detail| {
            emit_progress(app, stage, percent, message, detail);
        })?;

        emit_progress(app, "verify", 97, "正在验证离线 PP-FormulaNet 接口", None);
        let status = get_runtime_status_inner(app)?;
        if !status.installed {
            return Err(status.message);
        }
        if !status
            .installed_models
            .iter()
            .any(|model| model == ocr_offline::OFFLINE_DEFAULT_MODEL)
        {
            return Err("The bundled PP-FormulaNet M model was not installed".to_string());
        }
        emit_progress(
            app,
            "complete",
            100,
            "OCR 离线运行环境安装完成",
            Some("Python、PaddleOCR 与默认 M 模型均已内置，无需联网".to_string()),
        );
        Ok(status)
    }
}

fn spawn_worker(
    app: &AppHandle,
    paths: &RuntimePaths,
    worker_pid: Arc<AtomicU32>,
) -> Result<OcrWorker, String> {
    ensure_private_python_isolation(paths)?;
    let script = worker_script_path(app)?;
    fs::create_dir_all(&paths.logs)
        .map_err(|error| format!("Unable to create OCR log directory: {error}"))?;
    fs::create_dir_all(&paths.cache)
        .map_err(|error| format!("Unable to create OCR cache directory: {error}"))?;
    fs::create_dir_all(&paths.temp)
        .map_err(|error| format!("Unable to create OCR temporary directory: {error}"))?;
    let log_path = paths.logs.join("worker.log");
    let mut log_file = OpenOptions::new()
        .create(true)
        .append(true)
        .open(&log_path)
        .map_err(|error| format!("Unable to open OCR worker log: {error}"))?;
    writeln!(
        log_file,
        "\n===== VisualTeX OCR worker start: pid pending, unix_ms={} =====",
        SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|duration| duration.as_millis())
            .unwrap_or_default()
    )
    .map_err(|error| format!("Unable to initialize OCR worker log: {error}"))?;
    let log_file_error = log_file
        .try_clone()
        .map_err(|error| format!("Unable to clone OCR log handle: {error}"))?;

    let mut command = python_command(&paths.python);
    command
        .arg(&script)
        .env("PYTHONUNBUFFERED", "1")
        .env("PYTHONUTF8", "1")
        .env("PYTHONIOENCODING", "utf-8")
        .env("VISUALTEX_PARENT_PID", std::process::id().to_string())
        .env("PADDLE_PDX_MODEL_SOURCE", "BOS")
        .env("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")
        .env("PADDLE_PDX_CACHE_HOME", paths.cache.join("paddlex"))
        .env("PADDLE_HOME", paths.cache.join("paddle"))
        .env("XDG_CACHE_HOME", &paths.cache)
        .env("TMPDIR", &paths.temp)
        .env("TMP", &paths.temp)
        .env("TEMP", &paths.temp)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::from(log_file_error));
    #[cfg(not(windows))]
    command
        .env("VISUALTEX_OFFLINE_OCR", "1")
        .env("HF_HUB_OFFLINE", "1")
        .env("TRANSFORMERS_OFFLINE", "1")
        .env("MODELSCOPE_OFFLINE", "1");
    hide_windows_console(&mut command);
    let mut child = command
        .spawn()
        .map_err(|error| format!("Unable to start OCR worker: {error}"))?;
    worker_pid.store(child.id(), Ordering::SeqCst);

    let stdin = child
        .stdin
        .take()
        .ok_or_else(|| "OCR worker stdin is unavailable".to_string())?;
    let stdout = child
        .stdout
        .take()
        .ok_or_else(|| "OCR worker stdout is unavailable".to_string())?;
    let mut worker = OcrWorker {
        child,
        stdin: BufWriter::new(stdin),
        stdout: BufReader::new(stdout),
        pid_state: worker_pid,
        log_path: log_path.clone(),
        loaded_model: None,
    };

    let ready = read_worker_json(
        &mut worker.stdout,
        "OCR worker closed before sending its ready signal",
        "OCR worker ready signal",
    )
    .map_err(|error| format!("{error}; log={}", log_path.display()))?;
    if ready.get("event").and_then(Value::as_str) != Some("ready") {
        return Err(format!("Unexpected OCR worker ready response: {ready}"));
    }
    Ok(worker)
}

fn schedule_worker_rewarm(
    app: AppHandle,
    worker_state: Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: Arc<AtomicU32>,
    runtime_status: Arc<Mutex<Option<OcrRuntimeStatus>>>,
    desired_warmup_model: Arc<Mutex<Option<String>>>,
) {
    let _ = tauri::async_runtime::spawn(async move {
        let result = tauri::async_runtime::spawn_blocking(move || {
            let model = {
                let mut desired = desired_warmup_model
                    .lock()
                    .map_err(|_| "OCR warmup selection is unavailable".to_string())?;
                desired
                    .get_or_insert_with(|| DEFAULT_OCR_MODEL.to_string())
                    .clone()
            };

            // The recognition thread releases this lock as soon as the killed
            // worker pipe closes. A warmup task that was already running may
            // have recovered the worker itself, so preserve that live warmed
            // process instead of killing and loading it a second time.
            {
                let mut guard = worker_state
                    .lock()
                    .map_err(|_| "OCR worker lock is poisoned".to_string())?;
                let already_recovered = if let Some(worker) = guard.as_mut() {
                    worker.loaded_model.as_deref() == Some(model.as_str())
                        && worker
                            .child
                            .try_wait()
                            .map_err(|error| {
                                format!("Unable to inspect the recovered OCR worker: {error}")
                            })?
                            .is_none()
                } else {
                    false
                };
                if already_recovered {
                    return Ok(());
                }
                guard.take();
            }

            let paths = runtime_paths(&app)?;
            cleanup_worker_temp(&paths)?;
            warmup_worker(
                &app,
                &worker_state,
                &worker_pid,
                &runtime_status,
                &desired_warmup_model,
                &model,
            )
        })
        .await
        .map_err(|error| format!("OCR recovery task failed: {error}"))
        .and_then(|result| result);

        if let Err(error) = result {
            eprintln!("Unable to restore OCR prewarm after cancellation: {error}");
        }
    });
}

fn warmup_model_is_current(
    desired_warmup_model: &Arc<Mutex<Option<String>>>,
    model: &str,
) -> Result<bool, String> {
    desired_warmup_model
        .lock()
        .map(|desired| desired.as_deref() == Some(model))
        .map_err(|_| "OCR warmup selection is unavailable".to_string())
}

fn warmup_worker(
    app: &AppHandle,
    worker_state: &Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: &Arc<AtomicU32>,
    runtime_status: &Arc<Mutex<Option<OcrRuntimeStatus>>>,
    desired_warmup_model: &Arc<Mutex<Option<String>>>,
    model: &str,
) -> Result<(), String> {
    if !ALLOWED_MODELS.contains(&model) {
        return Err(format!("Unsupported PP-FormulaNet model: {model}"));
    }
    if !warmup_model_is_current(desired_warmup_model, model)? {
        return Ok(());
    }
    let paths = runtime_paths(app)?;
    let status = match read_cached_runtime_status(runtime_status)? {
        Some(status) => status,
        None => {
            let status = get_runtime_status_fast(app)?;
            write_cached_runtime_status(runtime_status, Some(status.clone()))?;
            status
        }
    };
    if !status.installed {
        return Err(format!("OCR runtime is not installed: {}", status.message));
    }

    {
        let mut guard = worker_state
            .lock()
            .map_err(|_| "OCR worker lock is poisoned".to_string())?;
        if !warmup_model_is_current(desired_warmup_model, model)? {
            return Ok(());
        }
        if guard
            .as_ref()
            .and_then(|worker| worker.loaded_model.as_deref())
            == Some(model)
        {
            return Ok(());
        }
        if guard.is_none() {
            *guard = Some(spawn_worker(app, &paths, worker_pid.clone())?);
        }
        let payload = json!({
            "id": format!("warmup-{}", std::process::id()),
            "action": "warmup",
            "model": model,
            "device": "cpu"
        });
        let first_result = guard
            .as_mut()
            .ok_or_else(|| "OCR worker failed to start".to_string())?
            .send(app, &payload);
        let response = match first_result {
            Ok(response) => response,
            Err(first_error) => {
                guard.take();
                *guard = Some(spawn_worker(app, &paths, worker_pid.clone())?);
                guard
                    .as_mut()
                    .ok_or_else(|| "OCR worker failed to restart".to_string())?
                    .send(app, &payload)
                    .map_err(|second_error| {
                        format!(
                            "OCR warmup worker failed twice. First: {first_error}. Second: {second_error}"
                        )
                    })?
            }
        };
        if response.get("ok").and_then(Value::as_bool) != Some(true) {
            let error = response
                .get("error")
                .and_then(Value::as_str)
                .unwrap_or("OCR model warmup failed");
            return Err(error.to_string());
        }
        if let Some(worker) = guard.as_mut() {
            worker.loaded_model = Some(model.to_string());
        }
    }

    let refreshed = get_runtime_status_fast(app)?;
    write_cached_runtime_status(runtime_status, Some(refreshed))?;
    Ok(())
}

fn run_recognition(
    app: &AppHandle,
    worker_state: &Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: &Arc<AtomicU32>,
    cancel_generation: &Arc<AtomicU64>,
    runtime_status: &Arc<Mutex<Option<OcrRuntimeStatus>>>,
    request: OcrImageRequest,
) -> Result<OcrRecognitionResult, String> {
    let request_generation = cancel_generation.load(Ordering::SeqCst);

    if request.bytes.is_empty() {
        return Err("The selected image is empty".to_string());
    }
    if request.bytes.len() > MAX_IMAGE_BYTES {
        return Err("The image is larger than the 20 MB limit".to_string());
    }
    if !ALLOWED_MODELS.contains(&request.model.as_str()) {
        return Err(format!(
            "Unsupported PP-FormulaNet model: {}",
            request.model
        ));
    }

    let paths = runtime_paths(app)?;
    if let Some(status) = read_cached_runtime_status(runtime_status)? {
        if !status.installed {
            return Err(format!("OCR runtime is not installed: {}", status.message));
        }
    }
    if !paths.python.exists() {
        return Err("OCR runtime is not installed: Python executable is missing".to_string());
    }
    fs::create_dir_all(&paths.input)
        .map_err(|error| format!("Unable to create OCR input directory: {error}"))?;
    fs::create_dir_all(&paths.processed)
        .map_err(|error| format!("Unable to create OCR processed directory: {error}"))?;

    let extension = request
        .extension
        .trim_start_matches('.')
        .to_ascii_lowercase();
    let allowed_extensions = ["png", "jpg", "jpeg", "webp", "bmp", "tif", "tiff"];
    if !allowed_extensions.contains(&extension.as_str()) {
        return Err(format!("Unsupported image type: .{extension}"));
    }

    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_err(|error| error.to_string())?
        .as_nanos();
    let request_id = format!("{}-{nonce}", std::process::id());
    let input_path = paths.input.join(format!("{request_id}.{extension}"));
    let processed_path = paths.processed.join(format!("{request_id}.png"));
    fs::write(&input_path, &request.bytes)
        .map_err(|error| format!("Unable to save OCR input image: {error}"))?;

    let payload = json!({
        "id": request_id,
        "action": "recognize",
        "image_path": input_path,
        "processed_path": processed_path,
        "model": request.model,
        "device": "cpu"
    });

    if cancel_generation.load(Ordering::SeqCst) != request_generation {
        let _ = fs::remove_file(&input_path);
        return Err(OCR_CANCELLED.to_string());
    }

    let response_result = (|| -> Result<Value, String> {
        let mut next_wait_notice = Instant::now();
        let mut guard = loop {
            match worker_state.try_lock() {
                Ok(guard) => break guard,
                Err(TryLockError::WouldBlock) => {
                    if cancel_generation.load(Ordering::SeqCst) != request_generation {
                        return Err(OCR_CANCELLED.to_string());
                    }
                    let now = Instant::now();
                    if now >= next_wait_notice {
                        let message = if request.model == "PP-FormulaNet_plus-L" {
                            "正在等待高精度 L 模型后台准备；首次使用可能需要下载约 731.5 MB 并完成模型初始化，准备完成后会自动继续识别".to_string()
                        } else {
                            format!(
                                "正在等待 {} 模型后台准备，完成后会自动继续识别",
                                request.model
                            )
                        };
                        emit_recognition_progress(
                            app,
                            &request_id,
                            "model-wait",
                            message,
                            &request.model,
                        );
                        next_wait_notice = now + OCR_WORKER_WAIT_NOTICE_INTERVAL;
                    }
                    thread::sleep(OCR_WORKER_WAIT_POLL_INTERVAL);
                }
                Err(TryLockError::Poisoned(_)) => {
                    return Err("OCR worker lock is poisoned".to_string());
                }
            }
        };

        if guard.is_none() {
            *guard = Some(spawn_worker(app, &paths, worker_pid.clone())?);
        }

        let first_result = guard
            .as_mut()
            .ok_or_else(|| "OCR worker failed to start".to_string())?
            .send(app, &payload);
        match first_result {
            Ok(response) => {
                if cancel_generation.load(Ordering::SeqCst) != request_generation {
                    return Err(OCR_CANCELLED.to_string());
                }
                if response.get("ok").and_then(Value::as_bool) == Some(true) {
                    if let Some(worker) = guard.as_mut() {
                        worker.loaded_model = Some(request.model.clone());
                    }
                }
                Ok(response)
            }
            Err(first_error) => {
                if cancel_generation.load(Ordering::SeqCst) != request_generation {
                    return Err(OCR_CANCELLED.to_string());
                }

                guard.take();
                *guard = Some(spawn_worker(app, &paths, worker_pid.clone())?);
                let response = guard
                    .as_mut()
                    .ok_or_else(|| "OCR worker failed to restart".to_string())?
                    .send(app, &payload)
                    .map_err(|second_error| {
                        format!(
                            "OCR worker failed twice. First: {first_error}. Second: {second_error}"
                        )
                    })?;
                if cancel_generation.load(Ordering::SeqCst) != request_generation {
                    return Err(OCR_CANCELLED.to_string());
                }
                if response.get("ok").and_then(Value::as_bool) == Some(true) {
                    if let Some(worker) = guard.as_mut() {
                        worker.loaded_model = Some(request.model.clone());
                    }
                }
                Ok(response)
            }
        }
    })();

    let _ = fs::remove_file(&input_path);
    let _ = fs::remove_file(&processed_path);
    let response_value = response_result?;

    let response: WorkerRecognitionResponse = serde_json::from_value(response_value)
        .map_err(|error| format!("Unable to decode OCR result: {error}"))?;
    if !response.ok {
        let mut message = response
            .error
            .unwrap_or_else(|| "PP-FormulaNet recognition failed".to_string());
        if let Some(details) = response.details {
            if !details.trim().is_empty() {
                let detail_tail = tail_text(&details, 3000);
                message.push('\n');
                message.push_str(detail_tail.trim());
            }
        }
        return Err(message);
    }

    let formulas = response.formulas.unwrap_or_default();
    if formulas.is_empty() {
        return Err("PP-FormulaNet returned no formulas".to_string());
    }

    Ok(OcrRecognitionResult {
        model: response.model.unwrap_or_else(|| request.model.clone()),
        elapsed_ms: response.elapsed_ms.unwrap_or_default(),
        processed_width: response.processed_width.unwrap_or_default(),
        processed_height: response.processed_height.unwrap_or_default(),
        background_inverted: response.background_inverted.unwrap_or(false),
        background_luminance: response.background_luminance.unwrap_or(255.0),
        formulas,
    })
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ExportFileRequest {
    path: String,
    text: Option<String>,
    base64: Option<String>,
}

fn write_export_file_request(request: ExportFileRequest) -> Result<(), String> {
    let trimmed_path = request.path.trim();
    if trimmed_path.is_empty() {
        return Err("Export path is empty".to_string());
    }

    let path = PathBuf::from(trimmed_path);
    if path.file_name().is_none() {
        return Err("Export path must include a file name".to_string());
    }
    if let Some(parent) = path.parent().filter(|parent| !parent.as_os_str().is_empty()) {
        fs::create_dir_all(parent)
            .map_err(|error| format!("Unable to create export directory: {error}"))?;
    }

    let bytes = match (request.text, request.base64) {
        (Some(text), None) => text.into_bytes(),
        (None, Some(encoded)) => BASE64_STANDARD
            .decode(encoded.trim())
            .map_err(|error| format!("Unable to decode export data: {error}"))?,
        (Some(_), Some(_)) => {
            return Err("Export request must contain text or base64 data, not both".to_string())
        }
        (None, None) => return Err("Export request contains no data".to_string()),
    };

    fs::write(&path, bytes)
        .map_err(|error| format!("Unable to write export file '{}': {error}", path.display()))
}

#[tauri::command]
fn write_export_file(request: ExportFileRequest) -> Result<(), String> {
    write_export_file_request(request)
}

#[tauri::command]
async fn get_ocr_runtime_status(
    app: AppHandle,
    state: State<'_, OcrState>,
    force_refresh: Option<bool>,
) -> Result<OcrRuntimeStatus, String> {
    state
        .runtime_status(app, force_refresh.unwrap_or(false))
        .await
}

#[tauri::command]
async fn install_ocr_runtime(
    app: AppHandle,
    state: State<'_, OcrState>,
) -> Result<OcrRuntimeStatus, String> {
    state.install_runtime(app).await
}

#[tauri::command]
async fn get_ocr_install_status(
    app: AppHandle,
    state: State<'_, OcrState>,
) -> Result<InstallSnapshot, String> {
    state.install_status(app).await
}

#[tauri::command]
fn cancel_ocr_install(state: State<'_, OcrState>) -> Result<(), String> {
    state.cancel_install()
}

#[tauri::command]
fn open_ocr_install_logs(app: AppHandle) -> Result<(), String> {
    let paths = runtime_paths(&app)?;
    fs::create_dir_all(&paths.logs)
        .map_err(|error| format!("Unable to create OCR installation log directory: {error}"))?;
    #[cfg(windows)]
    {
        let mut command = Command::new("explorer.exe");
        command.arg(&paths.logs);
        hide_windows_console(&mut command);
        return command
            .spawn()
            .map(|_| ())
            .map_err(|error| format!("Unable to open OCR installation logs: {error}"));
    }
    #[cfg(not(windows))]
    {
        let _ = paths;
        Err("Opening OCR installation logs from the app is currently available only on Windows".to_string())
    }
}

#[tauri::command]
async fn recognize_formula_image(
    app: AppHandle,
    state: State<'_, OcrState>,
    request: OcrImageRequest,
) -> Result<OcrRecognitionResult, String> {
    state.recognize(app, request).await
}

#[tauri::command]
fn cancel_ocr_recognition(app: AppHandle, state: State<'_, OcrState>) -> Result<(), String> {
    state.cancel(&app)
}

#[tauri::command]
async fn restart_ocr_worker(app: AppHandle, state: State<'_, OcrState>) -> Result<(), String> {
    state.restart(app).await
}

#[tauri::command]
async fn warmup_ocr_model(
    app: AppHandle,
    state: State<'_, OcrState>,
    model: String,
) -> Result<(), String> {
    state.warmup_model(app, model).await
}

#[tauri::command]
async fn reset_ocr_runtime(
    app: AppHandle,
    state: State<'_, OcrState>,
) -> Result<OcrRuntimeStatus, String> {
    state.reset_runtime(app).await
}

#[tauri::command]
async fn install_optional_ocr_model(
    app: AppHandle,
    state: State<'_, OcrState>,
    package_path: String,
) -> Result<OcrRuntimeStatus, String> {
    let package_path = package_path.trim();
    if package_path.is_empty() {
        return Err("No OCR model package was selected".to_string());
    }
    state
        .install_optional_model(app, PathBuf::from(package_path))
        .await
}

#[tauri::command]
async fn remove_optional_ocr_model(
    app: AppHandle,
    state: State<'_, OcrState>,
    model: String,
) -> Result<OcrRuntimeStatus, String> {
    state.remove_optional_model(app, model).await
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let background_mode = office::background::is_background_mode();
    let ocr_state = OcrState::default();
    let office_ocr_state = ocr_state.clone();
    let app = tauri::Builder::default()
        .plugin(tauri_plugin_single_instance::init(
            |app, arguments, _cwd| {
                if arguments
                    .iter()
                    .any(|argument| argument == office::background::BACKGROUND_ARGUMENT)
                {
                    return;
                }
                let _ = office::background::reveal_main_window(app);
            },
        ))
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .manage(ocr_state)
        .setup(move |app| {
            let office_state = office::initialize(app.handle(), office_ocr_state.clone())
                .map_err(std::io::Error::other)?;
            if let Err(error) = office::powerpoint_native::start_double_click_monitor(
                office_state.powerpoint_interactions.clone(),
            ) {
                eprintln!("Unable to start PowerPoint double-click monitor: {error}");
            }
            app.manage(office_state.clone());
            office::start(office_state);
            // Start loading the recommended M model as soon as the background
            // service is ready. This runs even before the desktop OCR dialog is
            // opened, so normal application startup time overlaps model setup.
            office_ocr_state.schedule_startup_warmup(app.handle().clone());
            if background_mode {
                office::background::hide_main_window(app.handle())
                    .map_err(std::io::Error::other)?;
            } else {
                if let Err(error) = office::background::resume_installed_launch_agent() {
                    eprintln!("Unable to resume VisualTeX Office background service: {error}");
                }
                office::background::reveal_main_window(app.handle())
                    .map_err(std::io::Error::other)?;
            }
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            write_export_file,
            get_ocr_runtime_status,
            install_ocr_runtime,
            get_ocr_install_status,
            cancel_ocr_install,
            open_ocr_install_logs,
            recognize_formula_image,
            cancel_ocr_recognition,
            restart_ocr_worker,
            warmup_ocr_model,
            reset_ocr_runtime,
            install_optional_ocr_model,
            remove_optional_ocr_model,
            office::lifecycle::set_app_theme,
            office::lifecycle::set_app_editor_layout,
            office::lifecycle::set_powerpoint_default_font_size,
            office::lifecycle::get_powerpoint_default_font_size,
            office::lifecycle::get_office_companion_status,
            office::lifecycle::start_office_companion,
            office::lifecycle::stop_office_companion,
            office::lifecycle::get_office_integration_status,
            office::lifecycle::get_office_platform_status,
            office::lifecycle::set_office_background_start,
            office::lifecycle::set_office_integration_mode,
            office::lifecycle::install_windows_ole_integration,
            office::lifecycle::uninstall_windows_ole_integration,
            office::lifecycle::repair_windows_office_integration,
            office::lifecycle::test_windows_office_runtime,
            office::lifecycle::open_windows_office_logs,
            office::lifecycle::install_office_integration,
            office::lifecycle::repair_office_integration,
            office::lifecycle::uninstall_office_integration,
            office::lifecycle::regenerate_office_certificate,
            office::lifecycle::open_word,
            office::lifecycle::open_powerpoint
        ])
        .build(tauri::generate_context!())
        .expect("error while building VisualTeX");

    app.run(move |app, event| match event {
        #[cfg(target_os = "macos")]
        tauri::RunEvent::WindowEvent {
            label,
            event: tauri::WindowEvent::CloseRequested { api, .. },
            ..
        } if label == "main" => {
            api.prevent_close();
            let _ = office::background::hide_main_window(app);
        }
        #[cfg(target_os = "macos")]
        tauri::RunEvent::Reopen { .. } => {
            let _ = office::background::reveal_main_window(app);
        }
        tauri::RunEvent::ExitRequested { .. } => {
            if let Some(state) = app.try_state::<OcrState>() {
                let _ = state.cancel_install();
            }
            #[cfg(target_os = "macos")]
            if let Err(error) = office::background::pause_launch_agent_for_quit() {
                eprintln!("Unable to pause VisualTeX Office background service: {error}");
            }
            if let Some(state) = app.try_state::<office::state::OfficeCompanionState>() {
                if let Err(error) = state.platform_backend.shutdown() {
                    eprintln!("Unable to stop the VisualTeX Office platform backend: {error}");
                }
            }
        }
        _ => {}
    });
}

#[cfg(test)]
mod protocol_tests {
    use super::*;
    use std::io::{BufReader, Cursor};

    #[test]
    fn worker_protocol_skips_non_utf8_diagnostics_before_json() {
        let bytes = b"\xd5\xfd\xca\xbd\xc8\xd5\xd6\xbe\n{\"event\":\"ready\",\"ok\":true}\n";
        let mut reader = BufReader::new(Cursor::new(bytes));

        let value = read_worker_json(&mut reader, "closed", "test response")
            .expect("reader should skip non-protocol diagnostic bytes");

        assert_eq!(value.get("event").and_then(Value::as_str), Some("ready"));
    }

    #[test]
    fn worker_protocol_decodes_ascii_escaped_unicode() {
        let bytes = b"{\"message\":\"\\u6b63\\u5728\\u52a0\\u8f7d\"}\n";
        let mut reader = BufReader::new(Cursor::new(bytes));

        let value = read_worker_json(&mut reader, "closed", "test response")
            .expect("escaped Unicode JSON should parse");

        assert_eq!(
            value.get("message").and_then(Value::as_str),
            Some("正在加载")
        );
    }

    #[test]
    fn ocr_event_bus_supports_cursor_filtering_and_bounded_history() {
        let events = OcrEventBus::default();
        let baseline = events.poll(u64::MAX, None);
        assert_eq!(baseline.cursor, 0);
        assert!(baseline.events.is_empty());

        events.publish(
            "ocr-install-progress",
            &json!({ "stage": "python", "percent": 5 }),
        );
        events.publish(
            "ocr-recognition-progress",
            &json!({ "stage": "model", "model": "PP-FormulaNet_plus-M" }),
        );

        let install_only = events.poll(0, Some("ocr-install-progress"));
        assert_eq!(install_only.cursor, 2);
        assert_eq!(install_only.events.len(), 1);
        assert_eq!(install_only.events[0].event, "ocr-install-progress");

        let incremental = events.poll(1, None);
        assert_eq!(incremental.events.len(), 1);
        assert_eq!(incremental.events[0].id, 2);

        for index in 0..(MAX_OCR_EVENTS + 20) {
            events.publish("ocr-recognition-progress", &json!({ "index": index }));
        }
        let bounded = events.poll(0, None);
        assert_eq!(bounded.events.len(), MAX_OCR_EVENTS);
        assert_eq!(bounded.cursor, (MAX_OCR_EVENTS + 22) as u64);
        assert!(bounded.events.first().is_some_and(|event| event.id > 1));
    }

    #[test]
    fn temporary_ocr_state_clone_cannot_terminate_shared_worker() {
        let state = OcrState::default();
        let temporary = state.clone();
        assert!(!is_final_ocr_state_owner(&state.worker));
        drop(temporary);
        assert!(is_final_ocr_state_owner(&state.worker));
    }

    #[test]
    fn runtime_status_cache_round_trips_and_clears() {
        let cache = Arc::new(Mutex::new(None));
        let expected = OcrRuntimeStatus {
            installed: true,
            python_path: Some("/tmp/visualtex-python".to_string()),
            python_version: Some("3.12.10".to_string()),
            paddle_version: Some(PADDLE_VERSION.to_string()),
            paddleocr_version: Some(PADDLEOCR_VERSION.to_string()),
            runtime_path: "/tmp/visualtex-ocr".to_string(),
            offline_bundle_available: true,
            installed_models: vec![ocr_offline::OFFLINE_DEFAULT_MODEL.to_string()],
            default_model: ocr_offline::OFFLINE_DEFAULT_MODEL.to_string(),
            message: "ready".to_string(),
        };

        write_cached_runtime_status(&cache, Some(expected.clone()))
            .expect("cache write should succeed");
        let cached = read_cached_runtime_status(&cache)
            .expect("cache read should succeed")
            .expect("cached status should exist");
        assert_eq!(cached.python_version, expected.python_version);
        assert!(cached.installed);

        write_cached_runtime_status(&cache, None).expect("cache clear should succeed");
        assert!(read_cached_runtime_status(&cache)
            .expect("cache read after clear should succeed")
            .is_none());
    }

    fn test_runtime_paths(root: &Path) -> RuntimePaths {
        let venv = root.join("venv");
        RuntimePaths {
            root: root.to_path_buf(),
            python: venv.join("Scripts").join("python.exe"),
            venv,
            input: root.join("input"),
            processed: root.join("processed"),
            logs: root.join("logs"),
            cache: root.join("cache"),
            temp: root.join("tmp"),
        }
    }

    #[test]
    fn runtime_probe_accepts_camel_case_and_legacy_snake_case() {
        let camel: RuntimeProbe = serde_json::from_str(
            r#"{"pythonVersion":"3.12.10","paddleVersion":"3.3.1","paddleocrVersion":"3.7.0"}"#,
        )
        .expect("camelCase probe should parse");
        let snake: RuntimeProbe = serde_json::from_str(
            r#"{"python_version":"3.12.10","paddle_version":"3.3.1","paddleocr_version":"3.7.0"}"#,
        )
        .expect("legacy snake_case probe should parse");
        assert_eq!(camel.python_version, snake.python_version);
        assert_eq!(camel.paddle_version, PADDLE_VERSION);
        assert_eq!(snake.paddleocr_version, PADDLEOCR_VERSION);
    }

    #[test]
    fn invalid_runtime_probe_preserves_raw_stdout_and_stderr() {
        let error = parse_runtime_probe_output(
            "Paddle diagnostic\n{\"paddleVersion\":\"3.3.1\"}",
            "No ccache found (warning only)",
        )
        .expect_err("missing pythonVersion should fail");
        assert!(error.contains("Raw stdout"));
        assert!(error.contains("Paddle diagnostic"));
        assert!(error.contains("Raw stderr"));
        assert!(error.contains("No ccache found"));
    }

    #[test]
    fn python_313_is_rejected_for_tokenizers_0191() {
        let python_312 = PythonProbe {
            version: "3.12.10".to_string(),
            major: 3,
            minor: 12,
            bits: 64,
            machine: "AMD64".to_string(),
            executable: "C:\\Python312\\python.exe".to_string(),
        };
        let python_313 = PythonProbe {
            version: "3.13.5".to_string(),
            major: 3,
            minor: 13,
            bits: 64,
            machine: "AMD64".to_string(),
            executable: "C:\\Python313\\python.exe".to_string(),
        };
        let python_312_32 = PythonProbe {
            version: "3.12.10".to_string(),
            major: 3,
            minor: 12,
            bits: 32,
            machine: "AMD64".to_string(),
            executable: "C:\\Python312-32\\python.exe".to_string(),
        };
        assert!(is_supported_windows_ocr_python(&python_312));
        assert!(!is_supported_windows_ocr_python(&python_313));
        assert!(!is_supported_windows_ocr_python(&python_312_32));
        assert!(should_replace_with_bundled_python(
            "The existing VisualTeX OCR environment uses a 32-bit interpreter"
        ));
    }

    #[test]
    fn tokenizers_install_forbids_source_compilation() {
        let cache = tempfile::tempdir().unwrap();
        let arguments = pip_install_arguments("tokenizers==0.19.1", true, cache.path());
        assert!(arguments.iter().any(|value| value == "--only-binary=:all:"));
        assert!(arguments.iter().any(|value| value == "--isolated"));
        assert!(arguments.iter().any(|value| value == "https://pypi.org/simple"));
        assert!(arguments.iter().any(|value| value == "--prefer-binary"));
        assert!(arguments.windows(2).any(|pair| pair == ["--progress-bar", "raw"]));
        assert!(arguments.iter().any(|value| value == "--retries"));
        assert!(arguments.iter().any(|value| value == "30"));
        assert!(!arguments.iter().any(|value| value.contains("no-binary")));
        assert_eq!(arguments.last().map(String::as_str), Some("tokenizers==0.19.1"));
    }

    #[test]
    fn large_packages_have_long_total_timeout_and_separate_idle_timeout() {
        let paddle = command_limits_for_step("paddle");
        let paddleocr = command_limits_for_step("paddleocr");
        assert_eq!(paddle.step_timeout, Duration::from_secs(30 * 60));
        assert_eq!(paddleocr.step_timeout, Duration::from_secs(30 * 60));
        assert_eq!(paddle.idle_timeout, Duration::from_secs(6 * 60));
        assert!(paddle.idle_timeout < paddle.step_timeout);
    }

    #[test]
    fn missing_tokenizers_resumes_that_step_instead_of_resetting_environment() {
        let missing: Result<InstalledPackageProbe, String> = Err(
            "ModuleNotFoundError: No module named 'tokenizers'".to_string(),
        );
        assert!(package_probe_requires_install(&missing));

        let root = tempfile::tempdir().unwrap();
        let mut snapshot = InstallSnapshot::new(&install_log_path(root.path()));
        snapshot.mark_step_complete("venv");
        snapshot.mark_step_complete("paddle");
        snapshot.mark_step_complete("paddleocr");
        snapshot.state = InstallState::InstallFailed;
        snapshot.current_step = Some("tokenizers".to_string());
        save_snapshot(root.path(), &snapshot).unwrap();
        let restored = load_snapshot(root.path()).expect("half-installed state should restore");
        assert!(restored.step_complete("paddle"));
        assert!(restored.step_complete("paddleocr"));
        assert!(!restored.step_complete("tokenizers"));
        assert_eq!(restored.current_step.as_deref(), Some("tokenizers"));
    }

    #[test]
    fn fast_runtime_probe_detects_missing_tokenizers_in_half_installed_environment() {
        let root = tempfile::tempdir().unwrap();
        let paths = test_runtime_paths(root.path());
        fs::create_dir_all(paths.python.parent().unwrap()).unwrap();
        fs::write(&paths.python, b"fake-python").unwrap();
        fs::write(paths.venv.join("pyvenv.cfg"), "version = 3.12.10\n").unwrap();
        let site_packages = paths.venv.join("Lib").join("site-packages");
        for module in ["paddle", "paddleocr", "ftfy", "wand"] {
            fs::create_dir_all(site_packages.join(module)).unwrap();
        }
        fs::create_dir_all(&site_packages).unwrap();
        fs::write(site_packages.join("imagesize.py"), b"# fake").unwrap();
        for dist_info in [
            format!("paddlepaddle-{PADDLE_VERSION}.dist-info"),
            format!("paddleocr-{PADDLEOCR_VERSION}.dist-info"),
            "imagesize-2.0.0.dist-info".to_string(),
            "ftfy-6.3.1.dist-info".to_string(),
            "Wand-0.7.2.dist-info".to_string(),
        ] {
            fs::create_dir_all(site_packages.join(dist_info)).unwrap();
        }

        let error = probe_runtime_from_files(&paths)
            .expect_err("missing tokenizers must not be reported as installed");
        assert!(error.contains("tokenizers"));

        fs::create_dir_all(site_packages.join("tokenizers")).unwrap();
        fs::create_dir_all(site_packages.join("tokenizers-0.19.1.dist-info")).unwrap();
        let probe = probe_runtime_from_files(&paths)
            .expect("complete dependency files should restore fast status");
        assert_eq!(probe.python_version, "3.12.10");
        assert_eq!(probe.paddle_version, PADDLE_VERSION);
        assert_eq!(probe.paddleocr_version, PADDLEOCR_VERSION);
    }

    #[test]
    fn package_probe_rejects_wrong_interpreter_and_external_module_path() {
        let root = tempfile::tempdir().unwrap();
        let paths = test_runtime_paths(root.path());
        let wrong_interpreter = InstalledPackageProbe {
            distribution: "tokenizers".to_string(),
            version: "0.19.1".to_string(),
            module_path: paths
                .venv
                .join("Lib/site-packages/tokenizers/__init__.py")
                .display()
                .to_string(),
            executable: "C:\\Python312\\python.exe".to_string(),
            python_version: "3.12.10".to_string(),
        };
        let error = validate_installed_package_probe(
            &paths,
            "tokenizers",
            Some("0.19.1"),
            &wrong_interpreter,
        )
        .expect_err("wrong interpreter should fail");
        assert!(error.contains("wrong interpreter"));

        let external_module = InstalledPackageProbe {
            executable: paths.python.display().to_string(),
            module_path: "C:\\Python312\\Lib\\site-packages\\tokenizers\\__init__.py"
                .to_string(),
            ..wrong_interpreter
        };
        let error = validate_installed_package_probe(
            &paths,
            "tokenizers",
            Some("0.19.1"),
            &external_module,
        )
        .expect_err("external module path should fail");
        assert!(error.contains("not installed inside"));
    }

    #[test]
    fn installation_blocks_full_runtime_probe_and_old_error_mixing() {
        assert!(should_run_full_runtime_probe(true, false));
        assert!(!should_run_full_runtime_probe(true, true));
        assert!(!should_run_full_runtime_probe(false, true));
    }

    #[test]
    fn interrupted_persisted_install_becomes_retryable_failure() {
        let root = tempfile::tempdir().unwrap();
        let mut snapshot = InstallSnapshot::new(&install_log_path(root.path()));
        snapshot.state = InstallState::Installing;
        snapshot.current_step = Some("tokenizers".to_string());
        snapshot.percent = 82;
        snapshot.mark_step_complete("paddle");
        snapshot.mark_step_complete("paddleocr");

        assert!(reconcile_interrupted_install_snapshot(&mut snapshot, false));
        assert_eq!(snapshot.state, InstallState::InstallFailed);
        assert_eq!(snapshot.current_step.as_deref(), Some("tokenizers"));
        assert_eq!(snapshot.percent, 82);
        assert!(snapshot.step_complete("paddle"));
        assert!(snapshot.error.as_deref().is_some_and(|value| value.contains("tokenizers")));

        let mut active = snapshot.clone();
        active.state = InstallState::Verifying;
        active.current_step = Some("verify".to_string());
        assert!(!reconcile_interrupted_install_snapshot(&mut active, true));
        assert_eq!(active.state, InstallState::Verifying);
    }

    #[test]
    fn damaged_venv_detection_excludes_incompatible_cancelled_and_timed_out_cases() {
        assert!(damaged_venv_python_error(
            "OCR virtual environment Python check failed with exit code: 1"
        ));
        assert!(damaged_venv_python_error(
            "OCR virtual environment returned invalid Python information"
        ));
        assert!(!damaged_venv_python_error(
            "The existing VisualTeX OCR environment uses Python 3.13.5"
        ));
        assert!(!damaged_venv_python_error(
            "OCR virtual environment Python check was cancelled"
        ));
        assert!(!damaged_venv_python_error(
            "OCR virtual environment Python check timed out"
        ));
    }

    #[test]
    fn runtime_status_manifest_round_trips_and_reads_legacy_fields() {
        let root = tempfile::tempdir().unwrap();
        let paths = test_runtime_paths(root.path());
        let expected = RuntimeProbe {
            python_version: "3.12.10".to_string(),
            paddle_version: PADDLE_VERSION.to_string(),
            paddleocr_version: PADDLEOCR_VERSION.to_string(),
        };
        write_runtime_status_manifest(&paths, &expected).unwrap();
        let current = read_runtime_status_manifest(&paths).expect("current manifest should load");
        assert_eq!(current.python_version, "3.12.10");

        let legacy = json!({
            "schema_version": OCR_RUNTIME_STATUS_SCHEMA,
            "python_path": paths.python.display().to_string(),
            "python_version": "3.12.9",
            "paddle_version": PADDLE_VERSION,
            "paddleocr_version": PADDLEOCR_VERSION,
        });
        fs::write(
            runtime_status_manifest_path(&paths),
            serde_json::to_vec_pretty(&legacy).unwrap(),
        )
        .unwrap();
        let legacy_probe =
            read_runtime_status_manifest(&paths).expect("legacy manifest should load");
        assert_eq!(legacy_probe.python_version, "3.12.9");
    }
}

#[cfg(all(test, unix))]
mod tests {
    use super::*;

    #[test]
    fn active_worker_can_be_terminated_without_taking_the_worker_lock() {
        let mut child = Command::new("sleep")
            .arg("30")
            .spawn()
            .expect("failed to start test process");
        let pid = child.id();
        let worker_pid = AtomicU32::new(pid);

        assert!(terminate_worker_process(&worker_pid).expect("termination failed"));
        let status = child.wait().expect("failed to wait for test process");

        assert!(!status.success());
        assert_eq!(worker_pid.load(Ordering::SeqCst), 0);
    }
}
