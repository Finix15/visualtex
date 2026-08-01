use crate::office::server::metadata_from_session;
use crate::office::sessions::{
    valid_uuid, CreateOfficeSessionInput, FormulaLine, OfficeFormulaSession, OfficeHost,
    OfficeSessionMode, OfficeSessionStatus, SessionError, VisualTeXFormulaMetadata,
};
use crate::office::state::OfficeCompanionState;
use base64::{
    engine::general_purpose::{STANDARD as BASE64_STANDARD, URL_SAFE_NO_PAD},
    Engine as _,
};
use flate2::{read::DeflateDecoder, write::DeflateEncoder, Compression};
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::fs::{self, OpenOptions};
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::{mpsc, Mutex, OnceLock};
use std::time::{Instant, SystemTime, UNIX_EPOCH};
use tauri::{AppHandle, Emitter, Manager, WebviewUrl, WebviewWindow, WebviewWindowBuilder};
use uuid::Uuid;

const OFFLINE_PROTOCOL_VERSION: u32 = 1;
const REQUEST_FILE: &str = "request.json";
const DISPATCH_FILE: &str = "dispatch.txt";
const RESULT_PNG_FILE: &str = "formula.png";
const RESULT_SVG_FILE: &str = "formula.svg";
const RESULT_WORD_SVG_DOCX_FILE: &str = "formula-svg.docx";
const DOCUMENT_IMPORT_MANIFEST_FILE: &str = "document-import.txt";
const DOCUMENT_IMPORT_PROGRESS_FILE: &str = "document-import-progress.txt";
const EDITOR_READY_FILE: &str = "editor-ready.json";
const EDITOR_PERFORMANCE_FILE: &str = "editor-performance.jsonl";
const OFFICE_EDITOR_ACTIVATE_EVENT: &str = "visualtex-office-editor-activate";
const OFFICE_EDITOR_CLEAR_EVENT: &str = "visualtex-office-editor-clear";
const WORD_POINTER_FILE: &str = "word-active-session.txt";
const POWERPOINT_POINTER_FILE: &str = "powerpoint-active-session.txt";
const WORD_RUNTIME_SUFFIX: &str =
    "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime";
const POWERPOINT_RUNTIME_SUFFIX: &str =
    "Library/Application Scripts/com.microsoft.Powerpoint/VisualTeXRuntime";
const METADATA_PREFIX: &str = "visualtex:v1:deflate:";
const PENDING_PREFIX: &str = "visualtex:pending:v1:";
const MAX_REQUEST_BYTES: u64 = 256 * 1024;
const MAX_METADATA_BYTES: usize = 2 * 1024 * 1024;
const MAX_OMML_BYTES: usize = 4 * 1024 * 1024;
const MAX_DOCUMENT_IMPORT_MANIFEST_BYTES: usize = 16 * 1024 * 1024;
const MAX_IDENTITY_CHARS: usize = 2048;
const MAX_SHAPE_NAME_CHARS: usize = 128;
const MAX_WORD_WIDTH_PT: f64 = 500.0;
const WORD_REFERENCE_FONT_SIZE_PT: f64 = 14.0;
// MathJax's TeX glyphs render about 9.5% narrower than Word's Cambria Math at
// the same nominal point size. Scale only Word image formulas so 14 pt means a
// comparable painted formula size; PowerPoint keeps its existing geometry.
const WORD_IMAGE_VISUAL_SCALE: f64 = 1.1;
const MIN_WORD_FONT_SIZE_PT: f64 = 1.0;
const MAX_WORD_FONT_SIZE_PT: f64 = 512.0;
const POWERPOINT_REFERENCE_FONT_SIZE_PT: f64 = 14.0;
const DEFAULT_POWERPOINT_FONT_SIZE_PT: f64 = 18.0;
const MIN_POWERPOINT_FONT_SIZE_PT: f64 = 1.0;
const MAX_POWERPOINT_FONT_SIZE_PT: f64 = 512.0;
static WORD_DISPATCH_LOCK: Mutex<()> = Mutex::new(());
static POWERPOINT_DISPATCH_LOCK: Mutex<()> = Mutex::new(());

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct MacOfflinePowerPointRequest {
    presentation_identity: String,
    slide_index: u32,
    slide_id: u32,
    shape_name: String,
    left: f64,
    top: f64,
    width: f64,
    height: f64,
    rotation: f64,
    z_order: u32,
    #[serde(default)]
    font_size_pt: Option<f64>,
    #[serde(default)]
    reference_width_pt: Option<f64>,
    #[serde(default)]
    reference_height_pt: Option<f64>,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct MacOfflineDocumentImportRequest {
    bookmark_name: String,
    default_font_size_pt: f64,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct MacOfflineSessionRequest {
    protocol_version: u32,
    session_id: String,
    host: String,
    mode: String,
    #[serde(default)]
    operation: Option<String>,
    formula_id: Option<String>,
    display_mode: String,
    numbered: bool,
    #[serde(default)]
    native_equation: bool,
    source_document_id: Option<String>,
    source_object_id: Option<String>,
    encoded_metadata: Option<String>,
    pending_marker: Option<String>,
    #[serde(default)]
    font_size_pt: Option<f64>,
    #[serde(default)]
    reference_width_pt: Option<f64>,
    #[serde(default)]
    reference_height_pt: Option<f64>,
    power_point: Option<MacOfflinePowerPointRequest>,
    #[serde(default)]
    document_import: Option<MacOfflineDocumentImportRequest>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MacOfflineDocumentImportPublicRequest {
    protocol_version: u32,
    session_id: String,
    host: String,
    source_document_id: String,
    bookmark_name: String,
    default_font_size_pt: f64,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MacOfflineDocumentImportProgress {
    current: usize,
    total: usize,
    stage: String,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct MacOfflineDocumentImportCommitInput {
    output_kind: String,
    items: Vec<MacOfflineDocumentImportCommitItem>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(tag = "kind", rename_all = "lowercase", rename_all_fields = "camelCase")]
pub enum MacOfflineDocumentImportCommitItem {
    Text {
        text: String,
        #[serde(default)]
        paragraph_id: Option<String>,
        #[serde(default)]
        paragraph_style: Option<String>,
        #[serde(default)]
        paragraph_alignment: Option<String>,
        #[serde(default)]
        list_kind: Option<String>,
        #[serde(default)]
        list_level: Option<u32>,
        #[serde(default)]
        paragraph_start: bool,
        #[serde(default)]
        paragraph_end: bool,
    },
    Formula {
        formula_id: String,
        latex: String,
        display_mode: String,
        #[serde(default)]
        numbered: bool,
        font_size_pt: f64,
        metadata: VisualTeXFormulaMetadata,
        omml_base64: String,
        omml_docx_base64: String,
        #[serde(default)]
        svg_base64: Option<String>,
        #[serde(default)]
        png_base64: Option<String>,
        #[serde(default)]
        width: Option<f64>,
        #[serde(default)]
        height: Option<f64>,
        #[serde(default)]
        baseline: Option<f64>,
        #[serde(default)]
        paragraph_id: Option<String>,
        #[serde(default)]
        paragraph_style: Option<String>,
        #[serde(default)]
        paragraph_alignment: Option<String>,
        #[serde(default)]
        list_kind: Option<String>,
        #[serde(default)]
        list_level: Option<u32>,
        #[serde(default)]
        paragraph_start: bool,
        #[serde(default)]
        paragraph_end: bool,
    },
}

#[derive(Debug, Clone)]
struct DocumentParagraphTransfer {
    id: String,
    style: String,
    alignment: String,
    list_kind: String,
    list_level: u32,
    start: bool,
    end: bool,
}

#[derive(Debug, Clone, Copy)]
struct WordGeometry {
    width: f64,
    height: f64,
    baseline: i32,
    font_size_pt: f64,
    reference_width_pt: f64,
    reference_height_pt: f64,
    reference_baseline_pt: f64,
}

#[derive(Debug, Clone, Copy)]
struct PowerPointGeometry {
    left: f64,
    top: f64,
    width: f64,
    height: f64,
    font_size_pt: f64,
    reference_width_pt: f64,
    reference_height_pt: f64,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MacOfflinePluginHealth {
    loaded: bool,
    plugin_version: Option<String>,
    source_revision: Option<String>,
    host: String,
    timestamp: Option<String>,
    status_path: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MacOfflineOfficeEditorActivation {
    session_id: String,
    host: OfficeHost,
    generation: u64,
    received_epoch_ms: u64,
}

#[derive(Debug, Clone)]
struct ActiveOfficeEditorSession {
    activation: MacOfflineOfficeEditorActivation,
    received_at: Instant,
    ready: bool,
}

#[derive(Debug, Default)]
struct OfficeEditorRuntime {
    next_generation: u64,
    word: Option<ActiveOfficeEditorSession>,
    powerpoint: Option<ActiveOfficeEditorSession>,
}

impl OfficeEditorRuntime {
    fn active(&self, host: OfficeHost) -> Option<&ActiveOfficeEditorSession> {
        match host {
            OfficeHost::Word => self.word.as_ref(),
            OfficeHost::Powerpoint => self.powerpoint.as_ref(),
        }
    }

    fn active_mut(&mut self, host: OfficeHost) -> &mut Option<ActiveOfficeEditorSession> {
        match host {
            OfficeHost::Word => &mut self.word,
            OfficeHost::Powerpoint => &mut self.powerpoint,
        }
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct OfficeEditorPerformanceRecord {
    schema: &'static str,
    session_id: String,
    host: OfficeHost,
    stage: String,
    epoch_ms: u64,
    elapsed_ms: f64,
    generation: Option<u64>,
    details: Value,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct MacOfflineOfficeEditorReadyInput {
    session_id: String,
    generation: u64,
    frontend_epoch_ms: u64,
    hydrate_ms: f64,
    editor_mounted_ms: f64,
    content_ready_ms: f64,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct MacOfflineOfficeEditorReadyMarker {
    schema: &'static str,
    session_id: String,
    host: OfficeHost,
    generation: u64,
    epoch_ms: u64,
    url_received_epoch_ms: u64,
    frontend_epoch_ms: u64,
    hydrate_ms: f64,
    editor_mounted_ms: f64,
    content_ready_ms: f64,
    show_focus_ms: f64,
}

fn office_editor_runtime() -> &'static Mutex<OfficeEditorRuntime> {
    static RUNTIME: OnceLock<Mutex<OfficeEditorRuntime>> = OnceLock::new();
    RUNTIME.get_or_init(|| Mutex::new(OfficeEditorRuntime::default()))
}

fn epoch_ms() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_millis() as u64)
        .unwrap_or_default()
}

fn performance_logger() -> &'static mpsc::Sender<OfficeEditorPerformanceRecord> {
    static LOGGER: OnceLock<mpsc::Sender<OfficeEditorPerformanceRecord>> = OnceLock::new();
    LOGGER.get_or_init(|| {
        let (sender, receiver) = mpsc::channel::<OfficeEditorPerformanceRecord>();
        std::thread::Builder::new()
            .name("visualtex-office-performance".to_string())
            .spawn(move || {
                while let Ok(record) = receiver.recv() {
                    let Ok(directory) = session_directory(record.host, &record.session_id) else {
                        continue;
                    };
                    if fs::create_dir_all(&directory).is_err() {
                        continue;
                    }
                    let path = directory.join(EDITOR_PERFORMANCE_FILE);
                    let Ok(mut line) = serde_json::to_vec(&record) else {
                        continue;
                    };
                    line.push(b'\n');
                    if let Ok(mut file) = OpenOptions::new().create(true).append(true).open(&path) {
                        if file.write_all(&line).is_ok() {
                            let _ = set_mode(&path, 0o600);
                        }
                    }
                }
            })
            .expect("VisualTeX Office performance logger thread must start");
        sender
    })
}

fn queue_editor_performance(
    host: OfficeHost,
    session_id: &str,
    stage: impl Into<String>,
    elapsed_ms: f64,
    generation: Option<u64>,
    details: Value,
) {
    let _ = performance_logger().send(OfficeEditorPerformanceRecord {
        schema: "visualtex-office-editor-performance-v1",
        session_id: session_id.to_string(),
        host,
        stage: stage.into(),
        epoch_ms: epoch_ms(),
        elapsed_ms,
        generation,
        details,
    });
}

fn user_home() -> Result<PathBuf, String> {
    std::env::var_os("HOME")
        .map(PathBuf::from)
        .filter(|path| path.is_absolute())
        .ok_or_else(|| "Unable to resolve the current user's home directory".to_string())
}

pub(crate) fn runtime_root(host: OfficeHost) -> Result<PathBuf, String> {
    let suffix = match host {
        OfficeHost::Word => WORD_RUNTIME_SUFFIX,
        OfficeHost::Powerpoint => POWERPOINT_RUNTIME_SUFFIX,
    };
    Ok(user_home()?.join(suffix))
}

fn host_from_request_name(value: &str) -> Result<OfficeHost, String> {
    match value {
        "word" => Ok(OfficeHost::Word),
        "powerpoint" => Ok(OfficeHost::Powerpoint),
        _ => Err("Offline Office request host must be word or powerpoint".to_string()),
    }
}

fn sessions_root(host: OfficeHost) -> Result<PathBuf, String> {
    Ok(runtime_root(host)?.join("OfficeSessions"))
}

fn ensure_runtime_root(host: OfficeHost) -> Result<PathBuf, String> {
    let root = runtime_root(host)?;
    let sessions = root.join("OfficeSessions");
    fs::create_dir_all(&sessions)
        .map_err(|error| format!("Unable to create {}: {error}", sessions.display()))?;
    set_mode(&root, 0o700)?;
    set_mode(&sessions, 0o700)?;
    Ok(root)
}

fn session_directory(host: OfficeHost, session_id: &str) -> Result<PathBuf, String> {
    validate_uuid(session_id, "Session id")?;
    Ok(sessions_root(host)?.join(session_id))
}

fn request_path(host: OfficeHost, session_id: &str) -> Result<PathBuf, String> {
    Ok(session_directory(host, session_id)?.join(REQUEST_FILE))
}

fn dispatch_path(host: OfficeHost, session_id: &str) -> Result<PathBuf, String> {
    Ok(session_directory(host, session_id)?.join(DISPATCH_FILE))
}

fn document_import_manifest_path(session_id: &str) -> Result<PathBuf, String> {
    Ok(session_directory(OfficeHost::Word, session_id)?.join(DOCUMENT_IMPORT_MANIFEST_FILE))
}

fn result_png_path(host: OfficeHost, session_id: &str) -> Result<PathBuf, String> {
    Ok(session_directory(host, session_id)?.join(RESULT_PNG_FILE))
}

fn result_svg_path(host: OfficeHost, session_id: &str) -> Result<PathBuf, String> {
    Ok(session_directory(host, session_id)?.join(RESULT_SVG_FILE))
}

fn result_word_svg_docx_path(session_id: &str) -> Result<PathBuf, String> {
    Ok(session_directory(OfficeHost::Word, session_id)?.join(RESULT_WORD_SVG_DOCX_FILE))
}

fn native_word_document_path(formula_id: &str) -> Result<PathBuf, String> {
    validate_uuid(formula_id, "Formula id")?;
    let directory = runtime_root(OfficeHost::Word)?.join("NativeDocuments");
    fs::create_dir_all(&directory)
        .map_err(|error| format!("Unable to create {}: {error}", directory.display()))?;
    set_mode(&directory, 0o700)?;
    Ok(directory.join(format!("{formula_id}.docx")))
}

fn cleanup_session_files_at(directory: &Path) -> Result<(), String> {
    for name in [
        REQUEST_FILE,
        DISPATCH_FILE,
        RESULT_PNG_FILE,
        RESULT_SVG_FILE,
        RESULT_WORD_SVG_DOCX_FILE,
        DOCUMENT_IMPORT_MANIFEST_FILE,
        "formula.docx",
    ] {
        let path = directory.join(name);
        match fs::remove_file(&path) {
            Ok(()) => {}
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => {}
            Err(error) => return Err(format!("Unable to remove {}: {error}", path.display())),
        }
    }
    if let Ok(entries) = fs::read_dir(directory) {
        for entry in entries.flatten() {
            let name = entry.file_name();
            let name = name.to_string_lossy();
            if name.starts_with("document-formula-") {
                let path = entry.path();
                if path.is_file() {
                    let _ = fs::remove_file(path);
                }
            }
        }
    }
    match fs::remove_dir(directory) {
        Ok(()) => Ok(()),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(()),
        Err(error) if error.kind() == std::io::ErrorKind::DirectoryNotEmpty => Ok(()),
        Err(error) => Err(format!(
            "Unable to remove offline Office Session directory {}: {error}",
            directory.display()
        )),
    }
}

fn cleanup_session_files(host: OfficeHost, session_id: &str) -> Result<(), String> {
    cleanup_session_files_at(&session_directory(host, session_id)?)
}

fn pointer_path(host: OfficeHost) -> Result<PathBuf, String> {
    Ok(sessions_root(host)?.join(match host {
        OfficeHost::Word => WORD_POINTER_FILE,
        OfficeHost::Powerpoint => POWERPOINT_POINTER_FILE,
    }))
}

fn validate_uuid(value: &str, label: &str) -> Result<(), String> {
    if valid_uuid(value) {
        Ok(())
    } else {
        Err(format!("{label} must be a canonical UUID"))
    }
}

fn validate_bounded_text(value: &str, maximum: usize, label: &str) -> Result<(), String> {
    if value.chars().count() > maximum || value.chars().any(char::is_control) {
        return Err(format!("{label} contains unsupported characters or is too long"));
    }
    Ok(())
}

fn validate_finite_geometry(value: f64, label: &str) -> Result<(), String> {
    if !value.is_finite() || value.abs() > 10_000_000.0 {
        Err(format!("PowerPoint {label} is invalid"))
    } else {
        Ok(())
    }
}

fn validate_request(request: &MacOfflineSessionRequest, session_id: &str) -> Result<(), String> {
    if request.protocol_version != OFFLINE_PROTOCOL_VERSION {
        return Err("Unsupported VisualTeX macOS offline protocol version".to_string());
    }
    validate_uuid(&request.session_id, "Request Session id")?;
    if request.session_id != session_id {
        return Err("Request Session id does not match the custom URL".to_string());
    }
    if !matches!(request.host.as_str(), "word" | "powerpoint") {
        return Err("Offline Office request host must be word or powerpoint".to_string());
    }
    if !matches!(request.mode.as_str(), "create" | "edit") {
        return Err("Offline Office request mode must be create or edit".to_string());
    }
    let operation = request.operation.as_deref().unwrap_or("formula");
    if operation == "documentImport" {
        if request.host != "word" || request.mode != "create" {
            return Err("Document import is supported only as a new Word operation".to_string());
        }
        let document_import = request
            .document_import
            .as_ref()
            .ok_or_else(|| "Document import request is missing its insertion anchor".to_string())?;
        let source_document_id = request
            .source_document_id
            .as_deref()
            .ok_or_else(|| "Document import request is missing the Word document identity".to_string())?;
        validate_bounded_text(source_document_id, MAX_IDENTITY_CHARS, "sourceDocumentId")?;
        validate_bounded_text(&document_import.bookmark_name, 40, "documentImport bookmarkName")?;
        if !document_import.bookmark_name.starts_with("VT_D_") {
            return Err("Document import bookmark name is invalid".to_string());
        }
        if !document_import.default_font_size_pt.is_finite()
            || !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT)
                .contains(&document_import.default_font_size_pt)
        {
            return Err("Document import default font size is outside the supported range".to_string());
        }
        if request.formula_id.is_some()
            || request.encoded_metadata.is_some()
            || request.pending_marker.is_some()
            || request.power_point.is_some()
            || request.document_import.is_none()
            || request.numbered
            || request.native_equation
        {
            return Err("Document import request contains formula-only fields".to_string());
        }
        return Ok(());
    }
    if operation != "formula" || request.document_import.is_some() {
        return Err("Unsupported offline Office operation".to_string());
    }
    if !matches!(request.display_mode.as_str(), "inline" | "block") {
        return Err("Offline Office displayMode must be inline or block".to_string());
    }
    if request.numbered && (request.host != "word" || request.display_mode != "block") {
        return Err("Only Word display formulas can be numbered".to_string());
    }
    if request.native_equation && request.host != "word" {
        return Err("Native equations are supported only by Word requests".to_string());
    }
    if let Some(formula_id) = request.formula_id.as_deref() {
        validate_uuid(formula_id, "Formula id")?;
    }
    for (value, label) in [
        (request.source_document_id.as_deref(), "sourceDocumentId"),
        (request.source_object_id.as_deref(), "sourceObjectId"),
        (request.pending_marker.as_deref(), "pendingMarker"),
    ] {
        if let Some(value) = value {
            validate_bounded_text(value, MAX_IDENTITY_CHARS, label)?;
        }
    }
    if let Some(marker) = request.pending_marker.as_deref() {
        if !marker.starts_with(PENDING_PREFIX) {
            return Err("Offline Office pending marker is invalid".to_string());
        }
    }
    if let Some(encoded) = request.encoded_metadata.as_deref() {
        if encoded.len() > MAX_METADATA_BYTES || !encoded.starts_with(METADATA_PREFIX) {
            return Err("Offline Office metadata envelope is invalid".to_string());
        }
    }
    for (value, label) in [
        (request.font_size_pt, "fontSizePt"),
        (request.reference_width_pt, "referenceWidthPt"),
        (request.reference_height_pt, "referenceHeightPt"),
    ] {
        if let Some(value) = value {
            if !value.is_finite() || value <= 0.0 {
                return Err(format!("Offline Office {label} must be a positive finite number"));
            }
        }
    }
    if let Some(font_size) = request.font_size_pt {
        if !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT).contains(&font_size) {
            return Err("Offline Office Word fontSizePt is outside the supported range".to_string());
        }
    }
    if request.host != "word"
        && (request.font_size_pt.is_some()
            || request.reference_width_pt.is_some()
            || request.reference_height_pt.is_some())
    {
        return Err("PowerPoint requests must not contain Word font-size metadata".to_string());
    }

    match (request.host.as_str(), request.power_point.as_ref()) {
        ("word", None) => {}
        ("word", Some(_)) => return Err("Word request must not contain PowerPoint geometry".to_string()),
        ("powerpoint", None) => return Err("PowerPoint request requires geometry".to_string()),
        ("powerpoint", Some(powerpoint)) => {
            validate_bounded_text(
                &powerpoint.presentation_identity,
                MAX_IDENTITY_CHARS,
                "PowerPoint presentation identity",
            )?;
            validate_bounded_text(
                &powerpoint.shape_name,
                MAX_SHAPE_NAME_CHARS,
                "PowerPoint shape name",
            )?;
            if powerpoint.slide_index == 0 || powerpoint.slide_id == 0 || powerpoint.z_order == 0 {
                return Err("PowerPoint slide and z-order references must be positive".to_string());
            }
            for (value, label) in [
                (powerpoint.left, "left"),
                (powerpoint.top, "top"),
                (powerpoint.width, "width"),
                (powerpoint.height, "height"),
                (powerpoint.rotation, "rotation"),
            ] {
                validate_finite_geometry(value, label)?;
            }
            if powerpoint.width <= 0.0 || powerpoint.height <= 0.0 {
                return Err("PowerPoint formula geometry must have positive dimensions".to_string());
            }
            for (value, label) in [
                (powerpoint.font_size_pt, "fontSizePt"),
                (powerpoint.reference_width_pt, "referenceWidthPt"),
                (powerpoint.reference_height_pt, "referenceHeightPt"),
            ] {
                if let Some(value) = value {
                    if !value.is_finite() || value <= 0.0 {
                        return Err(format!(
                            "PowerPoint formula {label} must be a positive finite number"
                        ));
                    }
                }
            }
            if let Some(font_size) = powerpoint.font_size_pt {
                if !(MIN_POWERPOINT_FONT_SIZE_PT..=MAX_POWERPOINT_FONT_SIZE_PT)
                    .contains(&font_size)
                {
                    return Err(
                        "PowerPoint formula fontSizePt is outside the supported range"
                            .to_string(),
                    );
                }
            }
        }
        _ => unreachable!(),
    }
    Ok(())
}

fn read_request(session_id: &str) -> Result<MacOfflineSessionRequest, String> {
    validate_uuid(session_id, "Session id")?;
    let mut candidates = Vec::new();
    for host in [OfficeHost::Word, OfficeHost::Powerpoint] {
        let path = request_path(host, session_id)?;
        match fs::symlink_metadata(&path) {
            Ok(metadata) => candidates.push((host, path, metadata)),
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => {}
            Err(error) => {
                return Err(format!(
                    "Unable to inspect offline Office request metadata at {}: {error}",
                    path.display()
                ))
            }
        }
    }
    let (expected_host, path, metadata) = match candidates.len() {
        1 => candidates.remove(0),
        0 => return Err("Offline Office request was not found in either host runtime directory".to_string()),
        _ => return Err("The same Offline Office Session exists in both host runtime directories".to_string()),
    };
    if metadata.file_type().is_symlink()
        || !metadata.is_file()
        || metadata.len() == 0
        || metadata.len() > MAX_REQUEST_BYTES
    {
        return Err("Offline Office request has an invalid size".to_string());
    }
    let bytes = fs::read(&path)
        .map_err(|error| format!("Unable to read offline Office request: {error}"))?;
    let request: MacOfflineSessionRequest = serde_json::from_slice(&bytes)
        .map_err(|error| format!("Offline Office request contains invalid JSON: {error}"))?;
    validate_request(&request, session_id)?;
    if host_from_request_name(&request.host)? != expected_host {
        return Err("Offline Office request host does not match its Application Scripts runtime directory".to_string());
    }
    Ok(request)
}

fn decode_metadata(encoded: &str) -> Result<VisualTeXFormulaMetadata, String> {
    let payload = encoded
        .strip_prefix(METADATA_PREFIX)
        .ok_or_else(|| "VisualTeX formula metadata prefix is invalid".to_string())?;
    if payload.is_empty() || payload.len() > MAX_METADATA_BYTES {
        return Err("VisualTeX formula metadata payload is invalid".to_string());
    }
    let compressed = URL_SAFE_NO_PAD
        .decode(payload)
        .map_err(|error| format!("Unable to decode VisualTeX formula metadata: {error}"))?;
    let decoder = DeflateDecoder::new(compressed.as_slice());
    let mut json = Vec::new();
    decoder
        .take((MAX_METADATA_BYTES + 1) as u64)
        .read_to_end(&mut json)
        .map_err(|error| format!("Unable to inflate VisualTeX formula metadata: {error}"))?;
    if json.len() > MAX_METADATA_BYTES {
        return Err("VisualTeX formula metadata expands beyond the allowed size".to_string());
    }
    let metadata: VisualTeXFormulaMetadata = serde_json::from_slice(&json)
        .map_err(|error| format!("VisualTeX formula metadata JSON is invalid: {error}"))?;
    validate_metadata(&metadata)?;
    Ok(metadata)
}

fn validate_metadata(metadata: &VisualTeXFormulaMetadata) -> Result<(), String> {
    if metadata.schema != "visualtex-formula" || metadata.schema_version != 1 {
        return Err("Unsupported VisualTeX formula metadata schema".to_string());
    }
    validate_uuid(&metadata.formula_id, "Metadata formulaId")?;
    if metadata.lines.is_empty() || metadata.lines.len() > 512 {
        return Err("VisualTeX formula metadata must contain 1 to 512 lines".to_string());
    }
    for line in &metadata.lines {
        validate_uuid(&line.id, "Metadata line id")?;
        if line.latex.len() > 1_000_000 {
            return Err("A VisualTeX formula line exceeds the 1 MB limit".to_string());
        }
    }
    if !matches!(metadata.display_mode.as_str(), "inline" | "block") {
        return Err("VisualTeX metadata displayMode is invalid".to_string());
    }
    for (value, label) in [
        (metadata.render_width_px, "renderWidthPx"),
        (metadata.render_height_px, "renderHeightPx"),
        (metadata.reference_width_pt, "referenceWidthPt"),
        (metadata.reference_height_pt, "referenceHeightPt"),
    ] {
        if let Some(value) = value {
            if !value.is_finite() || value <= 0.0 {
                return Err(format!("VisualTeX metadata {label} must be positive and finite"));
            }
        }
    }
    if let Some(value) = metadata.font_size_pt {
        if !value.is_finite()
            || !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT).contains(&value)
        {
            return Err("VisualTeX metadata fontSizePt is outside the Office range".to_string());
        }
    }
    if let Some(value) = metadata.reference_baseline_pt {
        if !value.is_finite() || !(-256.0..=0.0).contains(&value) {
            return Err("VisualTeX metadata referenceBaselinePt is invalid".to_string());
        }
    }
    Ok(())
}

fn replace_mathlive_latex_command(
    source: &str,
    command: &str,
    replacement: &str,
) -> String {
    let pattern = format!("\\{command}");
    let mut output = String::with_capacity(source.len());
    let mut cursor = 0;
    while let Some(relative) = source[cursor..].find(&pattern) {
        let start = cursor + relative;
        let end = start + pattern.len();
        output.push_str(&source[cursor..start]);
        let followed_by_command_letter = source[end..]
            .chars()
            .next()
            .is_some_and(|character| character.is_ascii_alphabetic());
        if followed_by_command_letter {
            output.push_str(&pattern);
        } else {
            output.push_str(replacement);
        }
        cursor = end;
    }
    output.push_str(&source[cursor..]);
    output
}

fn normalize_mathlive_upright_commands(source: &str) -> String {
    let mut normalized = source.to_string();
    for (command, replacement) in [
        ("capitalDifferentialD", "\\mathrm{D}"),
        ("differentialD", "\\mathrm{d}"),
        ("exponentialE", "\\mathrm{e}"),
        ("imaginaryI", "\\mathrm{i}"),
        ("imaginaryJ", "\\mathrm{j}"),
    ] {
        normalized = replace_mathlive_latex_command(&normalized, command, replacement);
    }
    for prefix in ["\\mathrm{d", "\\textrm{d"] {
        let mut output = String::with_capacity(normalized.len());
        let mut cursor = 0;
        while let Some(relative) = normalized[cursor..].find(prefix) {
            let start = cursor + relative;
            let variable_start = start + prefix.len();
            output.push_str(&normalized[cursor..start]);
            let Some(variable) = normalized[variable_start..].chars().next() else {
                output.push_str(prefix);
                cursor = variable_start;
                continue;
            };
            let variable_end = variable_start + variable.len_utf8();
            if variable.is_ascii_alphabetic()
                && normalized[variable_end..].starts_with('}')
            {
                output.push_str("\\mathrm{d}");
                output.push(variable);
                cursor = variable_end + 1;
            } else {
                output.push_str(prefix);
                cursor = variable_start;
            }
        }
        output.push_str(&normalized[cursor..]);
        normalized = output;
    }
    normalized
}

fn latex_character_is_escaped(source: &str, index: usize) -> bool {
    let bytes = source.as_bytes();
    let mut cursor = index;
    let mut slash_count = 0;
    while cursor > 0 && bytes[cursor - 1] == b'\\' {
        slash_count += 1;
        cursor -= 1;
    }
    slash_count % 2 == 1
}

fn read_latex_environment_token(
    source: &str,
    index: usize,
) -> Option<(bool, String, usize)> {
    let rest = source.get(index..)?;
    let (is_begin, name_start) = if rest.starts_with("\\begin{") {
        (true, index + "\\begin{".len())
    } else if rest.starts_with("\\end{") {
        (false, index + "\\end{".len())
    } else {
        return None;
    };
    let name_end = source[name_start..].find('}')? + name_start;
    let name = &source[name_start..name_end];
    if name.is_empty()
        || !name
            .chars()
            .enumerate()
            .all(|(position, character)| {
                character.is_ascii_alphabetic()
                    || (character == '*' && position == name.chars().count() - 1)
            })
    {
        return None;
    }
    Some((is_begin, name.to_string(), name_end + 1))
}

fn update_latex_environment_stack(
    environments: &mut Vec<String>,
    is_begin: bool,
    name: String,
) {
    if is_begin {
        environments.push(name);
    } else if let Some(index) = environments.iter().rposition(|value| value == &name) {
        environments.remove(index);
    }
}

fn has_top_level_alignment_marker(source: &str) -> bool {
    let mut brace_depth = 0_u32;
    let mut environments = Vec::new();
    let mut index = 0;
    while index < source.len() {
        if let Some((is_begin, name, end)) = read_latex_environment_token(source, index) {
            update_latex_environment_stack(&mut environments, is_begin, name);
            index = end;
            continue;
        }
        let character = source[index..].chars().next().expect("valid UTF-8");
        if character == '{' && !latex_character_is_escaped(source, index) {
            brace_depth += 1;
        } else if character == '}' && !latex_character_is_escaped(source, index) {
            brace_depth = brace_depth.saturating_sub(1);
        } else if character == '&'
            && !latex_character_is_escaped(source, index)
            && brace_depth == 0
            && environments.is_empty()
        {
            return true;
        }
        index += character.len_utf8();
    }
    false
}

fn top_level_relation_index(source: &str) -> Option<usize> {
    const RELATION_COMMANDS: &[&str] = &[
        "\\Longleftrightarrow",
        "\\Longrightarrow",
        "\\Leftrightarrow",
        "\\Rightarrow",
        "\\leftrightarrow",
        "\\rightarrow",
        "\\leftarrow",
        "\\subseteq",
        "\\supseteq",
        "\\notin",
        "\\approx",
        "\\equiv",
        "\\simeq",
        "\\propto",
        "\\mapsto",
        "\\subset",
        "\\supset",
        "\\cong",
        "\\neq",
        "\\leq",
        "\\geq",
        "\\sim",
        "\\to",
        "\\ne",
        "\\le",
        "\\ge",
        "\\in",
    ];
    let mut brace_depth = 0_u32;
    let mut environments = Vec::new();
    let mut index = 0;
    while index < source.len() {
        if let Some((is_begin, name, end)) = read_latex_environment_token(source, index) {
            update_latex_environment_stack(&mut environments, is_begin, name);
            index = end;
            continue;
        }
        let character = source[index..].chars().next().expect("valid UTF-8");
        if character == '{' && !latex_character_is_escaped(source, index) {
            brace_depth += 1;
            index += 1;
            continue;
        }
        if character == '}' && !latex_character_is_escaped(source, index) {
            brace_depth = brace_depth.saturating_sub(1);
            index += 1;
            continue;
        }
        if brace_depth == 0 && environments.is_empty() {
            if matches!(character, '=' | '<' | '>') {
                return Some(index);
            }
            if character == '\\' {
                for command in RELATION_COMMANDS {
                    if !source[index..].starts_with(command) {
                        continue;
                    }
                    let next = source[index + command.len()..].chars().next();
                    if next.is_some_and(|value| value.is_ascii_alphabetic()) {
                        continue;
                    }
                    return Some(index);
                }
            }
        }
        index += character.len_utf8();
    }
    None
}

fn add_latex_alignment_marker(source: &str) -> String {
    if source.is_empty() || has_top_level_alignment_marker(source) {
        return source.to_string();
    }
    let Some(index) = top_level_relation_index(source) else {
        return source.to_string();
    };
    format!("{}&{}", &source[..index], &source[index..])
}

fn wrap_latex_environment(name: &str, body: &str) -> String {
    format!("\\begin{{{name}}}\n{body}\n\\end{{{name}}}")
}

fn format_document_formula_rows(lines: &[String], align_relations: bool) -> String {
    lines
        .iter()
        .enumerate()
        .map(|(index, line)| {
            let content = if align_relations {
                add_latex_alignment_marker(line)
            } else {
                line.clone()
            };
            if index + 1 < lines.len() {
                format!("{content} \\\\")
            } else {
                content
            }
        })
        .collect::<Vec<_>>()
        .join("\n")
}

fn canonical_document_formula_latex(
    metadata: &VisualTeXFormulaMetadata,
) -> Result<String, String> {
    // `metadata.lines` stores logical editor rows. A single logical formula may
    // itself contain source-formatting newlines, especially inside equation or
    // equation* environments. Keep those internal newlines inside the same row;
    // splitting them here would rebuild one equation as several equations and
    // make the Rust validator disagree with the TypeScript serializer.
    let mut lines = metadata
        .lines
        .iter()
        .map(|line| {
            normalize_mathlive_upright_commands(
                &line.latex.replace("\r\n", "\n").replace('\r', "\n"),
            )
            .trim()
            .to_string()
        })
        .filter(|line| !line.is_empty())
        .collect::<Vec<_>>();
    if lines.is_empty() {
        lines.push(String::new());
    }
    let joined = lines.join("\n");
    let formatted = match metadata.code_format.as_str() {
        "raw" => joined,
        "inline-dollar" => lines
            .iter()
            .map(|line| format!("${line}$"))
            .collect::<Vec<_>>()
            .join("\n"),
        "inline-paren" => lines
            .iter()
            .map(|line| format!("\\({line}\\)"))
            .collect::<Vec<_>>()
            .join("\n"),
        "display-dollar" => lines
            .iter()
            .map(|line| format!("$$\n{line}\n$$"))
            .collect::<Vec<_>>()
            .join("\n\n"),
        "display-bracket" => lines
            .iter()
            .map(|line| format!("\\[\n{line}\n\\]"))
            .collect::<Vec<_>>()
            .join("\n\n"),
        "equation" => lines
            .iter()
            .map(|line| wrap_latex_environment("equation", line))
            .collect::<Vec<_>>()
            .join("\n\n"),
        "equation-star" => lines
            .iter()
            .map(|line| wrap_latex_environment("equation*", line))
            .collect::<Vec<_>>()
            .join("\n\n"),
        "align" => wrap_latex_environment(
            "align",
            &format_document_formula_rows(&lines, true),
        ),
        "align-star" => wrap_latex_environment(
            "align*",
            &format_document_formula_rows(&lines, true),
        ),
        "aligned" => format!(
            "\\[\n{}\n\\]",
            wrap_latex_environment(
                "aligned",
                &format_document_formula_rows(&lines, true),
            )
        ),
        "gather" => wrap_latex_environment(
            "gather",
            &format_document_formula_rows(&lines, false),
        ),
        "gather-star" => wrap_latex_environment(
            "gather*",
            &format_document_formula_rows(&lines, false),
        ),
        "multline" => wrap_latex_environment(
            "multline",
            &format_document_formula_rows(&lines, false),
        ),
        "multline-star" => wrap_latex_environment(
            "multline*",
            &format_document_formula_rows(&lines, false),
        ),
        "equation-split" => wrap_latex_environment(
            "equation",
            &wrap_latex_environment(
                "split",
                &format_document_formula_rows(&lines, true),
            ),
        ),
        "equation-star-split" => wrap_latex_environment(
            "equation*",
            &wrap_latex_environment(
                "split",
                &format_document_formula_rows(&lines, true),
            ),
        ),
        _ => {
            return Err(format!(
                "Document formula metadata codeFormat is unsupported: {}",
                metadata.code_format
            ))
        }
    };
    Ok(formatted)
}

fn normalized_serialized_latex(source: &str) -> String {
    source
        .replace("\r\n", "\n")
        .replace('\r', "\n")
        .trim()
        .to_string()
}

fn validate_document_formula_metadata_match(
    metadata: &VisualTeXFormulaMetadata,
    formula_id: &str,
    latex: &str,
    display_mode: &str,
    numbered: bool,
) -> Result<String, String> {
    validate_metadata(metadata)?;
    if metadata.formula_id != formula_id
        || metadata.display_mode != display_mode
        || metadata.numbered != numbered
    {
        return Err("Document formula metadata identity does not match its formula block".to_string());
    }
    // The document-import frontend builds metadata, SVG, PNG and OMML from one
    // normalized editor document. Treat that submitted serialization as the
    // source of truth instead of rebuilding it again in Rust. Re-serializing
    // here creates false mismatches for harmless differences such as CRLF/LF,
    // environment whitespace, alignat arguments and internal equation
    // newlines, even though every rendered artifact belongs to the same
    // formula. Rust still validates the full metadata schema and identity above
    // and requires the formula block and metadata to carry the same non-empty
    // normalized source.
    let formula_latex = normalized_serialized_latex(latex);
    let metadata_latex = normalized_serialized_latex(&metadata.latex);
    if formula_latex.is_empty() || metadata_latex != formula_latex {
        return Err("Document formula metadata LaTeX does not match its formula block".to_string());
    }
    Ok(formula_latex)
}

fn encode_metadata(metadata: &VisualTeXFormulaMetadata) -> Result<String, String> {
    validate_metadata(metadata)?;
    let json = serde_json::to_vec(metadata)
        .map_err(|error| format!("Unable to encode VisualTeX formula metadata: {error}"))?;
    let mut encoder = DeflateEncoder::new(Vec::new(), Compression::best());
    encoder
        .write_all(&json)
        .map_err(|error| format!("Unable to compress VisualTeX formula metadata: {error}"))?;
    let compressed = encoder
        .finish()
        .map_err(|error| format!("Unable to finish VisualTeX formula metadata: {error}"))?;
    Ok(format!("{METADATA_PREFIX}{}", URL_SAFE_NO_PAD.encode(compressed)))
}

fn hex_encode(value: &str) -> String {
    value.as_bytes().iter().map(|byte| format!("{byte:02x}")).collect()
}

fn import_request(
    state: &OfficeCompanionState,
    request: MacOfflineSessionRequest,
) -> Result<OfficeFormulaSession, String> {
    match state.session_store.get(&request.session_id) {
        Ok(existing) => return Ok(existing),
        Err(SessionError::NotFound) => {}
        Err(error) => return Err(error.to_string()),
    }

    let original_metadata = request
        .encoded_metadata
        .as_deref()
        .map(decode_metadata)
        .transpose()?;
    let metadata_formula_id = original_metadata.as_ref().map(|value| value.formula_id.clone());
    let formula_id = match (request.formula_id.clone(), metadata_formula_id) {
        (Some(request_id), Some(metadata_id)) if request_id != metadata_id => {
            return Err("Request formulaId does not match encoded metadata".to_string())
        }
        (Some(request_id), _) => request_id,
        (None, Some(metadata_id)) => metadata_id,
        (None, None) if request.mode == "create" => Uuid::new_v4().to_string(),
        (None, None) => return Err("Edit request does not contain a formulaId".to_string()),
    };
    validate_uuid(&formula_id, "Imported formula id")?;

    let host = match request.host.as_str() {
        "word" => OfficeHost::Word,
        "powerpoint" => OfficeHost::Powerpoint,
        _ => unreachable!(),
    };
    let mode = match request.mode.as_str() {
        "create" => OfficeSessionMode::Create,
        "edit" => OfficeSessionMode::Edit,
        _ => unreachable!(),
    };
    let lines = original_metadata
        .as_ref()
        .map(|metadata| {
            metadata
                .lines
                .iter()
                .map(|line| FormulaLine {
                    id: line.id.clone(),
                    latex: line.latex.clone(),
                })
                .collect::<Vec<_>>()
        })
        .unwrap_or_else(|| {
            vec![FormulaLine {
                id: Uuid::new_v4().to_string(),
                latex: String::new(),
            }]
        });
    let source_document_id = match host {
        OfficeHost::Word => request.source_document_id.clone(),
        OfficeHost::Powerpoint => request.power_point.as_ref().map(|powerpoint| {
            format!(
                "visualtex-ppt-native-presentation:{}",
                powerpoint.presentation_identity
            )
        }),
    };
    let source_object_id = match host {
        OfficeHost::Word => match mode {
            OfficeSessionMode::Create => request.pending_marker.clone(),
            OfficeSessionMode::Edit => request
                .source_object_id
                .clone()
                .or_else(|| request.encoded_metadata.clone()),
        },
        OfficeHost::Powerpoint => request.power_point.as_ref().map(|powerpoint| {
            format!(
                "visualtex-ppt-native-edit:{}:{}",
                powerpoint.slide_index,
                hex_encode(&powerpoint.shape_name)
            )
        }),
    };
    let title = original_metadata
        .as_ref()
        .map(|metadata| metadata.title.clone())
        .unwrap_or_else(|| match host {
            OfficeHost::Word => "Word Formula".to_string(),
            OfficeHost::Powerpoint => "PowerPoint Formula".to_string(),
        });
    let code_format = original_metadata
        .as_ref()
        .map(|metadata| metadata.code_format.clone())
        .unwrap_or_else(|| "latex".to_string());

    let session_id = request.session_id.clone();
    match state
        .session_store
        .create_external(
            session_id.clone(),
            CreateOfficeSessionInput {
                mode,
                host,
                formula_id: Some(formula_id),
                source_document_id,
                source_object_id,
                title: Some(title),
                lines: Some(lines),
                active_line_id: None,
                code_format: Some(code_format),
                display_mode: Some(request.display_mode),
                numbered: Some(request.numbered),
                export_width: None,
                export_height: None,
                original_metadata,
                auto_commit_on_close: Some(true),
            },
        )
    {
        Ok(session) => Ok(session),
        Err(SessionError::Conflict(_)) => state
            .session_store
            .get(&session_id)
            .map_err(|error| error.to_string()),
        Err(error) => Err(error.to_string()),
    }
}

pub(crate) fn parse_office_url(value: &str) -> Result<String, String> {
    const PREFIX: &str = "visualtex://office/open?session=";
    let session_id = value
        .strip_prefix(PREFIX)
        .ok_or_else(|| "VisualTeX URL must use visualtex://office/open".to_string())?;
    if session_id.contains(['&', '#', '?', '/', '%']) {
        return Err("VisualTeX URL contains unsupported query data".to_string());
    }
    validate_uuid(session_id, "VisualTeX URL Session id")?;
    Ok(session_id.to_string())
}

fn office_host_name(host: OfficeHost) -> &'static str {
    match host {
        OfficeHost::Word => "word",
        OfficeHost::Powerpoint => "powerpoint",
    }
}

fn editor_window_label(host: OfficeHost) -> &'static str {
    match host {
        OfficeHost::Word => "office-native-word-editor",
        OfficeHost::Powerpoint => "office-native-powerpoint-editor",
    }
}

fn editor_window_host(label: &str) -> Option<OfficeHost> {
    match label {
        "office-native-word-editor" => Some(OfficeHost::Word),
        "office-native-powerpoint-editor" => Some(OfficeHost::Powerpoint),
        _ => None,
    }
}

fn document_import_window_label(session_id: &str) -> String {
    format!("office-native-document-{}", session_id.replace('-', ""))
}

fn active_editor_session(host: OfficeHost) -> Option<ActiveOfficeEditorSession> {
    office_editor_runtime()
        .lock()
        .ok()
        .and_then(|runtime| runtime.active(host).cloned())
}

fn clear_editor_session(
    host: OfficeHost,
    session_id: &str,
    generation: u64,
) -> Result<(), String> {
    let mut runtime = office_editor_runtime()
        .lock()
        .map_err(|_| "VisualTeX Office editor state is unavailable".to_string())?;
    let matches = runtime.active(host).is_some_and(|active| {
        active.activation.session_id == session_id
            && active.activation.generation == generation
    });
    if !matches {
        return Err("The Office editor Session is no longer active".to_string());
    }
    *runtime.active_mut(host) = None;
    Ok(())
}

fn clear_any_editor_session(host: OfficeHost) -> Option<MacOfflineOfficeEditorActivation> {
    let mut runtime = office_editor_runtime().lock().ok()?;
    runtime
        .active_mut(host)
        .take()
        .map(|active| active.activation)
}

#[cfg(target_os = "macos")]
fn set_resident_editor_parked(
    window: &WebviewWindow,
    parked: bool,
) -> Result<(), String> {
    window
        .with_webview(move |webview| unsafe {
            let native_window: &objc2_app_kit::NSWindow =
                &*webview.ns_window().cast();
            // Alpha 0 is treated as fully occluded by WebKit and suspends the
            // resident page exactly like hide(). A tiny non-zero alpha keeps
            // JavaScript/rendering active while remaining visually imperceptible.
            native_window.setAlphaValue(if parked { 0.01 } else { 1.0 });
            native_window.setIgnoresMouseEvents(parked);
        })
        .map_err(|error| {
            format!("Unable to update the resident Office editor window: {error}")
        })
}

#[cfg(not(target_os = "macos"))]
fn set_resident_editor_parked(
    window: &WebviewWindow,
    parked: bool,
) -> Result<(), String> {
    if parked {
        window.hide()
    } else {
        window.show()
    }
    .map_err(|error| error.to_string())
}

#[cfg(target_os = "macos")]
fn wake_resident_editor_for_hydration(window: &WebviewWindow) -> Result<(), String> {
    window
        .with_webview(move |webview| unsafe {
            let native_window: &objc2_app_kit::NSWindow =
                &*webview.ns_window().cast();
            native_window.setAlphaValue(1.0);
            native_window.setIgnoresMouseEvents(true);
        })
        .map_err(|error| {
            format!("Unable to wake the resident Office editor window: {error}")
        })
}

#[cfg(not(target_os = "macos"))]
fn wake_resident_editor_for_hydration(window: &WebviewWindow) -> Result<(), String> {
    window.show().map_err(|error| error.to_string())
}

#[cfg(target_os = "macos")]
fn present_resident_editor_window(window: &WebviewWindow) -> Result<(), String> {
    window
        .with_webview(move |webview| unsafe {
            let native_window: &objc2_app_kit::NSWindow =
                &*webview.ns_window().cast();
            // Hydration deliberately makes the resident window mouse-inert.
            // Restoring only Tauri focus is not sufficient on macOS: an
            // NSWindow can remain visible with ignoresMouseEvents=true, causing
            // every click to pass through to Word. Reset the native state and
            // explicitly make it the key normal-level window.
            native_window.setAlphaValue(1.0);
            native_window.setIgnoresMouseEvents(false);
            native_window.setLevel(objc2_app_kit::NSNormalWindowLevel);
            native_window.orderFrontRegardless();
            native_window.makeKeyAndOrderFront(None);
        })
        .map_err(|error| {
            format!("Unable to present the resident Office editor window: {error}")
        })
}

#[cfg(not(target_os = "macos"))]
fn present_resident_editor_window(window: &WebviewWindow) -> Result<(), String> {
    window.show().map_err(|error| error.to_string())?;
    window.set_focus().map_err(|error| error.to_string())
}

fn set_resident_editor_content_visible(
    window: &WebviewWindow,
    visible: bool,
) -> Result<(), String> {
    let opacity = if visible { "1" } else { "0" };
    window
        .eval(format!(
            "if (document.body) {{ document.body.style.opacity = '{opacity}'; }}"
        ))
        .map_err(|error| {
            format!("Unable to update resident Office editor content: {error}")
        })
}

fn create_editor_window(app: &AppHandle, host: OfficeHost) -> Result<WebviewWindow, String> {
    let label = editor_window_label(host);
    if let Some(window) = app.get_webview_window(label) {
        return Ok(window);
    }

    let theme = crate::persisted_app_theme(app);
    let path = format!(
        "office-native-dialog.html?transport=tauri&officeHost={}&theme={theme}",
        office_host_name(host),
    );
    let window = WebviewWindowBuilder::new(app, label, WebviewUrl::App(path.into()))
        .title("VisualTeX Office Formula")
        .inner_size(1180.0, 820.0)
        .min_inner_size(720.0, 560.0)
        .center()
        .focused(false)
        .skip_taskbar(true)
        .visible(false)
        .build()
        .map_err(|error| format!("Unable to initialize the VisualTeX Office editor: {error}"))?;
    // A truly hidden WKWebView is suspended by WebKit before React can mount,
    // which deadlocks the readiness handshake. Keep the resident editor
    // ordered on-screen at an imperceptible alpha and mouse-inert while parked.
    set_resident_editor_parked(&window, true)?;
    window
        .show()
        .map_err(|error| format!("Unable to prewarm the VisualTeX Office editor: {error}"))?;
    Ok(window)
}

pub(crate) fn prewarm_office_editor_windows(app: &AppHandle) -> Result<(), String> {
    // Icon decoding and both WebKit process creations happen once while the
    // resident app is starting. Formula double-clicks only switch a Session in
    // an already initialized, hidden WebView.
    crate::office::background::install_application_icon(app)?;
    for host in [OfficeHost::Word, OfficeHost::Powerpoint] {
        create_editor_window(app, host)?;
    }
    Ok(())
}

fn open_editor_window(
    app: &AppHandle,
    host: OfficeHost,
    session_id: &str,
    received_epoch_ms: u64,
    received_at: Instant,
) -> Result<(), String> {
    let label = editor_window_label(host);
    let reused = app.get_webview_window(label).is_some();
    let window = create_editor_window(app, host)?;
    set_resident_editor_content_visible(&window, false)?;
    let activation = {
        let mut runtime = office_editor_runtime()
            .lock()
            .map_err(|_| "VisualTeX Office editor state is unavailable".to_string())?;
        // Never expose the previous formula while the new Session hydrates.
        // Keep the resident WebView alive but transparent and mouse-inert so
        // WebKit cannot suspend React before the readiness handshake.
        set_resident_editor_parked(&window, true)?;
        runtime.next_generation = runtime.next_generation.saturating_add(1).max(1);
        let activation = MacOfflineOfficeEditorActivation {
            session_id: session_id.to_string(),
            host,
            generation: runtime.next_generation,
            received_epoch_ms,
        };
        *runtime.active_mut(host) = Some(ActiveOfficeEditorSession {
            activation: activation.clone(),
            received_at,
            ready: false,
        });
        activation
    };
    let elapsed_ms = received_at.elapsed().as_secs_f64() * 1000.0;
    queue_editor_performance(
        host,
        session_id,
        if reused { "window-reused" } else { "window-created" },
        elapsed_ms,
        Some(activation.generation),
        json!({ "windowLabel": label }),
    );
    if let Err(error) = window.emit(OFFICE_EDITOR_ACTIVATE_EVENT, activation.clone()) {
        let _ = clear_editor_session(host, session_id, activation.generation);
        return Err(format!("Unable to activate the VisualTeX Office editor: {error}"));
    }
    queue_editor_performance(
        host,
        session_id,
        "activation-event-sent",
        received_at.elapsed().as_secs_f64() * 1000.0,
        Some(activation.generation),
        json!({}),
    );
    // requestAnimationFrame is suspended while the native window remains
    // effectively invisible. Reveal only the blanked document and keep it
    // mouse-inert so React/MathLive can hydrate without flashing stale content.
    crate::office::background::activate_foreground_app(app)?;
    wake_resident_editor_for_hydration(&window)?;
    window.center().map_err(|error| error.to_string())?;
    window.show().map_err(|error| error.to_string())?;
    window.unminimize().map_err(|error| error.to_string())?;
    Ok(())
}

fn set_word_document_import_preparing_status() -> Result<(), String> {
    let output = Command::new("/usr/bin/osascript")
        .args([
            "-e",
            "tell application \"Microsoft Word\"",
            "-e",
            "activate",
            "-e",
            "set status bar to \"VisualTeX 正在准备批量导入…\"",
            "-e",
            "end tell",
        ])
        .output()
        .map_err(|error| format!("Unable to activate Microsoft Word: {error}"))?;
    if output.status.success() {
        Ok(())
    } else {
        let detail = String::from_utf8_lossy(&output.stderr).trim().to_string();
        Err(if detail.is_empty() {
            "Unable to activate Microsoft Word".to_string()
        } else {
            format!("Unable to activate Microsoft Word: {detail}")
        })
    }
}

fn clear_word_document_import_status() {
    let _ = Command::new("/usr/bin/osascript")
        .args([
            "-e",
            "tell application \"Microsoft Word\" to set status bar to \"\"",
        ])
        .output();
}

fn open_document_import_window(app: &AppHandle, session_id: &str) -> Result<(), String> {
    crate::office::background::activate_foreground_app(app)?;
    crate::office::background::install_application_icon(app)?;

    let label = document_import_window_label(session_id);
    if let Some(window) = app.get_webview_window(&label) {
        window.show().map_err(|error| error.to_string())?;
        window.set_focus().map_err(|error| error.to_string())?;
        return Ok(());
    }
    let theme = crate::persisted_app_theme(app);
    let path = format!(
        "index.html?view=office-document-import&sessionId={session_id}&transport=tauri&theme={theme}"
    );
    let window = WebviewWindowBuilder::new(app, label, WebviewUrl::App(path.into()))
        .title("VisualTeX Word 文档批量导入")
        .inner_size(1260.0, 840.0)
        .min_inner_size(860.0, 620.0)
        .center()
        .build()
        .map_err(|error| format!("Unable to open the VisualTeX document importer: {error}"))?;
    window.show().map_err(|error| error.to_string())?;
    crate::office::background::activate_foreground_app(app)?;
    window.set_focus().map_err(|error| error.to_string())?;
    Ok(())
}

#[tauri::command]
pub fn get_macos_offline_office_editor_activation(
    window: WebviewWindow,
) -> Result<Option<MacOfflineOfficeEditorActivation>, String> {
    let host = editor_window_host(window.label())
        .ok_or_else(|| "Only a VisualTeX Office formula editor can query activation".to_string())?;
    Ok(active_editor_session(host).map(|active| active.activation))
}

#[tauri::command]
pub fn report_macos_offline_office_editor_ready(
    app: AppHandle,
    window: WebviewWindow,
    input: MacOfflineOfficeEditorReadyInput,
) -> Result<(), String> {
    validate_uuid(&input.session_id, "Office editor Session id")?;
    for (value, label) in [
        (input.hydrate_ms, "hydrateMs"),
        (input.editor_mounted_ms, "editorMountedMs"),
        (input.content_ready_ms, "contentReadyMs"),
    ] {
        if !value.is_finite() || !(0.0..=120_000.0).contains(&value) {
            return Err(format!("Office editor {label} is invalid"));
        }
    }
    if input.hydrate_ms > input.editor_mounted_ms
        || input.editor_mounted_ms > input.content_ready_ms
    {
        return Err("Office editor readiness stages are out of order".to_string());
    }

    let host = editor_window_host(window.label())
        .ok_or_else(|| "Only a VisualTeX Office formula editor can report readiness".to_string())?;
    let mut runtime = office_editor_runtime()
        .lock()
        .map_err(|_| "VisualTeX Office editor state is unavailable".to_string())?;
    let active = runtime
        .active_mut(host)
        .as_mut()
        .ok_or_else(|| "The Office editor has no active Session".to_string())?;
    if active.activation.session_id != input.session_id
        || active.activation.generation != input.generation
    {
        return Err("Ignoring stale Office editor readiness".to_string());
    }
    active.ready = true;
    let active = active.clone();
    let report_received_ms = active.received_at.elapsed().as_secs_f64() * 1000.0;
    let frontend_origin_ms = (report_received_ms - input.content_ready_ms).max(0.0);

    queue_editor_performance(
        host,
        &input.session_id,
        "frontend-hydrated",
        frontend_origin_ms + input.hydrate_ms,
        Some(input.generation),
        json!({ "durationMs": input.hydrate_ms }),
    );
    queue_editor_performance(
        host,
        &input.session_id,
        "frontend-editor-mounted",
        frontend_origin_ms + input.editor_mounted_ms,
        Some(input.generation),
        json!({ "durationMs": input.editor_mounted_ms }),
    );
    queue_editor_performance(
        host,
        &input.session_id,
        "frontend-content-ready",
        frontend_origin_ms + input.content_ready_ms,
        Some(input.generation),
        json!({
            "durationMs": input.content_ready_ms,
            "frontendEpochMs": input.frontend_epoch_ms,
        }),
    );

    crate::office::background::activate_foreground_app(&app)?;
    set_resident_editor_content_visible(&window, true)?;
    present_resident_editor_window(&window)?;
    window.center().map_err(|error| error.to_string())?;
    window.show().map_err(|error| error.to_string())?;
    window.unminimize().map_err(|error| error.to_string())?;
    crate::office::background::activate_foreground_app(&app)?;
    window.set_focus().map_err(|error| error.to_string())?;
    // Apply the native key/mouse state once more after Tauri focus. This closes
    // a race where a queued hydration callback could otherwise restore the
    // earlier mouse-inert state after the window became visible.
    present_resident_editor_window(&window)?;
    let show_focus_ms = active.received_at.elapsed().as_secs_f64() * 1000.0;
    drop(runtime);
    let ready_epoch_ms = epoch_ms();
    queue_editor_performance(
        host,
        &input.session_id,
        "window-show-focus",
        show_focus_ms,
        Some(input.generation),
        json!({}),
    );

    let marker = MacOfflineOfficeEditorReadyMarker {
        schema: "visualtex-office-editor-ready-v1",
        session_id: input.session_id.clone(),
        host,
        generation: input.generation,
        epoch_ms: ready_epoch_ms,
        url_received_epoch_ms: active.activation.received_epoch_ms,
        frontend_epoch_ms: input.frontend_epoch_ms,
        hydrate_ms: input.hydrate_ms,
        editor_mounted_ms: input.editor_mounted_ms,
        content_ready_ms: input.content_ready_ms,
        show_focus_ms,
    };
    std::thread::spawn(move || {
        let Ok(path) = session_directory(host, &input.session_id)
            .map(|directory| directory.join(EDITOR_READY_FILE))
        else {
            return;
        };
        let Ok(bytes) = serde_json::to_vec_pretty(&marker) else {
            return;
        };
        let _ = atomic_write(&path, &bytes, 0o600);
    });
    Ok(())
}

#[tauri::command]
pub fn close_macos_offline_office_editor_window(
    window: WebviewWindow,
    session_id: Option<String>,
    generation: Option<u64>,
) -> Result<(), String> {
    let app = window.app_handle().clone();
    let Some(host) = editor_window_host(window.label()) else {
        if window.label().starts_with("office-native-document-") {
            window.destroy().map_err(|error| {
                format!("Unable to close the VisualTeX document importer: {error}")
            })?;
            #[cfg(target_os = "macos")]
            {
                let main_visible = app
                    .get_webview_window("main")
                    .and_then(|main| main.is_visible().ok())
                    .unwrap_or(false);
                if !has_open_office_editor(&app)
                    && !main_visible
                    && crate::office::background::is_background_mode()
                {
                    app.set_activation_policy(tauri::ActivationPolicy::Accessory)
                        .map_err(|error| {
                            format!(
                                "Unable to return VisualTeX to Office background mode: {error}"
                            )
                        })?;
                }
            }
            return Ok(());
        }
        return Err("Only a VisualTeX Office formula editor can close itself".to_string());
    };
    let session_id = session_id
        .ok_or_else(|| "The Office formula editor close request is missing sessionId".to_string())?;
    let generation = generation
        .ok_or_else(|| "The Office formula editor close request is missing generation".to_string())?;
    {
        // Keep activation validation and hiding atomic with respect to a new
        // URL activation. A late close from generation N must never hide the
        // already-hydrating generation N+1.
        let mut runtime = office_editor_runtime()
            .lock()
            .map_err(|_| "VisualTeX Office editor state is unavailable".to_string())?;
        let matches = runtime.active(host).is_some_and(|active| {
            active.activation.session_id == session_id
                && active.activation.generation == generation
        });
        if !matches {
            return Err("The Office editor Session is no longer active".to_string());
        }
        set_resident_editor_content_visible(&window, false)?;
        set_resident_editor_parked(&window, true).map_err(|error| {
            format!("Unable to close the VisualTeX Office editor: {error}")
        })?;
        *runtime.active_mut(host) = None;
    }
    let _ = window.emit(
        OFFICE_EDITOR_CLEAR_EVENT,
        json!({ "sessionId": session_id, "generation": generation }),
    );

    #[cfg(target_os = "macos")]
    {
        let main_visible = app
            .get_webview_window("main")
            .and_then(|main| main.is_visible().ok())
            .unwrap_or(false);
        if !has_open_office_editor(&app)
            && !main_visible
            && crate::office::background::is_background_mode()
        {
            app.set_activation_policy(tauri::ActivationPolicy::Accessory)
                .map_err(|error| format!("Unable to return VisualTeX to Office background mode: {error}"))?;
        }
    }
    Ok(())
}

pub(crate) fn has_open_office_editor(app: &AppHandle) -> bool {
    [OfficeHost::Word, OfficeHost::Powerpoint]
        .into_iter()
        .any(|host| {
            active_editor_session(host).is_some()
                && app.get_webview_window(editor_window_label(host)).is_some()
        })
}

pub(crate) fn focus_open_office_editor(app: &AppHandle) -> bool {
    for host in [OfficeHost::Word, OfficeHost::Powerpoint] {
        if let Some(active) = active_editor_session(host) {
            let Some(window) = app.get_webview_window(editor_window_label(host)) else {
                continue;
            };
            // A parked active window is still hydrating. Treat it as owned so
            // a second native double-click route cannot launch a duplicate,
            // but never focus transparent stale content before readiness.
            if !active.ready {
                return true;
            }
            let _ = crate::office::background::activate_foreground_app(app);
            let _ = present_resident_editor_window(&window);
            let _ = window.set_focus();
            let _ = present_resident_editor_window(&window);
            return true;
        }
    }
    false
}

pub(crate) fn handle_open_url(app: &AppHandle, value: &str) -> Result<(), String> {
    let received_at = Instant::now();
    let received_epoch_ms = epoch_ms();
    let session_id = parse_office_url(value)?;
    let state = app
        .try_state::<OfficeCompanionState>()
        .ok_or_else(|| "VisualTeX Office state is not initialized".to_string())?;
    // The host is not trusted until request validation completes, so queue the
    // first stage immediately after read_request resolves it below.
    let request = read_request(&session_id)?;
    let host = host_from_request_name(&request.host)?;
    queue_editor_performance(
        host,
        &session_id,
        "url-received",
        0.0,
        None,
        json!({ "receivedEpochMs": received_epoch_ms }),
    );
    queue_editor_performance(
        host,
        &session_id,
        "request-read",
        received_at.elapsed().as_secs_f64() * 1000.0,
        None,
        json!({}),
    );
    ensure_runtime_root(host)?;

    crate::office::background::hide_main_window(app)?;
    if request.operation.as_deref() == Some("documentImport") {
        for host in [OfficeHost::Word, OfficeHost::Powerpoint] {
            if let Some(window) = app.get_webview_window(editor_window_label(host)) {
                let _ = set_resident_editor_content_visible(&window, false);
                let _ = set_resident_editor_parked(&window, true);
                if let Some(active) = clear_any_editor_session(host) {
                    let _ = window.emit(
                        OFFICE_EDITOR_CLEAR_EVENT,
                        json!({
                            "sessionId": active.session_id,
                            "generation": active.generation,
                        }),
                    );
                }
            }
        }
        for (label, window) in app.webview_windows() {
            if label.starts_with("office-native-document-")
                && label != document_import_window_label(&session_id)
            {
                let _ = window.destroy();
            }
        }
        return open_document_import_window(app, &session_id);
    }

    import_request(state.inner(), request)?;
    queue_editor_performance(
        host,
        &session_id,
        "request-imported",
        received_at.elapsed().as_secs_f64() * 1000.0,
        None,
        json!({}),
    );
    open_editor_window(app, host, &session_id, received_epoch_ms, received_at)
}

fn decode_png(value: &str) -> Result<Vec<u8>, String> {
    let payload = value
        .split_once(',')
        .filter(|(prefix, _)| prefix.starts_with("data:image/png;base64"))
        .map(|(_, payload)| payload)
        .unwrap_or(value);
    let bytes = BASE64_STANDARD
        .decode(payload.trim())
        .map_err(|error| format!("Unable to decode the Office PNG export: {error}"))?;
    if bytes.len() < 8 || &bytes[..8] != b"\x89PNG\r\n\x1a\n" {
        return Err("Office formula export is not a valid PNG image".to_string());
    }
    Ok(bytes)
}

#[cfg(unix)]
fn set_mode(path: &Path, mode: u32) -> Result<(), String> {
    use std::os::unix::fs::PermissionsExt;
    fs::set_permissions(path, fs::Permissions::from_mode(mode))
        .map_err(|error| format!("Unable to set permissions on {}: {error}", path.display()))
}

#[cfg(not(unix))]
fn set_mode(_path: &Path, _mode: u32) -> Result<(), String> {
    Ok(())
}

fn atomic_write(path: &Path, bytes: &[u8], mode: u32) -> Result<(), String> {
    let parent = path
        .parent()
        .ok_or_else(|| format!("Path has no parent: {}", path.display()))?;
    fs::create_dir_all(parent)
        .map_err(|error| format!("Unable to create {}: {error}", parent.display()))?;
    set_mode(parent, 0o700)?;
    let temporary = parent.join(format!(
        ".{}.{}.tmp",
        path.file_name().and_then(|value| value.to_str()).unwrap_or("visualtex"),
        Uuid::new_v4()
    ));
    let mut file = OpenOptions::new()
        .create_new(true)
        .write(true)
        .open(&temporary)
        .map_err(|error| format!("Unable to create {}: {error}", temporary.display()))?;
    file.write_all(bytes)
        .map_err(|error| format!("Unable to write {}: {error}", temporary.display()))?;
    file.sync_all()
        .map_err(|error| format!("Unable to sync {}: {error}", temporary.display()))?;
    set_mode(&temporary, mode)?;
    fs::rename(&temporary, path).map_err(|error| {
        let _ = fs::remove_file(&temporary);
        format!("Unable to replace {}: {error}", path.display())
    })?;
    set_mode(path, mode)
}

fn sanitize_dispatch_value(value: &str, label: &str) -> Result<String, String> {
    if value.contains(['\r', '\n', '\0']) {
        return Err(format!("{label} contains unsupported control characters"));
    }
    Ok(value.to_string())
}

fn dispatch_text(entries: &[(&str, String)]) -> Result<String, String> {
    let dynamic = entries
        .iter()
        .map(|(key, value)| ((*key).to_string(), value.clone()))
        .collect::<Vec<_>>();
    dynamic_dispatch_text(&dynamic)
}

fn dynamic_dispatch_text(entries: &[(String, String)]) -> Result<String, String> {
    let mut seen = std::collections::HashSet::new();
    let mut output = String::new();
    for (key, value) in entries {
        if !seen.insert(key.as_str())
            || key.is_empty()
            || !key.bytes().all(|byte| byte.is_ascii_alphanumeric())
        {
            return Err("VisualTeX dispatch contains an invalid key".to_string());
        }
        output.push_str(key);
        output.push('=');
        output.push_str(&sanitize_dispatch_value(value, key)?);
        output.push('\n');
    }
    Ok(output)
}

fn run_vba_callback(host: OfficeHost) -> Result<(), String> {
    let script = match host {
        OfficeHost::Word => r#"tell application "Microsoft Word"
if not (exists active document) then error "Microsoft Word has no active document"
run VB macro macro name "VisualTeX_ApplyPendingResult"
end tell"#,
        OfficeHost::Powerpoint => r#"tell application "Microsoft PowerPoint"
if not (exists active presentation) then error "Microsoft PowerPoint has no active presentation"
run VB macro macro name "VisualTeX_ApplyPendingResult" list of parameters {}
end tell"#,
    };
    run_office_vba_script(script, "Office VBA callback")
}

pub(crate) fn run_double_click_edit_macro(host: OfficeHost) -> Result<(), String> {
    let script = match host {
        OfficeHost::Word => r#"tell application "Microsoft Word"
if not (exists active document) then error "Microsoft Word has no active document"
run VB macro macro name "VisualTeX_DoubleClickEditSelected"
end tell"#,
        OfficeHost::Powerpoint => r#"tell application "Microsoft PowerPoint"
if not (exists active presentation) then error "Microsoft PowerPoint has no active presentation"
run VB macro macro name "VisualTeX_DoubleClickEditSelected" list of parameters {}
end tell"#,
    };
    run_office_vba_script(script, "Office double-click edit macro")
}

pub(crate) fn run_word_image_double_click_edit_macro() -> Result<(), String> {
    run_office_vba_script(
        r#"tell application "Microsoft Word"
if not (exists active document) then error "Microsoft Word has no active document"
run VB macro macro name "VisualTeX_EditSelectedImageFromNativeMonitor"
end tell"#,
        "Word image double-click edit macro",
    )
}

fn run_office_vba_script(script: &str, label: &str) -> Result<(), String> {
    let output = Command::new("/usr/bin/osascript")
        .arg("-e")
        .arg(script)
        .output()
        .map_err(|error| format!("Unable to launch the {label}: {error}"))?;
    if output.status.success() {
        Ok(())
    } else {
        let detail = String::from_utf8_lossy(&output.stderr).trim().to_string();
        Err(if detail.is_empty() {
            format!("The {label} failed")
        } else {
            format!("The {label} failed: {detail}")
        })
    }
}

fn with_dispatch_pointer<T>(
    host: OfficeHost,
    session_id: &str,
    operation: impl FnOnce() -> Result<T, String>,
) -> Result<T, String> {
    let lock = match host {
        OfficeHost::Word => &WORD_DISPATCH_LOCK,
        OfficeHost::Powerpoint => &POWERPOINT_DISPATCH_LOCK,
    };
    let _guard = lock
        .lock()
        .map_err(|_| "VisualTeX Office dispatch lock is unavailable".to_string())?;
    let pointer = pointer_path(host)?;
    atomic_write(&pointer, session_id.as_bytes(), 0o600)?;
    let result = operation();
    let _ = fs::remove_file(pointer);
    result
}

fn scale_word_reference_geometry(
    reference_width_pt: f64,
    reference_height_pt: f64,
    reference_baseline_pt: f64,
    font_size_pt: f64,
) -> Result<WordGeometry, String> {
    if !reference_width_pt.is_finite()
        || !reference_height_pt.is_finite()
        || reference_width_pt <= 0.0
        || reference_height_pt <= 0.0
        || !reference_baseline_pt.is_finite()
        || !(-256.0..=0.0).contains(&reference_baseline_pt)
        || !font_size_pt.is_finite()
        || !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT).contains(&font_size_pt)
    {
        return Err("Word formula point-size reference geometry is invalid".to_string());
    }
    let point_scale = font_size_pt / WORD_REFERENCE_FONT_SIZE_PT;
    let width = reference_width_pt * point_scale;
    let height = reference_height_pt * point_scale;
    if !width.is_finite()
        || !height.is_finite()
        || width <= 0.0
        || height <= 0.0
        || width > 10_000.0
        || height > 10_000.0
    {
        return Err("Word formula point-size geometry is invalid".to_string());
    }
    let baseline = (reference_baseline_pt * point_scale)
        .round()
        .clamp(-256.0, 0.0) as i32;
    Ok(WordGeometry {
        width,
        height,
        baseline,
        font_size_pt,
        reference_width_pt,
        reference_height_pt,
        reference_baseline_pt,
    })
}

fn calculate_word_svg_geometry(
    width: f64,
    height: f64,
    baseline: Option<f64>,
    font_size_pt: f64,
) -> Result<WordGeometry, String> {
    if !width.is_finite()
        || !height.is_finite()
        || width <= 0.0
        || height <= 0.0
        || baseline.is_some_and(|value| !value.is_finite() || value < 0.0 || value > height)
    {
        return Err("Word formula SVG geometry is invalid".to_string());
    }
    let natural_width = width * 0.75 * WORD_IMAGE_VISUAL_SCALE;
    let natural_height = height * 0.75 * WORD_IMAGE_VISUAL_SCALE;
    let reference_scale = f64::min(1.0, MAX_WORD_WIDTH_PT / natural_width);
    let reference_width_pt = natural_width * reference_scale;
    let reference_height_pt = natural_height * reference_scale;
    let reference_baseline_pt = baseline
        .map(|value| {
            let descent_ratio = (height - value) / height;
            // Preserve the fractional descent at the canonical 14 pt size.
            // Word's Font.Position is integral, so rounding this reference and
            // rounding again after point-size scaling visibly under-corrects
            // short subscript formulas such as L_z beside L^2.
            -(reference_height_pt * descent_ratio).max(0.0)
        })
        .unwrap_or(0.0)
        .clamp(-256.0, 0.0);
    scale_word_reference_geometry(
        reference_width_pt,
        reference_height_pt,
        reference_baseline_pt,
        font_size_pt,
    )
}

fn calculate_word_geometry(
    request: &MacOfflineSessionRequest,
    session: &OfficeFormulaSession,
) -> Result<WordGeometry, String> {
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "Word Session has no formula export".to_string())?;
    if !export.width.is_finite()
        || !export.height.is_finite()
        || export.width <= 0.0
        || export.height <= 0.0
    {
        return Err("Word formula export has invalid dimensions".to_string());
    }

    let font_size_pt = request
        .font_size_pt
        .filter(|value| {
            value.is_finite()
                && *value >= MIN_WORD_FONT_SIZE_PT
                && *value <= MAX_WORD_FONT_SIZE_PT
        })
        .or_else(|| {
            session
                .original_metadata
                .as_ref()
                .and_then(|metadata| metadata.font_size_pt)
                .filter(|value| {
                    value.is_finite()
                        && *value >= MIN_WORD_FONT_SIZE_PT
                        && *value <= MAX_WORD_FONT_SIZE_PT
                })
        })
        .unwrap_or(WORD_REFERENCE_FONT_SIZE_PT);
    let baseline = export
        .baseline
        .filter(|value| value.is_finite() && *value >= 0.0 && *value <= export.height);
    // Document import and edit replacement share this exact 14 pt reference
    // geometry path, including maximum width, descent and baseline rounding.
    calculate_word_svg_geometry(
        export.width,
        export.height,
        baseline,
        font_size_pt,
    )
}

fn calculate_powerpoint_geometry(
    request: &MacOfflinePowerPointRequest,
    session: &OfficeFormulaSession,
) -> Result<PowerPointGeometry, String> {
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "PowerPoint Session has no formula export".to_string())?;
    if !export.width.is_finite()
        || !export.height.is_finite()
        || export.width <= 0.0
        || export.height <= 0.0
    {
        return Err("PowerPoint formula export has invalid dimensions".to_string());
    }

    // MathJax exports at a stable 14 pt reference size. The imported SVG keeps
    // its vector paths, so a PowerPoint point size is represented by uniformly
    // scaling the natural SVG bounds rather than rasterizing or stretching it.
    let reference_width_pt = export.width * 0.75;
    let reference_height_pt = export.height * 0.75;
    if !reference_width_pt.is_finite()
        || !reference_height_pt.is_finite()
        || reference_width_pt <= 0.0
        || reference_height_pt <= 0.0
    {
        return Err("PowerPoint SVG reference geometry is invalid".to_string());
    }

    let original = session.original_metadata.as_ref();
    let previous_reference_height = request
        .reference_height_pt
        .filter(|value| value.is_finite() && *value > 0.0)
        .or_else(|| {
            original
                .and_then(|metadata| metadata.reference_height_pt)
                .filter(|value| value.is_finite() && *value > 0.0)
        })
        .or_else(|| {
            original
                .and_then(|metadata| metadata.render_height_px)
                .filter(|value| value.is_finite() && *value > 0.0)
                .map(|value| value * 0.75)
        });

    let declared_font_size = request
        .font_size_pt
        .filter(|value| {
            value.is_finite()
                && *value >= MIN_POWERPOINT_FONT_SIZE_PT
                && *value <= MAX_POWERPOINT_FONT_SIZE_PT
        })
        .or_else(|| {
            original
                .and_then(|metadata| metadata.font_size_pt)
                .filter(|value| {
                    value.is_finite()
                        && *value >= MIN_POWERPOINT_FONT_SIZE_PT
                        && *value <= MAX_POWERPOINT_FONT_SIZE_PT
                })
        });

    // The actual selected shape height wins when a user resized an existing SVG
    // manually. This converts that physical height back to an equivalent point
    // size and prevents the next edit from jumping to a stale stored value.
    let observed_font_size = previous_reference_height
        .map(|height| request.height / height * POWERPOINT_REFERENCE_FONT_SIZE_PT)
        .filter(|value| {
            value.is_finite()
                && *value >= MIN_POWERPOINT_FONT_SIZE_PT
                && *value <= MAX_POWERPOINT_FONT_SIZE_PT
        });
    let font_size_pt = observed_font_size
        .or(declared_font_size)
        .unwrap_or(DEFAULT_POWERPOINT_FONT_SIZE_PT);
    let point_scale = font_size_pt / POWERPOINT_REFERENCE_FONT_SIZE_PT;
    let target_width = reference_width_pt * point_scale;
    let target_height = reference_height_pt * point_scale;
    let center_x = request.left + request.width / 2.0;
    let center_y = request.top + request.height / 2.0;
    let left = center_x - target_width / 2.0;
    let top = center_y - target_height / 2.0;
    for (value, label) in [
        (left, "target left"),
        (top, "target top"),
        (target_width, "target width"),
        (target_height, "target height"),
    ] {
        validate_finite_geometry(value, label)?;
    }
    if target_width <= 0.0
        || target_height <= 0.0
        || target_width > 10_000.0
        || target_height > 10_000.0
    {
        return Err("PowerPoint target formula dimensions are invalid".to_string());
    }
    Ok(PowerPointGeometry {
        left,
        top,
        width: target_width,
        height: target_height,
        font_size_pt,
        reference_width_pt,
        reference_height_pt,
    })
}

fn materialize_result_png(session: &OfficeFormulaSession) -> Result<PathBuf, String> {
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "Office Session has no formula export".to_string())?;
    let png = export
        .png_base64
        .as_deref()
        .ok_or_else(|| "Offline Office Session requires a PNG export".to_string())
        .and_then(decode_png)?;
    let path = result_png_path(session.host, &session.id)?;
    atomic_write(&path, &png, 0o600)?;
    Ok(path)
}

fn decode_svg(value: &str) -> Result<Vec<u8>, String> {
    let bytes = BASE64_STANDARD
        .decode(value.trim())
        .map_err(|error| format!("Unable to decode the Office SVG export: {error}"))?;
    if bytes.is_empty() || bytes.len() > MAX_METADATA_BYTES * 4 {
        return Err("Office formula SVG export is empty or too large".to_string());
    }
    let svg = std::str::from_utf8(&bytes)
        .map_err(|_| "Office formula SVG export is not UTF-8".to_string())?;
    let normalized = svg.trim_start();
    if !normalized.starts_with("<svg")
        && !(normalized.starts_with("<?xml") && normalized.contains("<svg"))
    {
        return Err("Office formula export is not a valid SVG document".to_string());
    }
    let lower = normalized.to_ascii_lowercase();
    for forbidden in [
        "<!doctype",
        "<!entity",
        "<script",
        "<foreignobject",
        "href=\"http:",
        "href=\"https:",
        "href=\"//",
        "xlink:href=\"http:",
        "xlink:href=\"https:",
        "xlink:href=\"//",
    ] {
        if lower.contains(forbidden) {
            return Err("Office formula SVG export contains unsafe external content".to_string());
        }
    }
    Ok(bytes)
}

fn materialize_result_svg(session: &OfficeFormulaSession) -> Result<PathBuf, String> {
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "Office Session has no formula export".to_string())?;
    let svg = decode_svg(&export.svg_base64)?;
    let path = result_svg_path(session.host, &session.id)?;
    atomic_write(&path, &svg, 0o600)?;
    Ok(path)
}

fn crc32(bytes: &[u8]) -> u32 {
    let mut crc = 0xffff_ffff_u32;
    for byte in bytes {
        crc ^= u32::from(*byte);
        for _ in 0..8 {
            let mask = 0_u32.wrapping_sub(crc & 1);
            crc = (crc >> 1) ^ (0xedb8_8320 & mask);
        }
    }
    !crc
}

fn push_zip_u16(output: &mut Vec<u8>, value: u16) {
    output.extend_from_slice(&value.to_le_bytes());
}

fn push_zip_u32(output: &mut Vec<u8>, value: u32) {
    output.extend_from_slice(&value.to_le_bytes());
}

struct StoredZipEntry {
    name: Vec<u8>,
    crc: u32,
    size: u32,
    offset: u32,
}

fn build_stored_zip(entries: &[(&str, &[u8])]) -> Result<Vec<u8>, String> {
    let entry_count = u16::try_from(entries.len())
        .map_err(|_| "Word SVG staging package has too many ZIP entries".to_string())?;
    let mut output = Vec::new();
    let mut records = Vec::with_capacity(entries.len());

    for (name, contents) in entries {
        let name_bytes = name.as_bytes();
        let name_length = u16::try_from(name_bytes.len())
            .map_err(|_| "Word SVG staging package contains an overlong ZIP path".to_string())?;
        let size = u32::try_from(contents.len())
            .map_err(|_| "Word SVG staging package entry is too large".to_string())?;
        let offset = u32::try_from(output.len())
            .map_err(|_| "Word SVG staging package is too large".to_string())?;
        let checksum = crc32(contents);

        push_zip_u32(&mut output, 0x0403_4b50);
        push_zip_u16(&mut output, 20);
        push_zip_u16(&mut output, 0x0800);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 33);
        push_zip_u32(&mut output, checksum);
        push_zip_u32(&mut output, size);
        push_zip_u32(&mut output, size);
        push_zip_u16(&mut output, name_length);
        push_zip_u16(&mut output, 0);
        output.extend_from_slice(name_bytes);
        output.extend_from_slice(contents);

        records.push(StoredZipEntry {
            name: name_bytes.to_vec(),
            crc: checksum,
            size,
            offset,
        });
    }

    let central_offset = u32::try_from(output.len())
        .map_err(|_| "Word SVG staging package is too large".to_string())?;
    for record in &records {
        let name_length = u16::try_from(record.name.len())
            .map_err(|_| "Word SVG staging package contains an overlong ZIP path".to_string())?;
        push_zip_u32(&mut output, 0x0201_4b50);
        push_zip_u16(&mut output, 20);
        push_zip_u16(&mut output, 20);
        push_zip_u16(&mut output, 0x0800);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 33);
        push_zip_u32(&mut output, record.crc);
        push_zip_u32(&mut output, record.size);
        push_zip_u32(&mut output, record.size);
        push_zip_u16(&mut output, name_length);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 0);
        push_zip_u32(&mut output, 0);
        push_zip_u32(&mut output, record.offset);
        output.extend_from_slice(&record.name);
    }
    let central_size = u32::try_from(output.len())
        .map_err(|_| "Word SVG staging package is too large".to_string())?
        .checked_sub(central_offset)
        .ok_or_else(|| "Word SVG staging package central directory is invalid".to_string())?;

    push_zip_u32(&mut output, 0x0605_4b50);
    push_zip_u16(&mut output, 0);
    push_zip_u16(&mut output, 0);
    push_zip_u16(&mut output, entry_count);
    push_zip_u16(&mut output, entry_count);
    push_zip_u32(&mut output, central_size);
    push_zip_u32(&mut output, central_offset);
    push_zip_u16(&mut output, 0);
    Ok(output)
}

fn build_word_svg_docx(
    svg: &[u8],
    png: &[u8],
    width_points: f64,
    height_points: f64,
) -> Result<Vec<u8>, String> {
    let width_emu = (width_points * 12_700.0).round();
    let height_emu = (height_points * 12_700.0).round();
    if !width_emu.is_finite()
        || !height_emu.is_finite()
        || width_emu <= 0.0
        || height_emu <= 0.0
        || width_emu > i64::MAX as f64
        || height_emu > i64::MAX as f64
    {
        return Err("Word SVG staging package dimensions are invalid".to_string());
    }
    let width_emu = width_emu as i64;
    let height_emu = height_emu as i64;

    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="svg" ContentType="image/svg+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;
    let package_relationships = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;
    let document_relationships = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPng" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/formula.png"/>
  <Relationship Id="rIdSvg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/formula.svg"/>
</Relationships>"#;
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="{width_emu}" cy="{height_emu}"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:docPr id="1" name="VisualTeX Formula" descr="VisualTeX SVG formula"/>
            <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr><pic:cNvPr id="0" name="formula.svg"/><pic:cNvPicPr/></pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="rIdPng" cstate="print">
                      <a:extLst>
                        <a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext>
                      </a:extLst>
                    </a:blip>
                    <a:stretch><a:fillRect/></a:stretch>
                  </pic:blipFill>
                  <pic:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="{width_emu}" cy="{height_emu}"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:noFill/><a:ln><a:noFill/></a:ln>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
  </w:body>
</w:document>"#
    );

    build_stored_zip(&[
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", package_relationships),
        ("word/document.xml", document.as_bytes()),
        ("word/_rels/document.xml.rels", document_relationships),
        ("word/media/formula.png", png),
        ("word/media/formula.svg", svg),
    ])
}

fn materialize_word_svg_package(
    session: &OfficeFormulaSession,
    geometry: WordGeometry,
) -> Result<(PathBuf, PathBuf, PathBuf), String> {
    if session.host != OfficeHost::Word {
        return Err("Word SVG package materialization requires a Word Session".to_string());
    }
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "Word Session has no formula export".to_string())?;
    let svg = decode_svg(&export.svg_base64)?;
    let png = export
        .png_base64
        .as_deref()
        .ok_or_else(|| "Word SVG staging requires a PNG compatibility preview".to_string())
        .and_then(decode_png)?;
    let svg_path = result_svg_path(OfficeHost::Word, &session.id)?;
    let png_path = result_png_path(OfficeHost::Word, &session.id)?;
    let document_path = result_word_svg_docx_path(&session.id)?;
    atomic_write(&svg_path, &svg, 0o600)?;
    atomic_write(&png_path, &png, 0o600)?;
    let package = build_word_svg_docx(&svg, &png, geometry.width, geometry.height)?;
    atomic_write(&document_path, &package, 0o600)?;
    Ok((svg_path, document_path, png_path))
}

fn materialize_powerpoint_svg(session: &OfficeFormulaSession) -> Result<PathBuf, String> {
    if session.host != OfficeHost::Powerpoint {
        return Err("PowerPoint SVG materialization requires a PowerPoint Session".to_string());
    }
    materialize_result_svg(session)
}

fn commit_word(
    request: &MacOfflineSessionRequest,
    session: &OfficeFormulaSession,
    metadata: &str,
    canonical_latex: &str,
    geometry: WordGeometry,
) -> Result<(), String> {
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "Word Session has no formula export".to_string())?;
    let omml_base64 = export
        .omml_base64
        .as_deref()
        .ok_or_else(|| "Word formula export has no OMML payload".to_string())?;
    let omml_bytes = URL_SAFE_NO_PAD
        .decode(omml_base64)
        .map_err(|_| "Word formula OMML payload is not valid Base64URL".to_string())?;
    if omml_bytes.is_empty() || omml_bytes.len() > MAX_OMML_BYTES {
        return Err("Word formula OMML payload is empty or too large".to_string());
    }
    let omml = std::str::from_utf8(&omml_bytes)
        .map_err(|_| "Word formula OMML payload is not UTF-8".to_string())?;
    if !omml.trim_start().starts_with("<m:oMath")
        || !omml.contains("http://schemas.openxmlformats.org/officeDocument/2006/math")
        || omml.contains("<!DOCTYPE")
        || omml.contains("<!ENTITY")
    {
        return Err("Word formula OMML payload is not a safe Office Math fragment".to_string());
    }
    let omml_docx_base64 = export
        .omml_docx_base64
        .as_deref()
        .ok_or_else(|| "Word formula export has no native DOCX payload".to_string())?;
    let omml_docx = URL_SAFE_NO_PAD
        .decode(omml_docx_base64)
        .map_err(|_| "Word formula native DOCX payload is not valid Base64URL".to_string())?;
    if omml_docx.len() < 128 || omml_docx.len() > MAX_OMML_BYTES * 8 || !omml_docx.starts_with(b"PK\x03\x04") {
        return Err("Word formula native DOCX payload is invalid or too large".to_string());
    }
    let native_document_path = native_word_document_path(&session.formula_id)?;
    atomic_write(&native_document_path, &omml_docx, 0o600)?;

    // Word for Mac exposes SVG in the document format but its VBA
    // InlineShapes.AddPicture API rejects a raw .svg file. Materialize a tiny
    // DOCX containing an SVG blip plus a PNG compatibility preview, then let
    // Word import its already-parsed InlineShape through FormattedText.
    let (image_path, vector_document_path, fallback_image_path) =
        materialize_word_svg_package(session, geometry)?;
    let source_marker = request
        .source_object_id
        .clone()
        .or_else(|| request.encoded_metadata.clone())
        .unwrap_or_default();
    let pending_marker = request.pending_marker.clone().unwrap_or_default();
    let latex = canonical_latex.trim();
    if latex.is_empty() {
        return Err("Word native-equation conversion requires non-empty LaTeX".to_string());
    }
    let latex_base64 = URL_SAFE_NO_PAD.encode(latex.as_bytes());
    let dispatch = dispatch_text(&[
        ("protocolVersion", OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId", session.id.clone()),
        ("action", "commit".to_string()),
        ("host", "word".to_string()),
        ("mode", request.mode.clone()),
        ("formulaId", session.formula_id.clone()),
        ("displayMode", session.display_mode.clone()),
        ("numbered", if session.numbered { "1" } else { "0" }.to_string()),
        (
            "nativeEquation",
            if request.native_equation { "1" } else { "0" }.to_string(),
        ),
        ("imagePath", image_path.to_string_lossy().to_string()),
        (
            "vectorDocumentPath",
            vector_document_path.to_string_lossy().to_string(),
        ),
        (
            "fallbackImagePath",
            fallback_image_path.to_string_lossy().to_string(),
        ),
        ("metadata", metadata.to_string()),
        ("latexBase64", latex_base64),
        ("ommlBase64", omml_base64.to_string()),
        (
            "nativeDocumentPath",
            native_document_path.to_string_lossy().to_string(),
        ),
        ("pendingMarker", pending_marker),
        ("sourceMarker", source_marker),
        (
            "sourceDocumentId",
            request.source_document_id.clone().unwrap_or_default(),
        ),
        ("widthPoints", format!("{:.6}", geometry.width)),
        ("heightPoints", format!("{:.6}", geometry.height)),
        ("baseline", geometry.baseline.to_string()),
        ("fontSizePt", format!("{:.6}", geometry.font_size_pt)),
        (
            "referenceWidthPt",
            format!("{:.6}", geometry.reference_width_pt),
        ),
        (
            "referenceHeightPt",
            format!("{:.6}", geometry.reference_height_pt),
        ),
        (
            "referenceBaselinePt",
            format!("{:.6}", geometry.reference_baseline_pt),
        ),
    ])?;
    atomic_write(
        &dispatch_path(OfficeHost::Word, &session.id)?,
        dispatch.as_bytes(),
        0o600,
    )?;
    with_dispatch_pointer(OfficeHost::Word, &session.id, || {
        run_vba_callback(OfficeHost::Word)
    })?;
    Ok(())
}

fn commit_powerpoint(
    request: &MacOfflineSessionRequest,
    session: &OfficeFormulaSession,
    metadata: &str,
    geometry: PowerPointGeometry,
) -> Result<(), String> {
    let powerpoint = request
        .power_point
        .as_ref()
        .ok_or_else(|| "PowerPoint request geometry is missing".to_string())?;
    let image_path = materialize_powerpoint_svg(session)?;
    let fallback_image_path = materialize_result_png(session).ok();
    let dispatch = dispatch_text(&[
        ("protocolVersion", OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId", session.id.clone()),
        ("action", "commit".to_string()),
        ("host", "powerpoint".to_string()),
        ("mode", request.mode.clone()),
        ("formulaId", session.formula_id.clone()),
        ("displayMode", "block".to_string()),
        ("numbered", "0".to_string()),
        ("imagePath", image_path.to_string_lossy().to_string()),
        (
            "fallbackImagePath",
            fallback_image_path
                .as_ref()
                .map(|path| path.to_string_lossy().to_string())
                .unwrap_or_default(),
        ),
        ("metadata", metadata.to_string()),
        (
            "pendingMarker",
            request.pending_marker.clone().unwrap_or_default(),
        ),
        ("sourceMarker", request.encoded_metadata.clone().unwrap_or_default()),
        ("sourceShapeName", powerpoint.shape_name.clone()),
        ("shapeName", format!("VisualTeX_{}", session.formula_id)),
        ("targetLeft", format!("{:.6}", geometry.left)),
        ("targetTop", format!("{:.6}", geometry.top)),
        ("targetWidth", format!("{:.6}", geometry.width)),
        ("targetHeight", format!("{:.6}", geometry.height)),
        ("fontSizePt", format!("{:.6}", geometry.font_size_pt)),
        (
            "referenceWidthPt",
            format!("{:.6}", geometry.reference_width_pt),
        ),
        (
            "referenceHeightPt",
            format!("{:.6}", geometry.reference_height_pt),
        ),
        ("rotation", format!("{:.6}", powerpoint.rotation)),
        ("zOrder", powerpoint.z_order.to_string()),
        ("presentationIdentity", powerpoint.presentation_identity.clone()),
        ("slideIndex", powerpoint.slide_index.to_string()),
        ("slideId", powerpoint.slide_id.to_string()),
    ])?;
    atomic_write(
        &dispatch_path(OfficeHost::Powerpoint, &session.id)?,
        dispatch.as_bytes(),
        0o600,
    )?;
    with_dispatch_pointer(OfficeHost::Powerpoint, &session.id, || {
        run_vba_callback(OfficeHost::Powerpoint)
    })?;
    Ok(())
}

fn cancel_host(request: &MacOfflineSessionRequest) -> Result<(), String> {
    let host = if request.host == "word" {
        OfficeHost::Word
    } else {
        OfficeHost::Powerpoint
    };
    let entries = vec![
        ("protocolVersion", OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId", request.session_id.clone()),
        ("action", "cancel".to_string()),
        ("host", request.host.clone()),
        ("mode", request.mode.clone()),
        (
            "pendingMarker",
            request.pending_marker.clone().unwrap_or_default(),
        ),
        (
            "sourceDocumentId",
            request.source_document_id.clone().unwrap_or_default(),
        ),
    ];
    let dispatch = dispatch_text(&entries)?;
    atomic_write(
        &dispatch_path(host, &request.session_id)?,
        dispatch.as_bytes(),
        0o600,
    )?;
    if request.mode == "create" {
        with_dispatch_pointer(host, &request.session_id, || run_vba_callback(host))?;
    }
    Ok(())
}

fn document_import_request_data(
    request: &MacOfflineSessionRequest,
) -> Result<MacOfflineDocumentImportPublicRequest, String> {
    if request.operation.as_deref() != Some("documentImport") || request.host != "word" {
        return Err("Offline Office request is not a Word document import".to_string());
    }
    let document_import = request
        .document_import
        .as_ref()
        .ok_or_else(|| "Document import request is missing insertion information".to_string())?;
    Ok(MacOfflineDocumentImportPublicRequest {
        protocol_version: request.protocol_version,
        session_id: request.session_id.clone(),
        host: request.host.clone(),
        source_document_id: request
            .source_document_id
            .clone()
            .ok_or_else(|| "Document import request is missing the Word document identity".to_string())?,
        bookmark_name: document_import.bookmark_name.clone(),
        default_font_size_pt: document_import.default_font_size_pt,
    })
}

fn document_formula_file_path(
    session_id: &str,
    formula_id: &str,
    extension: &str,
) -> Result<PathBuf, String> {
    validate_uuid(formula_id, "Document formula id")?;
    if !extension.bytes().all(|byte| byte.is_ascii_alphanumeric()) {
        return Err("Document formula file extension is invalid".to_string());
    }
    Ok(session_directory(OfficeHost::Word, session_id)?.join(format!(
        "document-formula-{formula_id}.{extension}"
    )))
}

fn validate_document_omml_payload(value: &str) -> Result<(), String> {
    let bytes = URL_SAFE_NO_PAD
        .decode(value)
        .map_err(|_| "Document formula OMML payload is not valid Base64URL".to_string())?;
    if bytes.is_empty() || bytes.len() > MAX_OMML_BYTES {
        return Err("Document formula OMML payload is empty or too large".to_string());
    }
    let omml = std::str::from_utf8(&bytes)
        .map_err(|_| "Document formula OMML payload is not UTF-8".to_string())?;
    if !omml.trim_start().starts_with("<m:oMath")
        || !omml.contains("http://schemas.openxmlformats.org/officeDocument/2006/math")
        || omml.contains("<!DOCTYPE")
        || omml.contains("<!ENTITY")
    {
        return Err("Document formula OMML payload is not a safe Office Math fragment".to_string());
    }
    Ok(())
}

fn decode_document_native_docx(value: &str) -> Result<Vec<u8>, String> {
    let bytes = URL_SAFE_NO_PAD
        .decode(value)
        .map_err(|_| "Document formula native DOCX is not valid Base64URL".to_string())?;
    if bytes.len() < 128
        || bytes.len() > MAX_OMML_BYTES * 8
        || !bytes.starts_with(b"PK\x03\x04")
    {
        return Err("Document formula native DOCX payload is invalid or too large".to_string());
    }
    Ok(bytes)
}

fn decode_document_image_fallback_png(value: Option<&str>) -> Result<Vec<u8>, String> {
    let value = value.ok_or_else(|| {
        "Image document formula is missing its PNG compatibility preview".to_string()
    })?;
    let bytes = decode_png(value)?;
    if bytes.len() < 24 {
        return Err("Image document formula PNG compatibility preview is truncated".to_string());
    }
    let width = u32::from_be_bytes(bytes[16..20].try_into().map_err(|_| {
        "Image document formula PNG compatibility preview has an invalid IHDR".to_string()
    })?);
    let height = u32::from_be_bytes(bytes[20..24].try_into().map_err(|_| {
        "Image document formula PNG compatibility preview has an invalid IHDR".to_string()
    })?);
    if width <= 1 || height <= 1 {
        return Err(
            "Image document formula PNG compatibility preview must contain rendered formula pixels"
                .to_string(),
        );
    }
    Ok(bytes)
}

fn calculate_document_image_geometry(
    width: f64,
    height: f64,
    baseline: f64,
    font_size_pt: f64,
) -> Result<WordGeometry, String> {
    if !width.is_finite()
        || !height.is_finite()
        || !baseline.is_finite()
        || width <= 0.0
        || height <= 0.0
        || baseline < 0.0
        || baseline > height
    {
        return Err("Document formula SVG geometry is invalid".to_string());
    }
    calculate_word_svg_geometry(width, height, Some(baseline), font_size_pt)
}

fn resolve_document_paragraph_transfer(
    paragraph_id: &Option<String>,
    paragraph_style: &Option<String>,
    paragraph_alignment: &Option<String>,
    list_kind: &Option<String>,
    list_level: Option<u32>,
    paragraph_start: bool,
    paragraph_end: bool,
) -> Result<Option<DocumentParagraphTransfer>, String> {
    let Some(id) = paragraph_id.as_deref() else {
        if paragraph_style.is_some()
            || paragraph_alignment.is_some()
            || list_kind.is_some()
            || list_level.is_some()
            || paragraph_start
            || paragraph_end
        {
            return Err("Document paragraph metadata is missing its paragraph id".to_string());
        }
        return Ok(None);
    };
    validate_uuid(id, "Document paragraph id")?;
    let style = paragraph_style.as_deref().unwrap_or("normal");
    if !matches!(
        style,
        "normal" | "heading1" | "heading2" | "heading3" | "heading4" | "quote" | "code"
    ) {
        return Err("Document paragraph style is invalid".to_string());
    }
    let alignment = paragraph_alignment.as_deref().unwrap_or("left");
    if !matches!(alignment, "left" | "center" | "right" | "justify") {
        return Err("Document paragraph alignment is invalid".to_string());
    }
    let resolved_list_kind = list_kind.as_deref().unwrap_or("none");
    if !matches!(resolved_list_kind, "none" | "bullet" | "number") {
        return Err("Document paragraph list kind is invalid".to_string());
    }
    let resolved_list_level = list_level.unwrap_or(0);
    if (resolved_list_kind == "none" && resolved_list_level != 0)
        || (resolved_list_kind != "none" && !(1..=9).contains(&resolved_list_level))
    {
        return Err("Document paragraph list level is invalid".to_string());
    }
    Ok(Some(DocumentParagraphTransfer {
        id: id.to_string(),
        style: style.to_string(),
        alignment: alignment.to_string(),
        list_kind: resolved_list_kind.to_string(),
        list_level: resolved_list_level,
        start: paragraph_start,
        end: paragraph_end,
    }))
}

fn append_document_paragraph_entries(
    entries: &mut Vec<(String, String)>,
    prefix: &str,
    paragraph: Option<&DocumentParagraphTransfer>,
) {
    let value = |field: &str| format!("{prefix}{field}");
    if let Some(paragraph) = paragraph {
        entries.push((value("paragraphId"), paragraph.id.clone()));
        entries.push((value("paragraphStyle"), paragraph.style.clone()));
        entries.push((value("paragraphAlignment"), paragraph.alignment.clone()));
        entries.push((value("listKind"), paragraph.list_kind.clone()));
        entries.push((value("listLevel"), paragraph.list_level.to_string()));
        entries.push((
            value("paragraphStart"),
            if paragraph.start { "1" } else { "0" }.to_string(),
        ));
        entries.push((
            value("paragraphEnd"),
            if paragraph.end { "1" } else { "0" }.to_string(),
        ));
    }
}

fn commit_document_import_blocking(
    state: OfficeCompanionState,
    session_id: String,
    input: MacOfflineDocumentImportCommitInput,
) -> Result<(), String> {
    validate_uuid(&session_id, "Session id")?;
    let request = read_request(&session_id)?;
    let public_request = document_import_request_data(&request)?;
    if input.output_kind != "omml" && input.output_kind != "image" {
        return Err("Document formula output kind must be omml or image".to_string());
    }
    if input.items.is_empty() || input.items.len() > 2048 {
        return Err("Document import must contain 1 to 2048 blocks".to_string());
    }

    let mut entries = vec![
        ("protocolVersion".to_string(), OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId".to_string(), session_id.clone()),
        ("outputKind".to_string(), input.output_kind.clone()),
        (
            "sourceDocumentId".to_string(),
            public_request.source_document_id.clone(),
        ),
        ("bookmarkName".to_string(), public_request.bookmark_name.clone()),
        ("itemCount".to_string(), input.items.len().to_string()),
    ];
    let mut metadata_to_cache = Vec::new();
    let mut formula_count = 0usize;
    let mut text_bytes = 0usize;
    let mut active_paragraph_id: Option<String> = None;

    for (index, item) in input.items.iter().enumerate() {
        let prefix = format!("item{index}");
        match item {
            MacOfflineDocumentImportCommitItem::Text {
                text,
                paragraph_id,
                paragraph_style,
                paragraph_alignment,
                list_kind,
                list_level,
                paragraph_start,
                paragraph_end,
            } => {
                let paragraph = resolve_document_paragraph_transfer(
                    paragraph_id,
                    paragraph_style,
                    paragraph_alignment,
                    list_kind,
                    *list_level,
                    *paragraph_start,
                    *paragraph_end,
                )?;
                if let Some(paragraph) = paragraph.as_ref() {
                    if paragraph.start {
                        if active_paragraph_id.is_some() {
                            return Err("Document paragraphs overlap in the transfer stream".to_string());
                        }
                        active_paragraph_id = Some(paragraph.id.clone());
                    } else if active_paragraph_id.as_deref() != Some(paragraph.id.as_str()) {
                        return Err("Document paragraph continuation has no matching start".to_string());
                    }
                } else if active_paragraph_id.is_some() {
                    return Err("Document paragraph content is missing paragraph metadata".to_string());
                }
                text_bytes = text_bytes.saturating_add(text.len());
                if text_bytes > 4 * 1024 * 1024 {
                    return Err("Document import text exceeds the 4 MB limit".to_string());
                }
                entries.push((format!("{prefix}kind"), "text".to_string()));
                entries.push((
                    format!("{prefix}textBase64"),
                    URL_SAFE_NO_PAD.encode(text.as_bytes()),
                ));
                append_document_paragraph_entries(&mut entries, &prefix, paragraph.as_ref());
                if paragraph.as_ref().is_some_and(|value| value.end) {
                    active_paragraph_id = None;
                }
            }
            MacOfflineDocumentImportCommitItem::Formula {
                formula_id,
                latex,
                display_mode,
                numbered,
                font_size_pt,
                metadata,
                omml_base64,
                omml_docx_base64,
                svg_base64,
                png_base64,
                width,
                height,
                baseline,
                paragraph_id,
                paragraph_style,
                paragraph_alignment,
                list_kind,
                list_level,
                paragraph_start,
                paragraph_end,
            } => {
                let paragraph = resolve_document_paragraph_transfer(
                    paragraph_id,
                    paragraph_style,
                    paragraph_alignment,
                    list_kind,
                    *list_level,
                    *paragraph_start,
                    *paragraph_end,
                )?;
                if display_mode == "block" && paragraph.is_some() {
                    return Err("Display formulas must own their Word paragraph".to_string());
                }
                if let Some(paragraph) = paragraph.as_ref() {
                    if paragraph.start {
                        if active_paragraph_id.is_some() {
                            return Err("Document paragraphs overlap in the transfer stream".to_string());
                        }
                        active_paragraph_id = Some(paragraph.id.clone());
                    } else if active_paragraph_id.as_deref() != Some(paragraph.id.as_str()) {
                        return Err("Document paragraph continuation has no matching start".to_string());
                    }
                } else if active_paragraph_id.is_some() {
                    return Err("Document paragraph content is missing paragraph metadata".to_string());
                }
                formula_count += 1;
                if formula_count > 512 {
                    return Err("Document import supports at most 512 formulas".to_string());
                }
                validate_uuid(formula_id, "Document formula id")?;
                if latex.trim().is_empty() || latex.len() > 1_000_000 || latex.contains('\0') {
                    return Err("A document formula contains invalid or excessive LaTeX".to_string());
                }
                if !matches!(display_mode.as_str(), "inline" | "block") {
                    return Err("Document formula display mode is invalid".to_string());
                }
                if *numbered && display_mode != "block" {
                    return Err("Only document display formulas can be numbered".to_string());
                }
                if !font_size_pt.is_finite()
                    || !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT).contains(font_size_pt)
                {
                    return Err("Document formula font size is outside the supported range".to_string());
                }
                validate_document_omml_payload(omml_base64)?;
                let native_docx = decode_document_native_docx(omml_docx_base64)?;
                let native_document_path = native_word_document_path(formula_id)?;
                atomic_write(&native_document_path, &native_docx, 0o600)?;

                let mut resolved_metadata = metadata.clone();
                let canonical_latex = validate_document_formula_metadata_match(
                    &resolved_metadata,
                    formula_id,
                    latex,
                    display_mode,
                    *numbered,
                )?;
                resolved_metadata.latex = canonical_latex.clone();
                resolved_metadata.font_size_pt = Some(*font_size_pt);

                let mut image_path = String::new();
                let mut vector_document_path = String::new();
                let mut fallback_image_path = String::new();
                let geometry = if input.output_kind == "image" {
                    let svg_value = svg_base64
                        .as_deref()
                        .ok_or_else(|| "Image document formula is missing SVG data".to_string())?;
                    let svg = decode_svg(svg_value)?;
                    let png = decode_document_image_fallback_png(png_base64.as_deref())?;
                    let width = width.ok_or_else(|| "Image document formula width is missing".to_string())?;
                    let height = height.ok_or_else(|| "Image document formula height is missing".to_string())?;
                    let baseline = baseline.unwrap_or(height);
                    let geometry = calculate_document_image_geometry(
                        width,
                        height,
                        baseline,
                        *font_size_pt,
                    )?;
                    let svg_path = document_formula_file_path(&session_id, formula_id, "svg")?;
                    let png_path = document_formula_file_path(&session_id, formula_id, "png")?;
                    let vector_path = document_formula_file_path(&session_id, formula_id, "docx")?;
                    atomic_write(&svg_path, &svg, 0o600)?;
                    atomic_write(&png_path, &png, 0o600)?;
                    let package = build_word_svg_docx(
                        &svg,
                        &png,
                        geometry.width,
                        geometry.height,
                    )?;
                    atomic_write(&vector_path, &package, 0o600)?;
                    image_path = svg_path.to_string_lossy().to_string();
                    fallback_image_path = png_path.to_string_lossy().to_string();
                    vector_document_path = vector_path.to_string_lossy().to_string();
                    resolved_metadata.render_width_px = Some(width);
                    resolved_metadata.render_height_px = Some(height);
                    resolved_metadata.reference_width_pt = Some(geometry.reference_width_pt);
                    resolved_metadata.reference_height_pt = Some(geometry.reference_height_pt);
                    resolved_metadata.reference_baseline_pt = Some(geometry.reference_baseline_pt);
                    geometry
                } else {
                    resolved_metadata.reference_width_pt = None;
                    resolved_metadata.reference_height_pt = None;
                    resolved_metadata.reference_baseline_pt = None;
                    WordGeometry {
                        width: *font_size_pt,
                        height: (*font_size_pt * 1.8).max(18.0),
                        baseline: 0,
                        font_size_pt: *font_size_pt,
                        reference_width_pt: WORD_REFERENCE_FONT_SIZE_PT,
                        reference_height_pt: WORD_REFERENCE_FONT_SIZE_PT,
                        reference_baseline_pt: 0.0,
                    }
                };
                let encoded_metadata = encode_metadata(&resolved_metadata)?;
                metadata_to_cache.push(resolved_metadata);

                entries.push((format!("{prefix}kind"), "formula".to_string()));
                entries.push((format!("{prefix}formulaId"), formula_id.clone()));
                entries.push((
                    format!("{prefix}latexBase64"),
                    URL_SAFE_NO_PAD.encode(canonical_latex.as_bytes()),
                ));
                entries.push((format!("{prefix}displayMode"), display_mode.clone()));
                entries.push((
                    format!("{prefix}numbered"),
                    if *numbered { "1" } else { "0" }.to_string(),
                ));
                entries.push((
                    format!("{prefix}fontSizePt"),
                    format!("{:.6}", font_size_pt),
                ));
                entries.push((format!("{prefix}metadata"), encoded_metadata));
                entries.push((format!("{prefix}ommlBase64"), omml_base64.clone()));
                entries.push((
                    format!("{prefix}nativeDocumentPath"),
                    native_document_path.to_string_lossy().to_string(),
                ));
                entries.push((format!("{prefix}imagePath"), image_path));
                entries.push((
                    format!("{prefix}vectorDocumentPath"),
                    vector_document_path,
                ));
                entries.push((
                    format!("{prefix}fallbackImagePath"),
                    fallback_image_path,
                ));
                entries.push((
                    format!("{prefix}widthPoints"),
                    format!("{:.6}", geometry.width),
                ));
                entries.push((
                    format!("{prefix}heightPoints"),
                    format!("{:.6}", geometry.height),
                ));
                entries.push((
                    format!("{prefix}baseline"),
                    geometry.baseline.to_string(),
                ));
                entries.push((
                    format!("{prefix}referenceWidthPt"),
                    format!("{:.6}", geometry.reference_width_pt),
                ));
                entries.push((
                    format!("{prefix}referenceHeightPt"),
                    format!("{:.6}", geometry.reference_height_pt),
                ));
                entries.push((
                    format!("{prefix}referenceBaselinePt"),
                    format!("{:.6}", geometry.reference_baseline_pt),
                ));
                append_document_paragraph_entries(&mut entries, &prefix, paragraph.as_ref());
                if paragraph.as_ref().is_some_and(|value| value.end) {
                    active_paragraph_id = None;
                }
            }
        }
    }
    if active_paragraph_id.is_some() {
        return Err("Document paragraph transfer ended before its paragraph boundary".to_string());
    }
    if formula_count == 0 && text_bytes == 0 {
        return Err("Document import contains no visible content".to_string());
    }

    let manifest = dynamic_dispatch_text(&entries)?;
    if manifest.len() > MAX_DOCUMENT_IMPORT_MANIFEST_BYTES {
        return Err(
            "Document import expands beyond the 16 MB Word transfer limit; split the source into smaller imports"
                .to_string(),
        );
    }
    let manifest_path = document_import_manifest_path(&session_id)?;
    atomic_write(&manifest_path, manifest.as_bytes(), 0o600)?;
    let progress_path = session_directory(OfficeHost::Word, &session_id)?
        .join(DOCUMENT_IMPORT_PROGRESS_FILE);
    atomic_write(
        &progress_path,
        format!("current=0\ntotal={}\nstage=preparing\n", input.items.len()).as_bytes(),
        0o600,
    )?;
    let dispatch = dispatch_text(&[
        ("protocolVersion", OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId", session_id.clone()),
        ("action", "documentCommit".to_string()),
        ("host", "word".to_string()),
        (
            "sourceDocumentId",
            public_request.source_document_id.clone(),
        ),
        ("bookmarkName", public_request.bookmark_name.clone()),
        (
            "documentImportPath",
            manifest_path.to_string_lossy().to_string(),
        ),
    ])?;
    atomic_write(
        &dispatch_path(OfficeHost::Word, &session_id)?,
        dispatch.as_bytes(),
        0o600,
    )?;
    with_dispatch_pointer(OfficeHost::Word, &session_id, || {
        run_vba_callback(OfficeHost::Word)
    })?;

    for metadata in metadata_to_cache {
        let formula_id = metadata.formula_id.clone();
        let _ = state.formula_cache.put(&formula_id, metadata);
    }
    let _ = cleanup_session_files(OfficeHost::Word, &session_id);
    Ok(())
}

fn cancel_document_import_blocking(session_id: String) -> Result<(), String> {
    validate_uuid(&session_id, "Session id")?;
    let request = read_request(&session_id)?;
    let public_request = document_import_request_data(&request)?;
    let dispatch = dispatch_text(&[
        ("protocolVersion", OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId", session_id.clone()),
        ("action", "documentCancel".to_string()),
        ("host", "word".to_string()),
        (
            "sourceDocumentId",
            public_request.source_document_id,
        ),
        ("bookmarkName", public_request.bookmark_name),
    ])?;
    atomic_write(
        &dispatch_path(OfficeHost::Word, &session_id)?,
        dispatch.as_bytes(),
        0o600,
    )?;
    with_dispatch_pointer(OfficeHost::Word, &session_id, || {
        run_vba_callback(OfficeHost::Word)
    })?;
    let _ = cleanup_session_files(OfficeHost::Word, &session_id);
    Ok(())
}

fn complete_session(
    state: &OfficeCompanionState,
    session_id: &str,
) -> Result<OfficeFormulaSession, String> {
    state
        .session_store
        .patch(
            session_id,
            json!({ "status": "completed", "error": null }),
        )
        .map_err(|error| error.to_string())
}

fn fail_session(state: &OfficeCompanionState, session_id: &str, error: &str) {
    let _ = state.session_store.patch(
        session_id,
        json!({ "status": "failed", "error": error }),
    );
}

fn commit_session_blocking(
    state: OfficeCompanionState,
    session_id: String,
) -> Result<OfficeFormulaSession, String> {
    validate_uuid(&session_id, "Session id")?;
    let session = state
        .session_store
        .get(&session_id)
        .map_err(|error| error.to_string())?;
    if session.status == OfficeSessionStatus::Completed {
        let _ = cleanup_session_files(session.host, &session_id);
        return Ok(session);
    }
    if session.status != OfficeSessionStatus::Committing {
        return Err("Offline Office Session is not ready to commit".to_string());
    }
    let request = read_request(&session_id)?;
    let source_is_native_word_equation = session.host == OfficeHost::Word
        && request
            .source_object_id
            .as_deref()
            .is_some_and(|value| value.starts_with("VT_F_"));
    let word_format_conversion_requested = session.host == OfficeHost::Word
        && source_is_native_word_equation != request.native_equation;
    if session.mode == OfficeSessionMode::Edit
        && !session.dirty
        && !word_format_conversion_requested
    {
        let completed = complete_session(&state, &session_id)?;
        let _ = cleanup_session_files(session.host, &session_id);
        return Ok(completed);
    }
    let mut metadata = metadata_from_session(&session);
    let result = match session.host {
        OfficeHost::Word => {
            let geometry = calculate_word_geometry(&request, &session)?;
            metadata.latex = canonical_document_formula_latex(&metadata)?;
            metadata.font_size_pt = Some(geometry.font_size_pt);
            metadata.reference_width_pt = Some(geometry.reference_width_pt);
            metadata.reference_height_pt = Some(geometry.reference_height_pt);
            metadata.reference_baseline_pt = Some(geometry.reference_baseline_pt);
            let encoded = encode_metadata(&metadata)?;
            commit_word(
                &request,
                &session,
                &encoded,
                &metadata.latex,
                geometry,
            )
        }
        OfficeHost::Powerpoint => {
            let powerpoint = request
                .power_point
                .as_ref()
                .ok_or_else(|| "PowerPoint request geometry is missing".to_string())?;
            let geometry = calculate_powerpoint_geometry(powerpoint, &session)?;
            metadata.font_size_pt = Some(geometry.font_size_pt);
            metadata.reference_width_pt = Some(geometry.reference_width_pt);
            metadata.reference_height_pt = Some(geometry.reference_height_pt);
            metadata.reference_baseline_pt = None;
            let encoded = encode_metadata(&metadata)?;
            commit_powerpoint(&request, &session, &encoded, geometry)
        }
    };
    if let Err(error) = result {
        fail_session(&state, &session_id, &error);
        return Err(error);
    }
    if let Err(error) = state.formula_cache.put(&session.formula_id, metadata) {
        let message = format!("Formula metadata could not be saved: {error}");
        fail_session(&state, &session_id, &message);
        return Err(message);
    }
    let completed = complete_session(&state, &session_id)?;
    let _ = cleanup_session_files(session.host, &session_id);
    Ok(completed)
}

fn cancel_session_blocking(
    state: OfficeCompanionState,
    session_id: String,
) -> Result<OfficeFormulaSession, String> {
    validate_uuid(&session_id, "Session id")?;
    let request = read_request(&session_id)?;
    let host = host_from_request_name(&request.host)?;
    if let Err(error) = cancel_host(&request) {
        fail_session(&state, &session_id, &error);
        return Err(error);
    }
    let cancelled = state
        .session_store
        .patch(
            &session_id,
            json!({
                "status": "cancelled",
                "explicitCancel": true,
                "error": null
            }),
        )
        .map_err(|error| error.to_string())?;
    let _ = cleanup_session_files(host, &session_id);
    Ok(cancelled)
}

#[tauri::command]
pub fn get_macos_offline_document_import_request(
    session_id: String,
) -> Result<MacOfflineDocumentImportPublicRequest, String> {
    validate_uuid(&session_id, "Session id")?;
    let request = read_request(&session_id)?;
    document_import_request_data(&request)
}

#[tauri::command]
pub fn focus_macos_offline_document_import_target(
    window: WebviewWindow,
) -> Result<(), String> {
    if !window.label().starts_with("office-native-document-") {
        return Err("Only the VisualTeX document importer can focus Word".to_string());
    }
    set_word_document_import_preparing_status()?;
    window
        .hide()
        .map_err(|error| format!("Unable to hide the VisualTeX document importer: {error}"))
}

#[tauri::command]
pub fn restore_macos_offline_document_import_window(
    window: WebviewWindow,
) -> Result<(), String> {
    if !window.label().starts_with("office-native-document-") {
        return Err("Only the VisualTeX document importer can restore itself".to_string());
    }
    clear_word_document_import_status();
    let app = window.app_handle().clone();
    crate::office::background::activate_foreground_app(&app)?;
    window.show().map_err(|error| error.to_string())?;
    window.unminimize().map_err(|error| error.to_string())?;
    window.set_focus().map_err(|error| error.to_string())
}

#[tauri::command]
pub fn get_macos_offline_document_import_progress(
    session_id: String,
) -> Result<MacOfflineDocumentImportProgress, String> {
    validate_uuid(&session_id, "Session id")?;
    let path = session_directory(OfficeHost::Word, &session_id)?
        .join(DOCUMENT_IMPORT_PROGRESS_FILE);
    let source = match fs::read_to_string(path) {
        Ok(value) => value,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => {
            return Ok(MacOfflineDocumentImportProgress {
                current: 0,
                total: 0,
                stage: "preparing".to_string(),
            });
        }
        Err(error) => return Err(format!("Unable to read document import progress: {error}")),
    };
    let mut current = None;
    let mut total = None;
    let mut stage = None;
    for line in source.lines() {
        let Some((key, value)) = line.split_once('=') else {
            continue;
        };
        match key {
            "current" => current = value.parse::<usize>().ok(),
            "total" => total = value.parse::<usize>().ok(),
            "stage" => stage = Some(value.to_string()),
            _ => {}
        }
    }
    let current = current.ok_or_else(|| "Document import progress is missing current".to_string())?;
    let total = total.ok_or_else(|| "Document import progress is missing total".to_string())?;
    let stage = stage.ok_or_else(|| "Document import progress is missing stage".to_string())?;
    if total > 2048 || current > total || stage.len() > 32 || stage.chars().any(char::is_control) {
        return Err("Document import progress is invalid".to_string());
    }
    Ok(MacOfflineDocumentImportProgress {
        current,
        total,
        stage,
    })
}

#[tauri::command]
pub async fn commit_macos_offline_document_import(
    session_id: String,
    input: MacOfflineDocumentImportCommitInput,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<(), String> {
    let state = state.inner().clone();
    tokio::task::spawn_blocking(move || {
        commit_document_import_blocking(state, session_id, input)
    })
    .await
    .map_err(|error| format!("Offline document import task failed: {error}"))?
}

#[tauri::command]
pub async fn cancel_macos_offline_document_import(
    session_id: String,
) -> Result<(), String> {
    tokio::task::spawn_blocking(move || cancel_document_import_blocking(session_id))
        .await
        .map_err(|error| format!("Offline document import cancel task failed: {error}"))?
}

#[tauri::command]
pub fn get_macos_offline_office_session(
    session_id: String,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficeFormulaSession, String> {
    validate_uuid(&session_id, "Session id")?;
    state
        .session_store
        .get(&session_id)
        .map_err(|error| error.to_string())
}

#[tauri::command]
pub fn update_macos_offline_office_session(
    session_id: String,
    patch: Value,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficeFormulaSession, String> {
    validate_uuid(&session_id, "Session id")?;
    state
        .session_store
        .patch(&session_id, patch)
        .map_err(|error| error.to_string())
}

#[tauri::command]
pub fn delete_macos_offline_office_session(
    session_id: String,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<(), String> {
    validate_uuid(&session_id, "Session id")?;
    let session = state
        .session_store
        .get(&session_id)
        .map_err(|error| error.to_string())?;
    state
        .session_store
        .delete(&session_id)
        .map_err(|error| error.to_string())?;
    cleanup_session_files(session.host, &session_id)
}

#[tauri::command]
pub async fn commit_macos_offline_office_session(
    session_id: String,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficeFormulaSession, String> {
    let state = state.inner().clone();
    tokio::task::spawn_blocking(move || commit_session_blocking(state, session_id))
        .await
        .map_err(|error| format!("Offline Office commit task failed: {error}"))?
}

#[tauri::command]
pub async fn cancel_macos_offline_office_session(
    session_id: String,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficeFormulaSession, String> {
    let state = state.inner().clone();
    tokio::task::spawn_blocking(move || cancel_session_blocking(state, session_id))
        .await
        .map_err(|error| format!("Offline Office cancel task failed: {error}"))?
}

#[cfg(target_os = "macos")]
pub(crate) fn refresh_health_signal(host: &str) -> bool {
    let (process_name, script) = match host {
        "word" => (
            "Microsoft Word",
            r#"tell application "Microsoft Word" to run VB macro macro name "AutoExec""#,
        ),
        "powerpoint" => (
            "Microsoft PowerPoint",
            r#"tell application "Microsoft PowerPoint" to run VB macro macro name "Auto_Open" list of parameters {}"#,
        ),
        _ => return false,
    };
    let running = Command::new("/usr/bin/pgrep")
        .args(["-x", process_name])
        .output()
        .is_ok_and(|output| output.status.success());
    if !running {
        return false;
    }
    Command::new("/usr/bin/osascript")
        .arg("-e")
        .arg(script)
        .output()
        .is_ok_and(|output| output.status.success())
}

#[cfg(not(target_os = "macos"))]
pub(crate) fn refresh_health_signal(_host: &str) -> bool {
    false
}

pub(crate) fn health_path(host: &str) -> Result<PathBuf, String> {
    Ok(runtime_root(host_from_request_name(host)?)?
        .join("OfficePluginStatus")
        .join(format!("{host}.json")))
}

fn read_health(host: &str) -> Result<MacOfflinePluginHealth, String> {
    let path = health_path(host)?;
    let fallback = || MacOfflinePluginHealth {
        loaded: false,
        plugin_version: None,
        source_revision: None,
        host: host.to_string(),
        timestamp: None,
        status_path: path.display().to_string(),
    };
    let bytes = match fs::read(&path) {
        Ok(bytes) => bytes,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(fallback()),
        Err(error) => return Err(format!("Unable to read {} health: {error}", host)),
    };
    let value: Value = serde_json::from_slice(&bytes)
        .map_err(|error| format!("{host} health file contains invalid JSON: {error}"))?;
    let plugin_version = value
        .get("pluginVersion")
        .and_then(Value::as_str)
        .map(str::to_string);
    let source_revision = value
        .get("sourceRevision")
        .and_then(Value::as_str)
        .map(str::to_string);
    let reported_host = value
        .get("host")
        .and_then(Value::as_str)
        .unwrap_or(host)
        .to_string();
    let timestamp = value
        .get("timestamp")
        .and_then(Value::as_str)
        .filter(|value| {
            !value.is_empty() && value.len() <= 64 && !value.chars().any(char::is_control)
        })
        .map(str::to_string);
    let loaded = value.get("loaded").and_then(Value::as_bool).unwrap_or(false)
        && plugin_version.as_deref() == Some(env!("CARGO_PKG_VERSION"))
        && reported_host == host
        && timestamp.is_some();
    Ok(MacOfflinePluginHealth {
        loaded,
        plugin_version,
        source_revision,
        host: reported_host,
        timestamp,
        status_path: path.display().to_string(),
    })
}

#[tauri::command]
pub fn get_macos_offline_plugin_health() -> Result<Vec<MacOfflinePluginHealth>, String> {
    let _ = refresh_health_signal("word");
    let _ = refresh_health_signal("powerpoint");
    Ok(vec![read_health("word")?, read_health("powerpoint")?])
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn office_formula_editors_have_one_stable_window_label_per_host() {
        assert_eq!(editor_window_label(OfficeHost::Word), "office-native-word-editor");
        assert_eq!(
            editor_window_label(OfficeHost::Powerpoint),
            "office-native-powerpoint-editor"
        );
        assert_eq!(
            editor_window_host("office-native-word-editor"),
            Some(OfficeHost::Word)
        );
        assert_eq!(
            editor_window_host("office-native-powerpoint-editor"),
            Some(OfficeHost::Powerpoint)
        );
        assert_eq!(editor_window_host("office-native-document-session"), None);
    }

    #[test]
    fn editor_ready_marker_exposes_machine_readable_frontend_stages() {
        let marker = MacOfflineOfficeEditorReadyMarker {
            schema: "visualtex-office-editor-ready-v1",
            session_id: "12345678-1234-4234-9234-123456789abc".to_string(),
            host: OfficeHost::Word,
            generation: 4,
            epoch_ms: 900,
            url_received_epoch_ms: 100,
            frontend_epoch_ms: 850,
            hydrate_ms: 120.5,
            editor_mounted_ms: 150.25,
            content_ready_ms: 166.75,
            show_focus_ms: 171.0,
        };
        let value = serde_json::to_value(marker).expect("ready marker should serialize");
        assert_eq!(value["sessionId"], "12345678-1234-4234-9234-123456789abc");
        assert_eq!(value["host"], "word");
        assert_eq!(value["hydrateMs"], 120.5);
        assert_eq!(value["contentReadyMs"], 166.75);
        assert_eq!(value["showFocusMs"], 171.0);
    }

    #[test]
    fn runtime_roots_use_each_office_hosts_application_scripts_directory() {
        let word = runtime_root(OfficeHost::Word).expect("Word runtime root should resolve");
        let powerpoint =
            runtime_root(OfficeHost::Powerpoint).expect("PowerPoint runtime root should resolve");
        assert!(word.ends_with(WORD_RUNTIME_SUFFIX));
        assert!(powerpoint.ends_with(POWERPOINT_RUNTIME_SUFFIX));
        assert_ne!(word, powerpoint);
        assert!(!word.to_string_lossy().contains("UBF8T346G9.Office"));
        assert!(!powerpoint.to_string_lossy().contains("UBF8T346G9.Office"));
        assert!(!word.starts_with("/private/tmp"));
        assert!(!powerpoint.starts_with("/private/tmp"));
    }

    #[test]
    fn word_image_geometry_scales_from_the_14_point_reference() {
        let small = scale_word_reference_geometry(140.0, 28.0, -4.0, 10.5)
            .expect("10.5 pt geometry should scale");
        let large = scale_word_reference_geometry(140.0, 28.0, -4.0, 18.0)
            .expect("18 pt geometry should scale");

        assert!((small.width - 105.0).abs() < 0.001);
        assert!((small.height - 21.0).abs() < 0.001);
        assert_eq!(small.baseline, -3);
        assert!((large.width - 180.0).abs() < 0.001);
        assert!((large.height - 36.0).abs() < 0.001);
        assert_eq!(large.baseline, -5);
        assert!((large.width / small.width - 18.0 / 10.5).abs() < 0.001);
    }

    #[test]
    fn word_inline_baseline_rounds_only_after_font_size_scaling() {
        // Real MathJax geometry from the reported L^2 / L_z regression at the
        // canonical 14 pt export size. Their fractional descents are about
        // 0.825 pt and 3.25512 pt respectively.
        let superscript = calculate_word_svg_geometry(
            22.86186666666666,
            17.56613333333333,
            Some(16.56613333333333),
            11.0,
        )
        .expect("L^2 geometry should resolve");
        let subscript = calculate_word_svg_geometry(
            22.39893333333333,
            17.69493333333333,
            Some(13.749333333333333),
            11.0,
        )
        .expect("L_z geometry should resolve");

        assert!((superscript.reference_baseline_pt + 0.825).abs() < 0.001);
        assert!((subscript.reference_baseline_pt + 3.25512).abs() < 0.001);
        assert_eq!(superscript.baseline, -1);
        assert_eq!(subscript.baseline, -3);
        assert_eq!(subscript.baseline - superscript.baseline, -2);
    }

    #[test]
    fn document_image_requires_a_real_png_compatibility_preview() {
        assert!(decode_document_image_fallback_png(None).is_err());

        let transparent_placeholder = BASE64_STANDARD.encode(
            BASE64_STANDARD
                .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAEAQH/8l0Z8QAAAABJRU5ErkJggg==")
                .expect("transparent PNG fixture should decode"),
        );
        assert!(decode_document_image_fallback_png(Some(&transparent_placeholder)).is_err());

        let supplied = BASE64_STANDARD.encode(
            BASE64_STANDARD
                .decode("iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFElEQVR4nGNkYGD4z8DAwMDEAAUADigBA0dwHFEAAAAASUVORK5CYII=")
                .expect("PNG fixture should decode"),
        );
        let decoded = decode_document_image_fallback_png(Some(&supplied))
            .expect("supplied PNG compatibility preview should decode");
        assert!(decoded.starts_with(b"\x89PNG\r\n\x1a\n"));
        assert_eq!(u32::from_be_bytes(decoded[16..20].try_into().unwrap()), 2);
        assert_eq!(u32::from_be_bytes(decoded[20..24].try_into().unwrap()), 2);
    }

    #[cfg(target_os = "macos")]
    #[test]
    fn word_svg_staging_docx_is_a_valid_ooxml_zip() {
        let svg = br#"<svg xmlns="http://www.w3.org/2000/svg" width="16" height="8"><path d="M0 0h16v8H0z"/></svg>"#;
        let png = BASE64_STANDARD
            .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")
            .expect("PNG fixture should decode");
        let package = build_word_svg_docx(svg, &png, 16.0, 8.0)
            .expect("Word SVG package should build");
        assert!(package.starts_with(b"PK\x03\x04"));
        assert!(package.windows(b"word/media/formula.svg".len()).any(|value| {
            value == b"word/media/formula.svg"
        }));
        assert!(package.windows(b"drawing/2016/SVG/main".len()).any(|value| {
            value == b"drawing/2016/SVG/main"
        }));

        let directory = tempfile::tempdir().expect("temporary directory should exist");
        let path = directory.path().join("formula-svg.docx");
        fs::write(&path, &package).expect("Word SVG package should be writable");
        let output = Command::new("/usr/bin/unzip")
            .args(["-tqq"])
            .arg(&path)
            .output()
            .expect("macOS unzip should validate the package");
        assert!(
            output.status.success(),
            "generated Word SVG DOCX is invalid: {}",
            String::from_utf8_lossy(&output.stderr)
        );
    }

    #[cfg(target_os = "macos")]
    #[test]
    #[ignore = "writes a real Word SVG probe DOCX to VISUALTEX_WORD_SVG_PROBE_PATH"]
    fn write_word_svg_probe_docx() {
        let path = std::env::var("VISUALTEX_WORD_SVG_PROBE_PATH")
            .expect("set VISUALTEX_WORD_SVG_PROBE_PATH to an absolute .docx path");
        let svg = br#"<svg xmlns="http://www.w3.org/2000/svg" width="160" height="80" viewBox="0 0 160 80"><rect width="160" height="80" fill="white"/><path d="M20 55L55 20L90 55L125 20" fill="none" stroke="black" stroke-width="8"/></svg>"#;
        let png = BASE64_STANDARD
            .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")
            .expect("PNG fixture should decode");
        let package = build_word_svg_docx(svg, &png, 160.0, 80.0)
            .expect("Word SVG package should build");
        fs::write(path, package).expect("Word SVG probe should be writable");
    }

    #[test]
    fn native_word_documents_are_formula_scoped_and_outlive_sessions() {
        let formula_id = "12345678-1234-4234-9234-123456789abc";
        let path = native_word_document_path(formula_id)
            .expect("native Word document path should resolve");
        let runtime = runtime_root(OfficeHost::Word)
            .expect("Word runtime root should resolve");

        assert!(path.starts_with(&runtime));
        assert_eq!(
            path.parent().and_then(|value| value.file_name()).and_then(|value| value.to_str()),
            Some("NativeDocuments")
        );
        assert_eq!(
            path.file_name().and_then(|value| value.to_str()),
            Some("12345678-1234-4234-9234-123456789abc.docx")
        );
        assert!(!path.to_string_lossy().contains("OfficeSessions"));
    }

    #[test]
    fn completed_session_cleanup_removes_only_known_ephemeral_files() {
        let directory = std::env::temp_dir().join(format!(
            "visualtex-offline-cleanup-test-{}",
            Uuid::new_v4()
        ));
        fs::create_dir_all(&directory).expect("test Session directory should be created");
        for name in [REQUEST_FILE, DISPATCH_FILE, RESULT_PNG_FILE, RESULT_SVG_FILE] {
            fs::write(directory.join(name), b"temporary").expect("temporary file should exist");
        }
        fs::write(directory.join("keep.txt"), b"keep").expect("unknown file should exist");

        cleanup_session_files_at(&directory).expect("known files should be cleaned");
        for name in [REQUEST_FILE, DISPATCH_FILE, RESULT_PNG_FILE, RESULT_SVG_FILE] {
            assert!(!directory.join(name).exists());
        }
        assert!(directory.join("keep.txt").is_file());
        assert!(directory.is_dir());

        fs::remove_file(directory.join("keep.txt")).unwrap();
        cleanup_session_files_at(&directory).expect("empty Session directory should be removed");
        assert!(!directory.exists());
    }

    #[test]
    fn powerpoint_svg_decoder_accepts_internal_vector_references_only() {
        let safe = BASE64_STANDARD.encode(
            br##"<svg xmlns="http://www.w3.org/2000/svg"><defs><path id="g" d="M0 0h1v1z"/></defs><use href="#g"/></svg>"##,
        );
        let decoded = decode_svg(&safe).expect("generated SVG should be accepted");
        assert!(std::str::from_utf8(&decoded).unwrap().contains("<use"));

        let external = BASE64_STANDARD.encode(
            br#"<svg xmlns="http://www.w3.org/2000/svg"><image href="https://example.com/a.png"/></svg>"#,
        );
        assert!(decode_svg(&external).is_err());
        let scripted = BASE64_STANDARD.encode(
            br#"<svg xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script></svg>"#,
        );
        assert!(decode_svg(&scripted).is_err());
    }

    #[cfg(target_os = "macos")]
    #[test]
    #[ignore = "requires target/source PowerPoint Sessions and explicit environment variables"]
    fn live_powerpoint_svg_commit_uses_the_real_ppam_transaction() {
        let session_id = std::env::var("VISUALTEX_LIVE_PPT_SESSION")
            .expect("set VISUALTEX_LIVE_PPT_SESSION to an open PowerPoint create Session");
        let request = read_request(&session_id).expect("PowerPoint request should be readable");
        assert_eq!(request.host, "powerpoint");
        assert!(request.mode == "create" || request.mode == "edit");
        let formula_id = request
            .formula_id
            .clone()
            .expect("PowerPoint request should contain a formula id");
        let source_session_id = std::env::var("VISUALTEX_LIVE_PPT_EXPORT_SESSION")
            .expect("set VISUALTEX_LIVE_PPT_EXPORT_SESSION to a completed VisualTeX formula Session");
        validate_uuid(&source_session_id, "Source Session id").unwrap();
        let home = std::env::var("HOME").expect("HOME should be set on macOS");
        let source_session_path = PathBuf::from(home)
            .join("Library/Application Support/com.visualtex.studio/office/sessions")
            .join(&source_session_id)
            .join("session.json");
        let mut session: OfficeFormulaSession = serde_json::from_slice(
            &fs::read(&source_session_path).expect("source VisualTeX Session should be readable"),
        )
        .expect("source VisualTeX Session should decode");
        assert_eq!(session.host, OfficeHost::Powerpoint);
        let source_export = session
            .export_result
            .as_ref()
            .expect("source VisualTeX Session must contain a real formula export");
        decode_svg(&source_export.svg_base64)
            .expect("source VisualTeX Session must contain a validated SVG export");
        session.id = session_id.clone();
        session.mode = if request.mode == "edit" {
            OfficeSessionMode::Edit
        } else {
            OfficeSessionMode::Create
        };
        session.formula_id = formula_id;
        session.source_document_id = request.source_document_id.clone();
        session.source_object_id = request.source_object_id.clone();
        session.original_metadata = request
            .encoded_metadata
            .as_deref()
            .map(decode_metadata)
            .transpose()
            .expect("target PowerPoint metadata should decode");
        session.dirty = true;
        session.status = OfficeSessionStatus::Committing;
        session.explicit_cancel = false;
        session.error = None;

        let root = std::env::temp_dir().join(format!(
            "visualtex-live-powerpoint-svg-{}",
            Uuid::new_v4()
        ));
        let paths = crate::office::state::OfficePaths {
            certificate: root.join("localhost-cert.pem"),
            private_key: root.join("localhost-key.pem"),
            certificate_metadata: root.join("certificate.json"),
            install: root.join("install.json"),
            sessions: root.join("sessions"),
            recovery: root.join("recovery"),
            formula_cache: root.join("formulas"),
            root: root.clone(),
        };
        let session_store = crate::office::sessions::SessionStore::new(&paths)
            .expect("live Session store should initialize");
        let formula_cache = crate::office::formula_cache::FormulaMetadataCache::new(&paths)
            .expect("live formula cache should initialize");
        let state = OfficeCompanionState::new(
            None,
            crate::OcrState::default(),
            paths,
            "a".repeat(64),
            session_store,
            formula_cache,
            true,
        );
        let metadata = encode_metadata(&metadata_from_session(&session))
            .expect("live metadata should encode");
        let powerpoint = request
            .power_point
            .as_ref()
            .expect("live PowerPoint request should contain geometry");
        let geometry = calculate_powerpoint_geometry(powerpoint, &session)
            .expect("live PowerPoint geometry should resolve");
        commit_powerpoint(&request, &session, &metadata, geometry)
            .expect("real PowerPoint PPAM SVG transaction should succeed");
        let svg_path = result_svg_path(OfficeHost::Powerpoint, &session_id)
            .expect("SVG result path should resolve");
        assert_eq!(
            fs::read_to_string(svg_path).unwrap(),
            session.export_result.as_ref().unwrap().svg
        );
        fs::remove_dir_all(root).ok();
    }

    #[test]
    fn office_url_accepts_only_the_exact_canonical_form() {
        let id = "12345678-1234-4234-9234-123456789abc";
        assert_eq!(
            parse_office_url(&format!("visualtex://office/open?session={id}")),
            Ok(id.to_string())
        );
        assert!(parse_office_url(&format!("https://office/open?session={id}")).is_err());
        assert!(parse_office_url(&format!("visualtex://office/open?session={id}&x=1")).is_err());
        assert!(parse_office_url("visualtex://office/open?session=not-a-uuid").is_err());
    }

    #[test]
    fn offline_request_json_accepts_utf8_office_identities() {
        let session_id = "32345678-1234-4234-9234-123456789abc".to_string();
        let request = MacOfflineSessionRequest {
            protocol_version: OFFLINE_PROTOCOL_VERSION,
            session_id: session_id.clone(),
            host: "word".to_string(),
            mode: "create".to_string(),
            operation: None,
            formula_id: Some("12345678-1234-4234-9234-123456789abc".to_string()),
            display_mode: "inline".to_string(),
            numbered: false,
            native_equation: false,
            source_document_id: Some("/Users/测试/公式😀.docx".to_string()),
            source_object_id: Some("书签-公式".to_string()),
            encoded_metadata: None,
            pending_marker: Some(
                "visualtex:pending:v1:32345678-1234-4234-9234-123456789abc:12345678-1234-4234-9234-123456789abc"
                    .to_string(),
            ),
            font_size_pt: Some(10.5),
            reference_width_pt: Some(60.0),
            reference_height_pt: Some(15.0),
            power_point: None,
            document_import: None,
        };
        let json = serde_json::to_vec(&request).expect("UTF-8 request should encode");
        let decoded: MacOfflineSessionRequest =
            serde_json::from_slice(&json).expect("UTF-8 request should decode");
        validate_request(&decoded, &session_id).expect("UTF-8 request should validate");
        assert_eq!(
            decoded.source_document_id.as_deref(),
            Some("/Users/测试/公式😀.docx")
        );
    }

    #[test]
    fn metadata_codec_round_trips_the_shared_schema() {
        let metadata = VisualTeXFormulaMetadata {
            schema: "visualtex-formula".to_string(),
            schema_version: 1,
            formula_id: "12345678-1234-4234-9234-123456789abc".to_string(),
            title: "Formula".to_string(),
            latex: "x^2".to_string(),
            lines: vec![crate::office::sessions::MetadataLine {
                id: "22345678-1234-4234-9234-123456789abc".to_string(),
                latex: "x^2".to_string(),
            }],
            code_format: "latex".to_string(),
            display_mode: "inline".to_string(),
            numbered: false,
            render_width_px: Some(50.0),
            render_height_px: Some(20.0),
            font_size_pt: Some(10.5),
            reference_width_pt: Some(37.5),
            reference_height_pt: Some(15.0),
            reference_baseline_pt: Some(-3.0),
            created_with_version: "1.1.0".to_string(),
            updated_with_version: "1.1.0".to_string(),
            created_at: "unix-ms:1".to_string(),
            updated_at: "unix-ms:1".to_string(),
        };
        let encoded = encode_metadata(&metadata).expect("metadata should encode");
        let decoded = decode_metadata(&encoded).expect("metadata should decode");
        assert_eq!(decoded.formula_id, metadata.formula_id);
        assert_eq!(decoded.lines[0].latex, "x^2");
        assert_eq!(decoded.font_size_pt, Some(10.5));
        assert_eq!(decoded.reference_width_pt, Some(37.5));
        assert_eq!(decoded.reference_height_pt, Some(15.0));
        assert_eq!(decoded.reference_baseline_pt, Some(-3.0));
    }

    fn document_formula_metadata(
        code_format: &str,
        lines: &[&str],
        latex: &str,
        display_mode: &str,
        numbered: bool,
    ) -> VisualTeXFormulaMetadata {
        VisualTeXFormulaMetadata {
            schema: "visualtex-formula".to_string(),
            schema_version: 1,
            formula_id: "12345678-1234-4234-9234-123456789abc".to_string(),
            title: "Imported formula".to_string(),
            latex: latex.to_string(),
            lines: lines
                .iter()
                .map(|latex| crate::office::sessions::MetadataLine {
                    id: Uuid::new_v4().to_string(),
                    latex: (*latex).to_string(),
                })
                .collect(),
            code_format: code_format.to_string(),
            display_mode: display_mode.to_string(),
            numbered,
            render_width_px: None,
            render_height_px: None,
            font_size_pt: Some(14.0),
            reference_width_pt: None,
            reference_height_pt: None,
            reference_baseline_pt: None,
            created_with_version: "1.2.3".to_string(),
            updated_with_version: "1.2.3".to_string(),
            created_at: "unix-ms:1".to_string(),
            updated_at: "unix-ms:1".to_string(),
        }
    }

    #[test]
    fn document_formula_metadata_rebuilds_canonical_multiline_environments() {
        let cases = [
            (
                "align",
                vec!["a = b + c", "d = e"],
                r"\begin{align}
a &= b + c \\
d &= e
\end{align}",
            ),
            (
                "align-star",
                vec!["x = y", "y = z"],
                r"\begin{align*}
x &= y \\
y &= z
\end{align*}",
            ),
            (
                "aligned",
                vec!["p = q", "r = s"],
                r"\[
\begin{aligned}
p &= q \\
r &= s
\end{aligned}
\]",
            ),
            (
                "gather",
                vec!["a=b", "c=d"],
                r"\begin{gather}
a=b \\
c=d
\end{gather}",
            ),
            (
                "multline-star",
                vec!["a+b+c", "=d+e"],
                r"\begin{multline*}
a+b+c \\
=d+e
\end{multline*}",
            ),
            (
                "equation-split",
                vec!["a = b", "c = d"],
                r"\begin{equation}
\begin{split}
a &= b \\
c &= d
\end{split}
\end{equation}",
            ),
            (
                "equation-star-split",
                vec!["a = b", "c = d"],
                r"\begin{equation*}
\begin{split}
a &= b \\
c &= d
\end{split}
\end{equation*}",
            ),
        ];

        for (code_format, lines, expected) in cases {
            let metadata = document_formula_metadata(
                code_format,
                &lines,
                expected,
                "block",
                false,
            );
            assert_eq!(
                canonical_document_formula_latex(&metadata).unwrap(),
                expected,
                "{code_format} canonical source"
            );
            assert_eq!(
                validate_document_formula_metadata_match(
                    &metadata,
                    &metadata.formula_id,
                    expected,
                    "block",
                    false,
                )
                .unwrap(),
                expected,
                "{code_format} metadata match"
            );
        }
    }

    #[test]
    fn document_formula_metadata_preserves_internal_equation_newlines() {
        let body = r"u(x,y)=\sum_{n=1}^{+\infty}\sum_{m=1}^{+\infty}c_{nm}\sin\frac{n\pi}{a}x\sin\frac{m\pi}{b}y,\qquad
f(x,y)=\sum_{n=1}^{+\infty}\sum_{m=1}^{+\infty}d_{nm}\sin\frac{n\pi}{a}x\sin\frac{m\pi}{b}y.";
        let canonical = r"\begin{equation*}
u(x,y)=\sum_{n=1}^{+\infty}\sum_{m=1}^{+\infty}c_{nm}\sin\frac{n\pi}{a}x\sin\frac{m\pi}{b}y,\qquad
f(x,y)=\sum_{n=1}^{+\infty}\sum_{m=1}^{+\infty}d_{nm}\sin\frac{n\pi}{a}x\sin\frac{m\pi}{b}y.
\end{equation*}";
        let metadata = document_formula_metadata(
            "equation-star",
            &[body],
            canonical,
            "block",
            false,
        );

        assert_eq!(canonical_document_formula_latex(&metadata).unwrap(), canonical);
        assert_eq!(
            validate_document_formula_metadata_match(
                &metadata,
                &metadata.formula_id,
                canonical,
                "block",
                false,
            )
            .unwrap(),
            canonical,
        );

        let separate_rows = document_formula_metadata(
            "equation-star",
            &["a=b", "c=d"],
            r"\begin{equation*}
a=b
\end{equation*}

\begin{equation*}
c=d
\end{equation*}",
            "block",
            false,
        );
        assert_eq!(
            canonical_document_formula_latex(&separate_rows).unwrap(),
            separate_rows.latex,
            "separate logical rows must remain separate equation environments",
        );
    }

    #[test]
    fn document_formula_metadata_accepts_frontend_environment_spacing() {
        let frontend_source = r"\begin {equation}
E=mc^2
\end {equation}";
        let metadata = document_formula_metadata(
            "equation",
            &["E=mc^2"],
            frontend_source,
            "block",
            true,
        );

        assert_ne!(
            canonical_document_formula_latex(&metadata).unwrap(),
            frontend_source,
            "the Rust formatter intentionally normalizes environment spacing",
        );
        assert_eq!(
            validate_document_formula_metadata_match(
                &metadata,
                &metadata.formula_id,
                frontend_source,
                "block",
                true,
            )
            .unwrap(),
            frontend_source,
            "a valid frontend serialization must not be rejected because Rust formats it differently",
        );
    }

    #[test]
    fn document_formula_metadata_match_rejects_structural_drift() {
        let canonical = r"\begin{align}
a &= b \\
c &= d
\end{align}";
        let metadata = document_formula_metadata(
            "align",
            &["a = b", "c = d"],
            canonical,
            "block",
            true,
        );
        assert!(validate_document_formula_metadata_match(
            &metadata,
            &metadata.formula_id,
            canonical,
            "block",
            true,
        )
        .is_ok());
        assert!(validate_document_formula_metadata_match(
            &metadata,
            "22345678-1234-4234-9234-123456789abc",
            canonical,
            "block",
            true,
        )
        .is_err());
        assert!(validate_document_formula_metadata_match(
            &metadata,
            &metadata.formula_id,
            canonical,
            "inline",
            true,
        )
        .is_err());
        assert!(validate_document_formula_metadata_match(
            &metadata,
            &metadata.formula_id,
            canonical,
            "block",
            false,
        )
        .is_err());
        assert!(validate_document_formula_metadata_match(
            &metadata,
            &metadata.formula_id,
            r"\begin{align}
a &= b \\
c &= e
\end{align}",
            "block",
            true,
        )
        .is_err());

        let mut stale_metadata = metadata.clone();
        stale_metadata.latex = "a = b\nc = d".to_string();
        assert!(validate_document_formula_metadata_match(
            &stale_metadata,
            &stale_metadata.formula_id,
            canonical,
            "block",
            true,
        )
        .is_err());
    }

    #[test]
    fn dispatch_rejects_newlines_and_duplicate_keys() {
        assert!(dispatch_text(&[("sessionId", "a\nb".to_string())]).is_err());
        assert!(dispatch_text(&[
            ("sessionId", "a".to_string()),
            ("sessionId", "b".to_string())
        ])
        .is_err());
    }

    #[test]
    fn office_geometry_preserves_visual_point_size_and_powerpoint_center() {
        let word_geometry = calculate_word_svg_geometry(100.0, 20.0, Some(15.0), 14.0)
            .expect("Word image geometry should apply its visual calibration");
        assert!((word_geometry.width - 82.5).abs() < 0.001);
        assert!((word_geometry.height - 16.5).abs() < 0.001);
        assert_eq!(word_geometry.baseline, -4);
        assert!((word_geometry.reference_width_pt - 82.5).abs() < 0.001);
        assert!((word_geometry.reference_height_pt - 16.5).abs() < 0.001);

        let request = MacOfflinePowerPointRequest {
            presentation_identity: "Deck".to_string(),
            slide_index: 1,
            slide_id: 2,
            shape_name: "VisualTeX_12345678-1234-4234-9234-123456789abc".to_string(),
            left: 100.0,
            top: 200.0,
            width: 120.0,
            height: 40.0,
            rotation: 0.0,
            z_order: 2,
            font_size_pt: None,
            reference_width_pt: None,
            reference_height_pt: None,
        };
        let session = OfficeFormulaSession {
            id: "32345678-1234-4234-9234-123456789abc".to_string(),
            mode: OfficeSessionMode::Edit,
            host: OfficeHost::Powerpoint,
            formula_id: "12345678-1234-4234-9234-123456789abc".to_string(),
            source_document_id: None,
            source_object_id: None,
            title: "Formula".to_string(),
            lines: vec![],
            active_line_id: None,
            code_format: "latex".to_string(),
            display_mode: "block".to_string(),
            numbered: false,
            export_width: 0.0,
            export_height: 0.0,
            export_result: Some(crate::office::sessions::OfficeExportResult {
                svg: "<svg/>".to_string(),
                svg_base64: String::new(),
                png_base64: None,
            omml_base64: None,
            omml_docx_base64: None,
                width: 300.0,
                height: 50.0,
                baseline: None,
            }),
            original_metadata: Some(VisualTeXFormulaMetadata {
                schema: "visualtex-formula".to_string(),
                schema_version: 1,
                formula_id: "12345678-1234-4234-9234-123456789abc".to_string(),
                title: "Formula".to_string(),
                latex: String::new(),
                lines: vec![],
                code_format: "latex".to_string(),
                display_mode: "block".to_string(),
                numbered: false,
                render_width_px: Some(120.0),
                render_height_px: Some(40.0),
                font_size_pt: None,
                reference_width_pt: None,
                reference_height_pt: None,
                reference_baseline_pt: None,
                created_with_version: "1".to_string(),
                updated_with_version: "1".to_string(),
                created_at: "1".to_string(),
                updated_at: "1".to_string(),
            }),
            dirty: true,
            status: OfficeSessionStatus::Committing,
            auto_commit_on_close: true,
            explicit_cancel: false,
            error: None,
            created_at: 1,
            updated_at: 1,
            expires_at: 2,
        };
        let geometry =
            calculate_powerpoint_geometry(&request, &session).expect("geometry should scale");
        assert!((geometry.height - 50.0).abs() < 0.001);
        assert!((geometry.width - 300.0).abs() < 0.001);
        assert!((geometry.left + geometry.width / 2.0 - 160.0).abs() < 0.001);
        assert!((geometry.top + geometry.height / 2.0 - 220.0).abs() < 0.001);
        assert!((geometry.font_size_pt - 18.6666666667).abs() < 0.001);
        assert!((geometry.reference_width_pt - 225.0).abs() < 0.001);
        assert!((geometry.reference_height_pt - 37.5).abs() < 0.001);

        let mut create_request = request.clone();
        create_request.font_size_pt = Some(28.0);
        create_request.reference_width_pt = None;
        create_request.reference_height_pt = None;
        let mut create_session = session.clone();
        create_session.mode = OfficeSessionMode::Create;
        create_session.original_metadata = None;
        let create_geometry = calculate_powerpoint_geometry(&create_request, &create_session)
            .expect("declared PowerPoint point size should scale the SVG");
        assert!((create_geometry.font_size_pt - 28.0).abs() < 0.001);
        assert!((create_geometry.width - 450.0).abs() < 0.001);
        assert!((create_geometry.height - 75.0).abs() < 0.001);
        assert!((create_geometry.left + create_geometry.width / 2.0 - 160.0).abs() < 0.001);
        assert!((create_geometry.top + create_geometry.height / 2.0 - 220.0).abs() < 0.001);
    }
}
