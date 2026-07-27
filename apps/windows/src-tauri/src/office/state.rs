use crate::office::formula_cache::FormulaMetadataCache;
use crate::office::platform::{self, OfficePlatformBackend};
use crate::office::powerpoint_native::{
    PowerPointInteractionBus, PowerPointNativeSelection,
};
use crate::office::sessions::SessionStore;
use crate::OcrState;
use axum_server::Handle;
use serde::Serialize;
use std::collections::HashMap;
use std::fs::{self, OpenOptions};
use std::io::Write;
use std::net::SocketAddr;
use std::path::PathBuf;
use std::time::{SystemTime, UNIX_EPOCH};
use std::sync::{Arc, Mutex, RwLock};
use tauri::AppHandle;

pub const OFFICE_BIND_IP: [u8; 4] = [127, 0, 0, 1];
pub const OFFICE_PORT: u16 = 43_127;
pub const OFFICE_PROTOCOL_VERSION: u32 = 1;
pub const OFFICE_UI_VERSION: &str = env!("CARGO_PKG_VERSION");
pub const MAX_OFFICE_REQUEST_BYTES: usize = 22 * 1024 * 1024;

pub fn normalize_app_theme(theme: &str) -> &'static str {
    match theme.trim() {
        "beige" => "beige",
        "dark" => "dark",
        "purple" => "purple",
        "green" => "green",
        _ => "light",
    }
}

pub fn append_office_log(paths: &OfficePaths, file_name: &str, message: &str) {
    #[cfg(target_os = "windows")]
    let log_root = std::env::var_os("LOCALAPPDATA")
        .map(PathBuf::from)
        .unwrap_or_else(std::env::temp_dir)
        .join("VisualTeX")
        .join("office")
        .join("logs");
    #[cfg(not(target_os = "windows"))]
    let log_root = paths.root.join("logs");
    #[cfg(target_os = "windows")]
    let _ = paths;
    if fs::create_dir_all(&log_root).is_err() {
        return;
    }
    let timestamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|value| value.as_secs())
        .unwrap_or_default();
    let path = log_root.join(file_name);
    if let Ok(mut file) = OpenOptions::new().create(true).append(true).open(path) {
        let _ = writeln!(file, "[{timestamp}] {message}");
        let _ = file.flush();
    }
}

#[derive(Debug, Clone)]
pub struct OfficePaths {
    pub root: PathBuf,
    pub certificate: PathBuf,
    pub private_key: PathBuf,
    pub certificate_metadata: PathBuf,
    pub install: PathBuf,
    pub sessions: PathBuf,
    pub recovery: PathBuf,
    pub formula_cache: PathBuf,
    pub ui_root: PathBuf,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct OfficeCompanionStatus {
    pub running: bool,
    pub bind_address: String,
    pub port: u16,
    pub certificate_path: String,
    pub office_ui_version: String,
    pub protocol_version: u32,
    pub last_error: Option<String>,
}

impl OfficeCompanionStatus {
    pub fn stopped(paths: &OfficePaths) -> Self {
        Self {
            running: false,
            bind_address: "127.0.0.1".to_string(),
            port: OFFICE_PORT,
            certificate_path: paths.certificate.display().to_string(),
            office_ui_version: OFFICE_UI_VERSION.to_string(),
            protocol_version: OFFICE_PROTOCOL_VERSION,
            last_error: None,
        }
    }
}

#[derive(Clone)]
pub struct OfficeCompanionState {
    pub app: Option<AppHandle>,
    pub ocr: OcrState,
    pub paths: Arc<OfficePaths>,
    pub install_token: Arc<String>,
    pub status: Arc<RwLock<OfficeCompanionStatus>>,
    pub app_theme: Arc<RwLock<String>>,
    pub server_handle: Arc<Mutex<Option<Handle<SocketAddr>>>>,
    pub session_store: SessionStore,
    pub formula_cache: FormulaMetadataCache,
    pub platform_backend: Arc<dyn OfficePlatformBackend>,
    pub powerpoint_interactions: PowerPointInteractionBus,
    /// Native PowerPoint insertion is prepared before the Office.js command
    /// page writes durable tags and accessibility metadata. Keep the immutable
    /// prepared selection by Session id so retries never paste a second image.
    pub prepared_powerpoint_commits:
        Arc<Mutex<HashMap<String, PowerPointNativeSelection>>>,
    pub ocr_available: bool,
}

impl OfficeCompanionState {
    pub fn new(
        app: Option<AppHandle>,
        ocr: OcrState,
        paths: OfficePaths,
        install_token: String,
        session_store: SessionStore,
        formula_cache: FormulaMetadataCache,
        ocr_available: bool,
    ) -> Self {
        let status = OfficeCompanionStatus::stopped(&paths);
        let platform_backend = platform::create_backend(app.as_ref(), &paths);
        Self {
            app,
            ocr,
            paths: Arc::new(paths),
            install_token: Arc::new(install_token),
            status: Arc::new(RwLock::new(status)),
            app_theme: Arc::new(RwLock::new("light".to_string())),
            server_handle: Arc::new(Mutex::new(None)),
            session_store,
            formula_cache,
            platform_backend,
            powerpoint_interactions: PowerPointInteractionBus::default(),
            prepared_powerpoint_commits: Arc::new(Mutex::new(HashMap::new())),
            ocr_available,
        }
    }

    pub fn socket_addr() -> SocketAddr {
        SocketAddr::from((OFFICE_BIND_IP, OFFICE_PORT))
    }

    pub fn snapshot(&self) -> OfficeCompanionStatus {
        self.status
            .read()
            .map(|value| value.clone())
            .unwrap_or_else(|_| OfficeCompanionStatus::stopped(&self.paths))
    }

    pub fn update_status(&self, mutate: impl FnOnce(&mut OfficeCompanionStatus)) {
        if let Ok(mut status) = self.status.write() {
            mutate(&mut status);
        }
    }

    pub fn current_theme(&self) -> String {
        self.app_theme
            .read()
            .map(|theme| theme.clone())
            .unwrap_or_else(|_| "light".to_string())
    }

    pub fn set_current_theme(&self, theme: &str) -> String {
        let normalized = normalize_app_theme(theme).to_string();
        if let Ok(mut current) = self.app_theme.write() {
            *current = normalized.clone();
        }
        normalized
    }
}
