use std::collections::BTreeSet;
use std::ffi::OsStr;
use std::fs::{self, OpenOptions};
use std::io::Write;
use std::path::PathBuf;
use std::time::{SystemTime, UNIX_EPOCH};
use tauri::webview::PageLoadEvent;
use tauri::{AppHandle, Manager, WebviewUrl, WebviewWindow, WebviewWindowBuilder};

pub const OFFICE_BACKGROUND_ARGUMENT: &str = "--office-background";
pub const OFFICE_BOOTSTRAP_ARGUMENT: &str = "--office-bootstrap";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AppRunMode {
    Desktop,
    OfficeBackground,
    OfficeBootstrap,
}

impl AppRunMode {
    pub fn current() -> Self {
        Self::from_arguments(std::env::args_os())
    }

    pub fn from_arguments(arguments: impl IntoIterator<Item = impl AsRef<OsStr>>) -> Self {
        let mut background = false;
        for argument in arguments {
            let argument = argument.as_ref();
            if argument == OsStr::new(OFFICE_BOOTSTRAP_ARGUMENT) {
                return Self::OfficeBootstrap;
            }
            if argument == OsStr::new(OFFICE_BACKGROUND_ARGUMENT) {
                background = true;
            }
        }
        if background {
            Self::OfficeBackground
        } else {
            Self::Desktop
        }
    }

    pub fn label(self) -> &'static str {
        match self {
            Self::Desktop => "desktop",
            Self::OfficeBackground => "office-background",
            Self::OfficeBootstrap => "office-bootstrap",
        }
    }

    pub fn creates_main_window(self) -> bool {
        matches!(self, Self::Desktop)
    }

    pub fn starts_companion(self) -> bool {
        !matches!(self, Self::OfficeBootstrap)
    }

    pub fn schedules_ocr_warmup(self) -> bool {
        !matches!(self, Self::OfficeBootstrap)
    }
}

pub fn arguments_request_desktop(arguments: &[String]) -> bool {
    !arguments.iter().any(|argument| {
        argument == OFFICE_BACKGROUND_ARGUMENT || argument == OFFICE_BOOTSTRAP_ARGUMENT
    })
}

fn lifecycle_log_path() -> PathBuf {
    #[cfg(target_os = "windows")]
    let root = std::env::var_os("LOCALAPPDATA")
        .map(PathBuf::from)
        .unwrap_or_else(std::env::temp_dir)
        .join("VisualTeX")
        .join("logs");
    #[cfg(not(target_os = "windows"))]
    let root = std::env::temp_dir().join("VisualTeX").join("logs");
    root.join("app-lifecycle.log")
}

pub fn append_lifecycle_log(message: impl AsRef<str>) {
    let path = lifecycle_log_path();
    let Some(parent) = path.parent() else {
        return;
    };
    if fs::create_dir_all(parent).is_err() {
        return;
    }
    let timestamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|value| value.as_secs())
        .unwrap_or_default();
    if let Ok(mut file) = OpenOptions::new().create(true).append(true).open(path) {
        let _ = writeln!(
            file,
            "[{timestamp}] pid={} {}",
            std::process::id(),
            message.as_ref()
        );
        let _ = file.flush();
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MainAssetDiagnostic {
    pub index_present: bool,
    pub referenced_assets: Vec<String>,
    pub missing_assets: Vec<String>,
    pub detail: String,
}

fn extract_attribute_assets(source: &str, marker: &str, quote: char, output: &mut BTreeSet<String>) {
    let mut remainder = source;
    while let Some(index) = remainder.find(marker) {
        let value_start = index + marker.len();
        let tail = &remainder[value_start..];
        let Some(value_end) = tail.find(quote) else {
            break;
        };
        let raw = &tail[..value_end];
        let path = raw
            .split(['?', '#'])
            .next()
            .unwrap_or(raw)
            .trim_start_matches('/')
            .trim();
        if path.starts_with("assets/") && (path.ends_with(".js") || path.ends_with(".css")) {
            output.insert(path.to_string());
        }
        remainder = &tail[value_end + quote.len_utf8()..];
    }
}

pub fn referenced_frontend_assets(index_html: &str) -> Vec<String> {
    let mut assets = BTreeSet::new();
    extract_attribute_assets(index_html, "src=\"", '"', &mut assets);
    extract_attribute_assets(index_html, "href=\"", '"', &mut assets);
    extract_attribute_assets(index_html, "src='", '\'', &mut assets);
    extract_attribute_assets(index_html, "href='", '\'', &mut assets);
    assets.into_iter().collect()
}

pub fn diagnose_main_assets(
    index_html: Option<&[u8]>,
    mut asset_exists: impl FnMut(&str) -> bool,
) -> MainAssetDiagnostic {
    let Some(index_html) = index_html else {
        return MainAssetDiagnostic {
            index_present: false,
            referenced_assets: Vec::new(),
            missing_assets: vec!["index.html".to_string()],
            detail: "Tauri embedded main asset index.html is missing; Office companion resources are a separate installed resource tree and were not used for this check".to_string(),
        };
    };
    let Ok(index_html) = std::str::from_utf8(index_html) else {
        return MainAssetDiagnostic {
            index_present: true,
            referenced_assets: Vec::new(),
            missing_assets: vec!["index.html (invalid UTF-8)".to_string()],
            detail: "Tauri embedded index.html exists but is not valid UTF-8".to_string(),
        };
    };
    let referenced_assets = referenced_frontend_assets(index_html);
    let missing_assets = referenced_assets
        .iter()
        .filter(|path| !asset_exists(path))
        .cloned()
        .collect::<Vec<_>>();
    let detail = if missing_assets.is_empty() {
        format!(
            "Tauri embedded main assets verified: index.html plus {} referenced JS/CSS assets",
            referenced_assets.len()
        )
    } else {
        format!(
            "Tauri embedded index.html exists but referenced main assets are missing: {}",
            missing_assets.join(", ")
        )
    };
    MainAssetDiagnostic {
        index_present: true,
        referenced_assets,
        missing_assets,
        detail,
    }
}

pub fn validate_embedded_main_assets(app: &AppHandle) -> Result<MainAssetDiagnostic, String> {
    let resolver = app.asset_resolver();
    let index = resolver.get("index.html".to_string());
    let diagnostic = diagnose_main_assets(index.as_ref().map(|asset| asset.bytes.as_slice()), |path| {
        resolver.get(path.to_string()).is_some()
    });
    append_lifecycle_log(format!(
        "main-asset diagnostic index_present={} referenced={} missing={} detail={}",
        diagnostic.index_present,
        diagnostic.referenced_assets.len(),
        diagnostic.missing_assets.join("|"),
        diagnostic.detail
    ));
    if diagnostic.index_present && diagnostic.missing_assets.is_empty() {
        Ok(diagnostic)
    } else {
        Err(diagnostic.detail.clone())
    }
}

fn focus_existing_main_window(window: &WebviewWindow) -> Result<(), String> {
    if let Err(error) = window.show() {
        let message = format!("Unable to show existing main window: {error}");
        append_lifecycle_log(&message);
        return Err(message);
    }
    if let Err(error) = window.unminimize() {
        let message = format!("Unable to unminimize existing main window: {error}");
        append_lifecycle_log(&message);
        return Err(message);
    }
    if let Err(error) = window.set_focus() {
        let message = format!("Unable to focus existing main window: {error}");
        append_lifecycle_log(&message);
        return Err(message);
    }
    append_lifecycle_log("existing main window shown, unminimized and focused");
    Ok(())
}

pub fn ensure_main_window(app: &AppHandle) -> Result<WebviewWindow, String> {
    if let Some(window) = app.get_webview_window("main") {
        focus_existing_main_window(&window)?;
        return Ok(window);
    }

    validate_embedded_main_assets(app).map_err(|error| {
        append_lifecycle_log(format!("main window creation blocked by embedded asset failure: {error}"));
        error
    })?;

    append_lifecycle_log("creating main window from WebviewUrl::App(index.html)");
    let window = WebviewWindowBuilder::new(
        app,
        "main",
        WebviewUrl::App("index.html".into()),
    )
    .title("VisualTeX")
    .inner_size(1240.0, 820.0)
    .min_inner_size(640.0, 480.0)
    .resizable(true)
    .center()
    .focused(true)
    .visible(true)
    .on_page_load(|_window, payload| {
        append_lifecycle_log(format!(
            "main window page-load event={:?} url={}",
            payload.event(),
            payload.url()
        ));
        if matches!(payload.event(), PageLoadEvent::Finished) {
            append_lifecycle_log("main window finished loading embedded application entry");
        }
    })
    .build()
    .map_err(|error| {
        let message = format!("Unable to create VisualTeX main window: {error}");
        append_lifecycle_log(&message);
        message
    })?;

    focus_existing_main_window(&window)?;
    append_lifecycle_log("new main window created successfully");
    Ok(window)
}

pub fn background_retention_enabled() -> bool {
    #[cfg(target_os = "windows")]
    {
        crate::office::windows_backend::background_start_enabled()
    }
    #[cfg(not(target_os = "windows"))]
    {
        crate::office::background::status().installed
    }
}

pub fn destroy_main_window_for_background(app: &AppHandle) -> Result<(), String> {
    let Some(window) = app.get_webview_window("main") else {
        append_lifecycle_log("main window close requested for background retention, but main no longer exists");
        return Ok(());
    };
    window.destroy().map_err(|error| {
        let message = format!("Unable to destroy main window while retaining background companion: {error}");
        append_lifecycle_log(&message);
        message
    })?;
    append_lifecycle_log("main window destroyed; background companion remains available for a future desktop launch");
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn run_modes_are_distinct_and_bootstrap_has_priority() {
        assert_eq!(
            AppRunMode::from_arguments(["visualtex.exe"]),
            AppRunMode::Desktop
        );
        assert_eq!(
            AppRunMode::from_arguments(["visualtex.exe", OFFICE_BACKGROUND_ARGUMENT]),
            AppRunMode::OfficeBackground
        );
        assert_eq!(
            AppRunMode::from_arguments([
                "visualtex.exe",
                OFFICE_BACKGROUND_ARGUMENT,
                OFFICE_BOOTSTRAP_ARGUMENT,
            ]),
            AppRunMode::OfficeBootstrap
        );
        assert!(AppRunMode::Desktop.creates_main_window());
        assert!(!AppRunMode::OfficeBackground.creates_main_window());
        assert!(!AppRunMode::OfficeBootstrap.starts_companion());
        assert!(!AppRunMode::OfficeBootstrap.schedules_ocr_warmup());
    }

    #[test]
    fn background_single_instance_notification_can_request_desktop() {
        assert!(arguments_request_desktop(&[
            "C:\\VisualTeX\\visualtex.exe".to_string()
        ]));
        assert!(!arguments_request_desktop(&[
            "visualtex.exe".to_string(),
            OFFICE_BACKGROUND_ARGUMENT.to_string(),
        ]));
        assert!(!arguments_request_desktop(&[
            "visualtex.exe".to_string(),
            OFFICE_BOOTSTRAP_ARGUMENT.to_string(),
        ]));
    }

    #[test]
    fn main_asset_diagnostic_isolated_from_office_resource_success() {
        let diagnostic = diagnose_main_assets(None, |_| true);
        assert!(!diagnostic.index_present);
        assert_eq!(diagnostic.missing_assets, ["index.html"]);
        assert!(diagnostic.detail.contains("Tauri embedded main asset"));
        assert!(diagnostic.detail.contains("Office companion resources are a separate"));
    }

    #[test]
    fn main_asset_diagnostic_reports_missing_referenced_bundle() {
        let html = br#"<link rel="stylesheet" href="/assets/index-a.css"><script type="module" src="/assets/index-b.js"></script>"#;
        let diagnostic = diagnose_main_assets(Some(html), |path| path.ends_with(".css"));
        assert!(diagnostic.index_present);
        assert_eq!(
            diagnostic.referenced_assets,
            ["assets/index-a.css", "assets/index-b.js"]
        );
        assert_eq!(diagnostic.missing_assets, ["assets/index-b.js"]);
    }

    #[test]
    fn main_asset_diagnostic_accepts_complete_embedded_bundle() {
        let html = br#"<link rel='stylesheet' href='/assets/index-a.css?x=1'><script src='/assets/index-b.js#v'></script>"#;
        let diagnostic = diagnose_main_assets(Some(html), |_| true);
        assert!(diagnostic.index_present);
        assert!(diagnostic.missing_assets.is_empty());
        assert_eq!(diagnostic.referenced_assets.len(), 2);
    }
}
