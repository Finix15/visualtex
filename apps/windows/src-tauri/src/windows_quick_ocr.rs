#![cfg(target_os = "windows")]

use base64::{engine::general_purpose::STANDARD as BASE64_STANDARD, Engine as _};
use serde::Serialize;
use std::ffi::c_void;
use std::os::windows::process::CommandExt;
use std::process::{Command, Stdio};
use std::thread;
use std::time::Duration;

const CREATE_NO_WINDOW: u32 = 0x0800_0000;
const MIN_TIMEOUT_MS: u64 = 1_000;
const MAX_TIMEOUT_MS: u64 = 120_000;
const CF_UNICODETEXT: u32 = 13;
const GMEM_MOVEABLE: u32 = 0x0002;
const CLIPBOARD_OPEN_RETRIES: usize = 25;
const CLIPBOARD_OPEN_RETRY_DELAY_MS: u64 = 20;

#[link(name = "user32")]
unsafe extern "system" {
    fn OpenClipboard(hwnd: *mut c_void) -> i32;
    fn CloseClipboard() -> i32;
    fn EmptyClipboard() -> i32;
    fn SetClipboardData(format: u32, memory: *mut c_void) -> *mut c_void;
}

#[link(name = "kernel32")]
unsafe extern "system" {
    fn GlobalAlloc(flags: u32, bytes: usize) -> *mut c_void;
    fn GlobalLock(memory: *mut c_void) -> *mut c_void;
    fn GlobalUnlock(memory: *mut c_void) -> i32;
    fn GlobalFree(memory: *mut c_void) -> *mut c_void;
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct WindowsQuickOcrCapture {
    data_base64: String,
    extension: String,
}

fn encode_powershell(script: &str) -> String {
    let bytes = script
        .encode_utf16()
        .flat_map(|unit| unit.to_le_bytes())
        .collect::<Vec<_>>();
    BASE64_STANDARD.encode(bytes)
}

fn capture_script(launch_capture: bool, timeout_ms: u64) -> String {
    let launch = if launch_capture { "$true" } else { "$false" };
    format!(
        r#"
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
public static class VisualTeXClipboardNative {{
    [DllImport("user32.dll")]
    public static extern uint GetClipboardSequenceNumber();
}}
'@

$before = [VisualTeXClipboardNative]::GetClipboardSequenceNumber()
if ({launch}) {{
    Start-Process -FilePath 'explorer.exe' -ArgumentList 'ms-screenclip:' -WindowStyle Hidden | Out-Null
}}

$deadline = [DateTime]::UtcNow.AddMilliseconds({timeout_ms})
while ([DateTime]::UtcNow -lt $deadline) {{
    Start-Sleep -Milliseconds 120
    if ([VisualTeXClipboardNative]::GetClipboardSequenceNumber() -eq $before) {{
        continue
    }}
    try {{
        if (-not [System.Windows.Forms.Clipboard]::ContainsImage()) {{
            continue
        }}
        $image = [System.Windows.Forms.Clipboard]::GetImage()
        if ($null -eq $image) {{
            continue
        }}
        try {{
            $stream = New-Object System.IO.MemoryStream
            try {{
                $image.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
                [Console]::Out.Write([Convert]::ToBase64String($stream.ToArray()))
                exit 0
            }} finally {{
                $stream.Dispose()
            }}
        }} finally {{
            $image.Dispose()
        }}
    }} catch {{
        # The clipboard can be momentarily locked while Snipping Tool commits.
        # Keep polling until the timeout instead of surfacing a transient error.
    }}
}}
exit 2
"#,
    )
}

fn capture_windows_clipboard_image(
    launch_capture: bool,
    timeout_ms: u64,
) -> Result<Option<WindowsQuickOcrCapture>, String> {
    let timeout_ms = timeout_ms.clamp(MIN_TIMEOUT_MS, MAX_TIMEOUT_MS);
    let script = capture_script(launch_capture, timeout_ms);
    let encoded = encode_powershell(&script);
    let output = Command::new("powershell.exe")
        .args([
            "-NoProfile",
            "-STA",
            "-NonInteractive",
            "-EncodedCommand",
            encoded.as_str(),
        ])
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .creation_flags(CREATE_NO_WINDOW)
        .output()
        .map_err(|error| format!("Unable to start the Windows screenshot bridge: {error}"))?;

    if output.status.success() {
        let data_base64 = String::from_utf8(output.stdout)
            .map_err(|_| "Windows screenshot bridge returned invalid UTF-8 output".to_string())?
            .trim()
            .to_string();
        if data_base64.is_empty() {
            return Err("Windows screenshot bridge returned an empty image".to_string());
        }
        return Ok(Some(WindowsQuickOcrCapture {
            data_base64,
            extension: "png".to_string(),
        }));
    }

    if output.status.code() == Some(2) {
        return Ok(None);
    }

    let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();
    let stdout = String::from_utf8_lossy(&output.stdout).trim().to_string();
    let detail = if !stderr.is_empty() { stderr } else { stdout };
    Err(if detail.is_empty() {
        format!(
            "Windows screenshot bridge exited with status {:?}",
            output.status.code()
        )
    } else {
        format!("Windows screenshot bridge failed: {detail}")
    })
}

fn write_clipboard_text(text: &str) -> Result<(), String> {
    let mut opened = false;
    for _ in 0..CLIPBOARD_OPEN_RETRIES {
        if unsafe { OpenClipboard(std::ptr::null_mut()) } != 0 {
            opened = true;
            break;
        }
        thread::sleep(Duration::from_millis(CLIPBOARD_OPEN_RETRY_DELAY_MS));
    }
    if !opened {
        return Err(format!(
            "Unable to open the Windows clipboard: {}",
            std::io::Error::last_os_error()
        ));
    }

    let result = (|| {
        if unsafe { EmptyClipboard() } == 0 {
            return Err(format!(
                "Unable to clear the Windows clipboard: {}",
                std::io::Error::last_os_error()
            ));
        }

        let utf16 = text.encode_utf16().chain(std::iter::once(0)).collect::<Vec<_>>();
        let byte_len = utf16.len() * std::mem::size_of::<u16>();
        let memory = unsafe { GlobalAlloc(GMEM_MOVEABLE, byte_len) };
        if memory.is_null() {
            return Err(format!(
                "Unable to allocate Windows clipboard memory: {}",
                std::io::Error::last_os_error()
            ));
        }

        let locked = unsafe { GlobalLock(memory) };
        if locked.is_null() {
            unsafe {
                let _ = GlobalFree(memory);
            }
            return Err(format!(
                "Unable to lock Windows clipboard memory: {}",
                std::io::Error::last_os_error()
            ));
        }
        unsafe {
            std::ptr::copy_nonoverlapping(
                utf16.as_ptr(),
                locked.cast::<u16>(),
                utf16.len(),
            );
            let _ = GlobalUnlock(memory);
        }

        if unsafe { SetClipboardData(CF_UNICODETEXT, memory) }.is_null() {
            unsafe {
                let _ = GlobalFree(memory);
            }
            return Err(format!(
                "Unable to write text to the Windows clipboard: {}",
                std::io::Error::last_os_error()
            ));
        }

        // Ownership of memory transfers to the system after SetClipboardData.
        Ok(())
    })();

    unsafe {
        let _ = CloseClipboard();
    }
    result
}

#[tauri::command]
pub(crate) async fn capture_windows_quick_ocr(
    launch_capture: bool,
    timeout_ms: u64,
) -> Result<Option<WindowsQuickOcrCapture>, String> {
    tauri::async_runtime::spawn_blocking(move || {
        capture_windows_clipboard_image(launch_capture, timeout_ms)
    })
    .await
    .map_err(|error| format!("Windows screenshot bridge task failed: {error}"))?
}

#[tauri::command]
pub(crate) async fn write_windows_ocr_clipboard_text(text: String) -> Result<(), String> {
    tauri::async_runtime::spawn_blocking(move || write_clipboard_text(&text))
        .await
        .map_err(|error| format!("Windows clipboard task failed: {error}"))?
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn powershell_script_uses_native_screenclip_and_clipboard_sequence() {
        let script = capture_script(true, 60_000);
        assert!(script.contains("ms-screenclip:"));
        assert!(script.contains("GetClipboardSequenceNumber"));
        assert!(script.contains("Clipboard]::ContainsImage"));
        assert!(script.contains("ImageFormat]::Png"));
    }

    #[test]
    fn wait_only_script_does_not_launch_screenclip() {
        let script = capture_script(false, 60_000);
        assert!(script.contains("if ($false)"));
    }
}
