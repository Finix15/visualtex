#![cfg(target_os = "windows")]

use std::ffi::c_void;
use std::sync::mpsc::{self, Receiver, Sender};
use std::sync::OnceLock;
use std::thread;
use std::time::Duration;
use tauri::{AppHandle, Emitter};

const HOTKEY_ID: i32 = 0x5654; // "VT", within RegisterHotKey's application ID range
const MOD_ALT: u32 = 0x0001;
const MOD_CONTROL: u32 = 0x0002;
const MOD_NOREPEAT: u32 = 0x4000;
const VK_O: u32 = 0x4f;
const WM_HOTKEY: u32 = 0x0312;
const PM_REMOVE: u32 = 0x0001;
const SILENT_OCR_EVENT: &str = "visualtex-silent-ocr-global";

#[repr(C)]
#[derive(Clone, Copy, Default)]
struct Point {
    x: i32,
    y: i32,
}

#[repr(C)]
#[derive(Clone, Copy)]
struct Msg {
    hwnd: *mut c_void,
    message: u32,
    w_param: usize,
    l_param: isize,
    time: u32,
    pt: Point,
    l_private: u32,
}

impl Default for Msg {
    fn default() -> Self {
        Self {
            hwnd: std::ptr::null_mut(),
            message: 0,
            w_param: 0,
            l_param: 0,
            time: 0,
            pt: Point::default(),
            l_private: 0,
        }
    }
}

#[link(name = "user32")]
unsafe extern "system" {
    fn RegisterHotKey(hwnd: *mut c_void, id: i32, modifiers: u32, virtual_key: u32) -> i32;
    fn UnregisterHotKey(hwnd: *mut c_void, id: i32) -> i32;
    fn PeekMessageW(
        message: *mut Msg,
        hwnd: *mut c_void,
        min_filter: u32,
        max_filter: u32,
        remove_message: u32,
    ) -> i32;
}

enum HotkeyCommand {
    SetRegistered {
        enabled: bool,
        reply: Sender<Result<(), String>>,
    },
}

static COMMAND_SENDER: OnceLock<Sender<HotkeyCommand>> = OnceLock::new();

fn register_hotkey() -> Result<(), String> {
    let result = unsafe {
        RegisterHotKey(
            std::ptr::null_mut(),
            HOTKEY_ID,
            MOD_CONTROL | MOD_ALT | MOD_NOREPEAT,
            VK_O,
        )
    };
    if result != 0 {
        return Ok(());
    }
    Err(format!(
        "Unable to register Ctrl+Alt+O for silent OCR: {}",
        std::io::Error::last_os_error()
    ))
}

fn unregister_hotkey() -> Result<(), String> {
    let result = unsafe { UnregisterHotKey(std::ptr::null_mut(), HOTKEY_ID) };
    if result != 0 {
        return Ok(());
    }
    let error = std::io::Error::last_os_error();
    // ERROR_HOTKEY_NOT_REGISTERED means the desired disabled state is already
    // satisfied, for example after a failed registration attempt.
    if error.raw_os_error() == Some(1419) {
        return Ok(());
    }
    Err(format!("Unable to unregister the silent OCR hotkey: {error}"))
}

fn hotkey_thread(app: AppHandle, receiver: Receiver<HotkeyCommand>) {
    let mut registered = false;
    loop {
        match receiver.recv_timeout(Duration::from_millis(12)) {
            Ok(HotkeyCommand::SetRegistered { enabled, reply }) => {
                let result = if enabled == registered {
                    Ok(())
                } else if enabled {
                    register_hotkey().inspect(|_| registered = true)
                } else {
                    unregister_hotkey().inspect(|_| registered = false)
                };
                let _ = reply.send(result);
            }
            Err(mpsc::RecvTimeoutError::Timeout) => {}
            Err(mpsc::RecvTimeoutError::Disconnected) => {
                if registered {
                    let _ = unregister_hotkey();
                }
                return;
            }
        }

        let mut message = Msg::default();
        while unsafe {
            PeekMessageW(
                &mut message,
                std::ptr::null_mut(),
                0,
                0,
                PM_REMOVE,
            )
        } != 0
        {
            if message.message == WM_HOTKEY && message.w_param == HOTKEY_ID as usize {
                let _ = app.emit(SILENT_OCR_EVENT, ());
            }
        }
    }
}

pub(crate) fn initialize(app: &AppHandle) -> Result<(), String> {
    if COMMAND_SENDER.get().is_some() {
        return Ok(());
    }
    let (sender, receiver) = mpsc::channel();
    COMMAND_SENDER
        .set(sender)
        .map_err(|_| "Silent OCR hotkey bridge is already initialized".to_string())?;
    let app = app.clone();
    thread::Builder::new()
        .name("visualtex-silent-ocr-hotkey".to_string())
        .spawn(move || hotkey_thread(app, receiver))
        .map_err(|error| format!("Unable to start the silent OCR hotkey bridge: {error}"))?;
    Ok(())
}

pub(crate) fn set_registered(enabled: bool) -> Result<(), String> {
    let sender = COMMAND_SENDER
        .get()
        .ok_or_else(|| "Silent OCR hotkey bridge is unavailable".to_string())?;
    let (reply_sender, reply_receiver) = mpsc::channel();
    sender
        .send(HotkeyCommand::SetRegistered {
            enabled,
            reply: reply_sender,
        })
        .map_err(|_| "Silent OCR hotkey bridge has stopped".to_string())?;
    reply_receiver
        .recv_timeout(Duration::from_secs(2))
        .map_err(|_| "Silent OCR hotkey update timed out".to_string())?
}
