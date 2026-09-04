use serde::Deserialize;
use std::{
    ffi::OsString,
    io::{Read, Write},
    path::{Path, PathBuf},
    process::{Child, Command, Stdio},
    sync::{Arc, Mutex},
    thread,
};
use tauri::{AppHandle, Manager};

pub const MAX_STDOUT_BYTES: usize = 64 * 1024;
pub const MAX_STDERR_BYTES: usize = 1024 * 1024;

#[derive(Debug, Clone, Copy)]
pub enum WorkerOperation {
    Scan,
    Finalize,
}

impl WorkerOperation {
    fn argument(self) -> &'static str {
        match self {
            Self::Scan => "scan",
            Self::Finalize => "finalize",
        }
    }

    fn expected_status(self) -> &'static str {
        match self {
            Self::Scan => "awaitingOmml",
            Self::Finalize => "complete",
        }
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct WorkerStatus {
    pub protocol_version: u32,
    pub job_id: String,
    pub status: String,
}

#[derive(Debug)]
pub struct RunningWorker {
    pub child: Arc<Mutex<Option<Child>>>,
    completion: Arc<Mutex<Option<Result<WorkerStatus, String>>>>,
}

impl RunningWorker {
    pub fn is_running(&self) -> bool {
        let Ok(mut lock) = self.child.lock() else {
            return false;
        };
        let Some(child) = lock.as_mut() else {
            return false;
        };
        matches!(child.try_wait(), Ok(None))
    }

    pub fn completion(&self) -> Option<Result<WorkerStatus, String>> {
        self.completion.lock().ok()?.clone()
    }
}

fn production_worker(app: &AppHandle) -> Result<PathBuf, String> {
    app.path()
        .resource_dir()
        .map_err(|error| error.to_string())
        .map(|root| root.join("workers").join("mathtype-worker"))
}

fn is_regular_executable(path: &Path) -> bool {
    std::fs::symlink_metadata(path)
        .map(|metadata| metadata.file_type().is_file() && !metadata.file_type().is_symlink())
        .unwrap_or(false)
}

fn select_worker(production: PathBuf, development: Option<PathBuf>) -> Result<PathBuf, String> {
    if is_regular_executable(&production) {
        return Ok(production);
    }
    if let Some(path) = development {
        if path.is_absolute() && is_regular_executable(&path) {
            return Ok(path);
        }
        return Err("Development worker fallback must be an absolute regular file".to_string());
    }
    Err(format!(
        "Bundled MathType worker is missing: {}",
        production.display()
    ))
}

pub fn resolve_worker(app: &AppHandle) -> Result<PathBuf, String> {
    let production = production_worker(app)?;
    #[cfg(debug_assertions)]
    let development = std::env::var_os("VISUALTEX_MATHTYPE_WORKER_DEV").map(PathBuf::from);
    #[cfg(not(debug_assertions))]
    let development = None;
    select_worker(production, development)
}

fn read_limited<R: Read>(mut reader: R, maximum: usize) -> Result<Vec<u8>, String> {
    let mut result = Vec::new();
    let mut buffer = [0_u8; 8192];
    let mut oversized = false;
    loop {
        let count = reader
            .read(&mut buffer)
            .map_err(|error| error.to_string())?;
        if count == 0 {
            break;
        }
        if result.len() + count <= maximum {
            result.extend_from_slice(&buffer[..count]);
        } else {
            oversized = true;
        }
    }
    if oversized {
        return Err("Worker stream exceeded configured limit".to_string());
    }
    Ok(result)
}

fn redact_stderr(bytes: &[u8], job_root: &Path) -> Vec<u8> {
    let text = String::from_utf8_lossy(bytes);
    text.replace(&job_root.to_string_lossy().to_string(), "<job-root>")
        .into_bytes()
}

pub fn spawn_worker(
    executable: &Path,
    job_id: &str,
    job_root: &Path,
    operation: WorkerOperation,
) -> Result<
    (
        RunningWorker,
        thread::JoinHandle<Result<WorkerStatus, String>>,
    ),
    String,
> {
    let args: [OsString; 8] = [
        "--protocol".into(),
        "1".into(),
        "--job-id".into(),
        job_id.into(),
        "--job-root".into(),
        job_root.as_os_str().to_owned(),
        "--operation".into(),
        operation.argument().into(),
    ];
    let mut child = Command::new(executable)
        .args(args)
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .map_err(|error| format!("Unable to start MathType worker: {error}"))?;
    let stdout = child
        .stdout
        .take()
        .ok_or_else(|| "Worker stdout unavailable".to_string())?;
    let stderr = child
        .stderr
        .take()
        .ok_or_else(|| "Worker stderr unavailable".to_string())?;
    let shared = Arc::new(Mutex::new(Some(child)));
    let completion = Arc::new(Mutex::new(None));
    let process = RunningWorker {
        child: shared.clone(),
        completion: completion.clone(),
    };
    let root = job_root.to_owned();
    let expected_id = job_id.to_string();
    let expected_status = operation.expected_status();
    let thread_shared = shared.clone();
    let handle = thread::spawn(move || {
        let outcome = (|| {
            let stdout_reader = thread::spawn(move || read_limited(stdout, MAX_STDOUT_BYTES));
            let stderr_reader = thread::spawn(move || read_limited(stderr, MAX_STDERR_BYTES));
            let status = loop {
                let result = {
                    let mut lock = thread_shared
                        .lock()
                        .map_err(|_| "Worker lock poisoned".to_string())?;
                    let child = lock
                        .as_mut()
                        .ok_or_else(|| "Worker was cancelled".to_string())?;
                    child.try_wait().map_err(|error| error.to_string())?
                };
                if let Some(status) = result {
                    break status;
                }
                thread::sleep(std::time::Duration::from_millis(10));
            };
            let stdout = stdout_reader
                .join()
                .map_err(|_| "Worker stdout reader panicked".to_string())??;
            let stderr = stderr_reader
                .join()
                .map_err(|_| "Worker stderr reader panicked".to_string())??;
            let stderr = redact_stderr(&stderr, &root);
            let mut log = std::fs::File::create(root.join("worker.stderr.log"))
                .map_err(|error| error.to_string())?;
            log.write_all(&stderr).map_err(|error| error.to_string())?;
            log.sync_all().map_err(|error| error.to_string())?;
            *thread_shared
                .lock()
                .map_err(|_| "Worker lock poisoned".to_string())? = None;
            if !status.success() {
                return Err(format!("MathType worker exited with {status}"));
            }
            let response: WorkerStatus = serde_json::from_slice(&stdout)
                .map_err(|error| format!("Malformed worker response: {error}"))?;
            if response.protocol_version != 1
                || response.job_id != expected_id
                || response.status != expected_status
            {
                return Err("Worker response did not match the job protocol".to_string());
            }
            Ok(response)
        })();
        if let Ok(mut slot) = completion.lock() {
            *slot = Some(outcome.clone());
        }
        outcome
    });
    Ok((process, handle))
}

pub fn cancel_worker(worker: &RunningWorker) -> Result<(), String> {
    let mut lock = worker
        .child
        .lock()
        .map_err(|_| "Worker lock poisoned".to_string())?;
    if let Some(child) = lock.as_mut() {
        child
            .kill()
            .map_err(|error| format!("Unable to cancel worker: {error}"))?;
        let _ = child.wait();
    }
    *lock = None;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn bounded_reader_rejects_oversized_stdout_and_stderr() {
        assert!(read_limited(&[0_u8; 5][..], 4).is_err());
        assert_eq!(read_limited(&b"ok"[..], 4).unwrap(), b"ok");
    }

    #[test]
    fn stderr_log_redacts_the_private_job_path() {
        let root = Path::new("/private/app-data/legacy-equation-jobs/job-id");
        let result = redact_stderr(
            b"failed at /private/app-data/legacy-equation-jobs/job-id/request.json",
            root,
        );
        assert_eq!(
            String::from_utf8(result).unwrap(),
            "failed at <job-root>/request.json"
        );
    }

    #[test]
    fn missing_production_worker_fails_closed() {
        let temp = tempfile::tempdir().unwrap();
        assert!(select_worker(temp.path().join("missing"), None).is_err());
        let worker = temp.path().join("worker");
        std::fs::write(&worker, b"binary").unwrap();
        assert_eq!(
            select_worker(temp.path().join("missing"), Some(worker.clone())).unwrap(),
            worker
        );
    }

    #[test]
    fn worker_non_zero_exit_and_malformed_stdout_fail() {
        #[cfg(unix)]
        {
            let temp = tempfile::tempdir().unwrap();
            let crash = temp.path().join("crash");
            std::fs::write(&crash, "#!/bin/sh\nexit 9\n").unwrap();
            use std::os::unix::fs::PermissionsExt;
            std::fs::set_permissions(&crash, std::fs::Permissions::from_mode(0o700)).unwrap();
            let (_, handle) = spawn_worker(
                &crash,
                "00000000-0000-4000-8000-000000000000",
                temp.path(),
                WorkerOperation::Scan,
            )
            .unwrap();
            assert!(handle.join().unwrap().is_err());
            let malformed = temp.path().join("malformed");
            std::fs::write(&malformed, "#!/bin/sh\nprintf not-json\n").unwrap();
            std::fs::set_permissions(&malformed, std::fs::Permissions::from_mode(0o700)).unwrap();
            let (_, handle) = spawn_worker(
                &malformed,
                "00000000-0000-4000-8000-000000000000",
                temp.path(),
                WorkerOperation::Scan,
            )
            .unwrap();
            assert!(handle.join().unwrap().is_err());
        }
    }

    #[test]
    fn worker_rejects_oversized_stdout_and_stderr() {
        #[cfg(unix)]
        {
            use std::os::unix::fs::PermissionsExt;

            let temp = tempfile::tempdir().unwrap();
            let job_id = "00000000-0000-4000-8000-000000000000";
            for (name, script) in [
                (
                    "large-stdout",
                    "#!/bin/sh\n/usr/bin/head -c 70000 /dev/zero\n",
                ),
                (
                    "large-stderr",
                    "#!/bin/sh\n/usr/bin/head -c 1048600 /dev/zero >&2\nprintf '{\"protocolVersion\":1,\"jobId\":\"00000000-0000-4000-8000-000000000000\",\"status\":\"awaitingOmml\"}'\n",
                ),
            ] {
                let executable = temp.path().join(name);
                std::fs::write(&executable, script).unwrap();
                std::fs::set_permissions(
                    &executable,
                    std::fs::Permissions::from_mode(0o700),
                )
                .unwrap();
                let (_, handle) = spawn_worker(
                    &executable,
                    job_id,
                    temp.path(),
                    WorkerOperation::Scan,
                )
                .unwrap();
                assert!(handle.join().unwrap().is_err());
            }
        }
    }

    #[test]
    fn cancellation_targets_only_owned_process() {
        #[cfg(unix)]
        {
            let temp = tempfile::tempdir().unwrap();
            let sleeper = temp.path().join("sleeper");
            std::fs::write(&sleeper, "#!/bin/sh\nsleep 30\n").unwrap();
            use std::os::unix::fs::PermissionsExt;
            std::fs::set_permissions(&sleeper, std::fs::Permissions::from_mode(0o700)).unwrap();
            let (first, _) = spawn_worker(
                &sleeper,
                "00000000-0000-4000-8000-000000000001",
                temp.path(),
                WorkerOperation::Scan,
            )
            .unwrap();
            let (second, _) = spawn_worker(
                &sleeper,
                "00000000-0000-4000-8000-000000000002",
                temp.path(),
                WorkerOperation::Scan,
            )
            .unwrap();
            cancel_worker(&first).unwrap();
            assert!(second
                .child
                .lock()
                .unwrap()
                .as_mut()
                .unwrap()
                .try_wait()
                .unwrap()
                .is_none());
            cancel_worker(&second).unwrap();
        }
    }
}
