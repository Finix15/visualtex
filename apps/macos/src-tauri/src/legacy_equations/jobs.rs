use super::{
    package::{
        publish_candidate, read_bounded, remove_job_directory, replace_bounded, sha256_file,
        validate_regular_docx, write_bounded, MAX_BATCH_BYTES, MAX_MANIFEST_BYTES,
    },
    report::ConversionReport,
    worker::{cancel_worker, spawn_worker, RunningWorker, WorkerOperation},
};
use serde::{Deserialize, Serialize};
use std::{
    collections::HashMap,
    fs,
    path::{Component, Path, PathBuf},
    sync::{Arc, Mutex},
    time::{Duration, SystemTime, UNIX_EPOCH},
};
use uuid::Uuid;

pub const JOB_DIRECTORY_NAME: &str = "legacy-equation-jobs";
pub const JOB_TTL: Duration = Duration::from_secs(24 * 60 * 60);

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub enum JobStatus {
    Created,
    Running,
    AwaitingOmml,
    Finalizing,
    Complete,
    Cancelled,
    Failed,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct JobRecord {
    pub protocol_version: u32,
    pub job_id: Uuid,
    pub input_path: PathBuf,
    pub output_path: PathBuf,
    pub input_sha256: String,
    pub status: JobStatus,
    pub created_at_epoch_seconds: u64,
    pub updated_at_epoch_seconds: u64,
    pub error: Option<String>,
}

#[derive(Default)]
pub struct LegacyEquationState {
    records: Mutex<HashMap<Uuid, JobRecord>>,
    workers: Mutex<HashMap<Uuid, Arc<RunningWorker>>>,
}

fn epoch_seconds() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs()
}

pub fn parse_job_id(value: &str) -> Result<Uuid, String> {
    let id =
        Uuid::parse_str(value).map_err(|_| "Job ID must be a canonical UUID v4".to_string())?;
    if id.get_version_num() != 4 || id.hyphenated().to_string() != value {
        return Err("Job ID must be a canonical UUID v4".to_string());
    }
    Ok(id)
}

pub fn job_root(app_data: &Path) -> PathBuf {
    app_data.join(JOB_DIRECTORY_NAME)
}
pub fn job_path(root: &Path, id: Uuid) -> PathBuf {
    root.join(id.hyphenated().to_string())
}

fn validated_job_path(root: &Path, id: Uuid) -> Result<PathBuf, String> {
    let root_metadata = fs::symlink_metadata(root).map_err(|error| error.to_string())?;
    if !root_metadata.file_type().is_dir() || root_metadata.file_type().is_symlink() {
        return Err("Job root is not a safe directory".to_string());
    }
    let directory = job_path(root, id);
    let metadata = fs::symlink_metadata(&directory).map_err(|error| error.to_string())?;
    if !metadata.file_type().is_dir() || metadata.file_type().is_symlink() {
        return Err("Job directory is not a safe directory".to_string());
    }
    let canonical_root = root.canonicalize().map_err(|error| error.to_string())?;
    let canonical_directory = directory
        .canonicalize()
        .map_err(|error| error.to_string())?;
    if canonical_directory.parent() != Some(canonical_root.as_path()) {
        return Err("Job directory escaped the application data root".to_string());
    }
    Ok(canonical_directory)
}

fn normalize_output(path: &Path) -> Result<PathBuf, String> {
    if !path.is_absolute()
        || path
            .components()
            .any(|part| matches!(part, Component::ParentDir))
    {
        return Err("Output path must be absolute and normalized".to_string());
    }
    if path
        .extension()
        .and_then(|value| value.to_str())
        .map(str::to_ascii_lowercase)
        .as_deref()
        != Some("docx")
    {
        return Err("Output must use the .docx extension".to_string());
    }
    if path.exists() {
        return Err("Output already exists".to_string());
    }
    let parent = path
        .parent()
        .ok_or_else(|| "Output has no parent".to_string())?;
    let canonical_parent = parent
        .canonicalize()
        .map_err(|error| format!("Unable to resolve output parent: {error}"))?;
    Ok(canonical_parent.join(
        path.file_name()
            .ok_or_else(|| "Output filename missing".to_string())?,
    ))
}

impl LegacyEquationState {
    fn persist_status(root: &Path, record: &JobRecord) -> Result<(), String> {
        replace_bounded(
            &validated_job_path(root, record.job_id)?.join("status.json"),
            &serde_json::to_vec_pretty(record).map_err(|error| error.to_string())?,
            MAX_MANIFEST_BYTES,
        )
    }

    pub fn create(&self, root: &Path, input: &Path, output: &Path) -> Result<JobRecord, String> {
        let input = validate_regular_docx(input)?;
        let output = normalize_output(output)?;
        if input == output {
            return Err("Input and output must be different files".to_string());
        }
        fs::create_dir_all(root).map_err(|error| error.to_string())?;
        let root_metadata = fs::symlink_metadata(root).map_err(|error| error.to_string())?;
        if !root_metadata.file_type().is_dir() || root_metadata.file_type().is_symlink() {
            return Err("Job root is not a safe directory".to_string());
        }
        let id = Uuid::new_v4();
        let directory = job_path(root, id);
        fs::create_dir(&directory).map_err(|error| error.to_string())?;
        let directory = validated_job_path(root, id)?;
        let now = epoch_seconds();
        let record = JobRecord {
            protocol_version: 1,
            job_id: id,
            input_path: input.clone(),
            output_path: output,
            input_sha256: sha256_file(&input)?,
            status: JobStatus::Created,
            created_at_epoch_seconds: now,
            updated_at_epoch_seconds: now,
            error: None,
        };
        let payload = serde_json::to_vec_pretty(&record).map_err(|error| error.to_string())?;
        write_bounded(
            &directory.join("request.json"),
            &payload,
            MAX_MANIFEST_BYTES,
        )?;
        Self::persist_status(root, &record)?;
        self.records
            .lock()
            .map_err(|_| "Job state lock poisoned".to_string())?
            .insert(id, record.clone());
        Ok(record)
    }

    pub fn get(&self, root: &Path, id: Uuid) -> Result<JobRecord, String> {
        let completed = self
            .workers
            .lock()
            .map_err(|_| "Worker state lock poisoned".to_string())?
            .get(&id)
            .and_then(|worker| {
                if worker.is_running() {
                    None
                } else {
                    worker.completion()
                }
            });
        if let Some(completed) = completed {
            self.workers.lock().unwrap().remove(&id);
            if let Some(record) = self.records.lock().unwrap().get_mut(&id) {
                match completed {
                    Ok(_) => {
                        record.status = JobStatus::AwaitingOmml;
                        record.error = None;
                    }
                    Err(error) => {
                        record.status = JobStatus::Failed;
                        record.error = Some(error);
                    }
                }
                record.updated_at_epoch_seconds = epoch_seconds();
                Self::persist_status(root, record)?;
            }
        }
        if let Some(record) = self
            .records
            .lock()
            .map_err(|_| "Job state lock poisoned".to_string())?
            .get(&id)
            .cloned()
        {
            return Ok(record);
        }
        let directory = validated_job_path(root, id)?;
        let status_path = directory.join("status.json");
        let manifest_path = if status_path.exists() {
            status_path
        } else {
            directory.join("request.json")
        };
        let bytes = read_bounded(&manifest_path, MAX_MANIFEST_BYTES, "job manifest")?;
        let mut record: JobRecord = serde_json::from_slice(&bytes)
            .map_err(|error| format!("Malformed job manifest: {error}"))?;
        if record.job_id != id {
            return Err("Job manifest ID mismatch".to_string());
        }
        if record.status == JobStatus::Running || record.status == JobStatus::Finalizing {
            record.status = JobStatus::Failed;
            record.error = Some("Application restarted while the worker was running".to_string());
            Self::persist_status(root, &record)?;
        }
        self.records
            .lock()
            .map_err(|_| "Job state lock poisoned".to_string())?
            .insert(id, record.clone());
        Ok(record)
    }

    pub fn start(&self, root: &Path, id: Uuid, executable: &Path) -> Result<(), String> {
        let directory = validated_job_path(root, id)?;
        let (worker, handle) = spawn_worker(
            executable,
            &id.hyphenated().to_string(),
            &directory,
            WorkerOperation::Scan,
        )?;
        let worker = Arc::new(worker);
        self.workers
            .lock()
            .map_err(|_| "Worker state lock poisoned".to_string())?
            .insert(id, worker);
        if let Some(record) = self.records.lock().unwrap().get_mut(&id) {
            record.status = JobStatus::Running;
            record.updated_at_epoch_seconds = epoch_seconds();
            Self::persist_status(root, record)?;
        }
        // Completion is reconciled from the worker's bounded status plus its file protocol.
        let _ = handle;
        Ok(())
    }

    pub fn run_finalize_worker(
        &self,
        root: &Path,
        id: Uuid,
        executable: &Path,
    ) -> Result<(), String> {
        let current = self.get(root, id)?;
        if current.status != JobStatus::AwaitingOmml {
            return Err("Job is not awaiting OMML payloads".to_string());
        }
        if self.workers.lock().unwrap().contains_key(&id) {
            return Err("Job already has a running worker".to_string());
        }
        let directory = validated_job_path(root, id)?;
        let (worker, handle) = spawn_worker(
            executable,
            &id.hyphenated().to_string(),
            &directory,
            WorkerOperation::Finalize,
        )?;
        self.workers.lock().unwrap().insert(id, Arc::new(worker));
        if let Some(record) = self.records.lock().unwrap().get_mut(&id) {
            record.status = JobStatus::Finalizing;
            record.updated_at_epoch_seconds = epoch_seconds();
            Self::persist_status(root, record)?;
        }
        let outcome = handle
            .join()
            .map_err(|_| "MathType worker thread panicked".to_string())?;
        self.workers.lock().unwrap().remove(&id);
        if let Err(error) = outcome {
            if let Some(record) = self.records.lock().unwrap().get_mut(&id) {
                record.status = JobStatus::Failed;
                record.error = Some(error.clone());
                record.updated_at_epoch_seconds = epoch_seconds();
                Self::persist_status(root, record)?;
            }
            return Err(error);
        }
        Ok(())
    }

    pub fn read_batch(&self, root: &Path, id: Uuid, index: u32) -> Result<Vec<u8>, String> {
        read_bounded(
            &validated_job_path(root, id)?.join(format!("formula-batch-{index}.json")),
            MAX_BATCH_BYTES,
            "formula batch",
        )
    }

    pub fn read_optional_report(
        &self,
        root: &Path,
        id: Uuid,
        filename: &str,
    ) -> Result<Option<serde_json::Value>, String> {
        if !matches!(filename, "scan-report.json" | "conversion-report.json") {
            return Err("Unsupported legacy-equation report".to_string());
        }
        let path = validated_job_path(root, id)?.join(filename);
        if !path.exists() {
            return Ok(None);
        }
        let bytes = read_bounded(&path, MAX_MANIFEST_BYTES, "legacy-equation report")?;
        serde_json::from_slice(&bytes)
            .map(Some)
            .map_err(|error| format!("Malformed legacy-equation report: {error}"))
    }

    pub fn report_path(&self, root: &Path, id: Uuid) -> Result<PathBuf, String> {
        let directory = validated_job_path(root, id)?;
        for filename in ["conversion-report.txt", "conversion-report.json"] {
            let path = directory.join(filename);
            if path.exists() {
                read_bounded(&path, MAX_MANIFEST_BYTES, "conversion report")?;
                return Ok(path);
            }
        }
        Err("Conversion report is not available".to_string())
    }

    pub fn submit_batch(
        &self,
        root: &Path,
        id: Uuid,
        index: u32,
        payload: &[u8],
    ) -> Result<(), String> {
        serde_json::from_slice::<serde_json::Value>(payload)
            .map_err(|error| format!("Malformed OMML batch: {error}"))?;
        write_bounded(
            &validated_job_path(root, id)?.join(format!("omml-batch-{index}.json")),
            payload,
            MAX_BATCH_BYTES,
        )
    }

    pub fn finalize(&self, root: &Path, id: Uuid) -> Result<JobRecord, String> {
        let mut record = self.get(root, id)?;
        if record.output_path.exists() {
            return Err("Output already exists".to_string());
        }
        if sha256_file(&record.input_path)? != record.input_sha256 {
            return Err("Source document hash changed".to_string());
        }
        let directory = validated_job_path(root, id)?;
        let report_bytes = read_bounded(
            &directory.join("conversion-report.json"),
            MAX_MANIFEST_BYTES,
            "conversion report",
        )?;
        let report: ConversionReport = serde_json::from_slice(&report_bytes)
            .map_err(|error| format!("Malformed conversion report: {error}"))?;
        report.validate(&record.input_sha256)?;
        let candidate = directory.join("candidate.docx");
        if sha256_file(&candidate)? != report.output_sha256 {
            return Err("Candidate SHA-256 does not match report".to_string());
        }
        record.status = JobStatus::Finalizing;
        publish_candidate(&candidate, &record.output_path)?;
        if sha256_file(&record.input_path)? != record.input_sha256 {
            return Err("Source document changed during publish".to_string());
        }
        record.status = JobStatus::Complete;
        record.updated_at_epoch_seconds = epoch_seconds();
        Self::persist_status(root, &record)?;
        self.records.lock().unwrap().insert(id, record.clone());
        Ok(record)
    }

    pub fn cancel(&self, root: &Path, id: Uuid) -> Result<JobRecord, String> {
        if let Some(worker) = self
            .workers
            .lock()
            .map_err(|_| "Worker state lock poisoned".to_string())?
            .remove(&id)
        {
            cancel_worker(&worker)?;
        }
        let mut records = self
            .records
            .lock()
            .map_err(|_| "Job state lock poisoned".to_string())?;
        let record = records
            .get_mut(&id)
            .ok_or_else(|| "Unknown job".to_string())?;
        record.status = JobStatus::Cancelled;
        record.updated_at_epoch_seconds = epoch_seconds();
        Self::persist_status(root, record)?;
        Ok(record.clone())
    }

    pub fn delete(&self, root: &Path, id: Uuid) -> Result<(), String> {
        if self.workers.lock().unwrap().contains_key(&id) {
            return Err("Cannot delete a running job".to_string());
        }
        remove_job_directory(&validated_job_path(root, id)?)?;
        self.records.lock().unwrap().remove(&id);
        Ok(())
    }

    pub fn cleanup_expired(&self, root: &Path, now: u64) -> Result<usize, String> {
        if !root.exists() {
            return Ok(0);
        }
        let mut removed = 0;
        for entry in fs::read_dir(root).map_err(|error| error.to_string())? {
            let entry = entry.map_err(|error| error.to_string())?;
            let Some(name) = entry.file_name().to_str().map(str::to_owned) else {
                continue;
            };
            let Ok(id) = parse_job_id(&name) else {
                continue;
            };
            let Ok(record) = self.get(root, id) else {
                continue;
            };
            if now.saturating_sub(record.updated_at_epoch_seconds) >= JOB_TTL.as_secs()
                && !self.workers.lock().unwrap().contains_key(&id)
                && remove_job_directory(&entry.path()).is_ok()
            {
                self.records.lock().unwrap().remove(&id);
                removed += 1;
            }
        }
        Ok(removed)
    }

    pub fn cleanup_expired_now(&self, root: &Path) -> Result<usize, String> {
        self.cleanup_expired(root, epoch_seconds())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fixture() -> (tempfile::TempDir, LegacyEquationState, PathBuf, PathBuf) {
        let temp = tempfile::tempdir().unwrap();
        let input = temp.path().join("input.docx");
        fs::write(&input, b"source").unwrap();
        let output = temp.path().join("output.docx");
        (temp, LegacyEquationState::default(), input, output)
    }

    #[test]
    fn requires_canonical_v4_uuid() {
        assert!(parse_job_id("not-a-uuid").is_err());
        assert!(parse_job_id("00000000-0000-1000-8000-000000000000").is_err());
        assert!(parse_job_id("AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA").is_err());
        assert!(parse_job_id("aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa").is_ok());
    }

    #[test]
    fn rejects_same_input_output_and_existing_output() {
        let (temp, state, input, output) = fixture();
        let root = job_root(temp.path());
        assert!(state.create(&root, &input, &input).is_err());
        fs::write(&output, b"existing").unwrap();
        assert!(state.create(&root, &input, &output).is_err());
    }

    #[cfg(unix)]
    #[test]
    fn output_symlink_swap_cannot_redirect_job() {
        use std::os::unix::fs::symlink;

        let (temp, state, input, _) = fixture();
        let first = temp.path().join("first");
        let second = temp.path().join("second");
        fs::create_dir(&first).unwrap();
        fs::create_dir(&second).unwrap();
        let link = temp.path().join("output-link");
        symlink(&first, &link).unwrap();
        let root = job_root(temp.path());
        let record = state
            .create(&root, &input, &link.join("result.docx"))
            .unwrap();
        fs::remove_file(&link).unwrap();
        symlink(&second, &link).unwrap();

        assert_eq!(
            record.output_path,
            first.canonicalize().unwrap().join("result.docx")
        );
        assert!(!record.output_path.starts_with(&second));
    }

    #[cfg(unix)]
    #[test]
    fn job_directory_symlink_swap_is_rejected() {
        use std::os::unix::fs::symlink;

        let (temp, state, input, output) = fixture();
        let root = job_root(temp.path());
        let record = state.create(&root, &input, &output).unwrap();
        let directory = job_path(&root, record.job_id);
        let saved = temp.path().join("saved-job");
        fs::rename(&directory, &saved).unwrap();
        let outside = temp.path().join("outside");
        fs::create_dir(&outside).unwrap();
        symlink(&outside, &directory).unwrap();

        assert!(state
            .submit_batch(&root, record.job_id, 0, br#"{"unsafe":true}"#)
            .is_err());
        assert!(!outside.join("omml-batch-0.json").exists());
    }

    #[test]
    fn concurrent_jobs_keep_distinct_directories_and_payloads() {
        let (temp, state, input, output) = fixture();
        let root = job_root(temp.path());
        let first = state.create(&root, &input, &output).unwrap();
        let second = state
            .create(&root, &input, &temp.path().join("two.docx"))
            .unwrap();
        state
            .submit_batch(&root, first.job_id, 0, br#"{"job":"one"}"#)
            .unwrap();
        state
            .submit_batch(&root, second.job_id, 0, br#"{"job":"two"}"#)
            .unwrap();
        assert_ne!(
            fs::read(job_path(&root, first.job_id).join("omml-batch-0.json")).unwrap(),
            fs::read(job_path(&root, second.job_id).join("omml-batch-0.json")).unwrap()
        );
    }

    #[test]
    fn restart_marks_running_job_failed() {
        let (temp, state, input, output) = fixture();
        let root = job_root(temp.path());
        let mut record = state.create(&root, &input, &output).unwrap();
        record.status = JobStatus::Running;
        LegacyEquationState::persist_status(&root, &record).unwrap();
        let recovered = LegacyEquationState::default()
            .get(&root, record.job_id)
            .unwrap();
        assert_eq!(recovered.status, JobStatus::Failed);
    }

    #[test]
    fn expired_allowlisted_job_is_removed() {
        let (temp, state, input, output) = fixture();
        let root = job_root(temp.path());
        let mut record = state.create(&root, &input, &output).unwrap();
        record.updated_at_epoch_seconds = 1;
        LegacyEquationState::persist_status(&root, &record).unwrap();
        let fresh = LegacyEquationState::default();
        assert_eq!(
            fresh.cleanup_expired(&root, JOB_TTL.as_secs() + 2).unwrap(),
            1
        );
    }

    #[test]
    fn source_mutation_blocks_finalize_before_publish() {
        let (temp, state, input, output) = fixture();
        let root = job_root(temp.path());
        let record = state.create(&root, &input, &output).unwrap();
        fs::write(&input, b"changed").unwrap();
        assert!(state
            .finalize(&root, record.job_id)
            .unwrap_err()
            .contains("Source document hash changed"));
        assert!(!output.exists());
    }

    #[test]
    fn oversized_or_malformed_batch_is_rejected() {
        let (temp, state, input, output) = fixture();
        let root = job_root(temp.path());
        let record = state.create(&root, &input, &output).unwrap();
        assert!(state
            .submit_batch(&root, record.job_id, 0, b"not-json")
            .is_err());
        assert!(state
            .submit_batch(
                &root,
                record.job_id,
                1,
                &vec![b' '; MAX_BATCH_BYTES as usize + 1],
            )
            .is_err());
    }

    #[test]
    fn optional_ui_reports_are_bounded_and_must_be_json() {
        let (temp, state, input, output) = fixture();
        let root = job_root(temp.path());
        let record = state.create(&root, &input, &output).unwrap();
        let path = job_path(&root, record.job_id).join("scan-report.json");
        fs::write(&path, b"not-json").unwrap();
        assert!(state
            .read_optional_report(&root, record.job_id, "scan-report.json")
            .is_err());
        fs::write(&path, vec![0_u8; MAX_MANIFEST_BYTES as usize + 1]).unwrap();
        assert!(state
            .read_optional_report(&root, record.job_id, "scan-report.json")
            .is_err());
        assert!(state
            .read_optional_report(&root, record.job_id, "../request.json")
            .is_err());
    }
}
