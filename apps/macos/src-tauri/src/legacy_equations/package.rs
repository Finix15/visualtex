use sha2::{Digest, Sha256};
use std::{
    fs::{self, File, OpenOptions},
    io::{Read, Write},
    path::{Component, Path, PathBuf},
};

pub const MAX_MANIFEST_BYTES: u64 = 4 * 1024 * 1024;
pub const MAX_BATCH_BYTES: u64 = 16 * 1024 * 1024;
pub const ALLOWED_JOB_FILES: &[&str] = &[
    "request.json",
    "status.json",
    "scan-report.json",
    "conversion-report.json",
    "conversion-report.txt",
    "candidate.docx",
    "worker.stderr.log",
];

pub fn sha256_file(path: &Path) -> Result<String, String> {
    let mut file =
        File::open(path).map_err(|error| format!("Unable to read {}: {error}", path.display()))?;
    let mut digest = Sha256::new();
    let mut buffer = [0_u8; 64 * 1024];
    loop {
        let count = file.read(&mut buffer).map_err(|error| error.to_string())?;
        if count == 0 {
            break;
        }
        digest.update(&buffer[..count]);
    }
    Ok(hex::encode(digest.finalize()))
}

pub fn validate_regular_docx(path: &Path) -> Result<PathBuf, String> {
    if !path.is_absolute()
        || path
            .components()
            .any(|part| matches!(part, Component::ParentDir))
    {
        return Err("DOCX path must be absolute and normalized".to_string());
    }
    let metadata = fs::symlink_metadata(path)
        .map_err(|error| format!("Unable to inspect {}: {error}", path.display()))?;
    if !metadata.file_type().is_file() || metadata.file_type().is_symlink() {
        return Err("DOCX input must be a regular non-symlink file".to_string());
    }
    if path
        .extension()
        .and_then(|value| value.to_str())
        .map(str::to_ascii_lowercase)
        .as_deref()
        != Some("docx")
    {
        return Err("Only .docx input is supported".to_string());
    }
    path.canonicalize().map_err(|error| error.to_string())
}

pub fn read_bounded(path: &Path, maximum: u64, label: &str) -> Result<Vec<u8>, String> {
    let metadata = fs::symlink_metadata(path)
        .map_err(|error| format!("Unable to inspect {label}: {error}"))?;
    if !metadata.file_type().is_file()
        || metadata.file_type().is_symlink()
        || metadata.len() > maximum
    {
        return Err(format!("{label} is missing, unsafe, or oversized"));
    }
    fs::read(path).map_err(|error| format!("Unable to read {label}: {error}"))
}

fn write_synced(path: &Path, bytes: &[u8], create_new: bool) -> Result<(), String> {
    let mut options = OpenOptions::new();
    options.write(true);
    if create_new {
        options.create_new(true);
    } else {
        options.create(true).truncate(true);
    }
    let mut file = options.open(path).map_err(|error| error.to_string())?;
    file.write_all(bytes).map_err(|error| error.to_string())?;
    file.sync_all().map_err(|error| error.to_string())
}

pub fn write_bounded(path: &Path, bytes: &[u8], maximum: u64) -> Result<(), String> {
    if bytes.len() as u64 > maximum {
        return Err("Payload exceeds configured limit".to_string());
    }
    let parent = path
        .parent()
        .ok_or_else(|| "Output has no parent".to_string())?;
    fs::create_dir_all(parent).map_err(|error| error.to_string())?;
    if path.exists() {
        return Err("Destination already exists".to_string());
    }
    let temporary = parent.join(format!(
        ".{}.tmp",
        path.file_name()
            .and_then(|x| x.to_str())
            .unwrap_or("payload")
    ));
    write_synced(&temporary, bytes, true)?;
    fs::rename(&temporary, path).map_err(|error| error.to_string())?;
    if let Err(error) = File::open(parent).and_then(|directory| directory.sync_all()) {
        let _ = fs::remove_file(path);
        return Err(error.to_string());
    }
    Ok(())
}

pub fn replace_bounded(path: &Path, bytes: &[u8], maximum: u64) -> Result<(), String> {
    if bytes.len() as u64 > maximum {
        return Err("Payload exceeds configured limit".to_string());
    }
    let parent = path
        .parent()
        .ok_or_else(|| "Output has no parent".to_string())?;
    fs::create_dir_all(parent).map_err(|error| error.to_string())?;
    let temporary = parent.join(format!(
        ".{}.replace",
        path.file_name()
            .and_then(|x| x.to_str())
            .unwrap_or("payload")
    ));
    write_synced(&temporary, bytes, false)?;
    fs::rename(&temporary, path).map_err(|error| error.to_string())?;
    File::open(parent)
        .and_then(|directory| directory.sync_all())
        .map_err(|error| error.to_string())
}

pub fn publish_candidate(candidate: &Path, output: &Path) -> Result<(), String> {
    publish_candidate_with(
        candidate,
        output,
        |source, destination| fs::hard_link(source, destination),
        |parent| File::open(parent).and_then(|directory| directory.sync_all()),
    )
}

fn publish_candidate_with<L, S>(
    candidate: &Path,
    output: &Path,
    link: L,
    sync_parent: S,
) -> Result<(), String>
where
    L: FnOnce(&Path, &Path) -> std::io::Result<()>,
    S: FnOnce(&Path) -> std::io::Result<()>,
{
    if output.exists() {
        return Err("Output already exists".to_string());
    }
    let parent = output
        .parent()
        .ok_or_else(|| "Output has no parent".to_string())?;
    fs::create_dir_all(parent).map_err(|error| error.to_string())?;
    let metadata = fs::symlink_metadata(candidate).map_err(|error| error.to_string())?;
    if !metadata.file_type().is_file() || metadata.file_type().is_symlink() {
        return Err("Candidate must be a regular non-symlink file".to_string());
    }
    link(candidate, output).map_err(|error| format!("Atomic publish failed: {error}"))?;
    if let Err(error) = sync_parent(parent) {
        let _ = fs::remove_file(output);
        return Err(error.to_string());
    }
    Ok(())
}

pub fn job_directory_is_allowlisted(path: &Path) -> Result<bool, String> {
    for entry in fs::read_dir(path).map_err(|error| error.to_string())? {
        let entry = entry.map_err(|error| error.to_string())?;
        let metadata = fs::symlink_metadata(entry.path()).map_err(|error| error.to_string())?;
        if !metadata.file_type().is_file() || metadata.file_type().is_symlink() {
            return Ok(false);
        }
        let name = entry.file_name();
        let name = name
            .to_str()
            .ok_or_else(|| "Non-UTF-8 job filename".to_string())?;
        let batch = (name.starts_with("formula-batch-") || name.starts_with("omml-batch-"))
            && name.ends_with(".json");
        if !batch && !ALLOWED_JOB_FILES.contains(&name) {
            return Ok(false);
        }
    }
    Ok(true)
}

pub fn remove_job_directory(path: &Path) -> Result<(), String> {
    if !job_directory_is_allowlisted(path)? {
        return Err("Job directory contains an unknown or unsafe file".to_string());
    }
    for entry in fs::read_dir(path).map_err(|error| error.to_string())? {
        fs::remove_file(entry.map_err(|error| error.to_string())?.path())
            .map_err(|error| error.to_string())?;
    }
    fs::remove_dir(path).map_err(|error| error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn rejects_traversal_and_symlink_input() {
        assert!(validate_regular_docx(Path::new("../x.docx")).is_err());
        let temp = tempfile::tempdir().unwrap();
        let source = temp.path().join("source.docx");
        fs::write(&source, b"x").unwrap();
        #[cfg(unix)]
        {
            std::os::unix::fs::symlink(&source, temp.path().join("link.docx")).unwrap();
            assert!(validate_regular_docx(&temp.path().join("link.docx")).is_err());
        }
    }
    #[test]
    fn output_collision_and_write_failure_fail_closed() {
        let temp = tempfile::tempdir().unwrap();
        let candidate = temp.path().join("candidate.docx");
        let output = temp.path().join("output.docx");
        fs::write(&candidate, b"candidate").unwrap();
        fs::write(&output, b"existing").unwrap();
        assert!(publish_candidate(&candidate, &output).is_err());
        assert_eq!(fs::read(output).unwrap(), b"existing");
        assert!(write_bounded(&temp.path().join("large"), &[0; 5], 4).is_err());
        let regular_file = temp.path().join("not-a-directory");
        fs::write(&regular_file, b"x").unwrap();
        assert!(write_bounded(&regular_file.join("payload"), b"x", 4).is_err());
    }
    #[test]
    fn unknown_file_prevents_cleanup() {
        let temp = tempfile::tempdir().unwrap();
        fs::write(temp.path().join("request.json"), b"{}").unwrap();
        fs::write(temp.path().join("unknown"), b"keep").unwrap();
        assert!(remove_job_directory(temp.path()).is_err());
        assert!(temp.path().join("unknown").exists());
    }

    #[test]
    fn oversized_manifest_and_batch_are_rejected() {
        let temp = tempfile::tempdir().unwrap();
        let manifest = temp.path().join("request.json");
        fs::write(&manifest, vec![0_u8; 5]).unwrap();
        assert!(read_bounded(&manifest, 4, "manifest").is_err());
        assert!(write_bounded(&temp.path().join("batch.json"), &[0_u8; 5], 4).is_err());
    }

    #[test]
    fn publish_link_and_fsync_failures_do_not_leave_output() {
        let temp = tempfile::tempdir().unwrap();
        let candidate = temp.path().join("candidate.docx");
        let output = temp.path().join("output.docx");
        fs::write(&candidate, b"candidate").unwrap();
        assert!(publish_candidate_with(
            &candidate,
            &output,
            |_, _| Err(std::io::Error::other("injected publish failure")),
            |_| Ok(())
        )
        .is_err());
        assert!(!output.exists());
        assert!(publish_candidate_with(
            &candidate,
            &output,
            |source, destination| fs::hard_link(source, destination),
            |_| Err(std::io::Error::other("injected fsync failure")),
        )
        .is_err());
        assert!(!output.exists());
    }
}
