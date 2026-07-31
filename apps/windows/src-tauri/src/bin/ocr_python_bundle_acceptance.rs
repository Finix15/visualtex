#[path = "../ocr_python_bundle.rs"]
mod ocr_python_bundle;

use serde_json::Value;
use std::path::PathBuf;
use std::process::Command;

fn main() {
    if let Err(error) = run() {
        eprintln!("VisualTeX bundled OCR Python acceptance failed: {error}");
        std::process::exit(1);
    }
}

fn run() -> Result<(), String> {
    let bundle_root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("resources")
        .join("ocr-python")
        .join("windows-x64");
    let temporary = std::env::temp_dir()
        .join("VisualTeX 用户 空格路径")
        .join(format!(
            "ocr-python-acceptance-{}",
            uuid::Uuid::new_v4()
        ));
    std::fs::create_dir_all(&temporary)
        .map_err(|error| format!("Unable to create acceptance directory: {error}"))?;
    let runtime_root = temporary.join("ocr-runtime").join("python");
    let manifest =
        ocr_python_bundle::install_bundle_from_root(&bundle_root, &runtime_root)?;
    let python = runtime_root.join("python.exe");

    let probe = Command::new(&python)
        .arg("-c")
        .arg(
            "import json, platform, struct, sys; print(json.dumps({'pythonVersion': platform.python_version(), 'bits': struct.calcsize('P') * 8, 'executable': sys.executable, 'prefix': sys.prefix}))",
        )
        .env_clear()
        .env("PATH", "")
        .env("PYTHONUTF8", "1")
        .env("PYTHONIOENCODING", "utf-8")
        .output()
        .map_err(|error| format!("Unable to start bundled private Python: {error}"))?;
    if !probe.status.success() {
        return Err(format!(
            "Bundled private Python probe failed with {:?}\nstdout:\n{}\nstderr:\n{}",
            probe.status.code(),
            String::from_utf8_lossy(&probe.stdout),
            String::from_utf8_lossy(&probe.stderr)
        ));
    }
    let value: Value = serde_json::from_slice(&probe.stdout)
        .map_err(|error| format!("Bundled Python probe returned invalid JSON: {error}"))?;
    if value.get("pythonVersion").and_then(Value::as_str)
        != Some(manifest.python_version.as_str())
        || value.get("bits").and_then(Value::as_u64) != Some(64)
    {
        return Err(format!("Unexpected bundled Python probe: {value}"));
    }
    let executable = value
        .get("executable")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .replace('\\', "/")
        .to_ascii_lowercase();
    let expected = python
        .display()
        .to_string()
        .replace('\\', "/")
        .to_ascii_lowercase();
    if executable != expected {
        return Err(format!(
            "Bundled Python resolved the wrong executable. Expected {expected}, actual {executable}"
        ));
    }

    let fake_user_base = temporary.join("fake-user-base");
    let fake_python_path = temporary.join("hostile-python-path");
    std::fs::create_dir_all(&fake_python_path)
        .map_err(|error| format!("Unable to create hostile PYTHONPATH directory: {error}"))?;
    std::fs::write(
        fake_python_path.join("visualtex_pythonpath_pollution.py"),
        b"POLLUTED = True\n",
    )
    .map_err(|error| format!("Unable to create hostile PYTHONPATH package: {error}"))?;
    for relative in [
        PathBuf::from("Python312").join("site-packages"),
        PathBuf::from("Lib").join("site-packages"),
    ] {
        let site_packages = fake_user_base.join(relative);
        std::fs::create_dir_all(&site_packages)
            .map_err(|error| format!("Unable to create fake user site-packages: {error}"))?;
        std::fs::write(
            site_packages.join("visualtex_user_pollution.py"),
            b"POLLUTED = True\n",
        )
        .map_err(|error| format!("Unable to create fake user package: {error}"))?;
    }
    let isolation = Command::new(&python)
        .arg("-c")
        .arg(
            "import importlib.util, json, os, site, sys; user_site=site.getusersitepackages(); sites=[user_site] if isinstance(user_site,str) else list(user_site); norm=lambda value: os.path.normcase(os.path.abspath(value)); print(json.dumps({'userSiteEnabled': bool(site.ENABLE_USER_SITE), 'userSiteOnPath': any(norm(path)==norm(candidate) for path in sys.path for candidate in sites if candidate), 'pollutionVisible': importlib.util.find_spec('visualtex_user_pollution') is not None, 'sysPath': sys.path}))",
        )
        .env_clear()
        .env("PATH", "")
        .env("PYTHONUSERBASE", &fake_user_base)
        .env("PYTHONUTF8", "1")
        .env("PYTHONIOENCODING", "utf-8")
        .output()
        .map_err(|error| format!("Unable to run bundled Python isolation probe: {error}"))?;
    if !isolation.status.success() {
        return Err(format!(
            "Bundled Python isolation probe failed with {:?}\nstdout:\n{}\nstderr:\n{}",
            isolation.status.code(),
            String::from_utf8_lossy(&isolation.stdout),
            String::from_utf8_lossy(&isolation.stderr)
        ));
    }
    let isolation_value: Value = serde_json::from_slice(&isolation.stdout)
        .map_err(|error| format!("Bundled Python isolation probe returned invalid JSON: {error}"))?;
    if isolation_value
        .get("userSiteEnabled")
        .and_then(Value::as_bool)
        .unwrap_or(true)
        || isolation_value
            .get("userSiteOnPath")
            .and_then(Value::as_bool)
            .unwrap_or(true)
        || isolation_value
            .get("pollutionVisible")
            .and_then(Value::as_bool)
            .unwrap_or(true)
    {
        return Err(format!(
            "Bundled Python allowed user site-packages pollution: {isolation_value}"
        ));
    }

    let hostile_isolation = Command::new(&python)
        .args(["-I", "-X", "utf8", "-c"])
        .arg(
            "import importlib.util, json, os, sys; print(json.dumps({'userPollutionVisible': importlib.util.find_spec('visualtex_user_pollution') is not None, 'pythonPathPollutionVisible': importlib.util.find_spec('visualtex_pythonpath_pollution') is not None, 'sysPath': sys.path}))",
        )
        .env_clear()
        .env("PATH", &fake_python_path)
        .env("PYTHONPATH", &fake_python_path)
        .env("PYTHONHOME", temporary.join("invalid-python-home"))
        .env("PYTHONUSERBASE", &fake_user_base)
        .output()
        .map_err(|error| format!("Unable to run hostile-environment isolation probe: {error}"))?;
    if !hostile_isolation.status.success() {
        return Err(format!(
            "Hostile-environment isolation probe failed with {:?}\nstdout:\n{}\nstderr:\n{}",
            hostile_isolation.status.code(),
            String::from_utf8_lossy(&hostile_isolation.stdout),
            String::from_utf8_lossy(&hostile_isolation.stderr)
        ));
    }
    let hostile_value: Value = serde_json::from_slice(&hostile_isolation.stdout)
        .map_err(|error| format!("Hostile-environment probe returned invalid JSON: {error}"))?;
    if hostile_value
        .get("userPollutionVisible")
        .and_then(Value::as_bool)
        .unwrap_or(true)
        || hostile_value
            .get("pythonPathPollutionVisible")
            .and_then(Value::as_bool)
            .unwrap_or(true)
    {
        return Err(format!(
            "Bundled Python accepted external PYTHONPATH/PYTHONHOME/PYTHONUSERBASE pollution: {hostile_value}"
        ));
    }

    let pip = Command::new(&python)
        .args(["-I", "-X", "utf8", "-m", "pip", "--version"])
        .env_clear()
        .env("PATH", &fake_python_path)
        .env("PYTHONPATH", &fake_python_path)
        .env("PYTHONHOME", temporary.join("invalid-python-home"))
        .env("PYTHONUSERBASE", &fake_user_base)
        .output()
        .map_err(|error| format!("Unable to run bundled pip: {error}"))?;
    if !pip.status.success() {
        return Err(format!(
            "Bundled pip failed with {:?}\nstdout:\n{}\nstderr:\n{}",
            pip.status.code(),
            String::from_utf8_lossy(&pip.stdout),
            String::from_utf8_lossy(&pip.stderr)
        ));
    }
    let pip_output = String::from_utf8_lossy(&pip.stdout).replace('\\', "/");
    let expected_root = runtime_root.display().to_string().replace('\\', "/");
    if !pip_output
        .to_ascii_lowercase()
        .contains(&expected_root.to_ascii_lowercase())
    {
        return Err(format!(
            "Bundled pip resolved outside the private runtime: {pip_output}"
        ));
    }

    println!("VisualTeX bundled OCR Python acceptance passed");
    println!("python={}", manifest.python_version);
    println!("pip={}", manifest.pip_version);
    println!("runtime={}", runtime_root.display());
    println!("path_isolated=true");
    std::fs::remove_dir_all(&temporary).ok();
    Ok(())
}
