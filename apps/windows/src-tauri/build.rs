fn main() {
    #[cfg(target_os = "windows")]
    {
        let manifest = std::path::PathBuf::from(std::env::var_os("CARGO_MANIFEST_DIR").unwrap())
            .join("windows")
            .join("test-common-controls.manifest");
        println!("cargo:rerun-if-changed={}", manifest.display());
        // Rust lib-test executables do not pass through Tauri's application
        // resource embedding. The release/CI wrapper enables this only for
        // `cargo test --lib`; the desktop binary retains Tauri's own manifest.
        println!("cargo:rerun-if-env-changed=VISUALTEX_RUST_LIB_TEST");
        if std::env::var_os("VISUALTEX_RUST_LIB_TEST").as_deref()
            == Some(std::ffi::OsStr::new("1"))
        {
            println!("cargo:rustc-link-arg=/MANIFEST:EMBED");
            println!("cargo:rustc-link-arg=/MANIFESTINPUT:{}", manifest.display());
        }
    }
    tauri_build::build()
}
