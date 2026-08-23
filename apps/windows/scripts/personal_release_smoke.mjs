import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

const read = (path) => readFileSync(path, "utf8");
const tauri = JSON.parse(read("src-tauri/tauri.conf.json"));
const capability = read("src-tauri/capabilities/default.json");
const updater = read("src/update/updateService.ts");
const updateDialog = read("src/components/UpdateDialog.tsx");
const certificateRemoval = read("scripts/remove_windows_office_certificate.ps1");
const officeUninstall = read("scripts/uninstall_windows_vsto.ps1");
const buildScript = read("scripts/tauri_build.mjs");
const artifactVerifier = read("scripts/verify_windows_release_artifacts.ps1");
const rustBuild = read("src-tauri/build.rs");
const testManifest = read("src-tauri/windows/test-common-controls.manifest");
const workflow = read("../../.github/workflows/windows.yml");
const macUpdater = read("../macos/src/update/updateService.ts");
const macCapability = read("../macos/src-tauri/capabilities/default.json");

assert.match(updater, /Finix15\/visualtex\/releases\/latest/);
assert.doesNotMatch(updater, /paulhe666\/visualtex/);
assert.match(updateDialog, /github\.com\/Finix15\/visualtex/);
assert.doesNotMatch(updateDialog, /github\.com\/paulhe666\/visualtex/);
assert.match(capability, /github\.com\/Finix15\/visualtex/);
assert.doesNotMatch(capability, /github\.com\/paulhe666\/visualtex/);
assert.match(macUpdater, /Finix15\/visualtex\/releases\/latest/);
assert.doesNotMatch(macUpdater, /paulhe666\/visualtex/);
assert.match(macCapability, /github\.com\/Finix15\/visualtex/);

const csp = tauri.app.security.csp;
const devCsp = tauri.app.security.devCsp;
assert.equal(typeof csp, "string");
assert.match(csp, /default-src 'self'/);
assert.match(csp, /connect-src[^;]*https:\/\/api\.github\.com/);
assert.match(csp, /object-src 'none'/);
assert.match(csp, /frame-ancestors 'none'/);
assert.doesNotMatch(csp, /unsafe-eval/);
assert.match(devCsp, /http:\/\/localhost:1420/);
assert.match(devCsp, /ws:\/\/localhost:1420/);

assert.match(certificateRemoval, /CertificateCleanupPending/);
assert.match(certificateRemoval, /CurrentUser\\\$storeName\\\$thumbprint/);
assert.match(certificateRemoval, /CN=VisualTeX Local Office Companion/);
assert.match(officeUninstall, /certificate cleanup failed\. Uninstall will continue/);
assert.doesNotMatch(officeUninstall, /throw "VisualTeX Office certificate cleanup failed/);

assert.match(buildScript, /\.sha256/);
assert.match(artifactVerifier, /exact installer SHA-256 sidecar verified/);
assert.match(rustBuild, /VISUALTEX_RUST_LIB_TEST/);
assert.match(rustBuild, /rustc-link-arg=\/MANIFESTINPUT/);
assert.match(testManifest, /Microsoft\.Windows\.Common-Controls/);
assert.match(workflow, /cargo test --manifest-path src-tauri\/Cargo\.toml --lib/);
assert.match(workflow, /VISUALTEX_RUST_LIB_TEST: "1"/);
assert.doesNotMatch(workflow, /cargo test --manifest-path src-tauri\/Cargo\.toml --lib --no-run/);

console.log("Personal release channel, CSP, certificate, checksum and Rust-test gates passed.");
