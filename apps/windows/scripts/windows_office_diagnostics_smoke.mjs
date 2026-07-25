import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

async function source(path) {
  return (await readFile(path, "utf8")).replace(/\r\n?/g, "\n");
}

const certificate = await source("scripts/ensure_windows_office_certificate.ps1");
const installer = await source("scripts/install_windows_vsto.ps1");
const runtime = await source("scripts/test_windows_office_runtime.ps1");
const sessionClient = await source(
  "src-windows/VisualTeX.WindowsOffice.Contracts/VisualTeXSessionClient.cs",
);
const lifecycle = await source("src-tauri/src/office/lifecycle.rs");
const server = await source("src-tauri/src/office/server.rs");
const state = await source("src-tauri/src/office/state.rs");
const backend = await source("src-tauri/src/office/windows_backend.rs");
const platform = await source("src-tauri/src/office/platform/mod.rs");
const hooks = await source("src-tauri/windows/hooks.nsh");
const settings = await source("src/components/WindowsOfficeIntegrationSettings.tsx");
const bundle = await source("src-tauri/tauri.windows.conf.json");

for (const required of [
  "Split-Path -Parent $PSScriptRoot",
  "[string]$VisualTeXPath",
  'Get-IntegrationValue "ExecutablePath"',
  "Attempted paths:",
  "ResolveVisualTeXPathOnly",
  'Name "ExecutablePath"',
  'Name "AppDataRoot"',
  'Name "CertificatePath"',
  'Name "CertificateThumbprint"',
  'Name "CompanionPort"',
  'Name "ProtocolVersion"',
]) {
  assert.ok(certificate.includes(required), `certificate script missing ${required}`);
}
assert.ok(!certificate.includes('Programs\\VisualTeX\\VisualTeX.exe'));
assert.ok(!certificate.includes('VisualTeX\\visualtex.exe'));

const destructiveMarkers = [
  "Remove-LegacyOfficeJsState",
  "$oldProducts = @(Get-RelatedProductCodes)",
  'Invoke-MsiExec @("/x"',
];
const preflightMarkers = [
  "Assert-NoOfficeProcesses",
  "Resolve-OfficePlatform",
  "Assert-OfficeApplicationsInstalled",
  "Assert-VstoRuntimeInstalled",
  "Assert-NetFramework48Installed",
  "Assert-CurrentUserCertificateTrusted",
  "Assert-SharedCompanionConfiguration",
  "Assert-OfficeUiResources",
  "Assert-MsiArchitecture",
  "MSI SHA-256 mismatch",
];
const firstDestructive = Math.min(
  ...destructiveMarkers.map((marker) => {
    const index = installer.lastIndexOf(marker);
    assert.ok(index >= 0, `installer missing destructive marker ${marker}`);
    return index;
  }),
);
for (const marker of preflightMarkers) {
  const index = installer.lastIndexOf(marker, firstDestructive);
  assert.ok(index >= 0, `installer preflight missing/bad order: ${marker}`);
  assert.ok(index < firstDestructive, `installer runs ${marker} after destructive changes`);
}
for (const required of [
  "Static installation verification",
  "Companion runtime verification",
  "-CompanionOnly",
  "installed and runtime-verified",
  "Office add-ins installed successfully, but local companion runtime verification failed",
  "$script:staticInstallVerified",
  "$script:runtimeVerified",
  "Write-DiagnosticReport $false $script:staticInstallVerified $script:runtimeVerified",
]) {
  assert.ok(installer.includes(required), `installer missing ${required}`);
}
assert.ok(!/devenv|Visual Studio IDE/i.test(installer));
assert.ok(!installer.includes("uninstall_windows_ole.ps1"));

for (const required of [
  "$handler.UseProxy = $false",
  "$handler.Proxy = $null",
  "Get-PortOwner",
  "Get-ProcessIdentity",
  "PathRetryMilliseconds = 3000",
  "Get-CimInstance Win32_Process",
  "Get-MsiInstalledStateOnce",
  "WindowsInstaller.RelatedProducts",
  "Test-MsiInstalled([int]$TimeoutSeconds = 10)",
  "portOwnerPidMatchesVerifiedProcess",
  "startedProcessExitCode",
  "tlsPolicyErrors",
  "healthRaw",
  "exceptionChain",
  "installJsonPath",
  "CompanionOnly",
  "CompanionPort",
  "ProtocolVersion",
  "Certificate mismatch",
  'stage = "port-conflict"',
  "Port $port is already occupied by",
  "Port $port is owned by",
  "startup.log",
  "companion.log",
]) {
  assert.ok(runtime.includes(required), `runtime verifier missing ${required}`);
}

for (const required of [
  "UseProxy = false",
  "Proxy = null",
  "CompanionHealthDiagnostic",
  "VisualTeXCompanionException",
  'Stage = "configuration"',
  'diagnostic.Stage = "install-json"',
  'diagnostic.Stage = "port-listen"',
  'diagnostic.Stage = "https-handshake"',
  'diagnostic.Stage = "certificate-match"',
  'diagnostic.Stage = "health-json"',
  'diagnostic.Stage = "protocol-version"',
  "StartedProcessExitCode",
  "PortOwnerProcessId",
  "PortOwnerProcessPath",
  "PortOccupiedByOtherProcess",
  "TlsPolicyErrors",
  "HealthResponse",
  "ServerCertificateThumbprint",
  "QueryIntegrationRegistry",
  '"reg.exe"',
  '"--office-background"',
  "ProcessExited",
  "ProtocolVersionMismatch",
  "ServerCertificateMismatch",
  "InstallJsonInvalid",
]) {
  assert.ok(sessionClient.includes(required), `session client missing ${required}`);
}
assert.ok(!sessionClient.includes("TryReadHealthyAsync"));
assert.ok(!sessionClient.includes("Local companion did not become healthy"));
assert.ok(!sessionClient.includes('"Programs",\n                "VisualTeX"'));

for (const required of [
  "write_shared_configuration",
  "latest_bootstrap_log_tail",
  "stdout:",
  "stderr:",
  "open_windows_office_logs",
]) {
  assert.ok(lifecycle.includes(required), `lifecycle missing ${required}`);
}
for (const required of [
  '"ExecutablePath"',
  '"AppDataRoot"',
  '"CertificatePath"',
  '"CertificateThumbprint"',
  '"CompanionPort"',
  '"ProtocolVersion"',
]) {
  assert.ok(backend.includes(required), `shared registry writer missing ${required}`);
}
assert.ok(lifecycle.includes('"startup.log"'));
for (const required of [
  '"companion.log"',
  "binding TCP listener",
  "TLS certificate and private key loaded successfully",
  "Office UI resource validated",
  "companion stopped with error",
]) {
  assert.ok(server.includes(required), `server logging missing ${required}`);
}
assert.ok(state.includes('join("VisualTeX")'));
assert.ok(state.includes('join("office")'));
assert.ok(state.includes('join("logs")'));

for (const required of [
  "word_files_present",
  "word_registry_complete",
  "word_load_enabled",
  "powerpoint_files_present",
  "powerpoint_registry_complete",
  "powerpoint_load_enabled",
  "word_connected",
  "powerpoint_connected",
  "companion_process_running",
  "companion_port_listening",
  "companion_https_healthy",
  "companion_certificate_matches",
  "companion_protocol_matches",
  "ole_local_server_healthy",
]) {
  assert.ok(platform.includes(required), `platform status missing ${required}`);
}
assert.ok(backend.includes("registered_addin_file_exists"));
assert.ok(backend.includes("addin_registry_complete"));
assert.ok(!backend.includes("OLE_CATALOG_KEY"));
assert.ok(!backend.includes("register_ole_catalog"));
assert.ok(!backend.includes("office_catalog_path"));

assert.ok(hooks.includes('-VisualTeXPath "$INSTDIR\\VisualTeX.exe"'));
assert.ok(hooks.includes("visualtex_office_static_runtime_verified"));
assert.ok(hooks.includes("COMAddIn.Connect"));
assert.ok(settings.includes("wordFilesPresent"));
assert.ok(settings.includes("wordRegistryComplete"));
assert.ok(settings.includes("wordLoadEnabled"));
assert.ok(settings.includes("companionPortListening"));
assert.ok(settings.includes("companionHttpsHealthy"));
assert.ok(settings.includes("companionCertificateMatches"));
assert.ok(settings.includes("companionProtocolMatches"));
assert.ok(settings.includes("open_windows_office_logs"));
assert.ok(settings.includes("office-settings-diagnostic"));

assert.ok(!bundle.includes("TrustedCatalog"));
assert.ok(!bundle.includes("OfficeCatalog"));
assert.ok(!bundle.includes("office-manifests"));
assert.ok(!bundle.includes("uninstall_windows_ole.ps1"));

console.log("Windows Office staged installation and diagnostic coverage passed.");
