import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import {
  DEFAULT_ONBOARDING_STORAGE_KEY,
  WINDOWS_DESKTOP_ONBOARDING_STORAGE_KEY,
  detectDesktopPlatformFrom,
  onboardingStorageKey,
  shouldOpenOnboardingInitially,
} from "../src/platform.ts";
import { tutorialSteps } from "../src/components/OnboardingTour.tsx";
import {
  VISUALTEX_QQ_GROUP_NUMBER,
  VISUALTEX_QQ_GROUP_QR_DATA_URL,
} from "../src/assets/visualtexQqGroup.ts";

assert.equal(detectDesktopPlatformFrom("Win32", "Mozilla/5.0 (Windows NT 10.0)"), "windows");
assert.equal(onboardingStorageKey("windows", true), WINDOWS_DESKTOP_ONBOARDING_STORAGE_KEY);
assert.equal(onboardingStorageKey("windows", false), DEFAULT_ONBOARDING_STORAGE_KEY);
assert.equal(shouldOpenOnboardingInitially(false, false), true);
assert.equal(shouldOpenOnboardingInitially(true, false), false);

const windowsSteps = tutorialSteps("cn", "windows");
const windowsIds = windowsSteps.map((step) => step.id);
assert(windowsIds.includes("windows-office-manage"));
assert(!windowsIds.some((id) => id.startsWith("mac-")));
const windowsOfficeStep = windowsSteps.find((step) => step.id === "windows-office-manage");
assert(windowsOfficeStep?.title.includes("Word 和 PowerPoint"));
assert(windowsOfficeStep?.description.includes("OLE 或 OMML"));
assert(windowsOfficeStep?.description.includes("公式编号和插入引用"));
assert(windowsOfficeStep?.description.includes("转为原生 OLE"));

const matrixFontsStep = windowsSteps.find((step) => step.id === "matrix-fonts");
const inputBehaviorStep = windowsSteps.find((step) => step.id === "input-behavior");
const exportStep = windowsSteps.find((step) => step.id === "export");
assert(matrixFontsStep?.description.includes("10×10"));
assert(matrixFontsStep?.description.includes("黑板粗体"));
assert(inputBehaviorStep?.description.includes("上标与下标"));
assert(inputBehaviorStep?.description.includes("按 Enter 结束"));
assert(inputBehaviorStep?.description.includes("微分 d"));
assert(exportStep?.description.includes("Markdown、SVG 或 PNG"));
assert(exportStep?.description.includes("自选路径"));

assert.equal(VISUALTEX_QQ_GROUP_NUMBER, "1045801770");
assert(VISUALTEX_QQ_GROUP_QR_DATA_URL.startsWith("data:image/png;base64,"));
assert(Buffer.from(VISUALTEX_QQ_GROUP_QR_DATA_URL.split(",")[1], "base64").length > 10_000);

const updateDialogSource = await readFile("src/components/UpdateDialog.tsx", "utf8");
const stylesSource = await readFile("src/styles.css", "utf8");
const windowsSettingsSource = await readFile("src/components/WindowsOfficeIntegrationSettings.tsx", "utf8");
const mainSource = await readFile("src-tauri/src/main.rs", "utf8");
const lifecycleSource = await readFile("src-tauri/src/office/lifecycle.rs", "utf8");
const windowsBackendSource = await readFile("src-tauri/src/office/windows_backend.rs", "utf8");
const hooksSource = await readFile("src-tauri/windows/hooks.nsh", "utf8");
const installOleSource = await readFile("scripts/install_windows_ole.ps1", "utf8");
const installVstoSource = await readFile("scripts/install_windows_vsto.ps1", "utf8");
const runtimeTestSource = await readFile("scripts/test_windows_office_runtime.ps1", "utf8");
const uninstallVstoSource = await readFile("scripts/uninstall_windows_vsto.ps1", "utf8");
const removeCertificateSource = await readFile("scripts/remove_windows_office_certificate.ps1", "utf8");
const windowsBundleSource = await readFile("src-tauri/tauri.windows.conf.json", "utf8");
const certificateSource = await readFile("scripts/ensure_windows_office_certificate.ps1", "utf8");

assert(updateDialogSource.includes("update-community-card"));
assert(updateDialogSource.includes("VISUALTEX_QQ_GROUP_QR_DATA_URL"));
assert(updateDialogSource.includes("VISUALTEX_QQ_GROUP_NUMBER"));
assert(stylesSource.includes(".update-community-qr img"));
assert(stylesSource.includes(".onboarding-input-behavior-demo"));
assert(stylesSource.includes(".onboarding-export-demo"));
assert(windowsSettingsSource.includes('"set_office_background_start"'));
assert(mainSource.includes('#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]'));
assert(lifecycleSource.includes("pub fn set_office_background_start"));
assert(lifecycleSource.includes("set_background_start_enabled(enabled)"));
assert(lifecycleSource.includes("powershell_compatible_path"));
assert(lifecycleSource.includes("hidden_windows_command"));
assert(lifecycleSource.includes("CREATE_NO_WINDOW"));
assert(lifecycleSource.includes('run_windows_script(&app, "install_windows_vsto.ps1", &arguments)'));
assert(lifecycleSource.includes("OfficeIntegrationMode::Vsto"));
assert(lifecycleSource.includes('run_windows_script(&app, "test_windows_office_runtime.ps1", &arguments)'));
assert(!lifecycleSource.includes("OfficeIntegrationMode::Ole"));
assert(!windowsBundleSource.includes('"../scripts/install_windows_ole.ps1"'));
assert(!windowsBundleSource.includes('"../scripts/uninstall_windows_ole.ps1"'));
assert(!windowsBundleSource.includes('"../office/windows/ole/manifests/"'));
assert(windowsBundleSource.includes('"../scripts/test_windows_office_runtime.ps1"'));
assert(installVstoSource.includes("Assert-NoOfficeProcesses"));
assert(installVstoSource.includes("MSIRESTARTMANAGERCONTROL=Disable"));
assert(installVstoSource.includes("Assert-VstoRuntimeInstalled"));
assert(installVstoSource.includes("Assert-NetFramework48Installed"));
assert(installVstoSource.includes("Resolve-MachineOfficeInstallRoot"));
assert(installVstoSource.includes("ProgramW6432"));
assert(installVstoSource.includes("Resolve-PowerShellExecutable"));
assert(installVstoSource.includes("ArchitectureRelaunched"));
assert(installVstoSource.includes("Remove-LegacyPerUserOfficeRegistration"));
assert(installVstoSource.includes("Assert-ManagedComActivation"));
assert(installVstoSource.includes("Stop-VisualTeXProcessesForRepair"));
assert(installVstoSource.includes('$startParameters.Verb = "RunAs"'));
assert(installVstoSource.includes("Static files and registry installation verified successfully"));
assert(installVstoSource.includes("Static Office integration installed and verified"));
assert(installVstoSource.includes("deferred to the non-elevated installer stage"));
assert(!installVstoSource.includes("& $runtimeScript"));
assert(!installVstoSource.includes("chain-verified"));
assert(runtimeTestSource.includes("Get-ComAddInItem"));
assert(runtimeTestSource.includes("Resolve-OfficeExecutablePath"));
assert(runtimeTestSource.includes("Resolve-PowerShellExecutable"));
assert(runtimeTestSource.includes("ArchitectureRelaunched"));
assert(runtimeTestSource.includes("ProgramW6432"));
assert(runtimeTestSource.includes("GetActiveObject"));
assert(runtimeTestSource.includes("Start-CompanionAsInteractiveUser"));
assert(runtimeTestSource.includes("must run in the interactive user's non-elevated session"));
assert(!runtimeTestSource.includes("Shell.Application"));
assert(runtimeTestSource.includes('startupMode = "desktop-executable-rot"'));
assert(runtimeTestSource.includes("RuntimeVerificationPending"));
assert(runtimeTestSource.includes("desktop application did not enumerate"));
assert(runtimeTestSource.includes("Get-ManagedComRegistrationState"));
assert(runtimeTestSource.includes("Native Office integration installed and verified successfully"));
assert(windowsBackendSource.includes("hidden_command"));
assert(windowsBackendSource.includes("CREATE_NO_WINDOW"));
assert(windowsBackendSource.includes('HKLM\\Software\\Microsoft\\Office\\Word\\Addins\\VisualTeX.WordVsto'));
assert(windowsBackendSource.includes('HKLM\\Software\\Microsoft\\Office\\PowerPoint\\Addins\\VisualTeX.PowerPointVsto'));
assert(windowsBackendSource.includes('HKLM\\Software\\Classes\\CLSID'));
assert(windowsBackendSource.includes('Some("/reg:32")'));
assert(windowsBackendSource.includes('Some("/reg:64")'));
assert(windowsBackendSource.includes("resolve_office_registry_view"));
assert(!windowsBackendSource.includes('HKCU\\Software\\Microsoft\\Office\\Word\\Addins\\VisualTeX.WordVsto'));
assert(!windowsBackendSource.includes('HKCU\\Software\\Classes\\CLSID'));
assert(!windowsBackendSource.includes("OLE_CATALOG_KEY"));
assert(!windowsBackendSource.includes("register_ole_catalog"));
assert(hooksSource.includes("${NSD_Check} $VisualTeXOfficeNativeRadio"));
assert(hooksSource.includes("VisualTeXRepairMainUninstallRegistration"));
assert(hooksSource.includes('$INSTDIR == "$PROFILE\\AppData\\VisualTeX"'));
assert(hooksSource.includes("vsto-uninstall-bootstrap"));
assert(hooksSource.includes("NSIS_HOOK_POSTUNINSTALL"));
assert(hooksSource.includes('DeleteRegKey HKCU "Software\\visualtex\\VisualTeX"'));
assert(hooksSource.includes('RMDir /r "$INSTDIR"'));
assert(hooksSource.includes("OfficeSessions user data"));
const postUninstallHookSource = hooksSource.slice(hooksSource.indexOf("!macro NSIS_HOOK_POSTUNINSTALL"));
assert(postUninstallHookSource.includes("GetCurrentProcessId"));
assert(postUninstallHookSource.includes("Wait-Process -Id $0"));
assert(postUninstallHookSource.includes("Remove-Item -LiteralPath '$INSTDIR'"));
assert(!postUninstallHookSource.includes('RMDir /r "$APPDATA\\VisualTeX"'));
assert(uninstallVstoSource.includes("ArchitectureRelaunched"));
assert(uninstallVstoSource.includes("$process.WaitForExit()"));
assert(uninstallVstoSource.includes("vsto-uninstall-bootstrap-$stamp.log"));
assert(uninstallVstoSource.includes("remove_windows_office_certificate.ps1"));
assert(uninstallVstoSource.includes("Get-Process visualtex"));
assert(removeCertificateSource.includes("reg.exe"));
assert(removeCertificateSource.includes("SystemCertificates\\Root\\Certificates"));
assert(removeCertificateSource.includes("TimeoutSeconds"));
assert(!removeCertificateSource.includes("X509Store"));
assert(!removeCertificateSource.includes("certutil.exe"));
assert(hooksSource.includes('test_windows_office_runtime.ps1" -VisualTeXPath "$INSTDIR\\VisualTeX.exe" -CompanionOnly'));
assert(hooksSource.includes("visualtex_office_static_installed"));
assert(hooksSource.includes("visualtex_office_runtime_pending"));
assert(hooksSource.includes("RuntimeVerificationPending"));
assert(hooksSource.includes("这不代表插件安装失败"));
assert(/Companion runtime verification is not ready yet[\s\S]*?Goto visualtex_office_runtime_pending/.test(hooksSource));
assert(!/Companion runtime verification is not ready yet[\s\S]*?Goto visualtex_office_failed/.test(hooksSource));
assert(hooksSource.includes("non-elevated companion runtime verification passed"));
assert(hooksSource.includes('StrCpy $VisualTeXOfficeChoice "native"'));
assert(hooksSource.includes('-VisualTeXPath "$INSTDIR\\VisualTeX.exe"'));
assert(!hooksSource.includes('install_windows_ole.ps1'));
assert(installOleSource.includes("forwarding to the native Ribbon + OLE LocalServer installer"));
assert(certificateSource.includes("certutil.exe -user -f -addstore Root $certificatePath"));
assert(certificateSource.includes('Split-Path -Parent $PSScriptRoot'));

console.log("Windows onboarding and Office integration controls passed.");
