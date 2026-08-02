import { spawnSync } from "node:child_process";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import process from "node:process";
import { windowsPowerShellPath } from "./windows_powershell.mjs";

const require = createRequire(import.meta.url);

function run(command, args, env = process.env) {
  const isWindowsCmd =
    process.platform === "win32" && command.toLowerCase().endsWith(".cmd");
  const executable = isWindowsCmd ? (process.env.ComSpec ?? "cmd.exe") : command;
  const executableArgs = isWindowsCmd
    ? ["/d", "/s", "/c", command, ...args]
    : args;
  console.log(`\n> ${command} ${args.join(" ")}`);
  const result = spawnSync(executable, executableArgs, {
    stdio: "inherit",
    shell: false,
    env,
  });
  if (result.error) throw result.error;
  if (result.status !== 0) process.exit(result.status ?? 1);
}

function preparePatchedNsisTemplate() {
  const nativePackage =
    process.arch === "arm64"
      ? "@tauri-apps/cli-win32-arm64-msvc"
      : "@tauri-apps/cli-win32-x64-msvc";
  const nativeCliPath = require.resolve(nativePackage);
  const binary = readFileSync(nativeCliPath);

  let template;
  for (const newline of ["\r\n", "\n"]) {
    const startMarker = Buffer.from(`Unicode true${newline}ManifestDPIAware true`);
    const endMarker = Buffer.from(
      `  !insertmacro SetLnkAppUserModelId "$DESKTOP\\\${PRODUCTNAME}.lnk"${newline}FunctionEnd${newline}`,
    );
    const start = binary.indexOf(startMarker);
    const endStart = start >= 0 ? binary.indexOf(endMarker, start) : -1;
    if (start >= 0 && endStart >= 0) {
      template = binary
        .subarray(start, endStart + endMarker.length)
        .toString("utf8")
        .replace(/\r\n/g, "\n");
      break;
    }
  }
  if (!template) {
    throw new Error(
      `Unable to extract the exact NSIS installer template from ${nativeCliPath}`,
    );
  }

  const functionStart = "Function PageReinstall\n";
  if (template.split(functionStart).length !== 2) {
    throw new Error("The Tauri NSIS PageReinstall function changed unexpectedly");
  }
  const acceptanceStart = [
    "Function PageReinstall",
    "  ; Installed-release acceptance uses a clean custom directory and must not",
    "  ; enter maintenance mode for another installed VisualTeX copy.",
    "  ${GetParameters} $R7",
    "  ClearErrors",
    '  ${GetOptions} $R7 "/VISUALTEXACCEPTANCE" $R8',
    "  ${IfNot} ${Errors}",
    "    Abort",
    "  ${EndIf}",
    "",
  ].join("\n");
  template = template.replace(functionStart, acceptanceStart);

  const oldSelection = [
    "    ; Check the first radio button if this the first time",
    "    ; we enter this page or if the second button wasn't",
    "    ; selected the last time we were on this page",
    "    ${If} $ReinstallPageCheck <> 2",
    "      SendMessage $R2 ${BM_SETCHECK} ${BST_CHECKED} 0",
    "    ${Else}",
    "      SendMessage $R3 ${BM_SETCHECK} ${BST_CHECKED} 0",
    "    ${EndIf}",
    "",
    "    ${NSD_SetFocus} $R2",
  ].join("\n");
  const newSelection = [
    '    ; Same-version maintenance defaults to the second option, "Uninstall',
    '    ; VisualTeX". Preserve an explicit user selection when navigating back.',
    "    ; Upgrade/downgrade pages keep Tauri's original first-option default.",
    "    ${If} $R0 = 0",
    "      ${If} $ReinstallPageCheck = 1",
    "        SendMessage $R2 ${BM_SETCHECK} ${BST_CHECKED} 0",
    "        ${NSD_SetFocus} $R2",
    "      ${Else}",
    "        SendMessage $R3 ${BM_SETCHECK} ${BST_CHECKED} 0",
    "        StrCpy $ReinstallPageCheck 2",
    "        ${NSD_SetFocus} $R3",
    "      ${EndIf}",
    "    ${Else}",
    "      ${If} $ReinstallPageCheck <> 2",
    "        SendMessage $R2 ${BM_SETCHECK} ${BST_CHECKED} 0",
    "        ${NSD_SetFocus} $R2",
    "      ${Else}",
    "        SendMessage $R3 ${BM_SETCHECK} ${BST_CHECKED} 0",
    "        ${NSD_SetFocus} $R3",
    "      ${EndIf}",
    "    ${EndIf}",
  ].join("\n");
  if (!template.includes(oldSelection)) {
    throw new Error(
      "The Tauri NSIS maintenance selection block changed unexpectedly; refusing an unverified installer build",
    );
  }
  template = template.replace(oldSelection, newSelection);

  const output = resolve(
    "src-tauri",
    "target",
    "nsis-template",
    "visualtex-installer.nsi",
  );
  mkdirSync(dirname(output), { recursive: true });
  writeFileSync(output, template, "utf8");
  console.log(`Prepared verified VisualTeX NSIS template: ${output}`);
}

function releaseArguments(args) {
  const result = [];
  for (let index = 0; index < args.length; index += 1) {
    const argument = args[index];
    if (argument === "--debug" || argument === "-d") {
      throw new Error("VisualTeX Windows installer builds must use release mode");
    }
    if (argument === "--no-bundle") continue;
    if (argument === "--bundles" || argument === "-b") {
      index += 1;
      continue;
    }
    if (argument.startsWith("--bundles=")) continue;
    result.push(argument);
  }
  return result;
}

const npm = process.platform === "win32" ? "npm.cmd" : "npm";
const tauri = process.platform === "win32" ? "tauri.cmd" : "tauri";
const node = process.execPath;
const powershell = process.platform === "win32" ? windowsPowerShellPath() : "powershell";
const forwarded = releaseArguments(process.argv.slice(2));

if (process.platform !== "win32") {
  run(npm, ["run", "build:desktop"]);
  run(node, ["scripts/verify_frontend_dist.mjs"]);
  run(tauri, ["build", ...forwarded]);
  process.exit(0);
}

// 1. Remove only stale release/codegen outputs that can cause Tauri to reuse an
// old embedded asset table. User artifacts, logs and native build caches remain.
run(node, ["scripts/clean_windows_release_outputs.mjs"]);

// 2. Prepare the exact Tauri-version NSIS template with VisualTeX's verified
// maintenance default, then prepare OCR, Office UI, VSTO/OLE and externalBin inputs.
preparePatchedNsisTemplate();
run(npm, ["run", "build:bundle"]);

// 3. Build the main frontend exactly once, then validate every referenced asset
// before Tauri build.rs/codegen is allowed to execute.
run(npm, ["run", "build:desktop"]);
run(node, ["scripts/verify_frontend_dist.mjs"]);

// 4. Build the Rust application without bundling. beforeBuildCommand performs
// validation only and therefore cannot rebuild or replace dist.
run(tauri, ["build", "--no-bundle", ...forwarded], {
  ...process.env,
  VISUALTEX_FRONTEND_PREBUILT: "1",
});
run(node, [
  "scripts/verify_embedded_frontend_assets.mjs",
  "--exe",
  "src-tauri/target/release/visualtex.exe",
]);

// 5. Bundle directly with the verified custom NSIS template. This avoids
// depending on Tauri's ephemeral generated installer.nsi after makensis exits.
run(tauri, ["bundle", "--bundles", "nsis", ...forwarded], {
  ...process.env,
  VISUALTEX_FRONTEND_PREBUILT: "1",
});

// 6. Static package verification plus a clean-directory installed-runtime smoke
// test. The smoke test launches the exact installed visualtex.exe and requires
// the runtime asset resolver to log a successful index.html/JS/CSS preflight.
run(powershell, [
  "-NoProfile",
  "-ExecutionPolicy",
  "Bypass",
  "-File",
  "scripts/verify_windows_release_artifacts.ps1",
]);
if (process.env.VISUALTEX_SKIP_INSTALL_SMOKE !== "1") {
  run(powershell, [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    "scripts/test_windows_installed_release.ps1",
  ]);
}
