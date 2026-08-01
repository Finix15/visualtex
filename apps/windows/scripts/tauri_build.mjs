import { spawnSync } from "node:child_process";
import process from "node:process";

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

// 2. Prepare OCR, Office UI, VSTO/OLE and externalBin inputs exactly once.
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

// 5. Generate the NSIS source/resources from the already-built exact EXE, patch
// the same-version maintenance default, and rebuild the final installer.
run(tauri, ["bundle", "--bundles", "nsis", ...forwarded], {
  ...process.env,
  VISUALTEX_FRONTEND_PREBUILT: "1",
});
run("powershell.exe", [
  "-NoProfile",
  "-ExecutionPolicy",
  "Bypass",
  "-File",
  "scripts/patch_generated_nsis.ps1",
]);

// 6. Static package verification plus a clean-directory installed-runtime smoke
// test. The smoke test launches the exact installed visualtex.exe and requires
// the runtime asset resolver to log a successful index.html/JS/CSS preflight.
run("powershell.exe", [
  "-NoProfile",
  "-ExecutionPolicy",
  "Bypass",
  "-File",
  "scripts/verify_windows_release_artifacts.ps1",
]);
if (process.env.VISUALTEX_SKIP_INSTALL_SMOKE !== "1") {
  run("powershell.exe", [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    "scripts/test_windows_installed_release.ps1",
  ]);
}
