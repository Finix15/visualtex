import { spawnSync } from "node:child_process";
import { windowsPowerShellPath } from "./windows_powershell.mjs";

function run(command, args) {
  const isWindowsCmd =
    process.platform === "win32" && command.toLowerCase().endsWith(".cmd");
  const executable = isWindowsCmd ? (process.env.ComSpec ?? "cmd.exe") : command;
  const executableArgs = isWindowsCmd
    ? ["/d", "/s", "/c", command, ...args]
    : args;
  const result = spawnSync(executable, executableArgs, {
    stdio: "inherit",
    shell: false,
  });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    throw new Error(`${command} ${args.join(" ")} failed with ${result.status}`);
  }
}

const npm = process.platform === "win32" ? "npm.cmd" : "npm";
const powershell = process.platform === "win32" ? windowsPowerShellPath() : "powershell";

// Prepare only native/resource inputs consumed by Tauri. The main desktop
// frontend is built exactly once by tauri_build.mjs before Tauri codegen.
if (process.platform === "win32") {
  run(powershell, [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    "scripts/prepare_windows_ocr_python.ps1",
  ]);
  run(powershell, [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    "scripts/prepare_windows_vsto_runtime.ps1",
  ]);
  run(npm, ["run", "build:office:windows-native"]);
  run(powershell, [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    "scripts/build_windows_office.ps1",
    "-Configuration",
    "Release",
    "-SkipTests",
  ]);
}
