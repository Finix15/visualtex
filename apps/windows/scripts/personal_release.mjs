import { spawnSync } from "node:child_process";
import process from "node:process";

if (process.platform !== "win32") {
  throw new Error("The VisualTeX personal Windows release workflow is Windows-only.");
}

const mode = process.argv[2];
if (mode !== "--test" && mode !== "--release") {
  throw new Error("Usage: node scripts/personal_release.mjs --test|--release");
}

function run(command, args, extraEnv = {}) {
  const isCmd = command.toLowerCase().endsWith(".cmd");
  const executable = isCmd ? (process.env.ComSpec ?? "cmd.exe") : command;
  const executableArgs = isCmd ? ["/d", "/s", "/c", command, ...args] : args;
  console.log(`\n> ${command} ${args.join(" ")}`);
  const result = spawnSync(executable, executableArgs, {
    cwd: process.cwd(),
    env: { ...process.env, ...extraEnv },
    shell: false,
    stdio: "inherit",
  });
  if (result.error) throw result.error;
  if (result.status !== 0) process.exit(result.status ?? 1);
}

const npm = "npm.cmd";
const powershell = `${process.env.SystemRoot ?? "C:\\Windows"}\\System32\\WindowsPowerShell\\v1.0\\powershell.exe`;

run(npm, ["run", "build:desktop"]);
run(npm, ["run", "test:editor:run"]);
run(npm, ["run", "test:platform-onboarding"]);
run(npm, ["run", "test:windows-office-architecture"]);
run("node", ["scripts/personal_release_smoke.mjs"]);
run(powershell, [
  "-NoProfile",
  "-ExecutionPolicy",
  "Bypass",
  "-File",
  "scripts/check_powershell_syntax.ps1",
]);
run(powershell, [
  "-NoProfile",
  "-ExecutionPolicy",
  "Bypass",
  "-File",
  "scripts/test_certificate_cleanup_policy.ps1",
]);
run("dotnet", [
  "test",
  "src-windows/VisualTeX.WindowsOffice.Tests/VisualTeX.WindowsOffice.Tests.csproj",
  "--configuration",
  "Release",
]);
run("cargo", ["test", "--manifest-path", "src-tauri/Cargo.toml", "--lib"], {
  VISUALTEX_RUST_LIB_TEST: "1",
});

if (mode === "--release") {
  run(npm, ["run", "tauri:build"]);
}

console.log(
  mode === "--release"
    ? "VisualTeX personal release build and temporary installed acceptance passed."
    : "VisualTeX personal release test suite passed.",
);
