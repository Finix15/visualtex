import path from "node:path";

export function windowsPowerShellPath(env = process.env) {
  const windowsRoot = env.SystemRoot ?? env.WINDIR;
  if (!windowsRoot) {
    throw new Error("SystemRoot/WINDIR is unavailable; cannot locate Windows PowerShell");
  }
  return path.join(
    windowsRoot,
    "System32",
    "WindowsPowerShell",
    "v1.0",
    "powershell.exe",
  );
}
