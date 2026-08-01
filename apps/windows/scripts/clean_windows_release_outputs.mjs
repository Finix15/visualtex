import { readdir, rm } from "node:fs/promises";
import { join, resolve } from "node:path";
import process from "node:process";

if (process.platform !== "win32") {
  throw new Error("Windows release cleanup must run on Windows");
}

const root = process.cwd();
const releaseRoot = resolve(root, "src-tauri/target/release");
const fixedPaths = [
  resolve(root, "dist"),
  join(releaseRoot, "visualtex.exe"),
  join(releaseRoot, "frontend-dist-manifest.json"),
  join(releaseRoot, "nsis"),
  join(releaseRoot, "bundle", "nsis"),
];

async function remove(path) {
  await rm(path, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
  console.log(`Cleaned release output: ${path}`);
}

for (const path of fixedPaths) await remove(path);

for (const parent of [join(releaseRoot, "build"), join(releaseRoot, ".fingerprint")]) {
  const entries = await readdir(parent, { withFileTypes: true }).catch(() => []);
  for (const entry of entries) {
    if (entry.name.startsWith("visualtex-")) {
      await remove(join(parent, entry.name));
    }
  }
}

console.log(
  "Windows release outputs cleaned without touching artifacts/, build-logs/, user data, Office build outputs or unrelated Cargo dependencies.",
);
