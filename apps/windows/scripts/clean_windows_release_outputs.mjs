import { readdir, rm } from "node:fs/promises";
import { join, resolve } from "node:path";
import process from "node:process";
import { setTimeout as delay } from "node:timers/promises";

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
  const retryable = new Set(["EBUSY", "EPERM", "ENOTEMPTY"]);
  for (let attempt = 1; attempt <= 30; attempt += 1) {
    try {
      await rm(path, {
        recursive: true,
        force: true,
        maxRetries: 3,
        retryDelay: 250,
      });
      console.log(`Cleaned release output: ${path}`);
      return;
    } catch (error) {
      if (
        attempt === 30 ||
        !error ||
        typeof error !== "object" ||
        !("code" in error) ||
        !retryable.has(String(error.code))
      ) {
        throw error;
      }
      const waitMs = Math.min(5000, attempt * 400);
      console.warn(
        `Release output is temporarily locked; retrying in ${waitMs} ms (${attempt}/30): ${path}`,
      );
      await delay(waitMs);
    }
  }
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
