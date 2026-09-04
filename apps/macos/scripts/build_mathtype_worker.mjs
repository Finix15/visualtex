import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { chmodSync, copyFileSync, mkdirSync, mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const macosRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const repositoryRoot = resolve(macosRoot, "../..");
const sourceRoot = join(repositoryRoot, "tools", "mathtypejx");
const outputRoot = join(macosRoot, "src-tauri", "workers");
const python = process.env.VISUALTEX_MATHTYPE_BUILD_PYTHON;
if (!python) throw new Error("VISUALTEX_MATHTYPE_BUILD_PYTHON must name the pinned PyInstaller build environment");

const version = execFileSync(python, ["-m", "PyInstaller", "--version"], { encoding: "utf8" }).trim();
if (version !== "6.16.0") throw new Error(`PyInstaller 6.16.0 is required; found ${version}`);
const architecture = execFileSync("/usr/bin/uname", ["-m"], { encoding: "utf8" }).trim();
if (architecture !== "arm64") throw new Error(`Phase 5 worker must be built on arm64; found ${architecture}`);

const temporary = mkdtempSync(join(tmpdir(), "visualtex-mathtype-worker-"));
try {
  mkdirSync(outputRoot, { recursive: true });
  execFileSync(python, [
    "-m", "PyInstaller", "--clean", "--noconfirm", "--onefile",
    "--name", "mathtype-worker", "--paths", join(sourceRoot, "src"),
    "--collect-data", "mathtypejx", "--distpath", join(temporary, "dist"),
    "--workpath", join(temporary, "work"), "--specpath", join(temporary, "spec"),
    join(sourceRoot, "src", "mathtypejx", "worker.py"),
  ], {
    cwd: sourceRoot,
    stdio: "inherit",
    env: { ...process.env, PYINSTALLER_CONFIG_DIR: join(temporary, "config") },
  });
  const built = join(temporary, "dist", "mathtype-worker");
  const output = join(outputRoot, "mathtype-worker");
  copyFileSync(built, output);
  chmodSync(output, 0o755);
  const bytes = readFileSync(output);
  const manifest = {
    protocolVersion: 1, workerVersion: "0.1.0", target: "macos-arm64", architecture: "arm64",
    sha256: createHash("sha256").update(bytes).digest("hex"), fileSize: statSync(output).size,
    dependencies: [
      { name: "mathtypejx", version: "0.1.0", license: "MIT" },
      { name: "PyInstaller", version: "6.16.0", license: "GPL-2.0-or-later with bootloader exception" },
      { name: "lxml", version: "bundled", license: "BSD-3-Clause" },
      { name: "olefile", version: "bundled", license: "BSD-2-Clause" },
      { name: "Python", version: execFileSync(python, ["--version"], { encoding: "utf8" }).trim().replace(/^Python /, ""), license: "PSF-2.0" },
    ],
    bundledData: ["mathtypejx/mtef/xslt"],
    exclusions: ["MML2OMML.XSL", "private corpus", "virtualenv", "OCR runtime"],
  };
  writeFileSync(join(outputRoot, "mathtype-worker.manifest.json"), `${JSON.stringify(manifest, null, 2)}\n`);
  copyFileSync(join(sourceRoot, "LICENSE"), join(outputRoot, "mathtypejx.LICENSE"));
  copyFileSync(join(sourceRoot, "NOTICE"), join(outputRoot, "mathtypejx.NOTICE"));
  process.stdout.write(`Built ${output}\n${manifest.sha256}  mathtype-worker\n`);
} finally {
  rmSync(temporary, { recursive: true, force: true });
}
