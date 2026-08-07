import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import {
  chmodSync,
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  renameSync,
  writeFileSync,
} from "node:fs";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

if (process.platform !== "darwin") {
  throw new Error("The PowerPoint performance installer is available only on macOS.");
}

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const resourcesRoot = join(
  repositoryRoot,
  "office",
  "macos-offline",
  "resources",
);
const scratchRoot = join(
  homedir(),
  "Library",
  "Group Containers",
  "UBF8T346G9.Office",
  "VisualTeX",
  "Scratch",
);
const backupRoot = join(scratchRoot, "PowerPointPerformancePreOptimizationBackup");
const resourceWord = join(resourcesRoot, "VisualTeX.dotm");
const resourcePowerPoint = join(resourcesRoot, "VisualTeX.ppam");
const resourceManifest = join(resourcesRoot, "addins.json");
const installedPowerPoint = join(
  homedir(),
  "Library",
  "Group Containers",
  "UBF8T346G9.Office",
  "VisualTeX",
  "OfficeAddins",
  "VisualTeX.ppam",
);
const compiledPresentation = join(
  scratchRoot,
  "VisualTeXPowerPointPerformanceOptimized.pptm",
);
const testPowerPoint = join(scratchRoot, "VisualTeXPowerPointPerformanceOptimized.ppam");
const packageScript = join(repositoryRoot, "scripts", "package_macos_offline_addins.mjs");
const restore = process.argv.includes("--restore");

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

function atomicCopy(source, destination, mode) {
  mkdirSync(dirname(destination), { recursive: true });
  const staged = `${destination}.staged.${process.pid}`;
  copyFileSync(source, staged);
  if (mode !== undefined) chmodSync(staged, mode);
  renameSync(staged, destination);
  if (sha256(source) !== sha256(destination)) {
    throw new Error(`Copy verification failed: ${destination}`);
  }
}

function backupOnce(source, name) {
  const destination = join(backupRoot, name);
  if (!existsSync(destination)) atomicCopy(source, destination);
  return destination;
}

function requireFile(path, label) {
  if (!existsSync(path)) throw new Error(`${label} is missing: ${path}`);
}

function writeManifest() {
  const packageVersion = JSON.parse(
    readFileSync(join(repositoryRoot, "package.json"), "utf8"),
  ).version;
  const manifest = {
    schemaVersion: 1,
    pluginVersion: packageVersion,
    files: {
      "VisualTeX.dotm": { sha256: sha256(resourceWord) },
      "VisualTeX.ppam": { sha256: sha256(resourcePowerPoint) },
    },
  };
  writeFileSync(resourceManifest, `${JSON.stringify(manifest, null, 2)}\n`, "utf8");
}

requireFile(resourceWord, "Word resource");
requireFile(resourcePowerPoint, "PowerPoint resource");
requireFile(resourceManifest, "Office manifest");
requireFile(installedPowerPoint, "Installed PowerPoint add-in");
mkdirSync(backupRoot, { recursive: true });

if (restore) {
  for (const [name, destination, mode] of [
    ["VisualTeX.dotm", resourceWord, undefined],
    ["VisualTeX.ppam", resourcePowerPoint, undefined],
    ["addins.json", resourceManifest, undefined],
    ["VisualTeX.installed.ppam", installedPowerPoint, 0o600],
  ]) {
    const source = join(backupRoot, name);
    requireFile(source, `PowerPoint performance backup ${name}`);
    atomicCopy(source, destination, mode);
  }
  process.stdout.write(
    `Restored PowerPoint resources and installation from ${backupRoot}.\n`,
  );
  process.exit(0);
}

requireFile(compiledPresentation, "Compiled PowerPoint performance presentation");
backupOnce(resourceWord, "VisualTeX.dotm");
backupOnce(resourcePowerPoint, "VisualTeX.ppam");
backupOnce(resourceManifest, "addins.json");
backupOnce(installedPowerPoint, "VisualTeX.installed.ppam");

execFileSync(
  process.execPath,
  [
    packageScript,
    "--word",
    resourceWord,
    "--powerpoint",
    compiledPresentation,
    "--powerpoint-shell",
    join(backupRoot, "VisualTeX.ppam"),
    "--root-powerpoint-output",
    testPowerPoint,
  ],
  { stdio: "inherit", timeout: 180_000 },
);

// The official packager validates both hosts and rewrites both resource files.
// This performance iteration changes only PowerPoint, so restore the exact Word
// resource bytes and regenerate the manifest against the final pair.
atomicCopy(join(backupRoot, "VisualTeX.dotm"), resourceWord);
writeManifest();
atomicCopy(resourcePowerPoint, testPowerPoint);
atomicCopy(resourcePowerPoint, installedPowerPoint, 0o600);

if (sha256(resourcePowerPoint) !== sha256(installedPowerPoint)) {
  throw new Error("Installed PowerPoint add-in differs from the packaged resource.");
}
process.stdout.write(
  `${JSON.stringify(
    {
      status: "INSTALLED",
      compiledPresentation,
      packagedPowerPoint: resourcePowerPoint,
      installedPowerPoint,
      backupRoot,
      sha256: sha256(resourcePowerPoint),
    },
    null,
    2,
  )}\n`,
);
