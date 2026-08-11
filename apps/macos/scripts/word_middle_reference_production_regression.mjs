import { spawn, spawnSync } from "node:child_process";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  rmSync,
} from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const sleepSync = (ms) => {
  const buffer = new SharedArrayBuffer(4);
  Atomics.wait(new Int32Array(buffer), 0, 0, ms);
};

function option(name) {
  const index = process.argv.indexOf(name);
  if (index < 0) return "";
  const value = process.argv[index + 1];
  if (!value || value.startsWith("--")) {
    throw new Error(`${name} requires a value`);
  }
  return value;
}

const wordAddinPath = option("--word-addin");
const damagedDocumentBackupPath = option("--damaged-document");
if (!wordAddinPath || !existsSync(wordAddinPath)) {
  throw new Error(`Missing compiled Word add-in: ${wordAddinPath}`);
}
if (damagedDocumentBackupPath && !existsSync(damagedDocumentBackupPath)) {
  throw new Error(`Missing damaged-document backup: ${damagedDocumentBackupPath}`);
}

const home = homedir();
const runtimeTests = join(
  home,
  "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime/Tests",
);
const genericOpenXmlPath = join(runtimeTests, "document-import-openxml.xml");
const resultPath = join(
  runtimeTests,
  "word-middle-reference-production-regression-result.txt",
);
const startupRoot = join(
  home,
  "Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word",
);
const installedAddinPath = join(startupRoot, "VisualTeX.dotm");
const installedBackupPath = join(
  home,
  "Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch/VisualTeX-startup-before-middle-reference-regression.dotm",
);
const repairedScenePath = join(
  home,
  "Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch/Document1-numbering-reference-repair-evidence.docx",
);

function runAppleScript(lines, timeout = 30_000) {
  const args = [];
  for (const line of lines) args.push("-e", line);
  const result = spawnSync("/usr/bin/osascript", args, {
    encoding: "utf8",
    timeout,
  });
  if (result.status !== 0) {
    throw new Error(result.stderr.trim() || result.stdout.trim() || "AppleScript failed");
  }
  return result.stdout.trim();
}

function runJxa(source, timeout = 15_000) {
  const result = spawnSync(
    "/usr/bin/osascript",
    ["-l", "JavaScript", "-e", source],
    { encoding: "utf8", timeout },
  );
  if (result.status !== 0) return "";
  return result.stdout.trim();
}

function runWordMacro(name, timeout = 60_000) {
  let lastError = null;
  for (let attempt = 1; attempt <= 6; attempt += 1) {
    try {
      return runAppleScript([
        'tell application "Microsoft Word"',
        "activate",
        `run VB macro macro name ${JSON.stringify(name)}`,
        "end tell",
      ], timeout);
    } catch (error) {
      lastError = error;
      if (!/(-609|connection invalid|连接无效)/i.test(String(error?.message ?? error))) {
        throw error;
      }
      spawnSync("/usr/bin/open", ["-a", "Microsoft Word"], {
        encoding: "utf8",
        timeout: 10_000,
      });
      sleepSync(700);
    }
  }
  throw lastError ?? new Error(`Unable to run Word macro ${name}`);
}

function spawnWordMacro(name) {
  return spawn(
    "/usr/bin/osascript",
    [
      "-e",
      'tell application "Microsoft Word"',
      "-e",
      "activate",
      "-e",
      `run VB macro macro name ${JSON.stringify(name)}`,
      "-e",
      "end tell",
    ],
    { stdio: ["ignore", "pipe", "pipe"] },
  );
}

async function waitForChild(child, label, timeoutMs = 30_000) {
  return await new Promise((resolve, reject) => {
    let stdout = "";
    let stderr = "";
    child.stdout?.on("data", (chunk) => { stdout += chunk; });
    child.stderr?.on("data", (chunk) => { stderr += chunk; });
    const timer = setTimeout(() => {
      child.kill("SIGKILL");
      reject(new Error(`${label} timed out`));
    }, timeoutMs);
    child.on("error", (error) => {
      clearTimeout(timer);
      reject(error);
    });
    child.on("exit", (code) => {
      clearTimeout(timer);
      if (code === 0) resolve(stdout.trim());
      else reject(new Error(stderr.trim() || stdout.trim() || `${label} exited ${code}`));
    });
  });
}

function wordUiContains(...markers) {
  const serializedMarkers = JSON.stringify(markers.map((value) => value.toLowerCase()));
  const output = runJxa(String.raw`
const se = Application("System Events");
const word = se.processes.byName("Microsoft Word");
const markers = ${serializedMarkers};
let found = false;
let visited = 0;
function text(element, property) {
  try {
    const value = element[property]();
    return value === null || value === undefined ? "" : String(value);
  } catch (_) { return ""; }
}
function walk(element, depth) {
  if (found || depth > 10 || visited > 4000) return;
  visited += 1;
  const combined = ["name", "description", "help", "value"]
    .map((property) => text(element, property).toLowerCase())
    .join(" ");
  if (markers.some((marker) => combined.includes(marker))) {
    found = true;
    return;
  }
  let children = [];
  try { children = element.uiElements(); } catch (_) {}
  for (const child of children) walk(child, depth + 1);
}
if (word.exists()) {
  const windows = word.windows();
  for (const window of windows) walk(window, 0);
}
found ? "READY" : "WAIT";
`);
  return output === "READY";
}

async function waitForWordUi(markers, timeoutMs = 12_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (wordUiContains(...markers)) return;
    await sleep(150);
  }
  throw new Error(`Word UI did not show ${markers.join(" / ")}`);
}

function sendWordKey({ text = "", escape = false, enter = false }) {
  const lines = [
    'tell application "System Events"',
    'tell process "Microsoft Word"',
    "set frontmost to true",
  ];
  if (text) lines.push(`keystroke ${JSON.stringify(text)}`);
  if (escape) lines.push("key code 53");
  if (enter) lines.push("key code 36");
  lines.push("end tell", "end tell");
  runAppleScript(lines);
}

function wordUiSummary() {
  return runJxa(String.raw`
const se = Application("System Events");
const word = se.processes.byName("Microsoft Word");
const rows = [];
let visited = 0;
function text(element, property) {
  try {
    const value = element[property]();
    return value === null || value === undefined ? "" : String(value);
  } catch (_) { return ""; }
}
function walk(element, depth) {
  if (depth > 8 || visited > 1000 || rows.length > 120) return;
  visited += 1;
  const role = text(element, "role");
  const name = text(element, "name");
  const value = text(element, "value");
  if (role || name || value) rows.push([depth, role, name, value].join("|"));
  let children = [];
  try { children = element.uiElements(); } catch (_) {}
  for (const child of children) walk(child, depth + 1);
}
if (word.exists()) for (const window of word.windows()) walk(window, 0);
rows.join("\\n");
`);
}

async function invokePicker(action) {
  if (action === "cancel") {
    runWordMacro("VisualTeX_SetWordMiddleReferenceRegressionCancelSelection");
    runWordMacro("VisualTeX_OpenEquationCrossReference");
    return;
  }
  if (Number(action) !== 1) {
    throw new Error(`The production regression only supports picker item 1, got ${action}`);
  }
  runWordMacro("VisualTeX_SetWordMiddleReferenceRegressionSelection");
  runWordMacro("VisualTeX_OpenEquationCrossReference");
}

function visualTeXEditorReady() {
  const output = runJxa(String.raw`
const se = Application("System Events");
const app = se.processes.byName("visualtex");
if (!app.exists()) "WAIT";
else {
  let count = 0;
  try { count = app.windows().length; } catch (_) {}
  (count > 0 && app.frontmost()) ? "READY" : "WAIT";
}
`);
  return output === "READY";
}

async function applySelectedFormulaUnchanged() {
  runWordMacro("VisualTeX_EditSelected");
  const deadline = Date.now() + 20_000;
  while (Date.now() < deadline && !visualTeXEditorReady()) await sleep(150);
  if (!visualTeXEditorReady()) {
    throw new Error("VisualTeX formula editor did not become ready");
  }
  runAppleScript([
    'tell application "System Events"',
    'tell process "visualtex"',
    "set frontmost to true",
    "key code 36 using {command down}",
    "end tell",
    "end tell",
  ]);
  await sleep(2_000);
}

function selectImageFormulaById(formulaId) {
  runAppleScript([
    'tell application "Microsoft Word"',
    'set documentObject to active document',
    'set foundFormula to false',
    'repeat with shapeIndex from 1 to (count of inline shapes of documentObject)',
    'set formulaShape to inline shape shapeIndex of documentObject',
    `if (title of formulaShape as text) contains ${JSON.stringify(formulaId)} then`,
    'select text object of formulaShape',
    'set foundFormula to true',
    'exit repeat',
    'end if',
    'end repeat',
    'if foundFormula is false then error "Target VisualTeX image formula was not found"',
    'end tell',
  ], 30_000);
}

function selectDocumentEnd() {
  runAppleScript([
    'tell application "Microsoft Word"',
    'set documentObject to active document',
    'set endPosition to (end of content of text object of documentObject) - 1',
    'select (create range documentObject start endPosition end endPosition)',
    'end tell',
  ], 30_000);
}

async function insertFirstReferenceViaProductionPicker() {
  const child = spawnWordMacro("VisualTeX_OpenEquationCrossReference");
  await sleep(2_000);
  sendWordKey({ text: "1" });
  sendWordKey({ enter: true });
  await waitForChild(child, "VisualTeX_OpenEquationCrossReference", 30_000);
}

async function runPublicUpdateNumbers() {
  const child = spawnWordMacro("VisualTeX_UpdateEquationNumbers");
  await waitForWordUi(["updated", "visualtex equation numbers"], 20_000);
  sendWordKey({ enter: true });
  await waitForChild(child, "VisualTeX_UpdateEquationNumbers", 30_000);
}

function xmlDecode(value) {
  return value
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&amp;", "&")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'");
}

function visibleParagraphText(paragraphXml) {
  return [...paragraphXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
    .map((match) => xmlDecode(match[1]))
    .join("");
}

function snapshotInventory(xml) {
  const bookmarkNames = [...xml.matchAll(/<w:bookmarkStart\b[^>]*w:name="([^"]+)"/g)]
    .map((match) => match[1]);
  const countPrefix = (prefix) => bookmarkNames.filter((name) => name.startsWith(prefix)).length;
  const paragraphs = xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? [];
  const fieldCode = (paragraph) =>
    [...paragraph.matchAll(/<w:instrText(?:\s[^>]*)?>([\s\S]*?)<\/w:instrText>/g)]
      .map((match) => xmlDecode(match[1]))
      .join("")
      .replace(/\s+/g, " ")
      .trim();
  const seqParagraphs = paragraphs.filter((paragraph) => /(?:^|\s)SEQ\s+公式(?:\s|$)/i.test(fieldCode(paragraph)));
  const pollutedHelpers = seqParagraphs.filter((paragraph) => /(?:^|\s)REF\s+VT_N_/i.test(fieldCode(paragraph)));
  const unbookmarkedSeq = seqParagraphs.filter((paragraph) => !/w:name="VT_N_/i.test(paragraph));
  const helperNumbers = seqParagraphs
    .filter((paragraph) => /w:name="VT_N_/i.test(paragraph))
    .map((paragraph) => visibleParagraphText(paragraph).trim())
    .filter(Boolean);
  const bodyRefParagraphs = paragraphs.filter((paragraph) => {
    const code = fieldCode(paragraph);
    return /(?:^|\s)REF\s+VT_N_/i.test(code) &&
      !/(?:^|\s)SEQ\s+公式(?:\s|$)/i.test(code) &&
      !/<m:oMath\b/i.test(paragraph);
  });
  return {
    vtR: countPrefix("VT_R_"),
    vtN: countPrefix("VT_N_"),
    vtC: countPrefix("VT_C_"),
    seq: seqParagraphs.length,
    unbookmarkedSeq: unbookmarkedSeq.length,
    pollutedHelpers: pollutedHelpers.length,
    helperNumbers,
    bodyRefParagraphs: bodyRefParagraphs.length,
  };
}

function assertCleanFiveFormulaInventory(inventory, label, expectedBodyRefs) {
  const expectedCounts = [inventory.vtR, inventory.vtN, inventory.vtC, inventory.seq];
  if (expectedCounts.some((value) => value !== 5)) {
    throw new Error(`${label}: expected VT_R/VT_N/VT_C/SEQ = 5, got ${JSON.stringify(inventory)}`);
  }
  if (inventory.unbookmarkedSeq !== 0 || inventory.pollutedHelpers !== 0) {
    throw new Error(`${label}: detached or polluted SEQ helper remains: ${JSON.stringify(inventory)}`);
  }
  if (inventory.bodyRefParagraphs !== expectedBodyRefs) {
    throw new Error(`${label}: expected ${expectedBodyRefs} body REF paragraphs, got ${JSON.stringify(inventory)}`);
  }
}

function dumpSnapshot(label, expectedBodyRefs) {
  runWordMacro("VisualTeX_DumpActiveDocumentOpenXmlForRegression");
  if (!existsSync(genericOpenXmlPath)) {
    throw new Error(`OpenXML dump missing for ${label}`);
  }
  const xml = readFileSync(genericOpenXmlPath, "utf8");
  const target = join(runtimeTests, `word-middle-reference-${label}.xml`);
  copyFileSync(genericOpenXmlPath, target);
  const inventory = snapshotInventory(xml);
  assertCleanFiveFormulaInventory(inventory, label, expectedBodyRefs);
  return inventory;
}

function keepOnlyActiveWordDocument() {
  runAppleScript([
    'tell application "Microsoft Word"',
    'if (count of documents) = 0 then error "No active regression document"',
    'set keepName to name of active document as text',
    'repeat with i from (count of documents) to 1 by -1',
    'set d to document i',
    'if (name of d as text) is not keepName then close d saving no',
    'end repeat',
    'activate object document keepName',
    'end tell',
  ], 30_000);
}

function quitWord() {
  try {
    runAppleScript(['tell application "Microsoft Word" to quit saving no'], 20_000);
  } catch (_) {}
  spawnSync("/usr/bin/killall", ["Microsoft Word"], {
    encoding: "utf8",
    timeout: 10_000,
  });
}

async function waitForWordReady(timeoutMs = 30_000) {
  spawnSync("/usr/bin/open", ["-a", "Microsoft Word"], {
    encoding: "utf8",
    timeout: 10_000,
  });
  const deadline = Date.now() + timeoutMs;
  let consecutiveReady = 0;
  while (Date.now() < deadline) {
    try {
      const version = runAppleScript([
        'tell application "Microsoft Word"',
        "activate",
        "return version",
        "end tell",
      ], 5_000);
      if (version) {
        consecutiveReady += 1;
        if (consecutiveReady >= 2) return;
        await sleep(500);
        continue;
      }
    } catch (_) {}
    consecutiveReady = 0;
    await sleep(500);
  }
  throw new Error("Microsoft Word did not become stably automation-ready");
}

async function verifyDamagedSceneRepair() {
  if (!damagedDocumentBackupPath) return null;
  copyFileSync(damagedDocumentBackupPath, repairedScenePath);
  const repairedSceneName = repairedScenePath.split("/").at(-1);
  runAppleScript([
    'tell application "Microsoft Word"',
    `open file name ${JSON.stringify(repairedScenePath)}`,
    `activate object document ${JSON.stringify(repairedSceneName)}`,
    "end tell",
  ], 30_000);
  await sleep(700);
  keepOnlyActiveWordDocument();
  const before = (() => {
    runWordMacro("VisualTeX_DumpActiveDocumentOpenXmlForRegression");
    const xml = readFileSync(genericOpenXmlPath, "utf8");
    copyFileSync(genericOpenXmlPath, join(runtimeTests, "word-middle-reference-damaged-before-update.xml"));
    return snapshotInventory(xml);
  })();
  if (before.seq !== 6 || before.unbookmarkedSeq < 1 || before.pollutedHelpers < 1) {
    throw new Error(`Damaged-scene precondition changed: ${JSON.stringify(before)}`);
  }
  await runPublicUpdateNumbers();
  runWordMacro("VisualTeX_DumpActiveDocumentOpenXmlForRegression");
  const repairedXml = readFileSync(genericOpenXmlPath, "utf8");
  copyFileSync(genericOpenXmlPath, join(runtimeTests, "word-middle-reference-damaged-after-update.xml"));
  const after = snapshotInventory(repairedXml);
  assertCleanFiveFormulaInventory(after, "damaged-scene-after-update", 1);
  const normalizedNumbers = after.helperNumbers.map((value) => value.replace(/\s+/g, ""));
  const expected = ["1‐1", "1‐2", "1‐3", "1‐4", "1‐5"];
  if (JSON.stringify(normalizedNumbers) !== JSON.stringify(expected)) {
    throw new Error(`Damaged scene did not repair to 1-1..1-5: ${JSON.stringify(after)}`);
  }
  selectDocumentEnd();
  await insertFirstReferenceViaProductionPicker();
  runWordMacro("VisualTeX_DumpActiveDocumentOpenXmlForRegression");
  const afterReferenceXml = readFileSync(genericOpenXmlPath, "utf8");
  copyFileSync(
    genericOpenXmlPath,
    join(runtimeTests, "word-middle-reference-damaged-after-reference.xml"),
  );
  const afterReference = snapshotInventory(afterReferenceXml);
  assertCleanFiveFormulaInventory(
    afterReference,
    "damaged-scene-after-reference",
    2,
  );
  const afterReferenceNumbers = afterReference.helperNumbers
    .map((value) => value.replace(/\s+/g, ""));
  if (JSON.stringify(afterReferenceNumbers) !== JSON.stringify(expected)) {
    throw new Error(
      `Reference insertion changed Equation ordinals: ${JSON.stringify(afterReference)}`,
    );
  }

  selectImageFormulaById("039e4e2b-7acd-45f8-bb4e-b416e0a322cd");
  await applySelectedFormulaUnchanged();
  runWordMacro("VisualTeX_DumpActiveDocumentOpenXmlForRegression");
  const afterEditXml = readFileSync(genericOpenXmlPath, "utf8");
  copyFileSync(
    genericOpenXmlPath,
    join(runtimeTests, "word-middle-reference-damaged-after-edit.xml"),
  );
  const afterEdit = snapshotInventory(afterEditXml);
  assertCleanFiveFormulaInventory(afterEdit, "damaged-scene-after-edit", 2);

  await runPublicUpdateNumbers();
  runWordMacro("VisualTeX_DumpActiveDocumentOpenXmlForRegression");
  const finalXml = readFileSync(genericOpenXmlPath, "utf8");
  copyFileSync(
    genericOpenXmlPath,
    join(runtimeTests, "word-middle-reference-damaged-final.xml"),
  );
  const finalState = snapshotInventory(finalXml);
  assertCleanFiveFormulaInventory(finalState, "damaged-scene-final", 2);
  const finalNumbers = finalState.helperNumbers
    .map((value) => value.replace(/\s+/g, ""));
  if (JSON.stringify(finalNumbers) !== JSON.stringify(expected)) {
    throw new Error(
      `Edit + Update changed Equation ordinals: ${JSON.stringify(finalState)}`,
    );
  }

  runAppleScript([
    'tell application "Microsoft Word"',
    "save active document",
    "close active document saving no",
    "end tell",
  ]);
  return {
    before,
    afterRepair: after,
    afterReference,
    afterEdit,
    finalState,
    repairedScenePath,
  };
}

mkdirSync(runtimeTests, { recursive: true });
mkdirSync(startupRoot, { recursive: true });
let startupBackedUp = false;
const snapshots = {};
let damagedRepair = null;

try {
  quitWord();
  rmSync(installedBackupPath, { force: true });
  if (existsSync(installedAddinPath)) {
    copyFileSync(installedAddinPath, installedBackupPath);
    startupBackedUp = true;
    rmSync(installedAddinPath, { force: true });
  }
  copyFileSync(wordAddinPath, installedAddinPath);
  await waitForWordReady();

  runWordMacro(
    "VisualTeX_RunWordMiddleReferenceInsertProductionRegression",
    90_000,
  );
  const result = readFileSync(resultPath, "utf8").trim();
  if (!result.startsWith("PASS")) {
    throw new Error(`Middle-reference production regression failed: ${result}`);
  }
  console.log(result);
  console.log("Word middle-reference production regression: PASS");
} finally {
  quitWord();
  rmSync(installedAddinPath, { force: true });
  if (startupBackedUp && existsSync(installedBackupPath)) {
    copyFileSync(installedBackupPath, installedAddinPath);
  }
  rmSync(installedBackupPath, { force: true });
}
