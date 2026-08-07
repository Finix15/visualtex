import { execFileSync, spawnSync } from "node:child_process";
import {
  existsSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { decodeFormulaMetadata } from "../src/office/shared/formulaMetadata.ts";
import { normalizeFormulaEditorDocument } from "../src/office/shared/formulaEditorDocument.ts";

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const runtimeRoot = join(
  homedir(),
  "Library/Application Scripts/com.microsoft.Powerpoint/VisualTeXRuntime",
);
const sessionsRoot = join(runtimeRoot, "OfficeSessions");
const installedAddinPath = join(
  homedir(),
  "Library/Group Containers/UBF8T346G9.Office/VisualTeX/OfficeAddins/VisualTeX.ppam",
);
const resultPath = join(
  homedir(),
  "Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch/powerpoint-office-performance.json",
);
const hangTracePath = "/tmp/visualtex-powerpoint-hang.txt";
const visualTeXAppPath = join(
  repositoryRoot,
  "src-tauri/target/release/bundle/macos/VisualTeX.app",
);
const editorReadyFile = "editor-ready.json";
const editorPerformanceFile = "editor-performance.jsonl";
const sleep = (ms) => new Promise((resolvePromise) => setTimeout(resolvePromise, ms));

function run(program, args, timeout = 60_000) {
  return execFileSync(program, args, {
    encoding: "utf8",
    timeout,
    maxBuffer: 32 * 1024 * 1024,
  }).trim();
}

function bestEffort(program, args, timeout = 20_000) {
  try {
    return run(program, args, timeout);
  } catch {
    return "";
  }
}

function runAppleScript(lines, timeout = 60_000) {
  return run(
    "/usr/bin/osascript",
    lines.flatMap((line) => ["-e", line]),
    timeout,
  );
}

function currentSessionIds() {
  if (!existsSync(sessionsRoot)) return new Set();
  return new Set(
    readdirSync(sessionsRoot, { withFileTypes: true })
      .filter((entry) => entry.isDirectory())
      .map((entry) => entry.name),
  );
}

function requestForSession(sessionId) {
  const requestPath = join(sessionsRoot, sessionId, "request.json");
  if (!existsSync(requestPath)) return null;
  try {
    return JSON.parse(readFileSync(requestPath, "utf8"));
  } catch {
    return null;
  }
}

async function waitForNewSession(before, operation, formulaId, timeoutMs = 20_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    for (const sessionId of currentSessionIds()) {
      if (before.has(sessionId)) continue;
      const request = requestForSession(sessionId);
      if (
        request?.host === "powerpoint" &&
        request?.mode === operation &&
        (!formulaId || request.formulaId === formulaId)
      ) {
        return { sessionId, request };
      }
    }
    await sleep(25);
  }
  throw new Error(`PowerPoint did not create a ${operation} VisualTeX Session`);
}

async function waitForEditorReady(sessionId, timeoutMs = 10_000) {
  const readyPath = join(sessionsRoot, sessionId, editorReadyFile);
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (existsSync(readyPath)) {
      const marker = JSON.parse(readFileSync(readyPath, "utf8"));
      if (marker.sessionId === sessionId) return marker;
    }
    await sleep(20);
  }
  throw new Error(`PowerPoint editor did not become ready for ${sessionId}`);
}

function editorPerformanceRecords(sessionId) {
  const performancePath = join(
    sessionsRoot,
    sessionId,
    editorPerformanceFile,
  );
  if (!existsSync(performancePath)) return [];
  return readFileSync(performancePath, "utf8")
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => JSON.parse(line));
}

function visualTeXEditorWindowBounds() {
  const raw = runAppleScript([
    'tell application "System Events"',
    'tell process "visualtex"',
    'if (count of windows) is 0 then error "VisualTeX editor window is missing"',
    'set editorWindow to window 1',
    'set windowPosition to position of editorWindow',
    'set windowSize to size of editorWindow',
    'return (item 1 of windowPosition as text) & "|" & (item 2 of windowPosition as text) & "|" & (item 1 of windowSize as text) & "|" & (item 2 of windowSize as text)',
    "end tell",
    "end tell",
  ]);
  const [left, top, width, height] = raw.split("|").map(Number);
  if (
    [left, top, width, height].some((value) => !Number.isFinite(value)) ||
    width < 600 ||
    height < 500
  ) {
    throw new Error(`VisualTeX returned invalid editor bounds: ${raw}`);
  }
  return { left, top, width, height };
}

function focusedVisualTeXElement() {
  return runAppleScript([
    'tell application "System Events"',
    'tell process "visualtex"',
    'set focusedElement to value of attribute "AXFocusedUIElement"',
    'set roleValue to ""',
    'set subroleValue to ""',
    'set nameValue to ""',
    'set descriptionValue to ""',
    'try',
    'set roleValue to role of focusedElement as text',
    'end try',
    'try',
    'set subroleValue to subrole of focusedElement as text',
    'end try',
    'try',
    'set nameValue to name of focusedElement as text',
    'end try',
    'try',
    'set descriptionValue to description of focusedElement as text',
    'end try',
    'return roleValue & "|" & subroleValue & "|" & nameValue & "|" & descriptionValue',
    "end tell",
    "end tell",
  ]);
}

async function replaceActiveFormula(latex) {
  const clipboard = spawnSync("/usr/bin/pbcopy", [], {
    input: latex,
    encoding: "utf8",
    timeout: 5_000,
  });
  if (clipboard.status !== 0) {
    throw new Error(clipboard.stderr || "Unable to prepare the formula clipboard");
  }
  console.log(`POWERPOINT_FOCUS_BEFORE|${focusedVisualTeXElement()}`);
  const result = runAppleScript([
    'tell application "System Events"',
    'tell process "visualtex"',
    "set visible to true",
    "set frontmost to true",
    "delay 0.1",
    "set editorWindow to missing value",
    "set editorItems to {}",
    "repeat with candidateWindow in windows",
    "set candidateItems to entire contents of candidateWindow",
    "if (count of candidateItems) > 100 then",
    "set editorWindow to candidateWindow",
    "set editorItems to candidateItems",
    "exit repeat",
    "end if",
    "end repeat",
    'if editorWindow is missing value then error "The active VisualTeX Office editor was not found."',
    "set formulaField to missing value",
    "repeat with itemIndex from 1 to count of editorItems",
    "set candidateElement to item itemIndex of editorItems",
    "try",
    'if (role of candidateElement as text) is "AXTextField" then',
    "set formulaField to candidateElement",
    "exit repeat",
    "end if",
    "end try",
    "end repeat",
    'if formulaField is missing value then error "The VisualTeX formula input field was not found."',
    "set fieldPosition to position of formulaField",
    "set fieldSize to size of formulaField",
    "set clickX to (item 1 of fieldPosition) + ((item 1 of fieldSize) / 2)",
    "set clickY to (item 2 of fieldPosition) + ((item 2 of fieldSize) / 2)",
    "click at {clickX, clickY}",
    "delay 0.15",
    'keystroke "a" using {command down}',
    "delay 0.05",
    'keystroke "v" using {command down}',
    "delay 0.2",
    "set primaryButton to item 23 of editorItems",
    'if (role of primaryButton as text) is not "AXButton" then error "The VisualTeX primary action moved unexpectedly."',
    'if enabled of primaryButton is false then error "The VisualTeX primary action stayed disabled after formula input."',
    'return "REPLACED"',
    "end tell",
    "end tell",
  ]);
  if (result.trim() !== "REPLACED") {
    throw new Error(`Unexpected VisualTeX formula replacement result: ${result}`);
  }
  await sleep(100);
  console.log(`POWERPOINT_FOCUS_AFTER|${focusedVisualTeXElement()}`);
}

function selectFormulaShape(formulaId, slideIndex) {
  const shapeName = `VisualTeX_${formulaId}`;
  runAppleScript([
    'tell application "Microsoft PowerPoint"',
    `set currentSlide to slide ${Number(slideIndex)} of active presentation`,
    "go to slide (view of active window) number (slide index of currentSlide)",
    `if not (exists shape ${JSON.stringify(shapeName)} of currentSlide) then error "The VisualTeX formula shape is missing"`,
    `select shape ${JSON.stringify(shapeName)} of currentSlide`,
    "end tell",
  ]);
}

function processIds(processName) {
  return bestEffort("/usr/bin/pgrep", ["-x", processName])
    .split(/\s+/)
    .map((value) => Number.parseInt(value, 10))
    .filter((value) => Number.isInteger(value) && value > 0);
}

function sampleApplyHang(sessionId) {
  const sections = [`sessionId=${sessionId}`, `capturedEpochMs=${Date.now()}`];
  for (const processName of ["visualtex", "Microsoft PowerPoint", "osascript"]) {
    for (const pid of processIds(processName)) {
      let sample = "";
      try {
        sample = run("/usr/bin/sample", [String(pid), "1", "1"], 15_000);
      } catch (error) {
        sample = error instanceof Error ? error.message : String(error);
      }
      sections.push(`\n===== ${processName} pid=${pid} =====\n${sample}`);
    }
  }
  const combined = `${sections.join("\n")}\n`;
  writeFileSync(hangTracePath, combined, { mode: 0o600 });
  const symbols = combined
    .split(/\r?\n/)
    .map((line) =>
      line.match(/visualtex_lib::office::macos_offline::([^\s]+)/)?.[1],
    )
    .filter(Boolean);
  writeFileSync(
    "/tmp/visualtex-powerpoint-hang-symbols.txt",
    `${[...new Set(symbols)].join("\n")}\n`,
    { mode: 0o600 },
  );
}

async function applyActiveFormula(sessionId, timeoutMs = 30_000) {
  const startedEpochMs = Date.now();
  const clickResult = runAppleScript([
    'tell application "System Events"',
    'tell process "visualtex"',
    "set frontmost to true",
    "delay 0.05",
    "set editorItems to {}",
    "repeat with candidateWindow in windows",
    "set candidateItems to entire contents of candidateWindow",
    "if (count of candidateItems) > 100 then",
    "set editorItems to candidateItems",
    "exit repeat",
    "end if",
    "end repeat",
    'if (count of editorItems) <= 100 then error "The active VisualTeX Office editor was not found."',
    "set primaryButton to item 23 of editorItems",
    'if (role of primaryButton as text) is not "AXButton" then error "The VisualTeX primary action moved unexpectedly."',
    'if enabled of primaryButton is false then error "The VisualTeX primary action is disabled."',
    'perform action "AXPress" of primaryButton',
    'return "PRESSED"',
    "end tell",
    "end tell",
  ]);
  if (clickResult.trim() !== "PRESSED") {
    throw new Error(`Unexpected VisualTeX Apply press result: ${clickResult}`);
  }
  const deadline = Date.now() + timeoutMs;
  let hangSampled = false;
  while (Date.now() < deadline) {
    const record = editorPerformanceRecords(sessionId).find(
      (candidate) => candidate.stage === "apply-backend-complete",
    );
    if (record) {
      const completedEpochMs = Number(record.epochMs);
      return {
        startedEpochMs,
        completedEpochMs,
        clickToOfficeCompleteMs: completedEpochMs - startedEpochMs,
        backendElapsedMs: Number(record.elapsedMs),
        records: editorPerformanceRecords(sessionId).filter((candidate) =>
          String(candidate.stage).startsWith("apply-"),
        ),
      };
    }
    if (!hangSampled && Date.now() - startedEpochMs >= 5_000) {
      hangSampled = true;
      sampleApplyHang(sessionId);
    }
    await sleep(20);
  }
  throw new Error(`PowerPoint Apply did not complete for ${sessionId}`);
}

function enableMacrosIfPrompted() {
  bestEffort("/usr/bin/osascript", [
    "-e",
    'tell application "System Events"',
    "-e",
    'if exists process "Microsoft PowerPoint" then',
    "-e",
    'tell process "Microsoft PowerPoint"',
    "-e",
    "repeat with candidateWindow in windows",
    "-e",
    'repeat with buttonName in {"启用宏", "Enable Macros"}',
    "-e",
    "try",
    "-e",
    'if exists button (buttonName as text) of candidateWindow then click button (buttonName as text) of candidateWindow',
    "-e",
    "end try",
    "-e",
    "end repeat",
    "-e",
    "end repeat",
    "-e",
    "end tell",
    "-e",
    "end if",
    "-e",
    "end tell",
  ]);
}

async function waitForPowerPointUi(timeoutMs = 30_000) {
  const deadline = Date.now() + timeoutMs;
  let lastError = "";
  while (Date.now() < deadline) {
    try {
      const state = runAppleScript([
        'tell application "Microsoft PowerPoint" to activate',
        'tell application "System Events"',
        'tell process "Microsoft PowerPoint"',
        "set visible to true",
        "set frontmost to true",
        'if (count of menu bars) = 0 then error "PowerPoint UI is not ready"',
        'return "READY"',
        "end tell",
        "end tell",
      ], 5_000);
      if (state === "READY") return;
    } catch (error) {
      lastError = error instanceof Error ? error.message : String(error);
    }
    await sleep(500);
  }
  throw new Error(`PowerPoint did not become UI-ready: ${lastError}`);
}

function createTestPresentation() {
  return runAppleScript([
    'tell application "Microsoft PowerPoint"',
    "activate",
    "set testPresentation to make new presentation",
    "if (count of slides of testPresentation) is 0 then make new slide at end of testPresentation",
    "return name of testPresentation as text",
    "end tell",
  ]);
}

function runPowerPointMacro(name) {
  return runAppleScript([
    'tell application "Microsoft PowerPoint"',
    `run VB macro macro name ${JSON.stringify(name)} list of parameters {}`,
    "end tell",
  ]);
}

function formulaDocumentFromEditRequest(request) {
  const metadata = decodeFormulaMetadata(request?.encodedMetadata ?? "");
  const document = metadata
    ? normalizeFormulaEditorDocument(metadata.lines, metadata.codeFormat)
    : null;
  if (!metadata || !document) {
    throw new Error("PowerPoint edit request did not contain valid formula metadata");
  }
  return { metadata, document };
}

async function ensureAddinLoaded() {
  try {
    runPowerPointMacro("Auto_Open");
    return;
  } catch {
    if (!existsSync(installedAddinPath)) {
      throw new Error(`Installed PowerPoint add-in is missing: ${installedAddinPath}`);
    }
    run("/usr/bin/open", ["-b", "com.microsoft.Powerpoint", installedAddinPath]);
    await sleep(2_000);
    enableMacrosIfPrompted();
    await sleep(500);
    runPowerPointMacro("Auto_Open");
  }
}

async function main() {
  rmSync(resultPath, { force: true });
  bestEffort("/usr/bin/killall", ["Microsoft PowerPoint"]);
  if (processIds("visualtex").length === 0) {
    run("/usr/bin/open", ["-gj", visualTeXAppPath, "--args", "--office-background"]);
    await sleep(8_000);
  }
  run("/usr/bin/open", ["-b", "com.microsoft.Powerpoint"]);
  await waitForPowerPointUi();
  await ensureAddinLoaded();
  const presentationName = createTestPresentation();

  const beforeCreate = currentSessionIds();
  runPowerPointMacro("VisualTeX_NewFormula");
  const created = await waitForNewSession(beforeCreate, "create");
  const createReady = await waitForEditorReady(created.sessionId);
  const createLatex = String.raw`E=mc^2`;
  await replaceActiveFormula(createLatex);
  const createApply = await applyActiveFormula(created.sessionId);
  if (createApply.clickToOfficeCompleteMs > 1_500) {
    throw new Error(
      `PowerPoint create Apply exceeded 1500 ms: ${JSON.stringify(createApply)}`,
    );
  }
  await sleep(150);
  selectFormulaShape(
    created.request.formulaId,
    created.request.powerPoint.slideIndex,
  );

  const beforeEdit = currentSessionIds();
  const editInvokedEpochMs = Date.now();
  runPowerPointMacro("VisualTeX_DoubleClickEditSelected");
  const edited = await waitForNewSession(
    beforeEdit,
    "edit",
    created.request.formulaId,
  );
  const editReady = await waitForEditorReady(edited.sessionId);
  const editInvokeToReadyMs = Number(editReady.epochMs) - editInvokedEpochMs;
  if (editInvokeToReadyMs > 1_000) {
    throw new Error(`PowerPoint edit opening exceeded 1000 ms: ${editInvokeToReadyMs}`);
  }
  const createdState = formulaDocumentFromEditRequest(edited.request);
  if (
    createdState.metadata.formulaId !== created.request.formulaId ||
    createdState.document.lines[0]?.latex !== createLatex
  ) {
    throw new Error("PowerPoint did not persist the created formula metadata");
  }

  const editLatex = String.raw`E^2=p^2c^2+m^2c^4`;
  await replaceActiveFormula(editLatex);
  const editApply = await applyActiveFormula(edited.sessionId);
  if (editApply.clickToOfficeCompleteMs > 1_500) {
    throw new Error(
      `PowerPoint edit Apply exceeded 1500 ms: ${JSON.stringify(editApply)}`,
    );
  }
  await sleep(150);
  selectFormulaShape(
    created.request.formulaId,
    created.request.powerPoint.slideIndex,
  );
  const beforeVerify = currentSessionIds();
  runPowerPointMacro("VisualTeX_DoubleClickEditSelected");
  const verified = await waitForNewSession(
    beforeVerify,
    "edit",
    created.request.formulaId,
  );
  const verifyReady = await waitForEditorReady(verified.sessionId);
  const editedState = formulaDocumentFromEditRequest(verified.request);
  if (
    editedState.metadata.formulaId !== created.request.formulaId ||
    editedState.document.lines[0]?.latex !== editLatex
  ) {
    throw new Error("PowerPoint did not persist the edited formula metadata");
  }

  const result = {
    status: "PASS",
    revision: "powerpoint-office-performance-20260801-r4",
    presentationName,
    formulaId: created.request.formulaId,
    create: {
      sessionId: created.sessionId,
      requestToReadyMs: Number(createReady.requestToReadyMs),
      showFocusMs: Number(createReady.showFocusMs),
      apply: createApply,
      persistedLatex: createdState.document.lines[0].latex,
    },
    edit: {
      sessionId: edited.sessionId,
      requestToReadyMs: Number(editReady.requestToReadyMs),
      invokeToReadyMs: editInvokeToReadyMs,
      showFocusMs: Number(editReady.showFocusMs),
      apply: editApply,
      persistedLatex: editedState.document.lines[0].latex,
      verificationSessionId: verified.sessionId,
      verificationRequestToReadyMs: Number(verifyReady.requestToReadyMs),
    },
  };
  writeFileSync(resultPath, `${JSON.stringify(result, null, 2)}\n`, {
    mode: 0o600,
  });
  console.log(JSON.stringify(result, null, 2));
}

try {
  await main();
} finally {
  bestEffort("/usr/bin/osascript", [
    "-e",
    'tell application "Microsoft PowerPoint" to quit',
  ]);
  bestEffort("/usr/bin/killall", ["Microsoft PowerPoint"]);
}
