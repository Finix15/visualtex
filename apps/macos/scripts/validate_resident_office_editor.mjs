import { execFileSync } from "node:child_process";
import { randomUUID } from "node:crypto";
import {
  chmodSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const scriptDirectory = dirname(fileURLToPath(import.meta.url));
const appRoot = resolve(scriptDirectory, "..");
const validationApp = join(
  appRoot,
  "src-tauri/target/release/bundle/macos/VisualTeX Validation.app",
);
const validationBinary = join(validationApp, "Contents/MacOS/visualtex");
const validationBundleId = "com.visualtex.studio.validation";
const officeSessionsRoot = join(
  homedir(),
  "Library/Application Scripts/com.microsoft.Powerpoint/VisualTeXRuntime/OfficeSessions",
);
const validationOfficeRoot = join(
  homedir(),
  "Library/Application Support/com.visualtex.studio.validation/office",
);
const validationSessionsRoot = join(validationOfficeRoot, "sessions");
const validationPrewarmMarkers = {
  word: join(validationOfficeRoot, "word-editor-prewarmed.json"),
  powerpoint: join(validationOfficeRoot, "powerpoint-editor-prewarmed.json"),
};
const validationPrewarmDiagnostics = {
  word: join(validationOfficeRoot, "word-editor-prewarm-diagnostic.json"),
  powerpoint: join(validationOfficeRoot, "powerpoint-editor-prewarm-diagnostic.json"),
};
const resultPath = join(
  homedir(),
  "Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch/resident-editor-validation.json",
);
const sleep = (milliseconds) =>
  new Promise((resolvePromise) => setTimeout(resolvePromise, milliseconds));

function run(program, args, timeout = 30_000) {
  return execFileSync(program, args, {
    encoding: "utf8",
    timeout,
    maxBuffer: 16 * 1024 * 1024,
  }).trim();
}

function runAppleScript(lines, timeout = 30_000) {
  return run(
    "/usr/bin/osascript",
    lines.flatMap((line) => ["-e", line]),
    timeout,
  );
}

function validationPids() {
  try {
    return run("/usr/bin/pgrep", ["-f", validationBinary])
      .split(/\s+/)
      .map((value) => Number.parseInt(value, 10))
      .filter((value) => Number.isInteger(value) && value > 0);
  } catch {
    return [];
  }
}

async function stopValidationApp() {
  for (const pid of validationPids()) {
    try {
      process.kill(pid, "SIGTERM");
    } catch {
      // The process may have exited between pgrep and kill.
    }
  }
  const deadline = Date.now() + 5_000;
  while (Date.now() < deadline && validationPids().length > 0) {
    await sleep(50);
  }
  for (const pid of validationPids()) {
    try {
      process.kill(pid, "SIGKILL");
    } catch {
      // Best-effort cleanup for the isolated validation process only.
    }
  }
}

function findSourceRequest() {
  const requestedSession = process.env.VISUALTEX_VALIDATION_SOURCE_SESSION?.trim();
  const candidates = requestedSession
    ? [join(officeSessionsRoot, requestedSession, "request.json")]
    : readdirSync(officeSessionsRoot, { withFileTypes: true })
        .filter((entry) => entry.isDirectory())
        .map((entry) => join(officeSessionsRoot, entry.name, "request.json"))
        .filter(existsSync)
        .sort((left, right) => statSync(right).mtimeMs - statSync(left).mtimeMs);

  for (const path of candidates) {
    try {
      const request = JSON.parse(readFileSync(path, "utf8"));
      if (
        request?.host === "powerpoint" &&
        request?.mode === "edit" &&
        typeof request?.formulaId === "string" &&
        typeof request?.encodedMetadata === "string" &&
        request.encodedMetadata.startsWith("visualtex:v1:deflate:")
      ) {
        return { path, request };
      }
    } catch {
      // Ignore incomplete or stale runtime directories.
    }
  }
  throw new Error("No real editable PowerPoint VisualTeX request was found.");
}

function frontToBackWindowStack() {
  const source = `
import CoreGraphics
let windows = CGWindowListCopyWindowInfo([.optionOnScreenOnly, .excludeDesktopElements], kCGNullWindowID) as? [[String: Any]] ?? []
for (index, window) in windows.enumerated() {
  let owner = window[kCGWindowOwnerName as String] as? String ?? ""
  let name = window[kCGWindowName as String] as? String ?? ""
  let layer = window[kCGWindowLayer as String] as? Int ?? -1
  let pid = window[kCGWindowOwnerPID as String] as? Int ?? -1
  print("\\(index)|\\(pid)|\\(layer)|\\(owner)|\\(name)")
}
`;
  return run("/usr/bin/swift", ["-e", source], 15_000)
    .split("\n")
    .filter(Boolean)
    .map((line) => {
      const [index, pid, layer, owner, ...nameParts] = line.split("|");
      return {
        index: Number.parseInt(index, 10),
        pid: Number.parseInt(pid, 10),
        layer: Number.parseInt(layer, 10),
        owner,
        name: nameParts.join("|"),
      };
    });
}

function validationWindowStatus() {
  const raw = runAppleScript([
    'tell application "System Events"',
    `set matchingProcesses to every application process whose bundle identifier is ${JSON.stringify(validationBundleId)}`,
    'if (count of matchingProcesses) is 0 then return "0|false|"',
    "set validationProcess to item 1 of matchingProcesses",
    "set windowCount to count of windows of validationProcess",
    "set windowTitles to {}",
    "repeat with candidateWindow in windows of validationProcess",
    "try",
    "set end of windowTitles to name of candidateWindow as text",
    "on error",
    'set end of windowTitles to ""',
    "end try",
    "end repeat",
    'set AppleScript\'s text item delimiters to "<<<VT_WINDOW>>>"',
    "set joinedTitles to windowTitles as text",
    'set AppleScript\'s text item delimiters to ""',
    "return (windowCount as text) & \"|\" & (frontmost of validationProcess as text) & \"|\" & joinedTitles",
    "end tell",
  ]);
  const [windowCount, frontmost, ...titleParts] = raw.split("|");
  return {
    windowCount: Number.parseInt(windowCount, 10) || 0,
    frontmost: frontmost === "true",
    windowTitles: titleParts
      .join("|")
      .split("<<<VT_WINDOW>>>")
      .filter(Boolean),
  };
}

function activatePowerPointForBoundaryTest() {
  try {
    runAppleScript([
      'tell application "Microsoft PowerPoint" to activate',
      'tell application "System Events"',
      'tell process "Microsoft PowerPoint"',
      "set frontmost to true",
      "end tell",
      "end tell",
    ]);
    return true;
  } catch {
    return false;
  }
}

async function waitForValidationForeground(timeoutMs = 2_000) {
  const deadline = Date.now() + timeoutMs;
  let status = validationWindowStatus();
  while (Date.now() < deadline) {
    status = validationWindowStatus();
    if (
      status.frontmost &&
      status.windowTitles.some((title) => title.includes("PowerPoint"))
    ) {
      return status;
    }
    await sleep(25);
  }
  return status;
}

function readPrewarmMarker(path) {
  if (!existsSync(path)) return null;
  try {
    return JSON.parse(readFileSync(path, "utf8"));
  } catch {
    return null;
  }
}

async function waitForPrewarm(timeoutMs = 45_000) {
  const deadline = Date.now() + timeoutMs;
  let lastStatus = {
    windows: validationWindowStatus(),
    word: null,
    powerpoint: null,
    diagnostics: {
      word: readPrewarmMarker(validationPrewarmDiagnostics.word),
      powerpoint: readPrewarmMarker(validationPrewarmDiagnostics.powerpoint),
    },
    pids: validationPids(),
  };
  while (Date.now() < deadline) {
    const pids = validationPids();
    const word = readPrewarmMarker(validationPrewarmMarkers.word);
    const powerpoint = readPrewarmMarker(validationPrewarmMarkers.powerpoint);
    const windows = validationWindowStatus();
    const diagnostics = {
      word: readPrewarmMarker(validationPrewarmDiagnostics.word),
      powerpoint: readPrewarmMarker(validationPrewarmDiagnostics.powerpoint),
    };
    lastStatus = { windows, word, powerpoint, diagnostics, pids };
    const markerMatchesProcess = (marker) =>
      marker?.schema === "visualtex-office-editor-prewarmed-v1" &&
      Number.isInteger(marker?.processId) &&
      pids.includes(marker.processId);
    if (markerMatchesProcess(word) && markerMatchesProcess(powerpoint)) {
      return lastStatus;
    }
    await sleep(100);
  }
  throw new Error(
    `Validation editor prewarm timed out: ${JSON.stringify(lastStatus)}`,
  );
}

async function waitForReady(sessionDirectory, sessionId, timeoutMs = 10_000) {
  const readyPath = join(sessionDirectory, "editor-ready.json");
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (existsSync(readyPath)) {
      const marker = JSON.parse(readFileSync(readyPath, "utf8"));
      if (marker.sessionId === sessionId) return marker;
    }
    await sleep(20);
  }
  throw new Error(`Validation editor did not become ready for ${sessionId}.`);
}

async function main() {
  if (!existsSync(validationBinary)) {
    throw new Error(`Validation application is missing: ${validationApp}`);
  }
  const source = findSourceRequest();
  const sessionId = randomUUID();
  const sessionDirectory = join(officeSessionsRoot, sessionId);
  const validationSessionDirectory = join(validationSessionsRoot, sessionId);
  const request = { ...source.request, sessionId };
  const result = {
    status: "STARTED",
    sourceRequest: source.path,
    sessionId,
    validationApp,
  };

  await stopValidationApp();
  for (const markerPath of [
    ...Object.values(validationPrewarmMarkers),
    ...Object.values(validationPrewarmDiagnostics),
  ]) {
    rmSync(markerPath, { force: true });
  }
  rmSync(sessionDirectory, { recursive: true, force: true });
  rmSync(validationSessionDirectory, { recursive: true, force: true });
  mkdirSync(sessionDirectory, { recursive: true, mode: 0o700 });
  const requestPath = join(sessionDirectory, "request.json");
  writeFileSync(requestPath, `${JSON.stringify(request)}\n`, { mode: 0o600 });
  chmodSync(requestPath, 0o600);

  try {
    run("/usr/bin/open", ["-na", validationApp, "--args", "--office-background"]);
    const prewarmStarted = Date.now();
    result.prewarm = await waitForPrewarm();
    result.prewarm.elapsedMs = Date.now() - prewarmStarted;

    result.powerPointBoundaryActivated = activatePowerPointForBoundaryTest();
    const invocationStarted = Date.now();
    run(validationBinary, [`visualtex://office/open?session=${sessionId}`], 15_000);
    const ready = await waitForReady(sessionDirectory, sessionId);
    result.status = "PASS";
    result.invokeToReadyMs = Date.now() - invocationStarted;
    result.ready = ready;
    result.window = await waitForValidationForeground();
    result.windowStack = frontToBackWindowStack().filter(
      (item) =>
        item.owner.includes("VisualTeX") ||
        item.owner.includes("PowerPoint") ||
        item.name.includes("VisualTeX Office Formula"),
    );
    const validationPid = result.prewarm.pids[0] ?? -1;
    const editorWindow = result.windowStack.find(
      (item) =>
        item.pid === validationPid &&
        item.name.includes("VisualTeX Office Formula") &&
        item.layer > 0,
    );
    const powerPointWindow = result.windowStack.find((item) =>
      item.owner.includes("Microsoft PowerPoint"),
    );
    result.editorAbovePowerPoint = Boolean(
      editorWindow &&
        powerPointWindow &&
        editorWindow.layer > powerPointWindow.layer &&
        editorWindow.index < powerPointWindow.index,
    );

    if (ready.showFocusMs > 300) {
      throw new Error(
        `Resident editor showFocusMs ${ready.showFocusMs.toFixed(1)} exceeded 300 ms.`,
      );
    }
    if (ready.contentReadyMs > 300) {
      throw new Error(
        `Resident editor contentReadyMs ${ready.contentReadyMs.toFixed(1)} exceeded 300 ms.`,
      );
    }
    if (!ready.windowVisible) {
      throw new Error(
        `Resident PowerPoint editor did not report visible after ready: ${JSON.stringify(ready)}`,
      );
    }
    if (!ready.windowFocused && !result.editorAbovePowerPoint) {
      throw new Error(
        `Resident PowerPoint editor was neither focused nor ordered above PowerPoint: ${JSON.stringify({ ready, windowStack: result.windowStack })}`,
      );
    }
  } catch (error) {
    result.status = "FAIL";
    result.error = error instanceof Error ? error.message : String(error);
    throw error;
  } finally {
    mkdirSync(dirname(resultPath), { recursive: true });
    writeFileSync(resultPath, `${JSON.stringify(result, null, 2)}\n`, {
      mode: 0o600,
    });
    await stopValidationApp();
    if (result.status === "PASS") {
      rmSync(sessionDirectory, { recursive: true, force: true });
      rmSync(validationSessionDirectory, { recursive: true, force: true });
    }
  }

  process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
}

main().catch((error) => {
  process.stderr.write(`${error instanceof Error ? error.stack : String(error)}\n`);
  process.exitCode = 1;
});
