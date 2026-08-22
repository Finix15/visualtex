import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, unlinkSync } from "node:fs";
import { homedir } from "node:os";
import { join, resolve } from "node:path";

const root = resolve(import.meta.dirname, "..");
const office = join(root, "office", "macos-offline");
const scratch = join(homedir(), "Library", "Group Containers", "UBF8T346G9.Office", "VisualTeX", "Scratch");
// PowerPoint's macOS sandbox grants file access per selected location. /private/tmp
// is used as the VBE build staging area so the save can be completed unattended
// after the user has granted access once; the packager copies the result later.
const output = "/private/tmp/VisualTeXPowerPointBuild.pptm";
const modules = [
  join(office, "shared", "VTProtocol.bas"),
  join(office, "shared", "VTOfficePaths.bas"),
  join(office, "shared", "VTMetadata.bas"),
  join(office, "shared", "VTLauncher.bas"),
  join(office, "shared", "VTErrorHandling.bas"),
  join(office, "powerpoint", "VTPowerPointAdapter.bas"),
  join(office, "powerpoint", "VTPowerPointEvents.cls"),
  join(office, "powerpoint", "VTRibbonCallbacks.bas"),
];

function osa(lines, timeout = 120_000) {
  const args = lines.flatMap((line) => ["-e", line]);
  return execFileSync("/usr/bin/osascript", args, { encoding: "utf8", timeout });
}
function wait(ms) { Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms); }

mkdirSync(scratch, { recursive: true });
if (existsSync(output)) unlinkSync(output);
try { osa(['tell application "Microsoft PowerPoint" to quit']); } catch {}
wait(1500);
execFileSync("/usr/bin/open", ["-gj", "-a", "Microsoft PowerPoint"]);
for (let attempt = 0; attempt < 80; attempt += 1) {
  const ready = osa([
    'tell application "System Events"',
    'if exists process "Microsoft PowerPoint" then',
    'tell process "Microsoft PowerPoint"',
    'set visible to true', 'set frontmost to true',
    'if (count of menu bars) > 0 then return "READY"',
    'end tell', 'end if', 'end tell',
  ], 5000).trim();
  if (ready === "READY") break;
  wait(500);
}
osa([
  'tell application "Microsoft PowerPoint"',
  'set buildPresentation to make new presentation',
  `save buildPresentation in POSIX file ${JSON.stringify(output)} as save as Open XML presentation macro enabled`,
  'activate', 'end tell',
]);
wait(1800);
osa([
  'tell application "System Events"', 'tell process "Microsoft PowerPoint"',
  'set frontmost to true', 'set openedEditor to false',
  'repeat with toolsName in {"工具", "Tools"}',
  'if openedEditor is false then', 'try',
  'set toolsMenu to menu 1 of menu bar item (toolsName as text) of menu bar 1',
  'repeat with macroName in {"宏", "Macro"}', 'try',
  'set macroMenu to menu 1 of menu item (macroName as text) of toolsMenu',
  'repeat with editorName in {"Visual Basic 编辑器", "Visual Basic Editor"}', 'try',
  'click menu item (editorName as text) of macroMenu', 'set openedEditor to true', 'exit repeat',
  'end try', 'end repeat', 'end try', 'end repeat', 'end try', 'end if', 'end repeat',
  'if openedEditor is false then error "Unable to open the PowerPoint Visual Basic Editor."',
  'end tell', 'end tell',
]);
wait(1800);

for (const modulePath of modules) {
  osa([
    'tell application "System Events"', 'tell process "Microsoft PowerPoint"',
    'set frontmost to true',
    'set vbeWindow to first window whose name contains "Microsoft Visual Basic"',
    'perform action "AXRaise" of vbeWindow', 'set openedImport to false',
    'repeat with fileName in {"文件", "File"}', 'if openedImport is false then', 'try',
    'set fileMenu to menu 1 of menu bar item (fileName as text) of menu bar 1',
    'repeat with candidateItem in menu items of fileMenu',
    'set candidateName to name of candidateItem as text',
    'if candidateName starts with "导入文件" or candidateName starts with "Import File" then',
    'click candidateItem', 'set openedImport to true', 'exit repeat', 'end if',
    'end repeat', 'end try', 'end if', 'end repeat',
    'if openedImport is false then error "Unable to find File > Import File."',
    'delay 0.8', 'keystroke "g" using {command down, shift down}', 'delay 0.5',
    'set pathField to value of attribute "AXFocusedUIElement"',
    `set value of pathField to ${JSON.stringify(modulePath)}`,
    'key code 36', 'delay 0.8', 'key code 36', 'delay 1.3',
    'end tell', 'end tell',
  ]);
}

const compileMessage = osa([
  'tell application "System Events"', 'tell process "Microsoft PowerPoint"',
  'set frontmost to true',
  'set vbeWindow to first window whose name contains "Microsoft Visual Basic"',
  'perform action "AXRaise" of vbeWindow', 'set startedCompile to false',
  'repeat with debugName in {"调试", "Debug"}', 'if startedCompile is false then', 'try',
  'set debugMenu to menu 1 of menu bar item (debugName as text) of menu bar 1',
  'repeat with candidateItem in menu items of debugMenu',
  'set candidateName to name of candidateItem as text',
  'if candidateName starts with "编译 " or candidateName starts with "Compile " then',
  'click candidateItem', 'set startedCompile to true', 'exit repeat', 'end if',
  'end repeat', 'end try', 'end if', 'end repeat',
  'if startedCompile is false then error "Unable to find Debug > Compile Project."',
  'delay 2', 'set failureText to ""',
  'repeat with candidateWindow in windows', 'try',
  'if description of candidateWindow is "警告" or description of candidateWindow is "Warning" then set failureText to value of every static text of candidateWindow as text',
  'end try', 'end repeat', 'return failureText', 'end tell', 'end tell',
]).trim();
if (compileMessage) throw new Error(`PowerPoint VBE compile failed: ${compileMessage}`);

osa(['tell application "Microsoft PowerPoint"', 'save active presentation', 'end tell']);
wait(2000);
osa(['tell application "Microsoft PowerPoint" to quit']);
if (!existsSync(output)) throw new Error(`PowerPoint did not create ${output}`);
process.stdout.write(`Rebuilt and VBE-compiled ${output}\n`);
