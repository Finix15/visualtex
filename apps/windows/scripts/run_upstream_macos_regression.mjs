import { execFile } from "node:child_process";
import { readFile } from "node:fs/promises";
import { basename, resolve } from "node:path";
import process from "node:process";
import { promisify } from "node:util";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

const [sourceArgument, ...forwardedArguments] = process.argv.slice(2);
if (!sourceArgument) {
  throw new Error(
    "Usage: node scripts/run_upstream_macos_regression.mjs <script-path> [...arguments]",
  );
}

const gitSourcePrefix = "origin-main:";
const sourceFromGit = sourceArgument.startsWith(gitSourcePrefix);
const sourcePath = sourceFromGit
  ? sourceArgument.slice(gitSourcePrefix.length)
  : resolve(process.cwd(), sourceArgument);
const browserPath = resolveChromiumExecutable();
const profileName = `visualtex-upstream-${basename(sourcePath).replace(/\W+/g, "-")}`;
const browserProfile = createBrowserProfilePath(profileName);
const execFileAsync = promisify(execFile);
let source;
if (sourceFromGit) {
  const { stdout } = await execFileAsync(
    "git",
    ["show", `origin/main:${sourcePath}`],
    { cwd: process.cwd(), maxBuffer: 16 * 1024 * 1024 },
  );
  source = stdout;
} else {
  source = await readFile(sourcePath, "utf8");
}

const profilePattern = /const (chromeProfile|profile) = `\/tmp\/[^`]+`;/;
const browserPattern = /const chromePath = "[^"]+";/;
if (!profilePattern.test(source) || !browserPattern.test(source)) {
  throw new Error(
    `The upstream regression launcher declarations were not found in ${sourcePath}`,
  );
}

source = source.replace(
  profilePattern,
  (_match, variableName) =>
    `const ${variableName} = ${JSON.stringify(browserProfile)};`,
);
source = source.replace(
  browserPattern,
  `const chromePath = ${JSON.stringify(browserPath)};`,
);
source = source.replace(
  /await client\.send\("Page\.enable"\);\r?\n(?!\s*await client\.send\("Page\.navigate")/g,
  'await client.send("Page.enable");\n  await client.send("Page.navigate", { url: baseUrl });\n  await sleep(700);\n',
);
source = source.replace(
  /if \(!field\?\.isConnected\) return \{ ready: false \};\r?\n(\s*)field\.setValue\("", \{/g,
  (_match, indent) =>
    `if (!field?.isConnected) return { ready: false };\n${indent}field.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", code: "Escape", bubbles: true, composed: true }));\n${indent}field.executeCommand(["complete", "reject"]);\n${indent}field.mode = "math";\n${indent}field.setValue("", {`,
);
source = source.replace(
  'return { ready: field.isConnected && field.value === "" };',
  'return { ready: field.isConnected && field.value === "", value: field.value, connected: field.isConnected, lineCount: document.querySelectorAll("math-field").length, activeElement: document.activeElement?.tagName ?? "" };',
);
source = source.replace(
  'assert.equal(enabledSuperscript.inScript, false);',
  'assert.equal(enabledSuperscript.inScript, false, JSON.stringify(enabledSuperscript));',
);
source = source.replace(
  "    await rm(chromeProfile, { recursive: true, force: true });",
  "    try { await rm(chromeProfile, { recursive: true, force: true, maxRetries: 4, retryDelay: 150 }); } catch (error) { if (process.platform !== 'win32' || error?.code !== 'EBUSY') throw error; }",
);
source = source.replace(
  "    await rm(profile, { recursive: true, force: true });",
  "    try { await rm(profile, { recursive: true, force: true, maxRetries: 4, retryDelay: 150 }); } catch (error) { if (process.platform !== 'win32' || error?.code !== 'EBUSY') throw error; }",
);
source = source.replace(
  'localStorage.setItem("visualtex.onboarding.v3.completed", "true");',
  'localStorage.setItem("visualtex.onboarding.v3.completed", "true");\n        localStorage.setItem("visualtex.onboarding.windows.desktop.v1.1.0.completed", "true");',
);
if (basename(sourcePath) === "auto_escape_regression.mjs") {
  source = source.replace(
    /    await evaluate\(`document\.querySelector\("\.canvas-input-behavior-trigger"\)\.click\(\)`\);\r?\n    await evaluate\(`document\.querySelector\("\.export-menu-trigger"\)\.click\(\)`\);[\s\S]*?    assert\.match\(await typeText\(">="\)/,
    '    assert.match(await typeText(">=")',
  );
}
if (forwardedArguments.includes("auto-exit-switch")) {
  source = source.replace(
    "    await configure();",
    `    await configure({ autoExitSuperscript: true, autoExitSubscript: false });
    await reload();
    await prepareEmptyField();
    await typeCharacter("x", "KeyX", 88);
    await typeCharacter("^", "Digit6", 54);
    await typeCharacter("a", "KeyA", 65);
    const switchedSuperscript = await readState();
    console.log(JSON.stringify(switchedSuperscript));
    assert.equal(switchedSuperscript.value, "x^{a}");
    assert.equal(switchedSuperscript.inScript, false, JSON.stringify(switchedSuperscript));
    assert.equal(switchedSuperscript.position, switchedSuperscript.lastOffset);
    console.log("Auto-exit setting switch regression passed");
    return;`,
  );
}

process.argv = [process.execPath, sourcePath, ...forwardedArguments];
const encodedSource = Buffer.from(source, "utf8").toString("base64");
await import(`data:text/javascript;base64,${encodedSource}`);
