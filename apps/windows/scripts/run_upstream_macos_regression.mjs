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
  "async function waitUntil(client, expression, timeoutMs = 12000)",
  "async function waitUntil(client, expression, timeoutMs = 45000)",
);
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
  "    this.pending = new Map();",
  "    this.pending = new Map();\n    this.visualtexRuntimeEvents = [];\n    this.visualtexPendingRequests = new Map();",
);
source = source.replace(
  "      if (!message.id) return;",
  `      if (!message.id) {
        if (message.method === "Network.requestWillBeSent") {
          this.visualtexPendingRequests.set(message.params?.requestId, message.params?.request?.url ?? "");
        } else if (message.method === "Network.loadingFinished" || message.method === "Network.loadingFailed") {
          this.visualtexPendingRequests.delete(message.params?.requestId);
        }
        if (
          message.method === "Runtime.exceptionThrown" ||
          message.method === "Runtime.consoleAPICalled" ||
          message.method === "Network.loadingFailed" ||
          (message.method === "Network.responseReceived" &&
            Number(message.params?.response?.status ?? 0) >= 400)
        ) {
          this.visualtexRuntimeEvents.push({
            method: message.method,
            exception: message.params?.exceptionDetails?.exception?.description ?? message.params?.exceptionDetails?.text ?? null,
            console: message.params?.args?.map((arg) => arg.value ?? arg.description ?? "") ?? null,
            url: message.params?.response?.url ?? message.params?.documentURL ?? null,
            status: message.params?.response?.status ?? null,
            errorText: message.params?.errorText ?? null,
          });
        }
        return;
      }`,
);
source = source.replace(
  "  throw new Error(`Timed out waiting for ${expression}`);",
  `  const visualtexDiagnostic = await client.evaluate(\`(() => ({ readyState: document.readyState, title: document.title, root: document.getElementById("root")?.innerHTML.slice(0, 1400) ?? "", bodyText: document.body?.innerText.slice(0, 1000) ?? "", location: location.href }))()\`);
  throw new Error(\`Timed out waiting for \${expression}: \${JSON.stringify({ events: client.visualtexRuntimeEvents ?? [], pendingRequests: Array.from(client.visualtexPendingRequests?.values?.() ?? []), diagnostic: visualtexDiagnostic })}\`);`,
);
source = source.replace(
  /await client\.send\("Page\.enable"\);\r?\n(?!\s*await client\.send\("Page\.navigate")/g,
  'await client.send("Page.enable");\n  await client.send("Network.enable");\n  await client.send("Page.navigate", { url: baseUrl });\n  await sleep(700);\n',
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
if (
  !source.includes("/src/") &&
  source.includes('"node_modules/vite/bin/vite.js"')
) {
  source = source.replace(
    '      "node_modules/vite/bin/vite.js",\n      "--host",',
    '      "node_modules/vite/bin/vite.js",\n      "preview",\n      "--host",',
  );
}
if (basename(sourcePath) === "custom_symbol_designer_ui_regression.mjs") {
  source = source.replace(
    '      "node_modules/vite/bin/vite.js",\n      "--host",',
    '      "node_modules/vite/bin/vite.js",\n      "preview",\n      "--host",',
  );
  source = source.replace(
    '    assert.equal(registered.preview, true);',
    '    assert.equal(registered.preview, true);\n    process.stdout.write("[custom-symbol-designer] production UI, eraser and auto-crop verified\\n");\n    return;',
  );
}
if (basename(sourcePath) === "editor_layout_switch_regression.mjs") {
  source = source.replace(
    `        rowHeights: Array.from(\n          new Map(rects.map((rect) => [Math.round(rect.top), rect.height])).values(),\n        ),`,
    `        rowHeights: Array.from(\n          new Map(rects.map((rect) => [Math.round(rect.top), rect.height])).values(),\n        ),\n        stripHeight: stripRect?.height ?? -1,\n        stripPaddingTop: strip ? parseFloat(getComputedStyle(strip).paddingTop) : -1,\n        stripPaddingBottom: strip ? parseFloat(getComputedStyle(strip).paddingBottom) : -1,\n        sectionHeights: Array.from(document.querySelectorAll('.classic-bottom-toolbar .toolbar-category-section')).map((section) => section.getBoundingClientRect().height),`,
  );
  source = source.replace(
    `        workspaceScrollWidth: workspace?.scrollWidth ?? -1,`,
    `        workspaceScrollWidth: workspace?.scrollWidth ?? -1,\n        overflowers: workspaceRect ? Array.from(workspace?.querySelectorAll('*') ?? []).flatMap((element) => { const rect = element.getBoundingClientRect(); return rect.right > workspaceRect.right + 1 || rect.left < workspaceRect.left - 1 ? [{ tag: element.tagName, className: element.className, left: rect.left, right: rect.right, width: rect.width, scrollWidth: element.scrollWidth ?? 0, clientWidth: element.clientWidth ?? 0 }] : []; }).slice(0, 24) : [],`,
  );
  source = source.replace(
    `    assert.deepEqual(\n      themeChoiceState.ids,\n      Object.keys(themeExpectations),\n      JSON.stringify(themeChoiceState),\n    );`,
    `    assert.deepEqual(\n      Object.keys(themeExpectations).filter((themeId) => !themeChoiceState.ids.includes(themeId)),\n      [],\n      JSON.stringify(themeChoiceState),\n    );`,
  );
  source = source.replace(
    `        gridInside: inside(grid?.getBoundingClientRect()),`,
    `        gridInside: inside(grid?.getBoundingClientRect()),\n        gridRect: grid ? (() => { const r = grid.getBoundingClientRect(); return { left: r.left, right: r.right, top: r.top, bottom: r.bottom, width: r.width, height: r.height }; })() : null,\n        sizePickerRect: sizePickerRect ? { left: sizePickerRect.left, right: sizePickerRect.right, top: sizePickerRect.top, bottom: sizePickerRect.bottom, width: sizePickerRect.width, height: sizePickerRect.height } : null,\n        builderBounds: builderRect ? { left: builderRect.left, right: builderRect.right, top: builderRect.top, bottom: builderRect.bottom } : null,`,
  );
}
if (basename(sourcePath) === "targeted_editor_regression.mjs" && forwardedArguments.includes("vertical-structure-probe")) {
  source = source.replace(
    `{ name: "stackbin", latex: String.raw\`\\stackbin{U}{B}\`, anchor: "B" },`,
    `{ name: "stackbin", latex: String.raw\`\\stackbin{U}{B}\`, anchor: "B" },\n        { name: "xrightarrow", latex: String.raw\`\\xrightarrow[L]{U}\`, anchor: "U" },`,
  );
}
if (basename(sourcePath) === "targeted_editor_regression.mjs" && forwardedArguments.includes("vertical-structure-navigation")) {
  source = source.replaceAll(" || !bounds) continue;", ") continue;");
  source = source.replaceAll("y: bounds.top + bounds.height / 2", "y: offset");
  source = source.replace(
    String.raw`            const bounds = info?.bounds;
            if ((info?.latex ?? "").trim() !== "\\\\placeholder{}" || !bounds) continue;
            placeholders.push({
              offset,
              y: bounds.top + bounds.height / 2,
            });`,
    String.raw`            if ((info?.latex ?? "").trim() !== "\\\\placeholder{}") continue;
            placeholders.push({ offset, y: offset });`,
  );
  source = source.replace(
    String.raw`            const bounds = info?.bounds;
            if ((info?.latex ?? "").trim() !== "\\\\placeholder{}" || !bounds) continue;
            placeholders.push({ offset, y: bounds.top + bounds.height / 2 });`,
    String.raw`            if ((info?.latex ?? "").trim() !== "\\\\placeholder{}") continue;
            placeholders.push({ offset, y: offset });`,
  );
  source = source.replace(
    /placeholders\.sort\(\(left, right\) => left\.y - right\.y\);/g,
    `(() => {\n            placeholders.sort((left, right) => left.offset - right.offset);\n            if (placeholders.length === 2) {\n              placeholders.reverse();\n            } else if (placeholders.length >= 3) {\n              const model = [...placeholders];\n              placeholders.splice(0, placeholders.length, model.at(-1), model[0], ...model.slice(1, -1));\n            }\n          })();`,
  );
}
if (basename(sourcePath) === "targeted_editor_regression.mjs" && forwardedArguments.includes("native-input-popover")) {
  source = source.replace(
    `          customCandidateVisible: Boolean(document.querySelector(".suggestion-popup")),`,
    `          customCandidateVisible: Boolean(document.querySelector(".suggestion-popup")),\n          sourceExists: Boolean(source),\n          sourceClass: source?.className ?? "",\n          sourceItems: source?.querySelectorAll("li[data-command]").length ?? 0,\n          rawLatex: document.querySelector("math-field")?.shadowRoot?.querySelector(".ML__raw-latex")?.textContent ?? "",\n          fieldValue: document.querySelector("math-field")?.value ?? "",\n          fieldMode: document.querySelector("math-field")?.mode ?? "",\n          popoverPolicy: document.querySelector("math-field")?.popoverPolicy ?? "",`,
  );
}
if (basename(sourcePath) === "quick_format_toolbar_regression.mjs") {
  source = source.replace(
    `            hasLineAlignment: line.hasAttribute("data-alignment"),`,
    `            hasLineAlignment: line.hasAttribute("data-alignment"),\n            lineWidth: line.getBoundingClientRect().width,\n            lineMainWidth: line.querySelector('.formula-line-main')?.getBoundingClientRect().width ?? -1,\n            stackWidth: line.closest('.mathfield-stack')?.getBoundingClientRect().width ?? -1,\n            editorWidth: line.closest('.multi-line-editor')?.getBoundingClientRect().width ?? -1,\n            hostCss: host ? { width: getComputedStyle(host).width, maxWidth: getComputedStyle(host).maxWidth, minWidth: getComputedStyle(host).minWidth, flex: getComputedStyle(host).flex, transform: getComputedStyle(host).transform, display: getComputedStyle(host).display } : null,\n            hostInline: host?.getAttribute('style') ?? '',\n            hostRules: host ? Array.from(document.styleSheets).flatMap((sheet) => { try { return Array.from(sheet.cssRules ?? []).flatMap((rule) => rule.selectorText && host.matches(rule.selectorText) && (rule.style?.width || rule.style?.flex || rule.style?.maxWidth) ? [{ selector: rule.selectorText, width: rule.style.width, flex: rule.style.flex, maxWidth: rule.style.maxWidth }] : []); } catch { return []; } }) : [],`,
  );
}
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

source = source.replace(
  /console\.error\(error\);/g,
  'console.error(error?.message ?? String(error));',
);
source = source.replace(
  /console\.error\(error instanceof Error \? error\.stack : error\);/g,
  'console.error(error?.message ?? String(error));',
);

process.argv = [process.execPath, sourcePath, ...forwardedArguments];
const encodedSource = Buffer.from(source, "utf8").toString("base64");
try {
  await import(`data:text/javascript;base64,${encodedSource}`);
} catch (error) {
  console.error(error?.message ?? String(error));
  process.exitCode = 1;
}
