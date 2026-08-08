import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 700;
const vitePort = 19400 + offset;
const debugPort = 25400 + offset;
const baseUrl = `http://127.0.0.1:${vitePort}`;
const chromeProfile = `/tmp/visualtex-custom-symbol-prototype-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while Vite or Chrome starts.
    }
    await sleep(80);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

class CdpClient {
  constructor(url) {
    this.url = url;
    this.nextId = 1;
    this.pending = new Map();
  }

  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise((resolve, reject) => {
      this.socket.addEventListener("open", resolve, { once: true });
      this.socket.addEventListener("error", reject, { once: true });
    });
    this.socket.addEventListener("message", (event) => {
      const message = JSON.parse(event.data);
      if (!message.id) return;
      const pending = this.pending.get(message.id);
      if (!pending) return;
      this.pending.delete(message.id);
      if (message.error) pending.reject(new Error(message.error.message));
      else pending.resolve(message.result);
    });
  }

  send(method, params = {}) {
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`CDP ${method} timed out`));
      }, 15000);
      this.pending.set(id, {
        resolve: (value) => {
          clearTimeout(timeout);
          resolve(value);
        },
        reject: (error) => {
          clearTimeout(timeout);
          reject(error);
        },
      });
      this.socket.send(JSON.stringify({ id, method, params }));
    });
  }

  close() {
    this.socket?.close();
  }
}

async function main() {
  const vite = spawn(
    process.execPath,
    [
      "node_modules/vite/bin/vite.js",
      "preview",
      "--host",
      "127.0.0.1",
      "--port",
      String(vitePort),
      "--strictPort",
    ],
    { cwd: process.cwd(), stdio: "ignore" },
  );
  let chrome;
  let client;

  try {
    process.stdout.write("[custom-symbol] waiting for Vite\n");
    await waitFor(baseUrl);
    process.stdout.write("[custom-symbol] Vite ready\n");
    chrome = spawn(
      chromePath,
      [
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        `--remote-debugging-port=${debugPort}`,
        `--user-data-dir=${chromeProfile}`,
        "--window-size=1400,1000",
        baseUrl,
      ],
      { stdio: "ignore" },
    );
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    process.stdout.write("[custom-symbol] Chrome ready\n");
    const targets = await (
      await fetch(`http://127.0.0.1:${debugPort}/json/list`)
    ).json();
    const page = targets.find(
      (target) => target.type === "page" && target.url.startsWith(baseUrl),
    );
    assert.ok(page, "VisualTeX Chrome target must exist");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(650);
    process.stdout.write("[custom-symbol] page navigated\n");

    const evaluate = async (expression) => {
      const result = await client.send("Runtime.evaluate", {
        expression,
        awaitPromise: true,
        returnByValue: true,
      });
      if (result.exceptionDetails) {
        throw new Error(
          result.exceptionDetails.exception?.description ||
            result.exceptionDetails.text ||
            "Runtime.evaluate failed",
        );
      }
      return result.result.value;
    };

    await evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
      localStorage.setItem("visualtex.onboarding.macos.desktop.v1.2.0.completed", "true");
      localStorage.setItem("visualtex.office.macos.native-first-run.v1.2.0.completed", "true");
      const key = "visualtex-editor";
      const persisted = JSON.parse(localStorage.getItem(key) || "{}");
      persisted.state = {
        ...(persisted.state || {}),
        checkUpdatesOnStartup: false,
        sourceOpen: false,
        latexCodeFormat: "raw",
      };
      localStorage.setItem(key, JSON.stringify(persisted));
      return true;
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(650);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(650);
    process.stdout.write("[custom-symbol] preferences loaded\n");
    let editorReady = false;
    for (let attempt = 0; attempt < 160; attempt += 1) {
      editorReady = await evaluate(`Boolean(document.querySelector("math-field"))`);
      if (editorReady) break;
      await sleep(50);
    }
    if (!editorReady) {
      const pageState = await evaluate(`({
        title: document.title,
        body: document.body?.innerText?.slice(0, 1200) || "",
        url: location.href,
      })`);
      throw new Error(`Mathfield did not mount: ${JSON.stringify(pageState)}`);
    }

    const dispatchKey = async ({
      key,
      code,
      keyCode,
      text = "",
      modifiers = 0,
    }) => {
      const common = {
        key,
        code,
        modifiers,
        windowsVirtualKeyCode: keyCode,
        nativeVirtualKeyCode: keyCode,
      };
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        ...common,
        ...(text ? { text, unmodifiedText: text } : {}),
      });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(75);
    };

    const typeCharacter = (value, code, keyCode) =>
      dispatchKey({ key: value, code, keyCode, text: value });

    const typeRawCommand = async (command) => {
      await typeCharacter("\\", "Backslash", 220);
      for (const character of command) {
        const upper = character.toUpperCase();
        await typeCharacter(character, `Key${upper}`, upper.charCodeAt(0));
      }
      await typeCharacter(" ", "Space", 32);
    };

    const focusField = () =>
      evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        return { value: field.value, position: field.position, lastOffset: field.lastOffset };
      })()`);

    process.stdout.write("[custom-symbol] editor ready\n");
    await focusField();
    await typeCharacter("a", "KeyA", 65);
    await typeCharacter("b", "KeyB", 66);
    const ordinary = await evaluate(`document.querySelector("math-field").value`);
    assert.match(ordinary, /ab/, "Ordinary MathLive input changed after prototype registration");

    await dispatchKey({ key: "a", code: "KeyA", keyCode: 65, modifiers: 4 });
    await dispatchKey({ key: "Backspace", code: "Backspace", keyCode: 8 });
    await typeCharacter("A", "KeyA", 65);
    await typeRawCommand("vtxtestsymbol");
    await typeCharacter("B", "KeyB", 66);

    process.stdout.write("[custom-symbol] physical command typed\n");
    const custom = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      const BS = String.fromCharCode(92);
      const command = BS + "vtxtestsymbol";
      let macroEnd = -1;
      for (let offset = 1; offset <= field.lastOffset; offset += 1) {
        if (field.getElementInfo(offset)?.latex?.trim() === command) {
          macroEnd = offset;
          break;
        }
      }
      return {
        value: field.value,
        macroEnd,
        classCount: field.shadowRoot?.querySelectorAll(
          ".visualtex-custom-symbol-vtxtestsymbol",
        ).length || 0,
      };
    })()`);
    assert.match(custom.value, /\\vtxtestsymbol/, "Typed custom command was not preserved as LaTeX");
    assert.equal(custom.classCount, 1, "Custom command did not render its vector host class");
    assert.ok(custom.macroEnd > 0, "Could not find custom macro atom boundary");

    const navigation = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.position = ${custom.macroEnd};
      const macroEnd = field.position;
      field.executeCommand("moveToPreviousChar");
      const macroStart = field.position;
      field.executeCommand("moveToNextChar");
      return { macroStart, macroEnd, returnedEnd: field.position };
    })()`);
    assert.ok(
      navigation.macroEnd - navigation.macroStart > 1,
      "Custom symbol caret entered internal macro layers instead of jumping over the atom",
    );
    assert.equal(navigation.returnedEnd, navigation.macroEnd);

    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.position = ${custom.macroEnd};
      field.selection = { ranges: [[${custom.macroEnd}, ${custom.macroEnd}]], direction: "none" };
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      return true;
    })()`);
    await dispatchKey({ key: "Backspace", code: "Backspace", keyCode: 8 });
    process.stdout.write("[custom-symbol] atom navigation verified\n");
    const afterDelete = await evaluate(`document.querySelector("math-field").value`);
    assert.doesNotMatch(afterDelete, /\\vtxtestsymbol/, "Backspace did not delete the custom atom as one unit");

    await dispatchKey({ key: "z", code: "KeyZ", keyCode: 90, modifiers: 4 });
    const afterUndo = await evaluate(`document.querySelector("math-field").value`);
    assert.match(afterUndo, /\\vtxtestsymbol/, "VisualTeX history undo did not restore the custom atom");

    await dispatchKey({ key: "z", code: "KeyZ", keyCode: 90, modifiers: 12 });
    const afterRedo = await evaluate(`document.querySelector("math-field").value`);
    assert.doesNotMatch(afterRedo, /\\vtxtestsymbol/, "VisualTeX history redo did not re-delete the custom atom");

    await dispatchKey({ key: "z", code: "KeyZ", keyCode: 90, modifiers: 4 });
    const sourceToggleState = await evaluate(`(() => {
      const toggle =
        document.querySelector(".source-toggle") ||
        document.querySelector('[data-classic-bottom-view="source"]');
      if (toggle) toggle.click();
      return {
        found: Boolean(toggle),
        label: toggle?.getAttribute("aria-label") || toggle?.textContent || "",
        mode: toggle?.matches('[data-classic-bottom-view="source"]') ? "classic" : "standard",
      };
    })()`);
    assert.equal(sourceToggleState.found, true, `Source toggle missing: ${JSON.stringify(sourceToggleState)}`);
    let sourceText = "";
    let sourceMounted = false;
    for (let attempt = 0; attempt < 80; attempt += 1) {
      const sourceState = await evaluate(`(() => {
        const content = document.querySelector(".source-panel .cm-content");
        return {
          mounted: Boolean(content),
          text: content?.textContent || "",
        };
      })()`);
      sourceMounted = sourceState.mounted;
      sourceText = sourceState.text;
      if (sourceMounted && sourceText.includes("\\\\vtxtestsymbol")) break;
      await sleep(50);
    }
    process.stdout.write("[custom-symbol] history verified\n");
    if (!sourceMounted || !sourceText.includes("\\vtxtestsymbol")) {
      const sourceDebug = await evaluate(`(() => ({
        fieldValue: document.querySelector("math-field")?.value || "",
        sourcePanel: document.querySelector(".source-panel")?.className || "",
        sourceText: document.querySelector(".source-panel .cm-content")?.textContent || "",
        sourceOpenButton: document.querySelector(".source-toggle")?.outerHTML || "",
      }))()`);
      throw new Error(`CodeMirror source sync failed: ${JSON.stringify(sourceDebug)}`);
    }

    process.stdout.write("[custom-symbol] source sync verified\n");

    await evaluate(`(() => {
      const content = document.querySelector(".source-panel .cm-content");
      content.focus();
      const selection = window.getSelection();
      const range = document.createRange();
      range.selectNodeContents(content);
      selection.removeAllRanges();
      selection.addRange(range);
      return true;
    })()`);
    await client.send("Input.insertText", {
      text: "z=" + String.fromCharCode(92) + "vtxtestsymbol",
    });
    await sleep(180);
    await evaluate(`(() => {
      document.querySelector(".source-panel .primary-small-button")?.click();
      return true;
    })()`);
    await sleep(260);
    const sourceRoundTrip = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        value: field?.value || "",
        classCount: field?.shadowRoot?.querySelectorAll(
          ".visualtex-custom-symbol-vtxtestsymbol",
        ).length || 0,
      };
    })()`);
    assert.match(
      sourceRoundTrip.value,
      /z=\\vtxtestsymbol/,
      `Source apply did not preserve the custom command: ${JSON.stringify(sourceRoundTrip)}`,
    );
    assert.equal(
      sourceRoundTrip.classCount,
      1,
      "Source-applied custom command did not render as one custom symbol",
    );
    process.stdout.write("[custom-symbol] source round trip verified\n");

    const rendering = await evaluate(`(async () => {
      const field = document.querySelector("math-field");
      const BS = String.fromCharCode(92);
      const command = BS + "vtxtestsymbol";
      const set = async (value) => {
        field.setValue(value, {
          mode: "math", format: "latex", insertionMode: "replaceAll",
          selectionMode: "after", silenceNotifications: true,
        });
        await new Promise(requestAnimationFrame);
        await new Promise(requestAnimationFrame);
        const node = field.shadowRoot?.querySelector(".visualtex-custom-symbol-vtxtestsymbol");
        const box = node?.getBoundingClientRect();
        return {
          value: field.value,
          width: box?.width || 0,
          height: box?.height || 0,
          color: node ? getComputedStyle(node).color : "",
          overlay: node?.querySelector(".ML__rule")
            ? getComputedStyle(node.querySelector(".ML__rule")).backgroundColor
            : "",
        };
      };
      const root = await set(command);
      const subscript = await set("x_{" + command + "}");
      const scripts = await set(command + "_i^j");
      const colored = await set(BS + "textcolor{red}{" + command + "}");
      return { root, subscript, scripts, colored };
    })()`);
    assert.equal(rendering.root.value, "\\vtxtestsymbol");
    assert.equal(rendering.subscript.value, "x_{\\vtxtestsymbol}");
    assert.equal(rendering.scripts.value, "\\vtxtestsymbol_{i}^{j}");
    assert.ok(rendering.root.width > 0 && rendering.root.height > 0);
    assert.ok(
      rendering.subscript.width < rendering.root.width,
      "Custom symbol did not scale down in subscript mathstyle",
    );
    assert.match(rendering.colored.color, /^rgb\(/);
    assert.equal(
      rendering.colored.overlay,
      rendering.colored.color,
      "Custom geometry must inherit the exact MathLive formula color",
    );

    process.stdout.write(
      "Custom symbol prototype atom, source sync, history, scaling, and color regression passed\n",
    );
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    vite.kill("SIGTERM");
    await sleep(220);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => undefined);
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
