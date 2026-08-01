import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createServer } from "node:http";
import { readFile, rm } from "node:fs/promises";
import { extname, join, normalize } from "node:path";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const offset = process.pid % 600;
const port = 20200 + offset;
const debugPort = 21200 + offset;
const baseUrl = `http://127.0.0.1:${port}`;
const sessionId = "office-font-size-regression-session";
const officeUrl = `${baseUrl}/dialog/${sessionId}?runtime=vsto-desktop`;
const converterUrl = `${baseUrl}/dialog/${sessionId}?runtime=vsto-convert`;
const distRoot = join(process.cwd(), "dist-office-windows-native");
const chromeProfile = createBrowserProfilePath("visualtex-office-font-size");
const chromePath = resolveChromiumExecutable();

const mimeTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".woff2": "font/woff2",
};

function createSession(host, fontSizePt) {
  const lineId = `${host}-font-line`;
  return {
    id: sessionId,
    mode: "create",
    host,
    formulaId: `${host}-font-formula`,
    sourceDocumentId: `${host}-document`,
    sourceObjectId: null,
    title: host === "word" ? "Word Formula" : "PowerPoint Formula",
    lines: [{ id: lineId, latex: "x+y" }],
    activeLineId: lineId,
    codeFormat: "raw",
    displayMode: host === "word" ? "inline" : "block",
    numbered: false,
    fontSizePt,
    exportWidth: 0,
    exportHeight: 0,
    exportResult: null,
    originalMetadata: null,
    dirty: false,
    status: "created",
    autoCommitOnClose: true,
    explicitCancel: false,
    error: null,
    createdAt: Date.now(),
    updatedAt: Date.now(),
    expiresAt: Date.now() + 60_000,
  };
}

let session = createSession("powerpoint", 20);
const updates = [];
let completeCommits = false;
let closeRequests = 0;

function writeJson(response, status, value) {
  response.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
  });
  response.end(JSON.stringify(value));
}

async function readJsonBody(request) {
  const chunks = [];
  for await (const chunk of request) chunks.push(chunk);
  const source = Buffer.concat(chunks).toString("utf8");
  return source ? JSON.parse(source) : {};
}

const server = createServer(async (request, response) => {
  const url = new URL(request.url ?? "/", baseUrl);
  try {
    if (url.pathname === "/api/v1/theme") {
      writeJson(response, 200, { theme: "light" });
      return;
    }
    if (url.pathname === "/api/v1/preferences") {
      writeJson(response, 200, { powerpointDefaultFontSizePt: 28 });
      return;
    }
    if (url.pathname === `/api/v1/sessions/${sessionId}`) {
      if (request.method === "PATCH") {
        const update = await readJsonBody(request);
        updates.push(update);
        session = {
          ...session,
          ...update,
          updatedAt: Date.now(),
        };
        if (completeCommits && update.status === "committing") {
          session = {
            ...session,
            status: "completed",
            updatedAt: Date.now(),
          };
        }
      }
      writeJson(response, 200, session);
      return;
    }
    if (url.pathname === `/api/v1/app/sessions/${sessionId}/close`) {
      closeRequests += 1;
      writeJson(response, 200, { closed: true });
      return;
    }
    if (url.pathname.startsWith("/api/v1/ocr/")) {
      writeJson(response, 503, { error: "OCR is not needed in this regression" });
      return;
    }
    if (url.pathname.startsWith("/dialog/")) {
      const source = await readFile(join(distRoot, "dialog", "index.html"), "utf8");
      const meta = [
        '<meta name="visualtex-install-token" content="font-regression" />',
        '<meta name="visualtex-native-powerpoint-commit" content="false" />',
        '<meta name="visualtex-theme" content="light" />',
      ].join("\n");
      response.writeHead(200, {
        "Content-Type": "text/html; charset=utf-8",
        "Cache-Control": "no-store",
      });
      response.end(source.replace("</head>", `${meta}\n</head>`));
      return;
    }
    if (url.pathname.startsWith("/assets/")) {
      const relative = normalize(url.pathname.slice(1));
      if (relative.startsWith("..")) {
        response.writeHead(403).end();
        return;
      }
      const content = await readFile(join(distRoot, relative));
      response.writeHead(200, {
        "Content-Type": mimeTypes[extname(relative)] ?? "application/octet-stream",
        "Cache-Control": "no-store",
      });
      response.end(content);
      return;
    }
    response.writeHead(404).end();
  } catch (error) {
    response.writeHead(500, { "Content-Type": "text/plain; charset=utf-8" });
    response.end(String(error));
  }
});

async function waitFor(url, timeoutMs = 15_000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while the local server or browser starts.
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
      this.pending.set(id, { resolve, reject });
      this.socket.send(JSON.stringify({ id, method, params }));
    });
  }

  async evaluate(expression) {
    const result = await this.send("Runtime.evaluate", {
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
  }

  close() {
    this.socket?.close();
  }
}

async function waitForPage() {
  const startedAt = Date.now();
  while (Date.now() - startedAt < 15_000) {
    const targets = await (
      await fetch(`http://127.0.0.1:${debugPort}/json/list`)
    ).json();
    const target = targets.find(
      (candidate) => candidate.type === "page" && candidate.url.startsWith(officeUrl),
    );
    if (target) return target;
    await sleep(80);
  }
  throw new Error("Timed out waiting for Office font-size regression page");
}

async function waitForEvaluation(client, expression, description, timeoutMs = 10_000) {
  const startedAt = Date.now();
  let lastValue;
  while (Date.now() - startedAt < timeoutMs) {
    lastValue = await client.evaluate(expression);
    if (lastValue?.ready) return lastValue;
    await sleep(60);
  }
  throw new Error(`Timed out waiting for ${description}: ${JSON.stringify(lastValue)}`);
}

async function setFontSize(client, fontSizePt) {
  await client.evaluate(`(() => {
    const select = document.querySelector('[data-office-font-size]');
    if (!(select instanceof HTMLSelectElement)) return false;
    const setter = Object.getOwnPropertyDescriptor(
      HTMLSelectElement.prototype,
      'value',
    ).set;
    setter.call(select, ${JSON.stringify(String(fontSizePt))});
    select.dispatchEvent(new Event('change', { bubbles: true }));
    return true;
  })()`);
}

async function dispatchOfficeShortcut(client, overrides = {}) {
  return client.evaluate(`(() => {
    const target = document.querySelector('math-field') ?? document.body;
    const event = new KeyboardEvent('keydown', {
      key: 's',
      code: 'KeyS',
      ctrlKey: true,
      bubbles: true,
      cancelable: true,
      ...${JSON.stringify(overrides)},
    });
    const dispatchResult = target.dispatchEvent(event);
    return {
      defaultPrevented: event.defaultPrevented,
      dispatchResult,
      lateCaptureCount: window.__visualtexLateSaveShortcutCount ?? 0,
    };
  })()`);
}

async function main() {
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(port, "127.0.0.1", resolve);
  });
  let chrome;
  let client;
  try {
    await waitFor(`${baseUrl}/api/v1/preferences`);
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
        officeUrl,
      ],
      { stdio: "ignore" },
    );
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const target = await waitForPage();
    client = new CdpClient(target.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");

    const powerpointDefault = await waitForEvaluation(
      client,
      `(() => {
        const select = document.querySelector('[data-office-font-size]');
        const labels = [...(select?.querySelectorAll('option') ?? [])].map(
          (option) => option.textContent?.trim() ?? '',
        );
        const field = document.querySelector('math-field');
        return {
          ready:
            select instanceof HTMLSelectElement &&
            select.value === '28' &&
            select.querySelectorAll('optgroup').length === 2 &&
            labels.includes('小四（12 磅）') &&
            labels.includes('五号（10.5 磅）') &&
            labels.includes('初号（42 磅）') &&
            field?.position === field?.lastOffset,
          value: select?.value ?? '',
          selectedLabel: select?.selectedOptions[0]?.textContent?.trim() ?? '',
          optionGroupLabels: [...(select?.querySelectorAll('optgroup') ?? [])].map(
            (group) => group.label,
          ),
          host: document.querySelector('.office-dialog-header span')?.textContent ?? '',
          focused: field?.hasFocus?.() ?? false,
          position: field?.position ?? -1,
          lastOffset: field?.lastOffset ?? -1,
          documentFocused: document.hasFocus(),
          activeTag: document.activeElement?.tagName ?? '',
          activeClass: document.activeElement?.className ?? '',
          shadowActivePart:
            field?.shadowRoot?.activeElement?.getAttribute?.('part') ?? '',
        };
      })()`,
      "PowerPoint configured default font size with Chinese size options",
    );
    assert.equal(powerpointDefault.value, "28");
    assert.deepEqual(powerpointDefault.optionGroupLabels, ["中文字号", "磅值"]);

    await setFontSize(client, 12);
    const powerpointChineseSize = await waitForEvaluation(
      client,
      `(() => {
        const select = document.querySelector('[data-office-font-size]');
        return {
          ready:
            select?.value === '12' &&
            select?.selectedOptions[0]?.textContent?.includes('小四'),
          value: select?.value ?? '',
          selectedLabel: select?.selectedOptions[0]?.textContent?.trim() ?? '',
        };
      })()`,
      "PowerPoint Chinese font-size selection",
    );
    for (let attempt = 0; attempt < 80 && session.fontSizePt !== 12; attempt += 1) {
      await sleep(60);
    }
    assert.equal(session.fontSizePt, 12, "PowerPoint should persist 小四 as 12 pt");
    assert.ok(
      updates.some((update) => update.fontSizePt === 12),
      "PowerPoint autosave should include the selected Chinese size",
    );

    await setFontSize(client, 31.5);
    const powerpointSaved = await waitForEvaluation(
      client,
      `(() => ({
        ready: ${JSON.stringify(true)} && window.fetch !== undefined,
        value: document.querySelector('[data-office-font-size]')?.value ?? '',
      }))()`,
      "PowerPoint font input update",
    );
    assert.equal(powerpointSaved.value, "31.5");
    for (let attempt = 0; attempt < 80 && session.fontSizePt !== 31.5; attempt += 1) {
      await sleep(60);
    }
    assert.equal(session.fontSizePt, 31.5, "PowerPoint Session should persist selected size");
    assert.ok(
      updates.some((update) => update.fontSizePt === 31.5),
      "PowerPoint autosave should include selected font size",
    );

    await client.send("Page.reload", { ignoreCache: true });
    const powerpointReloaded = await waitForEvaluation(
      client,
      `(() => {
        const input = document.querySelector('[data-office-font-size]');
        return { ready: input?.value === '31.5', value: input?.value ?? '' };
      })()`,
      "PowerPoint Session font size after reload",
    );
    assert.equal(powerpointReloaded.value, "31.5");

    // Leave the PowerPoint page before replacing the mock Session. Its
    // pagehide autosave is valid for the old Session and must finish before
    // the test starts serving Word data.
    await client.send("Page.navigate", { url: "about:blank" });
    await sleep(250);
    session = createSession("word", 11.5);
    updates.length = 0;
    await client.send("Page.navigate", { url: officeUrl });
    const wordInherited = await waitForEvaluation(
      client,
      `(() => {
        const input = document.querySelector('[data-office-font-size]');
        const field = document.querySelector('math-field');
        return {
          ready:
            input?.value === '11.5' &&
            field?.position === field?.lastOffset,
          value: input?.value ?? '',
          host: document.querySelector('.office-dialog-header span')?.textContent ?? '',
          focused: field?.hasFocus?.() ?? false,
          position: field?.position ?? -1,
          lastOffset: field?.lastOffset ?? -1,
        };
      })()`,
      "Word current paragraph font size",
    );
    assert.equal(wordInherited.value, "11.5");
    assert.ok(
      !updates.some((update) => update.dirty === true),
      "Word initial Session load must not persist a transient dirty state",
    );

    await setFontSize(client, 10.5);
    const wordChineseSize = await waitForEvaluation(
      client,
      `(() => {
        const select = document.querySelector('[data-office-font-size]');
        return {
          ready:
            select?.value === '10.5' &&
            select?.selectedOptions[0]?.textContent?.includes('五号'),
          value: select?.value ?? '',
          selectedLabel: select?.selectedOptions[0]?.textContent?.trim() ?? '',
        };
      })()`,
      "Word Chinese font-size selection",
    );
    for (let attempt = 0; attempt < 80 && session.fontSizePt !== 10.5; attempt += 1) {
      await sleep(60);
    }
    assert.equal(session.fontSizePt, 10.5, "Word should persist 五号 as 10.5 pt");
    assert.ok(
      updates.some((update) => update.fontSizePt === 10.5),
      "Word autosave should include the selected Chinese size",
    );

    await setFontSize(client, 13);
    for (let attempt = 0; attempt < 80 && session.fontSizePt !== 13; attempt += 1) {
      await sleep(60);
    }
    assert.equal(session.fontSizePt, 13, "Word Session should persist selected size");
    assert.ok(
      updates.some((update) => update.fontSizePt === 13),
      "Word autosave should include selected font size",
    );

    const shortcutMetadata = await client.evaluate(`(() => {
      const primary = document.querySelector('.office-workspace-actions .primary-button');
      window.__visualtexLateSaveShortcutCount = 0;
      window.addEventListener('keydown', (event) => {
        if (event.ctrlKey && (event.code === 'KeyS' || event.key.toLowerCase() === 's')) {
          window.__visualtexLateSaveShortcutCount += 1;
        }
      }, true);
      return {
        ariaKeyShortcuts: primary?.getAttribute('aria-keyshortcuts') ?? '',
        title: primary?.getAttribute('title') ?? '',
        helperText: document.querySelector('.office-workspace-actions span')?.textContent ?? '',
      };
    })()`);
    assert.equal(shortcutMetadata.ariaKeyShortcuts, "Control+S");
    assert.ok(shortcutMetadata.title.includes("Ctrl+S"));
    assert.ok(shortcutMetadata.helperText.includes("Ctrl+S"));

    const shiftSave = await dispatchOfficeShortcut(client, { shiftKey: true });
    assert.equal(shiftSave.defaultPrevented, false);
    assert.equal(shiftSave.lateCaptureCount, 1);

    const altSave = await dispatchOfficeShortcut(client, { altKey: true });
    assert.equal(altSave.defaultPrevented, false);
    assert.equal(altSave.lateCaptureCount, 2);

    const metaSave = await dispatchOfficeShortcut(client, {
      ctrlKey: false,
      metaKey: true,
    });
    assert.equal(metaSave.defaultPrevented, false);
    assert.equal(metaSave.lateCaptureCount, 2);

    const composingSave = await dispatchOfficeShortcut(client, { isComposing: true });
    assert.equal(composingSave.defaultPrevented, false);
    assert.equal(composingSave.lateCaptureCount, 3);

    const repeatedSave = await dispatchOfficeShortcut(client, { repeat: true });
    assert.equal(repeatedSave.defaultPrevented, true);
    assert.equal(repeatedSave.dispatchResult, false);
    assert.equal(repeatedSave.lateCaptureCount, 3);
    await sleep(150);
    assert.ok(
      !updates.some((update) => update.status === "committing"),
      "Repeated Ctrl+S must be swallowed without starting a commit",
    );

    completeCommits = true;
    const exactSave = await dispatchOfficeShortcut(client);
    assert.equal(exactSave.defaultPrevented, true);
    assert.equal(exactSave.dispatchResult, false);
    assert.equal(exactSave.lateCaptureCount, 3);
    for (
      let attempt = 0;
      attempt < 120 &&
      !(session.status === "completed" && closeRequests === 1);
      attempt += 1
    ) {
      await sleep(60);
    }
    assert.equal(session.status, "completed", "Ctrl+S should commit the Office Session");
    assert.equal(closeRequests, 1, "Ctrl+S should close the Office editor after applying");
    assert.equal(
      updates.filter((update) => update.status === "committing").length,
      1,
      "One Ctrl+S press must enqueue exactly one commit",
    );

    // Reproduce the hidden-converter race directly. The previous PowerPoint
    // page has already placed 20/28/31.5 pt values in the component lifecycle;
    // a Word converter must still render from its immutable 11 pt Session,
    // never from the previous React state.
    try {
      await client.send("Page.navigate", { url: "about:blank" });
    } catch (error) {
      if (!String(error).includes("navigated or closed")) throw error;
    }
    client.close();
    client = undefined;
    await sleep(250);
    session = createSession("word", 11);
    session.autoCommitOnClose = false;
    completeCommits = false;
    updates.length = 0;
    const converterTarget = await (
      await fetch(
        `http://127.0.0.1:${debugPort}/json/new?${encodeURIComponent(converterUrl)}`,
        { method: "PUT" },
      )
    ).json();
    client = new CdpClient(converterTarget.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    for (
      let attempt = 0;
      attempt < 200 &&
      !(
        session.status === "committing" &&
        session.exportResult?.svg &&
        session.exportResult?.pngBase64
      );
      attempt += 1
    ) {
      await sleep(60);
    }
    assert.equal(session.status, "committing", "Word converter should commit");
    assert.equal(session.fontSizePt, 11, "Word converter must preserve the paragraph size");
    assert.ok(session.exportResult?.svg, "Word converter should produce SVG");
    assert.ok(session.exportResult?.pngBase64, "Word converter should produce PNG");
    assert.ok(
      Math.abs(session.exportResult.width - 35.504533333333335) < 0.2,
      `Word 11 pt converter width is wrong: ${session.exportResult.width}`,
    );
    assert.ok(
      Math.abs(session.exportResult.height - 13.557333333333332) < 0.2,
      `Word 11 pt converter height is wrong: ${session.exportResult.height}`,
    );
    assert.ok(
      session.exportResult.width < 45,
      "Word converter reused the PowerPoint 20 pt render state",
    );
    const wordConverterUpdate = updates.find(
      (update) => update.status === "committing" && update.exportResult,
    );
    assert.equal(wordConverterUpdate?.fontSizePt, 11);

    console.log(
      JSON.stringify(
        {
          powerpointDefault,
          powerpointChineseSize,
          powerpointSavedSize: 31.5,
          powerpointReloaded,
          wordInherited,
          wordChineseSize,
          wordSavedSize: 13,
          wordConverter: {
            fontSizePt: session.fontSizePt,
            width: session.exportResult.width,
            height: session.exportResult.height,
          },
        },
        null,
        2,
      ),
    );
    console.log("Office formula font-size regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    await new Promise((resolve) => server.close(resolve));
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => undefined);
  }
}

await main();
