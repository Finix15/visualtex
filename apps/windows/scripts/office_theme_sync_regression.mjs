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
const offset = process.pid % 700;
const port = 18100 + offset;
const debugPort = 19100 + offset;
const baseUrl = `http://127.0.0.1:${port}`;
const officeUrl = `${baseUrl}/dialog/theme-regression?runtime=vsto-desktop`;
const distRoot = join(process.cwd(), "dist-office-windows-native");
const chromeProfile = createBrowserProfilePath("visualtex-windows-office-theme");
const chromePath = resolveChromiumExecutable();
let currentTheme = "purple";
let currentEditorLayout = "classic";

const mimeTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".woff2": "font/woff2",
};

function writeJson(response, status, value) {
  response.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
  });
  response.end(JSON.stringify(value));
}

const server = createServer(async (request, response) => {
  const url = new URL(request.url ?? "/", baseUrl);
  try {
    if (url.pathname === "/api/v1/theme") {
      writeJson(response, 200, {
        theme: currentTheme,
        editorLayout: currentEditorLayout,
      });
      return;
    }
    if (url.pathname.startsWith("/api/v1/sessions/")) {
      writeJson(response, 404, { error: "Theme regression does not need a session" });
      return;
    }
    if (url.pathname.startsWith("/dialog/")) {
      const source = await readFile(join(distRoot, "dialog", "index.html"), "utf8");
      const meta = [
        '<meta name="visualtex-install-token" content="theme-regression" />',
        `<meta name="visualtex-theme" content="${currentTheme}" />`,
        `<meta name="visualtex-editor-layout" content="${currentEditorLayout}" />`,
      ].join("\n");
      const html = source.replace("</head>", `${meta}\n</head>`);
      response.writeHead(200, {
        "Content-Type": "text/html; charset=utf-8",
        "Cache-Control": "no-store",
      });
      response.end(html);
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
      // Retry while the server or browser starts.
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
  throw new Error("Timed out waiting for the Office theme regression page");
}

async function readAppearance(client) {
  return client.evaluate(`(() => {
    let editorLayout = null;
    try {
      editorLayout = JSON.parse(localStorage.getItem('visualtex-editor') || '{}')?.state?.editorLayout ?? null;
    } catch {}
    return {
      theme: document.documentElement.dataset.theme,
      editorLayout,
      background: getComputedStyle(document.documentElement).getPropertyValue('--bg').trim().toLowerCase(),
      surface: getComputedStyle(document.documentElement).getPropertyValue('--surface').trim().toLowerCase(),
      formulaSurface: getComputedStyle(document.documentElement).getPropertyValue('--formula-surface').trim().toLowerCase(),
      caret: getComputedStyle(document.documentElement).getPropertyValue('--formula-caret').trim().toLowerCase(),
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
    await waitFor(`${baseUrl}/api/v1/theme`);
    chrome = spawn(
      chromePath,
      [
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        `--remote-debugging-port=${debugPort}`,
        `--user-data-dir=${chromeProfile}`,
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
    await sleep(500);

    assert.deepEqual(await readAppearance(client), {
      theme: "purple",
      editorLayout: "classic",
      background: "#120e16",
      surface: "#362842",
      formulaSurface: "#433252",
      caret: "#d7c2ff",
    });

    currentTheme = "green";
    currentEditorLayout = "standard";
    let synchronized;
    for (let attempt = 0; attempt < 20; attempt += 1) {
      await sleep(100);
      synchronized = await readAppearance(client);
      if (
        synchronized.theme === "green" &&
        synchronized.editorLayout === "standard"
      ) break;
    }
    assert.deepEqual(synchronized, {
      theme: "green",
      editorLayout: "standard",
      background: "#0d120f",
      surface: "#25352d",
      formulaSurface: "#2a3b32",
      caret: "#8bd4ac",
    });

    process.stdout.write(
      "Windows Office theme and editor-layout inheritance regression passed\n",
    );
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    await new Promise((resolve) => server.close(resolve));
    await sleep(250);
    await rm(chromeProfile, {
      recursive: true,
      force: true,
      maxRetries: 4,
      retryDelay: 150,
    }).catch((error) => {
      if (process.platform !== "win32" || error?.code !== "EBUSY") throw error;
    });
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
