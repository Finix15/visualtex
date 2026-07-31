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
const port = 18300 + offset;
const debugPort = 19300 + offset;
const baseUrl = `http://127.0.0.1:${port}`;
const sessionId = "11111111-2222-4333-8444-555555555555";
const officeUrl = `${baseUrl}/dialog/${sessionId}?runtime=vsto-bulk-import`;
const distRoot = join(process.cwd(), "dist-office-windows-native");
const chromeProfile = createBrowserProfilePath("visualtex-document-import-preview-scale");
const chromePath = resolveChromiumExecutable();

const source = String.raw`# 长公式预览

这是用于检查批量导入预览字号的正文，文字应保持清晰但明显小于原来的尺寸。

$$
\frac{\displaystyle\sum_{n=1}^{N}\left(a_n+b_n+c_n+d_n\right)^2}{\displaystyle\prod_{k=1}^{M}\left(1+x_k^2+y_k^2\right)}+\int_{-\infty}^{\infty}e^{-x^2}\,\mathrm{d}x+\oiint_{\Sigma}\bm F\cdot\mathrm{d}\bm S
$$`;

const session = {
  id: sessionId,
  mode: "create",
  host: "word",
  formulaId: "preview-formula",
  sourceDocumentId: null,
  sourceObjectId: null,
  title: "Word 文档批量导入",
  lines: [{ id: "preview-line", latex: source }],
  activeLineId: "preview-line",
  codeFormat: "markdown",
  displayMode: "block",
  objectMode: "wordOmml",
  numbered: false,
  fontSizePt: 11,
  exportWidth: 0,
  exportHeight: 0,
  exportResult: null,
  originalMetadata: null,
  dirty: false,
  status: "editing",
  autoCommitOnClose: false,
  explicitCancel: false,
  error: null,
  createdAt: Date.now(),
  updatedAt: Date.now(),
  expiresAt: Date.now() + 60_000,
};

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
      writeJson(response, 200, { theme: "light", editorLayout: "classic" });
      return;
    }
    if (url.pathname === `/api/v1/sessions/${sessionId}`) {
      writeJson(response, 200, session);
      return;
    }
    if (url.pathname.startsWith("/dialog/")) {
      const htmlSource = await readFile(join(distRoot, "dialog", "index.html"), "utf8");
      const meta = [
        '<meta name="visualtex-install-token" content="preview-regression" />',
        '<meta name="visualtex-theme" content="light" />',
        '<meta name="visualtex-editor-layout" content="classic" />',
      ].join("\n");
      response.writeHead(200, {
        "Content-Type": "text/html; charset=utf-8",
        "Cache-Control": "no-store",
      });
      response.end(htmlSource.replace("</head>", `${meta}\n</head>`));
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
    } catch {}
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
      throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
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
    const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
    const target = targets.find(
      (candidate) => candidate.type === "page" && candidate.url.startsWith(officeUrl),
    );
    if (target) return target;
    await sleep(80);
  }
  throw new Error("Timed out waiting for the document import preview page");
}

async function main() {
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(port, "127.0.0.1", resolve);
  });
  let chrome;
  let client;
  try {
    chrome = spawn(
      chromePath,
      [
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        `--remote-debugging-port=${debugPort}`,
        `--user-data-dir=${chromeProfile}`,
        "--window-size=1440,1000",
        officeUrl,
      ],
      { stdio: "ignore" },
    );
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const target = await waitForPage();
    client = new CdpClient(target.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");

    let metrics;
    for (let attempt = 0; attempt < 80; attempt += 1) {
      metrics = await client.evaluate(`(() => {
        const documentNode = document.querySelector('.doc-import-preview-document');
        const row = document.querySelector('.doc-import-display-row');
        const svg = row?.querySelector('svg');
        const heading = document.querySelector('h1.doc-import-heading');
        if (!documentNode || !row || !svg || !heading) return null;
        const svgRect = svg.getBoundingClientRect();
        return {
          bodyFontSize: getComputedStyle(documentNode).fontSize,
          headingFontSize: getComputedStyle(heading).fontSize,
          svgWidth: svgRect.width,
          svgHeight: svgRect.height,
          rowClientWidth: row.clientWidth,
          rowScrollWidth: row.scrollWidth,
        };
      })()`);
      if (metrics) break;
      await sleep(100);
    }

    assert.ok(metrics, "Document import preview did not render");
    assert.equal(metrics.bodyFontSize, "9px");
    assert.equal(metrics.headingFontSize, "14px");
    assert.ok(metrics.svgHeight >= 6, `Formula became unreadably short: ${metrics.svgHeight}px`);
    assert.ok(
      metrics.svgWidth <= metrics.rowClientWidth + 1,
      `Long formula did not fit the preview row: ${metrics.svgWidth}px > ${metrics.rowClientWidth}px`,
    );
    assert.ok(
      metrics.rowScrollWidth <= metrics.rowClientWidth + 2,
      `Preview row still overflows horizontally: ${metrics.rowScrollWidth}px > ${metrics.rowClientWidth}px`,
    );
    process.stdout.write(
      `Document import preview scaling regression passed: ${JSON.stringify(metrics)}\n`,
    );
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    await new Promise((resolve) => server.close(resolve));
    await sleep(250);
    await rm(chromeProfile, { recursive: true, force: true, maxRetries: 4, retryDelay: 150 }).catch(
      (error) => {
        if (process.platform !== "win32" || error?.code !== "EBUSY") throw error;
      },
    );
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
