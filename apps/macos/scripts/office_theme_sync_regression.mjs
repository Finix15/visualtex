import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 700;
const previewPort = 8800 + offset;
const debugPort = 15800 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const officeUrl = `${baseUrl}/office-native-dialog.html?session=00000000-0000-4000-8000-000000000000&theme=purple`;
const chromeProfile = `/tmp/visualtex-office-theme-${process.pid}`;
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

async function pageTarget(predicate) {
  const targets = await (
    await fetch(`http://127.0.0.1:${debugPort}/json/list`)
  ).json();
  return targets.find((target) => target.type === "page" && predicate(target.url));
}

async function connectPage(target) {
  const client = new CdpClient(target.webSocketDebuggerUrl);
  await client.connect();
  await client.send("Runtime.enable");
  await client.send("Page.enable");
  return client;
}

async function main() {
  const preview = spawn(
    process.execPath,
    [
      "node_modules/vite/bin/vite.js",
      "preview",
      "--host",
      "127.0.0.1",
      "--port",
      String(previewPort),
      "--strictPort",
    ],
    { cwd: process.cwd(), stdio: "ignore" },
  );
  let chrome;
  let mainClient;
  let officeClient;

  try {
    await waitFor(baseUrl);
    chrome = spawn(
      chromePath,
      [
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        `--remote-debugging-port=${debugPort}`,
        `--user-data-dir=${chromeProfile}`,
        baseUrl,
      ],
      { stdio: "ignore" },
    );
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);

    const mainTarget = await pageTarget((url) => url === `${baseUrl}/` || url === baseUrl);
    assert.ok(mainTarget, "Main VisualTeX target must exist");
    mainClient = await connectPage(mainTarget);
    await mainClient.send("Target.createTarget", { url: officeUrl });
    let officeTarget;
    for (let attempt = 0; attempt < 80 && !officeTarget; attempt += 1) {
      officeTarget = await pageTarget((url) => url.startsWith(officeUrl));
      if (!officeTarget) await sleep(50);
    }
    assert.ok(officeTarget, "Office formula editor target must exist");
    officeClient = await connectPage(officeTarget);
    await sleep(300);

    const inherited = await officeClient.evaluate(`({
      theme: document.documentElement.dataset.theme,
      background: getComputedStyle(document.documentElement).getPropertyValue('--bg').trim().toLowerCase(),
      paper: getComputedStyle(document.documentElement).getPropertyValue('--surface').trim().toLowerCase(),
    })`);
    assert.deepEqual(inherited, {
      theme: "purple",
      background: "#120e16",
      paper: "#362842",
    });

    await mainClient.evaluate(`(() => {
      localStorage.setItem('visualtex.active-theme', 'green');
      const channel = new BroadcastChannel('visualtex-theme');
      channel.postMessage('green');
      channel.close();
    })()`);
    await sleep(180);
    const synchronized = await officeClient.evaluate(`({
      theme: document.documentElement.dataset.theme,
      background: getComputedStyle(document.documentElement).getPropertyValue('--bg').trim().toLowerCase(),
    })`);
    assert.deepEqual(synchronized, {
      theme: "green",
      background: "#0d120f",
    });

    process.stdout.write("Office theme inheritance and live synchronization regression passed\n");
  } finally {
    mainClient?.close();
    officeClient?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(180);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => undefined);
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
