import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 1000;
const previewPort = 18400 + offset;
const debugPort = 23600 + offset;
const sessionId = "12345678-1234-4234-9234-123456789abc";
const baseUrl = `http://127.0.0.1:${previewPort}/?view=office-document-import&sessionId=${sessionId}&transport=tauri`;
const chromeProfile = `/tmp/visualtex-document-import-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while the local process starts.
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
    this.events = [];
  }

  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise((resolve, reject) => {
      this.socket.addEventListener("open", resolve, { once: true });
      this.socket.addEventListener("error", reject, { once: true });
    });
    this.socket.addEventListener("message", (event) => {
      const message = JSON.parse(event.data);
      if (!message.id) {
        if (
          message.method === "Runtime.exceptionThrown" ||
          message.method === "Runtime.consoleAPICalled"
        ) {
          this.events.push(message);
        }
        return;
      }
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

  close() {
    this.socket?.close();
  }
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
  let client;

  try {
    await waitFor(`http://127.0.0.1:${previewPort}`);
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
        "about:blank",
      ],
      { stdio: "ignore" },
    );

    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (
      await fetch(`http://127.0.0.1:${debugPort}/json/list`)
    ).json();
    const page = targets.find((target) => target.type === "page");
    if (!page) throw new Error("No Chrome page target found");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.addScriptToEvaluateOnNewDocument", {
      source: `(() => {
        let callbackId = 1;
        const callbacks = new Map();
        window.__VISUALTEX_DOCUMENT_IMPORT_CALLS__ = [];
        Object.defineProperty(HTMLCanvasElement.prototype, "toBlob", {
          configurable: true,
          value(callback) {
            callback(null);
          },
        });
        window.__TAURI_INTERNALS__ = {
          metadata: {
            currentWindow: { label: "office-native-document-test" },
            currentWebview: { label: "office-native-document-test" },
          },
          transformCallback(callback, once = false) {
            const id = callbackId++;
            callbacks.set(id, { callback, once });
            return id;
          },
          unregisterCallback(id) {
            callbacks.delete(id);
          },
          async invoke(command, args) {
            window.__VISUALTEX_DOCUMENT_IMPORT_CALLS__.push({ command, args });
            if (command === "get_macos_offline_document_import_request") {
              return {
                protocolVersion: 1,
                sessionId: ${JSON.stringify(sessionId)},
                host: "word",
                sourceDocumentId: "visualtex-word-test-document",
                bookmarkName: "VT_D_12345678123442349234",
                defaultFontSizePt: 12,
              };
            }
            if (command === "commit_macos_offline_document_import") {
              window.__VISUALTEX_DOCUMENT_IMPORT_COMMIT__ = args;
              return null;
            }
            if (command === "cancel_macos_offline_document_import") return null;
            if (command === "close_macos_offline_office_editor_window") {
              window.__VISUALTEX_DOCUMENT_IMPORT_CLOSED__ = true;
              return null;
            }
            if (command === "plugin:event|listen" || command === "plugin:event|unlisten") {
              return 1;
            }
            throw new Error("Unexpected fake Tauri command: " + command);
          },
        };
      })();`,
    });
    await client.send("Page.navigate", { url: baseUrl });

    const evaluate = async (expression) => {
      const result = await client.send("Runtime.evaluate", {
        expression,
        awaitPromise: true,
        returnByValue: true,
      });
      if (result.exceptionDetails) {
        throw new Error(
          result.exceptionDetails.exception?.description ??
            result.exceptionDetails.text ??
            "Browser evaluation failed",
        );
      }
      return result.result.value;
    };

    const started = Date.now();
    while (Date.now() - started < 15_000) {
      const ready = await evaluate(
        `Boolean(document.querySelector(".document-import-app"))`,
      );
      if (ready) break;
      await sleep(80);
    }
    if (!(await evaluate(`Boolean(document.querySelector(".document-import-app"))`))) {
      const failure = await evaluate(`(() => ({
        text: document.body.innerText,
        html: document.getElementById("root")?.innerHTML?.slice(0, 2000) ?? "",
        calls: window.__VISUALTEX_DOCUMENT_IMPORT_CALLS__ ?? [],
      }))()`);
      const events = client.events.map((event) => ({
        method: event.method,
        exception:
          event.params?.exceptionDetails?.exception?.description ??
          event.params?.exceptionDetails?.text ??
          null,
        console:
          event.params?.args?.map((arg) => arg.value ?? arg.description ?? "") ??
          null,
      }));
      throw new Error(
        `Document importer did not mount: ${JSON.stringify({ failure, events })}`,
      );
    }

    const source = String.raw`正文中的行内公式 $p=mv$ 保持基线对齐。

\begin{equation}
E=mc^2
\end{equation}

\begin{align}
a &= b + c \\
d &= e
\end{align}

\begin{align*}
x &= y \\
y &= z
\end{align*}

结尾文字。`;
    await evaluate(`(() => {
      const textarea = document.querySelector(".document-import-source-pane textarea");
      if (!textarea) throw new Error("Missing document import source textarea");
      const setter = Object.getOwnPropertyDescriptor(
        HTMLTextAreaElement.prototype,
        "value",
      ).set;
      setter.call(textarea, ${JSON.stringify(source)});
      textarea.dispatchEvent(new Event("input", { bubbles: true }));
    })()`);

    const parseStarted = Date.now();
    while (Date.now() - parseStarted < 10_000) {
      if ((await evaluate(`document.querySelectorAll(".document-import-formula-card").length`)) === 4) {
        break;
      }
      await sleep(80);
    }

    const parsed = await evaluate(`(() => {
      const cards = [...document.querySelectorAll(".document-import-formula-card")];
      return {
        count: cards.length,
        modes: cards.map((card) => card.querySelector("select")?.value),
        numbered: cards.map((card) => card.querySelector('input[type="checkbox"]')?.checked ?? false),
        summary: document.querySelector(".document-import-summary")?.innerText ?? "",
      };
    })()`);
    if (
      parsed.count !== 4 ||
      parsed.modes.join(",") !== "inline,block,block,block" ||
      parsed.numbered.join(",") !== "false,true,true,false"
    ) {
      throw new Error(`Unexpected parsed formula blocks: ${JSON.stringify(parsed)}`);
    }

    await evaluate(`(() => {
      const imageRadio = document.querySelector(
        'input[name="document-formula-output"][type="radio"]:not(:checked)',
      );
      imageRadio?.click();
      const cards = [...document.querySelectorAll(".document-import-formula-card")];
      const sizes = [10.5, 18, 14, 16];
      cards.forEach((card, index) => {
        const input = card.querySelector('input[type="number"]');
        const setter = Object.getOwnPropertyDescriptor(
          HTMLInputElement.prototype,
          "value",
        ).set;
        setter.call(input, String(sizes[index]));
        input.dispatchEvent(new Event("input", { bubbles: true }));
        input.dispatchEvent(new Event("change", { bubbles: true }));
      });
      const insertButton = [...document.querySelectorAll("button")].find(
        (button) => button.textContent?.includes("插入 Word"),
      );
      if (!insertButton) throw new Error("Missing insert button");
      insertButton.click();
    })()`);

    const commitStarted = Date.now();
    let commit;
    while (Date.now() - commitStarted < 30_000) {
      commit = await evaluate(`window.__VISUALTEX_DOCUMENT_IMPORT_COMMIT__ ?? null`);
      if (commit) break;
      const error = await evaluate(
        `document.querySelector(".document-import-error")?.innerText ?? ""`,
      );
      if (error) throw new Error(`Document import UI reported: ${error}`);
      await sleep(100);
    }
    if (!commit) throw new Error("Document importer did not submit its Tauri commit");

    const input = commit.input;
    const formulas = input?.items?.filter((item) => item.kind === "formula") ?? [];
    const texts = input?.items?.filter((item) => item.kind === "text") ?? [];
    if (input?.outputKind !== "image" || formulas.length !== 4 || texts.length < 2) {
      throw new Error(`Unexpected document import commit: ${JSON.stringify(commit)}`);
    }
    if (
      formulas[0].displayMode !== "inline" ||
      formulas[0].numbered !== false ||
      formulas[0].fontSizePt !== 10.5 ||
      formulas[1].displayMode !== "block" ||
      formulas[1].numbered !== true ||
      formulas[1].fontSizePt !== 18 ||
      formulas[2].displayMode !== "block" ||
      formulas[2].numbered !== true ||
      formulas[2].fontSizePt !== 14 ||
      formulas[3].displayMode !== "block" ||
      formulas[3].numbered !== false ||
      formulas[3].fontSizePt !== 16
    ) {
      throw new Error(`Independent formula settings were lost: ${JSON.stringify(formulas)}`);
    }
    for (const formula of formulas) {
      if (
        !formula.formulaId ||
        !formula.metadata ||
        !formula.ommlBase64 ||
        !formula.ommlDocxBase64 ||
        !formula.svgBase64 ||
        formula.pngBase64 ||
        !(formula.width > 0) ||
        !(formula.height > 0)
      ) {
        throw new Error(
          `SVG fallback regression payload is invalid: ${JSON.stringify(formula)}`,
        );
      }
    }
    if (new Set(formulas.map((formula) => formula.formulaId)).size !== formulas.length) {
      throw new Error("Imported formulas did not receive independent identities");
    }
    const multilineExpectations = [
      {
        formula: formulas[2],
        codeFormat: "align",
        lines: ["a = b + c", "d = e"],
        environment: "align",
      },
      {
        formula: formulas[3],
        codeFormat: "align-star",
        lines: ["x = y", "y = z"],
        environment: "align*",
      },
    ];
    for (const expectation of multilineExpectations) {
      const metadataLines = expectation.formula.metadata.lines.map(
        (line) => line.latex,
      );
      if (
        expectation.formula.metadata.codeFormat !== expectation.codeFormat ||
        JSON.stringify(metadataLines) !== JSON.stringify(expectation.lines) ||
        expectation.formula.latex !== expectation.formula.metadata.latex ||
        !expectation.formula.latex.startsWith(
          `\\begin{${expectation.environment}}\n`,
        ) ||
        !expectation.formula.latex.endsWith(
          `\n\\end{${expectation.environment}}`,
        )
      ) {
        throw new Error(
          `Multiline formula canonical metadata is inconsistent: ${JSON.stringify(expectation.formula)}`,
        );
      }
    }

    const closed = await evaluate(
      `window.__VISUALTEX_DOCUMENT_IMPORT_CLOSED__ === true`,
    );
    if (!closed) throw new Error("Document importer did not request window close");

    console.log(JSON.stringify({ parsed, outputKind: input.outputKind, formulas: formulas.map((formula) => ({
      formulaId: formula.formulaId,
      displayMode: formula.displayMode,
      numbered: formula.numbered,
      fontSizePt: formula.fontSizePt,
      hasOmml: Boolean(formula.ommlBase64),
      hasSvg: Boolean(formula.svgBase64),
      usesBackendPngFallback: !formula.pngBase64,
    })) }, null, 2));
    console.log("Document import browser regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(500);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => {});
  }
}

await main();
