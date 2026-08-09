import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 650;
const vitePort = 23100 + offset;
const debugPort = 30100 + offset;
const baseUrl = `http://127.0.0.1:${vitePort}`;
const chromeProfile = `/tmp/visualtex-custom-symbol-designer-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      if ((await fetch(url)).ok) return;
    } catch {
      // Retry while local process starts.
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
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`CDP ${method} timed out`));
      }, 15000);
      this.pending.set(id, {
        resolve: (value) => {
          clearTimeout(timer);
          resolve(value);
        },
        reject: (error) => {
          clearTimeout(timer);
          reject(error);
        },
      });
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

async function waitUntil(client, expression, timeoutMs = 12000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const value = await client.evaluate(expression);
    if (value) return value;
    await sleep(60);
  }
  throw new Error(`Timed out waiting for ${expression}`);
}

async function setReactInput(client, selector, value) {
  await client.evaluate(`(() => {
    const input = document.querySelector(${JSON.stringify(selector)});
    if (!(input instanceof HTMLInputElement)) throw new Error(${JSON.stringify(`Input not found: ${selector}`)});
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
    setter?.call(input, ${JSON.stringify(String(value))});
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(new Event("change", { bubbles: true }));
    return input.value;
  })()`);
  await sleep(100);
}

async function setReactSelect(client, selector, value) {
  await client.evaluate(`(() => {
    const select = document.querySelector(${JSON.stringify(selector)});
    if (!select || select.tagName !== "SELECT") {
      throw new Error(${JSON.stringify(`Select not found: ${selector}`)} +
        " registrationPanel=" + Boolean(document.querySelector('[data-custom-symbol-registration-panel]')));
    }
    const setter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value")?.set;
    if (setter) setter.call(select, ${JSON.stringify(String(value))});
    else select.value = ${JSON.stringify(String(value))};
    select.dispatchEvent(new Event("input", { bubbles: true }));
    select.dispatchEvent(new Event("change", { bubbles: true }));
    return select.value;
  })()`);
  await sleep(100);
}

async function main() {
  const vite = spawn(
    process.execPath,
    [
      "node_modules/vite/bin/vite.js",
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
        "--window-size=1500,1000",
        baseUrl,
      ],
      { stdio: "ignore" },
    );
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (
      await fetch(`http://127.0.0.1:${debugPort}/json/list`)
    ).json();
    const page = targets.find(
      (target) => target.type === "page" && target.url.startsWith(baseUrl),
    );
    assert.ok(page, "VisualTeX browser target must exist");
    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(450);

    await client.evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
      localStorage.setItem("visualtex.onboarding.macos.desktop.v1.2.0.completed", "true");
      localStorage.setItem("visualtex.office.macos.native-first-run.v1.2.0.completed", "true");
      localStorage.removeItem("visualtex.custom-symbols.v1");
      const key = "visualtex-editor";
      const persisted = JSON.parse(localStorage.getItem(key) || "{}");
      persisted.state = {
        ...(persisted.state || {}),
        checkUpdatesOnStartup: false,
        theme: "proof",
      };
      localStorage.setItem(key, JSON.stringify(persisted));
      return true;
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(550);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(550);
    await waitUntil(client, `Boolean(document.querySelector("math-field"))`);

    const baseline = await client.evaluate(`(() => ({
      formula: document.querySelector("math-field")?.value || "",
      storage: localStorage.getItem("visualtex.custom-symbols.v1"),
      theme: document.documentElement.dataset.theme || "",
    }))()`);
    assert.equal(baseline.theme, "proof");
    process.stdout.write("[custom-symbol-designer] Proof editor ready\n");

    await client.evaluate(`(() => {
      const tilesTab = document.querySelector('[data-toolbar-view="tiles"]');
      if (tilesTab instanceof HTMLButtonElement) tilesTab.click();
      return true;
    })()`);
    await sleep(100);
    await waitUntil(client, `Boolean(document.querySelector('[data-tile-category="custom"]'))`);
    await client.evaluate(`document.querySelector('[data-tile-category="custom"]').click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-open-custom-symbol-designer]'))`);
    await client.evaluate(`document.querySelector('[data-open-custom-symbol-designer]').click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-custom-symbol-designer]'))`);

    const themePurity = await client.evaluate(`(() => {
      const selectors = [
        ".custom-symbol-designer-dialog",
        ".custom-symbol-designer-sidebar.is-materials",
        ".custom-symbol-designer-panel",
        ".custom-symbol-designer-canvas",
      ];
      return selectors.map((selector) => {
        const element = document.querySelector(selector);
        return {
          selector,
          background: element ? getComputedStyle(element).backgroundColor : "missing",
        };
      });
    })()`);
    for (const entry of themePurity) {
      assert.notEqual(entry.background, "missing", entry.selector);
      assert.notEqual(
        entry.background,
        "rgb(255, 255, 255)",
        `${entry.selector} leaked a fixed white background under Proof`,
      );
    }
    process.stdout.write("[custom-symbol-designer] Proof theme purity verified\n");

    await client.evaluate(`document.querySelector('[data-add-custom-symbol-material]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer]').length === 1`);
    const first = await client.evaluate(`(() => {
      const layer = document.querySelector('[data-custom-symbol-layer]');
      const canvasLayer = document.querySelector('[data-custom-symbol-canvas-layer]');
      const artworkLayer = document.querySelector('[data-custom-symbol-artwork-layer]');
      const canvas = document.querySelector('[data-custom-symbol-canvas]');
      const paper = document.querySelector('[data-custom-symbol-canvas-paper]');
      return {
        layerId: layer?.getAttribute("data-custom-symbol-layer") || "",
        hasCanvasLayer: Boolean(canvasLayer),
        hasPath: Boolean(artworkLayer?.querySelector("path")),
        hasBaseline: Boolean(document.querySelector('[data-custom-symbol-baseline]')),
        hasReferenceAlpha: Boolean(document.querySelector('[data-custom-symbol-reference-alpha]')),
        hasFitControl: Boolean(document.querySelector('[data-custom-symbol-fit-view]')),
        viewBox: canvas?.getAttribute("viewBox") || "",
        paperWidth: Number(paper?.getAttribute("width") || 0),
        paperHeight: Number(paper?.getAttribute("height") || 0),
        paperRect: paper ? (() => {
          const rect = paper.getBoundingClientRect();
          return { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom };
        })() : null,
        shellRect: (() => {
          const shell = document.querySelector('[data-custom-symbol-canvas-shell]');
          if (!shell) return null;
          const rect = shell.getBoundingClientRect();
          return { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom };
        })(),
      };
    })()`);
    assert.ok(first.layerId);
    assert.equal(first.hasCanvasLayer, true);
    assert.equal(first.hasPath, true);
    assert.equal(first.hasBaseline, true);
    assert.equal(first.hasReferenceAlpha, true, "Reference alpha should be visible by default");
    assert.equal(first.hasFitControl, true);
    const firstViewBox = first.viewBox.split(/\s+/).map(Number);
    assert.equal(firstViewBox.length, 4);
    assert.ok(firstViewBox.every(Number.isFinite));
    assert.ok(firstViewBox[0] < 0 && firstViewBox[1] < 0, "Fit view should leave visible space above and left of the mathematical canvas");
    assert.ok(firstViewBox[2] > first.paperWidth && firstViewBox[3] > first.paperHeight, "Fit view should leave visible space around the entire mathematical canvas");
    assert.ok(first.paperRect && first.shellRect);
    assert.ok(first.paperRect.left >= first.shellRect.left - 1, "Fitted canvas paper must not be clipped on the left");
    assert.ok(first.paperRect.top >= first.shellRect.top - 1, "Fitted canvas paper must not be clipped on the top");
    assert.ok(first.paperRect.right <= first.shellRect.right + 1, "Fitted canvas paper must not be clipped on the right");
    assert.ok(first.paperRect.bottom <= first.shellRect.bottom + 1, "Fitted canvas paper must not be clipped on the bottom");

    await client.evaluate(`document.querySelector('[data-toggle-custom-symbol-reference-alpha]').click()`);
    await waitUntil(client, `!document.querySelector('[data-custom-symbol-reference-alpha]')`);
    await client.evaluate(`document.querySelector('[data-toggle-custom-symbol-reference-alpha]').click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-custom-symbol-reference-alpha]'))`);

    const zoomBefore = await client.evaluate(`document.querySelector('[data-custom-symbol-canvas]')?.getAttribute('viewBox') || ''`);
    await client.evaluate(`document.querySelector('[data-custom-symbol-zoom-in]').click()`);
    await sleep(80);
    const zoomAfter = await client.evaluate(`document.querySelector('[data-custom-symbol-canvas]')?.getAttribute('viewBox') || ''`);
    assert.notEqual(zoomAfter, zoomBefore, "Zoom-in control must change the mathematical viewport");
    await client.evaluate(`document.querySelector('[data-custom-symbol-fit-view]').click()`);
    await sleep(80);

    const workspaceState = await client.evaluate(`(() => {
      const canvas = document.querySelector('[data-custom-symbol-canvas]');
      const output = document.querySelector('[data-custom-symbol-output-box]');
      const workspace = (canvas?.getAttribute('data-custom-symbol-workspace') || '').split(/\\s+/).map(Number);
      return {
        workspace,
        outputWidth: Number(output?.getAttribute('width') || 0),
        outputHeight: Number(output?.getAttribute('height') || 0),
        alphaRect: (() => {
          const reference = document.querySelector('[data-custom-symbol-reference]');
          if (!reference) return null;
          const rect = reference.getBoundingClientRect();
          return { width: rect.width, height: rect.height };
        })(),
      };
    })()`);
    assert.equal(workspaceState.workspace.length, 4);
    assert.ok(workspaceState.outputWidth >= 3200, "Default output width should leave generous room for mathematical construction");
    assert.ok(workspaceState.outputHeight >= 4500, "Default output height should leave generous room above and below the baseline");
    assert.ok(workspaceState.workspace[2] > workspaceState.outputWidth * 3, "Designer workspace should be several times wider than the final output box");
    assert.ok(workspaceState.workspace[3] > workspaceState.outputHeight * 3, "Designer workspace should be several times taller than the final output box");
    await client.evaluate(`document.querySelector('[data-custom-symbol-fit-workspace]').click()`);
    await sleep(80);
    const workspaceView = await client.evaluate(`document.querySelector('[data-custom-symbol-canvas]')?.getAttribute('viewBox') || ''`);
    assert.equal(
      workspaceView.split(/\s+/).map((value) => Math.round(Number(value))).join(','),
      workspaceState.workspace.map((value) => Math.round(value)).join(','),
      "Workspace control must reveal the full large design area",
    );
    await client.evaluate(`document.querySelector('[data-custom-symbol-fit-view]').click()`);
    await sleep(80);

    await setReactSelect(client, '[data-custom-symbol-reference-select]', String.raw`\displaystyle\sum`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-reference]')?.getAttribute('data-custom-symbol-reference-label') === 'Σ'`);
    const sumReference = await client.evaluate(`(() => {
      const reference = document.querySelector('[data-custom-symbol-reference]');
      if (!reference) return null;
      const rect = reference.getBoundingClientRect();
      return { width: rect.width, height: rect.height };
    })()`);
    assert.ok(sumReference && workspaceState.alphaRect);
    assert.ok(
      sumReference.height > workspaceState.alphaRect.height,
      "Large-operator reference should visibly differ from ordinary alpha size",
    );
    const referenceFitCases = [
      [String.raw`\displaystyle\int`, "∫"],
      [String.raw`\displaystyle\oint`, "∮"],
      [String.raw`\displaystyle\sum`, "Σ"],
      [String.raw`\displaystyle\prod`, "Π"],
      [String.raw`\displaystyle\bigcup`, "⋃"],
    ];
    for (const [latex, label] of referenceFitCases) {
      await setReactSelect(client, '[data-custom-symbol-reference-select]', latex);
      await waitUntil(
        client,
        `document.querySelector('[data-custom-symbol-reference]')?.getAttribute('data-custom-symbol-reference-label') === ${JSON.stringify(label)}`,
      );
      await client.evaluate(`document.querySelector('[data-custom-symbol-fit-view]').click()`);
      await sleep(70);
      const fit = await client.evaluate(`(() => {
        const reference = document.querySelector('[data-custom-symbol-reference]');
        const output = document.querySelector('[data-custom-symbol-output-box]');
        if (!reference || !output) return null;
        const rr = reference.getBoundingClientRect();
        const or = output.getBoundingClientRect();
        return {
          reference: { left: rr.left, top: rr.top, right: rr.right, bottom: rr.bottom },
          output: { left: or.left, top: or.top, right: or.right, bottom: or.bottom },
        };
      })()`);
      assert.ok(fit, `${label} reference and output box must exist`);
      assert.ok(fit.reference.left > fit.output.left + 2, `${label} reference must stay inside the output box on the left`);
      assert.ok(fit.reference.right < fit.output.right - 2, `${label} reference must stay inside the output box on the right`);
      assert.ok(fit.reference.top > fit.output.top + 2, `${label} reference must stay inside the output box at the top`);
      assert.ok(fit.reference.bottom < fit.output.bottom - 2, `${label} reference must stay inside the output box at the bottom`);
    }

    await setReactSelect(client, '[data-custom-symbol-reference-select]', String.raw`\displaystyle\int`);
    await setReactInput(client, '[data-custom-symbol-material-input]', String.raw`\int`);
    await client.evaluate(`document.querySelector('[data-add-custom-symbol-material]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer]').length === 2`);
    const integralConsistency = await client.evaluate(`(() => {
      const selected = document.querySelector('[data-custom-symbol-layer].is-selected');
      const artwork = selected
        ? document.querySelector('[data-custom-symbol-artwork-layer="' + selected.getAttribute('data-custom-symbol-layer') + '"]')
        : null;
      const reference = document.querySelector('[data-custom-symbol-reference]');
      return {
        source: selected?.getAttribute('data-layer-source-latex') || '',
        artworkPaths: Array.from(artwork?.querySelectorAll('path') || []).map((path) => path.getAttribute('d') || ''),
        referencePaths: Array.from(reference?.querySelectorAll('path') || []).map((path) => path.getAttribute('d') || ''),
      };
    })()`);
    assert.equal(integralConsistency.source, String.raw`\displaystyle\int`);
    assert.deepEqual(
      integralConsistency.artworkPaths,
      integralConsistency.referencePaths,
      "A quick/material \\int must use the exact same displaystyle MathJax glyph outline as the integral reference",
    );
    await client.evaluate(`document.querySelector('[data-delete-custom-symbol-layer]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer]').length === 1`);
    await client.evaluate(`document.querySelector('[data-custom-symbol-layer]').click()`);
    process.stdout.write("[custom-symbol-designer] material/reference integral outline consistency verified\n");

    await client.evaluate(`document.querySelector('[data-metric-preset="large"]').click()`);
    await waitUntil(client, `Number(document.querySelector('[data-designer-field="canvas-width"]')?.value || 0) >= 4.5`);
    await client.evaluate(`document.querySelector('[data-metric-preset="standard"]').click()`);
    await setReactSelect(client, '[data-custom-symbol-reference-select]', String.raw`\alpha`);
    await waitUntil(client, `Boolean(document.querySelector('[data-custom-symbol-reference-alpha]'))`);
    process.stdout.write("[custom-symbol-designer] large workspace, output presets and multi-size references verified\n");
    process.stdout.write("[custom-symbol-designer] viewport fit/zoom and alpha reference verified\n");
    process.stdout.write("[custom-symbol-designer] LaTeX glyph added\n");

    await setReactInput(client, '[data-designer-field="layer-x"]', 80);
    await setReactInput(client, '[data-designer-field="layer-scale-x"]', 1.25);
    await setReactInput(client, '[data-designer-field="layer-rotation"]', 12);
    const transformed = await client.evaluate(`(() => ({
      transform: document.querySelector('[data-custom-symbol-canvas-layer]')?.getAttribute("transform") || "",
      x: document.querySelector('[data-designer-field="layer-x"]')?.value || "",
      scaleX: document.querySelector('[data-designer-field="layer-scale-x"]')?.value || "",
      rotation: document.querySelector('[data-designer-field="layer-rotation"]')?.value || "",
    }))()`);
    assert.match(transformed.transform, /translate\(80 /);
    assert.match(transformed.transform, /rotate\(12\)/);
    assert.match(transformed.transform, /scale\(1\.25 /);
    process.stdout.write("[custom-symbol-designer] numeric transforms verified\n");

    const scaleBeforeHandle = await client.evaluate(`(() => ({
      x: Number(document.querySelector('[data-designer-field="layer-scale-x"]')?.value || 1),
      y: Number(document.querySelector('[data-designer-field="layer-scale-y"]')?.value || 1),
      handles: document.querySelectorAll('[data-custom-symbol-resize-handle]').length,
      hitTargets: document.querySelectorAll('[data-custom-symbol-resize-hit-target]').length,
    }))()`);
    assert.equal(scaleBeforeHandle.handles, 8, "Selected layer should expose eight visual resize handles");
    assert.equal(scaleBeforeHandle.hitTargets, 8, "Selected layer should expose eight forgiving resize hit targets");
    const resizeBox = await client.evaluate(`(() => {
      const handle = document.querySelector('[data-custom-symbol-resize-handle="se"]');
      const hitTarget = document.querySelector('[data-custom-symbol-resize-hit-target="se"]');
      if (!handle || !hitTarget) return null;
      const handleRect = handle.getBoundingClientRect();
      const hitRect = hitTarget.getBoundingClientRect();
      const x = handleRect.right + Math.max(1, (hitRect.right - handleRect.right) * 0.5);
      const y = handleRect.top + handleRect.height / 2;
      return {
        x,
        y,
        outsideVisualHandle: x > handleRect.right,
        insideHitTarget:
          x >= hitRect.left && x <= hitRect.right && y >= hitRect.top && y <= hitRect.bottom,
        visualWidth: handleRect.width,
        hitWidth: hitRect.width,
      };
    })()`);
    assert.ok(resizeBox, "Bottom-right resize handle must have a forgiving hit target");
    assert.equal(resizeBox.outsideVisualHandle, true, "Resize regression must begin outside the painted handle");
    assert.equal(resizeBox.insideHitTarget, true, "Resize regression must begin inside the expanded hit target");
    assert.ok(resizeBox.hitWidth >= resizeBox.visualWidth * 2.5, "Resize hit target should be substantially larger than the painted handle");
    await client.send("Input.dispatchMouseEvent", {
      type: "mousePressed",
      x: resizeBox.x,
      y: resizeBox.y,
      button: "left",
      buttons: 1,
      clickCount: 1,
    });
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseMoved",
      x: resizeBox.x + 35,
      y: resizeBox.y + 30,
      button: "left",
      buttons: 1,
    });
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseReleased",
      x: resizeBox.x + 35,
      y: resizeBox.y + 30,
      button: "left",
      buttons: 0,
      clickCount: 1,
    });
    await sleep(100);
    const scaleAfterHandle = await client.evaluate(`(() => ({
      x: Number(document.querySelector('[data-designer-field="layer-scale-x"]')?.value || 1),
      y: Number(document.querySelector('[data-designer-field="layer-scale-y"]')?.value || 1),
    }))()`);
    assert.ok(
      scaleAfterHandle.x !== scaleBeforeHandle.x || scaleAfterHandle.y !== scaleBeforeHandle.y,
      "Dragging a resize handle must change the selected layer scale",
    );
    process.stdout.write("[custom-symbol-designer] forgiving resize hit targets verified\n");

    const viewBoxBeforePan = await client.evaluate(`document.querySelector('[data-custom-symbol-canvas]')?.getAttribute('viewBox') || ''`);
    const emptyCanvasPoint = await client.evaluate(`(() => {
      const canvas = document.querySelector('[data-custom-symbol-canvas]');
      if (!canvas) return null;
      const rect = canvas.getBoundingClientRect();
      return { x: rect.left + 18, y: rect.top + rect.height / 2 };
    })()`);
    assert.ok(emptyCanvasPoint);
    await client.send("Input.dispatchMouseEvent", {
      type: "mousePressed",
      x: emptyCanvasPoint.x,
      y: emptyCanvasPoint.y,
      button: "left",
      buttons: 1,
      clickCount: 1,
    });
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseMoved",
      x: emptyCanvasPoint.x + 45,
      y: emptyCanvasPoint.y + 18,
      button: "left",
      buttons: 1,
    });
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseReleased",
      x: emptyCanvasPoint.x + 45,
      y: emptyCanvasPoint.y + 18,
      button: "left",
      buttons: 0,
      clickCount: 1,
    });
    await sleep(80);
    const viewBoxAfterPan = await client.evaluate(`document.querySelector('[data-custom-symbol-canvas]')?.getAttribute('viewBox') || ''`);
    assert.notEqual(viewBoxAfterPan, viewBoxBeforePan, "Dragging empty canvas space must pan the viewport");
    await client.evaluate(`document.querySelector('[data-custom-symbol-fit-view]').click()`);
    await sleep(80);
    await client.evaluate(`document.querySelector('[data-custom-symbol-layer]').click()`);
    process.stdout.write("[custom-symbol-designer] canvas pan verified\n");

    await client.evaluate(`document.querySelector('[data-duplicate-custom-symbol-layer]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer]').length === 2`);
    const duplicateState = await client.evaluate(`(() => ({
      list: document.querySelectorAll('[data-custom-symbol-layer]').length,
      canvas: document.querySelectorAll('[data-custom-symbol-canvas-layer]').length,
      selected: document.querySelector('[data-custom-symbol-layer].is-selected')?.getAttribute('data-custom-symbol-layer') || "",
    }))()`);
    assert.equal(duplicateState.list, 2);
    assert.equal(duplicateState.canvas, 2);
    assert.ok(duplicateState.selected);

    await client.evaluate(`document.querySelector('[data-custom-symbol-layer].is-selected [data-toggle-custom-symbol-layer-visibility]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-canvas-layer]').length === 1`);
    await client.evaluate(`document.querySelector('[data-custom-symbol-layer].is-selected [data-toggle-custom-symbol-layer-visibility]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-canvas-layer]').length === 2`);
    await client.evaluate(`document.querySelector('[data-custom-symbol-layer].is-selected [data-toggle-custom-symbol-layer-lock]').click()`);
    const locked = await client.evaluate(`document.querySelector('[data-custom-symbol-layer].is-selected')?.getAttribute('data-layer-locked')`);
    assert.equal(locked, "true");
    await client.evaluate(`document.querySelector('[data-custom-symbol-layer].is-selected [data-toggle-custom-symbol-layer-lock]').click()`);
    process.stdout.write("[custom-symbol-designer] duplicate, visibility and lock verified\n");

    const beforeDrag = Number(
      await client.evaluate(`document.querySelector('[data-designer-field="layer-x"]')?.value || "0"`),
    );
    const box = await client.evaluate(`(() => {
      const layer = document.querySelector('[data-custom-symbol-canvas-layer].is-selected') ||
        document.querySelector('[data-custom-symbol-canvas-layer]');
      if (!layer) return null;
      const rect = layer.getBoundingClientRect();
      return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 };
    })()`);
    assert.ok(box, "Selected canvas layer must have a hit box");
    await client.send("Input.dispatchMouseEvent", {
      type: "mousePressed",
      x: box.x,
      y: box.y,
      button: "left",
      buttons: 1,
      clickCount: 1,
    });
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseMoved",
      x: box.x + 45,
      y: box.y + 25,
      button: "left",
      buttons: 1,
    });
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseReleased",
      x: box.x + 45,
      y: box.y + 25,
      button: "left",
      buttons: 0,
      clickCount: 1,
    });
    await sleep(140);
    const afterDrag = Number(
      await client.evaluate(`document.querySelector('[data-designer-field="layer-x"]')?.value || "0"`),
    );
    assert.notEqual(afterDrag, beforeDrag, "Canvas drag must update mathematical X");
    process.stdout.write("[custom-symbol-designer] CTM drag verified\n");

    await client.evaluate(`document.querySelector('[data-delete-custom-symbol-layer]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer]').length === 1`);

    await client.evaluate(`(() => {
      const closeButtons = Array.from(document.querySelectorAll('.custom-symbol-designer-footer button'));
      closeButtons[0]?.click();
      return true;
    })()`);
    await waitUntil(client, `!document.querySelector('[data-custom-symbol-designer]')`);
    const afterClose = await client.evaluate(`(() => ({
      formula: document.querySelector("math-field")?.value || "",
      storage: localStorage.getItem("visualtex.custom-symbols.v1"),
    }))()`);
    assert.equal(afterClose.formula, baseline.formula);
    assert.equal(afterClose.storage, baseline.storage);

    await client.evaluate(`document.querySelector('[data-open-custom-symbol-designer]').click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-custom-symbol-designer]'))`);
    const reopened = await client.evaluate(`(() => ({
      layers: document.querySelectorAll('[data-custom-symbol-layer]').length,
      storage: localStorage.getItem("visualtex.custom-symbols.v1"),
      formula: document.querySelector("math-field")?.value || "",
    }))()`);
    assert.equal(reopened.layers, 1, "Designer draft should survive close/reopen in the same session");
    assert.equal(reopened.storage, baseline.storage);
    assert.equal(reopened.formula, baseline.formula);

    await client.evaluate(`document.querySelector('[data-custom-symbol-layer]').click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-custom-symbol-layer].is-selected'))`);
    await client.evaluate(`document.querySelector('[data-crop-preset="top"]').click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-designer-field="crop-height"]'))`);
    const croppedBeforeMove = await client.evaluate(`(() => {
      const interaction = document.querySelector('[data-custom-symbol-canvas-layer].is-selected');
      const id = interaction?.getAttribute('data-custom-symbol-canvas-layer') || '';
      const artwork = id ? document.querySelector('[data-custom-symbol-artwork-layer="' + id + '"]') : null;
      const clipRect = artwork?.querySelector('clipPath rect');
      return {
        transform: interaction?.getAttribute('transform') || "",
        clip: clipRect
          ? [clipRect.getAttribute('x'), clipRect.getAttribute('y'), clipRect.getAttribute('width'), clipRect.getAttribute('height')]
          : [],
        cropHeight: document.querySelector('[data-designer-field="crop-height"]')?.value || "",
      };
    })()`);
    assert.equal(croppedBeforeMove.clip.length, 4);
    assert.ok(Number(croppedBeforeMove.cropHeight) > 0);
    await setReactInput(client, '[data-designer-field="layer-x"]', 140);
    const croppedAfterMove = await client.evaluate(`(() => {
      const interaction = document.querySelector('[data-custom-symbol-canvas-layer].is-selected');
      const id = interaction?.getAttribute('data-custom-symbol-canvas-layer') || '';
      const artwork = id ? document.querySelector('[data-custom-symbol-artwork-layer="' + id + '"]') : null;
      const clipRect = artwork?.querySelector('clipPath rect');
      return {
        transform: interaction?.getAttribute('transform') || "",
        clip: clipRect
          ? [clipRect.getAttribute('x'), clipRect.getAttribute('y'), clipRect.getAttribute('width'), clipRect.getAttribute('height')]
          : [],
      };
    })()`);
    assert.match(croppedAfterMove.transform, /translate\(140 /);
    assert.deepEqual(
      croppedAfterMove.clip,
      croppedBeforeMove.clip,
      "Layer-local crop must move with the layer instead of changing its crop coordinates",
    );
    process.stdout.write("[custom-symbol-designer] local crop semantics verified\n");

    await client.evaluate(`document.querySelector('[data-split-custom-symbol-glyph="horizontal"]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer]').length === 4`);
    const sliced = await client.evaluate(`(() => {
      const list = Array.from(document.querySelectorAll('[data-custom-symbol-layer]'));
      const visibleCanvas = Array.from(document.querySelectorAll('[data-custom-symbol-artwork-layer]'));
      const paths = visibleCanvas.map((layer) => layer.querySelector('path')?.getAttribute('d') || "");
      return {
        listCount: list.length,
        hiddenCount: list.filter((layer) => layer.getAttribute('data-layer-visible') === 'false').length,
        visibleCanvasCount: visibleCanvas.length,
        cropCount: visibleCanvas.filter((layer) => layer.querySelector('clipPath rect')).length,
        uniquePaths: Array.from(new Set(paths.filter(Boolean))).length,
      };
    })()`);
    assert.equal(sliced.listCount, 4);
    assert.equal(sliced.hiddenCount, 1, "Original full glyph must remain as a hidden recovery layer");
    assert.equal(sliced.visibleCanvasCount, 3);
    assert.equal(sliced.cropCount, 3);
    assert.equal(
      sliced.uniquePaths,
      1,
      "All slices must keep the same full source path and differ only by clipRect",
    );
    process.stdout.write("[custom-symbol-designer] non-destructive three-way slicing verified\n");

    const geometryPresets = ["line", "circle", "ellipse", "rect", "triangle", "arrow", "arc"];
    for (const preset of geometryPresets) {
      await client.evaluate(`document.querySelector('[data-add-custom-symbol-geometry="${preset}"]').click()`);
      await sleep(45);
    }
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer-kind="geometry"]').length === 7`);
    const geometryState = await client.evaluate(`(() => ({
      listCount: document.querySelectorAll('[data-custom-symbol-layer]').length,
      geometryCanvasCount: document.querySelectorAll('[data-custom-symbol-artwork-kind="geometry"]').length,
      totalCanvasCount: document.querySelectorAll('[data-custom-symbol-artwork-layer]').length,
      geometryTags: Array.from(document.querySelectorAll('[data-custom-symbol-artwork-kind="geometry"]')).map((layer) =>
        Array.from(layer.children).map((child) => child.tagName.toLowerCase()).filter((tag) => tag !== 'defs')
      ).flat(),
    }))()`);
    assert.equal(geometryState.listCount, 11);
    assert.equal(geometryState.geometryCanvasCount, 7);
    assert.equal(geometryState.totalCanvasCount, 10);
    assert.ok(geometryState.geometryTags.includes("line"));
    assert.ok(geometryState.geometryTags.includes("circle"));
    assert.ok(geometryState.geometryTags.includes("ellipse"));
    assert.ok(geometryState.geometryTags.includes("polygon"));
    assert.ok(geometryState.geometryTags.filter((tag) => tag === "path").length >= 2);

    await client.evaluate(`document.querySelector('[data-layer-geometry-preset="rect"]')?.click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-custom-symbol-geometry-properties]'))`);
    await setReactInput(client, '[data-designer-field="geometry-width"]', 520);
    await setReactInput(client, '[data-designer-field="geometry-height"]', 130);
    await setReactInput(client, '[data-designer-field="geometry-stroke-width"]', 10);
    await setReactInput(client, '[data-designer-field="geometry-corner-radius"]', 6);
    const rectGeometry = await client.evaluate(`(() => {
      const rect = document.querySelector('[data-custom-symbol-artwork-preset="rect"] rect');
      return rect ? {
        width: Number(rect.getAttribute('width')),
        height: Number(rect.getAttribute('height')),
        strokeWidth: Number(rect.getAttribute('stroke-width')),
        rx: Number(rect.getAttribute('rx')),
        fill: rect.getAttribute('fill'),
      } : null;
    })()`);
    assert.ok(rectGeometry);
    assert.equal(rectGeometry.width, 520);
    assert.equal(rectGeometry.height, 130);
    assert.equal(rectGeometry.strokeWidth, 10);
    assert.equal(rectGeometry.rx, 6);
    assert.equal(rectGeometry.fill, "none");
    await client.evaluate(`document.querySelector('[data-geometry-fill]')?.click()`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-artwork-preset="rect"] rect')?.getAttribute('fill') !== 'none'`);
    process.stdout.write("[custom-symbol-designer] direct geometry width/height/stroke/fill controls verified\n");

    await client.evaluate(`document.querySelector('[data-custom-symbol-eraser-tool]').click()`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-eraser-tool]')?.getAttribute('aria-pressed') === 'true'`);
    await setReactInput(client, '[data-custom-symbol-eraser-size-number]', 18);
    const eraserCanvas = await client.evaluate(`(() => {
      const canvas = document.querySelector('[data-custom-symbol-canvas]');
      const rect = canvas?.getBoundingClientRect();
      return rect ? { left: rect.left, top: rect.top, width: rect.width, height: rect.height } : null;
    })()`);
    assert.ok(eraserCanvas);
    const eraseStart = { x: eraserCanvas.left + eraserCanvas.width * 0.42, y: eraserCanvas.top + eraserCanvas.height * 0.49 };
    const eraseEnd = { x: eraserCanvas.left + eraserCanvas.width * 0.58, y: eraserCanvas.top + eraserCanvas.height * 0.53 };
    await client.send("Input.dispatchMouseEvent", { type: "mouseMoved", x: eraseStart.x, y: eraseStart.y, button: "none", buttons: 0 });
    await waitUntil(client, `Boolean(document.querySelector('[data-custom-symbol-eraser-cursor]'))`);
    const cursorState = await client.evaluate(`(() => {
      const cursor = document.querySelector('[data-custom-symbol-eraser-cursor] circle');
      return cursor ? { radius: Number(cursor.getAttribute('r')) } : null;
    })()`);
    assert.ok(cursorState);
    assert.ok(Math.abs(cursorState.radius - 9) < 0.001, "Eraser cursor radius must match half the precise 18-unit erase width");
    await client.send("Input.dispatchMouseEvent", { type: "mousePressed", x: eraseStart.x, y: eraseStart.y, button: "left", buttons: 1, clickCount: 1 });
    await client.send("Input.dispatchMouseEvent", { type: "mouseMoved", x: eraseStart.x + (eraseEnd.x - eraseStart.x) * 0.3, y: eraseStart.y + 7, button: "left", buttons: 1 });
    await client.send("Input.dispatchMouseEvent", { type: "mouseMoved", x: eraseStart.x + (eraseEnd.x - eraseStart.x) * 0.62, y: eraseEnd.y - 9, button: "left", buttons: 1 });
    await client.send("Input.dispatchMouseEvent", { type: "mouseMoved", x: eraseEnd.x, y: eraseEnd.y, button: "left", buttons: 1 });
    await client.send("Input.dispatchMouseEvent", { type: "mouseReleased", x: eraseEnd.x, y: eraseEnd.y, button: "left", buttons: 0, clickCount: 1 });
    await waitUntil(client, `Boolean(document.querySelector('[data-layer-geometry-preset="eraser"]'))`);
    const eraserState = await client.evaluate(`(() => {
      const eraserArtwork = document.querySelector('#visualtex-custom-symbol-designer-erase-mask path');
      return {
        layerOperation: document.querySelector('[data-custom-symbol-layer-operation="erase"]')?.getAttribute('data-custom-symbol-layer-operation') || '',
        mask: Boolean(document.querySelector('#visualtex-custom-symbol-designer-erase-mask')),
        eraserPreset: document.querySelector('[data-layer-geometry-preset="eraser"]')?.getAttribute('data-layer-geometry-preset') || '',
        d: eraserArtwork?.getAttribute('d') || '',
        strokeWidth: Number(eraserArtwork?.getAttribute('stroke-width') || 0),
        legacyOverlay: Boolean(document.querySelector('.custom-symbol-designer-eraser-overlay, .custom-symbol-designer-live-eraser')),
        selectedCenterline: Boolean(document.querySelector('[data-custom-symbol-eraser-centerline]')),
      };
    })()`);
    assert.equal(eraserState.layerOperation, "erase");
    assert.equal(eraserState.mask, true);
    assert.equal(eraserState.eraserPreset, "eraser");
    assert.match(eraserState.d, /C/, "Dragged eraser strokes must persist as a smooth cubic path instead of a polyline");
    assert.equal(eraserState.strokeWidth, 18);
    assert.equal(eraserState.legacyOverlay, false, "Completed eraser strokes must not leave the old thick red overlay");
    assert.equal(eraserState.selectedCenterline, false, "Centerline stays hidden while eraser mode is active");

    const clickPoint = { x: eraserCanvas.left + eraserCanvas.width * 0.36, y: eraserCanvas.top + eraserCanvas.height * 0.58 };
    const eraserCountBeforeClick = await client.evaluate(`document.querySelectorAll('[data-layer-geometry-preset="eraser"]').length`);
    await client.send("Input.dispatchMouseEvent", { type: "mousePressed", x: clickPoint.x, y: clickPoint.y, button: "left", buttons: 1, clickCount: 1 });
    await client.send("Input.dispatchMouseEvent", { type: "mouseReleased", x: clickPoint.x, y: clickPoint.y, button: "left", buttons: 0, clickCount: 1 });
    await waitUntil(client, `document.querySelectorAll('[data-layer-geometry-preset="eraser"]').length === ${eraserCountBeforeClick + 1}`);
    await client.evaluate(`document.querySelector('[data-custom-symbol-eraser-tool]').click()`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-eraser-tool]')?.getAttribute('aria-pressed') === 'false'`);
    const selectedCenterlineVisible = await client.evaluate(`Boolean(document.querySelector('[data-custom-symbol-eraser-centerline]'))`);
    assert.equal(selectedCenterlineVisible, true, "Selected erase strokes should expose only a thin editable centerline after leaving eraser mode");
    process.stdout.write("[custom-symbol-designer] smooth precise vector eraser interaction verified\n");

    await setReactInput(client, '[data-designer-field="layer-rotation"]', 27);
    const geometryTransform = await client.evaluate(`document.querySelector('[data-custom-symbol-canvas-layer].is-selected')?.getAttribute('transform') || ""`);
    assert.match(geometryTransform, /rotate\(27\)/);
    process.stdout.write("[custom-symbol-designer] geometry layers verified\n");

    const finalIsolation = await client.evaluate(`(() => ({
      formula: document.querySelector("math-field")?.value || "",
      storage: localStorage.getItem("visualtex.custom-symbols.v1"),
    }))()`);
    assert.equal(finalIsolation.formula, baseline.formula);
    assert.equal(finalIsolation.storage, baseline.storage);
    process.stdout.write("[custom-symbol-designer] pre-registration isolation verified\n");

    await client.evaluate(`document.querySelector('[data-reset-custom-symbol-designer]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer]').length === 0`);
    await client.evaluate(`document.querySelector('[data-add-custom-symbol-material]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer]').length === 1`);
    const preRegistrationMetrics = await client.evaluate(`(() => ({
      widthEm: Number(document.querySelector('[data-designer-field="canvas-width"]')?.value || 0),
      ascentEm: Number(document.querySelector('[data-designer-field="canvas-ascent"]')?.value || 0),
      descentEm: Number(document.querySelector('[data-designer-field="canvas-descent"]')?.value || 0),
    }))()`);
    await setReactInput(client, '[data-custom-symbol-name-input]', "UI registered symbol");
    await setReactSelect(client, '[data-custom-symbol-role-select]', "relation");
    await setReactInput(client, '[data-custom-symbol-omml-fallback-input]', "\\approx");

    await setReactInput(client, '[data-custom-symbol-command-input]', "\\alpha");
    await client.evaluate(`document.querySelector('[data-register-custom-symbol]').click()`);
    await sleep(160);
    const builtinFailure = await client.evaluate(`(() => ({
      status: document.querySelector('[data-custom-symbol-registration-status]')?.getAttribute('data-custom-symbol-registration-status') || "",
      message: document.querySelector('[data-custom-symbol-registration-status]')?.textContent || "",
      storage: localStorage.getItem("visualtex.custom-symbols.v1"),
    }))()`);
    assert.equal(builtinFailure.status, "error");
    assert.match(builtinFailure.message, /alpha/i);
    assert.equal(builtinFailure.storage, baseline.storage);

    await setReactInput(client, '[data-custom-symbol-command-input]', "\\selfdef1");
    await client.evaluate(`document.querySelector('[data-register-custom-symbol]').click()`);
    await sleep(160);
    const numericFailure = await client.evaluate(`(() => ({
      status: document.querySelector('[data-custom-symbol-registration-status]')?.getAttribute('data-custom-symbol-registration-status') || "",
      storage: localStorage.getItem("visualtex.custom-symbols.v1"),
    }))()`);
    assert.equal(numericFailure.status, "error");
    assert.equal(numericFailure.storage, baseline.storage);

    await setReactInput(client, '[data-custom-symbol-command-input]', "\\selfdefa");
    await setReactInput(
      client,
      '[data-custom-symbol-omml-fallback-input]',
      "\\definitelyUnknownVisualTexCommand",
    );
    await client.evaluate(`document.querySelector('[data-register-custom-symbol]').click()`);
    await sleep(160);
    const fallbackFailure = await client.evaluate(`(() => ({
      status: document.querySelector('[data-custom-symbol-registration-status]')?.getAttribute('data-custom-symbol-registration-status') || "",
      storage: localStorage.getItem("visualtex.custom-symbols.v1"),
    }))()`);
    assert.equal(fallbackFailure.status, "error");
    assert.equal(fallbackFailure.storage, baseline.storage);
    process.stdout.write("[custom-symbol-designer] atomic registration failures verified\n");

    await setReactInput(client, '[data-custom-symbol-omml-fallback-input]', "\\approx");
    await client.evaluate(`document.querySelector('[data-register-custom-symbol]').click()`);
    await waitUntil(
      client,
      `document.querySelector('[data-custom-symbol-registration-status]')?.getAttribute('data-custom-symbol-registration-status') === 'success'`,
    );
    const registered = await client.evaluate(`(() => {
      const raw = localStorage.getItem("visualtex.custom-symbols.v1");
      const library = raw ? JSON.parse(raw) : null;
      const symbol = library?.symbols?.[0] || null;
      return {
        symbol,
        currentCanvasMetrics: {
          widthEm: Number(document.querySelector('[data-designer-field="canvas-width"]')?.value || 0),
          ascentEm: Number(document.querySelector('[data-designer-field="canvas-ascent"]')?.value || 0),
          descentEm: Number(document.querySelector('[data-designer-field="canvas-descent"]')?.value || 0),
        },
        dirty: document.querySelector('[data-custom-symbol-registration-panel]')?.getAttribute('data-registration-dirty') || "",
        preview: Boolean(document.querySelector('[data-custom-symbol-registered-preview] .math-preview')),
      };
    })()`);
    assert.equal(registered.symbol.command, "selfdefa");
    assert.equal(registered.symbol.name, "UI registered symbol");
    assert.equal(registered.symbol.role, "relation");
    assert.equal(registered.symbol.ommlFallback, "\\approx");
    assert.ok(registered.symbol.artwork.shapes.length > 0);
    assert.deepEqual(
      registered.currentCanvasMetrics,
      preRegistrationMetrics,
      "Registering a symbol must not shrink the designer work canvas",
    );
    assert.deepEqual(
      registered.symbol.designerSource?.metrics,
      preRegistrationMetrics,
      "Editable source must preserve the original large designer canvas metrics",
    );
    assert.ok(
      registered.symbol.metrics.widthEm < preRegistrationMetrics.widthEm,
      "Runtime registered width must auto-crop to visible artwork instead of using the designer canvas width",
    );
    assert.equal(registered.dirty, "false");
    assert.equal(registered.preview, true);

    const runtimeRegistration = await client.evaluate(`(async () => {
      const search = await import("/src/autocomplete/CommandSearchEngine.ts");
      const runtime = await import("/src/export/runtime.ts");
      const BS = String.fromCharCode(92);
      const command = BS + "selfdefa";
      const field = document.querySelector("math-field");
      field.setValue("A+" + command + "+B", {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      });
      field.dispatchEvent(
        new InputEvent("input", {
          bubbles: true,
          inputType: "insertText",
          data: "a",
        }),
      );
      await new Promise((resolve) => setTimeout(resolve, 90));
      const stored = JSON.parse(localStorage.getItem("visualtex.custom-symbols.v1"));
      const symbol = stored.symbols[0];
      const svg = runtime.latexToSvg(command, {
        displayMode: false,
        fontSizePt: 12,
        paddingPx: 0,
        background: "transparent",
      }).svg;
      const mathMl = runtime.latexToMathMl(command, false);
      return {
        value: field.value,
        renderedClass: Boolean(
          field.shadowRoot?.querySelector(".visualtex-custom-symbol-" + symbol.id),
        ),
        shadowStyle: Array.from(field.shadowRoot?.querySelectorAll("style") || []).some(
          (style) => style.textContent?.includes("visualtex-custom-symbol-" + symbol.id),
        ),
        search: search
          .searchCommands(BS + "selfdefa", {}, false, 10)
          .some((entry) => entry.command === command),
        svg: svg.includes('data-visualtex-custom-symbol="' + symbol.id + '"'),
        fallback: /2248/i.test(mathMl),
      };
    })()`);
    assert.equal(runtimeRegistration.value, "A+\\selfdefa+B");
    assert.equal(runtimeRegistration.renderedClass, true);
    assert.equal(runtimeRegistration.shadowStyle, true);
    assert.equal(runtimeRegistration.search, true);
    assert.equal(runtimeRegistration.svg, true);
    assert.equal(runtimeRegistration.fallback, true);
    process.stdout.write("[custom-symbol-designer] successful runtime registration verified\n");

    await setReactInput(client, '[data-custom-symbol-name-input]', "UI registered symbol updated");
    await setReactInput(client, '[data-designer-field="layer-rotation"]', 18);
    const dirtyState = await client.evaluate(`document.querySelector('[data-custom-symbol-registration-panel]')?.getAttribute('data-registration-dirty')`);
    assert.equal(dirtyState, "true");
    await client.evaluate(`document.querySelector('[data-register-custom-symbol]').click()`);
    await waitUntil(
      client,
      `document.querySelector('[data-custom-symbol-registration-status]')?.getAttribute('data-custom-symbol-registration-status') === 'success'`,
    );
    const updated = await client.evaluate(`(() => {
      const symbol = JSON.parse(localStorage.getItem("visualtex.custom-symbols.v1")).symbols[0];
      return {
        name: symbol.name,
        command: symbol.command,
        rotateDeg: symbol.artwork.shapes[0]?.transform?.rotateDeg ?? null,
        dirty: document.querySelector('[data-custom-symbol-registration-panel]')?.getAttribute('data-registration-dirty') || "",
      };
    })()`);
    assert.equal(updated.name, "UI registered symbol updated");
    assert.equal(updated.command, "selfdefa");
    assert.equal(updated.rotateDeg, 18);
    assert.equal(updated.dirty, "false");

    await setReactInput(client, '[data-custom-symbol-command-input]', "\\alpha");
    await client.evaluate(`document.querySelector('[data-register-custom-symbol]').click()`);
    await sleep(160);
    const failedUpdate = await client.evaluate(`(() => {
      const symbol = JSON.parse(localStorage.getItem("visualtex.custom-symbols.v1")).symbols[0];
      return {
        status: document.querySelector('[data-custom-symbol-registration-status]')?.getAttribute('data-custom-symbol-registration-status') || "",
        storedCommand: symbol.command,
        storedName: symbol.name,
        previewCode: document.querySelector('[data-custom-symbol-registered-preview] code')?.textContent || "",
      };
    })()`);
    assert.equal(failedUpdate.status, "error");
    assert.equal(failedUpdate.storedCommand, "selfdefa");
    assert.equal(failedUpdate.storedName, "UI registered symbol updated");
    assert.equal(failedUpdate.previewCode, "\\selfdefa");
    process.stdout.write("[custom-symbol-designer] atomic update failure verified\n");

    await client.evaluate(`document.querySelector('[data-reset-custom-symbol-designer]').click()`);
    await sleep(100);
    const resetSafety = await client.evaluate(`(() => ({
      layers: document.querySelectorAll('[data-custom-symbol-layer]').length,
      symbolCount: JSON.parse(localStorage.getItem("visualtex.custom-symbols.v1")).symbols.length,
      command: JSON.parse(localStorage.getItem("visualtex.custom-symbols.v1")).symbols[0].command,
    }))()`);
    assert.equal(resetSafety.layers, 0);
    assert.equal(resetSafety.symbolCount, 1);
    assert.equal(resetSafety.command, "selfdefa");

    const archiveState = await client.evaluate(`(() => {
      const symbol = JSON.parse(localStorage.getItem("visualtex.custom-symbols.v1")).symbols[0];
      const item = document.querySelector('[data-registered-custom-symbol-command="selfdefa"]');
      return {
        id: symbol.id,
        archiveVersion: symbol.designerSource?.version ?? null,
        assetCount: symbol.designerSource?.assets?.length ?? -1,
        layerCount: symbol.designerSource?.layers?.length ?? -1,
        sourceLatex: symbol.designerSource?.assets?.[0]?.sourceLatex ?? "",
        listItem: Boolean(item),
      };
    })()`);
    assert.equal(archiveState.archiveVersion, 1);
    assert.equal(archiveState.assetCount, 1);
    assert.equal(archiveState.layerCount, 1);
    assert.equal(archiveState.sourceLatex, "\\partial");
    assert.equal(archiveState.listItem, true);

    await client.evaluate(`document.querySelector('[data-registered-custom-symbol-command="selfdefa"] [data-edit-registered-custom-symbol]').click()`);
    await waitUntil(client, `document.querySelectorAll('[data-custom-symbol-layer]').length === 1`);
    const restoredRegistered = await client.evaluate(`(() => {
      const layer = document.querySelector('[data-custom-symbol-layer]');
      const panel = document.querySelector('[data-custom-symbol-registration-panel]');
      return {
        kind: layer?.getAttribute('data-layer-kind') || "",
        sourceLatex: layer?.getAttribute('data-layer-source-latex') || "",
        symbolId: panel?.getAttribute('data-registration-symbol-id') || "",
        command: document.querySelector('[data-custom-symbol-command-input]')?.value || "",
        legacyWarning: Boolean(document.querySelector('[data-custom-symbol-legacy-warning]')),
      };
    })()`);
    assert.equal(restoredRegistered.kind, "glyph");
    assert.equal(restoredRegistered.sourceLatex, "\\partial");
    assert.equal(restoredRegistered.symbolId, archiveState.id);
    assert.equal(restoredRegistered.command, "selfdefa");
    assert.equal(restoredRegistered.legacyWarning, false);
    process.stdout.write("[custom-symbol-designer] editable source archive restored\n");

    await client.evaluate(`document.querySelector('[data-registered-custom-symbol-command="selfdefa"] [data-duplicate-registered-custom-symbol]').click()`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-command-input]')?.value === "selfdefacopy"`);
    const duplicateDraft = await client.evaluate(`(() => ({
      symbolId: document.querySelector('[data-custom-symbol-registration-panel]')?.getAttribute('data-registration-symbol-id') || "",
      command: document.querySelector('[data-custom-symbol-command-input]')?.value || "",
      layers: document.querySelectorAll('[data-custom-symbol-layer]').length,
      symbolCount: JSON.parse(localStorage.getItem("visualtex.custom-symbols.v1")).symbols.length,
    }))()`);
    assert.equal(duplicateDraft.symbolId, "");
    assert.equal(duplicateDraft.command, "selfdefacopy");
    assert.equal(duplicateDraft.layers, 1);
    assert.equal(duplicateDraft.symbolCount, 1);
    process.stdout.write("[custom-symbol-designer] duplicate-as-draft isolation verified\n");

    await client.evaluate(`document.querySelector('[data-reset-custom-symbol-designer]').click()`);
    await client.evaluate(`document.querySelector('[data-registered-custom-symbol-command="selfdefa"] [data-edit-registered-custom-symbol]').click()`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-registration-panel]')?.getAttribute('data-registration-symbol-id') === ${JSON.stringify(archiveState.id)}`);
    const tileStorageBeforeInsert = await client.evaluate(`localStorage.getItem("visualtex-custom-formula-tiles")`);
    await client.evaluate(`document.querySelector('.custom-symbol-designer-footer button').click()`);
    await waitUntil(client, `!document.querySelector('[data-custom-symbol-designer]')`);
    await waitUntil(client, `Boolean(document.querySelector('[data-registered-custom-symbol-command="selfdefa"]'))`);
    const mainToolbarDeleteVisible = await client.evaluate(`(() => {
      const button = document.querySelector('[data-delete-registered-custom-symbol-toolbar]');
      if (!(button instanceof HTMLElement)) return false;
      const rect = button.getBoundingClientRect();
      const style = getComputedStyle(button);
      return rect.width > 0 && rect.height > 0 && style.display !== "none" && style.visibility !== "hidden";
    })()`);
    assert.equal(mainToolbarDeleteVisible, true, "Registered custom symbols must expose a visible delete button in the main toolbar");

    await client.evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      field.setValue("", {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      });
      field.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "deleteContentBackward", data: null }));
      field.position = field.lastOffset;
    })()`);
    await sleep(100);
    await client.evaluate(`document.querySelector('[data-registered-custom-symbol-command="selfdefa"] .formula-tile-button.is-registered-custom-symbol').click()`);
    await waitUntil(client, `document.querySelector("math-field")?.value === "\\\\selfdefa"`);
    const toolbarInsert = await client.evaluate(`(() => {
      const field = document.querySelector("math-field");
      const symbol = JSON.parse(localStorage.getItem("visualtex.custom-symbols.v1")).symbols[0];
      return {
        value: field.value,
        rendered: Boolean(field.shadowRoot?.querySelector('.visualtex-custom-symbol-' + symbol.id)),
        tileStorage: localStorage.getItem("visualtex-custom-formula-tiles"),
      };
    })()`);
    assert.equal(toolbarInsert.value, "\\selfdefa");
    assert.equal(toolbarInsert.rendered, true);
    assert.equal(toolbarInsert.tileStorage, tileStorageBeforeInsert);
    process.stdout.write("[custom-symbol-designer] registered toolbar insertion verified\n");

    await client.evaluate(`document.querySelector('[data-open-custom-symbol-designer]').click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-custom-symbol-designer]'))`);
    await client.evaluate(`document.querySelector('[data-registered-custom-symbol-command="selfdefa"] [data-edit-registered-custom-symbol]').click()`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-registration-panel]')?.getAttribute('data-registration-symbol-id') === ${JSON.stringify(archiveState.id)}`);

    await client.evaluate(`document.querySelector('[data-registered-custom-symbol-command="selfdefa"] [data-delete-registered-custom-symbol]').click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-custom-symbol-delete-warning]'))`);
    await client.evaluate(`document.querySelector('[data-cancel-delete-registered-custom-symbol]').click()`);
    await waitUntil(client, `!document.querySelector('[data-custom-symbol-delete-warning]')`);
    const afterCancelDelete = await client.evaluate(`JSON.parse(localStorage.getItem("visualtex.custom-symbols.v1")).symbols.length`);
    assert.equal(afterCancelDelete, 1);

    await client.evaluate(`document.querySelector('[data-registered-custom-symbol-command="selfdefa"] [data-delete-registered-custom-symbol]').click()`);
    await waitUntil(client, `Boolean(document.querySelector('[data-confirm-delete-registered-custom-symbol]'))`);
    await client.evaluate(`document.querySelector('[data-confirm-delete-registered-custom-symbol]').click()`);
    await waitUntil(client, `JSON.parse(localStorage.getItem("visualtex.custom-symbols.v1")).symbols.length === 0`);
    await waitUntil(client, `!document.querySelector('[data-registered-custom-symbol-command="selfdefa"]')`);
    await waitUntil(client, `document.querySelector('[data-custom-symbol-registration-status="success"]')?.textContent?.includes('selfdefa')`);
    const deletionState = await client.evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        formula: field.value,
        registeredListCount: document.querySelectorAll('[data-registered-custom-symbol]').length,
        toolbarButton: Boolean(document.querySelector('[data-registered-custom-symbol-command="selfdefa"]')),
        customClass: Boolean(field.shadowRoot?.querySelector('.visualtex-custom-symbol-${archiveState.id}')),
        designerLayers: document.querySelectorAll('[data-custom-symbol-layer]').length,
        registrationId: document.querySelector('[data-custom-symbol-registration-panel]')?.getAttribute('data-registration-symbol-id') || "",
      };
    })()`);
    assert.equal(deletionState.formula.trim(), "\\selfdefa");
    assert.equal(deletionState.registeredListCount, 0);
    assert.equal(deletionState.toolbarButton, false);
    assert.equal(deletionState.customClass, false);
    assert.equal(deletionState.designerLayers, 0);
    assert.equal(deletionState.registrationId, "");
    process.stdout.write("[custom-symbol-designer] two-step deletion and unresolved-source preservation verified\n");

    console.log(
      "Custom symbol designer UI composition, geometry, local slicing, transactional registration/update, editable-source restore, toolbar linkage, deletion safety, fallback, and Proof theme purity regression passed",
    );
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    vite.kill("SIGTERM");
    await sleep(240);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => undefined);
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
