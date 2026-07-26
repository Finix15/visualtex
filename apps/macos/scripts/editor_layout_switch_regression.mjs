import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const offset = process.pid % 700;
const previewPort = 8100 + offset;
const debugPort = 15100 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = `/tmp/visualtex-editor-layout-${process.pid}`;
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
        "--window-size=1440,1000",
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
    if (!page) throw new Error("No VisualTeX page target found");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(500);

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

    const waitForSelector = async (selector, timeoutMs = 6000) => {
      const startedAt = Date.now();
      while (Date.now() - startedAt < timeoutMs) {
        if (await evaluate(`Boolean(document.querySelector(${JSON.stringify(selector)}))`)) {
          return;
        }
        await sleep(50);
      }
      throw new Error(`Timed out waiting for selector: ${selector}`);
    };
    const setViewport = async (width, height = 820) => {
      await client.send("Emulation.setDeviceMetricsOverride", {
        width,
        height,
        deviceScaleFactor: 1,
        mobile: false,
      });
      await sleep(180);
    };

    await evaluate(`(() => {
      localStorage.clear();
      for (const key of [
        "visualtex.onboarding.v3.completed",
        "visualtex.office.macos.first-run.v1.completed",
        "visualtex.onboarding.macos.desktop.v1.2.0.completed",
        "visualtex.office.macos.native-first-run.v1.2.0.completed",
      ]) localStorage.setItem(key, "true");
      localStorage.setItem("visualtex-editor", JSON.stringify({
        state: {
          title: "Layout regression",
          lines: [{ id: "layout-line", latex: "a+b" }],
          activeLineId: "layout-line",
          formulaAlignment: "center",
          editorLayout: "standard",
          sourceOpen: false,
          language: "cn",
          zoom: 1,
        },
        version: 0,
      }));
      localStorage.setItem("visualtex-custom-formula-tiles", JSON.stringify({
        version: 3,
        sections: [{ id: "layout-section", name: "模板", createdAt: 0 }],
        tiles: [
          "\\\\frac{\\\\placeholder{}}{\\\\placeholder{}}",
          "\\\\int \\\\placeholder{}\\\\,\\\\mathrm{d}\\\\placeholder{}",
          "\\\\int_{\\\\placeholder{}}^{\\\\placeholder{}} \\\\placeholder{}\\\\,\\\\mathrm{d}\\\\placeholder{}",
          "\\\\lim_{\\\\placeholder{}\\\\to\\\\placeholder{}}\\\\placeholder{}",
          "\\\\sum_{\\\\placeholder{}}^{\\\\placeholder{}}\\\\placeholder{}",
          "\\\\sqrt{\\\\placeholder{}}+\\\\placeholder{}",
          "\\\\binom{\\\\placeholder{}}{\\\\placeholder{}}",
          "\\\\vec{\\\\placeholder{}}+\\\\placeholder{}",
        ].map((latex, index) => ({
          id: "layout-tile-" + index,
          latex,
          sectionId: "layout-section",
          color: null,
          createdAt: index,
        })),
      }));
    })()`);
    await client.send("Page.reload", { ignoreCache: true });
    await sleep(700);
    await waitForSelector('.workspace[data-editor-layout="standard"]');

    const readLayout = () => evaluate(`(() => {
      const workspace = document.querySelector('.workspace');
      const editor = document.querySelector('.formula-workspace.editor-pane');
      const standardToolbar = document.querySelector('.workspace > .formula-toolbar:not(.classic-tile-toolbar)');
      const classicTiles = document.querySelector('.classic-tile-toolbar');
      const dock = document.querySelector('.classic-bottom-dock');
      const bottomToolbar = document.querySelector('.classic-bottom-toolbar');
      const sourcePanel = document.querySelector('.classic-source-pane-slot .source-panel');
      const editorRect = editor?.getBoundingClientRect();
      const tileRect = classicTiles?.getBoundingClientRect();
      const dockRect = dock?.getBoundingClientRect();
      const editorScrollRect = document.querySelector('.editor-pane-scroll')?.getBoundingClientRect();
      const templateRects = Array.from(bottomToolbar?.querySelectorAll('.template-button') ?? [])
        .slice(0, 12)
        .map((button) => button.getBoundingClientRect());
      const normalTemplateRects = Array.from(
        bottomToolbar?.querySelectorAll('.template-button:not(.is-wide-preview)') ?? [],
      )
        .slice(0, 12)
        .map((button) => button.getBoundingClientRect());
      const bottomTabs = document.querySelector('.classic-bottom-tabs');
      const bottomTabGroup = document.querySelector('.classic-bottom-tab-group');
      const bottomTabsRect = bottomTabs?.getBoundingClientRect();
      const bottomTabGroupRect = bottomTabGroup?.getBoundingClientRect();
      const templateRows = Array.from(
        new Set(templateRects.map((rect) => Math.round(rect.top))),
      );
      const persisted = JSON.parse(localStorage.getItem('visualtex-editor') || '{}').state ?? {};
      return {
        layout: workspace?.dataset.editorLayout ?? '',
        persistedLayout: persisted.editorLayout ?? '',
        hasStandardToolbar: Boolean(standardToolbar),
        standardToolbarFixedView: standardToolbar?.dataset.toolbarFixedView ?? '',
        hasStandardViewTabs: Boolean(standardToolbar?.querySelector('.formula-toolbar-view-tabs')),
        hasClassicTiles: Boolean(classicTiles),
        classicTileView: classicTiles?.dataset.toolbarFixedView ?? '',
        tileRightOfEditor: Boolean(editorRect && tileRect && tileRect.left >= editorRect.right - 1),
        hasDock: Boolean(dock),
        dockBelowEditor: Boolean(dockRect && editorScrollRect && dockRect.top >= editorScrollRect.bottom - 1),
        hasBottomToolbar: Boolean(bottomToolbar),
        bottomToolbarLayout: bottomToolbar?.dataset.toolbarLayout ?? '',
        bottomToolbarView: bottomToolbar?.dataset.toolbarFixedView ?? '',
        templateRows,
        templateNormalMaxWidth: normalTemplateRects.length
          ? Math.max(...normalTemplateRects.map((rect) => rect.width))
          : -1,
        templateMaxHeight: templateRects.length
          ? Math.max(...templateRects.map((rect) => rect.height))
          : -1,
        bottomTabsCentered: Boolean(
          bottomTabsRect &&
          bottomTabGroupRect &&
          Math.abs(
            (bottomTabGroupRect.left + bottomTabGroupRect.right) / 2 -
              (bottomTabsRect.left + bottomTabsRect.right) / 2,
          ) <= 1
        ),
        hasDockCollapse: Boolean(document.querySelector('[data-classic-bottom-collapse]')),
        dockCollapsed: dock?.classList.contains('is-collapsed') ?? false,
        hasSourcePanel: Boolean(sourcePanel),
        hasSourceInternalCollapse: Boolean(
          document.querySelector('.classic-source-pane-slot .source-collapse-button'),
        ),
        alignmentControlsInHeader: Boolean(
          document.querySelector('.editor-pane-header .formula-alignment-controls'),
        ),
      };
    })()`);

    let state = await readLayout();
    assert.equal(state.layout, "standard", JSON.stringify(state));
    assert.equal(state.persistedLayout, "standard", JSON.stringify(state));
    assert.equal(state.hasStandardToolbar, true, JSON.stringify(state));
    assert.equal(state.standardToolbarFixedView, "", JSON.stringify(state));
    assert.equal(state.hasStandardViewTabs, true, JSON.stringify(state));
    assert.equal(state.hasClassicTiles, false, JSON.stringify(state));
    assert.equal(state.hasDock, false, JSON.stringify(state));
    assert.equal(state.alignmentControlsInHeader, true, JSON.stringify(state));

    const settingsClick = await evaluate(`(() => {
      const button = document.querySelector('.settings-toggle');
      button?.click();
      return {
        exists: Boolean(button),
        disabled: button?.disabled ?? null,
        className: button?.className ?? '',
        settingsDialogsBeforeFrame: document.querySelectorAll('.settings-dialog').length,
      };
    })()`);
    await sleep(150);
    const settingsState = await evaluate(`({
      dialogs: document.querySelectorAll('.settings-dialog').length,
      backdrops: document.querySelectorAll('.modal-backdrop').length,
      choices: document.querySelectorAll('[data-editor-layout-choice]').length,
      bodyText: document.body.innerText.slice(0, 300),
    })`);
    if (!settingsState.choices) {
      throw new Error(`Settings did not open: ${JSON.stringify({ settingsClick, settingsState })}`);
    }
    await waitForSelector('[data-editor-layout-choice="classic"]');
    await evaluate(`document.querySelector('[data-editor-layout-choice="classic"]')?.click()`);
    await sleep(180);

    state = await readLayout();
    assert.equal(state.layout, "classic", JSON.stringify(state));
    assert.equal(state.persistedLayout, "classic", JSON.stringify(state));
    assert.equal(state.hasStandardToolbar, false, JSON.stringify(state));
    assert.equal(state.hasClassicTiles, true, JSON.stringify(state));
    assert.equal(state.classicTileView, "tiles", JSON.stringify(state));
    assert.equal(state.tileRightOfEditor, true, JSON.stringify(state));
    assert.equal(state.hasDock, true, JSON.stringify(state));
    assert.equal(state.dockBelowEditor, true, JSON.stringify(state));
    assert.equal(state.hasBottomToolbar, true, JSON.stringify(state));
    assert.equal(state.bottomToolbarLayout, "horizontal", JSON.stringify(state));
    assert.equal(state.bottomToolbarView, "tools", JSON.stringify(state));
    assert.equal(state.templateRows.length, 3, JSON.stringify(state));
    assert.ok(state.templateNormalMaxWidth <= 60, JSON.stringify(state));
    assert.ok(state.templateMaxHeight <= 58, JSON.stringify(state));
    assert.equal(state.bottomTabsCentered, true, JSON.stringify(state));
    assert.equal(state.hasDockCollapse, true, JSON.stringify(state));
    assert.equal(state.dockCollapsed, false, JSON.stringify(state));
    assert.equal(state.alignmentControlsInHeader, true, JSON.stringify(state));

    for (const category of [
      "structure",
      "greek",
      "arrow",
      "structure",
      "greek",
      "arrow",
      "structure",
      "common",
    ]) {
      await evaluate(`document.querySelector('.classic-bottom-toolbar [data-category="${category}"]')?.click()`);
      await sleep(80);
      const categoryState = await evaluate(`(() => {
        const toolbar = document.querySelector('.classic-bottom-toolbar');
        const strips = Array.from(toolbar?.querySelectorAll('.template-strip') ?? []);
        const strip = strips[0];
        const buttons = Array.from(strip?.querySelectorAll('.template-button') ?? []);
        const ids = buttons.map((button) => button.dataset.commandId ?? '');
        return {
          category: strip?.dataset.activeCategory ?? '',
          stripCount: strips.length,
          buttonCount: buttons.length,
          previewCount: strip?.querySelectorAll('.template-button .math-preview').length ?? 0,
          ids,
          uniqueIds: new Set(ids).size,
        };
      })()`);
      assert.equal(categoryState.category, category, JSON.stringify(categoryState));
      assert.equal(categoryState.stripCount, 1, JSON.stringify(categoryState));
      assert.ok(categoryState.buttonCount > 0, JSON.stringify(categoryState));
      assert.equal(categoryState.previewCount, categoryState.buttonCount, JSON.stringify(categoryState));
      assert.equal(categoryState.uniqueIds, categoryState.buttonCount, JSON.stringify(categoryState));
    }

    await evaluate(`document.querySelector('.classic-bottom-toolbar [data-category="matrix"]')?.click()`);
    await sleep(160);
    const matrixState = await evaluate(`(() => {
      const builder = document.querySelector('.classic-bottom-toolbar .matrix-builder');
      const builderRect = builder?.getBoundingClientRect();
      const optionsColumn = builder?.querySelector('.matrix-options-column');
      const optionsColumnRect = optionsColumn?.getBoundingClientRect();
      const heading = builder?.querySelector('.matrix-builder-heading');
      const headingRect = heading?.getBoundingClientRect();
      const headingStrong = heading?.querySelector('strong');
      const badge = builder?.querySelector('.matrix-size-badge');
      const grid = builder?.querySelector('.matrix-size-grid');
      const sizePicker = builder?.querySelector('.matrix-size-picker');
      const sizePickerRect = sizePicker?.getBoundingClientRect();
      const insert = builder?.querySelector('.matrix-insert-button');
      const nextTemplate = document.querySelector(
        '.classic-bottom-toolbar .template-strip > .template-button',
      );
      const nextTemplateRect = nextTemplate?.getBoundingClientRect();
      const delimiterButtons = Array.from(builder?.querySelectorAll('.matrix-delimiter-options button') ?? []);
      const inside = (rect) => Boolean(
        builderRect && rect &&
        rect.left >= builderRect.left - 1 &&
        rect.right <= builderRect.right + 1 &&
        rect.top >= builderRect.top - 1 &&
        rect.bottom <= builderRect.bottom + 1
      );
      return {
        exists: Boolean(builder),
        builderRect: builderRect ? {
          width: builderRect.width,
          height: builderRect.height,
        } : null,
        optionsColumnInside: inside(optionsColumnRect),
        sizePickerInside: inside(sizePickerRect),
        hasTwoColumnOrder: Boolean(
          optionsColumnRect && sizePickerRect &&
          optionsColumnRect.right <= sizePickerRect.left
        ),
        headingInside: inside(headingRect),
        headingWhiteSpace: headingStrong ? getComputedStyle(headingStrong).whiteSpace : '',
        headingLineHeight: headingStrong ? parseFloat(getComputedStyle(headingStrong).lineHeight) : -1,
        headingHeight: headingStrong?.getBoundingClientRect().height ?? -1,
        badgeInside: inside(badge?.getBoundingClientRect()),
        gridInside: inside(grid?.getBoundingClientRect()),
        insertInside: inside(insert?.getBoundingClientRect()),
        insertWidth: insert?.getBoundingClientRect().width ?? -1,
        insertHeight: insert?.getBoundingClientRect().height ?? -1,
        nextTemplateGap: builderRect && nextTemplateRect
          ? nextTemplateRect.left - builderRect.right
          : -1,
        delimiterCount: delimiterButtons.length,
        delimiterHeights: delimiterButtons.map((button) =>
          button.getBoundingClientRect().height,
        ),
        delimiterInside: delimiterButtons.every((button) => inside(button.getBoundingClientRect())),
        delimiterPreviewInside: delimiterButtons.every((button) => {
          const buttonRect = button.getBoundingClientRect();
          const previewRect = button.querySelector('.math-preview-fit-content')?.getBoundingClientRect();
          return Boolean(
            previewRect &&
            previewRect.left >= buttonRect.left - 1 &&
            previewRect.right <= buttonRect.right + 1 &&
            previewRect.top >= buttonRect.top - 1 &&
            previewRect.bottom <= buttonRect.bottom + 1
          );
        }),
        delimiterScales: delimiterButtons.map((button) =>
          Number(button.querySelector('.math-preview')?.dataset.fitScale ?? '0'),
        ),
      };
    })()`);
    assert.equal(matrixState.exists, true, JSON.stringify(matrixState));
    assert.ok(matrixState.builderRect.width >= 340, JSON.stringify(matrixState));
    assert.ok(matrixState.builderRect.width <= 390, JSON.stringify(matrixState));
    assert.ok(matrixState.builderRect.height >= 130, JSON.stringify(matrixState));
    assert.equal(matrixState.optionsColumnInside, true, JSON.stringify(matrixState));
    assert.equal(matrixState.sizePickerInside, true, JSON.stringify(matrixState));
    assert.equal(matrixState.hasTwoColumnOrder, true, JSON.stringify(matrixState));
    assert.equal(matrixState.headingInside, true, JSON.stringify(matrixState));
    assert.equal(matrixState.headingWhiteSpace, "nowrap", JSON.stringify(matrixState));
    assert.ok(matrixState.headingHeight <= matrixState.headingLineHeight + 1, JSON.stringify(matrixState));
    assert.equal(matrixState.badgeInside, true, JSON.stringify(matrixState));
    assert.equal(matrixState.gridInside, true, JSON.stringify(matrixState));
    assert.equal(matrixState.insertInside, true, JSON.stringify(matrixState));
    assert.ok(matrixState.insertWidth >= 180, JSON.stringify(matrixState));
    assert.ok(matrixState.insertWidth <= 220, JSON.stringify(matrixState));
    assert.ok(matrixState.insertHeight <= 38, JSON.stringify(matrixState));
    assert.ok(matrixState.nextTemplateGap >= 0, JSON.stringify(matrixState));
    assert.ok(matrixState.nextTemplateGap <= 8, JSON.stringify(matrixState));
    assert.equal(matrixState.delimiterCount, 3, JSON.stringify(matrixState));
    assert.ok(
      matrixState.delimiterHeights.every((height) => height >= 58),
      JSON.stringify(matrixState),
    );
    assert.equal(matrixState.delimiterInside, true, JSON.stringify(matrixState));
    assert.equal(matrixState.delimiterPreviewInside, true, JSON.stringify(matrixState));
    assert.ok(
      matrixState.delimiterScales.every((scale) => scale > 0 && scale <= 1.151),
      JSON.stringify(matrixState),
    );

    await evaluate(`document.querySelector('.classic-bottom-toolbar [data-category="common"]')?.click()`);
    await sleep(80);

    const readResponsiveLayout = () => evaluate(`(() => {
      const workspace = document.querySelector('.workspace.is-classic-layout');
      const editor = document.querySelector('.formula-workspace.editor-pane');
      const editorHeader = document.querySelector('.editor-pane-header');
      const dock = document.querySelector('.classic-bottom-dock');
      const bottomToolbar = document.querySelector('.classic-bottom-toolbar');
      const tiles = document.querySelector('.classic-tile-toolbar');
      const workspaceRect = workspace?.getBoundingClientRect();
      const editorRect = editor?.getBoundingClientRect();
      const headerRect = editorHeader?.getBoundingClientRect();
      const dockRect = dock?.getBoundingClientRect();
      const bottomRect = bottomToolbar?.getBoundingClientRect();
      const tileRect = tiles?.getBoundingClientRect();
      const tileStyle = tiles ? getComputedStyle(tiles) : null;
      const bottomStyle = bottomToolbar ? getComputedStyle(bottomToolbar) : null;
      const contained = (outer, inner) => Boolean(
        outer && inner &&
        inner.left >= outer.left - 1 &&
        inner.right <= outer.right + 1 &&
        inner.top >= outer.top - 1 &&
        inner.bottom <= outer.bottom + 1
      );
      return {
        viewportWidth: window.innerWidth,
        workspaceWidth: workspaceRect?.width ?? -1,
        workspaceClientWidth: workspace?.clientWidth ?? -1,
        workspaceScrollWidth: workspace?.scrollWidth ?? -1,
        editorWidth: editorRect?.width ?? -1,
        editorVisible: Boolean(
          editorRect && headerRect &&
          editorRect.width > 180 && editorRect.height > 180 &&
          headerRect.height > 20
        ),
        dockInsideEditor: contained(editorRect, dockRect),
        bottomInsideEditor: contained(editorRect, bottomRect),
        bottomPosition: bottomStyle?.position ?? '',
        tilePosition: tileStyle?.position ?? '',
        tileWidth: tileRect?.width ?? -1,
        tileRightAligned: Boolean(
          workspaceRect && tileRect &&
          Math.abs(tileRect.right - workspaceRect.right) <= 1
        ),
        tileOverlaysEditor: Boolean(
          editorRect && tileRect && tileRect.left < editorRect.right - 1
        ),
      };
    })()`);

    for (const width of [1020, 820]) {
      await setViewport(width);
      const hasTiles = await evaluate(
        `Boolean(document.querySelector('.classic-tile-toolbar'))`,
      );
      if (!hasTiles) {
        await evaluate(`document.querySelector('.sidebar-toggle')?.click()`);
        await waitForSelector('.classic-tile-toolbar');
        await sleep(100);
      }
      const responsive = await readResponsiveLayout();
      assert.equal(responsive.viewportWidth, width, JSON.stringify(responsive));
      assert.equal(responsive.editorVisible, true, JSON.stringify(responsive));
      assert.equal(responsive.dockInsideEditor, true, JSON.stringify(responsive));
      assert.equal(responsive.bottomInsideEditor, true, JSON.stringify(responsive));
      assert.notEqual(responsive.bottomPosition, "absolute", JSON.stringify(responsive));
      assert.notEqual(responsive.bottomPosition, "fixed", JSON.stringify(responsive));
      assert.notEqual(responsive.tilePosition, "absolute", JSON.stringify(responsive));
      assert.notEqual(responsive.tilePosition, "fixed", JSON.stringify(responsive));
      assert.equal(responsive.tileRightAligned, true, JSON.stringify(responsive));
      assert.equal(responsive.tileOverlaysEditor, false, JSON.stringify(responsive));
      assert.ok(responsive.editorWidth >= 420, JSON.stringify(responsive));
      assert.ok(
        responsive.workspaceScrollWidth <= responsive.workspaceClientWidth + 1,
        JSON.stringify(responsive),
      );
    }

    await setViewport(680);
    const narrowResponsive = await readResponsiveLayout();
    assert.equal(narrowResponsive.viewportWidth, 680, JSON.stringify(narrowResponsive));
    assert.equal(narrowResponsive.editorVisible, true, JSON.stringify(narrowResponsive));
    assert.equal(narrowResponsive.dockInsideEditor, true, JSON.stringify(narrowResponsive));
    assert.equal(narrowResponsive.bottomInsideEditor, true, JSON.stringify(narrowResponsive));
    assert.notEqual(narrowResponsive.bottomPosition, "absolute", JSON.stringify(narrowResponsive));
    assert.notEqual(narrowResponsive.bottomPosition, "fixed", JSON.stringify(narrowResponsive));
    assert.equal(narrowResponsive.tilePosition, "absolute", JSON.stringify(narrowResponsive));
    assert.equal(narrowResponsive.tileRightAligned, true, JSON.stringify(narrowResponsive));
    assert.equal(narrowResponsive.tileOverlaysEditor, true, JSON.stringify(narrowResponsive));
    assert.ok(narrowResponsive.editorWidth >= 600, JSON.stringify(narrowResponsive));
    assert.ok(narrowResponsive.tileWidth <= 280, JSON.stringify(narrowResponsive));
    assert.ok(
      narrowResponsive.workspaceScrollWidth <= narrowResponsive.workspaceClientWidth + 1,
      JSON.stringify(narrowResponsive),
    );

    await setViewport(1440, 1000);

    await evaluate(`document.querySelector('[data-classic-bottom-view="source"]')?.click()`);
    await sleep(120);
    state = await readLayout();
    assert.equal(state.hasBottomToolbar, false, JSON.stringify(state));
    assert.equal(state.hasSourcePanel, true, JSON.stringify(state));
    assert.equal(state.hasSourceInternalCollapse, false, JSON.stringify(state));
    assert.equal(state.bottomTabsCentered, true, JSON.stringify(state));

    await evaluate(`document.querySelector('[data-classic-bottom-collapse]')?.click()`);
    await sleep(100);
    state = await readLayout();
    assert.equal(state.dockCollapsed, true, JSON.stringify(state));
    assert.equal(state.hasBottomToolbar, false, JSON.stringify(state));
    assert.equal(state.hasSourcePanel, false, JSON.stringify(state));
    assert.equal(state.hasDockCollapse, true, JSON.stringify(state));

    await evaluate(`document.querySelector('[data-classic-bottom-collapse]')?.click()`);
    await sleep(100);
    state = await readLayout();
    assert.equal(state.dockCollapsed, false, JSON.stringify(state));
    assert.equal(state.hasSourcePanel, true, JSON.stringify(state));

    await evaluate(`document.querySelector('[data-classic-bottom-view="tools"]')?.click()`);
    await sleep(120);
    state = await readLayout();
    assert.equal(state.hasBottomToolbar, true, JSON.stringify(state));
    assert.equal(state.hasSourcePanel, false, JSON.stringify(state));
    assert.equal(state.templateRows.length, 3, JSON.stringify(state));

    await evaluate(`document.querySelector('.classic-tile-toolbar [data-tile-category="custom"]')?.click()`);
    await sleep(350);
    await waitForSelector('.classic-tile-toolbar .formula-tile-button.is-custom');
    const classicTileFit = await evaluate(`Array.from(
      document.querySelectorAll('.classic-tile-toolbar .formula-tile-button.is-custom'),
    ).slice(0, 8).map((button) => {
      const preview = button.querySelector('.formula-tile-preview');
      const content = preview?.querySelector('.math-preview-fit-content');
      const previewRect = preview?.getBoundingClientRect();
      const contentRect = content?.getBoundingClientRect();
      return {
        id: button.dataset.formulaTileId ?? '',
        scale: Number(preview?.dataset.fitScale ?? '0'),
        fits: Boolean(
          previewRect && contentRect &&
          contentRect.left >= previewRect.left - 1 &&
          contentRect.right <= previewRect.right + 1 &&
          contentRect.top >= previewRect.top - 1 &&
          contentRect.bottom <= previewRect.bottom + 1
        ),
      };
    })`);
    assert.ok(classicTileFit.length > 0, JSON.stringify(classicTileFit));
    assert.ok(
      classicTileFit.every((item) =>
        item.scale >= 0.799 && item.scale <= 1.201 && item.fits
      ),
      JSON.stringify(classicTileFit),
    );
    const tileStability = await evaluate(`new Promise((resolve) => {
      const samples = [];
      let remaining = 36;
      const sample = () => {
        const panel = document.querySelector('.classic-tile-toolbar .formula-tiles-panel');
        const list = document.querySelector('.classic-tile-toolbar .formula-tile-list');
        const buttons = Array.from(
          document.querySelectorAll('.classic-tile-toolbar .formula-tile-button.is-custom'),
        ).slice(0, 8);
        samples.push({
          panelWidth: panel?.getBoundingClientRect().width ?? -1,
          listWidth: list?.getBoundingClientRect().width ?? -1,
          listHeight: list?.getBoundingClientRect().height ?? -1,
          scrollHeight: list?.scrollHeight ?? -1,
          buttons: buttons.map((button) => {
            const rect = button.getBoundingClientRect();
            return [rect.left, rect.top, rect.width, rect.height];
          }),
        });
        remaining -= 1;
        if (remaining > 0) requestAnimationFrame(sample);
        else resolve(samples);
      };
      requestAnimationFrame(sample);
    })`);
    const stabilitySignature = (sample) => JSON.stringify({
      panelWidth: Math.round(sample.panelWidth * 2) / 2,
      listWidth: Math.round(sample.listWidth * 2) / 2,
      listHeight: Math.round(sample.listHeight * 2) / 2,
      scrollHeight: sample.scrollHeight,
      buttons: sample.buttons.map((rect) => rect.map((value) => Math.round(value * 2) / 2)),
    });
    assert.equal(
      new Set(tileStability.map(stabilitySignature)).size,
      1,
      JSON.stringify(tileStability),
    );

    await evaluate(`document.querySelector('[data-editor-layout-choice="standard"]')?.click()`);
    await sleep(180);
    state = await readLayout();
    assert.equal(state.layout, "standard", JSON.stringify(state));
    assert.equal(state.persistedLayout, "standard", JSON.stringify(state));
    assert.equal(state.hasStandardToolbar, true, JSON.stringify(state));
    assert.equal(state.hasClassicTiles, false, JSON.stringify(state));
    assert.equal(state.hasDock, false, JSON.stringify(state));

    console.log("Editor layout switch regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(250);
    await rm(chromeProfile, { recursive: true, force: true });
  }
}

await main();
