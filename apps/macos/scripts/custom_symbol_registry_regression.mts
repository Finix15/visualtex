import assert from "node:assert/strict";
import { searchCommands } from "../src/autocomplete/CommandSearchEngine.ts";
import { parseLatexSourceDraft } from "../src/clipboard/LatexCopyService.ts";
import {
  deleteCustomSymbol,
  readCustomSymbolLibrary,
  registerCustomSymbol,
  replaceCustomSymbolLibrary,
  updateCustomSymbol,
} from "../src/math/customSymbolRegistry.ts";
import { latexToMathMl, latexToSvg } from "../src/export/runtime.ts";

const now = Date.now();
const baseDefinition = {
  id: "regression-selfdefa",
  command: "selfdefa",
  name: "Regression custom symbol A",
  role: "relation" as const,
  limitsBehavior: "auto" as const,
  metrics: { widthEm: 0.8, ascentEm: 0.62, descentEm: 0.12 },
  artwork: {
    shapes: [
      {
        kind: "circle" as const,
        cx: 400,
        cy: 370,
        r: 260,
        fill: false,
        strokeWidth: 72,
      },
      {
        kind: "line" as const,
        x1: 120,
        y1: 370,
        x2: 680,
        y2: 370,
        fill: false,
        strokeWidth: 72,
        lineCap: "round" as const,
      },
    ],
  },
  ommlFallback: "\\approx",
  createdAt: now,
  updatedAt: now,
};

function withCommand(command: string) {
  return {
    ...baseDefinition,
    id: `regression-${command}`,
    command,
  };
}

replaceCustomSymbolLibrary({ version: 1, symbols: [] });

try {
  assert.throws(() => registerCustomSymbol(withCommand("alpha")), /reserved/);
  assert.throws(() => registerCustomSymbol(withCommand("frac")), /reserved/);
  assert.throws(() => registerCustomSymbol(withCommand("color")), /reserved/);
  assert.throws(() => registerCustomSymbol(withCommand("selfdef1")), /letters only/);

  registerCustomSymbol(baseDefinition);
  let library = readCustomSymbolLibrary();
  assert.equal(library.symbols.length, 1);
  assert.equal(library.symbols[0]?.command, "selfdefa");
  assert.equal(library.symbols[0]?.role, "relation");

  const registeredSourceDraft = parseLatexSourceDraft("\\selfdefa", "raw");
  assert.equal(registeredSourceDraft.valid, true);
  const unknownSourceDraft = parseLatexSourceDraft("\\definitelyunknown", "raw");
  assert.equal(unknownSourceDraft.valid, false);
  assert.equal(unknownSourceDraft.error, "unknown-command");

  const suggestions = searchCommands("selfdefa", {}, false, 8);
  assert.equal(suggestions[0]?.command, "\\selfdefa");
  assert.equal(suggestions[0]?.id, "custom-symbol:regression-selfdefa");

  const svg = latexToSvg("A\\selfdefa B", {
    displayMode: false,
    fontSizePt: 12,
    paddingPx: 0,
    background: "transparent",
  });
  assert.match(svg.svg, /data-visualtex-custom-symbol="regression-selfdefa"/);
  assert.match(svg.svg, /<circle\b[^>]*cx="400"[^>]*cy="370"/);
  assert.match(svg.svg, /<line\b[^>]*x1="120"[^>]*x2="680"/);
  assert.doesNotMatch(svg.svg, /\\selfdefa/);

  registerCustomSymbol({
    ...baseDefinition,
    id: "regression-selferase",
    command: "selferase",
    name: "Regression erased custom symbol",
    role: "ordinary",
    artwork: {
      shapes: [
        ...baseDefinition.artwork.shapes,
        {
          kind: "path" as const,
          operation: "erase" as const,
          d: "M180 370L620 370",
          fill: false,
          strokeWidth: 120,
          lineCap: "round" as const,
        },
      ],
    },
    designerSource: {
      version: 1 as const,
      assets: [],
      layers: [
        {
          id: "eraser-layer-source",
          name: "Eraser stroke",
          kind: "geometry" as const,
          geometryPreset: "eraser" as const,
          visible: true,
          locked: false,
          transform: {},
          shape: {
            kind: "path" as const,
            operation: "erase" as const,
            d: "M180 370L620 370",
            fill: false,
            strokeWidth: 120,
            lineCap: "round" as const,
          },
          bounds: { x: 120, y: 310, width: 560, height: 120 },
        },
      ],
    },
    ommlFallback: null,
  });
  const erasedSvg = latexToSvg("\\selferase", {
    displayMode: false,
    fontSizePt: 12,
    paddingPx: 0,
    background: "transparent",
  });
  assert.match(erasedSvg.svg, /data-visualtex-custom-symbol="regression-selferase"/);
  assert.match(erasedSvg.svg, /<mask\b[^>]*id="visualtex-custom-symbol-erase-regression-selferase-/);
  assert.match(
    erasedSvg.svg,
    /<rect\b[^>]*fill="inherit"[^>]*mask="url\(#visualtex-custom-symbol-erase-regression-selferase-/,
    "Erased exported artwork must use a transparent SVG mask rather than a background-colored cover",
  );
  assert.match(erasedSvg.svg, /stroke="black"[^>]*stroke-width="120"/);
  const storedEraser = readCustomSymbolLibrary().symbols.find(
    (symbol) => symbol.id === "regression-selferase",
  );
  const storedEraserLayer = storedEraser?.designerSource?.layers[0];
  assert.equal(storedEraserLayer?.kind, "geometry");
  assert.equal(
    storedEraserLayer?.kind === "geometry" ? storedEraserLayer.geometryPreset : null,
    "eraser",
  );
  assert.equal(
    storedEraserLayer?.kind === "geometry" ? storedEraserLayer.shape.operation : null,
    "erase",
  );

  const mathMl = latexToMathMl("A\\selfdefa B", false);
  assert.match(mathMl, /<mo>&#x2248;<\/mo>/, "OMML/MathML fallback must be semantic, not vector data");

  registerCustomSymbol({
    ...baseDefinition,
    id: "regression-selfdefb",
    command: "selfdefb",
    name: "Regression custom symbol B",
    role: "ordinary",
    ommlFallback: null,
  });
  assert.throws(
    () => latexToMathMl("\\selfdefb", false),
    /did not resolve LaTeX command \\selfdefb/,
  );

  updateCustomSymbol("regression-selfdefa", {
    name: "Updated custom symbol A",
    metrics: { widthEm: 0.9, ascentEm: 0.66, descentEm: 0.1 },
  });
  library = readCustomSymbolLibrary();
  const updated = library.symbols.find((symbol) => symbol.id === "regression-selfdefa");
  assert.equal(updated?.name, "Updated custom symbol A");
  assert.equal(updated?.metrics.widthEm, 0.9);

  deleteCustomSymbol("regression-selfdefa");
  const deletedSourceDraft = parseLatexSourceDraft("\\selfdefa", "raw");
  assert.equal(deletedSourceDraft.valid, false);
  assert.equal(deletedSourceDraft.error, "unknown-command");
  assert.equal(
    searchCommands("selfdefa", {}, false, 8).some(
      (command) => command.command === "\\selfdefa",
    ),
    false,
  );
  assert.throws(
    () => latexToSvg("\\selfdefa", {
      displayMode: false,
      fontSizePt: 12,
      paddingPx: 0,
      background: "transparent",
    }),
    /MathJax could not parse|Undefined control sequence|Unknown symbol|unresolved LaTeX command/i,
  );

  replaceCustomSymbolLibrary({ version: 1, symbols: [] });
  const roleCases = [
    ["ordinary", "selfordinaryrole", "auto"],
    ["binary", "selfbinaryrole", "auto"],
    ["relation", "selfrelationrole", "auto"],
    ["operator", "selfoperatorrole", "limits"],
    ["open", "selfopenrole", "auto"],
    ["close", "selfcloserole", "auto"],
    ["punctuation", "selfpunctuationrole", "auto"],
  ] as const;
  for (const [role, command, limitsBehavior] of roleCases) {
    registerCustomSymbol({
      ...baseDefinition,
      id: `regression-${command}`,
      command,
      name: `Regression ${role} role`,
      role,
      limitsBehavior,
      ommlFallback: null,
    });
  }
  const roleSvgs = Object.fromEntries(
    roleCases.map(([role, command]) => {
      const latex = role === "operator"
        ? `A\\${command}_{i}^{j}B`
        : `A\\${command} B`;
      const output = latexToSvg(latex, {
        displayMode: role === "operator",
        fontSizePt: 12,
        paddingPx: 0,
        background: "transparent",
      });
      assert.match(
        output.svg,
        new RegExp(`data-visualtex-custom-symbol="regression-${command}"`),
        `${role} custom symbol SVG marker`,
      );
      return [role, output];
    }),
  );
  assert.ok(
    roleSvgs.binary.width > roleSvgs.ordinary.width,
    "Binary custom symbols must receive binary-operator spacing",
  );
  assert.ok(
    roleSvgs.relation.width > roleSvgs.ordinary.width,
    "Relation custom symbols must receive relation spacing",
  );

  console.log(
    "Custom symbol registry validation, persistence, runtime command search, SVG roles, fallback, update, and delete regression passed",
  );
} finally {
  replaceCustomSymbolLibrary({ version: 1, symbols: [] });
}
