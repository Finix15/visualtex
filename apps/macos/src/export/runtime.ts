import { mathjax } from "mathjax-full/js/mathjax.js";
import { TeX } from "mathjax-full/js/input/tex.js";
import { SVG } from "mathjax-full/js/output/svg.js";
import { liteAdaptor } from "mathjax-full/js/adaptors/liteAdaptor.js";
import { RegisterHTMLHandler } from "mathjax-full/js/handlers/html.js";
import { AllPackages } from "mathjax-full/js/input/tex/AllPackages.js";
import { normalizeMathLiveCanonicalUprightCommands } from "../editor/normalizeChineseLatex.ts";
import {
  normalizeMathJaxUnsupportedNaryCommands,
  registerMathJaxIntegralGlyphs,
} from "./mathJaxCompatibility.ts";
import type {
  PngExportOptions,
  PngExportResult,
  SvgExportOptions,
  SvgExportResult,
} from "./exportTypes";
import { errorMessage } from "../runtime/errorMessage";

const DEFAULT_OPTIONS: SvgExportOptions = {
  displayMode: true,
  fontSizePt: 12,
  paddingPx: 8,
  background: "transparent",
};

const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);
const texInput = new TeX({
  packages: AllPackages,
  formatError: (_jax: unknown, error: unknown) => {
    throw new Error(errorMessage(error, "MathJax could not parse this formula."), {
      cause: error,
    });
  },
});
const svgOutput = new SVG({
  fontCache: "local",
  internalSpeechTitles: false,
});
registerMathJaxIntegralGlyphs(svgOutput.font);
const mathDocument = mathjax.document("", {
  InputJax: texInput,
  OutputJax: svgOutput,
});

function positiveFinite(value: number, fallback: number) {
  return Number.isFinite(value) && value > 0 ? value : fallback;
}

function nonNegativeFinite(value: number, fallback: number) {
  return Number.isFinite(value) && value >= 0 ? value : fallback;
}

function isSingleCompleteEnvironment(source: string) {
  const first = source.match(/^\\begin\s*\{([^{}]+)\}/);
  if (!first) return false;

  const environmentToken = /\\(begin|end)\s*\{([^{}]+)\}/g;
  const stack: string[] = [];
  let match: RegExpExecArray | null;
  let outerEnd = -1;

  while ((match = environmentToken.exec(source))) {
    const [, kind, name] = match;
    if (kind === "begin") {
      stack.push(name);
      continue;
    }
    if (stack.at(-1) !== name) return false;
    stack.pop();
    if (stack.length === 0) {
      outerEnd = environmentToken.lastIndex;
      break;
    }
  }

  return outerEnd >= 0 && source.slice(outerEnd).trim().length === 0;
}

function prepareLatex(latex: string) {
  const normalized = normalizeMathJaxUnsupportedNaryCommands(
    normalizeMathLiveCanonicalUprightCommands(
      latex.replace(/\r\n?/g, "\n"),
    ),
  ).trim();
  if (!normalized) throw new Error("Cannot export an empty formula.");

  const lines = normalized
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
  if (lines.length <= 1) return normalized;

  // Preserve a source string that is itself one complete TeX environment.
  // A document with multiple VisualTeX formula rows may still contain an
  // inner matrix/cases environment on one row; that must not make all rows
  // collapse into a single horizontal TeX expression.
  if (isSingleCompleteEnvironment(normalized)) return normalized;

  // `aligned` uses a right/left pair around every alignment marker. Without
  // an explicit marker MathJax right-aligns rows of different widths. Keep
  // the whole document as one image, but anchor every formula row on its left
  // edge and preserve the editor's vertical ordering.
  return `\\begin{aligned}${lines.map((line) => `&${line}`).join("\\\\")}\\end{aligned}`;
}

function extractSvg(markup: string) {
  const start = markup.indexOf("<svg");
  const end = markup.lastIndexOf("</svg>");
  if (start < 0 || end < start) {
    throw new Error("MathJax did not produce an SVG element.");
  }
  return markup.slice(start, end + "</svg>".length);
}

function parseViewBox(svg: string) {
  const match = svg.match(
    /\bviewBox=["']\s*([-+\d.eE]+)\s+([-+\d.eE]+)\s+([-+\d.eE]+)\s+([-+\d.eE]+)\s*["']/,
  );
  if (!match) throw new Error("Exported SVG is missing a valid viewBox.");
  const values = match.slice(1).map(Number);
  if (values.some((value) => !Number.isFinite(value))) {
    throw new Error("Exported SVG has an invalid viewBox.");
  }
  const [x, y, width, height] = values;
  if (width <= 0 || height <= 0) {
    throw new Error("Exported SVG has non-positive dimensions.");
  }
  return { x, y, width, height };
}

function assertSelfContained(svg: string) {
  if (/<foreignObject\b/i.test(svg)) {
    throw new Error("SVG export must not contain foreignObject.");
  }
  if (/<link\b|@import\b/i.test(svg)) {
    throw new Error("SVG export must not depend on external CSS.");
  }
  if (/\b(?:href|xlink:href)=["'](?!#|data:)[^"']+/i.test(svg)) {
    throw new Error("SVG export contains an external resource reference.");
  }
  if (/url\(\s*["']?https?:/i.test(svg)) {
    throw new Error("SVG export contains a remote URL.");
  }
}

const WORD_EXPLICIT_BLACK = "#000000";

function wordCompatiblePaintValue(value: string) {
  const trimmed = value.trim();
  if (/^(?:none|transparent)$/i.test(trimmed)) {
    return trimmed.toLowerCase();
  }
  return WORD_EXPLICIT_BLACK;
}

function removeCssCustomProperties(value: string) {
  return value.replace(
    /(^|[;{])\s*--[a-zA-Z0-9_-]+\s*:[^;}]*;?/g,
    "$1",
  );
}

function forceStylePaintBlack(value: string) {
  return removeCssCustomProperties(value).replace(
    /(^|[;{]\s*)(color|fill|stroke)\s*:\s*([^;}]+)/gi,
    (_match, prefix: string, property: string, paint: string) =>
      `${prefix}${property}:${wordCompatiblePaintValue(paint)}`,
  );
}

/**
 * Word 16.89 can initially paint an SVG formula as transparent when its first
 * resolved colour comes from currentColor, a CSS variable, a white inherited
 * paint, or another deferred style carrier. Normalize every SVG paint carrier
 * before either the SVG or PNG is emitted so both compatibility representations
 * are byte-for-byte derived from the same explicit-black artwork.
 */
function forceWordCompatibleBlack(svg: string) {
  let output = svg.replace(/currentColor/gi, WORD_EXPLICIT_BLACK);
  output = output.replace(
    /\b(color|fill|stroke)=(['"])(.*?)\2/gi,
    (_match, property: string, quote: string, paint: string) =>
      `${property}=${quote}${wordCompatiblePaintValue(paint)}${quote}`,
  );
  output = output.replace(
    /\bstyle=(['"])(.*?)\1/gi,
    (_match, quote: string, style: string) =>
      `style=${quote}${forceStylePaintBlack(style)}${quote}`,
  );
  output = output.replace(
    /<style\b([^>]*)>([\s\S]*?)<\/style>/gi,
    (_match, attributes: string, css: string) =>
      `<style${attributes}>${forceStylePaintBlack(css)}</style>`,
  );

  const lower = output.toLowerCase();
  if (
    lower.includes("currentcolor") ||
    lower.includes("var(") ||
    /\b(?:color|fill|stroke)\s*[:=]\s*['"]?(?:inherit|white|#fff(?:fff)?)(?:['";\s>]|$)/i.test(
      output,
    )
  ) {
    throw new Error(
      "Word SVG export still contains a deferred or white paint style.",
    );
  }
  if (!/\b(?:fill|stroke)=["']#000000["']/i.test(output)) {
    throw new Error("Word SVG export is missing explicit black formula paint.");
  }
  return output;
}

function encodeUtf8Base64(value: string) {
  const bytes = new TextEncoder().encode(value);
  let binary = "";
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize));
  }
  return btoa(binary);
}

export function svgToBase64(svg: string) {
  return encodeUtf8Base64(svg);
}

export function latexToSvg(
  latex: string,
  options: SvgExportOptions = DEFAULT_OPTIONS,
): SvgExportResult {
  const source = prepareLatex(latex);
  const fontSizePt = positiveFinite(options.fontSizePt, DEFAULT_OPTIONS.fontSizePt);
  const paddingPx = nonNegativeFinite(options.paddingPx, DEFAULT_OPTIONS.paddingPx);
  const fontSizePx = fontSizePt * (96 / 72);
  const exPx = fontSizePx * 0.442;

  const container = mathDocument.convert(source, {
    display: options.displayMode,
    em: fontSizePx,
    ex: exPx,
    containerWidth: 100_000,
  });
  let svg = extractSvg(adaptor.outerHTML(container));
  const viewBox = parseViewBox(svg);

  const unitsPerPx = 1000 / fontSizePx;
  const paddingUnits = paddingPx * unitsPerPx;
  const padded = {
    x: viewBox.x - paddingUnits,
    y: viewBox.y - paddingUnits,
    width: viewBox.width + 2 * paddingUnits,
    height: viewBox.height + 2 * paddingUnits,
  };
  const width = Math.max(1, padded.width / unitsPerPx);
  const height = Math.max(1, padded.height / unitsPerPx);
  const baseline = Math.max(0, Math.min(height, -padded.y / unitsPerPx));

  svg = svg
    .replace(
      /\bviewBox=["'][^"']+["']/,
      `viewBox="${padded.x} ${padded.y} ${padded.width} ${padded.height}"`,
    )
    .replace(/^<svg\b([^>]*)>/, (_opening, rawAttributes: string) => {
      const attributes = rawAttributes
        .replace(
          /\s(?:xmlns|width|height|role|focusable|style)=["'][^"']*["']/g,
          "",
        )
        .trim();
      return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" role="img" focusable="false"${
        attributes ? ` ${attributes}` : ""
      }>`;
    })
    .replace(
      /currentColor/gi,
      options.forceExplicitBlack ? WORD_EXPLICIT_BLACK : "#111111",
    );

  const openingEnd = svg.indexOf(">");
  if (options.background === "white") {
    const background = `<rect x="${padded.x}" y="${padded.y}" width="${padded.width}" height="${padded.height}" fill="#ffffff"/>`;
    svg = `${svg.slice(0, openingEnd + 1)}${background}${svg.slice(openingEnd + 1)}`;
  } else {
    // PowerPoint otherwise hit-tests only the painted glyph paths of a
    // transparent SVG. A practically invisible filled rectangle makes the
    // entire formula bounds selectable and double-clickable at normal zoom.
    const hitTarget = `<rect x="${padded.x}" y="${padded.y}" width="${padded.width}" height="${padded.height}" fill="#000000" fill-opacity="0.001"/>`;
    svg = `${svg.slice(0, openingEnd + 1)}${hitTarget}${svg.slice(openingEnd + 1)}`;
  }

  if (options.forceExplicitBlack) {
    svg = forceWordCompatibleBlack(svg);
  }

  assertSelfContained(svg);
  return {
    svg,
    base64: svgToBase64(svg),
    width,
    height,
    baseline,
  };
}

function blobToBase64(blob: Blob) {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error ?? new Error("Unable to read PNG blob."));
    reader.onload = () => {
      const result = typeof reader.result === "string" ? reader.result : "";
      const comma = result.indexOf(",");
      resolve(comma >= 0 ? result.slice(comma + 1) : result);
    };
    reader.readAsDataURL(blob);
  });
}

function pngDataUrlToBlob(value: string) {
  const prefix = "data:image/png;base64,";
  if (!value.startsWith(prefix)) {
    throw new Error("Canvas did not produce a PNG data URL.");
  }
  const binary = atob(value.slice(prefix.length));
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return new Blob([bytes], { type: "image/png" });
}

async function encodeCanvasPng(canvas: HTMLCanvasElement) {
  if (typeof canvas.toBlob === "function") {
    const blob = await new Promise<Blob | null>((resolve) => {
      let settled = false;
      const finish = (value: Blob | null) => {
        if (settled) return;
        settled = true;
        window.clearTimeout(timeout);
        resolve(value);
      };
      const timeout = window.setTimeout(() => finish(null), 750);
      try {
        canvas.toBlob(finish, "image/png");
      } catch {
        finish(null);
      }
    });
    if (blob) return blob;
  }

  // WKWebView can expose canvas.toBlob() but return null for SVG-backed
  // canvases. toDataURL() uses a different WebKit encoding path and is stable
  // on the same canvas, so use it as the required Word compatibility fallback.
  return pngDataUrlToBlob(canvas.toDataURL("image/png"));
}

export async function svgToPng(
  svgResult: SvgExportResult,
  options: PngExportOptions = {},
): Promise<PngExportResult> {
  if (typeof document === "undefined" || typeof Image === "undefined") {
    throw new Error("PNG export requires a browser canvas environment.");
  }

  const scale = positiveFinite(options.scale ?? 2, 2);
  const width = Math.max(1, Math.ceil(svgResult.width * scale));
  const height = Math.max(1, Math.ceil(svgResult.height * scale));
  const image = new Image();
  image.decoding = "async";
  const source = `data:image/svg+xml;base64,${svgResult.base64}`;

  await new Promise<void>((resolve, reject) => {
    image.onload = () => resolve();
    image.onerror = () => reject(new Error("Unable to rasterize the generated SVG."));
    image.src = source;
  });

  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Unable to create a PNG canvas context.");
  if (options.background === "white") {
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, width, height);
  }
  context.drawImage(image, 0, 0, width, height);

  const pixels = context.getImageData(0, 0, width, height).data;
  let hasVisibleInk = false;
  for (let index = 0; index < pixels.length; index += 4) {
    const alpha = pixels[index + 3];
    if (
      alpha >= 16 &&
      (pixels[index] < 245 || pixels[index + 1] < 245 || pixels[index + 2] < 245)
    ) {
      hasVisibleInk = true;
      break;
    }
  }
  if (!hasVisibleInk) {
    throw new Error("PNG rasterization produced no visible formula ink.");
  }

  const blob = await encodeCanvasPng(canvas);
  return {
    blob,
    base64: await blobToBase64(blob),
    width,
    height,
  };
}
