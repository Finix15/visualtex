import {
  EXTENDED_INTEGRAL_MATHML_MACROS,
  EXTENDED_INTEGRAL_SVG_MACROS,
} from "./extendedIntegralCompatibility.ts";

export type VisualTexMathJaxMacro = string | readonly [replacement: string, argumentCount: number];

/**
 * Rendering-only compatibility aliases shared by every VisualTeX export path.
 * The saved LaTeX is intentionally left untouched. In particular, \bm must
 * retain the mathematical bold-italic semantics of amsmath/bm rather than be
 * degraded to upright \mathbf text.
 */
const STANDARD_COMPATIBILITY_MACROS = {
  bm: ["\\boldsymbol{#1}", 1] as const,
  // Common aliases from the LaTeX physics package. Documents copied from
  // lecture notes and papers frequently use \ip or \innerproduct without
  // carrying the original package preamble into VisualTeX.
  ip: ["\\left\\langle #1, #2 \\right\\rangle", 2] as const,
  innerproduct: ["\\left\\langle #1, #2 \\right\\rangle", 2] as const,
} satisfies Record<string, VisualTexMathJaxMacro>;

export const VISUALTEX_MATHML_MACROS = {
  ...EXTENDED_INTEGRAL_MATHML_MACROS,
  ...STANDARD_COMPATIBILITY_MACROS,
} satisfies Record<string, VisualTexMathJaxMacro>;

export const VISUALTEX_SVG_MACROS = {
  ...EXTENDED_INTEGRAL_SVG_MACROS,
  ...STANDARD_COMPATIBILITY_MACROS,
} satisfies Record<string, VisualTexMathJaxMacro>;

const unresolvedCommand = /\\[A-Za-z@]+/;
const mathText = /<mtext\b([^>]*)>([\s\S]*?)<\/mtext>/gi;

function stripXmlMarkup(value: string) {
  return value.replace(/<[^>]*>/g, "").replace(/&#x5c;|&#92;|&bsol;/gi, "\\");
}

/** Reject MathJax's permissive unknown-command fallback before it reaches Word. */
export function assertResolvedPresentationMathMl(mathMl: string) {
  mathText.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = mathText.exec(mathMl))) {
    const attributes = match[1] ?? "";
    const text = stripXmlMarkup(match[2] ?? "");
    if (unresolvedCommand.test(text) || /mathcolor=["']?red/i.test(attributes)) {
      const command = text.match(unresolvedCommand)?.[0] ?? (text.trim() || "unknown command");
      throw new Error(`MathJax did not resolve LaTeX command ${command}.`);
    }
  }
  if (/<merror\b/i.test(mathMl)) {
    throw new Error("MathJax produced an error node instead of semantic MathML.");
  }
}

/** The document-import preview uses SVG directly, so guard that branch too. */
export function assertResolvedMathJaxSvg(svg: string) {
  if (
    /<g\b[^>]*data-mml-node=["']mtext["'][^>]*(?:fill|stroke)=["']red["']/i.test(svg)
    || /<g\b[^>]*(?:fill|stroke)=["']red["'][^>]*data-mml-node=["']mtext["']/i.test(svg)
    || /data-mml-node=["']merror["']/i.test(svg)
  ) {
    const text = stripXmlMarkup(svg)
      .replace(/&lt;/gi, "<")
      .replace(/&gt;/gi, ">")
      .replace(/&amp;/gi, "&");
    const command = text.match(unresolvedCommand)?.[0];
    throw new Error(
      command
        ? `MathJax did not resolve LaTeX command ${command}.`
        : "MathJax left an unresolved LaTeX command in the rendered formula.",
    );
  }
}
