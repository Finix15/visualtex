export const EXTENDED_INTEGRAL_SYMBOLS = {
  oiint: "∯",
  oiiint: "∰",
  intclockwise: "∱",
  varointclockwise: "∲",
  ointctrclockwise: "∳",
  intctrclockwise: "⨑",
} as const;

/**
 * MathJax does not register the esint commands used by MathLive. Map the
 * commands to their native Unicode operators without applying any visual
 * scaling, stretching, positioning, or replacement glyph construction.
 * Saved LaTeX remains unchanged, and Word can still convert the semantic
 * MathML operators into native OMML n-ary structures.
 */
export const EXTENDED_INTEGRAL_MATHML_MACROS = {
  ...EXTENDED_INTEGRAL_SYMBOLS,
} as const;

export const EXTENDED_INTEGRAL_SVG_MACROS = {
  ...EXTENDED_INTEGRAL_SYMBOLS,
} as const;
