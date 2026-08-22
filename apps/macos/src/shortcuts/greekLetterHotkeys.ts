import type { LatexCommand } from "../types/command";

type GreekShortcutDefinition = {
  id: string;
  latex: string;
  labelVi: string;
  labelEn: string;
};

type GreekShortcutEvent = Pick<
  KeyboardEvent,
  | "code"
  | "key"
  | "ctrlKey"
  | "altKey"
  | "shiftKey"
  | "metaKey"
  | "isComposing"
  | "repeat"
>;

const lowerGreekByCode: Record<string, GreekShortcutDefinition> = {
  KeyA: { id: "alpha", latex: "\\alpha", labelVi: "Alpha", labelEn: "Alpha" },
  KeyB: { id: "beta", latex: "\\beta", labelVi: "Beta", labelEn: "Beta" },
  KeyG: { id: "gamma", latex: "\\gamma", labelVi: "Gamma", labelEn: "Gamma" },
  KeyD: { id: "delta", latex: "\\delta", labelVi: "Đồng bằng", labelEn: "Delta" },
  KeyE: { id: "epsilon", latex: "\\epsilon", labelVi: "Epsilon", labelEn: "Epsilon" },
  KeyZ: { id: "zeta", latex: "\\zeta", labelVi: "Zeta", labelEn: "Zeta" },
  KeyH: { id: "eta", latex: "\\eta", labelVi: "Eta", labelEn: "Eta" },
  KeyQ: { id: "theta", latex: "\\theta", labelVi: "Theta", labelEn: "Theta" },
  KeyI: { id: "iota", latex: "\\iota", labelVi: "Iota", labelEn: "Iota" },
  KeyK: { id: "kappa", latex: "\\kappa", labelVi: "Kappa", labelEn: "Kappa" },
  KeyL: { id: "lambda", latex: "\\lambda", labelVi: "Lambda", labelEn: "Lambda" },
  KeyM: { id: "mu", latex: "\\mu", labelVi: "Mu", labelEn: "Mu" },
  KeyN: { id: "nu", latex: "\\nu", labelVi: "Nữ", labelEn: "Nu" },
  KeyX: { id: "xi", latex: "\\xi", labelVi: "Xi", labelEn: "Xi" },
  // Standard LaTeX has no \\omicron command because lowercase omicron is
  // visually identical to a Latin o. Use the conventional LaTeX spelling.
  KeyO: { id: "omicron", latex: "o", labelVi: "Omicron", labelEn: "Omicron" },
  KeyP: { id: "pi", latex: "\\pi", labelVi: "Pi", labelEn: "Pi" },
  KeyR: { id: "rho", latex: "\\rho", labelVi: "Rho", labelEn: "Rho" },
  KeyS: { id: "sigma", latex: "\\sigma", labelVi: "Sigma", labelEn: "Sigma" },
  KeyT: { id: "tau", latex: "\\tau", labelVi: "Tàu", labelEn: "Tau" },
  KeyU: { id: "upsilon", latex: "\\upsilon", labelVi: "Upsilon", labelEn: "Upsilon" },
  KeyF: { id: "phi", latex: "\\phi", labelVi: "Phi", labelEn: "Phi" },
  KeyC: { id: "chi", latex: "\\chi", labelVi: "Chi", labelEn: "Chi" },
  KeyY: { id: "psi", latex: "\\psi", labelVi: "Psi", labelEn: "Psi" },
  KeyW: { id: "omega", latex: "\\omega", labelVi: "Omega", labelEn: "Omega" },
};

const upperGreekByCode: Record<string, GreekShortcutDefinition> = {
  KeyA: { id: "Alpha", latex: "A", labelVi: "Chữ viết hoa", labelEn: "Capital alpha" },
  KeyB: { id: "Beta", latex: "B", labelVi: "Vốn beta", labelEn: "Capital beta" },
  KeyG: { id: "Gamma", latex: "\\Gamma", labelVi: "Gamma vốn", labelEn: "Capital gamma" },
  KeyD: { id: "Delta", latex: "\\Delta", labelVi: "Đồng bằng thủ đô", labelEn: "Capital delta" },
  KeyE: { id: "Epsilon", latex: "E", labelVi: "Vốn epsilon", labelEn: "Capital epsilon" },
  KeyZ: { id: "Zeta", latex: "Z", labelVi: "Vốn zeta", labelEn: "Capital zeta" },
  KeyH: { id: "Eta", latex: "H", labelVi: "Vốn eta", labelEn: "Capital eta" },
  KeyQ: { id: "Theta", latex: "\\Theta", labelVi: "Vốn theta", labelEn: "Capital theta" },
  KeyI: { id: "Iota", latex: "I", labelVi: "Vốn iota", labelEn: "Capital iota" },
  KeyK: { id: "Kappa", latex: "K", labelVi: "Vốn kappa", labelEn: "Capital kappa" },
  KeyL: { id: "Lambda", latex: "\\Lambda", labelVi: "Vốn lambda", labelEn: "Capital lambda" },
  KeyM: { id: "Mu", latex: "M", labelVi: "Vốn mu", labelEn: "Capital mu" },
  KeyN: { id: "Nu", latex: "N", labelVi: "Vốn nu", labelEn: "Capital nu" },
  KeyX: { id: "Xi", latex: "\\Xi", labelVi: "Vốn xi", labelEn: "Capital xi" },
  KeyO: { id: "Omicron", latex: "O", labelVi: "Vốn omron", labelEn: "Capital omicron" },
  KeyP: { id: "Pi", latex: "\\Pi", labelVi: "Vốn pi", labelEn: "Capital pi" },
  // Greek capital rho and chi are conventionally represented by the same
  // glyphs as Latin P and X in LaTeX math mode.
  KeyR: { id: "Rho", latex: "P", labelVi: "Vốn rho", labelEn: "Capital rho" },
  KeyS: { id: "Sigma", latex: "\\Sigma", labelVi: "Sigma viết hoa", labelEn: "Capital sigma" },
  KeyT: { id: "Tau", latex: "T", labelVi: "Vốn tàu", labelEn: "Capital tau" },
  KeyU: { id: "Upsilon", latex: "\\Upsilon", labelVi: "Vốn upsilon", labelEn: "Capital upsilon" },
  KeyF: { id: "Phi", latex: "\\Phi", labelVi: "Vốn phi", labelEn: "Capital phi" },
  KeyC: { id: "Chi", latex: "X", labelVi: "Vốn chi", labelEn: "Capital chi" },
  KeyY: { id: "Psi", latex: "\\Psi", labelVi: "Vốn psi", labelEn: "Capital psi" },
  KeyW: { id: "Omega", latex: "\\Omega", labelVi: "Vốn omega", labelEn: "Capital omega" },
};

function toLatexCommand(definition: GreekShortcutDefinition): LatexCommand {
  return {
    id: `greek-hotkey-${definition.id}`,
    command: definition.latex,
    insertTemplate: definition.latex,
    previewLatex: definition.latex,
    labelVi: definition.labelVi,
    labelEn: definition.labelEn,
    aliases: [],
    keywords: ["Greek letters", "Greek"],
    category: "greek",
    defaultPriority: 0,
    supportedInMathMode: true,
  };
}

export function isGreekLetterHotkeyPrefix(event: GreekShortcutEvent) {
  return (
    !event.repeat &&
    !event.isComposing &&
    event.code === "KeyG" &&
    event.metaKey &&
    !event.ctrlKey &&
    !event.altKey &&
    !event.shiftKey
  );
}

export function greekLetterHotkeyCommandFromEvent(
  event: GreekShortcutEvent,
): LatexCommand | null {
  if (
    event.repeat ||
    event.isComposing ||
    event.ctrlKey ||
    event.altKey ||
    event.metaKey ||
    event.key === "Process"
  ) {
    return null;
  }
  const definition = event.shiftKey
    ? upperGreekByCode[event.code]
    : lowerGreekByCode[event.code];
  return definition ? toLatexCommand(definition) : null;
}

export const greekLetterHotkeyLowercaseCodes = Object.keys(lowerGreekByCode);
