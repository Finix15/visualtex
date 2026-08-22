import { commandRegistry } from "./commandRegistry";
import type { LatexCommand } from "../types/command";
import { readCustomSymbolLibrary } from "../math/customSymbolRegistry";

export function customSymbolCommands(): LatexCommand[] {
  return readCustomSymbolLibrary().symbols.map((symbol, index) => ({
    id: `custom-symbol:${symbol.id}`,
    command: `\\${symbol.command}`,
    insertTemplate: `\\${symbol.command}`,
    previewLatex: `\\${symbol.command}`,
    labelVi: symbol.name,
    labelEn: symbol.name,
    aliases: [symbol.command],
    keywords: ["Custom characters", "Custom symbol", "custom symbol"],
    category: "common",
    defaultPriority: Math.max(55, 84 - index),
    supportedInMathMode: true,
  }));
}

export function getRuntimeCommandRegistry() {
  const custom = customSymbolCommands();
  return custom.length ? [...commandRegistry, ...custom] : commandRegistry;
}

export function findRuntimeCommandByCommand(command: string) {
  const normalized = command.trim();
  return (
    getRuntimeCommandRegistry().find(
      (candidate) =>
        candidate.command === normalized ||
        candidate.insertTemplate === normalized,
    ) ?? null
  );
}
