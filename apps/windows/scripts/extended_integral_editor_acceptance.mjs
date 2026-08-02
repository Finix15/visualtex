import assert from "node:assert/strict";
import { convertLatexToMarkup } from "mathlive";

const cases = [
  ["oiint", "∯"],
  ["oiiint", "∰"],
];

for (const [command, symbol] of cases) {
  const markup = convertLatexToMarkup(`\\${command}_{\\Sigma} a`);
  assert.match(markup, new RegExp(symbol), `\\${command} must use MathLive's native symbol`);
  assert.doesNotMatch(
    markup,
    /visualtex-extended-integral|scaleX\(|scaleY\(|transform:/i,
    `\\${command} must not use VisualTeX glyph stretching`,
  );
}

console.log("Extended-integral editor acceptance passed without custom glyph stretching.");
