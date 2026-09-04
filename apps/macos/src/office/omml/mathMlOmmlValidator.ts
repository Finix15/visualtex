const MATH_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math";

const TOKEN_ELEMENTS = new Set(["mi", "mn", "mo", "mtext", "ms"]);
const STRUCTURAL_TOKEN_CHARACTERS = new Set([
  "^", "~", "¯", "‾", "―", "_", "˙", "¨", "´", "`", "ˇ", "˘", "←", "→", "↔", "⃗", "̂", "̃", "̇", "̈",
  "⏞", "︷", "︵", "⏟", "︸", "︶",
  "∑", "∏", "∐", "∫", "∮", "⋂", "⋃",
]);
const STRUCTURE_RULES: Record<string, string[]> = {
  mfrac: ["f"],
  msqrt: ["rad"],
  mroot: ["rad"],
  msub: ["sSub", "sSubSup", "nary"],
  msup: ["sSup", "sSubSup", "nary"],
  msubsup: ["sSubSup", "nary"],
  mmultiscripts: ["sPre"],
  mfenced: ["d"],
  menclose: ["borderBox", "box"],
};
const NARY = new Set([
  "∑", "∏", "∐", "⋂", "⋃", "⨀", "⨁", "⨂", "⨄", "⨆",
  "∫", "∬", "∭", "⨌", "∮", "∯", "∰", "∱", "∲", "∳",
  "⨋", "⨍", "⨎", "⨏", "⨐", "⨑", "⨒", "⨓", "⨔", "⨕",
  "⨖", "⨗", "⨘", "⨙", "⨚", "⨛", "⨜",
]);
const ACCENTS = new Set(["^", "~", "˙", "¨", "´", "`", "ˇ", "˘", "→", "←", "↔", "⃗", "¯", "‾", "_"]);

export interface MathMlOmmlValidationResult {
  valid: boolean;
  errors: string[];
  warnings: string[];
}

function parseXml(value: string, label: string) {
  if (typeof DOMParser === "undefined") {
    throw new Error("MathML-to-OMML validation requires a browser DOM parser.");
  }
  if (/<!DOCTYPE|<!ENTITY/i.test(value)) {
    throw new Error(`${label} must not contain a DTD or entity declaration.`);
  }
  const documentObject = new DOMParser().parseFromString(value, "application/xml");
  const parseError = documentObject.querySelector("parsererror");
  if (parseError) {
    throw new Error(`${label} parse error: ${parseError.textContent ?? "invalid XML"}`);
  }
  return documentObject.documentElement;
}

function descendants(root: Element, localName: string) {
  return Array.from(root.getElementsByTagNameNS("*", localName));
}

function normalized(value: string) {
  return value
    .replace(/[\u2061-\u2064\ufeff\u200b\u2009\u200a\u205f\u2003]/g, "")
    .replace(/[−–—]/g, "-")
    .replaceAll("，", ",")
    .replaceAll("（", "(")
    .replaceAll("）", ")")
    .replace(/\s+/g, "");
}

function mathMlTokens(root: Element) {
  return Array.from(root.getElementsByTagNameNS("*", "*"))
    .filter((element) => TOKEN_ELEMENTS.has(element.localName))
    .map((element) =>
      normalized(element.textContent ?? "")
        .split("")
        .filter((character) => !STRUCTURAL_TOKEN_CHARACTERS.has(character) && !NARY.has(character))
        .join(""),
    )
    .filter(Boolean);
}

function ommlText(root: Element) {
  const runs = descendants(root, "t").map((element) => element.textContent ?? "");
  const delimiters = descendants(root, "d").flatMap((element) => {
    const properties = Array.from(element.children).find((child) => child.localName === "dPr");
    if (!properties) return ["(", ")"];
    const value = (name: string, fallback: string) => {
      const property = Array.from(properties.children).find((child) => child.localName === name);
      return property ? property.getAttributeNS(MATH_NAMESPACE, "val") ?? property.getAttribute("m:val") ?? "" : fallback;
    };
    return [value("begChr", "("), value("endChr", ")")];
  });
  return normalized([...runs, ...delimiters].join(""));
}

function count(root: Element, name: string) {
  return descendants(root, name).length;
}

function visible(element: Element) {
  return normalized(element.textContent ?? "").length > 0 ||
    Array.from(element.attributes).some((attribute) => attribute.localName === "val" && attribute.value.length > 0);
}

export function validateMathMlToOmml(
  mathMl: string,
  omml: string,
): MathMlOmmlValidationResult {
  const errors: string[] = [];
  const warnings: string[] = [];
  let mathRoot: Element;
  let ommlRoot: Element;
  try {
    mathRoot = parseXml(mathMl, "MathML");
    ommlRoot = parseXml(omml, "OMML");
  } catch (error) {
    return { valid: false, errors: [error instanceof Error ? error.message : String(error)], warnings };
  }
  if (mathRoot.localName !== "math") errors.push("MathML root must be math.");
  if (ommlRoot.localName !== "oMath" || ommlRoot.namespaceURI !== MATH_NAMESPACE) {
    errors.push("OMML root must be m:oMath in the Office Math namespace.");
  }

  const outputText = ommlText(ommlRoot);
  const expectedTokens = new Map<string, number>();
  for (const token of mathMlTokens(mathRoot)) {
    expectedTokens.set(token, (expectedTokens.get(token) ?? 0) + 1);
  }
  for (const [token, expected] of expectedTokens) {
    let found = 0;
    for (let offset = 0; token && offset <= outputText.length - token.length;) {
      const index = outputText.indexOf(token, offset);
      if (index < 0) break;
      found += 1;
      offset = index + token.length;
    }
    if (found < expected) errors.push(`OMML token loss: ${JSON.stringify(token)} expected ${expected}, found ${found}.`);
  }
  for (const [mathName, ommlNames] of Object.entries(STRUCTURE_RULES)) {
    const expected = count(mathRoot, mathName);
    const found = ommlNames.reduce((total, name) => total + count(ommlRoot, name), 0);
    if (found < expected) errors.push(`OMML structure loss: ${mathName} expected ${expected}, found ${found}.`);
  }

  for (const slotName of ["num", "den", "deg", "lim", "fName"]) {
    for (const slot of descendants(ommlRoot, slotName)) {
      if (slotName === "deg") {
        const radical = slot.parentElement;
        const hidden = radical
          ? descendants(radical, "degHide").some((property) => {
              const value = property.getAttributeNS(MATH_NAMESPACE, "val") ?? property.getAttribute("m:val");
              return value === "1" || value === "on" || value === "true";
            })
          : false;
        if (hidden) continue;
      }
      if (!visible(slot)) errors.push(`Empty OMML critical slot: m:${slotName}.`);
    }
  }
  const requiredSubscripts = count(mathRoot, "msub") + count(mathRoot, "msubsup");
  const requiredSuperscripts = count(mathRoot, "msup") + count(mathRoot, "msubsup");
  const visibleSubscripts = descendants(ommlRoot, "sub").filter(visible).length;
  const visibleSuperscripts = descendants(ommlRoot, "sup").filter(visible).length;
  if (visibleSubscripts < requiredSubscripts) errors.push("Empty or missing OMML subscript slot.");
  if (visibleSuperscripts < requiredSuperscripts) errors.push("Empty or missing OMML superscript slot.");

  const roots = descendants(mathRoot, "mroot").length;
  if (descendants(ommlRoot, "deg").filter(visible).length < roots) {
    errors.push("OMML root degree loss.");
  }
  const mathRows = descendants(mathRoot, "mtr").length + descendants(mathRoot, "mlabeledtr").length;
  const mathCells = descendants(mathRoot, "mtd").length;
  const matrixRows = descendants(ommlRoot, "mr");
  const equationArrayRows = descendants(ommlRoot, "eqArr").flatMap((array) =>
    Array.from(array.children).filter((child) => child.localName === "e"),
  );
  const ommlRows = matrixRows.length + equationArrayRows.length;
  const ommlCells = matrixRows.reduce(
    (total, row) => total + Array.from(row.children).filter((child) => child.localName === "e").length,
    0,
  );
  if (mathRows && ommlRows < mathRows) errors.push("OMML matrix row loss.");
  if (mathCells && !equationArrayRows.length && ommlCells < mathCells) errors.push("OMML matrix cell loss.");

  const hasNaryLimits = ["munder", "mover", "munderover", "msub", "msup", "msubsup"].some((name) =>
    descendants(mathRoot, name).some((element) => {
      const base = Array.from(element.children)[0];
      return Boolean(base && NARY.has(normalized(base.textContent ?? "")));
    }),
  );
  if (hasNaryLimits && !count(ommlRoot, "nary")) errors.push("OMML n-ary limit loss.");

  const accentCount = ["mover", "munder"].flatMap((name) => descendants(mathRoot, name)).filter((element) => {
    const mark = Array.from(element.children)[1];
    return Boolean(mark && ACCENTS.has(normalized(mark.textContent ?? "")));
  }).length;
  const ommlAccentCount = count(ommlRoot, "acc") + count(ommlRoot, "bar") + count(ommlRoot, "groupChr");
  if (ommlAccentCount < accentCount) errors.push("OMML accent/bar loss.");

  if (descendants(mathRoot, "mprescripts").length && !count(ommlRoot, "sPre")) {
    errors.push("OMML prescript loss.");
  }
  return { valid: errors.length === 0, errors, warnings };
}

export function assertValidMathMlToOmml(mathMl: string, omml: string) {
  const validation = validateMathMlToOmml(mathMl, omml);
  if (!validation.valid) {
    throw new Error(`MathML-to-OMML validation failed: ${validation.errors.join(" ")}`);
  }
  return validation;
}
