import { spawn, spawnSync } from "node:child_process";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { homedir, tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { zipSync, strToU8 } from "fflate";
import { latexToSvg } from "../src/export/latexToSvg.ts";
import {
  createFormulaMetadata,
  decodeFormulaMetadata,
  encodeFormulaMetadata,
} from "../src/office/shared/formulaMetadata.ts";
import {
  normalizeFormulaEditorDocument,
  serializeFormulaEditorDocument,
} from "../src/office/shared/formulaEditorDocument.ts";
import { renderOfficeFormulaArtifacts } from "../src/office/shared/formulaRenderArtifacts.ts";

const repositoryRoot = resolve(new URL("..", import.meta.url).pathname);
const templatePath = join(
  repositoryRoot,
  "office/macos-offline/resources/VisualTeX.dotm",
);
const runtimeRoot = join(
  homedir(),
  "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime",
);
const pdfExportRequestPath = join(
  runtimeRoot,
  "document-import-regression-pdf-path.txt",
);
const pdfExportStatusPath = join(
  runtimeRoot,
  "document-import-regression-pdf-status.txt",
);
const imageEditStatusPath = join(
  runtimeRoot,
  "document-import-regression-image-edit-status.txt",
);
const formulaRegressionStatusPath = join(
  runtimeRoot,
  "document-import-regression-formula-status.txt",
);
const physicalScreenBoundsPath = join(
  runtimeRoot,
  "physical-double-click-screen-bounds.txt",
);
const workspaceVisualTeXBinary = join(
  repositoryRoot,
  "src-tauri/target/release/bundle/macos/VisualTeX.app/Contents/MacOS/visualtex",
);
const officeScratchRoot = join(
  homedir(),
  "Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch",
);
const wordStartupRoot = join(
  homedir(),
  "Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word",
);
const installedWordAddinPath = join(wordStartupRoot, "VisualTeX.dotm");
const installedWordAddinBackupPath = join(
  tmpdir(),
  `visualtex-installed-word-backup-${process.pid}.dotm`,
);
const coordinatePdfPath = join(
  officeScratchRoot,
  `document-import-geometry-${process.pid}.pdf`,
);
const reopenedDocumentPath = join(
  officeScratchRoot,
  `document-import-reopen-${process.pid}.docx`,
);
const sessionsRoot = join(runtimeRoot, "OfficeSessions");
const nativeRoot = join(runtimeRoot, "NativeDocuments");
const physicalDoubleClick = process.argv.includes("--physical-double-click");
const createImageRegression = process.argv.includes("--create-image");
const physicalTargets = new Set([
  "image-inline",
  "image-block",
  "image-align",
  "image-align-star",
  "omml-inline",
  "omml-block",
  "omml-align",
  "omml-align-star",
]);

function commandLineOption(name) {
  const exactIndex = process.argv.indexOf(name);
  const assigned = process.argv.find((argument) =>
    argument.startsWith(`${name}=`),
  );
  if (exactIndex >= 0 && assigned) {
    throw new Error(`Specify ${name} only once`);
  }
  if (exactIndex >= 0) {
    const value = process.argv[exactIndex + 1];
    if (!value || value.startsWith("--")) {
      throw new Error(`${name} requires a value`);
    }
    return value;
  }
  return assigned?.slice(name.length + 1) ?? "";
}

const physicalTarget = commandLineOption("--physical-target");
const itemLimitOption = commandLineOption("--item-limit");
const diagnosticItemLimit = itemLimitOption ? Number(itemLimitOption) : 17;
if (
  !Number.isInteger(diagnosticItemLimit) ||
  diagnosticItemLimit < 1 ||
  diagnosticItemLimit > 17
) {
  throw new Error("--item-limit must be an integer from 1 through 17");
}
if (physicalDoubleClick && !physicalTargets.has(physicalTarget)) {
  throw new Error(
    "--physical-double-click requires --physical-target " +
      [...physicalTargets].join("|"),
  );
}
if (!physicalDoubleClick && physicalTarget) {
  throw new Error("--physical-target requires --physical-double-click");
}
if (createImageRegression && physicalDoubleClick) {
  throw new Error("--create-image cannot be combined with --physical-double-click");
}
const physicalOutputKind = physicalTarget.split("-", 1)[0];
const outputKind = createImageRegression
  ? "image"
  : physicalDoubleClick
    ? physicalOutputKind
    : process.argv.includes("--image")
      ? "image"
      : "omml";
if (
  physicalDoubleClick &&
  process.argv.includes("--image") &&
  outputKind !== "image"
) {
  throw new Error("--image conflicts with the selected OMML physical target");
}
const referenceFontSizePt = 14;
const wordImageVisualScale = 1.1;
const wordDisplayPaddingPx = 2;
const nativeCalibrationWidthPt = 95.71632;
const editorReadyFileName = "editor-ready.json";
const editorPerformanceFileName = "editor-performance.jsonl";
const editorReadySchema = "visualtex-office-editor-ready-v1";
const editorPerformanceSchema = "visualtex-office-editor-performance-v1";
const warmEditorReadyLimitMs = 500;
const diagnosticSuccessPrefix = "VISUALTEX_DOCUMENT_IMPORT_DIAGNOSTIC_PASS:";
const transparentPng = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAEAQH/8l0Z8QAAAABJRU5ErkJggg==",
  "base64",
);
const sleep = (ms) => new Promise((resolvePromise) => setTimeout(resolvePromise, ms));

async function stopVisualTeXForManualWordCallback() {
  spawnSync("/usr/bin/killall", ["visualtex"], {
    encoding: "utf8",
    timeout: 10_000,
  });
  await sleep(1_000);
}

async function startVisualTeXForPhysicalRegression() {
  await stopVisualTeXForManualWordCallback();
  if (!existsSync(workspaceVisualTeXBinary)) {
    throw new Error(
      `The workspace VisualTeX validation app is missing: ${workspaceVisualTeXBinary}`,
    );
  }
  const child = spawn(workspaceVisualTeXBinary, ["--office-background"], {
    detached: true,
    stdio: "ignore",
  });
  child.unref();
  await sleep(2_000);
}

function physicallyDoubleClickSelectedWordFormula(testDocumentName) {
  rmSync(physicalScreenBoundsPath, { force: true });
  runAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(testDocumentName)}`,
    "activate",
    'run VB macro macro name "VisualTeX_WriteSelectedScreenBoundsForRegression"',
    "end tell",
  ], 30_000);
  if (!existsSync(physicalScreenBoundsPath)) {
    throw new Error("Word did not write physical double-click screen bounds");
  }
  const status = readFileSync(physicalScreenBoundsPath, "utf8").trim();
  const [result, leftText, topText, widthText, heightText] = status.split("|");
  const values = [leftText, topText, widthText, heightText].map(Number);
  if (
    result !== "PASS" ||
    values.some((value) => !Number.isFinite(value)) ||
    values[2] <= 0 ||
    values[3] <= 0
  ) {
    throw new Error(`Word returned invalid physical screen bounds: ${status}`);
  }
  const clickX = values[0] + values[2] / 2;
  const clickY = values[1] + values[3] / 2;
  const click = spawnSync(
    "/usr/bin/swift",
    [
      join(repositoryRoot, "scripts/macos_physical_double_click.swift"),
      clickX.toFixed(3),
      clickY.toFixed(3),
      "--appkit-y",
    ],
    {
      encoding: "utf8",
      timeout: 30_000,
      maxBuffer: 2 * 1024 * 1024,
    },
  );
  if (click.status !== 0) {
    throw new Error(
      click.stderr.trim() || click.stdout.trim() || "Quartz physical double-click failed",
    );
  }
  const clickResult = click.stdout.trim();
  return {
    wordScreenCenterX: clickX,
    wordScreenCenterY: clickY,
    screenBounds: values,
    quartzResult: clickResult,
  };
}
const legacyAlignLatex = String.raw`\begin{align}
1 &= 22 + 333 \\
44444 &= 55
\end{align}`;
const legacyAlignStarLatex = String.raw`\begin{align*}
666 &= 777 + 8 \\
999999 &= 0
\end{align*}`;

function runAppleScript(lines, timeout = 60_000) {
  const args = lines.flatMap((line) => ["-e", line]);
  const result = spawnSync("/usr/bin/osascript", args, {
    encoding: "utf8",
    timeout,
    maxBuffer: 8 * 1024 * 1024,
  });
  if (result.status !== 0) {
    const details = [
      result.stderr?.trim(),
      result.stdout?.trim(),
      result.error?.message,
      result.signal ? `signal=${result.signal}` : "",
      `status=${String(result.status)}`,
    ].filter(Boolean);
    throw new Error(details.join("\n") || "AppleScript failed");
  }
  return result.stdout.trim();
}

function base64Url(value) {
  return Buffer.from(value, "utf8").toString("base64url");
}

const MATH_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/math";
const WORD_NAMESPACE =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function escapeXml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function ommlRun(text, align = false) {
  if (!text) return "";
  const runProperties = align
    ? '<m:rPr><m:scr m:val="roman"/><m:sty m:val="p"/></m:rPr>'
    : "";
  const equationArrayAlignment = align ? "&" : "";
  return `<m:r>${runProperties}<m:t>${escapeXml(`${equationArrayAlignment}${text}`)}</m:t></m:r>`;
}

function ommlBodyForLatex(latex, alignRelation = false) {
  const normalized = latex.replaceAll("&", "").trim();
  if (alignRelation) {
    const relationIndex = normalized.indexOf("=");
    if (relationIndex >= 0) {
      return (
        ommlRun(normalized.slice(0, relationIndex)) +
        ommlRun("=", true) +
        ommlRun(normalized.slice(relationIndex + 1))
      );
    }
  }
  const superscript = latex.match(/^(.*?)([A-Za-z])\^\{?(\d+)\}?$/);
  return superscript
    ? `<m:r><m:t>${escapeXml(superscript[1])}</m:t></m:r>` +
      `<m:sSup><m:e><m:r><m:t>${escapeXml(superscript[2])}</m:t></m:r></m:e>` +
      `<m:sup><m:r><m:t>${escapeXml(superscript[3])}</m:t></m:r></m:sup></m:sSup>`
    : ommlRun(normalized);
}

function ommlForFormula(lines, codeFormat) {
  const relationAligned = [
    "align",
    "align-star",
    "aligned",
    "equation-split",
    "equation-star-split",
  ].includes(codeFormat);
  const converted = lines.map((line) =>
    ommlBodyForLatex(line, relationAligned),
  );
  const body =
    converted.length === 1 && !relationAligned
      ? converted[0]
      : `<m:eqArr><m:eqArrPr><m:baseJc m:val="center"/></m:eqArrPr>${converted
          .map((line) => `<m:e>${line}</m:e>`)
          .join("")}</m:eqArr>`;
  return `<m:oMath xmlns:m="${MATH_NAMESPACE}" xmlns:w="${WORD_NAMESPACE}">${body}</m:oMath>`;
}

function minimalDocxBytes(omml) {
  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    "</Types>";
  const relationships =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    "</Relationships>";
  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:m="${MATH_NAMESPACE}">` +
    `<w:body><w:p>${omml}</w:p><w:sectPr/></w:body></w:document>`;
  return zipSync(
    {
      "[Content_Types].xml": strToU8(contentTypes),
      "_rels/.rels": strToU8(relationships),
      "word/document.xml": strToU8(documentXml),
    },
    { level: 6 },
  );
}

function wordSvgDocxBytes(svg, png, widthPoints, heightPoints) {
  const widthEmu = Math.round(widthPoints * 12_700);
  const heightEmu = Math.round(heightPoints * 12_700);
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="svg" ContentType="image/svg+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
  const packageRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
  const documentRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPng" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/formula.png"/>
  <Relationship Id="rIdSvg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/formula.svg"/>
</Relationships>`;
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
  <w:body><w:p><w:r><w:drawing>
    <wp:inline distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="${widthEmu}" cy="${heightEmu}"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:docPr id="1" name="VisualTeX Formula" descr="VisualTeX SVG formula"/>
      <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:nvPicPr><pic:cNvPr id="0" name="formula.svg"/><pic:cNvPicPr/></pic:nvPicPr>
          <pic:blipFill><a:blip r:embed="rIdPng" cstate="print"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
          <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></pic:spPr>
        </pic:pic>
      </a:graphicData></a:graphic>
    </wp:inline>
  </w:drawing></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body>
</w:document>`;
  return zipSync(
    {
      "[Content_Types].xml": strToU8(contentTypes),
      "_rels/.rels": strToU8(packageRelationships),
      "word/document.xml": strToU8(documentXml),
      "word/_rels/document.xml.rels": strToU8(documentRelationships),
      "word/media/formula.png": new Uint8Array(png),
      "word/media/formula.svg": strToU8(svg),
    },
    { level: 6 },
  );
}

function calculateImageGeometry(svg, fontSizePt) {
  const naturalWidthPt = svg.width * 0.75 * wordImageVisualScale;
  const naturalHeightPt = svg.height * 0.75 * wordImageVisualScale;
  const referenceScale = Math.min(1, 500 / naturalWidthPt);
  const referenceWidthPt = naturalWidthPt * referenceScale;
  const referenceHeightPt = naturalHeightPt * referenceScale;
  const baselinePx = svg.baseline ?? svg.height;
  const descentRatio = Math.max(0, Math.min(1, (svg.height - baselinePx) / svg.height));
  const referenceBaselinePt = Math.max(
    -256,
    Math.min(0, -Math.max(0, referenceHeightPt * descentRatio)),
  );
  const pointScale = fontSizePt / referenceFontSizePt;
  return {
    widthPoints: referenceWidthPt * pointScale,
    heightPoints: referenceHeightPt * pointScale,
    baseline: Math.max(-256, Math.min(0, Math.round(referenceBaselinePt * pointScale))),
    referenceWidthPt,
    referenceHeightPt,
    referenceBaselinePt,
  };
}

function svgRelationshipPositions(svgMarkup) {
  const relationshipPattern =
    /<g data-mml-node="mtd" transform="translate\(([-+\d.]+),[-+\d.]+\)">(?:(?!<g data-mml-node="mtd")[\s\S])*?<g data-mml-node="mo" transform="translate\(([-+\d.]+),[-+\d.]+\)"><use data-c="3D"/g;
  return [...svgMarkup.matchAll(relationshipPattern)].map(
    (match) => Number(match[1]) + Number(match[2]),
  );
}

function assertAlignedSvg(svgMarkup, expectedRows, label) {
  const positions = svgRelationshipPositions(svgMarkup);
  if (
    positions.length !== expectedRows ||
    positions.some((position) => !Number.isFinite(position))
  ) {
    throw new Error(
      `${label} did not expose ${expectedRows} SVG relationship positions: ${JSON.stringify(positions)}`,
    );
  }
  const spread = Math.max(...positions) - Math.min(...positions);
  if (spread > 0.01) {
    throw new Error(
      `${label} SVG relationship columns are not aligned: ${JSON.stringify({ positions, spread })}`,
    );
  }
  return { positions, spread };
}

function rasterBounds(band, components) {
  const minX = Math.min(...components.map((component) => component.minX));
  const maxX = Math.max(...components.map((component) => component.maxX));
  return {
    minX,
    maxX,
    width: maxX - minX,
    centerX: (minX + maxX) / 2,
    minY: band.minY,
    maxY: band.maxY,
    height: band.height,
    centerY: band.centerY,
    components,
  };
}

function rasterEntryBounds(entries) {
  if (!entries.length) return null;
  const components = entries.map(({ component }) => component);
  const minX = Math.min(...components.map((component) => component.minX));
  const maxX = Math.max(...components.map((component) => component.maxX));
  const minY = Math.min(...entries.map(({ band }) => band.minY));
  const maxY = Math.max(...entries.map(({ band }) => band.maxY));
  return {
    minX,
    maxX,
    width: maxX - minX,
    centerX: (minX + maxX) / 2,
    minY,
    maxY,
    height: maxY - minY,
    centerY: (minY + maxY) / 2,
    components,
  };
}

function resolveImageRasterGeometry(
  rasterBands,
  textBoundaryCenter,
  wordGeometry = {},
) {
  const requiredGeometry = [
    "displayTop",
    "displayHeight",
    "numberedTop",
    "numberedHeight",
  ];
  if (
    requiredGeometry.some(
      (key) => !Number.isFinite(wordGeometry[key]) || wordGeometry[key] < 0,
    )
  ) {
    throw new Error(
      `Image raster geometry is missing Word formula bounds: ${JSON.stringify(wordGeometry)}`,
    );
  }

  const componentsInVerticalBox = (top, height) => {
    const bottom = top + height;
    const tolerance = 1.5;
    return rasterBands.flatMap((band) =>
      band.maxY >= top - tolerance && band.minY <= bottom + tolerance
        ? band.components.map((component) => ({ band, component }))
        : [],
    );
  };
  const centeredMarker = (entries, label, excludedComponent = null) => {
    const candidates = entries
      .filter(({ component }) => component !== excludedComponent)
      .map((entry) => ({
        ...entry,
        centerError: Math.abs(entry.component.centerX - textBoundaryCenter),
      }))
      .filter(({ centerError }) => centerError <= 8)
      .sort(
        (left, right) =>
          left.centerError - right.centerError ||
          right.component.width - left.component.width,
      );
    const best = candidates[0];
    if (!best) {
      throw new Error(
        `Unable to locate ${label} image formula center marker: ${JSON.stringify({ rasterBands, wordGeometry, textBoundaryCenter })}`,
      );
    }
    return rasterBounds(best.band, [best.component]);
  };

  const numberedEntries = componentsInVerticalBox(
    wordGeometry.numberedTop,
    wordGeometry.numberedHeight,
  );
  const numberCandidates = numberedEntries
    .filter(
      ({ component }) => component.minX > textBoundaryCenter + 100,
    )
    .sort(
      (left, right) => right.component.centerX - left.component.centerX,
    );
  const numberEntry = numberCandidates[0];
  if (!numberEntry) {
    throw new Error(
      `Unable to locate numbered image formula raster number: ${JSON.stringify({ rasterBands, wordGeometry, textBoundaryCenter })}`,
    );
  }

  const unnumberedEntries = componentsInVerticalBox(
    wordGeometry.displayTop,
    wordGeometry.displayHeight,
  );
  const numberedFormulaEntries = numberedEntries.filter(
    ({ component }) => component !== numberEntry.component,
  );
  const unnumbered = centeredMarker(
    unnumberedEntries,
    "unnumbered",
  );
  const numbered = centeredMarker(
    numberedFormulaEntries,
    "numbered",
  );
  return {
    unnumbered,
    numbered,
    unnumberedInk: rasterEntryBounds(unnumberedEntries),
    numberedInk: rasterEntryBounds(numberedFormulaEntries),
    equationNumber: {
      ...numberEntry.component,
      minY: numberEntry.band.minY,
      maxY: numberEntry.band.maxY,
      height: numberEntry.band.height,
      centerY: numberEntry.band.centerY,
    },
  };
}

function manifestText(entries) {
  const seen = new Set();
  return entries
    .map(([key, value]) => {
      if (!/^[A-Za-z0-9]+$/.test(key) || seen.has(key)) {
        throw new Error(`Invalid integration manifest key: ${key}`);
      }
      seen.add(key);
      const text = String(value);
      if (/[\r\n\0]/.test(text)) {
        throw new Error(`Invalid integration manifest value for ${key}`);
      }
      return `${key}=${text}`;
    })
    .join("\n") + "\n";
}

function parseRegressionReport(reportText) {
  const lines = reportText.trim().split(/\r?\n/);
  if (lines[0] !== "PASS") {
    throw new Error(`Word formula regression failed: ${reportText}`);
  }
  return Object.fromEntries(
    lines.slice(1).map((line) => {
      const separator = line.indexOf("=");
      if (separator <= 0) {
        throw new Error(`Invalid Word formula regression line: ${line}`);
      }
      return [line.slice(0, separator), line.slice(separator + 1)];
    }),
  );
}

function numericReportValue(report, key) {
  const value = Number(report[key]);
  if (!Number.isFinite(value)) {
    throw new Error(`Word formula regression omitted ${key}: ${JSON.stringify(report)}`);
  }
  return value;
}

function runFormulaRegressionReport(testDocumentName, formulas) {
  const formulaCount = formulas.length;
  const displayFormulaCount = formulas.filter(
    (formula) => formula.displayMode === "block",
  ).length;
  const alignedFormulaCountExpected = formulas.filter((formula) =>
    ["align", "align-star"].includes(formula.codeFormat),
  ).length;
  rmSync(formulaRegressionStatusPath, { force: true });
  runAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(testDocumentName)}`,
    'run VB macro macro name "VisualTeX_RunDocumentImportFormulaRegression"',
    "end tell",
  ], 60_000);
  if (!existsSync(formulaRegressionStatusPath)) {
    throw new Error("Word did not write the document-import formula regression report");
  }
  const report = parseRegressionReport(
    readFileSync(formulaRegressionStatusPath, "utf8"),
  );
  if (report.revision !== "word-double-click-routing-20260730-r66") {
    throw new Error(`Word loaded the wrong VisualTeX source revision: ${report.revision}`);
  }

  const documentMathCount = numericReportValue(report, "documentMathCount");
  const nativeFormulaCount = numericReportValue(report, "nativeFormulaCount");
  const nativeDisplayCount = numericReportValue(report, "nativeDisplayCount");
  const invalidNativeRangeCount = numericReportValue(
    report,
    "invalidNativeRangeCount",
  );
  const emptyMathCount = numericReportValue(report, "emptyMathCount");
  const alignedFormulaCount = numericReportValue(report, "alignedFormulaCount");
  const imageFormulaCount = numericReportValue(report, "imageFormulaCount");
  const imageDisplayCount = numericReportValue(report, "imageDisplayCount");
  const imageMacroButtonCount = numericReportValue(
    report,
    "imageMacroButtonCount",
  );
  const invalidImageMacroButtonCount = numericReportValue(
    report,
    "invalidImageMacroButtonCount",
  );
  const imageFormulaIds = (report.imageFormulaIds ?? "")
    .split(",")
    .filter(Boolean);
  const maximumImageSpaceBefore = numericReportValue(
    report,
    "maximumImageSpaceBefore",
  );
  const maximumImageSpaceAfter = numericReportValue(
    report,
    "maximumImageSpaceAfter",
  );

  if (outputKind === "omml") {
    if (
      documentMathCount !== formulaCount ||
      nativeFormulaCount !== formulaCount ||
      nativeDisplayCount !== displayFormulaCount ||
      invalidNativeRangeCount !== 0 ||
      emptyMathCount !== 0 ||
      alignedFormulaCount !== alignedFormulaCountExpected ||
      imageFormulaCount !== 0 ||
      imageMacroButtonCount !== 0 ||
      invalidImageMacroButtonCount !== 0
    ) {
      throw new Error(
        `Word OMML formula structure regression failed: ${JSON.stringify(report)}`,
      );
    }
    for (const [key, expected] of [
      // Numbered native equations retain their existing zero-spaced table
      // layout; unnumbered native equations inherit the configured Normal
      // style spacing. Image-only normalization must not change either case.
      ["minimumNativeSpaceBefore", 0],
      ["maximumNativeSpaceBefore", 6],
      ["minimumNativeSpaceAfter", 0],
      ["maximumNativeSpaceAfter", 9],
    ]) {
      if (Math.abs(numericReportValue(report, key) - expected) > 0.05) {
        throw new Error(
          `OMML paragraph spacing changed for ${key}: ${JSON.stringify(report)}`,
        );
      }
    }
  } else if (
    documentMathCount !== 0 ||
    nativeFormulaCount !== 0 ||
    nativeDisplayCount !== 0 ||
    invalidNativeRangeCount !== 0 ||
    emptyMathCount !== 0 ||
    alignedFormulaCount !== 0 ||
    imageFormulaCount !== formulaCount ||
    imageDisplayCount !== displayFormulaCount ||
    imageMacroButtonCount !== 0 ||
    invalidImageMacroButtonCount !== 0 ||
    JSON.stringify(imageFormulaIds) !==
      JSON.stringify(formulas.map((formula) => formula.formulaId)) ||
    maximumImageSpaceBefore > 0.01 ||
    maximumImageSpaceAfter > 0.01
  ) {
    throw new Error(
      `Word image formula structure regression failed: ${JSON.stringify(report)}`,
    );
  }
  return report;
}

function compactFormulaId(id) {
  return id.replaceAll("-", "");
}

function nativeBookmark(id) {
  return `VT_F_${compactFormulaId(id)}`;
}

function currentSessionIds() {
  return new Set(
    existsSync(sessionsRoot)
      ? readdirSync(sessionsRoot, { withFileTypes: true })
          .filter((entry) => entry.isDirectory())
          .map((entry) => entry.name)
      : [],
  );
}

function inspectWordFormulaContainers(testDocumentName, formulas, stage) {
  const report = runFormulaRegressionReport(testDocumentName, formulas);
  if (outputKind === "omml") {
    return {
      stage,
      inlineShapeCount: 0,
      macroButtonCount: 0,
      shapes: [],
    };
  }
  const inspection = runAppleScript([
    'tell application "Microsoft Word"',
    `set documentObject to document ${JSON.stringify(testDocumentName)}`,
    "set unitSeparator to ASCII character 31",
    "set recordSeparator to ASCII character 30",
    "set macroButtonCount to 0",
    "set documentFieldCount to count fields of documentObject",
    "repeat with fieldIndex from 1 to documentFieldCount",
    "set candidateField to field fieldIndex of documentObject",
    "if field type of candidateField is field macro button then set macroButtonCount to macroButtonCount + 1",
    "end repeat",
    "set reportText to (macroButtonCount as text) & unitSeparator & ((count of inline shapes of documentObject) as text)",
    "repeat with shapeIndex from 1 to count of inline shapes of documentObject",
    "set formulaShape to inline shape shapeIndex of documentObject",
    "set shapeRange to text object of formulaShape",
    "set shapeStart to start of content of shapeRange",
    "set shapeEnd to end of content of shapeRange",
    "set formulaParagraph to paragraph 1 of (create range documentObject start shapeStart end shapeStart)",
    "set paragraphRange to text object of formulaParagraph",
    "set paragraphStart to start of content of paragraphRange",
    "set paragraphEnd to end of content of paragraphRange",
    "set paragraphText to content of paragraphRange",
    "set paragraphFieldCount to count of fields of paragraphRange",
    "set metadataText to alternative text of formulaShape",
    "set reportText to reportText & recordSeparator & (shapeIndex as text) & unitSeparator & (shapeStart as text) & unitSeparator & (shapeEnd as text) & unitSeparator & (paragraphStart as text) & unitSeparator & (paragraphEnd as text) & unitSeparator & paragraphText & unitSeparator & (paragraphFieldCount as text) & unitSeparator & metadataText",
    "end repeat",
    "return reportText",
    "end tell",
  ]);
  const [summary, ...recordTexts] = inspection.split("\x1e");
  const [macroButtonCount, inlineShapeCount] = summary
    .split("\x1f")
    .map(Number);
  if (
    !Number.isInteger(macroButtonCount) ||
    macroButtonCount < 0 ||
    !Number.isInteger(inlineShapeCount) ||
    inlineShapeCount < 0
  ) {
    throw new Error(
      `${stage} returned an invalid Word field summary: ${JSON.stringify(inspection)}`,
    );
  }

  if (outputKind === "omml") {
    if (inlineShapeCount !== 0 || macroButtonCount !== 0 || recordTexts.length) {
      throw new Error(
        `${stage} OMML document contains image/MacroButton objects: ${JSON.stringify({
          inlineShapeCount,
          macroButtonCount,
          recordTexts,
        })}`,
      );
    }
    return { stage, inlineShapeCount, macroButtonCount, shapes: [] };
  }

  if (
    inlineShapeCount !== formulas.length ||
    macroButtonCount !== 0 ||
    recordTexts.length !== formulas.length
  ) {
    throw new Error(
      `${stage} did not retain plain field-free VisualTeX images: ${JSON.stringify({
        expected: formulas.length,
        inlineShapeCount,
        macroButtonCount,
        recordCount: recordTexts.length,
      })}`,
    );
  }
  const shapes = recordTexts.map((recordText, index) => {
    const [
      shapeIndexText,
      shapeStartText,
      shapeEndText,
      paragraphStartText,
      paragraphEndText,
      paragraphText,
      paragraphFieldCountText,
      encodedMetadata,
    ] = recordText.split("\x1f");
    const record = {
      shapeIndex: Number(shapeIndexText),
      shapeStart: Number(shapeStartText),
      shapeEnd: Number(shapeEndText),
      paragraphStart: Number(paragraphStartText),
      paragraphEnd: Number(paragraphEndText),
      paragraphText,
      paragraphFieldCount: Number(paragraphFieldCountText),
    };
    const expected = formulas[index];
    const metadata = decodeFormulaMetadata(encodedMetadata ?? "");
    if (
      record.shapeIndex !== index + 1 ||
      !Number.isInteger(record.shapeStart) ||
      !Number.isInteger(record.shapeEnd) ||
      !Number.isInteger(record.paragraphStart) ||
      !Number.isInteger(record.paragraphEnd) ||
      record.shapeStart >= record.shapeEnd ||
      record.paragraphStart > record.shapeStart ||
      record.paragraphEnd < record.shapeEnd ||
      record.paragraphFieldCount !== 0 ||
      metadata?.formulaId !== expected.formulaId
    ) {
      throw new Error(
        `${stage} image ${index + 1} is not one plain field-free ` +
          `VisualTeX InlineShape: ${JSON.stringify({ record, metadata })}`,
      );
    }
    return { ...record, formulaId: metadata.formulaId };
  });

  const paragraphGroups = new Map();
  for (const shape of shapes) {
    const key = `${shape.paragraphStart}:${shape.paragraphEnd}`;
    const group = paragraphGroups.get(key) ?? [];
    group.push(shape);
    paragraphGroups.set(key, group);
  }
  const normalizedParagraphText = (value, imageCount) => {
    let text = String(value ?? "").replace(
      /[\u0001\u0007\u0015\t\r\n\u00a0\u200b\u2060 ]/g,
      "",
    );
    // Word AppleScript exposes each plain InlineShape as one slash in Range.Text.
    // Remove exactly the known image-object placeholders, never arbitrary X/text.
    for (let index = 0; index < imageCount; index += 1) {
      text = text.replace("/", "");
    }
    return text;
  };

  const structuredShapes = shapes.map((shape, index) => {
    const formula = formulas[index];
    const key = `${shape.paragraphStart}:${shape.paragraphEnd}`;
    const paragraphShapes = paragraphGroups.get(key) ?? [];
    const visibleText = normalizedParagraphText(
      shape.paragraphText,
      paragraphShapes.length,
    );
    const paragraphModes = paragraphShapes.map(
      (paragraphShape) => formulas[paragraphShape.shapeIndex - 1]?.displayMode,
    );

    if (formula.displayMode === "block") {
      const validNumberText = /^\([^()]+\)$/.test(visibleText);
      const validDisplayStructure =
        paragraphShapes.length === 1 &&
        paragraphModes.every((mode) => mode === "block") &&
        (formula.numbered
          ? validNumberText
          : visibleText === "");
      if (!validDisplayStructure) {
        throw new Error(
          `${stage} display image ${index + 1} is not isolated in its own ` +
            `Word paragraph: ${JSON.stringify({ shape, paragraphShapes, visibleText })}`,
        );
      }
      return {
        ...shape,
        layoutStructure: formula.numbered
          ? "numbered-display-paragraph"
          : "dedicated-display-paragraph",
        paragraphFormulaCount: paragraphShapes.length,
        visibleParagraphText: visibleText,
      };
    }

    const validInlineStructure =
      visibleText.length > 0 &&
      paragraphModes.every((mode) => mode === "inline");
    if (!validInlineStructure) {
      throw new Error(
        `${stage} inline image ${index + 1} is not embedded in a body-text ` +
          `paragraph: ${JSON.stringify({ shape, paragraphShapes, visibleText })}`,
      );
    }
    return {
      ...shape,
      layoutStructure: "inline-text-flow",
      paragraphFormulaCount: paragraphShapes.length,
      visibleParagraphText: visibleText,
    };
  });
  return {
    stage,
    inlineShapeCount,
    macroButtonCount,
    shapes: structuredShapes,
  };
}

function saveAndReopenWordDocument(testDocumentName) {
  rmSync(reopenedDocumentPath, { force: true });
  try {
    runAppleScript([
      'tell application "Microsoft Word"',
      `set documentObject to document ${JSON.stringify(testDocumentName)}`,
      `save as documentObject file name ${JSON.stringify(reopenedDocumentPath)}`,
      "end tell",
    ], 90_000);
  } catch (error) {
    // Word for Mac can successfully complete SaveAs and then return -128 after
    // the old AppleScript document wrapper becomes invalid. The on-disk DOCX
    // is the source of truth; never retry SaveAs or keep using that wrapper.
    if (!existsSync(reopenedDocumentPath)) throw error;
  }
  runAppleScript(['tell application "Microsoft Word" to quit saving no'], 30_000);
  spawnSync("/bin/sleep", ["2"], { encoding: "utf8" });
  return runAppleScript([
    'tell application "Microsoft Word"',
    `open file name ${JSON.stringify(reopenedDocumentPath)}`,
    "set reopenedDocument to active document",
    "activate object reopenedDocument",
    "activate",
    "return name of reopenedDocument",
    "end tell",
  ], 90_000);
}

function validateFormulaEditSession(
  sessionId,
  formula,
  expectedCodeFormat,
  expectedLines,
) {
  const requestPath = join(sessionsRoot, sessionId, "request.json");
  const request = JSON.parse(readFileSync(requestPath, "utf8"));
  if (
    request.mode !== "edit" ||
    request.host !== "word" ||
    request.formulaId !== formula.formulaId ||
    request.displayMode !== formula.displayMode ||
    Boolean(request.numbered) !== formula.numbered
  ) {
    throw new Error(
      `Unexpected Word formula edit request: ${JSON.stringify(request)}`,
    );
  }
  const metadata = decodeFormulaMetadata(request.encodedMetadata ?? "");
  if (!metadata || metadata.formulaId !== formula.formulaId) {
    throw new Error(`Word edit Session lost formula metadata for ${formula.formulaId}`);
  }
  const normalized = normalizeFormulaEditorDocument(
    metadata.lines,
    metadata.codeFormat,
  );
  if (
    normalized.codeFormat !== expectedCodeFormat ||
    JSON.stringify(normalized.lines.map((line) => line.latex)) !==
      JSON.stringify(expectedLines)
  ) {
    throw new Error(
      `Word edit Session did not restore ${expectedCodeFormat}: ${JSON.stringify({
        metadata,
        normalized,
      })}`,
    );
  }
  if (normalized.lines[0]?.id !== formula.metadataLineId) {
    throw new Error("Word edit normalization did not preserve the imported formula line UUID");
  }
  if (
    Math.abs((request.fontSizePt ?? 0) - formula.fontSizePt) > 0.1
  ) {
    throw new Error(
      `Word edit Session lost formula font size: ${request.fontSizePt}`,
    );
  }
  return { request, metadata, normalized };
}

async function waitForNewSession(before, timeoutMs = 12_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    if (existsSync(sessionsRoot)) {
      const ready = [];
      for (const entry of readdirSync(sessionsRoot, { withFileTypes: true })) {
        if (!entry.isDirectory() || before.has(entry.name)) continue;
        const requestPath = join(sessionsRoot, entry.name, "request.json");
        if (!existsSync(requestPath)) continue;
        try {
          const request = JSON.parse(readFileSync(requestPath, "utf8"));
          if (
            request.operation === "documentImport" &&
            request.sessionId === entry.name &&
            request.host === "word"
          ) {
            ready.push({
              sessionId: entry.name,
              modifiedAt: statSync(requestPath).mtimeMs,
            });
          }
        } catch {
          // The Session directory can become visible before its atomic
          // request.json rename. Wait for a complete, validated request.
        }
      }
      if (ready.length > 0) {
        ready.sort((left, right) => right.modifiedAt - left.modifiedAt);
        return ready[0].sessionId;
      }
    }
    await sleep(100);
  }
  throw new Error("Word did not create a VisualTeX document import Session");
}

async function waitForWordCreateSession(before, timeoutMs = 12_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    for (const sessionId of currentSessionIds()) {
      if (before.has(sessionId)) continue;
      const requestPath = join(sessionsRoot, sessionId, "request.json");
      if (!existsSync(requestPath)) continue;
      try {
        const request = JSON.parse(readFileSync(requestPath, "utf8"));
        if (
          request.mode === "create" &&
          request.host === "word" &&
          request.sessionId === sessionId
        ) {
          return sessionId;
        }
      } catch {
        // The request may still be completing its atomic rename.
      }
    }
    await sleep(100);
  }
  throw new Error("Word did not create a VisualTeX formula creation Session");
}

async function waitForFormulaEditSession(before, formulaId, timeoutMs = 12_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    for (const sessionId of currentSessionIds()) {
      if (before.has(sessionId)) continue;
      const requestPath = join(sessionsRoot, sessionId, "request.json");
      if (!existsSync(requestPath)) continue;
      try {
        const request = JSON.parse(readFileSync(requestPath, "utf8"));
        if (request.mode === "edit" && request.formulaId === formulaId) {
          return sessionId;
        }
      } catch {
        // The request may still be in the middle of its atomic write.
      }
    }
    await sleep(100);
  }
  throw new Error(`Word did not create an edit Session for ${formulaId}`);
}

function editorPerformanceRecords(sessionId) {
  const performancePath = join(
    sessionsRoot,
    sessionId,
    editorPerformanceFileName,
  );
  if (!existsSync(performancePath)) return [];
  return readFileSync(performancePath, "utf8")
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line, index) => {
      try {
        return JSON.parse(line);
      } catch (error) {
        throw new Error(
          `Invalid editor performance record ${index + 1} for ${sessionId}: ${error}`,
        );
      }
    });
}

function validatedPhysicalEditorReadiness(sessionId, formulaId, marker, records) {
  if (
    marker.schema !== editorReadySchema ||
    marker.sessionId !== sessionId ||
    marker.host !== "word" ||
    !Number.isSafeInteger(marker.generation) ||
    marker.generation <= 0
  ) {
    throw new Error(
      `The physical edit wrote an invalid editor-ready marker: ${JSON.stringify(marker)}`,
    );
  }
  for (const key of [
    "epochMs",
    "urlReceivedEpochMs",
    "frontendEpochMs",
  ]) {
    if (!Number.isSafeInteger(marker[key]) || marker[key] <= 0) {
      throw new Error(`The editor-ready marker has an invalid ${key}`);
    }
  }
  for (const key of [
    "hydrateMs",
    "editorMountedMs",
    "contentReadyMs",
    "showFocusMs",
  ]) {
    if (!Number.isFinite(marker[key]) || marker[key] < 0) {
      throw new Error(`The editor-ready marker has an invalid ${key}`);
    }
  }
  if (
    marker.hydrateMs > marker.editorMountedMs ||
    marker.editorMountedMs > marker.contentReadyMs ||
    marker.contentReadyMs > marker.showFocusMs + 10
  ) {
    throw new Error(
      `The physical editor readiness stages are out of order: ${JSON.stringify(marker)}`,
    );
  }
  if (
    marker.urlReceivedEpochMs > marker.frontendEpochMs + 100 ||
    marker.frontendEpochMs > marker.epochMs + 100
  ) {
    throw new Error(
      `The physical editor readiness epochs are out of order: ${JSON.stringify(marker)}`,
    );
  }
  if (
    marker.contentReadyMs > warmEditorReadyLimitMs ||
    marker.showFocusMs > warmEditorReadyLimitMs
  ) {
    throw new Error(
      `The resident Word editor missed the ${warmEditorReadyLimitMs} ms warm target: ` +
        JSON.stringify(marker),
    );
  }

  const requiredStages = [
    "url-received",
    "request-read",
    "request-imported",
    "window-reused",
    "activation-event-sent",
    "frontend-hydrated",
    "frontend-editor-mounted",
    "frontend-content-ready",
    "window-show-focus",
  ];
  for (const record of records) {
    if (
      record.schema !== editorPerformanceSchema ||
      record.sessionId !== sessionId ||
      record.host !== "word" ||
      !Number.isFinite(record.elapsedMs) ||
      record.elapsedMs < 0
    ) {
      throw new Error(
        `The physical edit wrote an invalid performance record: ${JSON.stringify(record)}`,
      );
    }
  }
  const byStage = Object.fromEntries(
    requiredStages.map((stage) => [
      stage,
      records.filter((record) => record.stage === stage),
    ]),
  );
  for (const stage of requiredStages) {
    if (byStage[stage].length !== 1) {
      throw new Error(
        `The physical edit did not record exactly one ${stage} stage: ${JSON.stringify(records)}`,
      );
    }
  }
  if (records.some((record) => record.stage === "window-created")) {
    throw new Error(
      "The physical Word edit created a new WebView instead of reusing the resident editor",
    );
  }
  for (const stage of [
    "window-reused",
    "activation-event-sent",
    "frontend-hydrated",
    "frontend-editor-mounted",
    "frontend-content-ready",
    "window-show-focus",
  ]) {
    if (byStage[stage][0].generation !== marker.generation) {
      throw new Error(
        `The ${stage} performance record belongs to a stale editor generation`,
      );
    }
  }
  const stageElapsed = Object.fromEntries(
    requiredStages.map((stage) => [stage, byStage[stage][0].elapsedMs]),
  );
  const frontendOriginMs =
    stageElapsed["frontend-content-ready"] - marker.contentReadyMs;
  if (!Number.isFinite(frontendOriginMs) || frontendOriginMs < -1) {
    throw new Error(
      `The physical editor frontend timing origin is invalid: ${JSON.stringify({ frontendOriginMs, stageElapsed, marker })}`,
    );
  }
  for (const [stage, markerKey] of [
    ["frontend-hydrated", "hydrateMs"],
    ["frontend-editor-mounted", "editorMountedMs"],
    ["frontend-content-ready", "contentReadyMs"],
  ]) {
    const frontendRelativeMs = stageElapsed[stage] - frontendOriginMs;
    if (Math.abs(frontendRelativeMs - marker[markerKey]) > 1) {
      throw new Error(
        `The ${stage} timing disagrees with editor-ready.${markerKey}: ${JSON.stringify({ frontendRelativeMs, stageElapsed: stageElapsed[stage], frontendOriginMs, markerValue: marker[markerKey] })}`,
      );
    }
  }
  if (Math.abs(stageElapsed["window-show-focus"] - marker.showFocusMs) > 1) {
    throw new Error(
      "The window-show-focus timing disagrees with editor-ready.showFocusMs",
    );
  }
  const backendOrder = [
    "url-received",
    "request-read",
    "request-imported",
    "window-reused",
    "activation-event-sent",
    "window-show-focus",
  ];
  for (let index = 1; index < backendOrder.length; index += 1) {
    const previous = backendOrder[index - 1];
    const current = backendOrder[index];
    if (stageElapsed[current] + 1 < stageElapsed[previous]) {
      throw new Error(
        `The physical editor backend stages are out of order: ${JSON.stringify(stageElapsed)}`,
      );
    }
  }

  const requestPath = join(sessionsRoot, sessionId, "request.json");
  const requestWrittenEpochMs = statSync(requestPath).mtimeMs;
  const requestToUrlMs = marker.urlReceivedEpochMs - requestWrittenEpochMs;
  const requestToReadyMs = marker.epochMs - requestWrittenEpochMs;
  const urlToReadyEpochMs = marker.epochMs - marker.urlReceivedEpochMs;
  if (
    requestToUrlMs < -250 ||
    requestToUrlMs > 2_000 ||
    requestToReadyMs < -250 ||
    requestToReadyMs > 1_500 ||
    Math.abs(urlToReadyEpochMs - marker.showFocusMs) > 250
  ) {
    throw new Error(
      `The physical editor request/URL/readiness timing is invalid: ${JSON.stringify({
        requestWrittenEpochMs,
        requestToUrlMs,
        requestToReadyMs,
        urlToReadyEpochMs,
        marker,
      })}`,
    );
  }
  return {
    schema: editorReadySchema,
    sessionId,
    formulaId,
    generation: marker.generation,
    requestWrittenEpochMs,
    requestToUrlMs,
    requestToReadyMs,
    urlToReadyEpochMs,
    ...Object.fromEntries(
      [
        "urlReceivedEpochMs",
        "frontendEpochMs",
        "epochMs",
        "hydrateMs",
        "editorMountedMs",
        "contentReadyMs",
        "showFocusMs",
      ].map((key) => [key, marker[key]]),
    ),
    stages: stageElapsed,
  };
}

async function waitForPhysicalEditorReadiness(
  sessionId,
  formulaId,
  timeoutMs = 30_000,
) {
  const readyPath = join(sessionsRoot, sessionId, editorReadyFileName);
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    if (existsSync(readyPath)) {
      const marker = JSON.parse(readFileSync(readyPath, "utf8"));
      const records = editorPerformanceRecords(sessionId);
      const requiredStageNames = new Set(records.map((record) => record.stage));
      if (
        [
          "url-received",
          "request-read",
          "request-imported",
          "window-reused",
          "activation-event-sent",
          "frontend-hydrated",
          "frontend-editor-mounted",
          "frontend-content-ready",
          "window-show-focus",
        ].every((stage) => requiredStageNames.has(stage))
      ) {
        await sleep(100);
        return validatedPhysicalEditorReadiness(
          sessionId,
          formulaId,
          marker,
          editorPerformanceRecords(sessionId),
        );
      }
    }
    await sleep(50);
  }
  throw new Error(
    `VisualTeX did not write ${editorReadyFileName} and complete performance stages for ${formulaId}`,
  );
}

async function assertSinglePhysicalEditSession(before, sessionId, formulaId) {
  await sleep(300);
  const matchingSessionIds = [];
  for (const candidateSessionId of currentSessionIds()) {
    if (before.has(candidateSessionId)) continue;
    const requestPath = join(
      sessionsRoot,
      candidateSessionId,
      "request.json",
    );
    if (!existsSync(requestPath)) continue;
    try {
      const request = JSON.parse(readFileSync(requestPath, "utf8"));
      if (request.mode === "edit" && request.formulaId === formulaId) {
        matchingSessionIds.push(candidateSessionId);
      }
    } catch {
      // A different Session may still be finishing an atomic write.
    }
  }
  if (
    matchingSessionIds.length !== 1 ||
    matchingSessionIds[0] !== sessionId
  ) {
    throw new Error(
      `One physical double-click must create exactly one edit Session: ${JSON.stringify({
        sessionId,
        formulaId,
        matchingSessionIds,
      })}`,
    );
  }
}

function formulaItem({
  formulaId,
  latex,
  metadataLatex = latex,
  expectedCodeFormat,
  displayMode,
  numbered,
  fontSizePt,
  artifactDirectory,
}) {
  const normalized = normalizeFormulaEditorDocument(
    [{ id: crypto.randomUUID(), latex: metadataLatex }],
    "raw",
  );
  if (expectedCodeFormat && normalized.codeFormat !== expectedCodeFormat) {
    throw new Error(
      `Formula fixture did not normalize as ${expectedCodeFormat}: ${JSON.stringify(normalized)}`,
    );
  }
  const canonicalLatex = serializeFormulaEditorDocument(normalized);
  const normalizedLines = normalized.lines.map((line) => line.latex);
  const omml = ommlForFormula(normalizedLines, normalized.codeFormat);
  if (["align", "align-star"].includes(normalized.codeFormat)) {
    if (
      (omml.match(/<m:oMath\b/g) ?? []).length !== 1 ||
      (omml.match(/<m:eqArr>/g) ?? []).length !== 1 ||
      (omml.match(/&amp;/g) ?? []).length !== normalizedLines.length
    ) {
      throw new Error(
        `Aligned fixture is not one relationship-aligned OMML equation array: ${omml}`,
      );
    }
  }
  const nativePath = join(nativeRoot, `${formulaId}.docx`);
  writeFileSync(nativePath, minimalDocxBytes(omml));

  let imagePath = "";
  let vectorDocumentPath = "";
  let fallbackImagePath = "";
  let widthPoints = fontSizePt;
  let heightPoints = Math.max(18, fontSizePt * 1.8);
  let baseline = 0;
  let referenceWidthPt = referenceFontSizePt;
  let referenceHeightPt = referenceFontSizePt;
  let referenceBaselinePt = 0;
  let renderWidthPx;
  let renderHeightPx;
  let svgAlignment;

  if (outputKind === "image") {
    const svg = latexToSvg(canonicalLatex, {
      displayMode: displayMode === "block",
      fontSizePt: referenceFontSizePt,
      paddingPx: displayMode === "inline" ? 1 : wordDisplayPaddingPx,
      background: "transparent",
    });
    const geometry = calculateImageGeometry(svg, fontSizePt);
    ({
      widthPoints,
      heightPoints,
      baseline,
      referenceWidthPt,
      referenceHeightPt,
      referenceBaselinePt,
    } = geometry);
    renderWidthPx = svg.width;
    renderHeightPx = svg.height;
    if (["align", "align-star"].includes(normalized.codeFormat)) {
      svgAlignment = assertAlignedSvg(
        svg.svg,
        normalized.lines.length,
        `Initial ${normalized.codeFormat} image`,
      );
    }
    const stem = `document-formula-${compactFormulaId(formulaId)}`;
    imagePath = join(artifactDirectory, `${stem}.svg`);
    fallbackImagePath = join(artifactDirectory, `${stem}.png`);
    vectorDocumentPath = join(artifactDirectory, `${stem}-svg.docx`);
    writeFileSync(imagePath, svg.svg, { mode: 0o600 });
    writeFileSync(fallbackImagePath, transparentPng, { mode: 0o600 });
    writeFileSync(
      vectorDocumentPath,
      wordSvgDocxBytes(svg.svg, transparentPng, widthPoints, heightPoints),
      { mode: 0o600 },
    );
  }

  const metadata = createFormulaMetadata({
    formulaId,
    title: displayMode === "inline" ? "Integration inline formula" : "Integration display formula",
    lines: normalized.lines,
    codeFormat: normalized.codeFormat,
    sourceLatex: canonicalLatex,
    displayMode,
    numbered,
    fontSizePt,
    referenceWidthPt,
    referenceHeightPt,
    referenceBaselinePt,
    renderWidthPx,
    renderHeightPx,
  });
  return {
    formulaId,
    latex: canonicalLatex,
    metadataLatex: canonicalLatex,
    metadataLineId: normalized.lines[0].id,
    metadataLines: normalizedLines,
    codeFormat: normalized.codeFormat,
    pdfToken: normalizedLines.join(""),
    displayMode,
    numbered,
    fontSizePt,
    metadata: encodeFormulaMetadata(metadata),
    ommlBase64: Buffer.from(omml, "utf8").toString("base64url"),
    nativePath,
    imagePath,
    vectorDocumentPath,
    fallbackImagePath,
    widthPoints,
    heightPoints,
    baseline,
    referenceWidthPt,
    referenceHeightPt,
    referenceBaselinePt,
    svgAlignment,
  };
}

function editedImageFormulaArtifacts(
  formula,
  editSession,
  updatedLineLatex,
  artifactDirectory,
) {
  const updatedLines = editSession.normalized.lines.map((line, index) => ({
    ...line,
    latex: updatedLineLatex[index] ?? line.latex,
  }));
  if (updatedLines.length !== updatedLineLatex.length) {
    throw new Error(
      `Edited ${formula.codeFormat} fixture changed its row count unexpectedly`,
    );
  }
  const rendered = renderOfficeFormulaArtifacts({
    lines: updatedLines,
    codeFormat: editSession.normalized.codeFormat,
    displayMode: formula.displayMode,
    host: "word",
    includeWordOmml: false,
  });
  const svgAlignment = assertAlignedSvg(
    rendered.svg.svg,
    rendered.lines.length,
    `Edited ${rendered.codeFormat} image`,
  );
  const geometry = calculateImageGeometry(rendered.svg, formula.fontSizePt);
  const stem = `edited-${compactFormulaId(formula.formulaId)}`;
  const imagePath = join(artifactDirectory, `${stem}.svg`);
  const fallbackImagePath = join(artifactDirectory, `${stem}.png`);
  const vectorDocumentPath = join(artifactDirectory, `${stem}-svg.docx`);
  writeFileSync(imagePath, rendered.svg.svg, { mode: 0o600 });
  writeFileSync(fallbackImagePath, transparentPng, { mode: 0o600 });
  writeFileSync(
    vectorDocumentPath,
    wordSvgDocxBytes(
      rendered.svg.svg,
      transparentPng,
      geometry.widthPoints,
      geometry.heightPoints,
    ),
    { mode: 0o600 },
  );

  const omml = ommlForFormula(
    rendered.lines.map((line) => line.latex),
    rendered.codeFormat,
  );
  writeFileSync(formula.nativePath, minimalDocxBytes(omml), { mode: 0o600 });
  const metadata = createFormulaMetadata({
    formulaId: formula.formulaId,
    title: editSession.metadata.title,
    lines: rendered.lines,
    codeFormat: rendered.codeFormat,
    sourceLatex: rendered.canonicalLatex,
    displayMode: formula.displayMode,
    numbered: formula.numbered,
    fontSizePt: formula.fontSizePt,
    referenceWidthPt: geometry.referenceWidthPt,
    referenceHeightPt: geometry.referenceHeightPt,
    referenceBaselinePt: geometry.referenceBaselinePt,
    renderWidthPx: rendered.svg.width,
    renderHeightPx: rendered.svg.height,
    original: editSession.metadata,
  });
  return {
    lines: rendered.lines,
    codeFormat: rendered.codeFormat,
    canonicalLatex: rendered.canonicalLatex,
    metadata: encodeFormulaMetadata(metadata),
    ommlBase64: Buffer.from(omml, "utf8").toString("base64url"),
    imagePath,
    fallbackImagePath,
    vectorDocumentPath,
    ...geometry,
    renderWidthPx: rendered.svg.width,
    renderHeightPx: rendered.svg.height,
    svgAlignment,
  };
}

function commitEditedImageFormula(
  testDocumentName,
  sessionId,
  formula,
  editSession,
  updatedLineLatex,
) {
  const sessionDirectory = join(sessionsRoot, sessionId);
  const artifacts = editedImageFormulaArtifacts(
    formula,
    editSession,
    updatedLineLatex,
    sessionDirectory,
  );
  const request = editSession.request;
  const dispatch = manifestText([
    ["protocolVersion", "1"],
    ["sessionId", sessionId],
    ["action", "commit"],
    ["host", "word"],
    ["mode", "edit"],
    ["formulaId", formula.formulaId],
    ["displayMode", formula.displayMode],
    ["numbered", formula.numbered ? "1" : "0"],
    ["nativeEquation", "0"],
    ["imagePath", artifacts.imagePath],
    ["vectorDocumentPath", artifacts.vectorDocumentPath],
    ["fallbackImagePath", artifacts.fallbackImagePath],
    ["metadata", artifacts.metadata],
    ["latexBase64", base64Url(artifacts.canonicalLatex)],
    ["ommlBase64", artifacts.ommlBase64],
    ["nativeDocumentPath", formula.nativePath],
    ["pendingMarker", request.pendingMarker ?? ""],
    [
      "sourceMarker",
      request.sourceObjectId ?? request.encodedMetadata ?? "",
    ],
    ["sourceDocumentId", request.sourceDocumentId ?? ""],
    ["widthPoints", artifacts.widthPoints.toFixed(6)],
    ["heightPoints", artifacts.heightPoints.toFixed(6)],
    ["baseline", artifacts.baseline.toFixed(6)],
    ["fontSizePt", formula.fontSizePt.toFixed(6)],
    ["referenceWidthPt", artifacts.referenceWidthPt.toFixed(6)],
    ["referenceHeightPt", artifacts.referenceHeightPt.toFixed(6)],
    ["referenceBaselinePt", artifacts.referenceBaselinePt.toFixed(6)],
  ]);
  writeFileSync(join(sessionDirectory, "dispatch.txt"), dispatch, { mode: 0o600 });
  writeFileSync(join(sessionsRoot, "word-active-session.txt"), sessionId, {
    mode: 0o600,
  });
  runAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(testDocumentName)}`,
    'run VB macro macro name "VisualTeX_ApplyPendingResult"',
    "end tell",
  ], 90_000);

  formula.latex = artifacts.canonicalLatex;
  formula.metadataLatex = artifacts.canonicalLatex;
  formula.metadataLines = artifacts.lines.map((line) => line.latex);
  formula.codeFormat = artifacts.codeFormat;
  formula.metadata = artifacts.metadata;
  formula.ommlBase64 = artifacts.ommlBase64;
  formula.imagePath = artifacts.imagePath;
  formula.vectorDocumentPath = artifacts.vectorDocumentPath;
  formula.fallbackImagePath = artifacts.fallbackImagePath;
  formula.widthPoints = artifacts.widthPoints;
  formula.heightPoints = artifacts.heightPoints;
  formula.baseline = artifacts.baseline;
  formula.referenceWidthPt = artifacts.referenceWidthPt;
  formula.referenceHeightPt = artifacts.referenceHeightPt;
  formula.referenceBaselinePt = artifacts.referenceBaselinePt;
  formula.svgAlignment = artifacts.svgAlignment;
  return artifacts;
}

function appendParagraphMetadata(entries, index, paragraph) {
  if (!paragraph) return;
  const prefix = `item${index}`;
  entries.push([`${prefix}paragraphId`, paragraph.id]);
  entries.push([`${prefix}paragraphStyle`, paragraph.style ?? "normal"]);
  entries.push([`${prefix}paragraphAlignment`, paragraph.alignment ?? "left"]);
  entries.push([`${prefix}listKind`, paragraph.listKind ?? "none"]);
  entries.push([`${prefix}listLevel`, String(paragraph.listLevel ?? 0)]);
  entries.push([`${prefix}paragraphStart`, paragraph.start ? "1" : "0"]);
  entries.push([`${prefix}paragraphEnd`, paragraph.end ? "1" : "0"]);
}

function appendText(entries, index, text, paragraph) {
  entries.push([`item${index}kind`, "text"]);
  entries.push([`item${index}textBase64`, base64Url(text)]);
  appendParagraphMetadata(entries, index, paragraph);
}

function appendFormula(entries, index, formula, paragraph) {
  const prefix = `item${index}`;
  entries.push([`${prefix}kind`, "formula"]);
  entries.push([`${prefix}formulaId`, formula.formulaId]);
  entries.push([`${prefix}latexBase64`, base64Url(formula.latex)]);
  entries.push([`${prefix}displayMode`, formula.displayMode]);
  entries.push([`${prefix}numbered`, formula.numbered ? "1" : "0"]);
  entries.push([`${prefix}fontSizePt`, formula.fontSizePt.toFixed(6)]);
  entries.push([`${prefix}metadata`, formula.metadata]);
  entries.push([`${prefix}ommlBase64`, formula.ommlBase64]);
  entries.push([`${prefix}nativeDocumentPath`, formula.nativePath]);
  entries.push([`${prefix}imagePath`, formula.imagePath]);
  entries.push([`${prefix}vectorDocumentPath`, formula.vectorDocumentPath]);
  entries.push([`${prefix}fallbackImagePath`, formula.fallbackImagePath]);
  entries.push([`${prefix}widthPoints`, formula.widthPoints.toFixed(6)]);
  entries.push([`${prefix}heightPoints`, formula.heightPoints.toFixed(6)]);
  entries.push([`${prefix}baseline`, formula.baseline.toFixed(6)]);
  entries.push([`${prefix}referenceWidthPt`, formula.referenceWidthPt.toFixed(6)]);
  entries.push([`${prefix}referenceHeightPt`, formula.referenceHeightPt.toFixed(6)]);
  entries.push([`${prefix}referenceBaselinePt`, formula.referenceBaselinePt.toFixed(6)]);
  appendParagraphMetadata(entries, index, paragraph);
}

function createdImagePdfInkBounds(testDocumentName, label) {
  rmSync(coordinatePdfPath, { force: true });
  rmSync(pdfExportStatusPath, { force: true });
  writeFileSync(pdfExportRequestPath, coordinatePdfPath, { mode: 0o600 });
  runAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(testDocumentName)}`,
    'run VB macro macro name "VisualTeX_ExportActiveDocumentPdfForRegression"',
    "end tell",
  ], 90_000);
  const exportStatus = existsSync(pdfExportStatusPath)
    ? readFileSync(pdfExportStatusPath, "utf8").trim()
    : "missing-status";
  if (!exportStatus.startsWith("ok|") || !existsSync(coordinatePdfPath)) {
    throw new Error(`${label} PDF export failed: ${exportStatus}`);
  }
  const swiftGeometry = spawnSync(
    "/usr/bin/swift",
    [
      join(repositoryRoot, "scripts/pdf_formula_geometry.swift"),
      coordinatePdfPath,
      "--raster-only",
    ],
    {
      encoding: "utf8",
      timeout: 120_000,
      maxBuffer: 8 * 1024 * 1024,
    },
  );
  if (swiftGeometry.status !== 0) {
    throw new Error(
      swiftGeometry.stderr.trim() || `${label} PDF raster extraction failed`,
    );
  }
  const geometry = JSON.parse(swiftGeometry.stdout);
  const components = (geometry.rasterBands ?? []).flatMap(
    (band) => band.components ?? [],
  );
  if (components.length === 0) {
    throw new Error(`${label} PDF contains no visible formula ink`);
  }
  const minX = Math.min(...components.map((component) => component.minX));
  const maxX = Math.max(...components.map((component) => component.maxX));
  return {
    minX,
    maxX,
    width: maxX - minX,
    componentCount: components.length,
    rasterBands: geometry.rasterBands,
  };
}

function assertCreatedImageFormulaInk(bounds, formula, label) {
  const minimumVisibleWidth = Math.max(18, formula.widthPoints * 0.5);
  if (!Number.isFinite(bounds.width) || bounds.width < minimumVisibleWidth) {
    throw new Error(
      `${label} rendered as a fallback glyph instead of ${formula.latex}: ` +
        JSON.stringify({ bounds, expectedWidthPoints: formula.widthPoints }),
    );
  }
}

async function runCreatedImageFormulaRegression(beforeSessions) {
  runAppleScript([
    'tell application "Microsoft Word"',
    "activate",
    "end tell",
  ], 30_000);
  await sleep(3_000);
  let testDocumentName = runAppleScript([
    'tell application "Microsoft Word"',
    "make new document",
    "set testDocument to active document",
    "repeat 3 times",
    "activate object testDocument",
    "delay 0.2",
    "end repeat",
    "activate",
    'run VB macro macro name "VisualTeX_CreateInline"',
    "return name of testDocument",
    "end tell",
  ], 60_000);

  const sessionId = await waitForWordCreateSession(beforeSessions, 30_000);
  sessionDirectory = join(sessionsRoot, sessionId);
  const request = JSON.parse(
    readFileSync(join(sessionDirectory, "request.json"), "utf8"),
  );
  await stopVisualTeXForManualWordCallback();
  const pendingMarker = request.pendingMarker ?? request.sourceObjectId ?? "";
  const fontSizePt = Number(request.fontSizePt ?? 11);
  if (
    request.mode !== "create" ||
    request.host !== "word" ||
    request.sessionId !== sessionId ||
    request.displayMode !== "inline" ||
    request.numbered ||
    !request.formulaId ||
    !request.sourceDocumentId ||
    !pendingMarker ||
    !Number.isFinite(fontSizePt)
  ) {
    throw new Error(
      `Unexpected Word image creation request: ${JSON.stringify(request)}`,
    );
  }

  const formula = formulaItem({
    formulaId: request.formulaId,
    latex: "dfdfdf",
    displayMode: "inline",
    numbered: false,
    fontSizePt,
    artifactDirectory: sessionDirectory,
  });
  nativeFiles.push(formula.nativePath);
  const dispatch = manifestText([
    ["protocolVersion", "1"],
    ["sessionId", sessionId],
    ["action", "commit"],
    ["host", "word"],
    ["mode", "create"],
    ["formulaId", formula.formulaId],
    ["displayMode", formula.displayMode],
    ["numbered", "0"],
    ["nativeEquation", "0"],
    ["imagePath", formula.imagePath],
    ["vectorDocumentPath", formula.vectorDocumentPath],
    ["fallbackImagePath", formula.fallbackImagePath],
    ["metadata", formula.metadata],
    ["latexBase64", base64Url(formula.latex)],
    ["ommlBase64", formula.ommlBase64],
    ["nativeDocumentPath", formula.nativePath],
    ["pendingMarker", pendingMarker],
    ["sourceMarker", request.sourceObjectId ?? pendingMarker],
    ["sourceDocumentId", request.sourceDocumentId],
    ["widthPoints", formula.widthPoints.toFixed(6)],
    ["heightPoints", formula.heightPoints.toFixed(6)],
    ["baseline", formula.baseline.toFixed(6)],
    ["fontSizePt", formula.fontSizePt.toFixed(6)],
    ["referenceWidthPt", formula.referenceWidthPt.toFixed(6)],
    ["referenceHeightPt", formula.referenceHeightPt.toFixed(6)],
    ["referenceBaselinePt", formula.referenceBaselinePt.toFixed(6)],
  ]);
  writeFileSync(join(sessionDirectory, "dispatch.txt"), dispatch, {
    mode: 0o600,
  });
  writeFileSync(join(sessionsRoot, "word-active-session.txt"), sessionId, {
    mode: 0o600,
  });

  const callbackStatusPath = join(
    sessionDirectory,
    "word-callback-status.txt",
  );
  rmSync(callbackStatusPath, { force: true });
  runAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(testDocumentName)}`,
    'run VB macro macro name "VisualTeX_ApplyPendingResultForRegression"',
    "end tell",
  ], 90_000);
  if (!existsSync(callbackStatusPath)) {
    throw new Error("Word did not write the image-create callback status file");
  }
  const callbackStatus = readFileSync(callbackStatusPath, "utf8");
  if (!callbackStatus.startsWith("PASS")) {
    throw new Error(`Word image-create callback failed:\n${callbackStatus}`);
  }

  const afterCommit = runFormulaRegressionReport(testDocumentName, [formula]);
  const afterCommitInk = createdImagePdfInkBounds(
    testDocumentName,
    "Created dfdfdf formula after commit",
  );
  assertCreatedImageFormulaInk(
    afterCommitInk,
    formula,
    "Created dfdfdf formula after commit",
  );
  runAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(testDocumentName)}`,
    'run VB macro macro name "VisualTeX_MigrateImageMacroButtons"',
    "end tell",
  ], 60_000);
  const afterNativeNormalization = runFormulaRegressionReport(
    testDocumentName,
    [formula],
  );
  const afterNativeNormalizationInk = createdImagePdfInkBounds(
    testDocumentName,
    "Created dfdfdf formula after native normalization",
  );
  assertCreatedImageFormulaInk(
    afterNativeNormalizationInk,
    formula,
    "Created dfdfdf formula after native normalization",
  );
  testDocumentName = saveAndReopenWordDocument(testDocumentName);
  const afterSaveReopen = runFormulaRegressionReport(
    testDocumentName,
    [formula],
  );
  const afterSaveReopenInk = createdImagePdfInkBounds(
    testDocumentName,
    "Created dfdfdf formula after save and reopen",
  );
  assertCreatedImageFormulaInk(
    afterSaveReopenInk,
    formula,
    "Created dfdfdf formula after save and reopen",
  );

  console.log(
    JSON.stringify(
      {
        sessionId,
        formulaId: formula.formulaId,
        latex: formula.latex,
        reports: {
          afterCommit,
          afterNativeNormalization,
          afterSaveReopen,
        },
        visibleInk: {
          afterCommit: afterCommitInk,
          afterNativeNormalization: afterNativeNormalizationInk,
          afterSaveReopen: afterSaveReopenInk,
        },
      },
      null,
      2,
    ),
  );
  console.log("Word image formula creation integration passed");
}

const before = currentSessionIds();
let sessionDirectory = "";
let installedWordAddinBackedUp = false;
const nativeFiles = [];
const editSessionDirectories = [];

try {
  mkdirSync(sessionsRoot, { recursive: true });
  mkdirSync(nativeRoot, { recursive: true });
  mkdirSync(officeScratchRoot, { recursive: true });
  mkdirSync(wordStartupRoot, { recursive: true });
  try {
    runAppleScript([
      'tell application "Microsoft Word" to quit saving no',
    ], 20_000);
  } catch {
    // Continue with a hard process cleanup below.
  }
  spawnSync("/usr/bin/killall", ["Microsoft Word"], {
    encoding: "utf8",
    timeout: 10_000,
  });
  await sleep(2_000);
  if (existsSync(installedWordAddinPath)) {
    copyFileSync(installedWordAddinPath, installedWordAddinBackupPath);
    rmSync(installedWordAddinPath, { force: true });
    installedWordAddinBackedUp = true;
  }
  // Use the reviewed DOTM as a real global template for every integration
  // run. This is required for Word's native double-click event sink and legacy
  // image-field migration to remain available after DOCX save/reopen; the
  // user's previous Startup add-in is restored in finally.
  copyFileSync(templatePath, installedWordAddinPath);
  if (createImageRegression) {
    await runCreatedImageFormulaRegression(before);
  } else {
  let testDocumentName = runAppleScript([
    'tell application "Microsoft Word"',
    "make new document",
    "set testDocument to active document",
    "activate",
    'run VB macro macro name "VisualTeX_InsertLatexMarkdownDocument"',
    "return name of testDocument",
    "end tell",
  ], 60_000);

  const sessionId = await waitForNewSession(before);
  sessionDirectory = join(sessionsRoot, sessionId);
  const request = JSON.parse(
    readFileSync(join(sessionDirectory, "request.json"), "utf8"),
  );
  if (
    request.operation !== "documentImport" ||
    request.sessionId !== sessionId ||
    request.host !== "word"
  ) {
    throw new Error(`Unexpected Word document import request: ${JSON.stringify(request)}`);
  }
  await stopVisualTeXForManualWordCallback();

  const formulas = [
    formulaItem({
      formulaId: crypto.randomUUID(),
      latex: "101=202",
      displayMode: "inline",
      numbered: false,
      fontSizePt: 11,
      artifactDirectory: sessionDirectory,
    }),
    formulaItem({
      formulaId: crypto.randomUUID(),
      latex: "12345=67890",
      displayMode: "block",
      numbered: false,
      fontSizePt: 14,
      artifactDirectory: sessionDirectory,
    }),
    formulaItem({
      formulaId: crypto.randomUUID(),
      latex: "24680=13579",
      displayMode: "block",
      numbered: true,
      fontSizePt: 18,
      artifactDirectory: sessionDirectory,
    }),
    formulaItem({
      formulaId: crypto.randomUUID(),
      latex: String.raw`\sum_{i=1}^{\infty}{a_k\left( x-x_0 \right) ^k}`,
      displayMode: "inline",
      numbered: false,
      fontSizePt: 12,
      artifactDirectory: sessionDirectory,
    }),
    formulaItem({
      formulaId: crypto.randomUUID(),
      latex: String.raw`\sum_{i=1}^{\infty}{a_kP_k\left( x \right)}`,
      displayMode: "inline",
      numbered: false,
      fontSizePt: 12,
      artifactDirectory: sessionDirectory,
    }),
    formulaItem({
      formulaId: crypto.randomUUID(),
      latex: legacyAlignLatex,
      expectedCodeFormat: "align",
      displayMode: "block",
      numbered: false,
      fontSizePt: 14,
      artifactDirectory: sessionDirectory,
    }),
    formulaItem({
      formulaId: crypto.randomUUID(),
      latex: legacyAlignStarLatex,
      expectedCodeFormat: "align-star",
      displayMode: "block",
      numbered: false,
      fontSizePt: 14,
      artifactDirectory: sessionDirectory,
    }),
  ];
  nativeFiles.push(...formulas.map((formula) => formula.nativePath));

  const bodyParagraphId = crypto.randomUUID();
  const followingParagraphId = crypto.randomUUID();
  const headingParagraphId = crypto.randomUUID();
  const bulletParagraphId = crypto.randomUUID();
  const bulletParagraph2Id = crypto.randomUUID();
  const bulletFormulaParagraphId = crypto.randomUUID();
  const numberParagraphId = crypto.randomUUID();
  const endingParagraphId = crypto.randomUUID();
  const items = [];
  appendText(items, 0, "结构化测试", {
    id: headingParagraphId,
    style: "heading1",
    alignment: "left",
    listKind: "none",
    listLevel: 0,
    start: true,
    end: true,
  });
  appendText(items, 1, "开头文字：", {
    id: bodyParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "none",
    listLevel: 0,
    start: true,
    end: false,
  });
  appendFormula(items, 2, formulas[0], {
    id: bodyParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "none",
    listLevel: 0,
    start: false,
    end: false,
  });
  appendText(items, 3, "，行内公式之后。", {
    id: bodyParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "none",
    listLevel: 0,
    start: false,
    end: true,
  });
  appendFormula(items, 4, formulas[1]);
  appendText(items, 5, "未编号行间公式之后。", {
    id: followingParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "none",
    listLevel: 0,
    start: true,
    end: true,
  });
  appendFormula(items, 6, formulas[2]);
  appendText(items, 7, "多项式逼近的逼近系数和原函数原则上没有硬性关系。", {
    id: bulletParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "bullet",
    listLevel: 1,
    start: true,
    end: true,
  });
  appendText(items, 8, "多项式逼近是全局性的，而幂级数逼近有收敛半径。", {
    id: bulletParagraph2Id,
    style: "normal",
    alignment: "left",
    listKind: "bullet",
    listLevel: 1,
    start: true,
    end: true,
  });
  appendText(items, 9, "形式上的区别：幂级数的形式是", {
    id: bulletFormulaParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "bullet",
    listLevel: 1,
    start: true,
    end: false,
  });
  appendFormula(items, 10, formulas[3], {
    id: bulletFormulaParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "bullet",
    listLevel: 1,
    start: false,
    end: false,
  });
  appendText(items, 11, "而多项式级数的形式是", {
    id: bulletFormulaParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "bullet",
    listLevel: 1,
    start: false,
    end: false,
  });
  appendFormula(items, 12, formulas[4], {
    id: bulletFormulaParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "bullet",
    listLevel: 1,
    start: false,
    end: true,
  });
  appendText(items, 13, "编号列表正文", {
    id: numberParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "number",
    listLevel: 1,
    start: true,
    end: true,
  });
  appendFormula(items, 14, formulas[5]);
  appendFormula(items, 15, formulas[6]);
  appendText(items, 16, "结尾文字。", {
    id: endingParagraphId,
    style: "normal",
    alignment: "left",
    listKind: "none",
    listLevel: 0,
    start: true,
    end: true,
  });

  const manifestPath = join(sessionDirectory, "document-import.txt");
  const entries = [
    ["protocolVersion", "1"],
    ["sessionId", sessionId],
    ["outputKind", outputKind],
    ["sourceDocumentId", request.sourceDocumentId],
    ["bookmarkName", request.documentImport.bookmarkName],
    ["itemCount", String(diagnosticItemLimit)],
    ...items,
  ];
  writeFileSync(manifestPath, manifestText(entries), { mode: 0o600 });

  const dispatch = manifestText([
    ["protocolVersion", "1"],
    ["sessionId", sessionId],
    ["action", "documentCommit"],
    ["host", "word"],
    ["sourceDocumentId", request.sourceDocumentId],
    ["bookmarkName", request.documentImport.bookmarkName],
    ["documentImportPath", manifestPath],
  ]);
  writeFileSync(join(sessionDirectory, "dispatch.txt"), dispatch, { mode: 0o600 });
  writeFileSync(join(sessionsRoot, "word-active-session.txt"), sessionId, { mode: 0o600 });

  const bookmarkPreflight = runAppleScript([
    'tell application "Microsoft Word"',
    `set documentObject to document ${JSON.stringify(testDocumentName)}`,
    "activate object documentObject",
    `set targetExists to exists bookmark ${JSON.stringify(request.documentImport.bookmarkName)} of documentObject`,
    "set bookmarkNames to name of every bookmark of documentObject",
    'return (targetExists as text) & (ASCII character 31) & (bookmarkNames as text)',
    "end tell",
  ]);
  if (!bookmarkPreflight.startsWith("true\x1f")) {
    throw new Error(
      `The Word document-import bookmark disappeared before callback: ${bookmarkPreflight}`,
    );
  }

  const callbackStatusPath = join(
    sessionDirectory,
    "word-callback-status.txt",
  );
  rmSync(callbackStatusPath, { force: true });
  runAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(testDocumentName)}`,
    'run VB macro macro name "VisualTeX_ConfigureDocumentImportParagraphSpacingRegression"',
    'run VB macro macro name "VisualTeX_ApplyPendingResultForRegression"',
    "end tell",
  ], 90_000);
  if (!existsSync(callbackStatusPath)) {
    throw new Error("Word did not write the regression callback status file");
  }
  const callbackStatus = readFileSync(callbackStatusPath, "utf8");
  if (!callbackStatus.startsWith("PASS")) {
    throw new Error(`Word document-import callback failed:\n${callbackStatus}`);
  }
  if (diagnosticItemLimit < 17) {
    throw new Error(`${diagnosticSuccessPrefix}${diagnosticItemLimit}`);
  }
  let formulaRegressionReport = runFormulaRegressionReport(
    testDocumentName,
    formulas,
  );
  const initialFormulaContainerReport = inspectWordFormulaContainers(
    testDocumentName,
    formulas,
    "after-import",
  );

  const pdfPath = coordinatePdfPath;
  rmSync(pdfPath, { force: true });
  const bookmarkNames = formulas.map((formula) => nativeBookmark(formula.formulaId));
  const numberedCompactId = compactFormulaId(formulas[2].formulaId);
  const numberBookmarkName = `VT_R_${numberedCompactId}`;
  const inspectionLines = [
    'tell application "Microsoft Word"',
    `set documentObject to document ${JSON.stringify(testDocumentName)}`,
    "activate object documentObject",
    "activate",
    "set documentText to content of text object of documentObject",
    "set bookmarkNames to name of every bookmark of documentObject",
    "set tableCount to count of tables of documentObject",
    "set shapeCount to count of inline shapes of documentObject",
    "set pageSetupObject to page setup of section 1 of documentObject",
    "set pageWidthValue to page width of pageSetupObject",
    "set pageHeightValue to page height of pageSetupObject",
    "set leftMarginValue to left margin of pageSetupObject",
    "set rightMarginValue to right margin of pageSetupObject",
  ];
  if (outputKind === "omml") {
    inspectionLines.push(
      'set alternativeTexts to ""',
      `set inlineSize to font size of font object of text object of bookmark ${JSON.stringify(bookmarkNames[0])} of documentObject`,
      `set displaySize to font size of font object of text object of bookmark ${JSON.stringify(bookmarkNames[1])} of documentObject`,
      `set numberedSize to font size of font object of text object of bookmark ${JSON.stringify(bookmarkNames[2])} of documentObject`,
      `set displayFormulaRange to text object of bookmark ${JSON.stringify(bookmarkNames[1])} of documentObject`,
      "set displayStartPosition to start of content of displayFormulaRange",
      "set displayEndPosition to end of content of displayFormulaRange",
      "set displayStartRange to create range documentObject start displayStartPosition end displayStartPosition",
      "set displayEndRange to create range documentObject start displayEndPosition end displayEndPosition",
      "set displayLeft to get range information displayStartRange information type horizontal position relative to page",
      "set displayRight to get range information displayEndRange information type horizontal position relative to page",
      "set displayTop to get range information displayStartRange information type vertical position relative to page",
      "set inlineWidth to 0",
      "set inlineHeight to 0",
      "set displayHeight to 0",
      `set numberedFormulaRange to text object of bookmark ${JSON.stringify(bookmarkNames[2])} of documentObject`,
      "set numberedStartPosition to start of content of numberedFormulaRange",
      "set numberedEndPosition to end of content of numberedFormulaRange",
      "set numberedStartRange to create range documentObject start numberedStartPosition end numberedStartPosition",
      "set numberedEndRange to create range documentObject start numberedEndPosition end numberedEndPosition",
      "set numberedLeft to get range information numberedStartRange information type horizontal position relative to page",
      "set numberedRight to get range information numberedEndRange information type horizontal position relative to page",
      "set numberedTop to get range information numberedStartRange information type vertical position relative to page",
      "set numberedHeight to 0",
    );
  } else {
    inspectionLines.push(
      "set inlineShapeObject to inline shape 1 of documentObject",
      "set displayShapeObject to inline shape 2 of documentObject",
      "set numberedShapeObject to inline shape 3 of documentObject",
      "set listFormulaShapeObject1 to inline shape 4 of documentObject",
      "set listFormulaShapeObject2 to inline shape 5 of documentObject",
      "set alignShapeObject to inline shape 6 of documentObject",
      "set alignStarShapeObject to inline shape 7 of documentObject",
      'set alternativeTexts to (alternative text of inlineShapeObject) & "|" & (alternative text of displayShapeObject) & "|" & (alternative text of numberedShapeObject) & "|" & (alternative text of listFormulaShapeObject1) & "|" & (alternative text of listFormulaShapeObject2) & "|" & (alternative text of alignShapeObject) & "|" & (alternative text of alignStarShapeObject)',
      "set inlineSize to font size of font object of text object of inlineShapeObject",
      "set inlineWidth to width of inlineShapeObject",
      "set inlineHeight to height of inlineShapeObject",
      "set displaySize to font size of font object of text object of displayShapeObject",
      "set numberedSize to font size of font object of text object of numberedShapeObject",
      "set displayFormulaRange to text object of displayShapeObject",
      "set displayStartPosition to start of content of displayFormulaRange",
      "set displayStartRange to create range documentObject start displayStartPosition end displayStartPosition",
      "set displayLeft to get range information displayStartRange information type horizontal position relative to page",
      "set displayRight to displayLeft + (width of displayShapeObject)",
      "set displayTop to get range information displayStartRange information type vertical position relative to page",
      "set displayHeight to height of displayShapeObject",
      "set numberedFormulaRange to text object of numberedShapeObject",
      "set numberedStartPosition to start of content of numberedFormulaRange",
      "set numberedStartRange to create range documentObject start numberedStartPosition end numberedStartPosition",
      "set numberedLeft to get range information numberedStartRange information type horizontal position relative to page",
      "set numberedRight to numberedLeft + (width of numberedShapeObject)",
      "set numberedTop to get range information numberedStartRange information type vertical position relative to page",
      "set numberedHeight to height of numberedShapeObject",
    );
  }
  inspectionLines.push(
    `set numberRange to text object of bookmark ${JSON.stringify(numberBookmarkName)} of documentObject`,
    "set numberStartPosition to start of content of numberRange",
    "set numberEndPosition to end of content of numberRange",
    "set numberStartRange to create range documentObject start numberStartPosition end numberStartPosition",
    "set numberEndRange to create range documentObject start numberEndPosition end numberEndPosition",
    "set numberLeft to get range information numberStartRange information type horizontal position relative to page",
    "set numberRight to get range information numberEndRange information type horizontal position relative to page",
    "set numberTop to get range information numberStartRange information type vertical position relative to page",
    'return documentText & "\n---VT---\n" & (bookmarkNames as text) & "\n---VT---\n" & alternativeTexts & "\n---VT---\n" & shapeCount & "," & tableCount & "," & inlineSize & "," & displaySize & "," & numberedSize & "," & inlineWidth & "," & inlineHeight & "," & pageWidthValue & "," & pageHeightValue & "," & leftMarginValue & "," & rightMarginValue & "," & displayLeft & "," & displayRight & "," & displayTop & "," & displayHeight & "," & numberedLeft & "," & numberedRight & "," & numberedTop & "," & numberedHeight & "," & numberLeft & "," & numberRight & "," & numberTop',
    "end tell",
  );
  const inspection = runAppleScript(inspectionLines);
  rmSync(pdfExportStatusPath, { force: true });
  writeFileSync(pdfExportRequestPath, pdfPath, { mode: 0o600 });
  runAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(testDocumentName)}`,
    'run VB macro macro name "VisualTeX_ExportActiveDocumentPdfForRegression"',
    "end tell",
  ], 90_000);
  const exportStatus = existsSync(pdfExportStatusPath)
    ? readFileSync(pdfExportStatusPath, "utf8").trim()
    : "missing-status";
  if (!exportStatus.startsWith("ok|")) {
    throw new Error(`Word PDF regression export failed: ${exportStatus}`);
  }
  if (!existsSync(pdfPath)) {
    throw new Error(`Word did not export the coordinate verification PDF: ${pdfPath}`);
  }
  const swiftGeometryArguments = [
    join(repositoryRoot, "scripts/pdf_formula_geometry.swift"),
    pdfPath,
    ...(outputKind === "omml"
      ? [
          formulas[1].pdfToken,
          formulas[2].pdfToken,
          formulas[5].pdfToken,
          formulas[6].pdfToken,
        ]
      : ["--number-only"]),
  ];
  const swiftGeometry = spawnSync(
    "/usr/bin/swift",
    swiftGeometryArguments,
    {
      encoding: "utf8",
      timeout: 120_000,
      maxBuffer: 8 * 1024 * 1024,
    },
  );
  if (swiftGeometry.status !== 0) {
    throw new Error(
      swiftGeometry.stderr.trim() || "PDFKit formula geometry extraction failed",
    );
  }
  const renderedGeometry = JSON.parse(swiftGeometry.stdout);
  const ommlAlignmentGeometry = [];
  if (outputKind === "omml") {
    if (!Array.isArray(renderedGeometry.aligned) || renderedGeometry.aligned.length !== 2) {
      throw new Error(
        `PDF regression did not return both aligned OMML formulas: ${swiftGeometry.stdout}`,
      );
    }
    for (const [index, alignedFormula] of renderedGeometry.aligned.entries()) {
      const positions = alignedFormula.relationshipXs;
      if (
        !Array.isArray(positions) ||
        positions.length !== 2 ||
        positions.some((position) => !Number.isFinite(position))
      ) {
        throw new Error(
          `Aligned OMML formula ${index + 1} did not expose both PDF relationship positions: ${JSON.stringify(alignedFormula)}`,
        );
      }
      const spread = Math.max(...positions) - Math.min(...positions);
      if (spread > 0.5) {
        throw new Error(
          `Aligned OMML formula ${index + 1} has a misaligned relationship column: ${JSON.stringify({ positions, spread })}`,
        );
      }
      ommlAlignmentGeometry.push({ positions, spread });
    }
  }

  const [documentText, bookmarkText, alternativeText, numericText] =
    inspection.split("\n---VT---\n");
  const [
    shapeCount,
    tableCount,
    inlineSize,
    displaySize,
    numberedSize,
    inlineWidth,
    inlineHeight,
    pageWidth,
    pageHeight,
    leftMargin,
    rightMargin,
    displayLeft,
    displayRight,
    displayTop,
    displayHeight,
    numberedLeft,
    numberedRight,
    numberedTop,
    numberedHeight,
    numberLeft,
    numberRight,
    numberTop,
  ] = numericText.split(",").map(Number);

  for (const expected of ["开头文字：", "行内公式之后。", "未编号行间公式之后。", "结尾文字。"]) {
    if (!documentText.includes(expected)) {
      throw new Error(`Word import text is missing ${expected}: ${JSON.stringify(documentText)}`);
    }
  }
  if (outputKind === "omml") {
    for (const bookmarkName of bookmarkNames) {
      if (!bookmarkText.includes(bookmarkName)) {
        throw new Error(`Word import is missing formula bookmark ${bookmarkName}`);
      }
    }
    if (shapeCount !== 0) {
      throw new Error(`OMML import unexpectedly created ${shapeCount} inline shapes`);
    }
  } else {
    if (shapeCount !== formulas.length) {
      throw new Error(
        `Image import created ${shapeCount} inline shapes instead of ${formulas.length}`,
      );
    }
    const metadataPayloads = alternativeText.split("|");
    if (
      metadataPayloads.length !== formulas.length ||
      metadataPayloads.some((value) => !value.startsWith("visualtex:v1:deflate:"))
    ) {
      throw new Error(
        "Image formulas did not retain independent VisualTeX metadata payloads",
      );
    }
    metadataPayloads.forEach((payload, index) => {
      const metadata = decodeFormulaMetadata(payload);
      const expected = formulas[index];
      if (
        !metadata ||
        metadata.formulaId !== expected.formulaId ||
        metadata.displayMode !== expected.displayMode ||
        Boolean(metadata.numbered) !== expected.numbered ||
        Math.abs((metadata.fontSizePt ?? 0) - expected.fontSizePt) > 0.001
      ) {
        throw new Error(
          `Image formula ${index + 1} did not retain its independent identity, mode, numbering and font size`,
        );
      }
    });
  }
  for (const numberBookmark of [
    `VT_R_${numberedCompactId}`,
    `VT_N_${numberedCompactId}`,
    `VT_C_${numberedCompactId}`,
  ]) {
    if (!bookmarkText.includes(numberBookmark)) {
      throw new Error(`Numbered display formula is missing ${numberBookmark}`);
    }
  }
  const geometryValues = [
    pageWidth,
    pageHeight,
    leftMargin,
    rightMargin,
    inlineWidth,
    inlineHeight,
    displayLeft,
    displayRight,
    displayTop,
    displayHeight,
    numberedLeft,
    numberedRight,
    numberedTop,
    numberedHeight,
    numberLeft,
    numberRight,
    numberTop,
  ];
  if (geometryValues.some((value) => !Number.isFinite(value) || value < 0)) {
    throw new Error(`Word returned invalid formula geometry: ${JSON.stringify(geometryValues)}`);
  }
  const textBoundaryLeft = leftMargin;
  const textBoundaryRight = pageWidth - rightMargin;
  const textBoundaryCenter = (textBoundaryLeft + textBoundaryRight) / 2;
  const displayCenter = (displayLeft + displayRight) / 2;
  const numberedCenter = (numberedLeft + numberedRight) / 2;
  const displayCenterError = Math.abs(displayCenter - textBoundaryCenter);
  const numberedCenterError = Math.abs(numberedCenter - textBoundaryCenter);
  const displayToNumberedCenterError = Math.abs(displayCenter - numberedCenter);
  const imageRasterGeometry =
    outputKind === "image"
      ? resolveImageRasterGeometry(
          renderedGeometry.rasterBands ?? [],
          textBoundaryCenter,
          {
            pageHeight,
            displayTop,
            displayHeight,
            numberedTop,
            numberedHeight,
            numberTop,
          },
        )
      : null;
  const measuredUnnumberedCenter =
    outputKind === "omml"
      ? renderedGeometry.unnumbered.centerX
      : imageRasterGeometry.unnumbered.centerX;
  const measuredNumberedCenter =
    outputKind === "omml"
      ? renderedGeometry.numbered.centerX
      : imageRasterGeometry.numbered.centerX;
  const renderedUnnumberedCenterError = Math.abs(
    measuredUnnumberedCenter - textBoundaryCenter,
  );
  const renderedNumberedCenterError = Math.abs(
    measuredNumberedCenter - textBoundaryCenter,
  );
  const renderedFormulaCenterDifference = Math.abs(
    measuredUnnumberedCenter - measuredNumberedCenter,
  );
  let imageVisualCalibration = null;

  const sizes = [inlineSize, displaySize, numberedSize];
  const expectedSizes = [11, 14, 18];
  if (outputKind === "omml") {
    sizes.forEach((size, index) => {
      if (!Number.isFinite(size) || Math.abs(size - expectedSizes[index]) > 0.1) {
        throw new Error(`Formula ${index + 1} font size mismatch: ${size}`);
      }
    });
  } else {
    if (sizes.some((size) => !Number.isFinite(size) || size <= 0)) {
      throw new Error(`Word returned invalid image Range.Font.Size values: ${sizes.join(",")}`);
    }
    const actualDimensions = [
      [inlineWidth, inlineHeight],
      [displayRight - displayLeft, displayHeight],
      [numberedRight - numberedLeft, numberedHeight],
    ];
    actualDimensions.forEach(([width, height], index) => {
      const expected = formulas[index];
      if (
        Math.abs(width - expected.widthPoints) > 0.15 ||
        Math.abs(height - expected.heightPoints) > 0.15
      ) {
        throw new Error(
          `Image formula ${index + 1} visual geometry does not match its independent font size: ` +
            `${JSON.stringify({ width, height, expectedWidth: expected.widthPoints, expectedHeight: expected.heightPoints })}`,
        );
      }
    });
    for (const [label, inkBounds, wordShapeWidth] of [
      [
        "unnumbered display",
        imageRasterGeometry.unnumberedInk,
        displayRight - displayLeft,
      ],
      [
        "numbered display",
        imageRasterGeometry.numberedInk,
        numberedRight - numberedLeft,
      ],
    ]) {
      const minimumInkWidth = Math.max(20, wordShapeWidth * 0.3);
      if (!inkBounds || inkBounds.width < minimumInkWidth) {
        throw new Error(
          `${label} formula rendered as a narrow fallback glyph instead of ` +
            `its complete SVG: ${JSON.stringify({ inkBounds, wordShapeWidth, minimumInkWidth })}`,
        );
      }
    }
    const calibrationInk = imageRasterGeometry.unnumberedInk;
    const nativeWidthRatio = calibrationInk.width / nativeCalibrationWidthPt;
    const boxToInkHeightRatio = displayHeight / calibrationInk.height;
    if (nativeWidthRatio < 0.95 || nativeWidthRatio > 1.05) {
      throw new Error(
        `The 14 pt Word image formula does not visually match the native ` +
          `Cambria Math calibration: ${JSON.stringify({ calibrationInk, nativeCalibrationWidthPt, nativeWidthRatio })}`,
      );
    }
    if (boxToInkHeightRatio > 1.8) {
      throw new Error(
        `The Word image formula retains excessive transparent vertical padding: ` +
          `${JSON.stringify({ displayHeight, calibrationInk, boxToInkHeightRatio })}`,
      );
    }
    imageVisualCalibration = {
      nativeCalibrationWidthPt,
      imageInkWidthPt: calibrationInk.width,
      nativeWidthRatio,
      wordShapeHeightPt: displayHeight,
      imageInkHeightPt: calibrationInk.height,
      boxToInkHeightRatio,
      visualScale: wordImageVisualScale,
      displayPaddingPx: wordDisplayPaddingPx,
    };
  }

  const centerTolerancePt = outputKind === "image" ? 0.5 : 0.25;
  if (Math.abs(renderedGeometry.pageWidth - pageWidth) > 0.25) {
    throw new Error(
      `Word/PDF page width mismatch: Word=${pageWidth}, PDF=${renderedGeometry.pageWidth}`,
    );
  }
  if (outputKind === "omml") {
    if (renderedUnnumberedCenterError > centerTolerancePt) {
      throw new Error(
        `Unnumbered display formula is not centered: error=${renderedUnnumberedCenterError} pt`,
      );
    }
    if (renderedNumberedCenterError > centerTolerancePt) {
      throw new Error(
        `Numbered display formula is not centered: error=${renderedNumberedCenterError} pt`,
      );
    }
  } else {
    // Word's horizontal position for an InlineShape Range is a paragraph/text
    // anchor, not the centered visual image edge. The symmetric fixture places
    // its relationship sign at the image center, so the PDF raster marker must
    // cross the text-area center. Its glyph ink center can be a few points off
    // because the '=' outline and SVG padding are not geometrically symmetric.
    for (const [label, geometry] of [
      ["unnumbered", imageRasterGeometry.unnumbered],
      ["numbered", imageRasterGeometry.numbered],
    ]) {
      if (
        textBoundaryCenter < geometry.minX - centerTolerancePt ||
        textBoundaryCenter > geometry.maxX + centerTolerancePt
      ) {
        throw new Error(
          `${label} image formula center marker does not cross the text-area center: ${JSON.stringify({ geometry, textBoundaryCenter, centerTolerancePt })}`,
        );
      }
    }
  }
  if (renderedFormulaCenterDifference > centerTolerancePt) {
    throw new Error(
      `Numbering shifted the formula center marker: difference=${renderedFormulaCenterDifference} pt`,
    );
  }
  const equationNumberGeometry =
    outputKind === "omml"
      ? renderedGeometry.equationNumber
      : imageRasterGeometry.equationNumber;
  const numberedFormulaPdfCenterY =
    outputKind === "omml"
      ? renderedGeometry.numbered.centerY
      : imageRasterGeometry.numbered.centerY;
  const equationNumberInkCenterDifference = Math.abs(
    equationNumberGeometry.centerY - numberedFormulaPdfCenterY,
  );
  const equationNumberVerticalError =
    outputKind === "omml"
      ? equationNumberInkCenterDifference
      : Math.abs(numberTop - numberedTop);
  if (outputKind === "omml") {
    if (equationNumberVerticalError > 0.25) {
      throw new Error(
        `Equation number is not vertically centered with its formula: error=${equationNumberVerticalError} pt`,
      );
    }
  } else {
    const rasterTolerancePt = 1.5;
    if (equationNumberVerticalError > 0.25) {
      throw new Error(
        `Image equation number and formula outer boxes do not share a top edge: ${JSON.stringify({ numberedTop, numberTop, equationNumberVerticalError })}`,
      );
    }
    for (const [label, geometry] of [
      ["formula ink", imageRasterGeometry.numberedInk],
      ["equation number", equationNumberGeometry],
    ]) {
      if (!geometry || geometry.height > numberedHeight + rasterTolerancePt) {
        throw new Error(
          `Numbered image ${label} is taller than the Word image outer box: ${JSON.stringify({ geometry, numberedHeight, rasterTolerancePt })}`,
        );
      }
    }
    if (equationNumberInkCenterDifference > 0.5) {
      throw new Error(
        `Image equation number is not vertically centered with the formula ink: ${JSON.stringify({ equationNumberInkCenterDifference, numberedFormula: imageRasterGeometry.numberedInk, equationNumber: equationNumberGeometry })}`,
      );
    }
  }
  const measuredNumberedRight =
    outputKind === "omml"
      ? renderedGeometry.numbered.maxX
      : textBoundaryCenter + (numberedRight - numberedLeft) / 2;
  if (
    equationNumberGeometry.minX <= measuredNumberedRight + 4 ||
    equationNumberGeometry.maxX > textBoundaryRight + 0.5
  ) {
    throw new Error(
      `Equation number is outside the expected right-side region: ${JSON.stringify({
        formulaRight: measuredNumberedRight,
        numberLeft: equationNumberGeometry.minX,
        numberRight: equationNumberGeometry.maxX,
        textBoundaryRight,
      })}`,
    );
  }

  const editRegressions = [];
  if (outputKind === "image") {
    const imageEditCases = [
      {
        formula: formulas[0],
        shapeIndex: 1,
        codeFormat: "raw",
        expectedLines: formulas[0].metadataLines,
        recovery: true,
      },
      {
        formula: formulas[5],
        shapeIndex: 6,
        codeFormat: "align",
        expectedLines: formulas[5].metadataLines,
        updatedLines: ["1 = 22 + 333 + q", "44444 = 55 + r"],
      },
      {
        formula: formulas[6],
        shapeIndex: 7,
        codeFormat: "align-star",
        expectedLines: formulas[6].metadataLines,
        updatedLines: ["666 = 777 + 8 + s", "999999 = 0 + t"],
      },
    ];

    for (const editCase of imageEditCases) {
      const sessionsBeforeEdit = currentSessionIds();
      if (editCase.recovery) rmSync(imageEditStatusPath, { force: true });
      runAppleScript([
        'tell application "Microsoft Word"',
        `set documentObject to document ${JSON.stringify(testDocumentName)}`,
        "activate object documentObject",
        "activate",
        `set formulaShape to inline shape ${editCase.shapeIndex} of documentObject`,
        "select text object of formulaShape",
        editCase.recovery
          ? 'run VB macro macro name "VisualTeX_RunSelectedImageEditRecoveryRegression"'
          : 'run VB macro macro name "VisualTeX_DoubleClickEditSelected"',
        "end tell",
      ], 60_000);
      const editSessionId = await waitForFormulaEditSession(
        sessionsBeforeEdit,
        editCase.formula.formulaId,
      );
      editSessionDirectories.push(join(sessionsRoot, editSessionId));
      const editSession = validateFormulaEditSession(
        editSessionId,
        editCase.formula,
        editCase.codeFormat,
        editCase.expectedLines,
      );

      let restoredReference;
      if (editCase.recovery) {
        const imageEditStatus = existsSync(imageEditStatusPath)
          ? readFileSync(imageEditStatusPath, "utf8").trim()
          : "missing-status";
        if (!imageEditStatus.startsWith("ok|")) {
          throw new Error(
            `Word image edit recovery regression failed: ${imageEditStatus}`,
          );
        }
        restoredReference = imageEditStatus.slice(3);
        const expectedReference =
          `visualtex:formula-ref:v1:${editCase.formula.formulaId}:` +
          `${editCase.formula.displayMode}:${editCase.formula.numbered ? "1" : "0"}`;
        if (restoredReference !== expectedReference) {
          throw new Error(
            `Word did not restore the image formula Title before editing: ${JSON.stringify({
              restoredReference,
              expectedReference,
            })}`,
          );
        }
      }

      const regression = {
        kind: editCase.recovery
          ? "image-metadata-title-recovery"
          : "image-batch-edit-session",
        sessionId: editSessionId,
        formulaId: editCase.formula.formulaId,
        codeFormat: editSession.normalized.codeFormat,
        lines: editSession.normalized.lines.map((line) => line.latex),
        ...(restoredReference ? { restoredReference } : {}),
      };

      if (editCase.updatedLines) {
        const replacement = commitEditedImageFormula(
          testDocumentName,
          editSessionId,
          editCase.formula,
          editSession,
          editCase.updatedLines,
        );
        const sessionsBeforeReplacementEdit = currentSessionIds();
        runAppleScript([
          'tell application "Microsoft Word"',
          `set documentObject to document ${JSON.stringify(testDocumentName)}`,
          "activate object documentObject",
          "activate",
          `set formulaShape to inline shape ${editCase.shapeIndex} of documentObject`,
          "select text object of formulaShape",
          'run VB macro macro name "VisualTeX_DoubleClickEditSelected"',
          "end tell",
        ], 60_000);
        const replacementEditSessionId = await waitForFormulaEditSession(
          sessionsBeforeReplacementEdit,
          editCase.formula.formulaId,
        );
        editSessionDirectories.push(
          join(sessionsRoot, replacementEditSessionId),
        );
        const replacementEditSession = validateFormulaEditSession(
          replacementEditSessionId,
          editCase.formula,
          editCase.codeFormat,
          editCase.updatedLines,
        );
        Object.assign(regression, {
          kind: "image-align-edit-replacement",
          replacementSessionId: replacementEditSessionId,
          replacementLatex: replacement.canonicalLatex,
          replacementLines: replacementEditSession.normalized.lines.map(
            (line) => line.latex,
          ),
          svgRelationshipPositions: replacement.svgAlignment.positions,
          svgRelationshipSpread: replacement.svgAlignment.spread,
        });
      }
      editRegressions.push(regression);
    }
    formulaRegressionReport = runFormulaRegressionReport(
      testDocumentName,
      formulas,
    );
  } else {
    const nativeEditCases = [
      {
        formula: formulas[5],
        codeFormat: "align",
        lines: ["1 = 22 + 333", "44444 = 55"],
      },
      {
        formula: formulas[6],
        codeFormat: "align-star",
        lines: ["666 = 777 + 8", "999999 = 0"],
      },
    ];
    for (const editCase of nativeEditCases) {
      const sessionsBeforeEdit = currentSessionIds();
      runAppleScript([
        'tell application "Microsoft Word"',
        `set documentObject to document ${JSON.stringify(testDocumentName)}`,
        "activate object documentObject",
        "activate",
        `select text object of bookmark ${JSON.stringify(nativeBookmark(editCase.formula.formulaId))} of documentObject`,
        'run VB macro macro name "VisualTeX_DoubleClickEditSelected"',
        "end tell",
      ], 60_000);
      const editSessionId = await waitForFormulaEditSession(
        sessionsBeforeEdit,
        editCase.formula.formulaId,
      );
      editSessionDirectories.push(join(sessionsRoot, editSessionId));
      const editSession = validateFormulaEditSession(
        editSessionId,
        editCase.formula,
        editCase.codeFormat,
        editCase.lines,
      );
      editRegressions.push({
        kind: "omml-multiline-edit",
        sessionId: editSessionId,
        formulaId: editCase.formula.formulaId,
        codeFormat: editSession.normalized.codeFormat,
        lines: editSession.normalized.lines.map((line) => line.latex),
        displayMode: editSession.request.displayMode,
        numbered: editSession.request.numbered,
        fontSizePt: editSession.request.fontSizePt,
      });
    }
  }

  const postEditFormulaContainerReport = inspectWordFormulaContainers(
    testDocumentName,
    formulas,
    "after-edit",
  );
  testDocumentName = saveAndReopenWordDocument(testDocumentName);
  const reopenedFormulaContainerReport = inspectWordFormulaContainers(
    testDocumentName,
    formulas,
    "after-save-reopen",
  );
  formulaRegressionReport = runFormulaRegressionReport(
    testDocumentName,
    formulas,
  );

  if (physicalDoubleClick) {
    const physicalFormulaIndex = {
      "image-inline": 0,
      "image-block": 1,
      "image-align": 5,
      "image-align-star": 6,
      "omml-inline": 0,
      "omml-block": 1,
      "omml-align": 5,
      "omml-align-star": 6,
    }[physicalTarget];
    const physicalFormula = formulas[physicalFormulaIndex];
    await startVisualTeXForPhysicalRegression();
    const sessionsBeforePhysicalEdit = currentSessionIds();
    const physicalSelection = runAppleScript([
      'tell application "Microsoft Word"',
      `set documentObject to document ${JSON.stringify(testDocumentName)}`,
      "activate object documentObject",
      "activate",
      'run VB macro macro name "VisualTeX_AssertWordHostSelfTest"',
      ...(outputKind === "image"
        ? [
            `set formulaShape to inline shape ${physicalFormulaIndex + 1} of documentObject`,
            "set formulaRange to text object of formulaShape",
            "select formulaRange",
            'return "image" & (ASCII character 31) & (start of content of formulaRange as text) & (ASCII character 31) & (end of content of formulaRange as text) & (ASCII character 31) & (width of formulaShape as text) & (ASCII character 31) & (height of formulaShape as text)',
          ]
        : [
            `set formulaRange to text object of bookmark ${JSON.stringify(nativeBookmark(physicalFormula.formulaId))} of documentObject`,
            "select formulaRange",
            'return "omml" & (ASCII character 31) & (start of content of formulaRange as text) & (ASCII character 31) & (end of content of formulaRange as text)',
          ]),
      "end tell",
    ]);
    console.log(
      `WORD_PHYSICAL_DOUBLE_CLICK_READY|${JSON.stringify({
        documentName: testDocumentName,
        target: physicalTarget,
        outputKind,
        formulaId: physicalFormula.formulaId,
        selection: physicalSelection.split("\x1f"),
      })}`,
    );
    const physicalClick = physicallyDoubleClickSelectedWordFormula(
      testDocumentName,
    );

    const physicalEditSessionId = await waitForFormulaEditSession(
      sessionsBeforePhysicalEdit,
      physicalFormula.formulaId,
      600_000,
    );
    editSessionDirectories.push(join(sessionsRoot, physicalEditSessionId));
    const physicalEditSession = validateFormulaEditSession(
      physicalEditSessionId,
      physicalFormula,
      physicalFormula.codeFormat,
      physicalFormula.metadataLines,
    );
    const editorReadiness = await waitForPhysicalEditorReadiness(
      physicalEditSessionId,
      physicalFormula.formulaId,
    );
    await assertSinglePhysicalEditSession(
      sessionsBeforePhysicalEdit,
      physicalEditSessionId,
      physicalFormula.formulaId,
    );
    editRegressions.push({
      kind: `${physicalTarget}-physical-double-click`,
      target: physicalTarget,
      sessionId: physicalEditSessionId,
      formulaId: physicalFormula.formulaId,
      codeFormat: physicalEditSession.normalized.codeFormat,
      lines: physicalEditSession.normalized.lines.map((line) => line.latex),
      physicalClick,
      editorReadiness,
    });
  }

  console.log(
    JSON.stringify(
      {
        sessionId,
        outputKind,
        formulas: formulas.map((formula, index) => ({
          formulaId: formula.formulaId,
          displayMode: formula.displayMode,
          numbered: formula.numbered,
          fontSizePt: formula.fontSizePt,
          codeFormat: formula.codeFormat,
          lines: formula.metadataLines,
          ...(formula.svgAlignment
            ? { svgAlignment: formula.svgAlignment }
            : {}),
          ...(outputKind === "omml"
            ? { bookmark: nativeBookmark(formula.formulaId) }
            : {
                shapeIndex: index + 1,
                wordObjectType: "InlineShape",
                layoutStructure:
                  formula.displayMode === "block"
                    ? formula.numbered
                      ? "numbered-display-paragraph"
                      : "dedicated-display-paragraph"
                    : "inline-text-flow",
              }),
        })),
        shapeCount,
        tableCount,
        geometry: {
          pageWidth,
          leftMargin,
          rightMargin,
          textBoundaryLeft,
          textBoundaryRight,
          textBoundaryCenter,
          unnumberedDisplay: {
            left: displayLeft,
            right: displayRight,
            center: displayCenter,
            centerError: displayCenterError,
          },
          numberedDisplay: {
            left: numberedLeft,
            right: numberedRight,
            center: numberedCenter,
            centerError: numberedCenterError,
          },
          equationNumber: {
            left: numberLeft,
            right: numberRight,
          },
          displayToNumberedCenterError,
          renderedPdf: {
            pageWidth: renderedGeometry.pageWidth,
            pageHeight: renderedGeometry.pageHeight,
            unnumberedDisplay:
              outputKind === "omml"
                ? {
                    ...renderedGeometry.unnumbered,
                    centerError: renderedUnnumberedCenterError,
                  }
                : {
                    ...imageRasterGeometry.unnumbered,
                    wordShapeWidth: displayRight - displayLeft,
                    wordShapeHeight: displayHeight,
                    inkBounds: imageRasterGeometry.unnumberedInk,
                    centerError: renderedUnnumberedCenterError,
                  },
            numberedDisplay:
              outputKind === "omml"
                ? {
                    ...renderedGeometry.numbered,
                    centerError: renderedNumberedCenterError,
                  }
                : {
                    ...imageRasterGeometry.numbered,
                    wordShapeWidth: numberedRight - numberedLeft,
                    wordShapeHeight: numberedHeight,
                    inkBounds: imageRasterGeometry.numberedInk,
                    centerError: renderedNumberedCenterError,
                  },
            equationNumber: {
              ...equationNumberGeometry,
              verticalCenterError: equationNumberVerticalError,
              inkCenterDifference: equationNumberInkCenterDifference,
              rightBoundaryInset: textBoundaryRight - equationNumberGeometry.maxX,
            },
            formulaCenterDifference: renderedFormulaCenterDifference,
            centerTolerancePt,
            imageVisualCalibration,
            ommlAlignmentGeometry,
          },
        },
        documentText,
        formulaRegressionReport,
        formulaContainerReports: {
          afterImport: initialFormulaContainerReport,
          afterEdit: postEditFormulaContainerReport,
          afterSaveReopen: reopenedFormulaContainerReport,
        },
        editRegressions,
      },
      null,
      2,
    ),
  );
  console.log("Word document import integration passed");
  }
} catch (error) {
  if (
    error instanceof Error &&
    error.message.startsWith(diagnosticSuccessPrefix)
  ) {
    console.log(
      `Word document-import diagnostic passed ${error.message.slice(diagnosticSuccessPrefix.length)} items`,
    );
  } else {
    try {
      const wordState = runAppleScript([
        'tell application "Microsoft Word"',
        'if not (exists active document) then return "no-active-document"',
        "set documentObject to active document",
        "set unitSeparator to ASCII character 31",
        "set bookmarkNames to name of every bookmark of documentObject",
        "set documentText to content of text object of documentObject",
        "set paragraphCount to count paragraphs of documentObject",
        'return (name of documentObject as text) & unitSeparator & (paragraphCount as text) & unitSeparator & (bookmarkNames as text) & unitSeparator & documentText',
        "end tell",
      ], 15_000);
      console.error(`Word state after callback failure:\n${wordState}`);
    } catch (stateError) {
      console.error(
        `Unable to inspect Word after callback failure: ${stateError instanceof Error ? stateError.message : String(stateError)}`,
      );
    }
    if (sessionDirectory) {
    const stagePath = join(sessionDirectory, "document-import-stage.txt");
    if (existsSync(stagePath)) {
      console.error(`Last Word document-import stage:\n${readFileSync(stagePath, "utf8")}`);
    }
    const failurePath = join(sessionDirectory, "word-failure.log");
    if (existsSync(failurePath)) {
      console.error(`Word document-import failure:\n${readFileSync(failurePath, "utf8")}`);
    }
    }
    throw error;
  }
} finally {
  try {
    runAppleScript(['tell application "Microsoft Word" to quit saving no'], 20_000);
  } catch {
    // Continue with a hard process cleanup below.
  }
  spawnSync("/usr/bin/killall", ["Microsoft Word"], {
    encoding: "utf8",
    timeout: 10_000,
  });
  rmSync(pdfExportRequestPath, { force: true });
  rmSync(pdfExportStatusPath, { force: true });
  rmSync(imageEditStatusPath, { force: true });
  rmSync(formulaRegressionStatusPath, { force: true });
  rmSync(physicalScreenBoundsPath, { force: true });
  rmSync(coordinatePdfPath, { force: true });
  rmSync(installedWordAddinPath, { force: true });
  if (installedWordAddinBackedUp && existsSync(installedWordAddinBackupPath)) {
    copyFileSync(installedWordAddinBackupPath, installedWordAddinPath);
  }
  rmSync(installedWordAddinBackupPath, { force: true });
  if (sessionDirectory) rmSync(sessionDirectory, { recursive: true, force: true });
  for (const editSessionDirectory of editSessionDirectories) {
    rmSync(editSessionDirectory, { recursive: true, force: true });
  }
  rmSync(join(sessionsRoot, "word-active-session.txt"), { force: true });
  for (const nativeFile of nativeFiles) rmSync(nativeFile, { force: true });
}
