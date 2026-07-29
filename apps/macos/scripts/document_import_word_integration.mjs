import { spawnSync } from "node:child_process";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmSync,
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
const sessionsRoot = join(runtimeRoot, "OfficeSessions");
const nativeRoot = join(runtimeRoot, "NativeDocuments");
const outputKind = process.argv.includes("--image") ? "image" : "omml";
const referenceFontSizePt = 14;
const transparentPng = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAEAQH/8l0Z8QAAAABJRU5ErkJggg==",
  "base64",
);
const sleep = (ms) => new Promise((resolvePromise) => setTimeout(resolvePromise, ms));

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

function ommlForLatex(latex) {
  const superscript = latex.match(/^(.*?)([A-Za-z])\^\{?(\d+)\}?$/);
  const body = superscript
    ? `<m:r><m:t>${escapeXml(superscript[1])}</m:t></m:r>` +
      `<m:sSup><m:e><m:r><m:t>${escapeXml(superscript[2])}</m:t></m:r></m:e>` +
      `<m:sup><m:r><m:t>${escapeXml(superscript[3])}</m:t></m:r></m:sup></m:sSup>`
    : `<m:r><m:t>${escapeXml(latex)}</m:t></m:r>`;
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
  const naturalWidthPt = svg.width * 0.75;
  const naturalHeightPt = svg.height * 0.75;
  const referenceScale = Math.min(1, 500 / naturalWidthPt);
  const referenceWidthPt = naturalWidthPt * referenceScale;
  const referenceHeightPt = naturalHeightPt * referenceScale;
  const baselinePx = svg.baseline ?? svg.height;
  const descentRatio = Math.max(0, Math.min(1, (svg.height - baselinePx) / svg.height));
  const referenceBaselinePt = Math.max(
    -256,
    Math.min(0, -Math.max(0, Math.round(referenceHeightPt * descentRatio))),
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

function resolveImageRasterGeometry(rasterBands, textBoundaryCenter) {
  const numberedBand = rasterBands.find((band) =>
    band.components.some(
      (component) => component.minX > textBoundaryCenter + 100,
    ),
  );
  if (!numberedBand) {
    throw new Error(`Unable to locate numbered image formula raster band: ${JSON.stringify(rasterBands)}`);
  }
  const numberComponent = numberedBand.components.reduce((best, component) =>
    component.centerX > best.centerX ? component : best,
  );
  const numberedFormulaComponents = numberedBand.components.filter(
    (component) => component !== numberComponent,
  );
  if (!numberedFormulaComponents.length) {
    throw new Error("Numbered image formula raster band contains no formula ink");
  }
  const unnumberedCandidates = rasterBands
    .filter((band) => band !== numberedBand)
    .map((band) => ({
      band,
      bounds: rasterBounds(band, band.components),
    }))
    .filter(({ bounds }) => bounds.width >= 30)
    .sort(
      (left, right) =>
        Math.abs(left.bounds.centerX - textBoundaryCenter) -
        Math.abs(right.bounds.centerX - textBoundaryCenter),
    );
  const unnumbered = unnumberedCandidates[0]?.bounds;
  if (!unnumbered) {
    throw new Error(`Unable to locate unnumbered image formula raster band: ${JSON.stringify(rasterBands)}`);
  }
  return {
    unnumbered,
    numbered: rasterBounds(numberedBand, numberedFormulaComponents),
    equationNumber: {
      ...numberComponent,
      minY: numberedBand.minY,
      maxY: numberedBand.maxY,
      height: numberedBand.height,
      centerY: numberedBand.centerY,
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

function compactFormulaId(id) {
  return id.replaceAll("-", "");
}

function nativeBookmark(id) {
  return `VT_F_${compactFormulaId(id)}`;
}

async function waitForNewSession(before, timeoutMs = 12_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    if (existsSync(sessionsRoot)) {
      const next = readdirSync(sessionsRoot, { withFileTypes: true })
        .filter((entry) => entry.isDirectory() && !before.has(entry.name))
        .map((entry) => entry.name);
      if (next.length === 1) return next[0];
      if (next.length > 1) {
        next.sort((a, b) => {
          const aTime = readFileSync(join(sessionsRoot, a, "request.json"), "utf8");
          const bTime = readFileSync(join(sessionsRoot, b, "request.json"), "utf8");
          return bTime.length - aTime.length;
        });
        return next[0];
      }
    }
    await sleep(100);
  }
  throw new Error("Word did not create a VisualTeX document import Session");
}

function formulaItem({
  formulaId,
  latex,
  displayMode,
  numbered,
  fontSizePt,
  artifactDirectory,
}) {
  const line = { id: crypto.randomUUID(), latex };
  const omml = ommlForLatex(latex);
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

  if (outputKind === "image") {
    const svg = latexToSvg(latex, {
      displayMode: displayMode === "block",
      fontSizePt: referenceFontSizePt,
      paddingPx: displayMode === "inline" ? 1 : 10,
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
    lines: [line],
    codeFormat: "raw",
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
    latex,
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
  };
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

const before = new Set(
  existsSync(sessionsRoot)
    ? readdirSync(sessionsRoot, { withFileTypes: true })
        .filter((entry) => entry.isDirectory())
        .map((entry) => entry.name)
    : [],
);
let sessionDirectory = "";
let installedWordAddinBackedUp = false;
const nativeFiles = [];

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
  const testDocumentName = runAppleScript([
    'tell application "Microsoft Word"',
    `open file name ${JSON.stringify(templatePath)}`,
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

  const formulas = [
    formulaItem({
      formulaId: crypto.randomUUID(),
      latex: "101=202",
      displayMode: "inline",
      numbered: false,
      fontSizePt: 10.5,
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
  appendText(items, 14, "结尾文字。", {
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
    ["itemCount", "15"],
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

  runAppleScript([
    'tell application "Microsoft Word"',
    `activate object document ${JSON.stringify(testDocumentName)}`,
    'run VB macro macro name "VisualTeX_ApplyPendingResult"',
    "end tell",
  ], 90_000);

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
      'set alternativeTexts to (alternative text of inlineShapeObject) & "|" & (alternative text of displayShapeObject) & "|" & (alternative text of numberedShapeObject) & "|" & (alternative text of listFormulaShapeObject1) & "|" & (alternative text of listFormulaShapeObject2)',
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
      ? [formulas[1].latex, formulas[2].latex]
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
    if (shapeCount !== 5) {
      throw new Error(`Image import created ${shapeCount} inline shapes instead of 5`);
    }
    const metadataPayloads = alternativeText.split("|");
    if (
      metadataPayloads.length !== 5 ||
      metadataPayloads.some((value) => !value.startsWith("visualtex:v1:deflate:"))
    ) {
      throw new Error("Image formulas did not retain three independent VisualTeX metadata payloads");
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

  const sizes = [inlineSize, displaySize, numberedSize];
  const expectedSizes = [10.5, 14, 18];
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
  }

  const centerTolerancePt = outputKind === "image" ? 0.5 : 0.25;
  if (Math.abs(renderedGeometry.pageWidth - pageWidth) > 0.25) {
    throw new Error(
      `Word/PDF page width mismatch: Word=${pageWidth}, PDF=${renderedGeometry.pageWidth}`,
    );
  }
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
  if (renderedFormulaCenterDifference > centerTolerancePt) {
    throw new Error(
      `Numbering shifted the formula center: difference=${renderedFormulaCenterDifference} pt`,
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
  const equationNumberVerticalError = Math.abs(
    equationNumberGeometry.centerY - numberedFormulaPdfCenterY,
  );
  if (equationNumberVerticalError > 0.25) {
    throw new Error(
      `Equation number is not vertically centered with its formula: error=${equationNumberVerticalError} pt`,
    );
  }
  const measuredNumberedRight =
    outputKind === "omml"
      ? renderedGeometry.numbered.maxX
      : imageRasterGeometry.numbered.maxX;
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
          ...(outputKind === "omml"
            ? { bookmark: nativeBookmark(formula.formulaId) }
            : { inlineShapeIndex: index + 1 }),
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
                    centerError: renderedNumberedCenterError,
                  },
            equationNumber: {
              ...equationNumberGeometry,
              verticalCenterError: equationNumberVerticalError,
              rightBoundaryInset: textBoundaryRight - equationNumberGeometry.maxX,
            },
            formulaCenterDifference: renderedFormulaCenterDifference,
            centerTolerancePt,
          },
        },
        documentText,
      },
      null,
      2,
    ),
  );
  console.log("Word document import integration passed");
} catch (error) {
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
  rmSync(coordinatePdfPath, { force: true });
  if (installedWordAddinBackedUp && existsSync(installedWordAddinBackupPath)) {
    copyFileSync(installedWordAddinBackupPath, installedWordAddinPath);
  }
  rmSync(installedWordAddinBackupPath, { force: true });
  if (sessionDirectory) rmSync(sessionDirectory, { recursive: true, force: true });
  rmSync(join(sessionsRoot, "word-active-session.txt"), { force: true });
  for (const nativeFile of nativeFiles) rmSync(nativeFile, { force: true });
}
