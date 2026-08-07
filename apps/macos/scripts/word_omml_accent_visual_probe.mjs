import { spawnSync } from "node:child_process";
import { writeFileSync, readFileSync, existsSync, rmSync, mkdirSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { zipSync, strToU8 } from "fflate";

const scratch = join(
  homedir(),
  "Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch",
);
const runtime = join(
  homedir(),
  "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime",
);
const docxPath = join(scratch, "word-omml-accent-visual-probe.docx");
const pdfPath = join(scratch, "word-omml-accent-visual-probe.pdf");
const requestPath = join(runtime, "document-import-regression-pdf-path.txt");
const statusPath = join(runtime, "document-import-regression-pdf-status.txt");
mkdirSync(scratch, { recursive: true });
mkdirSync(runtime, { recursive: true });

function escapeXml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

const runProperties =
  '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>'; 
const controlProperties = `<m:ctrlPr>${runProperties}</m:ctrlPr>`;
const mathRun = (text) =>
  `<m:r>${runProperties}<m:t>${escapeXml(text)}</m:t></m:r>`;
const subscript = (base, sub) =>
  `<m:sSub><m:sSubPr>${controlProperties}</m:sSubPr>` +
  `<m:e>${base}</m:e><m:sub>${mathRun(sub)}</m:sub></m:sSub>`;
const bareSubscript = (base, sub) =>
  `<m:sSub><m:e>${base}</m:e><m:sub>${mathRun(sub)}</m:sub></m:sSub>`;
const accent = (character, base) =>
  `<m:acc><m:accPr><m:chr m:val="${escapeXml(character)}"/>${controlProperties}</m:accPr>` +
  `<m:e>${base}</m:e></m:acc>`;
const bareAccent = (character, base) =>
  `<m:acc><m:accPr><m:chr m:val="${escapeXml(character)}"/></m:accPr>` +
  `<m:e>${base}</m:e></m:acc>`;
const groupCharacter = (character, base) =>
  `<m:groupChr><m:groupChrPr><m:chr m:val="${escapeXml(character)}"/>` +
  '<m:pos m:val="top"/><m:vertJc m:val="bot"/>' +
  `${controlProperties}</m:groupChrPr><m:e>${base}</m:e></m:groupChr>`;
const paragraph = (body) =>
  '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="80" w:after="160"/></w:pPr>' +
  `<m:oMathPara><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr>` +
  `<m:oMath>${body}</m:oMath></m:oMathPara></w:p>`;

const formulas = [
  subscript(accent("^", mathRun("L")), "z"),
  subscript(accent("ˆ", mathRun("L")), "z"),
  subscript(accent("̂", mathRun("L")), "z"),
  subscript(groupCharacter("^", mathRun("L")), "z"),
  accent("→", mathRun("r")),
  groupCharacter("→", mathRun("r")),
  groupCharacter("→", subscript(mathRun("e"), "φ")),
];

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
const packageRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:body>
    ${formulas.map(paragraph).join("\n")}
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>
  </w:body>
</w:document>`;
const packageBytes = zipSync(
  {
    "[Content_Types].xml": strToU8(contentTypes),
    "_rels/.rels": strToU8(packageRelationships),
    "word/document.xml": strToU8(documentXml),
  },
  { level: 6 },
);
writeFileSync(docxPath, packageBytes);
rmSync(pdfPath, { force: true });
rmSync(statusPath, { force: true });
writeFileSync(requestPath, pdfPath, { mode: 0o600 });

function appleScript(lines) {
  const result = spawnSync(
    "/usr/bin/osascript",
    lines.flatMap((line) => ["-e", line]),
    { encoding: "utf8", timeout: 90_000, maxBuffer: 4 * 1024 * 1024 },
  );
  if (result.status !== 0) {
    throw new Error(result.stderr.trim() || result.stdout.trim() || "AppleScript failed");
  }
  return result.stdout.trim();
}

appleScript([
  'tell application "Microsoft Word"',
  `open POSIX file ${JSON.stringify(docxPath)}`,
  'set probeDocument to document "word-omml-accent-visual-probe.docx"',
  "activate object probeDocument",
  "activate",
  "set font size of font object of text object of probeDocument to 36",
  'run VB macro macro name "VisualTeX_ExportActiveDocumentPdfForRegression"',
  "close probeDocument saving no",
  'if exists document "文档1" then activate object document "文档1"',
  "end tell",
]);

if (!existsSync(statusPath)) {
  throw new Error("Word did not write the PDF export status");
}
const status = readFileSync(statusPath, "utf8").trim();
if (!status.startsWith("ok|") || !existsSync(pdfPath)) {
  throw new Error(`Word PDF export failed: ${status}`);
}
console.log(JSON.stringify({ docxPath, pdfPath, status }, null, 2));
