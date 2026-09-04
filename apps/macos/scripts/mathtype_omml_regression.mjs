import { createHash } from "node:crypto";
import { spawn } from "node:child_process";
import { access, readFile, rm } from "node:fs/promises";
import { resolve } from "node:path";
import process from "node:process";

const fixturePath = resolve("scripts/fixtures/mathtype-omml/mtef-v5-contract.json");
const sourceFixtureRoot = resolve("../../tools/mathtypejx/tests/fixtures");
const contract = JSON.parse(await readFile(fixturePath, "utf8"));
for (const fixture of contract.fixtures) {
  const digest = createHash("sha256")
    .update(await readFile(resolve(sourceFixtureRoot, fixture.ole)))
    .digest("hex");
  if (digest !== fixture.sha256) throw new Error(`${fixture.ole} no longer matches its MathML contract.`);
  if (fixture.mtefVersion !== 5) throw new Error("The checked-in contract must not claim unverified MTEF v3 coverage.");
}

const portOffset = process.pid % 1000;
const vitePort = 7400 + portOffset;
const debugPort = 12400 + portOffset;
const baseUrl = `http://127.0.0.1:${vitePort}`;
const profile = `/tmp/visualtex-mathtype-omml-${process.pid}`;
const browserCandidates = [
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
  "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
];
let chromePath;
for (const candidate of browserCandidates) {
  try { await access(candidate); chromePath = candidate; break; } catch {}
}
if (!chromePath) throw new Error("MathType OMML regression requires Google Chrome or Microsoft Edge.");
const sleep = (ms) => new Promise((resolvePromise) => setTimeout(resolvePromise, ms));

async function waitFor(url) {
  const deadline = Date.now() + 20_000;
  while (Date.now() < deadline) {
    try { const response = await fetch(url); if (response.ok) return response; } catch {}
    await sleep(100);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

class CdpClient {
  constructor(url) { this.url = url; this.id = 0; this.pending = new Map(); }
  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise((ok, fail) => { this.socket.addEventListener("open", ok, { once: true }); this.socket.addEventListener("error", fail, { once: true }); });
    this.socket.addEventListener("message", (event) => {
      const message = JSON.parse(event.data); const pending = this.pending.get(message.id); if (!pending) return;
      this.pending.delete(message.id); message.error ? pending.reject(new Error(message.error.message)) : pending.resolve(message.result);
    });
  }
  send(method, params = {}) { const id = ++this.id; return new Promise((resolvePromise, reject) => { this.pending.set(id, { resolve: resolvePromise, reject }); this.socket.send(JSON.stringify({ id, method, params })); }); }
  close() { this.socket?.close(); }
}

const structuralCorpus = [
  ["indexed-root", `<math xmlns="http://www.w3.org/1998/Math/MathML"><mroot><mi>x</mi><mn>3</mn></mroot></math>`, ["<m:rad>", "<m:deg>"]],
  ["scripts", `<math xmlns="http://www.w3.org/1998/Math/MathML"><msubsup><mi>x</mi><mi>i</mi><mn>2</mn></msubsup></math>`, ["<m:sSubSup>"]],
  ["matrix", `<math xmlns="http://www.w3.org/1998/Math/MathML"><mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr><mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable></math>`, ["<m:m>", "<m:mr>"]],
  ["delimiters", `<math xmlns="http://www.w3.org/1998/Math/MathML"><mrow><mo>{</mo><mi>x</mi><mo></mo></mrow></math>`, ["<m:d>", `m:begChr m:val="{"`, `m:endChr m:val=""`]],
  ["accents", `<math xmlns="http://www.w3.org/1998/Math/MathML"><mrow><mover><mi>x</mi><mo>¯</mo></mover><mover><mi>v</mi><mo>→</mo></mover></mrow></math>`, ["<m:bar>", "<m:groupChr>"]],
  ["nary-limits", `<math xmlns="http://www.w3.org/1998/Math/MathML"><mrow><munderover><mo>∑</mo><mi>i</mi><mi>n</mi></munderover><mi>x</mi></mrow></math>`, ["<m:nary>", "<m:sub>", "<m:sup>"]],
  ["prescripts", `<math xmlns="http://www.w3.org/1998/Math/MathML"><mmultiscripts><mi>C</mi><none/><none/><mprescripts/><mn>6</mn><mn>14</mn></mmultiscripts></math>`, ["<m:sPre>", "<m:t>6</m:t>", "<m:t>14</m:t>"]],
  ["text-unicode", `<math xmlns="http://www.w3.org/1998/Math/MathML"><mrow><mtext>kg·m/s² 中文 Việt</mtext><mo>+</mo><mi>α</mi><mo>+</mo><mi>𝑥</mi></mrow></math>`, ["<m:nor/>", "中文 Việt", "α", "𝑥"]],
  ["boxed", `<math xmlns="http://www.w3.org/1998/Math/MathML"><menclose notation="box"><mi>x</mi></menclose></math>`, ["<m:borderBox>"]],
];

let vite; let chrome; let client;
try {
  vite = spawn(process.execPath, ["node_modules/vite/bin/vite.js", "--host", "127.0.0.1", "--port", String(vitePort), "--strictPort"], { cwd: process.cwd(), stdio: "ignore" });
  await waitFor(baseUrl);
  chrome = spawn(chromePath, ["--headless=new", "--disable-gpu", "--no-first-run", `--remote-debugging-port=${debugPort}`, `--user-data-dir=${profile}`, baseUrl], { stdio: "ignore" });
  await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
  const pages = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
  client = new CdpClient(pages.find((page) => page.type === "page")?.webSocketDebuggerUrl);
  await client.connect(); await client.send("Runtime.enable");
  const expression = `(async()=>{const m=await import(${JSON.stringify(`${baseUrl}/src/office/omml/latexToOmml.ts`)});const v=await import(${JSON.stringify(`${baseUrl}/src/office/omml/mathMlOmmlValidator.ts`)});const inputs=${JSON.stringify([...contract.fixtures.map((fixture) => [fixture.ole, fixture.mathMl, []]), ...structuralCorpus])};const results=inputs.map(([name,mathMl,expected])=>{const block=m.mathMlToOmmlArtifacts(mathMl,'block');const inline=m.mathMlToOmmlArtifacts(mathMl,'inline');return{name,expected,block,inline,validation:v.validateMathMlToOmml(mathMl,block.omml)}});let malformedBlocked=false,tokenLossBlocked=false,dtdBlocked=false;try{m.mathMlToOmmlArtifacts('<math><mfrac/></math>','block')}catch{malformedBlocked=true}try{m.mathMlToOmmlArtifacts('<!DOCTYPE math [<!ENTITY x "unsafe">]><math><mi>&x;</mi></math>','block')}catch{dtdBlocked=true}try{v.assertValidMathMlToOmml('<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>x</mi></math>','<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>y</m:t></m:r></m:oMath>')}catch{tokenLossBlocked=true}return{results,malformedBlocked,tokenLossBlocked,dtdBlocked}})()`;
  const evaluation = await client.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
  if (evaluation.exceptionDetails) throw new Error(evaluation.exceptionDetails.exception?.description ?? evaluation.exceptionDetails.text);
  const value = evaluation.result.value;
  if (!value.malformedBlocked || !value.tokenLossBlocked || !value.dtdBlocked) throw new Error("Fail-closed validator checks did not block invalid conversion.");
  for (const result of value.results) {
    if (!result.validation.valid) throw new Error(`${result.name}: ${result.validation.errors.join(" ")}`);
    if (result.block.omml !== result.inline.omml) throw new Error(`${result.name}: inline/block semantic output diverged.`);
    for (const fragment of result.expected) if (!result.block.omml.includes(fragment)) throw new Error(`${result.name}: missing ${fragment}`);
    if (!result.block.ommlBase64 || !result.block.ommlDocxBase64) throw new Error(`${result.name}: incomplete artifacts.`);
  }
  console.log(`MathType MathML-to-OMML contract passed (${value.results.length} cases; 3 real MTEF v5 fixtures).`);
} finally {
  client?.close(); chrome?.kill("SIGTERM"); vite?.kill("SIGTERM"); await rm(profile, { recursive: true, force: true });
}
