import { spawn } from "node:child_process";
import process from "node:process";
import { rm } from "node:fs/promises";
import { existsSync } from "node:fs";

const port = 19600 + process.pid % 500;
const debugPort = 24600 + process.pid % 500;
const profile = `/tmp/visualtex-mathtype-ui-${process.pid}`;
const chrome = ["/Applications/Google Chrome.app/Contents/MacOS/Google Chrome", "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"];
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
async function wait(url) { for (let i=0;i<150;i++) { try { const r=await fetch(url); if(r.ok) return; } catch {} await sleep(100); } throw new Error(`timeout ${url}`); }
class Cdp { constructor(url){this.url=url;this.id=0;this.pending=new Map()} async connect(){this.ws=new WebSocket(this.url);await new Promise((r,j)=>{this.ws.onopen=r;this.ws.onerror=j});this.ws.onmessage=(e)=>{const m=JSON.parse(e.data);if(m.id){const p=this.pending.get(m.id);this.pending.delete(m.id);m.error?p.j(new Error(m.error.message)):p.r(m.result)}}} send(method,params={}){const id=++this.id;return new Promise((r,j)=>{this.pending.set(id,{r,j});this.ws.send(JSON.stringify({id,method,params}))})}}
const server=spawn(process.execPath,["node_modules/vite/bin/vite.js","preview","--host","127.0.0.1","--port",String(port),"--strictPort"],{stdio:"ignore"});
let browser;
try {
  await wait(`http://127.0.0.1:${port}`);
  const executable=chrome.find(existsSync); if(!executable) throw new Error("Chrome or Edge required");
  browser=spawn(executable,["--headless=new","--disable-gpu","--no-first-run",`--remote-debugging-port=${debugPort}`,`--user-data-dir=${profile}`,"about:blank"],{stdio:"ignore"});
  await wait(`http://127.0.0.1:${debugPort}/json/list`); const targets=await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json(); const c=new Cdp(targets.find(x=>x.type==="page").webSocketDebuggerUrl); await c.connect();
  await c.send("Runtime.enable"); await c.send("Page.enable");
  await c.send("Page.addScriptToEvaluateOnNewDocument",{source:`(()=>{let callbackId=1;const callbacks=new Map();const job={protocolVersion:1,jobId:"12345678-1234-4234-9234-123456789abc",inputPath:"/tmp/Công thức.docx",outputPath:"/tmp/Công thức_VisualTeX_OMML.docx",inputSha256:"a".repeat(64),status:"awaitingOmml",createdAtEpochSeconds:1,updatedAtEpochSeconds:1,error:null,scanReport:{protocolVersion:1,jobId:"12345678-1234-4234-9234-123456789abc",detected:0,batchCount:0,batchSize:100,byProgId:{"Equation.DSMT4":0,"Equation.3":0,"Equation.2":0,unknown:0},byRiskLevel:{"auto-replace":0,"spot-check":0,"manual-review":0,blocked:0}}};window.__LEGACY_CALLS__=[];window.__TAURI_INTERNALS__={metadata:{currentWindow:{label:"main"},currentWebview:{label:"main"}},transformCallback(callback,once=false){const id=callbackId++;callbacks.set(id,{callback,once});return id},unregisterCallback(id){callbacks.delete(id)},async invoke(command,args){window.__LEGACY_CALLS__.push({command,args});if(command.includes("dialog|open"))return "/tmp/Công thức.docx";if(command.includes("dialog|save"))return "/tmp/Công thức_VisualTeX_OMML.docx";if(command==="create_legacy_equation_job"||command==="get_legacy_equation_job")return job;if(command==="finalize_legacy_equation_job")return {...job,status:"complete",conversionReport:{detected:0,replaced:0,preserved:0,failed:0,skipped:0}};if(command==="cancel_legacy_equation_job")return {...job,status:"cancelled"};if(command.includes("plugin:event|listen")||command.includes("plugin:event|unlisten"))return 1;return null}}})()`});
  await c.send("Page.navigate",{url:`http://127.0.0.1:${port}/?view=legacy-equation-converter`}); await sleep(1200);
  const result=await c.send("Runtime.evaluate",{expression:`(()=>{const root=document.querySelector('[data-testid="legacy-equation-converter"]');const progress=document.querySelector('[role="progressbar"]');const buttons=[...document.querySelectorAll('button')];return {root:!!root,progress:progress?.getAttribute('aria-valuenow'),vi:document.body.textContent.includes('Chuyển công thức MathType cũ'),warning:document.body.textContent.includes('chưa có corpus MTEF v3 thực'),keyboard:buttons.every(b=>b.tabIndex>=0),html:!!document.querySelector('[dangerouslySetInnerHTML]'),text:document.body.textContent.slice(0,500)}})()`,returnByValue:true});
  const value=result.result.value; const localized=value.vi||value.text.includes("Convert legacy MathType equations"); const warning=value.warning||value.text.includes("real MTEF v3 corpus"); if(!value.root||value.progress!=="0"||!localized||!warning||!value.keyboard||value.html) throw new Error(JSON.stringify(value));
  await c.send("Runtime.evaluate",{expression:`document.querySelectorAll('button')[0].click()`,awaitPromise:true}); await sleep(150);
  await c.send("Runtime.evaluate",{expression:`document.querySelector('[data-testid="scan"]').click();document.querySelector('[data-testid="scan"]').click()`,awaitPromise:true}); await sleep(700);
  const review=await c.send("Runtime.evaluate",{expression:`({creates:window.__LEGACY_CALLS__.filter(x=>x.command==='create_legacy_equation_job').length,convert:!!document.querySelector('[data-testid="convert"]'),unicode:document.body.textContent.includes('Công thức.docx')})`,returnByValue:true});
  if(review.result.value.creates!==1||!review.result.value.convert||!review.result.value.unicode) throw new Error(JSON.stringify(review.result.value));
  await c.send("Runtime.evaluate",{expression:`document.querySelector('[data-testid="convert"]').click()`}); await sleep(500);
  const complete=await c.send("Runtime.evaluate",{expression:`({finalizes:window.__LEGACY_CALLS__.filter(x=>x.command==='finalize_legacy_equation_job').length,complete:document.querySelector('[role="progressbar"]').getAttribute('aria-valuenow')==='100'})`,returnByValue:true});
  if(complete.result.value.finalizes!==1||!complete.result.value.complete) throw new Error(JSON.stringify(complete.result.value));
  console.log("Legacy equation converter browser regression passed (route, zero-formula job, double-click guard, Unicode, locale, warning, keyboard, progress).")
} finally { server.kill("SIGTERM"); browser?.kill("SIGTERM"); await sleep(400); await rm(profile,{recursive:true,force:true}).catch(async()=>{await sleep(400);await rm(profile,{recursive:true,force:true})}); }
