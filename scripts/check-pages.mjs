#!/usr/bin/env node
/**
 * Page-count regression check in the running Obsidian (the page view needs a real layout engine):
 * opens every .docx under a vault folder in a background tab and compares
 *   - the page count with the one Word saved (docProps/app.xml), when there is one;
 *   - every drawn page's height with the paper height.
 * Exits 1 if any file has more pages than Word made, or a page of the wrong height.
 *
 *   npm run check:pages -- [vault folder]      (default: the whole vault; Obsidian must be running)
 */
import { execFileSync } from "node:child_process";

const folder = process.argv[2] ?? "";
const run = (code) => {
  const out = execFileSync("obsidian", ["eval", `code=${code}`], { encoding: "utf8", stdio: ["ignore", "pipe", "ignore"] });
  return out.split("\n").find((l) => l.startsWith("=> "))?.slice(3) ?? "";
};
const js = `window.__dwPages="run";(async()=>{try{
const files=app.vault.getFiles().filter(f=>f.extension==="docx"&&f.path.startsWith(${JSON.stringify(folder)})&&!f.path.startsWith(".trash"));const res=[];
for(const f of files){const leaf=app.workspace.getLeaf("tab");await leaf.openFile(f,{active:false});
for(let i=0;i<60;i++){await new Promise(r=>setTimeout(r,250));if(leaf.view.editor&&leaf.view.editor.dom.querySelector(".docx-page-first"))break;}
await new Promise(r=>setTimeout(r,1500));const v=leaf.view;const dom=v.editor?.dom;
if(!dom){res.push([f.path,null,null,[]]);leaf.detach();continue;}
const z=parseFloat(dom.style.zoom)||1;const top=dom.getBoundingClientRect().top;
const sp=[...dom.querySelectorAll(".docx-page-space")].map(e=>e.getBoundingClientRect()).sort((a,b)=>a.top-b.top);
let start=0;const wrong=[];sp.forEach((r,k)=>{const h=(r.top-top)/z-start;start=(r.bottom-top)/z;if(Math.abs(h-v.setup.height)>1.5)wrong.push([k+1,Math.round(h)]);});
res.push([f.path,v.docx.wordPages(),sp.length+1,wrong]);leaf.detach();}
window.__dwPages=JSON.stringify(res);}catch(e){window.__dwPages=JSON.stringify({error:String(e)})}})()`;
run(js);
let result = "run";
for (let i = 0; i < 2400 && result === "run"; i++) {
  await new Promise((r) => setTimeout(r, 500));
  result = run("window.__dwPages") || "run"; // eval output is sometimes lost
}
const rows = JSON.parse(result);
if (rows.error) throw new Error(rows.error);
let bad = 0;
for (const [path, word, ours, wrong] of rows) {
  const over = word !== null && ours > word;
  if (over || wrong.length || ours === null) bad++;
  const note = [over ? `${ours - word} page(s) more than Word` : "", wrong.length ? `wrong height: ${wrong.map(([p, h]) => `p${p}=${h}px`).join(" ")}` : "", ours === null ? "didn't open" : ""].filter(Boolean).join("; ");
  console.log(`${note ? "FAIL" : "ok  "}  Word ${word ?? "-"}  here ${ours ?? "-"}  ${path}${note ? `  (${note})` : ""}`);
}
console.log(`${rows.length} documents, ${bad} with problems`);
process.exit(bad ? 1 : 0);
