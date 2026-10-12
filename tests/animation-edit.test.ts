import { describe, expect, it } from "vitest";
import { NS } from "../src/core/utils/xml";
import { readAnimations } from "../src/domain/pptx/animation";
import { animationParts, updateAnimationPart, addAnimation, animationEntries, updateAnimation, removeAnimation, moveAnimation } from "../src/domain/pptx/animation-edit";
import { blankPptx } from "../src/infrastructure/blank";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";

const slide = (body = "") => new DOMParser().parseFromString(`<p:sld xmlns:p="${NS.p}"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="3"/></p:nvSpPr></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4"/></p:nvSpPr></p:sp></p:spTree></p:cSld>${body}</p:sld>`, "application/xml");
const serialize = (e: Node) => new XMLSerializer().serializeToString(e);

describe("animation editing", () => {
  it.each(["appear", "disappear", "fade-in", "fade-out", "fly-left", "fly-right", "fly-up", "fly-down", "fly-out-left", "fly-out-right", "fly-out-up", "fly-out-down", "zoom-in", "zoom-out", "spin"] as const)("adds a playable %s to the selected shape", (type) => {
    const doc = slide('<p:extLst/>');
    const content = serialize(doc.getElementsByTagNameNS(NS.p, "cSld")[0]);
    addAnimation(doc, "3", { type, trigger: "click", dur: 800, delay: 200 });
    const [step] = readAnimations(doc);
    expect(step.auto).toBe(false);
    const exit = type === "disappear" || type.endsWith("-out") || type.startsWith("fly-out");
    expect(step.effects[0]).toMatchObject({ spid: "3", start: 200, dur: type === "appear" || type === "disappear" ? 1 : 800, cls: exit ? "exit" : type === "spin" ? "emph" : "entr" });
    expect(animationEntries(doc)[0]).toMatchObject({ type, trigger: "click", delay: 200, simple: true });
    expect(serialize(doc.getElementsByTagNameNS(NS.p, "cSld")[0])).toBe(content);
    expect(Array.from(doc.documentElement.children).map((e) => e.localName)).toEqual(["cSld", "timing", "extLst"]);
    const ids = Array.from(doc.getElementsByTagNameNS(NS.p, "cTn")).map((e) => e.getAttribute("id"));
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("updates timing and effects, reorders and removes animations", () => {
    const doc = slide();
    addAnimation(doc, "3");
    addAnimation(doc, "4");
    updateAnimation(doc, animationEntries(doc)[1], { type: "fade-out", trigger: "after", dur: 1200, delay: 300 });
    expect(readAnimations(doc)[1]).toMatchObject({ auto: true, effects: [{ spid: "4", start: 300, dur: 1200, cls: "exit" }] });
    moveAnimation(doc, animationEntries(doc)[1], -1);
    expect(readAnimations(doc).map((s) => s.effects[0].spid)).toEqual(["4", "3"]);
    removeAnimation(doc, animationEntries(doc)[0]);
    expect(readAnimations(doc)).toHaveLength(1);
    removeAnimation(doc, animationEntries(doc)[0]);
    expect(readAnimations(doc)).toEqual([]);
  });

  it("preserves unrelated timing XML and interactive sequences", () => {
    const interactive = '<p:seq><p:cTn id="42" nodeType="interactiveSeq"><p:childTnLst/></p:cTn></p:seq>';
    const doc = slide(`<p:timing><p:tnLst>${interactive}</p:tnLst><p:bldLst/><p:extLst/></p:timing>`);
    const before = serialize(doc.getElementsByTagNameNS(NS.p, "seq")[0]);
    addAnimation(doc, "3");
    expect(serialize(doc.getElementsByTagNameNS(NS.p, "seq")[0])).toBe(before);
    expect(doc.getElementsByTagNameNS(NS.p, "bldLst")).toHaveLength(1);
    updateAnimation(doc, animationEntries(doc)[0], { dur: 1000 });
    removeAnimation(doc, animationEntries(doc)[0]);
    expect(serialize(doc.getElementsByTagNameNS(NS.p, "seq")[0])).toBe(before);
  });

  it("does not rewrite an existing step's trigger when appending another animation", () => {
    const doc = slide();
    addAnimation(doc, "3", { trigger: "after" });
    const first = animationEntries(doc)[0].el;
    const cond = first.getElementsByTagNameNS(NS.p, "cond")[0];
    cond.setAttribute("evt", "onEnd");
    const before = serialize(first);
    addAnimation(doc, "4");
    expect(serialize(first)).toBe(before);
  });

  it("retargets automatic steps after reorder/removal and retains unique timing IDs after an effect edit", () => {
    const doc = slide();
    addAnimation(doc, "3");
    addAnimation(doc, "4", { trigger: "after" });
    updateAnimation(doc, animationEntries(doc)[1], { type: "zoom-in", dur: 900 });
    const cond = () => animationEntries(doc).find((e) => e.spid === "4")!.el.getElementsByTagNameNS(NS.p, "cond")[0];
    expect(cond().getAttribute("evt")).toBe("onEnd");
    expect(cond().getElementsByTagNameNS(NS.p, "tn")[0].getAttribute("val")).toBe(animationEntries(doc)[0].el.firstElementChild!.getAttribute("id"));
    moveAnimation(doc, animationEntries(doc)[1], -1);
    expect(cond().getAttribute("evt")).toBe("onBegin");
    const ids = Array.from(doc.getElementsByTagNameNS(NS.p, "cTn")).map((e) => e.getAttribute("id"));
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("rejects nonexistent shape IDs and invalid timing without changing XML", () => {
    const doc = slide();
    const before = serialize(doc);
    expect(() => addAnimation(doc, "99")).toThrow();
    expect(() => addAnimation(doc, "3", { dur: NaN })).toThrow();
    expect(() => addAnimation(doc, "3", { delay: -100 })).toThrow();
    expect(serialize(doc)).toBe(before);
  });

  it("round trips created animations through the pptx package", async () => {
    const doc = await PptxDocument.load(await blankPptx());
    const path = doc.slides()[0].path;
    const xml = doc.xml(path)!;
    const id = xml.getElementsByTagNameNS(NS.p, "cNvPr")[1].getAttribute("id")!;
    addAnimation(xml, id, { type: "fade-in", trigger: "after", dur: 900 });
    const reloaded = await PptxDocument.load(await doc.save());
    expect(animationEntries(reloaded.xml(path)!)[0]).toMatchObject({ type: "fade-in", trigger: "after", dur: 900 });
    expect(readAnimations(reloaded.xml(path)!)[0].effects[0].behaviors.some((b) => b.type === "effect" && b.filter === "fade")).toBe(true);
  });
});


describe("advanced animation authoring", () => {
  it("starts with the previous animation in the same click step", () => {
    const doc = slide();
    addAnimation(doc, "3", { dur: 800 });
    addAnimation(doc, "4", { trigger: "with", delay: 200 });
    expect(readAnimations(doc)).toHaveLength(1);
    expect(readAnimations(doc)[0].effects.map(e => [e.spid, e.start])).toEqual([["3", 0], ["4", 200]]);
  });
  it("targets each paragraph separately", () => {
    const doc = slide();
    const shape = doc.getElementsByTagNameNS(NS.p, "sp")[0];
    const body = doc.createElementNS(NS.p, "p:txBody");
    for (let i = 0; i < 3; i++) body.appendChild(doc.createElementNS(NS.a, "a:p"));
    shape.append(body);
    addAnimation(doc, "3", { byParagraph: true });
    expect(readAnimations(doc).map(s => s.effects[0].para)).toEqual([[0, 0], [1, 1], [2, 2]]);
  });
  it("can time and remove imported compound steps without replacing their behaviors", () => {
    const doc = slide();
    addAnimation(doc, "3"); addAnimation(doc, "4", { trigger: "with" });
    const entry = animationEntries(doc)[0];
    updateAnimation(doc, entry, { dur: 1000, delay: 100 });
    expect(readAnimations(doc)[0].effects).toHaveLength(2);
    expect(readAnimations(doc)[0].effects.every(e => e.dur === 1000)).toBe(true);
    removeAnimation(doc, animationEntries(doc)[0]);
    expect(readAnimations(doc)).toEqual([]);
  });
});

it("edits one imported effect while retaining sibling effects and motion path XML", () => {
 const doc = slide(); addAnimation(doc,"3");addAnimation(doc,"4",{trigger:"with",type:"fly-left"});
 const entry=animationEntries(doc)[0]; const parts=animationParts(entry);
 const firstBefore=serialize(parts[0].el);
 updateAnimationPart(doc,parts[1],{dur:1400,delay:250});
 expect(serialize(parts[0].el)).toBe(firstBefore);
 expect(readAnimations(doc)[0].effects[1]).toMatchObject({dur:1400,start:250});
});

it("converts an existing text animation to paragraph steps at the same position",()=>{
 const doc=slide();const shape=doc.getElementsByTagNameNS(NS.p,"sp")[0];const body=doc.createElementNS(NS.p,"p:txBody");for(let i=0;i<2;i++)body.append(doc.createElementNS(NS.a,"a:p"));shape.append(body);
 addAnimation(doc,"3");addAnimation(doc,"4");updateAnimation(doc,animationEntries(doc)[0],{byParagraph:true});
 expect(readAnimations(doc).map(s=>[s.effects[0].spid,s.effects[0].para])).toEqual([["3",[0,0]],["3",[1,1]],["4",null]]);
 expect(doc.getElementsByTagNameNS(NS.p,"bldP")[0].getAttribute("build")).toBe("p");
});

it("changes imported effect duration without rebuilding custom keyframes or easing",()=>{
 const doc=slide();addAnimation(doc,"3",{type:"fly-left"});
 const anim=doc.getElementsByTagNameNS(NS.p,"anim")[0];anim.setAttribute("calcmode","fmla");
 const keys=serialize(anim.getElementsByTagNameNS(NS.p,"tavLst")[0]);
 updateAnimation(doc,animationEntries(doc)[0],{dur:1000});
 expect(doc.getElementsByTagNameNS(NS.p,"anim")[0].getAttribute("calcmode")).toBe("fmla");
 expect(serialize(doc.getElementsByTagNameNS(NS.p,"tavLst")[0])).toBe(keys);
});
