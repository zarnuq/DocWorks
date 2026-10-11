import { describe, it, expect } from "vitest";
import { evalFormula, readAnimations, readTransition, shapeIds } from "../src/domain/pptx/animation";

const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const slide = (inner: string) =>
  new DOMParser().parseFromString(
    `<p:sld xmlns:p="${P}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:cSld><p:spTree/></p:cSld>${inner}</p:sld>`,
    "application/xml"
  );

const tgt = (spid: string, para?: number) => `<p:tgtEl><p:spTgt spid="${spid}">${para === undefined ? "" : `<p:txEl><p:pRg st="${para}" end="${para}"/></p:txEl>`}</p:spTgt></p:tgtEl>`;
const set = (spid: string, val: string, delay = 0, para?: number) =>
  `<p:set><p:cBhvr><p:cTn dur="1" fill="hold"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst></p:cTn>${tgt(spid, para)}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${val}"/></p:to></p:set>`;
const fade = (spid: string, dir = "in", dur = 500) => `<p:animEffect transition="${dir}" filter="fade"><p:cBhvr><p:cTn dur="${dur}"/>${tgt(spid)}</p:cBhvr></p:animEffect>`;
const effect = (cls: string, node: string, body: string, delay = 0) =>
  `<p:par><p:cTn presetID="10" presetClass="${cls}" presetSubtype="0" fill="hold" nodeType="${node}"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst><p:childTnLst>${body}</p:childTnLst></p:cTn></p:par>`;
const group = (delay: number, ...effects: string[]) => `<p:par><p:cTn fill="hold"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst><p:childTnLst>${effects.join("")}</p:childTnLst></p:cTn></p:par>`;
const click = (conds: string, ...groups: string[]) => `<p:par><p:cTn fill="hold"><p:stCondLst>${conds}</p:stCondLst><p:childTnLst>${groups.join("")}</p:childTnLst></p:cTn></p:par>`;
const ON_CLICK = `<p:cond delay="indefinite"/>`;
const ON_BEGIN = `<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>`;
const timing = (...clicks: string[]) =>
  `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${clicks.join("")}</p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;

describe("readAnimations", () => {
  it("returns no steps for a slide without animations", () => {
    expect(readAnimations(slide(""))).toEqual([]);
  });

  it("reads one step per click with its effects' targets, class and timing", () => {
    const xml = timing(
      click(ON_CLICK, group(0, effect("entr", "clickEffect", set("3", "visible") + fade("3")))),
      click(ON_CLICK, group(0, effect("exit", "clickEffect", fade("3", "out", 700) + set("3", "hidden", 699))))
    );
    const steps = readAnimations(slide(xml));
    expect(steps).toHaveLength(2);
    expect(steps[0].auto).toBe(false);
    expect(steps[0].effects[0]).toMatchObject({ spid: "3", para: null, cls: "entr", start: 0, dur: 500 });
    expect(steps[1].effects[0]).toMatchObject({ cls: "exit", dur: 700 });
    expect(steps[1].effects[0].behaviors).toEqual([
      { type: "effect", filter: "fade", in: false, start: 0, dur: 700 },
      { type: "visibility", visible: false, start: 699, dur: 1 },
    ]);
  });

  it("adds group and effect delays for with-previous and after-previous effects", () => {
    const xml = timing(
      click(ON_CLICK, group(0, effect("entr", "clickEffect", fade("3")), effect("entr", "withEffect", fade("4"), 250)), group(500, effect("entr", "afterEffect", fade("5"))))
    );
    const [step] = readAnimations(slide(xml));
    expect(step.effects.map((e) => [e.spid, e.start])).toEqual([
      ["3", 0],
      ["4", 250],
      ["5", 500],
    ]);
  });

  it("marks a step that starts with the slide as automatic", () => {
    const steps = readAnimations(slide(timing(click(ON_BEGIN, group(0, effect("entr", "afterEffect", fade("3")))))));
    expect(steps[0].auto).toBe(true);
  });

  it("targets paragraphs of a shape (text by paragraph)", () => {
    const xml = timing(click(ON_CLICK, group(0, effect("entr", "clickEffect", set("3", "visible", 0, 2)))));
    expect(readAnimations(slide(xml))[0].effects[0]).toMatchObject({ spid: "3", para: [2, 2] });
  });

  it("reads motion, scale, rotation and property animations", () => {
    const anim = `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn dur="500" fill="hold"/>${tgt("3")}<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="1+#ppt_h/2"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav></p:tavLst></p:anim>`;
    const scale = `<p:animScale><p:cBhvr><p:cTn dur="250" autoRev="1"/>${tgt("3")}</p:cBhvr><p:by x="150000" y="150000"/></p:animScale>`;
    const rot = `<p:animRot by="21600000"><p:cBhvr><p:cTn dur="2000"/>${tgt("3")}<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>`;
    const motion = `<p:animMotion origin="layout" path="M 0 0 L 0.25 0.1 E" pathEditMode="relative"><p:cBhvr><p:cTn dur="2000"/>${tgt("3")}</p:cBhvr></p:animMotion>`;
    const [step] = readAnimations(slide(timing(click(ON_CLICK, group(0, effect("emph", "clickEffect", anim + scale + rot + motion))))));
    const b = step.effects[0].behaviors;
    expect(b[0]).toEqual({ type: "prop", attr: "ppt_y", keys: [{ t: 0, v: "1+#ppt_h/2" }, { t: 1, v: "#ppt_y" }], start: 0, dur: 500 });
    expect(b[1]).toEqual({ type: "scale", by: { x: 1.5, y: 1.5 }, autoRev: true, start: 0, dur: 250 });
    expect(b[2]).toEqual({ type: "rotate", by: 360, start: 0, dur: 2000 });
    expect(b[3]).toEqual({ type: "motion", points: [{ x: 0, y: 0 }, { x: 0.25, y: 0.1 }], start: 0, dur: 2000 });
    // autoRev plays twice as long.
    expect(step.effects[0].dur).toBe(2000);
  });

  it("ignores interactive (trigger) sequences", () => {
    const xml = `<p:timing><p:tnLst><p:par><p:cTn nodeType="tmRoot"><p:childTnLst><p:seq><p:cTn nodeType="interactiveSeq"><p:childTnLst>${click(ON_CLICK, group(0, effect("entr", "clickEffect", fade("3"))))}</p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
    expect(readAnimations(slide(xml))).toEqual([]);
  });
});

describe("readTransition", () => {
  it("returns null without a transition", () => {
    expect(readTransition(slide(""))).toBeNull();
  });

  it("reads the type, direction and speed", () => {
    expect(readTransition(slide(`<p:transition spd="slow"><p:push dir="u"/></p:transition>`))).toEqual({ type: "push", dir: "u", dur: 1000, advanceAfter: null });
    expect(readTransition(slide(`<p:transition><p:fade/></p:transition>`))).toEqual({ type: "fade", dir: null, dur: 500, advanceAfter: null });
  });

  it("reads the automatic advance time", () => {
    expect(readTransition(slide(`<p:transition advTm="3000"/>`))).toEqual({ type: "cut", dir: null, dur: 0, advanceAfter: 3000 });
  });

  it("prefers PowerPoint 2010's own transition and duration over the fallback", () => {
    const xml = `<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="slow" p14:dur="1600"><p14:vortex dir="r"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>`;
    expect(readTransition(slide(xml))).toEqual({ type: "vortex", dir: "r", dur: 1600, advanceAfter: null });
  });

  it("skips the sound and extension children", () => {
    expect(readTransition(slide(`<p:transition><p:sndAc><p:stSnd/></p:sndAc></p:transition>`))?.type).toBe("cut");
  });
});

describe("evalFormula", () => {
  const vars = { ppt_x: 0.5, ppt_y: 0.25, ppt_w: 0.2, ppt_h: 0.1 };
  it("evaluates PowerPoint's position formulas", () => {
    expect(evalFormula("#ppt_x", vars)).toBe(0.5);
    expect(evalFormula("1+#ppt_h/2", vars)).toBeCloseTo(1.05);
    expect(evalFormula("0-#ppt_w/2", vars)).toBeCloseTo(-0.1);
    expect(evalFormula("(#ppt_x+1)*2", vars)).toBeCloseTo(3);
    expect(evalFormula("0.5", vars)).toBe(0.5);
  });
  it("returns null for what it can't read", () => {
    expect(evalFormula("sin(#ppt_x)", vars)).toBeNull();
    expect(evalFormula("", vars)).toBeNull();
  });
});

describe("shapeIds", () => {
  it("lists a shape's id and the ids of the groups around it", () => {
    const doc = new DOMParser().parseFromString(
      `<p:spTree xmlns:p="${P}"><p:grpSp><p:nvGrpSpPr><p:cNvPr id="7" name="g"/></p:nvGrpSpPr><p:sp><p:nvSpPr><p:cNvPr id="9" name="s"/></p:nvSpPr></p:sp></p:grpSp></p:spTree>`,
      "application/xml"
    );
    expect(shapeIds(doc.getElementsByTagNameNS(P, "sp")[0])).toEqual(["9", "7"]);
  });
});
