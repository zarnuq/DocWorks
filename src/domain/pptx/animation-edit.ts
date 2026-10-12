/** Minimal XML edits for the main slide animation sequence. Imported complex steps retain their behaviors. */
import { NS } from "../../core/utils/xml";
import { Effect, readAnimations } from "./animation";

export const ANIMATION_TYPES = [
  ["Appear", "appear"], ["Disappear", "disappear"], ["Fade in", "fade-in"], ["Fade out", "fade-out"],
  ["Fly in from left", "fly-left"], ["Fly in from right", "fly-right"],
  ["Fly in from top", "fly-up"], ["Fly in from bottom", "fly-down"],
  ["Fly out to left", "fly-out-left"], ["Fly out to right", "fly-out-right"],
  ["Fly out to top", "fly-out-up"], ["Fly out to bottom", "fly-out-down"],
  ["Zoom in", "zoom-in"], ["Zoom out", "zoom-out"], ["Spin", "spin"],
] as const;
export type AnimationType = typeof ANIMATION_TYPES[number][1];
export interface AnimationSettings {
  type: AnimationType;
  trigger: "click" | "after" | "with";
  byParagraph?: boolean;
  paragraph?: number;
  dur: number;
  delay: number;
}
export interface AnimationEntry extends AnimationSettings {
  /** Whole click/automatic step; individual imported effects are available through animationParts. */
  el: Element;
  spid: string;
  editable: boolean;
  simple: boolean;
  label: string;
  targetName: string;
}

const child = (el: Element | null, name: string) => Array.from(el?.children ?? []).find((e) => e.namespaceURI === NS.p && e.localName === name) ?? null;
const mainOf = (doc: Document) => Array.from(doc.getElementsByTagNameNS(NS.p, "cTn")).find((e) => e.getAttribute("nodeType") === "mainSeq") ?? null;
const defaults: AnimationSettings = { type: "fade-in", trigger: "click", dur: 500, delay: 0 };
const validate = (s: AnimationSettings) => {
  if (!ANIMATION_TYPES.some(([, type]) => type === s.type)) throw new Error("Unsupported animation");
  if (s.trigger !== "click" && s.trigger !== "after" && s.trigger !== "with") throw new Error("Unsupported animation trigger");
  if (!Number.isFinite(s.dur) || s.dur < 1 || !Number.isFinite(s.delay) || s.delay < 0) throw new Error("Invalid animation timing");
};

function kindOf(e: Effect): AnimationType | null {
  const exit = e.cls === "exit";
  if (e.behaviors.length === 1 && e.behaviors[0].type === "visibility") return e.behaviors[0].visible ? "appear" : "disappear";
  const filter = e.behaviors.find((b) => b.type === "effect");
  if (filter?.type === "effect" && filter.filter === "fade") return exit ? "fade-out" : "fade-in";
  if (e.behaviors.some((b) => b.type === "rotate" && b.by === 360)) return "spin";
  const props = e.behaviors.filter((b) => b.type === "prop");
  if (props.some((b) => b.attr === "ppt_w") && props.some((b) => b.attr === "ppt_h")) return exit ? "zoom-out" : "zoom-in";
  const p = props[0];
  if (p && ["ppt_x", "ppt_y"].includes(p.attr)) {
    // Flying in starts off the slide; flying out ends there.
    const off = p.keys[exit ? p.keys.length - 1 : 0]?.v;
    const side = off === "0-#ppt_w/2" ? "left" : off === "1+#ppt_w/2" ? "right" : off === "0-#ppt_h/2" ? "up" : off === "1+#ppt_h/2" ? "down" : null;
    if (side) return `fly-${exit ? "out-" : ""}${side}` as AnimationType;
  }
  return null;
}

/** One row per step; compound steps expose individual effect controls without replacing behaviors. */
export function animationEntries(doc: Document): AnimationEntry[] {
  const main = mainOf(doc);
  const clicks = Array.from(child(main, "childTnLst")?.children ?? []).filter((e) => e.localName === "par");
  return readAnimations(doc).map((step, i) => {
    const e = step.effects[0];
    const type = e && kindOf(e);
    const click = clicks[i];
    const groups = Array.from(child(child(click, "cTn"), "childTnLst")?.children ?? []);
    const effectPars = groups.flatMap((g) => Array.from(child(child(g, "cTn"), "childTnLst")?.children ?? []));
    const behaviorCount = child(child(effectPars[0] ?? null, "cTn"), "childTnLst")?.children.length ?? 0;
    const expectedCount = type === "appear" || type === "disappear" || type === "spin" ? 1 : type === "zoom-in" || type === "zoom-out" ? 3 : 2;
    const editable = !!type && step.effects.length === 1 && groups.length === 1 && effectPars.length === 1 && behaviorCount === expectedCount && e.behaviors.length === expectedCount;
    const targetName = Array.from(doc.getElementsByTagNameNS(NS.p, "cNvPr")).find((n) => n.getAttribute("id") === e?.spid)?.getAttribute("name") || "Object";
    return { el: click, spid: e?.spid ?? "", type: type ?? "fade-in", trigger: step.auto ? "after" : "click", dur: Math.max(1, ...step.effects.map(effect => effect.dur)) || 500,
      delay: e?.start ?? 0, byParagraph: !!e?.para, editable: !!e, simple: editable, targetName, label: editable ? ANIMATION_TYPES.find(([, t]) => t === type)![0] : "Existing animation" };
  });
}

/** Namespace-aware element builder and fresh timing IDs, including IDs from interactive sequences. */
function builder(doc: Document) {
  let id = Math.max(0, ...Array.from(doc.getElementsByTagNameNS(NS.p, "cTn")).map((e) => Number(e.getAttribute("id")) || 0));
  return (name: string, attrs: Record<string, string> = {}, ...kids: Element[]) => {
    const el = doc.createElementNS(NS.p, `p:${name}`);
    if (name === "cTn") el.setAttribute("id", String(++id));
    for (const [key, value] of Object.entries(attrs)) el.setAttribute(key, value);
    el.append(...kids);
    return el;
  };
}

function ensureMain(doc: Document): Element {
  const main = mainOf(doc);
  if (main) return main;
  const e = builder(doc);
  let timing = child(doc.documentElement, "timing");
  if (!timing) {
    timing = e("timing");
    doc.documentElement.insertBefore(timing, child(doc.documentElement, "extLst"));
  }
  let list = child(timing, "tnLst");
  if (!list) { list = e("tnLst"); timing.insertBefore(list, timing.firstChild); }
  const m = e("cTn", { dur: "indefinite", nodeType: "mainSeq" }, e("childTnLst"));
  const nav = (evt: string) => e("cond", { evt, delay: "0" }, e("tgtEl", {}, e("sldTgt")));
  const seq = e("seq", { concurrent: "1", nextAc: "seek" }, m, e("prevCondLst", {}, nav("onPrev")), e("nextCondLst", {}, nav("onNext")));
  // A new main sequence shares the existing timing root without disturbing interactive branches.
  const root = Array.from(list.getElementsByTagNameNS(NS.p, "cTn")).find((n) => n.getAttribute("nodeType") === "tmRoot");
  if (root) {
    let children = child(root, "childTnLst");
    if (!children) { children = e("childTnLst"); root.append(children); }
    children.append(seq);
  } else list.append(e("par", {}, e("cTn", { dur: "indefinite", restart: "never", nodeType: "tmRoot" }, e("childTnLst", {}, seq))));
  return m;
}

function behaviors(doc: Document, spid: string, s: AnimationSettings, e = builder(doc)): Element {
  const target = () => e("tgtEl", {}, e("spTgt", { spid }, ...(s.paragraph !== undefined ? [e("txEl", {}, e("pRg", { st: String(s.paragraph), end: String(s.paragraph) }))] : [])));
  const common = (attr?: string, dur = s.dur, delay = 0) => e("cBhvr", {},
    e("cTn", { dur: String(Math.round(dur)), fill: "hold" }, e("stCondLst", {}, e("cond", { delay: String(delay) }))), target(),
    ...(attr ? [e("attrNameLst", {}, (() => { const a = e("attrName"); a.textContent = attr; return a; })())] : [])
  );
  const visibility = (visible: boolean, delay = 0) => e("set", {}, common("style.visibility", 1, delay), e("to", {}, e("strVal", { val: visible ? "visible" : "hidden" })));
  const prop = (attr: string, from: string, to: string) => e("anim", { calcmode: "lin", valueType: "num" }, common(attr),
    e("tavLst", {}, e("tav", { tm: "0" }, e("val", {}, e("strVal", { val: from }))), e("tav", { tm: "100000" }, e("val", {}, e("strVal", { val: to })))));
  const list = e("childTnLst");
  const exit = isExit(s.type);
  const hide = () => visibility(false, Math.round(s.dur) - 1);
  if (s.type === "disappear") list.append(visibility(false));
  else if (s.type !== "spin" && !exit) list.append(visibility(true));
  if (s.type === "fade-in" || s.type === "fade-out") {
    list.append(e("animEffect", { transition: s.type === "fade-out" ? "out" : "in", filter: "fade" }, common()));
    if (s.type === "fade-out") list.append(hide());
  } else if (s.type.startsWith("fly-")) {
    const dir = s.type.replace(/^fly-(out-)?/, "");
    const horiz = dir === "left" || dir === "right";
    const pos = horiz ? "ppt_x" : "ppt_y";
    const size = horiz ? "ppt_w" : "ppt_h";
    const off = dir === "left" || dir === "up" ? `0-#${size}/2` : `1+#${size}/2`;
    if (exit) list.append(prop(pos, `#${pos}`, off), hide());
    else list.append(prop(pos, off, `#${pos}`));
  } else if (s.type === "zoom-in") list.append(prop("ppt_w", "0.1*#ppt_w", "#ppt_w"), prop("ppt_h", "0.1*#ppt_h", "#ppt_h"));
  else if (s.type === "zoom-out") list.append(prop("ppt_w", "#ppt_w", "0.1*#ppt_w"), prop("ppt_h", "#ppt_h", "0.1*#ppt_h"), hide());
  else if (s.type === "spin") list.append(e("animRot", { by: "21600000" }, common("r")));
  return list;
}

const isExit = (t: AnimationType) => t === "disappear" || t === "fade-out" || t === "zoom-out" || t.startsWith("fly-out");
const effectAttrs = (s: AnimationSettings) => ({ presetID: s.type === "appear" || s.type === "disappear" ? "1" : s.type.startsWith("fade-") ? "10" : s.type.startsWith("fly-") ? "2" : s.type.startsWith("zoom-") ? "23" : "8",
  presetClass: isExit(s.type) ? "exit" : s.type === "spin" ? "emph" : "entr", presetSubtype: "0", fill: "hold", nodeType: s.trigger === "click" ? "clickEffect" : s.trigger === "with" ? "withEffect" : "afterEffect" });

export function addAnimation(doc: Document, spid: string, patch: Partial<AnimationSettings> = {}): void {
  const s = { ...defaults, ...patch };
  validate(s);
  if (!Array.from(doc.getElementsByTagNameNS(NS.p, "cNvPr")).some((n) => n.getAttribute("id") === spid && n.parentElement?.parentElement?.localName !== "spTree")) throw new Error("Select a shape to animate");
  if (s.byParagraph) {
    const shape = Array.from(doc.getElementsByTagNameNS(NS.p, "cNvPr")).find(n => n.getAttribute("id") === spid)!.parentElement!.parentElement!;
    const count = shape.getElementsByTagNameNS(NS.a, "p").length;
    if (!count) throw new Error("Select a text object to animate by paragraph");
    for (let paragraph = 0; paragraph < count; paragraph++) addAnimation(doc, spid, { ...s, byParagraph: false, paragraph });
    const timing = child(doc.documentElement,"timing")!;
    let builds=child(timing,"bldLst");
    const e=builder(doc);
    if(!builds){builds=e("bldLst");timing.insertBefore(builds,child(timing,"extLst"));}
    if(!Array.from(builds.children).some(b=>b.getAttribute("spid")===spid))builds.append(e("bldP",{spid,grpId:"0",build:"p",uiExpand:"1"}));
    return;
  }
  const main = ensureMain(doc);
  let list = child(main, "childTnLst");
  const e = builder(doc);
  if (!list) { list = e("childTnLst"); main.append(list); }
  const b = behaviors(doc, spid, s, e);
  const make = e;
  const effect = make("par", {}, make("cTn", effectAttrs(s), make("stCondLst", {}, make("cond", { delay: String(Math.round(s.delay)) })), b));
  const group = make("par", {}, make("cTn", { fill: "hold" }, make("stCondLst", {}, make("cond", { delay: "0" })), make("childTnLst", {}, effect)));
  const cond = s.trigger === "click" ? make("cond", { delay: "indefinite" }) : make("cond", { evt: "onBegin", delay: "0" }, make("tn", { val: main.getAttribute("id")! }));
  const click = make("par", {}, make("cTn", { fill: "hold", nodeType: "clickPar" }, make("stCondLst", {}, cond), make("childTnLst", {}, group)));
  if (s.trigger === "with" && list.lastElementChild) {
    const previous = child(list.lastElementChild, "cTn")!;
    child(previous, "childTnLst")!.append(group);
    return;
  }
  list.append(click);
  syncAutomaticStarts(doc, [click]);
}

export function updateAnimation(doc: Document, entry: AnimationEntry, patch: Partial<AnimationSettings>): void {
  if (!entry.editable || !doc.contains(entry.el)) throw new Error("This imported animation cannot be edited");
  const s = { ...entry, ...patch };
  validate(s);
  if (patch.byParagraph !== undefined && entry.simple) {
    if(patch.byParagraph && !entry.byParagraph) {
      const parent=entry.el.parentElement!;
      const existing=new Set(Array.from(parent.children));
      addAnimation(doc,entry.spid,{...s,byParagraph:true});
      for(const click of Array.from(parent.children).filter(c=>!existing.has(c)))parent.insertBefore(click,entry.el);
      entry.el.remove(); syncAutomaticStarts(doc); return;
    }
    if(!patch.byParagraph)for(const target of Array.from(entry.el.getElementsByTagNameNS(NS.p,"spTgt")))child(target,"txEl")?.remove();
  }
  const click = child(entry.el, "cTn")!;
  if (!entry.simple) {
    if (patch.type !== undefined) throw new Error("Change imported effect timing without replacing its behaviors");
    if (patch.dur !== undefined) {
      const scale = s.dur / entry.dur;
      for (const ct of Array.from(click.getElementsByTagNameNS(NS.p, "cTn"))) {
        const dur = ct.getAttribute("dur");
        if (dur && /^\d+$/.test(dur)) ct.setAttribute("dur", String(Math.max(1, Math.round(Number(dur) * scale))));
      }
    }
    for (const group of Array.from(child(click, "childTnLst")?.children ?? [])) {
      const ct = child(group, "cTn");
      const cond = child(child(ct, "stCondLst"), "cond");
      if (cond && patch.delay !== undefined) cond.setAttribute("delay", String(Math.max(0, Number(cond.getAttribute("delay") || 0) + s.delay - entry.delay)));
    }
    if (patch.trigger !== undefined) setStart(doc, entry.el, s.trigger);
    if(patch.trigger !== undefined)syncAutomaticStarts(doc, [entry.el]);
    return;
  }
  const group = child(child(child(click, "childTnLst")?.firstElementChild ?? null, "cTn"), "childTnLst")?.firstElementChild ?? null;
  const effect = child(group, "cTn")!;
  if (patch.type !== undefined) {
    const paragraph = child(effect.getElementsByTagNameNS(NS.p, "txEl")[0] ?? null, "pRg")?.getAttribute("st");
    const b = behaviors(doc, entry.spid, { ...s, paragraph: paragraph !== undefined && paragraph !== null ? Number(paragraph) : undefined });
    effect.replaceChild(b, child(effect, "childTnLst")!);
    for (const [key, value] of Object.entries(effectAttrs(s))) effect.setAttribute(key, value);
  }
  if(patch.dur !== undefined && patch.type === undefined)updateAnimationPart(doc,animationParts(entry)[0],{dur:s.dur});
  if(patch.delay !== undefined){
    const groupCT = child(child(click, "childTnLst")!.firstElementChild, "cTn")!;
    child(child(groupCT, "stCondLst"), "cond")!.setAttribute("delay", "0");
    child(child(effect, "stCondLst"), "cond")!.setAttribute("delay", String(Math.round(s.delay)));
  }
  if (patch.trigger !== undefined) setStart(doc, entry.el, s.trigger);
  if(patch.trigger !== undefined)syncAutomaticStarts(doc, [entry.el]);
}


function setStart(doc: Document, clickEl: Element, trigger: AnimationSettings["trigger"]): void {
  const ct = child(clickEl, "cTn")!;
  if (trigger === "with" && clickEl.previousElementSibling) {
    child(child(clickEl.previousElementSibling, "cTn"), "childTnLst")!.append(...Array.from(child(ct, "childTnLst")!.children));
    clickEl.remove();
    return;
  }
  const e = builder(doc);
  const cond = trigger === "click" ? e("cond", { delay: "indefinite" }) : e("cond", { evt: "onBegin", delay: "0" }, e("tn", { val: mainOf(doc)!.getAttribute("id")! }));
  child(ct, "stCondLst")!.replaceChildren(cond);
}

/** Automatic steps follow the previous step in Office as well as in our player. */
function syncAutomaticStarts(doc: Document, changed?: Element[]): void {
  const entries = animationEntries(doc);
  entries.forEach((entry, i) => {
    if (!entry.editable || entry.trigger !== "after" || (changed && !changed.includes(entry.el))) return;
    const e = builder(doc);
    const previous = i ? child(entries[i - 1].el, "cTn") : mainOf(doc);
    const cond = e("cond", { evt: i ? "onEnd" : "onBegin", delay: "0" }, e("tn", { val: previous!.getAttribute("id")! }));
    child(child(entry.el, "cTn"), "stCondLst")!.replaceChildren(cond);
  });
}

export function removeAnimation(doc: Document, entry: AnimationEntry): void {
  if (!doc.contains(entry.el) || !entry.editable) throw new Error("This imported animation cannot be removed");
  entry.el.remove();
  syncAutomaticStarts(doc);
}

export function moveAnimation(doc: Document, entry: AnimationEntry, direction: -1 | 1): void {
  const entries = animationEntries(doc);
  const index = entries.findIndex((e) => e.el === entry.el);
  const neighbor = entries[index + direction];
  if (index < 0 || !neighbor || !entry.editable || !neighbor.editable) return;
  const parent = entry.el.parentElement!;
  parent.insertBefore(entry.el, direction === -1 ? neighbor.el : neighbor.el.nextSibling);
  syncAutomaticStarts(doc);
}

export interface AnimationPart { el: Element; spid: string; dur: number; delay: number; label: string; motionPath?: string }
/** Individual effects inside an imported compound step, retaining their complete behavior XML. */
export function animationParts(entry: AnimationEntry): AnimationPart[] {
 return Array.from(entry.el.getElementsByTagNameNS(NS.p,"cTn")).filter(e=>e.hasAttribute("presetClass")).map(el=> {
  const spid=el.getElementsByTagNameNS(NS.p,"spTgt")[0]?.getAttribute("spid")??"";
  const timings=Array.from(el.getElementsByTagNameNS(NS.p,"cTn"));
  const dur=Math.max(1,...timings.map(ct=>(Number(ct.getAttribute("dur"))||0)+(Number(child(child(ct,"stCondLst"),"cond")?.getAttribute("delay"))||0)));
  const delay=Number(child(child(el,"stCondLst"),"cond")?.getAttribute("delay"))||0;
  return {el,spid,dur,delay,label:`${el.getAttribute("presetClass")} · object ${spid}`,motionPath:el.getElementsByTagNameNS(NS.p,"animMotion")[0]?.getAttribute("path")??undefined};
 });
}
export function updateAnimationPart(doc:Document,part:AnimationPart,patch:Partial<Pick<AnimationPart,"dur"|"delay"|"motionPath">>):void {
 if(!doc.contains(part.el))throw new Error("Animation is no longer on this slide");
 if(patch.dur!==undefined&&(!Number.isFinite(patch.dur)||patch.dur<1))throw new Error("Duration must be positive");
 if(patch.delay!==undefined&&(!Number.isFinite(patch.delay)||patch.delay<0))throw new Error("Delay cannot be negative");
 if(patch.motionPath!==undefined && (!/^[MLCQAZEHmlcqazeh\d.,+\-\s]+$/.test(patch.motionPath)||!/^M\s/i.test(patch.motionPath)))throw new Error("Enter a motion path beginning with M");
 if(patch.dur!==undefined) {
  const scale=patch.dur/part.dur;
  for(const ct of Array.from(part.el.getElementsByTagNameNS(NS.p,"cTn"))) {
   const dur=ct.getAttribute("dur");if(dur&&/^\d+$/.test(dur)&&ct.parentElement?.parentElement?.localName!=="set")ct.setAttribute("dur",String(Math.max(1,Math.round(Number(dur)*scale))));
   for(const cond of Array.from(child(ct,"stCondLst")?.children??[])) {const delay=cond.getAttribute("delay");if(delay&&/^\d+$/.test(delay))cond.setAttribute("delay",String(Math.round(Number(delay)*scale)));}
  }
 }
 if(patch.delay!==undefined){let list=child(part.el,"stCondLst");if(!list){list=builder(doc)("stCondLst");part.el.insertBefore(list,part.el.firstChild);}let cond=child(list,"cond");if(!cond){cond=builder(doc)("cond");list.append(cond);}cond.setAttribute("delay",String(Math.round(patch.delay)));}
 if(patch.motionPath!==undefined)part.el.getElementsByTagNameNS(NS.p,"animMotion")[0]?.setAttribute("path",patch.motionPath);
}
