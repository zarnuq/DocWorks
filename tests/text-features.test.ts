import { describe, it, expect, beforeEach } from "vitest";
import { parseSlide, readParagraphs as readModel, TextStyle } from "../src/domain/pptx/slide-model";
import type { Para, RunProps } from "../src/domain/pptx/slide-model";
import { addTextBox, writeBodyPr, writeParagraphs } from "../src/domain/pptx/slide-edit";
import { insertField, readParagraphs, renderParagraphs, applyRunStyle, refreshMarks } from "../src/presentation/pptx/text-html";
import { loadSample, roundTrip } from "./helpers";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const xml = (s: string) => new DOMParser().parseFromString(s, "application/xml").documentElement;
const style: TextStyle = { chain: [], fontRefColor: null, colors: { scheme: {}, map: {} }, fonts: { major: "Calibri", minor: "Calibri" } };
const body = (paras: string) => xml(`<p:txBody xmlns:p="${P}" xmlns:a="${A}"><a:bodyPr/><a:lstStyle/>${paras}</p:txBody>`);
const order = (el: Element | undefined) => Array.from(el?.children ?? []).map((c) => c.localName);

/** A fresh text box on the sample deck, re-parsed after each change. */
async function textBox() {
  const doc = await loadSample();
  const path = doc.slides()[1].path;
  const el = addTextBox(doc.xml(path)!, { x: 0, y: 0, w: 3000000, h: 1000000 }, "Hello");
  const model = () => parseSlide(doc, path).shapes.find((s) => s.el === el)!;
  return { doc, path, el, model };
}

describe("bullet colour", () => {
  it("reads a:buClr; a:buClrTx or none follows the text", () => {
    const [own, follow, none] = readModel(body(
      `<a:p><a:pPr><a:buClr><a:srgbClr val="00AA00"/></a:buClr><a:buChar char="•"/></a:pPr><a:r><a:t>x</a:t></a:r></a:p>` +
      `<a:p><a:pPr><a:buClrTx/><a:buChar char="•"/></a:pPr><a:r><a:t>x</a:t></a:r></a:p>` +
      `<a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:t>x</a:t></a:r></a:p>`), style);
    expect(own.bulletColor).toBe("#00AA00");
    expect(follow.bulletColor).toBeNull();
    expect(none.bulletColor).toBeNull();
  });
});

describe("numbered lists", () => {
  it("reads buAutoNum as a numbering scheme with its start", () => {
    const [p] = readModel(body(`<a:p><a:pPr><a:buAutoNum type="alphaLcParenR" startAt="3"/></a:pPr><a:r><a:t>x</a:t></a:r></a:p>`), style);
    expect(p.bullet).toBe("#alphaLcParenR@3");
  });

  it("writes numbering and bullets in schema order", async () => {
    const { model } = await textBox();
    const t = model().text!;
    const run = t.paras[0].runs[0];
    writeParagraphs(t, [
      { align: "l", lvl: 0, bullet: "#romanUcPeriod@2", runs: [run] },
      { align: "l", lvl: 0, bullet: "•", runs: [{ ...run, text: "b" }] },
    ]);
    const ps = Array.from(t.el.getElementsByTagNameNS(A, "pPr"));
    const auto = ps[0].getElementsByTagNameNS(A, "buAutoNum")[0];
    expect([auto.getAttribute("type"), auto.getAttribute("startAt")]).toEqual(["romanUcPeriod", "2"]);
    expect(ps[1].getElementsByTagNameNS(A, "buChar")[0].getAttribute("char")).toBe("•");
    expect(model().text!.paras.map((p) => p.bullet)).toEqual(["#romanUcPeriod@2", "•"]);
  });

  it("labels numbered paragraphs per level, skipping empty ones", () => {
    const host = document.createElement("div");
    const defaults: RunProps = { b: false, i: false, u: false, strike: false, sz: 18, color: "#000000", font: "Calibri", spc: 0, cap: "none", link: null };
    const para = (text: string, lvl: number, bullet: string | null): Para => ({
      runs: text ? [{ ...defaults, text }] : [],
      align: "l",
      lvl,
      bullet,
      marL: 0,
      indent: 0,
      defaults,
    });
    renderParagraphs(host, [
      para("a", 0, "#arabicPeriod"),
      para("", 0, "#arabicPeriod"),
      para("b", 0, "#arabicPeriod"),
      para("c", 1, "#alphaLcParenR"),
      para("d", 1, "#alphaLcParenR"),
      para("e", 0, "#arabicPeriod@5"),
      para("f", 0, "•"),
      para("g", 0, "#romanUcPeriod"),
    ], 2);
    const labels = Array.from(host.querySelectorAll("p")).map((p) => p.dataset.label ?? "");
    expect(labels).toEqual(["1.", "", "2.", "a)", "b)", "3.", "•", "I."]);
    expect(host.querySelector("p")!.dataset.bullet).toBe("#arabicPeriod");
    // Typing into the empty paragraph renumbers.
    host.querySelectorAll("p")[1].textContent = "new";
    refreshMarks(host);
    expect(Array.from(host.querySelectorAll("p")).map((p) => p.dataset.label)[2]).toBe("3.");
  });
});

describe("character spacing", () => {
  it("reads, renders, reads back and writes a:rPr spc", async () => {
    const [p] = readModel(body(`<a:p><a:r><a:rPr spc="300"/><a:t>x</a:t></a:r></a:p>`), style);
    expect(p.runs[0].spc).toBe(3);
    const host = document.createElement("div");
    renderParagraphs(host, [p], 2);
    expect(host.querySelector("span")!.style.letterSpacing).toBe("6px");
    const read = readParagraphs(host, 2, () => p.defaults);
    expect(read[0].runs[0].spc).toBe(3);

    const { model } = await textBox();
    const t = model().text!;
    writeParagraphs(t, [{ ...t.paras[0], runs: [{ ...t.paras[0].runs[0], spc: -1.5 }] }]);
    expect(t.el.getElementsByTagNameNS(A, "rPr")[0].getAttribute("spc")).toBe("-150");
  });
});

describe("paragraph spacing", () => {
  it("writes line spacing and space before / after first in a:pPr", async () => {
    const { model } = await textBox();
    const t = model().text!;
    writeParagraphs(t, [{ ...t.paras[0], bullet: "•", lnSpc: { pct: 1.5 }, spcBef: { pts: 6 }, spcAft: { pts: 12 } }]);
    const pPr = t.el.getElementsByTagNameNS(A, "pPr")[0];
    expect(order(pPr)).toEqual(["lnSpc", "spcBef", "spcAft", "buChar"]);
    expect(model().text!.paras[0]).toMatchObject({ lnSpc: { pct: 1.5 }, spcBef: { pts: 6 }, spcAft: { pts: 12 } });
    // Unchanged (inherited) spacing is not written.
    const t2 = model().text!;
    writeParagraphs(t2, [{ ...t2.paras[0], lnSpc: { pct: 1.5 } }]);
    expect(t2.el.getElementsByTagNameNS(A, "lnSpc")).toHaveLength(1);
  });
});

describe("text box settings", () => {
  it("writes anchor, insets, wrap and autofit on a:bodyPr", async () => {
    const { model } = await textBox();
    const t = model().text!;
    writeBodyPr(t.el, { anchor: "b", insets: { l: 0, t: 45720, r: 0, b: 45720 }, wrap: false, autofit: { type: "shrink", fontScale: 0.8, lnSpcReduction: 0.1 } });
    const bodyPr = t.el.getElementsByTagNameNS(A, "bodyPr")[0];
    expect(bodyPr.getAttribute("anchor")).toBe("b");
    expect(bodyPr.getAttribute("wrap")).toBe("none");
    const fit = bodyPr.getElementsByTagNameNS(A, "normAutofit")[0];
    expect([fit.getAttribute("fontScale"), fit.getAttribute("lnSpcReduction")]).toEqual(["80000", "10000"]);
    expect(model().text).toMatchObject({ anchor: "b", wrap: false, autofit: "shrink", fontScale: 0.8, lnSpcReduction: 0.1, insets: { l: 0, r: 0 } });
    writeBodyPr(t.el, { autofit: { type: "resize" }, wrap: true });
    expect(order(bodyPr)).toEqual(["spAutoFit"]);
    expect(model().text).toMatchObject({ autofit: "resize", wrap: true, fontScale: 1 });
  });
});

describe("hyperlinks", () => {
  it("writes a link as an external relationship and reads it back", async () => {
    const { doc, path, model } = await textBox();
    const t = model().text!;
    writeParagraphs(t, [{ ...t.paras[0], runs: [{ ...t.paras[0].runs[0], link: "https://example.com/" }] }]);
    const back = await roundTrip(doc);
    const run = parseSlide(back, path).shapes.find((s) => s.text?.paras[0]?.runs[0]?.link)!.text!.paras[0].runs[0];
    expect(run.link).toBe("https://example.com/");
    const rel = back.rels(path).find((r) => r.type === "hyperlink")!;
    expect(rel).toMatchObject({ target: "https://example.com/", external: true });
  });

  it("round-trips links through the editor DOM", () => {
    const [p] = readModel(body(`<a:p><a:r><a:t>x</a:t></a:r></a:p>`), style);
    const host = document.createElement("div");
    document.body.appendChild(host);
    renderParagraphs(host, [p], 2);
    const r = document.createRange();
    r.selectNodeContents(host.querySelector("span")!);
    getSelection()!.removeAllRanges();
    getSelection()!.addRange(r);
    applyRunStyle(host, { link: "https://a.b/" }, 2);
    expect(readParagraphs(host, 2, () => p.defaults)[0].runs[0].link).toBe("https://a.b/");
    host.remove();
  });
});

describe("fields", () => {
  let host: HTMLElement;
  beforeEach(() => {
    document.body.innerHTML = "";
    host = document.body.appendChild(document.createElement("div"));
  });

  it("reads fields, shows the live slide number and writes them back", async () => {
    const [p] = readModel(body(`<a:p><a:fld id="{A}" type="slidenum"><a:rPr/><a:t>‹#›</a:t></a:fld></a:p>`), style);
    expect(p.runs[0]).toMatchObject({ field: "slidenum", text: "‹#›" });
    renderParagraphs(host, [p], 2, undefined, 0, { slideNumber: 7 });
    const span = host.querySelector<HTMLElement>("[data-field]")!;
    expect(span.textContent).toBe("7");
    expect(span.contentEditable).toBe("false");
    expect(readParagraphs(host, 2, () => p.defaults)[0].runs[0]).toMatchObject({ field: "slidenum" });

    const { model } = await textBox();
    const t = model().text!;
    writeParagraphs(t, [{ ...t.paras[0], runs: [...t.paras[0].runs, { ...t.paras[0].runs[0], text: "3", field: "slidenum" }] }]);
    const fld = t.el.getElementsByTagNameNS(A, "fld")[0];
    expect(fld.getAttribute("type")).toBe("slidenum");
    expect(fld.getAttribute("id")).toMatch(/^\{[0-9A-F-]{36}\}$/);
    expect(model().text!.paras[0].runs.map((r) => r.field ?? null)).toEqual([null, "slidenum"]);
  });

  it("inserts a field at the caret", () => {
    host.innerHTML = '<p data-lvl="0"><span>ab</span></p>';
    const text = host.querySelector("span")!.firstChild!;
    const r = document.createRange();
    r.setStart(text, 1);
    r.collapse(true);
    getSelection()!.removeAllRanges();
    getSelection()!.addRange(r);
    insertField(host, "slidenum", "4");
    const runs = readParagraphs(host, 2, () => ({ b: false, i: false, u: false, strike: false, sz: 18, color: "#000000", font: "Calibri", spc: 0, cap: "none", link: null }))[0].runs;
    expect(runs.map((x) => [x.text, x.field ?? null])).toEqual([["a", null], ["4", "slidenum"], ["b", null]]);
  });
});
