import { describe, it, expect, beforeEach } from "vitest";
import { renderParagraphs, refreshMarks, insertPlainText, readParagraphs, applyRunStyle, selectionRunProps, updateParagraphs } from "../src/presentation/pptx/text-html";
import type { Para, RunProps } from "../src/domain/pptx/slide-model";

const PX_PER_PT = 2;
const defaults: RunProps = { b: false, i: false, u: false, strike: false, sz: 18, color: "#000000", font: "Calibri", spc: 0, cap: "none", link: null };
const defaultsFor = (lvl: number): RunProps => ({ ...defaults, sz: 18 - lvl * 2 });
const run = (text: string, extra: Partial<RunProps> = {}) => ({ ...defaults, text, ...extra });
const para = (runs: Para["runs"], extra: Partial<Para> = {}): Para => ({
  runs,
  align: "l",
  lvl: 0,
  bullet: null,
  marL: 0,
  indent: 0,
  defaults,
  ...extra,
});

let host: HTMLElement;
beforeEach(() => {
  document.body.innerHTML = "";
  host = document.body.appendChild(document.createElement("div"));
});
const read = () => readParagraphs(host, PX_PER_PT, defaultsFor);
const select = (node: Node, start: number, endNode: Node, end: number) => {
  const r = document.createRange();
  r.setStart(node, start);
  r.setEnd(endNode, end);
  const sel = window.getSelection()!;
  sel.removeAllRanges();
  sel.addRange(r);
};

describe("renderParagraphs / readParagraphs", () => {
  it("round-trips formatted runs", () => {
    const paras = [
      para([run("Plain "), run("bold", { b: true }), run(" red", { color: "#FF0000", sz: 24, font: "Georgia" })]),
      para([run("under", { u: true, strike: true, i: true })], { align: "ctr", lvl: 1, bullet: "•", defaults: defaultsFor(1) }),
    ];
    paras[1].runs[0].sz = 16;
    renderParagraphs(host, paras, PX_PER_PT);
    expect(read()).toEqual(paras.map(({ runs, align, lvl, bullet }) => ({ runs, align, lvl, bullet, lnSpc: null, spcBef: null, spcAft: null })));
  });

  it("styles only runs that differ from the paragraph defaults", () => {
    renderParagraphs(host, [para([run("a"), run("b", { b: true })])], PX_PER_PT);
    const p = host.querySelector("p")!;
    expect(p.style.fontSize).toBe("36px");
    const spans = p.querySelectorAll("span");
    expect(spans[0].getAttribute("style") ?? "").toBe("");
    expect(spans[1].style.fontWeight).toBe("bold");
  });

  it("lets a run turn off an underline its text box style sets (CSS can't undo a parent's underline)", () => {
    const underlined = { ...defaults, u: true };
    const paras = [para([run("plain", { u: false }), run(" under", { u: true })], { defaults: underlined })];
    renderParagraphs(host, paras, PX_PER_PT);
    const p = host.querySelector("p")!;
    expect(p.style.textDecoration).toBe("");
    const spans = p.querySelectorAll("span");
    expect(spans[0].style.textDecoration).toBe("none");
    expect(spans[1].style.textDecoration).toBe("underline");
    expect(readParagraphs(host, PX_PER_PT, () => underlined)[0].runs.map((r) => r.u)).toEqual([false, true]);
  });

  it("renders empty paragraphs with a <br> and reads them back empty", () => {
    renderParagraphs(host, [para([])], PX_PER_PT);
    expect(host.innerHTML).toContain("<br>");
    expect(read()[0].runs).toEqual([]);
  });

  it("marks empty paragraphs so their bullets can be hidden, like PowerPoint", () => {
    renderParagraphs(host, [para([], { bullet: "•" }), para([run("x")], { bullet: "•" })], PX_PER_PT);
    const [empty, full] = Array.from(host.querySelectorAll("p"));
    expect(empty.hasAttribute("data-empty")).toBe(true);
    expect(full.hasAttribute("data-empty")).toBe(false);
    empty.textContent = "typed";
    refreshMarks(host);
    expect(empty.hasAttribute("data-empty")).toBe(false);
    expect(read()[0].bullet).toBe("•");
  });

  it("falls back to metric-compatible fonts", () => {
    renderParagraphs(host, [para([run("x", { font: "Georgia" })])], PX_PER_PT);
    const p = host.querySelector("p")!;
    expect(p.style.fontFamily).toMatch(/^"?Calibri"?, "?Carlito"?/);
    expect(read()[0].runs[0].font).toBe("Georgia");
  });

  it("reads line breaks inside a paragraph", () => {
    renderParagraphs(host, [para([run("one\ntwo")])], PX_PER_PT);
    expect(read()[0].runs.map((r) => r.text)).toEqual(["one", "\n", "two"]);
  });

  it("uses the new level's defaults for unstyled text", () => {
    renderParagraphs(host, [para([run("x")])], PX_PER_PT);
    host.querySelector("p")!.dataset.lvl = "2";
    expect(read()[0].runs[0].sz).toBe(14);
  });

  it("understands browser-generated markup", () => {
    host.innerHTML =
      '<p data-lvl="0"><b>b</b><i>i</i><u>u</u><strike>s</strike><font color="#00ff00" face="Arial">f</font>' +
      '<span style="font-weight: normal">n</span></p><div>loose div</div>';
    const [p, div] = read();
    expect(p.runs.map((r) => [r.text, r.b, r.i, r.u, r.strike])).toEqual([
      ["b", true, false, false, false],
      ["i", false, true, false, false],
      ["u", false, false, true, false],
      ["s", false, false, false, true],
      ["f", false, false, false, false],
      ["n", false, false, false, false],
    ]);
    expect(p.runs[4]).toMatchObject({ color: "#00FF00", font: "Arial" });
    expect(div.runs[0].text).toBe("loose div");
  });

  it("wraps stray top-level text into a paragraph", () => {
    host.innerHTML = "typed";
    expect(read()).toEqual([{ runs: [run("typed")], align: "l", lvl: 0, bullet: null }]);
  });

  it("merges adjacent runs with identical formatting", () => {
    host.innerHTML = '<p><span>a</span><span>b</span><b>c</b></p>';
    expect(read()[0].runs.map((r) => r.text)).toEqual(["ab", "c"]);
  });
});

describe("applyRunStyle", () => {
  it("wraps only the selected part of a text node", () => {
    renderParagraphs(host, [para([run("hello world")])], PX_PER_PT);
    const t = host.querySelector("span")!.firstChild!;
    select(t, 6, t, 11);
    applyRunStyle(host, { b: true }, PX_PER_PT);
    expect(read()[0].runs).toEqual([run("hello "), run("world", { b: true })]);
  });

  it("spans multiple paragraphs", () => {
    renderParagraphs(host, [para([run("one")]), para([run("two")])], PX_PER_PT);
    const [a, b] = Array.from(host.querySelectorAll("span")).map((s) => s.firstChild!);
    select(a, 1, b, 2);
    applyRunStyle(host, { color: "#123456", sz: 30 }, PX_PER_PT);
    const [p1, p2] = read();
    expect(p1.runs).toEqual([run("o"), run("ne", { color: "#123456", sz: 30 })]);
    expect(p2.runs).toEqual([run("tw", { color: "#123456", sz: 30 }), run("o")]);
  });

  it("ignores selections outside the host", () => {
    renderParagraphs(host, [para([run("x")])], PX_PER_PT);
    const outside = document.body.appendChild(document.createTextNode("outside"));
    select(outside, 0, outside, 3);
    applyRunStyle(host, { b: true }, PX_PER_PT);
    expect(read()[0].runs[0].b).toBe(false);
  });

  it("reports the selection's common formatting for toggles", () => {
    renderParagraphs(host, [para([run("ab", { b: true }), run("cd")])], PX_PER_PT);
    const [s1, s2] = Array.from(host.querySelectorAll("span")).map((s) => s.firstChild!);
    select(s1, 0, s1, 2);
    expect(selectionRunProps(host, PX_PER_PT, defaultsFor)?.b).toBe(true);
    select(s1, 0, s2, 1);
    expect(selectionRunProps(host, PX_PER_PT, defaultsFor)?.b).toBe(false);
  });
});

describe("updateParagraphs", () => {
  it("changes alignment, level and bullets of selected paragraphs only", () => {
    renderParagraphs(host, [para([run("a")]), para([run("b")]), para([run("c")])], PX_PER_PT);
    const spans = Array.from(host.querySelectorAll("span")).map((s) => s.firstChild!);
    select(spans[0], 0, spans[1], 1);
    updateParagraphs(host, (p) => ({ ...p, align: "r", lvl: p.lvl + 1, bullet: "•" }));
    expect(read().map((p) => [p.align, p.lvl, p.bullet])).toEqual([
      ["r", 1, "•"],
      ["r", 1, "•"],
      ["l", 0, null],
    ]);
  });

  it("treats a caret as selecting its paragraph", () => {
    renderParagraphs(host, [para([run("a")]), para([run("b")])], PX_PER_PT);
    const t = host.querySelectorAll("span")[1].firstChild!;
    select(t, 0, t, 0);
    updateParagraphs(host, (p) => ({ ...p, align: "ctr" }));
    expect(read().map((p) => p.align)).toEqual(["l", "ctr"]);
  });
});

describe("renderParagraphs spacing", () => {
  it("sizes lines from the runs, not the level default", () => {
    renderParagraphs(host, [para([run("x", { sz: 10 })], { defaults: { ...defaults, sz: 28 } })], PX_PER_PT);
    // single spacing = 1.2 x the largest run
    expect(host.querySelector("p")!.style.lineHeight).toBe(`${1.2 * 10 * PX_PER_PT}px`);
  });

  it("applies percentage and point line spacing, minus autofit reduction", () => {
    renderParagraphs(host, [para([run("x")], { lnSpc: { pct: 0.9 } }), para([run("y")], { lnSpc: { pts: 30 } })], PX_PER_PT, undefined, 0.1);
    const [a, b] = Array.from(host.querySelectorAll("p"));
    expect(parseFloat(a.style.lineHeight)).toBeCloseTo(1.2 * 0.8 * 18 * PX_PER_PT, 5);
    expect(b.style.lineHeight).toBe(`${30 * PX_PER_PT}px`);
  });

  it("uses the end-of-paragraph size for empty paragraphs", () => {
    renderParagraphs(host, [para([], { endSz: 12 })], PX_PER_PT);
    expect(host.querySelector("p")!.style.lineHeight).toBe(`${1.2 * 12 * PX_PER_PT}px`);
  });

  it("adds space before and after paragraphs, but not before the first (like PowerPoint)", () => {
    const spaced = para([run("x")], { spcBef: { pts: 10 }, spcAft: { pct: 0.5 } });
    renderParagraphs(host, [spaced, spaced], PX_PER_PT);
    const [first, p] = Array.from(host.querySelectorAll("p"));
    expect(first.style.marginTop).toBe("");
    expect(p.style.marginTop).toBe(`${10 * PX_PER_PT}px`);
    expect(parseFloat(p.style.marginBottom)).toBeCloseTo(0.5 * 1.2 * 18 * PX_PER_PT, 5);
  });

  it("sizes the bullet like the first run", () => {
    renderParagraphs(host, [para([run("x", { sz: 10 })], { bullet: "•", defaults: { ...defaults, sz: 28 } })], PX_PER_PT);
    expect(host.querySelector("p")!.style.getPropertyValue("--pptx-bullet-size")).toBe(`${10 * PX_PER_PT}px`);
  });
});

describe("insertPlainText", () => {
  it("replaces the selection with text, turning newlines into line breaks", () => {
    renderParagraphs(host, [para([run("hello world")])], PX_PER_PT);
    const text = host.querySelector("span")!.firstChild!;
    select(text, 6, text, 11);
    insertPlainText("big\nday");
    expect(read()[0].runs.map((r) => r.text).join("")).toBe("hello big\nday");
    // caret ends after the pasted text
    const caret = window.getSelection()!.getRangeAt(0);
    expect(caret.collapsed).toBe(true);
    expect(caret.startContainer.textContent).toBe("day");
  });
});
