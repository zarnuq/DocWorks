/*
 * Paragraph model <-> contentEditable HTML. Uses plain DOM APIs only (no
 * Obsidian helpers) so it runs under jsdom in tests.
 *
 * Each paragraph is a <p data-lvl data-bullet> styled with its level defaults;
 * runs are <span>s styled only where they differ from those defaults. Reading
 * back resolves unstyled text against the (possibly changed) level's defaults.
 */
import type { Align, Para, Run, RunProps, Spacing } from "../../domain/pptx/slide-model";
import type { ParaInput } from "../../domain/pptx/slide-edit";
import { fontStack, letterSpacing } from "./text-fit";

const CSS_ALIGN: Record<Align, string> = { l: "left", ctr: "center", r: "right", just: "justify" };
const FROM_CSS_ALIGN: Record<string, Align> = { left: "l", start: "l", center: "ctr", right: "r", end: "r", justify: "just" };
const BLOCK = new Set(["P", "DIV", "LI", "H1", "H2", "H3", "H4", "H5", "H6"]);
const EMU_PER_PT = 12700;
const PROP_KEYS: (keyof RunProps)[] = ["b", "i", "u", "strike", "sz", "color", "font"];

export type DefaultsFor = (lvl: number) => RunProps;

const decoration = (u: boolean, strike: boolean) => [u && "underline", strike && "line-through"].filter(Boolean).join(" ") || "none";

/** Write the CSS for every property in `props` onto `el`. */
function styleRun(el: HTMLElement, props: Partial<RunProps>, pxPerPt: number, decorationBase?: { u: boolean; strike: boolean }): void {
  if (props.b !== undefined) el.style.fontWeight = props.b ? "bold" : "normal";
  if (props.i !== undefined) el.style.fontStyle = props.i ? "italic" : "normal";
  if (props.u !== undefined || props.strike !== undefined) {
    el.style.textDecoration = decoration(props.u ?? decorationBase?.u ?? false, props.strike ?? decorationBase?.strike ?? false);
  }
  if (props.sz !== undefined) el.style.fontSize = `${props.sz * pxPerPt}px`;
  if (props.color !== undefined) el.style.color = props.color;
  if (props.font !== undefined) el.style.fontFamily = fontStack(props.font);
}

function diff(run: RunProps, base: RunProps): Partial<RunProps> {
  const out: Partial<RunProps> = {};
  for (const k of PROP_KEYS) {
    if (String(run[k]).toUpperCase() !== String(base[k]).toUpperCase()) (out as Record<string, unknown>)[k] = run[k];
  }
  if (out.u !== undefined || out.strike !== undefined) Object.assign(out, { u: run.u, strike: run.strike });
  return out;
}

export function renderParagraphs(host: HTMLElement, paras: Para[], pxPerPt: number, pxPerEmu = pxPerPt / EMU_PER_PT, lnSpcReduction = 0): void {
  host.replaceChildren();
  paras.forEach((para, index) => {
    const p = host.createEl("p");
    p.dataset.lvl = String(para.lvl);
    if (para.bullet) p.dataset.bullet = para.bullet;
    p.style.textAlign = CSS_ALIGN[para.align];
    // Hanging indent: wrapped lines start at marL, the first line (bullet) at marL + indent.
    p.style.paddingLeft = `${para.marL * pxPerEmu}px`;
    p.style.textIndent = `${para.indent * pxPerEmu}px`;
    p.style.setProperty("--pptx-bullet-width", `${Math.abs(para.indent) * pxPerEmu}px`);
    styleRun(p, para.defaults, pxPerPt);
    // PowerPoint sizes lines from the text itself (single = 1.2 x font size), not the level default.
    const size = para.runs.length ? Math.max(...para.runs.map((r) => r.sz)) : (para.endSz ?? para.defaults.sz);
    const single = 1.2 * size * pxPerPt;
    const space = (s: Spacing | null | undefined, reduce = 0) => (!s ? 0 : "pts" in s ? s.pts * pxPerPt : (s.pct - reduce) * single);
    p.style.lineHeight = `${para.lnSpc ? space(para.lnSpc, lnSpcReduction) : (1 - lnSpcReduction) * single}px`;
    // PowerPoint ignores space before the first paragraph of a text box.
    if (para.spcBef && index > 0) p.style.marginTop = `${space(para.spcBef)}px`;
    if (para.spcAft) p.style.marginBottom = `${space(para.spcAft)}px`;
    if (para.bullet) p.style.setProperty("--pptx-bullet-size", `${(para.runs[0]?.sz ?? size) * pxPerPt}px`);

    for (const run of para.runs) {
      run.text.split("\n").forEach((segment, i) => {
        if (i > 0) p.createEl("br");
        if (!segment) return;
        const span = p.createSpan();
        span.textContent = segment;
        styleRun(span, diff(run, para.defaults), pxPerPt);
        const ls = letterSpacing(segment, run, run.sz * pxPerPt);
        if (ls !== null && Math.abs(ls) > 0.01) span.style.letterSpacing = `${ls}px`;
      });
    }
    if (!p.querySelector("span")) p.createEl("br");
  });
  markEmptyParagraphs(host);
}

/** Flag paragraphs without text: PowerPoint draws no bullet for them. */
export function markEmptyParagraphs(host: HTMLElement): void {
  for (const p of Array.from(host.children)) p.toggleAttribute("data-empty", !p.textContent);
}

function cssColorToHex(value: string): string | null {
  if (/^#[0-9a-f]{6}$/i.test(value)) return value.toUpperCase();
  if (/^#[0-9a-f]{3}$/i.test(value)) return ("#" + value.slice(1).replace(/./g, "$&$&")).toUpperCase();
  const m = /rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)/.exec(value);
  return m ? "#" + m.slice(1, 4).map((n) => Number(n).toString(16).padStart(2, "0")).join("").toUpperCase() : null;
}

/** Formatting set explicitly on `el` and its ancestors below `stop`; closest wins. */
function explicitProps(el: Element | null, stop: Element, pxPerPt: number): Partial<RunProps> {
  const out: Partial<RunProps> = {};
  const set = <K extends keyof RunProps>(k: K, v: RunProps[K] | null | undefined) => {
    if (v !== null && v !== undefined && out[k] === undefined) out[k] = v;
  };
  for (; el && el !== stop; el = el.parentElement) {
    const tag = el.tagName;
    if (tag === "B" || tag === "STRONG") set("b", true);
    if (tag === "I" || tag === "EM") set("i", true);
    if (tag === "U") set("u", true);
    if (tag === "S" || tag === "STRIKE" || tag === "DEL") set("strike", true);
    if (tag === "FONT") {
      set("color", cssColorToHex(el.getAttribute("color") ?? ""));
      set("font", el.getAttribute("face")?.split(",")[0].trim() || null);
    }
    const style = (el as HTMLElement).style;
    if (!style) continue;
    const weight = style.fontWeight;
    if (weight) set("b", weight === "bold" || weight === "bolder" || Number(weight) >= 600);
    if (style.fontStyle) set("i", style.fontStyle !== "normal");
    const deco = style.getPropertyValue("text-decoration-line") || style.getPropertyValue("text-decoration");
    if (deco) {
      set("u", deco.includes("underline"));
      set("strike", deco.includes("line-through"));
    }
    if (style.color) set("color", cssColorToHex(style.color));
    if (style.fontFamily) set("font", style.fontFamily.split(",")[0].trim().replace(/^["']|["']$/g, ""));
    const size = style.fontSize;
    if (size.endsWith("px")) set("sz", Math.round((parseFloat(size) / pxPerPt) * 100) / 100);
    else if (size.endsWith("pt")) set("sz", parseFloat(size));
  }
  return out;
}

const sameProps = (a: RunProps, b: RunProps) => PROP_KEYS.every((k) => a[k] === b[k]);

/** Node-type check that works across windows (unlike instanceof). */
const isElement = (n: Node): n is Element => n.nodeType === Node.ELEMENT_NODE;

/** Top-level blocks of the editor; loose inline nodes are grouped into one paragraph. */
function paragraphGroups(host: HTMLElement): { block: Element | null; nodes: Node[] }[] {
  const groups: { block: Element | null; nodes: Node[] }[] = [];
  let loose: Node[] | null = null;
  host.childNodes.forEach((node) => {
    if (isElement(node) && BLOCK.has(node.tagName)) {
      loose = null;
      groups.push({ block: node, nodes: [node] });
    } else {
      if (!loose) groups.push({ block: null, nodes: (loose = []) });
      loose.push(node);
    }
  });
  return groups;
}

function readGroup(host: HTMLElement, block: Element | null, nodes: Node[], pxPerPt: number, defaultsFor: DefaultsFor): ParaInput {
  const data = (block as HTMLElement | null)?.dataset;
  const lvl = Number(data?.lvl) || 0;
  const defaults = defaultsFor(lvl);
  const stop = block ?? host;
  const runs: Run[] = [];

  const visit = (node: Node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      if (node.textContent) runs.push({ ...defaults, ...explicitProps(node.parentElement, stop, pxPerPt), text: node.textContent });
    } else if (isElement(node)) {
      if (node.tagName === "BR") runs.push({ ...defaults, ...explicitProps(node.parentElement, stop, pxPerPt), text: "\n" });
      else node.childNodes.forEach(visit);
    }
  };
  nodes.forEach((n) => (n === block ? n.childNodes.forEach(visit) : visit(n)));

  // A trailing <br> is the browser's placeholder that keeps an empty line visible.
  if (runs.at(-1)?.text === "\n") runs.pop();

  const merged: Run[] = [];
  for (const run of runs) {
    const prev = merged.at(-1);
    if (prev && prev.text !== "\n" && run.text !== "\n" && sameProps(prev, run)) prev.text += run.text;
    else merged.push(run);
  }

  return {
    runs: merged,
    align: FROM_CSS_ALIGN[(block as HTMLElement | null)?.style.textAlign ?? ""] ?? "l",
    lvl,
    bullet: data?.bullet || null,
  };
}

export function readParagraphs(host: HTMLElement, pxPerPt: number, defaultsFor: DefaultsFor): ParaInput[] {
  return paragraphGroups(host).map(({ block, nodes }) => readGroup(host, block, nodes, pxPerPt, defaultsFor));
}

// -----------------------------------------------------------------------------
// Editing commands on the current selection
// -----------------------------------------------------------------------------

function rangeIn(host: HTMLElement): Range | null {
  const sel = window.getSelection();
  if (!sel || !sel.rangeCount) return null;
  const range = sel.getRangeAt(0);
  return host.contains(range.commonAncestorContainer) ? range : null;
}

function selectedTextNodes(host: HTMLElement, range: Range): Text[] {
  const walker = document.createTreeWalker(host, NodeFilter.SHOW_TEXT);
  const out: Text[] = [];
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (n.textContent && range.intersectsNode(n)) out.push(n as Text);
  }
  return out;
}

/** Apply run formatting to the selected text, splitting text nodes at the edges. */
export function applyRunStyle(host: HTMLElement, patch: Partial<RunProps>, pxPerPt: number): void {
  const range = rangeIn(host);
  if (!range || range.collapsed) return;
  const spans: HTMLElement[] = [];

  for (let node of selectedTextNodes(host, range)) {
    const start = node === range.startContainer ? range.startOffset : 0;
    const end = node === range.endContainer ? range.endOffset : node.length;
    if (start >= end) continue;
    if (end < node.length) node.splitText(end);
    if (start > 0) node = node.splitText(start);

    const current = explicitProps(node.parentElement, host, pxPerPt);
    const span = createSpan();
    node.parentNode!.insertBefore(span, node);
    span.appendChild(node);
    styleRun(span, patch, pxPerPt, { u: current.u ?? false, strike: current.strike ?? false });
    spans.push(span);
  }

  if (spans.length) {
    const sel = window.getSelection()!;
    const r = document.createRange();
    r.setStartBefore(spans[0]);
    r.setEndAfter(spans[spans.length - 1]);
    sel.removeAllRanges();
    sel.addRange(r);
  }
}

/** Formatting shared by the whole selection (booleans are true only if true everywhere). */
export function selectionRunProps(host: HTMLElement, pxPerPt: number, defaultsFor: DefaultsFor): RunProps | null {
  const range = rangeIn(host);
  if (!range) return null;
  const nodes = selectedTextNodes(host, range);
  const props = nodes.map((n) => {
    const block = paragraphGroups(host).find((g) => g.nodes.some((x) => x === n || x.contains(n)))?.block ?? null;
    return { ...defaultsFor(Number((block as HTMLElement | null)?.dataset.lvl) || 0), ...explicitProps(n.parentElement, block ?? host, pxPerPt) };
  });
  if (!props.length) return null;
  return {
    ...props[0],
    b: props.every((p) => p.b),
    i: props.every((p) => p.i),
    u: props.every((p) => p.u),
    strike: props.every((p) => p.strike),
  };
}

type ParaAttrs = Pick<ParaInput, "align" | "lvl" | "bullet">;

/** Update alignment / level / bullet of every paragraph touched by the selection. */
export function updateParagraphs(host: HTMLElement, fn: (p: ParaAttrs) => ParaAttrs): void {
  const range = rangeIn(host);
  if (!range) return;
  for (const p of Array.from(host.children) as HTMLElement[]) {
    if (!BLOCK.has(p.tagName) || !range.intersectsNode(p)) continue;
    const next = fn({
      align: FROM_CSS_ALIGN[p.style.textAlign] ?? "l",
      lvl: Number(p.dataset.lvl) || 0,
      bullet: p.dataset.bullet || null,
    });
    p.style.textAlign = CSS_ALIGN[next.align];
    p.dataset.lvl = String(Math.max(0, Math.min(8, next.lvl)));
    if (next.bullet) p.dataset.bullet = next.bullet;
    else delete p.dataset.bullet;
  }
}
