/*
 * Paragraph model <-> contentEditable HTML. Uses plain DOM APIs only (no
 * Obsidian helpers) so it runs under jsdom in tests.
 *
 * Each paragraph is a <p data-lvl data-bullet> styled with its level defaults;
 * runs are <span>s styled only where they differ from those defaults. Reading
 * back resolves unstyled text against the (possibly changed) level's defaults.
 * data-bullet is the model's bullet ("•", "#arabicPeriod"); data-label is what
 * is drawn (the char, or the computed number). Fields are non-editable
 * <span data-field>; links are <span data-link> ("" = explicitly no link).
 */
import type { Align, Para, Run, RunProps, Spacing } from "../../domain/pptx/slide-model";
import type { ParaInput } from "../../domain/pptx/slide-edit";
import { fontStack, letterSpacing } from "./text-fit";

const CSS_ALIGN: Record<Align, string> = { l: "left", ctr: "center", r: "right", just: "justify" };
const FROM_CSS_ALIGN: Record<string, Align> = { left: "l", start: "l", center: "ctr", right: "r", end: "r", justify: "just" };
const BLOCK = new Set(["P", "DIV", "LI", "H1", "H2", "H3", "H4", "H5", "H6"]);
const EMU_PER_PT = 12700;
const PROP_KEYS: (keyof RunProps)[] = ["b", "i", "u", "strike", "sz", "color", "font", "spc", "link"];

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
  if (props.spc !== undefined) {
    el.dataset.spc = String(props.spc);
    el.style.letterSpacing = `${props.spc * pxPerPt}px`;
  }
  if (props.link !== undefined) {
    el.dataset.link = props.link ?? "";
    el.classList.toggle("pptx-link", !!props.link);
  }
}

function diff(run: RunProps, base: RunProps): Partial<RunProps> {
  const out: Partial<RunProps> = {};
  for (const k of PROP_KEYS) {
    if (String(run[k]).toUpperCase() !== String(base[k]).toUpperCase()) (out as Record<string, unknown>)[k] = run[k];
  }
  if (out.u !== undefined || out.strike !== undefined) Object.assign(out, { u: run.u, strike: run.strike });
  return out;
}

export interface RenderTextOptions {
  /** Shown by slide-number fields. */
  slideNumber?: number;
}

/** What a field shows: the live slide number or today's date, else its cached text. */
function fieldText(run: Run, opts: RenderTextOptions): string {
  if (run.field === "slidenum" && opts.slideNumber) return String(opts.slideNumber);
  if (run.field?.startsWith("datetime")) return new Date().toLocaleDateString();
  return run.text;
}

export function renderParagraphs(
  host: HTMLElement,
  paras: Para[],
  pxPerPt: number,
  pxPerEmu = pxPerPt / EMU_PER_PT,
  lnSpcReduction = 0,
  opts: RenderTextOptions = {}
): void {
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
    p.dataset.spacing = JSON.stringify({ lnSpc: para.lnSpc ?? null, spcBef: para.spcBef ?? null, spcAft: para.spcAft ?? null });

    for (const run of para.runs) {
      if (run.field !== undefined) {
        const span = p.createSpan({ text: fieldText(run, opts) });
        span.dataset.field = run.field;
        span.contentEditable = "false";
        styleRun(span, diff(run, para.defaults), pxPerPt);
        continue;
      }
      run.text.split("\n").forEach((segment, i) => {
        if (i > 0) p.createEl("br");
        if (!segment) return;
        const span = p.createSpan();
        span.textContent = segment;
        styleRun(span, diff(run, para.defaults), pxPerPt);
        // Match PowerPoint's widths (font fitting) plus the run's own character spacing.
        const ls = (letterSpacing(segment, run, run.sz * pxPerPt) ?? 0) + run.spc * pxPerPt;
        if (Math.abs(ls) > 0.01) span.style.letterSpacing = `${ls}px`;
      });
    }
    if (!p.querySelector("span")) p.createEl("br");
  });
  refreshMarks(host);
}

const ROMAN: [number, string][] = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
const roman = (n: number) => ROMAN.reduce((out, [v, s]) => {
  for (; n >= v; n -= v) out += s;
  return out;
}, "");
const alpha = (n: number) => {
  let out = "";
  for (; n > 0; n = Math.floor((n - 1) / 26)) out = String.fromCharCode(97 + ((n - 1) % 26)) + out;
  return out;
};

/** Label for number `n` in a DrawingML auto-number scheme ("arabicPeriod", "alphaUcParenR"...). */
export function numberLabel(scheme: string, n: number): string {
  const m = /^(arabic|alphaLc|alphaUc|romanLc|romanUc)(Period|ParenR|ParenBoth|Plain)$/.exec(scheme) ?? [scheme, "arabic", "Period"];
  const kind = m[1];
  let text = kind === "arabic" ? String(n) : kind.startsWith("alpha") ? alpha(n) : roman(n);
  if (kind.endsWith("Uc")) text = text.toUpperCase();
  return m[2] === "Period" ? `${text}.` : m[2] === "ParenR" ? `${text})` : m[2] === "ParenBoth" ? `(${text})` : text;
}

/**
 * Recompute what each paragraph's marker shows: nothing for empty paragraphs
 * (PowerPoint draws no bullet there), the bullet character, or the paragraph's
 * number in its level's run of numbered paragraphs.
 */
export function refreshMarks(host: HTMLElement): void {
  const counters: ({ scheme: string; n: number } | null)[] = [];
  for (const p of Array.from(host.children) as HTMLElement[]) {
    const empty = !p.textContent;
    p.toggleAttribute("data-empty", empty);
    const lvl = Number(p.dataset.lvl) || 0;
    counters.length = lvl + 1;
    const bullet = p.dataset.bullet ?? "";
    const auto = /^#(\w+)(?:@(\d+))?$/.exec(bullet);
    if (!auto) {
      counters[lvl] = null;
      if (bullet) p.dataset.label = bullet;
      else delete p.dataset.label;
      continue;
    }
    if (empty) {
      p.dataset.label = "";
      continue;
    }
    const prev = counters[lvl];
    const n = prev && prev.scheme === auto[1] ? prev.n + 1 : Number(auto[2] ?? 1);
    counters[lvl] = { scheme: auto[1], n };
    p.dataset.label = numberLabel(auto[1], n);
  }
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
    const data = (el as HTMLElement).dataset;
    if (data) {
      if (data.spc !== undefined) set("spc", Number(data.spc));
      if (data.link !== undefined && !("link" in out)) out.link = data.link || null;
    }
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
      const field = (node as HTMLElement).dataset?.field;
      if (field !== undefined) runs.push({ ...defaults, ...explicitProps(node, stop, pxPerPt), text: node.textContent ?? "", field });
      else if (node.tagName === "BR") runs.push({ ...defaults, ...explicitProps(node.parentElement, stop, pxPerPt), text: "\n" });
      else node.childNodes.forEach(visit);
    }
  };
  nodes.forEach((n) => (n === block ? n.childNodes.forEach(visit) : visit(n)));

  // A trailing <br> is the browser's placeholder that keeps an empty line visible.
  if (runs.at(-1)?.text === "\n") runs.pop();

  const merged: Run[] = [];
  for (const run of runs) {
    const prev = merged.at(-1);
    const plain = (r: Run) => r.text !== "\n" && r.field === undefined;
    if (prev && plain(prev) && plain(run) && sameProps(prev, run)) prev.text += run.text;
    else merged.push(run);
  }

  return {
    runs: merged,
    align: FROM_CSS_ALIGN[(block as HTMLElement | null)?.style.textAlign ?? ""] ?? "l",
    lvl,
    bullet: data?.bullet || null,
    ...spacingOf(block as HTMLElement | null),
  };
}

/** Line / paragraph spacing stored on a rendered paragraph (absent = leave as is). */
function spacingOf(p: HTMLElement | null): Partial<Pick<ParaInput, "lnSpc" | "spcBef" | "spcAft">> {
  try {
    return p?.dataset.spacing ? (JSON.parse(p.dataset.spacing) as Pick<ParaInput, "lnSpc" | "spcBef" | "spcAft">) : {};
  } catch {
    return {};
  }
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

type ParaAttrs = Pick<ParaInput, "align" | "lvl" | "bullet" | "lnSpc">;

/** Update alignment / level / bullet of every paragraph touched by the selection. */
export function updateParagraphs(host: HTMLElement, fn: (p: ParaAttrs) => ParaAttrs): void {
  const range = rangeIn(host);
  if (!range) return;
  for (const p of Array.from(host.children) as HTMLElement[]) {
    if (!BLOCK.has(p.tagName) || !range.intersectsNode(p)) continue;
    const spacing = spacingOf(p);
    const next = fn({
      align: FROM_CSS_ALIGN[p.style.textAlign] ?? "l",
      lvl: Number(p.dataset.lvl) || 0,
      bullet: p.dataset.bullet || null,
      lnSpc: spacing.lnSpc ?? null,
    });
    p.dataset.spacing = JSON.stringify({ ...spacing, lnSpc: next.lnSpc ?? null });
    p.style.textAlign = CSS_ALIGN[next.align];
    p.dataset.lvl = String(Math.max(0, Math.min(8, next.lvl)));
    if (next.bullet) p.dataset.bullet = next.bullet;
    else delete p.dataset.bullet;
  }
  refreshMarks(host);
}

/** Insert a non-editable field (slide number, date) at the caret. */
export function insertField(host: HTMLElement, type: string, text: string): void {
  const range = rangeIn(host);
  if (!range) return;
  range.deleteContents();
  const span = createSpan({ text });
  span.dataset.field = type;
  span.contentEditable = "false";
  range.insertNode(span);
  range.setStartAfter(span);
  range.collapse(true);
  const sel = window.getSelection()!;
  sel.removeAllRanges();
  sel.addRange(range);
}

/** Replace the current selection with plain text (newlines become <br>) and leave the caret after it. */
export function insertPlainText(text: string): void {
  const sel = window.getSelection();
  if (!sel?.rangeCount) return;
  const range = sel.getRangeAt(0);
  range.deleteContents();
  const frag = createFragment();
  text.replace(/\r\n?/g, "\n").split("\n").forEach((line, i) => {
    if (i > 0) frag.appendChild(createEl("br"));
    if (line) frag.appendChild(document.createTextNode(line));
  });
  const last = frag.lastChild;
  range.insertNode(frag);
  if (last?.nodeType === Node.TEXT_NODE) range.setStart(last, (last as Text).length);
  else if (last) range.setStartAfter(last);
  range.collapse(true);
  sel.removeAllRanges();
  sel.addRange(range);
}
