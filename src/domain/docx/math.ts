import type { DOMOutputSpec } from "prosemirror-model";

/**
 * Office Math (OMML) equations: drawn as MathML (which the browser lays out: stacked fractions, limits, roots,
 * matrices), and edited as LaTeX, as Google Docs' equation editor takes it. `latexToOmml` writes new XML only
 * when an equation is edited; untouched equations keep their XML.
 */

const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MML = "http://www.w3.org/1998/Math/MathML";

/** LaTeX symbol commands and their characters; the first command for a character is the one written back. */
const SYMBOLS: Record<string, string> = {
  alpha: "α", beta: "β", gamma: "γ", delta: "δ", epsilon: "ϵ", varepsilon: "ε", zeta: "ζ", eta: "η", theta: "θ", vartheta: "ϑ",
  iota: "ι", kappa: "κ", lambda: "λ", mu: "μ", nu: "ν", xi: "ξ", pi: "π", varpi: "ϖ", rho: "ρ", sigma: "σ", varsigma: "ς",
  tau: "τ", upsilon: "υ", phi: "ϕ", varphi: "φ", chi: "χ", psi: "ψ", omega: "ω",
  Gamma: "Γ", Delta: "Δ", Theta: "Θ", Lambda: "Λ", Xi: "Ξ", Pi: "Π", Sigma: "Σ", Upsilon: "Υ", Phi: "Φ", Psi: "Ψ", Omega: "Ω",
  pm: "±", mp: "∓", times: "×", div: "÷", cdot: "⋅", ast: "∗", star: "⋆", circ: "∘", bullet: "∙", oplus: "⊕", otimes: "⊗",
  leq: "≤", le: "≤", geq: "≥", ge: "≥", neq: "≠", ne: "≠", approx: "≈", equiv: "≡", sim: "∼", simeq: "≃", cong: "≅", propto: "∝",
  ll: "≪", gg: "≫", infty: "∞", partial: "∂", nabla: "∇", in: "∈", notin: "∉", ni: "∋", subset: "⊂", subseteq: "⊆",
  supset: "⊃", supseteq: "⊇", cup: "∪", cap: "∩", setminus: "∖", forall: "∀", exists: "∃", neg: "¬", wedge: "∧", land: "∧",
  vee: "∨", lor: "∨", to: "→", rightarrow: "→", leftarrow: "←", gets: "←", leftrightarrow: "↔", Rightarrow: "⇒",
  Leftarrow: "⇐", Leftrightarrow: "⇔", mapsto: "↦", uparrow: "↑", downarrow: "↓", cdots: "⋯", ldots: "…", vdots: "⋮",
  ddots: "⋱", angle: "∠", perp: "⊥", parallel: "∥", emptyset: "∅", varnothing: "∅", hbar: "ℏ", ell: "ℓ", Re: "ℜ", Im: "ℑ",
  aleph: "ℵ", prime: "′", degree: "°", therefore: "∴", because: "∵", langle: "⟨", rangle: "⟩", lfloor: "⌊", rfloor: "⌋",
  lceil: "⌈", rceil: "⌉", quad: " ", ",": " ", ";": " ", ":": " ", "!": "", " ": " ",
};
const CHAR_CMD = new Map<string, string>();
for (const [cmd, ch] of Object.entries(SYMBOLS)) if (ch && /^[a-zA-Z]+$/.test(cmd) && !CHAR_CMD.has(ch)) CHAR_CMD.set(ch, cmd);
CHAR_CMD.set(" ", ",");

const NARY: Record<string, string> = { sum: "∑", prod: "∏", coprod: "∐", int: "∫", iint: "∬", iiint: "∭", oint: "∮", bigcup: "⋃", bigcap: "⋂", bigvee: "⋁", bigwedge: "⋀" };
const NARY_CMD = new Map(Object.entries(NARY).map(([k, v]) => [v, k]));
const INTEGRALS = new Set(["∫", "∬", "∭", "∮"]);
/** Combining accent characters (OMML acc) with their LaTeX command and a spacing form to draw. */
const ACCENTS: [string, string, string][] = [
  ["hat", "̂", "^"],
  ["check", "̌", "ˇ"],
  ["tilde", "̃", "~"],
  ["bar", "̄", "¯"],
  ["vec", "⃗", "→"],
  ["dot", "̇", "˙"],
  ["ddot", "̈", "¨"],
  ["acute", "́", "´"],
  ["grave", "̀", "`"],
];
const FUNCS = new Set(["sin", "cos", "tan", "cot", "sec", "csc", "sinh", "cosh", "tanh", "arcsin", "arccos", "arctan", "log", "ln", "lg", "exp", "lim", "max", "min", "sup", "inf", "det", "gcd", "deg", "dim", "ker", "arg", "Pr"]);
const ESCAPED = "{}&%#_$";

const kids = (el: Element | null | undefined, local?: string) => (el ? Array.from(el.children).filter((c) => c.namespaceURI === M && (!local || c.localName === local)) : []);
const kid = (el: Element | null | undefined, local: string) => kids(el, local)[0] ?? null;
const val = (el: Element | null | undefined, pr: string, local: string): string | null => {
  const v = kid(kid(el, pr), local);
  return v ? (v.getAttributeNS(M, "val") ?? v.getAttribute("m:val") ?? "") : null;
};
const on = (v: string | null) => v !== null && v !== "0" && v !== "off" && v !== "false";
const parse = (xml: string) => new DOMParser().parseFromString(xml, "application/xml").documentElement;
/** Plain text runs: m:nor, or m:sty plain. */
const isText = (r: Element) => !!kid(kid(r, "rPr"), "nor") || val(r, "rPr", "sty") === "p";

// ----------------------------------------------------------------------------------------------- MathML

type Spec = DOMOutputSpec;

export function mathSpec(xml: string): DOMOutputSpec {
  const root = parse(xml);
  const display = root.localName === "oMathPara";
  const maths = display ? kids(root, "oMath") : [root];
  return ["span", { class: display ? "docx-math is-display" : "docx-math" }, ...maths.map((m): Spec => [`${MML} math`, display ? { display: "block" } : {}, ...mml(m)])];
}

/** Text split into identifiers, numbers and operators. */
function tokens(text: string, upright: boolean): Spec[] {
  if (upright) return [["mi", { mathvariant: "normal" }, text]];
  const parts: string[] = text.match(/\d+(?:\.\d+)?|\p{L}|\s+|./gu) ?? [];
  return parts.flatMap((t): Spec[] => {
    if (/^\s+$/.test(t)) return [["mspace", { width: "0.25em" }]];
    if (/^\d/.test(t)) return [["mn", t]];
    if (/\p{L}/u.test(t)) return [["mi", t]];
    return [["mo", t]];
  });
}

function mml(el: Element | null, upright = false): Spec[] {
  if (!el) return [];
  const row = (e: Element | null, up = upright): Spec => ["mrow", ...kids(e).flatMap((c) => mml(c, up))];
  const all = () => kids(el).flatMap((c) => mml(c, upright));
  switch (el.localName) {
    case "r": {
      const text = kids(el, "t").map((t: Element) => t.textContent ?? "").join("");
      return isText(el) ? [["mtext", text]] : tokens(text, upright);
    }
    case "sSub":
      return [["msub", row(kid(el, "e")), row(kid(el, "sub"))]];
    case "sSup":
      return [["msup", row(kid(el, "e")), row(kid(el, "sup"))]];
    case "sSubSup":
      return [["msubsup", row(kid(el, "e")), row(kid(el, "sub")), row(kid(el, "sup"))]];
    case "sPre":
      return [["mmultiscripts", row(kid(el, "e")), ["mprescripts"], row(kid(el, "sub")), row(kid(el, "sup"))]];
    case "f": {
      const type = val(el, "fPr", "type");
      if (type === "lin") return [["mrow", row(kid(el, "num")), ["mo", "/"], row(kid(el, "den"))]];
      return [["mfrac", type === "noBar" ? { linethickness: "0" } : {}, row(kid(el, "num")), row(kid(el, "den"))]];
    }
    case "rad": {
      const deg = kid(el, "deg");
      return on(val(el, "radPr", "degHide")) || !deg?.children.length ? [["msqrt", row(kid(el, "e"))]] : [["mroot", row(kid(el, "e")), row(deg)]];
    }
    case "nary": {
      const chr = val(el, "naryPr", "chr") || "∫";
      const op: Spec = ["mo", { largeop: "true" }, chr];
      const sub = on(val(el, "naryPr", "subHide")) ? null : kid(el, "sub");
      const sup = on(val(el, "naryPr", "supHide")) ? null : kid(el, "sup");
      const under = (val(el, "naryPr", "limLoc") ?? (INTEGRALS.has(chr) ? "subSup" : "undOvr")) === "undOvr";
      const has = (e: Element | null) => !!e?.children.length;
      const big: Spec =
        has(sub) && has(sup)
          ? [under ? "munderover" : "msubsup", op, row(sub), row(sup)]
          : has(sub)
            ? [under ? "munder" : "msub", op, row(sub)]
            : has(sup)
              ? [under ? "mover" : "msup", op, row(sup)]
              : op;
      return [big, row(kid(el, "e"))];
    }
    case "d": {
      const beg = val(el, "dPr", "begChr") ?? "(";
      const end = val(el, "dPr", "endChr") ?? ")";
      const sep = val(el, "dPr", "sepChr") ?? "|";
      const parts = kids(el, "e").flatMap((e, i): Spec[] => (i ? [["mo", sep], row(e)] : [row(e)]));
      return [["mrow", ...(beg ? [["mo", beg] as Spec] : []), ...parts, ...(end ? [["mo", end] as Spec] : [])]];
    }
    case "acc": {
      const chr = val(el, "accPr", "chr") ?? "̂";
      const shown = ACCENTS.find((a) => a[1] === chr)?.[2] ?? chr;
      return [["mover", { accent: "true" }, row(kid(el, "e")), ["mo", shown]]];
    }
    case "bar":
      return val(el, "barPr", "pos") === "top" ? [["mover", { accent: "true" }, row(kid(el, "e")), ["mo", "¯"]]] : [["munder", row(kid(el, "e")), ["mo", "_"]]];
    case "func":
      return [["mrow", row(kid(el, "fName"), true), ["mo", "⁡"], row(kid(el, "e"))]];
    case "limLow":
      return [["munder", row(kid(el, "e")), row(kid(el, "lim"), false)]];
    case "limUpp":
      return [["mover", row(kid(el, "e")), row(kid(el, "lim"), false)]];
    case "groupChr": {
      const chr = val(el, "groupChrPr", "chr") ?? "⏟";
      return [[val(el, "groupChrPr", "pos") === "top" ? "mover" : "munder", row(kid(el, "e")), ["mo", chr]]];
    }
    case "m":
      return [["mtable", ...kids(el, "mr").map((mr): Spec => ["mtr", ...kids(mr, "e").map((e): Spec => ["mtd", row(e)])])]];
    case "eqArr":
      return [["mtable", ...kids(el, "e").map((e): Spec => ["mtr", ["mtd", row(e)]])]];
    default:
      return el.localName.endsWith("Pr") ? [] : all();
  }
}

// ----------------------------------------------------------------------------------------------- OMML -> LaTeX

/** An equation as LaTeX; a display equation's lines are joined with \\. */
export function ommlToLatex(xml: string): string {
  const root = parse(xml);
  return root.localName === "oMathPara" ? kids(root, "oMath").map((m) => tex(m)).join("\\\\") : tex(root);
}

const group = (s: string) => `{${s}}`;
/** A script: one character bare, else braced. */
const script = (s: string) => (Array.from(s).length === 1 ? s : group(s));
const runText = (s: string) =>
  Array.from(s)
    .map((ch) => {
      const cmd = CHAR_CMD.get(ch);
      if (cmd) return /^[a-zA-Z]+$/.test(cmd) ? `\\${cmd} ` : `\\${cmd}`;
      return ESCAPED.includes(ch) ? `\\${ch}` : ch;
    })
    .join("");
const delim = (ch: string) => (ch === "" ? "." : ch === "{" || ch === "}" ? `\\${ch}` : (CHAR_CMD.has(ch) ? `\\${CHAR_CMD.get(ch)} ` : ch));

function tex(el: Element | null): string {
  if (!el) return "";
  const all = (e: Element | null) => kids(e).map((c) => tex(c)).join("");
  const plain = (e: Element | null) => (e?.textContent ?? "").trim();
  switch (el.localName) {
    case "r": {
      const text = kids(el, "t").map((t: Element) => t.textContent ?? "").join("");
      return isText(el) ? `\\text{${text}}` : runText(text);
    }
    case "sSub":
    case "sSup":
    case "sSubSup": {
      const base = all(kid(el, "e"));
      const b = Array.from(base).length === 1 || /^\\[a-zA-Z]+ ?$/.test(base) ? base : group(base);
      const sub = kid(el, "sub") ? `_${script(all(kid(el, "sub")))}` : "";
      const sup = kid(el, "sup") ? `^${script(all(kid(el, "sup")))}` : "";
      return b + sub + sup;
    }
    case "f":
      return `\\frac{${all(kid(el, "num"))}}{${all(kid(el, "den"))}}`;
    case "rad": {
      const deg = all(kid(el, "deg"));
      return on(val(el, "radPr", "degHide")) || !deg ? `\\sqrt{${all(kid(el, "e"))}}` : `\\sqrt[${deg}]{${all(kid(el, "e"))}}`;
    }
    case "nary": {
      const chr = val(el, "naryPr", "chr") || "∫";
      const sub = on(val(el, "naryPr", "subHide")) ? "" : all(kid(el, "sub"));
      const sup = on(val(el, "naryPr", "supHide")) ? "" : all(kid(el, "sup"));
      const op = NARY_CMD.has(chr) ? `\\${NARY_CMD.get(chr)}` : runText(chr);
      return `${op}${sub ? `_{${sub}}` : ""}${sup ? `^{${sup}}` : ""}{${all(kid(el, "e"))}}`;
    }
    case "d": {
      const sep = val(el, "dPr", "sepChr") ?? "|";
      const body = kids(el, "e")
        .map((e) => tex(e))
        .join(sep);
      return `\\left${delim(val(el, "dPr", "begChr") ?? "(")}${body}\\right${delim(val(el, "dPr", "endChr") ?? ")")}`;
    }
    case "acc": {
      const chr = val(el, "accPr", "chr") ?? "̂";
      return `\\${ACCENTS.find((a) => a[1] === chr)?.[0] ?? "hat"}{${all(kid(el, "e"))}}`;
    }
    case "bar":
      return `\\${val(el, "barPr", "pos") === "top" ? "overline" : "underline"}{${all(kid(el, "e"))}}`;
    case "func": {
      const fName = kid(el, "fName");
      const low = kid(fName, "limLow");
      const name = plain(low ? kid(low, "e") : fName);
      const head = low ? `${FUNCS.has(name) ? `\\${name}` : `\\operatorname{${name}}`}_{${all(kid(low, "lim"))}}` : FUNCS.has(name) ? `\\${name}` : `\\operatorname{${name}}`;
      return `${head}{${all(kid(el, "e"))}}`;
    }
    case "limLow":
      return `\\underset{${all(kid(el, "lim"))}}{${all(kid(el, "e"))}}`;
    case "limUpp":
      return `\\overset{${all(kid(el, "lim"))}}{${all(kid(el, "e"))}}`;
    case "m":
      return `\\begin{matrix}${kids(el, "mr")
        .map((mr) => kids(mr, "e").map((e) => all(e)).join("&"))
        .join("\\\\")}\\end{matrix}`;
    case "eqArr":
      return `\\begin{gathered}${kids(el, "e")
        .map((e) => all(e))
        .join("\\\\")}\\end{gathered}`;
    default:
      return el.localName.endsWith("Pr") ? "" : all(el);
  }
}

// ----------------------------------------------------------------------------------------------- LaTeX -> OMML

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const tag = (name: string, ...body: string[]) => `<m:${name}>${body.join("")}</m:${name}>`;
const prop = (name: string, v: string) => `<m:${name} m:val="${esc(v)}"/>`;
const FONT = `<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>`;
const run = (text: string, upright = false) => (text ? `<m:r>${upright ? "<m:rPr><m:nor/></m:rPr>" : ""}${FONT}<m:t xml:space="preserve">${esc(text)}</m:t></m:r>` : "");

/** Parsed content: plain characters (merged into runs) or finished OMML. */
type Item = { text: string } | { xml: string };
const xmlOf = (items: Item[]) => items.map((i) => ("xml" in i ? i.xml : run(i.text))).join("");

class LatexParser {
  private i = 0;
  constructor(private s: string) {}

  private peek(): string {
    return this.s[this.i] ?? "";
  }
  private skipSpace(): void {
    while (/\s/.test(this.peek())) this.i++;
  }
  /** A command name after a backslash: letters, or one other character. */
  private command(): string {
    const m = /^[a-zA-Z]+/.exec(this.s.slice(this.i));
    const name = m ? m[0] : this.s[this.i] ?? "";
    this.i += name.length;
    if (m && this.peek() === " ") this.i++;
    return name;
  }
  private lookingAt(str: string): boolean {
    return this.s.startsWith(str, this.i);
  }

  /** Items up to the end, a closing brace, \right, \end, & or \\ (not consumed). */
  seq(): Item[] {
    const out: Item[] = [];
    const push = (it: Item) => {
      const last = out[out.length - 1];
      if ("text" in it && last && "text" in last) last.text += it.text;
      else out.push(it);
    };
    for (;;) {
      this.skipSpace();
      const c = this.peek();
      if (!c || c === "}" || (c === "&" && this.inTable) || this.lookingAt("\\right") || this.lookingAt("\\end") || (this.lookingAt("\\\\") && this.inTable)) return out;
      if (c === "^" || c === "_") {
        // A script goes on the last character or item before it.
        const last = out.pop();
        let base = "";
        if (last && "text" in last) {
          const chars = Array.from(last.text);
          base = run(chars.pop()!);
          if (chars.length) out.push({ text: chars.join("") });
        } else if (last) base = last.xml;
        push({ xml: this.scripts(base) });
        continue;
      }
      push(this.atom());
    }
  }

  private inTable = false;

  /** `^x`, `_x` or both after a base, as sSub / sSup / sSubSup. */
  private scripts(base: string): string {
    const { sub, sup } = this.limits();
    const e = tag("e", base);
    if (sub !== null && sup !== null) return tag("sSubSup", e, tag("sub", sub), tag("sup", sup));
    return sub !== null ? tag("sSub", e, tag("sub", sub)) : tag("sSup", e, tag("sup", sup ?? ""));
  }

  private limits(): { sub: string | null; sup: string | null } {
    let sub: string | null = null;
    let sup: string | null = null;
    for (;;) {
      this.skipSpace();
      const c = this.peek();
      if (c === "_" && sub === null) {
        this.i++;
        sub = this.arg();
      } else if (c === "^" && sup === null) {
        this.i++;
        sup = this.arg();
      } else return { sub, sup };
    }
  }

  /** A braced group or a single atom, as OMML. */
  arg(): string {
    this.skipSpace();
    if (this.peek() === "{") return xmlOf(this.group());
    if (!this.peek() || this.peek() === "}") return "";
    return xmlOf([this.atom()]);
  }

  private group(): Item[] {
    this.i++; // {
    const items = this.seq();
    if (this.peek() === "}") this.i++;
    return items;
  }

  /** Raw text up to the matching brace (for \text, \operatorname, \begin). */
  private raw(): string {
    this.skipSpace();
    if (this.peek() !== "{") return "";
    let depth = 0;
    const start = ++this.i;
    for (; this.i < this.s.length; this.i++) {
      if (this.s[this.i] === "{") depth++;
      else if (this.s[this.i] === "}" && depth-- === 0) break;
    }
    const text = this.s.slice(start, this.i);
    if (this.peek() === "}") this.i++;
    return text;
  }

  private delimiter(): string {
    this.skipSpace();
    const c = this.peek();
    if (c === "\\") {
      this.i++;
      const name = this.command();
      return name === "{" || name === "}" || name === "|" ? name : (SYMBOLS[name] ?? "");
    }
    this.i++;
    return c === "." ? "" : c;
  }

  private atom(): Item {
    const c = this.peek();
    if (c === "{") return { xml: xmlOf(this.group()) };
    if (c !== "\\") {
      this.i++;
      return { text: c };
    }
    this.i++;
    const name = this.command();
    if (name in SYMBOLS) return { text: SYMBOLS[name] };
    if (name.length === 1 && ESCAPED.includes(name)) return { text: name };
    if (name === "frac" || name === "dfrac" || name === "tfrac") return { xml: tag("f", tag("num", this.arg()), tag("den", this.arg())) };
    if (name === "sqrt") {
      this.skipSpace();
      let deg = "";
      if (this.peek() === "[") {
        const end = this.s.indexOf("]", this.i);
        deg = xmlOf(new LatexParser(this.s.slice(this.i + 1, end < 0 ? undefined : end)).seq());
        this.i = end < 0 ? this.s.length : end + 1;
      }
      return { xml: tag("rad", deg ? "" : tag("radPr", prop("degHide", "1")), tag("deg", deg), tag("e", this.arg())) };
    }
    if (name in NARY) {
      const { sub, sup } = this.limits();
      const pr = tag("naryPr", prop("chr", NARY[name]), sub === null ? prop("subHide", "1") : "", sup === null ? prop("supHide", "1") : "");
      return { xml: tag("nary", pr, tag("sub", sub ?? ""), tag("sup", sup ?? ""), tag("e", this.arg())) };
    }
    if (name === "left") {
      const beg = this.delimiter();
      const body = xmlOf(this.seq());
      let end = ")";
      if (this.lookingAt("\\right")) {
        this.i += 6;
        end = this.delimiter();
      }
      const pr = [beg !== "(" ? prop("begChr", beg) : "", end !== ")" ? prop("endChr", end) : ""].join("");
      return { xml: tag("d", pr ? tag("dPr", pr) : "", tag("e", body)) };
    }
    const accent = ACCENTS.find((a) => a[0] === name || (name === "widehat" && a[0] === "hat") || (name === "widetilde" && a[0] === "tilde"));
    if (accent) return { xml: tag("acc", accent[0] === "hat" ? "" : tag("accPr", prop("chr", accent[1])), tag("e", this.arg())) };
    if (name === "overline" || name === "underline") return { xml: tag("bar", tag("barPr", prop("pos", name === "overline" ? "top" : "bot")), tag("e", this.arg())) };
    if (name === "underset" || name === "overset") {
      const lim = this.arg();
      return { xml: tag(name === "underset" ? "limLow" : "limUpp", tag("e", this.arg()), tag("lim", lim)) };
    }
    if (FUNCS.has(name) || name === "operatorname") return { xml: this.func(name === "operatorname" ? this.raw() : name) };
    if (name === "text" || name === "mathrm" || name === "textrm") return { xml: run(this.raw(), true) };
    if (name === "begin") return { xml: this.environment(this.raw()) };
    // Unknown command: keep what was typed.
    return { text: `\\${name}` };
  }

  private func(name: string): string {
    this.skipSpace();
    let fName = run(name, true);
    if (this.peek() === "_") {
      this.i++;
      fName = tag("limLow", tag("e", fName), tag("lim", this.arg()));
    }
    return tag("func", tag("fName", fName), tag("e", this.arg()));
  }

  private environment(env: string): string {
    const was = this.inTable;
    this.inTable = true;
    const rows: string[][] = [[]];
    for (;;) {
      rows[rows.length - 1].push(xmlOf(this.seq()));
      if (this.peek() === "&") this.i++;
      else if (this.lookingAt("\\\\")) {
        this.i += 2;
        rows.push([]);
      } else break;
    }
    this.inTable = was;
    if (this.lookingAt("\\end")) {
      this.i += 4;
      this.raw();
    }
    if (env === "gathered" || env === "aligned" || env === "cases") {
      const arr = tag("eqArr", ...rows.map((r) => tag("e", r.join(run(" ")))));
      return env === "cases" ? tag("d", tag("dPr", prop("begChr", "{"), prop("endChr", "")), tag("e", arr)) : arr;
    }
    const cols = Math.max(...rows.map((r) => r.length));
    const m = tag(
      "m",
      tag("mPr", tag("mcs", tag("mc", tag("mcPr", prop("count", String(cols)), prop("mcJc", "center"))))),
      ...rows.map((r) => tag("mr", ...Array.from({ length: cols }, (_, k) => tag("e", r[k] ?? ""))))
    );
    const fence: Record<string, [string, string]> = { pmatrix: ["(", ")"], bmatrix: ["[", "]"], Bmatrix: ["{", "}"], vmatrix: ["|", "|"], Vmatrix: ["‖", "‖"] };
    const f = fence[env];
    if (!f) return m;
    const pr = [f[0] !== "(" ? prop("begChr", f[0]) : "", f[1] !== ")" ? prop("endChr", f[1]) : ""].join("");
    return tag("d", pr ? tag("dPr", pr) : "", tag("e", m));
  }
}

/** LaTeX as an OMML equation (m:oMath), or a display equation (m:oMathPara) with one line per \\. */
export function latexToOmml(latex: string, display = false): string {
  const ns = `xmlns:m="${M}" xmlns:w="${W}"`;
  const one = (src: string) => xmlOf(new LatexParser(src).seq());
  if (!display) return `<m:oMath ${ns}>${one(latex)}</m:oMath>`;
  const lines = latex.split(/\\\\(?![^{]*\})/);
  return `<m:oMathPara ${ns}>${lines.map((l) => tag("oMath", one(l))).join("")}</m:oMathPara>`;
}
