/*
 * A small Markdown reader for pasting notes into documents and slides: block
 * structure line by line (each non-blank line is its own paragraph, as
 * Obsidian shows single line breaks) and inline emphasis / code / links.
 * Wikilinks and embeds become their display text; vault pictures and math ($…$, $$…$$) are kept as such for the
 * editor to place. No HTML, no Obsidian.
 */
import { IMAGE_MIME } from "./images";

export interface MdRun {
  text: string;
  b?: true;
  i?: true;
  strike?: true;
  code?: true;
  highlight?: true;
  href?: string;
  /** An equation: `text` is its LaTeX. */
  math?: true;
  /** A picture in the vault: its link path as written (`text` is its alt text or name). */
  image?: string;
}

export type MdBlock =
  | { kind: "blank" }
  | { kind: "hr" }
  | { kind: "para"; runs: MdRun[] }
  | { kind: "heading"; level: number; runs: MdRun[] }
  | { kind: "quote"; runs: MdRun[]; /** The first line of a callout ("> [!type]- Title"): its type and fold mark; `runs` is its title. */ callout?: { type: string; fold: string } }
  | { kind: "item"; list: "bullet" | "ordered" | "task"; depth: number; number?: number; checked?: boolean; runs: MdRun[] }
  | { kind: "code"; lang: string; text: string }
  | { kind: "math"; latex: string }
  | { kind: "table"; rows: MdRun[][][] };

type Marks = Omit<MdRun, "text">;
const MARK_KEYS = ["b", "i", "strike", "code", "highlight", "href", "math", "image"] as const;
const DELIMS: [string, Marks][] = [["***", { b: true, i: true }], ["**", { b: true }], ["__", { b: true }], ["~~", { strike: true }], ["==", { highlight: true }], ["*", { i: true }], ["_", { i: true }]];
const ESCAPABLE = /[\\`*_~=[\]()#|>!+.$-]/;
const isSpace = (c: string | undefined) => c === undefined || /\s/.test(c);
const isWord = (c: string | undefined) => c !== undefined && /[\p{L}\p{N}]/u.test(c);

const isPicture = (path: string) => !/^[a-z][a-z0-9+.-]*:/i.test(path) && !!IMAGE_MIME[path.split("#")[0].split(".").pop()!.toLowerCase()];

/** Where math opened with `d` ("$" or "$$") at `from` closes, or -1: no space inside the marks, and no price ("$5 and $10"). */
function mathEnd(s: string, d: string, from: number): number {
  if (isSpace(s[from])) return -1;
  for (let j = s.indexOf(d, from + 1); j >= 0; j = s.indexOf(d, j + 1)) {
    if (isSpace(s[j - 1]) || s[j - 1] === "\\") continue;
    if (d === "$" && /\d/.test(s[j + 1] ?? "")) continue;
    return j;
  }
  return -1;
}

/** Where the delimiter `d` opened at `from` closes, or -1. */
function closing(s: string, d: string, from: number): number {
  for (let j = s.indexOf(d, from); j >= 0; j = s.indexOf(d, j + 1)) {
    if (j === from || isSpace(s[j - 1]) || s[j - 1] === "\\") continue;
    // A single * or _ must not be half of a doubled one.
    if (d.length === 1 && (s[j + 1] === d || s[j - 1] === d)) continue;
    if (d === "_" && isWord(s[j + 1])) continue;
    return j;
  }
  return -1;
}

function walk(s: string, marks: Marks, push: (text: string, marks: Marks) => void): void {
  let buf = "";
  const flush = () => {
    push(buf, marks);
    buf = "";
  };
  let i = 0;
  outer: while (i < s.length) {
    const c = s[i];
    if (c === "\\" && ESCAPABLE.test(s[i + 1] ?? "")) {
      buf += s[i + 1];
      i += 2;
      continue;
    }
    if (c === "`") {
      const end = s.indexOf("`", i + 1);
      if (end > i + 1) {
        flush();
        push(s.slice(i + 1, end), { ...marks, code: true });
        i = end + 1;
        continue;
      }
    }
    if (c === "$") {
      const d = s.startsWith("$$", i) ? "$$" : "$";
      const end = mathEnd(s, d, i + d.length);
      if (end > 0) {
        flush();
        push(s.slice(i + d.length, end).trim(), { ...marks, math: true });
        i = end + d.length;
        continue;
      }
    }
    const embed = s.startsWith("![[", i);
    if (embed || s.startsWith("[[", i)) {
      const start = i + (embed ? 3 : 2);
      const end = s.indexOf("]]", start);
      if (end > start) {
        const [target, alias] = s.slice(start, end).split("|");
        if (embed && isPicture(target)) {
          flush();
          push(target, { ...marks, image: target });
        } else buf += embed || alias === undefined ? target : alias;
        i = end + 2;
        continue;
      }
    }
    const link = /^!?\[([^\]]*)\]\(([^)\s]*)(?:\s+"[^"]*")?\)/.exec(s.slice(i));
    if (link) {
      flush();
      let path = link[2];
      try {
        path = decodeURIComponent(path);
      } catch {
        // kept as written
      }
      if (c === "!" && isPicture(path)) push(link[1] || path, { ...marks, image: path });
      else if (c === "!") push(link[1], marks);
      else walk(link[1], { ...marks, href: link[2] }, push);
      i += link[0].length;
      continue;
    }
    for (const [d, add] of DELIMS) {
      if (!s.startsWith(d, i) || isSpace(s[i + d.length])) continue;
      if (d === "_" && isWord(s[i - 1])) continue;
      const end = closing(s, d, i + d.length);
      if (end < 0) continue;
      flush();
      walk(s.slice(i + d.length, end), { ...marks, ...add }, push);
      i = end + d.length;
      continue outer;
    }
    // An unmatched run of delimiter characters is literal text.
    const run = /^([*_=~])\1*/.exec(s.slice(i));
    buf += run ? run[0] : c;
    i += run ? run[0].length : 1;
  }
  flush();
}

/** Inline Markdown as runs; neighbouring runs with the same marks are merged. */
export function parseInline(src: string): MdRun[] {
  const out: MdRun[] = [];
  walk(src, {}, (text, marks) => {
    if (!text) return;
    const prev = out.at(-1);
    if (prev && !marks.math && !marks.image && MARK_KEYS.every((k) => prev[k] === marks[k])) prev.text += text;
    else out.push({ text, ...marks });
  });
  return out;
}

const FENCE = /^\s*(`{3,}|~{3,})\s*([\w+#-]*)/;
const HR = /^\s{0,3}([-*_])(\s*\1){2,}\s*$/;
const HEADING = /^\s{0,3}(#{1,6})\s+(.*?)(?:\s+#+)?\s*$/;
const ITEM = /^([ \t]*)([-*+]|\d+[.)])\s+(?:\[([ xX])\]\s+)?(.*)$/;
const QUOTE = /^\s*(?:>\s?)+(.*)$/;
const TABLE_RULE = /^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$/;

function tableCells(line: string): string[] {
  const cells: string[] = [];
  let cell = "";
  let s = line.trim().replace(/^\|/, "");
  if (s.endsWith("|") && !s.endsWith("\\|")) s = s.slice(0, -1);
  for (let i = 0; i < s.length; i++) {
    if (s[i] === "\\" && s[i + 1] === "|") {
      cell += "\\|";
      i++;
    } else if (s[i] === "|") {
      cells.push(cell);
      cell = "";
    } else cell += s[i];
  }
  cells.push(cell);
  return cells.map((c) => c.trim());
}

const indentWidth = (s: string) => [...s].reduce((w, c) => w + (c === "\t" ? 4 : 1), 0);

export function parseMarkdown(text: string): MdBlock[] {
  const lines = text.replace(/\r\n?/g, "\n").split("\n");
  const blocks: MdBlock[] = [];
  let stack: number[] = [];
  for (let n = 0; n < lines.length; n++) {
    const line = lines[n];
    const fence = FENCE.exec(line);
    if (fence) {
      const body: string[] = [];
      while (++n < lines.length && !lines[n].trim().startsWith(fence[1])) body.push(lines[n]);
      blocks.push({ kind: "code", lang: fence[2], text: body.join("\n") });
      stack = [];
      continue;
    }
    // Display math: $$ … $$ on one line or over several.
    const t = line.trim();
    if (t.startsWith("$$") && (t.indexOf("$$", 2) < 0 || (t.length > 4 && t.indexOf("$$", 2) === t.length - 2))) {
      if (t.length > 4 && t.indexOf("$$", 2) === t.length - 2) {
        blocks.push({ kind: "math", latex: t.slice(2, -2).trim() });
        stack = [];
        continue;
      }
      const start = n;
      const body = [t.slice(2)];
      while (++n < lines.length) {
        const next = lines[n].trimEnd();
        body.push(next.endsWith("$$") ? next.slice(0, -2) : lines[n]);
        if (next.endsWith("$$")) break;
      }
      if (n < lines.length) {
        blocks.push({ kind: "math", latex: body.join("\n").trim() });
        stack = [];
        continue;
      }
      n = start;
    }
    const item = ITEM.exec(line);
    if (item && !HR.test(line)) {
      const w = indentWidth(item[1]);
      while (stack.length && w < stack[stack.length - 1]) stack.pop();
      if (!stack.length || w > stack[stack.length - 1]) stack.push(w);
      const ordered = /\d/.test(item[2]);
      const task = !ordered && item[3] !== undefined;
      blocks.push({
        kind: "item",
        list: task ? "task" : ordered ? "ordered" : "bullet",
        depth: stack.length - 1,
        ...(ordered ? { number: parseInt(item[2], 10) } : {}),
        ...(task ? { checked: item[3] !== " " } : {}),
        runs: parseInline(ordered && item[3] !== undefined ? `[${item[3]}] ${item[4]}` : item[4]),
      });
      continue;
    }
    stack = [];
    let m: RegExpExecArray | null;
    if (!line.trim()) blocks.push({ kind: "blank" });
    else if (HR.test(line)) blocks.push({ kind: "hr" });
    else if ((m = HEADING.exec(line))) blocks.push({ kind: "heading", level: m[1].length, runs: parseInline(m[2]) });
    else if (line.includes("|") && TABLE_RULE.test(lines[n + 1] ?? "") && lines[n + 1].includes("-")) {
      const rows = [tableCells(line)];
      for (n += 2; n < lines.length && lines[n].includes("|") && lines[n].trim(); n++) rows.push(tableCells(lines[n]));
      n--;
      blocks.push({ kind: "table", rows: rows.map((r) => r.map(parseInline)) });
    } else if ((m = QUOTE.exec(line))) {
      // A callout's [!type] marker gives way to its title (or the type's name).
      const callout = /^\[!([\w-]+)\]([+-]?)\s*/.exec(m[1]);
      const rest = callout ? m[1].slice(callout[0].length) || callout[1][0].toUpperCase() + callout[1].slice(1) : m[1];
      blocks.push({ kind: "quote", runs: parseInline(rest), ...(callout ? { callout: { type: callout[1].toLowerCase(), fold: callout[2] } } : {}) });
    } else blocks.push({ kind: "para", runs: parseInline(line.trim()) });
  }
  return blocks;
}

const BLOCK_SYNTAX = [/^\s{0,3}#{1,6}\s+\S/m, /^\s*([-*+]|\d+[.)])\s+\S/m, /^\s*>/m, /^\s*(```|~~~)/m, /^\s*\|?\s*:?-{3,}:?\s*\|/m, /^\s*\$\$/m];
const INLINE_SYNTAX = /\*\*\S[^*\n]*\*\*|__\S[^_\n]*__|~~\S[^~\n]*~~|==\S[^=\n]*==|`[^`\n]+`|\[[^\]\n]+\]\([^)\s]+\)|\[\[[^\]\n]+\]\]/;
// Inline math counts only when it has TeX in it (a command or a power): "$HOME/bin:$PATH" is a shell line.
const INLINE_MATH = /\$([^\s$](?:[^$\n]*[^\s$\\])?)\$(?!\d)/g;

/** Whether pasted plain text reads as Markdown (so is worth formatting). */
export function looksLikeMarkdown(text: string): boolean {
  return BLOCK_SYNTAX.some((re) => re.test(text)) || INLINE_SYNTAX.test(text) || Array.from(text.matchAll(INLINE_MATH)).some((m) => /\\[A-Za-z]|\^/.test(m[1]));
}

export const runsText = (runs: MdRun[]) => runs.map((r) => r.text).join("");

/** A block's inline runs (a table's cells in reading order). */
export const blockRuns = (b: MdBlock): MdRun[] => (b.kind === "table" ? b.rows.flat(2) : "runs" in b ? b.runs : []);

/** The vault pictures the Markdown shows, as their link paths (each once), for the editor to load before pasting. */
export function markdownImages(text: string): string[] {
  return [...new Set(parseMarkdown(text).flatMap(blockRuns).flatMap((r) => (r.image ? [r.image] : [])))];
}

/** The text without Markdown syntax: bullets, task boxes and heading marks go, numbers stay, one line per line. */
export function markdownToPlain(text: string): string {
  return parseMarkdown(text)
    .map((b) => {
      switch (b.kind) {
        case "blank":
        case "hr":
          return "";
        case "code":
          return b.text;
        case "math":
          return b.latex;
        case "table":
          return b.rows.map((r) => r.map(runsText).join("\t")).join("\n");
        case "item":
          return b.list === "ordered" ? `${b.number}. ${runsText(b.runs)}` : runsText(b.runs);
        default:
          return runsText(b.runs);
      }
    })
    .join("\n");
}
