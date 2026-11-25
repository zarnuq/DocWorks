/*
 * Named ranges, as Google Sheets' Data > Named ranges: list, add, change the range, rename (formulas follow) and
 * delete. Hidden and built-in names (_xlnm.*) are left alone. Named functions (Data > Named functions) are names
 * defined as LAMBDA, as Excel stores them: _xlfn.LAMBDA(_xlpm.a, _xlpm.b, body).
 */
import { NS_S } from "../../infrastructure/xlsx/xlsx-workbook";
import { Edits } from "./commands";
import { renameNameInFormula, tokenize } from "./formula";
import { isBuiltinFunction } from "./formula-evaluator";
import { quoteSheet } from "./references";
import { rewriteFormulas } from "./sheet-edit";
import { Workbook } from "./workbook-model";

interface NamedRange {
  name: string;
  range: string;
}

const userNames = (wb: Workbook) => wb.names().filter((n) => !n.name.startsWith("_xlnm.") && n.el.getAttribute("hidden") !== "1");

const isLambdaText = (formula: string) => /^\s*=?\s*(_xlfn\.)?LAMBDA\s*\(/i.test(formula);

export function listNames(wb: Workbook): NamedRange[] {
  return userNames(wb).filter((n) => !isLambdaText(n.formula)).map((n) => ({ name: n.name, range: n.formula }));
}

export interface NamedFunction {
  name: string;
  /** Argument placeholders. */
  params: string[];
  /** The formula definition, as typed (no _xlpm. / _xlfn. prefixes). */
  body: string;
  description: string;
}

/** Text without the file's prefixes on parameters and newer functions. */
const plain = (text: string) => text.replace(/_xlpm\./gi, "").replace(/(_xlfn\.|_xlws\.)+/gi, "");

export function listFunctions(wb: Workbook): NamedFunction[] {
  const out: NamedFunction[] = [];
  for (const n of userNames(wb)) {
    if (!isLambdaText(n.formula)) continue;
    let toks: ReturnType<typeof tokenize>;
    try {
      toks = tokenize(n.formula.trim().replace(/^=/, ""));
    } catch {
      continue;
    }
    // The LAMBDA's own arguments: split at its top-level commas.
    const args: string[] = [];
    let depth = 0, current = "";
    for (const t of toks.slice(1)) {
      if (t.kind === "punct" && t.text === "(" && depth++ === 0) continue;
      if (t.kind === "punct" && t.text === ")" && --depth === 0) break;
      if (t.kind === "punct" && t.text === "," && depth === 1) {
        args.push(current);
        current = "";
      } else current += t.text;
    }
    args.push(current);
    out.push({ name: n.name, params: args.slice(0, -1).map((a) => plain(a).trim()), body: plain(args[args.length - 1]).trim(), description: n.el.getAttribute("comment") ?? "" });
  }
  return out;
}

const PARAM = /^[A-Za-z_][A-Za-z0-9_]*$/;

/** The stored definition: parameters and the body's uses of them prefixed _xlpm., newer functions _xlfn. */
export function lambdaText(params: string[], body: string): string {
  const keys = new Set(params.map((p) => p.toLowerCase()));
  let marked: string;
  try {
    marked = tokenize(body.trim().replace(/^=/, "")).map((t) => (t.kind === "name" && keys.has(t.text.toLowerCase()) ? `_xlpm.${t.text}` : t.text)).join("");
  } catch (err) {
    throw new Error(`There's a problem with the formula definition: ${(err as Error).message}`);
  }
  const stored = Edits.normalizeFormula(`=${marked}`);
  return `_xlfn.LAMBDA(${[...params.map((p) => `_xlpm.${p}`), stored].join(",")})`;
}

/** Add (`previous` null) or change a named function; renaming rewrites the formulas that call it. */
export function setFunction(e: Edits, previous: string | null, fn: NamedFunction): void {
  validateName(e.wb, fn.name, previous ?? undefined);
  if (isBuiltinFunction(fn.name)) throw new Error(`${fn.name.toUpperCase()} is already a function.`);
  const seen = new Set<string>();
  for (const p of fn.params) {
    if (!PARAM.test(p) || /^[A-Za-z]{1,3}[0-9]+$/.test(p)) throw new Error(`"${p}" can't be an argument placeholder.`);
    if (seen.has(p.toLowerCase())) throw new Error(`"${p}" is used twice.`);
    seen.add(p.toLowerCase());
  }
  if (!fn.body.trim()) throw new Error("Enter a formula definition.");
  const text = lambdaText(fn.params, fn.body);
  if (previous === null) addDefinedName(e, fn.name, text, fn.description);
  else e.structural(() => {
    const el = nameEl(e, previous);
    el.textContent = text;
    if (fn.description) el.setAttribute("comment", fn.description);
    else el.removeAttribute("comment");
    if (previous !== fn.name) {
      el.setAttribute("name", fn.name);
      rewriteFormulas(e, (t) => renameFunctionInFormula(t, previous, fn.name));
    }
    e.wb.pkg.touch(e.wb.pkg.workbookPath);
  });
}

/** Calls of a named function renamed in a formula. */
function renameFunctionInFormula(src: string, from: string, to: string): string {
  try {
    return tokenize(src).map((t) => (t.kind === "func" && t.text.toLowerCase() === from.toLowerCase() ? to : t.text)).join("");
  } catch {
    return src;
  }
}

/** Throws with a message for an invalid or taken name (`except` is the name being renamed). */
export function validateName(wb: Workbook, name: string, except?: string): void {
  if (!name) throw new Error("Enter a name.");
  if (name.length > 255) throw new Error("A name can be at most 255 characters.");
  if (/^[0-9.]/.test(name)) throw new Error("A name must start with a letter or an underscore.");
  if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name)) throw new Error("A name can only contain letters, numbers, underscores and periods.");
  if (/^[A-Za-z]{1,3}[0-9]+$/.test(name) || /^[Rr][0-9]*[Cc][0-9]*$/.test(name)) throw new Error("A name can't look like a cell reference.");
  if (/^(true|false)$/i.test(name)) throw new Error("TRUE and FALSE can't be names.");
  const lower = name.toLowerCase();
  if (lower !== except?.toLowerCase() && wb.names().some((n) => n.name.toLowerCase() === lower)) throw new Error(`The name "${name}" already exists.`);
}

/** "Sheet!A1:B2" with absolute references, or an error for anything that isn't one sheet-qualified range. */
function normalizeRange(range: string): string {
  let toks: ReturnType<typeof tokenize> = [];
  try {
    toks = tokenize(range.trim().replace(/^=/, ""));
  } catch {
    // reported below
  }
  const t = toks[0];
  if (toks.length !== 1 || t.kind !== "ref" || !t.sheet) throw new Error("Enter a range like Sheet1!A1:B10.");
  const local = t.text.slice(t.text.lastIndexOf("!") + 1).replace(/\$?([A-Za-z]+)\$?([0-9]+)/g, (_, c: string, r: string) => `$${c.toUpperCase()}$${r}`);
  return `${quoteSheet(t.sheet)}!${local}`;
}

function nameEl(e: Edits, name: string): Element {
  const found = userNames(e.wb).find((n) => n.name.toLowerCase() === name.toLowerCase());
  if (!found) throw new Error(`No name "${name}".`);
  return found.el;
}

export function addName(e: Edits, name: string, range: string): void {
  validateName(e.wb, name);
  addDefinedName(e, name, normalizeRange(range));
}

function addDefinedName(e: Edits, name: string, text: string, comment = ""): void {
  e.structural(() => {
    const doc = e.wb.pkg.workbook;
    let names = doc.getElementsByTagNameNS(NS_S, "definedNames")[0];
    if (!names) {
      names = doc.createElementNS(NS_S, "definedNames");
      const sheets = doc.getElementsByTagNameNS(NS_S, "sheets")[0];
      sheets.parentNode!.insertBefore(names, sheets.nextSibling);
    }
    const el = doc.createElementNS(NS_S, "definedName");
    el.setAttribute("name", name);
    if (comment) el.setAttribute("comment", comment);
    el.textContent = text;
    // Excel keeps names sorted; inserting in order keeps a later save's diff small.
    const next = Array.from(names.children).find((n) => (n.getAttribute("name") ?? "").toLowerCase() > name.toLowerCase()) ?? null;
    names.insertBefore(el, next);
    e.wb.pkg.touch(e.wb.pkg.workbookPath);
  });
}

export function setNameRange(e: Edits, name: string, range: string): void {
  const text = normalizeRange(range);
  e.structural(() => {
    nameEl(e, name).textContent = text;
    e.wb.pkg.touch(e.wb.pkg.workbookPath);
  });
}

export function renameName(e: Edits, from: string, to: string): void {
  if (from === to) return;
  validateName(e.wb, to, from);
  e.structural(() => {
    nameEl(e, from).setAttribute("name", to);
    rewriteFormulas(e, (text) => renameNameInFormula(text, from, to));
  });
}

export function deleteName(e: Edits, name: string): void {
  e.structural(() => {
    const el = nameEl(e, name);
    const parent = el.parentElement!;
    el.remove();
    if (!parent.children.length) parent.remove();
    e.wb.pkg.touch(e.wb.pkg.workbookPath);
  });
}
