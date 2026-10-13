/*
 * Named ranges, as Google Sheets' Data > Named ranges: list, add, change the range, rename (formulas follow) and
 * delete. Hidden and built-in names (_xlnm.*) are left alone.
 */
import { NS_S } from "../../infrastructure/xlsx/xlsx-workbook";
import { Edits } from "./commands";
import { renameNameInFormula, tokenize } from "./formula";
import { quoteSheet } from "./references";
import { rewriteFormulas } from "./sheet-edit";
import { Workbook } from "./workbook-model";

export interface NamedRange {
  name: string;
  range: string;
}

const userNames = (wb: Workbook) => wb.names().filter((n) => !n.name.startsWith("_xlnm.") && n.el.getAttribute("hidden") !== "1");

export function listNames(wb: Workbook): NamedRange[] {
  return userNames(wb).map((n) => ({ name: n.name, range: n.formula }));
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
  const text = normalizeRange(range);
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
