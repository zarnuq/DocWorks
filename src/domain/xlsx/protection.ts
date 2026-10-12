/** Excel's sheet protection flags describe operations that are forbidden (true), rather than allowed. */
import type { Edits } from "./commands";
import { child } from "./sheet-edit";
import type { Sheet } from "./workbook-model";

export type ProtectedAction = "formatCells" | "formatColumns" | "formatRows" | "insertColumns" | "insertRows" | "deleteColumns" | "deleteRows" | "sort" | "autoFilter" | "objects" | "sheet";
const truth = (value: string | null | undefined, fallback: boolean) => value == null ? fallback : value === "1" || value === "true";
export function isProtected(sheet: Sheet): boolean {
  return truth(sheet.first(sheet.doc.documentElement, "sheetProtection")?.getAttribute("sheet"), false);
}
export function assertAction(sheet: Sheet, action: ProtectedAction): void {
  const p = sheet.first(sheet.doc.documentElement, "sheetProtection");
  if (isProtected(sheet) && truth(p?.getAttribute(action), true)) throw new Error("This sheet is protected. This operation is not allowed.");
}
export function assertCell(sheet: Sheet, row: number, col: number): void {
  if (!isProtected(sheet)) return;
  const xf = sheet.wb.styles.xf(sheet.styleAt(row, col)).el;
  const protection = sheet.first(xf, "protection");
  // A cell without an explicit protection setting inherits its named style; the default is locked.
  const base = sheet.wb.styles.doc.getElementsByTagNameNS(xf.namespaceURI, "cellStyleXfs")[0]?.children[Number(xf.getAttribute("xfId") ?? 0)];
  const inherited = sheet.first(base, "protection");
  const locked = xf.getAttribute("applyProtection") === "0" ? inherited?.getAttribute("locked") : protection?.getAttribute("locked") ?? inherited?.getAttribute("locked");
  if (truth(locked, true)) throw new Error("This cell is protected. Unlock it before editing.");
}

const PASSWORD_ATTRS = ["password", "hashValue", "algorithmName"];

/** Password-bearing protection is only removed by `unprotectWithPassword`. */
export function setProtected(e: Edits, sheet: Sheet, on: boolean): void {
  const existing = sheet.first(sheet.doc.documentElement, "sheetProtection");
  if (!on && existing && PASSWORD_ATTRS.some((a) => existing.hasAttribute(a))) {
    throw new Error("This sheet has password protection. Enter its password to unprotect it.");
  }
  apply(e, sheet, on);
}

export function hasPassword(sheet: Sheet): boolean {
  const p = sheet.first(sheet.doc.documentElement, "sheetProtection");
  return !!p && PASSWORD_ATTRS.some((a) => p.hasAttribute(a));
}

function apply(e: Edits, sheet: Sheet, on: boolean): void {
  const path = sheet.ref.path;
  e.structural(() => {
    const s = e.wb.sheets.find((s) => s.ref.path === path)!;
    if (on) child(s, "sheetProtection", true)!.setAttribute("sheet", "1");
    else child(s, "sheetProtection")?.remove();
    s.touch();
  });
}

/** Excel's legacy 16-bit password hash (the `password` attribute). */
function legacyHash(pw: string): string {
  let h = 0;
  const rot = () => (h = ((h >> 14) & 1) | ((h << 1) & 0x7fff));
  for (let i = pw.length - 1; i >= 0; i--) {
    rot();
    h ^= pw.charCodeAt(i);
  }
  rot();
  h ^= pw.length ^ 0xce4b;
  return h.toString(16).toUpperCase().padStart(4, "0");
}

const ALGORITHMS: Record<string, string> = { "SHA-512": "SHA-512", "SHA-384": "SHA-384", "SHA-256": "SHA-256", "SHA-1": "SHA-1" };
const bytesOf = (b64: string) => Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
const join = (a: Uint8Array, b: Uint8Array) => {
  const out = new Uint8Array(a.length + b.length);
  out.set(a);
  out.set(b, a.length);
  return out;
};

/** Whether `pw` matches the sheet's protection password (legacy or salted, spun hash as Excel writes it). */
export async function checkPassword(sheet: Sheet, pw: string): Promise<boolean> {
  const p = sheet.first(sheet.doc.documentElement, "sheetProtection");
  if (!p) return true;
  const legacy = p.getAttribute("password");
  if (legacy) return legacyHash(pw) === legacy.toUpperCase().padStart(4, "0");
  const algo = ALGORITHMS[p.getAttribute("algorithmName") ?? ""];
  const want = p.getAttribute("hashValue");
  if (!algo || !want) return false;
  const utf16 = new Uint8Array(pw.length * 2);
  for (let i = 0; i < pw.length; i++) {
    utf16[i * 2] = pw.charCodeAt(i) & 0xff;
    utf16[i * 2 + 1] = pw.charCodeAt(i) >> 8;
  }
  let h = new Uint8Array(await crypto.subtle.digest(algo, join(bytesOf(p.getAttribute("saltValue") ?? ""), utf16)));
  const spins = Number(p.getAttribute("spinCount") ?? 0);
  const n = new Uint8Array(4);
  for (let i = 0; i < spins; i++) {
    new DataView(n.buffer).setUint32(0, i, true);
    h = new Uint8Array(await crypto.subtle.digest(algo, join(h, n)));
  }
  return btoa(String.fromCharCode(...h)) === want;
}

/** Remove password protection when `pw` is right; false (nothing changed) when it isn't. */
export async function unprotectWithPassword(e: Edits, sheet: Sheet, pw: string): Promise<boolean> {
  if (!(await checkPassword(sheet, pw))) return false;
  apply(e, sheet, false);
  return true;
}
