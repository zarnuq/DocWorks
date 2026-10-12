/** Resolve structured references against worksheet table relationships without rewriting the formula text. */
import type { Node } from "./formula";
import { parseRange } from "./references";
import type { Sheet, Workbook } from "./workbook-model";

export function tableReference(wb: Workbook, sheet: Sheet, row: number, col: number, text: string): Node {
  const at = text.indexOf("[");
  const name = at < 0 ? text : text.slice(0, at);
  const spec = at < 0 ? "[#Data]" : text.slice(at);
  const bad = (value: string): Node => ({ type: "err", value });
  for (const owner of wb.sheets) for (const rel of wb.pkg.rels(owner.ref.path)) {
    if (rel.type !== "table" || rel.external) continue;
    const table = wb.pkg.xml(rel.target)?.documentElement;
    const area = parseRange(table?.getAttribute("ref") ?? "");
    if (!table || !area) continue;
    if (name ? ![table.getAttribute("name"), table.getAttribute("displayName")].some((n) => n?.toLowerCase() === name.toLowerCase()) : owner !== sheet || row < area.r1 || row > area.r2 || col < area.c1 || col > area.c2) continue;
    // Leaf selectors (a quote escapes special characters in column names).
    const leaves = [...spec.matchAll(/\[((?:'[^]|[^\][])*)\]/g)].map((m) => m[1]);
    if (!leaves.length) return { type: "unsupported", text };
    const selectors = leaves.filter((s) => s.startsWith("#")).map((s) => s.toLowerCase());
    if (selectors.some((s) => !["#all", "#data", "#headers", "#totals", "#this row"].includes(s))) return { type: "unsupported", text };
    const thisRow = selectors.includes("#this row") || leaves.some((s) => s.startsWith("@")) || spec.startsWith("[@[");
    const names = leaves.filter((s) => !s.startsWith("#")).map((s) => s.replace(/^@/, "").replace(/'([\][#'@])/g, "$1")).filter(Boolean);
    const columns = Array.from(table.getElementsByTagNameNS(table.namespaceURI, "tableColumn")).map((c) => c.getAttribute("name")?.toLowerCase());
    const indices = names.map((n) => columns.indexOf(n.toLowerCase()));
    if (indices.some((n) => n < 0)) return bad("#REF!");
    // Comma-separated disjoint columns / row regions need a union, which is not a rectangle.
    if (indices.length > 1 && !spec.includes("]:[")) return { type: "unsupported", text };
    const c1 = indices.length ? area.c1 + Math.min(...indices) : area.c1;
    const c2 = indices.length ? area.c1 + Math.max(...indices) : area.c2;
    const headers = Number(table.getAttribute("headerRowCount") ?? 1);
    const totals = Number(table.getAttribute("totalsRowCount") ?? 0);
    let r1 = area.r1 + headers, r2 = area.r2 - totals;
    if (thisRow) {
      if (row < r1 || row > r2) return bad("#VALUE!");
      r1 = r2 = row;
    } else if (selectors.includes("#all")) { r1 = area.r1; r2 = area.r2; }
    else {
      if (selectors.includes("#headers")) {
        if (!headers) return bad("#REF!");
        r1 = area.r1;
        if (!selectors.includes("#data")) r2 = area.r1 + headers - 1;
      }
      if (selectors.includes("#totals")) {
        if (!totals) return bad("#REF!");
        if (selectors.includes("#headers") && !selectors.includes("#data")) return { type: "unsupported", text };
        r2 = area.r2;
        if (!selectors.includes("#data")) r1 = area.r2 - totals + 1;
      }
    }
    if (r1 > r2) return bad("#REF!");
    return { type: "ref", sheet: owner.name, range: { r1, c1, r2, c2 }, area: true };
  }
  return bad("#NAME?");
}
