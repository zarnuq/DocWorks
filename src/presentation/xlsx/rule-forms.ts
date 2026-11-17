/*
 * The conditional formatting and data validation dialogs' contents (Google Sheets' sidebars), as plain DOM so
 * the view can put them in a modal and tests can drive them under jsdom.
 */
import { CfRule, NewRule, NewRuleType } from "../../domain/xlsx/conditional";
import { NewValidation, Validation } from "../../domain/xlsx/validation";
import { parseRange, Rect, rangeName } from "../../domain/xlsx/references";
import type { PivotFn, PivotLayout } from "../../domain/xlsx/pivot";

/** "A1:B2 D4" or "A1:B2, D4" as ranges; null when any part isn't a range. */
function parseRanges(text: string): Rect[] | null {
  const parts = text.replace(/\$/g, "").split(/[\s,]+/).filter(Boolean);
  const out = parts.map((p) => parseRange(p.toUpperCase()));
  return out.length && out.every(Boolean) ? (out as Rect[]) : null;
}

const CONDITIONS: [string, string][] = [
  ["Is empty", "containsBlanks"],
  ["Is not empty", "notContainsBlanks"],
  ["Text contains", "containsText"],
  ["Text does not contain", "notContainsText"],
  ["Text starts with", "beginsWith"],
  ["Text ends with", "endsWith"],
  ["Text is exactly", "cellIs:equal:text"],
  ["Greater than", "cellIs:greaterThan"],
  ["Greater than or equal to", "cellIs:greaterThanOrEqual"],
  ["Less than", "cellIs:lessThan"],
  ["Less than or equal to", "cellIs:lessThanOrEqual"],
  ["Is equal to", "cellIs:equal"],
  ["Is not equal to", "cellIs:notEqual"],
  ["Is between", "cellIs:between"],
  ["Is not between", "cellIs:notBetween"],
  ["Custom formula is", "expression"],
  ["Color scale", "colorScale"],
];

const OPERATORS: Record<string, string> = {
  greaterThan: "Greater than", greaterThanOrEqual: "Greater than or equal to", lessThan: "Less than", lessThanOrEqual: "Less than or equal to",
  equal: "Is equal to", notEqual: "Is not equal to", between: "Is between", notBetween: "Is not between",
};

/** A rule in words, as the sidebar lists it. */
function describeRule(r: CfRule): string {
  const f = r.formulas.map((x) => x.replace(/^"(.*)"$/, "$1"));
  switch (r.type) {
    case "cellIs":
      return `${OPERATORS[r.operator ?? "equal"] ?? r.operator} ${r.operator === "between" || r.operator === "notBetween" ? `${f[0]} and ${f[1]}` : f[0]}`;
    case "expression":
      return `Custom formula is =${r.formulas[0] ?? ""}`;
    case "containsText":
      return `Text contains "${r.text ?? ""}"`;
    case "notContainsText":
      return `Text does not contain "${r.text ?? ""}"`;
    case "beginsWith":
      return `Text starts with "${r.text ?? ""}"`;
    case "endsWith":
      return `Text ends with "${r.text ?? ""}"`;
    case "containsBlanks":
      return "Is empty";
    case "notContainsBlanks":
      return "Is not empty";
    case "colorScale":
      return "Color scale";
    case "dataBar":
      return "Data bar";
    case "top10":
      return `${r.el.getAttribute("bottom") === "1" ? "Bottom" : "Top"} ${r.el.getAttribute("rank") ?? 10}${r.el.getAttribute("percent") === "1" ? "%" : ""}`;
    case "duplicateValues":
      return "Duplicate values";
    case "uniqueValues":
      return "Unique values";
    default:
      return r.type;
  }
}

function row(host: HTMLElement, label: string): HTMLElement {
  const el = host.createDiv({ cls: "xlsx-rule-row" });
  el.createSpan({ text: label });
  return el;
}

interface ConditionalFormOptions {
  rules: CfRule[];
  range: string;
  /** Swatch colours of a rule's format. */
  swatch(rule: CfRule): { fill?: string; color?: string };
  add(ranges: Rect[], rule: NewRule): void;
  remove(rule: CfRule): void;
}

export function conditionalForm(host: HTMLElement, o: ConditionalFormOptions): void {
  host.addClass("xlsx-rules");
  const list = host.createDiv({ cls: "xlsx-rule-list" });
  if (!o.rules.length) list.createDiv({ cls: "xlsx-rule-empty", text: "No rules on this sheet." });
  for (const r of o.rules) {
    const item = list.createDiv({ cls: "xlsx-rule-item" });
    const sw = o.swatch(r);
    const swatch = item.createSpan({ cls: "xlsx-rule-swatch", text: "123" });
    swatch.setCssProps({ "--xlsx-swatch-fill": sw.fill ?? "transparent", "--xlsx-swatch-color": sw.color ?? "inherit" });
    const text = item.createDiv({ cls: "xlsx-rule-text" });
    text.createDiv({ text: describeRule(r) });
    text.createDiv({ cls: "xlsx-rule-range", text: r.ranges.map(rangeName).join(", ") });
    item.createEl("button", { text: "Delete", attr: { "aria-label": `Delete rule: ${describeRule(r)}` } }).addEventListener("click", () => o.remove(r));
  }

  host.createDiv({ cls: "xlsx-rule-heading", text: "Add a rule" });
  const range = row(host, "Apply to range").createEl("input", { type: "text", value: o.range, attr: { "aria-label": "Apply to range" } });
  const condition = row(host, "Format cells if").createEl("select", { cls: "dropdown", attr: { "aria-label": "Condition" } });
  for (const [label, value] of CONDITIONS) condition.createEl("option", { text: label, value });
  condition.value = "notContainsBlanks";
  const valueRow = row(host, "Value");
  const value1 = valueRow.createEl("input", { type: "text", attr: { "aria-label": "Value or formula" } });
  const value2 = valueRow.createEl("input", { type: "text", attr: { "aria-label": "Second value" } });

  const styleRow = row(host, "Formatting style");
  const pick = (label: string, initial: string, on: boolean) => {
    const wrap = styleRow.createEl("label", { cls: "xlsx-rule-pick" });
    const box = wrap.createEl("input", { type: "checkbox", attr: { "aria-label": `${label} on` } });
    box.checked = on;
    wrap.appendText(label);
    const input = wrap.createEl("input", { type: "color", value: initial, attr: { "aria-label": label } });
    input.addEventListener("input", () => (box.checked = true));
    return { box, input };
  };
  const fill = pick("Fill", "#b7e1cd", true);
  const text = pick("Text", "#000000", false);
  const flags = (["bold", "italic", "underline", "strike"] as const).map((k) => {
    const wrap = styleRow.createEl("label", { cls: "xlsx-rule-pick" });
    const box = wrap.createEl("input", { type: "checkbox", attr: { "aria-label": k[0].toUpperCase() + k.slice(1) } });
    wrap.appendText(k === "strike" ? "S̶" : k === "bold" ? "B" : k === "italic" ? "I" : "U");
    return [k, box] as const;
  });

  const scaleRow = row(host, "Color scale");
  const scale = (label: string, initial: string) => scaleRow.createEl("input", { type: "color", value: initial, attr: { "aria-label": label } });
  const min = scale("Minpoint color", "#ffffff");
  const midWrap = scaleRow.createEl("label", { cls: "xlsx-rule-pick" });
  const useMid = midWrap.createEl("input", { type: "checkbox", attr: { "aria-label": "Use a midpoint" } });
  midWrap.appendText("Midpoint");
  const mid = scale("Midpoint color", "#ffd666");
  const max = scale("Maxpoint color", "#57bb8a");

  const error = host.createDiv({ cls: "xlsx-rule-error" });
  const sync = () => {
    const v = condition.value;
    const needs = v === "containsBlanks" || v === "notContainsBlanks" || v === "colorScale" ? 0 : v.endsWith("between") || v.endsWith("Between") ? 2 : 1;
    valueRow.toggle(needs > 0);
    value2.toggle(needs === 2);
    value1.placeholder = v === "expression" ? "=A1>10" : "Value or formula";
    styleRow.toggle(v !== "colorScale");
    scaleRow.toggle(v === "colorScale");
  };
  condition.addEventListener("change", sync);
  sync();

  host.createEl("button", { text: "Add rule", cls: "mod-cta" }).addEventListener("click", () => {
    const ranges = parseRanges(range.value);
    if (!ranges) return void error.setText("Enter a range like A1:B10.");
    const [type, operator, kind] = condition.value.split(":");
    if (type === "colorScale") {
      o.add(ranges, { type: "colorScale", colors: [min.value, ...(useMid.checked ? [mid.value] : []), max.value].map((c) => c.toUpperCase()) });
      return;
    }
    const values = [value1.value, value2.value];
    if (type !== "containsBlanks" && type !== "notContainsBlanks" && !values[0].trim()) return void error.setText("Enter a value.");
    // "Text is exactly" compares as text even when it looks like a number.
    if (kind === "text") values[0] = `="${values[0].replace(/"/g, '""')}"`;
    const style: NewRule["style"] = {};
    if (fill.box.checked) style.fill = fill.input.value.toUpperCase();
    if (text.box.checked) style.color = text.input.value.toUpperCase();
    for (const [k, box] of flags) if (box.checked) style[k] = true;
    if (!Object.keys(style).length) return void error.setText("Choose a formatting style.");
    o.add(ranges, { type: type as NewRuleType, operator, values: type === "expression" && !values[0].startsWith("=") ? [`=${values[0]}`] : values, style });
  });
}

interface ValidationFormOptions {
  current: Validation | null;
  range: string;
  add(ranges: Rect[], rule: NewValidation): void;
  remove(rule: Validation): void;
}

const CRITERIA: [string, string][] = [
  ["Dropdown", "list"],
  ["Dropdown (from a range)", "listRange"],
  ["Number is between", "decimal:between"],
  ["Number is not between", "decimal:notBetween"],
  ["Number is greater than", "decimal:greaterThan"],
  ["Number is greater than or equal to", "decimal:greaterThanOrEqual"],
  ["Number is less than", "decimal:lessThan"],
  ["Number is less than or equal to", "decimal:lessThanOrEqual"],
  ["Number is equal to", "decimal:equal"],
  ["Custom formula is", "custom"],
];

export function validationForm(host: HTMLElement, o: ValidationFormOptions): void {
  host.addClass("xlsx-rules");
  const range = row(host, "Apply to range").createEl("input", { type: "text", value: o.range, attr: { "aria-label": "Apply to range" } });
  const criteria = row(host, "Criteria").createEl("select", { cls: "dropdown", attr: { "aria-label": "Criteria" } });
  for (const [label, value] of CRITERIA) criteria.createEl("option", { text: label, value });
  const itemsRow = row(host, "Items");
  const items = itemsRow.createEl("input", { type: "text", attr: { "aria-label": "Items", placeholder: "Yes, No, Maybe" } });
  const valueRow = row(host, "Value");
  const value1 = valueRow.createEl("input", { type: "text", attr: { "aria-label": "Value" } });
  const value2 = valueRow.createEl("input", { type: "text", attr: { "aria-label": "Second value" } });
  const rejectRow = row(host, "If the data is invalid");
  const reject = rejectRow.createEl("select", { cls: "dropdown", attr: { "aria-label": "If the data is invalid" } });
  reject.createEl("option", { text: "Show a warning", value: "warning" });
  reject.createEl("option", { text: "Reject the input", value: "stop" });

  const cur = o.current;
  if (cur) {
    if (cur.type === "list") {
      criteria.value = cur.items ? "list" : "listRange";
      items.value = cur.items ? cur.items.join(", ") : cur.formulas[0] ?? "";
    } else if (cur.type === "custom") {
      criteria.value = "custom";
      value1.value = `=${cur.formulas[0] ?? ""}`;
    } else if (cur.type === "decimal" || cur.type === "whole") {
      criteria.value = `decimal:${cur.operator}`;
      [value1.value, value2.value] = [cur.formulas[0] ?? "", cur.formulas[1] ?? ""];
    }
    reject.value = cur.reject ? "stop" : "warning";
  }
  const sync = () => {
    const v = criteria.value;
    itemsRow.toggle(v === "list" || v === "listRange");
    items.placeholder = v === "listRange" ? "Sheet1!A1:A10" : "Yes, No, Maybe";
    valueRow.toggle(v.startsWith("decimal") || v === "custom");
    value2.toggle(v.endsWith("etween"));
  };
  criteria.addEventListener("change", sync);
  sync();

  const error = host.createDiv({ cls: "xlsx-rule-error" });
  const buttons = host.createDiv({ cls: "xlsx-rule-buttons" });
  if (cur) buttons.createEl("button", { text: "Remove rule" }).addEventListener("click", () => o.remove(cur));
  buttons.createEl("button", { text: "Done", cls: "mod-cta" }).addEventListener("click", () => {
    const ranges = parseRanges(range.value);
    if (!ranges) return void error.setText("Enter a range like A1:B10.");
    const v = criteria.value;
    const stop = reject.value === "stop";
    if (v === "list") {
      const list = items.value.split(",").map((s) => s.trim()).filter(Boolean);
      if (!list.length) return void error.setText("Enter at least one item.");
      return o.add(ranges, { type: "list", items: list, reject: stop });
    }
    if (v === "listRange") {
      if (!items.value.trim()) return void error.setText("Enter a range for the items.");
      return o.add(ranges, { type: "list", source: items.value.trim().replace(/^=/, ""), reject: stop });
    }
    if (v === "custom") {
      if (!value1.value.trim()) return void error.setText("Enter a formula.");
      return o.add(ranges, { type: "custom", formulas: [value1.value.trim().replace(/^=/, "")], reject: stop });
    }
    const operator = v.split(":")[1];
    const values = (operator.endsWith("etween") ? [value1.value, value2.value] : [value1.value]).map((x) => x.trim());
    if (values.some((x) => !Number.isFinite(Number(x)) || x === "")) return void error.setText("Enter numbers.");
    o.add(ranges, { type: "decimal", operator, formulas: values, reject: stop });
  });
}

interface PivotFormOptions {
  fields: string[];
  layout: PivotLayout;
  /** A field's items, for its filter's checklist (left out: filters can't be edited). */
  items?(field: number): (string | number | boolean)[];
  apply(layout: PivotLayout): void;
}

const PIVOT_FNS: [string, PivotFn][] = [["SUM", "sum"], ["COUNTA", "count"], ["AVERAGE", "average"], ["MAX", "max"], ["MIN", "min"]];

/** Google's pivot table editor: row fields, column fields, values (field and summary) and filters (a field and the items it shows). */
export function pivotForm(host: HTMLElement, o: PivotFormOptions): void {
  host.addClass("xlsx-rules");
  const rows = new Set(o.layout.rows), cols = new Set(o.layout.cols);
  const values = o.layout.values.map((v) => ({ field: v.field, fn: v.fn }));
  // A field is on one axis: ticking it on one unticks it on the other.
  const axis = (title: string, own: Set<number>, other: Set<number>, others: () => HTMLInputElement[]) => {
    const box = row(host, title);
    return o.fields.map((name, i) => {
      const label = box.createEl("label", { cls: "xlsx-pivot-field" });
      const input = label.createEl("input", { type: "checkbox", attr: { "aria-label": `${title}: ${name}` } });
      input.checked = own.has(i);
      label.appendText(name);
      input.addEventListener("change", () => {
        if (!input.checked) return void own.delete(i);
        own.add(i);
        if (other.delete(i)) others()[i].checked = false;
      });
      return input;
    });
  };
  const rowBoxes: HTMLInputElement[] = axis("Rows", rows, cols, () => colBoxes);
  const colBoxes: HTMLInputElement[] = axis("Columns", cols, rows, () => rowBoxes);
  const valueBox = row(host, "Values").createDiv({ cls: "xlsx-pivot-values" });
  const drawValues = () => {
    valueBox.empty();
    values.forEach((v, k) => {
      const line = valueBox.createDiv({ cls: "xlsx-pivot-value" });
      const field = line.createEl("select", { cls: "dropdown", attr: { "aria-label": "Value field" } });
      o.fields.forEach((name, i) => field.createEl("option", { text: name, value: String(i) }));
      field.value = String(v.field);
      field.addEventListener("change", () => (v.field = Number(field.value)));
      const fn = line.createEl("select", { cls: "dropdown", attr: { "aria-label": "Summarize by" } });
      for (const [label, value] of PIVOT_FNS) fn.createEl("option", { text: label, value });
      fn.value = v.fn;
      fn.addEventListener("change", () => (v.fn = fn.value as PivotFn));
      line.createEl("button", { text: "Remove", attr: { "aria-label": "Remove value" } }).addEventListener("click", () => {
        values.splice(k, 1);
        drawValues();
      });
    });
    valueBox.createEl("button", { text: "Add value", cls: "xlsx-pivot-add" }).addEventListener("click", () => {
      // The first field not already in use; COUNTA works whatever the column holds (change it to SUM for numbers).
      const used = new Set([...rows, ...cols, ...values.map((v) => v.field)]);
      const free = o.fields.findIndex((_, i) => !used.has(i));
      values.push({ field: free < 0 ? 0 : free, fn: "count" });
      drawValues();
    });
  };
  drawValues();
  // Filters: one per field, showing the items left ticked.
  const filters = (o.layout.filters ?? []).map((f) => ({ field: f.field, hidden: new Set(f.hidden) }));
  // A source that can't be found (a named table, a renamed sheet) leaves the checklist empty; Done reports it.
  const itemsOf = (field: number) => {
    try {
      return o.items?.(field) ?? [];
    } catch {
      return [];
    }
  };
  const filterBox = row(host, "Filters").createDiv({ cls: "xlsx-pivot-values" });
  const drawFilters = (open = -1) => {
    filterBox.empty();
    filters.forEach((f, k) => {
      const line = filterBox.createDiv({ cls: "xlsx-pivot-value" });
      const field = line.createEl("select", { cls: "dropdown", attr: { "aria-label": "Filter field" } });
      o.fields.forEach((name, i) => (i === f.field || !filters.some((x) => x.field === i)) && field.createEl("option", { text: name, value: String(i) }));
      field.value = String(f.field);
      field.addEventListener("change", () => {
        f.field = Number(field.value);
        f.hidden.clear();
        drawFilters();
      });
      const items = itemsOf(f.field);
      // What the filter shows, in words (the items may be unreadable: then only whether any are hidden is known).
      const shows = () => (!f.hidden.size ? "Showing all items" : items.length ? `Showing ${items.filter((v) => !f.hidden.has(v)).length} of ${items.length} items` : "Some items hidden");
      const status = line.createEl("button", { cls: "xlsx-pivot-filter-status", text: shows(), attr: { "aria-expanded": String(open === k) } });
      status.addEventListener("click", () => drawFilters(open === k ? -1 : k));
      line.createEl("button", { text: "Remove", attr: { "aria-label": "Remove filter" } }).addEventListener("click", () => {
        filters.splice(k, 1);
        drawFilters();
      });
      if (open !== k) return;
      // ponytail: a condition only ticks the items there are now; it isn't kept, so new values in the source show
      // until the filter is edited again. Storing it means OOXML's pivot <filters>.
      const cond = filterBox.createDiv({ cls: "xlsx-pivot-value" });
      const kind = cond.createEl("select", { cls: "dropdown", attr: { "aria-label": "Condition" } });
      for (const [text, value] of [["Text contains", "contains"], ["Text does not contain", "lacks"], ["Is equal to", "equal"], ["Greater than", "greater"], ["Less than", "less"]]) kind.createEl("option", { text, value });
      const want = cond.createEl("input", { type: "text", attr: { "aria-label": "Condition value" } });
      cond.createEl("button", { text: "Show matching", cls: "xlsx-pivot-filter-match" }).addEventListener("click", () => {
        const text = want.value.trim().toLowerCase(), n = Number(want.value);
        const match = (v: string | number | boolean) =>
          kind.value === "contains" ? String(v).toLowerCase().includes(text)
          : kind.value === "lacks" ? !String(v).toLowerCase().includes(text)
          : kind.value === "equal" ? String(v).toLowerCase() === text
          : typeof v === "number" && Number.isFinite(n) && (kind.value === "greater" ? v > n : v < n);
        f.hidden = new Set(items.filter((v) => !match(v)));
        drawFilters(k);
      });
      const list = filterBox.createDiv({ cls: "xlsx-pivot-filter-items" });
      for (const v of items) {
        const label = list.createEl("label", { cls: "xlsx-pivot-field" });
        const box = label.createEl("input", { type: "checkbox" });
        box.checked = !f.hidden.has(v);
        label.appendText(String(v));
        box.addEventListener("change", () => {
          if (box.checked) f.hidden.delete(v);
          else f.hidden.add(v);
          status.setText(shows());
        });
      }
    });
    const free = o.fields.findIndex((_, i) => !filters.some((f) => f.field === i));
    if (o.items && free >= 0) filterBox.createEl("button", { text: "Add filter", cls: "xlsx-pivot-add-filter" }).addEventListener("click", () => {
      filters.push({ field: free, hidden: new Set() });
      drawFilters();
    });
  };
  drawFilters();
  const error = host.createDiv({ cls: "xlsx-rule-error" });
  host.createDiv({ cls: "xlsx-rule-buttons" }).createEl("button", { text: "Done", cls: "mod-cta" }).addEventListener("click", () => {
    if (!values.length) return void error.setText("Add at least one value.");
    // Fields nest in the order they had, newly ticked ones last.
    o.apply({ rows: [...rows], cols: [...cols], values: values.map((v) => ({ ...v })), filters: filters.map((f) => ({ field: f.field, hidden: [...f.hidden] })) });
  });
}
