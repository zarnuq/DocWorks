/*
 * The conditional formatting and data validation dialogs' contents (Google Sheets' sidebars), as plain DOM so
 * the view can put them in a modal and tests can drive them under jsdom.
 */
import { CfRule, NewRule, NewRuleType } from "../../domain/xlsx/conditional";
import { NewValidation, Validation } from "../../domain/xlsx/validation";
import { parseRange, Rect, rangeName } from "../../domain/xlsx/references";
import type { PivotFn, PivotLayout } from "../../domain/xlsx/pivot";

/** "A1:B2 D4" or "A1:B2, D4" as ranges; null when any part isn't a range. */
export function parseRanges(text: string): Rect[] | null {
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
export function describeRule(r: CfRule): string {
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

export interface ConditionalFormOptions {
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

export interface ValidationFormOptions {
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

export interface PivotFormOptions {
  fields: string[];
  layout: PivotLayout;
  apply(layout: PivotLayout): void;
}

const PIVOT_FNS: [string, PivotFn][] = [["SUM", "sum"], ["COUNTA", "count"], ["AVERAGE", "average"], ["MAX", "max"], ["MIN", "min"]];

/** Google's pivot table editor: row fields, a column field and values (field and summary). */
export function pivotForm(host: HTMLElement, o: PivotFormOptions): void {
  host.addClass("xlsx-rules");
  const rows = new Set(o.layout.rows);
  let col: number | null = o.layout.cols[0] ?? null;
  const values = o.layout.values.map((v) => ({ field: v.field, fn: v.fn }));
  const rowBox = row(host, "Rows");
  const boxes = o.fields.map((name, i) => {
    const label = rowBox.createEl("label", { cls: "xlsx-pivot-field" });
    const box = label.createEl("input", { type: "checkbox" });
    box.checked = rows.has(i);
    label.appendText(name);
    box.addEventListener("change", () => {
      if (box.checked) {
        rows.add(i);
        if (col === i) {
          colSelect.value = "";
          col = null;
        }
      } else rows.delete(i);
    });
    return box;
  });
  const colSelect = row(host, "Columns").createEl("select", { cls: "dropdown", attr: { "aria-label": "Column field" } });
  colSelect.createEl("option", { text: "None", value: "" });
  o.fields.forEach((name, i) => colSelect.createEl("option", { text: name, value: String(i) }));
  colSelect.value = col === null ? "" : String(col);
  colSelect.addEventListener("change", () => {
    col = colSelect.value === "" ? null : Number(colSelect.value);
    if (col !== null && rows.delete(col)) boxes[col].checked = false;
  });
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
      const used = new Set([...rows, ...(col === null ? [] : [col]), ...values.map((v) => v.field)]);
      const free = o.fields.findIndex((_, i) => !used.has(i));
      values.push({ field: free < 0 ? 0 : free, fn: "count" });
      drawValues();
    });
  };
  drawValues();
  const error = host.createDiv({ cls: "xlsx-rule-error" });
  host.createDiv({ cls: "xlsx-rule-buttons" }).createEl("button", { text: "Done", cls: "mod-cta" }).addEventListener("click", () => {
    if (!values.length) return void error.setText("Add at least one value.");
    o.apply({ rows: o.fields.map((_, i) => i).filter((i) => rows.has(i)), cols: col === null ? [] : [col], values: values.map((v) => ({ ...v })) });
  });
}
