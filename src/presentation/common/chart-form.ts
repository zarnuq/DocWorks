/** The chart editor's settings (Google's chart editor "Setup" type and "Customize" options), as plain DOM. */
import { App, Modal } from "obsidian";
import type { ChartData } from "../../domain/pptx/chart";
import type { ChartFormat, ChartKind, LegendPos, SeriesKind, Stacking } from "../../domain/pptx/chart-format";

export function defaultChartFormat(): ChartFormat {
  const axis = { title: "", visible: true, gridlines: false, min: null, max: null, reversed: false, format: null };
  return { kind: "column", stacking: "none", title: "", legend: "b", labels: false, colors: [], category: { ...axis }, value: { ...axis, gridlines: true } };
}

/** Insert › Chart's types, as Google's submenu lists them. */
export const NEW_CHART_TYPES: [label: string, kind: ChartKind][] = [["Bar", "bar"], ["Column", "column"], ["Line", "line"], ["Pie", "pie"]];

/** A new chart's format of `kind` (a bar chart lists its first category at the top, as Google's). */
export function newChartFormat(kind: ChartKind): ChartFormat {
  const f = defaultChartFormat();
  f.kind = kind;
  f.category.reversed = kind === "bar";
  return f;
}

/** Controls that edit `format` in place; `data` names the series (or slices) for the colour pickers. */
export function chartFormatForm(host: HTMLElement, format: ChartFormat, data: ChartData, palette: string[]): { renderColors: () => void } {
  const field = (label: string) => {
    const row = host.createDiv({ cls: "dw-chart-option" });
    row.createSpan({ text: label });
    return row;
  };
  // Chart type, as Google's chart editor lists them (stacking is part of the type).
  const types: [string, ChartKind, Stacking][] = [["Column", "column", "none"], ["Stacked column", "column", "stacked"], ["100% stacked column", "column", "percent"], ["Bar", "bar", "none"], ["Stacked bar", "bar", "stacked"], ["100% stacked bar", "bar", "percent"], ["Line", "line", "none"], ["Area", "area", "none"], ["Stacked area", "area", "stacked"], ["100% stacked area", "area", "percent"], ["Pie", "pie", "none"], ["Doughnut", "doughnut", "none"], ["Combo (columns and lines)", "combo", "none"], ["Scatter", "scatter", "none"]];
  const picker = field("Chart type").createEl("select", { cls: "dropdown", attr: { "aria-label": "Chart type" } });
  for (const [text, kind, stacking] of types) picker.createEl("option", { text, value: `${kind}:${stacking}` });
  if (!types.some(([, k, st]) => k === format.kind && st === format.stacking)) picker.createEl("option", { text: "Other (kept as is)", value: `${format.kind}:${format.stacking}` });
  picker.value = `${format.kind}:${format.stacking}`;
  picker.addEventListener("change", () => {
    const [kind, stacking] = picker.value.split(":") as [ChartKind, Stacking];
    // Google lists a bar chart's first category at the top.
    if ((kind === "bar") !== (format.kind === "bar")) format.category.reversed = kind === "bar";
    format.kind = kind;
    format.stacking = stacking;
    format.colors = [];
    // A new combo starts as Google's: the last series a line, the rest columns.
    if (kind === "combo" && !format.seriesKinds) format.seriesKinds = data.series.map((_, i, all) => (i === all.length - 1 && all.length > 1 ? "line" : "column"));
    renderColors();
  });
  const text = (label: string, value: string, set: (v: string) => void) => {
    const input = field(label).createEl("input", { type: "text", value, attr: { "aria-label": label } });
    input.addEventListener("input", () => set(input.value));
    return input;
  };
  text("Chart title", format.title, (v) => (format.title = v));
  const legend = field("Legend").createEl("select", { cls: "dropdown", attr: { "aria-label": "Legend" } });
  for (const [label, value] of [["Bottom", "b"], ["Top", "t"], ["Left", "l"], ["Right", "r"], ["None", "none"]]) legend.createEl("option", { text: label, value });
  legend.value = format.legend;
  legend.addEventListener("change", () => (format.legend = legend.value as LegendPos));
  const check = (label: string, value: boolean, set: (v: boolean) => void) => {
    const box = field(label).createEl("input", { type: "checkbox", attr: { "aria-label": label } });
    box.checked = value;
    box.addEventListener("change", () => set(box.checked));
    return box;
  };
  check("Data labels", format.labels, (v) => (format.labels = v));
  // Axis settings, hidden for pies (they have no axes), as in Google's chart editor.
  const axisRows: HTMLElement[] = [];
  axisRows.push(check("Gridlines", format.value.gridlines, (v) => (format.value.gridlines = v)).parentElement!);
  axisRows.push(text("Category axis title", format.category.title, (v) => (format.category.title = v)).parentElement!);
  axisRows.push(text("Value axis title", format.value.title, (v) => (format.value.title = v)).parentElement!);
  for (const end of ["min", "max"] as const) {
    const label = end === "min" ? "Value axis minimum" : "Value axis maximum";
    const input = field(label).createEl("input", { type: "number", value: format.value[end] === null ? "" : String(format.value[end]), attr: { "aria-label": label, placeholder: "Auto", step: "any" } });
    input.addEventListener("input", () => (format.value[end] = input.value === "" || !Number.isFinite(Number(input.value)) ? null : Number(input.value)));
    axisRows.push(input.parentElement!);
  }
  const colorRow = host.createDiv({ cls: "dw-chart-colors" });
  /** One colour per series, or per slice on a pie; a combo also picks each series' type. */
  const renderColors = () => {
    colorRow.replaceChildren();
    const pie = format.kind === "pie" || format.kind === "doughnut";
    for (const row of axisRows) row.toggle(!pie);
    (pie ? data.categories : data.series.map((s) => s.name)).forEach((name, i) => {
      const label = colorRow.createEl("label", { text: name });
      const input = label.createEl("input", { type: "color", value: (format.colors[i] ?? palette[i % palette.length]).slice(0, 7).toLowerCase(), attr: { "aria-label": `Colour of ${name}` } });
      input.addEventListener("input", () => (format.colors[i] = input.value.toUpperCase()));
      if (format.kind === "combo" && format.seriesKinds) {
        const kinds = format.seriesKinds;
        const type = label.createEl("select", { cls: "dropdown", attr: { "aria-label": `Type of ${name}` } });
        type.createEl("option", { text: "Column", value: "column" });
        type.createEl("option", { text: "Line", value: "line" });
        type.value = kinds[i] ?? "column";
        type.addEventListener("change", () => (kinds[i] = type.value as SeriesKind));
      }
    });
  };
  renderColors();
  return { renderColors };
}

/**
 * Insert / Edit chart: the chart's options above its data grid (categories × series). `apply` gets copies of the
 * edited data and format; what it throws shows in the dialog, else the dialog closes. A new chart can add series;
 * an imported one keeps its series count (its workbook's ranges).
 */
export function chartDialog(app: App, opts: { path?: string; data?: ChartData; format?: ChartFormat; palette: string[]; apply: (data: ChartData, format: ChartFormat) => Promise<void> }): void {
  const data: ChartData = opts.data ? structuredClone(opts.data) : { categories: ["Category 1", "Category 2", "Category 3"], series: [{ name: "Series 1", values: [10, 20, 15] }] };
  const format: ChartFormat = opts.format ? structuredClone(opts.format) : defaultChartFormat();
  const modal = new Modal(app);
  modal.titleEl.setText(opts.path ? "Edit chart" : "Insert chart");
  const host = modal.contentEl;
  const form = chartFormatForm(host, format, data, opts.palette);
  const grid = host.createDiv({ cls: "dw-chart-grid" });
  const render = () => {
    grid.replaceChildren();
    const table = grid.createEl("table");
    const heading = table.createEl("tr");
    heading.createEl("th", { text: "Category" });
    data.series.forEach(s => { const field = heading.createEl("th").createEl("input", { type: "text", value: s.name, attr: { "aria-label": "Series name" } }); field.addEventListener("input", () => { s.name = field.value; }); });
    data.categories.forEach((name, i) => {
      const row = table.createEl("tr");
      const category = row.createEl("td").createEl("input", { type: "text", value: name, attr: { "aria-label": `Category ${i + 1}` } });
      category.addEventListener("input", () => { data.categories[i] = category.value; });
      data.series.forEach(s => { const field = row.createEl("td").createEl("input", { type: "number", value: String(s.values[i]), attr: { "aria-label": `${s.name}, row ${i + 1}`, step: "any" } }); field.addEventListener("input", () => { s.values[i] = field.value === "" ? NaN : Number(field.value); }); });
      const remove = row.createEl("td").createEl("button", { text: "Remove row" });
      remove.disabled = data.categories.length === 1;
      remove.addEventListener("click", () => { data.categories.splice(i, 1); data.series.forEach(s => s.values.splice(i, 1)); render(); });
    });
    form.renderColors();
  };
  render();
  host.createEl("button", { text: "Add row" }).addEventListener("click", () => { data.categories.push(`Category ${data.categories.length + 1}`); data.series.forEach(s => s.values.push(0)); render(); });
  if (!opts.path)
    host.createEl("button", { text: "Add series" }).addEventListener("click", () => { data.series.push({ name: `Series ${data.series.length + 1}`, values: data.categories.map(() => 0) }); render(); });
  const error = host.createDiv({ cls: "pptx-props-hint" });
  const apply = host.createEl("button", { text: opts.path ? "Apply" : "Insert chart", cls: "mod-cta" });
  apply.addEventListener("click", () => {
    void (async () => {
      apply.disabled = true;
      try {
        await opts.apply(structuredClone(data), structuredClone(format));
        modal.close();
      } catch (err) {
        error.setText((err as Error).message);
      } finally {
        apply.disabled = false;
      }
    })();
  });
  host.createEl("button", { text: "Cancel" }).addEventListener("click", () => modal.close());
  modal.open();
}
