import { FileView, Menu, Modal, Notice, Scope, Setting, TFile, WorkspaceLeaf } from "obsidian";
import { XlsxWorkbook } from "../../infrastructure/xlsx/xlsx-workbook";
import { Sheet, Workbook } from "../../domain/xlsx/workbook-model";
import { Edits } from "../../domain/xlsx/commands";
import { FormulaSyntaxError } from "../../domain/xlsx/formula";
import { Rect, cellName, colName, parseRange, rangeName } from "../../domain/xlsx/references";
import { adjustDecimals, formatGeneral, parseInput } from "../../domain/xlsx/number-format";
import { BorderStyle, StyleChange } from "../../domain/xlsx/styles";
import { Clip, copyRange, pasteClip, pasteText } from "../../domain/xlsx/clipboard";
import { CellMatch, findCells, replaceCells } from "../../domain/xlsx/find";
import * as ops from "../../domain/xlsx/sheet-edit";
import { numText } from "../../domain/xlsx/formula-evaluator";
import { FONTS, PaperTheme, Toolbar, promptText, showValue } from "../common/toolbar";
import { exportPdf } from "../common/pdf";
import { Grid, Hit } from "./grid";
import { Picture, sheetPictures } from "../../domain/xlsx/drawing";
import { addRule, CfStyle, ConditionalFormats, deleteRule, readRules } from "../../domain/xlsx/conditional";
import { checkValue, invalidateValidations, listItems, removeValidation, setValidation, Validation, validationAt } from "../../domain/xlsx/validation";
import { conditionalForm, validationForm } from "./rule-forms";
import { chartModel, chartSource, deleteChart, formatChart, insertChart, setChartSource, SheetChart, sheetCharts } from "../../domain/xlsx/charts";
import type { ChartModel } from "../../domain/pptx/chart";
import { renderChart } from "../pptx/chart-renderer";
import { chartFormatForm, defaultChartFormat } from "../common/chart-form";
import { IMAGE_MIME } from "../../core/utils/images";

export const VIEW_TYPE_XLSX = "doc-works-xlsx-view";
const THEME_KEY = "doc-works-xlsx-theme";
const FONT_SIZES = [6, 7, 8, 9, 10, 11, 12, 14, 18, 24, 36];
const WHOLE_COL = (c1: number, c2: number): Rect => ({ r1: 0, c1, r2: 1048575, c2 });

const NUMBER_FORMATS: [string, string][] = [
  ["Automatic", "General"],
  ["Plain text", "@"],
  ["Number (1,000.12)", "#,##0.00"],
  ["Percent (10.12%)", "0.00%"],
  ["Scientific (1.01E+03)", "0.00E+00"],
  ["Accounting ($ (1,000.12))", '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'],
  ["Financial ((1,000.12))", "#,##0.00;(#,##0.00)"],
  ["Currency ($1,000.12)", '"$"#,##0.00'],
  ["Currency rounded ($1,000)", '"$"#,##0'],
  ["Date (9/26/2008)", "m/d/yyyy"],
  ["Time (3:59:00 PM)", "h:mm:ss AM/PM"],
  ["Date time (9/26/2008 15:59:00)", "m/d/yyyy h:mm:ss"],
  ["Duration (24:01:00)", "[h]:mm:ss"],
];

export class XlsxView extends FileView {
  allowNoFile = false;
  private edits: Edits | null = null;
  private sheetIndex = 0;
  private grid: Grid | null = null;
  private theme: PaperTheme = this.app.loadLocalStorage(THEME_KEY) === "light" ? "light" : "dark";
  private nameBox!: HTMLInputElement;
  private formulaInput!: HTMLTextAreaElement;
  private tabsEl!: HTMLElement;
  private statusEl!: HTMLElement;
  private saveIndicatorEl: HTMLElement | null = null;
  private fontSelect: HTMLSelectElement | null = null;
  private sizeSelect: HTMLSelectElement | null = null;
  private formatSelect: HTMLSelectElement | null = null;
  private toggles = new Map<string, HTMLElement>();
  private clip: Clip | null = null;
  /** The next paste event (from the menu) pastes values only. */
  private valuesOnly = false;
  private dirty = false;
  private saveTimer: number | null = null;
  private saving: Promise<void> = Promise.resolve();
  private lastWrite = 0;
  private findEl: HTMLElement | null = null;
  /** Object URLs of images by zip path ("" while loading). */
  private imageUrls = new Map<string, string>();
  private cfCache = new Map<string, ConditionalFormats>();
  private invalidCache = new Map<string, boolean>();
  /** Charts per sheet path, with their models, until the next edit. */
  private chartCache = new Map<string, { chart: SheetChart; model: ChartModel }[]>();
  /** Drawn charts by chart path: redrawn only when their data, format or size change. */
  private chartSvgs = new Map<string, { sig: string; el: HTMLElement }>();
  private pictureCache = new Map<string, Picture[]>();

  constructor(leaf: WorkspaceLeaf) {
    super(leaf);
    // Spreadsheet shortcuts win over app hotkeys while the grid has focus; in the editor they're normal text keys.
    this.scope = new Scope(this.app.scope);
    const keys: [("Mod" | "Shift" | "Alt")[], string, () => void][] = [
      [["Mod"], "z", () => this.undo()],
      [["Mod", "Shift"], "z", () => this.redo()],
      [["Mod"], "y", () => this.redo()],
      [["Mod"], "b", () => this.toggle("bold")],
      [["Mod"], "i", () => this.toggle("italic")],
      [["Mod"], "u", () => this.toggle("underline")],
      [["Mod", "Shift"], "x", () => this.toggle("strike")],
      [["Mod"], "a", () => this.grid?.select({ row: 0, col: 0 }, { row: 1048575, col: 16383 })],
      [["Mod"], "d", () => this.fillFrom("down")],
      [["Mod"], "r", () => this.fillFrom("right")],
      [["Mod"], ";", () => this.insertNow("date")],
      [["Mod", "Shift"], ";", () => this.insertNow("time")],
    ];
    for (const [mods, key, run] of keys) {
      this.scope.register(mods, key, () => {
        if (!this.grid || document.activeElement !== this.grid.input || this.grid.edit) return true;
        run();
        return false;
      });
    }
    this.scope.register(["Mod"], "f", () => {
      this.openFind();
      return false;
    });
    this.registerEvent(
      this.app.vault.on("modify", (f) => {
        if (f === this.file && Date.now() - this.lastWrite > 2000) void this.onExternalChange();
      })
    );
  }

  getViewType(): string {
    return VIEW_TYPE_XLSX;
  }
  getIcon(): string {
    return "table";
  }
  getDisplayText(): string {
    return this.file?.basename ?? "Spreadsheet";
  }
  canAcceptExtension(extension: string): boolean {
    return extension === "xlsx";
  }

  onPaneMenu(menu: Menu, source: string): void {
    super.onPaneMenu(menu, source);
    menu.addItem((i) => i.setTitle("Export to PDF").setIcon("file-output").setSection("action").onClick(() => void this.exportPdf()));
  }

  private get wb(): Workbook {
    return this.edits!.wb;
  }

  private get sheet() {
    const sheets = this.wb.sheets;
    this.sheetIndex = Math.min(this.sheetIndex, sheets.length - 1);
    return sheets[this.sheetIndex];
  }

  // ---------------------------------------------------------------------------
  // Lifecycle
  // ---------------------------------------------------------------------------

  async onLoadFile(file: TFile): Promise<void> {
    this.contentEl.empty();
    this.contentEl.addClass("dw-root", "xlsx-editor-root");
    const status = this.contentEl.createDiv({ cls: "dw-loading", text: "Loading spreadsheet..." });
    try {
      this.edits = new Edits(Workbook.open(await XlsxWorkbook.load(await this.app.vault.readBinary(file))));
    } catch (err) {
      status.setText(`Could not open this .xlsx file: ${(err as Error).message}`);
      console.error(err);
      this.edits = null;
      return;
    }
    status.remove();
    this.edits.onReload = () => this.afterReload();
    const view = this.wb.pkg.workbook.getElementsByTagNameNS("*", "workbookView")[0];
    this.sheetIndex = Number(view?.getAttribute("activeTab") ?? 0) || 0;
    if (this.sheet.ref.state !== "visible") this.sheetIndex = Math.max(0, this.wb.sheets.findIndex((s) => s.ref.state === "visible"));

    this.buildToolbar(this.contentEl.createDiv({ cls: "dw-toolbar xlsx-toolbar" }));
    this.buildFormulaBar(this.contentEl.createDiv({ cls: "xlsx-formula-bar" }));
    this.grid = new Grid(this.contentEl.createDiv({ cls: "xlsx-body" }), {
      sheet: () => this.sheet,
      isStale: (c) => this.edits!.calc.isStale(c),
      isCircular: (c) => this.edits!.calc.isCircular(c),
      commit: (row, col, raw) => this.commit(row, col, raw),
      selectionChanged: () => this.syncSelection(),
      editing: (text) => {
        if (text !== null && document.activeElement !== this.formulaInput) this.formulaInput.value = text;
        if (text === null) this.syncSelection();
      },
      clear: (r) => this.change(() => this.edits!.clear(this.sheet, r)),
      fill: (src, target) => this.change(() => this.edits!.fill(this.sheet, src, target)),
      resize: (axis, a, b, px) => this.change(() => (axis === "col" ? ops.setColWidth(this.edits!, this.sheet, a, b, px) : ops.setRowHeight(this.edits!, this.sheet, a, b, px))),
      autofit: (cols) => this.autofit(cols),
      contextMenu: (e, hit) => this.contextMenu(e, hit),
      filterMenu: (col, anchor) => this.filterMenu(col, anchor),
      pictures: () => this.pictures(),
      imageUrl: (path) => this.imageUrl(path),
      conditional: (row, col) => this.conditional(this.sheet, row, col),
      dropdown: (row, col) => validationAt(this.sheet, row, col)?.type === "list",
      invalid: (row, col) => this.invalid(row, col),
      openDropdown: (row, col, at) => this.openDropdown(row, col, at),
      charts: () => this.charts(),
      chartMenu: (key, e) => this.chartMenu(key, e),
    });
    const bottom = this.contentEl.createDiv({ cls: "xlsx-bottom" });
    this.tabsEl = bottom.createDiv({ cls: "xlsx-tabs" });
    this.statusEl = bottom.createDiv({ cls: "xlsx-status" });
    const input = this.grid.input;
    input.addEventListener("copy", (e) => this.onCopy(e, false));
    input.addEventListener("cut", (e) => this.onCopy(e, true));
    input.addEventListener("paste", (e) => this.onPaste(e));
    this.setTheme(this.theme);
    this.renderTabs();
    const frozen = this.sheet.frozen;
    this.grid.select({ row: frozen.rows, col: frozen.cols });
    this.grid.render();
    window.requestAnimationFrame(() => this.grid?.render());
    this.grid.focusGrid();
  }

  async onUnloadFile(_file: TFile): Promise<void> {
    if (this.grid?.edit) this.grid.commitEdit();
    if (this.saveTimer) window.clearTimeout(this.saveTimer);
    if (this.dirty) await this.saveNow();
    await this.saving;
    this.closeFind();
    this.grid?.destroy();
    this.grid = null;
    for (const url of this.imageUrls.values()) if (url) URL.revokeObjectURL(url);
    this.imageUrls.clear();
    this.pictureCache.clear();
    this.edits = null;
    this.clip = null;
  }

  async requestSave(): Promise<void> {
    if (this.saveTimer) window.clearTimeout(this.saveTimer);
    if (this.dirty) await this.saveNow();
  }

  onResize(): void {
    this.grid?.render();
  }

  /** The file changed outside this view: reload it, unless there are edits it would throw away. */
  private async onExternalChange(): Promise<void> {
    if (!this.file) return;
    if (this.dirty) {
      new Notice(`${this.file.name} changed on disk. Your unsaved edits here will replace that version.`);
      return;
    }
    const pos = this.grid ? { ...this.grid.active } : null;
    const index = this.sheetIndex;
    await this.onLoadFile(this.file);
    this.sheetIndex = Math.min(index, this.wb.sheets.length - 1);
    if (pos) this.grid?.select(pos);
    this.renderTabs();
  }

  // ---------------------------------------------------------------------------
  // Editing
  // ---------------------------------------------------------------------------

  /** Run an edit; report errors (like a malformed formula) instead of throwing; redraw and save. */
  private change(fn: () => void): boolean {
    if (!this.edits) return false;
    try {
      fn();
    } catch (err) {
      new Notice(err instanceof FormulaSyntaxError ? `There's a problem with this formula: ${err.message}` : (err as Error).message);
      return false;
    }
    this.cfCache.clear();
    this.invalidCache.clear();
    this.chartCache.clear();
    for (const sheet of this.wb.sheets) invalidateValidations(sheet);
    this.grid?.invalidate();
    this.syncSelection();
    this.markDirty();
    return true;
  }

  /** Conditional formatting of a cell, from rules worked out per sheet until the next edit. */
  private conditional(sheet: Sheet, row: number, col: number): CfStyle | null {
    let cf = this.cfCache.get(sheet.ref.path);
    if (!cf) {
      cf = new ConditionalFormats(sheet, (r, c, f) => this.edits!.calc.evaluate(sheet, r, c, f));
      this.cfCache.set(sheet.ref.path, cf);
    }
    return cf.empty ? null : cf.at(row, col);
  }

  private commit(row: number, col: number, raw: string): boolean {
    const sheet = this.sheet;
    const rule = validationAt(sheet, row, col);
    if (rule?.reject && !raw.startsWith("=") && rule.type !== "custom") {
      // Refused before it's written: the editor stays open, as in Sheets.
      if (!checkValue(rule, parseInput(raw, this.wb.date1904).value, sheet, row, col, this.evaluator(sheet))) {
        new Notice(this.invalidMessage(rule));
        return false;
      }
    }
    if (!this.change(() => this.edits!.input(sheet, row, col, raw))) return false;
    // A custom formula (or a formula's result) can only be checked once the value is in place.
    if (rule?.reject && (raw.startsWith("=") || rule.type === "custom") && !checkValue(rule, sheet.at(row, col)?.value ?? null, sheet, row, col, this.evaluator(sheet))) {
      this.undo();
      new Notice(this.invalidMessage(rule));
      return false;
    }
    return true;
  }

  private evaluator(sheet: Sheet) {
    return (r: number, c: number, f: string) => this.edits!.calc.evaluate(sheet, r, c, f);
  }

  private invalidMessage(rule: Validation): string {
    const custom = rule.el.getAttribute("error");
    if (custom) return custom;
    if (rule.type === "list") return "Input must be an item on the specified list.";
    if (rule.type === "decimal" || rule.type === "whole") return `Input must be a number ${rule.operator === "between" ? `between ${rule.formulas[0]} and ${rule.formulas[1]}` : `that meets the rule (${rule.operator} ${rule.formulas[0] ?? ""})`}.`;
    return "The data you entered violates the data validation rules set on this cell.";
  }

  /** Cells whose value breaks a warning-only rule (or was there before a rule). */
  private invalid(row: number, col: number): boolean {
    const sheet = this.sheet;
    const key = `${sheet.ref.path}:${row}:${col}`;
    let bad = this.invalidCache.get(key);
    if (bad === undefined) {
      const rule = validationAt(sheet, row, col);
      const value = sheet.at(row, col)?.value ?? null;
      bad = !!rule && value !== null && !checkValue(rule, value, sheet, row, col, this.evaluator(sheet));
      this.invalidCache.set(key, bad);
    }
    return bad;
  }

  private openDropdown(row: number, col: number, at: { x: number; y: number }): void {
    const sheet = this.sheet;
    const rule = validationAt(sheet, row, col);
    if (!rule) return;
    const current = sheet.display(row, col).text;
    const menu = new Menu();
    for (const item of listItems(rule, sheet))
      menu.addItem((i) => i.setTitle(item).setChecked(item === current).onClick(() => {
        this.commit(row, col, item);
        this.grid?.focusGrid();
      }));
    if (current) menu.addSeparator().addItem((i) => i.setTitle("Clear").onClick(() => this.commit(row, col, "")));
    menu.showAtPosition(at);
  }

  // ---------------------------------------------------------------------------
  // Charts
  // ---------------------------------------------------------------------------

  private sheetChartModels(): { chart: SheetChart; model: ChartModel }[] {
    const s = this.sheet;
    let list = this.chartCache.get(s.ref.path);
    if (!list) {
      list = sheetCharts(this.wb, s).map((chart) => ({ chart, model: chartModel(this.wb, chart.path) }));
      this.chartCache.set(s.ref.path, list);
    }
    return list;
  }

  private charts() {
    return this.sheetChartModels().map(({ chart, model }) => ({
      key: chart.path,
      ...chart.box,
      draw: (el: HTMLElement) => {
        const sig = JSON.stringify([model.data, model.format, Math.round(chart.box.w), Math.round(chart.box.h)]);
        let cached = this.chartSvgs.get(chart.path);
        if (!cached || cached.sig !== sig) {
          const host = createDiv();
          // Points, as on slides: the chart's pixel size at 96 dpi.
          renderChart(host, model, { w: chart.box.w * 0.75, h: chart.box.h * 0.75 });
          this.chartSvgs.set(chart.path, (cached = { sig, el: host }));
        }
        for (const child of Array.from(cached.el.children)) el.appendChild(child.cloneNode(true));
      },
    }));
  }

  private chartMenu(path: string, e: MouseEvent): void {
    const found = this.sheetChartModels().find((c) => c.chart.path === path);
    if (!found) return;
    const menu = new Menu();
    menu.addItem((i) => i.setTitle("Edit chart").setIcon("pencil").onClick(() => this.chartDialog(path)));
    menu.addItem((i) => i.setTitle("Delete chart").setIcon("trash").onClick(() => this.change(() => deleteChart(this.edits!, this.sheet, found.chart))));
    menu.showAtMouseEvent(e);
  }

  private insertChart(): void {
    if (!this.edits || !this.grid) return;
    let sel = this.grid.selection();
    // A single cell: the block of data around it, as Sheets guesses.
    if (sel.r1 === sel.r2 && sel.c1 === sel.c2) sel = this.sheet.used() ?? sel;
    let path = "";
    if (this.change(() => (path = insertChart(this.edits!, this.sheet, sel)))) this.chartDialog(path);
  }

  private chartDialog(path: string): void {
    const found = this.sheetChartModels().find((c) => c.chart.path === path);
    if (!found || !this.edits) return;
    const format = structuredClone(found.model.format ?? defaultChartFormat());
    const modal = new Modal(this.app);
    modal.titleEl.setText("Chart editor");
    const host = modal.contentEl;
    const sourceRow = host.createDiv({ cls: "dw-chart-option" });
    sourceRow.createSpan({ text: "Data range" });
    const original = chartSource(this.wb, path);
    const source = sourceRow.createEl("input", { type: "text", value: original, attr: { "aria-label": "Data range" } });
    chartFormatForm(host, format, found.model.data, found.model.palette ?? []);
    const error = host.createDiv({ cls: "pptx-props-hint" });
    new Setting(host).addButton((b) =>
      b.setButtonText("Apply").setCta().onClick(() => {
        const ok = this.change(() => {
          if (source.value.trim() && source.value.trim() !== original) setChartSource(this.edits!, path, source.value.trim());
          formatChart(this.edits!, path, format);
        });
        if (ok) modal.close();
        else error.setText("Check the data range.");
      })
    );
    modal.open();
  }

  // ---------------------------------------------------------------------------
  // Conditional formatting and data validation dialogs
  // ---------------------------------------------------------------------------

  private selectedRange(): string {
    return rangeName(this.grid!.selection());
  }

  private conditionalDialog(): void {
    if (!this.edits) return;
    const sheet = this.sheet;
    const modal = new Modal(this.app);
    modal.titleEl.setText("Conditional format rules");
    modal.modalEl.addClass("xlsx-rules-modal");
    const draw = () => {
      modal.contentEl.empty();
      conditionalForm(modal.contentEl, {
        rules: readRules(this.sheet),
        range: this.selectedRange(),
        swatch: (r) => (r.type === "colorScale" ? { fill: r.colors[r.colors.length - 1] } : r.dxf === null ? {} : sheet.wb.styles.dxf(r.dxf)),
        add: (ranges, rule) => {
          if (this.change(() => addRule(this.edits!, this.sheet, ranges, rule))) draw();
        },
        remove: (rule) => {
          if (this.change(() => deleteRule(this.edits!, this.sheet, rule))) draw();
        },
      });
    };
    draw();
    modal.open();
  }

  private validationDialog(): void {
    if (!this.edits) return;
    const sel = this.grid!.selection();
    const modal = new Modal(this.app);
    modal.titleEl.setText("Data validation");
    modal.modalEl.addClass("xlsx-rules-modal");
    validationForm(modal.contentEl, {
      current: validationAt(this.sheet, sel.r1, sel.c1),
      range: rangeName(sel),
      add: (ranges, rule) => {
        if (this.change(() => setValidation(this.edits!, this.sheet, ranges, rule))) modal.close();
      },
      remove: (rule) => {
        if (this.change(() => removeValidation(this.edits!, this.sheet, rule.ranges))) modal.close();
      },
    });
    modal.open();
  }

  private pictures(): Picture[] {
    const s = this.sheet;
    let pics = this.pictureCache.get(s.ref.path);
    if (!pics) {
      pics = sheetPictures(this.wb, s);
      this.pictureCache.set(s.ref.path, pics);
    }
    return pics;
  }

  /** An image's object URL, loading it on first use (the grid redraws when it arrives). */
  private imageUrl(path: string): string | null {
    const url = this.imageUrls.get(path);
    if (url !== undefined) return url || null;
    this.imageUrls.set(path, "");
    const edits = this.edits;
    void edits?.wb.pkg.media(path).then((bytes) => {
      if (this.edits !== edits) return;
      const ext = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
      this.imageUrls.set(path, URL.createObjectURL(new Blob([bytes as BlobPart], { type: IMAGE_MIME[ext] ?? `image/${ext}` })));
      this.grid?.schedule();
    }).catch(() => undefined);
    return null;
  }

  private afterReload(): void {
    this.pictureCache.clear();
    this.cfCache.clear();
    this.invalidCache.clear();
    this.chartCache.clear();
    this.sheetIndex = Math.min(this.sheetIndex, this.wb.sheets.length - 1);
    this.renderTabs();
  }

  private undo(): void {
    if (this.edits?.canUndo) this.change(() => this.edits!.undo());
  }
  private redo(): void {
    if (this.edits?.canRedo) this.change(() => this.edits!.redo());
  }

  private get sel(): Rect {
    return this.grid!.selection();
  }

  private format(change: StyleChange): void {
    this.change(() => this.edits!.format(this.sheet, this.sel, change));
  }

  private toggle(key: "bold" | "italic" | "underline" | "strike" | "wrap"): void {
    this.format({ [key]: this.edits!.toggled(this.sheet, this.sel, key) });
  }

  private borders(kind: "all" | "outer" | "none" | "top" | "bottom" | "left" | "right"): void {
    const r = this.sel;
    const e = this.edits!;
    const s = this.sheet;
    const thin: BorderStyle = "thin";
    this.change(() =>
      e.run(() => {
        const edge = (rect: Rect, borders: StyleChange["borders"]) => e.format(s, rect, { borders });
        if (kind === "all") edge(r, { top: thin, bottom: thin, left: thin, right: thin });
        else if (kind === "none") edge(r, { top: null, bottom: null, left: null, right: null });
        if (kind === "outer" || kind === "top") edge({ ...r, r2: r.r1 }, { top: thin });
        if (kind === "outer" || kind === "bottom") edge({ ...r, r1: r.r2 }, { bottom: thin });
        if (kind === "outer" || kind === "left") edge({ ...r, c2: r.c1 }, { left: thin });
        if (kind === "outer" || kind === "right") edge({ ...r, c1: r.c2 }, { right: thin });
        return [];
      })
    );
  }

  private decimals(delta: number): void {
    const s = this.sheet;
    const cell = s.at(this.grid!.active.row, this.grid!.active.col);
    const code = this.wb.styles.numFmt(cell?.style ?? 0);
    this.format({ numFmt: adjustDecimals(code, delta, typeof cell?.value === "number" ? cell.value : 0) });
  }

  private fillFrom(dir: "down" | "right"): void {
    const r = this.sel;
    if (dir === "down" && r.r2 > r.r1) this.change(() => this.edits!.fill(this.sheet, { ...r, r2: r.r1 }, r));
    if (dir === "right" && r.c2 > r.c1) this.change(() => this.edits!.fill(this.sheet, { ...r, c2: r.c1 }, r));
  }

  private insertNow(kind: "date" | "time"): void {
    const d = new Date();
    const pad = (n: number) => String(n).padStart(2, "0");
    const text = kind === "date" ? `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}` : `${d.getHours()}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
    const { row, col } = this.grid!.active;
    this.commit(row, col, text);
  }

  /** Widen columns to fit their longest shown text. */
  private autofit(cols: number[]): void {
    const s = this.sheet;
    const ctx = createEl("canvas").getContext("2d");
    if (!ctx) return;
    this.change(() =>
      this.edits!.run(() => {
        for (const col of cols) {
          let w = 0;
          for (const c of s.cellsIn(WHOLE_COL(col, col))) {
            const css = this.wb.styles.css(c.style);
            ctx.font = `${css["font-style"] ?? ""} ${css["font-weight"] ?? ""} ${css["font-size"]} ${css["font-family"]}`;
            w = Math.max(w, ctx.measureText(s.display(c.row, c.col).text).width);
          }
          if (w) ops.setColWidth(this.edits!, this.sheet, col, col, Math.ceil(w + 10));
        }
        return [];
      })
    );
  }

  // ---------------------------------------------------------------------------
  // Clipboard
  // ---------------------------------------------------------------------------

  private onCopy(e: ClipboardEvent, cut: boolean): void {
    if (!this.grid || this.grid.edit || !e.clipboardData) return;
    e.preventDefault();
    this.clip = copyRange(this.sheet, this.sel, cut);
    e.clipboardData.setData("text/plain", this.clip.text);
    e.clipboardData.setData("text/html", this.clip.html);
  }

  private onPaste(e: ClipboardEvent): void {
    if (!this.grid || this.grid.edit || !e.clipboardData) return;
    e.preventDefault();
    const valuesOnly = this.valuesOnly;
    this.valuesOnly = false;
    this.pasteFrom(e.clipboardData.getData("text/plain"), valuesOnly);
  }

  private pasteFrom(text: string, valuesOnly = false): void {
    const norm = (t: string) => t.replace(/\r\n?/g, "\n").replace(/\n$/, "");
    let area: Rect | null = null;
    this.change(() => {
      if (this.clip && norm(text) === norm(this.clip.text) && !valuesOnly) {
        area = pasteClip(this.edits!, this.sheet, this.sel, this.clip);
        if (this.clip.cut) this.clip = null;
      } else if (text) area = pasteText(this.edits!, this.sheet, this.sel, text);
    });
    if (area) this.grid!.selectRect(area);
  }

  /**
   * Cut / copy / paste from a menu: the app runs the edit command on the grid's input, which raises the same
   * clipboard events as the keyboard shortcuts (no direct clipboard reads or writes).
   */
  private editCommand(command: "copy" | "cut" | "paste", valuesOnly = false): void {
    if (!this.grid) return;
    this.valuesOnly = valuesOnly;
    this.grid.input.focus();
    const contents = (window.require("electron") as { remote: { getCurrentWebContents(): Record<typeof command, () => void> } }).remote.getCurrentWebContents();
    contents[command]();
    // The event arrives shortly after; a later Ctrl+V must not paste values only.
    window.setTimeout(() => (this.valuesOnly = false), 1000);
  }

  // ---------------------------------------------------------------------------
  // Toolbar, formula bar, tabs, status
  // ---------------------------------------------------------------------------

  private buildToolbar(bar: HTMLElement): void {
    bar.addEventListener("mousedown", (e) => {
      if ((e.target as HTMLElement).closest("button")) e.preventDefault();
    });
    const tb = new Toolbar(bar);
    tb.icon("undo-2", "Undo (Ctrl+Z)", () => this.undo());
    tb.icon("redo-2", "Redo (Ctrl+Y)", () => this.redo());
    tb.gap();
    this.formatSelect = tb.select("Number format", "123", NUMBER_FORMATS, (code) => this.format({ numFmt: code }));
    tb.button("$", "Format as currency", () => this.format({ numFmt: '"$"#,##0.00' }));
    tb.button("%", "Format as percent", () => this.format({ numFmt: "0%" }));
    tb.button(".0←", "Decrease decimal places", () => this.decimals(-1));
    tb.button(".00→", "Increase decimal places", () => this.decimals(1));
    tb.gap();
    this.fontSelect = tb.select("Font", "Font", FONTS.map((f) => [f, f]), (fontName) => this.format({ fontName }));
    this.sizeSelect = tb.select("Font size", "Size", FONT_SIZES.map((s) => [String(s), String(s)]), (s) => this.format({ fontSize: Number(s) }));
    tb.gap();
    this.toggles.set("bold", tb.button("B", "Bold (Ctrl+B)", () => this.toggle("bold"), "bold"));
    this.toggles.set("italic", tb.button("I", "Italic (Ctrl+I)", () => this.toggle("italic"), "italic"));
    this.toggles.set("strike", tb.button("S", "Strikethrough (Ctrl+Shift+X)", () => this.toggle("strike"), "strike"));
    this.toggles.set("underline", tb.button("U", "Underline (Ctrl+U)", () => this.toggle("underline"), "underline"));
    tb.color("A", "Text color", "#000000", (color) => this.format({ color }));
    tb.color("▇", "Fill color", "#FFFF00", (fill) => this.format({ fill }));
    tb.button("No fill", "Remove fill color", () => this.format({ fill: null }));
    tb.iconSelect("square-dashed", "Borders", [["All borders", "all"], ["Outer borders", "outer"], ["Top", "top"], ["Bottom", "bottom"], ["Left", "left"], ["Right", "right"], ["No borders", "none"]], (v) =>
      this.borders(v as "all")
    );
    tb.gap();
    tb.icon("align-left", "Align left", () => this.format({ hAlign: "left" }));
    tb.icon("align-center", "Align center", () => this.format({ hAlign: "center" }));
    tb.icon("align-right", "Align right", () => this.format({ hAlign: "right" }));
    tb.iconSelect("align-vertical-justify-center", "Vertical align", [["Top", "top"], ["Middle", "center"], ["Bottom", "bottom"]], (v) => this.format({ vAlign: v as "top" }));
    this.toggles.set("wrap", tb.icon("wrap-text", "Wrap text", () => this.toggle("wrap")));
    tb.icon("merge", "Merge cells", () => this.mergeToggle());
    tb.gap();
    tb.iconSelect("sigma", "Functions", ["SUM", "AVERAGE", "COUNT", "MAX", "MIN", "IF", "VLOOKUP", "CONCATENATE"].map((f) => [f, f]), (f) => this.insertFunction(f));
    tb.icon("arrow-down-az", "Sort ascending by this column", () => this.sort(false));
    tb.icon("arrow-up-za", "Sort descending by this column", () => this.sort(true));
    tb.icon("filter", "Create or remove a filter", () => this.toggleFilter());
    tb.iconSelect("snowflake", "Freeze", [["No rows", "r0"], ["1 row", "r1"], ["2 rows", "r2"], ["Up to current row", "rc"], ["No columns", "c0"], ["1 column", "c1"], ["2 columns", "c2"], ["Up to current column", "cc"]], (v) =>
      this.freeze(v)
    );
    tb.icon("chart-column", "Insert chart", () => this.insertChart());
    tb.icon("palette", "Conditional formatting", () => this.conditionalDialog());
    tb.icon("list-checks", "Data validation (dropdowns and rules)", () => this.validationDialog());
    tb.icon("search", "Find and replace (Ctrl+F)", () => this.openFind());
    const { syncTheme, status } = tb.end(this.theme, (t) => this.setTheme(t), "grid");
    this.syncTheme = syncTheme;
    this.saveIndicatorEl = status;
  }

  private syncTheme: ((t: PaperTheme) => void) | null = null;

  private setTheme(theme: PaperTheme): void {
    this.theme = theme;
    this.app.saveLocalStorage(THEME_KEY, theme);
    this.contentEl.toggleClass("xlsx-theme-dark", theme === "dark");
    this.contentEl.toggleClass("xlsx-theme-light", theme === "light");
    this.syncTheme?.(theme);
    if (this.grid) {
      this.grid.dark = theme === "dark";
      this.grid.invalidate();
    }
  }

  private buildFormulaBar(bar: HTMLElement): void {
    this.nameBox = bar.createEl("input", { cls: "xlsx-name-box", type: "text", attr: { "aria-label": "Name box: type a cell or range to go there" } });
    this.nameBox.addEventListener("keydown", (e) => {
      if (e.key === "Escape") {
        this.syncSelection();
        this.grid?.focusGrid();
      }
      if (e.key !== "Enter") return;
      e.preventDefault();
      this.goTo(this.nameBox.value.trim());
    });
    this.nameBox.addEventListener("focus", () => this.nameBox.select());
    bar.createSpan({ cls: "xlsx-fx", text: "fx" });
    this.formulaInput = bar.createEl("textarea", { cls: "xlsx-formula-input", attr: { rows: "1", "aria-label": "Formula bar", spellcheck: "false" } });
    this.formulaInput.addEventListener("input", () => this.grid?.setEditText(this.formulaInput.value));
    this.formulaInput.addEventListener("focus", () => {
      if (this.grid && !this.grid.edit) this.grid.setEditText(this.formulaInput.value);
    });
    this.formulaInput.addEventListener("keydown", (e) => {
      if (!this.grid) return;
      if (e.key === "Escape") {
        e.preventDefault();
        this.grid.cancelEdit();
        this.grid.focusGrid();
      } else if (e.key === "Enter" && !e.altKey && !e.ctrlKey) {
        e.preventDefault();
        this.grid.setEditText(this.formulaInput.value);
        if (this.grid.commitEdit()) {
          this.grid.move(1, 0, false);
          this.grid.focusGrid();
        }
      }
    });
  }

  /** Name box navigation: "B7", "A1:C3", "Sheet2!A1" or a defined name. */
  private goTo(text: string): void {
    const grid = this.grid!;
    let target = text;
    const name = this.wb.names().find((n) => n.name.toLowerCase() === text.toLowerCase());
    if (name) target = name.formula;
    const bang = target.lastIndexOf("!");
    if (bang > 0) {
      const sheetName = target.slice(0, bang).replace(/^'|'$/g, "").replace(/''/g, "'");
      const index = this.wb.sheets.findIndex((s) => s.name.toLowerCase() === sheetName.toLowerCase());
      if (index < 0) return void new Notice(`There's no sheet named "${sheetName}".`);
      this.showSheet(index);
      target = target.slice(bang + 1);
    }
    const r = parseRange(target.replace(/\$/g, ""));
    if (!r) return void new Notice(`"${text}" isn't a cell, range or named range.`);
    grid.selectRect(r);
    grid.reveal({ row: r.r1, col: r.c1 });
    grid.focusGrid();
  }

  private syncSelection(): void {
    const grid = this.grid;
    if (!grid || !this.edits) return;
    const s = this.sheet;
    const r = grid.selection();
    const one = r.r1 === r.r2 && r.c1 === r.c2;
    if (document.activeElement !== this.nameBox) this.nameBox.value = one || s.mergeAt(r.r1, r.c1) ? cellName(grid.active.row, grid.active.col) : rangeName(r);
    const cell = s.at(grid.active.row, grid.active.col);
    if (!grid.edit && document.activeElement !== this.formulaInput) {
      grid.input.value = "";
      this.formulaInput.value = cell?.formula !== undefined ? "=" + cell.formula : cell ? (grid.edit ?? "") || this.rawText(grid.active.row, grid.active.col) : "";
    }
    const f = this.wb.styles.format(cell?.style ?? 0);
    for (const [k, el] of this.toggles) el.toggleClass("is-active", !!f[k as "bold"]);
    if (this.fontSelect) showValue(this.fontSelect, f.fontName);
    if (this.sizeSelect) showValue(this.sizeSelect, String(f.fontSize));
    if (this.formatSelect) showValue(this.formatSelect, NUMBER_FORMATS.find(([, c]) => c === f.numFmt) ? f.numFmt : f.numFmt === "General" ? null : f.numFmt);
    this.updateStatus(r);
  }

  private rawText(row: number, col: number): string {
    const s = this.sheet;
    const cell = s.at(row, col);
    if (!cell) return "";
    if (typeof cell.value === "number") return numText(cell.value);
    return s.display(row, col).text;
  }

  /** Sum / average / count of the selection, like the spreadsheet status bar. */
  private updateStatus(r: Rect): void {
    const s = this.sheet;
    const cells = (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1) > 1 ? s.cellsIn(r) : [];
    const nums = cells.map((c) => c.value).filter((v): v is number => typeof v === "number");
    const filled = cells.filter((c) => c.value !== null).length;
    this.statusEl.empty();
    if (!nums.length && filled < 2) return;
    const sum = nums.reduce((a, b) => a + b, 0);
    const parts = nums.length ? [`Sum: ${formatGeneral(sum)}`, `Average: ${formatGeneral(sum / nums.length)}`, `Count: ${filled}`] : [`Count: ${filled}`];
    this.statusEl.setText(parts.join("   "));
  }

  private renderTabs(): void {
    if (!this.edits) return;
    const bar = this.tabsEl;
    bar.empty();
    const add = bar.createEl("button", { cls: "dw-btn xlsx-tab-add", text: "+", attr: { "aria-label": "Add sheet" } });
    add.addEventListener("click", () => {
      this.change(() => ops.addSheet(this.edits!, this.sheetIndex));
      this.showSheet(this.sheetIndex + 1);
    });
    this.wb.sheets.forEach((s, i) => {
      if (s.ref.state !== "visible") return;
      const tab = bar.createEl("button", { cls: i === this.sheetIndex ? "xlsx-tab is-active" : "xlsx-tab", text: s.name, attr: { "aria-label": `Sheet ${s.name}` } });
      tab.addEventListener("click", () => this.showSheet(i));
      tab.addEventListener("dblclick", () => void this.renameSheet(i));
      tab.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        this.sheetMenu(i).showAtMouseEvent(e);
      });
    });
    const hidden = this.wb.sheets.filter((s) => s.ref.state === "hidden");
    if (hidden.length) {
      const more = bar.createEl("button", { cls: "dw-btn", text: `${hidden.length} hidden ▾`, attr: { "aria-label": "Show a hidden sheet" } });
      more.addEventListener("click", (e) => {
        const menu = new Menu();
        for (const s of hidden) menu.addItem((it) => it.setTitle(`Show "${s.name}"`).onClick(() => this.change(() => ops.setSheetHidden(this.edits!, this.wb.sheets[this.wb.sheets.indexOf(s)], false))));
        menu.showAtMouseEvent(e);
      });
    }
  }

  private showSheet(index: number): void {
    if (!this.grid || index < 0 || index >= this.wb.sheets.length) return;
    if (this.grid.edit && !this.grid.commitEdit()) return;
    this.sheetIndex = index;
    this.renderTabs();
    const f = this.sheet.frozen;
    this.grid.scrollHome();
    this.grid.select({ row: f.rows, col: f.cols });
    this.grid.invalidate();
    this.grid.focusGrid();
  }

  private async renameSheet(i: number): Promise<void> {
    const name = await promptText(this.app, "Rename sheet", this.wb.sheets[i].name);
    if (name !== null) this.change(() => ops.renameSheet(this.edits!, this.wb.sheets[i], name));
  }

  private sheetMenu(i: number): Menu {
    const e = this.edits!;
    const s = () => this.wb.sheets[i];
    const menu = new Menu();
    menu.addItem((it) => it.setTitle("Rename").setIcon("pencil").onClick(() => void this.renameSheet(i)));
    menu.addItem((it) => it.setTitle("Duplicate").setIcon("copy").onClick(() => {
      this.change(() => ops.duplicateSheet(e, s()));
      this.showSheet(i + 1);
    }));
    menu.addItem((it) => it.setTitle("Delete").setIcon("trash").onClick(() => {
      if (this.change(() => ops.deleteSheet(e, s()))) this.showSheet(Math.max(0, Math.min(i, this.wb.sheets.length - 1)));
    }));
    menu.addItem((it) => it.setTitle("Hide").setIcon("eye-off").onClick(() => {
      if (this.change(() => ops.setSheetHidden(e, s(), true))) this.showSheet(this.wb.sheets.findIndex((x) => x.ref.state === "visible"));
    }));
    menu.addSeparator();
    if (i > 0) menu.addItem((it) => it.setTitle("Move left").setIcon("arrow-left").onClick(() => {
      this.change(() => ops.moveSheet(e, i, i - 1));
      this.showSheet(i - 1);
    }));
    if (i < this.wb.sheets.length - 1) menu.addItem((it) => it.setTitle("Move right").setIcon("arrow-right").onClick(() => {
      this.change(() => ops.moveSheet(e, i, i + 1));
      this.showSheet(i + 1);
    }));
    return menu;
  }

  // ---------------------------------------------------------------------------
  // Menus and commands
  // ---------------------------------------------------------------------------

  private contextMenu(e: MouseEvent, hit: Hit): void {
    const ed = this.edits!;
    const r = this.sel;
    const menu = new Menu();
    const item = (title: string, icon: string, run: () => void) => menu.addItem((it) => it.setTitle(title).setIcon(icon).onClick(run));
    const rows = `${r.r2 - r.r1 + 1} ${r.r2 > r.r1 ? "rows" : "row"}`;
    const cols = `${r.c2 - r.c1 + 1} ${r.c2 > r.c1 ? "columns" : "column"}`;
    item("Cut", "scissors", () => this.editCommand("cut"));
    item("Copy", "copy", () => this.editCommand("copy"));
    item("Paste", "clipboard", () => this.editCommand("paste"));
    item("Paste values only", "clipboard-type", () => this.editCommand("paste", true));
    menu.addSeparator();
    if (hit.area !== "colHeader") {
      item(`Insert ${rows} above`, "arrow-up", () => this.change(() => ops.insertRows(ed, this.sheet, r.r1, r.r2 - r.r1 + 1)));
      item(`Insert ${rows} below`, "arrow-down", () => this.change(() => ops.insertRows(ed, this.sheet, r.r2 + 1, r.r2 - r.r1 + 1)));
    }
    if (hit.area !== "rowHeader") {
      item(`Insert ${cols} left`, "arrow-left", () => this.change(() => ops.insertCols(ed, this.sheet, r.c1, r.c2 - r.c1 + 1)));
      item(`Insert ${cols} right`, "arrow-right", () => this.change(() => ops.insertCols(ed, this.sheet, r.c2 + 1, r.c2 - r.c1 + 1)));
    }
    menu.addSeparator();
    if (hit.area !== "colHeader") item(`Delete ${rows}`, "trash", () => this.change(() => ops.deleteRows(ed, this.sheet, r.r1, r.r2 - r.r1 + 1)));
    if (hit.area !== "rowHeader") item(`Delete ${cols}`, "trash", () => this.change(() => ops.deleteCols(ed, this.sheet, r.c1, r.c2 - r.c1 + 1)));
    item("Clear contents", "eraser", () => this.change(() => ed.clear(this.sheet, r)));
    menu.addSeparator();
    if (hit.area === "rowHeader") {
      item(`Hide ${rows}`, "eye-off", () => this.change(() => ops.setHidden(ed, this.sheet, "row", r.r1, r.r2, true)));
      item("Unhide rows", "eye", () => this.change(() => ops.setHidden(ed, this.sheet, "row", Math.max(0, r.r1 - 1), r.r2 + 1, false)));
      item("Resize rows…", "move-vertical", () => void this.askSize("row", r.r1, r.r2));
    } else if (hit.area === "colHeader") {
      item(`Hide ${cols}`, "eye-off", () => this.change(() => ops.setHidden(ed, this.sheet, "col", r.c1, r.c2, true)));
      item("Unhide columns", "eye", () => this.change(() => ops.setHidden(ed, this.sheet, "col", Math.max(0, r.c1 - 1), r.c2 + 1, false)));
      item("Resize columns…", "move-horizontal", () => void this.askSize("col", r.c1, r.c2));
    } else {
      item("Sort range ascending", "arrow-down-az", () => this.sort(false));
      item("Sort range descending", "arrow-up-za", () => this.sort(true));
      item(this.sheet.merges.some((m) => m.r1 <= r.r2 && r.r1 <= m.r2 && m.c1 <= r.c2 && r.c1 <= m.c2) ? "Unmerge cells" : "Merge cells", "merge", () => this.mergeToggle());
    }
    item("Freeze up to here", "snowflake", () => this.change(() => ops.freeze(ed, this.sheet, hit.area === "colHeader" ? 0 : r.r2 + 1, hit.area === "rowHeader" ? 0 : r.c2 + 1)));
    menu.showAtMouseEvent(e);
  }

  private async askSize(axis: "row" | "col", a: number, b: number): Promise<void> {
    const s = this.sheet;
    const current = axis === "col" ? s.colWidth(a) : s.rowHeight(a);
    const v = await promptText(this.app, axis === "col" ? "Column width in pixels" : "Row height in pixels", String(current));
    const px = Number(v);
    if (v === null || !Number.isFinite(px) || px <= 0) return;
    this.change(() => (axis === "col" ? ops.setColWidth(this.edits!, s, a, b, px) : ops.setRowHeight(this.edits!, s, a, b, px)));
  }

  private mergeToggle(): void {
    const r = this.sel;
    const merged = this.sheet.merges.some((m) => m.r1 <= r.r2 && r.r1 <= m.r2 && m.c1 <= r.c2 && r.c1 <= m.c2);
    this.change(() => (merged ? ops.unmerge(this.edits!, this.sheet, r) : ops.merge(this.edits!, this.sheet, r)));
  }

  private insertFunction(name: string): void {
    const grid = this.grid!;
    const r = this.sel;
    const multi = r.r1 !== r.r2 || r.c1 !== r.c2;
    if (multi) {
      // Like Sheets' Σ: put the function below a column selection (or right of a row) over that range.
      const below = r.r2 - r.r1 >= r.c2 - r.c1;
      const at = below ? { row: r.r2 + 1, col: r.c1 } : { row: r.r1, col: r.c2 + 1 };
      grid.select(at);
      grid.beginEdit(`=${name}(${rangeName(below ? { ...r, c2: r.c1 } : { ...r, r2: r.r1 })})`, false);
    } else grid.beginEdit(`=${name}(`, false);
  }

  private sort(desc: boolean): void {
    const s = this.sheet;
    let r = this.sel;
    const col = this.grid!.active.col;
    if (r.r1 === r.r2) {
      // A single cell sorts the sheet's data below any frozen header rows.
      const used = s.used();
      if (!used) return;
      r = { r1: Math.max(s.frozen.rows, used.r1), c1: used.c1, r2: used.r2, c2: used.c2 };
    } else if (r.r2 === 1048575) {
      const used = s.used();
      if (!used) return;
      r = { ...r, r1: Math.max(r.r1, s.frozen.rows), r2: used.r2, c1: Math.max(r.c1, used.c1), c2: Math.min(r.c2, used.c2) };
    }
    if (s.filter && r.r1 <= s.filter.r1 && r.r2 >= s.filter.r1) r = { ...r, r1: s.filter.r1 + 1 };
    this.change(() => ops.sortRange(this.edits!, s, r, [{ col, desc }]));
  }

  private toggleFilter(): void {
    const s = this.sheet;
    if (s.filter) return void this.change(() => ops.setFilter(this.edits!, s, null));
    let r = this.sel;
    if (r.r1 === r.r2 && r.c1 === r.c2) {
      const used = s.used();
      if (!used) return void new Notice("Add some data first: a filter needs a header row and values below it.");
      r = used;
    }
    this.change(() => ops.setFilter(this.edits!, s, r));
  }

  private freeze(v: string): void {
    const s = this.sheet;
    const { row, col } = this.grid!.active;
    const rows = v[0] === "r" ? (v[1] === "c" ? row + 1 : Number(v[1])) : s.frozen.rows;
    const cols = v[0] === "c" ? (v[1] === "c" ? col + 1 : Number(v[1])) : s.frozen.cols;
    this.change(() => ops.freeze(this.edits!, s, rows, cols));
  }

  private filterMenu(col: number, _anchor: HTMLElement): void {
    const s = this.sheet;
    const f = s.filter;
    if (!f) return;
    const values = new Set<string>();
    for (let r = f.r1 + 1; r <= f.r2; r++) values.add(s.display(r, col).text);
    const current = ops.filterValues(s, col);
    const modal = new Modal(this.app);
    modal.titleEl.setText(`Filter by ${colName(col)}`);
    const list = modal.contentEl.createDiv({ cls: "xlsx-filter-list" });
    const boxes: [string, HTMLInputElement][] = [];
    for (const v of [...values].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))) {
      const label = list.createEl("label");
      const box = label.createEl("input", { type: "checkbox" });
      box.checked = !current || current.has(v);
      label.appendText(v || "(Blanks)");
      boxes.push([v, box]);
    }
    new Setting(modal.contentEl)
      .addButton((b) => b.setButtonText("Select all").onClick(() => boxes.forEach(([, x]) => (x.checked = true))))
      .addButton((b) => b.setButtonText("Clear").onClick(() => boxes.forEach(([, x]) => (x.checked = false))));
    new Setting(modal.contentEl)
      .addButton((b) => b.setButtonText("Sort ascending").onClick(() => {
        modal.close();
        this.change(() => ops.sortRange(this.edits!, s, { ...f, r1: f.r1 + 1 }, [{ col }]));
      }))
      .addButton((b) => b.setButtonText("Sort descending").onClick(() => {
        modal.close();
        this.change(() => ops.sortRange(this.edits!, s, { ...f, r1: f.r1 + 1 }, [{ col, desc: true }]));
      }))
      .addButton((b) =>
        b.setButtonText("OK").setCta().onClick(() => {
          modal.close();
          const allowed = new Set(boxes.filter(([, x]) => x.checked).map(([v]) => v));
          this.change(() => ops.filterColumn(this.edits!, s, col, allowed.size === boxes.length ? null : allowed));
        })
      );
    modal.open();
  }

  // ---------------------------------------------------------------------------
  // Find and replace
  // ---------------------------------------------------------------------------

  private openFind(): void {
    if (!this.edits) return;
    if (this.findEl) {
      this.findEl.querySelector("input")?.focus();
      return;
    }
    const bar = (this.findEl = createDiv({ cls: "pptx-find xlsx-find" }));
    this.contentEl.insertBefore(bar, this.contentEl.querySelector(".xlsx-body"));
    const input = bar.createEl("input", { type: "search", attr: { placeholder: "Find", "aria-label": "Find" } });
    const replace = bar.createEl("input", { type: "text", attr: { placeholder: "Replace with", "aria-label": "Replace with" } });
    const check = (text: string, title: string) => {
      const label = bar.createEl("label", { cls: "pptx-find-case", attr: { "aria-label": title } });
      const box = label.createEl("input", { type: "checkbox" });
      label.appendText(text);
      return box;
    };
    const matchCase = check("Aa", "Match case");
    const entire = check("Whole cell", "Match entire cell contents");
    const formulas = check("Formulas", "Also search within formulas");
    const allSheets = check("All sheets", "Search all sheets");
    const count = bar.createSpan({ cls: "pptx-find-count" });
    let matches: CellMatch[] = [];
    let index = -1;
    const opts = () => ({ matchCase: matchCase.checked, entire: entire.checked, formulas: formulas.checked });
    const refresh = () => {
      matches = findCells(allSheets.checked ? this.wb.sheets : [this.sheet], input.value, opts());
      index = Math.min(index, matches.length - 1);
      count.setText(input.value ? (matches.length ? `${Math.max(index, 0) + 1} of ${matches.length}` : "No results") : "");
    };
    const show = (m: CellMatch) => {
      const i = this.wb.sheets.indexOf(m.sheet);
      if (i !== this.sheetIndex) this.showSheet(i);
      this.grid!.select({ row: m.row, col: m.col });
      this.grid!.reveal({ row: m.row, col: m.col });
    };
    const step = (dir: number) => {
      refresh();
      if (!matches.length) return;
      index = (index + dir + matches.length) % matches.length;
      count.setText(`${index + 1} of ${matches.length}`);
      show(matches[index]);
    };
    const button = (text: string, title: string, fn: () => void) => bar.createEl("button", { text, cls: "dw-btn", attr: { "aria-label": title } }).addEventListener("click", fn);
    button("↑", "Previous match (Shift+Enter)", () => step(-1));
    button("↓", "Next match (Enter)", () => step(1));
    button("Replace", "Replace this match", () => {
      refresh();
      const m = matches[index];
      if (!m) return step(1);
      this.change(() => replaceCells(this.edits!, [m], input.value, replace.value, opts()));
      step(0);
    });
    button("Replace all", "Replace every match", () => {
      refresh();
      let n = 0;
      this.change(() => (n = replaceCells(this.edits!, matches, input.value, replace.value, opts())));
      new Notice(`Replaced ${n} ${n === 1 ? "cell" : "cells"}.`);
      refresh();
    });
    button("×", "Close (Esc)", () => this.closeFind());
    input.addEventListener("input", () => {
      index = -1;
      refresh();
    });
    for (const box of [matchCase, entire, formulas, allSheets]) box.addEventListener("change", () => refresh());
    bar.addEventListener("keydown", (e) => {
      if (e.key === "Escape") this.closeFind();
      else if (e.key === "Enter" && e.target === input) step(e.shiftKey ? -1 : 1);
    });
    input.focus();
  }

  private closeFind(): void {
    this.findEl?.remove();
    this.findEl = null;
    this.grid?.focusGrid();
  }

  // ---------------------------------------------------------------------------
  // PDF
  // ---------------------------------------------------------------------------

  /** Export the current sheet's used range: fit to page width (landscape Letter), rows split across pages. */
  async exportPdf(): Promise<void> {
    if (!this.edits || !this.file) return;
    if (this.grid?.edit) this.grid.commitEdit();
    const s = this.sheet;
    const used = s.used();
    if (!used) return void new Notice("This sheet is empty.");
    const size = { width: 1056, height: 816 };
    const margin = 48;
    const totalW = s.colX(used.c2 + 1) - s.colX(used.c1);
    const scale = Math.min(1, (size.width - 2 * margin) / totalW);
    const pages: HTMLElement[] = [];
    const painter = new Grid(createDiv(), {
      sheet: () => s,
      conditional: (row, col) => this.conditional(s, row, col),
      isStale: () => false,
      isCircular: () => false,
      commit: () => false,
      selectionChanged: () => undefined,
      editing: () => undefined,
      clear: () => undefined,
      fill: () => undefined,
      resize: () => undefined,
      autofit: () => undefined,
      contextMenu: () => undefined,
      filterMenu: () => undefined,
    });
    let start = used.r1;
    while (start <= used.r2) {
      let end = start, h = 0;
      while (end <= used.r2 && (h + s.rowHeight(end)) * scale <= size.height - 2 * margin) h += s.rowHeight(end++);
      if (end === start) end++;
      const page = createDiv({ cls: "dw-print-page xlsx-print-page" });
      const sheetEl = page.createDiv({ cls: "xlsx-print-sheet" });
      sheetEl.setCssProps({ "--xlsx-scale": String(scale), "--xlsx-margin": `${margin}px`, "--xlsx-print-w": `${totalW}px`, "--xlsx-print-h": `${h}px` });
      painter.drawRange(sheetEl, { r1: start, c1: used.c1, r2: end - 1, c2: used.c2 });
      pages.push(page);
      start = end;
    }
    await exportPdf(this.app, this.file, pages, size);
  }

  // ---------------------------------------------------------------------------
  // Saving
  // ---------------------------------------------------------------------------

  private markDirty(): void {
    this.dirty = true;
    this.saveIndicatorEl?.setText("Saving…");
    if (this.saveTimer) window.clearTimeout(this.saveTimer);
    this.saveTimer = window.setTimeout(() => void this.saveNow(), 800);
  }

  private saveNow(): Promise<void> {
    this.saving = this.saving.then(async () => {
      if (!this.edits || !this.file) return;
      if (this.saveTimer) window.clearTimeout(this.saveTimer);
      const file = this.file;
      try {
        this.dirty = false;
        const bytes = await this.edits.wb.pkg.save();
        this.lastWrite = Date.now();
        await this.app.vault.modifyBinary(file, bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
        this.lastWrite = Date.now();
      } catch (err) {
        this.dirty = true;
        console.error("Failed to save .xlsx:", err);
        new Notice(`Failed to save ${file.name}: ${(err as Error).message}`);
      } finally {
        this.saveIndicatorEl?.setText(this.dirty ? "Saving…" : "Saved");
      }
    });
    return this.saving;
  }
}
