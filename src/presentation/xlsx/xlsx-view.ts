import { insertPivot, Pivot, pivotAt, readPivots, refreshPivot } from "../../domain/xlsx/pivot";
import { formulaSparklines, readSparklines, Sparkline } from "../../domain/xlsx/sparklines";
import { addComment, commentsOf, editComment, removeThread, resolveThread, CommentThread } from "../../domain/xlsx/comments";
import { assertAction, hasPassword, isProtected, setProtected, unprotectWithPassword } from "../../domain/xlsx/protection";
import { App, FileView, FuzzySuggestModal, Menu, Modal, Notice, Scope, Setting, TFile, WorkspaceLeaf } from "obsidian";
import { detectSeparator, removeDuplicates, splitSources, splitText } from "../../domain/xlsx/cleanup";
import { importCsv, parseCsv, sheetToCsv } from "../../domain/xlsx/csv";
import { addName, deleteName, listNames, renameName, setNameRange } from "../../domain/xlsx/names";
import { groupRange, setGroupCollapsed, ungroupRange } from "../../domain/xlsx/outline";
import { XlsxWorkbook } from "../../infrastructure/xlsx/xlsx-workbook";
import { Sheet, Workbook } from "../../domain/xlsx/workbook-model";
import { Edits } from "../../domain/xlsx/commands";
import { FormulaSyntaxError } from "../../domain/xlsx/formula";
import { Rect, cellName, colName, parseCellRef, parseRange, quoteSheet as quoteSheetName, rangeName } from "../../domain/xlsx/references";
import { linksOf, setLink } from "../../domain/xlsx/links";
import { adjustDecimals, formatGeneral, parseInput } from "../../domain/xlsx/number-format";
import { BorderStyle, StyleChange } from "../../domain/xlsx/styles";
import { Clip, copyRange, pasteClip, pasteText } from "../../domain/xlsx/clipboard";
import { SheetSearchMatch, findSheetMatches, replaceCells } from "../../domain/xlsx/find";
import { showFindPreview } from "../common/find-preview";
import * as ops from "../../domain/xlsx/sheet-edit";
import { numText } from "../../domain/xlsx/formula-evaluator";
import { SaveStatus } from "../common/save-status";
import { runEditCommand } from "../common/spelling";
import { escapeToEditor, focusRegions, returnFocusAfterOverlays } from "../common/keyboard";

import { FONTS, menuItem, PaperTheme, Toolbar, promptText, showValue, shieldEditingKeys, subMenu, addFileItems } from "../common/toolbar";
import { Note, notesOf, setNote } from "../../domain/xlsx/notes";
import { pagePlan, PrintSettings, printSettings, setPrintSettings } from "../../domain/xlsx/print";
import { DEFAULT_SETTINGS, DocWorksSettings, paperFor } from "../common/settings";
import { canExportPdf, renderPdf } from "../common/pdf";
import { ExportFormat, exportDialog, writeExport } from "../common/export-dialog";
import { Grid, Hit } from "./grid";
import { deletePicture, insertPicture, movePicture, Picture, sheetPictures } from "../../domain/xlsx/drawing";
import { pickImage } from "../common/images";
import { addRule, CfStyle, ConditionalFormats, deleteRule, readRules } from "../../domain/xlsx/conditional";
import { checkValue, invalidateValidations, listItems, removeValidation, setValidation, Validation, validationAt } from "../../domain/xlsx/validation";
import { conditionalForm, pivotForm, validationForm } from "./rule-forms";
import { chartModel, chartSource, deleteChart, formatChart, insertChart, moveChart, setChartSource, SheetChart, sheetCharts } from "../../domain/xlsx/charts";
import type { ChartModel } from "../../domain/pptx/chart";
import { renderChart } from "../pptx/chart-renderer";
import { chartFormatForm, defaultChartFormat } from "../common/chart-form";
import { IMAGE_MIME } from "../../core/utils/images";

export const VIEW_TYPE_XLSX = "doc-works-xlsx-view";
const THEME_KEY = "doc-works-xlsx-theme";
const H_ALIGNS: [string, string][] = [["Left", "left"], ["Center", "center"], ["Right", "right"]];
const V_ALIGNS: [string, string][] = [["Top", "top"], ["Middle", "center"], ["Bottom", "bottom"]];
const FUNCTIONS = ["SUM", "AVERAGE", "COUNT", "MAX", "MIN", "IF", "VLOOKUP", "CONCATENATE"];
const FREEZES: [string, string][] = [["No rows", "r0"], ["1 row", "r1"], ["2 rows", "r2"], ["Up to current row", "rc"], ["No columns", "c0"], ["1 column", "c1"], ["2 columns", "c2"], ["Up to current column", "cc"]];
const ZOOM_KEY = "doc-works-xlsx-zoom";
/** Google Sheets' zoom levels. */
const ZOOMS = ["0.5", "0.75", "0.9", "1", "1.25", "1.5", "2"];
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
  /** Stops giving focus back to the document after dialogs close. */
  private stopFocusReturn: (() => void) | null = null;
  allowNoFile = false;
  private edits: Edits | null = null;
  private sheetIndex = 0;
  private grid: Grid | null = null;
  private theme: PaperTheme;
  private nameBox!: HTMLInputElement;
  private formulaInput!: HTMLTextAreaElement;
  private tabsEl!: HTMLElement;
  private statusEl!: HTMLElement;
  private saveStatus: SaveStatus | null = null;
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
  private sparkCache = new Map<string, Sparkline[]>();
  private invalidCache = new Map<string, boolean>();
  /** Charts per sheet path, with their models, until the next edit. */
  private chartCache = new Map<string, { chart: SheetChart; model: ChartModel }[]>();
  /** Drawn charts by chart path: redrawn only when their data, format or size change. */
  private chartSvgs = new Map<string, { sig: string; el: HTMLElement }>();
  private pictureCache = new Map<string, Picture[]>();

  constructor(leaf: WorkspaceLeaf, private settings: DocWorksSettings = { ...DEFAULT_SETTINGS }) {
    super(leaf);
    // The setting, or (by default) the theme last picked here.
    this.theme = paperFor(settings, this.app.loadLocalStorage(THEME_KEY) === "light" ? "light" : "dark");
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
      [["Mod", "Alt"], "m", () => this.commentDialog(this.sel.r1, this.sel.c1)],
      [["Shift"], "F2", () => void this.editNote(this.sel.r1, this.sel.c1)],
    ];
    for (const [mods, key, run] of keys) {
      this.scope.register(mods, key, () => {
        if (!this.grid || document.activeElement !== this.grid.input || this.grid.edit) return true;
        run();
        return false;
      });
    }
    // The grid's own keys (its input is a textarea) that a user hotkey could take.
    shieldEditingKeys(this.scope, [[["Shift"], "Tab"], [[], "Tab"]]);
    focusRegions(this.scope, this.contentEl, () => {
      // Leaving the grid commits what is being typed, as Sheets does.
      if (this.grid?.edit) this.grid.commitEdit();
      return [this.contentEl.querySelector<HTMLElement>(".dw-menubar"), this.contentEl.querySelector<HTMLElement>(".dw-toolbar"), this.nameBox?.parentElement, this.grid?.input, this.tabsEl, this.findEl];
    });
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
    addFileItems(this, menu, source);
    menu.addItem((i) => i.setTitle("Export…").setIcon("file-output").setSection("action").onClick(() => void this.exportAs()));
    menu.addItem((i) => i.setTitle("Download sheet as CSV…").setIcon("file-spreadsheet").setSection("action").onClick(() => void this.exportAs("csv")));
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

  private painter: { styles: number[][]; sticky: boolean } | null = null;
  private brushBtn: HTMLButtonElement | null = null;

  async onLoadFile(file: TFile): Promise<void> {
    this.stopFocusReturn?.();
    this.stopFocusReturn = returnFocusAfterOverlays(() => this.app.workspace.getActiveViewOfType(XlsxView) === this, () => this.grid?.focusGrid());
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
      conditional: (row, col) => this.withLink(row, col, this.conditional(this.sheet, row, col)),
      openLink: (row, col) => this.openLink(row, col),
      editLink: () => void this.editLink(),
      sparklines: () => this.sparklines(this.sheet),
      dropdown: (row, col) => validationAt(this.sheet, row, col)?.type === "list",
      moveChart: (key, box) => this.change(() => moveChart(this.edits!, key, box)),
      note: (row, col) => this.notes().get(cellName(row, col))?.text ?? null,
      comment: (row, col) => this.comments().get(cellName(row, col))?.some(t => !t.resolved) ?? false,
      invalid: (row, col) => this.invalid(row, col),
      openDropdown: (row, col, at) => this.openDropdown(row, col, at),
      charts: () => this.charts(),
      chartMenu: (key, e) => this.chartMenu(key, e),
      movePicture: (key, box) => this.change(() => movePicture(this.edits!, this.sheet, key, box)),
      pictureMenu: (key, e) => this.pictureMenu(key, e),
      deletePicture: (key) => this.change(() => deletePicture(this.edits!, this.sheet, key)),
      toggleGroup: (axis, g) => this.change(() => setGroupCollapsed(this.edits!, this.sheet, axis, g, !g.collapsed)),
    });
    if (this.savedZoom() !== "1") this.grid.setZoom(Number(this.savedZoom()));
    const bottom = this.contentEl.createDiv({ cls: "xlsx-bottom" });
    this.tabsEl = bottom.createDiv({ cls: "xlsx-tabs" });
    escapeToEditor(this.tabsEl, () => this.grid?.focusGrid());
    this.statusEl = bottom.createDiv({ cls: "xlsx-status" });
    const input = this.grid.input;
    input.addEventListener("copy", (e) => this.onCopy(e, false));
    input.addEventListener("cut", (e) => this.onCopy(e, true));
    input.addEventListener("paste", (e) => this.onPaste(e));
    // Paint format lands on the range selected next, when the mouse is released (as in Google Sheets).
    this.grid.el.addEventListener("mouseup", () => this.paintSelection());
    input.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && this.painter) this.stopPainting();
    });
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
    // A closed file stops retrying (the save above was its last attempt).
    this.saveStatus?.destroy();
    this.stopFocusReturn?.();
    this.stopFocusReturn = null;
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
  /** Tools › Paint format: copy the selected cells' formatting (double-click: keep painting until Escape). */
  private startPainting(sticky: boolean): void {
    if (!this.grid) return;
    const sel = this.grid.selection(), used = this.sheet.used();
    const r2 = Math.min(sel.r2, Math.max(sel.r1, used?.r2 ?? sel.r1)), c2 = Math.min(sel.c2, Math.max(sel.c1, used?.c2 ?? sel.c1));
    const styles = Array.from({ length: r2 - sel.r1 + 1 }, (_, i) => Array.from({ length: c2 - sel.c1 + 1 }, (_, j) => this.sheet.at(sel.r1 + i, sel.c1 + j)?.style ?? 0));
    this.painter = { styles, sticky };
    this.brushBtn?.addClass("is-active");
    this.grid.focusGrid();
  }

  private stopPainting(): void {
    this.painter = null;
    this.brushBtn?.removeClass("is-active");
  }

  private paintSelection(): void {
    const painter = this.painter;
    if (!painter || !this.grid) return;
    const t = this.grid.selection(), h = painter.styles.length, w = painter.styles[0].length;
    this.change(() => this.edits!.restyle(this.sheet, t, (row, col) => painter.styles[(row - t.r1) % h][(col - t.c1) % w]));
    if (!painter.sticky) this.stopPainting();
  }

  private change(fn: () => void): boolean {
    if (!this.edits) return false;
    try {
      fn();
    } catch (err) {
      new Notice(err instanceof FormulaSyntaxError ? `There's a problem with this formula: ${err.message}` : (err as Error).message);
      return false;
    }
    this.cfCache.clear();
    this.sparkCache.clear();
    this.notesCache = null;
    this.commentsCache = null;
    this.linksCache = null;
    this.invalidCache.clear();
    this.chartCache.clear();
    for (const sheet of this.wb.sheets) invalidateValidations(sheet);
    this.grid?.invalidate();
    this.syncSelection();
    this.markDirty();
    return true;
  }

  /** A sheet's sparklines, read once per sheet until the next edit. */
  private sparklines(sheet: Sheet): Sparkline[] {
    let out = this.sparkCache.get(sheet.ref.path);
    if (!out) {
      const calc = this.edits!.calc;
      out = [...readSparklines(sheet), ...formulaSparklines(sheet, (s, r, c, f) => calc.evaluateArray(s, r, c, f))];
      this.sparkCache.set(sheet.ref.path, out);
    }
    return out;
  }

  /** Conditional formatting of a cell, from rules worked out per sheet until the next edit. */
  private linksCache: { sheet: string; links: Map<string, string> } | null = null;
  /** The current sheet's cell links by reference (see domain/xlsx/links). */
  private links(): Map<string, string> {
    const path = this.sheet.ref.path;
    if (!this.linksCache || this.linksCache.sheet !== path) this.linksCache = { sheet: path, links: linksOf(this.edits!.wb.pkg, path) };
    return this.linksCache.links;
  }

  /** A linked cell is drawn as Google Sheets draws links: blue and underlined. */
  private withLink(row: number, col: number, cf: CfStyle | null): CfStyle | null {
    if (!this.links().has(cellName(row, col))) return cf;
    return { ...(cf ?? {}), color: this.theme === "dark" ? "#8AB4F8" : "#1155CC", underline: true };
  }

  /** Ctrl+click: a web address opens in the browser, a place in the workbook is shown. */
  private openLink(row: number, col: number): boolean {
    const link = this.links().get(cellName(row, col));
    if (!link) return false;
    if (!link.startsWith("#")) {
      window.open(link);
      return true;
    }
    const at = link.lastIndexOf("!");
    const name = at < 0 ? null : link.slice(1, at).replace(/^'|'$/g, "").replace(/''/g, "'");
    const index = name === null ? this.sheetIndex : this.wb.sheets.findIndex((s) => s.name === name);
    const target = parseCellRef(link.slice(at + 1).split(":")[0].replace(/\$/g, ""));
    if (index < 0 || !target) return true;
    if (index !== this.sheetIndex) this.showSheet(index);
    this.grid?.select({ row: target.row, col: target.col });
    return true;
  }

  /** Insert › Link (Ctrl+K): the active cell's link; an empty address removes it. An empty cell shows the address. */
  private async editLink(): Promise<void> {
    if (!this.grid || !this.edits) return;
    const { row, col } = this.grid.active;
    const ref = cellName(row, col);
    const url = await promptText(this.app, "Link address (empty removes the link)", this.links().get(ref) ?? "https://");
    if (url === null || !this.edits) return;
    const href = url.trim() && url.trim() !== "https://" ? url.trim() : null;
    const cell = this.sheet.at(row, col);
    if (href && (!cell || cell.value === null || cell.value === "")) this.change(() => this.edits!.input(this.sheet, row, col, href));
    this.change(() => this.edits!.structural(() => setLink(this.edits!.wb.pkg, this.sheet.ref.path, ref, href)));
    this.grid.render();
  }

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
    this.refocusAfter(menu);
    menu.showAtPosition(at);
  }

  // ---------------------------------------------------------------------------
  // Charts
  // ---------------------------------------------------------------------------

  private sheetChartModels(s = this.sheet): { chart: SheetChart; model: ChartModel }[] {
    let list = this.chartCache.get(s.ref.path);
    if (!list) {
      list = sheetCharts(this.wb, s).map((chart) => ({ chart, model: chartModel(this.wb, chart.path) }));
      this.chartCache.set(s.ref.path, list);
    }
    return list;
  }

  private charts(sheet = this.sheet) {
    return this.sheetChartModels(sheet).map(({ chart, model }) => ({
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

  private pictureMenu(key: string, e: MouseEvent): void {
    const menu = new Menu();
    menu.addItem((i) => i.setTitle("Delete image").setIcon("trash").onClick(() => this.change(() => deletePicture(this.edits!, this.sheet, key))));
    this.refocusAfter(menu);
    menu.showAtMouseEvent(e);
  }

  private chartMenu(path: string, e: MouseEvent): void {
    const found = this.sheetChartModels().find((c) => c.chart.path === path);
    if (!found) return;
    const menu = new Menu();
    menu.addItem((i) => i.setTitle("Edit chart").setIcon("pencil").onClick(() => this.chartDialog(path)));
    menu.addItem((i) => i.setTitle("Delete chart").setIcon("trash").onClick(() => this.change(() => deleteChart(this.edits!, this.sheet, found.chart))));
    this.refocusAfter(menu);
    menu.showAtMouseEvent(e);
  }

  /** Google's "Group rows 2 - 5" / "Ungroup rows 2 - 5" on a header menu. */
  private groupItems(menu: Menu, axis: "row" | "col", a: number, b: number): void {
    const label = axis === "row" ? `rows ${a + 1} - ${b + 1}` : `columns ${colName(a)} - ${colName(b)}`;
    menu.addItem((it) => it.setTitle(`Group ${label}`).setIcon("list-tree").onClick(() => this.change(() => groupRange(this.edits!, this.sheet, axis, a, b))));
    const levels = axis === "row" ? this.sheet.rowLevels : this.sheet.colLevels;
    let grouped = false;
    for (const i of levels.keys()) if (i >= a && i <= b) grouped = true;
    if (grouped) menu.addItem((it) => it.setTitle(`Ungroup ${label}`).setIcon("list-x").onClick(() => this.change(() => ungroupRange(this.edits!, this.sheet, axis, a, b))));
  }

  /** The selection, or the used area when a single cell is selected (as Sheets' data tools do). */
  private dataRange(): Rect {
    const sel = this.grid!.selection();
    return sel.r1 === sel.r2 && sel.c1 === sel.c2 ? this.sheet.used() ?? sel : sel;
  }

  /** Data › Data cleanup › Remove duplicates. */
  private removeDuplicatesDialog(): void {
    if (!this.edits || !this.grid) return;
    const r = this.dataRange();
    const modal = new Modal(this.app);
    modal.titleEl.setText("Remove duplicates");
    modal.contentEl.createEl("p", { text: `Data range: ${this.sheet.name}!${rangeName(r)}` });
    let header = false;
    const picked = new Set<number>();
    const list = modal.contentEl.createDiv();
    const drawColumns = () => {
      list.empty();
      for (let col = r.c1; col <= r.c2; col++) {
        const title = header && this.sheet.display(r.r1, col).text ? `${this.sheet.display(r.r1, col).text} (${colName(col)})` : `Column ${colName(col)}`;
        new Setting(list).setName(title).addToggle((t) => t.setValue(picked.has(col)).onChange((on: boolean) => (on ? picked.add(col) : picked.delete(col))));
      }
    };
    for (let col = r.c1; col <= r.c2; col++) picked.add(col);
    new Setting(modal.contentEl).setName("Data has header row").addToggle((t) =>
      t.setValue(header).onChange((v: boolean) => {
        header = v;
        drawColumns();
      })
    );
    modal.contentEl.append(list);
    drawColumns();
    new Setting(modal.contentEl)
      .addButton((b) => b.setButtonText("Remove duplicates").setCta().onClick(() => {
        if (!picked.size) return void new Notice("Pick at least one column to analyze.");
        let result = { removed: 0, kept: 0 };
        if (!this.change(() => (result = removeDuplicates(this.edits!, this.sheet, r, { header, cols: [...picked].sort((x, y) => x - y) })))) return;
        modal.close();
        const rows = (n: number, word: string) => `${n} ${word} ${n === 1 ? "row" : "rows"}`;
        new Notice(result.removed ? `${rows(result.removed, "duplicate")} found and removed. ${rows(result.kept, "unique")} remain.` : "No duplicate rows found.");
      }))
      .addButton((b) => b.setButtonText("Cancel").onClick(() => modal.close()));
    modal.open();
  }

  /** Data › Split text to columns, with Sheets' separator choices. */
  private splitTextDialog(): void {
    if (!this.edits || !this.grid) return;
    const sel = this.grid.selection();
    if (sel.c1 !== sel.c2) return void new Notice("Select one column to split.");
    const r = sel.r1 === sel.r2 ? { ...sel, r2: this.sheet.used()?.r2 ?? sel.r2 } : sel;
    const texts = splitSources(this.sheet, r);
    const modal = new Modal(this.app);
    modal.titleEl.setText("Split text to columns");
    const detected = detectSeparator(texts) ?? ",";
    let choice = "auto", custom = "";
    const NAMES: Record<string, string> = { ",": "Comma", ";": "Semicolon", ".": "Period", " ": "Space" };
    new Setting(modal.contentEl).setName("Separator").addDropdown((d) =>
      d.addOptions({ auto: `Detect automatically (${NAMES[detected]})`, ",": "Comma", ";": "Semicolon", ".": "Period", " ": "Space", custom: "Custom" }).setValue(choice).onChange((v: string) => (choice = v))
    );
    new Setting(modal.contentEl).setName("Custom separator").addText((t) => t.setValue(custom).onChange((v: string) => (custom = v)));
    const sep = () => (choice === "auto" ? detected : choice === "custom" ? custom : choice);
    const width = (s: string) => Math.max(0, ...texts.map((t) => t.split(s).length));
    const filledRight = () => {
      const w = width(sep());
      return w > 1 && this.sheet.cellsIn({ r1: r.r1, r2: r.r2, c1: r.c1 + 1, c2: r.c1 + w - 1 }).some((c) => c.value !== null || c.formula !== undefined);
    };
    modal.contentEl.createEl("p", { cls: "pptx-props-hint", text: "Parts go into the columns to the right, replacing what is there." });
    new Setting(modal.contentEl)
      .addButton((b) => b.setButtonText("Split").setCta().onClick(() => {
        if (!sep()) return void new Notice("Enter a separator.");
        if (filledRight()) new Notice("Existing data to the right was replaced.");
        if (this.change(() => splitText(this.edits!, this.sheet, r, sep()))) modal.close();
      }))
      .addButton((b) => b.setButtonText("Cancel").onClick(() => modal.close()));
    modal.open();
  }

  /** Data › Named ranges: list, add, edit, rename and delete. */
  private namedRangesDialog(): void {
    if (!this.edits || !this.grid) return;
    const modal = new Modal(this.app);
    modal.titleEl.setText("Named ranges");
    modal.modalEl.addClass("xlsx-names-modal");
    const draw = () => {
      modal.contentEl.empty();
      for (const n of listNames(this.wb)) {
        let name = n.name, range = n.range;
        new Setting(modal.contentEl)
          .addText((t) => {
            t.inputEl.setAttribute("aria-label", "Name");
            t.setValue(name).onChange((v: string) => (name = v.trim()));
          })
          .addText((t) => {
            t.inputEl.setAttribute("aria-label", "Range");
            t.setValue(range).onChange((v: string) => (range = v.trim()));
          })
          .addButton((b) => b.setButtonText("Save").onClick(() => {
            const ok = this.change(() => {
              if (range !== n.range) setNameRange(this.edits!, n.name, range);
              if (name !== n.name) renameName(this.edits!, n.name, name);
            });
            if (ok) draw();
          }))
          .addButton((b) => b.setButtonText("Delete").onClick(() => {
            if (this.change(() => deleteName(this.edits!, n.name))) draw();
          }));
      }
      let name = "", range = `${quoteSheetName(this.sheet.name)}!${rangeName(this.grid!.selection())}`;
      new Setting(modal.contentEl).setName("Add a range").setHeading();
      new Setting(modal.contentEl)
        .addText((t) => {
          t.inputEl.setAttribute("aria-label", "New name");
          t.setPlaceholder("Name").onChange((v: string) => (name = v.trim()));
        })
        .addText((t) => {
          t.inputEl.setAttribute("aria-label", "New range");
          t.setValue(range).onChange((v: string) => (range = v.trim()));
        })
        .addButton((b) => b.setButtonText("Add").setCta().onClick(() => {
          if (this.change(() => addName(this.edits!, name, range))) draw();
        }));
    };
    draw();
    modal.open();
  }

  /** File › Import: pick a .csv in the vault for a new sheet. */
  private pickCsv(): void {
    const files = this.app.vault.getFiles().filter((f) => f.extension.toLowerCase() === "csv");
    if (!files.length) return void new Notice("No .csv files in this vault.");
    new FilePicker(this.app, files, (f) => void this.importCsvFile(f)).open();
  }

  private async importCsvFile(file: TFile): Promise<void> {
    if (!this.edits) return;
    const rows = parseCsv(await this.app.vault.read(file));
    let name = "";
    if (!this.change(() => (name = importCsv(this.edits!, rows, file.basename, this.sheetIndex)))) return;
    this.showSheet(this.wb.sheets.findIndex((s) => s.name === name));
  }

  /** Google's Insert › Pivot table: a new sheet with a pivot table over the selection, then its editor. */
  private insertPivotTable(): void {
    if (!this.edits || !this.grid) return;
    let sel = this.grid.selection();
    if (sel.r1 === sel.r2 && sel.c1 === sel.c2) sel = this.sheet.used() ?? sel;
    let name = "";
    if (!this.change(() => (name = insertPivot(this.edits!, this.sheet, sel)))) return;
    const index = this.wb.sheets.findIndex((s) => s.name === name);
    this.showSheet(index);
    const pivot = readPivots(this.wb.sheets[index])[0];
    if (pivot) this.pivotDialog(pivot);
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

  /** Insert › Image: the picture at its own size (at most 400 px wide), over the cells at the active cell. */
  private async insertImageFile(file: File): Promise<void> {
    if (!this.grid || !this.edits) return;
    const bytes = new Uint8Array(await file.arrayBuffer());
    const ext = (/\.([a-z0-9]+)$/i.exec(file.name)?.[1] ?? /^image\/([a-z]+)/.exec(file.type)?.[1] ?? "png").toLowerCase();
    let size = { w: 300, h: 200 };
    try {
      const bitmap = await createImageBitmap(file);
      const scale = Math.min(1, 400 / bitmap.width);
      size = { w: Math.round(bitmap.width * scale), h: Math.round(bitmap.height * scale) };
      bitmap.close();
    } catch {
      // Not decodable here (or no createImageBitmap): Google's default box.
    }
    const at = this.grid.active;
    this.change(() => insertPicture(this.edits!, this.sheet, bytes, ext, at, size, file.name.replace(/\.[^.]+$/, "")));
    this.pictureCache.clear();
    this.grid.render();
  }

  private pictures(s = this.sheet): Picture[] {
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
    this.sparkCache.clear();
    this.notesCache = null;
    this.commentsCache = null;
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
    const text = e.clipboardData.getData("text/plain");
    // A copied picture or screenshot: Google places it over the cells at the active cell.
    const image = Array.from(e.clipboardData.files ?? []).find((f) => f.type.startsWith("image/"));
    if (image && !text.trim()) void this.insertImageFile(image);
    else this.pasteFrom(text, valuesOnly);
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

  /** The zoom picked last (100% by default). */
  private savedZoom(): string {
    const z = this.app.loadLocalStorage(ZOOM_KEY) as string | null;
    return z && ZOOMS.includes(z) ? z : "1";
  }

  // ---------------------------------------------------------------------------
  // Notes
  // ---------------------------------------------------------------------------

  private notesCache: { sheet: string; notes: Map<string, Note> } | null = null;
  /** The current sheet's notes (read again after every change). */
  private notes(): Map<string, Note> {
    const path = this.sheet.ref.path;
    if (!this.notesCache || this.notesCache.sheet !== path) this.notesCache = { sheet: path, notes: notesOf(this.edits!.wb.pkg, path) };
    return this.notesCache.notes;
  }

  private commentsCache: { sheet: string; threads: Map<string, CommentThread[]> } | null = null;
  private comments(): Map<string, CommentThread[]> {
    const path = this.sheet.ref.path;
    if (!this.commentsCache || this.commentsCache.sheet !== path) this.commentsCache = { sheet: path, threads: commentsOf(this.wb.pkg, path) };
    return this.commentsCache.threads;
  }

  /** Discussions stay attached to the cell while the user replies, edits or resolves a thread. */
  private commentDialog(row: number, col: number, hit?: { messageId?: string; start: number; end: number }): void {
    if (!this.edits) return;
    const edits = this.edits, path = this.sheet.ref.path, ref = cellName(row, col);
    const sheet = () => edits.wb.sheets.find(s => s.ref.path === path)!;
    const modal = new Modal(this.app);
    modal.titleEl.setText(`Comments on ${ref}`);
    modal.contentEl.addClass("xlsx-comments-modal");
    const change = (fn: () => void) => {
      if (this.edits !== edits || !sheet()) return;
      if (this.change(fn)) draw();
    };
    const button = (parent: HTMLElement, text: string, fn: () => void) => {
      const b = parent.createEl("button", { text }); b.addEventListener("click", fn); return b;
    };
    const compose = (parent: HTMLElement, label: string, post: (text: string) => void, value = "") => {
      const input = parent.createEl("textarea", { value, attr: { "aria-label": label, rows: "3", placeholder: label } });
      button(parent, value ? "Save" : "Post", () => { if (input.value.trim()) post(input.value); });
    };
    const draw = () => {
      modal.contentEl.empty();
      const threads = commentsOf(edits.wb.pkg, path).get(ref) ?? [];
      if (!threads.length) modal.contentEl.createEl("p", { text: "No comments on this cell." });
      for (const thread of threads) {
        const box = modal.contentEl.createDiv({ cls: "xlsx-comment-thread" });
        if (thread.resolved) box.createDiv({ text: "Resolved", cls: "dw-muted" });
        for (const message of thread.messages) {
          const item = box.createDiv({ cls: "xlsx-comment-message" });
          item.createEl("strong", { text: message.author });
          if (message.date) item.createEl("time", { text: new Date(message.date).toLocaleString(), attr: { datetime: message.date } });
          const textEl = item.createDiv({ cls: "xlsx-comment-text" });
          if (hit?.messageId === message.id) {
            textEl.appendText(message.text.slice(0, hit.start));
            textEl.createEl("mark", { text: message.text.slice(hit.start, hit.end) });
            textEl.appendText(message.text.slice(hit.end));
          } else textEl.setText(message.text);
          button(item, "Edit", () => {
            item.empty(); compose(item, "Edit comment", text => change(() => editComment(edits, sheet(), message.id, text)), message.text);
            button(item, "Cancel", draw);
          });
        }
        const actions = box.createDiv({ cls: "xlsx-comment-actions" });
        button(actions, thread.resolved ? "Reopen" : "Resolve", () => change(() => resolveThread(edits, sheet(), thread.id, !thread.resolved)));
        button(actions, "Delete thread", () => change(() => removeThread(edits, sheet(), thread.id)));
        if (!thread.resolved) compose(box, "Reply", text => change(() => addComment(edits, sheet(), ref, text, this.author(), thread.id)));
      }
      compose(modal.contentEl, "New comment", text => change(() => addComment(edits, sheet(), ref, text, this.author())));
      button(modal.contentEl, "Close", () => modal.close());
    };
    draw(); modal.open();
    modal.contentEl.querySelector("mark")?.scrollIntoView({ block: "nearest" });
  }

  private setNoteText(row: number, col: number, text: string | null): void {
    const path = this.sheet.ref.path;
    this.change(() => this.edits!.structural(() => { assertAction(this.sheet, "objects"); setNote(this.edits!.wb.pkg, path, cellName(row, col), text, this.author()); }));
    this.notesCache = null;
    this.commentsCache = null;
    this.grid?.invalidate();
  }

  /** Google Sheets' Insert › Note (Shift+F2): the cell's note in a box to edit. */
  private async editNote(row: number, col: number): Promise<void> {
    const current = this.notes().get(cellName(row, col))?.text ?? "";
    const text = await promptNote(this.app, `Note on ${cellName(row, col)}`, current);
    if (text === null || text === current) return;
    this.setNoteText(row, col, text.trim() ? text : null);
  }

  /** The name notes are written with: the comment author setting, else "Doc Works". */
  private author(): string {
    return this.settings.author.trim() || "Doc Works";
  }

  /**
   * Cut / copy / paste from a menu: the app runs the edit command on the grid's input, which raises the same
   * clipboard events as the keyboard shortcuts (no direct clipboard reads or writes).
   */
  private editCommand(command: "copy" | "cut" | "paste", valuesOnly = false): void {
    if (!this.grid) return;
    this.valuesOnly = valuesOnly;
    this.grid.input.focus();
    // ponytail: no Electron on mobile; the browser's own copy / cut command raises the same clipboard events.
    if (!runEditCommand(command)) new Notice("Paste with your keyboard or the system menu in the cell editor.");
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
    const zoomTo = (z: string) => {
      zoom.value = z;
      this.grid?.setZoom(Number(z));
      this.app.saveLocalStorage(ZOOM_KEY, z);
    };

    // --- The row: everyday controls, as Google Sheets' toolbar ---
    tb.icon("undo-2", "Undo (Ctrl+Z)", () => this.undo());
    tb.icon("redo-2", "Redo (Ctrl+Y)", () => this.redo());
    const brush = (this.brushBtn = tb.icon("paintbrush", "Paint format (double-click to keep painting)", () => this.startPainting(false)));
    brush.addEventListener("dblclick", () => this.startPainting(true));
    tb.pin();
    tb.gap();
    const zoom = tb.select("Zoom", "Zoom", ZOOMS.map((z) => [`${Math.round(Number(z) * 100)}%`, z]), zoomTo);
    zoom.value = this.savedZoom();
    tb.gap();
    tb.button("$", "Format as currency", () => this.format({ numFmt: '"$"#,##0.00' }));
    tb.button("%", "Format as percent", () => this.format({ numFmt: "0%" }));
    tb.button(".0←", "Decrease decimal places", () => this.decimals(-1));
    tb.button(".00→", "Increase decimal places", () => this.decimals(1));
    this.formatSelect = tb.select("Number format", "123", NUMBER_FORMATS, (code) => this.format({ numFmt: code }));
    tb.gap();
    this.fontSelect = tb.select("Font", "Font", FONTS.map((f) => [f, f]), (fontName) => this.format({ fontName }));
    this.sizeSelect = tb.select("Font size", "Size", FONT_SIZES.map((s) => [String(s), String(s)]), (s) => this.format({ fontSize: Number(s) }));
    tb.pin();
    tb.gap();
    this.toggles.set("bold", tb.button("B", "Bold (Ctrl+B)", () => this.toggle("bold"), "bold"));
    this.toggles.set("italic", tb.button("I", "Italic (Ctrl+I)", () => this.toggle("italic"), "italic"));
    this.toggles.set("strike", tb.button("S", "Strikethrough (Ctrl+Shift+X)", () => this.toggle("strike"), "strike"));
    this.toggles.set("underline", tb.button("U", "Underline (Ctrl+U)", () => this.toggle("underline"), "underline"));
    tb.color("A", "Text color", "#000000", (color) => this.format({ color }));
    tb.gap();
    tb.color("▇", "Fill color", "#FFFF00", (fill) => this.format({ fill }));
    tb.button("No fill", "Remove fill color", () => this.format({ fill: null }));
    tb.iconSelect("square-dashed", "Borders", [["All borders", "all"], ["Outer borders", "outer"], ["Top", "top"], ["Bottom", "bottom"], ["Left", "left"], ["Right", "right"], ["No borders", "none"]], (v) =>
      this.borders(v as "all")
    );
    tb.icon("merge", "Merge cells", () => this.mergeToggle());
    tb.gap();
    tb.iconSelect("align-left", "Horizontal align", H_ALIGNS, (v) => this.format({ hAlign: v as "left" }));
    tb.iconSelect("align-vertical-justify-center", "Vertical align", V_ALIGNS, (v) => this.format({ vAlign: v as "top" }));
    this.toggles.set("wrap", tb.icon("wrap-text", "Wrap text", () => this.toggle("wrap")));
    tb.gap();
    tb.icon("link", "Insert link (Ctrl+K)", () => void this.editLink());
    tb.icon("message-square-plus", "Insert comment", () => this.commentDialog(this.sel.r1, this.sel.c1));
    tb.icon("chart-column", "Insert chart", () => this.insertChart());
    tb.icon("filter", "Create or remove a filter", () => this.toggleFilter());
    tb.iconSelect("sigma", "Functions", FUNCTIONS.map((f) => [f, f]), (f) => this.insertFunction(f));

    // --- The menus: every command ---
    const menu = (name: string, build: (m: Menu) => void) =>
      tb.menu(name, (m) => {
        build(m);
        this.refocusAfter(m);
      });
    tb.fileMenu(this, (m) => {
      menuItem(m, "Export…", () => void this.exportAs());
      menuItem(m, "Download sheet as CSV…", () => void this.exportAs("csv"));
    });
    menu("Edit", (m) => {
      menuItem(m, "Undo", () => this.undo());
      menuItem(m, "Redo", () => this.redo());
      m.addSeparator();
      menuItem(m, "Cut", () => this.editCommand("cut"));
      menuItem(m, "Copy", () => this.editCommand("copy"));
      menuItem(m, "Paste", () => this.editCommand("paste"));
      menuItem(m, "Paste values only", () => this.editCommand("paste", true));
      menuItem(m, "Select all", () => this.grid?.selectAll());
      menuItem(m, "Delete", () => this.grid && this.change(() => this.edits!.clear(this.sheet, this.grid!.selection())));
      m.addSeparator();
      menuItem(m, "Find and replace", () => this.openFind());
    });
    menu("View", (m) => {
      subMenu(m, "Freeze", (s) => {
        for (const [label, v] of FREEZES) menuItem(s, label, () => this.freeze(v));
      });
      subMenu(m, "Zoom", (s) => {
        for (const z of ZOOMS) menuItem(s, `${Math.round(Number(z) * 100)}%`, () => zoomTo(z), { checked: zoom.value === z });
      });
      m.addSeparator();
      menuItem(m, "Light", () => this.setTheme("light"), { checked: this.theme === "light" });
      menuItem(m, "Dark", () => this.setTheme("dark"), { checked: this.theme === "dark" });
    });
    menu("Insert", (m) => {
      menuItem(m, "Sheet", () => {
        this.change(() => ops.addSheet(this.edits!, this.sheetIndex));
        this.showSheet(this.sheetIndex + 1);
      });
      menuItem(m, "Image", () => pickImage((f) => void this.insertImageFile(f)));
      menuItem(m, "Chart", () => this.insertChart());
      menuItem(m, "Pivot table", () => this.insertPivotTable());
      subMenu(m, "Function", (s) => {
        for (const f of FUNCTIONS) menuItem(s, f, () => this.insertFunction(f));
      });
      m.addSeparator();
      menuItem(m, "Link", () => void this.editLink());
      menuItem(m, "Comment", () => this.commentDialog(this.sel.r1, this.sel.c1));
      menuItem(m, "Note", () => void this.editNote(this.sel.r1, this.sel.c1));
      menuItem(m, "Dropdown", () => this.validationDialog());
    });
    menu("Format", (m) => {
      subMenu(m, "Number", (s) => {
        for (const [label, code] of NUMBER_FORMATS) menuItem(s, label, () => this.format({ numFmt: code }));
      });
      subMenu(m, "Text", (s) => {
        for (const [label, key] of [["Bold", "bold"], ["Italic", "italic"], ["Underline", "underline"], ["Strikethrough", "strike"]] as const) menuItem(s, label, () => this.toggle(key));
      });
      subMenu(m, "Alignment", (s) => {
        for (const [label, h] of H_ALIGNS) menuItem(s, label, () => this.format({ hAlign: h as "left" }));
        s.addSeparator();
        for (const [label, v] of V_ALIGNS) menuItem(s, label, () => this.format({ vAlign: v as "top" }));
      });
      menuItem(m, "Wrap text", () => this.toggle("wrap"));
      menuItem(m, "Merge cells", () => this.mergeToggle());
      m.addSeparator();
      menuItem(m, "Conditional formatting", () => this.conditionalDialog());
      m.addSeparator();
      menuItem(m, "Clear formatting", () => this.grid && this.change(() => this.edits!.restyle(this.sheet, this.grid!.selection(), () => 0)));
    });
    menu("Data", (m) => {
      const r = this.sel;
      menuItem(m, "Sort A → Z", () => this.sort(false));
      menuItem(m, "Sort Z → A", () => this.sort(true));
      menuItem(m, this.sheet.filter ? "Remove filter" : "Create a filter", () => this.toggleFilter());
      menuItem(m, "Data validation", () => this.validationDialog());
      m.addSeparator();
      menuItem(m, "Remove duplicates…", () => this.removeDuplicatesDialog());
      menuItem(m, "Split text to columns…", () => this.splitTextDialog());
      menuItem(m, "Named ranges…", () => this.namedRangesDialog());
      m.addSeparator();
      this.groupItems(m, "row", r.r1, r.r2);
      this.groupItems(m, "col", r.c1, r.c2);
      m.addSeparator();
      menuItem(m, "Import CSV as new sheet…", () => this.pickCsv());
    });
    menu("Tools", (m) => {
      menuItem(m, "Paint format", () => this.startPainting(false));
      menuItem(m, "Find and replace", () => this.openFind());
    });

    this.saveStatus = tb.end().status;
  }

  private setTheme(theme: PaperTheme): void {
    this.theme = theme;
    this.app.saveLocalStorage(THEME_KEY, theme);
    this.contentEl.toggleClass("xlsx-theme-dark", theme === "dark");
    this.contentEl.toggleClass("xlsx-theme-light", theme === "light");
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
    for (const [k, el] of this.toggles) {
      el.toggleClass("is-active", !!f[k as "bold"]);
      el.setAttribute("aria-pressed", String(!!f[k as "bold"]));
    }
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
      const tab = bar.createEl("button", { cls: i === this.sheetIndex ? "xlsx-tab is-active" : "xlsx-tab", text: s.name, attr: { "aria-label": `Sheet ${s.name}`, "aria-current": i === this.sheetIndex ? "true" : "false" } });
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
        this.refocusAfter(menu);
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

  private async unprotect(sheet: Sheet): Promise<void> {
    const pw = await promptText(this.app, "Sheet password", "", true);
    if (pw === null || !this.edits) return;
    if (await unprotectWithPassword(this.edits, sheet, pw)) this.change(() => {});
    else new Notice("The password is incorrect.");
  }

  private sheetMenu(i: number): Menu {
    const e = this.edits!;
    const s = () => this.wb.sheets[i];
    const menu = new Menu();
    menu.addItem((it) => it.setTitle(isProtected(s()) ? "Unprotect sheet" : "Protect sheet").setIcon("lock").onClick(() => {
      if (isProtected(s()) && hasPassword(s())) void this.unprotect(s());
      else this.change(() => setProtected(e, s(), !isProtected(s())));
    }));
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

  /** Google's pivot table editor, as a dialog; Done works the table out again with the new fields. */
  private pivotDialog(p: Pivot): void {
    const modal = new Modal(this.app);
    modal.titleEl.setText("Pivot table editor");
    pivotForm(modal.contentEl, {
      fields: p.fields,
      layout: { rows: p.rows, cols: p.cols, values: p.values },
      apply: (layout) => {
        if (this.change(() => refreshPivot(this.edits!, this.sheet, p, layout))) modal.close();
      },
    });
    modal.open();
  }

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
    if (hit.area === "cell") {
      if (!isProtected(this.sheet)) {
        item("Unlock cells", "lock-open", () => this.format({ locked: false }));
        item("Lock cells", "lock", () => this.format({ locked: true }));
      }
      item(this.comments().has(cellName(r.r1, r.c1)) ? "View comments" : "Insert comment", "message-square", () => this.commentDialog(r.r1, r.c1));
      const note = this.notes().get(cellName(r.r1, r.c1));
      item(note ? "Edit note" : "Insert note", "sticky-note", () => void this.editNote(r.r1, r.c1));
      if (note) item("Delete note", "trash", () => this.setNoteText(r.r1, r.c1, null));
    }
    const pivot = hit.area === "cell" ? pivotAt(this.sheet, r.r1, r.c1) : null;
    if (pivot) {
      menu.addSeparator();
      item("Edit pivot table", "table-2", () => this.pivotDialog(pivot));
      item("Refresh pivot table", "refresh-cw", () => this.change(() => refreshPivot(ed, this.sheet, pivot)));
    }
    menu.addSeparator();
    if (hit.area === "rowHeader") {
      item(`Hide ${rows}`, "eye-off", () => this.change(() => ops.setHidden(ed, this.sheet, "row", r.r1, r.r2, true)));
      item("Unhide rows", "eye", () => this.change(() => ops.setHidden(ed, this.sheet, "row", Math.max(0, r.r1 - 1), r.r2 + 1, false)));
      item("Resize rows…", "move-vertical", () => void this.askSize("row", r.r1, r.r2));
      this.groupItems(menu, "row", r.r1, r.r2);
    } else if (hit.area === "colHeader") {
      item(`Hide ${cols}`, "eye-off", () => this.change(() => ops.setHidden(ed, this.sheet, "col", r.c1, r.c2, true)));
      item("Unhide columns", "eye", () => this.change(() => ops.setHidden(ed, this.sheet, "col", Math.max(0, r.c1 - 1), r.c2 + 1, false)));
      item("Resize columns…", "move-horizontal", () => void this.askSize("col", r.c1, r.c2));
      this.groupItems(menu, "col", r.c1, r.c2);
    } else {
      item("Sort range ascending", "arrow-down-az", () => this.sort(false));
      item("Sort range descending", "arrow-up-za", () => this.sort(true));
      item(this.sheet.merges.some((m) => m.r1 <= r.r2 && r.r1 <= m.r2 && m.c1 <= r.c2 && r.c1 <= m.c2) ? "Unmerge cells" : "Merge cells", "merge", () => this.mergeToggle());
    }
    item("Freeze up to here", "snowflake", () => this.change(() => ops.freeze(ed, this.sheet, hit.area === "colHeader" ? 0 : r.r2 + 1, hit.area === "rowHeader" ? 0 : r.c2 + 1)));
    this.refocusAfter(menu);
    menu.showAtMouseEvent(e);
  }

  /** Once a menu closes, the grid takes the keys again (unless what the menu did opened a dialog or field). */
  private refocusAfter(menu: Menu): void {
    menu.onHide(() =>
      window.setTimeout(() => {
        const a = document.activeElement;
        if (!a || a === document.body || this.contentEl.contains(a)) this.grid?.focusGrid();
      })
    );
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
      const box = label.createEl("input", { type: "checkbox", attr: { "aria-label": title } });
      label.appendText(text);
      return box;
    };
    const matchCase = check("Aa", "Match case");
    const entire = check("Whole cell", "Match entire cell contents");
    const formulas = check("Formulas", "Also search within formulas");
    const allSheets = check("All sheets", "Search all sheets");
    const notes = check("Notes", "Search cell notes");
    const comments = check("Comments", "Search comments");
    notes.checked = comments.checked = true;
    const count = bar.createSpan({ cls: "pptx-find-count" });
    const preview = bar.createDiv({ cls: "dw-find-preview", attr: { "aria-live": "polite" } });
    let matches: SheetSearchMatch[] = [];
    let index = -1;
    const opts = () => ({ matchCase: matchCase.checked, entire: entire.checked, formulas: formulas.checked, notes: notes.checked, comments: comments.checked });
    const refresh = () => {
      matches = findSheetMatches(allSheets.checked ? this.wb.sheets : [this.sheet], input.value, opts());
      index = Math.min(index, matches.length - 1);
      count.setText(input.value ? (matches.length ? `${Math.max(index, 0) + 1} of ${matches.length}` : "No results") : "");
      preview.empty();
      replaceButton.disabled = matches[Math.max(index, 0)]?.kind !== "cell";
      replaceAllButton.disabled = !matches.some(m => m.kind === "cell");
    };
    const show = (m: SheetSearchMatch) => {
      const i = this.wb.sheets.indexOf(m.sheet);
      if (i !== this.sheetIndex) this.showSheet(i);
      this.grid!.select({ row: m.row, col: m.col });
      this.grid!.reveal({ row: m.row, col: m.col });
      input.focus({ preventScroll: true });
      replaceButton.disabled = m.kind !== "cell";
      if (m.kind !== "cell") showFindPreview(preview, `${m.sheet.name}!${cellName(m.row, m.col)} · ${m.kind === "note" ? "Note" : m.resolved ? "Resolved comment" : "Comment"} · ${m.author}`, m.text, m.start, m.end,
        m.kind === "note" ? "Open note" : "Open comment", () => {
          const sheetIndex = this.wb.sheets.indexOf(m.sheet);
          if (sheetIndex !== this.sheetIndex) this.showSheet(sheetIndex);
          if (m.kind === "note") void this.editNote(m.row, m.col);
          else this.commentDialog(m.row, m.col, m);
        });
    };
    const step = (dir: number) => {
      refresh();
      if (!matches.length) return;
      index = (index + dir + matches.length) % matches.length;
      count.setText(`${index + 1} of ${matches.length}`);
      show(matches[index]);
    };
    const button = (text: string, title: string, fn: () => void) => {
      const b = bar.createEl("button", { text, cls: "dw-btn", attr: { "aria-label": title } });
      b.addEventListener("click", fn);
      return b;
    };
    button("↑", "Previous match (Shift+Enter)", () => step(-1));
    button("↓", "Next match (Enter)", () => step(1));
    const replaceButton = button("Replace", "Replace this cell-content match", () => {
      refresh();
      const m = matches[index];
      if (!m || m.kind !== "cell") return;
      this.change(() => replaceCells(this.edits!, [m], input.value, replace.value, opts()));
      step(0);
    });
    const replaceAllButton = button("Replace all", "Replace matching cell contents", () => {
      refresh();
      let n = 0;
      this.change(() => (n = replaceCells(this.edits!, matches.filter(m => m.kind === "cell"), input.value, replace.value, opts())));
      new Notice(`Replaced ${n} ${n === 1 ? "cell" : "cells"}.`);
      refresh();
    });
    button("×", "Close (Esc)", () => this.closeFind());
    bar.append(preview);
    input.addEventListener("input", () => {
      index = -1;
      refresh();
      if (matches.length) { index = 0; show(matches[0]); }
    });
    for (const box of [matchCase, entire, formulas, allSheets, notes, comments]) box.addEventListener("change", () => { refresh(); if (matches.length) { index = Math.max(index, 0); show(matches[index]); } });
    bar.addEventListener("keydown", (e) => {
      if (e.key === "Escape") this.closeFind();
      else if (e.key === "Enter" && e.target === input) { e.preventDefault(); step(e.shiftKey ? -1 : 1); }
    });
    refresh();
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
  /** Google Sheets' print screen: what to print, paper, orientation, scale, margins and gridlines; then the PDF. */
  /** File › Download, as Google Sheets: PDF (with the print settings) or the current sheet as CSV. */
  async exportAs(initial?: "pdf" | "csv"): Promise<void> {
    const file = this.file;
    if (!this.edits || !file) return;
    if (this.grid?.edit) this.grid.commitEdit();
    const print = { settings: { ...printSettings(this.sheet), margins: { ...printSettings(this.sheet).margins } }, what: "sheet" as PrintWhat };
    const hasSelection = this.sel.r1 !== this.sel.r2 || this.sel.c1 !== this.sel.c2;
    const formats: ExportFormat[] = [{ id: "csv", label: "Comma-separated values, current sheet (.csv)", ext: "csv", outside: true, name: `${file.basename} - ${this.sheet.name}` }];
    if (canExportPdf()) formats.unshift({ id: "pdf", label: "PDF document (.pdf)", ext: "pdf", outside: true, options: (el) => printOptions(el, print, hasSelection) });
    const choice = await exportDialog(this.app, file, formats, initial);
    if (!choice) return;
    if (choice.format === "csv") return void (await writeExport(this.app, choice.target, sheetToCsv(this.sheet), choice.open));
    const data = await this.sheetPdf(print.settings, print.what);
    if (data) await writeExport(this.app, choice.target, data, choice.open);
  }

  /** The sheet, selection or workbook printed to PDF with `settings` (kept on the sheet, as in Excel). */
  private async sheetPdf(settings: PrintSettings, what: PrintWhat): Promise<Uint8Array | null> {
    if (!this.edits) return null;
    // The settings belong to the sheet (as in Excel): keep them when changed.
    if (JSON.stringify(settings) !== JSON.stringify(printSettings(this.sheet))) this.change(() => setPrintSettings(this.edits!, this.sheet, settings));
    const jobs: { sheet: Sheet; range: Rect }[] = [];
    if (what === "selection") jobs.push({ sheet: this.sheet, range: this.sel });
    else for (const sheet of what === "workbook" ? this.edits.wb.sheets.filter((s) => s.ref.state === "visible") : [this.sheet]) {
      let used = sheet.used();
      for (const box of [...this.pictures(sheet), ...this.charts(sheet)]) {
        const end = { r1: sheet.rowAt(box.y), c1: sheet.colAt(box.x), r2: sheet.rowAt(box.y + box.h - 0.01), c2: sheet.colAt(box.x + box.w - 0.01) };
        used = used ? { r1: Math.min(used.r1, end.r1), c1: Math.min(used.c1, end.c1), r2: Math.max(used.r2, end.r2), c2: Math.max(used.c2, end.c2) } : end;
      }
      if (used) jobs.push({ sheet, range: used });
    }
    if (!jobs.length) {
      new Notice("There's nothing to print.");
      return null;
    }
    const pages: HTMLElement[] = [];
    let size = { width: 816, height: 1056 };
    for (const { sheet, range } of jobs) {
      const ps = what === "workbook" && sheet !== this.sheet ? printSettings(sheet) : settings;
      const cols: number[] = [], rows: number[] = [];
      for (let c = range.c1; c <= range.c2; c++) cols.push(sheet.colWidth(c));
      for (let r = range.r1; r <= range.r2; r++) rows.push(sheet.rowHeight(r));
      const plan = pagePlan(cols, rows, ps);
      // ponytail: one paper size per PDF (the first sheet's); sheets on other paper are scaled onto it.
      if (!pages.length) size = { width: plan.width, height: plan.height };
      const pictures = this.pictures(sheet);
      await Promise.all(pictures.map(async (p) => {
        if (this.imageUrls.get(p.image)) return;
        try {
          const bytes = await this.wb.pkg.media(p.image);
          const ext = p.image.slice(p.image.lastIndexOf(".") + 1).toLowerCase();
          this.imageUrls.set(p.image, URL.createObjectURL(new Blob([bytes as BlobPart], { type: IMAGE_MIME[ext] ?? `image/${ext}` })));
        } catch { /* Keep unavailable images blank, as the grid does. */ }
      }));
      const painter = new Grid(createDiv(), {
        sheet: () => sheet,
        pictures: () => pictures,
        imageUrl: (path) => this.imageUrls.get(path) ?? null,
        charts: () => this.charts(sheet),
        conditional: (row, col) => this.conditional(sheet, row, col),
        sparklines: () => this.sparklines(sheet),
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
      painter.printGridlines = ps.gridlines;
      for (const p of plan.pages) {
        const r: Rect = { r1: range.r1 + p.r1, c1: range.c1 + p.c1, r2: range.r1 + p.r2, c2: range.c1 + p.c2 };
        const page = createDiv({ cls: "dw-print-page xlsx-print-page" });
        const sheetEl = page.createDiv({ cls: "xlsx-print-sheet" });
        sheetEl.setCssProps({
          "--xlsx-scale": String(plan.scale),
          "--xlsx-margin-top": `${ps.margins.top * 96}px`,
          "--xlsx-margin-left": `${ps.margins.left * 96}px`,
          "--xlsx-print-w": `${sheet.colX(r.c2 + 1) - sheet.colX(r.c1)}px`,
          "--xlsx-print-h": `${sheet.rowY(r.r2 + 1) - sheet.rowY(r.r1)}px`,
        });
        painter.drawRange(sheetEl, r);
        pages.push(page);
      }
    }
    return renderPdf(pages, size);
  }

  // ---------------------------------------------------------------------------
  // Saving
  // ---------------------------------------------------------------------------

  private markDirty(): void {
    this.dirty = true;
    this.saveStatus?.saving();
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
        if (this.dirty) this.saveStatus?.saving();
        else this.saveStatus?.saved();
      } catch (err) {
        this.dirty = true;
        console.error("Failed to save .xlsx:", err);
        this.saveStatus?.failed(err, () => void this.saveNow(), file.name);
      }
    });
    return this.saving;
  }
}

/** Pick one of `files` by fuzzy search. */
class FilePicker extends FuzzySuggestModal<TFile> {
  constructor(app: App, private files: TFile[], private choose: (f: TFile) => void) {
    super(app);
  }
  getItems(): TFile[] {
    return this.files;
  }
  getItemText(f: TFile): string {
    return f.path;
  }
  onChooseItem(f: TFile): void {
    this.choose(f);
  }
}

/** A box for a note's text (several lines); resolves null when cancelled. */
function promptNote(app: App, title: string, value: string): Promise<string | null> {
  return new Promise((resolve) => {
    let result: string | null = null;
    const modal = new Modal(app);
    modal.titleEl.setText(title);
    const area = modal.contentEl.createEl("textarea", { cls: "xlsx-note-input", attr: { rows: "5", "aria-label": "Note" } });
    area.value = value;
    // Ctrl+Enter saves, as in Sheets; Enter is a new line.
    area.addEventListener("keydown", (e) => {
      if (e.key !== "Enter" || !(e.ctrlKey || e.metaKey)) return;
      e.preventDefault();
      result = area.value;
      modal.close();
    });
    new Setting(modal.contentEl)
      .addButton((b) => b.setButtonText("Save").setCta().onClick(() => {
        result = area.value;
        modal.close();
      }))
      .addButton((b) => b.setButtonText("Cancel").onClick(() => modal.close()));
    modal.onClose = () => resolve(result);
    modal.open();
    window.setTimeout(() => area.focus());
  });
}

const MARGINS: Record<string, PrintSettings["margins"]> = {
  normal: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7 },
  narrow: { top: 0.75, bottom: 0.75, left: 0.25, right: 0.25 },
  wide: { top: 1, bottom: 1, left: 1, right: 1 },
};

type PrintWhat = "sheet" | "selection" | "workbook";

/** The PDF options of the export dialog: what to print, paper, orientation, scale, margins and gridlines. */
function printOptions(el: HTMLElement, state: { settings: PrintSettings; what: PrintWhat }, hasSelection: boolean): void {
  const ps = state.settings;
  const dropdown = (name: string, options: Record<string, string>, value: string, set: (v: string) => void) =>
    new Setting(el).setName(name).addDropdown((d) => d.addOptions(options).setValue(value).onChange(set));
  dropdown("Print", { sheet: "Current sheet", ...(hasSelection ? { selection: "Selected cells" } : {}), workbook: "Workbook" }, state.what, (v) => (state.what = v as PrintWhat));
  dropdown("Paper size", { letter: "Letter (8.5\" × 11\")", a4: "A4 (8.27\" × 11.69\")", legal: "Legal (8.5\" × 14\")" }, ps.paper, (v) => (ps.paper = v as PrintSettings["paper"]));
  dropdown("Page orientation", { portrait: "Portrait", landscape: "Landscape" }, ps.landscape ? "landscape" : "portrait", (v) => (ps.landscape = v === "landscape"));
  dropdown("Scale", { normal: "Normal", width: "Fit to width", page: "Fit to page" }, ps.fit, (v) => (ps.fit = v as PrintSettings["fit"]));
  new Setting(el).setName("Scale (%)").setDesc("Used when the scale is normal.").addText((t) => {
    t.inputEl.type = "number";
    t.setValue(String(ps.scale)).onChange((v) => (ps.scale = Math.min(400, Math.max(10, Math.round(Number(v) || 100)))));
  });
  const preset = Object.keys(MARGINS).find((k) => JSON.stringify(MARGINS[k]) === JSON.stringify(ps.margins)) ?? "custom";
  dropdown("Margins", { normal: "Normal", narrow: "Narrow", wide: "Wide", ...(preset === "custom" ? { custom: "As saved" } : {}) }, preset, (v) => {
    if (MARGINS[v]) ps.margins = { ...MARGINS[v] };
  });
  new Setting(el).setName("Show gridlines").addToggle((t) => t.setValue(ps.gridlines).onChange((v) => (ps.gridlines = v)));
}
