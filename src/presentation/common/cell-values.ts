/*
 * Live cell values in notes: inline code `cell: Budget.xlsx#Budget!H9` shows that cell's value as the sheet shows it,
 * and follows the file when it's saved. Reading view through a post processor, Live Preview through an editor
 * extension (the code itself shows while the cursor is on it).
 */
import { App, editorInfoField, MarkdownRenderChild, parseLinktext, Plugin, TFile } from "obsidian";
import { Prec, RangeSetBuilder } from "@codemirror/state";
import { Decoration, DecorationSet, EditorView, ViewPlugin, ViewUpdate, WidgetType } from "@codemirror/view";
import { Workbook } from "../../domain/xlsx/workbook-model";
import { XlsxWorkbook } from "../../infrastructure/xlsx/xlsx-workbook";
import { sheetTarget } from "./embeds";

export interface CellRef {
  path: string;
  subpath: string;
}

/** `cell: Budget.xlsx#Budget!H9` (the link may be in [[ ]]); null for other code. */
export function parseCellCode(text: string): CellRef | null {
  const m = /^cell:\s*(.+)$/.exec(text.trim());
  if (!m) return null;
  const { path, subpath } = parseLinktext(m[1].trim().replace(/^\[\[|\]\]$/g, ""));
  return path && subpath ? { path, subpath } : null;
}

/** The shown value of the cell (the top left one of a range); null when the sheet or cell isn't named right. */
export function cellValue(wb: Workbook, subpath: string): string | null {
  const t = sheetTarget(subpath, wb.sheets.map((s) => s.name));
  if (!t.range) return null;
  const sheet = t.sheet === null ? wb.sheets.find((s) => s.ref.state === "visible") ?? wb.sheets[0] : wb.sheets.find((s) => s.name === t.sheet);
  return sheet ? sheet.display(t.range.r1, t.range.c1).text : null;
}

/** Workbooks read for cell values, until the file changes. */
const books = new Map<string, { mtime: number; wb: Promise<Workbook> }>();

function workbook(app: App, file: TFile): Promise<Workbook> {
  const cached = books.get(file.path);
  if (cached && cached.mtime === file.stat.mtime) return cached.wb;
  const wb = app.vault.readBinary(file).then(async (bytes) => Workbook.open(await XlsxWorkbook.load(bytes)));
  books.set(file.path, { mtime: file.stat.mtime, wb });
  wb.catch(() => books.delete(file.path));
  return wb;
}

/** Put the cell's value in `el` (or why there's none, with `is-error`). */
export async function renderCellValue(app: App, el: HTMLElement, ref: CellRef, sourcePath: string): Promise<void> {
  const file = app.metadataCache.getFirstLinkpathDest(ref.path, sourcePath);
  let text: string | null = null, error = "";
  if (!file || file.extension.toLowerCase() !== "xlsx") error = `${ref.path} not found`;
  else {
    try {
      text = cellValue(await workbook(app, file), ref.subpath);
      if (text === null) error = `No cell ${ref.subpath.slice(1)} in ${file.name}`;
    } catch {
      error = `Couldn't read ${file.name}`;
    }
  }
  el.setText(error || text || "");
  el.toggleClass("is-error", !!error);
}

class CellValueChild extends MarkdownRenderChild {
  constructor(el: HTMLElement, private app: App, private ref: CellRef, private sourcePath: string) {
    super(el);
  }

  onload(): void {
    void renderCellValue(this.app, this.containerEl, this.ref, this.sourcePath);
    this.registerEvent(this.app.vault.on("modify", (f) => {
      if (f.path === this.app.metadataCache.getFirstLinkpathDest(this.ref.path, this.sourcePath)?.path) void renderCellValue(this.app, this.containerEl, this.ref, this.sourcePath);
    }));
  }
}

class CellWidget extends WidgetType {
  constructor(private app: App, private ref: CellRef, private source: string, private code: string, private version: number) {
    super();
  }

  eq(other: CellWidget): boolean {
    return other.code === this.code && other.version === this.version && other.source === this.source;
  }

  toDOM(): HTMLElement {
    const el = createSpan({ cls: "dw-cell-value", attr: { title: this.code } });
    void renderCellValue(this.app, el, this.ref, this.source);
    return el;
  }
}

/** Bumped when any workbook changes, so Live Preview draws its values again. */
let version = 0;
const CODE = /`(cell:[^`\n]+)`/g;

function cellValuesPlugin(app: App) {
  return ViewPlugin.fromClass(
    class {
      decorations: DecorationSet;
      private version = version;

      constructor(view: EditorView) {
        this.decorations = this.build(view);
      }

      update(u: ViewUpdate): void {
        if (u.docChanged || u.viewportChanged || u.selectionSet || this.version !== version) {
          this.version = version;
          this.decorations = this.build(u.view);
        }
      }

      build(view: EditorView): DecorationSet {
        const builder = new RangeSetBuilder<Decoration>();
        const source = view.state.field(editorInfoField, false)?.file?.path ?? "";
        const sel = view.state.selection.main;
        for (const { from, to } of view.visibleRanges) {
          const text = view.state.doc.sliceString(from, to);
          for (const m of text.matchAll(CODE)) {
            const start = from + m.index, end = start + m[0].length;
            const ref = parseCellCode(m[1]);
            // The cursor on it shows the code to edit.
            if (!ref || (sel.from <= end && sel.to >= start)) continue;
            builder.add(start, end, Decoration.replace({ widget: new CellWidget(app, ref, source, m[1], version) }));
          }
        }
        return builder.finish();
      }
    },
    { decorations: (v) => v.decorations }
  );
}

export function registerCellValues(plugin: Plugin): void {
  const app = plugin.app;
  plugin.registerMarkdownPostProcessor((el, ctx) => {
    for (const code of Array.from(el.querySelectorAll("code"))) {
      if (code.closest("pre")) continue;
      const ref = parseCellCode(code.textContent ?? "");
      if (!ref) continue;
      const span = createSpan({ cls: "dw-cell-value", attr: { title: code.textContent ?? "" } });
      code.replaceWith(span);
      ctx.addChild(new CellValueChild(span, app, ref, ctx.sourcePath));
    }
  });
  // Highest precedence: plugins that style inline code (Code Styler) replace the same text.
  plugin.registerEditorExtension(Prec.highest(cellValuesPlugin(app)));
  // A saved workbook: Live Preview draws its cells again (reading view children follow the file themselves).
  plugin.registerEvent(app.vault.on("modify", (f) => {
    if (!(f instanceof TFile) || f.extension.toLowerCase() !== "xlsx") return;
    version++;
    app.workspace.iterateAllLeaves((leaf) => (leaf.view as { editor?: { cm?: EditorView } }).editor?.cm?.dispatch({}));
  }));
}
