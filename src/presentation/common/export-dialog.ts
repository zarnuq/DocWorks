/*
 * One export dialog for every format, as Google's File › Download: pick the format (with its own options), the file
 * name and the vault folder (or, for PDF on desktop, a place outside the vault), see whether it replaces a file,
 * and choose whether to open the result. `writeExport` then saves it with the same feedback everywhere.
 */
import { App, FuzzySuggestModal, Modal, normalizePath, Notice, Setting, TFile, TFolder } from "obsidian";
import { canExportPdf } from "./pdf";

export interface ExportFormat {
  id: string;
  /** Shown in the format list, e.g. "PDF document (.pdf)". */
  label: string;
  ext: string;
  /** Suggested file name (without extension); default: the source's name. */
  name?: string;
  /** May be saved outside the vault (a single file the system save dialog can take). */
  outside?: boolean;
  /** Draw this format's options; they write into the caller's own state. */
  options?: (el: HTMLElement) => void;
}

export type ExportTarget = { vault: string } | { outside: string };

export interface ExportChoice {
  format: string;
  target: ExportTarget;
  open: boolean;
}

const OPEN_KEY = "doc-works-export-open";
const folderOf = (path: string) => (path.includes("/") ? path.slice(0, path.lastIndexOf("/")) : "");
const join = (folder: string, name: string) => normalizePath(folder ? `${folder}/${name}` : name);

class FolderPicker extends FuzzySuggestModal<TFolder> {
  constructor(app: App, private choose: (f: TFolder) => void) {
    super(app);
    this.setPlaceholder("Choose a folder");
  }
  getItems(): TFolder[] {
    return this.app.vault.getAllFolders(true);
  }
  getItemText(f: TFolder): string {
    return f.isRoot() ? "/" : f.path;
  }
  onChooseItem(f: TFolder): void {
    this.choose(f);
  }
}

/** Ask where and how to export `source`; resolves null when cancelled. */
export function exportDialog(app: App, source: TFile, formats: ExportFormat[], initial?: string): Promise<ExportChoice | null> {
  return new Promise((resolve) => {
    let result: ExportChoice | null = null;
    let format = formats.find((f) => f.id === initial) ?? formats[0];
    let name = format.name ?? source.basename;
    let nameInput: HTMLInputElement | null = null;
    let folder = folderOf(source.path);
    let outside: string | null = null;
    // Exports used to open their result: on until turned off.
    let open = app.loadLocalStorage(OPEN_KEY) !== "off";
    const modal = new Modal(app);
    modal.titleEl.setText("Export");
    const body = modal.contentEl;

    if (formats.length > 1) {
      new Setting(body).setName("Format").addDropdown((d) => {
        d.selectEl.setAttribute("aria-label", "Format");
        d.addOptions(Object.fromEntries(formats.map((f) => [f.id, f.label]))).setValue(format.id).onChange((v: string) => {
          const before = format;
          format = formats.find((f) => f.id === v) ?? format;
          if (!format.outside) outside = null;
          // A name the user hasn't changed follows the format's suggestion.
          if (name === (before.name ?? source.basename) && nameInput) nameInput.value = name = format.name ?? source.basename;
          drawOptions();
          drawWhere();
        });
      });
    }
    const optionsEl = body.createDiv({ cls: "dw-export-options" });
    new Setting(body).setName("File name").addText((t) => {
      nameInput = t.inputEl;
      t.inputEl.setAttribute("aria-label", "File name");
      t.setValue(name);
      t.inputEl.addEventListener("input", () => {
        name = t.inputEl.value.trim();
        drawWhere();
      });
    });
    const whereEl = body.createDiv();
    new Setting(body).setName("Open after export").addToggle((t) =>
      t.setValue(open).onChange((v: boolean) => {
        open = v;
        app.saveLocalStorage(OPEN_KEY, v ? "on" : "off");
      })
    );
    const buttons = new Setting(body);
    let exportBtn: HTMLButtonElement | null = null;
    buttons
      .addButton((b) => {
        exportBtn = (b as unknown as { buttonEl?: HTMLButtonElement }).buttonEl ?? null;
        b.setButtonText("Export").setCta().onClick(() => {
          if (!name) return void new Notice("Enter a file name.");
          result = { format: format.id, target: outside ? { outside: `${name}.${format.ext}` } : { vault: join(folder, `${name}.${format.ext}`) }, open: open && !outside };
          modal.close();
        });
      })
      .addButton((b) => b.setButtonText("Cancel").onClick(() => modal.close()));

    const drawOptions = () => {
      optionsEl.empty();
      format.options?.(optionsEl);
    };
    const drawWhere = () => {
      whereEl.empty();
      const path = join(folder, `${name}.${format.ext}`);
      const s = new Setting(whereEl).setName("Save to");
      whereEl.createDiv({ cls: "dw-export-path", text: outside ? "Outside the vault (you'll choose where)" : path });
      s.addButton((b) => b.setButtonText("Choose folder…").onClick(() =>
        new FolderPicker(app, (f) => {
          folder = f.isRoot() ? "" : f.path;
          outside = null;
          drawWhere();
        }).open()
      ));
      if (format.outside && canExportPdf()) {
        // The app's download asks where to save, with the system dialog.
        s.addButton((b) => b.setButtonText("Outside the vault…").onClick(() => {
          outside = `${name}.${format.ext}`;
          drawWhere();
        }));
      }
      const exists = !outside && !!app.vault.getAbstractFileByPath(path);
      if (exists) whereEl.createDiv({ cls: "dw-export-warning", text: `A file named “${name}.${format.ext}” already exists there. Exporting replaces it.` });
      const label = exists ? "Replace" : "Export";
      const btn = exportBtn ?? Array.from(buttons.settingEl.querySelectorAll("button"))[0];
      if (btn) btn.textContent = label;
    };
    drawOptions();
    drawWhere();
    modal.onClose = () => resolve(result);
    modal.open();
  });
}

/**
 * Save an export: a vault file is created or replaced (then opened if asked); a place outside the vault goes
 * through the app's own download. Shows the outcome; resolves the vault path written, or null.
 */
export async function writeExport(app: App, target: ExportTarget, data: Uint8Array | string, open: boolean): Promise<string | null> {
  try {
    if ("outside" in target) {
      const bytes = typeof data === "string" ? new TextEncoder().encode(data) : data;
      const url = URL.createObjectURL(new Blob([bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer]));
      const link = document.body.createEl("a", { href: url, attr: { download: target.outside.split(/[\\/]/).pop() ?? "export" } });
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
      new Notice("Choose where to save the file");
      return null;
    }
    const path = target.vault;
    const existing = app.vault.getAbstractFileByPath(path);
    let file: TFile;
    if (typeof data === "string") {
      if (existing instanceof TFile) await app.vault.modify((file = existing), data);
      else file = await app.vault.create(path, data);
    } else {
      const buf = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer;
      if (existing instanceof TFile) await app.vault.modifyBinary((file = existing), buf);
      else file = await app.vault.createBinary(path, buf);
    }
    new Notice(`Exported to ${path}`);
    if (open) await app.workspace.getLeaf("tab").openFile(file);
    return path;
  } catch (err) {
    new Notice(`Couldn't export: ${err instanceof Error ? err.message : String(err)}`);
    return null;
  }
}
