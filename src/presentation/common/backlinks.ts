/*
 * Backlinks into Office files: the notes that link or embed a file, by place (slide, cell range, heading), from
 * Obsidian's link index. Each editor lists them in a "Linked from" drawer and marks the linked places.
 */
import { App, type FileView, parseLinktext, setIcon, TFile } from "obsidian";
import { openDrawer } from "./narrow";

export interface Backlink {
  source: TFile;
  /** "#3", "#Budget!B4", "#Heading" or "" for the whole file. */
  subpath: string;
  /** 0-based line of the link in the note. */
  line: number;
  /** That line's text. */
  text: string;
}

export interface BacklinkGroup {
  subpath: string;
  label: string;
  links: Backlink[];
}

type Kind = "docx" | "pptx" | "xlsx";
type LinkCache = { link: string; position?: { start: { line: number } } };

/** Every link and embed of `file` in the vault's notes. */
export async function backlinksTo(app: App, file: TFile): Promise<Backlink[]> {
  const out: Backlink[] = [];
  for (const [path, targets] of Object.entries(app.metadataCache.resolvedLinks)) {
    if (!targets[file.path]) continue;
    const source = app.vault.getAbstractFileByPath(path);
    if (!(source instanceof TFile)) continue;
    const cache = app.metadataCache.getFileCache(source);
    const refs = [...((cache?.links ?? []) as LinkCache[]), ...((cache?.embeds ?? []) as LinkCache[])]
      .map((l) => ({ ...parseLinktext(l.link), line: l.position?.start.line ?? 0 }))
      .filter((l) => app.metadataCache.getFirstLinkpathDest(l.path, path) === file)
      .sort((a, b) => a.line - b.line);
    if (!refs.length) continue;
    const lines = (await app.vault.cachedRead(source)).split("\n");
    for (const r of refs) out.push({ source, subpath: r.subpath, line: r.line, text: (lines[r.line] ?? "").trim() });
  }
  return out.sort((a, b) => a.source.path.localeCompare(b.source.path) || a.line - b.line);
}

/** A place as the editor names it: "Slide 3", "Budget!B4:C8", "Results › Totals", "Whole file". */
export function placeLabel(kind: Kind, subpath: string): string {
  if (!subpath) return "Whole file";
  if (kind === "pptx" && /^#\d+$/.test(subpath)) return `Slide ${subpath.slice(1)}`;
  return subpath.slice(1).split("#").join(" › ");
}

/** Backlinks by place: the whole file first, then places in order (slides by number). */
export function groupBacklinks(kind: Kind, links: Backlink[]): BacklinkGroup[] {
  const groups = new Map<string, BacklinkGroup>();
  for (const l of links) {
    const g = groups.get(l.subpath) ?? { subpath: l.subpath, label: placeLabel(kind, l.subpath), links: [] };
    g.links.push(l);
    groups.set(l.subpath, g);
  }
  return [...groups.values()].sort((a, b) => (a.subpath === "" ? -1 : b.subpath === "" ? 1 : a.label.localeCompare(b.label, undefined, { numeric: true })));
}

/** The "Linked from" list: a place jumps there, a note opens at the link's line. */
export function renderBacklinks(panel: HTMLElement, groups: BacklinkGroup[], on: { goTo: (subpath: string) => void; open: (link: Backlink) => void }): void {
  panel.empty();
  panel.addClass("dw-backlinks");
  if (!groups.length) return void panel.createDiv({ cls: "dw-muted", text: "No notes link to this file yet." });
  for (const g of groups) {
    const place = panel.createEl("button", { cls: "dw-backlinks-place", text: g.label });
    if (g.subpath) place.addEventListener("click", () => on.goTo(g.subpath));
    else place.disabled = true;
    for (const l of g.links) {
      const note = panel.createDiv({ cls: "dw-backlinks-note", attr: { role: "button", tabindex: "0", title: l.text } });
      setIcon(note.createSpan({ cls: "dw-backlinks-icon" }), "file-text");
      note.createSpan({ cls: "dw-backlinks-name", text: l.source.basename });
      note.createSpan({ cls: "dw-backlinks-text", text: l.text });
      note.addEventListener("click", () => on.open(l));
      note.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          on.open(l);
        }
      });
    }
  }
}

/**
 * A view's backlinks: kept current from the link index (`changed` redraws the view's markers) and shown in a
 * "Linked from" drawer at the right edge.
 */
export class LinkedFrom {
  groups: BacklinkGroup[] = [];
  private panel: HTMLElement | null = null;
  private close: (() => void) | null = null;

  constructor(private view: FileView, private kind: Kind, private changed: () => void) {
    view.registerEvent(view.app.metadataCache.on("resolved", () => void this.refresh()));
  }

  /** How many links point at `subpath` ("#3", "#Heading"...). */
  count(subpath: string): number {
    return this.groups.find((g) => g.subpath === subpath)?.links.length ?? 0;
  }

  async refresh(): Promise<void> {
    const file = this.view.file;
    const groups = file ? groupBacklinks(this.kind, await backlinksTo(this.view.app, file)) : [];
    if (file !== this.view.file) return;
    this.groups = groups;
    if (this.panel) this.render(this.panel);
    this.changed();
  }

  /** Open the drawer, or close it when open. */
  toggle(button?: HTMLElement | null): void {
    if (this.close) return this.close();
    const root = this.view.contentEl;
    const panel = (this.panel = root.createDiv({ cls: "dw-backlinks-drawer", attr: { tabindex: "-1", "aria-label": "Linked from" } }));
    this.close = openDrawer(panel, {
      root,
      toggle: button,
      onClose: () => {
        panel.remove();
        this.panel = this.close = null;
      },
    });
    this.render(panel);
    panel.focus();
    void this.refresh();
  }

  private render(panel: HTMLElement): void {
    const list = createDiv();
    renderBacklinks(list, this.groups, {
      goTo: (subpath) => {
        this.close?.();
        this.view.setEphemeralState({ subpath });
      },
      open: (l) => void this.view.app.workspace.getLeaf("tab").openFile(l.source, { eState: { line: l.line } }),
    });
    panel.empty();
    const title = panel.createDiv({ cls: "dw-backlinks-title", text: "Linked from" });
    const x = title.createEl("button", { cls: "clickable-icon", attr: { "aria-label": "Close" } });
    setIcon(x, "x");
    x.addEventListener("click", () => this.close?.());
    panel.appendChild(list);
  }
}
