/** The Office-style interface's chrome shared by the three editors: ribbon (with File and side buttons), status bar, save status. */
import type { App, FileView, Menu } from "obsidian";
import type { Commands } from "./commands";
import { Ribbon, type Tab } from "./ribbon";
import { SaveStatus } from "./save-status";
import { StatusBar, type StatusItem } from "./status-bar";
import { addFileItems, menuItem } from "./toolbar";
import type { VersionedView } from "./version-dialog";

export interface OfficeUiOptions {
  view: FileView & VersionedView;
  app: App;
  /** Where the ribbon goes (above the editor) and where the status bar goes (appended: below it). */
  host: HTMLElement;
  root: HTMLElement;
  cmds: Commands;
  tabs: Tab[];
  /** Storage key prefix, e.g. "doc-works-rb-pptx". */
  key: string;
  readOnly: boolean;
  qat: { defaults: string[]; choices: string[] };
  /** The app's own items at the top of File, before Obsidian's file items. */
  fileIds: string[];
  sides: [icon: string, title: string, run: (button: HTMLElement) => void][];
  status: StatusItem[];
  zoom?: { get: () => string; set: (z: string) => void; levels: string[] };
}

export function buildOfficeUi(o: OfficeUiOptions): { ribbon: Ribbon; statusBar: StatusBar; save: SaveStatus | null } {
  const ribbon = new Ribbon(o.host, o.cmds, o.tabs, {
    readOnly: o.readOnly,
    app: o.app,
    key: o.key,
    qat: o.qat,
    fileMenu: (m: Menu) => {
      for (const id of o.fileIds.filter((x) => o.cmds[x] && (!o.readOnly || o.cmds[x].view))) {
        const c = o.cmds[id];
        menuItem(m, `${c.label}…`.replace(/……$/, "…"), () => c.run?.(), { icon: c.icon });
      }
      m.addSeparator();
      // Obsidian's own file items (Rename, Delete, Move…), as the Google layout's File menu shows them.
      addFileItems(o.view, m, "more-options", o.readOnly, null);
    },
  });
  for (const [icon, title, run] of o.sides) {
    const b: HTMLElement = ribbon.side(icon, title, () => run(b));
  }
  const statusBar = new StatusBar(o.root, o.status, o.app, o.key.replace("-rb-", "-sb-"));
  if (o.zoom) statusBar.zoom(o.zoom);
  const save = new SaveStatus(statusBar.save);
  if (o.readOnly) statusBar.save.hide();
  return { ribbon, statusBar, save: o.readOnly ? null : save };
}
