# Word-style interface (beta) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A per-file-type beta setting that gives Word documents desktop Word's ribbon, Quick Access Toolbar, right-click menu + mini toolbar, status bar and shortcuts, over the commands Doc Works already has.

**Architecture:** Generic pieces in `src/presentation/common/` (ribbon, quick access toolbar, own context menu, status bar), all driven by a per-view command table (`Commands`). Word's layout (tabs, groups, menus, shortcuts) is data in `src/presentation/docx/word-ui.ts`; `docx-view.ts` builds the command table from its existing methods and picks Google's `Toolbar` or the `Ribbon` from the setting.

**Tech Stack:** TypeScript, Obsidian API (Menu, setIcon, Scope, settings definitions), ProseMirror keymaps, vitest + jsdom with `tests/obsidian-stub.ts`.

**Spec:** `docs/superpowers/specs/2026-11-09-office-ribbon-design.md` (reference outlines: `docs/office-ribbon/word.md`).

## Global Constraints

- Default stays the Google layout; the toggle is labelled "Word-style interface (beta)", default off.
- Only `main.ts`, `*-view.ts` and `presentation/common/` import `obsidian`; `word-ui.ts` is plain data/types.
- styles.css: no `!important`, no `text-indent`, no multi-column properties. Colours/spacing from Obsidian CSS variables.
- View preferences via `app.loadLocalStorage` / `app.saveLocalStorage`, never `window.localStorage`.
- No transactions on load; opening a file or switching interface never saves it.
- Shortcuts that clash with app hotkeys go through the view's `Scope`.
- No new runtime dependencies, no network. Icons are Lucide names.
- Every task ends green: `npm test`, `npm run typecheck`, `npm run lint` (0 errors).
- Owner rule: never run git; the owner commits. Steps below end with "verify", not "commit".

## Review Focus

1. Read-only files with the Word interface: only look-only commands appear (ribbon, QAT, menu) and nothing saves. Test in Task 8.
2. Narrow pane / phone: ribbon folds without overlapping and every command stays reachable through a collapsed group. Test in Task 3.
3. A QAT list stored with ids that no longer exist (or from the other app) loads without errors and drops them. Test in Task 4.
4. Switching the toggle with the file open (reopenViews) keeps the document, doesn't mark it dirty, and leaves no stale listeners (window keydown / resize). Test in Task 8.
5. Right-click menu near the window's right/bottom edge stays on screen; Escape / outside click / scroll closes both menu and mini toolbar. Test in Task 5.

---

### Task 1: Settings sections, per-type paper, Word interface toggle

**Files:**
- Modify: `src/presentation/common/settings.ts`, `src/main.ts` (settings tab, onload), `src/presentation/docx/docx-view.ts:402`, `src/presentation/pptx/pptx-view.ts:469`, `src/presentation/xlsx/xlsx-view.ts:134`
- Test: `tests/settings.test.ts` (new)

**Interfaces:**
- Produces: `type FileKind = "docx" | "pptx" | "xlsx"`; `DocWorksSettings` gains `paperDocx`, `paperPptx`, `paperXlsx: PaperSetting` and `officeDocx: boolean` (default false), loses `paper`; `migrateSettings(raw: Record<string, unknown> | null): DocWorksSettings`; `paperFor(settings, kind: FileKind, viewDefault: "light" | "dark"): "light" | "dark"`.

- [ ] Step 1: failing tests in `tests/settings.test.ts`:
  - `migrateSettings({ paper: "dark" })` → `paperDocx/paperPptx/paperXlsx` all `"dark"`, no `paper` key, `officeDocx === false`.
  - `migrateSettings({ paper: "default" })` and `migrateSettings(null)` → all three `"default"`.
  - `migrateSettings({ paperDocx: "light", paper: "dark" })` keeps `paperDocx: "light"` (new keys win).
  - `paperFor({ ...DEFAULT_SETTINGS, paperXlsx: "light" }, "xlsx", "dark")` → `"light"`; with `"default"` → the `viewDefault`.
- [ ] Step 2: run `npx vitest run tests/settings.test.ts` → FAIL (no export).
- [ ] Step 3: implement in `settings.ts`; `main.ts` onload uses `migrateSettings(await this.loadData())`; views call `paperFor(settings, "<kind>", <same default as now>)`.
- [ ] Step 4: settings tab: `getSettingDefinitions` returns groups (`SettingDefinitionGroup`) General (Read only, Version history, Keep versions for, Comment author), Word documents (Open .docx, Paper colour, Word-style interface (beta), Word page size), PowerPoint presentations (Open .pptx, Paper colour), Excel spreadsheets (Open .xlsx, Paper colour). `display()` uses `new Setting(el).setName(..).setHeading()` per section. Paper options per type: `default` (label "Light" for Word / "Dark" for others, i.e. today's behaviour), `light`, `dark`, `obsidian`. Changing `officeDocx` calls `reopenViews()` like Read only. Toggle desc: "Desktop Word's ribbon, Quick Access Toolbar, status bar, right-click menu and shortcuts instead of the Google Docs layout."
- [ ] Step 5: `npm test && npm run typecheck && npm run lint` → all pass.

### Task 2: Command table and ribbon core

**Files:**
- Create: `src/presentation/common/commands.ts`, `src/presentation/common/ribbon.ts`
- Modify: `styles.css` (ribbon section)
- Test: `tests/ribbon.test.ts` (new)

**Interfaces:**
- Produces (`commands.ts`):
  ```ts
  export interface Cmd {
    label: string;            // Word's caption
    icon: string;             // Lucide
    run?: () => void;
    menu?: (m: MenuLike) => void;   // dropdown / split part
    active?: () => boolean;   // toggle state
    enabled?: () => boolean;
    view?: boolean;           // kept for read-only files
    shortcut?: string;        // tooltip text, e.g. "Ctrl+B"
  }
  export type Commands = Record<string, Cmd>;
  export interface MenuLike { item(label: string, run: () => void, opts?: { checked?: boolean; disabled?: boolean; icon?: string }): void; separator(): void; sub(label: string, build: (m: MenuLike) => void): void }
  export function obsidianMenu(m: Menu): MenuLike; // adapter over menuItem/subMenu from toolbar.ts
  ```
- Produces (`ribbon.ts`):
  ```ts
  export type Control =
    | { kind: "large" | "small" | "icon"; id: string }          // button; split when the Cmd has both run and menu
    | { kind: "select"; id: string; options: () => [string, string][]; value: () => string | null; set: (v: string) => void; width?: number }
    | { kind: "color"; id: string; initial: string; set: (hex: string) => void }
    | { kind: "check"; id: string }
    | { kind: "gallery"; id: string; items: () => [label: string, value: string][]; value: () => string | null; set: (v: string) => void }
    | { kind: "column"; items: Control[] };                       // stack of up to 3 small controls
  export interface Group { id: string; label: string; icon: string; launcher?: string; controls: Control[] }
  export interface Tab { id: string; label: string; groups: Group[]; context?: () => boolean }
  export class Ribbon {
    constructor(el: HTMLElement, cmds: Commands, tabs: Tab[], opts: { readOnly: boolean; app: AppStorage; key: string; fileMenu: (m: Menu) => void });
    select(tabId: string): void; readonly current: string;
    sync(): void;   // toggles, values, contextual tabs (a contextual tab that appears is selected, as Word does; when it goes, back to the last core tab)
    fit(): void;    // Task 3
    side(icon: string, title: string, onClick: () => void): HTMLButtonElement;   // right end of the tab row
    readonly tabRow: HTMLElement; readonly panel: HTMLElement;
    overflows: () => boolean;  // replaced in tests
  }
  ```
  `AppStorage` = `{ loadLocalStorage(k: string): unknown; saveLocalStorage(k: string, v: unknown): void }`.

- [ ] Step 1: failing tests in `tests/ribbon.test.ts` (plain DOM, stub `Menu`):
  - renders tab buttons in order with `role="tab"`, File first; clicking "Insert" shows only Insert's groups (`.dw-rb-group` captions).
  - a `large` control shows icon + caption; `small` caption; `icon` only `aria-label`; clicking runs `cmd.run`.
  - a Cmd with `run` and `menu` renders a split button: main part runs, ▾ opens `Menu.last` with the menu's items.
  - `active()` true → `is-active` + `aria-pressed="true"` after `sync()`.
  - contextual tab with `context: () => flag`: hidden while false; after flag = true and `sync()`, shown, marked `.is-contextual`, and selected; after false + `sync()`, previous tab selected again.
  - `readOnly: true` drops controls whose Cmd lacks `view`, then empty groups, then empty tabs.
  - group with `launcher: "ParagraphDialog"` renders a corner button that runs that Cmd.
  - double-clicking the selected tab toggles `.is-collapsed` on the ribbon and saves `"collapsed"` under `opts.key + "-collapsed"`; a new Ribbon with the same storage starts collapsed.
- [ ] Step 2: run `npx vitest run tests/ribbon.test.ts` → FAIL.
- [ ] Step 3: implement. Keyboard: tab row uses `rovingToolbar` semantics with `role="tablist"`; panel is `rovingToolbar`. Buttons `mousedown` preventDefault (keep editor selection). CSS classes prefixed `dw-rb-`; heights: tab row 28 px, panel 92 px (large button 66 px tall), group caption 11 px `var(--text-muted)`, separators `var(--background-modifier-border)`, contextual tab label `var(--text-accent)`.
- [ ] Step 4: tests pass; full suite, typecheck, lint pass.

### Task 3: Ribbon folding (Word-style)

**Files:** Modify `src/presentation/common/ribbon.ts`, `styles.css`; Test `tests/ribbon.test.ts`

**Interfaces:** Produces `Ribbon.fit()`; groups get `data-size="full" | "compact" | "collapsed"`.

- [ ] Step 1: failing tests (stub `overflows` with a counter so the n-th check returns false):
  - overflow once → rightmost group goes `compact` (small controls lose captions, `.dw-rb-compact`), others `full`.
  - overflow while every group is compact → rightmost group `collapsed`: a single large button with the group's icon and label; clicking opens a `.dw-rb-flyout` holding that group's live controls; clicking a control in it runs the Cmd and closes the flyout; Escape closes it and focuses the button.
  - groups shrink right to left; the first group (Clipboard) collapses last.
  - `fit()` with no overflow restores every group to `full` and moves controls back from any flyout.
- [ ] Step 2: run → FAIL. Step 3: implement (ResizeObserver on the panel calls `fit`; never read layout in `sync`). On `.is-narrow` the tab row gets `overflow-x: auto`. Step 4: tests + suite pass.

### Task 4: Quick Access Toolbar

**Files:** Modify `src/presentation/common/ribbon.ts` (QAT lives in the tab row), `styles.css`; Test `tests/ribbon.test.ts`

**Interfaces:**
- Consumes `Ribbon` opts; adds `opts.qat: { defaults: string[]; choices: string[] }`.
- Produces `Ribbon.qatIds(): string[]`; storage key `opts.key + "-qat"` (JSON array of ids).

- [ ] Step 1: failing tests:
  - with no stored list, QAT shows `defaults` in order (`.dw-rb-qat button[aria-label]`), left of File.
  - right-clicking a ribbon button whose id isn't in the QAT opens `Menu.last` with "Add to Quick Access Toolbar"; running it appends the id, re-renders, saves the list.
  - right-clicking a QAT button → "Remove from Quick Access Toolbar" removes and saves.
  - the QAT ▾ opens "Customize Quick Access Toolbar": `choices` with checkmarks (toggling adds/removes), then "More Commands…" opening a submenu of every Cmd label (alphabetical) not already shown.
  - a stored list `["Bold", "NoSuchId", 7]` loads as `["Bold"]` without throwing.
  - read-only: QAT shows only `view` commands.
- [ ] Step 2: run → FAIL. Step 3: implement. Step 4: pass.

### Task 5: Own context menu with mini toolbar

**Files:** Create `src/presentation/common/context-menu.ts`; Modify `styles.css`; Test `tests/context-menu.test.ts`

**Interfaces:**
- Produces:
  ```ts
  export interface CtxItem { label: string; run?: () => void; icon?: string; checked?: boolean; disabled?: boolean; sub?: CtxItem[] }
  export type CtxEntry = CtxItem | "---";
  export function showContextMenu(at: { x: number; y: number }, entries: CtxEntry[], opts?: { mini?: (el: HTMLElement) => void; onClose?: () => void }): { close(): void };
  ```
  Uses Obsidian's menu classes (`menu`, `menu-item`, `menu-separator`, `menu-item-icon`, `menu-item-title`) inside `.dw-ctx` so it looks like Obsidian's own menu.

- [ ] Step 1: failing tests:
  - renders items/separators in order; separators at the start/end or doubled are dropped.
  - clicking an item runs it and closes (menu and mini toolbar removed from the DOM, `onClose` called once).
  - `mini` callback receives an element positioned above the menu (`.dw-ctx-mini`).
  - ArrowDown/ArrowUp move `.selected` skipping disabled items; Enter runs; ArrowRight on an item with `sub` opens the submenu (`.dw-ctx-sub`), ArrowLeft closes it; Escape closes all.
  - mousedown outside, window `blur`, and `scroll` (capture) close it.
  - placed at x/y; when `x + width > innerWidth` (stub `getBoundingClientRect`) it is moved left to fit; same for bottom.
  - opening a second menu closes the first.
- [ ] Step 2: run → FAIL. Step 3: implement (one module-level open instance). Step 4: pass.

### Task 6: Status bar

**Files:** Create `src/presentation/common/status-bar.ts`; Modify `styles.css`; Test `tests/status-bar.test.ts`

**Interfaces:**
- Produces:
  ```ts
  export interface StatusItem { id: string; label: string; render: (el: HTMLElement) => void; side: "left" | "right" }
  export class StatusBar {
    constructor(el: HTMLElement, items: StatusItem[], storage: AppStorage, key: string);
    update(): void;                    // re-renders every shown item
    zoom(opts: { get: () => string; set: (z: string) => void; levels: string[] }): void;  // − slider + percentage at the right end
    readonly save: HTMLElement;        // where SaveStatus mounts
  }
  ```
- [ ] Step 1: failing tests:
  - renders left items then right items in the given order; `update()` calls each `render` again.
  - right-click opens `Menu.last` "Customize Status Bar" with each item's label checked; unchecking hides it and saves the hidden ids under `key + "-status"`; a new StatusBar with that storage starts with it hidden.
  - zoom: − / + step through `levels` ("fit" excluded from stepping), the range input maps to the nearest level, the percentage button opens the Zoom menu (`Menu.last` lists `zoomLabel` of each level with the current checked).
- [ ] Step 2: run → FAIL. Step 3: implement. Step 4: pass.

### Task 7: Word layout data and command table

**Files:**
- Create: `src/presentation/docx/word-ui.ts` (tabs, context-menu layout, QAT defaults/choices, Word shortcut table, status items ids)
- Modify: `src/presentation/docx/docx-view.ts` (new `private wordCommands(ctx): Commands`, reusing existing methods; small helpers it needs: `growFont(dir: 1 | -1)` stepping `FONT_SIZES` via `cmd.setMark(m.size)`, `deleteCurrentComment()`, `acceptAll(accept: boolean)` extracted from the review panel's `everything`)
- Test: `tests/word-ui.test.ts`

**Interfaces:**
- Produces (`word-ui.ts`): `wordTabs(c: WordControls): Tab[]` where `WordControls` carries the select/colour/gallery callbacks the view supplies (`font`, `size`, `style`, `highlight`, `color`, `zoom`, `inTable`, `onPicture`, `onDrawing`); `WORD_QAT = { defaults: ["FileSave", "Undo", "Redo"], choices: ["FileSave", "Undo", "Redo", "FilePrint", "FindDialog", "WordCount", "ReviewTrackChanges"] }`; `WORD_SHORTCUTS: [keys: string, id: string][]`; `WORD_SHORTCUT_SECTIONS: ShortcutSection[]`.
- Command ids: Microsoft's where they exist (see `docs/office-ribbon/word.md`): `FileSave, Undo, Redo, FilePrint, Paste, PasteTextOnly, Cut, Copy, FormatPainter, Font, FontSize, FontSizeIncreaseWord, FontSizeDecreaseWord, ClearFormatting, Bold, Italic, Underline, Strikethrough, Subscript, Superscript, InlineCode, TextHighlightColorPicker, FontColorPicker, BulletsGalleryWord, NumberingGalleryWord, IndentDecreaseWord, IndentIncreaseWord, AlignLeft, AlignCenter, AlignRight, AlignJustify, LineSpacingGallery, ParagraphDialog, QuickStylesGallery, UpdateStyle, FindDialog, ReplaceDialog, SelectAll, Delete, PageBreakInsertWord, TableInsertGallery, PictureInsertFromFile, ShapesInsertGallery, ChartInsert, HyperlinkInsert, ReviewNewComment, HeaderInsertGallery, FooterInsertGallery, TextBoxInsert, HorizontalLineInsert, EquationInsert, SymbolInsert, WatermarkGallery, TableColumnsGallery, BreaksGallery, TextWrapGallery, TableOfContentsInsert, TableOfContentsUpdate, FootnoteInsert, CreateNoteFromSelection, CopyLinkToHeading, CompanionNote, LinkedFrom, WordCount, ReviewDeleteComment, ShowComments, ReviewTrackChanges, ReviewingPane, ReviewAcceptAll, ReviewRejectAll, VersionHistory, ViewPrintLayoutView, ViewWebLayoutView, DarkModeOn, ViewRulerWord, NavigationPaneShowHide, ZoomDialog, ZoomCurrent100, ZoomPageWidth, KeyboardShortcuts, TableShading, TableBorders, TableDeleteRows, TableDeleteColumns, TableDelete, TableRowsInsertAboveWord, TableRowsInsertBelowWord, TableColumnsInsertLeft, TableColumnsInsertRight, MergeCellsRight, MergeCellsDown, SplitCells, TableColumnWidth, TableRepeatHeaderRows, TablePropertiesDialog, PictureChange, AltText, PictureCrop, EditShape, ExportFile`.

- [ ] Step 1: failing tests in `tests/word-ui.test.ts` with a fake `Commands` of every id above:
  - `wordTabs(...)` core tab labels in order: `["Home", "Insert", "Design", "Layout", "References", "Review", "View", "Help"]`; contextual: `["Table Design", "Layout", "Picture Format", "Shape Format"]`.
  - Home group labels in order: `["Clipboard", "Font", "Paragraph", "Styles", "Editing"]`; Review: `["Proofing", "Comments", "Tracking", "Changes", "Compare"]` (exact ones per spec tables).
  - every control id used by `wordTabs` exists in the id list (no typos).
  - coverage: the set of ids reachable from tabs + context menu layout + File ∪ {Undo, Redo via QAT} contains every id in a `GOOGLE_PARITY` list in the test that names the Google layout's commands: every toolbar control and every Edit/View/Insert/Format/Tools/Help menu item in `buildToolbar`, plus Code, column width, repeat header row, row break toggle, Delete, word-count toggle (status bar customize).
  - `WORD_SHORTCUTS` keys are unique and every id exists.
- [ ] Step 2: run → FAIL. Step 3: implement `word-ui.ts` from the spec tables and `wordCommands` in the view (each entry calls the same method / `cmd.*` the Google toolbar calls; `view: true` on look-only commands: Copy, SelectAll, FindDialog, WordCount, views, zoom, ruler, navigation, dark mode, KeyboardShortcuts, VersionHistory, ExportFile, FilePrint, CopyLinkToHeading, CreateNoteFromSelection, CompanionNote, LinkedFrom). Step 4: pass.

### Task 8: Word view uses the ribbon when the setting is on

**Files:** Modify `src/presentation/docx/docx-view.ts` (onLoadFile, buildToolbar call site, `this.toolbar` uses, focusRegions, word count, `onunload`); `styles.css`; Test `tests/views.test.ts`

**Interfaces:**
- Consumes Tasks 2–7. Adds `private ribbon: Ribbon | null`, `private statusBar: StatusBar | null`; a small private `ui()` returning `{ sync(): void; fit(): void }` from whichever is built so existing `this.toolbar?.sync()` / `fit()` calls go through it.

- [ ] Step 1: failing view tests (new `describe("DocxView Word-style interface")`, settings `{ ...DEFAULT_SETTINGS, officeDocx: true }`):
  - opens with `.dw-rb` and `.dw-statusbar`, no `.dw-menubar` / `.dw-toolbar`; Home selected; Undo/Redo/Save in the QAT.
  - clicking Home › Bold toggles bold on the selection (doc has a `bold` mark) and the button shows `is-active` after the editor's selection update.
  - placing the cursor in a table shows "Table Design" and "Layout" contextual tabs; selecting a picture shows "Picture Format".
  - status bar shows `Page 1 of N` and `N words`; Customize Status Bar unchecking Word Count hides it.
  - read-only (`readOnly: true, officeDocx: true`): no Insert/Layout/References/Review tabs; QAT has no Save; nothing is written (`app.vault.modifyBinary` not called).
  - `officeDocx: false` (default): the existing Google menu/tool tests still pass unchanged (no new assertions needed beyond the suite).
  - re-running `onLoadFile` with `officeDocx` flipped switches layouts, the document text is unchanged, and the view isn't dirty.
- [ ] Step 2: run → FAIL. Step 3: implement; `Ctrl+F1` toggles ribbon collapse through the view `Scope`. F6 regions: `.dw-rb`, editor, outline, review, find, `.dw-statusbar`. Step 4: suite, typecheck, lint pass.

### Task 9: Word right-click menu and mini toolbar

**Files:** Modify `src/presentation/docx/docx-view.ts` (`contextMenu` branches on the setting), `src/presentation/docx/word-ui.ts` (`wordContextMenu(state): CtxEntry[]` layout by ids), `styles.css`; Test `tests/views.test.ts`

**Interfaces:** Consumes `showContextMenu`, Word command table.

- [ ] Step 1: failing tests:
  - Word mode, right-click on text with a selection: `.dw-ctx` shows, in order, `Cut, Copy, Paste, Keep Text Only, ---, Link…, New Comment, ---, Paragraph…, Create Note from Selection`; `.dw-ctx-mini` holds Font, Size, Grow, Shrink, Styles, Bold, Italic, Underline, Highlight, Font Color, Bullets, Numbering, New Comment; clicking mini Bold bolds the selection and keeps the menu open (Word keeps it).
  - right-click on plain text with no selection in Word mode also opens `.dw-ctx` (Google mode still returns early as today).
  - in a table: adds `Insert ▸ (Insert Rows Above, Insert Rows Below, Insert Columns to the Left, Insert Columns to the Right)`, `Delete Cells ▸ (Delete Row, Delete Column, Delete Table)`, `Merge Cells ▸ (Right, Down)`, `Split Cells`, `Cell Shading ▸`.
  - on a picture: `Cut, Copy, Change Picture…, Wrap Text ▸, Crop, Edit Alt Text…, Delete`; no mini toolbar.
  - misspelled word: suggestions first, then "Add to Dictionary", as now.
  - Google mode: `Menu.last` used, behaviour unchanged.
- [ ] Step 2: run → FAIL. Step 3: implement. Step 4: pass.

### Task 10: Word shortcuts in Word mode

**Files:** Modify `src/presentation/docx/docx-view.ts` (extraKeys at line 324, scope registrations), `src/presentation/docx/docx-editor.ts` (`docxPlugins(ctx, extraKeys, opts?: { googleKeys?: boolean })`: when false, leave out `Mod-Shift-x`… only if Word assigns that key differently; keep shared ones), `src/presentation/common/shortcuts.ts` (`WORD_SHORTCUTS` sections); Test `tests/views.test.ts`

- [ ] Step 1: failing tests (Word mode, key events dispatched to the editor DOM):
  - Ctrl+E centres the paragraph; Ctrl+J justifies; Ctrl+L / Ctrl+R left / right.
  - Ctrl+] grows the selection's size one step in `FONT_SIZES`; Ctrl+[ shrinks.
  - Ctrl+Shift+L makes a bulleted list; Ctrl+Shift+8 does nothing in Word mode.
  - Ctrl+Alt+1 sets Heading 1; Ctrl+Shift+N sets Normal.
  - Ctrl+Shift+E toggles Suggesting.
  - Ctrl+E reaches the editor even though Obsidian binds it: registered on the view `Scope` (stub records registrations; assert `["Mod"], "e"` is registered only in Word mode).
  - Help › Keyboard Shortcuts (Help tab) lists the Word set.
  - Google mode: Ctrl+Shift+8 still makes a bulleted list.
- [ ] Step 2: run → FAIL. Step 3: implement from the spec's table, skipping rows whose command doesn't exist. Step 4: pass.

### Task 11: Live check and docs

**Files:** Modify `README.md` (one paragraph: Settings › Word documents › Word-style interface (beta)), `AGENTS.md` / `CLAUDE.md` layout lines for the new modules.

- [ ] Step 1: `npm run build && cp main.js manifest.json styles.css ~/notes/.obsidian/plugins/doc-works/ && obsidian plugin:reload id=doc-works && obsidian dev:errors` → no errors.
- [ ] Step 2: in a new tab, open a scratch copy in `~/notes/doc-works-test/`, turn the setting on via `obsidian eval` on the plugin settings + `reopenViews()`; screenshot each tab, a contextual tab, the folded ribbon in a narrow split, the right-click menu with mini toolbar, the status bar; dark and light Obsidian themes.
- [ ] Step 3: `npm run check:files -- ~/notes/doc-works-test` unchanged (UI only); restore the setting to the owner's value, close test tabs, delete scratch copies.

---

## Part 2: PowerPoint and Excel (spec section "PowerPoint and Excel")

Same Global Constraints and Review Focus, read per app. Each task: failing tests first, then code, then `npm test`, typecheck, lint.

### Task 12: Shared office chrome + settings
- Create `src/presentation/common/office-ui.ts`: `buildOfficeUi(opts: { host: HTMLElement; root: HTMLElement; app: App; cmds: Commands; tabs: Tab[]; key: string; readOnly: boolean; qat: { defaults: string[]; choices: string[] }; fileIds: string[]; fileExtras: (m: Menu) => void; sides: [icon: string, title: string, run: () => void][]; status: StatusItem[]; zoom?: { get(): string; set(z: string): void; levels: string[] } }): { ribbon: Ribbon; statusBar: StatusBar; save: SaveStatus | null }`. Move DocxView.buildRibbon's generic part onto it (Word tests unchanged and green).
- Settings: `officePptx`, `officeXlsx` (default false, migrated), toggles in their sections, reopen on change. Tests in `tests/settings.test.ts`.

### Task 13: `pptx/powerpoint-ui.ts` (data) — tests `tests/powerpoint-ui.test.ts`: tab and group order per spec, contextual tests, ids all in `PPT_COMMAND_IDS`, Google parity list from `PptxView.buildToolbar` / slide menu covered, shortcuts unique.
### Task 14: PptxView ribbon + status bar — tests in `tests/powerpoint-view.test.ts` (ribbon instead of toolbar, every id provided, Bold works, Shape/Picture Format contextual, Slide N of M, read-only, layout switch on reopen without saving).
### Task 15: PptxView right-click + mini toolbar + PowerPoint shortcuts (scope-registered, only in PowerPoint mode).
### Task 16: `xlsx/excel-ui.ts` (data) — tests `tests/excel-ui.test.ts` as Task 13.
### Task 17: XlsxView ribbon + status bar (Average / Count / Sum of the selection) — tests `tests/excel-view.test.ts`.
### Task 18: XlsxView right-click + mini toolbar + Excel shortcuts.
### Task 19: Live check of both in Obsidian, README / CLAUDE.md.
