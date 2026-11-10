# Office-style interface (beta): design

## Intent

Some people know desktop Word, PowerPoint and Excel far better than Google Docs, Slides and Sheets. Doc Works'
editors copy Google's interface: a menu row above one compact toolbar. This adds a second interface, chosen per
file type in settings, that copies the desktop Office apps' layout instead: ribbon tabs with labelled groups, a
Quick Access Toolbar, Office-style right-click menus and a status bar.

It is a UI addition only. It reaches the same commands the Google layout reaches; it adds no editing features.
The Google layout stays exactly as it is and stays the default. The new interface is marked beta.

This document covers the settings change, the shared ribbon pieces, and Word. PowerPoint and Excel come next,
after the owner's questions on them; they reuse everything here.

## Goals

- A per-file-type setting, a toggle labelled with "(beta)", default off (Google layout).
- Word documents: desktop Word's ribbon, Quick Access Toolbar, right-click behaviour and status bar, laid out as
  close to 1:1 as the commands Doc Works has allow.
- Looks at home in Obsidian: theme colours, fonts and spacing from Obsidian's CSS variables; only Word's layout
  is copied.
- Works in narrow panes and on phones by folding groups the way Word does.

## Non-goals

- New editing commands. A Word control with no Doc Works command is left out, not greyed out. Tabs and groups
  that end up empty are dropped (Mailings, Draw, Design's Document Formatting, Layout's Page Setup sizes…).
- The File backstage page. File is a drop-down menu.
- Alt-key letter shortcuts (KeyTips).
- Word's colours, icons or name in the UI. Icons are Lucide (Obsidian's set). Settings call it "Word-style".

## Anti-goals

- Any change to the Google layout's behaviour or look when the setting is off.
- Opening or switching interfaces must not change or save the file.
- A second copy of each command's logic: both layouts must call the same code, so a fix in one reaches both.

## Reference

Microsoft's Office Fluent UI command identifiers (MIT) list every tab, group and control of the desktop apps in
ribbon order, including contextual tab sets and right-click menus. Outlines are in `docs/office-ribbon/`
(`word.md`, `powerpoint.md`, `excel.md`). They give ids and order, not captions or button sizes; captions and
large/small sizes follow desktop Word (Microsoft 365, Current Channel).

## Settings (decided with the owner)

The settings tab is split into sections:

- **General:** Read only, Version history, Keep versions for, Comment author.
- **Word documents:** Open Word documents (.docx); Paper colour; Word-style interface (beta); Word page size.
- **PowerPoint presentations:** Open PowerPoint presentations (.pptx); Paper colour; PowerPoint-style interface (beta).
- **Excel spreadsheets:** Open Excel spreadsheets (.xlsx); Paper colour; Excel-style interface (beta).

Paper colour is now per file type (it was one setting for all three). The interface toggles are Obsidian
toggles (the "slider"), default off. Changing one reopens open views of that type, as Read only does.

My call: a toggle appears only once its interface exists, so this work ships the Word one; the PowerPoint and
Excel sections get theirs with their ribbons. No setting that does nothing reaches a release.

Builder's notes (my calls):

- Use Obsidian's setting groups (`SettingDefinitionGroup`, available in the 1.13 typings) in
  `getSettingDefinitions`, and headings in the older `display()` path.
- Paper colour options per type: Light, Dark, Match Obsidian; defaults Light for Word, Dark for slides and
  sheets (today's "default"). The old single `paper` setting migrates: "light"/"dark"/"obsidian" copy to all
  three; "default" gives the per-type defaults. Each view's own Light/Dark switch keeps working as now.
- Manifest/README wording must not use "Microsoft" as a product claim; "Word-style" is descriptive.

## Shared pieces (all three apps)

### Commands, once

Each view gets a table of its commands (id → caption, Lucide icon, run, and optional `active` / `enabled` /
dropdown builder). The ribbon, the Quick Access Toolbar, the mini toolbar and the Office right-click menu use
it. The Google toolbar may keep building its controls as it does now; where it is cheap, it can call the same
table entries, but it must not change behaviour. Command ids follow Microsoft's control ids where one exists
(`Bold`, `FormatPainter`, `TableInsertGallery`…) so the spec outlines can be cross-checked.

### Ribbon (`presentation/common/ribbon.ts`)

- A tab row: File (drop-down), then the app's tabs, then contextual tabs (shown only while their test holds,
  coloured with the accent like Word's), with the existing side icons (Version history, Companion note,
  Linked from, Comments) at the right end of the row, where Word puts Comments and Share.
- A panel showing the selected tab's groups. A group has a caption under it and may have a dialog launcher
  (the small ↘ in its corner) when Doc Works has a matching dialog (Paragraph…).
- Control kinds: large button (icon over caption), small button (icon + caption, three to a column), icon-only
  small button, split button (main action + ▾ menu), drop-down (Font, Size, Zoom), colour split button (last
  colour + ▾ picker), checkbox (View › Ruler), and an in-ribbon gallery (Styles) with a ▾ for the full list.
- Toggle state (Bold, Align, Ruler…) refreshes from the view's selection updates, as `Toolbar.sync` does now.
- Folding (decided: the way Word does it): when the panel overflows, groups shrink from the right, first
  dropping small buttons' captions, then collapsing a whole group into one button showing the group's icon
  and caption, which opens the group's controls in a drop-down panel. On a phone the tab row scrolls sideways.
- Read-only files: only commands that just look at the file (today's `VIEW_ITEMS` idea): File, Home's Copy and
  Find, View, Help.
- Keyboard: the tab row is a tablist (arrow keys), Tab moves into the panel, the panel is one roving toolbar
  as today; F6 regions include ribbon, document, status bar; Escape in a drop-down returns to its button.
- Double-clicking a tab collapses the ribbon to the tab row (Word's behaviour; cheap). Remembered per app with
  `app.saveLocalStorage`.

### Quick Access Toolbar (decided: customizable)

- Sits at the left end of the tab row, before File, separated by a divider (my call: Obsidian has no title bar
  to put it in).
- Default: Save, Undo, Redo. Save saves now (autosave still runs). A ▾ at its end opens "Customize Quick Access
  Toolbar": a checklist of Word's list limited to what Doc Works has (Save, Undo, Redo, Print = export to PDF,
  Find, Word Count, Track Changes), then "More Commands…" listing every command in the table.
- Right-click any ribbon control: "Add to Quick Access Toolbar" (or "Remove…" when it's already there).
  Right-click a QAT button: "Remove from Quick Access Toolbar".
- Stored per app as a list of command ids with `app.saveLocalStorage` (`doc-works-qat-docx`); unknown ids are
  ignored.

### Status bar (decided: yes)

A bar along the bottom of the view, as Word's: left side shows facts, right side view buttons and zoom.

- Word: "Page N of M" (from `pageAt`; clicking opens the outline, as Word's opens Navigation), "N words" (or
  "N of M words" with a selection; clicking opens Word count), "Suggesting" when in that mode, save status.
  Right: view buttons (Print Layout = page view, Web Layout = page view off), zoom − / slider / + and the
  percentage (clicking opens the Zoom menu).
- The Google layout keeps its word count and save status where they are.

### Right-click (decided: Word's behaviour, own menu)

On right-click: Word's mini toolbar (a floating bar of Font, Size, Grow/Shrink, Styles, Bold, Italic,
Underline, Highlight, Font Color, Bullets, Numbering, New Comment) above the menu, and Word's menu items in
Word's order, built from the same command table. Both close together.

Decided: the Office interface always draws its own right-click menu in the page (plain DOM, styled with
Obsidian's menu classes/variables so it looks like Obsidian's menus), whatever Obsidian's "native menus"
setting is, because an OS menu can't sit beside the mini toolbar. It supports separators, disabled items,
checkmarks, submenus (hover and arrow keys), keyboard navigation (arrows, Enter, Escape), closes on outside
click / Escape / scroll / blur, and keeps within the window. On plain text with no selection the Word interface
also shows this menu (Word does), unlike the Google layout which leaves that to the native menu; spelling
suggestions stay in it. My call: only on right-click; Word also shows a faded mini toolbar on selection, which
is left out for now.

## Word

Every command the Google layout reaches today must be reachable in the Word interface (my call: "UI addition"
means nothing is lost). A test lists the Google layout's commands and checks each has a place in the ribbon,
the File menu, the right-click menu or the status bar. Commands Word keeps elsewhere go where Word keeps them:

- Inline Code: Home › Font, after Superscript (no Word counterpart; small icon button).
- Edit › Delete: right-click menu and Home › Editing › Select ▾ (Word has no Delete button; Delete key works).
- Table column width…, Repeat Header Rows, "Allow row to break across pages": Table Layout › Cell Size
  (Width…) and Data (Repeat Header Rows), plus a Table Properties… launcher holding the row-break toggle, as
  Word's Table Properties does.
- Word count while typing: right-clicking the status bar shows Word's "Customize Status Bar" checklist (Page
  Number, Word Count, Track Changes, Zoom, Zoom Slider, View Shortcuts), stored with `app.saveLocalStorage`.
- Paint format double-click, Update style to match, Copy link to heading, Create note from selection,
  Companion note, Linked from: as placed in the tables below.

The mapping below lists only what's kept. Captions are Word's; the right column is today's command.

### File (drop-down)

The existing file menu (Export…, Version history…, Rename…, Delete file, Obsidian's file-menu items), with
"Print…" = Export to PDF, the same as the Google layout's File menu.

### Home

| Group | Controls (order as Word) | Doc Works |
|---|---|---|
| Clipboard ↘ none | Paste (large split: Paste, Keep Text Only), Cut, Copy, Format Painter | Edit menu items; paint format (double-click keeps painting) |
| Font | Font, Font Size, Grow Font, Shrink Font, Clear All Formatting; Bold, Italic, Underline, Strikethrough, Subscript, Superscript, Text Highlight Color (▾ incl. No Color), Font Color | existing marks; Grow/Shrink step through `FONT_SIZES` (UI over the existing size mark, my call) |
| Paragraph ↘ Paragraph… | Bullets, Numbering, Decrease Indent, Increase Indent; Align Left, Center, Align Right, Justify, Line and Paragraph Spacing (▾ incl. Line Spacing Options… = Paragraph…) | existing |
| Styles | Styles gallery (Normal, Title, Heading 1…), ▾ adds "Update <style> to Match Selection" | `STYLES`, `updateStyle` |
| Editing | Find, Replace, Select ▾ (Select All) | find bar, `selectAll` |

### Insert

| Group | Controls | Doc Works |
|---|---|---|
| Pages | Page Break | `insertPageBreak` |
| Tables | Table (grid picker) | `fillTableMenu` |
| Illustrations | Pictures, Shapes (drawing editor), Chart ▾ | `pickImage`, `drawingDialog`, chart types |
| Links | Link | `editLink` |
| Comments | Comment | `addComment` |
| Header & Footer | Header, Footer (both open the Header & footer dialog) | `headerFooterDialog` |
| Text | Text Box (drawing editor), Horizontal Line (Word keeps it under Borders; here for lack of a Borders control, my call) | `drawingDialog`, `insertRule` |
| Symbols | Equation, Symbol | `equationDialog`, special characters |

### Design

| Page Background | Watermark | `watermarkDialog` |
|---|---|---|

### Layout

| Group | Controls | Doc Works |
|---|---|---|
| Page Setup | Columns (▾ One, Two, Three, More Columns… = Columns…), Breaks ▾ (Page, Column, Next Page, Continuous) | `columnsDialog`, break commands |
| Paragraph ↘ Paragraph… | Indent Left / Right, Spacing Before / After (number fields) only if the paragraph commands can set them directly; otherwise the launcher alone | Paragraph dialog |
| Arrange | Wrap Text ▾ (picture selected) | `setWrap` |

### References

| Group | Controls | Doc Works |
|---|---|---|
| Table of Contents | Table of Contents (insert), Update Table | TOC commands |
| Footnotes | Insert Footnote | `addFootnote` |
| Notes (Doc Works' own, my call; no Word counterpart) | Create Note from Selection, Copy Link to Heading, Companion Note, Linked From | existing |

### Review

| Group | Controls | Doc Works |
|---|---|---|
| Proofing | Word Count | `wordCountDialog` |
| Comments | New Comment, Delete (comment at cursor), Show Comments (review panel) | `addComment`, `commentActions`, `toggleReview` |
| Tracking | Track Changes (toggle = Suggesting mode), Reviewing Pane | mode, review panel |
| Changes | Accept ▾ (Accept All Changes), Reject ▾ (Reject All Changes) | review panel's "everything" |
| Compare | Compare ▾ (Version History…) | `versionHistoryDialog` |

### View

| Group | Controls | Doc Works |
|---|---|---|
| Views | Print Layout, Web Layout | page view on / off |
| Dark Mode (Word 365 has View › Dark Mode) | Dark Mode toggle | paper theme |
| Show | Ruler, Navigation Pane (checkboxes) | `toggleRuler`, `toggleOutline` |
| Zoom | Zoom (menu), 100%, Page Width | `setZoom` (Page Width = "fit") |

### Help

| Help | Keyboard Shortcuts | shortcuts dialog |
|---|---|---|

### Keyboard shortcuts (decided: Word's in Word mode)

While the Word interface is on, the editor uses Word's shortcuts instead of Google's where they differ; tooltips
and Help › Keyboard Shortcuts show the active set. Keys that Obsidian's hotkeys may take (Ctrl+E is Obsidian's
editing/reading toggle) are registered on the view's `Scope`, never as plain keydown listeners, and act only
while the document has focus. Word's set, for commands Doc Works has:

| Keys | Command |
|---|---|
| Ctrl+B / I / U | Bold / Italic / Underline (same as now) |
| Ctrl+= / Ctrl+Shift+= | Subscript / Superscript |
| Ctrl+] / Ctrl+[ ; Ctrl+Shift+> / < | Grow / Shrink Font |
| Ctrl+Space | Clear character formatting |
| Ctrl+Shift+C / Ctrl+Shift+V | Copy / paste formatting (Format Painter) |
| Ctrl+L / E / R / J | Align left / center / right / justify |
| Ctrl+1 / Ctrl+2 / Ctrl+5 | Line spacing single / double / 1.5 |
| Ctrl+0 | Add or remove 12 pt space before |
| Ctrl+M / Ctrl+Shift+M | Increase / decrease indent |
| Ctrl+Shift+L | Bulleted list |
| Ctrl+Alt+1 / 2 / 3, Ctrl+Shift+N | Heading 1 / 2 / 3, Normal |
| Ctrl+Enter / Ctrl+Shift+Enter | Page break / column break |
| Ctrl+K | Insert link (same) |
| Ctrl+Alt+M | New comment (same in Word) |
| Ctrl+Alt+F | Insert footnote |
| Ctrl+Shift+E | Track Changes (Suggesting) |
| Ctrl+F / Ctrl+H | Find / Replace |
| Alt+= | Equation |

Commands whose Word shortcut is already the same keep working; Google-only keys (Ctrl+Shift+7/8 lists,
Ctrl+Alt+0–6 styles) are off in Word mode. The builder checks each against the existing keymap in
`docx-editor.ts` and leaves out any whose command doesn't exist.

### Contextual tabs (decided: as Word)

These replace the Google layout's selection-only groups while the Word interface is on.

- **Table Design** (cursor in a table): Shading (cell fill), Borders ▾ (All, Outside, No Border).
- **Layout** (table; Word names the second table tab "Layout" too): Delete ▾ (Rows, Columns, Table), Insert
  Above, Insert Below, Insert Left, Insert Right, Merge Cells (merge right / down, or the cell selection if a
  command for it exists), Split Cells.
- **Picture Format** (picture selected): Change Picture, Alt Text, Wrap Text ▾, Crop.
- **Shape Format** (drawing selected): Edit Shape (drawing editor). Text boxes get the same tab.

### Right-click menu (Word's items, those Doc Works has)

Text: spelling suggestions (as now), Cut, Copy, Paste Options (Paste, Keep Text Only), Link / Edit Link /
Open Link / Remove Link, New Comment, Paragraph…, Create Note from Selection. In a table add Insert ▸ (Above,
Below, Left, Right), Delete Cells ▸ (Row, Column, Table), Merge Cells, Split Cells, Cell Shading ▸. On a
picture: Cut, Copy, Change Picture, Wrap Text ▸, Crop, Edit Alt Text…, Delete. On a comment: its actions.

## Constraints

- Project rules in CLAUDE.md / AGENTS.md apply: only `main.ts`, `*-view.ts` and `presentation/common/` import
  `obsidian`; no `!important`, `text-indent` or multi-column CSS; preferences via `app.loadLocalStorage`; no
  transactions on load; editor shortcuts through the view's `Scope`; TDD with view tests on
  `tests/obsidian-stub.ts`; `npm test`, `npm run typecheck`, `npm run lint` (0 errors); checked live in Obsidian.
- No new runtime dependencies. No network.
- Mobile: touch targets, `touch-action` on drag targets; drop-downs work by tap.

## Left to the builder

- Exact pixel sizes, icons per command, caption wording where Word's is long.
- How groups are described in code (data vs builder calls), as long as each command's logic lives once.
- Shrink order details within the folding rule above.
- Whether the Google toolbar is moved onto the command table now or later.

## Later (not in this work)

PowerPoint and Excel interfaces (next, after questions), KeyTips, the File backstage page, the selection mini
toolbar, Word controls without Doc Works commands.

## PowerPoint and Excel (added 2026-11-09)

The owner wants all three apps. Every decision above carries over unchanged:
- File drop-down
- desktop tabs in Obsidian colours
- contextual tabs
- Word-style folding
- customizable Quick Access Toolbar
- status bar
- own right-click menu with mini toolbar
- the app's own shortcuts in its mode
- nothing lost from the Google layout, enforced by a parity test
- a per-type "(beta)" toggle

Settings get `officePptx` / `officeXlsx` ("PowerPoint-style interface (beta)", "Excel-style interface (beta)").
Layouts follow `docs/office-ribbon/powerpoint.md` and `excel.md`.

### PowerPoint (`pptx/powerpoint-ui.ts`)

- Tabs: Home, Insert, Design, Transitions, Animations, Slide Show, Review, View, Help.
- Contextual tabs: Shape Format (shape), Picture Format (picture).
- Home: Clipboard, Slides (New Slide, Duplicate, Delete, Hide, Section), Font, Paragraph, Drawing (Shapes, Arrange, Shape Fill, Shape Outline), Editing.
- Insert: Slides, Tables, Images, Illustrations (Shapes, Chart), Links, Comments, Text (Text Box, Field), Symbols, Media.
- Design: Themes (Theme…), Customize (Slide Size / Format Background open the slide panel).
- Transitions and Animations each open the Motion panel.
- Slide Show: From Beginning, From Current Slide, Hide Slide.
- Review: Comments, Notes (Obsidian group: Copy as Markdown outline, Copy link to slide, Create note, Companion note, Linked from).
- View: Normal, Slide Sorter (grid view), Format Pane, Dark Mode, Zoom, Fit to Window.
- Status bar: Slide N of M, Notes toggle (speaker notes pane), Comments, view buttons (Normal, Slide Sorter, Slide Show), zoom.
- Right-click on a shape or text:
  - Cut, Copy, Paste, Keep Text Only
  - Group ▸, Bring to Front ▸, Send to Back ▸
  - Link…, New Comment
  - Format Shape… (format pane), Create Note from Selection
  - On a picture: Change Picture…, Crop, Reset Picture
  - Mini toolbar: Font, Size, Grow, Shrink, B, I, U, Align Center, Bullets, Shape Fill, Shape Outline, Font Color.
  - Thumbnail menus stay as they are.
- Shortcuts:
  - Ctrl+E/L/R/J align; Ctrl+] / [ and Ctrl+Shift+> / < font size; Ctrl+Space clear formatting
  - Ctrl+M new slide; Ctrl+D duplicate; Ctrl+G / Ctrl+Shift+G group and ungroup
  - Ctrl+Shift+C / V copy and paste format; Ctrl+= / Ctrl+Shift+= sub- and superscript
  - Ctrl+K link; Ctrl+Alt+M comment
  - Google's Ctrl+Shift+7/8 lists and Ctrl+Alt+1 grid view are off.

### Excel (`xlsx/excel-ui.ts`)

- Tabs: Home, Insert, Page Layout, Formulas, Data, Review, View, Help.
- Contextual tab: PivotTable Analyze (cursor in a pivot table).
- Home:
  - Clipboard
  - Font (Font, Size, Grow, Shrink, B, I, U, Strikethrough, Borders ▾, Fill Color, Font Color)
  - Alignment (Top / Middle / Bottom, Wrap Text, Left / Center / Right, Merge & Center)
  - Number (format box, Accounting $, Percent, Increase / Decrease Decimal)
  - Styles (Conditional Formatting)
  - Cells (Insert ▾, Delete ▾, Format ▾: row height, column width, hide / unhide, rename sheet, protect sheet)
  - Editing (AutoSum ▾, Clear ▾, Sort & Filter ▾, Find & Select ▾)
- Insert: Tables (PivotTable), Illustrations (Pictures), Charts, Links, Comments, Symbols, Cell Controls (Checkbox, Dropdown).
- Page Layout: Page Setup (opens the print settings), Sheet Options (Gridlines).
- Formulas: Function Library (Insert Function ▾, AutoSum), Defined Names (Name Manager).
- Data:
  - Sort & Filter (Sort A to Z, Sort Z to A, Filter)
  - Data Tools (Text to Columns, Remove Duplicates, Data Validation)
  - Outline (Group, Ungroup)
  - Get Data (From Text/CSV)
- Review: Comments (New Comment, Show Comments), Notes (New Note), Protect (Protect Sheet), Obsidian group (Copy as Markdown table, Copy link to cells, Create note, Companion note, Linked from).
- View: Freeze Panes ▾, Zoom, 100%, Dark Mode.
- Status bar: mode (Ready / Edit), Average / Count / Sum of the selection's numbers (each switchable in Customize Status Bar, as Excel), zoom. The sheet tabs stay where they are.
- Right-click on cells keeps every item of today's cell / header menu, in Excel's order and wording:
  - Cut, Copy, Paste, Paste Values
  - Insert…, Delete…, Clear Contents
  - Sort ▸, Filter
  - New Comment, New Note, Link
  - Hide / Unhide, Row Height / Column Width
  - Lock / Unlock, pivot items, Freeze, Create note
  - Mini toolbar: Font, Size, Grow, Shrink, $, %, comma, Fill Color, Font Color, B, I, Align Center, Borders, Merge, Increase / Decrease Decimal.
- Shortcuts:
  - Ctrl+5 strikethrough
  - Ctrl+Shift+$ currency, Ctrl+Shift+% percent, Ctrl+Shift+~ General
  - Ctrl+Shift+& outline border, Ctrl+Shift+_ remove borders
  - Alt+= AutoSum, Ctrl+Shift+L filter
  - Ctrl+9 / Ctrl+0 hide rows / columns
  - Ctrl+- delete, Ctrl+Shift++ insert
  - Shift+F2 note (as now)
  - Google-only keys are off.
