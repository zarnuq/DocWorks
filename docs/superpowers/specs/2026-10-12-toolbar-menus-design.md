# Toolbar regroup + narrow layout (UX sub-project 5)

## Goal

The three editors' toolbars are one long row of 40–50 controls that wraps to 2–3 lines in a narrow pane.
Replace them with Google's layout: a slim menu row holding every command, plus one compact row of everyday
controls that folds into a ⋮ overflow menu instead of wrapping. Side panels float over the document in a
narrow pane. No new dependencies, no network.

## Decisions (owner-approved)

- Shape: Google-style menu row + compact row + ⋮ overflow.
- Menus: Google's per-app set, plus a small File menu (Export…).
- Contextual controls: Google's set (Docs picture / table cell, Slides shape / picture; Sheets none).
- Narrow pane (< 700 px): side panels become drawers over the document.
- Architecture: extend the existing `Toolbar` builder (approach A), not a command registry.

## Architecture (`presentation/common/toolbar.ts`)

`Toolbar` keeps its current API and gains:

- `menus(row: HTMLElement)` / `menu(name, build: (m: Menu) => void)`: a menu-row button that builds an
  Obsidian `Menu` each time it opens (so checks and disabled items reflect the current state) and shows it
  below the button. Obsidian's typed `Menu` has no submenus: long menus use `setSection` separators; items
  marked ▸ (e.g. Format › Align, Insert › Break) open a second `Menu` beside the first, via a shared helper
  `subMenu(m, title, build)`.
- `pin()`: marks the current group as never folding (Undo/Redo, Style/Font, the end area).
- `context(test: () => boolean)`: marks the current group as contextual; `sync()` shows it only when `test()`
  holds. Views call `tb.sync()` from their existing selection updates (`updateFormatControls` in docx and
  pptx).
- Overflow: a ⋮ button before `.dw-end`. A `ResizeObserver` on the bar (plus `sync()`) hides groups from
  the right, never pinned ones, until the row fits on one line (`flex-wrap: nowrap`). Opening ⋮ lists the
  hidden groups' controls as menu items, separated per group:
  - button → item with its aria-label (and `is-active` → checked), click proxies `button.click()`;
  - select / iconSelect → a ▸ item whose submenu lists the options; picking one sets `value` and dispatches
    `change`;
  - colour → item that opens the hidden `<input type="color">` (`showPicker()`, else `click()`).
  No control is duplicated: the menu proxies the real hidden element, so handlers and state stay in one place.
- The Light/Dark segmented control leaves `end()`; each view's View menu sets the theme, checking the current
  one when the menu opens. `end()` takes no theme arguments and returns `{ status }`; views drop `syncTheme`.
- Keyboard: the menu row is a second `role="menubar"` roving group (Left/Right, Enter/Space/ArrowDown
  opens the menu). `rovingToolbar` skips hidden controls already; the ⋮ button is part of the row. F6 regions
  gain the menu row ahead of the toolbar.

## Per-editor contents

Rows list groups left to right; `·` separates groups. Menus hold every command, including those in the row.

### Documents

- Row: Undo Redo (pinned) · Style Font Size (pinned) · B I U, text colour, highlight · Link, Comment,
  Image · Align▾, Line spacing▾ · bullets, numbered · outdent, indent · Clear formatting · end: Editing /
  Suggesting select, word count, Saved.
- Contextual: picture (NodeSelection on an image) → Wrap▾, Crop, Alt text; table cell → cell fill, border▾.
- File: Export…
- Edit: Undo, Redo, Cut, Copy, Paste, Select all, Find and replace.
- View: Editing / Suggesting, Page view, Ruler, Outline, Review panel, Light / Dark paper.
- Insert: Image, Table▸, Drawing, Equation, Horizontal line, Break▸, Footnote, Comment, Link, Table of
  contents▸, Header & footer, Watermark.
- Format: Text▸ (Bold, Italic, Underline, Strikethrough, Subscript, Superscript, Code), Paragraph style▸
  (+ Update to match), Align▸, Line spacing▸, Paragraph…, Columns…, Bulleted / Numbered list, Clear formatting.
- Tools: Paint format, Word count, Find and replace.

### Slides

- Row: Undo Redo Paint format (pinned) · Zoom · Text box, Image, Shape▾, Table · Comment · Font Size
  (pinned) · B I U, text colour · Align▾, Line spacing▾ · bullets, numbered · outdent, indent · Clear
  formatting · end: ▶ Present, Saved.
- Contextual: shape → fill, border colour, border weight▾, dash▾; picture → Crop, Replace, Reset▾ (crop,
  colour adjustments).
- File: Export…
- Edit: Undo, Redo, Cut, Copy, Paste, Delete, Find and replace.
- View: Present, Present from start, Grid view, Format panel, Zoom▸, Light / Dark.
- Insert: Text box, Image, Audio / video, Shape, Table, Chart, Link, Field▸, Comment, New slide.
- Format: Text▸, Align▸, Line spacing▸, Character spacing▸, Bullets / Numbering, Clear formatting, Format
  options (opens the panel).
- Slide: today's Slide▾ menu (new, duplicate, delete, hide, move, sections), Theme…, Motion…
- Arrange: Order▸, Align▸, Distribute▸, Group, Ungroup (the existing arrange control).

### Sheets

- Row: Undo Redo (pinned) · Zoom · $ % .0 .00 123▾ · Font Size (pinned) · B I S U, text colour, fill,
  Borders▾, Merge · H-align▾, V-align▾, Wrap · Comment, Chart, Filter, Functions▾ · end: Saved.
- File: Export…
- Edit: Undo, Redo, Cut, Copy, Paste, Find and replace.
- View: Freeze▸, Zoom▸, Light / Dark.
- Insert: Sheet, Chart, Pivot table, Function▸, Comment, Note, Dropdown (validation).
- Format: Number▸, Text▸, Alignment▸, Wrap, Merge▸, Conditional formatting, Clear formatting.
- Data: Sort A→Z, Sort Z→A, Filter, Data validation, Remove duplicates, Split text to columns, Named ranges,
  Group / Ungroup rows and columns, Import CSV.

The Data iconSelect and Freeze iconSelect leave the row (they move to Data / View).

## Narrow layout

- A `ResizeObserver` on the view's content element toggles `is-narrow` below 700 px.
- In `is-narrow`, docked side panels (docx outline + review, pptx format panel) are hidden; their toggles
  open them as drawers: `position: absolute` over the document's edge, shadowed, closed by Escape, a click
  outside or the toggle. The remembered wide-mode open state is not changed by narrow mode.
- Slide thumbnails become a ~48 px strip of slide numbers (the listbox keeps its keyboard behaviour); the
  full strip comes back above 700 px.
- `fitDocument` (docx) and the slide fit ignore drawers, so the page doesn't reflow when one opens.

## Testing

- `tests/toolbar.test.ts` (new): menus build on open with current state; overflow hides unpinned groups
  from the right and lists them in ⋮ (buttons click through, selects dispatch `change`); pinned groups never
  hide; `context` groups follow `sync()`. jsdom has no layout: overflow is driven through an injectable
  width measure.
- `tests/views.test.ts` / `xlsx-view.test.ts`: each editor's menus contain the commands listed above; picking
  a menu item runs the same action as before (spot checks per menu); contextual groups appear for a picture /
  shape selection; `is-narrow` turns panels into drawers and Escape closes them.
- Existing toolbar tests that look up buttons by aria-label keep working (controls keep their labels).
- Live: each editor at full width and at ~500 px; overflow, menus, drawers, keyboard (Tab into the menu row,
  arrows, F6), light and dark themes; screenshots.

## Out of scope

Properties sidebar consistency (sub-project 6), "search the menus", new commands that don't exist today
(page setup, spelling tool).
