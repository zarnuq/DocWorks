# Doc Works

**Open, read and edit Word and PowerPoint files right inside Obsidian.** No Office, no LibreOffice, no conversion. Your `.docx` and `.pptx` files stay as they are, next to your notes.

[![Release](https://img.shields.io/github/v/release/zarnuq/DocWorks)](https://github.com/zarnuq/DocWorks/releases)
[![Downloads](https://img.shields.io/github/downloads/zarnuq/DocWorks/total)](https://github.com/zarnuq/DocWorks/releases)
[![License: GPL-3.0](https://img.shields.io/badge/license-GPL--3.0-blue)](LICENSE)

![Editing a PowerPoint slide in Obsidian](docs/pptx.gif)

## Why Doc Works

- **Slides that look like PowerPoint.** It renders themes, layouts, gradients, shapes, SmartArt, bullets and spacing, and lines wrap where PowerPoint wraps them, even without Office fonts installed.
- **Edits that don't break your files.** Only the parts you change are rewritten. Everything else, including animations, charts, comments and fields, is saved back byte-for-byte. Opening a file never modifies it.
- **Self-contained.** One bundled plugin with no system dependencies. It works offline and has no telemetry.
- **Word and PowerPoint in one plugin**, with undo/redo, autosave and light/dark themes.

If Doc Works saves you a trip to Office, please ⭐ the repo. It helps other people find it.

| PowerPoint | Word |
|---|---|
| ![Slide editor](docs/pptx.png) | ![Document editor](docs/docx.png) |

## Features

### Word (.docx)

- Paragraph styles (Normal, Title, Heading 1–4, Quote)
- Character formatting: bold / italic / underline / strike / sub- and superscript / code font, font, size, text color, highlight; clear formatting
- Paragraph formatting: alignment, line spacing, indent
- Bullet and numbered lists with real Word numbering:
  - Tab / Shift+Tab change the list level
  - Markdown shortcuts: `- `, `1. `, `# `, `> `
- Tables: insert, add/delete rows and columns, cell shading
- Images (insert, paste, drop, resize, alt text, replace), page breaks (Ctrl+Enter), horizontal lines
- Display uses the document's own styles (fonts, sizes, colors, spacing) and list markers from `numbering.xml`
- Light and dark paper, autosave

Lossless by design: unedited paragraphs and tables are copied back unchanged. Edited ones keep every property the editor doesn't model. Hyperlinks, bookmarks, fields, footnote references, content controls and section breaks survive edits. Opening a file never modifies it.

### PowerPoint (.pptx)

- Slide thumbnails, a main slide canvas and editable speaker notes (a notes page is created when a slide has none)
- Slideshow (▶ Present, F5 from the start, Shift+F5 from the current slide): arrow keys / space / click to advance, type a number + Enter to jump, Esc to exit
- Presenter view (press P in the slideshow): current and next slide, notes, slide counter, timer and clock
- Hide slides from the slideshow
- Faithful rendering: placeholder inheritance from layouts and masters, theme colors, fonts and style matrix, `lumMod` / `tint` / `alpha`, gradients, all 187 DrawingML preset shapes and custom shapes (SVG), dashes, arrowheads, shadows, pictures cropped to shapes, tables, groups, SmartArt, background colors and images
- Text wraps like PowerPoint even without Office fonts: built-in width tables for ~30 common fonts (Calibri, Cambria, Aptos, Arial, Segoe UI, Verdana...) adjust letter spacing to the original widths
- Text editing directly on the slide:
  - Bold / italic / underline / strike, font, size, color and character spacing
  - Clear formatting, alignment, bullets and numbered lists (1. a) i. ... with list levels via Tab / Shift+Tab), line spacing
  - Hyperlinks (Ctrl+click follows them; clickable in the slideshow) and fields: live slide number and date
  - Text box settings: vertical alignment, margins, wrapping, autofit (shrink text on overflow / resize shape to fit text)
- Shapes:
  - Insert any of the 187 preset shapes from the Shapes gallery (drag to draw, Shift for squares / 45° lines), text boxes and images (file, paste or drop)
  - Lines and connectors glue to shapes' connection points and follow them when the shapes move
  - Move (drag the edge or the top-left handle, arrow keys to nudge), resize (Shift keeps the aspect ratio), rotate (top handle, Shift snaps to 15°), adjust the shape (yellow handles), delete (Delete key)
  - Format panel: fill (none, solid with transparency, linear or radial gradient), line (color, width, dash, arrowheads), shadow, exact position, size and rotation, flip, bring to front / send to back
- Arranging:
  - Select several shapes (Shift / Ctrl+click, drag a box on the slide, Ctrl+A) and move, nudge, delete or restack them together
  - Group / ungroup (Ctrl+G / Ctrl+Shift+G); groups move and scale as one
  - Align (left, center, right, top, middle, bottom; one shape aligns to the slide) and distribute evenly
  - Snapping guides to other shapes' edges and centers and the slide center while dragging (hold Alt to turn off)
  - Cut, copy, paste and duplicate (Ctrl+X / C / V / D, or right-click) between slides and between presentations; pictures and links come along
- Slides: new (copies the current layout's placeholders), duplicate, delete, reorder
- Undo / redo for everything (Ctrl+Z / Ctrl+Y)
- Navigate slides with arrow keys or PageUp / PageDown
- Light and dark workspace, debounced autosave, plus a save on blur and on quit

Saving changes only what you edited. Each XML part stays a live DOM, and text is rewritten with minimal diffs. Unedited text, animations, transitions, charts and SmartArt keep their content.

### Not supported (yet)

- PowerPoint:
  - Editing tables or SmartArt (they render read-only); charts show a placeholder
  - Animations and transitions
  - Moving shapes that sit inside a group (their text is still editable)
- Word: tracked-change wrappers and smart tags are dropped from a paragraph you edit; there are no comments, headers or footers in the editor (they are preserved in the file).

## Install

**Community plugins** (once listed): Settings → Community plugins → Browse → search "Doc Works".

**BRAT:** add `zarnuq/DocWorks` in the [BRAT](https://github.com/TfTHacker/obsidian42-brat) plugin.

**Manual:** download `main.js`, `manifest.json` and `styles.css` from the [latest release](https://github.com/zarnuq/DocWorks/releases), put them in `<vault>/.obsidian/plugins/doc-works/`, then enable **Doc Works** in Settings → Community plugins.

Then click any `.docx` or `.pptx` file in your vault. If another plugin already opens `.docx` or `.pptx`, it keeps that file type and Doc Works shows a notice; disable the other plugin to use Doc Works for it.

## Contributing

Bug reports with a sample file that renders wrong are the most useful thing you can send: open an issue and attach the file, or a screenshot next to PowerPoint's rendering. PRs are welcome. See `CLAUDE.md` for the architecture and the rules the writers follow.

## Development

```bash
npm run dev        # esbuild watch
npm run build      # production bundle -> main.js
npm test           # vitest + jsdom
npm run typecheck  # tsc --noEmit
npm run lint       # eslint with Obsidian's plugin rules (eslint-plugin-obsidianmd)
DOC_WORKS_DECKS=/path/to/decks DOC_WORKS_DOCS=/path/to/docs npm test   # also round-trip real files
DOC_WORKS_EXPORT=/tmp/out npx vitest run tests/export-samples.test.ts   # write heavily edited copies...
soffice --headless --convert-to pdf --outdir /tmp/out/pdf /tmp/out/*      # ...and check another app opens them
```

The test fixtures `tests/fixtures/sample.pptx` and `sample.docx` are generated by the `make_*_fixture.py` scripts next to them, using `python-pptx` and `python-docx`. `src/presentation/pptx/font-metrics.ts` is generated by `scripts/make_font_metrics.py` (fonttools; font sources listed in the script), and `src/domain/pptx/preset-shapes.ts` by `scripts/make_preset_shapes.py` from the ECMA-376 preset shape definitions.

## Releasing

Bump `version` in `manifest.json`, `package.json` and `versions.json`, then push a tag with the same version (`git tag -a 0.1.0 -m 0.1.0 && git push origin 0.1.0`). The release workflow tests, builds and drafts a GitHub release with `main.js`, `manifest.json` and `styles.css`.

## License

GPL-3.0
