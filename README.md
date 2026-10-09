# Doc Works

**Open and edit Word and PowerPoint files right inside Obsidian.**

[![Downloads](https://img.shields.io/badge/dynamic/json?logo=obsidian&color=%23483699&label=downloads&style=for-the-badge&query=%24%5B%22doc-works%22%5D.downloads&url=https%3A%2F%2Fraw.githubusercontent.com%2Fobsidianmd%2Fobsidian-releases%2Fmaster%2Fcommunity-plugin-stats.json)](https://obsidian.md/plugins?id=doc-works)
[![Release](https://img.shields.io/github/v/release/zarnuq/DocWorks?style=for-the-badge&logo=data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9IiNmZmZmZmYiIHN0cm9rZS13aWR0aD0iMiIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIiBjbGFzcz0ibHVjaWRlIGx1Y2lkZS1naXQtbWVyZ2UiPjxjaXJjbGUgY3g9IjE4IiBjeT0iMTgiIHI9IjMiLz48Y2lyY2xlIGN4PSI2IiBjeT0iNiIgcj0iMyIvPjxwYXRoIGQ9Ik02IDIxVjlhOSA5IDAgMCAwIDkgOSIvPjwvc3ZnPg==)](https://github.com/zarnuq/DocWorks/releases/latest)
[![Latest Release](https://img.shields.io/github/release-date/zarnuq/DocWorks?style=for-the-badge&label=Latest%20Release&logo=data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9IiNmZmZmZmYiIHN0cm9rZS13aWR0aD0iMiIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIiBjbGFzcz0ibHVjaWRlIGx1Y2lkZS1jYWxlbmRhci1jaGVjay0yIj48cGF0aCBkPSJNMjEgMTRWNmEyIDIgMCAwIDAtMi0ySDVhMiAyIDAgMCAwLTIgMnYxNGEyIDIgMCAwIDAgMiAyaDgiLz48bGluZSB4MT0iMTYiIHgyPSIxNiIgeTE9IjIiIHkyPSI2Ii8+PGxpbmUgeDE9IjgiIHgyPSI4IiB5MT0iMiIgeTI9IjYiLz48bGluZSB4MT0iMyIgeDI9IjIxIiB5MT0iMTAiIHkyPSIxMCIvPjxwYXRoIGQ9Im0xNiAyMCAyIDIgNC00Ii8+PC9zdmc+)](https://github.com/zarnuq/DocWorks/releases/latest)
[![License: GPL-3.0](https://img.shields.io/badge/license-GPL--3.0-blue?style=for-the-badge)](LICENSE)

![Editing a PowerPoint slide in Obsidian](docs/pptx.png)

- **Looks like PowerPoint and Word.** Themes, layouts, shapes, SmartArt, tables and document styles render faithfully, even without Office fonts.
- **Doesn't break your files.** Only what you edit is rewritten; everything else is saved back byte-for-byte.
- **Self-contained.** No Office, no LibreOffice, no system dependencies. Works offline.

If it saves you a trip to Office, please ⭐ the repo.

![Editing a Word document with comments, footnotes and tracked changes](docs/docx.png)

## Features

**Word:** styles, fonts and formatting, lists, links, tables (merge, borders, header rows), images, page view with headers, footers and page numbers, comments, footnotes, table of contents, tracked changes (accept / reject), find and replace, word count.

**PowerPoint:** text editing on the slide, all 187 preset shapes, connectors, group / align / distribute with snapping, tables with table styles, picture crop, layouts, backgrounds, slide size, sections, speaker notes, slideshow with presenter view, find and replace, zoom.

**Both:** New presentation / New document commands, undo / redo, autosave, light and dark themes.

**Not yet:** editing SmartArt and charts, animations, recording your own edits as tracked changes.

## Install

- **Community plugins** (once listed): search "Doc Works".
- **BRAT:** add `zarnuq/DocWorks`.
- **Manual:** copy `main.js`, `manifest.json` and `styles.css` from the [latest release](https://github.com/zarnuq/DocWorks/releases) into `<vault>/.obsidian/plugins/doc-works/` and enable it.

If another plugin already opens `.docx` or `.pptx`, disable it to use Doc Works for that type.

## Contributing

Found a file that renders wrong? Open an issue with the file or a screenshot next to Office's rendering. PRs welcome; see `CLAUDE.md` for the architecture.

```bash
npm run dev    # watch build
npm test       # vitest
npm run lint   # Obsidian plugin rules
```

## License

GPL-3.0
