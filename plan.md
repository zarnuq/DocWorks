# XLSX editor implementation plan

Status: phases 0–4 implemented (2026-10-10), plus pictures and PDF export from phase 5. Remaining: charts, comments, conditional formatting, validation, print settings, protection, phase 6 hardening. See AGENTS.md.

## Goal

Add `.xlsx` as a third editable file type in Doc Works. Users should be able to open an existing workbook, edit cells and formulas, format ranges, manage sheets, and save it without losing unrelated workbook features. Use Google Sheets interaction patterns where they fit a local Obsidian editor, while preserving Excel file semantics.

This is a new editor comparable in scope to an entire DOCX or PPTX implementation. The shared Office package infrastructure helps with loading and saving; the spreadsheet grid, calculation engine, reference rewriting, and workbook editing are substantial new systems.

## Constraints

- No new runtime dependencies. Reuse JSZip and the existing XML, history, toolbar, font, and PDF utilities where appropriate.
- Keep spreadsheet domain logic independent of Obsidian. Follow the existing infrastructure/domain/presentation boundaries.
- Write failing tests before implementing parser, writer, calculation, and reference transformation behaviour.
- Opening, scrolling, selecting, and switching sheets must never modify or save a workbook.
- Patch only the parts affected by an edit. Preserve unknown XML, relationships, embedded assets, and unsupported features.
- Do not use a DOCX table or ProseMirror document as the spreadsheet data model.
- Do not run git commands. Leave release commands for the owner.

## Release scope

| Milestone | User-visible outcome |
| --- | --- |
| Viewer foundation | Open workbooks, navigate sheets, display cell values and cached formula results with formatting. |
| First editing release | Edit values and supported formulas, format cells, copy/paste ranges, undo/redo, and save reliably. |
| Everyday spreadsheet release | Add safe structural edits, sheet management, fill, sorting, filters, validation, and wider formula coverage. |
| Rich workbook release | Add charts, pictures, comments, conditional formatting, and print/PDF controls. |

The first editing release must include useful calculation support. Displaying imported formula caches alone is a viewer capability, not a finished spreadsheet editor.

Initially preserve rather than edit pivot tables, slicers, external data connections, advanced array formulas, unsupported chart types, and other unimplemented workbook extensions. Do not claim Excel or Google Sheets feature parity. Legacy `.xls`, binary `.xlsb`, macro-enabled `.xlsm`, encrypted workbooks, macros, and real-time collaboration are separate projects.

## Architecture

### Package layer

Create `src/infrastructure/xlsx/xlsx-workbook.ts` on top of `OoxmlPackage`:

- Discover the workbook and sheets through package relationships rather than fixed filenames.
- Read workbook metadata, worksheet parts, styles, shared strings, theme, defined names, and calculation settings.
- Keep stable sheet IDs distinct from tab order, names, and relationship IDs.
- Preserve drawings, comments, tables, external links, and extension parts even before they can be edited.
- Add `blankXlsx()` to `src/infrastructure/blank.ts`.

The current shared package layer eagerly parses all XML and serializes every parsed part on save. Benchmark it against large worksheets before committing to that approach. Add original-part retention and explicit dirty-part handling if needed, with DOCX/PPTX regression tests. A spreadsheet edit must not require serializing every untouched worksheet. Avoid changing the contract of the existing formats without corresponding tests.

Distinguish byte-identical unmodified ZIP entries from identical ZIP containers: ZIP metadata/compression can differ. Test preservation at the part level as well as semantic correctness of edited parts.

### Domain layer

Proposed modules under `src/domain/xlsx/`:

| Module | Responsibility |
| --- | --- |
| `workbook-model.ts` | Sparse cells, sheets, ranges, merges, dimensions, and source XML references. |
| `xlsx-reader.ts` / `xlsx-writer.ts` | Read modeled features and minimally patch changed XML. |
| `references.ts` | A1 parsing, absolute/mixed references, ranges, quoted sheet names, and reference transforms. |
| `styles.ts` / `number-format.ts` | Shared style records, effective formatting, and displayed values. |
| `formula-parser.ts` / `formula-evaluator.ts` | Formula AST, typed evaluation, errors, and supported functions. |
| `calculation.ts` | Dependencies, invalidation, recalculation, and formula cache status. |
| `commands.ts` / `history.ts` | Atomic edits and bounded undo/redo using affected cells/parts. |
| `clipboard.ts` | Internal range payloads, external TSV/HTML, and formula translation. |
| `sheet-edit.ts` | Sheet operations, row/column edits, merges, sort, and filters. |

Store cell input, formula, calculated/cached value, displayed text, and formatting separately. Distinguish absent cells, styled empty cells, empty strings, booleans, numeric values, errors, and formulas. Dates are numeric values interpreted using number formats and workbook date settings.

Use sparse storage. Index row heights and column widths for viewport lookup; do not allocate a rectangular array for an entire sheet. Preserve rich shared strings and avoid changing every cell that references a shared string or style when editing one cell.

### Presentation layer

Create `src/presentation/xlsx/xlsx-view.ts` with a virtualized grid, formula bar, name box, sheet tabs, toolbar, and selection/status UI.

- Start with DOM virtualization and a single active cell editor. Render only visible cells plus a small overscan region.
- Keep frozen panes, headers, selections, and the editor aligned during scrolling and resizing.
- Keep selection and formula editing state outside recycled cell DOM nodes.
- Support keyboard navigation, IME composition, clipboard operations, and accessible grid semantics from the start.
- Use Obsidian `Scope` for shortcuts that conflict with application hotkeys.
- Reuse the view autosave pattern, including pending-save flush, external modification handling, and clear save failures.
- Use command diffs for undo; avoid full-workbook snapshots per keystroke.

Integrate `doc-works-xlsx-view`, `.xlsx` extension registration, background tab naming, shutdown save, a New spreadsheet command, and folder-menu creation in `src/main.ts`.

## Implementation phases

### 0. Fixtures, compatibility contract, and performance prototype

- [ ] Add `tests/fixtures/make_xlsx_fixture.py` using dev-only tooling and small hand-authored XML cases.
- [ ] Collect fixtures exported by Excel, Google Sheets, and LibreOffice when available.
- [ ] Cover sparse cells, multiple sheets, shared/inline/rich strings, formulas and caches, shared formulas, dates, styles, merges, hidden rows/columns/sheets, freeze panes, drawings, tables, and unsupported parts.
- [ ] Add `DOC_WORKS_BOOKS` real-file coverage and extend `scripts/check-files.sh` to process XLSX originals and edited copies.
- [ ] Prototype grid scrolling and measure package loading before selecting storage/caching details.

Exit: documented baseline for load time, memory, rendering, and no-op preservation; supported/editable/preserved-only features are explicit.

### 1. Reliable workbook viewer

- [ ] Read workbook relationships, sheet metadata, cells, strings, styles, dimensions, merges, and view settings.
- [ ] Display common number, percentage, currency, date/time, and scientific formats; preserve unimplemented custom formats.
- [ ] Handle both workbook date systems and date serial edge cases without timezone shifts.
- [ ] Display formula caches without silently treating absent caches as zero.
- [ ] Implement virtualized scrolling, sheet tabs, name-box navigation, selection, frozen panes, and hidden dimensions.
- [ ] Add the Obsidian view and blank workbook integration.

Exit: real workbooks open without mutation; large sparse sheets stay responsive; unsupported content remains intact through package round trips.

### 2. Cell editing, formatting, and saving

- [ ] Implement cell/formula-bar editing, Enter/Tab/Escape, multiline input, IME, delete contents, and selection extension.
- [ ] Separate literal text from numbers and formula input, including leading apostrophes and leading zeros.
- [ ] Add atomic range edits, TSV/HTML paste, internal copy/paste, undo/redo, and find/replace.
- [ ] Add font, size, bold/italic, foreground/background, borders, alignment, wrapping, and common number formats.
- [ ] Deduplicate new styles without mutating shared records used by other cells.
- [ ] Save minimal XML changes, preserve shared-string references, and keep worksheet metadata consistent.
- [ ] Implement save debouncing, flush-on-close, failed-save recovery, and external modification conflict handling.

Exit: edit/save/reopen retains cell types and formatting, undo restores exact affected content, and untouched parts survive. Keep this an internal milestone until phase 3 is ready.

### 3. Calculation engine and first editing release

- [ ] Parse formulas into an AST; never use JavaScript `eval` or string replacement to interpret formulas.
- [ ] Support arithmetic, comparisons, concatenation, precedence, literals, ranges, absolute/mixed references, and cross-sheet references.
- [ ] Implement common functions first: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `IF`, `IFERROR`, `AND`, `OR`, `NOT`, `ROUND`, and basic text functions.
- [ ] Specify spreadsheet coercion and error behaviour, including blank cells, division by zero, missing references, lazy conditional branches, and circular references.
- [ ] Build dependency tracking and recalculate affected cells. Index range dependencies without expanding entire-column references into millions of cells.
- [ ] Read shared formulas and safely translate/de-share a group when an individual member changes.
- [ ] Preserve unsupported formulas and their caches. Show when a result is unavailable or stale, including supported formulas depending on unsupported results.
- [ ] Update calculated caches and workbook recalculation metadata on relevant edits. Handle obsolete calculation-chain data; requesting recalculation in Excel is not a substitute for local calculation.

Exit: an editable budget/tracker workbook recalculates correctly across sheets, matches reference results for supported formulas, survives save/reopen, and clearly identifies unsupported calculation.

### 4. Structural editing and everyday sheet operations

- [ ] Add/rename/reorder/duplicate/delete sheets, including defined names and references.
- [ ] Insert/delete/move rows and columns; resize, hide/unhide, and freeze/unfreeze them.
- [ ] Centralize reference transformations for formulas, names, merges, tables, validation, conditional formatting, chart sources, and drawing anchors.
- [ ] Distinguish copy translation, cut/move semantics, and insertion/deletion adjustments. Produce reference errors where appropriate.
- [ ] Block only operations that cannot safely update an affected unsupported feature; do not silently corrupt its references.
- [ ] Add merge/unmerge, fill down/right, drag fill, series fill, and paste-special values/formulas/formatting.
- [ ] Add sorting and filters with explicit header/range handling and correct row integrity.
- [ ] Add named ranges, common validation rules/dropdowns, and basic spreadsheet tables with structured references.
- [ ] Expand functions with conditional aggregation and common lookup/date/text operations, backed by compatibility fixtures.

Exit: sorting, moving, duplicating, and deleting data do not detach formulas, formatting, or related objects from their intended ranges. Each operation is one undoable command.

### 5. Rich content and printing

- [ ] Add common conditional formatting rules; preserve unsupported rule types and priorities.
- [ ] Add cell notes/comments, with separate handling for legacy notes and threaded comments.
- [ ] Add pictures and common editable charts tied to worksheet ranges.
- [ ] Reuse chart XML/data utilities where practical. Audit the current PPTX chart renderer first; horizontal bars, stacking, axes, labels, and styling are not already complete.
- [ ] Keep chart data/caches synchronized with source cells and support chart copy/paste between workbooks and presentations.
- [ ] Add print area, page orientation, margins, scaling, repeated rows/columns, manual page breaks, and PDF export.

Exit: chart changes follow cell edits, print output follows workbook settings, and edited files open in reference applications without new repair warnings.

### 6. Compatibility and release hardening

- [ ] Expand real-file regression coverage and publish a feature support matrix.
- [ ] Test keyboard-only use, screen-reader navigation, IME, zoom, themes, multiple panes, sheet switching, and reload with pending saves.
- [ ] Respect workbook/sheet protection settings; do not silently remove them to allow editing.
- [ ] Test formula errors, incomplete/corrupt packages, missing relationships, large pastes, and interrupted saves with actionable messages.
- [ ] Verify imported hyperlinks and external workbook references do not trigger automatic external data refreshes.
- [ ] Document limits on unsupported formulas, advanced workbook features, and large files.
- [ ] Update README, AGENTS.md, CLAUDE.md, release metadata, and screenshots when implementation is ready.

## Testing and acceptance gates

For each writer/calculation milestone:

1. Add failing tests, implement the behaviour, then run the focused suite.
2. Check no-op and single-edit round trips, including untouched XML/binary parts, relationship validity, schema ordering, and formula/style consistency.
3. Run `npm test`, `npm run typecheck`, and `npm run lint` with zero errors; build the plugin.
4. Run the extended real-file checks and compare original/edited workbooks in LibreOffice. Check calculated values separately from rendered PDFs. Use Excel and Google Sheets for additional compatibility checks when available; LibreOffice is not a complete oracle.
5. Test live in Obsidian using new scratch-file tabs, checking `view.file.path` before scripted edits. Verify save/reopen, undo/redo, and no saves from navigation. Close test tabs and delete scratch files afterwards.
6. Re-run DOCX/PPTX regression coverage for shared package, chart, clipboard, or PDF changes.

Use benchmarks with approximately 100,000 populated cells, many formulas, multiple sheets, and very large sparse ranges. Proposed targets on a recorded reference machine: ordinary edits visible within 50 ms, smooth scrolling near 60 fps, and cancellable/yielding long operations. These are targets to validate, not current performance claims. Measure bundle growth and memory alongside speed. If main-thread parsing/calculation exceeds the budget, add a bundled worker with a defined message protocol; do not move DOM-dependent XML code into a worker unchanged.

## Effort and sequencing

These are rough planning estimates for one experienced developer working on this codebase, including tests and compatibility work, not delivery commitments.

| Scope | Estimated focused effort |
| --- | --- |
| Phases 0–1: package foundation and useful viewer | 3–5 weeks |
| Phases 2–3: reliable editing and useful calculation | 6–10 additional weeks |
| Phase 4: structural edits and everyday tools | 5–9 additional weeks |
| Phases 5–6: rich features and broad hardening | 6–12 additional weeks |

Working estimate: 9–15 weeks for the first useful editing release, and 20–36 weeks for the broader scope above. Re-estimate after phase 0. Full Excel compatibility remains substantially larger and is not bounded by these estimates.

The highest uncertainty is formula compatibility and safe reference updates, followed by preservation of complex imported workbooks and large-sheet performance. UI controls alone are a small fraction of the work. Reusing ZIP/XML plumbing saves time, but does not make XLSX comparable to adding one feature to DOCX or PPTX.

## First implementation slice

Start with a fixture containing two sheets, styled literals, a shared string, one cached formula, a merged range, and an untouched drawing. Implement workbook discovery and a read-only virtualized view. Prove no-op preservation, opening without saving, and bounded DOM size before building the cell editor. Then implement one cell-value edit with undo and save/reopen before expanding formatting or formulas.
