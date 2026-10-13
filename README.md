# Foolscap

Spreadsheet and word processor that run in the browser. **Sheets** is working and
reads/writes real `.xlsx`; **Docs** is a placeholder.

Named after the paper size, because it is paper either way.

```bash
npm install
npm run dev      # http://localhost:5173
npm test         # 43 unit tests
npm run build    # typecheck + production bundle
```

## Stack

| Concern | Choice | Why |
| --- | --- | --- |
| UI | React 19 + TypeScript + Vite | Fast HMR, no framework overhead |
| Formula engine | [HyperFormula](https://hyperformula.handsontable.com/) | 390+ Excel functions, dependency graph, incremental recalc, built-in undo/redo |
| File I/O | [ExcelJS](https://github.com/exceljs/exceljs) | Reads and writes real `.xlsx` with styles |
| State | Zustand | Plain store, no provider tree |
| Grid | Hand-written, virtualised | Needed for Excel-specific UX (fill handle, spill, range drag) that off-the-shelf grids fight you on |

## License

**GPL-3.0-or-later.** Not a free choice — HyperFormula is dual-licensed GPLv3 or
commercial, and [`src/sheets/engine/engine.ts`](src/sheets/engine/engine.ts) uses
the `gpl-v3` key. GPLv3 is copyleft, so anything distributed that bundles it has
to be GPLv3-compatible. A permissive license on this repo would be incorrect
while that engine is in the bundle.

Everything else here is MIT or Apache-2.0. To relicense permissively, swap the
formula engine for an MIT one — `@formulajs/formulajs` plus a hand-rolled
dependency graph, or `fast-formula-parser` — or buy a HyperFormula commercial
key. The engine wrapper is deliberately thin to keep that swap cheap.

Microsoft Excel and Word are trademarks of Microsoft. This project is not
affiliated with Microsoft; it reads and writes the `.xlsx` format, which is an
open ECMA/ISO standard.

## Architecture

```
src/
  app/                 shell: title bar, Sheets/Docs switch
  shared/              types, A1-reference helpers
  sheets/
    engine/
      engine.ts        HyperFormula wrapper: read/write, input coercion, ref offsetting
      format.ts        value -> display string, Excel serial dates
    grid/
      geometry.ts      cumulative offset axes, binary-search hit testing
      Grid.tsx         virtualised grid, selection, keyboard, fill handle, resize
      CellEditor.tsx   in-cell editor
    io/xlsx.ts         .xlsx import/export, CSV export (lazy-loaded chunk)
    ui/                ribbon, formula bar, sheet tabs, status bar
    store.ts           zustand store: selection, styles, structure, clipboard
```

**Split of responsibility.** HyperFormula owns values, formulas and
recalculation. The store owns everything presentational — styles, column widths,
row heights, frozen panes, sheet metadata. The two are joined by a `rev`
counter: the engine mutates outside React, so every engine write bumps `rev` and
the grid re-reads it.

## What works

- **Grid** — 5000 x 200 cells, virtualised. Column/row resize (double-click a
  column edge to autofit), frozen panes, text spill into empty neighbours.
- **Formulas** — the full HyperFormula library (`SUM`, `VLOOKUP`, `IF`, `INDEX`,
  `XLOOKUP`, …), cross-sheet refs (`Sheet1!A1`), `#DIV/0!`-style errors with the
  reason on hover, autocomplete in the formula bar.
- **Editing** — in-cell and formula-bar editing, fill handle drag, `Ctrl+D` fill
  down, copy/cut/paste with relative-reference rewriting, TSV paste from Excel or
  Sheets, undo/redo.
- **Formatting** — bold/italic/underline/strike, font size, text and fill
  colours, horizontal alignment, wrap, number formats (number, currency,
  percent, date, text).
- **Structure** — insert/delete rows and columns (values, formulas, styles and
  sizes all shift together), multiple sheets with rename and delete.
- **Files** — open `.xlsx` (values, formulas, styles, widths, frozen panes),
  save `.xlsx`, export CSV.
- **Status bar** — live count/sum/average/min/max over the selection.

## Keyboard

| Key | Action |
| --- | --- |
| Arrows / `Tab` / `Enter` | Move | 
| `Shift`+Arrows | Extend selection |
| `Ctrl+A` | Select all |
| `F2` / `Enter` / type | Edit cell |
| `Esc` | Cancel edit, or clear the copy marquee |
| `Delete` | Clear contents |
| `Ctrl+C/X/V` | Copy / cut / paste |
| `Ctrl+Z` / `Ctrl+Y` | Undo / redo |
| `Ctrl+B/I/U` | Bold / italic / underline |
| `Ctrl+D` | Fill down |
| `Ctrl+Home` | Go to A1 |
| `PageUp/Down` | Page scroll |

## Known gaps

- No charts, pivot tables, conditional formatting, data validation or filters.
- No merged cells (the `merges` field exists in the model but nothing reads it).
- `Ctrl+Arrow` does not jump to the data edge.
- Cell styles are per-cell records; a whole-column format is stored one cell at a
  time, so formatting a select-all is clipped to the used range.
- `.xlsx` import drops charts, images, pivot tables and conditional formats, and
  shared-formula cells import as their cached value rather than a live formula.
- Nothing persists across a reload — there is no autosave, only explicit export.

## Next: Docs

Planned shape, mirroring the Sheets split:

- Rich-text document model (ProseMirror or Lexical for the editing surface)
- Paragraph styles, headings, lists, tables
- Page layout with a print view and pagination
- `.docx` round-trip via `docx` (write) and `mammoth` (read)
- Reuse `src/app/` shell, the ribbon pattern, and the status bar
