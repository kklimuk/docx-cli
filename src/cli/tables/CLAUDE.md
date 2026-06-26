# src/cli/tables — table restructuring & formatting verbs

`docx tables <verb>` reshapes or formats an existing table: `insert-row` /
`delete-row` / `insert-column` / `delete-column` / `set-widths` / `merge` /
`unmerge` / `borders` (restructure), plus `format` (cell/row/table-level
formatting — shading, vertical/horizontal alignment, per-cell borders, row
height, repeat-header, table alignment, table style). There is no `list` verb —
`docx read --ast` already returns the full table structure (grid widths,
`gridSpan`, `vMerge`, `vAlign`, `shading`, cell `borders`, row `height`/
`repeatHeader`, table `align`/`style`, and tracked-change markers).

## `format` is locator-scoped, not flag-scoped

`tables format --at LOCATOR` picks the targets by the locator's granularity:
cell properties (`--shade`/`--valign`/`--halign`/`--cell-borders`) broadcast over
EVERY cell the locator covers — a cell, a `rRcC-rRcC` range, a `cC` column, a
`rR` row, or the whole `tN` table — so "shade the whole table" is one call that
writes per-cell `<w:shd>` (table-level `<w:shd>` isn't read back; per-cell is, via
`docx:cell shading`). `--align`/`--style` need `--at tN` (whole table);
`--row-height`/`--repeat-header` need a row (`--at tN:rR`) or `--at tN` (all rows).
The cell set is the merge-aware grid's PHYSICAL cells, deduped by node — a
`gridSpan` cell covered by a range/column is touched once. `--halign` is a
PARAGRAPH property (`<w:pPr><w:jc>`), so it fans `Edit.paragraphProperties` over
the cells' paragraphs (NOT a `<w:tcPr>` child); `--valign` (`<w:tcPr><w:vAlign>`)
is the true cell property. The two are deliberately distinct from `--align`
(`<w:tblPr><w:jc>`, the table on the page).

The model and mutation primitives these verbs build on live in
[`@core/table`](../../core/table/): the merge-aware grid (`buildGrid`,
`cellAt`), the `<w:tcPr>`/`<w:tblPr>` surgery (`setGridSpan`, `setVMerge`,
`emptyCell`, `appendTblGridChange`, …), and the `resolveTableNode` /
`parseTableAt` / `parseRowAt` / `parseColumnAt` / `parseCellRangeAt` /
`parseCellAt` locator helpers. Every locator helper accepts the chained nested
form (`t0:r0c1:t0`, `t0:r0c1:t0:r1`, …) so every verb addresses nested tables
with the same syntax. This folder owns only the CLI surface: arg-parse, the
merge-correctness gates ("don't corrupt a merge", below), `--dry-run`/`--output`
glue, and `noteStructuralChange` (the audit-comment policy for changes Word
won't round-trip).

## Structural edits never corrupt a merge

`insert-column` rejects a position that **bisects** a horizontal span;
`delete-column` rejects a column that **passes through** one; `delete-row`
rejects a row holding the **restart** of a vMerge that would orphan its
continuations; `merge` rejects a region whose edges **cross** an existing merge.
Those point the agent at `tables unmerge` first. `delete-row`/`delete-column`
also refuse to empty the table (delete the last row/column → use `docx delete
tN`).

**`insert-row` is the exception — it extends rather than rejects.** A row
inserted inside a vertical merge gets a `vMerge="continue"` cell in each merged
column so the merge grows through the new row; a row inserted below the merge is
a normal row. This matches Word's UI behavior exactly (verified empirically: a
Word UI row-insert inside a vmerge extends it, and our output round-trips
through Word unchanged), and is why `insert-row` builds its row via `buildRow`
(consulting the row now below) rather than a flat band of empty cells.

`buildRow` also inherits each new cell's paragraph **properties** from the
reference (sibling) row's cell — it clones that cell's first `<w:p>`'s `<w:pPr>`
(alignment/spacing/indent) via `cellParagraph`. Without this the new row copied
only the gridSpan/merge *structure* and the cells fell back to the left-aligned
default, so an inserted numeric column sat visibly out of column with the
right-aligned rows above (the invoice "Calibration kit" row defect). `<w:pPr>` is
pure properties (no content), so cloning it wholesale is safe; cell-level
`<w:tcPr>` (shading/vAlign) is deliberately NOT inherited — striped/alternating
rows make that unsafe to assume.

## Track-changes: native where a construct exists, audit comment otherwise

See [src/cli/track-changes](../track-changes/CLAUDE.md) for the lifecycle. The
split below was settled **empirically** by driving Microsoft Word (AppleScript
`accept all revisions` / `reject all revisions`) and comparing its result to
ours — Word, not the ECMA-376 schema, is the arbiter of what actually
round-trips. ECMA-376 defines more table-revision markup than Word honors.

- **Rows** → native `<w:trPr><w:ins>` / `<w:del>` (kinds `rowIns`/`rowDel`).
  Word authors and round-trips these.
- **Columns** → per-cell `<w:tcPr><w:cellIns>` / `<w:cellDel>` (`cellIns`/
  `cellDel`); `insert-column` also emits a paired `<w:tblGridChange>` so the
  width growth is reversible. `delete-column` leaves the grid intact — the
  column is trimmed by the accept-time **grid resync** in `apply.ts`
  (`resyncTableGrids`). Word's UI doesn't author `cellIns`/`cellDel`, but it
  **reads and accept/rejects** the ones we author (verified).
- **set-widths** → `<w:tblGridChange>` (grid snapshot) **plus a per-cell
  `<w:tcPrChange>`** (each `<w:tcW>`). Both are needed: Word's reject is driven
  by the per-cell `tcPrChange` — a `tblGridChange` alone is *not* reverted by
  Word. This mirrors exactly what Word emits for a width change.
- **merge / unmerge / borders** → applied **immediately** + a `[docx-cli]` audit
  comment via `noteStructuralChange` ([common.ts](common.ts)). Verified: Word
  does **not** track cell merges/splits (it warns "this action won't be marked
  as a change" and applies a plain `<w:gridSpan>`), and it does **not** revert a
  hand-authored `<w:tblPrChange>` for borders. Authoring a revision Word won't
  round-trip would be dishonest, so we match Word and just note it.
- **format** → split by the level the property lives at, since that decides
  whether Word round-trips a revision:
  - cell `<w:tcPr>` props (`--shade`/`--valign`/`--cell-borders`) → native
    **`<w:tcPrChange>`**, one per cell, exactly like `set-widths` (snapshot the
    prior `<w:tcPr>` BEFORE mutating, then `appendTcPrChange`). Surfaces as the
    `cell-fmt` tcN kind; accept drops the snapshot, reject restores it.
  - `--halign` → native **`<w:pPrChange>`** per cell paragraph, via
    `Edit.paragraphProperties` (the same revision `edit --alignment` emits — Word
    authors and round-trips it).
  - table `<w:tblPr>` props (`--align`/`--style`) and row `<w:trPr>` props
    (`--row-height`/`--repeat-header`) → applied **immediately** + ONE `[docx-cli]`
    audit comment per invocation (Word reverts neither a hand-authored
    `<w:tblPrChange>` nor `<w:trPrChange>` — same reasoning as `borders`).

We additionally **read** (list / accept / reject) `tblPrChange` and `tcPrChange`
that *Word* authors (e.g. from an interactive width edit). Our `tcPrChange`
emitters are `set-widths` and `format`; we never emit `tblPrChange`/`trPrChange`.

When adding a tracked structural kind, follow the "Adding a tracked-change kind"
checklist in the track-changes CLAUDE.md — register it in `core/ast/read.ts`
(table walk) **and** `apply.ts` (`visitTable`) in the same order, or `tcN` ids
drift between `list` and `accept` — and verify it against Word before trusting
the schema.
