# src/cli/tables — table restructuring verbs

`docx tables <verb>` reshapes an existing table: `insert-row` / `delete-row` /
`insert-column` / `delete-column` / `set-widths` / `merge` / `unmerge` /
`borders`. There is no `list` verb — `docx read --ast` already returns the full
table structure (grid widths, `gridSpan`, `vMerge`, and tracked-change markers).

## The merge-aware grid model is the foundation

[grid.ts](grid.ts) (`buildGrid`) resolves a `<w:tbl>` into logical coordinates:
each physical `<w:tc>` occupies `colSpan` logical columns from `colStart`, and a
`vMerge="continue"` cell is still a real `<w:tc>`. Every verb queries the model
(`cellAt`, `colCount`) to map a `tN:rR` / `tN:cC` / `tN:rR1cC1-rR2cC2` locator
onto the physical cells it mutates. Don't index `<w:tc>` positionally — a
spanned column has fewer physical cells than logical ones.

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

## Fresh XML and `<w:tcPr>`/`<w:tblPr>` surgery live in mutate.tsx

[mutate.tsx](mutate.tsx) holds the constructors (`emptyCell`, `gridColElement`)
and the in-place setters (`setGridSpan`, `setVMerge`, `setCellWidth`,
`setTableLayout`, `setTablePropertiesChild`). The setters splice individual
children into existing `<w:tcPr>`/`<w:tblPr>` **in CT_TcPr/CT_TblPr schema order**
(§17.4.42 / §17.4.60) without disturbing siblings — so unmodeled cell/table
properties survive (the in-place-mutation invariant). `grid.ts` stays pure
read-model (`.ts`, no JSX); construction happens in the `.tsx` files.

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

We additionally **read** (list / accept / reject) `tblPrChange` and `tcPrChange`
that *Word* authors (e.g. from an interactive width edit), even though our only
emitter of `tcPrChange` is set-widths and we never emit `tblPrChange`.

When adding a tracked structural kind, follow the "Adding a tracked-change kind"
checklist in the track-changes CLAUDE.md — register it in `core/ast/read.ts`
(table walk) **and** `apply.ts` (`visitTable`) in the same order, or `tcN` ids
drift between `list` and `accept` — and verify it against Word before trusting
the schema.
