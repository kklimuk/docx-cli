# src/core/table — table emit, grid model, mutation primitives

Four files behind the `@core/table` barrel ([index.tsx](index.tsx)): the file itself holds the `<BlankTable>` / `<Table>` / `<TableRow>` / `<TableCell>` emitters; [grid.ts](grid.ts) is the pure merge-aware read-model (`buildGrid`, `cellAt`, `resolveTableNode`); [cell-content.ts](cell-content.ts) owns direct-cell block invariants (sole/empty paragraph detection, empty-paragraph reuse, start/end insertion via `applyCellInsertion`, the multi-insert `CellInsertionCursor`, `cellInsertionAnchor`, mandatory trailing paragraph repair); [mutate.tsx](mutate.tsx) is the `<w:tcPr>` / `<w:trPr>` / `<w:tblPr>` surgery (`setGridSpan`, `setVMerge`, `setCellWidth`, `setCellShading`, `setCellVAlign`, `setCellBorders`, `setRowHeight`, `setRepeatHeader`, `setTableLayout`, `setTableJustification`, `setTableStyle`, `setTablePropertiesChild`, `emptyCell`, `gridColElement`, `appendTblGridChange`, `appendTcPrChange`, `markRowTracked`, `markCellTracked`, `clearCellContent`). The three `set*Child` splice helpers keep CT_TcPr / CT_TrPr / CT_TblPr child order.

`@core/table` is what `cli/tables/` (including `tables create`) builds on. The CLI verbs there are thin glue — arg-parse + merge-correctness gates — over this folder's primitives.

## One insert vs. many: `applyCellInsertion` and `CellInsertionCursor`

`applyCellInsertion` places blocks at a cell boundary by re-deriving that
boundary from the cell's CURRENT blocks. That is already right for a single
insert (`docx insert --at CELL`), for every APPEND in a batch (the cell's last
block IS the one the previous entry appended), and for the FIRST prepend. It is
wrong in exactly one case: a SECOND prepend, which would re-derive to the block
the first one just placed and land ahead of it, reversing JSONL order.

So `CellInsertionCursor` carries the one piece of state that case needs — a start
cursor — and delegates everything else back to `applyCellInsertion`. **That rule
lives here, not in the CLI**: `cli/insert/batch.ts` opens one cursor per cell and
calls `insert`. The cursor is a live `XmlNode` ref into the cell's child list, so
an instance is valid only during the splice phase that owns that list; a detached
ref raises a stale-reference `CellTargetError` rather than silently
repositioning. Call `ensureTerminalParagraph()` once the cell's whole batch is in
— deferred so a synthetic paragraph isn't appended and then repositioned by a
later entry.

Resist re-adding an end cursor: `tests/cli/batch.test.ts` pins both orderings
(stacked prepends, and a reuse-prepend followed by an append), and the end
boundary is derivable by construction.

`CellTargetError` sits in its own [../locators/cell-target-error.ts](../locators/cell-target-error.ts)
because both sides raise it — `locators/resolve.ts` for a refused locator,
`cell-content.ts` for a stale boundary — and `resolve.ts` already value-imports
`cell-content.ts`; homing the class in either would close an import cycle.

## Grid is the foundation; never index `<w:tc>` positionally

`buildGrid` resolves a `<w:tbl>` into logical coordinates: each physical `<w:tc>` occupies `colSpan` logical columns from `colStart`, and a `vMerge="continue"` cell is still a real `<w:tc>` in its row. Every consumer queries the model (`cellAt`, `colCount`) to map a `tN:rR` / `tN:cC` / `tN:rR1cC1-rR2cC2` locator onto physical cells. Don't index `<w:tc>` positionally — a spanned column has fewer physical cells than logical ones.

`grid.ts` stays pure `.ts` (no JSX, no fresh-XML construction). Construction lives in the `.tsx` files so the read-only model is reusable without dragging in the JSX runtime.

## mutate.tsx splices in CT_TcPr / CT_TblPr schema order

The setters add/remove a *single* child (`gridSpan` / `vMerge` / `tcW` / `tblBorders` / `tblLayout`) inside an existing `<w:tcPr>` or `<w:tblPr>` at the position ECMA-376 §17.4.42 (CT_TcPr) / §17.4.60 (CT_TblPr) demands, leaving siblings untouched. That's how unmodeled cell/table properties (borders, shading, vAlign, custom styles) survive — per the in-place-mutation invariant. `pruneEmptyTcPr` deletes the whole `<w:tcPr>` if a setter empties it, so a cell that loses its last property doesn't carry an empty wrapper.

## Adding a cell property

Extend `TableCell` in `ast/types.ts`, populate it in `readTableCell` in `ast/read.ts`, widen the cell block in `cli/info/schema.ts`, then emit it in `TableCellProperties` (in [index.tsx](index.tsx)) — that component drives `<w:tcPr>` child order, so insert your new tag at the right slot per §17.4.42. If it's also mutable post-creation, add a `setXxx(cell, value)` to [mutate.tsx](mutate.tsx) following `setCellWidth`'s pattern (`ensureTcPr` → `setTcPrChild` with the tag in `TC_PR_ORDER` → `pruneEmptyTcPr`).

## Locator helpers live in core/locators

The `parseTableAt` / `parseRowAt` / `parseColumnAt` / `parseCellRangeAt` / `parseCellAt` helpers used by `docx tables` plus top-level bare-cell `insert`/`edit` live in [../locators/resolve.ts](../locators/resolve.ts) (re-exported from `@core/locators` / `@core`), not here. `resolveCellReference` adds the conservative content-target gate: merged/grid-shifted cells reject; `resolveCellParagraphReference` requires one direct paragraph for bare-cell edit. They share the recursive cell-chain shape with `locatorToBlockTarget`. They unwrap chained nested forms (`t0:r0c1:t0`, `t0:r0c1:t0:r1`, …) so every consumer addresses nested tables with the same syntax.
