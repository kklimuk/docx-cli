# src/core/table — table emit, grid model, mutation primitives

Three files behind the `@core/table` barrel ([index.tsx](index.tsx)): the file itself holds the `<BlankTable>` / `<Table>` / `<TableRow>` / `<TableCell>` emitters; [grid.ts](grid.ts) is the pure merge-aware read-model (`buildGrid`, `cellAt`, `resolveTableNode`); [mutate.tsx](mutate.tsx) is the `<w:tcPr>` / `<w:trPr>` / `<w:tblPr>` surgery (`setGridSpan`, `setVMerge`, `setCellWidth`, `setCellShading`, `setCellVAlign`, `setCellBorders`, `setRowHeight`, `setRepeatHeader`, `setTableLayout`, `setTableJustification`, `setTableStyle`, `setTablePropertiesChild`, `emptyCell`, `gridColElement`, `appendTblGridChange`, `appendTcPrChange`, `markRowTracked`, `markCellTracked`, `clearCellContent`). The three `set*Child` splice helpers keep CT_TcPr / CT_TrPr / CT_TblPr child order.

`@core/table` is what `cli/tables/` and `cli/insert --table` build on. The CLI verbs there are thin glue — arg-parse + merge-correctness gates — over this folder's primitives.

## Grid is the foundation; never index `<w:tc>` positionally

`buildGrid` resolves a `<w:tbl>` into logical coordinates: each physical `<w:tc>` occupies `colSpan` logical columns from `colStart`, and a `vMerge="continue"` cell is still a real `<w:tc>` in its row. Every consumer queries the model (`cellAt`, `colCount`) to map a `tN:rR` / `tN:cC` / `tN:rR1cC1-rR2cC2` locator onto physical cells. Don't index `<w:tc>` positionally — a spanned column has fewer physical cells than logical ones.

`grid.ts` stays pure `.ts` (no JSX, no fresh-XML construction). Construction lives in the `.tsx` files so the read-only model is reusable without dragging in the JSX runtime.

## mutate.tsx splices in CT_TcPr / CT_TblPr schema order

The setters add/remove a *single* child (`gridSpan` / `vMerge` / `tcW` / `tblBorders` / `tblLayout`) inside an existing `<w:tcPr>` or `<w:tblPr>` at the position ECMA-376 §17.4.42 (CT_TcPr) / §17.4.60 (CT_TblPr) demands, leaving siblings untouched. That's how unmodeled cell/table properties (borders, shading, vAlign, custom styles) survive — per the in-place-mutation invariant. `pruneEmptyTcPr` deletes the whole `<w:tcPr>` if a setter empties it, so a cell that loses its last property doesn't carry an empty wrapper.

## Adding a cell property

Extend `TableCell` in `ast/types.ts`, populate it in `readTableCell` in `ast/read.ts`, widen the cell block in `cli/info/schema.ts`, then emit it in `TableCellProperties` (in [index.tsx](index.tsx)) — that component drives `<w:tcPr>` child order, so insert your new tag at the right slot per §17.4.42. If it's also mutable post-creation, add a `setXxx(cell, value)` to [mutate.tsx](mutate.tsx) following `setCellWidth`'s pattern (`ensureTcPr` → `setTcPrChild` with the tag in `TC_PR_ORDER` → `pruneEmptyTcPr`).

## Locator helpers live in core/locators

The `parseTableAt` / `parseRowAt` / `parseColumnAt` / `parseCellRangeAt` / `parseCellAt` helpers that every `docx tables` verb uses to validate `--at` strings live in [../locators/resolve.ts](../locators/resolve.ts) (re-exported from `@core/locators` / `@core`), not here. They share the recursive cell-chain shape with `locatorToBlockTarget`, so they're not table-specific in mechanics — just in current consumers. They unwrap chained nested forms (`t0:r0c1:t0`, `t0:r0c1:t0:r1`, …) so every consumer addresses nested tables with the same syntax.
