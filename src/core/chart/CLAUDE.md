# src/core/chart — native Word charts (read-only, first slice of #43)

A `<w:drawing>` holding `<c:chart r:id>` is a native chart: the reader mints it a positional `chartN` (`WalkState.chartIndex`, document order over the body, table cells and text boxes — headers/footers aren't walked yet) and fills a `NativeChartRun` from the chart part.

- [read.ts](read.ts) — `readChartContent(part, partName)`: the plot area's chart groups (`chartTypes`, through any `mc:AlternateContent`), the title text, every `c:ser` from its CACHES (`strCache`/`numCache`/`multiLvlStrCache`/`strLit`/`numLit`, padded to `ptCount`, a missing `c:pt` = `null`, never 0), and the data source (`c:externalData r:id` → the chart part's rels → `embedded` workbook part / `external` link / `none`). The caches are what Word and LibreOffice draw; the workbook is located, never opened.
- [charts.ts](charts.ts) — the `Charts` lens (`list()` / `get(chartN)`) over the AST; `chartTypeName` / `chartPointCount` are shared with the `docx:chart` read hint so the hint, `charts list` and `charts extract` can't disagree.
- [csv.ts](csv.ts) — `chartToCsv`: `category,<series…>` header + one row per point, RFC 4180 quoting. This is the shape a future `charts set-data --csv` should take back in.

The chart parts come from `ChartsView` ([../ast/document/charts.ts](../ast/document/charts.ts)), preloaded because the body walk is synchronous. Nothing here mutates a chart: there is no `writeTo`, so an unrelated save keeps the chart part, its rels and its workbook byte-identical. A write path (set-data, add, delete) must update caches AND the embedded workbook together — both engines render the cache, so a cache-only edit looks right until someone clicks Edit Data.
