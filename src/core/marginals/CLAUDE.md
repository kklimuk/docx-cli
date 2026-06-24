# src/core/marginals — header/footer emit, config, the `Marginals` lens

A **marginal** is the shared abstraction over page headers and footers (the
header/footer analog of `Note` for footnote/endnote — there's no single English
word, hence the coinage). Three files behind the `@core` barrel:

- [config.ts](config.ts) — `MarginalKind = "header" | "footer"`, `MarginalType =
  "default" | "first" | "even"`, and the `MarginalConfig` lookup via
  `marginalConfig(kind)` (part-name stem, root tag `w:hdr`/`w:ftr`, reference tag
  `w:headerReference`/`w:footerReference`, locator prefix `hdr`/`ftr`, relationship
  type + content type). Leaf data, same shape as `noteConfig`. Plus the part-name
  helpers `isMarginalPartName` / `marginalPartNameFromTarget`.
- [text.ts](text.ts) — `marginalText(tree)` + `fieldToken(instr)`: the read-side
  extraction that turns a `<w:hdr>`/`<w:ftr>` tree into `Marginal.text`, rendering
  `<w:fldSimple>` fields as `{page}` / `{pages}` / `{date}` / `{time}` /
  `{styleref:NAME}` / `{filename}` / `{title}` / `{author}` tokens (not the cached
  value). `{time}` is read-only (no authoring flag mints it). Pure, no JSX.
- [index.tsx](index.tsx) — the **`Marginals` lens** (`new Marginals(document).set(…)`
  / `.clear(…)`) + the part-body emitters (`<w:hdr>`/`<w:ftr>` root, the content
  paragraph incl. the two-zone right-tab path, and the `fldSimple` field vocabulary).

## View (state) + lens (operation)

Header/footer **part trees** are state, owned by [`MarginalsView`](../ast/document/marginals.ts)
(embedded on `Document`). Unlike the single-part views, ONE `MarginalsView` owns
MANY parts keyed by part name (`word/header1.xml`, `word/footer1.xml`, …);
`fromPackage` scans `pkg.listParts()` for them, `writeTo` serializes all. The AST
reader reads part text through `partTree`; the lens writes through `setPart` +
`nextPartName`. There's no `register` (the multi-part analog) — the lens mints each
part's relationship + content-type as it allocates the part.

The **operation** is the `Marginals` lens, constructed at the call site like
`Images`. `set` reaches through `Document` into the `MarginalsView`, the
relationships + content-types views (part registration), the settings view (the
even/odd toggle), and the live `<w:sectPr>` nodes (the per-section references).

## How a header/footer is wired

Per-section references live in `word/document.xml`'s `<w:sectPr>`s as
`<w:headerReference w:type=… r:id=…>` / `<w:footerReference …>` — FIRST in
CT_SectPr order (splice via `insertSectPrChildInOrder`, never `push`). One part can
be referenced from MANY sections (document-wide = one part, one rId, N references),
so `set` with no `--at` reuses a single part across every section.

- `first` type also needs `<w:titlePg/>` on the section (or Word ignores it).
- `even` type also needs the document-level `<w:evenAndOddHeaders/>` in
  `settings.xml` (`SettingsView.ensureEvenAndOddHeaders`). There is NO `odd` type —
  `default` IS the odd-page marginal once even exists.

## Word validity

- Part root declares `xmlns:w` AND `xmlns:r` (a future header-body hyperlink uses
  `r:id`; an undeclared prefix is malformed XML) — like the notes parts.
- Each part needs a content-type `<Override>` + a `<Relationship>`, or Word reports
  "unreadable content."
- `<w:fldSimple>` is already a `RUN_BEARING_WRAPPER_TAG`, so no offset-bridge change.

## Tracked changes

The reference add/remove rides the existing `<w:sectPrChange>` machinery for free:
references are sectPr children, so `wrapSectPrChange` snapshots them and
`restoreSectPrSnapshot` restores them on reject. `set`/`clear` call
`wrapSectPrChange` (under `--track`) only when a section's reference actually
changes. The part-body CONTENT is authored fresh and is NOT individually
`<w:ins>`/`<w:del>`-tracked (a documented v1 limitation — honest structural
tracking is the sectPrChange; the orphaned part on reject is harmless per the
unreferenced-part invariant).

## clear leaves orphans

`clear` removes only the `<w:…Reference>`; the part + relationship + content-type
override stay as harmless orphans (the invariant-safe choice — never dangle a
referenced part). Re-setting after a clear mints a fresh part.

## Adding a field to the vocabulary

Add a variant to `MarginalField` (index.tsx), a `case` in `fieldRuns` (the
`<w:fldSimple w:instr=…>`), a token in `fieldToken` (text.ts), and a flag in
`cli/headers/set.tsx`'s `resolveField`.

## Adding a config property

Extend `MarginalConfig` in [config.ts](config.ts) and add the field to both
`HEADER_CONFIG` and `FOOTER_CONFIG`; everything picks it up via `marginalConfig(kind)`.
