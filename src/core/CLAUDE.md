# src/core — emitters, AST reader, mutators

`src/core` uses relative imports between siblings (CLI code uses the `@core/*` aliases). JSX is for emitters only: any file that constructs fresh XML can be `.tsx`; readers, locators, and pure analysis stay `.ts` — never JSX in the AST reader.

## In-place XML mutation, not AST round-trip

The typed AST from `read` is a *view* over the parsed XML tree. Mutations operate on the underlying `XmlNode` references via `BlockReference.parent.splice(...)`. Anything we don't model survives because we never re-emit untouched regions. Only emit fresh XML for nodes we're inserting (via JSX) — never round-trip whole subtrees through the AST.

## RUN_BEARING_WRAPPER_TAGS — the AST↔XML offset bridge

Defined in [parser/run-ops.ts](parser/run-ops.ts). AST text and `find`'s offsets descend into every tag in this set; the XML-side walkers in `cli/comments/helpers.tsx`, `cli/replace/replace-span.tsx`, and `cli/hyperlinks/wrap.tsx` all do the same via `isRunBearingWrapper(tag)` / `sumRunBearingTextLength(children)`. They must stay in sync — if the AST descends into a wrapper the XML walkers don't (or vice versa), `find → replace` / `find → comments add` misaligns by the wrapper's inner-text length. Current set: `<w:ins>`, `<w:del>`, `<w:moveFrom>`, `<w:moveTo>`, `<w:hyperlink>`, `<w:fldSimple>`, `<w:smartTag>`. Any tag NOT in the set is preserved by the catchall `push(child)` in every walker (see `tests/cli/preserve-unknown.test.ts`).

**Adding a run-bearing wrapper:** add to `RUN_BEARING_WRAPPER_TAGS` in `parser/run-ops.ts` and recurse into it in `walkRunContainer` in `ast/read.ts` — that's the only edit (every offset-aware walker reads the predicate). Add a regression test in `tests/cli/preserve-unknown.test.ts` or `tests/cli/transparent-wrappers.test.ts`.

## fast-xml-builder owns escaping

On the JSX path, never manually escape — the builder handles entities (uses `&apos;` for `'`, which fast-xml-parser decodes back).

## Adding an OOXML tag

Add it to the appropriate namespace's tag list in [jsx/index.ts](jsx/index.ts). The mapped-type pattern (`namespace("w", W_TAGS)`) survives `noUncheckedIndexedAccess` because the keys are a literal-string union.

## Adding an AST field

Add to [ast/types.ts](ast/types.ts), populate in [ast/read.ts](ast/read.ts), then update `cli/info/schema.ts`. The `--ts` output reads `types.ts` live via Bun's text import, so it stays in sync.

## Sections

`<w:sectPr>` surfaces as `SectionBreak` blocks (`sN`) with optional `columns` and `sectionType` (`continuous`/`nextPage`/`evenPage`/`oddPage`/`nextColumn`). Per ECMA-376 §17.6.22 `<w:type>` describes where the **current** section *begins*, not where the next starts. Trailing `<w:sectPr>` (mandatory) and inline `<w:pPr><w:sectPr>` (defines the section ENDING at that paragraph) both enumerate; the inline case puts an extra `sN` block right after its owning `pN`.

CRUD: `insert --after pN --section` emits a sentinel paragraph carrying an inline sectPr; `edit --at sN [--columns N] [--type T]` mutates the targeted sectPr in place; `delete --at sN` strips the inline sectPr (paragraph stays) and rejects on the trailing one. Helpers in [sections.tsx](sections.tsx). Under tracking, `edit --at sN` emits a real `<w:sectPrChange>` snapshot (accept removes it; reject restores its children); `track-changes list` enriches sectPrChange entries with `prior`/`current` — `read --accepted/--baseline` only switches RUN-level views, so the enriched list is the only way to see prior section state.

**Adding a section property:** extend `SectionBreak` in `ast/types.ts` and `readSectionProperties` in [sections.tsx](sections.tsx); widen `cli/info/schema.ts`; add an `applyXxx` mutator alongside `applyColumns`/`applySectionType`; thread the flag through `parseSectionFlags` in both `cli/insert/index.tsx` and `cli/edit/index.tsx`. `wrapSectPrChange` already snapshots all sectPr children, so the property round-trips through tracking for free.

## Styles & numbering

`ensureStyle(view, id)` in [styles.tsx](styles.tsx) lazily provisions styles.xml from the `BASELINE` catalog. **Adding a baseline style:** extend the `BaselineStyleId` union and add a JSX-built definition to `BASELINE` — pStyle children must follow ECMA-376 §17.3.1.26 order (pStyle → keepNext → keepLines → numPr → pBdr → spacing → ind → jc). New tags go in `W_TAGS`. `insert --style`/`edit --style` auto-define baseline styles via `ensureReferencedStyle`; custom values are referenced but left undefined.

`allocateNum(view, "bullet"|"ordered")` in [numbering.tsx](numbering.tsx) lazily provisions numbering.xml. **Adding a format:** extend `AbstractNumKind` and add an abstractNum builder; `ensureAbstractNum` reuse keys on `lvl[0]/numFmt/w:val`, so give each kind a distinct level-0 numFmt. Cover it in `tests/core/numbering.test.ts`.

## Tables

`<BlankTable>` / `<Table>` / `<TableRow>` / `<TableCell>` emitters in [table.tsx](table.tsx). **Adding a cell property:** extend `TableCell` in `ast/types.ts`, populate it in `readTableCell` in `ast/read.ts`, widen the cell block in `cli/info/schema.ts`, and emit it in `TableCellProperties` — that component drives `<w:tcPr>` order, so respect ECMA-376 §17.4.42 (CT_TcPr). Row/column structural mutations (insert-row, delete-col, merge, unmerge) belong to the `docx tables` verbs in `src/cli/tables/`, not `insert/edit`.

## Images

Images are their own package — see [image/CLAUDE.md](image/CLAUDE.md).
