# src/core — `Document`, embedded views, lenses, emitters

`src/core` uses relative imports between siblings (CLI code uses the `@core/*` aliases). JSX is for emitters only: any file that constructs fresh XML can be `.tsx`; readers, locators, and pure analysis stay `.ts` — never JSX in the AST reader.

## Document, embedded views, cross-cutting lenses

`Document` (the composition root, at [ast/document/index.ts](ast/document/index.ts)) holds one **tree-owning view per OPC part** as a field — see [ast/CLAUDE.md](ast/CLAUDE.md) for the list and the reader. **Cross-cutting concerns are lenses** under the feature folder ([insert/](insert/), [edit/](edit/), [image/](image/), [hyperlinks/](hyperlinks/), [equation/](equation/), [track-changes/](track-changes/), [comments/](comments/)) — stateless classes constructed at the call site (`new Insert(document).paragraph(blockRef, spec, …)`, `new Edit(document).range(rangeRef, spec, …)`, `new Images(document).add(source)`, `new TrackChanges(document).accept([...])`) that reach through Document into multiple embedded views. The split — tree-owning views ARE state (embedded), lenses are operations over state (constructed) — is the same one stated in the root [CLAUDE.md](../../CLAUDE.md). Cross-view deps are passed as method args (`NotesView.ensureNoteStyles(stylesView)`), never as constructor or field references.

## In-place XML mutation, not AST round-trip

The typed AST from `read` is a *view* over the parsed XML tree. Mutations operate on the underlying `XmlNode` references via `BlockReference.parent.splice(...)`. Anything we don't model survives because we never re-emit untouched regions. Only emit fresh XML for nodes we're inserting (via JSX) — never round-trip whole subtrees through the AST.

## RUN_BEARING_WRAPPER_TAGS — the AST↔XML offset bridge

Defined in [parser/run-ops.ts](parser/run-ops.ts). AST text and `find`'s offsets descend into every tag in this set; the XML-side walkers in [comments/markers.tsx](comments/markers.tsx), [find/replace-span.tsx](find/replace-span.tsx), and [hyperlinks/wrap.tsx](hyperlinks/wrap.tsx) all do the same via `isRunBearingWrapper(tag)` / `sumRunBearingTextLength(children)`. They must stay in sync — if the AST descends into a wrapper the XML walkers don't (or vice versa), `find → replace` / `find → comments add` misaligns by the wrapper's inner-text length. Current set: `<w:ins>`, `<w:del>`, `<w:moveFrom>`, `<w:moveTo>`, `<w:hyperlink>`, `<w:fldSimple>`, `<w:smartTag>`. Any tag NOT in the set is preserved by the catchall `push(child)` in every walker (see `tests/cli/invariants.test.ts`).

**Adding a run-bearing wrapper:** add to `RUN_BEARING_WRAPPER_TAGS` in `parser/run-ops.ts` and recurse into it in `walkRunContainer` in `ast/read.ts` — that's the only edit (every offset-aware walker reads the predicate). Add a regression test in `tests/cli/invariants.test.ts` (unmodeled-XML survival + transparent wrappers).

## fast-xml-builder owns escaping

On the JSX path, never manually escape — the builder handles entities (uses `&apos;` for `'`, which fast-xml-parser decodes back).

## Adding an OOXML tag

Add it to the appropriate namespace's tag list in [jsx/index.ts](jsx/index.ts). The mapped-type pattern (`namespace("w", W_TAGS)`) survives `noUncheckedIndexedAccess` because the keys are a literal-string union.

## Adding an AST field

See [ast/CLAUDE.md](ast/CLAUDE.md) — `types.ts` is the source, `read.ts` populates, `cli/info/schema.ts` widens.

## Sections

`<w:sectPr>` surfaces as `SectionBreak` blocks (`sN`) with optional `columns` and `sectionType` (`continuous`/`nextPage`/`evenPage`/`oddPage`/`nextColumn`). Per ECMA-376 §17.6.22 `<w:type>` describes where the **current** section *begins*, not where the next starts. Trailing `<w:sectPr>` (mandatory) and inline `<w:pPr><w:sectPr>` (defines the section ENDING at that paragraph) both enumerate; the inline case puts an extra `sN` block right after its owning `pN`.

CRUD: `insert --after pN --section` emits a sentinel paragraph carrying an inline sectPr; `edit --at sN [--columns N] [--type T]` mutates the targeted sectPr in place; `delete --at sN` strips the inline sectPr (paragraph stays) and rejects on the trailing one. Helpers in [sections.tsx](sections.tsx). Under tracking, `edit --at sN` emits a real `<w:sectPrChange>` snapshot (accept removes it; reject restores its children); `track-changes list` enriches sectPrChange entries with `prior`/`current` — `read --accepted/--baseline` only switches RUN-level views, so the enriched list is the only way to see prior section state.

**Adding a section property:** extend `SectionBreak` in `ast/types.ts` and `readSectionProperties` in [sections.tsx](sections.tsx); widen `cli/info/schema.ts`; add an `applyXxx` mutator alongside `applyColumns`/`applySectionType`; thread the flag through `parseSectionFlags` in both `cli/insert/index.tsx` and `cli/edit/index.tsx`. `wrapSectPrChange` already snapshots all sectPr children, so the property round-trips through tracking for free.

## Styles & numbering

`document.ensureStyles().ensureStyle(id)` ([ast/document/styles.tsx](ast/document/styles.tsx)) lazily provisions `word/styles.xml` from the `BASELINE` catalog if it isn't there yet. **Adding a baseline style:** extend the `BaselineStyleId` union and add a JSX-built definition to `BASELINE` — pStyle children must follow ECMA-376 §17.3.1.26 order (pStyle → keepNext → keepLines → numPr → pBdr → spacing → ind → jc). New tags go in `W_TAGS`. `insert --style`/`edit --style` auto-define baseline styles via `StylesView.ensureReferencedStyle`; custom values are referenced but left undefined.

`document.ensureNumbering().allocate("bullet" | "ordered")` ([ast/document/numbering.tsx](ast/document/numbering.tsx)) lazily provisions `word/numbering.xml`. **Adding a format:** extend `AbstractNumKind` and add an abstractNum builder; `ensureAbstractNum` reuse keys on `lvl[0]/numFmt/w:val`, so give each kind a distinct level-0 numFmt. Cover it in `tests/core/numbering.test.ts`.

## Task lists (GFM checkboxes)

Task lists are their own package — see [task-list/CLAUDE.md](task-list/CLAUDE.md). The reader recognizes two shapes (SDT content control + Word-for-Web's Wingdings-bullet + strike model); we emit only the SDT shape. The `taskState` prop is plumbed through `Paragraph` in [blocks.tsx](blocks.tsx) which prepends the `<TaskCheckbox>` primitive from `@core/task-list`. Tracked toggles surface as the `checkboxToggle` `TrackedChangeKind` and have accept/reject machinery that infers the prior `w14:checked` value from the deleted glyph.

## Tables

Tables are their own package — see [table/CLAUDE.md](table/CLAUDE.md).

## Images

Images are their own package — see [image/CLAUDE.md](image/CLAUDE.md).

## Equations

Equations are their own package — see [equation/CLAUDE.md](equation/CLAUDE.md). The reader walks `<m:oMath>` / `<m:oMathPara>` to reconstruct LaTeX (`run.latex` on `EquationRun`); markdown render emits `$…$` / `$$…$$`. The writer (`latexToOmml`, used by `insert --equation` and `edit --at eqN --equation`) goes LaTeX → MathML via temml → OMML via our own adapter. The walker tolerates both Pandoc's lean OMML and Word's verbose `<m:ctrlPr>`-laden output. Tracked equation edits land via the same paragraph-level `<w:ins>` / `<w:del>` wrappers we use for text runs — `TrackChanges.applyInsertion`, `TrackChanges.applyDeletion`, and `Equations.edit` all treat `<m:oMath>` / `<m:oMathPara>` as trackable run-level siblings.
