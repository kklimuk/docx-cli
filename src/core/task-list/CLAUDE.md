# src/core/task-list — GFM task-list read + emit + tracked toggle

Three files behind the `@core/task-list` barrel ([index.ts](index.ts)):
[detect.ts](detect.ts) is the read-side recognizer; [emit.tsx](emit.tsx) is
the `<w:sdt><w14:checkbox/>` primitive; [toggle.ts](toggle.ts) is the
mutate-side toolkit (accept / reject / flip). Every shape in the file was
**validated empirically against Microsoft Word for Mac, Word for Web, and
LibreOffice** — see `/tmp/checkbox-track-probe/` for the AppleScript probes.

## The two recognized shapes

GFM `- [ ]` / `- [x]` doesn't have a single canonical OOXML rendering — the
producer ecosystem split. We **read both** shapes, **emit only shape 1**:

1. **SDT content control** — `<w:sdt><w:sdtPr><w14:checkbox/></w:sdtPr><w:sdtContent>…☐|☒…</w:sdtContent></w:sdt>` followed by a whitespace `<w:r>`, leading a list paragraph's content. Pandoc, LibreOffice, and Word for Mac/Windows desktop all emit this; Word for Web preserves it byte-for-byte on no-edit round-trip. `w14:checked` is the source of truth for state; the glyph in `sdtContent` is decoration that mirrors the attribute.
2. **Wingdings ☐ bullet + paragraph-mark strike** — Word for Web's *Home → Checklist* button emits a regular bulleted list whose level-0 bullet character is U+F0A8 (Wingdings ☐) per `numbering.xml`, with `<w:strike w:val="1"/>` on the paragraph-mark `<w:rPr>` marking the item done. No SDT, no `<w14:checkbox>` — the "checkbox-ness" is conveyed structurally via the bullet glyph.

`detectTaskListState` populates `paragraph.taskState` from whichever shape it sees and returns the set of nodes to skip during the run walk so the SDT glyph and the trailing space run don't leak into AST `runs`. Under tracking the next sibling may be a `<w:ins>`/`<w:del>` wrapper rather than a bare `<w:r>` — the skip logic descends one level for the leading whitespace so a tracked-insert task still strips it.

## Why we emit only shape 1

It's the format Pandoc produces, all desktop Word versions render it interactively (clickable to toggle, real tracked-change semantics), and Word for Web preserves it on no-edit round-trip. The bounded failure mode: a user opens our doc in Word for Web, types a NEW task line via the Checklist button, gets a mixed-format doc — but the original SDTs survive untouched. Reading both shapes means we still surface `taskState` for every task in that mixed doc.

## Tracked-toggle pairing

When the user toggles a `<w14:checkbox>` under track-changes, Word emits an `<w:ins>` (new glyph ☒ or ☐) and `<w:del>` (old glyph ☐ or ☒) pair INSIDE `<w:sdtContent>` AND flips the `w14:checked` attribute in place — no `<w14:checkedChange>` element exists in the spec. `findCheckboxToggle` recognizes the pair structure; the AST reader's `walkRunContainer` calls it to register a `checkboxToggle` `tcN` into `document.trackedChangeReferences`. That reader walk is the single source of `tcN` ids — `list` and `accept`/`reject` (in [src/core/track-changes/apply.ts](../../core/track-changes/apply.ts)) read the map rather than re-walking, so they can't drift.

- `acceptCheckboxToggle` keeps the ins glyph (unwrap) and drops the del. The attribute is already in the "after" state; no fix needed.
- `rejectCheckboxToggle` drops the ins, unwraps the del, renames the `<w:delText>` back to `<w:t>`, AND **flips `w14:checked` back** — inferred from the deleted glyph (☐ → "0", ☒ → "1") since Word stores no separate prior-value record. This is the only tracked-change kind in the codebase that has to infer a metadata field from its content; document it loudly if you touch it.

## `flipCheckbox{Untracked,Tracked}` — authoring side

`flipCheckboxUntracked(paragraph, checked)` updates the SDT in place (attribute + glyph). `flipCheckboxTracked(paragraph, checked, mintMeta)` writes Word's tracked-toggle shape: replaces `sdtContent`'s children with `<w:ins>new</w:ins><w:del>old</w:del>` and flips the attribute. Each side of the pair gets a fresh revision id from `mintMeta` (Word emits distinct ids — the meta-minter pattern is shared with `core/track-changes/replace.tsx`). Both return `false` on no-op (no SDT, or already in the target state) so callers can produce a clean ack.

## Structural inserts/deletes — known gap

Adding/removing an entire checkbox SDT (Word wraps the SDT in `<w:customXmlDel/InsRangeStart/End>` brackets) round-trips through the XmlNode tree but isn't enumerated as a dedicated tracked-change kind. The reader's `walkRunContainer` skips checkbox-SDTs whose only tracking is structural so no phantom `tcN` is minted for them; `accept --all` won't honor them. Bounded gap — toggle and edit-task-text cover the common cases.

## Adding a recognized shape

Extend `detectTaskListState` in [detect.ts](detect.ts) with the new pattern. Set `paragraph.taskState`, populate `skip` with any decoration nodes that shouldn't surface as runs, and (if the shape can carry tracked changes) call `registerToggle` so the AST reader registers a `tcN`. Keep the emit side single-format (SDT) unless the consumer ecosystem really demands the new shape on author.
