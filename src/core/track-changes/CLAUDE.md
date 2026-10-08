# src/core/track-changes — `TrackChanges` lens, accept/reject machinery, tracked-range shapes

Five files behind the `@core/track-changes` barrel ([index.tsx](index.tsx)):

- [index.tsx](index.tsx) — the `TrackChanges` lens + small free helpers (`resolveAuthor`, `resolveDate`, `convertTextToDelText`)
- [apply.ts](apply.ts) — accept/reject state machine: `previewTrackedChanges`, `applyTrackedChanges`, `applyTrackedDecisions` (the mixed accept+reject pass behind `track-changes apply` — both lists resolved against the pre-mutation tree, applied in one reverse-preorder traversal via the shared `applyResolvedTargets` core), `collectTrackedChanges`, the `actionFor` dispatch table, body-side note pairing, table-grid resync
- [emit.tsx](emit.tsx) — leaf primitives: `<Ins meta>`, `<Del meta>`, `markParagraphMarkAs(paragraph, kind, meta)` (drops a self-closing ins/del marker into `<w:pPr><w:rPr>`)
- [replace.tsx](replace.tsx) — range edit/delete shapes empirically validated against Word: `applyTrackedRangeReplace`, `applyTrackedRangeDelete`, `applyUntrackedRangeReplace`, `applyUntrackedRangeDelete`, the shared `assertParagraphOnlyTrackedRange` guard + `TrackedRangeConflictError`, `applyFormattingPreservingEdit` (the word-level diff for `edit --text`), and `removeParagraphLine` — the cell-safe single-paragraph removal shared by `docx delete --at pN` and `docx edit --at pN --text ""` (a `<w:tc>`'s last paragraph is blanked, not deleted, so we never emit an invalid empty `<w:tc/>`)
- [preserve-formatting.tsx](preserve-formatting.tsx) — the LCS-based word-level diff that drives `applyFormattingPreservingEdit`: `extractOldTokens`, `tokenize`, `diffTokens`, `buildTrackedRuns`, `buildUntrackedRuns`

## TrackChanges lens API

`new TrackChanges(document).…`:
- `mintMeta(authorFlag?)` — one-shot `TrackedMeta`. For multi-revision operations call `createAllocator()` instead.
- `createAllocator()` — `{ next(): number }` seeded from `computeMaxRevisionId` (scans document.xml + footnotes.xml + endnotes.xml so a new id can't collide with one already in a note body).
- `setEnabled(on)` — toggles `<w:trackChanges/>` in settings.xml; turning on materializes the part.
- `list()` — `collectTrackedChanges(document)`; reads `document.trackedChangeReferences` (the reader's map), never re-walks.
- `preview(target, verb)` — accept/reject preview for `--dry-run`. Throws `TrackedChangeNotFoundError` on unknown id.
- `accept(target)` / `reject(target)` — apply; returns `ChangeRecord[]`. Caller saves.
- `apply(accepts, rejects)` — the `track-changes apply` finalize: accept one handle list AND reject another in ONE pre-mutation-resolved pass (no inter-call renumbering, no half-finalized file). Returns the combined `ChangeRecord[]`. Caller saves.
- `applyInsertion(paragraph, authorFlag?)` / `applyDeletion(paragraph, authorFlag?)` — wrap a freshly-built paragraph's trackable children (`w:r`, `m:oMath`, `m:oMathPara`) in `<w:ins>` / `<w:del>` and mark its paragraph break. Used by `Insert.paragraph` and `delete --at pN` under tracking. Non-trackable siblings (e.g. `<w:pPr>`) pass through at their existing positions via `wrapContiguousTrackable`.
- `applyContentInsertion(paragraph, authorFlag?)` — `applyInsertion` WITHOUT the paragraph-mark ins. Used when bare-cell insert reuses Word's mandatory empty paragraph: reject removes the new content but preserves the valid `<w:p>` container.
- `applyContentDeletion(paragraph, authorFlag?)` — `applyDeletion` WITHOUT the paragraph-mark del: wraps content in `<w:del>` but keeps the `<w:p>`. Used by `removeParagraphLine` (replace.tsx) for a table cell's last paragraph, where accept-all must leave a valid empty paragraph rather than merge the cell away.

## The reader is the single source of `tcN` ids

`document.trackedChangeReferences` (a `Map<string, TrackedChangeReference>` on Document) is populated **exclusively** by [core/ast/read.ts](../ast/read.ts) in one walk over the body, then `readNoteTrackedChanges` appends standalone note-body revisions. `apply.ts::collectTrackedChanges` reconstructs `ChangeFound` records from the map (via `changeFoundFromReference`) — it never re-walks. So `list` and `accept`/`reject --at tcN` can't disagree on what `tcN` means.

## actionFor — the dispatch table

`actionFor(target, verb)` in [apply.ts](apply.ts) is the single source for "given a kind + accept/reject, what action does this become?" Memorize the patterns:

- **Additive vs subtractive run wrappers**: `ins` / `moveTo` are additive (accept = `unwrap`, reject = `delete`). `del` / `moveFrom` are subtractive (accept = `delete`, reject = `unwrap` after `<w:delText>` → `<w:t>` rename).
- **Paragraph-mark ins/del** (a self-closing wrapper inside `<w:pPr><w:rPr>`, recognized in `changeFoundFromReference` by `parent === paragraphRef.node…rPr.children`): accept-ins = `delete` the marker; reject-ins = drop the inserted break — but **conditionally**: a PURELY-inserted paragraph (all content is `<w:ins>`/`<w:moveTo>` — a fresh `insert`, or new paragraphs 2..N of a range replace) splices out wholesale, while a paragraph that ALSO holds reject-surviving content (a `<w:del>` carrying the original text — the *transition* paragraph of a 1→N range replace — or a plain untracked run) keeps its `<w:p>` and only strips the break marker (`paragraphHasSurvivingContent` in apply.ts). Unconditional `deleteParagraph` here silently destroyed the transition's restored original on reject-all, so every N≥2 multi-block edit lost its baseline (an empty `<w:p/>` where the paragraph had been). accept-del = `merge` with next paragraph (per ECMA-376 §17.13.5.4 — the merged paragraph adopts the **NEXT** paragraph's `<w:pPr>`, since its mark is the one that survives; keeping the deleted paragraph's pPr silently stripped a following list item's `<w:numPr>` / heading style on accept, collapsing a run of clauses into one un-numbered paragraph — see `mergeParagraphWithNext`); reject-del = `delete` the marker.
- **Property-change snapshots** (`sectPrChange`, `tblGridChange`, `tblPrChange`, `tcPrChange`): accept drops the snapshot, reject restores its children to the parent property element.
- **Table-structural** (`rowIns` / `rowDel` / `cellIns` / `cellDel`): the marker lives inside `<w:trPr>` or `<w:tcPr>`; the row/cell node is carried separately on `ChangeFound.tableRow`/`tableCell`. Accept-ins/reject-del = `stripMarker` (keep the row/cell); reject-ins/accept-del = `deleteRow` / `deleteCell` (remove the whole container).
- **`checkboxToggle`**: dispatches to `acceptCheckboxToggle` / `rejectCheckboxToggle` in [task-list/toggle.ts](../task-list/toggle.ts). Reject infers the prior `w14:checked` value from the deleted glyph — see [task-list/CLAUDE.md](../task-list/CLAUDE.md).

## Reverse pre-order traversal

`applyTrackedChanges` walks targets `[...targets].reverse()` so a nested `<w:ins>` inside another `<w:ins>` is processed *before* its parent — the inner mutation leaves the outer node's stored `parent` array intact for the outer pass. Forget this and a `<w:ins>` inside `<w:hyperlink>` inside `<w:ins>` blows up on the second pass.

## Note pairing post-pass

Footnotes/endnotes added or deleted under tracking have a body-side `<w:ins>` / `<w:del>` (in footnotes.xml / endnotes.xml) paired with a reference-side wrapper in document.xml. The reader hides the body-side from `tcN` ids so add/delete surfaces as ONE change. After applying targets, `applyNotePairing` consumes the `(kind, noteId)` pairs `collectAffectedNotes` snapshotted **before** mutation (post-mutation the `<w:del>` is gone and the link is lost):

- If no live `<w:footnoteReference>`/`<w:endnoteReference>` remains for that id → GC the entire `<w:footnote>` / `<w:endnote>` from the part. Word does the same.
- Otherwise → `normalizeNoteBody`: unwrap any remaining `<w:ins>` / `<w:del>` in the note's paragraphs (with `<w:delText>` → `<w:t>` rename on del), and strip any paragraph-mark ins/del markers.

`computeMaxRevisionId` scans all three parts (document + footnotes + endnotes) so allocated ids never collide with one already living in a note body.

## Table-grid resync

Cell removals (accept-cellDel / reject-cellIns) shrink rows without touching `<w:tblGrid>`. After applying targets, `resyncTableGrids` walks every `<w:tbl>` and reconciles the grid column count with the widest row, so the table model stays consistent. Standalone `tblGridChange` snapshots are handled by `restorePropertySnapshot` and don't need this pass.

## Tracked range edit / delete — Word-canonical shapes

`applyTrackedRangeReplace` (range edit) and `applyTrackedRangeDelete` (range delete) implement the exact XML shape Word produces for the equivalent operations. Both were empirically validated — see `scripts/word-redlines.sh` for the AppleScript oracle and `/tmp/range-probe/*.docx` for the probe outputs. Don't tweak the shape on a hunch; verify against a fresh probe.

The shorthand:
- **Range edit** (M old → N new paragraphs): old 1..M-1 get content `<w:del>` + paragraph-mark `<w:del>` (accept-all merges them forward); the transition (old M) has its content `<w:del>`'d, then the first new's content appended inline `<w:ins>`-wrapped, with paragraph-mark `<w:ins>` if N≥2; new 2..N each splice in as fresh `<w:p>` with content `<w:ins>` and paragraph-mark `<w:ins>` (except the trailing one, whose mark is bare).
- **Range delete** (M paragraphs): every paragraph in `[startIndex, endIndex]` gets content `<w:del>`; paragraphs `startIndex..endIndex-1` get paragraph-mark `<w:del>`. The last one's mark stays bare so accept-all leaves an empty residue paragraph (markdown render hides it).

**Re-editing an already-tracked paragraph — drop your own unaccepted insertions.** Every deletion-direction transform (`convertParagraphContentToDeleted` for the range paths, `applyDeletion` / `applyContentDeletion` for single delete) FIRST runs `dropOwnInsertions` (in [index.tsx](index.tsx)) over the paragraph's children: a top-level `<w:ins>` is content an earlier tracked edit in this same pass added but that hasn't been accepted, so re-deleting/replacing the paragraph must **un-insert** it (remove the wrapper + runs), not leave it visible. Without this, a paragraph edited twice under tracking (a single-block `--markdown` then a multi-block split — the contract-markup A-r1 shape — or an `edit` then a `delete`) kept the first edit's `<w:ins>` AND the second's, so Accept All showed the sentence twice and merged the following headings into a run-on. Dropping is correct on both sides of the review — accept and reject each make an unaccepted-then-replaced insertion vanish (reject restores the true baseline from the sibling `<w:del>`). This assumes the doc-level single-author ("Reviewer") model; a genuine cross-author prior insertion would need Word's nested del-of-ins shape, which we don't model. `wrapContiguousTrackable` still passes existing `<w:del>` / `<w:hyperlink>` / `<w:moveFrom>` / `<w:fldSimple>` / `<w:smartTag>` wrappers through untouched — only additive `<w:ins>` is dropped, and only in the deletion direction (never in `applyInsertion` / `wrapNewParagraphContentAsInserted`, which build or preserve NEW insertions, e.g. CriticMarkup from `edit --markdown`).

## Span replace inside another revision's insertion

The span engine behind `replace` and `edit --at pN:S-E` (`replaceSpanInParagraph` in [find/replace-span/](../find/replace-span/index.ts)) DOES model the cross-author shape, matching Word. It works in three steps, one file each, newspaper-ordered behind the folder's `index.ts`:

- **Where the span sits** ([target.ts](../find/replace-span/target.ts), `spanTarget`): every visible run with its offset and its FULL wrapper ancestry (`RunSlot.path`, outermost first — the XML twin of the reader's `within` chain), and the path all the runs the span touches share. An empty span (`pN:S-S`, a pure insertion) at a run boundary uses the path its two neighbors share, so an insertion at a link's or revision's EDGE lands beside it, and one strictly inside a link (between two of its runs) lands in the link with its style.
- **Which level to rebuild at** ([level.ts](../find/replace-span/level.ts), `chooseLevel`): the deepest shared container (a link the span sits in, so the replacement inherits the link) — except that under tracking, and for any empty span, the level RISES to the container above the OUTERMOST revision wrapper (`<w:ins>`/`<w:del>`/`<w:moveFrom>`/`<w:moveTo>`) on that path. So the replacement's own `<w:ins>` is never a descendant of another revision at any depth — inside one it was attributed to that author (issue #13), or, inside a deletion, vanished on accept. Every wrapper below the level splits; a link the span sat in below it is re-created around the replacement (`linkShell`) so it keeps the link.
- **The rebuild** ([rebuild.tsx](../find/replace-span/rebuild.tsx) → [split.tsx](../find/replace-span/split.tsx)): one walk over the level's children slices runs at the span's edges and splits every visible wrapper the span crosses, recursively (`splitAroundSpan` ↔ `partitionAroundSpan`) — moved whole, matched text inside a nested link/smart tag survived the replace with exit 0. The cut is settled by `settleCut`: dropped untracked; left as is when ANY enclosing wrapper is a deletion (`SpanContext.deleted`, propagated down the recursion — that text is already gone on accept, so it is never re-wrapped in our `<w:del>`); otherwise a `<w:del>` nested where the text was (inside another author's insertion — Word's shape). A split revision's tail gets a freshly allocated `w:id` under tracking (same author/date). Zero-width children (bookmarks, equations, a deletion hidden in the view) move whole to the side their offset falls on (inside the span → between the halves, so document order holds). A link the span STARTS inside keeps the replacement in its leading half.

**Same author merges.** When the tracked author equals the NEAREST enclosing revision's `w:author` on the shared path, that revision is an `<w:ins>`/`<w:moveTo>` (exact string equality), and no other revision sits between it and any run the span touches, editing it is not a new revision — Word just edits your own pending insertion. So the span is rebuilt UNTRACKED at the shared level inside it (`Level.ownInsertion`): the cut is removed outright and the replacement lands as plain runs in the existing insertion, keeping its `w:id`/date — no split, no nested `<w:del>`, no new `<w:ins>`, even when the span crosses a link nested in the insertion. An insertion point at the own insertion's edge extends it. An insertion the merge leaves holding only blank runs (you deleted all your own pending text) is removed (`dropIfBlank`) rather than lingering as an empty revision. A span that crosses out of the own insertion (or sits in a deletion) keeps the tracked shape above.

Known gaps: an untracked split of a revision wrapper (an untracked insertion point inside another author's insertion) keeps the same `w:id` on both halves (no allocator untracked). The whole-paragraph `edit --text` diff (`applyFormattingPreservingEdit`) still follows the single-author model above.

## Nested revisions — view visibility is the whole chain

A revision wrapper can sit inside another (an editor's `<w:del>` inside a reviewer's `<w:ins>`, as the span replace above emits; or an `<w:ins>` inside a `<w:del>`). The reader keeps the AST honest — a run's `trackedChange` is its INNERMOST wrapper (its own `tcN`/kind/author) — and records the enclosing revision on `TrackedChange.within`. Visibility uses the whole chain: content shows in a view only if EVERY enclosing revision does (accepted hides `del`/`moveFrom`, baseline hides `ins`/`moveTo`, current shows all), because rejecting an insertion removes everything inside it and accepting a deletion removes everything inside it. The per-kind rule is stated ONCE, `isRevisionKindVisible` in [parser/revision-view.ts](../parser/revision-view.ts); `isRevisionVisible` ([ast/revision-visibility.ts](../ast/revision-visibility.ts)) applies it along the chain for the AST — `paragraphTextAccepted`/`Baseline` (wc), find's `runViewText`, `iterateBlocks({ view })`, and the markdown renderer (every run kind: text, breaks, tabs, images) all use it — and `isWrapperVisibleInView` / `sumVisibleTextLength` apply it for the XML offset walkers (replace-span, replace-across, comments markers, set/clear formatting), which recurse only into visible wrappers and skip a wrapper by its VISIBLE length. So `read --baseline`/`--accepted` print exactly what `reject --all`/`accept --all` leave (asserted in `tests/cli/track-changes.test.ts`), and a span `find` reports lands on the same text in every walker.

## The shared range guard

`assertParagraphOnlyTrackedRange(rangeRef)` throws `TrackedRangeConflictError` if any block in `[startIndex, endIndex]` is non-`<w:p>` (most commonly a table). The tracked-range walkers inject `<w:pPr>` into every span block, which would corrupt `<w:tbl>`. Untracked range paths splice cleanly across any block tag, so the guard only fires when tracking is on. Used by `Edit.range` and `cli/delete/index.tsx::commitRangeDelete`; each catches the error and maps to its own error code.

## Word-level diff for `edit --text`

`applyFormattingPreservingEdit` is what `Edit.paragraph` falls through to when the spec is `--text` without override formatting flags and `--no-formatting` isn't set. It:

1. **Extracts old tokens** via `extractOldTokens` — recursively walks the paragraph descending through visible-in-accepted-view run-bearing wrappers (`<w:ins>`, `<w:moveTo>`, `<w:hyperlink>`, `<w:fldSimple>`, `<w:smartTag>`; skips `<w:del>`, `<w:moveFrom>`), concatenates every visible run's text once, builds a per-character rPr index, and tokenizes the full string. This handles source docs that split words across `<w:r>` boundaries (e.g. `"me" + "ssenger"` in adjacent italic runs) — flat per-run tokenization would shred those into spurious del+ins pairs.
2. **Tokenizes the new text** identically.
3. **LCS diff** aligns new tokens against old (`diffTokens`).
4. **Re-emits runs** grouped by rPr: untracked path emits only kept + inserted tokens (inherited rPr from nearest matched neighbor); tracked path emits kept tokens as plain runs, inserts in `<w:ins>`, deletes in `<w:del>`. Pure pPr properties (`--style`, `--alignment`) apply in place via `applyParagraphOptionsInPlace`.

rPr equality is structural (XML-string compare on cloned nodes). The codebase's emitter produces canonical output so this works in practice; non-canonical rPr from an external producer might leave equivalent rPr blocks unmerged (each token gets its own run). Acceptable.

## Adding a new tracked-change kind

See [cli/track-changes/CLAUDE.md "Adding a tracked-change kind"](../../cli/track-changes/CLAUDE.md). The summary: extend `TrackedChange["kind"]` in `core/ast/types.ts`, widen `cli/info/schema.ts`, register in `walkRunContainer` in `core/ast/read.ts` (the single source of `tcN` ids), and handle in `actionFor` / `applyAccept` / `applyReject` here. If the wrapper tag is ambiguous (table-structural, checkboxToggle), set `reference.kind` explicitly in the reader so `changeFoundFromReference` doesn't have to guess; otherwise `trackedChangeKindForTag` covers it. Surface in `paragraphTextAccepted` / `paragraphTextBaseline` ([core/ast/text.ts](../ast/text.ts)) and `cli/read/markdown.ts` if the kind affects visible text.
