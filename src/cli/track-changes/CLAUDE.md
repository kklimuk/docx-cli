# src/cli/track-changes

`docx track-changes FILE on|off` toggles `<w:trackChanges/>` in settings.xml. `list` enumerates revision wrappers; `accept`/`reject FILE (--at tcN | --all)` apply them. The unwrap/delete logic lives in [apply.ts](apply.ts).

## Track-changes is doc-level, not per-command

When `<w:trackChanges/>` is set, every mutating command (`insert`/`edit`/`delete`/`replace`) automatically emits `<w:ins>`/`<w:del>` — there is no per-command override flag. For a one-off untracked edit: `track-changes off`, edit, `track-changes on`.

## Accept/reject mechanics

Accept/reject **bypass tracking** — they mutate the XML tree directly without wrapping the change in `<w:ins>`/`<w:del>`.

- Additive wrappers (`<w:ins>`, `<w:moveTo>`): accept-unwrap, reject-delete.
- Subtractive wrappers (`<w:del>`, `<w:moveFrom>`): accept-delete, reject-unwrap (with `<w:delText>` → `<w:t>` rename so restored text is plain).
- moveFrom/moveTo halves process independently — `--all` handles a complete move; targeting one half by `tcN` leaves the other in place.
- Targets walk fresh each invocation (stored `view.trackedChangeReferences` goes stale across mutations) and process in reverse pre-order so nested changes apply before parents.

Paragraph-mark trackings (a self-closing `<w:ins>`/`<w:del>` inside `<w:pPr><w:rPr>`) are full citizens: accept-ins / reject-del removes the marker (paragraph stays); reject-ins removes the entire owning paragraph (the inserted break disappears — for `insert --section` sentinels this collapses the section break too); accept-del merges the owning paragraph with the next (next paragraph's runs append here, next paragraph removed) per ECMA-376 §17.13.5.4.

Table-structural trackings are full citizens too. **Which constructs we emit was settled empirically against Microsoft Word (AppleScript accept/reject), not from the schema** — see [src/cli/tables](../tables/CLAUDE.md): `rowIns`/`rowDel` (`<w:trPr><w:ins>`/`<w:del>` — accept/reject acts on the whole `<w:tr>`), `cellIns`/`cellDel` (`<w:tcPr><w:cellIns>`/`<w:cellDel>` — per-cell record of a tracked column insert/delete), `tblGridChange` (`<w:tblGrid><w:tblGridChange>` — grid-width snapshot, restored on reject like `sectPrChange`), and `tcPrChange` (`<w:tcPr><w:tcPrChange>` — prior-`tcPr` snapshot; `set-widths` pairs one per cell with the `tblGridChange`, since Word's reject of a width change is driven by the per-cell `tcPrChange`). `tblPrChange` is read/accept/rejected (Word authors it) but we don't emit it — Word doesn't revert a hand-authored one. Cell merges and border changes are NOT tracked (Word applies them immediately and warns "won't be marked as a change"), so `merge`/`unmerge`/`borders` apply in place + drop a `[docx-cli]` audit comment. Cell removals (accept-cellDel / reject-cellIns) shrink rows without touching `<w:tblGrid>`, so `runApply` runs `resyncTableGrids` afterward to trim the grid to the widest row. Row markers are ambiguous by tag (`<w:ins>` is also run-level), so the `TrackedChangeReference` carries an explicit `kind` that `list.ts` prefers; `apply.ts` disambiguates by walk context.

Footnote/endnote trackings span two parts. **Empirically validated against Microsoft Word** (see [scripts/word-redlines.sh](../../../scripts/word-redlines.sh) and the AppleScript probes in `/tmp/fn-probe/{add,delete,edit}.docx`): adding a footnote under tracking wraps the reference run in `document.xml` AND the body's content (`<w:footnoteRef/>` + text runs) in `footnotes.xml` in independent `<w:ins>` wrappers; deleting wraps both sides in `<w:del>` (with `<w:t>`→`<w:delText>` rename on the body) and marks the paragraph-mark of the body's `<w:p>` with a third `<w:del>` inside `<w:pPr><w:rPr>`; editing the body only touches `footnotes.xml`, emitting `<w:ins>NEW</w:ins><w:del>OLD</w:del>` with ins preceding del. The two sides are **paired structurally by the footnote id** (the `w:id` on `<w:footnoteReference>` / `<w:footnote>`), not by revision id — each side allocates its own revision id from the same monotonic counter (`computeMaxRevisionId` walks document/footnotes/endnotes so ids never collide).
- `collectTrackedChanges` walks document.xml, then footnotes/endnotes; the body-side `<w:ins>`/`<w:del>` of a paired note is hidden from `list`/`apply` so add/delete surface as ONE tcN (the reference-side). Body-only revisions (footnote edits) stay visible — they're standalone.
- `applyNotePairing` (post-pass after all targets are applied) handles the paired side via the footnote-id linkage: if no live `<w:footnoteReference>` remains, GC the entire `<w:footnote>` from footnotes.xml (Word does the same — orphan bodies disappear on accept-del / reject-ins); otherwise unwrap any remaining body-side `<w:ins>`/`<w:del>` and drop the paragraph-mark del marker. The emitters live in [src/core/notes/helpers.tsx](../../core/notes/helpers.tsx).

Still out of scope: `<w:rPrChange>`, `<w:pPrChange>`. Author/date resolve via `--author NAME` → `DOCX_AUTHOR` env → `"docx-cli"` default (`core/track-changes/index.ts → resolveAuthor`); `DOCX_CLI_NOW` injects a fixed date for tests. `delete tN` (the top-level block delete) still rejects with `TRACKED_CHANGE_CONFLICT` when tracking is on — tracked deletion of a whole table goes through `tables delete-row`/`delete-column` instead.

## tcN walk order — must mirror the AST reader

The `apply.ts` walk and `walkRunContainer` in `core/ast/read.ts` must visit revisions in the same order so `tcN` ids agree between `list` and `accept --at tcN`. Per paragraph: run-level wrappers first, then inline `<w:sectPrChange>`, then paragraph-mark `<w:ins>`/`<w:del>`.

## Adding a tracked-change kind

Extend `TrackedChange["kind"]` (and `TrackedChangeKind`) in `core/ast/types.ts`, widen the JSON schema enum in `cli/info/schema.ts`, update `walkRunContainer` in `core/ast/read.ts`, update `paragraphTextAccepted`/`paragraphTextBaseline` in `core/ast/text.ts` and `isRunVisible`/`criticMarkerFor`/`trackedChangeLabelFor` in `cli/read/markdown.ts`. If the kind takes part in accept/reject, update `actionFor` and `applyAccept`/`applyReject` in [apply.ts](apply.ts) and `trackedChangeKindForTag` in [list.ts](list.ts). Slot it into the same walk order in both the reader and apply.ts.

Table-structural kinds (rowIns/rowDel/cellIns/cellDel/tblGridChange) live outside the run/paragraph machinery: register them in `readTable` (`core/ast/read.ts`) and `visitTable` (apply.ts) in the same order, give the `TrackedChangeReference` an explicit `kind` (their tags are ambiguous), handle them in `applyTableChange` (apply.ts), and filter whole tracked rows by view in `renderTable` (`cli/read/markdown.ts`) — they don't touch `text.ts` (no run text). GFM tables can't express cell-level changes, so only whole-row insert/delete are view-filtered.
