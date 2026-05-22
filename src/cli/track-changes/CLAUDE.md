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

Still out of scope: `<w:rPrChange>`, `<w:pPrChange>`. Author/date resolve via `--author NAME` → `DOCX_AUTHOR` env → `"docx-cli"` default (`core/track-changes/index.ts → resolveAuthor`); `DOCX_CLI_NOW` injects a fixed date for tests. `delete tN` rejects with `TRACKED_CHANGE_CONFLICT` when tracking is on (tracked table-row deletion isn't supported).

## tcN walk order — must mirror the AST reader

The `apply.ts` walk and `walkRunContainer` in `core/ast/read.ts` must visit revisions in the same order so `tcN` ids agree between `list` and `accept --at tcN`. Per paragraph: run-level wrappers first, then inline `<w:sectPrChange>`, then paragraph-mark `<w:ins>`/`<w:del>`.

## Adding a tracked-change kind

Extend `TrackedChange["kind"]` (and `TrackedChangeKind`) in `core/ast/types.ts`, widen the JSON schema enum in `cli/info/schema.ts`, update `walkRunContainer` in `core/ast/read.ts`, update `paragraphTextAccepted`/`paragraphTextBaseline` in `core/ast/text.ts` and `isRunVisible`/`criticMarkerFor`/`trackedChangeLabelFor` in `cli/read/markdown.ts`. If the kind takes part in accept/reject, update `actionFor` and `applyAccept`/`applyReject` in [apply.ts](apply.ts) and `trackedChangeKindForTag` in [list.ts](list.ts). Slot it into the same walk order in both the reader and apply.ts.
