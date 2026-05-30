# src/core/notes — footnote + endnote emit, lookup, tree ops

Four files behind the `@core/notes` barrel ([index.ts](index.ts)):

- [config.ts](config.ts) — `NoteKind = "footnote" | "endnote"`, the `NoteConfig` shape (tag names, style ids, part names, relationship type), and the `FOOTNOTE_CONFIG` / `ENDNOTE_CONFIG` lookup via `noteConfig(kind)`. Leaf data, same shape as the `BASELINE` catalog in `core/styles.tsx`.
- [empty.tsx](empty.tsx) — `ensureNotesPart` (lazy-provisions `word/footnotes.xml` / `word/endnotes.xml` the first time an agent runs `footnotes add` against a doc that had none — seeds Word's reserved `id=-1 separator` + `id=0 continuationSeparator` entries so LibreOffice and Word render the note area correctly), `ensureNoteStyles`, lookup ops (`nextNoteId`, `findNoteByNumericId`).
- [emit.tsx](emit.tsx) — the emitter components: `NoteReferenceRun`, `NoteBody`, `TrackedNoteBody`, plus tracked-body mutators `wrapNoteBodyAsDeleted` / `wrapNoteBodyAsEdited`.
- [splice.tsx](splice.tsx) — tree mutators: `insertNoteReferenceAtOffset`, `removeNoteReferences` and their walker helpers.

Footnotes and endnotes share their entire mechanics — separate parts (`word/footnotes.xml` / `word/endnotes.xml`), same XML shape, same id-allocation rules, same reference-run pattern. The package is parameterized on `NoteKind` everywhere; every public function takes a `NoteKind` and pulls tag names / style ids / part names from `noteConfig`.

`ensureNotesPart` lazy-provisions the part the first time an agent runs `footnotes add` against a doc that had none. Word writes two reserved entries before any user notes — `id=-1 separator` (the rule above the note area) and `id=0 continuationSeparator` (used when notes wrap across pages); both have `w:type` set and are filtered out by the AST reader. We seed both so LibreOffice and Word render the note area correctly from the start.

## Tracked footnotes

`TrackedNoteBody` (add side) and `wrapNoteBodyAsDeleted` / `wrapNoteBodyAsEdited` (delete / edit sides) mirror Microsoft Word's empirical XML shape — see `scripts/word-redlines.sh` for the AppleScript oracle. Each operation gets distinct revision ids on each side (`computeMaxRevisionId` walks document/footnotes/endnotes so allocation doesn't collide). The reference-side and body-side revisions are **paired structurally by the footnote id** (the `w:id` on `<w:footnoteReference>` and `<w:footnote>`), not by revision id; the accept/reject coordinator (`core/track-changes/apply.ts::applyNotePairing`) consumes `noteConfig` from here to walk the pairing and GC orphan bodies.

## Adding a property to a NoteConfig

Extend the `NoteConfig` shape in [config.ts](config.ts) and add the new field to both `FOOTNOTE_CONFIG` and `ENDNOTE_CONFIG`. Public functions pick it up automatically via `noteConfig(kind)` — no other edits needed.
