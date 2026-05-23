# src/core/notes — footnote + endnote emit, lookup, tree ops

One file behind the `@core/notes` barrel ([index.ts](index.ts)): [helpers.tsx](helpers.tsx). Footnotes and endnotes share their entire mechanics — separate parts (`word/footnotes.xml` / `word/endnotes.xml`), same XML shape, same id-allocation rules, same reference-run pattern. The package is parameterized on `NoteKind = "footnote" | "endnote"` via the `noteConfig(kind)` lookup; every public function takes a `NoteKind` and pulls tag names / style ids / part names from the config.

## Lifecycle, lookup, emit, tree ops

The file follows the newspaper convention: `ensureNotesPart` (the entry that everything else builds on) is at the top, its internal helpers follow, then `ensureNoteStyles`, then lookup ops (`nextNoteId`, `findNoteByNumericId`), then emitters (`NoteReferenceRun`, `NoteBody`, `TrackedNoteBody`), then tree mutators (`insertNoteReferenceAtOffset`, `removeNoteReferences`) with their walker helpers, then tracked-body mutators (`wrapNoteBodyAsDeleted`, `wrapNoteBodyAsEdited`), and finally `noteConfig` + types + the `FOOTNOTE_CONFIG` / `ENDNOTE_CONFIG` catalog at the bottom — leaf data, same shape as the `BASELINE` catalog in `core/styles.tsx`.

`ensureNotesPart` lazy-provisions the part the first time an agent runs `footnotes add` against a doc that had none. Word writes two reserved entries before any user notes — `id=-1 separator` (the rule above the note area) and `id=0 continuationSeparator` (used when notes wrap across pages); both have `w:type` set and are filtered out by the AST reader. We seed both so LibreOffice and Word render the note area correctly from the start.

## Tracked footnotes

`TrackedNoteBody` (add side) and `wrapNoteBodyAsDeleted` / `wrapNoteBodyAsEdited` (delete / edit sides) mirror Microsoft Word's empirical XML shape — see `scripts/word-redlines.sh` for the AppleScript oracle. Each operation gets distinct revision ids on each side (`computeMaxRevisionId` walks document/footnotes/endnotes so allocation doesn't collide). The reference-side and body-side revisions are **paired structurally by the footnote id** (the `w:id` on `<w:footnoteReference>` and `<w:footnote>`), not by revision id; the accept/reject coordinator (`cli/track-changes/apply.ts::applyNotePairing`) consumes `noteConfig` from here to walk the pairing and GC orphan bodies.

## Adding a property to a NoteConfig

Extend the `NoteConfig` shape at the bottom and add the new field to both `FOOTNOTE_CONFIG` and `ENDNOTE_CONFIG`. Public functions pick it up automatically via `noteConfig(kind)` — no other edits needed.
