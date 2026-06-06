# src/core/notes — footnote + endnote emit, lookup, tree ops

Four files behind the `@core/notes` barrel ([index.ts](index.ts)):

- [config.ts](config.ts) — `NoteKind = "footnote" | "endnote"`, the `NoteConfig` shape (tag names, style ids, part names, relationship type), and the `FOOTNOTE_CONFIG` / `ENDNOTE_CONFIG` lookup via `noteConfig(kind)`. Leaf data, same shape as the `BASELINE` catalog in `core/styles.tsx`.
- [empty.tsx](empty.tsx) — `ensureNotesPart` (lazy-provisions `word/footnotes.xml` / `word/endnotes.xml` the first time an agent runs `footnotes add` against a doc that had none — seeds Word's reserved `id=-1 separator` + `id=0 continuationSeparator` entries so LibreOffice and Word render the note area correctly), `ensureNoteStyles`, lookup ops (`nextNoteId`, `findNoteByNumericId`).
- [emit.tsx](emit.tsx) — the emitter components: `NoteReferenceRun`, `NoteBody`, `TrackedNoteBody`, plus tracked-body mutators `wrapNoteBodyAsDeleted` / `wrapNoteBodyAsEdited`.
- [splice.tsx](splice.tsx) — tree mutators: `insertNoteReferenceAtOffset`, `removeNoteReferences` and their walker helpers.

Footnotes and endnotes share their entire mechanics — separate parts (`word/footnotes.xml` / `word/endnotes.xml`), same XML shape, same id-allocation rules, same reference-run pattern. The package is parameterized on `NoteKind` everywhere; every public function takes a `NoteKind` and pulls tag names / style ids / part names from `noteConfig`.

`ensureNotesPart` lazy-provisions the part the first time an agent runs `footnotes add` against a doc that had none. Word writes two reserved entries before any user notes — `id=-1 separator` (the rule above the note area) and `id=0 continuationSeparator` (used when notes wrap across pages); both have `w:type` set and are filtered out by the AST reader. We seed both so LibreOffice and Word render the note area correctly from the start.

## Word validity: three things a notes part needs, or Word rejects the file

Word reports "unreadable content" (and silently "repairs") a document whose notes
violate any of these — and LibreOffice is permissive enough that render/roundtrip
tests miss all three. `tests/cli/docx-validity.test.ts` is the guard (namespace +
rId-resolution checks, plus `xmllint` when present). The rules:

1. **One footnote definition PER reference.** Markdown lets `[^x]` be cited many
   times; OOXML/Word require a distinct `<w:footnote>` per `<w:footnoteReference>`
   (Word "repairs" N:1 by cloning). The markdown importer mints one clone per
   reference — see `core/markdown/import.tsx` (`countFootnoteReferences` +
   `footnoteRefCursor`).
2. **`<w:footnotePr>` / `<w:endnotePr>` in `settings.xml`** declaring the reserved
   separator (`id=-1`) + continuationSeparator (`id=0`). Without the settings-level
   pointer Word can't bind the separators and flags the file. `Document.ensureNoteInfrastructure()`
   adds both (via `SettingsView.ensureNotePr`) and provisions BOTH parts (Word always
   pairs footnotes+endnotes) whenever either is needed.
3. **Every namespace prefix used in the part must be declared on its root.** A
   note-body `<w:hyperlink r:id="…">` uses the `r:` prefix, so `buildEmptyNotesRoot`
   declares `xmlns:r` on `<w:footnotes>`/`<w:endnotes>` (like `document.xml` does) —
   an undeclared prefix is malformed XML. The hyperlink's relationship lives in the
   part's OWN rels (`word/_rels/footnotes.xml.rels`), owned by `NotesView`
   (`ensureRelationships` / `writeTo`), NOT `document.xml.rels`.

## Tracked footnotes

`TrackedNoteBody` (add side) and `wrapNoteBodyAsDeleted` / `wrapNoteBodyAsEdited` (delete / edit sides) mirror Microsoft Word's empirical XML shape — see `scripts/word-redlines.sh` for the AppleScript oracle. Each operation gets distinct revision ids on each side (`computeMaxRevisionId` walks document/footnotes/endnotes so allocation doesn't collide). The reference-side and body-side revisions are **paired structurally by the footnote id** (the `w:id` on `<w:footnoteReference>` and `<w:footnote>`), not by revision id; the accept/reject coordinator (`core/track-changes/apply.ts::applyNotePairing`) consumes `noteConfig` from here to walk the pairing and GC orphan bodies.

## Adding a property to a NoteConfig

Extend the `NoteConfig` shape in [config.ts](config.ts) and add the new field to both `FOOTNOTE_CONFIG` and `ENDNOTE_CONFIG`. Public functions pick it up automatically via `noteConfig(kind)` — no other edits needed.
