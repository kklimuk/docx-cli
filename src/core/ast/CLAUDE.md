# src/core/ast — `Document`, the embedded views, and the reader

`Document` is the composition root: one OPC package plus one **tree-owning view per part**. The reader (`read.ts`) is Document's construction pass.

## Layout

- [document/](document/) — `Document` ([document/index.ts](document/index.ts)) + the embedded views. Each view owns one OPC part's `XmlNode` tree plus any maps keyed to it:
  - `Body` ([document/body.ts](document/body.ts)) — `word/document.xml` body: `blocks`, `blockReferences` (pN/tN/sN), `imageById`, `hyperlinkById`, `equationReferences`. The `*iterateBlocks` walker is canonical (descends into table cells); use it over hand-rolling.
  - `RelationshipsView`, `ContentTypesView`, `StylesView`, `NumberingView`, `CommentsView`, `NotesView` (×2 — footnotes/endnotes), `SettingsView`, `CorePropertiesView`. Same lifecycle on each: `static fromPackage(pkg)`, `static fromXml(xml)`, instance `writeTo(pkg)`, and (for the optional ones) `static register(deps)` that mints the relationship + content-type and returns a fresh view.
  - `MarginalsView` ([document/marginals.ts](document/marginals.ts)) — the multi-part exception: ONE view owns every `word/header{N}.xml` / `word/footer{N}.xml` tree, keyed by part name. `fromPackage` scans `pkg.listParts()`; `writeTo` serializes all. No `register` — the [`Marginals`](../../marginals/CLAUDE.md) lens mints each part's rel + content-type as it allocates the part.
  - `Pkg` ([document/package.ts](document/package.ts)) — the generic OPC zip transport. Imports only `XmlNode` + JSZip. `Pkg.open` / `Pkg.empty` / `readPart` / `writeText` / `writeBytes` / `save`. Stays domain-agnostic.
- [read.ts](read.ts) — the single walk over `<w:body>` and the note parts; populates Body's maps, registers every tracked change into `Document.trackedChangeReferences`, and asks each part-view for its AST slice (`CommentsView.toComments`, `NotesView.toNotes`; `readMarginals` resolves each section's header/footer references → `MarginalsView` part trees → `Body.headers`/`footers`).
- [types.ts](types.ts) — the AST types (`Block`, `Paragraph`, `Run`, `TextRun`, `Table`, `Comment`, `Note`, `TrackedChange`, …). What `read --ast` and `info schema` expose.
- [text.ts](text.ts) — `paragraphText` / `paragraphTextAccepted` / `paragraphTextBaseline` over a single `Paragraph`. Free functions because they need no `Document`.
- [sym.ts](sym.ts) — Wingdings/Symbol-font `<w:sym>` decoding.

## The reader IS Document's construction pass

`Document.open(path)` parses each part, constructs the embedded views, then calls `buildBody(document, path)`. The reader does three coupled things in one walk:

1. **Bind cross-part lookups it needs first** — `document.relationships.index(document.contentTypes)` populates the rId → media/hyperlink maps the body walk dereferences.
2. **Build the AST and Body's maps** — `walkRunContainer` descends `<w:body>`, building `Paragraph.runs` (TextRun / ImageRun / EquationRun / …) and writing `blockReferences` / `imageById` / `hyperlinkById` / `equationReferences` as it goes.
3. **Register every tracked change into `Document.trackedChangeReferences`** — run-level wrappers (`<w:ins>`/`<w:del>`/`<w:moveFrom>`/`<w:moveTo>`), paragraph-mark wrappers in `<w:pPr><w:rPr>` (both top-level AND cell paragraphs), `<w:sectPrChange>`, table-structural revisions (`rowIns`/`rowDel`/`cellIns`/`cellDel`/`tblGridChange`/`tblPrChange`/`tcPrChange`), and `checkboxToggle` SDTs. After the body walk, `readNoteTrackedChanges` appends standalone footnote/endnote edits (paired notes — those with a reference-side wrapper in the body — are hidden so the change surfaces as one tcN).

This is the **only walk that assigns `tcN` ids**. `cli/track-changes/list.ts` and `core/track-changes/apply.ts` both read `Document.trackedChangeReferences` (or `new TrackChanges(document).list()`) and never re-walk — so they can't drift from the AST.

## Why the reader is free functions and not a `Body` method

The walk reads `document.relationships` + `document.numbering` (image/list resolution) and writes `document.trackedChangeReferences` — a Document-level map spanning body + notes. `readNoteTrackedChanges` operates on the footnote/endnote views directly. So the reader is genuinely cross-part orchestration; making it a `Body` method would force `Body` to reach sideways into its siblings, which is exactly the coupling the embedded-views split removed everywhere else. Free functions stay.

## Adding an AST field

Add to [types.ts](types.ts), populate in [read.ts](read.ts), then widen `cli/info/schema.ts`. The `info schema --ts` output reads `types.ts` live via Bun's text import, so it stays in sync.

## Adding a tracked-change kind

See [src/cli/track-changes/CLAUDE.md](../../cli/track-changes/CLAUDE.md). The summary: register it in the right place in the reader's body walk (the single source of `tcN` ids), handle it in `actionFor` / `applyAccept` / `applyReject` (`core/track-changes/apply.ts`), and surface it in `paragraphTextAccepted` / `paragraphTextBaseline` ([text.ts](text.ts)) + `cli/read/markdown.ts`. One place to register — no second walker to mirror.
