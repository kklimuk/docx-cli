# src/core/markdown — `MarkdownImport` lens, walker, CriticMarkup

Five files behind the `@core/markdown` barrel ([index.ts](index.ts)):

- [import.tsx](import.tsx) — the `MarkdownImport` lens. Assembles the unified pipeline (`remark-parse` + `remark-gfm` + `remark-math` + our `remarkCriticMarkup`), pre-walks the mdast tree once for side effects (mint footnote ids, register bodies into `footnotes.xml`, fetch image bytes and mint rels), then hands off to the block walker.
- [walker.tsx](walker.tsx) — block dispatcher (`walkRoot`, `walkBlock`). One handler per mdast block type, each composing the existing emitters in `@core` (`Paragraph` / `HorizontalRule` / `Table` / `BlankTable` / `buildCodeBlockParagraphs` / `latexToOmml`).
- [inline.tsx](inline.tsx) — phrasing walker (`walkInline`). Returns `<w:p>`-child siblings: a mix of `<w:r>`, `<w:hyperlink>`, `<m:oMath>`, `<w:ins>`, `<w:del>`. Threads an immutable `InlineFormat` through nested strong / em / delete / inlineCode wrappers.
- [critic.ts](critic.ts) — `remarkCriticMarkup` plugin. Post-parse, visits every `text` node and splits it on `{++…++}` / `{--…--}` markers, producing `criticInsert` / `criticDelete` phrasing nodes the inline walker dispatches on.
- [errors.ts](errors.ts) — `MarkdownImportError` (code is a strict subset of `cli/respond.ts`'s `ErrorCode`).

## The lens does an async pre-walk, then a sync block walk

`MarkdownImport.blocks(source)` is the only public entry point. It runs in three phases:

1. **Parse** — `parseToMdast(source)`: build the unified processor, call `.parse()`, then `.runSync()` to apply the CriticMarkup transformer. Returns the mutated mdast `Root`.
2. **Pre-walk (async)** — `collectFootnoteDefinitions` then `registerFootnotes` (mint numeric ids, append `<w:footnote>` bodies to `footnotes.xml`, populate `ctx.mintedNoteIds`); `preloadImages` (for each `<image>` URL: if the URL is hash-shaped — `<sha256>.<ext>` — first try to reuse an existing image from the target document via `body.findImageByHash`; otherwise fetch via `loadImageSource`, mint rel + image part via `Images.add`, compute EMU extents, cache as `ResolvedImage`).
3. **Block walk (sync)** — `walkRoot(tree, ctx)` reads from `ctx.mintedNoteIds` / `ctx.imageCache` without itself touching async; produces the `<w:body>`-ready `XmlNode[]` the caller splices.

Phases 2 and 3 share one `WalkContext` (see [inline.tsx](inline.tsx)). The CLI verbs (`insert`, `create`, `edit`) all consume this same flow; `edit` only passes the resulting blocks through `Edit.paragraph` / `Edit.range` as a pre-built `markdown-blocks` spec so the `Edit` lens itself stays sync.

`blocks(source, options)` takes two routing knobs the note-body callers (`footnotes`/`endnotes add|edit --markdown`) MUST pass, because their result is spliced into `footnotes.xml`/`endnotes.xml`, NOT `document.xml`: `options.relationships` overrides where hyperlink rels are minted (pass `notesView.ensureRelationships()` so a `<w:hyperlink r:id>` resolves against `word/_rels/<part>.xml.rels` — otherwise the rId dangles and Word reports "unreadable content"), and `options.stripImages` drops every `image`/`imageReference` before the walk (note bodies are text + links; an image's media rel is minted by `Images.add` into the document rels regardless of `options.relationships`, so it would dangle in the note part — strip it instead). The default body verbs (`insert`/`create`/`edit`) splice into `document.xml` and pass neither.

## CriticMarkup tokenization happens at text-node level

`{++text++}` and `{--text--}` are tokenized only inside parsed `Text` nodes — markers that straddle markdown formatting (`{++**bold**++}`) are NOT recognized because the strong wrapper is already established at parse time. To track formatted content, place markers INSIDE the formatting (`**{++bold++}**`). Documented limitation; a follow-up could implement a micromark extension for proper inline-level tokenization.

When `<w:trackChanges/>` is on, `criticInsert` wraps its runs in `<w:ins>` (text stays `<w:t>`); `criticDelete` wraps in `<w:del>` with `<w:t>` → `<w:delText>` rename. When tracking is off, `criticInsert` flattens to plain text (CriticMarkup's "accepted" view) and `criticDelete` drops entirely. Both behaviors are determined once at lens construction via `document.isTrackChangesEnabled()` — there's no per-call override.

## Adding an mdast node type

Two cases, depending on whether it's block- or inline-level:

- **Block** (`type` in `RootContentMap`): add a `case` in `walkBlock` in [walker.tsx](walker.tsx). Compose an existing emitter where possible — heading reuses `<Paragraph>` with `pStyle`, blockquote reuses `paragraphBlock` with a Quote `styleOverride`, table goes through `<Table>` + `<TableCell>` from `@core/table`, etc. If you need to write to a separate part (numbering.xml, footnotes.xml), use the `document.ensureX()` accessor — never reach for the part directly.
- **Inline** (`type` in `PhrasingContentMap`): add a `case` in `walkPhrasing` in [inline.tsx](inline.tsx). Decide whether the new node carries `InlineFormat` through recursion (like `strong` / `emphasis`) or is a leaf that emits a single sibling (like `inlineMath` / `footnoteReference`). For new wrapper kinds (something that contains children), follow the `delete` pattern: `node.children.flatMap(child => walkPhrasing(child, ctx, { ...format, newFlag: true }))`.

If the new construct is async-loadable (e.g. another media kind), add a corresponding `preload*` pass in [import.tsx](import.tsx) and a cache on `WalkContext` — the sync walker reads from the cache.

## Blockquote depth + the deliberate escape

OOXML has no `<w:blockquote>` container — quote treatment is per-paragraph. The walker encodes a markdown blockquote as:

- **Quoted paragraph** → `pStyle="Quote"` + `<w:ind w:left={720 * depth}>`
- **List item inside a quote** → `pStyle="QuoteListParagraph"` (a baseline style that extends `ListParagraph` with italic; see [styles.tsx](../ast/document/styles.tsx)) + `<w:ind w:left={720 * depth}>` on the paragraph
- **Nested blockquote** → same as quoted paragraph but at deeper `depth`

The AST reader recovers `paragraph.quoteDepth` by reading the pStyle prefix (`Quote*`) plus the paragraph's `<w:ind w:left>` value (`leftTwips / 720`, floor to 1). The markdown renderer prepends one `> ` per depth. Round-trip is lossless for paragraphs, lists, and nested quotes.

**The escape:** code blocks, tables, math (`$$..$$`), headings, thematic breaks, and any other non-paragraph / non-list / non-blockquote child inside a `> ` block are **emitted at top level** by `blockquoteBlocks` (it calls `walkBlock(child, ctx)` without quote framing). They break the quote at that point — adjacent quoted content before and after surfaces as separate blockquotes on round-trip.

This is a deliberate v0.12 design choice. The cost of full round-trip would be combinatorial baseline styles (`QuoteCodeBlock-LANG` × 37 languages, `QuoteTable`, `QuoteMath`, …); the realistic frequency of these nested constructs in actual blockquotes is low enough that escape-and-document is the better trade. Section 7.6 of [tests/fixtures/setup/markdown-import.source.md](../../../tests/fixtures/setup/markdown-import.source.md) demonstrates the behavior.

## Image round-trip via content hash

`docx read --markdown` emits every embedded image as `![alt](<sha256>.<ext>)` — the same content-addressed naming `docx images extract` writes to disk. On the import side, `tryReuseImageByHash` in [import.tsx](import.tsx) sniffs that URL shape, looks the hash up in the target document's `body` via `findImageByHash`, and reuses the existing relationship id when it matches. The result: a same-doc round-trip (`read --markdown → edit --markdown-file → write`) doesn't duplicate any `word/media/imageN` parts, and cross-doc round-trips work whenever the target happens to carry an identical image.

If the URL is hash-shaped but the doc carries no matching image, we surface a clear `IMAGE_SOURCE` error pointing at the hash. The walker doesn't silently fall through to `loadImageSource("abc123…ef0.png")` because that would just produce an ENOENT with no hint about what went wrong.

## What we don't support yet

- **Footnote/endnote body fidelity** — `registerFootnotes` in [import.tsx](import.tsx) walks each definition's inline content through `walkInline` (untracked), so **bold/italic and hyperlinks inside note bodies are preserved**: note-body link rels are minted into the part's OWN rels (`word/_rels/footnotes.xml.rels`) via `NotesView.ensureRelationships()` + a `WalkContext.relationships` swap, since a `<w:hyperlink r:id>` inside `footnotes.xml` resolves against the footnotes rels, not `document.xml.rels`. Two caveats remain: (a) **under track-changes**, note bodies still flatten to plain text (the verified `TrackedNoteBody` shape is single-run; rich tracked bodies aren't Word-verified) — links/formatting are lost only in that case; (b) **images** inside note bodies are dropped (`stripImagePhrasing`). Also: footnote reference **labels** are renumbered to positional `[^fnN]`/`[^enN]` on import (a named `[^pt]` becomes `[^fn2]`); the markdown id is not preserved. NOTE: the *reader* (`NotesView.toNotes`) still surfaces note bodies as plain `text`, so `read --ast` / `read --markdown` don't yet round-trip note-body links back out — a read-side follow-up.
- **Multiple references to one footnote** — markdown lets `[^x]` be cited repeatedly; OOXML/Word require a *distinct* footnote definition per reference (Word "repairs" N:1 linkage by cloning). `registerFootnotes` mints one definition clone per reference (`countFootnoteReferences`) and the walker consumes them in document order (`footnoteRefCursor`).
- **Frontmatter into core properties** — YAML frontmatter is dropped. A follow-up could surface `title:` / `author:` / `date:` into `docProps/core.xml`.
- **Comment references** — `[^c1]` (the dialect `read --markdown --comments` emits) is treated identically to a footnote reference. Round-tripping `read --comments` import-bound markdown imports the comment body as a footnote, not as a real `<w:comment>` — adding span-anchored comment imports needs the post-splice comment-anchor pass to live on the lens.
- **Locator HTML comments** — `<!-- p3 -->` markers are dropped by the inline / block walkers. The locator metadata is already implicit in block order, so round-trips reconstruct the locators on the read side.
- **Image dimensions overrides** — there's no `--image-width` style flag on `--markdown`; we use the source's intrinsic pixel dimensions. Pass `--width` via standalone `insert --image` for explicit sizing.
- **Strict CriticMarkup substitution** (`{~~old~>new~~}`) — only `{++…++}` and `{--…--}` are tokenized; substitution would need a third mdast node kind plus a paired `<w:del>` / `<w:ins>` emit.
