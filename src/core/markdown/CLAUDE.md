# src/core/markdown — `MarkdownImport` lens, walker, inline-surgery (CriticMarkup + spans)

Five files behind the `@core/markdown` barrel ([index.ts](index.ts)):

- [import.tsx](import.tsx) — the `MarkdownImport` lens. Assembles the unified pipeline (`remark-parse` + `remark-gfm` + `remark-math` + our `remarkInlineSurgery`), pre-walks the mdast tree once for side effects (mint footnote ids, register bodies into `footnotes.xml`, fetch image bytes and mint rels), then hands off to the block walker.
- [walker.tsx](walker.tsx) — block dispatcher (`walkRoot`, `walkBlock`). One handler per mdast block type, each composing the existing emitters in `@core` (`Paragraph` / `HorizontalRule` / `Table` / `BlankTable` / `buildCodeBlockParagraphs` / `latexToOmml`).
- [inline.tsx](inline.tsx) — phrasing walker (`walkInline`). Returns `<w:p>`-child siblings: a mix of `<w:r>`, `<w:hyperlink>`, `<m:oMath>`, `<w:ins>`, `<w:del>`. Threads an immutable `InlineFormat` through nested strong / em / delete / inlineCode / bracketed-span wrappers.
- [inline-surgery.ts](inline-surgery.ts) — `remarkInlineSurgery` plugin. Post-parse, rewrites each phrasing-children array with two reducers: text-marker surgery gathers CriticMarkup (`{++…++}` / `{--…--}`) and legacy Pandoc spans (`[text]{.class key="value"}`); **HTML-element surgery (`gatherHtmlSpans`)** re-pairs the `<span>` / `<mark>` / `<sup>` / `<sub>` / `<u>` tags `read` now emits (remark leaves inline HTML as FLAT, unpaired `html` tokens — open tag, mdast content, close tag as separate siblings — so we re-pair them, the same stack-reduce as the text markers). Both feed `bracketedSpan` (etc.) parent nodes the inline walker dispatches on. Each tag's attributes are parsed with the project's `XmlNode.parse` (fast-xml-parser), not a bespoke regex. Scanning whole sibling arrays (not single `text` nodes) is what lets a marker straddle inline formatting.
- [errors.ts](errors.ts) — `MarkdownImportError` (code is a strict subset of `cli/respond.ts`'s `ErrorCode`).

## The lens does an async pre-walk, then a sync block walk

`MarkdownImport.blocks(source)` is the only public entry point. It runs in three phases:

1. **Parse** — `parseToMdast(source)`: build the unified processor, call `.parse()`, then `.runSync()` to apply the inline-surgery transformer. Returns the mutated mdast `Root`.
2. **Pre-walk (async)** — `collectFootnoteDefinitions` then `registerFootnotes` (mint numeric ids, append `<w:footnote>` bodies to `footnotes.xml`, populate `ctx.mintedNoteIds`); `preloadImages` (for each `<image>` URL: if the URL is hash-shaped — `<sha256>.<ext>` — first try to reuse an existing image from the target document via `body.findImageByHash`; otherwise fetch via `loadImageSource`, mint rel + image part via `Images.add`, compute EMU extents, cache as `ResolvedImage`).
3. **Block walk (sync)** — `walkRoot(tree, ctx)` reads from `ctx.mintedNoteIds` / `ctx.imageCache` without itself touching async; produces the `<w:body>`-ready `XmlNode[]` the caller splices.

Phases 2 and 3 share one `WalkContext` (see [inline.tsx](inline.tsx)). The CLI verbs (`insert`, `create`, `edit`) all consume this same flow; `edit` only passes the resulting blocks through `Edit.paragraph` / `Edit.range` as a pre-built `markdown-blocks` spec so the `Edit` lens itself stays sync.

`blocks(source, options)` takes two routing knobs the note-body callers (`footnotes`/`endnotes add|edit --markdown`) MUST pass, because their result is spliced into `footnotes.xml`/`endnotes.xml`, NOT `document.xml`: `options.relationships` overrides where hyperlink rels are minted (pass `notesView.ensureRelationships()` so a `<w:hyperlink r:id>` resolves against `word/_rels/<part>.xml.rels` — otherwise the rId dangles and Word reports "unreadable content"), and `options.stripImages` drops every `image`/`imageReference` before the walk (note bodies are text + links; an image's media rel is minted by `Images.add` into the document rels regardless of `options.relationships`, so it would dangle in the note part — strip it instead). The default body verbs (`insert`/`create`/`edit`) splice into `document.xml` and pass neither.

## inline-surgery: CriticMarkup + bracketed spans, gathered across phrasing siblings

`remarkInlineSurgery` runs **after** remark's own parse and rewrites whole phrasing-children arrays. Because it tokenizes a sibling sequence — not one `text` node at a time — a marker may **straddle** inline formatting: `{++**bold**++}` parses to `[text("{++"), strong, text("++}")]`, and the matched markers gather the `strong` between them (the old text-split plugin couldn't, and leaked the markers as literal text). The same machinery parses Pandoc spans `[text]{attrs}`; the walker overlays a span's parsed attributes onto the inherited `InlineFormat`.

Robustness is the contract: markers that don't balance, and spans whose attributes parse to nothing, **degrade to their literal text** — never a throw, never lost content. `inlineCode` (and every non-text node) is an opaque atom, so code spans (`` `{++x++}` ``) are excluded for free. Any unexpected failure restores the original children array.

When `<w:trackChanges/>` is on, `criticInsert` wraps its (now possibly multi-run, formatted) content in `<w:ins>`; `criticDelete` wraps in `<w:del>` with **every descendant** `<w:t>` → `<w:delText>` (recursive rename — deleted content can nest hyperlinks/formatted runs, and a bare `<w:t>` inside `<w:del>` is invalid OOXML). When tracking is off, `criticInsert` flattens to plain runs (CriticMarkup's "accepted" view, formatting preserved) and `criticDelete` drops entirely. Decided once at lens construction via `document.isTrackChangesEnabled()` — no per-call override.

## Run-formatting encoding: hybrid HTML (the read↔import contract)

Run-level formatting with no native markdown syntax is emitted as **HTML a markdown reader actually renders** — not Pandoc `[text]{…}` spans, which show literal brackets in GitHub / VS Code / Obsidian. `read --markdown` emits it (via `wrapRunFormatting` in [cli/read/markdown.ts](../../cli/read/markdown.ts)) and the import walker parses it back (via `gatherHtmlSpans` in [inline-surgery.ts](inline-surgery.ts), which converts each tag to a `bracketedSpan` the inline walker already overlays). **The tag → property mapping is the contract; keep `wrapRunFormatting` and `gatherHtmlSpans`/`applyCssStyle` in sync.** `read --ast` is the lossless format; markdown is the human/agent comprehension format.

Three carriers — semantic tags (render everywhere, incl. GitHub), `<span style>` (CSS-expressible props; render in editors/browsers; GitHub strips `style`), and `data-*` attributes (OOXML-only props CSS can't say; ignored by renderers, kept in source so markdown stays lossless too):

| Run prop | Emitted as | Notes |
| --- | --- | --- |
| highlight | `<mark>` / `<mark data-highlight="green">` | bare `<mark>` = yellow; other of the 16 names via `data-highlight` (enum-validated; **no hex** — use shade) |
| superscript / subscript | `<sup>` / `<sub>` | run-level only — math is `$…$` |
| underline | `<u>` / `<u data-underline="double" data-underline-color="FF0000">` | bare `<u>` = single; all 18 `ST_Underline` styles + color via `data-*` |
| color (hex) | `<span style="color:#FF0000">` | |
| shade (bg hex) | `<span style="background-color:#FFE599">` | `<w:shd w:fill>` — arbitrary hex |
| theme color | `<span data-color-theme="accent1" data-color-theme-tint= data-color-theme-shade=>` | `<w:color w:themeColor/…>` — CSS has no theme concept, so `data-*` only (byte-exact) |
| font | `<span style="font-family:Arial">` | quoted when it has spaces (`'Times New Roman'`); literal fonts only (theme fonts inherit) |
| size | `<span style="font-size:12pt">` | `<w:sz>` half-points; **omitted when it equals the document baseline** (see below) |
| smallCaps / allCaps | `<span style="font-variant:small-caps">` / `text-transform:uppercase` | |

bold/italic/strike/links/code stay native markdown (`**`/`*`/`~~`/`[](…)`/`` ` ``). Wrappers nest innermost→outermost `<span>` → `<u>` → `<sup>`/`<sub>` → `<mark>`, a fixed order `gatherHtmlSpans` reverses. The `<w:rPr>` emitter is still duplicated across `core/blocks.tsx::RunProperties` (AST→XML) and `inline.tsx::RunProperties` (markdown→XML); both MUST emit identical child order (CT_RPr §17.3.2.28) — enforced by the two-emitter convergence test in [tests/cli/markdown.test.ts](../../../tests/cli/markdown.test.ts).

**Document baseline note.** A leading `<!-- docx:base font="Arial" size="8pt" -->` (`formatBaseNote`) declares the dominant font/size across the doc; `read` then omits those from every matching run so the body reads clean instead of repeating `font-family:Arial` on every span, AND so an agent can see the doc's baseline to match new content. The declared value is the per-run majority (a >50%-of-text dominant), **else the document default** (`docDefaults`/Normal, via `StylesView.defaultFont`/`defaultSizeHalfPoints`) when it DEVIATES from the canonical template (Calibri 11pt) — i.e. someone ran `styles set-default-font`. The universal template default is suppressed as noise (a `docx:base` on every plain doc would be useless); a deviation is the meaningful signal, and surfacing it is what makes `set-default-font` observable on read (the write-read loop) — it has no other agent-visible read-back (`read --ast` doesn't model docDefaults, `styles --at Normal` reads the Normal style's rPr, not docDefaults). Deviation-only suppression is computed in `renderMarkdown`'s `noteBaseline` (the `deviation()` helper); a theme-only docDefaults (no explicit `w:ascii`) stays unsurfaced (we don't resolve the theme), but `set-default-font` always writes explicit ascii so its effect is caught. It is a **visibility hint, not parse-back** (per "comments are never anything but hints" in the root CLAUDE.md): the importer DROPS it (the leading note flows through as a block `html` node and `walkBlock` drops it), so a full `read → create` rebuild falls back to the template `docDefaults` for the dominant font/size — which is exactly why declaring the deviating default in the note is safe (it can't drift the round-trip). `read --ast` stays lossless (every run's font/size is there) and in-place `edit` never touches runs, so only the from-scratch rebuild is lossy. black (`000000`/`auto`) and the `text1`/`dark1` theme are dropped as noise (the universal default).

**Enum validation is on every ingress.** `highlight`/`underline`/`vertAlign` are closed OOXML enums; an out-of-range value would make Word silently drop schema-invalid XML, so both the import path (`inline.tsx::validateSpanFormatting` → `MarkdownImportError` USAGE — fed by both HTML and legacy-Pandoc spans) AND the `--runs` JSON path (`cli/parse-helpers.ts::parseRunsArg` → USAGE fail) validate against the shared sets in [@core/run-formatting](../run-formatting.ts). Hex colors / theme tokens / font names pass through unvalidated (Word degrades unknowns gracefully).

**Metacharacters are escaped PARSER-DRIVEN — content is never lost or corrupted (invariant II).** Every run's text — plain prose AND the text inside an HTML wrapper — is backslash-escaped per a mask from `inlineEscapeMask` ([escape.ts](escape.ts)). There are NO hand-rolled pairing/flanking/link rules to drift from remark: the mask is built by **parsing the run content with the importer's own parser** (remark-parse + gfm + math) and escaping exactly the punctuation the parser CONSUMES into a construct — every character it keeps as a literal `text` node is left byte-clean. That is why a `[ x ]` checkbox, a `[Fill in …]` placeholder, a lone `$5`, an `R&D`, or a `5 < 10` all stay clean (the parser reads them as plain text) while a paired `$…$`, a `[label](url)`, a `<div>`, or an `&amp;` get escaped. `read` applies the mask via `applyEscapeMask`, threading each run's offset into the scope-wide content.

- **Scope.** The mask is parsed over the whole paragraph — and for a table cell, over every paragraph in it joined by `\n` (the cell renders as one `<br>`-joined line, so a `$` in one cell-paragraph pairs with a `$` in the next; the `\n` lets math pair across the boundary while stopping a `]`-ending paragraph from fusing with a `(`/`[`-opening one into a phantom link). `renderCell` gives each paragraph its base offset (`+ 1` per `\n`).
- **Wide leaf constructs.** `inlineMath` and `inlineCode` hold their content as a `value` string, not a child `text` node, so the mask marks their INTERIOR literal and escapes only the boundary `$`/backtick — escaping the boundary breaks the construct, so a `[placeholder]` that fell between two paired `$` across a cell needn't be touched. Every other construct (link, emphasis, image, html) keeps its interior as a child `text` node, so the span scan already escapes only its delimiters.
- **Decode triggers.** Two things the parser HIDES by decoding (so the span scan can't see them) are re-flagged explicitly: a character reference (`&amp;` → `&`) escapes its `&`, and a backslash escape (`\*` → `*`) escapes its `\`.
- **Equations.** A run-level equation in the scope emits a real `$`/`$$` that would pair with an otherwise-lone text `$`, and isn't part of the parsed content — so `hasEquation` escapes every text `$` in that scope.

remark decodes `\x` → `x` on import (CommonMark backslash-escapes every ASCII punctuation char), so escaped or not, the text round-trips byte-exact. Inline **code RUNS** (`runStyle: "Code"`) are excluded from the parsed content entirely — they ride verbatim between backticks. `escape.ts`'s parser plugin list MUST mirror `parseToMdast` in [import.tsx](import.tsx), or what `read` leaves unescaped could diverge from what the importer reads back. Attribute values are HTML-escaped (`&quot;` etc., via `htmlAttr`) so a crafted font name with a `"` can't close the attribute early and inject a sibling. So unlike the old Pandoc spans (which had to DROP formatting on metacharacter text), the wrapper keeps both the text AND its formatting. Unmatched/unknown HTML degrades safely: the inline walker drops a raw `html` node, so an unrecognized tag loses its formatting but keeps its content. Legacy Pandoc `[text]{…}` spans are still parsed on import for backward compatibility, but no longer emitted.

## Structural visibility annotations (read-only)

`read --markdown` surfaces structural facts the GFM body can't show as HTML
comments shaped `<!-- docx:TYPE [bareId] key="value" … -->` (emitted via
`formatNote` in [cli/read/annotations.ts](../../cli/read/annotations.ts), escaped
with `htmlAttr`): `docx:section` (section breaks), `docx:page` (page geometry),
`docx:table` (uneven column widths + border summary), `docx:cell` (per-cell
merge/shading, carrying the cell address), and `docx:image` (size always +
float/wrap/align/overflow deviation-only, carrying the `imgN` id).

**These are read-time VISIBILITY hints, not round-trip carriers — ONE contract.**
The importer DROPS every one: `walkBlock`'s `case "html"` returns `[]`, and the
inline walker drops inline comments. The structure survives normal edits via
in-place XML mutation; `read --ast` is the lossless view; and the authoring verbs
(`docx sections`, `docx tables …`) manage it. So a from-scratch
`create` is deliberately lossy for non-Markdown structure (a section break read →
created vanishes rather than corrupting into a border paragraph) — but the agent
SAW it in the read output. Re-emitted fresh each read, so they never accrete.

**Naming rule — bare = locator, `docx:` = metadata.** A bare comment is a
LOCATOR (an address: `<!-- p0 -->`, `<!-- t0:r0c0:p0 -->`). Anything docx-cli adds
BEYOND addressing is a `docx:TYPE` annotation (`docx:section`, `docx:page`,
`docx:table`, `docx:cell`, `docx:base`), which may carry the relevant locator as a
bare leading token (`<!-- docx:cell t0:r0c0 gridSpan="2" shading="FFE699" -->`).
Metadata never rides a bare locator comment — that keeps locators short and makes
`docx:` a clean grep for "everything docx-cli added."

**Deviation-only:** emit an attribute only when it differs from the document
default (the `docx:base` size-suppression lesson) — a note that repeats the
default is noise, not signal. Even columns, default Letter geometry, plain
borderless tables, `align=left` → emit nothing.

**No exceptions — `docx:base` is a hint too.** The run-formatting baseline note
(`docx:base font=… size=…`) was once parse-back, but no longer: the importer drops
it like every other comment (see the "Document baseline note" section above), so a
from-scratch `read → create` rebuild falls back to the template `docDefaults`
(Calibri 11pt) for the dominant font/size. `read --ast` stays lossless and
in-place `edit` never rewrites runs, so only the full rebuild is lossy. There is
NO comment, anywhere, that drives reconstruction.

**Section breaks specifically** no longer render as bare `---`. `---` round-trips
as a thematic break (`<HorizontalRule>` → a border paragraph), so emitting it for
a section silently corrupted layout on `read → create` AND was indistinguishable
from a real thematic break. Now a section is an own-line `docx:section` visibility
comment and a hand-authored `---` unambiguously means a thematic break.

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
- **Locator HTML comments** — `<!-- p3 -->` markers are dropped by the inline / block walkers. The locator metadata is already implicit in block order, so round-trips reconstruct the locators on the read side. The structural visibility annotations (`docx:section`/`docx:page`/`docx:table`) are dropped the same way — see "Structural visibility annotations" below.
- **Image dimensions overrides** — there's no `--image-width` style flag on `--markdown`; we use the source's intrinsic pixel dimensions. Pass `--width` via standalone `insert --image` for explicit sizing.
- **Strict CriticMarkup substitution** (`{~~old~>new~~}`) — only `{++…++}` and `{--…--}` are tokenized; substitution would need a third mdast node kind plus a paired `<w:del>` / `<w:ins>` emit.
