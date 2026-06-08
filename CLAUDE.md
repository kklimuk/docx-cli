# docx-cli

CLI for AI agents to read, edit, and comment on `.docx` files. JSON-AST output, locator-based addressing, full format fidelity via in-place XML mutation.

**Bun, not Node.** Use `Bun.file`, `Bun.write`, `Bun.env`, `Bun.$`. Bun loads `.env` automatically — no dotenv.

Subsystem-specific guidance lives in nested CLAUDE.md files that load when you edit those folders. If you need to add a new CLAUDE.md to describe a new practice for a part of the system, do so.

## Conventions

These conventions are NOT SUGGESTIONS. These are rules.

- **All stdout goes through `respond()` (JSON ack) or `writeStdout()` (text)** from `src/cli/respond.ts` — never `process.stdout.write`. Both use `Bun.write(Bun.stdout, ...)`; the 64 KB truncation that bites on early exit is real and silent, and these helpers are the only safe path.
- **File naming**: kebab-case, named after the primary export (`xml-node.ts` → `XmlNode`).
- **Newspaper ordering.** The entry point (primary export) goes at the top; its dependencies follow in the order it uses them, then _their_ dependencies, and so on — a file reads top-to-bottom like a newspaper. Use hoisted `function` declarations for internal helpers so this works at runtime; arrow functions only for inline callbacks and short utilities. Types are usually not the primary exports and should go below the functions/classes that are.
- **Feature nesting** When a file accumulates too many dependencies to be read well with newspaper ordering (> 300 lines), split them into a separate folder/file named after the feature they're working on. It should be a folder if it is going to represent a logical feature of dependencies. This nesting can continue indefinitely if subfeatures have subfeatures of their own.
- **JSX is for emitters only.** Files that construct fresh XML can be `.tsx`; readers/locators/analysis stay `.ts`. Components are PascalCase, accept props, may return `NullableXmlNode` (null skipped by flatten). Attribute names with colons use the hyphen shortcut (`w-val="x"` → `w:val="x"`) or JSX spread.
- **Component vs view vs lens vs free function**: four shapes, one decision tree.
  - A pure `props → XmlNode` builder is a PascalCase **component** — destructure its props in the signature (no `props.x` access), and don't take a `Document` (or any package state).
  - Stateful OOXML state lives in **tree-owning views**, embedded as fields on `Document`: `Body`, plus one view per OPC part — `StylesView`, `NumberingView`, `CommentsView`, `NotesView`, `RelationshipsView`, `ContentTypesView`, `SettingsView`, `CorePropertiesView`. Each owns its part's `XmlNode` tree and any maps keyed to it, and exposes a `fromPackage`/`fromXml`/`writeTo` lifecycle (`register` too, for the lazily-provisioned ones). Cross-view dependencies (e.g., `NotesView.ensureNoteStyles(stylesView)`) are passed as method arguments — no view reaches up to `Document`.
  - **Cross-cutting lenses** (`Images`, `Hyperlinks`, `Equations`, `TrackChanges`, `Comments`) are NOT fields on `Document` — they're stateless, constructed at the call site: `new Images(document).add(source)`, `new TrackChanges(document).accept(["tc0"])`. They hold only a back-reference; the embedded views are the state they reach through.
  - **Free functions** are reserved for: pure builders (the components above), the AST reader (`src/core/ast/read.ts` — Document's construction pass; populates the embedded views from XML and is the sole assigner of `tcN` ids), and emitter helpers in `src/core/blocks`/`table`/`sections` that thread a `Document` because they touch many slices in one call. If a free function's body operates on one slice of `document`, make it a method on that slice's view instead.
- **JSX.Element = XmlNode** (single, not nullable). `Fragment` returns a `#fragment` sentinel unwrapped in `flatten()` and `serialize()`. Components return `null` to render nothing; `jsx()` converts that to an empty fragment. The `jsx`/`jsxs`/`jsxDEV` runtime exports are distinct functions, not `= jsx` aliases (knip flags aliased re-exports as duplicates) — don't collapse them.
- **Path aliases**: `@core` → `src/core/index.ts`, `@core/*` → `src/core/*`. Use these in `src/cli/*`; `src/core` itself uses relative sibling imports. Import the body emitters from the `@core/blocks` and `@core/table` subpaths, not the `@core` barrel — `ast/types` already exports `Paragraph`/`Table`/`TableCell`/`TableRow` as _types_, and barrel-merging the same-named value emitters is confusing.
- **Variable names**: descriptive, no single/two-letter (`paragraph` not `p`). Exception: regex-match destructuring (`const [, prefix, idx] = match`).
- **Inline props in the signature.** When a component's props type is used only by that component, write it inline (`function HeadingStyle({ styleId }: { styleId: BaselineStyleId; … })`) rather than declaring a separate named `Props` type. Extract a named type only when it's shared.
- **knip runs strict** (`bun run check`, no rule overrides in `knip.json`). An unused export is dead code — delete it (this is a CLI app, not a library; there are no external `@core` consumers). The one exception: an export staged for a named upcoming tier with no caller yet gets a `@public` JSDoc tag whose comment names the future consumer (knip honors `@public`) — e.g. `HorizontalRule` (S8) and the `r`/`a`/`wp`/`pic` image namespaces (S5). Don't silence knip by re-adding rule suppressions.
- **Style**: tabs, double quotes (Biome enforced). Early returns over else-if chains.

## Invariants

These invariants are NOT SUGGESTIONS. These MUST be followed.

- **In-place XML mutation, not AST round-trip** — the AST is a view; mutate `XmlNode` refs, only emit fresh XML for inserted nodes. See [src/core](src/core/CLAUDE.md).
- **Never delete a relationship or part something still references** — a dangling rId corrupts the file ("unreadable content"), but an unreferenced part is harmless. So pruning is gated on a reference check that scans **everything we don't model** (VML `<v:imagedata>` fallbacks, OLE objects, `<w:background>`, chart rels), not just the construct we authored: `isRelationshipReferenced(documentTree, rId)` before dropping a relationship, `hasRelationshipWithTarget` before deleting a shared media part (both in `core/relationships.ts`). When in doubt, leave the orphan.
- **`RUN_BEARING_WRAPPER_TAGS`** in `src/core/parser/run-ops.ts` is the AST↔XML offset bridge; every offset-aware walker reads it. See [src/core](src/core/CLAUDE.md).
- **Stable positional ids** (`p0`, `t0`, `c0`, `img0`, `link0`, `tc0`). Block ids shift after structural edits — re-read between non-trivial mutations. Comment ids are `max-existing + 1`.
- **Hyperlinks own a relationship, not their text.** `hyperlinks replace` updates the `<Relationship>` `Target` (mints a new rId if multiple `<w:hyperlink>` share one); `delete` unwraps and prunes the rId when unreferenced.
- **paraId is required for resolve/reply** — auto-injected via `ensureCommentParaId()` rather than failing.
- **Track-changes is doc-level** — see [src/cli/track-changes](src/cli/track-changes/CLAUDE.md).
- **Sections are blocks; CRUD goes through the standard verbs** — see [src/core](src/core/CLAUDE.md).
- **Table structure is a merge-aware logical grid** — `gridSpan`/`vMerge` map logical (row,col) onto physical `<w:tc>`. The `docx tables` verbs reshape rows/columns/merges/widths/borders through that model — see [src/cli/tables](src/cli/tables/CLAUDE.md).
- **Markdown is GFM + math + CriticMarkup + inline HTML formatting, parsed via remark.** The `MarkdownImport` lens (`src/core/markdown/`) composes existing emitters; the walker is sync (image fetch + footnote registration happen in an async pre-walk). Run formatting with no native markdown syntax (color/theme color, highlight, shading, underline, super/subscript, caps, font, size) is emitted as **HTML a reader renders** — `<mark>`/`<sup>`/`<sub>` semantic tags, `<span style>` for CSS props, `data-*` attributes for OOXML-only props — and a leading `<!-- docx:base font=… size=… -->` note declares the document's dominant font/size once (omitted per-run, re-applied on import). The `remarkInlineSurgery` transform gathers CriticMarkup, legacy Pandoc spans, AND HTML tags across phrasing siblings (remark leaves inline HTML as flat unpaired tokens, re-paired via `gatherHtmlSpans`). `read --ast` is the lossless format. See [src/core/markdown](src/core/markdown/CLAUDE.md).
- **`docx render` is the only external-runtime command.** Every other verb works purely against the .docx zip + XML; `render` shells out to Word (macOS/Windows) or LibreOffice (cross-platform) to produce a PDF, then rasterizes in-process via the bundled `@hyzyla/pdfium` WASM package (no system tools needed for the rasterizer). The lens lives at [`@core/render`](src/core/render/CLAUDE.md) (`renderDocxPages` is the entry point); [`src/cli/render`](src/cli/render/CLAUDE.md) is a thin arg-parse shell. Agents that consume PNGs (and you, when verifying a fixture) use this command; the rest of the CLI never invokes it.
- **No undo, no journal.** Mutating commands overwrite `FILE` in place; git is the history. `-o/--output PATH` writes a parallel file; `--dry-run` previews (wins over `--output`).

## Commands

`docx <verb>` and `docx <noun> <verb>`. Every command has `--help`. Mutating commands accept `--dry-run` and `-o/--output PATH` — **except `create`, whose positional FILE is already the output path** (it has no `-o`). Real-tracked-change mutators (`edit`/`insert`/`delete`/`replace`, the `tables` verbs, `images delete`) also accept `--track` to record one invocation as tracked even when the doc toggle is off.

**Locator flags are unified:** address an existing thing with **`--at LOCATOR`** everywhere (edit, delete, tables, comments, footnotes/endnotes, images, hyperlinks, track-changes). `insert` uses `--after`/`--before` (relative placement), `read` uses `--from`/`--to` (a block slice), `comments add` also takes `--anchor PHRASE`, and `wc` takes a positional `[LOCATOR]`. No more `--id`/`--range`/`--to` for addressing.

**`--batch FILE.jsonl` (one JSONL object per line; `-` = stdin)** lets `edit`, `insert`, `replace`, and the `comments` verbs apply many changes from a single read — keys mirror the command's flags. The shared reader is `readJsonlObjects` (`cli/parse-helpers.ts`); each command's batch path lives in a sibling `batch.ts`. **All locators in a batch address the document AS READ** — the load-bearing invariant — and each command keeps that true differently: `edit` resolves every entry to a live `XmlNode` ref *before any mutation* (same-paragraph spans apply right-to-left; a paragraph takes one whole-paragraph edit OR several non-overlapping spans); `replace` re-reads the live tree between entries (`document.reread()` re-walks `documentTree`, no re-parse, so node identity holds) for sed-like sequential semantics; `insert` pins anchors to live refs, builds all blocks with zero body mutation, then splices with a per-anchor offset so stacked inserts keep entry order. Range/section/equation `edit`s aren't batchable (do them one at a time).

**Output model — exit code is the success signal** (`0` ok, `1` general, `2` usage, `3` not-found, in `src/cli/respond.ts`). The `ok` field appears ONLY on the `--verbose` ack. So:
- A mutator that mints a new addressable handle (`comments add`→`cN`, `comments reply`→`cN`, `footnotes/endnotes add`→`fnN`/`enN`, `hyperlinks add`→`linkN`, `insert`→the new `pN`) prints the bare locator(s) — one per line — by default (`respondMinted`); `--verbose` → full `{ok:true,…}` ack.
- A mutator with no new handle (`edit`/`delete`/`replace`/`comments resolve`/`tables *`/`track-changes *`/`toggle`) is silent on success (`respondAck`; `--verbose` → ack).
- Query commands are text-first where there's a clean plain form: `find`→locator lines, `wc`→count, `outline`→indented `locator␉text` tree, `read`→markdown — each with `--json`/`--ast` for the structured payload. List verbs print a bare JSON array whose items' `id` is the `--at` handle.
- Errors print `{code, error, hint?}` (no `ok`) + a nonzero exit. Dry-run previews drop `ok` too.

| Surface         | Verbs                                                                                                                              |
| --------------- | ---------------------------------------------------------------------------------------------------------------------------------- |
| top-level       | `create` (`--from PATH.md`) `read` `insert` (`--markdown TEXT` / `--markdown-file PATH`; `--batch`) `edit` (`--markdown` on `pN` or `pN-pM`; `--clear` to strip formatting; span/cell locators `pN:S-E`/`tN:rRcC:pK`; `--batch`) `delete` `find` (text, or `--highlight`/`--color`/`--bold`/`--italic`/`--underline` to locate by formatting) `replace` (`--batch`) `wc` `outline` `render` (per-page PNG via Word or LibreOffice) (insert: `--code [--language LANG]` / `--code-file PATH`) |
| `comments`      | `add` `reply` `resolve` `delete` `list`                                                                                            |
| `footnotes`     | `add` `edit` `delete` `list`                                                                                                       |
| `endnotes`      | `add` `edit` `delete` `list`                                                                                                       |
| `images`        | `add` (alias for `insert --image`) `list` `extract` `replace` `delete`                                                            |
| `hyperlinks`    | `add` `list` `replace` `delete`                                                                                                    |
| `tables`        | `insert-row` `delete-row` `insert-column` `delete-column` `set-widths` `merge` `unmerge` `borders`                                 |
| `track-changes` | `FILE on\|off` (toggle), `list FILE`, `accept`/`reject FILE (--at tcN \| --all)`                                                   |
| `info`          | `schema` `locators`                                                                                                                |

`docx read FILE` renders GFM by default (`--ast` for JSON-AST). Tracked changes have three views: default `accepted` (drops subtractive, inlines additive — clean text), `--current` (CriticMarkup `{++ins++}` / `{--del--}` with `[^tcN]` footnotes), `--baseline` (drops additive, inlines subtractive). `--comments` appends `[^cN]` footnotes. `wc` accepts the same `--accepted`/`--baseline`/`--current` flags. Render/locator logic in `cli/read/markdown.ts`; feature detection in `core/ast/read.ts`.

## Locators

```
pN              paragraph N        pN:S-E          chars S..E within paragraph N
pN-pM           paragraph range    pN:S-pM:E       cross-paragraph (chars)
                                   tN / tN:rRcC    table N / cell at row R col C (chainable :pK)
sN              section break N    cN imgN linkN tcN   comment / image / hyperlink / tracked-change ids
                                   fnN enN          footnote / endnote ids (`footnotes`/`endnotes` verbs)
tN:rR tN:cC     table row R / column C (the `tables` verbs) tN:rR1cC1-rR2cC2   rectangular cell region (merge)
```

Nested tables chain the same syntax (Word emits them for compound rubric/layout cells). A cell prefix can be followed by another `tN:…` to address a table inside that cell, with `:rR`, `:cC`, `:rRcC`, `:rRcC:pK`, etc. composing arbitrarily deep — e.g. `t0:r2c1:t0:r0c0:p0` is the first paragraph of the (0,0) cell of the first table nested inside the (2,1) cell of the document's first table. Every surface that takes a locator (`read`, `wc`, `find`, `insert`, `edit`, the `tables` and `comments` verbs) accepts the chained form.

Span comments and `hyperlinks add` split runs at offsets, preserving `<w:rPr>` on both halves (`core/comments/markers.tsx`, `core/hyperlinks/wrap.tsx`).

## Testing

```bash
bun run test:unit         # core + cli (~3s)
bun run test:integration  # LibreOffice round-trip (auto-skips if no soffice)
bun test                  # everything
bun run check             # biome + knip + tsc
```

Fixture authoring (core-emitters-first, CLI-second) and rebuild instructions live in [tests/fixtures/setup/CLAUDE.md](tests/fixtures/setup/CLAUDE.md).

## Build

- `bun run build` → `dist/index.js` — bundled JS that npm publishes (runs under Bun). **Required**: path aliases (`@core/*`) and JSX runtime resolution don't work when consumed from `node_modules`; the bundle pre-resolves everything. Never ship raw `src/`.
- `bun run build:binary` → `dist/docx` — standalone executable for GitHub Releases.

## Docs layout

Three docs, each for a different reader — don't cross the streams. When you change one, check the others:

- **README.md** — user-facing: install, examples, command reference, "How It Works."
- **CONTRIBUTING.md** — dev-facing: setup/test commands, LibreOffice install, project-structure tree, CI table.
- **CLAUDE.md** _(this file + nested)_ — agent-facing: conventions, invariants, per-subsystem playbooks.

New CLI command → README + `src/cli/CLAUDE.md`. New invariant → CLAUDE.md only. New build/test step → CONTRIBUTING.md + this Testing section. New runtime dep → README + CONTRIBUTING.md.
