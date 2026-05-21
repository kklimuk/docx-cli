# docx-cli

CLI for AI agents (Claude, Codex) to read, edit, and comment on `.docx` files. JSON-AST output, locator-based addressing, full format fidelity via in-place XML mutation.

Default to **Bun**, not Node. Use `Bun.file`, `Bun.write`, `Bun.env`, `Bun.$` over Node equivalents. Bun loads `.env` automatically — don't use dotenv.

## Docs layout

Three docs, each for a different reader. Don't cross the streams.

- **README.md** — user-facing. Hero (Loom embed + tagline), install, the NDA quick-example, full command reference, markdown-rendering details, locators, "How It Works" for engineers digging in. Anything a drive-by visitor or returning user wants.
- **CONTRIBUTING.md** — dev-facing. `bun install`/`bun dev`/test commands, LibreOffice install for integration tests, the `src/` architecture tree, CI job table. Pointers back to this file for deeper conventions. External visitors don't need this in the README.
- **CLAUDE.md** _(this file)_ — agent-facing. Conventions, architectural invariants, what to do when editing different parts of the codebase, the full project structure walkthrough. Lives at the repo root so it's auto-loaded by Claude Code.

When you change one, check whether the others need a parallel edit:

- New CLI command → README (Commands block + maybe Quick example) + CLAUDE.md ("When Editing")
- New architectural invariant → CLAUDE.md only (it's not user-facing)
- New build/test step → CONTRIBUTING.md + CLAUDE.md (Testing section)
- New top-level dependency or runtime requirement → README (Stack — currently lives in README; check if still there) + CONTRIBUTING.md if it affects dev setup

## Project Structure

```
src/
  index.ts                    # binary entrypoint (#!/usr/bin/env bun)
  cli/
    index.ts                  # parseArgs dispatcher, COMMANDS map
    help.ts                   # top-level --help
    respond.ts                # respond() / fail() — JSON ack + error helpers
    create/                   # docx create FILE
    read/                     # docx read FILE (markdown default, --ast for JSON)
      markdown.ts             # GFM renderer — the default for `docx read`
    insert/                   # docx insert FILE
      emit.tsx                # <Paragraph> + <RunElement> JSX components (shared by edit too)
    edit/                     # docx edit FILE
    delete/                   # docx delete FILE
    find/                     # docx find FILE QUERY
    replace/                  # docx replace FILE PATTERN REPLACEMENT
      replace-span.tsx        # in-place run-splitting for span replacements
    wc/                       # docx wc FILE [LOCATOR]
      count.ts                # word-count helpers over the typed AST
    outline/                  # docx outline FILE
      build.ts                # heading-tree builder
    comments/                 # docx comments <verb>
      index.ts                # sub-dispatcher for add/reply/resolve/delete/list
      helpers.tsx             # paraId, ensureCommentsPart, run-splitting marker injection
      add | reply | resolve | delete | list
    images/                   # docx images <verb>
      list | extract | replace
    hyperlinks/               # docx hyperlinks <verb>
      index.ts                # sub-dispatcher for add/list/replace/delete
      wrap.tsx                # in-place run-splitting for hyperlinks add
      add | list | replace | delete
    track-changes/            # docx track-changes <verb>
      index.ts                # sub-dispatcher: routes accept/reject/list to verb files; falls through to toggle
      toggle.tsx              # FILE on|off (sets/clears <w:trackChanges/>)
      list.ts                 # JSON inventory of every revision wrapper (<w:ins>/<w:del>/<w:moveFrom>/<w:moveTo>)
      accept.ts | reject.ts   # thin wrappers over apply.ts
      apply.ts                # collectTrackedChanges + unwrap/delete/renameDelTextToText (handles all four kinds)
    info/                     # docx info <topic>
      index.ts                # sub-dispatcher for schema/locators
      schema.ts               # docx info schema [--ts]  (TS source via Bun text import)
      locators.ts             # docx info locators [--json]
  core/
    package/                  # JSZip wrapper: open, read/write XML parts, save
    parser/
      xml-node.ts             # XmlNode class — instance methods + static parse/serialize
      run-ops.ts              # runTextLength / sliceRun / sumRunBearingTextLength + RUN_BEARING_WRAPPER_TAGS allow-list
      index.ts                # re-export
    jsx/
      index.ts                # h, Fragment, makeTag, namespaces (w, r, a, wp, pic, cp, dc, …)
      jsx-runtime.ts          # auto-runtime — converts component-null to empty fragments
      jsx-dev-runtime.ts      # alias for dev-mode auto-runtime
    ast/
      types.ts                # Doc / Block / Run / Comment / Footnote types (read live by `docx info schema --ts`)
      doc-view.ts             # DocView, openDocView, saveDocView, enrichImageHashes — also loads footnotes.xml/endnotes.xml
      read.ts                 # XML → typed AST walker; handles oMath/oMathPara, footnote/endnote refs, charts/SmartArt/shapes, tracked moves, transparent wrappers (fldSimple/smartTag), single-character markers (noBreakHyphen/softHyphen/sym), w:cr / w:ptab / w:pict / w:object. Surfaces inline + trailing <w:sectPr> as sN blocks (with columns / sectionType), and registers <w:sectPrChange> + paragraph-mark <w:ins>/<w:del> in trackedChangeReferences.
      text.ts                 # paragraphText / paragraphTextAccepted / paragraphTextBaseline / flattenParagraphs / findBlockById — shared text helpers (accepted/baseline drop the matching tracked-change kinds)
      sym.ts                  # decodeSym(font, charHex) — Symbol / Wingdings / Wingdings 2 / Webdings / Zapf Dingbats lookup tables for <w:sym>
    sections.tsx              # sectPr emitters + mutators: SentinelSectionParagraph, applyColumns, applySectionType, removeInlineSectPr, isTrailingSectPr, wrapSectPrChange, readSectionProperties
    relationships.ts          # mintRelationshipId, addHyperlinkRelationship — _rels helpers
    locators/
      parse.ts                # parseLocator("p3:5-20") → Locator union
      resolve.ts              # resolveBlock / resolveComment / resolveImage / resolveHyperlink
tests/
  core/                       # XmlNode + locator parser unit tests
  cli/
    harness.ts                # Bun.spawn wrapper — runs binary, parses JSON output
    *.test.ts                 # one file per command surface
  integration/
    libreoffice-roundtrip.test.ts  # auto-skips if `soffice` not on PATH
  fixtures/                   # .docx files exercising different OOXML features (comments, tracked changes, footnotes, equations, tables, sections, ...). All fixtures are MIT-licensed; `minimal.docx`, `notes.docx`, `sections.docx`, `tracked-moves.docx`, `transparent-wrappers.docx`, `chained-tracked-edits.docx`, `normalize-query.docx`, `comments-batch.docx`, `word-formatted.docx`, and `multi-tracked.docx` are built by scripts in `scripts/make-*-fixture.ts`.
scripts/
  move.ts                     # move file + auto-update imports
  fxp-smoke.ts                # JSX → XML smoke test
  inspect-fixtures.ts         # summarize each fixture's features
  make-fixture.ts             # rebuild minimal.docx
  make-notes-fixture.ts       # rebuild notes.docx (footnotes + endnotes)
  make-tracked-moves-fixture.ts        # rebuild tracked-moves.docx
  make-transparent-wrappers-fixture.ts # rebuild transparent-wrappers.docx (fldSimple/smartTag)
  make-sections-fixture.ts             # rebuild sections.docx (all sectionTypes + sectPrChange) — dogfoods the CLI
  make-chained-tracked-edits-fixture.ts # rebuild chained-tracked-edits.docx (tracking on, baseline for chained replaces)
  make-normalize-query-fixture.ts       # rebuild normalize-query.docx (smart quotes, em-dash, * between digits)
  make-comments-batch-fixture.ts        # rebuild comments-batch.docx (3 paragraphs for batch + anchor tests)
  make-word-formatted-fixture.ts        # rebuild word-formatted.docx (mixed bold/italic/color spans for B3 diff tests)
  make-multi-tracked-fixture.ts         # rebuild multi-tracked.docx (3 paragraphs × 6 tracked changes for batch accept/reject tests)
  escape-check.ts             # XML entity round-trip sanity
```

## Key Conventions

These are not suggestions. Follow them.

- **Bun not Node**. `Bun.file(path).type` for MIME. All stdout goes through `respond()` (JSON ack) or `writeStdout()` (text) from `src/cli/respond.ts`, both of which use `Bun.write(Bun.stdout, ...)` — never `process.stdout.write` directly. The 64 KB truncation that bites on early exit is real and silent; the helpers are the only safe path.
- **File naming**: kebab-case, named after the primary export (`xml-node.ts` → `XmlNode`, `paragraph.ts` → `Paragraph`).
- **Feature structure**: each command is a kebab-case folder under `src/cli/`. `index.ts(x)` is the public surface (`export async function run(args: string[]): Promise<number>`). Sub-files for shared helpers, no cross-feature imports unless centrally exposed via `@core/*`.
- **Path aliases**: `@core` → `src/core/index.ts`, `@core/*` → `src/core/*/index.{ts,tsx}` or `src/core/*.{ts,tsx}` (configured in tsconfig). Use these in `src/cli/*` to avoid `@core` chains. `src/core` itself uses relative imports between siblings.
- **JSX is for emitters only**. Any file that constructs fresh XML can be `.tsx` — wherever it lives in the tree. Readers, locators, and pure analysis stay `.ts`; never JSX in the AST reader. Components are PascalCase (`<Paragraph>`, `<RunProperties>`); they accept props and may return `XmlNode | null` (null skipped by flatten). Tag namespaces are imported as lowercase (`w.p`, `cp.coreProperties`); attribute names with colons go through hyphen shortcut (`w-val="800080"` → `w:val="800080"`) or JSX spread (`{...{"w:val": "x"}}`).
- **Variable names**: no single/two-letter names. Use descriptive (`paragraph` not `p`, `commentId` not `cid`). Exception: standard regex-match destructuring (`const [, prefix, idx] = match`) since the convention is unambiguous.
- **Spacing**: tabs, double quotes (Biome enforced).
- **Top-down file ordering**: exports first, dependencies below — like a newspaper. Use `function` declarations (hoisted) for internal helpers so this works at runtime. Arrow functions only for inline callbacks and short utilities.
- **Early returns**: bail on validation failures and null checks. No else-if chains when an early return reads cleaner.

## Architectural Invariants

These are not architectural suggestions, but requirements. If you disagree, make the case about needing a different pattern for a feature.

- **In-place XML mutation, not AST round-trip emission**. The typed AST returned by `read` is a _view_ over the parsed XML tree. Mutations (insert/edit/delete/comments add) operate on the underlying `XmlNode` references via `BlockReference.parent.splice(...)`. Anything we don't model survives because we never re-emit untouched regions. Only emit fresh XML for nodes we're inserting (via JSX) — never round-trip whole subtrees through the AST.
- **`RUN_BEARING_WRAPPER_TAGS` is the bridge between AST and XML offsets**. Defined in `src/core/parser/run-ops.ts`. AST text and `find`'s offsets descend into every tag in this set; the XML-side walkers in `cli/comments/helpers.tsx`, `cli/replace/replace-span.tsx`, and `cli/hyperlinks/wrap.tsx` all do the same via `isRunBearingWrapper(tag)` / `sumRunBearingTextLength(children)`. They have to stay in sync — if the AST descends into a wrapper that the XML walkers don't (or vice versa), `find → replace` / `find → comments add` will misalign by the wrapper's inner-text length. Current set: `<w:ins>`, `<w:del>`, `<w:moveFrom>`, `<w:moveTo>`, `<w:hyperlink>`, `<w:fldSimple>`, `<w:smartTag>`. Adding any new wrapper that holds runs requires a single edit in `run-ops.ts` — every walker reads from the same predicate. Conversely: any tag NOT in the set is preserved by the catchall `push(child)` in every walker (we don't destroy what we don't model — see `tests/cli/preserve-unknown.test.ts`).
- **fast-xml-builder owns escaping**. On the JSX path, never manually escape — the builder handles entities correctly (uses `&apos;` for `'`, which fast-xml-parser decodes back). The static templates in `cli/create/template.tsx` carry no user-supplied content, so they don't need escaping at all.
- **JSX.Element = XmlNode** (single, not nullable union). `Fragment` returns a sentinel `#fragment` XmlNode that gets unwrapped both in `flatten()` (composition) and `XmlNode.serialize()` (top-level). Components that want to "render nothing" return `null`; `jsx()` converts that to an empty fragment so callers always see `XmlNode`.
- **Stable positional ids** (`p0`, `t0`, `c0`, `img0`, `link0`, `tc0`). Block ids shift after structural edits — agents must re-read between non-trivial mutations. Comment numeric ids are allocated as `max-existing + 1`. Image, hyperlink, and tracked-change ids are positional (document order).
- **Hyperlinks own a relationship, not their text.** `hyperlinks replace` updates the `Target` of the underlying `<Relationship>` — but if multiple `<w:hyperlink>` elements share the same `r:id`, it allocates a new rId so the others stay pointing at the original URL. `hyperlinks delete` unwraps the `<w:hyperlink>` (text survives as plain runs) and prunes the rId from the rels file when no other element references it.
- **paraId is required for resolve/reply**. Comments authored by external tools may lack `w14:paraId` on their bodies. The `resolve` and `reply` verbs auto-inject one via `ensureCommentParaId()` (also adds `xmlns:w14` to the `<w:comments>` root if missing). Do this rather than failing — agents shouldn't have to recreate comments.
- **Track-changes is doc-level, not per-command**. When `<w:trackChanges/>` is set in `settings.xml`, every mutating command (`insert`/`edit`/`delete`/`replace`) automatically emits `<w:ins>`/`<w:del>` markers — there is no per-command override flag. To make a one-off untracked edit, run `docx track-changes FILE off`, edit, then `track-changes on`. **Accept/reject bypass tracking** (`docx track-changes accept|reject FILE --at tcN | --all`): they directly mutate the XML tree without wrapping the change itself in `<w:ins>`/`<w:del>`. Additive wrappers (`<w:ins>`, `<w:moveTo>`) accept-unwrap and reject-delete; subtractive wrappers (`<w:del>`, `<w:moveFrom>`) accept-delete and reject-unwrap (with `<w:delText>` → `<w:t>` rename so the restored text is plain). moveFrom/moveTo halves are processed independently — `--all` handles a complete move; targeting one half by `tcN` leaves the other in place. Targets walk fresh on each invocation (stored `view.trackedChangeReferences` would go stale across mutations) and process in reverse pre-order so nested changes are applied before their parents. Walk order mirrors the AST reader's so `tcN` ids agree between `track-changes list` and `accept --at tcN`: per paragraph it visits run-level wrappers first, then any inline `<w:sectPrChange>`, then any paragraph-mark `<w:ins>`/`<w:del>`. Paragraph-mark trackings (a self-closing `<w:ins>`/`<w:del>` inside `<w:pPr><w:rPr>`) are full citizens of accept/reject: accept-ins / reject-del simply remove the marker (paragraph stays); reject-ins removes the entire owning paragraph (the inserted break disappears — for sentinels created by `insert --section` this collapses the section break too); accept-del merges the owning paragraph with the next paragraph (the next paragraph's runs are appended to this one and the next paragraph is removed) per ECMA-376 §17.13.5.4. Still out of scope: `<w:rPrChange>`, `<w:pPrChange>`. Author/date come from the per-call `--author NAME` flag, then the `DOCX_AUTHOR` env var, then the `"docx-cli"` default (resolved in `core/track-changes/index.ts → resolveAuthor`); `DOCX_CLI_NOW` injects a fixed date for tests. `delete tN` rejects with `TRACKED_CHANGE_CONFLICT` when tracking is on (tracked table-row deletion isn't supported).
- **Hyperlink and image edits emit audit comments under track-changes**. OOXML has no `w:hyperlinkChange` / `w:drawingChange` element — Word itself silently bypasses tracking for hyperlink edits and image swaps. We compromise: when `<w:trackChanges/>` is on, `hyperlinks add/replace/delete` and `images replace` each auto-emit a `[docx-cli] …` comment anchored to the affected span/run, attributed via the same `--author` chain. The mutation itself stays silent (no fake `<w:ins>`/`<w:del>` since OOXML has no honest construct for it). Helpers live in `cli/comments/helpers.tsx` (`emitAuditComment`, `findContainingParagraph`, `findElementOffsetsInParagraph`, `addCommentMarkersAroundRun`). When track-changes is off, no comment is emitted.
- **Sections are blocks; CRUD goes through the standard verbs.** `<w:sectPr>` surfaces in the AST as `SectionBreak` blocks with `id` (`sN`), optional `columns`, and optional `sectionType` (`continuous` / `nextPage` / `evenPage` / `oddPage` / `nextColumn`). Per ECMA-376 §17.6.22 the `<w:type>` describes where the **current** section _begins_, not where the next one starts: `continuous` → this section begins on the same page as the prior section; `nextPage` → on a new page; `nextColumn` → at the top of the next available column; `evenPage` / `oddPage` → on the next even / odd page (with a blank parity-filler page inserted by the viewer if needed). Trailing `<w:sectPr>` (mandatory in OOXML) and inline `<w:pPr><w:sectPr>` (defines the section ENDING at that paragraph) both enumerate; the inline case puts an extra `sN` block right after its owning `pN`. CRUD: `insert --after pN --section [--columns N] [--type T]` emits a sentinel paragraph carrying an inline sectPr; `edit --at sN [--columns N] [--type T]` mutates the targeted sectPr's `<w:cols>` / `<w:type>` children in place; `delete --at sN` strips the inline sectPr from its owning paragraph (the paragraph stays — agents can `delete --at pN` separately) and rejects on the trailing one. Helpers in [src/core/sections.tsx](src/core/sections.tsx): `SentinelSectionParagraph`, `applyColumns`, `applySectionType`, `removeInlineSectPr`, `isTrailingSectPr`, `wrapSectPrChange`, `readSectionProperties`. Track-changes integration: `insert --section` is tracked via the existing paragraph-mark `<w:ins>` mechanism (sentinel has no runs so only the paragraph mark gets marked); `edit --at sN` under tracking emits a real `<w:sectPrChange>` snapshot inside the live sectPr (accept removes the snapshot; reject restores its children) — `tcN` enumeration includes sectPrChange, and `track-changes list` enriches sectPrChange entries with `prior: { columns?, sectionType? }` (extracted from the snapshot) and `current: { columns?, sectionType? }` (extracted from the live siblings) so agents can see the diff without re-reading XML — note `read --accepted/--baseline` only switches RUN-level views, not section properties, so the enriched list is the only way to see the prior section state; `delete --at sN` under tracking emits a `[docx-cli] section break removed` audit comment when the owning paragraph has runs to anchor on (sentinel paragraphs can't be anchored, so the mutation is silent in that case — consistent with hyperlinks/images-delete pattern).
- **No undo, no journal**. Mutating commands overwrite `FILE` in place. Pass `-o/--output PATH` to write to a parallel file instead, or `--dry-run` to preview. There is no snapshot ring, restore command, or trash directory — git is the version history. When both `--dry-run` and `--output` are passed, `--dry-run` wins (nothing is written to either path); the dry-run payload echoes `output` so the agent knows where a real run would have written.

## Commands

`docx <verb>` and `docx <noun> <verb>`. Every command has `--help`. Mutating commands accept `--dry-run` and `-o/--output PATH` (write to a parallel file instead of overwriting `FILE`). JSON output by default; structured `{ok: false, code, error, hint}` on failure.

| Surface         | Verbs                                                                            |
| --------------- | -------------------------------------------------------------------------------- |
| top-level       | `create` `read` `insert` `edit` `delete` `find` `replace` `wc` `outline`         |
| `comments`      | `add` `reply` `resolve` `delete` `list`                                          |
| `images`        | `list` `extract` `replace`                                                       |
| `hyperlinks`    | `add` `list` `replace` `delete`                                                  |
| `track-changes` | `FILE on\|off` (toggle), `list FILE`, `accept`/`reject FILE (--at tcN \| --all)` |
| `info`          | `schema` `locators`                                                              |

`insert --url URL --text "label"` wraps the inserted run in a `<w:hyperlink>`. To wrap an existing span, use `hyperlinks add --at pN:S-E --url URL`.

`insert --section [--columns N] [--type continuous|nextPage|evenPage|oddPage|nextColumn]` inserts a sentinel paragraph that carries an inline `<w:sectPr>` — defining the section ENDING at that paragraph. Combine with the existing `--page-break` / `--column-break` flags (also content-flag-mutex) to author multi-column layouts: e.g. set the doc-level trailing section to 2 columns with `edit --at s0 --columns 2 --type continuous` and use `--column-break` to control flow inside it. `wc FILE` (whole-document scope only) reports a `sections` count alongside `words`.

`docx read FILE` renders the body as GFM (the default since v0.9; pass `--ast` for the JSON-AST view). Each paragraph trails its locator as an HTML comment (`<!-- p3 -->`) — invisible in rendered view, parseable from raw text. Headings → `#`, lists → `-`, tables → pipe tables (with per-cell-paragraph locators, multi-paragraph cells joined by `<br>`), images → `![alt](imgN)`, run color / highlight → `<span style="color:#hex">…</span>` / `<span style="background-color:NAME">…</span>`. `--from`/`--to` slice top-level blocks (paragraph/table/cell/span/range locators all collapse to enclosing top-level block). Tracked changes have three views: default (`current`) renders additive wrappers (`<w:ins>`, `<w:moveTo>`) as `{++text++}[^tcN]` and subtractive wrappers (`<w:del>`, `<w:moveFrom>`) as `{--text--}[^tcN]` (CriticMarkup) with `[^tcN]: insertion|deletion|moveTo|moveFrom by author (date)` definitions at the end; `--accepted` drops subtractive wrappers and inlines additive as plain; `--baseline` drops additive and inlines subtractive as plain. `--accepted` and `--baseline` are mutually exclusive. `--comments` appends GFM footnote refs (`[^cN]`) at the end of each commented span and emits `[^cN]: "span" — author (date): body` definitions at end of output (replies marked `↳ cP`). Footnotes / endnotes (the document's own) render unconditionally as `[^fnN]` / `[^enN]` with definitions at the end. Equations (OOMath / `<m:oMath>` / `<m:oMathPara>`) surface as `` `equation: text` `` (concatenated `<m:t>` plaintext — degraded but readable; structure like sub/sup collapses to literal characters). Charts / SmartArt / shapes / other non-picture drawings (including legacy `<w:pict>` / `<w:object>` embeds) render as `` `[chart]` `` / `` `[smartart]` `` / `` `[shape]` `` / `` `[drawing]` `` placeholders. `<w:sym>` symbol-font references decode to Unicode via tables in [src/core/ast/sym.ts](src/core/ast/sym.ts); `<w:noBreakHyphen>` and `<w:softHyphen>` fold into adjacent TextRun text. All locator/render logic in `cli/read/markdown.ts`; footnote/endnote/equation/chart detection in `core/ast/read.ts`.

`docx wc FILE [LOCATOR]` accepts the same `--accepted` / `--baseline` flags. Default is the `current` view (counts every tracked-change wrapper's text — everything on disk); `--accepted` skips subtractive wrappers (`<w:del>`, `<w:moveFrom>`); `--baseline` skips additive wrappers (`<w:ins>`, `<w:moveTo>`). Helpers: `paragraphTextAccepted` / `paragraphTextBaseline` in [src/core/ast/text.ts](src/core/ast/text.ts).

Exit codes: `0` ok, `1` general, `2` usage, `3` not-found (file/locator/comment/image/hyperlink). Defined in `src/cli/respond.ts` (`EXIT` const + `ErrorCode` union).

## Locators

```
pN                   paragraph N
pN:S-E               chars S..E within paragraph N
pN:S-pM:E            cross-paragraph range
tN                   table N; tN:rRcC for cell at row R, col C; chainable :pK
sN                   section break N (inline sectPr lives in pN's pPr; trailing
                     sectPr is the body's last child)
cN, imgN, linkN, tcN comment / image / hyperlink / tracked-change ids
```

Span comments split runs at offsets, preserving `<w:rPr>` on both halves (logic in `cli/comments/helpers.tsx → addCommentMarkersToParagraph`). `hyperlinks add` does similar run-splitting in `cli/hyperlinks/wrap.tsx → wrapSpanInHyperlink`.

## Testing

```bash
bun run test:unit         # core + cli (~3s)
bun run test:integration  # LibreOffice round-trip (auto-skips if no soffice)
bun test                  # everything
bun run check             # biome + knip + tsc
```

LibreOffice install for local integration tests:

- macOS: `brew install --cask libreoffice`
- Linux: `sudo apt-get install libreoffice-core libreoffice-writer`

## Build & Publish

Two distinct artifacts:

- **`bun run build`** → `dist/index.js` — single bundled JS (~530 KB, shebang preserved). What npm publishes; runs under Bun. Triggered automatically by `prepack` before `npm pack`/`npm publish`.
- **`bun run build:binary`** → `dist/docx` — `bun build --compile` standalone executable. What GitHub Releases ship per-platform.

The bundled JS path matters because **path aliases (`@core/*`) and JSX runtime resolution don't work when the package is consumed from `node_modules`** — the consumer's tsconfig doesn't have our paths. The bundle pre-resolves everything; without it, `bun run docx ...` from a globally-installed package fails with `Cannot find module 'react/jsx-dev-runtime'`. So: never ship raw `src/`; always bundle.

## CI

GitHub Actions runs on push to `main` and on PRs (`.github/workflows/ci.yml`):

- `check` — biome + knip + tsc
- `unit-tests` — `bun run test:unit`
- `integration-tests` — installs LibreOffice, runs `bun run test:integration`
- `build-binary` — smoke-builds `dist/docx` via `bun run build:binary`

Tag pushes (`v*`) trigger:

- `.github/workflows/release.yml` — matrix-builds five platform binaries (linux/darwin × x64/arm64 + windows-x64) and publishes via `softprops/action-gh-release`. The `install.sh` at the repo root downloads the right binary for the user's platform.
- `.github/workflows/publish.yml` — runs `npm publish --access public` against the `Publishing` environment (uses npm trusted publishing via OIDC; no `NPM_TOKEN` secret needed).

## When Editing

- New CLI command? Add a folder under `src/cli/`, register in the COMMANDS map in `src/cli/index.ts`, add tests under `tests/cli/`, document in README and CLAUDE.md. If it emits OOXML, the file is `.tsx` and uses JSX components from `@core/jsx`.
- New OOXML tag? Add to the appropriate namespace's tag list in `src/core/jsx/index.ts`. The mapped-type pattern (`namespace("w", W_TAGS)`) survives `noUncheckedIndexedAccess` because the keys are a literal-string union.
- New AST field? Add to `src/core/ast/types.ts`, populate in `src/core/ast/read.ts`, then update `src/cli/info/schema.ts` JSON schema. The `--ts` output reads `types.ts` live via Bun's text import attribute, so it stays in sync automatically.
- New run-bearing wrapper (a paragraph-level element that holds `<w:r>` children that should be visible to `find` / `wc` / `read`)? Add to `RUN_BEARING_WRAPPER_TAGS` in `src/core/parser/run-ops.ts` and recurse into it in `walkRunContainer` in `src/core/ast/read.ts`. Every offset-aware walker reads from the predicate, so that's the only edit. Add a regression test in `tests/cli/preserve-unknown.test.ts` (if you also chose to leave a related tag _out_ of the set, to prove pass-through works) or `tests/cli/transparent-wrappers.test.ts` (round-trip find→replace with the wrapper present).
- New tracked-change kind? Extend `TrackedChange["kind"]` (and `TrackedChangeKind`) in `src/core/ast/types.ts`, widen the JSON schema enum in `src/cli/info/schema.ts`, update `walkRunContainer` in `src/core/ast/read.ts`, update `paragraphTextAccepted` / `paragraphTextBaseline` in `src/core/ast/text.ts` and `isRunVisible` / `criticMarkerFor` / `trackedChangeLabelFor` in `src/cli/read/markdown.ts`. If the new kind also takes part in `track-changes accept`/`reject`, update `actionFor` and `applyAccept`/`applyReject` in `src/cli/track-changes/apply.ts` and `trackedChangeKindForTag` in `src/cli/track-changes/list.ts`. **Reader and apply.ts walks must mirror each other** so `tcN` ids agree between `track-changes list` and `accept --at tcN`. Today's order: per paragraph, run-level wrappers first, then `<w:sectPrChange>`, then paragraph-mark `<w:ins>`/`<w:del>`. New tracked-change kinds need to slot into the same order in both walks.
- New section property (margins, page size, orientation, etc.)? Extend `SectionBreak` in `src/core/ast/types.ts` and `readSectionProperties` in `src/core/sections.tsx`; widen the JSON schema in `src/cli/info/schema.ts`; add an `applyXxx` mutator alongside `applyColumns` / `applySectionType` in `src/core/sections.tsx`; thread the new flag through `parseSectionFlags` in both `src/cli/insert/index.tsx` and `src/cli/edit/index.tsx`. Under track-changes, `wrapSectPrChange` already snapshots all sectPr children — new properties round-trip through the snapshot for free, and `track-changes list` will surface them in `prior` / `current`.
- New fixture? Drop in `tests/fixtures/`, add to the FIXTURES list in `tests/integration/libreoffice-roundtrip.test.ts` if it should round-trip. Prefer building it programmatically via a `scripts/make-*-fixture.ts` script (so the source is reproducible and stays MIT-compatible). The `scripts/inspect-fixtures.ts` summarizes what each one exercises.
- Don't touch `bun.lock` manually; let `bun install` manage it.
