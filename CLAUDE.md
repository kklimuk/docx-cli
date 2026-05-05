# docx-cli

CLI for AI agents (Claude, Codex) to read, edit, and comment on `.docx` files. JSON-AST output, locator-based addressing, full format fidelity via in-place XML mutation.

Default to **Bun**, not Node. Use `Bun.file`, `Bun.write`, `Bun.env`, `Bun.$` over Node equivalents. Bun loads `.env` automatically — don't use dotenv.

## Project Structure

```
src/
  index.ts                    # binary entrypoint (#!/usr/bin/env bun)
  cli/
    index.ts                  # parseArgs dispatcher, COMMANDS map
    help.ts                   # top-level --help
    respond.ts                # respond() / fail() — JSON ack + error helpers
    create/                   # docx create FILE
    read/                     # docx read FILE [--markdown ...]
      markdown.ts             # GFM renderer used by `read --markdown`
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
      list.ts                 # JSON inventory of every <w:ins>/<w:del>
      accept.ts | reject.ts   # thin wrappers over apply.ts
      apply.ts                # collectTrackedChanges + unwrap/delete/renameDelTextToText helpers
    info/                     # docx info <topic>
      index.ts                # sub-dispatcher for schema/locators
      schema.ts               # docx info schema [--ts]  (TS source via Bun text import)
      locators.ts             # docx info locators [--json]
  core/
    package/                  # JSZip wrapper: open, read/write XML parts, save
    parser/
      xml-node.ts             # XmlNode class — instance methods + static parse/serialize
      index.ts                # re-export
    jsx/
      index.ts                # h, Fragment, makeTag, namespaces (w, r, a, wp, pic, cp, dc, …)
      jsx-runtime.ts          # auto-runtime — converts component-null to empty fragments
      jsx-dev-runtime.ts      # alias for dev-mode auto-runtime
    ast/
      types.ts                # Doc / Block / Run / Comment / Footnote types (read live by `docx info schema --ts`)
      doc-view.ts             # DocView, openDocView, saveDocView, enrichImageHashes — also loads footnotes.xml/endnotes.xml
      read.ts                 # XML → typed AST walker; detects oMath/oMathPara, footnote/endnote refs, charts/SmartArt/shapes
      text.ts                 # paragraphText / flattenParagraphs / findBlockById — shared text helpers
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
  fixtures/                   # .docx files exercising different OOXML features (comments, tracked changes, footnotes, equations, tables, ...). `notes.docx` is from the Pandoc test corpus.
scripts/
  move.ts                     # move file + auto-update imports
  fxp-smoke.ts                # JSX → XML smoke test
  inspect-fixtures.ts         # summarize each fixture's features
  make-fixture.ts             # rebuild minimal.docx
  escape-check.ts             # XML entity round-trip sanity
```

## Key Conventions

These are not suggestions. Follow them.

- **Bun not Node**. `Bun.file(path).type` for MIME. All stdout goes through `respond()` (JSON ack) or `writeStdout()` (text) from `src/cli/respond.ts`, both of which use `Bun.write(Bun.stdout, ...)` — never `process.stdout.write` directly. The 64 KB truncation that bites on early exit is real and silent; the helpers are the only safe path.
- **File naming**: kebab-case, named after the primary export (`xml-node.ts` → `XmlNode`, `paragraph.ts` → `Paragraph`).
- **Feature structure**: each command is a kebab-case folder under `src/cli/`. `index.ts(x)` is the public surface (`export async function run(args: string[]): Promise<number>`). Sub-files for shared helpers, no cross-feature imports unless centrally exposed via `@core/*`.
- **Path aliases**: `@core` → `src/core/index.ts`, `@core/*` → `src/core/*/index.{ts,tsx}` or `src/core/*.{ts,tsx}` (configured in tsconfig). Use these in `src/cli/*` to avoid `../../core` chains. `src/core` itself uses relative imports between siblings.
- **JSX is for emitters only**. Any file that constructs fresh XML can be `.tsx` — wherever it lives in the tree. Readers, locators, and pure analysis stay `.ts`; never JSX in the AST reader. Components are PascalCase (`<Paragraph>`, `<RunProperties>`); they accept props and may return `XmlNode | null` (null skipped by flatten). Tag namespaces are imported as lowercase (`w.p`, `cp.coreProperties`); attribute names with colons go through hyphen shortcut (`w-val="800080"` → `w:val="800080"`) or JSX spread (`{...{"w:val": "x"}}`).
- **Variable names**: no single/two-letter names. Use descriptive (`paragraph` not `p`, `commentId` not `cid`). Exception: standard regex-match destructuring (`const [, prefix, idx] = match`) since the convention is unambiguous.
- **Spacing**: tabs, double quotes (Biome enforced).
- **Top-down file ordering**: exports first, dependencies below — like a newspaper. Use `function` declarations (hoisted) for internal helpers so this works at runtime. Arrow functions only for inline callbacks and short utilities.
- **Early returns**: bail on validation failures and null checks. No else-if chains when an early return reads cleaner.

## Architectural Invariants

These are not architectural suggestions, but requirements. If you disagree, make the case about needing a different pattern for a feature.

- **In-place XML mutation, not AST round-trip emission**. The typed AST returned by `read` is a _view_ over the parsed XML tree. Mutations (insert/edit/delete/comments add) operate on the underlying `XmlNode` references via `BlockReference.parent.splice(...)`. Anything we don't model survives because we never re-emit untouched regions. Only emit fresh XML for nodes we're inserting (via JSX) — never round-trip whole subtrees through the AST.
- **fast-xml-builder owns escaping**. On the JSX path, never manually escape — the builder handles entities correctly (uses `&apos;` for `'`, which fast-xml-parser decodes back). The static templates in `cli/create/template.tsx` carry no user-supplied content, so they don't need escaping at all.
- **JSX.Element = XmlNode** (single, not nullable union). `Fragment` returns a sentinel `#fragment` XmlNode that gets unwrapped both in `flatten()` (composition) and `XmlNode.serialize()` (top-level). Components that want to "render nothing" return `null`; `jsx()` converts that to an empty fragment so callers always see `XmlNode`.
- **Stable positional ids** (`p0`, `t0`, `c0`, `img0`, `link0`, `tc0`). Block ids shift after structural edits — agents must re-read between non-trivial mutations. Comment numeric ids are allocated as `max-existing + 1`. Image, hyperlink, and tracked-change ids are positional (document order).
- **Hyperlinks own a relationship, not their text.** `hyperlinks replace` updates the `Target` of the underlying `<Relationship>` — but if multiple `<w:hyperlink>` elements share the same `r:id`, it allocates a new rId so the others stay pointing at the original URL. `hyperlinks delete` unwraps the `<w:hyperlink>` (text survives as plain runs) and prunes the rId from the rels file when no other element references it.
- **paraId is required for resolve/reply**. Comments authored by external tools may lack `w14:paraId` on their bodies. The `resolve` and `reply` verbs auto-inject one via `ensureCommentParaId()` (also adds `xmlns:w14` to the `<w:comments>` root if missing). Do this rather than failing — agents shouldn't have to recreate comments.
- **Track-changes is doc-level, not per-command**. When `<w:trackChanges/>` is set in `settings.xml`, every mutating command (`insert`/`edit`/`delete`/`replace`) automatically emits `<w:ins>`/`<w:del>` markers — there is no per-command override flag. To make a one-off untracked edit, run `docx track-changes FILE off`, edit, then `track-changes on`. **Accept/reject bypass tracking** (`docx track-changes accept|reject FILE --at tcN | --all`): they directly mutate the XML tree without wrapping the change itself in `<w:ins>`/`<w:del>` — accept-ins unwraps, accept-del deletes, reject-ins deletes, reject-del unwraps after renaming `<w:delText>` → `<w:t>`. Targets walk fresh on each invocation (stored `view.trackedChangeReferences` would go stale across mutations) and process in reverse pre-order so nested changes are applied before their parents. Out of scope (silently skipped by `--all`): tracked paragraph marks, `<w:rPrChange>`, `<w:pPrChange>`, `<w:moveFrom>`/`<w:moveTo>`. Author/date come from the per-call `--author NAME` flag, then the `DOCX_AUTHOR` env var, then the `"docx-cli"` default (resolved in `core/track-changes/index.ts → resolveAuthor`); `DOCX_CLI_NOW` injects a fixed date for tests. `delete tN` rejects with `TRACKED_CHANGE_CONFLICT` when tracking is on (tracked table-row deletion isn't supported).
- **Hyperlink and image edits emit audit comments under track-changes**. OOXML has no `w:hyperlinkChange` / `w:drawingChange` element — Word itself silently bypasses tracking for hyperlink edits and image swaps. We compromise: when `<w:trackChanges/>` is on, `hyperlinks add/replace/delete` and `images replace` each auto-emit a `[docx-cli] …` comment anchored to the affected span/run, attributed via the same `--author` chain. The mutation itself stays silent (no fake `<w:ins>`/`<w:del>` since OOXML has no honest construct for it). Helpers live in `cli/comments/helpers.tsx` (`emitAuditComment`, `findContainingParagraph`, `findElementOffsetsInParagraph`, `addCommentMarkersAroundRun`). When track-changes is off, no comment is emitted.
- **No undo, no journal**. Mutating commands overwrite `FILE` in place. Pass `-o/--output PATH` to write to a parallel file instead, or `--dry-run` to preview. There is no snapshot ring, restore command, or trash directory — git is the version history. When both `--dry-run` and `--output` are passed, `--dry-run` wins (nothing is written to either path); the dry-run payload echoes `output` so the agent knows where a real run would have written.

## Commands

`docx <verb>` and `docx <noun> <verb>`. Every command has `--help`. Mutating commands accept `--dry-run` and `-o/--output PATH` (write to a parallel file instead of overwriting `FILE`). JSON output by default; structured `{ok: false, code, error, hint}` on failure.

| Surface         | Verbs                                                                                |
| --------------- | ------------------------------------------------------------------------------------ |
| top-level       | `create` `read` `insert` `edit` `delete` `find` `replace` `wc` `outline`             |
| `comments`      | `add` `reply` `resolve` `delete` `list`                                              |
| `images`        | `list` `extract` `replace`                                                           |
| `hyperlinks`    | `add` `list` `replace` `delete`                                                      |
| `track-changes` | `FILE on\|off` (toggle), `list FILE`, `accept`/`reject FILE (--at tcN \| --all)`    |
| `info`          | `schema` `locators`                                                                  |

`insert --url URL --text "label"` wraps the inserted run in a `<w:hyperlink>`. To wrap an existing span, use `hyperlinks add --at pN:S-E --url URL`.

`docx read FILE --markdown` renders the body as GFM (instead of JSON). Each paragraph trails its locator as an HTML comment (`<!-- p3 -->`) — invisible in rendered view, parseable from raw text. Headings → `#`, lists → `-`, tables → pipe tables (with per-cell-paragraph locators, multi-paragraph cells joined by `<br>`), images → `![alt](imgN)`, run color / highlight → `<span style="color:#hex">…</span>` / `<span style="background-color:NAME">…</span>`. `--from`/`--to` slice top-level blocks (paragraph/table/cell/span/range locators all collapse to enclosing top-level block). Tracked changes have three views: default (`current`) renders insertions as `{++text++}[^tcN]` and deletions as `{--text--}[^tcN]` (CriticMarkup) with `[^tcN]: insertion|deletion by author (date)` definitions at the end; `--accepted` shows the post-accept view (drops `del` runs, inlines `ins` as plain); `--baseline` shows the pre-change view (drops `ins` runs, inlines `del` as plain). `--accepted` and `--baseline` are mutually exclusive. `--comments` appends GFM footnote refs (`[^cN]`) at the end of each commented span and emits `[^cN]: "span" — author (date): body` definitions at end of output (replies marked `↳ cP`). Footnotes / endnotes (the document's own) render unconditionally as `[^fnN]` / `[^enN]` with definitions at the end. Equations (OOMath / `<m:oMath>` / `<m:oMathPara>`) surface as `` `equation: text` `` (concatenated `<m:t>` plaintext — degraded but readable; structure like sub/sup collapses to literal characters). Charts / SmartArt / shapes / other non-picture drawings render as `` `[chart]` `` / `` `[smartart]` `` / `` `[shape]` `` / `` `[drawing]` `` placeholders. All locator/render logic in `cli/read/markdown.ts`; footnote/endnote/equation/chart detection in `core/ast/read.ts`.

`docx wc FILE [LOCATOR]` accepts the same `--accepted` / `--baseline` flags. Default is the `current` view (counts plain + ins + del, i.e., everything on disk); `--accepted` skips `<w:del>` runs; `--baseline` skips `<w:ins>` runs. Helpers: `paragraphTextAccepted` / `paragraphTextBaseline` in [src/core/ast/text.ts](src/core/ast/text.ts).

Exit codes: `0` ok, `1` general, `2` usage, `3` not-found (file/locator/comment/image/hyperlink). Defined in `src/cli/respond.ts` (`EXIT` const + `ErrorCode` union).

## Locators

```
pN                   paragraph N
pN:S-E               chars S..E within paragraph N
pN:S-pM:E            cross-paragraph range
tN                   table N; tN:rRcC for cell at row R, col C; chainable :pK
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
- New fixture? Drop in `tests/fixtures/`, add to the FIXTURES list in `tests/integration/libreoffice-roundtrip.test.ts` if it should round-trip. The `scripts/inspect-fixtures.ts` summarizes what each one exercises.
- Don't touch `bun.lock` manually; let `bun install` manage it.
