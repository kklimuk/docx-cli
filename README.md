# docx-cli

A Bun-built CLI for AI agents (Claude, Codex) to read, edit, and comment on `.docx` files with full format fidelity. Outputs JSON-AST for precise locator-based editing; preserves anything it doesn't model by mutating XML in place.

## Install

**Standalone binary** (no Bun required):

```sh
curl -fsSL https://raw.githubusercontent.com/kklimuk/docx-cli/main/install.sh | sh
```

Honors `PREFIX` (default `$HOME/.local/bin`) and `VERSION` (default `latest`):

```sh
PREFIX=/usr/local sh -c "$(curl -fsSL https://raw.githubusercontent.com/kklimuk/docx-cli/main/install.sh)"
VERSION=v0.2.0 sh -c "$(curl -fsSL https://raw.githubusercontent.com/kklimuk/docx-cli/main/install.sh)"
```

Pre-built binaries are published for linux/x64, linux/arm64, darwin/x64, darwin/arm64, windows/x64.

**npm** (requires Bun >= 1.3):

```sh
bun add -g bun-docx
# or
bunx bun-docx read doc.docx
```

## Commands

```sh
docx create FILE [--title T] [--author A] [--text "..."]
docx read FILE [--markdown [--from pN] [--to pN] [--changes] [--comments]]
docx insert FILE --after p3 --text "..." [--style HeadingN] [--color HEX] [--bold] [--italic] [--url URL]
docx insert FILE --after p3 --runs '[{"type":"text","text":"X","bold":true}]'
docx edit   FILE --at p3 --text "..." | --runs '[...]'
docx delete FILE --at p3

docx find FILE QUERY [--regex] [--ignore-case] [--all] [--nth N]
docx replace FILE PATTERN REPLACEMENT [--regex] [--ignore-case] [--all] [--limit N] [--dry-run]

docx wc      FILE [LOCATOR]
docx outline FILE

docx comments add     FILE --range p3:5-20 --text "..." [--author NAME]
docx comments reply   FILE --to c0 --text "..."
docx comments resolve FILE --id c0 [--unset]
docx comments delete  FILE --id c0
docx comments list    FILE [--include-resolved] [--thread c0]

docx images list    FILE
docx images extract FILE --to ./media [--id imgN]
docx images replace FILE --at imgN --with ./new.png

docx hyperlinks list    FILE
docx hyperlinks add     FILE --at pN:S-E --url URL
docx hyperlinks replace FILE --at linkN --with URL
docx hyperlinks delete  FILE --at linkN

docx track-changes FILE on|off
docx info schema [--ts]
docx info locators [--json]
```

Every command has `--help`. Mutating commands accept `--dry-run` and `-o/--output PATH` (write to a parallel file instead of overwriting `FILE`). JSON output by default for `read` and `*.list`; structured `{ok, code, error, hint}` on failure.

When `<w:trackChanges/>` is set in the doc (toggle via `docx track-changes FILE on`), `insert`/`edit`/`delete`/`replace` automatically emit `<w:ins>`/`<w:del>` markers. Author resolution: per-call `--author NAME` overrides `$DOCX_AUTHOR`, which falls back to `docx-cli`. To make a one-off untracked edit, flip the flag off, edit, then flip it back on. `find` results inside tracked-change wrappers carry a `trackedChanges` array so agents can decide what to do with hits in pending insertions/deletions.

OOXML has no native tracked-change form for hyperlink edits or image swaps, so when track-changes is on, `hyperlinks add/replace/delete` and `images replace` auto-emit a `[docx-cli] …` comment anchored to the affected span/run instead. The comment carries the same `--author` attribution as the other tracked operations. Word itself silently bypasses tracking for these — we trade silence for an explicit audit trail.

### Markdown rendering

`docx read FILE --markdown` renders the document body as GitHub-flavored Markdown instead of JSON. Useful when you (or an LLM) want to skim a doc quickly without parsing the AST. Each rendered paragraph is followed by an HTML comment with its locator (`<!-- p3 -->`) so the markdown is invisible-pinned: humans see clean prose in a renderer, agents parse the locators from raw text.

- Headings → `#`/`##`/`###` based on `style="HeadingN"`
- Lists → `- ` indented per `level`
- Bold/italic/strike → `**…**` / `*…*` / `~~…~~`; underline → `<u>…</u>`
- Run color → `<span style="color:#hex">…</span>`; highlight → `<span style="background-color:NAME">…</span>`
- Hyperlinks → `[text](url)`
- Images → `![alt](imgN)`
- Tables → GitHub pipe tables; multi-paragraph cells joined with `<br>`; per-cell-paragraph locators inline
- Section breaks → `---`
- Equations (`<m:oMath>`/`<m:oMathPara>`) → `` `equation: text` `` (concatenated `<m:t>` plaintext; structure like sub/sup/fractions collapses to literal characters — degraded but readable)
- Footnotes / endnotes → inline `[^fnN]` / `[^enN]` refs with GFM footnote definitions at end of output
- Charts / SmartArt / shapes / other non-picture drawings → `` `[chart]` `` / `` `[smartart]` `` / `` `[shape]` `` / `` `[drawing]` `` placeholders

`--from LOC` and `--to LOC` slice by top-level block (both inclusive). Accepts paragraph, table, cell, span, and range locators; cell/span/range collapse to their enclosing top-level block. Comment/image/hyperlink locators are rejected.

`--changes` renders tracked insertions and deletions as `<ins>`/`<del>` instead of producing the accepted view (which silently drops deletions and inlines insertions).

`--comments` appends a GFM footnote reference (`[^cN]`) at the end of each commented span and emits one footnote definition per comment at the end of the output:

```
… some commented text[^c0] …

[^c0]: "commented span" — Author Name (2024-01-15T...): comment body
[^c1]: "another span" — Author Name (2024-01-15T...) ↳ c0: reply body
```

Footnotes/endnotes (the document's own `<w:footnoteReference>` / `<w:endnoteReference>`) are rendered unconditionally — `[^fn1]` / `[^en1]` inline + `[^fn1]: body` definitions at the end of the output, alongside any `--comments` footnotes. They use `fn`/`en` prefixes so the namespaces don't collide.

### Locators

```
pN              paragraph N (e.g., p3)
pN:S-E          characters S..E within paragraph N
pN:S-pM:E       cross-paragraph range
tN              table N; tN:rRcC for cell at row R, col C
cN, imgN, linkN comment / image / hyperlink ids
```

Run `docx info locators` for the full reference.

## Development

```sh
bun install && bun run prepare      # set up + git hooks
bun dev <subcommand>                # run via source
bun run check                       # biome + knip + tsc
bun run test:unit                   # core + cli tests (fast)
bun run test:integration            # LibreOffice round-trip (needs `soffice` on PATH)
bun test                            # everything
bun run build                       # produce dist/docx via bun build --compile
```

### LibreOffice (for integration tests)

- **macOS**: `brew install --cask libreoffice`
- **Linux**: `sudo apt-get install libreoffice-core libreoffice-writer`
- **Windows**: <https://www.libreoffice.org/download/>

## Architecture

```
src/
  index.ts               # binary entrypoint
  cli/
    index.ts             # parseArgs dispatch
    help.ts              # top-level --help
    respond.ts           # JSON ack / structured error helpers
    create/              # create FILE
    read/                # read FILE [--markdown ...]  (markdown.ts renderer)
    insert/              # insert FILE  (uses ./emit Paragraph component)
    edit/                # edit FILE
    delete/              # delete FILE
    find/                # find FILE QUERY
    replace/             # replace FILE PATTERN REPLACEMENT
    wc/                  # wc FILE [LOCATOR]
    outline/             # outline FILE
    comments/            # add | reply | resolve | delete | list
    images/              # list | extract | replace
    hyperlinks/          # add | list | replace | delete
    track-changes/       # FILE on|off
    info/                # schema | locators (reference output)
  core/
    package/             # JSZip open/close, named-part read/write
    parser/              # XmlNode class + parse/serialize + JSX factory
    jsx/                 # h, Fragment, namespaces (w, r, a, wp, pic, ...)
    ast/                 # types + DocView + XML→AST reader (text.ts: shared paragraph helpers)
    locators/            # parse "p3:5-20" + resolve to refs
tests/
  core/, cli/, integration/
  fixtures/
```

## Stack

- **Runtime**: Bun (`node:util` parseArgs, JSX with custom factory, native zlib)
- **Parser**: [`jszip`](https://www.npmjs.com/package/jszip) + [`fast-xml-parser`](https://www.npmjs.com/package/fast-xml-parser) + [`fast-xml-builder`](https://www.npmjs.com/package/fast-xml-builder)
- **Quality**: Biome + Knip + tsc; LibreOffice headless for round-trip integration tests
- **Standard**: ECMA-376 Part 1 §17 (WordprocessingML), Transitional profile

## How It Works

**In-place XML mutation.** The AST returned by `read` is a _view_ over the parsed XML tree, not a separate model. When you `edit` or `comments add`, we mutate the underlying XML nodes directly and serialize back. Anything we don't model in the AST (custom styles, theme colors, schema extensions) survives because we never re-emit untouched regions.

**JSX for emitters.** Constructing OOXML fragments imperatively (`<w:rPr>` → `<w:b/>` → `<w:color w:val="800080"/>`) gets verbose. We write emitters in JSX with a custom factory: `<w.rPr><w.b/><w.color w-val="800080"/></w.rPr>` becomes the right `XmlNode` tree. Component names are PascalCase (`<Paragraph>`, `<RunProperties>`); they return `XmlNode | null` so empty wrappers get omitted automatically.

**Span-aware comments.** `comments add --range p3:5-20` finds the runs that contain offsets 5 and 20, splits them at the boundaries (preserving rPr formatting on both halves), and inserts `<w:commentRangeStart>` / `<w:commentRangeEnd>` markers between the slices. The `<w:commentReference>` run goes after the end marker.

**ParaId auto-injection.** Comments authored by tools like mammoth or older Word versions lack `w14:paraId`, which `commentsExtended.xml` requires for resolve/reply. We detect this on resolve/reply and inject a fresh paraId, also adding the `xmlns:w14` namespace declaration to the `<w:comments>` root if missing.

**Cross-format image replacement.** `images replace --at img0 --with new.png` detects the new MIME type via `Bun.file().type`, renames the part (`word/media/image1.jpeg` → `word/media/image1.png`), rewrites the relationship `Target`, and ensures `[Content_Types].xml` has a `<Default>` for the new extension.

**Hyperlink CRUD.** `hyperlinks list` enumerates `<w:hyperlink>` elements with positional `linkN` ids; `hyperlinks add --at p3:5-20 --url URL` wraps an existing span (splitting runs at offsets); `hyperlinks replace --at link0 --with URL` updates the rels `Target`, allocating a new rId if the existing one is shared so siblings stay pointed at the original URL; `hyperlinks delete --at link0` unwraps the link (text survives) and prunes the rels entry when no longer referenced.

## CI

GitHub Actions (`.github/workflows/ci.yml`) runs four jobs on push to `main` and on PRs:

| Job                 | What                                                        |
| ------------------- | ----------------------------------------------------------- |
| `check`             | `biome check . && knip-bun && tsc --noEmit`                 |
| `unit-tests`        | `bun run test:unit` (core + cli, fast)                      |
| `integration-tests` | Installs LibreOffice, runs `bun run test:integration`       |
| `build-binary`      | Smoke-builds via `bun build --compile` and runs `--version` |

`.github/workflows/release.yml` triggers on `v*` tags, matrix-builds the five binaries, and uploads them to a GitHub Release via [`softprops/action-gh-release`](https://github.com/softprops/action-gh-release).
