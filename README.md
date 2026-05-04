# docx-cli

A Bun-built CLI for AI agents (Claude, Codex) to read, edit, and comment on `.docx` files with full format fidelity. Outputs JSON-AST for precise locator-based editing; preserves anything it doesn't model by mutating XML in place.

## Stack

- **Runtime**: Bun (`node:util` parseArgs, JSX with custom factory, native zlib)
- **Parser**: [`jszip`](https://www.npmjs.com/package/jszip) + [`fast-xml-parser`](https://www.npmjs.com/package/fast-xml-parser) + [`fast-xml-builder`](https://www.npmjs.com/package/fast-xml-builder)
- **Quality**: Biome + Knip + tsc; LibreOffice headless for round-trip integration tests
- **Standard**: ECMA-376 Part 1 §17 (WordprocessingML), Transitional profile

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
bun add -g @kklimuk/docx-cli
# or
bunx @kklimuk/docx-cli read doc.docx
```

## Commands

```sh
docx create FILE [--title T] [--author A] [--text "..."]
docx read FILE
docx insert FILE --after p3 --text "..." [--style HeadingN] [--color HEX] [--bold] [--italic]
docx insert FILE --after p3 --runs '[{"type":"text","text":"X","bold":true}]'
docx edit   FILE --at p3 --text "..." | --runs '[...]'
docx delete FILE --at p3

docx comments add     FILE --range p3:5-20 --text "..." [--author NAME]
docx comments reply   FILE --to c0 --text "..."
docx comments resolve FILE --id c0 [--unset]
docx comments delete  FILE --id c0
docx comments restore FILE --id c0
docx comments list    FILE [--include-resolved] [--thread c0]

docx images list    FILE
docx images extract FILE --to ./media [--id imgN]
docx images replace FILE --at imgN --with ./new.png

docx track-changes FILE on|off
docx schema [--ts]
docx locators [--json]
```

Every command has `--help`. All mutating commands accept `--dry-run`. JSON output by default for `read` and `*.list`; structured `{ok, code, error, hint}` on failure.

### Locators

```
pN              paragraph N (e.g., p3)
pN:S-E          characters S..E within paragraph N
pN:S-pM:E       cross-paragraph range
tN              table N; tN:rRcC for cell at row R, col C
cN, imgN        comment / image ids
```

Run `docx locators` for the full reference.

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
    read/                # read FILE
    insert/              # insert FILE  (uses ./emit Paragraph component)
    edit/                # edit FILE
    delete/              # delete FILE
    comments/            # add | reply | resolve | delete | restore | list
    images/              # list | extract | replace
    track-changes/       # FILE on|off
    schema/              # AST JSON Schema + TS source
    locators-cmd/        # grammar reference
  core/
    package/             # JSZip open/close, named-part read/write
    parser/              # XmlNode class + parse/serialize + JSX factory
    jsx/                 # h, Fragment, namespaces (w, r, a, wp, pic, ...)
    ast/                 # types + DocView + XML→AST reader
    locators/            # parse "p3:5-20" + resolve to refs
tests/
  core/, cli/, integration/    # 51 tests, 9 files
  fixtures/                    # 11 .docx fixtures across producers
```

## How It Works

**In-place XML mutation.** The AST returned by `read` is a *view* over the parsed XML tree, not a separate model. When you `edit` or `comments add`, we mutate the underlying XML nodes directly and serialize back. Anything we don't model in the AST (custom styles, theme colors, schema extensions) survives because we never re-emit untouched regions.

**JSX for emitters.** Constructing OOXML fragments imperatively (`<w:rPr>` → `<w:b/>` → `<w:color w:val="800080"/>`) gets verbose. We write emitters in JSX with a custom factory: `<w.rPr><w.b/><w.color w-val="800080"/></w.rPr>` becomes the right `XmlNode` tree. Component names are PascalCase (`<Paragraph>`, `<RunProperties>`); they return `XmlNode | null` so empty wrappers get omitted automatically.

**Span-aware comments.** `comments add --range p3:5-20` finds the runs that contain offsets 5 and 20, splits them at the boundaries (preserving rPr formatting on both halves), and inserts `<w:commentRangeStart>` / `<w:commentRangeEnd>` markers between the slices. The `<w:commentReference>` run goes after the end marker.

**ParaId auto-injection.** Comments authored by tools like mammoth or older Word versions lack `w14:paraId`, which `commentsExtended.xml` requires for resolve/reply. We detect this on resolve/reply and inject a fresh paraId, also adding the `xmlns:w14` namespace declaration to the `<w:comments>` root if missing.

**Cross-format image replacement.** `images replace --at img0 --with new.png` detects the new MIME type via `Bun.file().type`, renames the part (`word/media/image1.jpeg` → `word/media/image1.png`), rewrites the relationship `Target`, and ensures `[Content_Types].xml` has a `<Default>` for the new extension.

**Trash for restore.** `comments delete` journals the removed comment XML + anchor info to `<dir>/.docx-cli/trash.json` so `comments restore --id cN` can re-anchor it at its original location.

## CI

GitHub Actions (`.github/workflows/ci.yml`) runs four jobs on push to `main` and on PRs:

| Job | What |
|---|---|
| `check` | `biome check . && knip-bun && tsc --noEmit` |
| `unit-tests` | `bun run test:unit` (core + cli, fast) |
| `integration-tests` | Installs LibreOffice, runs `bun run test:integration` |
| `build-binary` | Smoke-builds via `bun build --compile` and runs `--version` |

`.github/workflows/release.yml` triggers on `v*` tags, matrix-builds the five binaries, and uploads them to a GitHub Release via [`softprops/action-gh-release`](https://github.com/softprops/action-gh-release).
