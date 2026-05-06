# docx-cli

[![Watch: Claude filling out and redlining an NDA](https://cdn.loom.com/sessions/thumbnails/da70269a970f42caa138fb3389b4b9cc-f477402396d4154d-full-play.gif#t=0.1)](https://www.loom.com/share/da70269a970f42caa138fb3389b4b9cc)

**Let your AI agent review Word documents the way a human would.** Leave comments, suggest redlines, never break the formatting or remove content. You accept or reject in Word.

- Hand a `.docx` to Claude or Codex and get back a redlined copy with comments — open it in Word, accept or reject as usual.
- Agents see a structured AST with stable character offsets (`p3:5-20`); humans see normal Word formatting on disk.
- Custom styles, theme colors, embedded objects — all of it survives. We mutate XML in place rather than re-emitting from a lossy model.

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

## Quick example: filling out an NDA

The repo includes a Common Paper Mutual NDA template at `tests/fixtures/mnda.docx`. Below are the primitives an agent would compose to fill in the cover page and leave redline edits — the same flow shown in the video above. Every command was verified end-to-end against the fixture:

```sh
# Make a copy first — there's no undo
cp tests/fixtures/mnda.docx mnda-filled.docx

# Read the cover-page table so the agent knows what placeholders exist
docx read mnda-filled.docx --markdown --to t1

# Fill the yellow-highlighted bracketed placeholders
docx replace mnda-filled.docx "Fill in: today’s date" "May 6, 2026"
docx replace mnda-filled.docx "fill in state and/or county" "California"
docx replace mnda-filled.docx "fill in state" "California"
docx replace mnda-filled.docx "Fill in, if any." "None."

# Verify nothing's left to fill
docx find mnda-filled.docx '\[(Fill|fill)[^]]*\]' --regex --all

# Flip on tracked changes for the redline pass
docx track-changes mnda-filled.docx on

# Tighten "having a reasonable need to know" in the Use & Protection clause
docx replace mnda-filled.docx \
    "having a reasonable need to know" \
    "with a documented need to know"

# Leave a comment for the human reviewer
docx comments add mnda-filled.docx --range p7:0-30 \
    --text "Should we narrow 'representatives' to a named list?"
```

Open `mnda-filled.docx` in Word: tracked changes and comments appear in the review pane, ready to accept, reject, or reply. Or run `docx track-changes accept --all mnda-filled.docx` to bake them in from the CLI.

## Commands

```sh
docx create FILE [--title T] [--author A] [--text "..."]
docx read FILE [--markdown [--from pN] [--to pN] [--accepted | --baseline] [--comments]]
docx insert FILE --after p3 --text "..." [--style HeadingN] [--color HEX] [--bold] [--italic] [--url URL]
docx insert FILE --after p3 --runs '[{"type":"text","text":"X","bold":true}]'
docx edit   FILE --at p3 --text "..." | --runs '[...]'
docx delete FILE --at p3

docx find FILE QUERY [--regex] [--ignore-case] [--all] [--nth N]
docx replace FILE PATTERN REPLACEMENT [--regex] [--ignore-case] [--all] [--limit N] [--dry-run]

docx wc      FILE [LOCATOR] [--accepted | --baseline]
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
docx track-changes list   FILE
docx track-changes accept FILE (--at tcN | --all)
docx track-changes reject FILE (--at tcN | --all)
docx info schema [--ts]
docx info locators [--json]
```

Every command has `--help`. Mutating commands accept `--dry-run` and `-o/--output PATH` (write to a parallel file instead of overwriting `FILE`). JSON output by default for `read` and `*.list`; structured `{ok, code, error, hint}` on failure.

### Markdown rendering

`docx read FILE --markdown` renders the document body as GitHub-flavored Markdown instead of JSON. Useful when an LLM wants to skim a doc without parsing the AST. Each rendered paragraph is followed by an HTML comment with its locator (`<!-- p3 -->`) so the markdown is invisible-pinned: humans see clean prose in a renderer, agents parse the locators from raw text.

`--from LOC` and `--to LOC` slice by top-level block (both inclusive). Accepts paragraph, table, cell, span, and range locators; cell/span/range collapse to their enclosing top-level block.

**Tracked changes — three views.** By default, additive wrappers (`<w:ins>`, `<w:moveTo>`) render as CriticMarkup `{++text++}[^tcN]` and subtractive wrappers (`<w:del>`, `<w:moveFrom>`) as `{--text--}[^tcN]`, with `[^tcN]: …` definitions appended at end. Two flags switch view (mutually exclusive):

- `--accepted` — post-accept view: `<w:del>` and `<w:moveFrom>` runs are dropped; `<w:ins>` and `<w:moveTo>` runs render as plain text.
- `--baseline` — pre-change view: the inverse.

`docx wc` accepts the same `--accepted` / `--baseline` flags with parallel semantics.

`--comments` appends a GFM footnote reference (`[^cN]`) at the end of each commented span and emits one footnote definition per comment at the end of the output. Footnotes/endnotes (the document's own `<w:footnoteReference>` / `<w:endnoteReference>`) are rendered unconditionally as `[^fnN]` / `[^enN]`. Run `docx info schema` for the full mapping table.

### Locators

```
pN                   paragraph N (e.g., p3)
pN:S-E               characters S..E within paragraph N
pN:S-pM:E            cross-paragraph range
tN                   table N; tN:rRcC for cell at row R, col C
cN, imgN, linkN, tcN comment / image / hyperlink / tracked-change ids
```

Run `docx info locators` for the full reference.

## How It Works

**In-place XML mutation.** The AST returned by `read` is a _view_ over the parsed XML tree, not a separate model. When you `edit` or `comments add`, we mutate the underlying XML nodes directly and serialize back. Anything we don't model in the AST (custom styles, theme colors, schema extensions) survives because we never re-emit untouched regions.

**JSX for emitters.** Constructing OOXML fragments imperatively (`<w:rPr>` → `<w:b/>` → `<w:color w:val="800080"/>`) gets verbose. We write emitters in JSX with a custom factory: `<w.rPr><w.b/><w.color w-val="800080"/></w.rPr>` becomes the right `XmlNode` tree. Component names are PascalCase (`<Paragraph>`, `<RunProperties>`); they return `XmlNode | null` so empty wrappers get omitted automatically.

**Span-aware comments.** `comments add --range p3:5-20` finds the runs that contain offsets 5 and 20, splits them at the boundaries (preserving rPr formatting on both halves), and inserts `<w:commentRangeStart>` / `<w:commentRangeEnd>` markers between the slices.

**ParaId auto-injection.** Comments authored by tools like mammoth or older Word versions lack `w14:paraId`, which `commentsExtended.xml` requires for resolve/reply. We detect this on resolve/reply and inject a fresh paraId, also adding the `xmlns:w14` namespace declaration to the `<w:comments>` root if missing.

**Cross-format image replacement.** `images replace --at img0 --with new.png` detects the new MIME type via `Bun.file().type`, renames the part (`word/media/image1.jpeg` → `word/media/image1.png`), rewrites the relationship `Target`, and ensures `[Content_Types].xml` has a `<Default>` for the new extension.

**Hyperlink CRUD.** `hyperlinks list` enumerates `<w:hyperlink>` elements with positional `linkN` ids; `hyperlinks add --at p3:5-20 --url URL` wraps an existing span (splitting runs at offsets); `hyperlinks replace --at link0 --with URL` updates the rels `Target`, allocating a new rId if the existing one is shared so siblings stay pointed at the original URL; `hyperlinks delete --at link0` unwraps the link (text survives) and prunes the rels entry when no longer referenced.

## Stack

- **Runtime**: Bun (`node:util` parseArgs, JSX with custom factory, native zlib)
- **Parser**: [`jszip`](https://www.npmjs.com/package/jszip) + [`fast-xml-parser`](https://www.npmjs.com/package/fast-xml-parser) + [`fast-xml-builder`](https://www.npmjs.com/package/fast-xml-builder)
- **Quality**: Biome + Knip + tsc; LibreOffice headless for round-trip integration tests
- **Standard**: ECMA-376 Part 1 §17 (WordprocessingML), Transitional profile

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, architecture overview, and CI.
