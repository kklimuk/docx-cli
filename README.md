# docx-cli

[![Watch: Claude filling out and redlining an NDA](https://cdn.loom.com/sessions/thumbnails/da70269a970f42caa138fb3389b4b9cc-f477402396d4154d-full-play.gif#t=0.1)](https://www.loom.com/share/da70269a970f42caa138fb3389b4b9cc)

**A `.docx` CLI built for AI agents.** Leave comments, suggest redlines, and edit Word documents without breaking the formatting or losing content — a human accepts or rejects in Word afterward.

- Hand a `.docx` to Claude or Codex and get back a redlined copy with comments — open it in Word, accept or reject as usual.
- Agents address text by **stable locators** with character offsets (`p3:5-20`); humans see normal Word formatting on disk.
- Custom styles, theme colors, embedded objects — all of it survives. The CLI mutates XML in place rather than re-emitting from a lossy model.

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
# Make a copy first — there's no undo (git is the history; the CLI overwrites in place)
cp tests/fixtures/mnda.docx mnda-filled.docx

# Read the cover-page table so the agent knows what placeholders exist
docx read mnda-filled.docx --from t1 --to t1

# Fill the yellow-highlighted bracketed placeholders
docx replace mnda-filled.docx "Fill in: today's date" "May 6, 2026"
docx replace mnda-filled.docx "fill in state and/or county" "California"
docx replace mnda-filled.docx "fill in state" "California"
docx replace mnda-filled.docx "Fill in, if any." "None."

# Verify nothing's left to fill (bare locator lines, one per match; nothing → exit 0)
docx find mnda-filled.docx '\[(Fill|fill)[^]]*\]' --regex --all

# Flip on tracked changes for the redline pass
docx track-changes mnda-filled.docx on

# Tighten "having a reasonable need to know" in the Use & Protection clause
docx replace mnda-filled.docx \
    "having a reasonable need to know" \
    "with a documented need to know"

# Leave a comment for the human reviewer — addresses an existing span with --at
docx comments add mnda-filled.docx --at p7:0-30 \
    --text "Should we narrow 'representatives' to a named list?"
```

Open `mnda-filled.docx` in Word: tracked changes and comments appear in the review pane, ready to accept, reject, or reply. Or run `docx track-changes accept mnda-filled.docx --all` to bake them in from the CLI.

## `docx <command> --help` is the authoritative contract

> **Agents: run `docx <command> --help` before composing a call.** Every command's `--help` is the source of truth for its flags, locator forms, and exact output shape — this README is a map, not the territory. Two more must-reads:
>
> - **`docx info locators`** — the canonical locator grammar (`--json` for a machine-readable form). The top-level `docx --help` says it outright: *"It is highly recommended to agents to run `docx info locators` to understand their capabilities."*
> - **`docx info schema`** — the AST type definitions (`--ts` for TypeScript source) that `read --ast` emits.

## Command reference

`docx <verb>` and `docx <noun> <verb>`. Every command has `--help`. Two groups: **read/query** commands print data to stdout; **mutate** commands change the file (and accept `--dry-run`, `-o/--output PATH`, `-v/--verbose`).

### Read & query (print to stdout, never write the file)

```sh
docx read    FILE [--from LOC] [--to LOC] [--accepted | --baseline | --current] [--comments]
docx read    FILE --ast                  # JSON-AST instead of Markdown (disables the markdown-only flags)
docx find    FILE QUERY [--regex] [--ignore-case] [--all] [--nth N] [--current | --baseline] [--exact] [--json]
docx find    FILE (--highlight COLOR|any | --color HEX | --bold | --italic | --underline) [--all] [--json]   # find by formatting (no QUERY)
docx wc      FILE [LOCATOR] [--accepted | --baseline | --current] [--json]
docx outline FILE [--style-prefix S] [--json]
docx styles  FILE [--used] [--at STYLEID] [--json]   # the style catalog (not in the body) — what --style NAMEs exist
docx styles  --catalog [--json]                      # built-in styles you can apply on demand (Title, Heading1–9, Quote, …), no FILE needed
docx render  FILE [--out DIR] [--engine word|libreoffice|auto] [--dpi N] [--pages 1-N] [--format png|jpg]

docx comments      list FILE [--include-resolved] [--thread cN]
docx footnotes     list FILE
docx endnotes      list FILE
docx images        list FILE
docx hyperlinks    list FILE
docx track-changes list FILE

docx info schema   [--ts]
docx info locators [--json]
```

`docx read` surfaces structural facts the Markdown body can't show as HTML-comment
annotations (`<!-- docx:TYPE … -->`). These are **read-time visibility hints** — the
agent can SEE the structure, but the importer drops them (the structure survives
normal edits in place, `read --ast` is the lossless view, and `docx sections` /
`docx tables …` manage it). They're emitted **deviation-only**
(only when a value differs from the document default, so a plain document stays
clean):

- **Section breaks** render as `<!-- docx:section sN cols="2" type="continuous" -->`
  on their own line — never a bare `---` (that's a thematic break, and emitting it
  for a section silently turned layout into border paragraphs). A hand-authored
  `---` now unambiguously means a thematic break.
- **Page geometry** rides a leading `<!-- docx:page sN orientation="landscape"
  size="…in" margins="…in" text-width="…in" -->` note when the page deviates from
  US-Letter-portrait-1″ — `text-width` is the usable column width, and the leading
  `sN` is the section to re-apply against. Exact twips are in `read --ast` (on each
  section break: `pageWidth`/`pageHeight`/`pageOrientation`/`margin*`). Set it with
  `docx sections --at sN --orientation/--size/--margins` (the trailing `sN` is the
  whole document) or at `create` time; under track-changes it records as one
  `<w:sectPrChange>` (accept/reject in Word).
- **Tables** carry a leading `<!-- docx:table t0 widths="1,2,3in" borders="double" -->`
  when columns are uneven or borders deviate from the default, plus a per-cell
  `<!-- docx:cell t0:r0c0 gridSpan="2" vMerge="continue" shading="FFE699" -->`
  note on merged/shaded cells — so structure invisible in GFM is visible
  (`Table.borders` / `TableCell.shading` in `read --ast`).
- **Images** trail a `<!-- docx:image img0 size="6.2x4.1in" float="yes" wrap="square" align="center" overflow="yes" -->`
  note: `size` always (the `![](hash)` alone doesn't say "6in wide"), and
  `float`/`wrap`/`align`/`overflow` only when they deviate (an inline, in-bounds
  image shows just its size). `overflow` flags an image wider than the usable text
  column (`ImageRun.floating`/`wrap`/`align` + EMU extents in `read --ast`).

### Mutate (change FILE in place; `--dry-run`, `-v` everywhere; `-o PATH` on every mutator except `create`, whose positional FILE is already the output)

```sh
docx create FILE [--title T] [--author A] [--text "..." | --text-file PATH | --from PATH.md | --from -] [--orientation O] [--size SIZE] [--margins M] [--force]
docx insert FILE (--after | --before) LOCATOR <content>   # LOCATOR = pN | tN | sN | tN:rRcC:pK
docx insert FILE (--at-start | --at-end) <content>        # no locator — prepend / append to the document
docx edit   FILE --at LOCATOR <content>                   # LOCATOR = pN | pN:S-E | pN-pM | sN | eqN | tN:rRcC:pK[:S-E]
docx delete FILE --at LOCATOR                             # LOCATOR = pN | pN-pM | tN | sN
docx sections FILE --at LOCATOR [--columns N] [--type T] [--orientation O] [--size SIZE] [--margins M]   # LOCATOR = pN-pM | pN (wrap a range in N columns) | sN (edit an existing section's columns/type/page geometry). Multi-column layout AND page setup (margins/orientation/size) live HERE; the trailing sN is the whole-document page geometry.
docx styles set-default-font FILE "Font Name" [--size N] [--all]   # document-wide font: sets styles.xml docDefaults + theme major/minor; --all also repoints styles/runs that pin their own font
docx replace FILE PATTERN REPLACEMENT [--regex] [--ignore-case] [--all] [--limit N] [--current | --baseline] [--exact] [--track] [--dry-run]

# Batch — apply many changes from ONE read (no re-reading between edits). Keys
# on each JSONL line mirror the command's flags; all locators address the doc as
# read. insert/edit also accept --batch - to read JSONL from stdin.
docx edit    FILE --batch fills.jsonl       # { at, <one of: text|clear|markdown|runs|code|task>, style?, … }
docx insert  FILE --batch additions.jsonl   # { after|before, <content>, style?, color?, … }
docx replace FILE --batch script.jsonl      # { pattern, replacement, regex?, all?, limit?, … } applied in order
docx delete  FILE --batch drop.jsonl        # { at } per line — whole blocks (pN/tN/cell), resolved live-first

# All four of insert/edit/delete/replace accept --track to record that one
# invocation as a tracked change even when the doc's track-changes toggle is off.
#
# insert/edit content selectors (run "docx insert --help" / "docx edit --help" for the full list):
#   --text "..." [--style NAME] [--alignment A] [--color HEX] [--bold] [--italic] [--url URL]
#       (a newline in --text becomes a line break <w:br/>, a tab becomes <w:tab/> — verse/addresses stay line-per-line)
#   paragraph spacing/indent (insert + edit, alone or with content, per-entry in --batch, across a range):
#       --space-before PT --space-after PT --line-spacing N(=1|1.5|2|single|double, or 15pt / "15pt atLeast")
#       --indent-left IN --indent-right IN --first-line IN --hanging IN  (points / inches; first-line ⊥ hanging;
#       left/right/first-line accept a negative value to outdent into the margin; hanging stays non-negative)
#       Under track-changes these record a tracked <w:pPrChange> (accept/reject in Word) — even when they ride
#       along with --text; read surfaces them as a deviation-only <!-- docx:p … space-after="6pt" --> hint.
#   edit --tabs right   fix a line whose tabbed-over content WRAPS (read flags it as `docx:layout … warn`,
#       and prints ONE consolidated fix-all summary at the top): swaps the fragile LEFT tab for a RIGHT tab
#       flush at the margin so a long value (e.g. a city) never wraps. Rides along with --text, works
#       per-entry in --batch, and on a RANGE (edit --at pN-pM --tabs right) cures every tab line at once.
#   edit --text ""      rejected (it would leave an invisible blank paragraph) — use `delete` to remove the
#       line, or `--runs '[]'` to blank it but keep an empty spacer.
#   --runs '[{"type":"text","text":"X","bold":true}]'
#   --text-file PATH                               # (insert/create) LITERAL multi-paragraph text, NOT parsed — every char verbatim,
#       each newline = a new paragraph. For prose GFM would corrupt: "3. note" stays "3.", *x* / [t](u) / bare URLs / {++x++} untouched.
#   --markdown "..." | --markdown-file PATH        # GFM + math + CriticMarkup + inline HTML formatting → blocks
#   --code "..." | --code-file PATH [--language LANG]
#   --equation "x^2 + y^2" [--display]   (insert; edit also accepts --inline)
#   --clear bold,italic,highlight,color,size,font,…|all   (edit; strip run formatting, keep text)
#   --bold --italic --underline --strike --color HEX --highlight NAME --shade HEX --font NAME --size PT
#       --caps --smallcaps --superscript --subscript   (edit; SET run formatting on EXISTING text —
#       the inverse of --clear. Alone they format a span/paragraph/range in place; with --text they
#       fill AND format. Like --clear, applied directly — not recorded as a tracked change.)
#   --task checked|unchecked | --list bullet|ordered [--list-level N]   (insert)
#   --task checked|unchecked                                            (edit, flip in place)
#   --table --rows N --cols N [--widths "A,B,C"] [--table-width V] [--borders S] [--layout L]   (insert)
#   --image SRC [--alt T] [--width IN] [--height IN] [--caption "Figure 1: …"]   (insert; SRC = path, data: URI, or http(s) URL; --caption adds a Word "Caption"-styled line under the figure)
#   --page-break | --column-break | --section [--columns N] [--type T]   (insert)

docx comments add     FILE --at LOCATOR --text "..." [--author NAME] [--current | --baseline]
docx comments add     FILE --anchor "phrase" --text "..." [--occurrence N]
docx comments add     FILE --batch reviews.jsonl                    # JSONL: { at | anchor (+occurrence), text, author? }
docx comments reply   FILE --at cN --text "..."
docx comments resolve FILE --at cN [--at cM ...] [--unset] | --batch resolutions.jsonl
docx comments delete  FILE --at cN [--at cM ...]          | --batch removals.jsonl

docx footnotes add    FILE --at pN[:offset] (--text "..." | --runs JSON | --markdown TEXT)
docx footnotes edit   FILE --at fnN (--text "..." | --runs JSON | --markdown TEXT)
docx footnotes delete FILE --at fnN
docx endnotes  add    FILE --at pN[:offset] (--text "..." | --runs JSON | --markdown TEXT)
docx endnotes  edit   FILE --at enN (--text "..." | --runs JSON | --markdown TEXT)
docx endnotes  delete FILE --at enN

docx images extract FILE --to DIR [--at imgN]            # --to = output directory; --at picks one image
docx images replace FILE --at imgN --with ./new.png
docx images delete  FILE --at imgN

docx hyperlinks add     FILE --at pN:S-E --url URL
docx hyperlinks replace FILE --at linkN --with URL
docx hyperlinks delete  FILE --at linkN

docx tables insert-row    FILE --at tN [--position INDEX] [--cells "a,b,c"]
docx tables delete-row    FILE --at tN:rR
docx tables insert-column FILE --at tN [--position INDEX] [--width TWIPS]
docx tables delete-column FILE --at tN:cC
docx tables set-widths    FILE --at tN --widths "25%,25%,50%" | "1440,..." | auto
docx tables merge         FILE --at tN:rR1cC1-rR2cC2
docx tables unmerge       FILE --at tN:rRcC
docx tables borders       FILE --at tN [--style single|double|none] [--size N] [--color HEX]

docx track-changes on|off FILE
docx track-changes accept FILE (--at tcN [--at tcM ...] | --at revN | --all)
docx track-changes reject FILE (--at tcN [--at tcM ...] | --at revN | --all)
# A del+ins REPLACE pair shares a "group": "revN" in `list`; `--at revN` accepts/rejects
# both halves in one call (no re-list between them — tcN ids renumber after each accept).
```

> **One rule to memorize: addressing an existing thing is always `--at`.**
> `comments reply/resolve/delete`, `footnotes/endnotes edit/delete`, `images extract/replace/delete`, `hyperlinks replace/delete`, `tables *`, `track-changes accept/reject`, `edit`, and `delete` all take `--at LOCATOR`. The exceptions are positional or directional by nature: `insert` uses `--after`/`--before LOCATOR` (or `--at-start`/`--at-end` for the document boundaries, no locator); `read` slices with `--from`/`--to LOCATOR`; `wc` takes a positional `[LOCATOR]`; `find`/`replace` take a positional `QUERY`/`PATTERN`. `images extract --to DIR` is an *output directory*, not a locator.

## Output contract

The CLI is built for non-interactive agents. **Exit code is the success signal**, output is data:

| Exit | Meaning | Error codes |
| ---- | ------- | ----------- |
| `0`  | success | — |
| `2`  | usage / bad locator | `USAGE`, `INVALID_LOCATOR` |
| `3`  | addressed thing not found | `FILE_NOT_FOUND`, `PART_NOT_FOUND`, `BLOCK_NOT_FOUND`, `COMMENT_NOT_FOUND`, `IMAGE_NOT_FOUND`, `HYPERLINK_NOT_FOUND`, `TRACKED_CHANGE_NOT_FOUND`, `MATCH_NOT_FOUND` |
| `1`  | general failure | `NOT_A_ZIP`, `TRACKED_CHANGE_CONFLICT`, `TABLE_STRUCTURE`, `IMAGE_SOURCE`, `RENDER_ENGINE`, `RENDER_FAILED`, `UNHANDLED` |

**Errors** print `{code, error, hint?}` JSON to stdout with a nonzero exit — note there is **no `ok` field**; the exit code plus `code` are the unambiguous signal.

**The `ok` field appears in exactly one place: the `--verbose` success ack** (`{ok:true, operation, path, …}`). Without `-v`, success output is shaped for the next command:

| Command class | Default stdout on success | `--verbose` |
| ------------- | ------------------------- | ----------- |
| **Mutator that mints a new handle** — `comments add`→`cN`, `comments reply`→`cN`, `footnotes/endnotes add`→`fnN`/`enN`, `hyperlinks add`→`linkN`, `insert`→the new `pN` | the bare locator(s), **one per line** (a multi-block `--markdown` insert prints several) | full `{ok:true,…}` ack |
| **Mutator with no new handle** — `edit`, `delete`, `replace`, `create`, `comments resolve/delete`, `images replace/delete`, `hyperlinks replace/delete`, `footnotes/endnotes edit/delete`, `tables *`, `track-changes accept/reject` & toggle | **one-line confirmation** — `<operation> <target>` (e.g. `edit t1:r0c1:p0`, `edit 7 changes`, `replace 0 occurrences replaced`) (exit `0`) | full `{ok:true,…}` ack |
| `find` | matched span locators, one per line (no matches → nothing, exit `0`) | `--json` → `{ totalMatches, query, view, matches:[…], normalizedQuery? }` |
| `wc` | the bare count (whole-doc adds a tab-separated `sections` column, like `wc`) | `--json` → `{ words, scope, view, sections? }` |
| `outline` | indented `LOCATOR⇥TEXT` tree (two spaces per level) | `--json` → nested `[{ id, locator, level, style, text, children }]` |
| `read` | GFM Markdown, each paragraph trailed by `<!-- pN -->` | `--ast` → the JSON AST body (`docx info schema`) |
| `render` | image paths, one per line | `--verbose` → `{ok, operation, path, engine, output, pages}` |
| `* list` (all six `list` verbs) | a **bare JSON array**; each item's `id` is its `--at` handle | — |

`--dry-run` always prints a preview object (no `ok`) and writes nothing; it wins over `-o/--output`.

## Discovering ids

Locators come in two flavors. **Positional block ids** (`pN`, `tN`, `sN`) are derived from document order and **shift after structural edits** — re-read between non-trivial mutations. **Entity ids** (`cN`, `imgN`, `linkN`, `fnN`, `enN`, `tcN`, `eqN`) are surfaced by a `list` verb (or `read --ast`) and are what you pass to `--at`:

| Id | Discover with | Used by |
| -- | ------------- | ------- |
| `pN` / `tN` / `sN` (block ids) | `docx read FILE` (the `<!-- pN -->` trailers), `docx read FILE --ast`, `docx outline FILE` (heading `pN`s), or `docx render` page images | `read`, `edit`, `insert`, `delete`, `wc`, `find` results |
| `cN` (comment) | `docx comments list FILE` | `comments reply/resolve/delete --at` |
| `fnN` / `enN` (foot/endnote) | `docx footnotes list FILE` / `docx endnotes list FILE` | `footnotes/endnotes edit/delete --at` |
| `imgN` (image) | `docx images list FILE` | `images extract/replace/delete --at` |
| `linkN` (hyperlink) | `docx hyperlinks list FILE` | `hyperlinks replace/delete --at` |
| `tcN` (tracked change) | `docx track-changes list FILE` | `track-changes accept/reject --at` |
| `eqN` (equation) | `docx read FILE --ast` (run `latex` field) | `edit --at eqN --equation` |

Each `list` verb prints a bare JSON array where every item's `id` is exactly the handle you feed back to `--at` — pipe through `jq` to filter (`docx comments list doc.docx | jq '.[] | select(.author=="Jane")'`).

## Locators

`docx info locators` (`--json` for machine-readable) is the canonical reference. The grammar in brief:

```
pN              paragraph N                  pN:S-E          chars S..E within paragraph N
pN-pM           whole-paragraph range        pN:S-pM:E       cross-paragraph character range
sN              section break N              tN              table N
tN:rRcC         cell at row R, col C         tN:rRcC:pK      paragraph K of that cell (chainable)
tN:rR / tN:cC   table row R / column C       tN:rR1cC1-rR2cC2  rectangular cell region (merge)
cN  imgN  linkN  fnN  enN  tcN  eqN          entity ids (comment / image / hyperlink /
                                             footnote / endnote / tracked-change / equation)
```

**Offset semantics: character offsets are 0-based, start-inclusive, end-exclusive** — `p3:5-20` is the 15 characters at indices 5..19 of paragraph 3. Offsets count the *visible* text of the paragraph in the selected view (accepted by default).

**Nested tables chain the same syntax** arbitrarily deep — `t0:r2c1:t0:r0c0:p0` is the first paragraph of the (0,0) cell of the first table nested inside the (2,1) cell of the document's first table.

**Not every command accepts every form** — each command's `--at`/`--from`/positional help lists exactly what it takes. The shapes:

| Form | Accepted by |
| ---- | ----------- |
| `pN`, `tN`, `sN`, `tN:rRcC:pK` (blocks) | `read --from/--to`, `insert --after/--before`, `wc`, `comments add` |
| `pN`, `pN:S-E`, `pN-pM`, `sN`, `eqN`, `tN:rRcC:pK`, `tN:rRcC:pK:S-E` | `edit --at` (span/cell forms strip or replace just that range) |
| `pN`, `pN-pM`, `tN`, `sN` | `delete --at` |
| `pN:S-E`, `pN:S-pM:E`, `tN:rRcC:pK:S-E` (spans) | `comments add --at`, `hyperlinks add --at` (single paragraph), `find`/`wc` results |
| `pN[:offset]` (point) | `footnotes/endnotes add --at` |
| `cN` / `fnN` / `enN` / `imgN` / `linkN` / `tcN` (entities) | the matching noun's `--at` (the `c`/`fn`/`en`/`img`/`link`/`tc` prefix is optional) |
| `tN`, `tN:rR`, `tN:cC`, `tN:rRcC`, `tN:rR1cC1-rR2cC2` | the `tables` verbs |

## Common workflows

**find → comment.** `find` emits bare locators that drop straight into `comments add --at` (same default view, so offsets line up — no coordinate translation):

```sh
docx comments add doc.docx --at "$(docx find doc.docx 'fatally flawed' | head -1)" \
    --text "Cite a source here?"
# or anchor by phrase directly:
docx comments add doc.docx --anchor "fatally flawed" --text "Cite a source here?"
```

**read → edit markdown round-trip.** `read` emits a markdown dialect that `edit --markdown` re-parses, so render → LLM-rewrite → splice-back is lossless for paragraphs/lists/quotes:

```sh
docx read doc.docx --from p3 --to p3              # → markdown (with <!-- p3 --> trailer)
# … hand to an LLM, get a revised block back …
docx edit doc.docx --at p3 --markdown-file revised.md   # multi-block source expands naturally
```

Use `--markdown-file` (not `--markdown TEXT`) when the source starts with `-` — Node's `parseArgs` rejects leading-dash flag values.

**track-changes review loop.** Toggle tracking on, make edits (they auto-emit `<w:ins>`/`<w:del>`), then inventory and resolve:

```sh
docx track-changes doc.docx on
docx replace doc.docx "old phrasing" "new phrasing" --all
docx track-changes list doc.docx                  # → JSON array of { id:tcN, kind, author, text, … }
docx read doc.docx --current                       # → CriticMarkup {++ins++}[^tcN] / {--del--}[^tcN]
docx track-changes accept doc.docx --at tc0 --at tc2   # or --all
```

`read` has three tracked-change views: default **`--accepted`** renders clean text — drops subtractive edits and inlines additive ones (the post-accept document); **`--current`** shows CriticMarkup with `[^tcN]` footnotes; **`--baseline`** does the reverse of accepted (the pre-change document). `find`, `replace`, `wc`, and `comments add` honor the same `--accepted`/`--baseline`/`--current` flags so offsets stay consistent across commands. Add `--comments` to `read` to append `[^cN]` footnotes for comment spans.

## How It Works

**In-place XML mutation.** The AST returned by `read` is a _view_ over the parsed XML tree, not a separate model. When you `edit` or `comments add`, the CLI mutates the underlying XML nodes directly and serializes back. Anything not modeled in the AST (custom styles, theme colors, schema extensions) survives because untouched regions are never re-emitted. Never delete a relationship something still references — that corrupts the file — so part/relationship pruning is gated on a reference scan; unreferenced orphans are left in place.

**JSX for emitters.** Constructing OOXML fragments imperatively (`<w:rPr>` → `<w:b/>` → `<w:color w:val="800080"/>`) is verbose, so fresh XML is authored in JSX with a custom factory: `<w.rPr><w.b/><w.color w-val="800080"/></w.rPr>` becomes the right `XmlNode` tree.

**Span-aware comments & hyperlinks.** `comments add --at p3:5-20` (and `hyperlinks add`) find the runs containing offsets 5 and 20, split them at the boundaries (preserving `<w:rPr>` on both halves), and insert markers between the slices. Comments authored by older tools that lack `w14:paraId` (required by `commentsExtended.xml`) get a fresh paraId injected automatically on resolve/reply.

**Tracked changes.** With `<w:trackChanges/>` set, `insert`/`edit`/`delete`/`replace` emit native `<w:ins>`/`<w:del>` (attributed via `--author`, `$DOCX_AUTHOR`, or `Reviewer`); pass `--track` to one of those commands (or the `tables` verbs / `images delete`) to track just that invocation even when the doc toggle is off. `edit --at pN --text` runs a word-level diff so unchanged words keep their formatting and only changed words are wrapped — the same shape Word produces mid-tracking. `accept`/`reject` handle run-level ins/del/moveFrom/moveTo, `sectPrChange`, paragraph-mark ins/del, and the table-structural revisions (rowIns/rowDel, cellIns/cellDel, tblGridChange, tcPrChange). OOXML has no tracked-change construct for hyperlink edits or image swaps, so under tracking those emit a `[docx-cli]` audit comment instead of a fake revision (image *deletion* is honest removal — it wraps a real `<w:del>`).

**Rich content.** Images insert from a path, `data:` URI, or `http(s)` URL (bounded fetch; HEIC→JPEG transcode; SVG sanitized; non-public/metadata addresses refused at every redirect hop). Equations round-trip OOXML `<m:oMath>` ↔ LaTeX (reconstructed, not legacy plaintext) — authored via temml (LaTeX→MathML) plus an in-house MathML→OMML adapter, no LGPL deps. Code blocks emit one `CodeBlock`-styled paragraph per line with optional lowlight syntax highlighting (37 bundled languages); they collapse back to a GFM fenced block on read. GFM task lists round-trip Word's checkbox content control (and the Word-for-Web Wingdings-glyph variant), surfacing as `taskState` in the AST. Tables operate on a merge-aware logical grid so `gridSpan`/`vMerge` cells map onto physical `<w:tc>`, and structural edits refuse to bisect an existing merge.

**Markdown dialect.** `create --from`, `insert/edit --markdown`, and the note bodies all parse the same GFM + math + CriticMarkup + inline-HTML-formatting dialect (remark + remark-gfm + remark-math + an in-house inline-surgery transform), composing the existing OOXML emitters. `read` emits a compatible dialect, so the read → edit → write loop round-trips (lossless for paragraphs, lists, and nested blockquotes; code/tables/math/headings inside a blockquote intentionally escape to top level on import). `read --ast` is the fully lossless JSON form.

**Literal text — the parser-free channel.** When you want prose inserted *exactly* — reviewer notes, quoted excerpts, anything where Markdown would misfire — use `create --text-file PATH` / `insert --text-file PATH` (or `-` for stdin). Every character lands verbatim and each newline starts a new paragraph; nothing is interpreted, so `3. note` stays `3.` (no ordered-list renumber), and `*x*`, `[t](u)`, bare URLs, and `{++x++}` are kept as written. This exists because GFM corruption isn't always escapable: bare URLs autolink with no escape sequence at all, and CriticMarkup eats `{++…++}` regardless of backslashes — so a literal path is the only safe way to author untouched prose.

**Document-wide font.** `docx styles set-default-font FILE "Times New Roman"` sets the font in the two places a font actually lives — `word/styles.xml` `<w:docDefaults>` (the formal default) *and* the theme font scheme (`word/theme/theme1.xml`, major + minor `<a:latin>`), since real Word docs resolve their fonts *through* the theme and touching only one silently loses to the other. Body text and theme-following headings both adopt it; styles or runs that pin their own font (a code block's monospace, a deliberately-Arial run) are preserved and named in the ack, with `--all` to repoint even those. `--size N` sets the default size on the same write.

**Run formatting beyond bold/italic.** Properties markdown has no native syntax for — text color, theme color, highlight, shading, underline (all 18 styles + color), super/subscript, small/all caps, font, and size — are emitted as the **HTML a markdown reader actually renders**, so the output looks right in GitHub, VS Code, Obsidian, and browsers (Pandoc `[text]{…}` spans render as literal brackets in all of those). `read` emits semantic tags where they exist — `<mark>overdue</mark>`, `<sup>x</sup>`, `<sub>2</sub>` — a `<span style="color:#C00000">…</span>` for the CSS-expressible properties, and `data-*` attributes for the OOXML-only ones CSS can't express (theme colors, underline styles); `insert/edit --markdown` parses them back losslessly, and a leading `<!-- docx:base font="Arial" size="8pt" -->` note declares the document's dominant font/size once so the body isn't buried in per-run repetition. Bold/italic/strike/code/links stay native (`**`/`*`/`~~`/`` ` ``/`[](…)`). Because the inline-surgery transform scans whole sibling sequences, a CriticMarkup marker or span can straddle other formatting — `{++**bold insertion**++}` is tracked correctly. An invalid enum value (e.g. a bogus highlight name) fails with a clear error rather than silently vanishing. Inserted plain content inherits the surrounding paragraph's font/size so it blends in.

**Visual verification.** `docx render` is the only command that needs an external runtime: it drives Microsoft Word (macOS via `osascript`, Windows via PowerShell COM — the ground-truth renderer) or LibreOffice (`soffice`, cross-platform) to produce a PDF, then rasterizes in-process via the bundled `@hyzyla/pdfium` WASM package — no poppler/pdftoppm/ImageMagick needed. Agents that consume PNGs use this to verify edits, diff accept/reject before-vs-after, or generate screenshots.

## Stack

- **Runtime**: Bun (`node:util` parseArgs, JSX with custom factory, native zlib)
- **Parser**: [`jszip`](https://www.npmjs.com/package/jszip) + [`fast-xml-parser`](https://www.npmjs.com/package/fast-xml-parser) + [`fast-xml-builder`](https://www.npmjs.com/package/fast-xml-builder)
- **Markdown**: [`unified`](https://www.npmjs.com/package/unified) + [`remark-parse`](https://www.npmjs.com/package/remark-parse) + [`remark-gfm`](https://www.npmjs.com/package/remark-gfm) + [`remark-math`](https://www.npmjs.com/package/remark-math)
- **Math**: [`temml`](https://www.npmjs.com/package/temml) (MIT) compiles LaTeX → MathML; an in-house MathML → OMML adapter handles the OOXML side bidirectionally
- **Render**: [`@hyzyla/pdfium`](https://www.npmjs.com/package/@hyzyla/pdfium) (MIT wrapper + Apache-2.0 PDFium-as-WASM) for the PDF → PNG/JPG step, plus [`pngjs`](https://www.npmjs.com/package/pngjs) / [`jpeg-js`](https://www.npmjs.com/package/jpeg-js) for image encoding
- **Images**: [`heic-convert`](https://www.npmjs.com/package/heic-convert) (wasm libheif) transcodes HEIC/HEIF input to JPEG on insert
- **Quality**: Biome + Knip + tsc; LibreOffice headless for round-trip integration tests
- **Standard**: ECMA-376 Part 1 §17 (WordprocessingML), Transitional profile

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, architecture overview, and CI.
