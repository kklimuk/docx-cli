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
docx read mnda-filled.docx --to t1

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
docx create FILE [--title T] [--author A] [--text "..." | --from PATH.md]
docx read FILE [--from pN] [--to pN] [--accepted | --baseline | --current] [--comments]
docx read FILE --ast    # JSON-AST opt-in (every other read flag is markdown-only)
docx insert FILE --after p3 --text "..." [--style HeadingN] [--color HEX] [--bold] [--italic] [--url URL]
docx insert FILE --after p3 --runs '[{"type":"text","text":"X","bold":true}]'
docx insert FILE --after p3 --page-break | --column-break
docx insert FILE --after p3 --section [--columns N] [--type continuous|nextPage|evenPage|oddPage|nextColumn]
docx insert FILE --after p3 --table --rows N --cols N [--widths "A,B,C"] [--table-width 100%] [--borders single|none|double] [--layout autofit|fixed]
docx insert FILE --after p3 --image SRC [--alt TEXT] [--width INCHES] [--height INCHES]   # SRC = path, data: URI, or http(s) URL
docx insert FILE --after p3 --code "..." [--language LANG]    # one CodeBlock paragraph per \n; --language → syntax highlight
docx insert FILE --after p3 --code-file PATH [--language LANG] | --code-file -    # same, from file or stdin
docx insert FILE --after p3 --task checked|unchecked --text "..." [--list-level N]   # GFM task list item
docx insert FILE --after p3 --list bullet|ordered --text "..." [--list-level N]      # plain list item
docx insert FILE --after p3 --equation "x^2 + y^2" [--display]   # LaTeX → OMML via temml + own MathML→OMML adapter
docx insert FILE --after p3 --markdown "..." | --markdown-file PATH | --markdown-file -    # GFM + math + CriticMarkup → one or more blocks
docx edit   FILE --at p3 --task checked|unchecked         # flip an existing task's state; track-changes emits checkboxToggle
docx edit   FILE --at eq3 --equation "x^3" [--display|--inline]  # replace equation content and/or toggle display mode
docx edit   FILE --at p3 --text "..." [--no-formatting]   # word-level diff preserves bold/italic on unchanged words
docx edit   FILE --at p3 --runs '[...]'
docx edit   FILE --at p3 --code "..." [--language LANG]   # replace paragraph with a code block (expands to N lines)
docx edit   FILE --at p3 --code-file PATH [--language LANG]
docx edit   FILE --at p2-p5 --text "..."                  # range replace: collapse paragraph range to one paragraph
docx edit   FILE --at p2-p5 --code-file new.go --language go  # range replace with a fresh code block
docx edit   FILE --at p3 --markdown "..." | --markdown-file PATH    # replace paragraph (or pN-pM range) with parsed markdown
docx edit   FILE --at s0 [--columns N] [--type T]         # mutate section properties
docx delete FILE --at p3
docx delete FILE --at p2-p5                               # range delete: drop a contiguous paragraph span
docx delete FILE --at s0                                  # strip an inline sectPr

docx find FILE QUERY [--regex] [--ignore-case] [--all] [--nth N] [--current | --baseline] [--exact]
docx replace FILE PATTERN REPLACEMENT [--regex] [--ignore-case] [--all] [--limit N] [--current | --baseline] [--exact] [--dry-run]

docx wc      FILE [LOCATOR] [--accepted | --baseline | --current]
docx outline FILE
docx render  FILE [--out DIR] [--engine word|libreoffice|auto] [--dpi N] [--pages 1-N] [--format png|jpg]    # per-page PNG/JPG via Word or LibreOffice

docx comments add     FILE --range p3:5-20 --text "..." [--author NAME] [--current | --baseline]
docx comments add     FILE --anchor "phrase" --text "..." [--occurrence N]
docx comments add     FILE --batch reviews.jsonl              # JSONL: { range | anchor (+occurrence), text, author? }
docx comments reply   FILE --to c0 --text "..."
docx comments resolve FILE --id c0 [--unset]
docx comments resolve FILE --id c1 --id c3 [--unset]          # repeatable
docx comments resolve FILE --batch resolutions.jsonl
docx comments delete  FILE --id c0
docx comments delete  FILE --id c1 --id c3                    # repeatable
docx comments delete  FILE --batch removals.jsonl
docx comments list    FILE [--include-resolved] [--thread c0]

docx footnotes add    FILE --at pN[:offset] --text "..."
docx footnotes edit   FILE --id fnN --text "..."
docx footnotes delete FILE --id fnN
docx footnotes list   FILE
docx endnotes  add    FILE --at pN[:offset] --text "..."
docx endnotes  edit   FILE --id enN --text "..."
docx endnotes  delete FILE --id enN
docx endnotes  list   FILE

docx images list    FILE
docx images extract FILE --to ./media [--id imgN]
docx images replace FILE --at imgN --with ./new.png
docx images delete  FILE --at imgN

docx hyperlinks list    FILE
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

docx track-changes FILE on|off
docx track-changes list   FILE
docx track-changes accept FILE (--at tcN [--at tcM ...] | --all)
docx track-changes reject FILE (--at tcN [--at tcM ...] | --all)
docx info schema [--ts]
docx info locators [--json]
```

Every command has `--help`. Mutating commands accept `--dry-run`, `-o/--output PATH` (write to a parallel file instead of overwriting `FILE`), and `-v/--verbose` (print the JSON ack — see "Quiet by default" below).

**Quiet by default.** Mutators (`create`, `insert`, `edit`, `delete`, `replace`, `comments add/reply/resolve/delete`, `footnotes add/edit/delete`, `endnotes add/edit/delete`, `images replace/delete`, `hyperlinks add/replace/delete`, `tables *`, `track-changes` toggle/accept/reject) print nothing on success and exit 0. Errors always print as `{ok: false, code, error, hint}`. Pass `-v`/`--verbose` to get the full JSON ack. Read commands (`read`, `find`, `wc`, `outline`, `info *`, `*-list`) print their data unconditionally. Batch operations that mint new ids — `comments add --batch`, `comments delete --batch`, `comments resolve --batch`, and `comments delete/resolve` with multiple `--id` — always print the affected ids since the agent can't reconstruct them otherwise.

### Markdown rendering

`docx read FILE` renders the document body as GitHub-flavored Markdown — the default since v0.9. Useful when an LLM wants to skim a doc without parsing the AST. Each rendered paragraph is followed by an HTML comment with its locator (`<!-- p3 -->`) so the markdown is invisible-pinned: humans see clean prose in a renderer, agents parse the locators from raw text. Pass `--ast` for the structured JSON when programmatic walks need the typed tree (`docx read FILE --ast | jq '.blocks[] | select(.type == "paragraph")'`).

`--from LOC` and `--to LOC` slice by top-level block (both inclusive). Accepts paragraph, table, cell, span, and range locators; cell/span/range collapse to their enclosing top-level block.

**Tracked changes — three views.** Default is the **accepted** view: `<w:del>` and `<w:moveFrom>` runs are dropped; `<w:ins>` and `<w:moveTo>` runs render as plain text. Three mutually-exclusive flags select the view (the default needs no flag):

- `--accepted` — explicit alias for the default. Post-accept view.
- `--baseline` — pre-change view: drops `<w:ins>` and `<w:moveTo>`, renders `<w:del>` and `<w:moveFrom>` as plain text.
- `--current` — raw concatenation, with diff markup. Additive wrappers render as CriticMarkup `{++text++}[^tcN]` and subtractive as `{--text--}[^tcN]`, with `[^tcN]: …` definitions appended at end.

The same `--current`/`--baseline` flags apply to `find`, `replace`, `wc`, and `comments add` so offsets stay consistent across commands (default everywhere is the accepted view). `comments add` uses the same view to resolve `--anchor` matches and `--range` offsets — agents can pipe `find` output to `comments add` without coordinate translation.

`docx wc` also accepts `sN` as a locator — the count covers every paragraph and table in that section's content range. Whole-document `wc` returns a `sections` count alongside `words`.

`find` and `replace` auto-normalize their query/pattern by default: balanced markdown emphasis around non-whitespace (`**X**`, `__X__`, `*X*`, `` `X` ``) is stripped; smart and straight quotes are equivalent; em-dash and en-dash collapse to a hyphen. Pass `--exact` to disable normalization. `--regex` is always verbatim. `find`'s response surfaces `normalizedQuery` / `normalizationApplied` when normalization changed the query so the agent sees what was actually searched.

`--comments` appends a GFM footnote reference (`[^cN]`) at the end of each commented span and emits one footnote definition per comment at the end of the output. Footnotes/endnotes (the document's own `<w:footnoteReference>` / `<w:endnoteReference>`) are rendered unconditionally as `[^fnN]` / `[^enN]`. Run `docx info schema` for the full mapping table.

### Markdown authoring

`docx create FILE --from PATH.md`, `docx insert FILE --after pN --markdown "..."` / `--markdown-file PATH`, and `docx edit FILE --at pN --markdown ...` (or `--at pN-pM` for range replace) all accept the same dialect: GitHub-flavored Markdown via `remark-gfm` (tables, strikethrough, footnotes, task lists, autolinks), inline + display math via `remark-math` (`$x^2$`, `$$x^2$$`), and CriticMarkup (`{++ins++}` / `{--del--}`). The walker emits real OOXML — fenced code becomes `CodeBlock-LANG` paragraphs, tables become `<w:tbl>` with even-grid columns, lists provision `numbering.xml` lazily, footnote refs/defs author `footnotes.xml`, images fetch bytes (path, `data:`, or `http(s)`) and mint media parts. CriticMarkup wraps in real `<w:ins>` / `<w:del>` under tracking, and respects the accepted view (`{++X++}` keeps `X`, `{--X--}` drops it) when tracking is off.

`docx read FILE` (which already defaults to markdown render) emits a compatible dialect, so the **read → edit → write** loop round-trips: render a doc to markdown, hand it to an LLM, splice the result back via `--markdown-file`. Use `--markdown-file` (not `--markdown TEXT`) when the source starts with `-` — Node's `parseArgs` rejects leading-dash flag values.

**Blockquote round-trip** is lossless for **paragraphs, lists, and nested blockquotes** (encoded as `pStyle="Quote"` / `"QuoteListParagraph"` + `<w:ind w:left={720 * depth}>`, recovered as `paragraph.quoteDepth` on read). **Code blocks, tables, math, headings, and HRs inside a blockquote intentionally escape** — they emit at top level on import, breaking the surrounding quote. Adjacent quoted content before and after surfaces as separate blockquotes on re-read. See [src/core/markdown/CLAUDE.md](src/core/markdown/CLAUDE.md) for why.

### Visual verification

`docx render FILE` produces one PNG (or JPG) per page using Microsoft Word as the ground-truth renderer on macOS / Windows, with LibreOffice as the cross-platform fallback. Auto-detect picks the highest-fidelity engine available on the machine; `--engine word|libreoffice` overrides. The JSON ack lists the page paths so agents reading PNGs can iterate over them programmatically — same shape as `comments add --batch` printing minted ids:

```sh
$ docx render report.docx --out ./snapshots
{"ok":true,"operation":"render","path":"report.docx","engine":"word-mac",
 "output":"./snapshots","pages":["./snapshots/page-001.png", ...]}
```

**Runtime requirements** (none for any other command — `render` is the only verb that needs an external runtime):

- **Word engine**: Microsoft Word installed locally. macOS drives Word via `osascript` (first run triggers a one-time Automation permission prompt that has to be granted); Windows drives Word via PowerShell COM. Linux isn't supported by the Word path — use `--engine libreoffice` instead.
- **LibreOffice engine**: `soffice` on PATH or installed at the canonical location (`brew install --cask libreoffice` on macOS, `apt install libreoffice` on Linux, libreoffice.org on Windows).

PDF rasterization is built in — the PDFium WASM binary (BSD-3-Clause / Apache-2.0 dual) is embedded into every shipped artifact: `dist/index.js` carries it as a sibling `dist/pdfium-<hash>.wasm` asset; the `bun build --compile` standalone binary embeds it directly inside the executable. No `poppler` / `pdftoppm` / `imagemagick` install required.

Use cases: visual verification of an agent's edits, comparing tracked-change accept/reject before vs after, generating screenshots for PR descriptions or bug reports.

### Bulk + anchored comments

`comments add --anchor "phrase"` resolves the phrase via the same matcher as `docx find` (default accepted view, query normalization on) and anchors the comment to that match. Pass `--occurrence N` (1-indexed) to disambiguate when the phrase matches more than once; without it, an ambiguous anchor errors atomically before any writes.

`comments add --batch FILE.jsonl` (or `--batch -` for stdin) takes one JSON object per line: `{range | anchor (+ optional occurrence), text, author?}`. The whole batch validates against the pre-mutation tree first and aborts cleanly if any entry fails. `comments delete` and `comments resolve` accept `--batch FILE.jsonl` (`{"id": "cN"}` per line) or repeated `--id c1 --id c3`. Atomic in all cases — no partial writes on error.

### Formatting preservation

`docx edit --at pN --text "..."` runs a word-level diff between the existing paragraph's text and the new text, preserving `<w:rPr>` formatting (bold, italic, color, etc.) on unchanged words. New words inherit formatting via position-pairing with the deleted span (Kth inserted word inherits from the Kth deleted word, falling back to neighboring kept words when the edit group has no deletes). Pass `--no-formatting` to fall back to a single fresh run with no formatting; explicit `--color`/`--bold`/`--italic` also bypass preservation and apply uniformly to the new paragraph. Under tracking, the result is per-word `<w:del>`/`<w:ins>` markers (the same shape Word produces when an author edits a few words mid-tracking) instead of whole-paragraph del+ins.

### Atomic batch accept/reject

`track-changes accept --at tc1 --at tc2 --at tc3` resolves all targets against the pre-mutation tree, deduplicates, and applies in a single call. Mid-batch renumbering doesn't shift the still-pending ids out from under the agent. Mutually exclusive with `--all`. Same shape for `reject`.

### Locators

```
pN                   paragraph N (e.g., p3)
pN-pM                paragraph range (whole paragraphs pN..pM as a unit)
pN:S-E               characters S..E within paragraph N
pN:S-pM:E            cross-paragraph character range
tN                   table N; tN:rRcC for cell at row R, col C
sN                   section break N — column count, type (continuous /
                     nextPage / nextColumn / evenPage / oddPage)
cN, imgN, linkN, tcN comment / image / hyperlink / tracked-change ids
```

Run `docx info locators` for the full reference.

## How It Works

**In-place XML mutation.** The AST returned by `read` is a _view_ over the parsed XML tree, not a separate model. When you `edit` or `comments add`, we mutate the underlying XML nodes directly and serialize back. Anything we don't model in the AST (custom styles, theme colors, schema extensions) survives because we never re-emit untouched regions.

**JSX for emitters.** Constructing OOXML fragments imperatively (`<w:rPr>` → `<w:b/>` → `<w:color w:val="800080"/>`) gets verbose. We write emitters in JSX with a custom factory: `<w.rPr><w.b/><w.color w-val="800080"/></w.rPr>` becomes the right `XmlNode` tree. Component names are PascalCase (`<Paragraph>`, `<RunProperties>`); they return `XmlNode | null` so empty wrappers get omitted automatically.

**Span-aware comments.** `comments add --range p3:5-20` finds the runs that contain offsets 5 and 20, splits them at the boundaries (preserving rPr formatting on both halves), and inserts `<w:commentRangeStart>` / `<w:commentRangeEnd>` markers between the slices.

**ParaId auto-injection.** Comments authored by tools like mammoth or older Word versions lack `w14:paraId`, which `commentsExtended.xml` requires for resolve/reply. We detect this on resolve/reply and inject a fresh paraId, also adding the `xmlns:w14` namespace declaration to the `<w:comments>` root if missing.

**Cross-format image replacement.** `images replace --at img0 --with new.png` detects the new MIME type via `Bun.file().type`, renames the part (`word/media/image1.jpeg` → `word/media/image1.png`), rewrites the relationship `Target`, and ensures `[Content_Types].xml` has a `<Default>` for the new extension.

**GFM task lists.** Paragraphs in a list whose leading content is a Word checkbox content control (`<w:sdt><w14:checkbox/></w:sdt>`) surface as `paragraph.taskState: "checked" | "unchecked"` in the AST and render as `- [ ]` / `- [x]` in markdown. The reader strips the SDT subtree and its trailing space run from `runs` so the AST carries only the task text; the SDT survives untouched in the underlying `XmlNode` tree so a read→edit→save round-trip preserves every checkbox. The reader also recognizes the **Word-for-Web Checklist** shape — a bulleted list whose level-0 bullet character is Wingdings ☐ (U+F0A8), with `<w:strike>` on the paragraph-mark marking "done" — since Web silently strips SDT content controls when it authors new task lists. Both shapes were confirmed empirically against Microsoft Word for Mac desktop, Word for Web, and LibreOffice. On the emit side we always produce SDT. **Authoring**: `insert --task checked|unchecked --text "..."` creates a fresh task line (inherits the anchor's numId if it's already a list, otherwise allocates a new bullet list); `--list-level N` nests; `--list bullet|ordered` makes a plain list item without a checkbox. `edit --at pN --task checked|unchecked` flips an existing task's state in place. **Tracked toggles** (checking or unchecking under track-changes) surface as the `checkboxToggle` tracked-change kind in `track-changes list`; accept keeps the new state, reject restores the prior glyph AND flips `w14:checked` back (inferred from the kept glyph, since Word stores no separate prior-value record).

**Math equations.** `<m:oMath>` and `<m:oMathPara>` (the Office Math markup Word, Pandoc, and LibreOffice all emit) surface as `EquationRun.latex` in the AST — reconstructed LaTeX, not the legacy plaintext concatenation. Markdown render emits `$…$` inline and `$$…$$` for display equations, so a doc full of academic equations round-trips through Pandoc with high fidelity. The walker handles fractions, super/sub/mixed scripts, roots (incl. nth roots), n-ary operators (sum/product/integral/contour), accents (hat/bar/vec — both Word's combining-diacritic and Pandoc's spacing-overscript encodings), delimited expressions, matrices, aligned equation arrays, and function operators (`\sin`/`\lim`/`\log` promoted from upright-styled runs). Unrecognized OMML constructs degrade per-subtree to plaintext (`EquationRun.text` is kept as a fallback) so a niche construct doesn't corrupt the surrounding equation. **Authoring**: `insert --equation "x^2 + y^2 = r^2" [--display]` inserts an inline or display equation from LaTeX (parsed by [temml](https://github.com/ronkok/Temml) for the LaTeX-side spec coverage, then walked into OMML by our own MathML → OMML adapter — no LGPL deps). `edit --at eqN --equation NEW_LATEX` replaces the content; `--display` / `--inline` toggle the mode. `eqN` locators address equations in document order.

**Code blocks + inline code.** `insert --code TEXT` (or `--code-file PATH`, `-` for stdin) splits content on `\n` and emits one `<w:p>` per source line, all styled `CodeBlock` (Courier New, indent, adjacent-paragraph spacing collapse) with runs styled `Code` (monospace character style — defensive in case Word doesn't cascade the paragraph font). Both styles get provisioned in `styles.xml` automatically. `--language LANG` syntax-highlights via [lowlight](https://github.com/wooorm/lowlight) (highlight.js); 37 common languages are bundled — `bash`, `c`, `cpp`, `csharp`, `css`, `diff`, `go`, `graphql`, `ini`, `java`, `javascript`, `json`, `kotlin`, `less`, `lua`, `makefile`, `markdown`, `objectivec`, `perl`, `php`, `php-template`, `plaintext`, `python`, `python-repl`, `r`, `ruby`, `rust`, `scss`, `shell`, `sql`, `swift`, `typescript`, `vbnet`, `wasm`, `xml`, `yaml`. Unknown languages degrade silently to uncolored runs. The palette is GitHub-light inspired (keywords red, strings dark-blue, comments gray, …) and unmapped highlight.js classes fall through with no color. For an inline ``code`` span inside a normal paragraph, use `--runs` JSON with `runStyle: "Code"` (the S8 markdown walker will make this ergonomic via the `\`code\`` shorthand). On `docx read`, consecutive `CodeBlock` paragraphs collapse into one GFM fenced block (`` ``` `` … `` ``` ``); inline `runStyle: "Code"` runs render with backticks.

**Image insertion.** `insert --image SRC` resolves SRC from a file path, a `data:` URI, or an `http(s)` URL (bounded fetch: 10s timeout, 25 MB cap streamed per-chunk so the limit holds even if `Content-Length` lies), writes the bytes to `word/media/imageN.ext`, mints an `image` relationship, and registers the extension's content-type `<Default>`. Pixel dimensions are read from the PNG/JPEG/GIF header and converted to EMU (1px = 9525 EMU at 96 dpi) for `<wp:extent>`; `--width`/`--height` override in inches, and supplying one alone scales the other to preserve aspect. The drawing is a standard inline `<w:drawing><wp:inline>` picture (`a:`/`pic:` namespaces declared on the subtree). Under track-changes the inserted run is wrapped in `<w:ins>` like any other inserted content. **HEIC/HEIF** input (common from iPhones) is transcoded to JPEG before embedding — Word can't render HEIC — so students can drop a `.heic` straight in; detection is by file header, not just extension. **SVG input is sanitized** before embedding (`<script>`, `on*` handlers, `<foreignObject>`, animation events, external `href`/`xlink:href`, and `data:image/svg+xml` self-references are stripped; XXE is rejected at parse time by `fast-xml-parser`) so an attacker-controlled SVG can't smuggle active content into the doc. **Remote fetches block non-public addresses** — private, loopback, link-local, and cloud-metadata ranges are refused, and HTTP redirects are followed manually with the same check at every hop, so an agent steered into `http://169.254.169.254/...` or `http://10.0.0.1/admin` is short-circuited.

**Image deletion.** `images delete --at imgN` removes the inline drawing and its run, pruning the media part and relationship when nothing else references them. Under track-changes it wraps the run in a real `<w:del>` instead (accept removes, reject restores), keeping the part until the change is accepted.

**Hyperlink CRUD.** `hyperlinks list` enumerates `<w:hyperlink>` elements with positional `linkN` ids; `hyperlinks add --at p3:5-20 --url URL` wraps an existing span (splitting runs at offsets); `hyperlinks replace --at link0 --with URL` updates the rels `Target`, allocating a new rId if the existing one is shared so siblings stay pointed at the original URL; `hyperlinks delete --at link0` unwraps the link (text survives) and prunes the rels entry when no longer referenced.

**Table restructuring.** The `tables` verbs operate on a merge-aware logical grid: `gridSpan` (horizontal) and `vMerge` (vertical) cells map to physical `<w:tc>` elements so row/column locators (`tN:rR`, `tN:cC`, region `tN:rR1cC1-rR2cC2`) resolve correctly. `merge`/`unmerge` reshape `gridSpan`/`vMerge`; `set-widths` rewrites `<w:tblGrid>` plus per-cell `<w:tcW>`; structural edits refuse to bisect or orphan an existing merge (with a hint to `unmerge` first). Under track-changes, the tracked representation was verified against Microsoft Word (accept/reject), since Word — not the ECMA-376 schema — decides what actually round-trips. Row insert/delete emit native `<w:trPr><w:ins>`/`<w:del>`; column insert/delete emit per-cell `<w:tcPr><w:cellIns>`/`<w:cellDel>` (paired with a `<w:tblGridChange>` on insert); `set-widths` emits `<w:tblGridChange>` plus a per-cell `<w:tcPrChange>` (which is what Word's reject actually reverts) — all addressable as `tcN` and resolvable via `track-changes accept`/`reject` (which resyncs the grid). Cell merges and border changes are *not* tracked by Word (it warns "this action won't be marked as a change" and applies them immediately), so `merge`/`unmerge`/`borders` match that — applied in place with a `[docx-cli]` audit comment (mirroring hyperlink/image edits).

## Stack

- **Runtime**: Bun (`node:util` parseArgs, JSX with custom factory, native zlib)
- **Parser**: [`jszip`](https://www.npmjs.com/package/jszip) + [`fast-xml-parser`](https://www.npmjs.com/package/fast-xml-parser) + [`fast-xml-builder`](https://www.npmjs.com/package/fast-xml-builder)
- **Markdown**: [`unified`](https://www.npmjs.com/package/unified) + [`remark-parse`](https://www.npmjs.com/package/remark-parse) + [`remark-gfm`](https://www.npmjs.com/package/remark-gfm) + [`remark-math`](https://www.npmjs.com/package/remark-math) drive the import path (`docx create --from`, `docx insert/edit --markdown`)
- **Math**: [`temml`](https://www.npmjs.com/package/temml) (MIT) compiles LaTeX → MathML; our own MathML → OMML adapter handles the OOXML side bidirectionally
- **Render**: [`@hyzyla/pdfium`](https://www.npmjs.com/package/@hyzyla/pdfium) (MIT wrapper + Apache-2.0 PDFium-as-WASM) for `docx render`'s PDF → PNG/JPG step, plus [`pngjs`](https://www.npmjs.com/package/pngjs) / [`jpeg-js`](https://www.npmjs.com/package/jpeg-js) for the image encoding — no `poppler`/`pdftoppm`/`imagemagick` install required
- **Images**: [`heic-convert`](https://www.npmjs.com/package/heic-convert) (wasm libheif) transcodes HEIC/HEIF input to JPEG on insert
- **Quality**: Biome + Knip + tsc; LibreOffice headless for round-trip integration tests
- **Standard**: ECMA-376 Part 1 §17 (WordprocessingML), Transitional profile

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, architecture overview, and CI.
