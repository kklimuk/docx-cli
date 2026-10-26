# docx-cli reference

The full command map, the output contract, the locator grammar, and how docx-cli works under the hood. For install and the overview, see the [README](README.md).

## `docx <command> --help` is the authoritative contract

> **Agents: run `docx <command> --help` before composing a call.** Every command's `--help` is the source of truth for its flags, locator forms, and exact output shape — this page is a map, not the territory. Two more must-reads:
>
> - **`docx info locators`** — the canonical locator grammar (`--json` for a machine-readable form). The top-level `docx --help` says it outright: _"It is HIGHLY RECOMMENDED to run `docx info locators` to understand the addressing model."_
> - **`docx info schema`** — the AST type definitions (`--ts` for TypeScript source) that `read --ast` emits.

## Command reference

`docx <verb>` and `docx <noun> <verb>`. Every command has `--help`. Two groups: **read/query** commands print data to stdout; **mutate** commands change the file (and accept `--dry-run`, `-o/--output PATH`, `-v/--verbose`).

### Read & query (print to stdout, never write the file)

```sh
docx read    FILE [--from LOC] [--to LOC] [--accepted | --baseline | --current] [--comments]
docx read    FILE --ast                  # JSON-AST instead of Markdown (disables the markdown-only flags)
docx find    FILE QUERY [--regex] [--ignore-case] [--all] [--nth N] [--current | --baseline] [--exact] [--json]
docx find    FILE (--highlight COLOR|any | --color HEX | --bold | --italic | --underline) [--all] [--json]   # find by formatting (no QUERY)
docx find    FILE --batch queries.jsonl [--json]   # many independent text/format queries, one read; --batch - reads stdin
docx wc      FILE [LOCATOR] [--accepted | --baseline | --current] [--json]
docx outline FILE [--style-prefix S] [--json]
docx diff    FILE --against SRC [--from LOC] [--to LOC] [--comments] [--json]   # what changed vs another version
docx validate FILE [--json]              # ECMA-376 schema check, per WML part (exit 0 = clean)
docx raw get FILE --at LOCATOR [--json]  # a block's exact XML — the read half of the raw patch loop
docx styles  FILE [--used] [--at STYLEID] [--json]   # the style catalog (not in the body) — what --style NAMEs exist
docx styles  --catalog [--json]                      # built-in styles you can apply on demand (Title, Heading1–9, Quote, …), no FILE needed
docx styles  set    FILE --at STYLEID [--bold --color HEX --size PT --font NAME --font-east-asia NAME --font-complex-script NAME --space-before PT --indent-left IN …]   # restyle every paragraph/run that uses the style
docx styles  create FILE STYLEID [--type paragraph|character] [--name "…"] [--based-on STYLEID] [--next STYLEID] [formatting]   # define a new custom style
docx render  FILE [--out DIR] [--engine word|libreoffice|auto] [--dpi N] [--pages 1-N] [--format png|jpg]
docx upgrade [--to TAG] [--dry-run]                                        # update the installed binary in place (no FILE)

docx comments      list FILE [--open] [--thread cN]
docx footnotes     list FILE
docx endnotes      list FILE
docx headers       list FILE
docx footers       list FILE
docx images        list FILE
docx hyperlinks    list FILE
docx track-changes list FILE

docx info schema   [--ts]
docx info locators [--json]
```

`find --batch` reads one JSON object per line: either `{"query":"…"}` with
optional `regex`/`ignoreCase`/`nth`/view fields, or a formatting filter such as
`{"highlight":"yellow"}` / `{"bold":true}`. Plain output flattens locators in
entry order; `--json` preserves each query's `totalMatches` and `matches` inside
a `batch` array. It opens the document once and never mutates it.

`docx diff` shows **what you changed** as a git-style unified diff — it renders both
the current file and a baseline to their `read` markdown and diffs them, so `+` lines
are what you added and `-` lines what you removed (a one-line edit shows inline as
`[-removed-]{+added+}`). There's no undo/history, so **snapshot before editing** and
diff against the snapshot:

```sh
cp report.docx report.orig.docx          # keep the original
docx edit report.docx --at p3 --text "…" # make changes
docx diff report.docx --against report.orig.docx
```

`--against` also takes a saved `docx read` output (a text file), or `-` to read the
baseline from stdin — so you can diff against a git branch without any git integration:
`git show main:report.docx | docx diff report.docx --against -`. Positional locators
(`<!-- p3 -->`, cell/row ids) are normalized out so a structural edit doesn't renumber
every following line; formatting/structure changes (shading, borders, tracked-changes
state) still show. It's a read-only report, not a patch to apply.

`docx read` surfaces structural facts the Markdown body can't show as HTML-comment
annotations (`<!-- docx:TYPE … -->`). These are **read-time visibility hints** — the
agent can SEE the structure, but the importer drops them (the structure survives
normal edits in place, `read --ast` is the lossless view, and `docx sections` /
`docx tables …` manage it). They're emitted **deviation-only**
(only when a value differs from the document default, so a plain document stays
clean):

- **Per-paragraph style/spacing/indent** — the most common annotation — rides a
  `<!-- docx:p pN style="Caption" align="center" space-after="6pt" line-spacing="1" indent-left="0.25in" -->`
  note, emitted deviation-only (only the attrs that differ from the style/document
  default). Each attribute maps to the matching `edit`/`insert` flag (`--style`,
  `--alignment`, `--space-before`/`--space-after`, `--line-spacing`,
  `--indent-left`/`--indent-right`/`--first-line`/`--hanging`), so an agent reads a
  value and re-applies it. The paragraph's locator rides this note as its leading `pN`
  token, so an annotated paragraph does NOT also get a bare `<!-- pN -->` (only
  undeviating paragraphs get the bare locator). Full properties are in `read --ast`.
- **Section breaks** render as `<!-- docx:section sN cols="2" type="continuous" -->`
  on their own line — never a bare `---` (that's a thematic break, and emitting it
  for a section silently turned layout into border paragraphs). A hand-authored
  `---` now unambiguously means a thematic break.
- **Page geometry** rides a leading `<!-- docx:page sN orientation="landscape"
size="…in" margins="…in" text-width="…in" -->` note when the page deviates from
  US-Letter-portrait-1″ — `text-width` is the usable column width, and the leading
  `sN` is the section to re-apply against. A `varies="by-section"` attribute is added
  when a later section's page setup differs from the leading one — and in that case
  the note fires **even if page 1 is plain default Letter-portrait-1″** (it then shows
  just `text-width` + `varies="by-section"`), warning that the geometry shown describes
  only the leading section; use `read --ast` for every section's exact geometry.
  Exact twips are in `read --ast` (on each
  section break: `pageWidth`/`pageHeight`/`pageOrientation`/`margin*`). Set it for the
  WHOLE document with `docx sections --orientation/--size/--margins` (no `--at` →
  every section gets it, so a multi-section doc doesn't leave the trailing section
  behind), one section with `docx sections --at sN …`, or at `create` time; under
  track-changes it records as one `<w:sectPrChange>` per section (accept/reject in Word).
  Changing margins/size also **auto-realigns right-edge tab columns** (résumé
  dates/locations): a LEFT tab calibrated to the old margins would overflow and wrap
  at the new width, so page setup converts each to a RIGHT tab flush at the new
  margin and reports how many it fixed — no second `--tabs right` step needed.
- **Tables** carry a leading `<!-- docx:table t0 widths="1,2,3in" borders="double" -->`
  when columns are uneven or borders deviate from the default. Within each cell,
  a trailing `<!-- docx:cell t0:r0c0 gridSpan="2" vMerge="continue" shading="FFE699" -->`
  note follows the cell content and its paragraph locators — so structure invisible in GFM is visible
  (`Table.borders` / `TableCell.shading` in `read --ast`). A plain empty cell emits
  `<!-- t0:r0c0 -->`: pass that bare handle directly to `edit --at` or `insert --at`
  to fill its mandatory empty paragraph without creating a blank extra line. Where a
  bare cell target is refused — a merged cell, or any cell of a merged/grid-shifted
  row — the empty cell emits the explicit `<!-- t0:r0c1:p0 -->` paragraph locator
  instead (in front of its `docx:cell` note, when it has one). Pass back whichever
  form you see rather than deriving one by counting columns: in a merged row the
  printed cell ids no longer line up with logical column numbers.
- **Content controls** (Word's `<w:sdt>` — template placeholders, form fields,
  data-bound regions) are transparent: the paragraphs and tables inside one read
  as ordinary `pN`/`tN` blocks and edits land inside the control, with a
  `content-control="Governing Law"` pair on the block's note so you know Word may
  have locked it.
- **Images** trail a `<!-- docx:image img0 size="6.2x4.1in" float="yes" wrap="square" align="center" overflow="yes" -->`
  note: `size` always (the `![](hash)` alone doesn't say "6in wide"), and
  `float`/`wrap`/`align`/`overflow` only when they deviate (an inline, in-bounds
  image shows just its size). `overflow` flags an image wider than the usable text
  column (`ImageRun.floating`/`wrap`/`align` + EMU extents in `read --ast`).
- **Headers / footers** surface as `<!-- docx:header hdr0 text="Quarterly Report" -->`
  / `<!-- docx:footer ftr0 text="Page {page} of {pages}" -->` notes — led by the
  marginal's `hdrN`/`ftrN` id, the handle `raw get --at` and `headers`/`footers
  set/clear --at` accept (the `type` attr
  appears only for `first`/`even`). Fields read as tokens — `{page}` `{pages}`
  `{date}` `{time}` `{styleref:NAME}` `{filename}` `{title}` `{author}` (`{time}`
  read-only). A marginal that's
  the same on every section rides the top; one that **differs by section** renders at
  that section's **start** (alongside the `docx:section` note, which also renders at
  the section's start with `applies-to="… (below)"`), so each hint reads right before
  the content it governs. The text lives in the comment attribute
  so the importer drops it (it can't re-inject into the body); full entries are in
  `read --ast` under `headers`/`footers` (`Marginal[]`). Set with `docx
headers`/`docx footers`.
- **Track-changes state** always rides a head `<!-- docx:track-changes on|off -->`
  line — it's the one read hint that states its default too (an agent shouldn't have
  to infer "off" from a missing line), so you can see whether subsequent edits will be
  redlined without inspecting `settings.xml`.
  Toggle it with `docx track-changes on|off FILE`; the three tracked-change read views
  (`--accepted`/`--current`/`--baseline`) are covered under the review loop below.

### Mutate (change FILE in place; `--dry-run`, `-v` everywhere; `-o PATH` on every mutator except `create`, whose positional FILE is already the output)

```sh
docx create FILE [--title T] [--author A] [--text "..." | --text-file PATH | --from PATH.md | --from -] [--orientation O] [--size SIZE] [--margins M] [--header "..."] [--footer "..." | --page-numbers] [--force]
docx insert FILE --at LOCATOR <content>                  # after a block; INTO a bare simple cell (fills/reuses blank p0)
docx insert FILE (--after | --before) LOCATOR <content>   # explicit block side or bare-cell end/start boundary
docx insert FILE (--at-start | --at-end) <content>        # no locator — prepend / append to the document
docx edit   FILE --at LOCATOR <content>                   # LOCATOR = pN | pN:S-E | pN-pM | tN:rRcC | tN:rRcC:pK[:S-E]; bare CELL requires one direct paragraph
docx delete FILE --at LOCATOR                             # LOCATOR = pN | pN-pM | tN | sN | tN:rRcC:pK (cell paragraph)
docx sections FILE [--at LOCATOR] [--columns N] [--type T] [--orientation O] [--size SIZE] [--margins M]   # LOCATOR = pN-pM | pN (wrap a range in N columns) | sN (edit one section's columns/type/page geometry). Multi-column layout AND page setup live HERE. PAGE GEOMETRY (margins/orientation/size) with NO --at applies to the WHOLE document (every section); --at sN targets one. Columns/type need --at.
docx styles set-default-font FILE "Font Name" [--size N] [--all]   # document-wide font: sets styles.xml docDefaults + theme major/minor; --all also repoints styles/runs that pin their own font
docx replace FILE PATTERN REPLACEMENT [--at pN] [--regex] [--ignore-case] [--all] [--limit N] [--current | --baseline] [--exact] [--clear ATTRS] [run-formatting] [--track] [--dry-run]
#   Keeps the run's formatting (bold/font) and any tabs by default — the no-rebuild way to fill a
#   formatted/tabbed template line (e.g. "**Org Name**⇥Date"); don't hand-build --runs to refill it.
#   Add --bold/--italic/--color/--highlight/--font/--size/… to format ONLY the replacement text;
#   --clear highlight (or a comma list / all) strips inherited properties first. Clear runs before set,
#   so an explicit format flag wins. The same fields work per-entry under replace --batch.
#   A TAB matches as one character (pattern "City State Zip" fills a tab-separated line), and
#   TAB/NBSP/bullet-glyph variants match their plain equivalents unless --exact.
#   --at pN (or a cell paragraph tT:rRcC:pN) CONFINES the replace to one paragraph — use it when the
#   SAME placeholder repeats across the doc (a résumé's "City, State" in every entry) and you want THE
#   one in a specific paragraph, instead of find → edit --at pN:S-E span surgery. Batch entries take "at" too.
#   MULTI-LINE (editor-style): "\n" in PATTERN matches a line break OR a paragraph boundary
#   (consecutive paragraphs in the body or one table cell; never across a table/section break/cell wall); the REPLACEMENT's own
#   "\n"s then define the result — single-line replacement MERGES the paragraphs (first one's
#   formatting governs), "\n" in the replacement inserts a paragraph mark (splits). Untracked
#   only: refuses under tracking rather than skip the journal. Block ids shift; re-read after.

# Query batch — many independent searches against ONE immutable read.
docx find    FILE --batch queries.jsonl      # { query, regex?, ignoreCase?, nth?, … } OR { highlight|color|bold|italic|underline }; --json preserves per-query results

# Mutation batch — apply many changes from ONE read. Keys on each JSONL line
# mirror the command's flags; all locators address the doc as read. Every batch
# surface accepts --batch - to read JSONL from stdin.
docx edit    FILE --batch fills.jsonl       # { at, <one of: text|clear|markdown|runs>, style?, … }
docx insert  FILE --batch additions.jsonl   # { at|after|before, <content>, style?, color?, … }
docx replace FILE --batch script.jsonl      # { pattern, replacement, at?, regex?, all?, limit?, clear?, bold?, color?, … } applied in order
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
#   edit --text ""      REMOVES the line (same as `delete`; a table cell's last paragraph is blanked, not
#       deleted, so the cell stays valid). In --batch, `{"at":"pN","text":""}` or `{"at":"pN","delete":true}`
#       removes a line — so a form-fill is ONE sweep: fill the cells with values, drop the leftover
#       placeholder lines. Use `--runs '[]'` to blank a paragraph but keep an empty spacer. (Empty
#       `--text` can't ride along with --clear/run-formatting/--style/--alignment/--tabs — those exit
#       with a USAGE error; use `--runs '[]'` to keep a formatted empty spacer instead.) A SPAN's
#       `--text ""` (pN:S-E) still deletes just those characters.
#   --runs '[{"type":"text","text":"X","bold":true}]'
#   --text-file PATH                               # (insert/create) LITERAL multi-paragraph text, NOT parsed — every char verbatim,
#       each newline = a new paragraph. For prose GFM would corrupt: "3. note" stays "3.", *x* / [t](u) / bare URLs / {++x++} untouched.
#   "MARKDOWN" (positional, edit/insert) | --markdown "..." | --markdown-file PATH
#       GFM + math + CriticMarkup + inline HTML formatting → blocks. Markdown is the default:
#       `docx edit doc.docx --at p3 "Net **30** days"`, `docx insert doc.docx --after p7 "- another bullet"`.
#       New content looks like the document: a "- item" beside a list joins THAT list (numbering,
#       indent, font); an edit keeps the font that was at that position (a "# heading" over a
#       Georgia line stays Georgia); otherwise it takes the document's dominant font/size (the
#       `docx:base` read shows), not the style chain's default.
#   --code "..." | --code-file PATH [--language LANG]
#   --equation "x^2 + y^2" [--display]   (insert; edit also accepts --inline)
#   --clear bold,italic,highlight,color,size,font,…|all   (edit; strip run formatting, keep text)
#   --bold --italic --underline --strike --color HEX --highlight NAME --shade HEX --font NAME --size PT
#       --caps --smallcaps --superscript --subscript   (edit; SET run formatting on EXISTING text —
#       the inverse of --clear. Alone they format a span/paragraph/range in place; with --text they
#       fill AND format. Like --clear, applied directly — not recorded as a tracked change.)
#   --font-east-asia NAME   sets the East-Asian-script font (<w:rFonts w:eastAsia>) — separate from
#       --font because Word resolves w:eastAsia independently of w:ascii/w:hAnsi; --font alone has no
#       visible effect on Chinese/Japanese/Korean text. Supported by edit, edit --batch,
#       replace, replace --batch, styles set, and styles create.
#   --font-complex-script NAME   sets w:cs independently (e.g. Amiri for Arabic). The legacy
#       --font still sets ASCII, high-ANSI, and complex-script slots; this explicit override wins
#       for complex-script text when both flags are supplied, regardless of flag order.
#       Supports the same commands as --font-east-asia; each clears only its matching theme reference.
#       AST/--runs use fontEastAsia and fontComplexScript; Markdown preserves them as span data attributes.
#       NOTE: in a single no-content call (or one --batch entry) these run-format SET flags and the
#       paragraph properties (--style/--alignment/--space-*/--line-spacing/--indent-*/--first-line/
#       --hanging/--tabs) can't ride together — use separate calls/entries, or add --text to set both.
#   --list bullet|ordered [--list-level N]   (insert; for a task-list CHECKBOX use `docx tasks add`)
#   --rows N --cols N [--widths "A,B,C"] [--table-width V] [--borders S] [--layout L]   (docx tables create)
#   --image SRC [--alt T] [--width IN] [--height IN] [--caption "Figure 1: …"]   (docx images add; SRC = path, data: URI, or http(s) URL; --caption adds a Word "Caption"-styled line under the figure)
#   --page-break | --column-break | --section [--columns N] [--type T]   (insert)

docx comments add     FILE --at LOCATOR --text "..." [--author NAME] [--current | --baseline]
docx comments add     FILE --anchor "phrase" --text "..." [--occurrence N]
docx comments add     FILE --batch reviews.jsonl                    # JSONL: { at | anchor (+occurrence), text, author? }
docx comments reply   FILE --at cN --text "..."
docx comments reply   FILE --batch replies.jsonl   # {"at":"c0","text":"...","resolve":true} per line — answer (and close) every thread in one call
docx comments resolve FILE --at cN [--at cM ...] [--unset] | --batch resolutions.jsonl
docx comments delete  FILE --at cN [--at cM ...]          | --batch removals.jsonl

docx footnotes add    FILE --at pN[:offset] (--text "..." | --runs JSON | --markdown TEXT)
docx footnotes edit   FILE --at fnN (--text "..." | --runs JSON | --markdown TEXT)
docx footnotes delete FILE --at fnN
docx endnotes  add    FILE --at pN[:offset] (--text "..." | --runs JSON | --markdown TEXT)
docx endnotes  edit   FILE --at enN (--text "..." | --runs JSON | --markdown TEXT)
docx endnotes  delete FILE --at enN

# Headers & footers (one shared impl — "marginals"). Default placement is every
# page, all sections (--at sN targets one). Content: ONE primary source, except
# --text + one field = two-zone (text left, field right at a content-edge tab).
docx headers set   FILE [--at sN] [--type default|first|even | --first-page | --even | --odd] \
                        [--text "..."] [--align left|center|right] \
                        [--page-number [--of-pages] | --date [--date-format FMT] | --style-ref STYLE | --field filename|title|author] \
                        [--track] [--author NAME]
docx headers clear FILE [--at sN] [--type T | --first-page | --even | --odd]
docx footers set   FILE …   # identical flags, kind=footer (e.g. --page-number --of-pages → "Page X of Y")
docx footers clear FILE …

docx images add     FILE (--after | --before) LOCATOR --image SRC [--alt T] [--width IN] [--height IN] [--caption "..."]
docx images add     FILE (--at-start | --at-end) --image SRC [options]   # SRC = path, data: URI, or http(s) URL
docx images extract FILE --to DIR [--at imgN]            # --to = output directory; --at picks one image
docx images replace FILE --at imgN --with ./new.png
docx images delete  FILE --at imgN

docx hyperlinks add     FILE --at pN:S-E --url URL
docx hyperlinks replace FILE --at linkN --with URL
docx hyperlinks delete  FILE --at linkN

docx tables create        FILE (--after|--before LOCATOR | --at-start|--at-end) --rows N --cols M [--widths "A,B,C"] [--table-width V] [--borders S] [--layout L]
docx tables insert-row    FILE --at tN [--position INDEX] [--cells "a,b,c"]
docx tables delete-row    FILE --at tN:rR
docx tables insert-column FILE --at tN [--position INDEX] [--width TWIPS]
docx tables delete-column FILE --at tN:cC
docx tables set-widths    FILE --at tN --widths "25%,25%,50%" | "1440,..." | auto
docx tables merge         FILE --at tN:rR1cC1-rR2cC2
docx tables unmerge       FILE --at tN:rRcC
docx tables borders       FILE --at tN [--style single|double|none] [--size N] [--color HEX]
docx tables format        FILE --at LOCATOR [--shade HEX|NAME] [--valign top|center|bottom]
                          [--halign left|center|right|justify] [--cell-borders SIDES]
                          [--cell-margins M|T,R,B,L|none]
                          [--align left|center|right] [--style ID] [--row-height M] [--repeat-header]
                          # --cell-margins with --at tN sets the TABLE default cell margins
                          # (<w:tblCellMar>); a row/column/cell/range sets those cells' own.

docx lists set            FILE --at pN [--start N] [--format FMT] [--restart] [--continue]
                          # renumber a NUMBERED list (--at = any item; applies to the whole list).
                          # FMT = decimal | lower-alpha | upper-alpha | lower-roman | upper-roman.
                          # --restart splits a fresh list off here; --continue picks up the previous
                          # list's numbering instead of restarting. Untracked (Word records no revision).

docx tasks add            FILE (--after | --before) LOCATOR (--text LABEL | --runs JSON) [--checked | --unchecked] [--list-level N]
docx tasks add            FILE (--at-start | --at-end) (--text LABEL | --runs JSON) [--checked | --unchecked]
                          # insert a GFM task-list checkbox item (☐ default, ☒ with --checked). Consecutive
                          # adds after a list anchor build ONE contiguous list.
docx tasks check          FILE --at pN            # mark an existing task done (☒), in place
docx tasks uncheck        FILE --at pN            # clear an existing task (☐), in place
                          # check/uncheck take --track to redline the toggle even with the doc toggle off
                          # (surfaces as a checkboxToggle in `track-changes list`).

docx track-changes on|off FILE
docx track-changes list   FILE [--json]
docx track-changes accept FILE (--at tcN [--at tcM ...] | --at revN | --all)
docx track-changes reject FILE (--at tcN [--at tcM ...] | --at revN | --all)
docx track-changes apply  FILE [--accept H ...] [--reject H ...]
# `list` defaults to a text table, one LOGICAL change per line (revN collapses a del+ins
# pair onto one line); `--json` for the raw array. A del+ins REPLACE pair shares a
# "group": "revN"; `--at revN` accepts/rejects both halves in one call.
# To FINALIZE a review (accept some, reject the rest), use `apply` — it takes both decision
# lists in ONE call, resolved against the original ids, so nothing renumbers mid-operation
# and the file is never left half-finalized. Doing it as separate accept then reject calls
# renumbers the ids between them. After a subset accept/reject/apply, the confirmation
# re-lists what remains with its renumbered handles.
```

#### Raw OOXML escape hatch (last resort)

```sh
docx raw get FILE --at LOCATOR                     # a block/section/relationship's exact XML
docx raw edit FILE --at LOCATOR --find S --with S  # patch it in place, one gated call
docx raw replace FILE --at LOCATOR --xml '…'       # swap it wholesale (also --batch)
docx raw insert FILE --after pN --xml '…'          # splice new block XML (or a <Relationship/>)
docx raw part list|get|add|replace|edit FILE …     # OPC parts, addressed by --name
docx validate FILE [--json]                        # schema-check against ECMA-376 (transitional)
```

For constructs no modeled verb covers — drop caps, TOC field codes, embedded
objects. Every mutation runs a gate pipeline (well-formedness, addressable
roots, ECMA-376 child order, reference/id integrity, and a schema check
against the bundled transitional XSDs — a change may not ADD errors) and
**nothing is written when any gate fails**. Inline `--xml`/`--find`/`--with`
are verbatim; raw changes are never tracked. Details and examples:
`docx raw --help`.

> **One rule to memorize: addressing an existing thing is always `--at`.**
> `comments reply/resolve/delete`, `footnotes/endnotes edit/delete`, `images extract/replace/delete`, `hyperlinks replace/delete`, `tables *`, `track-changes accept/reject`, `edit`, `delete`, and `insert` all take `--at LOCATOR`. For insert, `--at` means after an ordinary block and into a bare cell; use `--before`/`--after` when the side matters, or `--at-start`/`--at-end` for document boundaries. The remaining exceptions are positional by nature: `read` slices with `--from`/`--to LOCATOR`; `wc` takes a positional `[LOCATOR]`; `find`/`replace` take a positional `QUERY`/`PATTERN` (and `replace` accepts an optional `--at pN` to _confine_ the substitution to one paragraph). `images extract --to DIR` is an _output directory_, not a locator.

## Output contract

The CLI is built for non-interactive agents. **Exit code is the success signal**, output is data:

| Exit | Meaning                   | Error codes                                                                                                                                                         |
| ---- | ------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `0`  | success                   | —                                                                                                                                                                   |
| `2`  | usage / bad locator       | `USAGE`, `INVALID_LOCATOR`, `INVALID_XML`                                                                                                                           |
| `3`  | addressed thing not found | `FILE_NOT_FOUND`, `PART_NOT_FOUND`, `BLOCK_NOT_FOUND`, `COMMENT_NOT_FOUND`, `IMAGE_NOT_FOUND`, `HYPERLINK_NOT_FOUND`, `TRACKED_CHANGE_NOT_FOUND`, `MATCH_NOT_FOUND` |
| `1`  | general failure           | `NOT_A_ZIP`, `TRACKED_CHANGE_CONFLICT`, `TABLE_STRUCTURE`, `VALIDATION_FAILED`, `IMAGE_SOURCE`, `RENDER_ENGINE`, `RENDER_FAILED`, `UNHANDLED`                       |

**Errors** print `{code, error, hint?}` JSON to stdout with a nonzero exit — note there is **no `ok` field**; the exit code plus `code` are the unambiguous signal.

**The `ok` field appears in exactly one place: the `--verbose` success ack** (`{ok:true, operation, path, …}`). Without `-v`, success output is shaped for the next command:

| Command class                                                                                                                                                                                                                                                              | Default stdout on success                                                                                                                                                                   | `--verbose`                                                               |
| -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------- |
| **Mutator that mints a new handle** — `comments add`→`cN`, `comments reply`→`cN`, `footnotes/endnotes add`→`fnN`/`enN`, `hyperlinks add`→`linkN`, `insert`→the new `pN`                                                                                                    | the bare locator(s), **one per line** (a multi-block `--markdown` insert prints several)                                                                                                    | full `{ok:true,…}` ack                                                    |
| **Mutator with no new handle** — `edit`, `delete`, `replace`, `create`, `comments resolve/delete`, `images replace/delete`, `hyperlinks replace/delete`, `footnotes/endnotes edit/delete`, `headers/footers set/clear`, `tables *`, `track-changes accept/reject` & toggle | **one-line confirmation** — `<operation> <target>` (e.g. `edit t1:r0c1:p0`, `edit 7 changes`, `replace 3 occurrences replaced`) (exit `0`)                                                  | full `{ok:true,…}` ack                                                    |
| `find`                                                                                                                                                                                                                                                                     | matched span locators, one per line (no matches → nothing, exit `0`)                                                                                                                        | `--json` → `{ totalMatches, query, view, matches:[…], normalizedQuery? }` |
| `wc`                                                                                                                                                                                                                                                                       | the bare count (whole-doc adds a tab-separated `sections` column, like `wc`)                                                                                                                | `--json` → `{ words, scope, view, sections? }`                            |
| `outline`                                                                                                                                                                                                                                                                  | indented `LOCATOR⇥TEXT` tree (two spaces per level)                                                                                                                                         | `--json` → nested `[{ id, locator, level, style, text, children }]`       |
| `read`                                                                                                                                                                                                                                                                     | GFM Markdown; each paragraph carries its `pN` locator once, while a plain empty table cell carries one bare `<!-- tN:rRcC -->` fill handle | `--ast` → the JSON AST body (`docx info schema`)                          |
| `render`                                                                                                                                                                                                                                                                   | image paths, one per line                                                                                                                                                                   | `--verbose` → `{ok, operation, path, engine, output, pages}`              |
| `* list` (all eight `list` verbs)                                                                                                                                                                                                                                          | a **bare JSON array**; each item's `id` is its `--at` handle (`track-changes list` prints a text table; add `--json` for the array)                                                                                                                               | —                                                                         |

`--dry-run` always prints a preview object (no `ok`) and writes nothing; it wins over `-o/--output`.

## Discovering ids

Locators come in two flavors. **Positional block ids** (`pN`, `tN`, `sN`) are derived from document order and **shift after structural edits** — re-read between non-trivial mutations. **Entity ids** (`cN`, `imgN`, `linkN`, `fnN`, `enN`, `tcN`, `eqN`, `hdrN`, `ftrN`) are surfaced by a `list` verb (or `read --ast`) and are what you pass to `--at`:

| Id                              | Discover with                                                                                                                            | Used by                                                                       |
| ------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------- |
| `pN` / `tN` / `sN` (block ids)  | `docx read FILE` (the `<!-- pN -->` trailers), `docx read FILE --ast`, `docx outline FILE` (heading `pN`s), or `docx render` page images | `read`, `edit`, `insert`, `delete`, `wc`, `find` results                      |
| `cN` (comment)                  | `docx comments list FILE`                                                                                                                | `comments reply/resolve/delete --at`                                          |
| `fnN` / `enN` (foot/endnote)    | `docx footnotes list FILE` / `docx endnotes list FILE`                                                                                   | `footnotes/endnotes edit/delete --at`                                         |
| `hdrN` / `ftrN` (header/footer) | `docx headers list FILE` / `docx footers list FILE` (or `read --ast` / the `docx:header`/`docx:footer` note)                              | `raw get --at`, `headers/footers set/clear --at` (or scope by section: `--at sN --type T`) |
| `imgN` (image)                  | `docx images list FILE`                                                                                                                  | `images extract/replace/delete --at`                                          |
| `linkN` (hyperlink)             | `docx hyperlinks list FILE`                                                                                                              | `hyperlinks replace/delete --at`                                              |
| `tcN` (tracked change)          | `docx track-changes list FILE`                                                                                                           | `track-changes accept/reject --at`                                            |
| `eqN` (equation)                | `docx read FILE --ast` (run `latex` field)                                                                                               | `equations edit --at eqN`                                                     |
| `tbxN` (text box)               | `docx read FILE` (the `<!-- docx:textbox tbxN anchor="pM" -->` block; its paragraphs carry `<!-- tbxN:pK -->`) or `read --ast`          | `tbxN:pK` on `edit`/`insert`/`delete`/`replace --at`/`wc`; `find`/`replace`/`wc` cover text boxes by default (not `comments add`/`footnotes add` — Word discards those inside a box) |

Each `list` verb prints a bare JSON array (`track-changes list` needs `--json`; its default is a one-line-per-change table) where every item's `id` is exactly the handle you feed back to `--at` — pipe through `jq` to filter (`docx comments list doc.docx | jq '.[] | select(.author=="Jane")'`).

## Locators

`docx info locators` (`--json` for machine-readable) is the canonical reference. The grammar in brief:

```
pN              paragraph N                  pN:S-E          chars S..E within paragraph N
pN-pM           whole-paragraph range        pN:S-pM:E       cross-paragraph character range
sN              section break N              tN              table N
tN:rRcC         cell — direct fill target    tN:rRcC:pK      paragraph K of that cell (chainable)
tN:rR / tN:cC   table row R / column C       tN:rR1cC1-rR2cC2  rectangular cell region (merge)
tbxN            text box N (its whole story) tbxN:pK         paragraph K inside that text box (chainable)
cN  imgN  linkN  fnN  enN  tcN  eqN          entity ids (comment / image / hyperlink /
                                             footnote / endnote / tracked-change / equation)
```

**Offset semantics: character offsets are 0-based, start-inclusive, end-exclusive** — `p3:5-20` is the 15 characters at indices 5..19 of paragraph 3. Offsets count the _visible_ text of the paragraph in the selected view (accepted by default).

**Nested tables chain the same syntax** arbitrarily deep — `t0:r2c1:t0:r0c0:p0` is the first paragraph of the (0,0) cell of the first table nested inside the (2,1) cell of the document's first table.

**Text boxes are their own story, like a cell.** A Word text box (a "Confidential" stamp, a pull quote, a letterhead block) floats over the paragraph that anchors it; its paragraphs are `tbxN:pK`, numbered in document order (`tbx0:p1:15-19` for a span, `tbx0:t0:r0c0:p0` for a table inside the box). `read` prints the story right after its anchor block, bracketed by `<!-- docx:textbox tbx0 anchor="p1" … -->` / `<!-- docx:textbox-end tbx0 -->`, and `find`/`replace --all`/`wc` include it — a human sees that text on the page, so the tool counts what `read` shows (`wc FILE tbx0` counts one box). Editing the anchor paragraph's text, or tracked-deleting it, keeps the box (a tracked delete hides it in the accepted view and `reject` brings it back). Two things Word itself refuses inside a text box and silently discards on its next save — comments and footnotes/endnotes — are refused up front with a pointer to the anchor paragraph; tracked changes and hyperlinks inside a box survive Word. Word writes every box twice (a modern shape plus a legacy fallback copy); edits land in the copy Word reads and the fallback is re-synced on save, which is what Word does when it edits a box itself (an untouched box's fallback it leaves verbatim).

**Bare cell content targets are conservative.** `edit --at CELL` requires exactly one direct paragraph. `insert --at CELL` fills/reuses a normal blank cell and otherwise appends; `--before CELL`/`--after CELL` use its start/end boundary. Merged or grid-shifted bare cells are rejected — use `CELL:pK` for precise paragraph targeting.

**Not every command accepts every form** — each command's `--at`/`--from`/positional help lists exactly what it takes. The shapes:

| Form                                                                 | Accepted by                                                                         |
| -------------------------------------------------------------------- | ----------------------------------------------------------------------------------- |
| `pN`, `tN`, `sN`, `tN:rRcC:pK`, `tbxN:pK` (blocks)                   | `read --from/--to` (top-level only), `insert --at/--after/--before`, `wc`, `comments add` |
| `tN:rRcC` (bare simple cell)                                          | `insert --at/--after/--before`; `edit --at` when the cell has one direct paragraph  |
| `pN`, `pN:S-E`, `pN-pM`, `tN:rRcC:pK`, `tN:rRcC:pK:S-E` | `edit --at` (span/cell forms strip or replace just that range; sections → `docx sections`, equations → `docx equations edit`) |
| `pN`, `pN-pM`, `tN`, `sN`, `tN:rRcC:pK`                              | `delete --at`                                                                       |
| `pN:S-E`, `pN:S-pM:E`, `tN:rRcC:pK:S-E` (spans)                      | `comments add --at`, `hyperlinks add --at` (single paragraph), `find`/`wc` results  |
| `pN[:offset]` (point)                                                | `footnotes/endnotes add --at`                                                       |
| `cN` / `fnN` / `enN` / `imgN` / `linkN` / `tcN` (entities)           | the matching noun's `--at` (the `c`/`fn`/`en`/`img`/`link`/`tc` prefix is optional) |
| `tN`, `tN:rR`, `tN:cC`, `tN:rRcC`, `tN:rR1cC1-rR2cC2`                | the `tables` verbs                                                                  |

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

A source starting with `-` (a bullet) works inline too — `docx insert doc.docx --after p3 "- item"`; `--markdown-file` is for long or multi-block sources.

**track-changes review loop.** Toggle tracking on, make edits (they auto-emit `<w:ins>`/`<w:del>`), then inventory and resolve:

```sh
docx track-changes on doc.docx
docx replace doc.docx "old phrasing" "new phrasing" --all
docx track-changes list doc.docx                  # → one line per logical change (revN); --json for { id:tcN, kind, author, text, … }
docx read doc.docx --current                       # → CriticMarkup {++ins++}[^tcN] / {--del--}[^tcN]
docx track-changes accept doc.docx --at rev0 --at rev2  # or --all
docx track-changes apply doc.docx --accept rev0 --reject rev1   # finalize: accept some, reject the rest, in ONE call
```

`read` has three tracked-change views: default **`--accepted`** renders clean text — drops subtractive edits and inlines additive ones (the post-accept document); **`--current`** shows CriticMarkup with `[^tcN]` footnotes; **`--baseline`** does the reverse of accepted (the pre-change document). `find`, `replace`, `wc`, and `comments add` honor the same `--accepted`/`--baseline`/`--current` flags so offsets stay consistent across commands. Add `--comments` to `read` to append `[^cN]` footnotes for comment spans.

## How It Works

**In-place XML mutation.** The AST returned by `read` is a _view_ over the parsed XML tree, not a separate model. When you `edit` or `comments add`, the CLI mutates the underlying XML nodes directly and serializes back. Anything not modeled in the AST (custom styles, theme colors, schema extensions) survives because untouched regions are never re-emitted. Never delete a relationship something still references — that corrupts the file — so part/relationship pruning is gated on a reference scan; unreferenced orphans are left in place.

**JSX for emitters.** Constructing OOXML fragments imperatively (`<w:rPr>` → `<w:b/>` → `<w:color w:val="800080"/>`) is verbose, so fresh XML is authored in JSX with a custom factory: `<w.rPr><w.b/><w.color w-val="800080"/></w.rPr>` becomes the right `XmlNode` tree.

**Span-aware comments & hyperlinks.** `comments add --at p3:5-20` (and `hyperlinks add`) find the runs containing offsets 5 and 20, split them at the boundaries (preserving `<w:rPr>` on both halves), and insert markers between the slices. Comments authored by older tools that lack `w14:paraId` (required by `commentsExtended.xml`) get a fresh paraId injected automatically on resolve/reply.

**Tracked changes.** With `<w:trackRevisions/>` set, `insert`/`edit`/`delete`/`replace` emit native `<w:ins>`/`<w:del>` (attributed via `--author`, `$DOCX_AUTHOR`, or `Reviewer`); pass `--track` to one of those commands (or the `tables` verbs / `images delete`) to track just that invocation even when the doc toggle is off. `edit --at pN --text` runs a word-level diff so unchanged words keep their formatting and only changed words are wrapped — the same shape Word produces mid-tracking. `accept`/`reject` handle run-level ins/del/moveFrom/moveTo, `sectPrChange`, paragraph-mark ins/del, and the table-structural revisions (rowIns/rowDel, cellIns/cellDel, tblGridChange, tcPrChange). OOXML has no tracked-change construct for hyperlink edits or image swaps, so under tracking those emit a `[docx-cli]` audit comment instead of a fake revision (image _deletion_ is honest removal — it wraps a real `<w:del>`).

**Text boxes and markup compatibility.** Word writes every modern shape and text box as an `<mc:AlternateContent>` pair (a `wps` shape under `<mc:Choice>`, a legacy VML copy under `<mc:Fallback>`), and the markup-compatibility spec lets that wrapper sit at any level — around a block, a paragraph's runs, or a run's children — and nest. The reader resolves the wrapper wherever it appears (first `Choice`, else `Fallback`), so nothing under it is dropped, and every offset walker descends the same branch (a wrapper with several `Choice`s counts exactly the one the reader read); a `<w:txbxContent>` story surfaces as a `textBox` run whose blocks are addressable `tbxN:pK` paragraphs. Edits land in the `Choice` copy (the one Word reads) and a text box's `Fallback` twin is mirrored from it on save, exactly as Word does when it edits the box. A wrapper that is not a text box has no mirrorable story, so an edit inside it changes the `Choice` only — the same two-faced file Word leaves when it preserves a fallback it didn't regenerate.

**Rich content.** Images insert from a path, `data:` URI, or `http(s)` URL (bounded fetch; HEIC→JPEG transcode; SVG sanitized; non-public/metadata addresses refused at every redirect hop). Equations round-trip OOXML `<m:oMath>` ↔ LaTeX (reconstructed, not legacy plaintext) — authored via temml (LaTeX→MathML) plus an in-house MathML→OMML adapter (the math path pulls in no LGPL code). Code blocks emit one `CodeBlock`-styled paragraph per line with optional lowlight syntax highlighting (37 bundled languages); they collapse back to a GFM fenced block on read. GFM task lists round-trip Word's checkbox content control (and the Word-for-Web Wingdings-glyph variant), surfacing as `taskState` in the AST. Tables operate on a merge-aware logical grid so `gridSpan`/`vMerge` cells map onto physical `<w:tc>`, and structural edits refuse to bisect an existing merge.

**Markdown dialect.** `create --from`, `insert/edit --markdown`, and the note bodies all parse the same GFM + math + CriticMarkup + inline-HTML-formatting dialect (remark + remark-gfm + remark-math + an in-house inline-surgery transform), composing the existing OOXML emitters. `read` emits a compatible dialect, so the read → edit → write loop round-trips (lossless for paragraphs, lists, and nested blockquotes; code/tables/math/headings inside a blockquote intentionally escape to top level on import). `read --ast` is the fully lossless JSON form.

**Literal text — the parser-free channel.** When you want prose inserted _exactly_ — reviewer notes, quoted excerpts, anything where Markdown would misfire — use `create --text-file PATH` / `insert --text-file PATH` (or `-` for stdin). Every character lands verbatim and each newline starts a new paragraph; nothing is interpreted, so `3. note` stays `3.` (no ordered-list renumber), and `*x*`, `[t](u)`, bare URLs, and `{++x++}` are kept as written. This exists because GFM corruption isn't always escapable: bare URLs autolink with no escape sequence at all, and CriticMarkup eats `{++…++}` regardless of backslashes — so a literal path is the only safe way to author untouched prose.

**Document-wide font.** `docx styles set-default-font FILE "Times New Roman"` sets the font in the two places a font actually lives — `word/styles.xml` `<w:docDefaults>` (the formal default) _and_ the theme font scheme (`word/theme/theme1.xml`, major + minor `<a:latin>`), since real Word docs resolve their fonts _through_ the theme and touching only one silently loses to the other. Body text and theme-following headings both adopt it; styles or runs that pin their own font (a code block's monospace, a deliberately-Arial run) are preserved and named in the ack, with `--all` to repoint even those. `--size N` sets the default size on the same write.

**Edit & create styles.** `docx styles set FILE --at Heading1 --color 1F4E79 --size 16 --bold` rewrites the _style definition_ in `word/styles.xml`, so every paragraph or run that uses the style picks up the change at once ("make all Heading 1s green") — the same run-/paragraph-formatting flags as `edit` (color/font/size/highlight/underline/caps + alignment/spacing/indentation for paragraph styles). `docx styles create FILE Callout --color C00000 --bold` mints a new paragraph or character style that `insert/edit --style Callout` can then apply. Editing an un-materialized built-in (`--at Heading3` on a doc that never used it) provisions it first; a paragraph with its own _direct_ formatting keeps it (the override wins — the definition edit never touches the body). Style edits are **not** tracked changes even under track-changes — matching Word, which applies style-definition edits to `styles.xml` directly with no redline.

**List numbering.** `docx lists set FILE --at p12 --start 5` makes a numbered list begin at 5; `--format upper-roman` (or `lower-alpha`/`upper-alpha`/`lower-roman`) switches the glyph; `--restart` splits a fresh list off at that item; `--continue` makes a list pick up the previous list's numbering instead of restarting. `--at` names any item — the change applies to the whole list. The start round-trips through the markdown ordinal (the body reads `5. 6. 7.`); the glyph and any continue link, which GFM can't express, surface as a deviation-only `<!-- docx:list p12 format="upper-roman" -->` / `<!-- docx:list p20 continues -->` hint (dropped on import, like every `docx:` note — `read --ast` carries `list.start`/`list.format` losslessly). Untracked, matching Word, which records no revision for a list-numbering change.

**Run formatting beyond bold/italic.** Properties markdown has no native syntax for — text color, theme color, highlight, shading, underline (all 18 styles + color), super/subscript, small/all caps, font, and size — are emitted as the **HTML a markdown reader actually renders**, so the output looks right in GitHub, VS Code, Obsidian, and browsers (Pandoc `[text]{…}` spans render as literal brackets in all of those). `read` emits semantic tags where they exist — `<mark>overdue</mark>`, `<sup>x</sup>`, `<sub>2</sub>` — a `<span style="color:#C00000">…</span>` for the CSS-expressible properties, and `data-*` attributes for the OOXML-only ones CSS can't express (theme colors, underline styles); `insert/edit --markdown` parses them back losslessly, and a leading `<!-- docx:base font="Arial" size="8pt" -->` note declares the document's dominant font/size once so the body isn't buried in per-run repetition. Bold/italic/strike/code/links stay native (`**`/`*`/`~~`/`` ` ``/`[](…)`). Because the inline-surgery transform scans whole sibling sequences, a CriticMarkup marker or span can straddle other formatting — `{++**bold insertion**++}` is tracked correctly. An invalid enum value (e.g. a bogus highlight name) fails with a clear error rather than silently vanishing. Inserted plain content inherits the surrounding paragraph's font/size so it blends in; a markdown bullet beside a list joins that list; and new headings/list items fall back to the `docx:base` font/size rather than their style's.

**Visual verification.** `docx render` is the only command that needs an external application installed: it drives Microsoft Word (macOS via `osascript`, Windows via PowerShell COM — the ground-truth renderer) or LibreOffice (`soffice`, cross-platform) to produce a PDF, then rasterizes in-process via the bundled `@hyzyla/pdfium` WASM package — no poppler/pdftoppm/ImageMagick needed. Agents that consume PNGs use this to verify edits, diff accept/reject before-vs-after, or generate screenshots.
