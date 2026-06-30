---
name: docx-cli
description: "Read, edit, redline, comment on, and create Microsoft Word .docx files. Use to fill out or edit a Word doc, redline a contract with tracked changes, add/resolve comments, replace text keeping its formatting, restyle headings/fonts, edit tables, or read/extract a .docx as Markdown or text. Also BUILD a new .docx — from Markdown or programmatically (code that outputs a Word report with headings, tables, images). Not for PDF, Google Docs, Excel, PowerPoint, or .doc."
---

# docx-cli

`docx` is a command-line tool for reading, editing, redlining, and commenting on
Microsoft Word `.docx` files. It edits the underlying OOXML **in place** (it never
rebuilds the document from a lossy view), addresses everything with **stable
locators**, and signals success through an **exit code** plus a one-line
confirmation — so even small, cheap models can drive it reliably.

## 0. Make sure the binary is on PATH

Run `docx --version`. If you get "command not found", install it:

```sh
curl -fsSL https://raw.githubusercontent.com/kklimuk/docx-cli/main/install.sh | sh
```

Or, from this skill folder, run `bash scripts/bootstrap.sh` — it checks the
installed version against the latest release and self-updates. Every verb works
against the `.docx` zip directly; only `docx render` needs Word (macOS/Windows)
or LibreOffice installed.

## 1. The contract is `--help` / `docx info` — start there

The help text is authoritative and versioned with the binary. This skill is thin
on purpose and defers to it. Before doing anything, run (none of these need a FILE):

```sh
docx --help              # every command + a one-line capability hint each
docx info locators       # the addressing grammar — READ THIS, it is the backbone
docx info schema         # the JSON-AST shape that "docx read --ast" emits
```

Then `docx <command> --help` for any verb before you use it.

## 2. Locators — how you address things

- `pN` paragraph, `tN` table, `sN` section; `p3:5-20` = characters 5..19 of `p3`;
  `pN-pM` a block range; `tN:rRcC` a table cell.
- Entities: `cN` comment, `imgN` image, `linkN` hyperlink, `fnN`/`enN`
  foot/endnote, `tcN` tracked change, `eqN` equation.
- Get them from `docx read FILE` (locators ride the Markdown as `<!-- pN -->`
  comments) or `docx read FILE --ast` (lossless JSON).
- **Ids are positional and SHIFT after structural edits.** Re-read between
  mutations — OR apply many changes from ONE read with `--batch` (below).
- Pass a locator with `--at` (edit / delete / comments / footnotes / images /
  hyperlinks / tables / track-changes), `--after`/`--before` (insert), or
  `--from`/`--to` (read a slice).
- Don't hand-count character offsets: `docx find FILE "phrase"` returns the exact
  span locator (e.g. `p3:5-20`) to paste into `--at`.

## 3. Golden workflows

### Fill out a form or contract (keeps formatting)
`docx replace` swaps only the text and preserves the run's bold/font and any tab
stops — so it fills bold, tabbed template lines without rebuilding runs.
```sh
docx read contract.docx                                  # see content + locators
docx replace contract.docx "[Client Name]" "Acme, Inc."  # one field
docx replace contract.docx --batch fills.jsonl           # many fields, one read/write
```

### Redline with tracked changes
```sh
docx track-changes on contract.docx       # turn tracking on (doc-level)
docx replace contract.docx "Net 90" "Net 30"   # now auto-emits <w:ins>/<w:del>
docx edit --at p12:0-40 contract.docx --text "…" --track   # or redline one edit
docx track-changes list contract.docx     # the tcN handles
docx read contract.docx --current         # view redlines as CriticMarkup
docx track-changes accept contract.docx --at tc3   # or --all / reject
```

### Comment on clauses
```sh
docx comments add contract.docx --anchor "limitation of liability" --text "Cap is too low."
docx comments list contract.docx
docx comments reply contract.docx --at c0 --text "Agreed, raising to \$5M."
docx comments resolve contract.docx --at c0
```

### Read / extract
```sh
docx read FILE            # Markdown (default; tracked changes shown accepted-clean)
docx read FILE --ast      # lossless JSON AST
docx wc FILE              # word count (whole doc or a slice)
docx outline FILE         # headings as a locator tree
```

### Build from scratch / verify layout
```sh
docx create out.docx --from draft.md      # GFM + math + CriticMarkup + inline HTML
docx render FILE --out pages/             # PNG per page — only when LAYOUT is the question
```

## 4. Apply many changes from one read — `--batch`

`edit`, `insert`, `replace`, `delete`, and the `comments` verbs take
`--batch FILE.jsonl` (one JSON change per line; `-` reads stdin). Every locator
in the batch addresses the document **as read**, so ids stay valid across the whole
batch — one read, one write, no re-reading between changes. Keys mirror the
command's flags. This is the right tool for filling a form or applying a review.

## 5. Output & safety contract

- **Exit code is the success signal:** `0` ok, `1` error, `2` usage, `3`
  not-found. Every command also prints a one-line text confirmation — you never
  have to re-read just to learn whether a mutation landed.
- Mutators overwrite `FILE` **in place** (git is your history). `-o/--output PATH`
  writes a copy instead; `--dry-run` previews without writing.
- A command that mints a new handle (`comments add`→`cN`, `insert`→`pN`,
  `footnotes add`→`fnN`, …) prints the bare locator(s), one per line.
- Re-read after structural edits (ids shift), or batch from one read.
- Need exact literal text in (a URL, prose GFM would mangle)? `insert` and
  `create` take `--text-file PATH` (`-` = stdin): every character lands verbatim,
  each newline a new paragraph. No escaping burden.

## 6. Going deeper

- `references/commands.md` — the full command surface at a glance.
- `references/troubleshooting.md` — install, PATH, render runtime, common errors.
- Or just run `docx <command> --help` — the authoritative, versioned contract.
