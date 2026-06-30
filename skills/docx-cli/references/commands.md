# docx-cli command reference

The **authoritative, versioned** reference is `docx <command> --help` (and
`docx --help` for the index). This file is a quick map; when a detail matters,
run the help. None of the `info` commands need a FILE:

```sh
docx --help            # every command, one capability hint each
docx info locators     # the addressing grammar (the backbone)
docx info schema       # the JSON-AST shape that `docx read --ast` emits
docx info skill        # this skill, regenerated from the binary
```

## Read & query (never mutate)

| Command             | What it does                                                                                                                                                                |
| ------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `read FILE`         | Render as Markdown with `pN` locators. `--from/--to` slice; `--accepted` (default)/`--current`/`--baseline` tracked views; `--comments`; `--ast` for the lossless JSON AST. |
| `find FILE [QUERY]` | Find spans by text OR by formatting (`--highlight/--color/--bold/--italic/--underline`); returns locators to feed `--at`.                                                   |
| `wc FILE [LOCATOR]` | Word count for the whole doc or a slice.                                                                                                                                    |
| `outline FILE`      | Headings as a locator tree.                                                                                                                                                 |
| `render FILE`       | Render each page to PNG/JPG via Word or LibreOffice — for verifying LAYOUT only.                                                                                            |
| `styles FILE`       | List/describe styles (`--used`, `--at ID`).                                                                                                                                 |
| `info <topic>`      | `schema` / `locators` / `skill` — reference material, no FILE.                                                                                                              |

## Mutate (overwrite FILE in place; `-o PATH` writes a copy; `--dry-run` previews)

| Command                              | What it does                                                                                                                      |
| ------------------------------------ | --------------------------------------------------------------------------------------------------------------------------------- |
| `create FILE`                        | Create a new `.docx` (`--from PATH.md` / `--from -`; `--text-file` for literal text).                                             |
| `edit FILE`                          | Replace/strip text or formatting at a locator (`--clear`, `--track`, `--batch`).                                                  |
| `insert FILE`                        | Insert a paragraph, image, table, equation, code, markdown, or page break (`--after/--before`, `--track`, `--batch`).             |
| `delete FILE`                        | Remove a paragraph, range, table, or section break (`--at`, `--track`, `--batch`).                                                |
| `replace FILE PATTERN REPL`          | Substitute text spans sed-style — KEEPS the run's formatting and tabs (`--regex`, `--track`, `--batch`). The form-fill workhorse. |
| `sections FILE`                      | Multi-column layout, section breaks, and page setup (margins/orientation/size). The only way to do columns.                       |
| `styles set/create/set-default-font` | Restyle every heading at once, mint a style, or set the document font.                                                            |
| `comments`                           | `add` (`--at` / `--anchor PHRASE` / `--batch`), `reply`, `resolve` (`--unset`), `delete`, `list`.                                 |
| `footnotes` / `endnotes`             | `add` (`--at` / `--anchor`), `edit`, `delete`, `list`.                                                                            |
| `headers` / `footers`                | `set` / `list` / `clear` — text, page numbers, dates, fields; e.g. `footers set FILE --page-number --of-pages`.                   |
| `images`                             | `add` (`--caption` for a figure), `extract`, `replace`, `delete`, `list`.                                                         |
| `hyperlinks`                         | `add` (`--url`), `list`, `replace` (`--with`), `delete`.                                                                          |
| `tables`                             | Insert/delete rows & columns, merge/unmerge, set widths, borders, formatting.                                                     |
| `lists`                              | Renumber a numbered list (`set --at pN --start 5` / `--format upper-roman` / `--restart` / `--continue`).                         |
| `track-changes`                      | `on`/`off`, `list`, `accept`/`reject` (`--at tcN` / `--all`).                                                                     |

## The two gotchas that can trip you up

1. **Ids shift after structural edits.** Re-read between mutations, or apply many
   changes from ONE read with `--batch FILE.jsonl` — every locator addresses the
   document as read, so ids stay valid across the whole batch.
2. **`docx read` is the source of truth for CONTENT.** Only `docx render` (slow —
   spins up Word) shows LAYOUT: columns, page/section breaks, image placement,
   table geometry. Render once at the end, not after every edit.
