# docx-cli troubleshooting

## `docx: command not found`

The binary isn't on PATH. Install it (or run `scripts/bootstrap.sh`, which also
self-updates a stale install):

```sh
curl -fsSL https://raw.githubusercontent.com/kklimuk/docx-cli/main/install.sh | sh
```

It installs to `~/.local/bin/docx` by default. Make sure that directory is on
your `PATH` (`export PATH="$HOME/.local/bin:$PATH"`). Set `PREFIX=/usr/local/bin`
before the install to choose another location.

## `docx render` fails or hangs

`render` is the **only** command that needs an external app. Everything else
operates on the `.docx` zip directly. For `render` you need one of:

- **Word** (macOS or Windows) — highest-fidelity output.
- **LibreOffice** (`soffice`) — cross-platform; auto-used when Word isn't present.

If neither is installed, every other verb still works — you just can't rasterize
pages. And you usually don't need to: `docx read` already proves content changes
through the write→read loop. Render only when the open question is LAYOUT
(columns, page/section breaks, image sizing, table geometry), and then once.

## "Did my change apply?" — read the exit code

Exit code is the contract: `0` ok, `1` error, `2` usage, `3` not-found. Every
command also prints a one-line confirmation, so you never have to re-read just to
learn whether a mutation landed. For the full structured ack, add `--verbose`.

## A locator points at the wrong thing

Block ids (`pN`, `tN`, `sN`, …) are **positional** and shift after structural
edits (insert/delete that add or remove blocks). Two fixes:

1. **Re-read** (`docx read FILE` or `--ast`) after a structural edit to get fresh ids.
2. **Batch** instead: `--batch FILE.jsonl` applies many changes from a single
   read, so every locator stays valid across the whole batch. This is the right
   tool for filling a form or applying a multi-point review.

To get an exact span locator without counting characters, run
`docx find FILE "phrase"` and paste the result into `--at`.

## My literal text got mangled (a URL autolinked, prose got reformatted)

The `--markdown`/`--from` channels parse GFM (+ math + CriticMarkup + inline
HTML), which can transform literal prose. When you need text in **verbatim**, use
the parser-free channel: `insert` and `create` take `--text-file PATH` (`-` =
stdin). Every character lands as-is; each newline starts a new paragraph.

## I edited a file and lost some unusual formatting

You shouldn't — docx-cli mutates the underlying XML **in place** and leaves
constructs it doesn't model untouched. If you rebuilt the document instead
(`read` → `create`), that round-trips through a view and can drop what the view
doesn't carry. Prefer in-place `edit`/`replace`/`insert`/`delete` over
read-then-recreate. `git` is your undo (there's no journal).
