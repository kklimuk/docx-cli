# Contributing to docx-cli

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
  index.ts                    # binary entrypoint (#!/usr/bin/env bun)
  cli/
    index.ts                  # parseArgs dispatcher, COMMANDS map
    help.ts                   # top-level --help
    respond.ts                # respond() / fail() — JSON ack + error helpers
    create/                   # docx create FILE (+ canonical-parts.ts, template.tsx)
    read/                     # docx read FILE (markdown default, --ast for JSON; markdown.ts renderer)
    insert/                   # docx insert FILE
    edit/                     # docx edit FILE
    delete/                   # docx delete FILE
    find/                     # docx find FILE QUERY
    replace/                  # docx replace FILE PATTERN REPLACEMENT (replace-span.tsx: run-splitting)
    wc/                       # docx wc FILE [LOCATOR] (count.ts)
    outline/                  # docx outline FILE (build.ts: heading-tree builder)
    comments/                 # add | reply | resolve | delete | list (helpers.tsx: paraId, run-splitting)
    images/                   # list | extract | replace | delete
    hyperlinks/               # add | list | replace | delete (wrap.tsx: run-splitting)
    tables/                   # row/column insert+delete | set-widths | merge | unmerge | borders (grid.ts: merge-aware model)
    track-changes/            # on|off | list | accept | reject (apply.ts holds the unwrap/delete logic)
    info/                     # schema | locators (reference output)
  core/
    package/                  # JSZip wrapper: open, read/write XML parts, save
    parser/                   # XmlNode class + parse/serialize; run-ops.ts (run text/offsets + RUN_BEARING_WRAPPER_TAGS)
    jsx/                      # Fragment, namespace() + tag namespaces (w, a, wp, pic, cp, dc, …), auto-runtime
    ast/                      # types + DocView + XML→AST reader; text.ts (paragraph helpers), sym.ts (<w:sym> decode)
    locators/                 # parse "p3:5-20" → Locator + resolve to refs
    image/                    # drawing.tsx (<Image> + addImagePart + collectImageRuns), source.ts (load path/data:/http + HEIC→JPEG), formats.ts (one mime↔ext table)
    blocks.tsx                # <Paragraph> / <RunElement> / <ListParagraph> / <HorizontalRule> emitters
    table.tsx                 # <BlankTable> / <Table> / <TableRow> / <TableCell> emitters
    sections.tsx              # sectPr emitters + mutators
    styles.tsx                # ensureStyle — lazy styles.xml provisioning + baseline catalog
    numbering.tsx             # allocateNum — lazy numbering.xml provisioning
    relationships.ts          # addHyperlinkRelationship + _rels helpers
tests/
  core/                       # XmlNode + locator + numbering unit tests
  cli/                        # one file per command surface (harness.ts wraps Bun.spawn)
  integration/                # LibreOffice round-trip (auto-skips if no soffice)
  fixtures/                   # .docx files
    setup/                    # builders: setup/<name>.{ts,tsx} writes ../<name>.docx (helpers/ for canonical parts; inspect.ts audit)
scripts/                      # repo utilities (move.ts, word-redlines.sh, smoke probes)
```

See [CLAUDE.md](CLAUDE.md) and the nested `CLAUDE.md` files under `src/` for conventions, architectural invariants, and per-subsystem playbooks.

## CI

GitHub Actions (`.github/workflows/ci.yml`) runs four jobs on push to `main` and on PRs:

| Job                 | What                                                        |
| ------------------- | ----------------------------------------------------------- |
| `check`             | `biome check . && knip-bun && tsc --noEmit`                 |
| `unit-tests`        | `bun run test:unit` (core + cli, fast)                      |
| `integration-tests` | Installs LibreOffice, runs `bun run test:integration`       |
| `build-binary`      | Smoke-builds via `bun build --compile` and runs `--version` |

`.github/workflows/release.yml` triggers on `v*` tags, matrix-builds the five binaries, and uploads them to a GitHub Release via [`softprops/action-gh-release`](https://github.com/softprops/action-gh-release).
