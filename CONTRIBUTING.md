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
    track-changes/       # on|off | list | accept | reject (apply.ts holds the unwrap/delete logic)
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

See [CLAUDE.md](CLAUDE.md) for deeper conventions, architectural invariants, and the project structure walkthrough.

## CI

GitHub Actions (`.github/workflows/ci.yml`) runs four jobs on push to `main` and on PRs:

| Job                 | What                                                        |
| ------------------- | ----------------------------------------------------------- |
| `check`             | `biome check . && knip-bun && tsc --noEmit`                 |
| `unit-tests`        | `bun run test:unit` (core + cli, fast)                      |
| `integration-tests` | Installs LibreOffice, runs `bun run test:integration`       |
| `build-binary`      | Smoke-builds via `bun build --compile` and runs `--version` |

`.github/workflows/release.yml` triggers on `v*` tags, matrix-builds the five binaries, and uploads them to a GitHub Release via [`softprops/action-gh-release`](https://github.com/softprops/action-gh-release).
