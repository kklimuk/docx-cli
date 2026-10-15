# Contributing to docx-cli

## Development

```sh
bun install && bun run prepare      # set up + git hooks
bun dev <subcommand>                # run via source
bun run check                       # biome + knip + tsc
bun run test:unit                   # core + cli tests (fast)
bun run test:benchmark              # comparison harness checks + types; no model calls
bun run test:integration            # LibreOffice round-trip (needs `soffice` on PATH)
bun test                            # everything
bun run build                       # produce dist/index.js (npm bundle)
bun run build:binary                # produce dist/docx via bun build --compile
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
    diff/                     # docx diff FILE --against SRC (markers.ts: locator normalization; renders read view, diffs via @core/diff)
    comments/                 # add | reply | resolve | delete | list (helpers.tsx: paraId, run-splitting)
    images/                   # list | extract | replace | delete
    hyperlinks/               # add | list | replace | delete (wrap.tsx: run-splitting)
    tables/                   # row/column insert+delete | set-widths | merge | unmerge | borders (grid.ts: merge-aware model)
    track-changes/            # on|off | list | accept | reject (apply.ts holds the unwrap/delete logic)
    info/                     # schema | locators | skill (reference output; skill.ts emits the canonical SKILL.md)
    raw/                      # get | insert | replace — the raw-OOXML escape hatch (commit.ts: reread→assert→schema-gate→save tail)
    validate/                 # ECMA-376 schema check, per WML part (thin shell over @core/raw/validate)
  core/
    diff/                     # pure text→text unified-diff engine (jsdiff wrapper + HTML-aware word refinement); backs docx diff
    package/                  # JSZip wrapper: open, read/write XML parts, save
    parser/                   # XmlNode class + parse/serialize; run-ops.ts (run text/offsets + RUN_BEARING_WRAPPER_TAGS)
    jsx/                      # Fragment, namespace() + tag namespaces (w, a, wp, pic, cp, dc, …), auto-runtime
    ast/                      # types + DocView + XML→AST reader; text.ts (paragraph helpers), sym.ts (<w:sym> decode)
    locators/                 # parse "p3:5-20" → Locator + resolve to refs
    image/                    # drawing.tsx (<Image> + addImagePart + collectImageRuns), source.ts (load path/data:/http + HEIC→JPEG), formats.ts (one mime↔ext table)
    equation/                 # latexToOmml / ommlToLatex bidirectional (temml + own MathML↔OMML adapter)
    markdown/                 # MarkdownImport lens: remark-parse + remark-gfm + remark-math + CriticMarkup plugin → walker → emitters
    raw/                      # Raw lens: fragment gate pipeline (element-order, namespaces+dcx:raw marker, reference audit) + libxml2-wasm XSD engine over bundled ECMA-376 schemas (schemas/)
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
skills/docx-cli/              # the published Agent Skill — SKILL.md (generated), references/
                              #   scripts/install.sh   the ONE installer (also published as a release asset)
                              #   scripts/bootstrap.sh session-start version check; delegates to install.sh
install.sh                    # byte-identical deprecation shim for bootstraps <= v0.22.0 (see CLAUDE.md)
.claude-plugin/               # Claude Code: marketplace.json + plugin.json (skills/ auto-discovered)
.codex-plugin/                # Codex: plugin.json skill-only manifest ("skills": "./skills/")
```

See [CLAUDE.md](CLAUDE.md) and the nested `CLAUDE.md` files under `src/` for conventions, architectural invariants, and per-subsystem playbooks.

## Stack

- **Runtime**: Bun (`node:util` parseArgs, JSX with custom factory, native zlib), compiled to a standalone binary for releases
- **Parser**: [`jszip`](https://www.npmjs.com/package/jszip) + [`fast-xml-parser`](https://www.npmjs.com/package/fast-xml-parser) + [`fast-xml-builder`](https://www.npmjs.com/package/fast-xml-builder)
- **Markdown**: [`unified`](https://www.npmjs.com/package/unified) + [`remark-parse`](https://www.npmjs.com/package/remark-parse) + [`remark-gfm`](https://www.npmjs.com/package/remark-gfm) + [`remark-math`](https://www.npmjs.com/package/remark-math)
- **Diff**: [`diff`](https://www.npmjs.com/package/diff) (jsdiff, BSD-3) backs `docx diff` — unified line hunks + HTML-aware word-level refinement of the read view
- **Math**: [`temml`](https://www.npmjs.com/package/temml) (MIT) compiles LaTeX → MathML; an in-house MathML → OMML adapter handles the OOXML side bidirectionally
- **Render**: [`@hyzyla/pdfium`](https://www.npmjs.com/package/@hyzyla/pdfium) (MIT wrapper; PDFium WASM redistributed under BSD-3-Clause) for the PDF → PNG/JPG step, plus [`pngjs`](https://www.npmjs.com/package/pngjs) / [`jpeg-js`](https://www.npmjs.com/package/jpeg-js) for image encoding
- **Images**: [`heic-convert`](https://www.npmjs.com/package/heic-convert) (wasm libheif, LGPL-3.0) transcodes HEIC/HEIF input to JPEG on insert
- **Validation**: [`libxml2-wasm`](https://www.npmjs.com/package/libxml2-wasm) (MIT, libxml2 compiled to WASM) + the bundled ECMA-376 5th-edition transitional XSDs power `docx validate` and the `docx raw` schema gate — in-process, ~55 ms one-time schema compile, single-digit ms per document
- **Quality**: Biome + Knip + tsc; LibreOffice headless for round-trip integration tests
- **Standard**: ECMA-376 Part 1 §17 (WordprocessingML), Transitional profile

## Agent skill (anti-drift)

`skills/docx-cli/SKILL.md` is **generated by the binary** — `src/cli/info/skill.ts`
is the source of truth. After changing the skill content, regenerate and commit:

```bash
docx info skill > skills/docx-cli/SKILL.md   # or: bun src/index.ts info skill > …
```

A unit test (`tests/cli/info.test.ts`) fails if the committed `SKILL.md` drifts from
`docx info skill`, so `bun run check` / CI catches a stale copy. The skill content is
version-independent; on a release version bump, also bump `version` in
`.codex-plugin/plugin.json` (Claude's `.claude-plugin/` files carry no pinned version).

## CI

GitHub Actions (`.github/workflows/ci.yml`) runs these jobs on push to `main` and on PRs:

| Job                 | What                                                        |
| ------------------- | ----------------------------------------------------------- |
| `check`             | `biome check . && knip-bun && tsc --noEmit`                 |
| `unit-tests`        | `bun run test:unit` and `bun run test:benchmark` (no model calls) |
| `macos-output`      | Output contract and real-pipe regressions on macOS with Bun 1.3.10 and latest |
| `integration-tests` | Installs LibreOffice, runs `bun run test:integration`       |
| `build-binary`      | Builds on Linux and native Intel/ARM macOS; signs/verifies macOS binaries; runs version/help, read, and schema-validation smoke checks |

`.github/workflows/release.yml` triggers on `v*` tags, matrix-builds the five binaries, and uploads them to a GitHub Release via [`softprops/action-gh-release`](https://github.com/softprops/action-gh-release).

macOS release binaries are ad-hoc signed with `sh scripts/sign-macos-binary.sh BINARY`
after compilation and any other binary modifications. The script supplies Bun's
JavaScriptCore entitlements and requires strict signature verification to pass.
Both macOS release jobs run the artifact natively before upload. `SHA256SUMS` is
then generated from the downloaded, signed artifacts; never modify binaries after
signing or checksum generation. The same signing and smoke checks run on PRs.
To reproduce locally on macOS, run `bun run build:binary`, then
`sh scripts/sign-macos-binary.sh ./dist/docx` and the `build-binary` smoke commands
in `.github/workflows/ci.yml`. These checks cover the runner's macOS version;
compatibility with a newer macOS release needs a separate run on that release.
