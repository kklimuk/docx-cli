# docx-cli

CLI for AI agents (Claude, Codex) to read, edit, and comment on `.docx` files. JSON-AST output, locator-based addressing, full format fidelity via in-place XML mutation.

Default to **Bun**, not Node. Use `Bun.file`, `Bun.write`, `Bun.env`, `Bun.$` over Node equivalents. Bun loads `.env` automatically — don't use dotenv.

## Project Structure

```
src/
  index.ts                    # binary entrypoint (#!/usr/bin/env bun)
  cli/
    index.ts                  # parseArgs dispatcher, COMMANDS map
    help.ts                   # top-level --help
    respond.ts                # respond() / fail() — JSON ack + error helpers
    create/                   # docx create FILE
    read/                     # docx read FILE
    insert/                   # docx insert FILE
      emit.tsx                # <Paragraph> + <RunElement> JSX components (shared by edit too)
    edit/                     # docx edit FILE
    delete/                   # docx delete FILE
    comments/                 # docx comments <verb>
      index.ts                # sub-dispatcher for add/reply/resolve/delete/restore/list
      helpers.tsx             # paraId, ensureCommentsPart, run-splitting marker injection
      add | reply | resolve | delete | restore | list
      trash.ts                # <dir>/.docx-cli/trash.json journal for restore
    images/                   # docx images <verb>
      list | extract | replace
    track-changes/            # docx track-changes FILE on|off
    schema/                   # docx schema [--ts]  (TS source via Bun text import)
    locators-cmd/             # docx locators [--json]
  core/
    package/                  # JSZip wrapper: open, read/write XML parts, save
    parser/
      xml-node.ts             # XmlNode class — instance methods + static parse/serialize
      index.ts                # re-export
    jsx/
      index.ts                # h, Fragment, makeTag, namespaces (w, r, a, wp, pic, cp, dc, …)
      jsx-runtime.ts          # auto-runtime — converts component-null to empty fragments
      jsx-dev-runtime.ts      # alias for dev-mode auto-runtime
    ast/
      types.ts                # Doc / Block / Run / Comment types (read live by `docx schema --ts`)
      doc-view.ts             # DocView, openDocView, saveDocView, enrichImageHashes
      read.ts                 # XML → typed AST walker; populates back-refs
    locators/
      parse.ts                # parseLocator("p3:5-20") → Locator union
      resolve.ts              # resolveBlock / resolveComment / resolveImage
tests/
  core/                       # XmlNode + locator parser unit tests
  cli/
    harness.ts                # Bun.spawn wrapper — runs binary, parses JSON output
    *.test.ts                 # one file per command surface
  integration/
    libreoffice-roundtrip.test.ts  # auto-skips if `soffice` not on PATH
  fixtures/                   # 11 .docx files: minimal + comments-* + tracked-changes + …
scripts/
  move.ts                     # move file + auto-update imports
  fxp-smoke.ts                # JSX → XML smoke test
  inspect-fixtures.ts         # summarize each fixture's features
  make-fixture.ts             # rebuild minimal.docx
  escape-check.ts             # XML entity round-trip sanity
```

## Key Conventions

These are not suggestions. Follow them.

- **Bun not Node**. `Bun.file(path).type` for MIME, `Bun.escapeHTML` for HTML escape (note: prefer fast-xml-builder's auto-escape on the JSX path; `Bun.escapeHTML` produces `&#x27;` for `'` which fast-xml-parser doesn't decode back). All stdout goes through `respond()` (JSON ack) or `writeStdout()` (text) from `src/cli/respond.ts`, both of which use `Bun.write(Bun.stdout, ...)` — never `process.stdout.write` directly. The 64 KB truncation that bites on early exit is real and silent; the helpers are the only safe path.
- **File naming**: kebab-case, named after the primary export (`xml-node.ts` → `XmlNode`, `paragraph.ts` → `Paragraph`).
- **Feature structure**: each command is a kebab-case folder under `src/cli/`. `index.ts(x)` is the public surface (`export async function run(args: string[]): Promise<number>`). Sub-files for shared helpers, no cross-feature imports unless centrally exposed via `@core/*`.
- **Path aliases**: `@core` → `src/core/index.ts`, `@core/*` → `src/core/*/index.{ts,tsx}` or `src/core/*.{ts,tsx}` (configured in tsconfig). Use these in `src/cli/*` to avoid `../../core` chains. `src/core` itself uses relative imports between siblings.
- **JSX is for emitters only**. `.tsx` files only in `src/cli/*` (and `src/core/jsx`). Never in the AST reader. Components are PascalCase (`<Paragraph>`, `<RunProperties>`); they accept props and may return `XmlNode | null` (null skipped by flatten). Tag namespaces are imported as lowercase (`w.p`, `cp.coreProperties`); attribute names with colons go through hyphen shortcut (`w-val="800080"` → `w:val="800080"`) or JSX spread (`{...{"w:val": "x"}}`).
- **Variable names**: no single/two-letter names. Use descriptive (`paragraph` not `p`, `commentId` not `cid`). Exception: standard regex-match destructuring (`const [, prefix, idx] = match`) since the convention is unambiguous.
- **Spacing**: tabs, double quotes (Biome enforced).
- **Top-down file ordering**: exports first, dependencies below — like a newspaper. Use `function` declarations (hoisted) for internal helpers so this works at runtime. Arrow functions only for inline callbacks and short utilities.
- **Early returns**: bail on validation failures and null checks. No else-if chains when an early return reads cleaner.

## Architectural Invariants

- **In-place XML mutation, not AST round-trip emission**. The typed AST returned by `read` is a _view_ over the parsed XML tree. Mutations (insert/edit/delete/comments add) operate on the underlying `XmlNode` references via `BlockReference.parent.splice(...)`. Anything we don't model survives because we never re-emit untouched regions. Only emit fresh XML for nodes we're inserting (via JSX) — never round-trip whole subtrees through the AST.
- **fast-xml-builder owns escaping**. On the JSX path, never manually escape — the builder handles entities correctly (uses `&apos;` for `'`, which fast-xml-parser decodes back). On the very-rare static-template path (only used in `cli/create/template.tsx` for `<?xml?>` prefix), use `Bun.escapeHTML` only for non-quote characters since `&#x27;` doesn't round-trip cleanly.
- **JSX.Element = XmlNode** (single, not nullable union). `Fragment` returns a sentinel `#fragment` XmlNode that gets unwrapped both in `flatten()` (composition) and `XmlNode.serialize()` (top-level). Components that want to "render nothing" return `null`; `jsx()` converts that to an empty fragment so callers always see `XmlNode`.
- **Stable positional ids** (`p0`, `t0`, `c0`, `img0`). Block ids shift after structural edits — agents must re-read between non-trivial mutations. Comment numeric ids are allocated as `max-existing + 1`. Image ids are positional (document order).
- **paraId is required for resolve/reply**. Comments authored by external tools may lack `w14:paraId` on their bodies. The `resolve` and `reply` verbs auto-inject one via `ensureCommentParaId()` (also adds `xmlns:w14` to the `<w:comments>` root if missing). Do this rather than failing — agents shouldn't have to recreate comments.

## Commands

`docx <verb>` and `docx <noun> <verb>`. Every command has `--help`. Mutating commands accept `--dry-run`. JSON output by default; structured `{ok: false, code, error, hint}` on failure.

| Surface    | Verbs                                                                        |
| ---------- | ---------------------------------------------------------------------------- |
| top-level  | `create` `read` `insert` `edit` `delete` `track-changes` `schema` `locators` |
| `comments` | `add` `reply` `resolve` `delete` `restore` `list`                            |
| `images`   | `list` `extract` `replace`                                                   |

Exit codes: `0` ok, `1` general, `2` usage, `3` not-found (file/locator/comment/image), `4` permission, `5` already-applied. Defined in `src/cli/respond.ts` (`EXIT` const + `ErrorCode` union).

## Locators

```
pN              paragraph N
pN:S-E          chars S..E within paragraph N
pN:S-pM:E       cross-paragraph range
tN              table N; tN:rRcC for cell at row R, col C; chainable :pK
cN, imgN        comment / image ids
```

Span comments split runs at offsets, preserving `<w:rPr>` on both halves (logic in `cli/comments/helpers.tsx → addCommentMarkersToParagraph`).

## Testing

```bash
bun run test:unit         # core + cli (~3s)
bun run test:integration  # LibreOffice round-trip (auto-skips if no soffice)
bun test                  # everything
bun run check             # biome + knip + tsc
```

LibreOffice install for local integration tests:

- macOS: `brew install --cask libreoffice`
- Linux: `sudo apt-get install libreoffice-core libreoffice-writer`

## Build & Publish

Two distinct artifacts:

- **`bun run build`** → `dist/index.js` — single bundled JS (~530 KB, shebang preserved). What npm publishes; runs under Bun. Triggered automatically by `prepack` before `npm pack`/`npm publish`.
- **`bun run build:binary`** → `dist/docx` — `bun build --compile` standalone executable. What GitHub Releases ship per-platform.

The bundled JS path matters because **path aliases (`@core/*`) and JSX runtime resolution don't work when the package is consumed from `node_modules`** — the consumer's tsconfig doesn't have our paths. The bundle pre-resolves everything; without it, `bun run docx ...` from a globally-installed package fails with `Cannot find module 'react/jsx-dev-runtime'`. So: never ship raw `src/`; always bundle.

## CI

GitHub Actions runs on push to `main` and on PRs (`.github/workflows/ci.yml`):

- `check` — biome + knip + tsc
- `unit-tests` — `bun run test:unit`
- `integration-tests` — installs LibreOffice, runs `bun run test:integration`
- `build-binary` — smoke-builds `dist/docx` via `bun run build:binary`

Tag pushes (`v*`) trigger:

- `.github/workflows/release.yml` — matrix-builds five platform binaries (linux/darwin × x64/arm64 + windows-x64) and publishes via `softprops/action-gh-release`. The `install.sh` at the repo root downloads the right binary for the user's platform.
- `.github/workflows/publish.yml` — runs `npm publish --access public` against the `Publishing` environment (uses npm trusted publishing via OIDC; no `NPM_TOKEN` secret needed).

## When Editing

- New CLI command? Add a folder under `src/cli/`, register in the COMMANDS map in `src/cli/index.ts`, add tests under `tests/cli/`, document in README. If it emits OOXML, the file is `.tsx` and uses JSX components from `@core/jsx`.
- New OOXML tag? Add to the appropriate namespace's tag list in `src/core/jsx/index.ts`. The mapped-type pattern (`namespace("w", W_TAGS)`) survives `noUncheckedIndexedAccess` because the keys are a literal-string union.
- New AST field? Add to `src/core/ast/types.ts`, populate in `src/core/ast/read.ts`, then update `src/cli/schema/index.ts` JSON schema. The `--ts` output reads `types.ts` live via Bun's text import attribute, so it stays in sync automatically.
- New fixture? Drop in `tests/fixtures/`, add to the FIXTURES list in `tests/integration/libreoffice-roundtrip.test.ts` if it should round-trip. The `scripts/inspect-fixtures.ts` summarizes what each one exercises.
- Don't touch `bun.lock` manually; let `bun install` manage it.
