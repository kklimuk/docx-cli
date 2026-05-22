# docx-cli

CLI for AI agents to read, edit, and comment on `.docx` files. JSON-AST output, locator-based addressing, full format fidelity via in-place XML mutation.

**Bun, not Node.** Use `Bun.file`, `Bun.write`, `Bun.env`, `Bun.$`. Bun loads `.env` automatically — no dotenv.

Subsystem-specific guidance lives in nested CLAUDE.md files that load when you edit those folders. If you need to add a new CLAUDE.md to describe a new practice for a part of the system, do so.

## Conventions

These conventions are NOT SUGGESTIONS. These are rules.

- **All stdout goes through `respond()` (JSON ack) or `writeStdout()` (text)** from `src/cli/respond.ts` — never `process.stdout.write`. Both use `Bun.write(Bun.stdout, ...)`; the 64 KB truncation that bites on early exit is real and silent, and these helpers are the only safe path.
- **File naming**: kebab-case, named after the primary export (`xml-node.ts` → `XmlNode`).
- **JSX is for emitters only.** Files that construct fresh XML can be `.tsx`; readers/locators/analysis stay `.ts`. Components are PascalCase, accept props, may return `XmlNode | null` (null skipped by flatten). Attribute names with colons use the hyphen shortcut (`w-val="x"` → `w:val="x"`) or JSX spread.
- **Component vs function**: a pure `props → XmlNode` builder is a PascalCase component — destructure its props in the signature (no `props.x` access). Anything that takes a `DocView` or mutates package state (minting relationships, provisioning styles/numbering, splicing into the tree) is an _operation_, not a component — keep it a plain `function`. Same test decides what belongs in the `src/core` emitters (`blocks`/`table`/`sections`/`styles`/`numbering`) vs a CLI command's glue (e.g. `buildTextParagraph`/`wrapFirstRunInHyperlink` stay functions because they thread `view`).
- **JSX.Element = XmlNode** (single, not nullable). `Fragment` returns a `#fragment` sentinel unwrapped in `flatten()` and `serialize()`. Components return `null` to render nothing; `jsx()` converts that to an empty fragment. The `jsx`/`jsxs`/`jsxDEV` runtime exports are distinct functions, not `= jsx` aliases (knip flags aliased re-exports as duplicates) — don't collapse them.
- **Path aliases**: `@core` → `src/core/index.ts`, `@core/*` → `src/core/*`. Use these in `src/cli/*`; `src/core` itself uses relative sibling imports. Import the body emitters from the `@core/blocks` and `@core/table` subpaths, not the `@core` barrel — `ast/types` already exports `Paragraph`/`Table`/`TableCell`/`TableRow` as _types_, and barrel-merging the same-named value emitters is confusing.
- **Variable names**: descriptive, no single/two-letter (`paragraph` not `p`). Exception: regex-match destructuring (`const [, prefix, idx] = match`).
- **Newspaper ordering.** The entry point (primary export) goes at the top; its dependencies follow in the order it uses them, then _their_ dependencies, and so on — a file reads top-to-bottom like a newspaper. Use hoisted `function` declarations for internal helpers so this works at runtime; arrow functions only for inline callbacks and short utilities. When a file accumulates too many dependencies to read this way, split them into a separate folder/file named after the feature they're working on. [src/core/styles.tsx](src/core/styles.tsx) is the canonical example (`ensureStyle` → its helpers → the `BASELINE` catalog → the individual style components).
- **Inline props in the signature.** When a component's props type is used only by that component, write it inline (`function HeadingStyle({ styleId }: { styleId: BaselineStyleId; … })`) rather than declaring a separate named `Props` type. Extract a named type only when it's shared.
- **knip runs strict** (`bun run check`, no rule overrides in `knip.json`). An unused export is dead code — delete it (this is a CLI app, not a library; there are no external `@core` consumers). The one exception: an export staged for a named upcoming tier with no caller yet gets a `@public` JSDoc tag whose comment names the future consumer (knip honors `@public`) — e.g. `HorizontalRule` (S8) and the `r`/`a`/`wp`/`pic` image namespaces (S5). Don't silence knip by re-adding rule suppressions.
- **Style**: tabs, double quotes (Biome enforced). Early returns over else-if chains.

## Invariants

These invariants are NOT SUGGESTIONS. These MUST be followed.

- **In-place XML mutation, not AST round-trip** — the AST is a view; mutate `XmlNode` refs, only emit fresh XML for inserted nodes. See [src/core](src/core/CLAUDE.md).
- **`RUN_BEARING_WRAPPER_TAGS`** in `src/core/parser/run-ops.ts` is the AST↔XML offset bridge; every offset-aware walker reads it. See [src/core](src/core/CLAUDE.md).
- **Stable positional ids** (`p0`, `t0`, `c0`, `img0`, `link0`, `tc0`). Block ids shift after structural edits — re-read between non-trivial mutations. Comment ids are `max-existing + 1`.
- **Hyperlinks own a relationship, not their text.** `hyperlinks replace` updates the `<Relationship>` `Target` (mints a new rId if multiple `<w:hyperlink>` share one); `delete` unwraps and prunes the rId when unreferenced.
- **paraId is required for resolve/reply** — auto-injected via `ensureCommentParaId()` rather than failing.
- **Track-changes is doc-level** — see [src/cli/track-changes](src/cli/track-changes/CLAUDE.md).
- **Sections are blocks; CRUD goes through the standard verbs** — see [src/core](src/core/CLAUDE.md).
- **Table structure is a merge-aware logical grid** — `gridSpan`/`vMerge` map logical (row,col) onto physical `<w:tc>`. The `docx tables` verbs reshape rows/columns/merges/widths/borders through that model — see [src/cli/tables](src/cli/tables/CLAUDE.md).
- **No undo, no journal.** Mutating commands overwrite `FILE` in place; git is the history. `-o/--output PATH` writes a parallel file; `--dry-run` previews (wins over `--output`).

## Commands

`docx <verb>` and `docx <noun> <verb>`. Every command has `--help`. Mutating commands accept `--dry-run` and `-o/--output PATH`. JSON by default; `{ok: false, code, error, hint}` on failure. Exit codes: `0` ok, `1` general, `2` usage, `3` not-found (defined in `src/cli/respond.ts`).

| Surface         | Verbs                                                                                              |
| --------------- | -------------------------------------------------------------------------------------------------- |
| top-level       | `create` `read` `insert` `edit` `delete` `find` `replace` `wc` `outline`                           |
| `comments`      | `add` `reply` `resolve` `delete` `list`                                                            |
| `images`        | `list` `extract` `replace`                                                                         |
| `hyperlinks`    | `add` `list` `replace` `delete`                                                                    |
| `tables`        | `insert-row` `delete-row` `insert-column` `delete-column` `set-widths` `merge` `unmerge` `borders` |
| `track-changes` | `FILE on\|off` (toggle), `list FILE`, `accept`/`reject FILE (--at tcN \| --all)`                   |
| `info`          | `schema` `locators`                                                                                |

`docx read FILE` renders GFM by default (`--ast` for JSON-AST). Tracked changes have three views: default `current` (CriticMarkup `{++ins++}` / `{--del--}` with `[^tcN]` footnotes), `--accepted` (drops subtractive, inlines additive), `--baseline` (drops additive, inlines subtractive). `--comments` appends `[^cN]` footnotes. `wc` accepts the same `--accepted`/`--baseline` flags. Render/locator logic in `cli/read/markdown.ts`; feature detection in `core/ast/read.ts`.

## Locators

```
pN              paragraph N        pN:S-E          chars S..E within paragraph N
pN:S-pM:E       cross-paragraph    tN / tN:rRcC    table N / cell at row R col C (chainable :pK)
sN              section break N    cN imgN linkN tcN   comment / image / hyperlink / tracked-change ids
tN:rR tN:cC     table row R / column C (the `tables` verbs) tN:rR1cC1-rR2cC2   rectangular cell region (merge)
```

Span comments and `hyperlinks add` split runs at offsets, preserving `<w:rPr>` on both halves (`cli/comments/helpers.tsx`, `cli/hyperlinks/wrap.tsx`).

## Testing

```bash
bun run test:unit         # core + cli (~3s)
bun run test:integration  # LibreOffice round-trip (auto-skips if no soffice)
bun test                  # everything
bun run check             # biome + knip + tsc
```

## Build

- `bun run build` → `dist/index.js` — bundled JS that npm publishes (runs under Bun). **Required**: path aliases (`@core/*`) and JSX runtime resolution don't work when consumed from `node_modules`; the bundle pre-resolves everything. Never ship raw `src/`.
- `bun run build:binary` → `dist/docx` — standalone executable for GitHub Releases.

## Docs layout

Three docs, each for a different reader — don't cross the streams. When you change one, check the others:

- **README.md** — user-facing: install, examples, command reference, "How It Works."
- **CONTRIBUTING.md** — dev-facing: setup/test commands, LibreOffice install, project-structure tree, CI table.
- **CLAUDE.md** _(this file + nested)_ — agent-facing: conventions, invariants, per-subsystem playbooks.

New CLI command → README + `src/cli/CLAUDE.md`. New invariant → CLAUDE.md only. New build/test step → CONTRIBUTING.md + this Testing section. New runtime dep → README + CONTRIBUTING.md.
