# src/core/diff — the unified-diff engine

A **pure text→text** unified-diff engine, thin over jsdiff (`diff`). It knows
nothing about docx: it takes two strings and produces hunks / unified-diff text.
The `docx diff` command renders both document sides to their `read` markdown,
normalizes markers (`cli/diff/markers.ts`), then hands the two strings here.

**Must not import from `src/cli` or `src/core/ast/document`.** `renderMarkdown`
lives in the CLI layer, so rendering happens there and only strings cross into
core — keeping this engine reusable and the dependency graph acyclic.

- `diffHunks(old, new)` → structured hunks (empty ⇒ identical); the `--json` shape.
- `renderUnified(hunks, {oldLabel, newLabel, wordDiff})` → unified-diff text.
  `wordDiff` (default on) refines an isolated 1-delete/1-insert pair into one
  `[-removed-]{+added+}` line — read renders each table ROW as one long line, so
  line-level diff on tables is coarse without it.
- `diffStats(hunks)` → `{hunks, added, removed}` for the at-a-glance header.

The word pass tokenizes HTML tags and `<!-- … -->` comments as ATOMIC units —
tuned for the dense `<span>`/`<mark>`/`<u>` runs and `docx:` annotations the read
view emits, so a change straddling markup never splits a tag. Tokenization is
lossless for any input (a bare `<` is kept), so plain, non-HTML text diffs
correctly too.

jsdiff supplies `structuredPatch` (line hunks) and `diffWordsWithSpace` (the
word pass). If the hand-off ever needs battle-tested edge cases we don't cover,
they already live in the dependency — don't reinvent them here.
