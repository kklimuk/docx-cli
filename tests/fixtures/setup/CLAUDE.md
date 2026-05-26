# tests/fixtures/setup — fixture builders

Programs that rebuild the `.docx` files in `tests/fixtures/`. One builder per fixture, named after the fixture: `tests/fixtures/setup/<name>.{ts,tsx}` writes `tests/fixtures/<name>.docx`. Shared canonical-parts helpers live in `./helpers`; the `inspect.ts` audit tool summarizes what each fixture exercises.

## Building a fixture

```sh
bun tests/fixtures/setup/<name>.ts        # rebuilds tests/fixtures/<name>.docx
```

Each script is self-contained — it (re)creates the fixture from scratch and prints `Wrote <path> (N bytes)`. Run only the fixture you intend to change; the others are committed binaries and shouldn't churn for unrelated edits.

To audit what every fixture exercises:

```sh
bun tests/fixtures/setup/inspect.ts
```

## Authoring a new fixture

When a new feature needs a fixture, follow this order:

1. **Prototype with core emitters** (`src/core/blocks`, `src/core/table`, JSX, `ensureStyle`, `allocateNum`, …). Hand-rolled XML must go through the helpers in `./helpers` so the Word-canonical part set ships with the file.
2. **Re-author the fixture with the CLI** once the surface verbs exist. This is the real check: if you can't reproduce the same shape with `docx create | insert | edit | …`, the CLI is missing something — fix the CLI, don't paper over it in the fixture script.
3. **Keep the CLI version as the canonical fixture.** The core-emitters prototype is dev scaffolding; delete it once the CLI version produces an equivalent file.

Exception: a small number of fixtures intentionally stay on core emitters because they capture a **read-only** shape we deliberately don't author (e.g. [task-lists-web.tsx](task-lists-web.tsx) reproduces the Word-for-Web Wingdings checklist shape that our CLI doesn't emit). Call this out in the script's docstring.

## Hand-rolled XML rules

Hand-rolled scripts MUST use the helpers in [helpers/index.ts](helpers/index.ts) (`wrapDocument`, `buildCoreProps`, `buildContentTypes`, `buildRootRels`, `buildDocumentRels`, `addCanonicalParts`) so the fixture ships the Word-canonical part set (see [src/cli/create/CLAUDE.md](../../../src/cli/create/CLAUDE.md)). Scripts that dogfood the CLI (`bun ${cliEntry} create ...`) inherit canonical parts automatically.

Inter-element whitespace inside `<w:body>` and `<cp:coreProperties>` triggers Word's "unreadable content" prompt — those parents have `complexContent` and treat whitespace between children as illegal character data. fast-xml-builder omits inter-element whitespace by default, so the JSX path is safe; only hand-rolled XML here needs care. `wrapDocument` / `buildCoreProps` collapse via `/\n\s*/g` and emit single-line output, so pretty-print indentation in the source is fine as long as it flows through these helpers.

## Round-trip coverage

If a new fixture should round-trip through LibreOffice, add it to `CORE_FIXTURES` in [tests/integration/libreoffice-roundtrip.test.ts](../../integration/libreoffice-roundtrip.test.ts).
