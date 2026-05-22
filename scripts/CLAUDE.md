# scripts — fixture builders & utilities

`make-*-fixture.{ts,tsx}` rebuild the fixtures under `tests/fixtures/`. `move.ts` moves a file and auto-updates imports; `inspect-fixtures.ts` summarizes what each fixture exercises.

## Adding a fixture

Prefer building it programmatically via a `make-*-fixture.ts` script (reproducible, MIT-compatible). Add it to the FIXTURES list in `tests/integration/libreoffice-roundtrip.test.ts` if it should round-trip.

Hand-rolled scripts MUST use the helpers in [fixture-helpers.ts](fixture-helpers.ts) (`wrapDocument`, `buildCoreProps`, `buildContentTypes`, `buildRootRels`, `buildDocumentRels`, `addCanonicalParts`) so the fixture ships the Word-canonical part set (see `src/cli/create/CLAUDE.md`). Scripts that dogfood the CLI (`bun ${cliEntry} create ...`) inherit canonical parts automatically.

## Hand-rolled XML must avoid inter-element whitespace inside `<w:body>` and `<cp:coreProperties>`

These parents have `complexContent` and don't allow text — Word treats whitespace between child elements as illegal character data and triggers the "unreadable content" prompt. fast-xml-builder omits inter-element whitespace by default, so the JSX path is safe; only hand-rolled XML here needs care. `wrapDocument`/`buildCoreProps` collapse via `/\n\s*/g` and emit single-line output. Pretty-print indentation in fixture sources is fine as long as it flows through these helpers.

Don't touch `bun.lock` manually; let `bun install` manage it.
