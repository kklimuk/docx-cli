# tests/integration

## Default run is a curated subset; CI sweeps everything

Each soffice convert costs ~1s, so [libreoffice-roundtrip.test.ts](libreoffice-roundtrip.test.ts) round-trips only `CORE_FIXTURES` by default — one fixture per distinct surface we emit or must preserve (canonical parts, styles, numbering, footnotes, tracked runs, sections, comments, tables). `DOCX_LO_ALL=1` adds `EXTRA_FIXTURES` (the overlapping duplicates) for the full sweep; CI sets it (`.github/workflows/ci.yml`) so coverage isn't lost. When you add a fixture that exercises a **new** emit/preserve surface, add it to `CORE_FIXTURES`; if it just overlaps an existing one, `EXTRA_FIXTURES`.

## Integration tests must isolate the LibreOffice user profile

macOS `soffice --headless` locks the default user profile (`~/Library/Application Support/LibreOffice/4`); any concurrent or stale soffice causes new spawns to exit non-zero. [libreoffice-roundtrip.test.ts](libreoffice-roundtrip.test.ts) creates one `mkdtemp` profile in `beforeAll` and passes `-env:UserInstallation=file://<dir>` on every spawn. Any new test that calls `soffice` directly must follow the same pattern, or it'll flake the moment two test runs (or two test files) overlap.
