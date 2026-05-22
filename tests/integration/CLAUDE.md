# tests/integration

## Default run is a curated subset; CI sweeps everything

Each soffice convert costs ~1s, so [libreoffice-roundtrip.test.ts](libreoffice-roundtrip.test.ts) round-trips only `CORE_FIXTURES` by default — one fixture per distinct surface we emit or must preserve (canonical parts, styles, numbering, footnotes, tracked runs, sections, comments, tables). `DOCX_LO_ALL=1` adds `EXTRA_FIXTURES` (the overlapping duplicates) for the full sweep; CI sets it (`.github/workflows/ci.yml`) so coverage isn't lost. When you add a fixture that exercises a **new** emit/preserve surface, add it to `CORE_FIXTURES`; if it just overlaps an existing one, `EXTRA_FIXTURES`.

## Conversions run in a parallel pool; each soffice needs its own profile

The slow part is the per-fixture `soffice --convert-to` (~1s each), so
[libreoffice-roundtrip.test.ts](libreoffice-roundtrip.test.ts) preps every
fixture serially (fast in-process `runCli` — which shares a global output sink,
so it is NOT concurrency-safe) and then runs the conversions through a bounded
worker pool (`CONCURRENCY` workers) in `beforeAll`, storing results for the
per-fixture tests to assert. This roughly halves the suite.

macOS `soffice --headless` locks the default user profile
(`~/Library/Application Support/LibreOffice/4`); any concurrent or stale soffice
makes new spawns exit non-zero. So each pool **worker** gets its own `mkdtemp`
profile (`newProfile()`) passed as `-env:UserInstallation=file://<dir>` —
concurrent converts never share one. Verified stable across repeated full
runs. Any new code that spawns `soffice` must take its own exclusive profile,
or it'll flake the moment two converts overlap.
