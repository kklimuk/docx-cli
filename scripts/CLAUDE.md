# scripts — repo utilities

- [move.ts](move.ts) — move a TS file and auto-update imports (TS LanguageService-driven).
- [escape-check.ts](escape-check.ts), [fxp-smoke.ts](fxp-smoke.ts), [jsx-smoke.tsx](jsx-smoke.tsx) — ad-hoc probes for XML escaping, fast-xml-parser behavior, and the JSX runtime.
- [word-redlines.sh](word-redlines.sh) — AppleScript oracle that drives Microsoft Word to produce ground-truth track-changes XML (referenced from [src/core/track-changes/replace.tsx](../src/core/track-changes/replace.tsx) and [src/core/notes/CLAUDE.md](../src/core/notes/CLAUDE.md)).
- `data/` — sample `.docx` files used for one-off inspection (not test fixtures; those live in [tests/fixtures/](../tests/fixtures/)).

Fixture builders moved to [tests/fixtures/setup/](../tests/fixtures/setup/CLAUDE.md).
