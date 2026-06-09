# tests/cli — CLI test harness

`runCli(...args)` from [harness.ts](harness.ts) is the default — it runs the CLI
**in-process** (`main(argv)` with output captured via respond.ts's sinks), not
by spawning a subprocess. This keeps the suite fast (the whole unit suite is a
few seconds; spawning `bun src/index.ts` per call was ~10× slower). It still
exercises the real arg-parse → dispatch → command → respond path; only the OS
process boundary is skipped.

## How the tests are organized — ONE rule

**One file per CLI verb / tight verb-family, named `<verb>.test.ts`, holding ALL
of that verb's behavior** — including its tracked-change, `--batch`, formatting,
and style-provisioning variants. Those are properties _of_ the verb, not separate
concerns. To place a new test: find the verb it exercises → that's the file. A
"tracked `tables delete-row`" test goes in [tables.test.ts](tables.test.ts) (it
asserts the `rowDel` shape), not in a generic tracking file.

**Five standalone files** hold behavior that no single verb owns — reach for one
only when the thing under test holds _regardless of which verb runs_:

- [track-changes.test.ts](track-changes.test.ts) — the tracking **engine**:
  toggle / `list` / `accept` / `reject`, `--track`-forces-tracking, and moves.
  (Verb-specific tracked _XML shape_ — edit's del/ins consolidation, footnotes'
  paired body+ref ins — stays in the verb file; the accept/reject engine is here.)
- [batch.test.ts](batch.test.ts) — the cross-verb batch **mechanism**
  (resolve-first / re-read-between / pin-then-splice).
- [output-contract.test.ts](output-contract.test.ts) — the output **contract**
  (quiet-vs-verbose, exit codes, bare-minted-locator, 64 KB boundary, `--help`
  matrix), exercised at the real process boundary.
- [invariants.test.ts](invariants.test.ts) — the in-place-mutation **invariant**
  (unmodeled-XML survival, transparent wrappers, docx validity).
- [end-to-end.test.ts](end-to-end.test.ts) — the full document **lifecycle**
  (author → revise → review → read). It drives many verbs in sequence and asserts
  the result survives the write-read loop. Per-verb mechanics stay in the verb
  file; this file guards that the verbs **compose**.

Two sub-domains split out of their nominal verb because each is large enough to
own a file: [markdown.test.ts](markdown.test.ts) (import + round-trip + run
formatting) and [equations.test.ts](equations.test.ts).

`tests/core/*` stays **pure unit tests** below the CLI surface — never merge CLI
e2e into core. `tests/integration/` stays the LibreOffice slow path.

Within a file, group by `describe` (named for the behavior under test); shared
helpers and fixtures sit near the top or just above the block that needs them.
Reach for shared helpers in [helpers.ts](helpers.ts) (`readMarkdown`,
`readDocumentXml`, `trackedKinds`, `freshFixture`) before re-deriving them.

## When to use spawnCli instead

`spawnCli(...args)` spawns the real binary. Use it **only** when the process
boundary itself is under test — exit-code propagation, argv handling, the 64 KB
`Bun.stdout` truncation, quiet-vs-verbose acks, and stdin `-` ingress (via
`spawnCliStdin(input, …)`, the only way to feed stdin). Those live in
[output-contract.test.ts](output-contract.test.ts), which also bypasses the
harness's `--verbose` auto-injection. Keep the subprocess set small; everything
else uses `runCli`.

## Gotchas

- `runCli` resets the `verboseAck` module-global before each call and restores
  the output sinks in a `finally`. If you add new module-level mutable state to
  a command, reset it the same way or it leaks between in-process calls.
- All CLI output must go through `respond()` / `writeStdout()` / `writeStderr()`
  (respond.ts) — that's the single chokepoint the sinks capture. A stray
  `Bun.write(Bun.stdout, …)` or `console.log` would be invisible to `runCli`.
- The harness auto-injects `--verbose` for mutating verbs (see
  `shouldInjectVerbose`) so tests can assert on the JSON ack; register new
  mutating subverbs there.
