# tests/cli — CLI test harness

`runCli(...args)` from [harness.ts](harness.ts) is the default — it runs the CLI
**in-process** (`main(argv)` with output captured via respond.ts's sinks), not
by spawning a subprocess. This keeps the suite fast (the whole unit suite is a
few seconds; spawning `bun src/index.ts` per call was ~10× slower). It still
exercises the real arg-parse → dispatch → command → respond path; only the OS
process boundary is skipped.

## When to use spawnCli instead

`spawnCli(...args)` spawns the real binary. Use it **only** when the process
boundary itself is under test — exit-code propagation, argv handling, the 64 KB
`Bun.stdout` truncation. Those live in [binary-smoke.test.ts](binary-smoke.test.ts)
and [quiet-default.test.ts](quiet-default.test.ts) (which also bypasses the
harness's `--verbose` auto-injection). Keep the subprocess set small; everything
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
