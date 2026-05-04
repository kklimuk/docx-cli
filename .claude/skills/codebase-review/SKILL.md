---
name: codebase-review
description: "Full codebase audit — architecture, structural health, technical debt. Use when the user asks for a 'codebase review', 'architecture review', 'codebase audit', 'full review', 'engineering critique', 'refactoring plan', or 'what would a senior engineer think of this codebase'. Do NOT use for reviewing a PR or branch diff — that's /code-review."
context: fork
agent: general-purpose
allowed-tools: Read Grep Glob Bash(git log:*) Bash(git status:*) Bash(bun run check:*)
---

# Staff+ Code Review

A systematic process for reviewing codebases the way a senior staff engineer would — focusing on structural health, maintainability, and the kinds of issues that compound over time rather than surface-level style nits.

When running locally as a forked subagent, the main session does not see any files you read or any reasoning you do — only the final report you return. When running in CI, the workflow consumes the report directly. Either way, this is a serious, deliberate audit: take the time to read every relevant file, hold the whole picture in mind, and produce a thorough, considered report. The consumer of this report uses it as a worklist, so it must be complete and self-contained.

## Philosophy

A good code review isn't a list of complaints. It's a conversation about where a codebase is headed and whether its current structure supports that trajectory. The best reviews identify patterns (good and bad), not just individual issues. They distinguish between "this is wrong" and "this will hurt you in six months."

Think of it like a doctor's checkup: you're looking for systemic health, not just symptoms. A duplicated function is a symptom; the absence of an extraction pattern is the disease.

## The Review Process

### Phase 1: Orientation

Before critiquing anything, understand the codebase on its own terms.

1. **Read the project documentation** — CLAUDE.md, README, any architecture docs. Understand the stated intent, conventions, and constraints.
2. **Map the architecture** — identify entry points, the module dependency graph, and data flow. Use `Glob` and `Grep` to build a mental model before diving into individual files.
3. **Identify the tech stack and its idioms** — a Bun project has different conventions than a Node project. A Rails-influenced TypeScript codebase will value convention over configuration. A Python+FastAPI service has different bones than a Django app. Meet the code where it is.

### Phase 2: The Review

Read every file in the source directory systematically. For each file, hold these questions in mind:

#### Structural Issues (highest impact)

- **Type/interface duplication**: Is the same shape defined in multiple places? When one definition changes and the others don't, you get subtle runtime bugs that pass type-checking. Look for interfaces or type aliases with identical or near-identical fields across files.
- **Dead code**: Exported functions that nothing imports. Feature flags that are always true. Utility functions written for a refactor that never landed. Dead code is cognitive tax on every future reader. If the project has a dead-code linter (`knip`, `vulture`), run it to get a machine-verified list of unused exports, files, and dependencies — don't rely on manual inspection alone.
- **Missing single source of truth**: When a constant, type, or configuration value is defined in more than one place, which one wins? There should be one canonical location, and everything else should import from it.
- **Pipeline/decomposition clarity**: Is the main orchestration function doing too much? Good code reads like a table of contents — the orchestrator calls well-named functions in sequence, and each function does one thing. Use comment delimiters (`// ─── Section ───`) as a code smell — if you need a comment to separate sections, they should probably be separate functions or files.
- **File/folder decomposition**: When a module grows, follow the sibling file + subfolder convention: `foo.ts` sits alongside `foo/` which holds its implementation details. The file is the public interface; the folder contains internal helpers. Don't let a single file accumulate unrelated responsibilities just because they're in the same domain.
- **Separation of concerns in pipelines**: Pipeline modules that produce results should use generators (or async generators) that yield individual results. The caller (orchestrator) handles I/O, counting, and progress logging — the generator only knows how to process and yield. This keeps analysis/processing logic testable without touching the filesystem or mixing in persistence concerns.

#### Robustness Issues

- **Silent catch blocks**: `catch {}` or `catch { /* ignore */ }` (or `except Exception: pass`) is almost always wrong. Every catch should either rethrow, log with context, or have a comment explaining exactly why swallowing the error is safe. The one legitimate case is when you're parsing untrusted input and non-matching input is expected — and even then, a comment is warranted.
- **Missing guard clauses**: Look for array access (`arr[0]`) or property access on values that could be undefined/empty without a preceding check. Particularly in loops processing external data.
- **Asymmetric error handling**: If two similar code paths handle errors differently, the weaker one will eventually bite. Both should be equally robust.
- **Stale/retry logic**: If the code has retry or recovery mechanisms, are they aggressive enough? Compare similar mechanisms across the codebase — if one path has sophisticated recovery and another has a minimal version, flag it.

#### Maintainability Issues

- **Structural duplication**: Not just copy-paste, but repeated patterns that should be extracted. If two functions follow the same resolve → check → transform → collect pattern with different types, that's a helper waiting to be born.
- **Import hygiene**: Unused imports, inconsistent alias usage, deep relative paths where aliases exist. These are small individually but signal a codebase that isn't being actively maintained.
- **Naming**: Abbreviated names (`cfg`, `ctx`, `mgr`) make code harder to grep and harder to read six months later. Full descriptive names throughout.
- **Definition ordering**: Within files, are the most important things (exports, entry points) at the top, or do you have to scroll past helpers to find the main function?

#### Testing & Verification

- **Test coverage gaps**: Not just "are there tests" but "do the tests cover the interesting cases?" Edge cases, error paths, and schema validation are where bugs hide.
- **Fixture freshness**: If tests use fixtures, do the fixtures match the current data shape? Stale fixtures mean tests pass but don't actually validate current behavior.
- **Schema validation in tests**: Every parsed output should be validated against the schema. This catches drift between what parsers produce and what the rest of the system expects.

#### Documentation Drift

- **CLAUDE.md accuracy**: Compare the architecture section, command examples, data model, and key patterns against the actual code. Flag any commands, file paths, type shapes, or pipeline descriptions that no longer match reality. CLAUDE.md is the primary onboarding document — if it's wrong, every future session starts with a lie.
- **README.md accuracy**: If a README exists, check that setup instructions, usage examples, and feature descriptions reflect the current state. Outdated READMEs are worse than no README — they actively mislead.
- **Rules in `.claude/rules/`** (if present): Read each rule file and verify its guidance still applies. Rules that reference deleted files, renamed functions, or superseded patterns should be updated or removed.
- **Consistency across docs**: If CLAUDE.md says one thing and the code does another, flag the drift. The fix is always to update the docs to match the code, not the other way around.

### Phase 3: Prioritized Report

Return the **complete formatted report** as your final message — not a summary or TL;DR. Whatever consumes the report (a main Claude session locally, or a CI workflow) uses it as a worklist, so it must be self-contained.

Present findings organized by impact, not by file. Group them into tiers:

1. **Critical** — will cause bugs or data loss (type mismatches, missing error handling on critical paths)
2. **High** — structural issues that compound (duplication, dead code, missing single source of truth)
3. **Medium** — maintainability concerns (naming, ordering, import hygiene)
4. **Low** — style preferences and minor improvements

For each finding, include enough detail that the main session can apply the fix without re-reading the entire codebase:

- **Location** — exact `path:line` (or `path:start-end` for ranges); for cross-file findings, list every site
- **Tier** — Critical / High / Medium / Low
- **Current code** — short snippet(s) of the problematic code
- **What and why** — explain the failure mode. Don't just say "this is duplicated" — explain what goes wrong when the definitions drift apart. Don't just say "add a guard clause" — explain what input would cause the crash.
- **Fix** — specific code change, ideally as a before/after snippet, plus any cross-file changes that must happen in lockstep
- **Surrounding context** — callers, related files, tests that may break, ordering constraints, anything else needed to apply the fix safely

For pattern-level findings ("this pattern repeats across N files"), list all the affected files and call out the shared root cause once, rather than repeating the same finding N times.

### Phase 4: Iterative Fixing

The fix phase happens in whatever consumes this report — not here. Your job ends when you return the report. Locally, the main session will work through findings in tier order using your report as the worklist, re-reading specific files only as needed. In CI, the workflow will surface findings as appropriate. Either way, the consumer should be able to:

1. Fix in priority order, one concern per edit
2. Preserve or update tests as needed
3. Run lint, tests, type-check, and dead code check after each batch (e.g. `bun run check` + `bun test`)
4. Update CLAUDE.md / README.md / `.claude/rules/` to reflect any changes
5. Respect the project's own conventions (use what CLAUDE.md prescribes — don't introduce new tooling)

Your job is to make sure the report has enough information for that to be straightforward.

## What NOT To Do

- **Don't nitpick formatting** — that's the linter's job. If the project has a formatter configured, trust it.
- **Don't suggest rewriting in a different language/framework** — work within the existing tech choices.
- **Don't conflate personal preference with engineering quality** — "I prefer X" is not a review finding. "X prevents Y class of bug" is.
- **Don't pile on** — if there are 30 issues, prioritize the top 7-10 that matter most. The user can always ask for more.
- **Don't suggest adding dependencies for simple things** — especially for testing. Use what's already in the project.

## Adapting to the User

Pay attention to the user's background and experience level. A user who mentions Ruby on Rails experience will appreciate DRY principles, convention over configuration, and "sharp knives" philosophy. A user from the Java world will resonate with interface segregation and dependency inversion. A user who's new to programming needs gentler framing and more explanation of the "why."

The best review meets the developer where they are and uses concepts they already value to motivate improvements.
