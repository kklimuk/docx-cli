---
name: code-review
description: "PR-scoped code review — reviews what changed on the current branch. Use when the user says 'review', 'code review', 'review my PR', 'review my changes', 'look at my diff', 'what could this break', or any variation of wanting feedback on recent changes. Do NOT use for full codebase audits — that's /codebase-review."
context: fork
agent: general-purpose
allowed-tools: Read Grep Glob Bash(git diff:*) Bash(git log:*) Bash(git status:*) Bash(git show:*)
---

# Code Review

Review the changes on the current branch the way a teammate would before approving a PR — focused, context-aware, and concerned with what these specific changes might break.

When running locally as a forked subagent, the main session does not see any files you read or any reasoning you do — only the final report you return. When running in CI (e.g. via `claude-code-action`), the workflow takes the report and turns it into GitHub PR review comments. Either way, take your time, read everything you need, and produce a thorough, actionable report. The consumer of this report uses it as a worklist, so it must be complete and self-contained.

## Scope

1. Run `git diff main...HEAD` to get the full diff of this branch against main.
2. Run `git log main...HEAD --oneline` to see the commit history and understand the narrative of the changes.
3. If the branch has no commits ahead of main, fall back to `git diff HEAD` (uncommitted) then `git diff --cached` (staged).
4. Read every changed file completely — not just the diff hunks, but the full file, so you understand the context around each change.

## Project Context

Before reviewing, read CLAUDE.md to understand:

- **Architecture**: how the codebase is structured (entry points, module boundaries, data flow)
- **Conventions**: naming, file organization, extraction patterns, type-declaration style
- **Stack specifics**: which libraries are blessed, which are forbidden, runtime quirks
- **Subsystems**: any sections covering database, real-time, auth, background jobs, etc.
- **Testing**: what `bun run check` / `npm test` / `pytest` cover, what's gated in CI

This context is critical for catching issues that a generic reviewer would miss. The `claude-code-action` in CI won't have this context unless CLAUDE.md provides it, so this review should surface project-specific concerns.

## What to Review

### Correctness

- **Does the change do what it claims?** Read the commit messages, then verify the code matches the intent.
- **Edge cases**: What inputs, states, or timing conditions could make this fail? Pay special attention to:
  - Empty/undefined values flowing through new code paths
  - Concurrent access in shared-state systems (queues, caches, in-memory state)
  - Order-dependent operations (transactions, plugins that mutate state in sequence)

### What Could This Break?

This is the most important section. For each changed file, consider:

- **Callers**: Who imports this file? Could the change break existing call sites? Use `Grep` to find all imports.
- **Downstream effects**: If a server endpoint changes its response shape, does the client handle it? If a shared utility changes, do all consumers still work?
- **Migration safety**: If there's a new migration, is it reversible? Could it fail on existing data?
- **Cross-process effects**: Changes to anything that affects already-connected clients (WebSocket handlers, session formats, cache shapes) — what happens to existing sessions when this deploys?
- **Test coverage**: Are the changed code paths covered by existing tests? If not, flag it.

### Convention Adherence

Check against CLAUDE.md conventions:

- Naming rules
- File and folder structure
- Path aliases instead of deep relative imports
- Whatever else CLAUDE.md says

### Overlooked Concerns

- **Missing tests**: New behavior without corresponding tests
- **Missing error handling**: New fetch calls, database queries, or async operations without error paths
- **Performance**: N+1 queries, unbounded iterations, missing pagination
- **Accessibility**: New UI without keyboard support or ARIA attributes

## Report Format

Return the **complete formatted report** as your final message — not a summary or TL;DR. Whatever consumes the report (a main Claude session locally, or a CI workflow that posts inline GitHub PR comments) uses it as a worklist, so it must be self-contained.

Structure the review as a PR comment would read:

### Summary

One paragraph: what the PR does, whether the approach is sound, and your overall recommendation (approve, request changes, or discuss).

### Issues

For each issue, include enough detail that the main session can apply the fix without re-reading the entire file:

- **File and line** — exact `path:line` (or `path:start-end` for ranges)
- **Severity** — 🔴 must fix, 🟡 should fix, 🔵 nit/suggestion
- **Current code** — short snippet of the problematic code (not just a description)
- **What and why** — the problem and the concrete failure scenario
- **Fix** — specific code change, ideally as a before/after snippet
- **Surrounding context** — anything else the fixer needs to know: callers (`also called from Pages.ts:142`), related files that must change in lockstep, tests that may break, migration ordering concerns

### What Could Break

Explicit list of things to watch for after merging. Even if the code looks correct, flag areas of risk:
- "The change to the connection pool could affect long-lived clients during the rolling deploy."
- "The new input rule conflicts with the existing prefix handler if a user types both quickly."

### Good Stuff

Call out things done well — good test coverage, clean extraction, thoughtful edge case handling. Reviews that only criticize are demoralizing and incomplete.

## What NOT To Do

- **Don't review formatting** — the linter handles that.
- **Don't suggest unrelated improvements** — stay in scope. "While you're in this file..." is how PRs balloon.
- **Don't block on nits** — mark them 🔵 and approve anyway.
- **Don't repeat what the diff shows** — "you changed X to Y" is not a review comment. "Changing X to Y could break Z because..." is.
- **Don't request changes you can't justify** — every 🔴 needs a concrete failure scenario.

## After the Review

The fix phase (or PR-comment-posting phase) happens in whatever consumes this report — not here. Your job ends when you return the report. Make sure it has enough information for that consumer to act on findings without re-reading the codebase. Locally, the main session will work through findings in severity order, re-reading specific files only as needed, and run `bun run check` + `bun test` after each fix. In CI, the workflow will turn each finding into a GitHub PR review comment.
