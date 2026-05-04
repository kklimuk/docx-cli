---
name: commit
description: "Create well-structured git commits from the current working tree. Use when the user says 'commit', 'save my work', 'let's commit this', 'make a commit', or any variation of wanting to commit code to git."
---

# Commit

Create clean, well-structured git commits that tell a coherent story.

## Process

### 1. Assess the Working Tree

Run `git status` (never `-uall`) and `git diff` to understand what changed. Also run `git log --oneline -5` to match the repo's existing commit message style.

If there are no changes, say so and stop.

### 2. Group Changes by Intent

Look at the changed files and mentally group them:

- **Feature**: new functionality (client + server + tests for the same feature = one commit)
- **Fix**: bug fixes
- **Refactor**: structural changes that don't change behavior
- **Docs**: documentation-only changes (CLAUDE.md, README, comments)
- **Test**: test-only additions or changes
- **Chore**: config, dependencies, tooling

Rules for grouping:
- **Prefer fewer commits.** A feature that touches 15 files is still one commit if it's one logical change.
- **Only split when intent is genuinely different.** "Add collaborative editing" is one commit even if it touches client, server, DB, and tests. But "add collaborative editing" + "fix unrelated CSS bug" should be two commits.
- **For a first commit or large initial build, one commit is fine.** Don't artificially split an initial implementation.
- **Docs updates that accompany code changes go in the same commit.** Only separate docs commits for docs-only changes.

### 3. Present the Plan

Before committing, show the user:
- How many commits you plan to make
- For each commit: the message and which files are included
- Ask for confirmation

### 4. Create the Commits

For each commit:
1. Stage the specific files with `git add <file1> <file2> ...` (never `git add -A` or `git add .`)
2. Commit with a message using this format:

```
<type>: <concise description>

<optional body — explain WHY, not WHAT. The diff shows what.>

Co-Authored-By: Claude <noreply@anthropic.com>
```

Types: `feat`, `fix`, `refactor`, `docs`, `test`, `chore`

Message guidelines:
- Subject line under 72 characters
- Use imperative mood ("add", not "added" or "adds")
- The subject should complete the sentence "This commit will..."
- Body is optional — use it for non-obvious context (e.g., "the old approach caused X" or "this unblocks Y")
- Always include the Co-Authored-By trailer

### 5. Verify

After all commits, run `git log --oneline -10` to show the result.

## Safety Rules

- **Never commit `.env`, credentials, or secrets.** Check staged files for these patterns and warn.
- **Never use `git add -A` or `git add .`** — always stage specific files.
- **Never amend a commit** unless the user explicitly asks.
- **Never force push.**
- **Never skip hooks** (`--no-verify`).
- **If a pre-commit hook fails**, fix the issue and create a NEW commit (don't amend). To diagnose, read the pre-commit hook (`.husky/pre-commit` or `.pre-commit-config.yaml`) to see what it runs, then run each command individually to find the failure.

## What NOT To Do

- Don't write commit messages that describe every file changed. The diff does that.
- Don't split a single feature across 5 commits just because it touches 5 directories.
- Don't use vague messages like "update code" or "fix stuff".
- Don't commit generated files (`db/schema.ts`, `dist/`, `node_modules/`, `__pycache__/`).
