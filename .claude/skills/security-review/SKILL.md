---
name: security-review
description: "Review code for security vulnerabilities. Use when the user says 'security review', 'security audit', 'check for vulnerabilities', 'pentest the code', 'OWASP check', or any variation of wanting a security assessment."
context: fork
agent: general-purpose
allowed-tools: Read Grep Glob Bash(git diff:*) Bash(git log:*) Bash(git status:*) Bash(git show:*) WebFetch
---

# Security Review

Audit changed files for security vulnerabilities, focusing on the OWASP Top 10 and issues specific to the project's stack.

When running locally as a forked subagent, the main session does not see any files you read or any reasoning you do — only the final report you return. When running in CI (e.g. via `claude-code-action`), the workflow takes the report and turns it into GitHub PR review comments. Either way, take your time, read every changed file completely, and produce a thorough, actionable report. The consumer of this report uses it as a worklist, so it must be complete and self-contained.

## Scope

Determine the diff to review:

1. Run `git diff main...HEAD --name-only` to get files changed on this branch vs main.
2. If that fails (no `main`, detached worktree, etc.), fall back to `git diff HEAD --name-only` for uncommitted changes, then `git diff --cached --name-only` for staged files.
3. If no diff is available, ask the user which files to review.

Read every changed file completely before starting the review. Read CLAUDE.md first to understand the project's stack and any subsystems with security-sensitive surface area (auth, real-time, payments, file uploads).

## What to Look For

### Injection & Input Handling

- **SQL injection**: Raw SQL with string interpolation instead of parameterized queries. Check for any template literals that build SQL, and verify that database libraries are being used in their parameterized form.
- **Command injection**: User input passed to shell commands (`Bun.$`, `child_process`, `subprocess`, `os.system`) without sanitization.
- **XSS**: User-controlled data rendered as `dangerouslySetInnerHTML`, or reflected into HTML/JS without escaping. Check `contentEditable` fields that accept pasted HTML.
- **Path traversal**: User input used in file paths without validation. Check for `..` traversal.
- **Prototype pollution**: `Object.assign` or spread on user-controlled objects without allowlisting keys.

### Authentication & Authorization

- **Missing auth checks**: Endpoints that read/write data without verifying the caller's identity or org membership.
- **IDOR (Insecure Direct Object Reference)**: Endpoints that accept an ID parameter and return/modify the resource without verifying the caller has access. Particularly dangerous when URL params (like a slug) aren't validated against the actual resource ownership.
- **Privilege escalation**: Actions that should be restricted (delete, move, admin operations) but aren't gated on role/permission.

### Data Exposure

- **Over-fetching**: API responses that include more data than the client needs (e.g., internal IDs, secrets, full document state when only a title is needed).
- **Error leakage**: Stack traces, SQL errors, or internal paths exposed in error responses.
- **Sensitive data in logs**: Passwords, tokens, or PII logged to console.

### Real-time / WebSocket Security

(Only relevant if the project has a WebSocket layer — see CLAUDE.md.)

- **Channel authorization**: Can a client subscribe to any channel by guessing the name? Are channel subscriptions validated against user permissions?
- **Message spoofing**: Can a client broadcast messages to channels they shouldn't have write access to?
- **Payload validation**: Are incoming WebSocket messages validated before processing?

### Denial of Service

- **Unbounded queries**: Endpoints that return all records without pagination or limits.
- **Regex DoS**: User input used in regex patterns without sanitization.
- **Resource exhaustion**: File uploads, large request bodies, or expensive operations without rate limiting or size limits.

### Cryptography & Secrets

- **Hardcoded secrets**: API keys, passwords, or tokens in source code.
- **Weak randomness**: `Math.random()` / `random.random()` used for security-sensitive operations instead of `crypto.randomUUID()` / `secrets.token_*()`.
- **Missing TLS**: WebSocket connections using `ws://` in production contexts.

### Dependencies

- **Known vulnerabilities**: If `bun audit` / `npm audit` / `pip-audit` is available, check for known CVEs.
- **Prototype pollution via deps**: Libraries that merge user input deeply.

## Report Format

Return the **complete formatted report** as your final message — not a summary or TL;DR. Whatever consumes the report (a main Claude session locally, or a CI workflow that posts inline GitHub PR comments) uses it as a worklist, so it must be self-contained.

Organize findings by severity:

### Critical
Exploitable now with no authentication required. Data loss, unauthorized access, or remote code execution.

### High
Exploitable with some preconditions (e.g., needs authenticated user, specific timing). Privilege escalation, significant data leakage.

### Medium
Defense-in-depth issues. Missing validation that's currently protected by another layer but shouldn't rely on it.

### Low
Hardening recommendations. Not exploitable today but reduce attack surface.

For each finding, include enough detail that the consumer can apply the fix without re-reading the entire file:

1. **File and line** — exact `path:line` (or `path:start-end` for ranges); list every site for cross-file findings
2. **Severity** — Critical / High / Medium / Low
3. **Vulnerability type** — OWASP category or CWE
4. **Current code** — short snippet of the vulnerable code (not just a description)
5. **Exploit scenario** — concrete steps showing how an attacker would use this
6. **Fix** — specific code change, ideally as a before/after snippet
7. **Surrounding context** — callers, related files that must change in lockstep, validation layers the fix depends on, tests that should be added

## What NOT to Do

- Don't flag style issues — that's the code review's job.
- Don't suggest adding WAFs, rate limiters, or infrastructure changes unless the code-level fix is insufficient.
- Don't report theoretical issues that require physical access or compromised infrastructure.
- Don't pile on — prioritize the top findings that matter most.

## After the Report

The fix phase (or PR-comment-posting phase) happens in whatever consumes this report — not here. Your job ends when you return the report. Make sure it has enough information for that consumer to act on findings without re-reading the codebase. Locally, the main session will work through findings in severity order with minimal, targeted fixes and run `bun run check` + `bun test` after each. In CI, the workflow will turn each finding into a GitHub PR review comment.
