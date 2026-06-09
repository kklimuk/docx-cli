export const meta = {
	name: "adversarial-code-review",
	description:
		"Skeptical multi-agent review of the uncommitted docx-cli changes: review each slice, adversarially verify every finding by reproducing it, then synthesize a prioritized report.",
	phases: [
		{ title: "Review", detail: "one skeptical reviewer per change slice" },
		{ title: "Verify", detail: "reproduce/refute each finding against the code" },
		{ title: "Synthesize", detail: "prioritized report of confirmed findings" },
	],
};

const parsed = typeof args === "string" ? JSON.parse(args) : args || {};
const BASE = parsed.baseRef || "HEAD";

// Each slice is an independent concern in the diff. Reviewers read the real
// diff via `git diff` and the surrounding files, grounded in the repo's
// invariants (root CLAUDE.md + the relevant nested one).
const TARGETS = [
	{
		key: "tab-newline-emitter",
		title: "Formatting-preserve tab/newline emitter",
		files: [
			"src/core/track-changes/preserve-formatting.tsx",
			"src/core/edit/index.tsx",
			"src/cli/edit/index.tsx",
		],
		claude: ["src/core/track-changes/CLAUDE.md", "src/core/CLAUDE.md"],
		focus:
			"The change removes the `\\t`/`\\n` exclusion from canPreserveFormatting and adds runContentChildren/textSegment so a tab/newline inside `edit --text` splits into <w:tab/>/<w:br/> WITHIN each rPr-bearing <w:r> (instead of flattening the paragraph to one rPr-less run). Verify: (1) is the emitted OOXML schema-valid — does <w:tab/>/<w:br/> belong inside <w:r> next to <w:t>/<w:delText>, and in the right CT_R child order relative to rPr? (2) the DELETED-run path (tracked deletes) now also splits — is a <w:tab/> inside a <w:del>'d run correct, or should deleted tabs/breaks be represented differently? (3) edge cases: leading/trailing tab, consecutive tabs (\\t\\t), only-tab text, a literal \\r\\n (does split on /(\\n|\\t)/ leave a stray \\r in <w:t>?), empty segments. (4) does extractOldTokens see <w:tab/> at all — the diff aligns NEW text (with tabs) against OLD visible text where <w:tab/> contributes no character, so a kept line gets re-tokenized as insert/delete; confirm rPr inheritance still lands bold on the leading span and non-bold after the tab.",
	},
	{
		key: "edit-clear-combined",
		title: "edit --clear combined with content (single + batch + span)",
		files: ["src/cli/edit/index.tsx", "src/cli/edit/batch.ts", "src/core/edit/index.tsx"],
		claude: ["src/cli/CLAUDE.md", "src/core/CLAUDE.md"],
		focus:
			"`--clear` is now repeatable (string[]) and composes with content in one invocation (whole-paragraph AND span). Verify: clear-after-content ordering is correct for both commitBlockEdit (clears resultNode) and commitSpanEdit (clears [start, start+text.length]); the span window is right when the new text length differs from the old; batch entries apply clear correctly under the resolve-first invariant; CONTENT_KEYS vs clear separation can't silently drop an edit; invalid --clear tags are rejected cleanly. Look for a span whose clear range can run past the paragraph end after a shortening edit.",
	},
	{
		key: "delete-batch",
		title: "delete --batch (resolve-first, live refs)",
		files: ["src/cli/delete/index.tsx", "src/cli/delete/batch.ts"],
		claude: ["src/cli/CLAUDE.md", "src/core/CLAUDE.md"],
		focus:
			"New `delete --batch` removes many blocks in one read. The load-bearing invariant: all locators address the document AS READ. Verify: it resolves every entry to a LIVE XmlNode ref before any mutation, then splices by live indexOf so earlier deletions don't invalidate later locators; it rejects the unsupported shapes (ranges, sections, equations, spans) and the `--at`+`--batch` combination; deleting the same block twice or overlapping is handled; relationship/part pruning (the dangling-rId invariant) still happens for deleted images/hyperlinks if applicable. Check for a double-free / splice-after-detach.",
	},
	{
		key: "find-changes",
		title: "find: default-all, bare --highlight, no-match signal",
		files: ["src/cli/find/index.ts"],
		claude: ["src/cli/CLAUDE.md"],
		focus:
			"find now returns ALL matches by default (was first-only), bare `--highlight` matches ANY color, and a no-match run writes `no matches` to stderr. Verify: the default-all change didn't break the documented output contract (locator lines on stdout, the JSON shape under --json); `no matches` goes to STDERR not stdout (so `--json` consumers still get a clean array/empty result on stdout); bare --highlight vs --highlight=COLOR both work and don't collide; exit codes still mean what respond.ts says (0 ok even on zero matches? or 3 not-found?). Flag any inconsistency between find's exit/stream behavior and replace/read.",
	},
	{
		key: "respond-output-model",
		title: "respond.ts: success-confirmation text + dash-leading value merge",
		files: ["src/cli/respond.ts"],
		claude: ["src/cli/CLAUDE.md"],
		focus:
			"Two changes: (1) respondAck now prints a one-line `<operation> <target>` text confirmation by default via summarizeAck/ackTarget (was silent); (2) mergeDashLeadingValues in tryParseArgs merges `--flag value` into `--flag=value` when value is dash-led but not flag-shaped, so values like `-6` or `--del--` reach string flags. Verify HARD: does mergeDashLeadingValues ever swallow a REAL adjacent flag (e.g. `edit --text -v` where -v is meant as --verbose, or `--author -x`)? flagShaped treats any letter-after-dashes as a flag — so `--text -v` is preserved but `--text -6` merges; is that the intended boundary and does it break any negative-number-but-also-a-flag case? Confirm the text confirmation goes through the Bun.write sink (not process.stdout) and only on the non-verbose path, and that it can't fire for query commands or dry-run.",
	},
	{
		key: "render-word-mac",
		title: "render word-mac engine changes",
		files: ["src/core/render/engines/word-mac.ts"],
		claude: ["src/core/render/CLAUDE.md", "src/cli/render/CLAUDE.md"],
		focus:
			"Changes to the Word-for-Mac render engine (the osascript-driven single app instance). Verify: correctness of any serialization/locking (Word-mac is NOT concurrency-safe — a single app instance via osascript), temp-file cleanup, error propagation when Word isn't installed or the AppleScript fails, and that a render failure can't silently produce a 0-page or stale PDF. Check for a leaked Word process or a race if two renders overlap.",
	},
	{
		key: "help-docs",
		title: "help.ts + CLAUDE.md/README accuracy",
		files: ["src/cli/help.ts", "CLAUDE.md", "README.md"],
		claude: ["CLAUDE.md"],
		focus:
			"Top-level help was rewritten: a BATCH callout was added, the tracked-changes section condensed, and the VERIFY-LAYOUT guidance scoped to layout-only. Verify ACCURACY against the actual CLI: every flag/verb the help names exists and behaves as described (especially the `--batch` claim for edit/insert/replace/delete/comments, and the `delete --batch` one-liner); the render guidance doesn't contradict respond.ts behavior; no stale references to removed flags (--id/--range). Treat any help claim the code doesn't back as a finding.",
	},
];

const FINDINGS_SCHEMA = {
	type: "object",
	additionalProperties: false,
	properties: {
		findings: {
			type: "array",
			items: {
				type: "object",
				additionalProperties: false,
				properties: {
					severity: { type: "string", enum: ["blocker", "major", "minor", "nit"] },
					title: { type: "string" },
					file: { type: "string" },
					location: { type: "string", description: "line number or symbol" },
					issue: { type: "string", description: "what is wrong and why it matters" },
					evidence: { type: "string", description: "the specific code/behavior proving it" },
					repro: { type: "string", description: "a concrete command or input that would expose it, if any" },
					suggestedFix: { type: "string" },
				},
				required: ["severity", "title", "file", "issue", "suggestedFix"],
			},
		},
		summary: { type: "string" },
	},
	required: ["findings", "summary"],
};

const VERDICT_SCHEMA = {
	type: "object",
	additionalProperties: false,
	properties: {
		verdict: { type: "string", enum: ["confirmed", "refuted", "uncertain"] },
		correctedSeverity: {
			type: "string",
			enum: ["blocker", "major", "minor", "nit", "invalid"],
		},
		reasoning: { type: "string" },
		reproResult: { type: "string", description: "what happened when you tried to reproduce it" },
	},
	required: ["verdict", "correctedSeverity", "reasoning"],
};

phase("Review");
log(`Reviewing ${TARGETS.length} change slices against base ${BASE}`);

const reviewed = await pipeline(
	TARGETS,
	(target) =>
		agent(
			`You are a SKEPTICAL senior reviewer of uncommitted changes to docx-cli (a Bun CLI for editing .docx files). Your job is to find real defects, not to praise.

SLICE: ${target.title}

Read, in this order:
1. The repo invariants: the root CLAUDE.md, and these nested ones: ${target.claude.join(", ")}.
2. The actual diff for this slice: run \`git diff ${BASE} -- ${target.files.join(" ")}\` from the repo root.
3. The full current contents of the changed files (for context the diff doesn't show).

FOCUS ON: ${target.focus}

Ground rules:
- Only report defects in the CHANGED code (the diff), or pre-existing issues the change makes reachable. Do NOT report style nits the linter would catch (biome runs clean).
- For each finding give concrete evidence (quote the code) and, where possible, a repro: an exact \`dist/docx …\` command or input that would expose it. The binary at dist/docx is freshly built from this working tree — you MAY run it to check a hypothesis (use a /tmp copy of a fixture under .claude/skills/adversarial-review/fixtures/).
- Severity: blocker = corrupts a file / breaks the write-read loop / violates a stated invariant; major = wrong behavior on a realistic input; minor = wrong behavior on an edge case; nit = clarity/robustness.
- If the slice is clean, return an empty findings array and say so in summary. Do not invent findings.

Return structured findings.`,
			{ label: `review:${target.key}`, phase: "Review", schema: FINDINGS_SCHEMA },
		).then((review) => ({ target, review })),
	({ target, review }) => {
		const findings = (review?.findings ?? []).map((finding, i) => ({ ...finding, _i: i }));
		if (findings.length === 0) return { target: target.key, title: target.title, verified: [] };
		return parallel(
			findings.map((finding) => () =>
				agent(
					`You are an ADVERSARIAL verifier. A reviewer claimed a defect in docx-cli. Try HARD to REFUTE it — assume it is wrong until you can reproduce it. Default to "refuted" if you cannot concretely confirm it.

CLAIMED DEFECT (slice: ${target.title}):
- severity: ${finding.severity}
- title: ${finding.title}
- file: ${finding.file} ${finding.location ? `(${finding.location})` : ""}
- issue: ${finding.issue}
- evidence: ${finding.evidence ?? "(none given)"}
- repro: ${finding.repro ?? "(none given)"}

To verify: read the actual code (\`git diff ${BASE} -- ${finding.file}\` and the current file), and where a repro is given or constructable, RUN it against the freshly-built dist/docx binary on a /tmp copy of a fixture (fixtures live in .claude/skills/adversarial-review/fixtures/ and tests/fixtures/). Inspect the resulting document.xml (unzip -p FILE word/document.xml) if the claim is about emitted XML.

Decide:
- confirmed: you reproduced it or proved it from the code beyond doubt.
- refuted: the claim is wrong, already handled, or not reachable.
- uncertain: plausible but you couldn't prove it either way.
Set correctedSeverity to your assessment (or "invalid" if refuted). Explain with what you actually observed.`,
					{ label: `verify:${target.key}#${finding._i}`, phase: "Verify", schema: VERDICT_SCHEMA },
				).then((verdict) => ({ finding, verdict })),
			),
		).then((verified) => ({ target: target.key, title: target.title, verified: verified.filter(Boolean) }));
	},
);

phase("Synthesize");

const allVerified = reviewed.filter(Boolean).flatMap((slice) =>
	(slice.verified ?? []).map((v) => ({ slice: slice.title, ...v })),
);
const confirmed = allVerified.filter(
	(v) => v.verdict?.verdict === "confirmed" && v.verdict?.correctedSeverity !== "invalid",
);
const uncertain = allVerified.filter((v) => v.verdict?.verdict === "uncertain");

log(
	`${allVerified.length} findings; ${confirmed.length} confirmed, ${uncertain.length} uncertain, ${allVerified.length - confirmed.length - uncertain.length} refuted`,
);

const report = await agent(
	`You are writing the final adversarial code-review report for a body of uncommitted docx-cli changes. Below are the findings that survived adversarial verification (each was independently challenged; refuted ones are excluded).

CONFIRMED findings (JSON):
${JSON.stringify(confirmed.map((c) => ({ slice: c.slice, severity: c.verdict.correctedSeverity, finding: c.finding, reasoning: c.verdict.reasoning, reproResult: c.verdict.reproResult })), null, 2)}

UNCERTAIN findings (JSON):
${JSON.stringify(uncertain.map((c) => ({ slice: c.slice, finding: c.finding, reasoning: c.verdict.reasoning })), null, 2)}

Write a concise Markdown report:
1. **Verdict** — one line: is this body of work safe to commit as-is, or are there must-fix blockers/majors first?
2. **Must-fix** — confirmed blockers + majors, each: file:line, the problem, the fix. Ordered by severity. If none, say "none".
3. **Should-fix** — confirmed minors/nits worth doing.
4. **Worth a look** — uncertain findings the author should sanity-check.
Be specific and terse. Do not pad. If the changes are clean, say so plainly.`,
	{ label: "synthesize", phase: "Synthesize" },
);

return {
	base: BASE,
	counts: {
		total: allVerified.length,
		confirmed: confirmed.length,
		uncertain: uncertain.length,
		refuted: allVerified.length - confirmed.length - uncertain.length,
	},
	confirmed,
	uncertain,
	report,
};
