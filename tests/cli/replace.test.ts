import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

type Body = {
	blocks: Array<{
		id: string;
		type: string;
		runs?: Array<{
			type: string;
			text: string;
			bold?: boolean;
			color?: string;
		}>;
	}>;
};

describe("docx replace", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("replace");
		docPath = join(workspace, "out.docx");
		await runCli(
			"create",
			docPath,
			"--text",
			"The quick brown fox jumps over the lazy dog.",
		);
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Another fox sneaks by, then a Fox departs.",
		);
	});

	test("default replaces first match only", async () => {
		const result = await runCli("replace", docPath, "fox", "cat");
		expect(result.parsed).toMatchObject({
			ok: true,
			totalMatches: 2,
			replaced: 1,
		});
		const read = await runCli("read", docPath, "--ast");
		const text = (read.parsed as Body).blocks
			.flatMap((block) => block.runs ?? [])
			.filter((run) => run.type === "text")
			.map((run) => run.text)
			.join("");
		expect(text).toContain("brown cat jumps");
		expect(text).toContain("Another fox sneaks");
	});

	test("--all replaces every match", async () => {
		await runCli("replace", docPath, "fox", "cat", "--all");
		const read = await runCli("read", docPath, "--ast");
		const text = (read.parsed as Body).blocks
			.flatMap((block) => block.runs ?? [])
			.filter((run) => run.type === "text")
			.map((run) => run.text)
			.join("");
		expect(text).toContain("brown cat jumps");
		expect(text).toContain("Another cat sneaks");
		expect(text).not.toContain("fox");
	});

	test("--ignore-case catches mixed case across runs", async () => {
		await runCli("replace", docPath, "fox", "WOLF", "--all", "--ignore-case");
		const read = await runCli("read", docPath, "--ast");
		const text = (read.parsed as Body).blocks
			.flatMap((block) => block.runs ?? [])
			.filter((run) => run.type === "text")
			.map((run) => run.text)
			.join("");
		expect(text).toContain("brown WOLF jumps");
		expect(text).toContain("Another WOLF sneaks");
		expect(text).toContain("then a WOLF departs");
	});

	test("--regex with capture-group backrefs", async () => {
		await runCli("replace", docPath, "(quick) (brown)", "$2 $1", "--regex");
		const read = await runCli("read", docPath, "--ast");
		const text = (read.parsed as Body).blocks
			.flatMap((block) => block.runs ?? [])
			.filter((run) => run.type === "text")
			.map((run) => run.text)
			.join("");
		expect(text).toContain("The brown quick fox");
	});

	test("--regex with $& full-match reference", async () => {
		await runCli("replace", docPath, "fox", "[$&]", "--regex", "--all");
		const read = await runCli("read", docPath, "--ast");
		const text = (read.parsed as Body).blocks
			.flatMap((block) => block.runs ?? [])
			.filter((run) => run.type === "text")
			.map((run) => run.text)
			.join("");
		expect(text).toContain("[fox]");
		expect(text).toContain("Fox departs"); // case-sensitive: capital Fox left untouched
		expect(text).not.toContain("[Fox]");
	});

	test("--limit caps at N matches", async () => {
		const result = await runCli(
			"replace",
			docPath,
			"the",
			"THE",
			"--regex",
			"--ignore-case",
			"--limit",
			"2",
		);
		const payload = result.parsed as { totalMatches: number; replaced: number };
		expect(payload.totalMatches).toBeGreaterThanOrEqual(3);
		expect(payload.replaced).toBe(2);
	});

	test("--dry-run does not modify the file", async () => {
		const before = await Bun.file(docPath).arrayBuffer();
		await runCli("replace", docPath, "fox", "cat", "--all", "--dry-run");
		const after = await Bun.file(docPath).arrayBuffer();
		expect(after.byteLength).toBe(before.byteLength);
	});

	test("zero matches returns ok with replaced: 0", async () => {
		const result = await runCli("replace", docPath, "absent", "x");
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			ok: true,
			totalMatches: 0,
			replaced: 0,
		});
	});

	test("preserves rPr on surrounding text", async () => {
		const colorWorkspace = tempWorkspace("replace-color");
		const colorPath = join(colorWorkspace, "color.docx");
		await Bun.write(colorPath, Bun.file("tests/fixtures/minimal.docx"));
		// minimal.docx has p1 = "Use important terms in purple bold."
		// where "important" is purple+bold and the rest is unstyled.
		await runCli("replace", colorPath, "important", "essential");
		const read = await runCli("read", colorPath, "--ast");
		const paragraph = (read.parsed as Body).blocks.find(
			(block) => block.id === "p1",
		);
		const replacementRun = paragraph?.runs?.find(
			(run) => run.text === "essential",
		);
		expect(replacementRun?.color).toBe("800080");
		expect(replacementRun?.bold).toBe(true);
	});

	test("multiple matches in one paragraph apply in reverse order", async () => {
		const workspace = tempWorkspace("replace-multi");
		const multiPath = join(workspace, "multi.docx");
		await runCli("create", multiPath, "--text", "abcabcabc");
		await runCli("replace", multiPath, "abc", "Z", "--all");
		const read = await runCli("read", multiPath, "--ast");
		const text = (read.parsed as Body).blocks
			.flatMap((block) => block.runs ?? [])
			.filter((run) => run.type === "text")
			.map((run) => run.text)
			.join("");
		expect(text).toBe("ZZZ");
	});

	test("invalid --limit returns USAGE", async () => {
		const result = await runCli(
			"replace",
			docPath,
			"fox",
			"cat",
			"--limit",
			"-1",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("invalid regex returns USAGE", async () => {
		const result = await runCli(
			"replace",
			docPath,
			"(unclosed",
			"x",
			"--regex",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("replaces text inside table cells", async () => {
		const workspace = tempWorkspace("replace-cells");
		const cellPath = join(workspace, "cells.docx");
		await Bun.write(cellPath, Bun.file("tests/fixtures/tables-and-lists.docx"));

		const result = await runCli(
			"replace",
			cellPath,
			"Breadboard",
			"Protoboard",
			"--all",
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as {
			replaced: number;
			matches: Array<{ blockId: string }>;
		};
		expect(payload.replaced).toBeGreaterThanOrEqual(2);
		expect(
			payload.matches.some((match) => match.blockId.startsWith("t0:r")),
		).toBe(true);

		const read = await runCli("read", cellPath, "--ast");
		const text = JSON.stringify(read.parsed);
		expect(text).toContain("Protoboard");
		expect(text).not.toContain("Breadboard");
	});
});

const NORMALIZE_FIXTURE = "tests/fixtures/normalize-query.docx";
// Layout (built by tests/fixtures/setup/normalize-query.ts):
//   p0: 'The plan: "hello" world—ready to ship. The figure: 5 * 3 = 15.'
//       (smart quotes around hello)
//   p1: 'plan: "hello" today.' (straight quotes)

async function paragraphText(
	docPath: string,
	blockId: string,
): Promise<string> {
	const read = await runCli("read", docPath, "--ast");
	const blocks = (
		read.parsed as {
			blocks: Array<{
				id: string;
				runs?: Array<{ type: string; text: string }>;
			}>;
		}
	).blocks;
	const block = blocks.find((candidate) => candidate.id === blockId);
	return (block?.runs ?? [])
		.filter((run) => run.type === "text")
		.map((run) => run.text)
		.join("");
}

describe("docx replace — pattern normalization", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("replace-norm");
		docPath = join(workspace, "out.docx");
		await Bun.write(docPath, Bun.file(NORMALIZE_FIXTURE));
	});

	test("strips markdown emphasis from the pattern; replacement is literal", async () => {
		// Default: replace just the first match (p0's smart-quote "hello").
		const result = await runCli("replace", docPath, "**hello**", "goodbye");
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as {
			normalizedPattern?: string;
			normalizationApplied?: string[];
		};
		expect(payload.normalizedPattern).toBe("hello");
		expect(payload.normalizationApplied).toContain("strip-md-emphasis");

		// p0's surrounding smart quotes are preserved (they weren't part of
		// the matched span); "goodbye" replaces just "hello".
		expect(await paragraphText(docPath, "p0")).toContain("“goodbye”");
	});

	test("smart-quote pattern matches straight-quote document text via canonicalization", async () => {
		// Smart-quote pattern with --all hits both p0 (smart in doc) and
		// p1 (straight in doc) thanks to canonicalization.
		const result = await runCli(
			"replace",
			docPath,
			"“hello”", // smart quotes in the pattern.
			"goodbye",
			"--all",
		);
		expect(result.exitCode).toBe(0);

		// Replacement is LITERAL: the matched span (smart quote + hello +
		// smart quote in p0; straight quote + hello + straight quote in p1)
		// is replaced wholesale by the literal "goodbye". Surrounding
		// punctuation is preserved.
		expect(await paragraphText(docPath, "p0")).toBe(
			"The plan: goodbye world—ready to ship. The figure: 5 * 3 = 15.",
		);
		expect(await paragraphText(docPath, "p1")).toBe("plan: goodbye today.");
	});

	test("--exact disables pattern normalization", async () => {
		const result = await runCli(
			"replace",
			docPath,
			"**hello**",
			"goodbye",
			"--exact",
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as {
			totalMatches: number;
			replaced: number;
			normalizedPattern?: string;
		};
		expect(payload.totalMatches).toBe(0);
		expect(payload.replaced).toBe(0);
		expect(payload.normalizedPattern).toBeUndefined();
		// Both paragraphs unchanged.
		expect(await paragraphText(docPath, "p0")).toContain("“hello”");
		expect(await paragraphText(docPath, "p1")).toContain('"hello"');
	});
});

// Repro of the agent-feedback case: under track-changes ON, two consecutive
// replace calls in the same paragraph used to corrupt offsets — the second
// match's start was computed against a string that included the first
// replace's <w:ins>, so the splice landed mid-word inside the inserted run.
//
// The default (accepted) view fix: replace's offsets ignore the just-emitted
// <w:ins> and existing <w:del> wrappers, so chained edits stay safe.

const CHAINED_FIXTURE = "tests/fixtures/chained-tracked-edits.docx";
// Layout (built by tests/fixtures/setup/chained-tracked-edits.ts):
//   p0: "Cost of living, anti-price-gouging, and housing reform."
//   p1: "Old plan: ship Tuesday."
//   track-changes: ON, no tracked changes recorded yet.

describe("docx replace — chained edits under tracking", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("chained-replace");
		docPath = join(workspace, "out.docx");
		await Bun.write(docPath, Bun.file(CHAINED_FIXTURE));
	});

	test("two replaces in the same paragraph keep offsets stable in accepted view", async () => {
		const first = await runCli(
			"replace",
			docPath,
			"Cost of living",
			"Affordability",
		);
		expect(first.exitCode).toBe(0);

		// The second pattern is a phrase to the right of the first edit.
		// In the buggy version, the second replace's offset would be computed
		// against a haystack that included the just-inserted "Affordability"
		// AND the still-present <w:del>"Cost of living", landing the splice
		// inside the <w:ins>. With the accepted-view fix, neither the
		// pre-existing <w:del> nor the new <w:ins> shifts subsequent offsets.
		const second = await runCli(
			"replace",
			docPath,
			"anti-price-gouging",
			"price control",
		);
		expect(second.exitCode).toBe(0);

		// Read the accepted view of just p0 — under accepted view the
		// <w:del>s are dropped and the <w:ins>s are inlined as plain text.
		const result = await runCli("read", docPath, "--from", "p0", "--to", "p0");
		expect(result.exitCode).toBe(0);
		const accepted = result.stdout
			.split("\n")
			.map((line) => line.replace(/\s*<!--\s*[a-z0-9]+\s*-->\s*$/, ""))
			// Drop the head `<!-- docx:track-changes on -->` orientation hint (this
			// fixture has tracking on); it's not part of the paragraph text.
			.filter((line) => !/^<!--\s*docx:[^>]*-->$/.test(line.trim()))
			.join("\n")
			.trim();
		expect(accepted).toBe("Affordability, price control, and housing reform.");
	});

	test("--current view (legacy behavior) sees the raw concatenation", async () => {
		// Use p1 of the fixture: "Old plan: ship Tuesday."
		await runCli("replace", docPath, "Old", "New");

		// In --current view, find sees both ins and del text, so the next
		// query against "Old" still matches the deleted run.
		const find = await runCli("find", docPath, "Old", "--current");
		const payload = find.parsed as {
			matches: Array<{
				blockId: string;
				trackedChanges?: Array<{ kind: string }>;
			}>;
		};
		expect(payload.matches).toHaveLength(1);
		expect(payload.matches[0]?.blockId).toBe("p1");
		expect(payload.matches[0]?.trackedChanges?.[0]?.kind).toBe("del");

		// The default (accepted) view, by contrast, no longer sees "Old".
		const findDefault = await runCli("find", docPath, "Old");
		const defaultPayload = findDefault.parsed as { matches: unknown[] };
		expect(defaultPayload.matches).toEqual([]);

		// And the accepted view of p1 reads cleanly.
		expect(await paragraphText(docPath, "p1")).toContain("New plan");
	});
});

// replace chooses which tracked view the PATTERN matches against. The default
// (accepted) view can't see deleted text; --baseline can. This is the only path
// that substitutes text living inside a <w:del>.
describe("docx replace — view selection (--baseline / --current)", () => {
	async function trackedDeletionDoc(label: string): Promise<string> {
		const path = join(tempWorkspace(label), "out.docx");
		await runCli("create", path, "--text", "The quick brown fox jumps.");
		await runCli("track-changes", path, "on");
		await runCli("replace", path, "quick ", ""); // tracked-delete "quick "
		return path;
	}

	test("--baseline matches text that lives only inside <w:del>", async () => {
		const path = await trackedDeletionDoc("replace-baseline");

		// The accepted (default) view no longer sees the deleted word.
		const accepted = await runCli("replace", path, "quick", "QUICK");
		expect((accepted.parsed as { totalMatches: number }).totalMatches).toBe(0);

		// The baseline view matches the deleted text and substitutes it.
		const baseline = await runCli(
			"replace",
			path,
			"quick",
			"QUICK",
			"--baseline",
		);
		const payload = baseline.parsed as {
			view: string;
			totalMatches: number;
			replaced: number;
		};
		expect(payload.view).toBe("baseline");
		expect(payload.totalMatches).toBe(1);
		expect(payload.replaced).toBe(1);
	});

	test("--current and --baseline together are a USAGE error", async () => {
		const path = await trackedDeletionDoc("replace-view-mutex");
		const result = await runCli(
			"replace",
			path,
			"a",
			"b",
			"--current",
			"--baseline",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});
});

// `--at LOCATOR` confines a replace to one paragraph — the résumé fix for a
// placeholder that repeats across entries (`City, State` in every one), so a
// bare first-match replace can't safely target THE one being filled.
describe("docx replace — --at paragraph scope", () => {
	/** Three paragraphs that each contain the same "City, State" placeholder. */
	async function repeatedPlaceholder(label: string): Promise<string> {
		const path = join(tempWorkspace(label), "doc.docx");
		await runCli("create", path, "--text", "Resume header");
		await runCli("insert", path, "--after", "p0", "--text", "City, State one");
		await runCli("insert", path, "--after", "p1", "--text", "City, State two");
		await runCli(
			"insert",
			path,
			"--after",
			"p2",
			"--text",
			"City, State three",
		);
		return path;
	}

	test("replaces only the match in the scoped paragraph", async () => {
		const path = await repeatedPlaceholder("at-scope");
		const result = await runCli(
			"replace",
			path,
			"--at",
			"p2",
			"City, State",
			"Boston, MA",
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as {
			at: string;
			totalMatches: number;
			replaced: number;
			matches: Array<{ blockId: string }>;
		};
		// Scope cut the 3 doc-wide matches down to the one in p2.
		expect(payload.at).toBe("p2");
		expect(payload.totalMatches).toBe(1);
		expect(payload.replaced).toBe(1);
		expect(payload.matches[0]?.blockId).toBe("p2");

		// p1 and p3 still hold the placeholder; only p2 changed.
		const remaining = await runCli("find", path, "City, State", "--json");
		expect(
			(remaining.parsed as { matches: Array<{ blockId: string }> }).matches.map(
				(match) => match.blockId,
			),
		).toEqual(["p1", "p3"]);
	});

	test("a nonexistent scope paragraph is BLOCK_NOT_FOUND", async () => {
		const path = await repeatedPlaceholder("at-missing");
		const result = await runCli(
			"replace",
			path,
			"--at",
			"p99",
			"City, State",
			"X",
		);
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({ code: "BLOCK_NOT_FOUND" });
	});

	test("a range or span scope is rejected (single paragraph only)", async () => {
		const path = await repeatedPlaceholder("at-range");
		for (const bad of ["p1-p3", "p1:0-5"]) {
			const result = await runCli(
				"replace",
				path,
				"--at",
				bad,
				"City, State",
				"X",
			);
			expect(result.parsed).toMatchObject({ code: "INVALID_LOCATOR" });
		}
	});

	test("a nested-cell paragraph locator passes shape validation (not INVALID_LOCATOR)", async () => {
		// tT:rRcC:tU:rVcW:pN (a paragraph in a nested table cell) is a valid
		// paragraph the rest of the locator system addresses. The shape predicate
		// must accept it — a missing one errors BLOCK_NOT_FOUND (existence), NOT
		// INVALID_LOCATOR (shape). Regression: the predicate once required exactly
		// one cell-nesting level.
		const path = await repeatedPlaceholder("at-nested");
		const result = await runCli(
			"replace",
			path,
			"--at",
			"t0:r0c0:t1:r0c0:p0",
			"City, State",
			"X",
		);
		expect(result.parsed).toMatchObject({ code: "BLOCK_NOT_FOUND" });
		expect((result.parsed as { code: string }).code).not.toBe(
			"INVALID_LOCATOR",
		);
	});

	test("--at on a table cell paragraph scopes to that cell", async () => {
		const path = join(tempWorkspace("at-cell"), "cells.docx");
		await Bun.write(path, Bun.file("tests/fixtures/tables-and-lists.docx"));
		// Find a cell-paragraph match to scope to.
		const found = await runCli("find", path, "Breadboard", "--json");
		const cellMatch = (
			found.parsed as { matches: Array<{ blockId: string }> }
		).matches.find((match) => match.blockId.startsWith("t0:r"));
		expect(cellMatch).toBeDefined();
		const cellId = cellMatch?.blockId as string;
		const result = await runCli(
			"replace",
			path,
			"--at",
			cellId,
			"Breadboard",
			"Protoboard",
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as {
			matches: Array<{ blockId: string }>;
		};
		expect(payload.matches.every((match) => match.blockId === cellId)).toBe(
			true,
		);
	});

	test("--at is rejected alongside --batch (scope is per-entry there)", async () => {
		const path = await repeatedPlaceholder("at-batch-conflict");
		const batchPath = join(tempWorkspace("at-batch-conflict-jsonl"), "b.jsonl");
		await Bun.write(batchPath, '{"pattern":"City, State","replacement":"X"}\n');
		const result = await runCli(
			"replace",
			path,
			"--batch",
			batchPath,
			"--at",
			"p1",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("batch entries carry their own at, filling distinct paragraphs in one call", async () => {
		const path = await repeatedPlaceholder("at-batch");
		const batchPath = join(tempWorkspace("at-batch-jsonl"), "fill.jsonl");
		await Bun.write(
			batchPath,
			`${[
				'{"at":"p1","pattern":"City, State","replacement":"Boston, MA"}',
				'{"at":"p3","pattern":"City, State","replacement":"Austin, TX"}',
			].join("\n")}\n`,
		);
		const result = await runCli("replace", path, "--batch", batchPath);
		expect(result.exitCode).toBe(0);

		// p1 and p3 filled distinctly; p2 left untouched.
		const remaining = await runCli("find", path, "City, State", "--json");
		expect(
			(remaining.parsed as { matches: Array<{ blockId: string }> }).matches.map(
				(match) => match.blockId,
			),
		).toEqual(["p2"]);
		const boston = await runCli("find", path, "Boston, MA", "--json");
		expect(
			(boston.parsed as { matches: Array<{ blockId: string }> }).matches[0]
				?.blockId,
		).toBe("p1");
	});

	test("a batch entry with a malformed at scope is a per-entry error", async () => {
		const path = await repeatedPlaceholder("at-batch-bad");
		const batchPath = join(tempWorkspace("at-batch-bad-jsonl"), "b.jsonl");
		await Bun.write(
			batchPath,
			'{"at":"p1-p3","pattern":"City, State","replacement":"X"}\n',
		);
		const result = await runCli("replace", path, "--batch", batchPath);
		expect(result.parsed).toMatchObject({ code: "INVALID_LOCATOR" });
	});

	test("a batch entry with a parseable-but-nonexistent at errors (not a silent no-op)", async () => {
		// The single-shot path errors BLOCK_NOT_FOUND on a typo'd scope; the batch
		// path must too, else a fat-fingered `at` matches nothing, mutates nothing,
		// and falsely reports success — the exact write→read-loop trap.
		const path = await repeatedPlaceholder("at-batch-missing");
		const before = await runCli("find", path, "City, State", "--json");
		const batchPath = join(tempWorkspace("at-batch-missing-jsonl"), "b.jsonl");
		await Bun.write(
			batchPath,
			'{"at":"p99","pattern":"City, State","replacement":"X"}\n',
		);
		const result = await runCli("replace", path, "--batch", batchPath);
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({ code: "BLOCK_NOT_FOUND" });
		// Nothing mutated — the document is untouched.
		const after = await runCli("find", path, "City, State", "--json");
		expect((after.parsed as { matches: unknown[] }).matches.length).toBe(
			(before.parsed as { matches: unknown[] }).matches.length,
		);
	});
});
