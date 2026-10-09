import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";
import {
	buildRawDoc,
	readDocumentXml,
	wordTextBoxParagraphXml,
} from "./helpers";

type Body = {
	blocks: Array<{
		id: string;
		type: string;
		runs?: Array<{
			type: string;
			text: string;
			bold?: boolean;
			italic?: boolean;
			color?: string;
			highlight?: string;
			font?: string;
			sizeHalfPoints?: number;
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

	test("zero matches exits nonzero (MATCH_NOT_FOUND) — no silent no-op", async () => {
		const result = await runCli("replace", docPath, "absent", "x");
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({ code: "MATCH_NOT_FOUND" });
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

describe("docx replace — replacement formatting", () => {
	async function replacementRun(
		path: string,
		text: string,
	): Promise<NonNullable<Body["blocks"][number]["runs"]>[number] | undefined> {
		const read = await runCli("read", path, "--ast");
		return (read.parsed as Body).blocks
			.flatMap((block) => block.runs ?? [])
			.find((run) => run.text === text);
	}

	test("format flags and --clear apply to only the replacement runs", async () => {
		const path = join(tempWorkspace("replace-format"), "out.docx");
		await runCli("create", path, "--text", "TODO remains");
		await runCli(
			"edit",
			path,
			"--at",
			"p0:0-4",
			"--bold",
			"--highlight",
			"yellow",
			"--font",
			"Courier New",
			"--size",
			"9",
		);

		const result = await runCli(
			"replace",
			path,
			"TODO",
			"Done",
			"--clear",
			"bold,highlight",
			"--italic",
			"--color",
			"00AA00",
			"--size",
			"12",
		);
		expect(result.exitCode).toBe(0);
		const run = await replacementRun(path, "Done");
		expect(run).toMatchObject({
			italic: true,
			color: "00AA00",
			font: "Courier New",
			sizeHalfPoints: 24,
		});
		expect(run?.bold).toBeUndefined();
		expect(run?.highlight).toBeUndefined();
	});

	test("batch entries default to first and opt into every match with all", async () => {
		const firstPath = join(tempWorkspace("replace-batch-first"), "out.docx");
		await runCli("create", firstPath, "--text", "TODO TODO");
		const firstBatch = join(
			tempWorkspace("replace-batch-first-jsonl"),
			"batch.jsonl",
		);
		await Bun.write(
			firstBatch,
			'{"pattern":"TODO","replacement":"Done","bold":true}\n',
		);
		const first = await runCli("replace", firstPath, "--batch", firstBatch);
		expect(first.parsed).toMatchObject({
			batch: [{ totalMatches: 2, replaced: 1 }],
		});
		expect(await paragraphText(firstPath, "p0")).toBe("Done TODO");
		expect((await replacementRun(firstPath, "Done"))?.bold).toBe(true);

		const allPath = join(tempWorkspace("replace-batch-all"), "out.docx");
		await runCli("create", allPath, "--text", "TODO TODO");
		const allBatch = join(
			tempWorkspace("replace-batch-all-jsonl"),
			"batch.jsonl",
		);
		await Bun.write(
			allBatch,
			'{"pattern":"TODO","replacement":"Done","italic":true,"all":true}\n',
		);
		const all = await runCli("replace", allPath, "--batch", allBatch);
		expect(all.parsed).toMatchObject({
			batch: [{ totalMatches: 2, replaced: 2 }],
		});
		expect(await paragraphText(allPath, "p0")).toBe("Done Done");
		const read = await runCli("read", allPath, "--ast");
		const doneRuns = (read.parsed as Body).blocks
			.flatMap((block) => block.runs ?? [])
			.filter((run) => run.text.includes("Done"));
		expect(doneRuns.every((run) => run.italic)).toBe(true);
	});

	test("formatting applies to every segment of a multi-line replacement", async () => {
		const path = join(tempWorkspace("replace-format-lines"), "out.docx");
		await runCli("create", path, "--text", "alpha beta gamma");
		const result = await runCli(
			"replace",
			path,
			"beta",
			"one\ntwo",
			"--bold",
			"--color",
			"C00000",
		);
		expect(result.exitCode).toBe(0);
		expect(await replacementRun(path, "one")).toMatchObject({
			bold: true,
			color: "C00000",
		});
		expect(await replacementRun(path, "two")).toMatchObject({
			bold: true,
			color: "C00000",
		});
	});

	test("tracked replacement keeps formatting on the inserted text", async () => {
		const path = join(tempWorkspace("replace-format-track"), "out.docx");
		await runCli("create", path, "--text", "TODO");
		await runCli("track-changes", path, "on");
		const result = await runCli(
			"replace",
			path,
			"TODO",
			"Done",
			"--bold",
			"--color",
			"007700",
		);
		expect(result.exitCode).toBe(0);
		expect(await replacementRun(path, "Done")).toMatchObject({
			bold: true,
			color: "007700",
		});
		const xml = await readDocumentXml(path);
		expect(xml).toMatch(/<w:ins[^>]*>.*?<w:rPr>.*?<w:b\s*\/>/s);
	});

	test("invalid single-shot and batch formatting fail with USAGE", async () => {
		const path = join(tempWorkspace("replace-format-invalid"), "out.docx");
		await runCli("create", path, "--text", "TODO");
		const single = await runCli("replace", path, "TODO", "Done", "--size", "0");
		expect(single.parsed).toMatchObject({ code: "USAGE" });

		const batchPath = join(
			tempWorkspace("replace-format-invalid-jsonl"),
			"batch.jsonl",
		);
		await Bun.write(
			batchPath,
			'{"pattern":"TODO","replacement":"Done","color":42}\n',
		);
		const batch = await runCli("replace", path, "--batch", batchPath);
		expect(batch.parsed).toMatchObject({ code: "USAGE" });

		const booleanPath = join(
			tempWorkspace("replace-format-boolean-jsonl"),
			"batch.jsonl",
		);
		await Bun.write(
			booleanPath,
			'{"pattern":"TODO","replacement":"Done","all":"false","bold":"false"}\n',
		);
		const boolean = await runCli("replace", path, "--batch", booleanPath);
		expect(boolean.parsed).toMatchObject({ code: "USAGE" });

		const malformedSizePath = join(
			tempWorkspace("replace-format-size-jsonl"),
			"batch.jsonl",
		);
		await Bun.write(
			malformedSizePath,
			'{"pattern":"TODO","replacement":"Done","size":"12garbage"}\n',
		);
		const malformedSize = await runCli(
			"replace",
			path,
			"--batch",
			malformedSizePath,
		);
		expect(malformedSize.parsed).toMatchObject({ code: "USAGE" });

		const incompleteUnderlinePath = join(
			tempWorkspace("replace-format-underline-jsonl"),
			"batch.jsonl",
		);
		await Bun.write(
			incompleteUnderlinePath,
			'{"pattern":"TODO","replacement":"Done","underlineColor":"FF0000"}\n',
		);
		const incompleteUnderline = await runCli(
			"replace",
			path,
			"--batch",
			incompleteUnderlinePath,
		);
		expect(incompleteUnderline.parsed).toMatchObject({ code: "USAGE" });

		const limitPath = join(
			tempWorkspace("replace-format-limit-jsonl"),
			"batch.jsonl",
		);
		await Bun.write(
			limitPath,
			'{"pattern":"TODO","replacement":"Done","limit":true}\n',
		);
		const limit = await runCli("replace", path, "--batch", limitPath);
		expect(limit.parsed).toMatchObject({ code: "USAGE" });
		expect(await paragraphText(path, "p0")).toBe("TODO");
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
		// --exact keeps "**hello**" literal, so it matches neither smart- nor
		// straight-quoted "hello" → 0 matches, which now exits nonzero (no silent
		// no-op) rather than a cheerful replaced:0.
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({ code: "MATCH_NOT_FOUND" });
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

		// The accepted (default) view no longer sees the deleted word → 0 matches,
		// which now exits nonzero (MATCH_NOT_FOUND) instead of a silent no-op.
		const accepted = await runCli("replace", path, "quick", "QUICK");
		expect(accepted.exitCode).toBe(3);
		expect(accepted.parsed).toMatchObject({ code: "MATCH_NOT_FOUND" });

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

	test("a batch entry that matches nothing exits nonzero (no silent no-op)", async () => {
		// A pattern that isn't literal document text matches 0. The batch must exit
		// nonzero (MATCH_NOT_FOUND) even though other entries applied — else the agent
		// reads the clean batch ack as total success and ships a partial result.
		const path = await repeatedPlaceholder("batch-noop");
		const batchPath = join(tempWorkspace("batch-noop-jsonl"), "b.jsonl");
		await Bun.write(
			batchPath,
			'{"pattern":"City, State","replacement":"Denver, CO","all":true}\n{"pattern":"<mark>absent</mark>","replacement":"x"}\n',
		);
		const result = await runCli("replace", path, "--batch", batchPath);
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({ code: "MATCH_NOT_FOUND" });
		// The matching entry still applied + saved (sed-like); only the miss is flagged.
		const after = await runCli("find", path, "Denver, CO", "--json");
		expect(
			(after.parsed as { matches: unknown[] }).matches.length,
		).toBeGreaterThan(0);
	});
});

// Tabs are one offset character (they render as "\t" in read), so a
// tab-separated line — the classic résumé placeholder — is fillable by a
// space-typed pattern, and a cut that spans a tab tracks it honestly.
describe("docx replace — tab-separated lines", () => {
	test("fills a TAB-separated line from a space-typed pattern", async () => {
		const path = join(tempWorkspace("replace-tab"), "out.docx");
		await runCli("create", path, "--text", "City\tState\tZip");
		const result = await runCli(
			"replace",
			path,
			"City State Zip",
			"Austin, TX",
		);
		expect(result.exitCode).toBe(0);
		expect(await paragraphText(path, "p0")).toContain("Austin, TX");
	});

	test("a tab inside a tracked deletion is wrapped in <w:del>, not left behind", async () => {
		const path = join(tempWorkspace("replace-tab-track"), "out.docx");
		await runCli("create", path, "--text", "aa\tbb ZZ");
		await runCli("track-changes", path, "on");
		// The space-typed pattern spans "aa", the tab, and "bb" — all three land in
		// the deletion (the tab is not a zero-width straggler outside the <w:del>).
		const result = await runCli("replace", path, "aa bb", "");
		expect(result.exitCode).toBe(0);
		const xml = await readDocumentXml(path);
		expect(xml).toContain("<w:del");
		expect(xml).toMatch(
			/<w:del[^>]*>\s*<w:r>\s*<w:tab\s*\/>\s*<\/w:r>\s*<\/w:del>/,
		);
	});

	test("a later replace stays aligned across a hidden deleted tab", async () => {
		const path = join(tempWorkspace("replace-tab-align"), "out.docx");
		await runCli("create", path, "--text", "aa\tbb ZZ");
		await runCli("track-changes", path, "on");
		await runCli("replace", path, "aa bb", ""); // tracked-delete text containing a tab
		// In the accepted view the deleted tab is 0-width, so ZZ is where find/replace
		// expect it — the substitution must land, proving offsets didn't drift.
		const result = await runCli("replace", path, "ZZ", "ZZZ");
		expect(result.exitCode).toBe(0);
		expect(await paragraphText(path, "p0")).toContain("ZZZ");
	});
});

// Editor-style multi-line replace: a "\n" in the pattern matches a line break
// or a paragraph boundary; the replacement's newlines then define the resulting
// paragraph structure — exactly as if the span were selected in Word and the
// replacement typed. Untracked only (refuses under tracking).
describe("docx replace — across paragraphs (editor-style)", () => {
	async function threeParagraphs(label: string): Promise<string> {
		const path = join(tempWorkspace(label), "out.docx");
		await runCli("create", path, "--text", "AAA header");
		await runCli("insert", path, "--at-end", "--text", "BBB middle");
		await runCli("insert", path, "--at-end", "--text", "CCC footer");
		return path;
	}

	async function blockTexts(path: string): Promise<string[]> {
		const read = await runCli("read", path, "--ast");
		const blocks = (
			read.parsed as {
				blocks: Array<{
					type: string;
					runs?: Array<{ type: string; text: string }>;
				}>;
			}
		).blocks;
		return blocks
			.filter((block) => block.type === "paragraph")
			.map((block) =>
				(block.runs ?? [])
					.filter((run) => run.type === "text")
					.map((run) => run.text)
					.join(""),
			);
	}

	test("single-line replacement across a boundary MERGES the paragraphs", async () => {
		const path = await threeParagraphs("across-merge");
		const result = await runCli("replace", path, "header\nBBB", "joined");
		expect(result.exitCode).toBe(0);
		expect(await blockTexts(path)).toEqual(["AAA joined middle", "CCC footer"]);
	});

	test("a \\n in the replacement SPLITS a paragraph", async () => {
		const path = join(tempWorkspace("across-split"), "out.docx");
		await runCli("create", path, "--text", "alpha beta gamma");
		const result = await runCli("replace", path, "beta", "one\ntwo");
		expect(result.exitCode).toBe(0);
		expect(await blockTexts(path)).toEqual(["alpha one", "two gamma"]);
	});

	test("middle paragraphs die; the last keeps its own tail", async () => {
		const path = await threeParagraphs("across-middle");
		const result = await runCli(
			"replace",
			path,
			"header\nBBB middle\nCCC",
			"intro\nCCC",
		);
		expect(result.exitCode).toBe(0);
		expect(await blockTexts(path)).toEqual(["AAA intro", "CCC footer"]);
	});

	test("the first paragraph's style governs a merge; tail formatting survives", async () => {
		const path = join(tempWorkspace("across-style"), "out.docx");
		await runCli("create", path, "--text", "Heading Line");
		await runCli("edit", path, "--at", "p0", "--style", "Heading1");
		await runCli(
			"insert",
			path,
			"--after",
			"p0",
			"--markdown",
			"Body **bold**",
		);

		const result = await runCli("replace", path, "Line\nBody", "Joined");
		expect(result.exitCode).toBe(0);
		const read = await runCli("read", path, "--ast");
		const blocks = (
			read.parsed as {
				blocks: Array<{
					type: string;
					style?: string;
					runs?: Array<{ type: string; text: string; bold?: boolean }>;
				}>;
			}
		).blocks;
		const merged = blocks[0];
		expect(merged?.style).toBe("Heading1");
		const boldRun = merged?.runs?.find((run) => run.bold);
		expect(boldRun?.text).toBe("bold");
	});

	test("refuses under tracking instead of skipping the journal", async () => {
		const path = await threeParagraphs("across-tracked");
		await runCli("track-changes", path, "on");
		const result = await runCli("replace", path, "header\nBBB", "joined");
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
		// Nothing changed.
		expect(await blockTexts(path)).toEqual([
			"AAA header",
			"BBB middle",
			"CCC footer",
		]);
	});

	test("rejects --at (a multi-line pattern spans paragraphs)", async () => {
		const path = await threeParagraphs("across-at");
		const result = await runCli(
			"replace",
			path,
			"--at",
			"p0",
			"header\nBBB",
			"joined",
		);
		expect(result.exitCode).toBe(2);
	});

	test("0 cross-paragraph matches exits MATCH_NOT_FOUND", async () => {
		const path = await threeParagraphs("across-miss");
		const result = await runCli("replace", path, "header\nZZZ", "joined");
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({ code: "MATCH_NOT_FOUND" });
	});

	test("a batch entry with \\n merges too (JSONL newlines are real)", async () => {
		const path = await threeParagraphs("across-batch");
		const batchPath = join(tempWorkspace("across-batch-jsonl"), "b.jsonl");
		await Bun.write(
			batchPath,
			'{"pattern":"header\\nBBB","replacement":"joined"}\n',
		);
		const result = await runCli("replace", path, "--batch", batchPath);
		expect(result.exitCode).toBe(0);
		expect(await blockTexts(path)).toEqual(["AAA joined middle", "CCC footer"]);
	});

	test("a CRLF batch replacement splits cleanly (no stray \\r in the runs)", async () => {
		const path = await threeParagraphs("across-crlf");
		const batchPath = join(tempWorkspace("across-crlf-jsonl"), "b.jsonl");
		// JSONL authored on Windows: JSON.parse hands the across path a real
		// "\r\n" (inline argv would have been normalized by decodeInlineEscapes).
		await Bun.write(
			batchPath,
			'{"pattern":"BBB middle","replacement":"one\\r\\ntwo"}\n',
		);
		const result = await runCli("replace", path, "--batch", batchPath);
		expect(result.exitCode).toBe(0);
		expect(await blockTexts(path)).toEqual([
			"AAA header",
			"one",
			"two",
			"CCC footer",
		]);
	});

	test("a \\n pattern merges consecutive paragraphs INSIDE a table cell", async () => {
		const path = join(tempWorkspace("across-cell"), "out.docx");
		await runCli("create", path, "--text", "intro");
		await runCli(
			"tables",
			"create",
			path,
			"--at-end",
			"--rows",
			"1",
			"--cols",
			"1",
		);
		await runCli("edit", path, "--at", "t0:r0c0:p0", "--text", "alpha");
		await runCli("insert", path, "--after", "t0:r0c0:p0", "--text", "beta");

		const result = await runCli("replace", path, "alpha\nbeta", "joined");
		expect(result.exitCode).toBe(0);
		const read = await runCli("read", path, "--ast");
		const table = (
			read.parsed as {
				blocks: Array<{
					type: string;
					rows?: Array<{
						cells: Array<{
							blocks: Array<{
								type: string;
								runs?: Array<{ type: string; text: string }>;
							}>;
						}>;
					}>;
				}>;
			}
		).blocks.find((block) => block.type === "table");
		const cellTexts = (table?.rows?.[0]?.cells[0]?.blocks ?? [])
			.filter((block) => block.type === "paragraph")
			.map((block) =>
				(block.runs ?? [])
					.filter((run) => run.type === "text")
					.map((run) => run.text)
					.join(""),
			);
		expect(cellTexts).toEqual(["joined"]);
	});

	test('replace "\\n" with a space --all merges every paragraph', async () => {
		const path = await threeParagraphs("across-merge-all");
		const result = await runCli("replace", path, "\n", " ", "--all");
		expect(result.exitCode).toBe(0);
		expect(await blockTexts(path)).toEqual([
			"AAA header BBB middle CCC footer",
		]);
	});
});

test("replace and replace batch apply complex-script override only to replacement", async () => {
	const workspace = tempWorkspace("complex-script-replace");
	const docPath = join(workspace, "out.docx");
	expect(
		(await runCli("create", docPath, "--text", "Hello target world")).exitCode,
	).toBe(0);
	expect(
		(
			await runCli(
				"replace",
				docPath,
				"target",
				"مرحبا",
				"--font-complex-script",
				"Amiri",
				"--font",
				"Arial",
			)
		).exitCode,
	).toBe(0);
	const xml = await readDocumentXml(docPath);
	expect(xml).toContain('w:cs="Amiri"');
	expect(xml).toContain('w:ascii="Arial"');
	const batchPath = join(workspace, "batch.jsonl");
	await Bun.write(
		batchPath,
		JSON.stringify({
			pattern: "مرحبا",
			replacement: "أهلا",
			"font-complex-script": "Noto Naskh Arabic",
		}),
	);
	expect(
		(await runCli("replace", docPath, "--batch", batchPath)).exitCode,
	).toBe(0);
	expect(await readDocumentXml(docPath)).toContain('w:cs="Noto Naskh Arabic"');
	expect(
		(
			await runCli(
				"replace",
				docPath,
				"--batch",
				batchPath,
				"--font-complex-script",
				"Amiri",
			)
		).exitCode,
	).toBe(2);
});

// Issue #13: a tracked replace whose match lies WHOLLY inside another author's
// pending <w:ins>. Main put the replacement as a bare run inside that ins
// (credited to the other author) and, when the match sat in a later run of the
// ins (here: after a nested <w:del>), started its offset walk at that run's
// offset and cut the wrong words. Word's shape: the other author's ins keeps
// its id and the text before the match plus the editor's nested <w:del> of
// the cut; the replacement rides the editor's own top-level <w:ins>; the text
// after the match rides a copy of the ins with a fresh id. When the editor IS
// that author the ins is rebuilt in place (no split, no <w:del>).
describe("docx replace --track inside another author's insertion (#13)", () => {
	const INS_A =
		'<w:ins w:id="1" w:author="Reviewer A" w:date="2026-09-01T00:00:00Z">';
	const SINGLE_RUN = `<w:p><w:r><w:t xml:space="preserve">The Client </w:t></w:r>${INS_A}<w:r><w:t>may withhold a disputed amount.</w:t></w:r></w:ins></w:p>`;
	const NESTED_DEL = `<w:p><w:r><w:t xml:space="preserve">Fees are due. </w:t></w:r>${INS_A}<w:r><w:t xml:space="preserve">The Client may withhold </w:t></w:r><w:del w:id="2" w:author="Reviewer B" w:date="2026-09-02T00:00:00Z"><w:r><w:delText xml:space="preserve">any </w:delText></w:r></w:del><w:r><w:t>disputed amount.</w:t></w:r></w:ins></w:p>`;

	async function trackedReplace(
		body: string,
		label: string,
		author: string,
		pattern = "disputed amount",
		...flags: string[]
	): Promise<string> {
		const docPath = await buildRawDoc(body, label);
		const result = await runCli(
			"replace",
			docPath,
			pattern,
			"disputed sum",
			"--track",
			"--author",
			author,
			...flags,
		);
		expect(result.exitCode).toBe(0);
		// Every emitted shape must pass the ECMA-376 schema gate Word applies.
		expect((await runCli("validate", docPath)).exitCode).toBe(0);
		return docPath;
	}

	/** The first `<w:p>` of document.xml with the editor's (now) dates pinned,
	 *  so the whole revision structure can be asserted as one string. */
	async function paragraphXml(
		docPath: string,
		editor: string,
	): Promise<string> {
		const xml = await readDocumentXml(docPath);
		const paragraph = xml.match(/<w:p>.*?<\/w:p>/s)?.[0] ?? "";
		return paragraph.replaceAll(
			new RegExp(`(w:author="${editor}") w:date="[^"]*"`, "g"),
			'$1 w:date="NOW"',
		);
	}

	async function readPart(docPath: string, name: string): Promise<string> {
		return await (await Pkg.open(docPath)).readText(name);
	}

	async function viewText(docPath: string, view: string): Promise<string> {
		const result = await runCli("read", docPath, view);
		return result.stdout
			.split("\n")
			.filter((line) => line.trim() && !/^<!--.*-->$/.test(line.trim()))
			.map((line) => line.replace(/\s*<!--\s*p\d+\s*-->\s*$/, ""))
			.join("\n");
	}

	async function listChanges(
		docPath: string,
	): Promise<
		Array<{ id: string; kind: string; text: string; author: string }>
	> {
		const result = await runCli("track-changes", "list", docPath);
		return result.parsed as Array<{
			id: string;
			kind: string;
			text: string;
			author: string;
		}>;
	}

	test("case 1: the other author's ins is split around the editor's del + ins, tail gets a fresh id", async () => {
		const docPath = await trackedReplace(SINGLE_RUN, "ins-split", "Editor");
		expect(await paragraphXml(docPath, "Editor")).toBe(
			'<w:p><w:r><w:t xml:space="preserve">The Client </w:t></w:r>' +
				`${INS_A}<w:r><w:t xml:space="preserve">may withhold a </w:t></w:r>` +
				'<w:del w:id="2" w:author="Editor" w:date="NOW"><w:r><w:delText xml:space="preserve">disputed amount</w:delText></w:r></w:del></w:ins>' +
				'<w:ins w:id="3" w:author="Editor" w:date="NOW"><w:r><w:t xml:space="preserve">disputed sum</w:t></w:r></w:ins>' +
				'<w:ins w:id="4" w:author="Reviewer A" w:date="2026-09-01T00:00:00Z"><w:r><w:t xml:space="preserve">.</w:t></w:r></w:ins></w:p>',
		);
		expect(await viewText(docPath, "--accepted")).toBe(
			"The Client may withhold a disputed sum.",
		);

		const changes = await listChanges(docPath);
		expect(
			changes.map((change) => [change.kind, change.author, change.text]),
		).toEqual([
			["ins", "Reviewer A", "may withhold a "],
			["del", "Editor", "disputed amount"],
			["ins", "Editor", "disputed sum"],
			["ins", "Reviewer A", "."],
		]);
		// revN pairing is per author: the editor's del+ins is one group, and
		// Reviewer A's head ins never pairs with the editor's nested del.
		const grouped = changes as Array<{ author: string; group?: string }>;
		const editorGroups = grouped
			.filter((change) => change.author === "Editor")
			.map((change) => change.group);
		expect(editorGroups).toEqual(["rev0", "rev0"]);
		expect(
			grouped
				.filter((change) => change.author === "Reviewer A")
				.every((change) => change.group === undefined),
		).toBe(true);
	});

	test("case 2: a nested deletion by a third author stays put and does not shift the cut", async () => {
		const docPath = await trackedReplace(
			NESTED_DEL,
			"ins-nested-del",
			"Editor",
		);
		expect(await paragraphXml(docPath, "Editor")).toBe(
			'<w:p><w:r><w:t xml:space="preserve">Fees are due. </w:t></w:r>' +
				`${INS_A}<w:r><w:t xml:space="preserve">The Client may withhold </w:t></w:r>` +
				'<w:del w:id="2" w:author="Reviewer B" w:date="2026-09-02T00:00:00Z"><w:r><w:delText xml:space="preserve">any </w:delText></w:r></w:del>' +
				'<w:del w:id="3" w:author="Editor" w:date="NOW"><w:r><w:delText xml:space="preserve">disputed amount</w:delText></w:r></w:del></w:ins>' +
				'<w:ins w:id="4" w:author="Editor" w:date="NOW"><w:r><w:t xml:space="preserve">disputed sum</w:t></w:r></w:ins>' +
				'<w:ins w:id="5" w:author="Reviewer A" w:date="2026-09-01T00:00:00Z"><w:r><w:t xml:space="preserve">.</w:t></w:r></w:ins></w:p>',
		);
		expect(await viewText(docPath, "--accepted")).toBe(
			"Fees are due. The Client may withhold disputed sum.",
		);
		const changes = await listChanges(docPath);
		expect(changes.map((change) => [change.kind, change.author])).toEqual([
			["ins", "Reviewer A"],
			["del", "Reviewer B"],
			["del", "Editor"],
			["ins", "Editor"],
			["ins", "Reviewer A"],
		]);
	});

	test("case 3: the author's OWN ins is rebuilt in place — same id, no del, no split", async () => {
		const docPath = await trackedReplace(SINGLE_RUN, "ins-own", "Reviewer A");
		// Nothing to pin: no Editor revision exists, and A's date must survive.
		expect(await paragraphXml(docPath, "Editor")).toBe(
			'<w:p><w:r><w:t xml:space="preserve">The Client </w:t></w:r>' +
				'<w:ins w:id="1" w:author="Reviewer A" w:date="2026-09-01T00:00:00Z">' +
				'<w:r><w:t xml:space="preserve">may withhold a </w:t></w:r>' +
				'<w:r><w:t xml:space="preserve">disputed sum</w:t></w:r>' +
				'<w:r><w:t xml:space="preserve">.</w:t></w:r></w:ins></w:p>',
		);
		expect(await viewText(docPath, "--accepted")).toBe(
			"The Client may withhold a disputed sum.",
		);
		expect(await listChanges(docPath)).toHaveLength(1);
	});

	test("the author match is exact: 'Reviewer a' is another author", async () => {
		const docPath = await trackedReplace(SINGLE_RUN, "ins-case", "Reviewer a");
		const xml = await readDocumentXml(docPath);
		expect(xml).toContain('<w:del w:id="2" w:author="Reviewer a"');
		expect(xml).toContain('<w:ins w:id="3" w:author="Reviewer a"');
		expect(xml.match(/<w:ins /g)).toHaveLength(3);
	});

	test("rejecting the editor's revisions restores the other author's text intact", async () => {
		const docPath = await trackedReplace(SINGLE_RUN, "ins-reject", "Editor");
		const editorIds = (await listChanges(docPath))
			.filter((change) => change.author === "Editor")
			.map((change) => change.id);
		expect(editorIds).toHaveLength(2);
		const rejected = await runCli(
			"track-changes",
			"reject",
			docPath,
			...editorIds.flatMap((id) => ["--at", id]),
		);
		expect(rejected.exitCode).toBe(0);

		expect(await viewText(docPath, "--accepted")).toBe(
			"The Client may withhold a disputed amount.",
		);
		expect(await viewText(docPath, "--baseline")).toBe("The Client");
		const xml = await readDocumentXml(docPath);
		expect(xml).not.toContain('w:author="Editor"');
		expect(xml).toContain(
			`${INS_A}<w:r><w:t xml:space="preserve">may withhold a </w:t></w:r><w:r><w:t xml:space="preserve">disputed amount</w:t></w:r></w:ins>`,
		);
		expect((await listChanges(docPath)).map((change) => change.text)).toEqual([
			"may withhold a disputed amount",
			".",
		]);
	});

	test("--all: two matches inside the same ins split it twice, right to left", async () => {
		const docPath = await trackedReplace(
			`<w:p>${INS_A}<w:r><w:t>disputed amount beta disputed amount.</w:t></w:r></w:ins></w:p>`,
			"ins-all",
			"Editor",
			"disputed amount",
			"--all",
		);
		expect(await viewText(docPath, "--accepted")).toBe(
			"disputed sum beta disputed sum.",
		);
		const changes = await listChanges(docPath);
		// The head of A's ins holds only the editor's nested del, so the reader
		// lists it with no visible text of its own.
		expect(
			changes.map((change) => [change.kind, change.author, change.text]),
		).toEqual([
			["ins", "Reviewer A", ""],
			["del", "Editor", "disputed amount"],
			["ins", "Editor", "disputed sum"],
			["ins", "Reviewer A", " beta "],
			["del", "Editor", "disputed amount"],
			["ins", "Editor", "disputed sum"],
			["ins", "Reviewer A", "."],
		]);
		const ids = [
			...(await readDocumentXml(docPath)).matchAll(/w:id="(\d+)"/g),
		].map((match) => match[1]);
		expect(new Set(ids).size).toBe(ids.length);
	});

	test("a third author's deletion INSIDE the match keeps its place between two editor dels", async () => {
		const docPath = await trackedReplace(
			`<w:p>${INS_A}<w:r><w:t xml:space="preserve">pay disputed </w:t></w:r><w:del w:id="2" w:author="Reviewer B" w:date="2026-09-02T00:00:00Z"><w:r><w:delText xml:space="preserve">old </w:delText></w:r></w:del><w:r><w:t xml:space="preserve">amount now.</w:t></w:r></w:ins></w:p>`,
			"ins-del-inside",
			"Editor",
		);
		expect(await paragraphXml(docPath, "Editor")).toBe(
			`<w:p>${INS_A}<w:r><w:t xml:space="preserve">pay </w:t></w:r>` +
				'<w:del w:id="3" w:author="Editor" w:date="NOW"><w:r><w:delText xml:space="preserve">disputed </w:delText></w:r></w:del>' +
				'<w:del w:id="2" w:author="Reviewer B" w:date="2026-09-02T00:00:00Z"><w:r><w:delText xml:space="preserve">old </w:delText></w:r></w:del>' +
				'<w:del w:id="4" w:author="Editor" w:date="NOW"><w:r><w:delText xml:space="preserve">amount</w:delText></w:r></w:del></w:ins>' +
				'<w:ins w:id="5" w:author="Editor" w:date="NOW"><w:r><w:t xml:space="preserve">disputed sum</w:t></w:r></w:ins>' +
				'<w:ins w:id="6" w:author="Reviewer A" w:date="2026-09-01T00:00:00Z"><w:r><w:t xml:space="preserve"> now.</w:t></w:r></w:ins></w:p>',
		);
		expect(await viewText(docPath, "--accepted")).toBe("pay disputed sum now.");
		// Current view keeps the original word order of the deleted text.
		const current = await runCli("read", docPath, "--current");
		expect(current.stdout.replaceAll(/\[\^tc\d+\]/g, "")).toContain(
			"{++pay ++}{--disputed --}{--old --}{--amount--}{++disputed sum++}{++ now.++}",
		);
	});

	test("the author's OWN ins keeps a footnote reference and a text-box anchor that sat inside the cut", async () => {
		const docPath = join(tempWorkspace("ins-own-anchors"), "out.docx");
		await runCli("create", docPath, "--text", "seed");
		expect(
			(
				await runCli(
					"footnotes",
					"add",
					docPath,
					"--at",
					"p0",
					"--text",
					"the note",
				)
			).exitCode,
		).toBe(0);
		const footnotesBefore = await readPart(docPath, "word/footnotes.xml");
		// Word's text-box run, anchored mid-word ("am|ount"), and the footnote
		// reference run the CLI emits — both inside Reviewer A's insertion.
		const box = wordTextBoxParagraphXml(["box story"], { leadingText: "am" });
		const namespaces = box.slice("<w:p".length, box.indexOf(">"));
		const shapeRun = box.slice(
			box.indexOf("<w:r>"),
			box.lastIndexOf("</w:r>") + "</w:r>".length,
		);
		const paragraph =
			`<w:p${namespaces}><w:r><w:t xml:space="preserve">The Client </w:t></w:r>${INS_A}` +
			'<w:r><w:t xml:space="preserve">may withhold a disputed</w:t></w:r>' +
			'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>' +
			`<w:r><w:t xml:space="preserve"> </w:t></w:r>${shapeRun}<w:r><w:t>ount.</w:t></w:r></w:ins></w:p>`;
		expect(
			(
				await runCli(
					"raw",
					"replace",
					docPath,
					"--at",
					"p0",
					"--xml",
					paragraph,
				)
			).exitCode,
		).toBe(0);
		const before = await runCli("read", docPath);
		expect(before.stdout).toContain(
			"The Client may withhold a disputed[^fn1] amount.",
		);
		expect(before.stdout).toContain("box story");

		const result = await runCli(
			"replace",
			docPath,
			"disputed amount",
			"disputed sum",
			"--track",
			"--author",
			"Reviewer A",
		);
		expect(result.exitCode).toBe(0);
		const after = await runCli("read", docPath);
		expect(after.stdout).toContain(
			"The Client may withhold a disputed sum[^fn1].",
		);
		expect(after.stdout).toContain("docx:textbox tbx0");
		expect(after.stdout).toContain("box story");
		const xml = await readDocumentXml(docPath);
		expect(xml).not.toContain("<w:del ");
		expect(xml.match(/<w:ins /g)).toHaveLength(1);
		expect(xml.match(/<w:footnoteReference /g)).toHaveLength(1);
		expect(xml.match(/<w:txbxContent>/g)).toHaveLength(2);
		expect(await readPart(docPath, "word/footnotes.xml")).toBe(footnotesBefore);
	});

	test("the author's OWN moveTo is split, never merged into (moved text is not new words)", async () => {
		const MOVED =
			'<w:p><w:r><w:t xml:space="preserve">Origin: </w:t></w:r><w:moveFromRangeStart w:id="0" w:author="Reviewer" w:date="2026-05-05T12:00:00Z" w:name="move1"/><w:moveFrom w:id="1" w:author="Reviewer" w:date="2026-05-05T12:00:00Z"><w:r><w:delText>the moved sentence</w:delText></w:r></w:moveFrom><w:moveFromRangeEnd w:id="0"/></w:p>' +
			'<w:p><w:r><w:t xml:space="preserve">Destination: </w:t></w:r><w:moveToRangeStart w:id="2" w:author="Reviewer" w:date="2026-05-05T12:00:00Z" w:name="move1"/><w:moveTo w:id="3" w:author="Reviewer" w:date="2026-05-05T12:00:00Z"><w:r><w:t>the moved sentence</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="2"/><w:r><w:t>.</w:t></w:r></w:p>';
		const docPath = await buildRawDoc(MOVED, "moveto-own");
		const result = await runCli(
			"replace",
			docPath,
			"moved",
			"relocated",
			"--track",
			"--author",
			"Reviewer",
		);
		expect(result.exitCode).toBe(0);
		expect((await runCli("validate", docPath)).exitCode).toBe(0);

		const xml = (await readDocumentXml(docPath)).replaceAll(
			/ w:date="[^"]*"/g,
			"",
		);
		const destination = xml.match(
			/<w:p>(?:(?!<w:p>).)*Destination.*?<\/w:p>/s,
		)?.[0];
		expect(destination).toBe(
			'<w:p><w:r><w:t xml:space="preserve">Destination: </w:t></w:r><w:moveToRangeStart w:id="2" w:author="Reviewer" w:name="move1"/>' +
				'<w:moveTo w:id="3" w:author="Reviewer"><w:r><w:t xml:space="preserve">the </w:t></w:r>' +
				'<w:del w:id="4" w:author="Reviewer"><w:r><w:delText xml:space="preserve">moved</w:delText></w:r></w:del></w:moveTo>' +
				'<w:ins w:id="5" w:author="Reviewer"><w:r><w:t xml:space="preserve">relocated</w:t></w:r></w:ins>' +
				'<w:moveTo w:id="6" w:author="Reviewer"><w:r><w:t xml:space="preserve"> sentence</w:t></w:r></w:moveTo>' +
				'<w:moveToRangeEnd w:id="2"/><w:r><w:t>.</w:t></w:r></w:p>',
		);
		expect(await viewText(docPath, "--accepted")).toBe(
			"Origin:\nDestination: the relocated sentence.",
		);
		expect((await listChanges(docPath)).map((change) => change.kind)).toEqual([
			"moveFrom",
			"moveTo",
			"del",
			"ins",
			"moveTo",
		]);
	});

	test("a comment anchored on the replaced phrase keeps its markers; a marker-only tail is not a revision", async () => {
		const docPath = await buildRawDoc(
			`<w:p><w:r><w:t xml:space="preserve">The Client </w:t></w:r>${INS_A}<w:r><w:t>may withhold a disputed amount</w:t></w:r></w:ins></w:p>`,
			"ins-comment-tail",
		);
		expect(
			(
				await runCli(
					"comments",
					"add",
					docPath,
					"--anchor",
					"disputed amount",
					"--text",
					"why?",
				)
			).exitCode,
		).toBe(0);
		const result = await runCli(
			"replace",
			docPath,
			"disputed amount",
			"disputed sum",
			"--track",
			"--author",
			"Editor",
		);
		expect(result.exitCode).toBe(0);
		expect((await runCli("validate", docPath)).exitCode).toBe(0);

		// The tail (comment range end + reference run) is pure markers: it stays
		// in Reviewer A's ins after the cut — no fresh <w:ins> around zero
		// content, so no empty revision, and the comment still brackets the text.
		expect(await paragraphXml(docPath, "Editor")).toBe(
			'<w:p><w:r><w:t xml:space="preserve">The Client </w:t></w:r>' +
				`${INS_A}<w:r><w:t xml:space="preserve">may withhold a </w:t></w:r><w:commentRangeStart w:id="0"/>` +
				'<w:del w:id="2" w:author="Editor" w:date="NOW"><w:r><w:delText xml:space="preserve">disputed amount</w:delText></w:r></w:del>' +
				'<w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r></w:ins>' +
				'<w:ins w:id="3" w:author="Editor" w:date="NOW"><w:r><w:t xml:space="preserve">disputed sum</w:t></w:r></w:ins></w:p>',
		);
		const changes = await listChanges(docPath);
		expect(changes.map((change) => [change.kind, change.author])).toEqual([
			["ins", "Reviewer A"],
			["del", "Editor"],
			["ins", "Editor"],
		]);
		expect(changes.every((change) => change.text.length > 0)).toBe(true);
		const list = await runCli("comments", "list", docPath);
		const comments = list.parsed as Array<{
			id: string;
			anchor: { startBlockId: string };
		}>;
		expect(comments.map((comment) => [comment.id, comment.anchor])).toEqual([
			["c0", expect.objectContaining({ startBlockId: "p0" })],
		]);

		// Rejecting everything removes A's insertion with the comment markers it
		// held — nothing dangles.
		const rejected = await runCli("track-changes", "reject", docPath, "--all");
		expect(rejected.exitCode).toBe(0);
		const after = await readDocumentXml(docPath);
		expect(after).not.toContain("<w:commentRangeStart ");
		expect(after).not.toContain("<w:commentRangeEnd ");
		expect(after).not.toContain("<w:commentReference ");
		expect(await viewText(docPath, "--accepted")).toBe("The Client");
	});

	test("the author's OWN ins drops a tab that was part of the match (only zero-width children survive)", async () => {
		const docPath = await buildRawDoc(
			`<w:p>${INS_A}<w:r><w:t>Total:</w:t><w:tab/><w:t>100 due</w:t></w:r></w:ins></w:p>`,
			"ins-own-tab",
		);
		// The agent types `\t`; the CLI ingress decodes it to a real tab.
		const result = await runCli(
			"replace",
			docPath,
			"Total:\\t100",
			"Sum 200",
			"--track",
			"--author",
			"Reviewer A",
		);
		expect(result.exitCode).toBe(0);
		expect(await viewText(docPath, "--accepted")).toBe("Sum 200 due");
		expect(await paragraphXml(docPath, "Editor")).toBe(
			`<w:p>${INS_A}<w:r><w:t xml:space="preserve">Sum 200</w:t></w:r><w:r><w:t xml:space="preserve"> due</w:t></w:r></w:ins></w:p>`,
		);
	});

	test("a comment that opened inside the cut brackets the replacement in the author's OWN ins", async () => {
		const docPath = await buildRawDoc(
			`<w:p>${INS_A}<w:r><w:t>pay the disputed amount now.</w:t></w:r></w:ins></w:p>`,
			"ins-own-comment",
		);
		expect(
			(
				await runCli(
					"comments",
					"add",
					docPath,
					"--anchor",
					"amount",
					"--text",
					"which?",
				)
			).exitCode,
		).toBe(0);
		const result = await runCli(
			"replace",
			docPath,
			"disputed amount",
			"disputed sum",
			"--track",
			"--author",
			"Reviewer A",
		);
		expect(result.exitCode).toBe(0);
		expect(await paragraphXml(docPath, "Editor")).toContain(
			'<w:commentRangeStart w:id="0"/><w:r><w:t xml:space="preserve">disputed sum</w:t></w:r><w:commentRangeEnd w:id="0"/>',
		);
		const comments = (await runCli("comments", "list", docPath))
			.parsed as Array<{
			anchor: { startOffset: number; endOffset: number };
		}>;
		expect(comments.map((comment) => comment.anchor)).toEqual([
			expect.objectContaining({ startOffset: 8, endOffset: 20 }),
		]);
		const rendered = await runCli("read", docPath, "--comments");
		expect(rendered.stdout).toContain("pay the disputed sum[^c0] now.");
		expect(rendered.stdout).toContain('[^c0]: "disputed sum"');
	});

	test("a tail holding only a rendered-page-break run is not a revision", async () => {
		const docPath = await trackedReplace(
			`<w:p><w:r><w:t xml:space="preserve">The Client </w:t></w:r>${INS_A}<w:r><w:t>may withhold a disputed amount</w:t></w:r><w:r><w:lastRenderedPageBreak/></w:r></w:ins></w:p>`,
			"ins-pagebreak-tail",
			"Editor",
		);
		expect(await paragraphXml(docPath, "Editor")).toBe(
			'<w:p><w:r><w:t xml:space="preserve">The Client </w:t></w:r>' +
				`${INS_A}<w:r><w:t xml:space="preserve">may withhold a </w:t></w:r>` +
				'<w:del w:id="2" w:author="Editor" w:date="NOW"><w:r><w:delText xml:space="preserve">disputed amount</w:delText></w:r></w:del>' +
				"<w:r><w:lastRenderedPageBreak/></w:r></w:ins>" +
				'<w:ins w:id="3" w:author="Editor" w:date="NOW"><w:r><w:t xml:space="preserve">disputed sum</w:t></w:r></w:ins></w:p>',
		);
		expect((await listChanges(docPath)).map((change) => change.author)).toEqual(
			["Reviewer A", "Editor", "Editor"],
		);
	});

	test("a span CROSSING out of the ins keeps the general walker's shape", async () => {
		// Only a match wholly inside the wrapper is gated; this one starts in the
		// plain run before it. The half of the ins after the cut is a second
		// revision, so it gets a fresh id (it reused w:id="1" before #16).
		const docPath = await trackedReplace(
			SINGLE_RUN,
			"ins-crossing",
			"Editor",
			"Client may",
		);
		expect(await paragraphXml(docPath, "Editor")).toBe(
			'<w:p><w:r><w:t xml:space="preserve">The </w:t></w:r>' +
				'<w:del w:id="2" w:author="Editor" w:date="NOW"><w:r><w:delText xml:space="preserve">Client </w:delText></w:r></w:del>' +
				'<w:ins w:id="3" w:author="Editor" w:date="NOW"><w:r><w:t xml:space="preserve">disputed sum</w:t></w:r></w:ins>' +
				`${INS_A}<w:del w:id="4" w:author="Editor" w:date="NOW"><w:r><w:delText xml:space="preserve">may</w:delText></w:r></w:del></w:ins>` +
				'<w:ins w:id="5" w:author="Reviewer A" w:date="2026-09-01T00:00:00Z"><w:r><w:t xml:space="preserve"> withhold a disputed amount.</w:t></w:r></w:ins></w:p>',
		);
	});

	test("untracked, the replacement lands inside the ins exactly as before", async () => {
		const docPath = await buildRawDoc(SINGLE_RUN, "ins-untracked");
		const result = await runCli(
			"replace",
			docPath,
			"disputed amount",
			"disputed sum",
		);
		expect(result.exitCode).toBe(0);
		expect(await paragraphXml(docPath, "Editor")).toBe(
			'<w:p><w:r><w:t xml:space="preserve">The Client </w:t></w:r>' +
				`${INS_A}<w:r><w:t xml:space="preserve">may withhold a </w:t></w:r>` +
				'<w:r><w:t xml:space="preserve">disputed sum</w:t></w:r>' +
				'<w:r><w:t xml:space="preserve">.</w:t></w:r></w:ins></w:p>',
		);
	});
});

// Issue #16: an UNTRACKED replace whose match starts in plain text inside
// another author's pending <w:ins> and ends inside a <w:hyperlink> nested in
// that same ins. Main split the ins around the match (both halves kept
// w:id="1"), put the replacement BETWEEN the halves as plain original text,
// and never descended the nested link — so the link's part of the match
// survived and the order scrambled ("may [a disputed](#x)keep aamount.").
// Untracked, a revision wrapper is now cut IN PLACE (no split, so no id is
// duplicated) and the replacement lands where the match started — inside the
// ins, credited to its author, as it already did for a match wholly inside it.
describe("docx replace across a wrapper nested in another author's insertion (#16)", () => {
	const INS_A =
		'<w:ins w:id="1" w:author="Reviewer A" w:date="2026-09-01T00:00:00Z">';
	// A <w:hyperlink> directly under <w:ins> is what the issue reported, but the
	// ECMA schema only allows the reverse nesting, so this body fails `validate`
	// before any edit — those cases skip the schema check.
	const INS_LINK = `<w:p><w:r><w:t xml:space="preserve">The Client </w:t></w:r>${INS_A}<w:r><w:t xml:space="preserve">may withhold </w:t></w:r><w:hyperlink w:anchor="x"><w:r><w:t>a disputed</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> amount.</w:t></w:r></w:ins></w:p>`;

	async function paragraphXml(docPath: string): Promise<string> {
		const xml = await readDocumentXml(docPath);
		return xml.match(/<w:p>.*?<\/w:p>/s)?.[0] ?? "";
	}

	function revisionIds(xml: string): string[] {
		return [...xml.matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map(
			(match) => match[1] ?? "",
		);
	}

	async function replaceIn(
		body: string,
		label: string,
		pattern: string,
		replacement: string,
		...flags: string[]
	): Promise<string> {
		const docPath = await buildRawDoc(body, label);
		const result = await runCli(
			"replace",
			docPath,
			pattern,
			replacement,
			...flags,
		);
		expect(result.exitCode).toBe(0);
		if (body !== INS_LINK) {
			expect((await runCli("validate", docPath)).exitCode).toBe(0);
		}
		return docPath;
	}

	test("the link's part of the match is cut and the replacement stays in the ins", async () => {
		const docPath = await replaceIn(
			INS_LINK,
			"ins-link",
			"withhold a",
			"keep a",
		);
		expect(await paragraphXml(docPath)).toBe(
			'<w:p><w:r><w:t xml:space="preserve">The Client </w:t></w:r>' +
				`${INS_A}<w:r><w:t xml:space="preserve">may </w:t></w:r>` +
				'<w:r><w:t xml:space="preserve">keep a</w:t></w:r>' +
				'<w:hyperlink w:anchor="x"><w:r><w:t xml:space="preserve"> disputed</w:t></w:r></w:hyperlink>' +
				'<w:r><w:t xml:space="preserve"> amount.</w:t></w:r></w:ins></w:p>',
		);
		const read = await runCli("read", docPath);
		expect(read.stdout).toContain(
			"The Client may keep a[ disputed](#x) amount.",
		);
		const ids = revisionIds(await paragraphXml(docPath));
		expect(new Set(ids).size).toBe(ids.length);
	});

	test("a match starting inside the nested link keeps the replacement in the link", async () => {
		const docPath = await replaceIn(
			INS_LINK,
			"ins-link-start",
			"disputed amount",
			"sum",
		);
		expect(await paragraphXml(docPath)).toBe(
			'<w:p><w:r><w:t xml:space="preserve">The Client </w:t></w:r>' +
				`${INS_A}<w:r><w:t xml:space="preserve">may withhold </w:t></w:r>` +
				'<w:hyperlink w:anchor="x"><w:r><w:t xml:space="preserve">a </w:t></w:r><w:r><w:t xml:space="preserve">sum</w:t></w:r></w:hyperlink>' +
				'<w:r><w:t xml:space="preserve">.</w:t></w:r></w:ins></w:p>',
		);
	});

	test("a match crossing OUT of the ins cuts it in place without duplicating its id", async () => {
		const docPath = await replaceIn(
			`<w:p><w:r><w:t xml:space="preserve">The Client </w:t></w:r>${INS_A}<w:r><w:t>may withhold</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> a disputed amount.</w:t></w:r></w:p>`,
			"ins-out",
			"withhold a",
			"keep a",
		);
		expect(await paragraphXml(docPath)).toBe(
			'<w:p><w:r><w:t xml:space="preserve">The Client </w:t></w:r>' +
				`${INS_A}<w:r><w:t xml:space="preserve">may </w:t></w:r>` +
				'<w:r><w:t xml:space="preserve">keep a</w:t></w:r></w:ins>' +
				'<w:r><w:t xml:space="preserve"> disputed amount.</w:t></w:r></w:p>',
		);
	});

	test("an ins the match covers entirely is removed, not left empty", async () => {
		const docPath = await replaceIn(
			`<w:p><w:r><w:t xml:space="preserve">The Client may </w:t></w:r>${INS_A}<w:r><w:t>not</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> pay.</w:t></w:r></w:p>`,
			"ins-covered",
			"may not pay",
			"must pay",
		);
		expect(await paragraphXml(docPath)).toBe(
			'<w:p><w:r><w:t xml:space="preserve">The Client </w:t></w:r>' +
				'<w:r><w:t xml:space="preserve">must pay</w:t></w:r>' +
				'<w:r><w:t xml:space="preserve">.</w:t></w:r></w:p>',
		);
	});

	test("a match in a LATER run of a multi-run ins cuts the right words", async () => {
		const docPath = await replaceIn(
			`<w:p><w:r><w:t xml:space="preserve">The Client </w:t></w:r>${INS_A}<w:r><w:t xml:space="preserve">may withhold </w:t></w:r><w:r><w:t>a disputed amount.</w:t></w:r></w:ins></w:p>`,
			"ins-multirun",
			"disputed",
			"contested",
		);
		const read = await runCli("read", docPath);
		expect(read.stdout).toContain(
			"The Client may withhold a contested amount.",
		);
	});

	test("a match in a later run of a multi-run hyperlink cuts the right words", async () => {
		const docPath = await replaceIn(
			'<w:p><w:hyperlink w:anchor="x"><w:r><w:t xml:space="preserve">alpha </w:t></w:r><w:r><w:t>beta gamma</w:t></w:r></w:hyperlink></w:p>',
			"link-multirun",
			"gamma",
			"DELTA",
		);
		const read = await runCli("read", docPath);
		expect(read.stdout).toContain("[alpha beta DELTA](#x)");
	});

	test("a run after a link nested in the ins keeps its offset", async () => {
		const docPath = await buildRawDoc(
			`<w:p>${INS_A}<w:r><w:t xml:space="preserve">alpha </w:t></w:r><w:hyperlink w:anchor="x"><w:r><w:t xml:space="preserve">link </w:t></w:r></w:hyperlink><w:r><w:t>beta gamma</w:t></w:r></w:ins></w:p>`,
			"ins-link-after",
		);
		expect((await runCli("replace", docPath, "gamma", "DELTA")).exitCode).toBe(
			0,
		);
		const read = await runCli("read", docPath, "--accepted");
		expect(read.stdout).toContain("alpha [link ](#x)beta DELTA");
	});

	test("tracked, the half of the ins after the match gets a fresh id", async () => {
		const docPath = await replaceIn(
			`<w:p><w:r><w:t xml:space="preserve">The Client </w:t></w:r>${INS_A}<w:r><w:t>may withhold a disputed amount.</w:t></w:r></w:ins></w:p>`,
			"ins-tracked-out",
			"Client may",
			"Client must",
			"--track",
		);
		const ids = revisionIds(await paragraphXml(docPath));
		expect(new Set(ids).size).toBe(ids.length);
	});
});
