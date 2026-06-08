import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

/**
 * Per-command `--batch` for edit / replace / insert — apply many changes from
 * one read. The load-bearing property under test is **locator stability**: all
 * locators in a batch address the document AS READ, and applying one change must
 * never silently misdirect another. The mechanisms differ per command:
 *   - edit:    resolve every entry to a live node ref before any mutation;
 *              same-paragraph spans apply right-to-left.
 *   - replace: re-read the live tree between entries (sequential / sed-like).
 *   - insert:  pin anchors to live refs, build all blocks, then splice with a
 *              per-anchor offset so stacked inserts keep entry order.
 */

type RunAst = { type: string; text?: string; highlight?: string };
type Block = { id: string; type: string; runs?: RunAst[] };

async function blocks(path: string): Promise<Block[]> {
	const result = await runCli("read", path, "--ast");
	return (result.parsed as { blocks: Block[] }).blocks;
}

function textOf(block: Block): string {
	return (block.runs ?? [])
		.filter((run) => run.type === "text")
		.map((run) => run.text ?? "")
		.join("");
}

async function blockText(path: string, id: string): Promise<string> {
	const block = (await blocks(path)).find((candidate) => candidate.id === id);
	if (!block) throw new Error(`block ${id} not found`);
	return textOf(block);
}

/** Paragraph texts in document order — excludes the trailing section-break (sN)
 *  block `create` leaves behind, so assertions compare just the prose. */
async function paragraphTexts(path: string): Promise<string[]> {
	return (await blocks(path))
		.filter((block) => block.type === "paragraph")
		.map(textOf);
}

/** Build a doc with paragraphs p0..pN-1 carrying the given texts. */
async function docWithParagraphs(
	label: string,
	texts: string[],
): Promise<string> {
	const path = join(tempWorkspace(label), "doc.docx");
	await runCli("create", path, "--text", texts[0] ?? "");
	let anchor = "p0";
	for (let index = 1; index < texts.length; index++) {
		await runCli(
			"insert",
			path,
			"--after",
			anchor,
			"--text",
			texts[index] ?? "",
		);
		anchor = `p${index}`;
	}
	return path;
}

async function writeJsonl(label: string, lines: unknown[]): Promise<string> {
	const path = join(tempWorkspace(label), "batch.jsonl");
	await Bun.write(path, lines.map((line) => JSON.stringify(line)).join("\n"));
	return path;
}

describe("edit --batch", () => {
	test("applies span edits across several paragraphs in one read", async () => {
		const path = await docWithParagraphs("edit-multi", [
			"Name: ____",
			"City: ____",
			"Role: ____",
		]);
		const batch = await writeJsonl("edit-multi", [
			{ at: "p0:6-10", text: "Ada" },
			{ at: "p1:6-10", text: "Paris" },
			{ at: "p2:6-10", text: "Eng" },
		]);
		const result = await runCli("edit", path, "--batch", batch);
		expect(result.exitCode).toBe(0);
		expect(await blockText(path, "p0")).toBe("Name: Ada");
		expect(await blockText(path, "p1")).toBe("City: Paris");
		expect(await blockText(path, "p2")).toBe("Role: Eng");
	});

	test("two spans in ONE paragraph stay correct (right-to-left apply)", async () => {
		// If the lower-offset span were applied first, the higher span's offsets
		// (computed against the original text) would drift. Descending order keeps
		// every locator valid against the doc as read.
		const path = await docWithParagraphs("edit-twospan", ["AAAA BBBB CCCC"]);
		const batch = await writeJsonl("edit-twospan", [
			{ at: "p0:0-4", text: "W" }, // listed low-offset first on purpose
			{ at: "p0:10-14", text: "ZZZZZZ" }, // longer than original — would shift
		]);
		const result = await runCli("edit", path, "--batch", batch);
		expect(result.exitCode).toBe(0);
		expect(await blockText(path, "p0")).toBe("W BBBB ZZZZZZ");
	});

	test("a markdown entry that EXPANDS one paragraph doesn't misdirect a later entry", async () => {
		// p0 expands into two heading blocks (positional ids shift), but the p2
		// edit was pinned to a live node ref before any mutation, so it still lands
		// on the original third paragraph.
		const path = await docWithParagraphs("edit-expand", [
			"First",
			"Second",
			"Third",
		]);
		const batch = await writeJsonl("edit-expand", [
			{ at: "p0", markdown: "## New A\n\n## New B" },
			{ at: "p2", text: "ThirdEdited" },
		]);
		const result = await runCli("edit", path, "--batch", batch);
		expect(result.exitCode).toBe(0);
		expect(await paragraphTexts(path)).toEqual([
			"New A",
			"New B",
			"Second",
			"ThirdEdited",
		]);
	});

	test("clear strips formatting in a batch", async () => {
		const path = await docWithParagraphs("edit-clear", ["Intro."]);
		await runCli(
			"insert",
			path,
			"--after",
			"p0",
			"--runs",
			JSON.stringify([{ type: "text", text: "loud", highlight: "yellow" }]),
		);
		const batch = await writeJsonl("edit-clear", [
			{ at: "p1", clear: "highlight" },
		]);
		const result = await runCli("edit", path, "--batch", batch);
		expect(result.exitCode).toBe(0);
		const p1 = (await blocks(path)).find((block) => block.id === "p1");
		expect(p1?.runs?.[0]?.highlight).toBeUndefined();
	});

	test("dry-run lists locators and does not write", async () => {
		const path = await docWithParagraphs("edit-dry", ["AAAA BBBB"]);
		const batch = await writeJsonl("edit-dry", [{ at: "p0:0-4", text: "ZZ" }]);
		const result = await runCli("edit", path, "--batch", batch, "--dry-run");
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { dryRun?: boolean }).dryRun).toBe(true);
		expect(await blockText(path, "p0")).toBe("AAAA BBBB"); // untouched
	});

	test("rejects a whole-paragraph edit + a span on the same paragraph", async () => {
		const path = await docWithParagraphs("edit-conflict", ["AAAA BBBB"]);
		const batch = await writeJsonl("edit-conflict", [
			{ at: "p0", text: "whole" },
			{ at: "p0:0-4", text: "ZZ" },
		]);
		const result = await runCli("edit", path, "--batch", batch);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code?: string }).code).toBe("USAGE");
	});

	test("rejects overlapping spans on the same paragraph", async () => {
		const path = await docWithParagraphs("edit-overlap", ["AAAA BBBB"]);
		const batch = await writeJsonl("edit-overlap", [
			{ at: "p0:0-5", text: "X" },
			{ at: "p0:3-8", text: "Y" },
		]);
		const result = await runCli("edit", path, "--batch", batch);
		expect(result.exitCode).toBe(2);
	});

	test("rejects mixing --batch with single-shot flags", async () => {
		const path = await docWithParagraphs("edit-mix", ["AAAA"]);
		const batch = await writeJsonl("edit-mix", [{ at: "p0:0-4", text: "Z" }]);
		const result = await runCli("edit", path, "--batch", batch, "--at", "p0");
		expect(result.exitCode).toBe(2);
	});

	test("a malformed entry fails with its index", async () => {
		const path = await docWithParagraphs("edit-bad", ["AAAA"]);
		const batch = await writeJsonl("edit-bad", [
			{ at: "p0:0-4", text: "ok" },
			{ at: "p9:0-1", text: "nope" }, // p9 doesn't exist
		]);
		const result = await runCli("edit", path, "--batch", batch);
		expect(result.exitCode).toBe(3);
		expect((result.parsed as { error?: string }).error).toContain("entry 1");
	});
});

describe("replace --batch", () => {
	test("applies several substitutions in one read", async () => {
		const path = await docWithParagraphs("rep-multi", ["Q2 results for FY24"]);
		const batch = await writeJsonl("rep-multi", [
			{ pattern: "Q2", replacement: "Q3" },
			{ pattern: "FY24", replacement: "FY25" },
		]);
		const result = await runCli("replace", path, "--batch", batch);
		expect(result.exitCode).toBe(0);
		expect(await blockText(path, "p0")).toBe("Q3 results for FY25");
	});

	test("later entries see earlier substitutions (sequential, re-read between)", async () => {
		// "A"→"B" then "B"→"C" must yield "C" — only correct if the second find
		// runs against the post-first-edit tree.
		const path = await docWithParagraphs("rep-seq", ["A A A"]);
		const batch = await writeJsonl("rep-seq", [
			{ pattern: "A", replacement: "B", all: true },
			{ pattern: "B", replacement: "C", all: true },
		]);
		const result = await runCli("replace", path, "--batch", batch);
		expect(result.exitCode).toBe(0);
		expect(await blockText(path, "p0")).toBe("C C C");
	});

	test("dry-run reports counts and does not write", async () => {
		const path = await docWithParagraphs("rep-dry", ["foo foo foo"]);
		const batch = await writeJsonl("rep-dry", [
			{ pattern: "foo", replacement: "bar", all: true },
		]);
		const result = await runCli("replace", path, "--batch", batch, "--dry-run");
		expect(result.exitCode).toBe(0);
		const parsed = result.parsed as {
			dryRun?: boolean;
			batch?: Array<{ replaced: number }>;
		};
		expect(parsed.dryRun).toBe(true);
		expect(parsed.batch?.[0]?.replaced).toBe(3);
		expect(await blockText(path, "p0")).toBe("foo foo foo"); // untouched
	});
});

describe("insert --batch", () => {
	test("inserts at several anchors and reports minted locators", async () => {
		const path = await docWithParagraphs("ins-multi", ["First", "Second"]);
		const batch = await writeJsonl("ins-multi", [
			{ after: "p0", text: "After first" },
			{ before: "p1", text: "Before second" },
		]);
		const result = await runCli("insert", path, "--batch", batch);
		expect(result.exitCode).toBe(0);
		// p0=First, then "After first", then "Before second", then Second
		expect(await paragraphTexts(path)).toEqual([
			"First",
			"After first",
			"Before second",
			"Second",
		]);
	});

	test("stacked inserts after the SAME anchor keep entry order", async () => {
		// Two inserts after p0: naive indexOf+1 each time would reverse them. The
		// per-anchor offset must keep "One" before "Two".
		const path = await docWithParagraphs("ins-stack", ["Top", "Bottom"]);
		const batch = await writeJsonl("ins-stack", [
			{ after: "p0", text: "One" },
			{ after: "p0", text: "Two" },
		]);
		const result = await runCli("insert", path, "--batch", batch);
		expect(result.exitCode).toBe(0);
		expect(await paragraphTexts(path)).toEqual(["Top", "One", "Two", "Bottom"]);
	});

	test("an insert at an earlier anchor doesn't misdirect a later anchor", async () => {
		// Inserting after p0 shifts p2's position; because anchors are pinned to
		// live refs, the p2 insert still lands after the original third paragraph.
		const path = await docWithParagraphs("ins-shift", ["A", "B", "C"]);
		const batch = await writeJsonl("ins-shift", [
			{ after: "p0", text: "afterA" },
			{ after: "p2", text: "afterC" },
		]);
		const result = await runCli("insert", path, "--batch", batch);
		expect(result.exitCode).toBe(0);
		expect(await paragraphTexts(path)).toEqual([
			"A",
			"afterA",
			"B",
			"C",
			"afterC",
		]);
	});

	test("dry-run does not write", async () => {
		const path = await docWithParagraphs("ins-dry", ["Only"]);
		const batch = await writeJsonl("ins-dry", [{ after: "p0", text: "New" }]);
		const result = await runCli("insert", path, "--batch", batch, "--dry-run");
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { dryRun?: boolean }).dryRun).toBe(true);
		expect(await paragraphTexts(path)).toEqual(["Only"]); // untouched
	});

	test("stacked `before` and mixed after/before on one anchor keep order", async () => {
		// `afterOffset` only tracks `after`; `before` relies on recomputed indexOf.
		const path = await docWithParagraphs("ins-before", ["Top", "Bottom"]);
		const batch = await writeJsonl("ins-before", [
			{ before: "p1", text: "One" },
			{ before: "p1", text: "Two" },
			{ after: "p0", text: "Mid" },
		]);
		const result = await runCli("insert", path, "--batch", batch);
		expect(result.exitCode).toBe(0);
		// after p0 → Mid right after Top; before p1 → One, Two stack just before Bottom.
		expect(await paragraphTexts(path)).toEqual([
			"Top",
			"Mid",
			"One",
			"Two",
			"Bottom",
		]);
	});

	test("under tracking, every entry's revision ids are unique (no w:id collisions)", async () => {
		// All blocks are built (ids minted) before any splice, so a per-entry
		// allocator would re-scan the same tree max and collide. One shared
		// allocator keeps w:ids monotonic across entries.
		const path = await docWithParagraphs("ins-track", ["Doc"]);
		await runCli("track-changes", path, "on");
		const batch = await writeJsonl("ins-track", [
			{ after: "p0", text: "Alpha" },
			{ after: "p0", text: "Beta" },
			{ after: "p0", text: "Gamma" },
		]);
		const result = await runCli("insert", path, "--batch", batch);
		expect(result.exitCode).toBe(0);
		const changes = (await runCli("track-changes", "list", path))
			.parsed as Array<{ revisionId: string }>;
		const ids = changes.map((c) => c.revisionId);
		expect(new Set(ids).size).toBe(ids.length); // all unique
		expect(ids.length).toBeGreaterThanOrEqual(3);
	});

	test("rejects a single-shot paragraph flag (--style) passed alongside --batch", async () => {
		const path = await docWithParagraphs("ins-flag", ["Doc"]);
		const batch = await writeJsonl("ins-flag", [{ after: "p0", text: "x" }]);
		const result = await runCli(
			"insert",
			path,
			"--batch",
			batch,
			"--style",
			"Heading1",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code?: string }).code).toBe("USAGE");
	});
});
