import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";
import { readMarkdown } from "./helpers";

// End-to-end document lifecycle: author → revise → review → read. Each test
// drives several CLI verbs in sequence and proves the result is observable on
// the next read — the load-bearing "changes survive the write-read loop, and
// wherever possible in Markdown" invariant. Per-verb mechanics live in the
// matching <verb>.test.ts; this file guards that the verbs COMPOSE.

/** Author a small report: a title, a Heading2, a body paragraph, a 2-item list. */
async function authoredReport(label: string): Promise<string> {
	const path = join(tempWorkspace(label), "report.docx");
	await runCli("create", path, "--text", "Quarterly Report");
	await runCli(
		"insert",
		path,
		"--after",
		"p0",
		"--markdown",
		"## Summary\n\nRevenue grew in the Alpha region. Alpha is strong.\n\n- one\n- two",
	);
	return path;
}

async function blockIds(path: string): Promise<string[]> {
	const read = await runCli("read", path, "--ast");
	return (read.parsed as { blocks: Array<{ id: string }> }).blocks.map(
		(block) => block.id,
	);
}

describe("end-to-end document lifecycle", () => {
	test("author → reader: created content is fully observable in Markdown + AST + queries", async () => {
		const path = await authoredReport("e2e-author");

		const markdown = await readMarkdown(path);
		expect(markdown).toContain("Quarterly Report");
		expect(markdown).toContain("## Summary");
		expect(markdown).toContain("Revenue grew in the Alpha region.");
		expect(markdown).toContain("- one");
		expect(markdown).toContain("- two");

		// The same content is addressable through the AST and the query verbs.
		const ids = await blockIds(path);
		expect(ids[0]).toBe("p0");
		expect(ids.length).toBeGreaterThanOrEqual(5);

		const wc = await runCli("wc", path); // harness injects --json
		expect((wc.parsed as { words: number }).words).toBeGreaterThan(0);

		const outline = await runCli("outline", path); // harness injects --json
		expect(JSON.stringify(outline.parsed)).toContain("Summary");
	});

	test("editor → reader: edits and replacements survive the write-read loop", async () => {
		const path = await authoredReport("e2e-edit");

		// Reword the heading and swap a term everywhere it appears.
		await runCli("edit", path, "--at", "p1", "--text", "Executive Summary");
		await runCli("replace", path, "Alpha", "Beta", "--all");

		const markdown = await readMarkdown(path);
		expect(markdown).toContain("## Executive Summary");
		expect(markdown).toContain("Revenue grew in the Beta region.");
		expect(markdown).toContain("Beta is strong");
		expect(markdown).not.toContain("Alpha");
	});

	test("reviewer → reader: tracked edits + comments + footnotes survive, list, and accept", async () => {
		const path = await authoredReport("e2e-review");

		// Annotate paragraphs that the tracked edit below will NOT rewrite, so
		// the anchors aren't clobbered: comment on the title, footnote on the
		// final list item.
		await runCli(
			"comments",
			"add",
			path,
			"--anchor",
			"Quarterly Report",
			"--text",
			"Confirm the final numbers.",
		);
		await runCli(
			"footnotes",
			"add",
			path,
			"--anchor",
			"two",
			"--text",
			"Source: FY24 close.",
		);

		// Turn tracking on and rewrite the body paragraph (p2) as a tracked change.
		await runCli("track-changes", path, "on");
		await runCli(
			"edit",
			path,
			"--at",
			"p2",
			"--text",
			"Revenue surged across every Beta region.",
		);

		// The change and the comment are both enumerable before accepting.
		const list = await runCli("track-changes", "list", path);
		expect((list.parsed as unknown[]).length).toBeGreaterThan(0);
		const comments = await runCli("comments", "list", path);
		expect((comments.parsed as unknown[]).length).toBe(1);

		// Accept everything; the accepted text and the footnote marker survive.
		await runCli("track-changes", "accept", path, "--all");
		const markdown = await readMarkdown(path);
		expect(markdown).toContain("Revenue surged across every Beta region.");
		expect(markdown).toContain("[^fn1]");

		// The comment outlives the accept (it annotates untouched text).
		const afterAccept = await runCli("comments", "list", path);
		expect((afterAccept.parsed as unknown[]).length).toBe(1);
	});

	test("re-read keeps locators valid across sequential structural edits", async () => {
		const path = await authoredReport("e2e-reread");

		// A structural insert shifts positional ids; the title stays p0, but the
		// list items move — so re-read (here via find) before addressing them.
		await runCli(
			"insert",
			path,
			"--after",
			"p0",
			"--text",
			"Prepared by Finance.",
		);
		const ids = await blockIds(path);
		expect(ids[0]).toBe("p0");

		const found = await runCli("find", path, "one"); // harness injects --json
		const locator = (found.parsed as { matches: Array<{ locator: string }> })
			.matches[0]?.locator;
		expect(locator).toBeTruthy();

		const blockId = (locator as string).split(":")[0] ?? "";
		await runCli("edit", path, "--at", blockId, "--text", "first item");

		const markdown = await readMarkdown(path);
		expect(markdown).toContain("Prepared by Finance.");
		expect(markdown).toContain("first item");
		expect(markdown).not.toContain("- one");
	});
});
