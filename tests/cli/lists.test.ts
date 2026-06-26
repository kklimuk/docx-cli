import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";
import { freshFixture, readMarkdown } from "./helpers";

// `docx lists set` controls a numbered list's numbering: its start value, glyph
// format (decimal / lower-alpha / upper-roman / …), and whether it restarts or
// continues the previous list. Start round-trips through the markdown ordinal;
// format and continue surface as the dropped-on-import `docx:list` hint.

const FIXTURES = join(import.meta.dir, "..", "fixtures");
const LISTS = join(FIXTURES, "lists.docx");
// In lists.docx the ordered list is p8 (first item) … p13.
const ORDERED_FIRST = "p8";

/** Build a doc with two ordered lists separated by a heading, for restart /
 * continue. Returns the path; the lists are p0–p2 (numId 1) and p4–p5 (numId 2). */
async function twoOrderedLists(label: string): Promise<string> {
	const dir = tempWorkspace(label);
	const md = join(dir, "src.md");
	await Bun.write(
		md,
		"1. alpha\n2. beta\n3. gamma\n\n## Break\n\n1. delta\n2. epsilon\n",
	);
	const docPath = join(dir, "doc.docx");
	const created = await runCli("create", docPath, "--force", "--from", md);
	expect(created.exitCode).toBe(0);
	return docPath;
}

describe("docx lists set", () => {
	test("--start renumbers the list and round-trips through the markdown ordinal", async () => {
		const path = await freshFixture("lists-start", LISTS);
		const result = await runCli(
			"lists",
			"set",
			path,
			"--at",
			ORDERED_FIRST,
			"--start",
			"5",
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { operation: string }).operation).toBe(
			"lists.set",
		);
		expect((result.parsed as { applied: string[] }).applied).toContain(
			"start=5",
		);

		const markdown = await readMarkdown(path);
		expect(markdown).toContain("5. Preheat the oven");
		expect(markdown).toContain("6. Combine the dry ingredients");
		expect(markdown).toContain("7. Bake for forty minutes");
		// The run-start carries a docx:list hint naming the start (feeds --start back).
		expect(markdown).toMatch(/docx:list p8 start="5"/);
	});

	test("--format overrides the glyph (surfaced as the docx:list hint + AST)", async () => {
		const path = await freshFixture("lists-format", LISTS);
		const result = await runCli(
			"lists",
			"set",
			path,
			"--at",
			ORDERED_FIRST,
			"--format",
			"upper-roman",
		);
		expect(result.exitCode).toBe(0);

		const markdown = await readMarkdown(path);
		// GFM can't render roman, so the body stays decimal-looking; the hint carries it.
		expect(markdown).toMatch(/docx:list p8 format="upper-roman"/);

		const ast = await runCli("read", path, "--ast");
		const block = (
			ast.parsed as {
				blocks: Array<{ id: string; list?: { format?: string } }>;
			}
		).blocks.find((b) => b.id === ORDERED_FIRST);
		expect(block?.list?.format).toBe("upper-roman");
	});

	test("--restart splits a list: earlier items keep numbering, the rest restart", async () => {
		const path = await freshFixture("lists-restart", LISTS);
		// p9 is the second top-level item; restart it as a fresh list at 3.
		const result = await runCli(
			"lists",
			"set",
			path,
			"--at",
			"p9",
			"--restart",
			"--start",
			"3",
		);
		expect(result.exitCode).toBe(0);

		const markdown = await readMarkdown(path);
		expect(markdown).toContain("1. Preheat the oven"); // p8 keeps the old list
		expect(markdown).toContain("3. Combine the dry ingredients"); // p9 begins anew at 3
		expect(markdown).toContain("4. Bake for forty minutes"); // p13 follows the new list
		expect(markdown).toMatch(/docx:list p9 start="3"/);
	});

	test("--continue makes a list pick up the previous list's numbering", async () => {
		const path = await twoOrderedLists("lists-continue");
		const result = await runCli(
			"lists",
			"set",
			path,
			"--at",
			"p4",
			"--continue",
		);
		expect(result.exitCode).toBe(0);

		const markdown = await readMarkdown(path);
		expect(markdown).toContain("3. gamma"); // first list ends at 3
		expect(markdown).toContain("4. delta"); // second list continues, not restarts
		expect(markdown).toContain("5. epsilon");
		expect(markdown).toMatch(/docx:list p4 continues/);
	});

	test("rejects a bulleted list (numbering controls are for ordered lists)", async () => {
		const path = await freshFixture("lists-bullet", LISTS);
		// p1 is the first bullet item.
		const result = await runCli(
			"lists",
			"set",
			path,
			"--at",
			"p1",
			"--start",
			"5",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { error: string }).error).toMatch(/bulleted list/);
	});

	test("rejects --restart with --continue, and a no-op call", async () => {
		const path = await freshFixture("lists-validation", LISTS);
		const both = await runCli(
			"lists",
			"set",
			path,
			"--at",
			ORDERED_FIRST,
			"--restart",
			"--continue",
		);
		expect(both.exitCode).toBe(2);
		expect((both.parsed as { error: string }).error).toMatch(/only one of/);

		const noop = await runCli("lists", "set", path, "--at", ORDERED_FIRST);
		expect(noop.exitCode).toBe(2);
		expect((noop.parsed as { error: string }).error).toMatch(/Nothing to do/);
	});

	test("--continue with no preceding list fails (not found)", async () => {
		const path = await freshFixture("lists-no-preceding", LISTS);
		// p8 is the FIRST ordered list — nothing precedes it.
		const result = await runCli(
			"lists",
			"set",
			path,
			"--at",
			ORDERED_FIRST,
			"--continue",
		);
		expect(result.exitCode).toBe(3);
		expect((result.parsed as { error: string }).error).toMatch(
			/no preceding list/,
		);
	});

	test("--dry-run previews without writing", async () => {
		const path = await freshFixture("lists-dry", LISTS);
		const before = await readMarkdown(path);
		const result = await runCli(
			"lists",
			"set",
			path,
			"--at",
			ORDERED_FIRST,
			"--start",
			"9",
			"--dry-run",
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { dryRun: boolean }).dryRun).toBe(true);
		expect(await readMarkdown(path)).toBe(before);
	});
});
