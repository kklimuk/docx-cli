import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

describe("docx wc", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("wc");
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
			"Sphinx of black quartz, judge my vow.",
		);
		await runCli(
			"insert",
			docPath,
			"--after",
			"p1",
			"--text",
			"Pack my box with five dozen liquor jugs.",
		);
	});

	test("no locator counts the whole document", async () => {
		const result = await runCli("wc", docPath);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			ok: true,
			operation: "wc",
			scope: "document",
			words: 9 + 7 + 8,
		});
	});

	test("paragraph locator counts that paragraph", async () => {
		const result = await runCli("wc", docPath, "p1");
		expect(result.parsed).toMatchObject({
			locator: "p1",
			scope: "paragraph",
			words: 7,
		});
	});

	test("span locator counts a paragraph slice", async () => {
		// "The quick brown fox" → chars 0..19, four words
		const result = await runCli("wc", docPath, "p0:0-19");
		expect(result.parsed).toMatchObject({
			locator: "p0:0-19",
			scope: "paragraphSpan",
			words: 4,
		});
	});

	test("cross-paragraph range counts from offset to offset", async () => {
		// p0 starts at "The quick brown fox jumps over the lazy dog."
		// p2 ends at "Pack my box with five dozen liquor jugs."
		// Range "p0:16-p2:11" covers "fox jumps over the lazy dog." +
		// "Sphinx of black quartz, judge my vow." + "Pack my box"
		const result = await runCli("wc", docPath, "p0:16-p2:11");
		expect(result.parsed).toMatchObject({
			locator: "p0:16-p2:11",
			scope: "range",
			words: 6 + 7 + 3,
		});
	});

	test("table locator sums every cell paragraph", async () => {
		const result = await runCli(
			"wc",
			"tests/fixtures/tables-and-lists.docx",
			"t0",
		);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			locator: "t0",
			scope: "table",
			words: 32,
		});
	});

	test("cell and cell-paragraph locators count just the cell", async () => {
		const cellResult = await runCli(
			"wc",
			"tests/fixtures/tables-and-lists.docx",
			"t0:r1c0",
		);
		expect(cellResult.parsed).toMatchObject({
			locator: "t0:r1c0",
			scope: "cell",
			words: 7,
		});

		const cellParagraphResult = await runCli(
			"wc",
			"tests/fixtures/tables-and-lists.docx",
			"t0:r1c0:p0",
		);
		expect(cellParagraphResult.parsed).toMatchObject({
			locator: "t0:r1c0:p0",
			scope: "paragraph",
			words: 7,
		});
	});

	test("comment and image locators are rejected as USAGE errors", async () => {
		const commentResult = await runCli("wc", docPath, "c0");
		expect(commentResult.exitCode).toBe(2);
		expect(commentResult.parsed).toMatchObject({ ok: false, code: "USAGE" });

		const imageResult = await runCli("wc", docPath, "img0");
		expect(imageResult.exitCode).toBe(2);
		expect(imageResult.parsed).toMatchObject({ ok: false, code: "USAGE" });
	});

	test("missing block returns BLOCK_NOT_FOUND with exit 3", async () => {
		const result = await runCli("wc", docPath, "p99");
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({
			ok: false,
			code: "BLOCK_NOT_FOUND",
		});
	});

	test("invalid locator returns INVALID_LOCATOR", async () => {
		const result = await runCli("wc", docPath, "p3:bogus");
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({
			ok: false,
			code: "INVALID_LOCATOR",
		});
	});
});

describe("docx wc — tracked changes", () => {
	async function trackedFixture(label: string): Promise<string> {
		const workspace = tempWorkspace(label);
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "The quick brown fox jumps.");
		await runCli("track-changes", docPath, "on");
		// `replace` under track-changes emits a <w:del> for "quick brown " and
		// a <w:ins> for "old slow ", giving us both kinds in one paragraph.
		await runCli("replace", docPath, "quick brown ", "old slow ");
		return docPath;
	}

	test("default is the accepted view (skips del; keeps ins)", async () => {
		const docPath = await trackedFixture("wc-default");
		const result = await runCli("wc", docPath);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { view: string }).view).toBe("accepted");
		// "The old slow fox jumps." = 5 words (skip the 2-word del). Same
		// as --accepted; default flipped from "current" to match
		// `read` / `find` / `replace`.
		expect((result.parsed as { words: number }).words).toBe(5);
	});

	test("--current counts the on-disk view (plain + ins + del)", async () => {
		const docPath = await trackedFixture("wc-current");
		const result = await runCli("wc", docPath, "--current");
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { words: number; view: string }).view).toBe(
			"current",
		);
		// Plain "The fox jumps." (3 words) + ins "old slow " (2) + del "quick brown " (2) = 7
		expect((result.parsed as { words: number }).words).toBe(7);
	});

	test("--accepted skips deletions, keeps insertions (explicit alias)", async () => {
		const docPath = await trackedFixture("wc-accepted");
		const result = await runCli("wc", docPath, "--accepted");
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { view: string }).view).toBe("accepted");
		// "The old slow fox jumps." = 5 words (skip the 2-word del)
		expect((result.parsed as { words: number }).words).toBe(5);
	});

	test("--baseline skips insertions, keeps deletions", async () => {
		const docPath = await trackedFixture("wc-baseline");
		const result = await runCli("wc", docPath, "--baseline");
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { view: string }).view).toBe("baseline");
		// "The quick brown fox jumps." = 5 words (skip the 2-word ins)
		expect((result.parsed as { words: number }).words).toBe(5);
	});

	test("--accepted and --baseline together are rejected", async () => {
		const result = await runCli(
			"wc",
			"tests/fixtures/tracked-changes.docx",
			"--accepted",
			"--baseline",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ ok: false, code: "USAGE" });
	});

	test("tracked-changes fixture: current/accepted both count ins, baseline skips it", async () => {
		// Fixture has only <w:ins> ("two exciting"), no <w:del>.
		// "This is a text with two exciting insertions." — 8 words on disk.
		const current = await runCli("wc", "tests/fixtures/tracked-changes.docx");
		expect((current.parsed as { words: number }).words).toBe(8);

		const accepted = await runCli(
			"wc",
			"tests/fixtures/tracked-changes.docx",
			"--accepted",
		);
		expect((accepted.parsed as { words: number }).words).toBe(8);

		const baseline = await runCli(
			"wc",
			"tests/fixtures/tracked-changes.docx",
			"--baseline",
		);
		// "This is a text with insertions." = 6 words.
		expect((baseline.parsed as { words: number }).words).toBe(6);
	});
});
