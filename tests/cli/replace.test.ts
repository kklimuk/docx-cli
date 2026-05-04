import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

type Doc = {
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
		const read = await runCli("read", docPath);
		const text = (read.parsed as Doc).blocks
			.flatMap((block) => block.runs ?? [])
			.filter((run) => run.type === "text")
			.map((run) => run.text)
			.join("");
		expect(text).toContain("brown cat jumps");
		expect(text).toContain("Another fox sneaks");
	});

	test("--all replaces every match", async () => {
		await runCli("replace", docPath, "fox", "cat", "--all");
		const read = await runCli("read", docPath);
		const text = (read.parsed as Doc).blocks
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
		const read = await runCli("read", docPath);
		const text = (read.parsed as Doc).blocks
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
		const read = await runCli("read", docPath);
		const text = (read.parsed as Doc).blocks
			.flatMap((block) => block.runs ?? [])
			.filter((run) => run.type === "text")
			.map((run) => run.text)
			.join("");
		expect(text).toContain("The brown quick fox");
	});

	test("--regex with $& full-match reference", async () => {
		await runCli("replace", docPath, "fox", "[$&]", "--regex", "--all");
		const read = await runCli("read", docPath);
		const text = (read.parsed as Doc).blocks
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
		const read = await runCli("read", colorPath);
		const paragraph = (read.parsed as Doc).blocks.find(
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
		const read = await runCli("read", multiPath);
		const text = (read.parsed as Doc).blocks
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
		expect(result.parsed).toMatchObject({ ok: false, code: "USAGE" });
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
		expect(result.parsed).toMatchObject({ ok: false, code: "USAGE" });
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

		const read = await runCli("read", cellPath);
		const text = JSON.stringify(read.parsed);
		expect(text).toContain("Protoboard");
		expect(text).not.toContain("Breadboard");
	});
});
