import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

describe("docx insert / edit / delete", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("ied");
		docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Original body");
	});

	test("insert --after places paragraph after the locator", async () => {
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Appended",
			"--style",
			"Heading2",
			"--color",
			"CC0000",
			"--bold",
		);
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath);
		const doc = read.parsed as {
			blocks: Array<{
				id: string;
				type: string;
				style?: string;
				runs?: Array<{ text: string; color?: string; bold?: boolean }>;
			}>;
		};
		const paragraphs = doc.blocks.filter((block) => block.type === "paragraph");
		expect(paragraphs[0]?.runs?.[0]?.text).toBe("Original body");
		expect(paragraphs[1]?.style).toBe("Heading2");
		expect(paragraphs[1]?.runs?.[0]).toMatchObject({
			text: "Appended",
			color: "CC0000",
			bold: true,
		});
	});

	test("insert --before places paragraph before the locator", async () => {
		await runCli("insert", docPath, "--before", "p0", "--text", "Prepended");
		const read = await runCli("read", docPath);
		const doc = read.parsed as {
			blocks: Array<{ type: string; runs?: Array<{ text: string }> }>;
		};
		const paragraphs = doc.blocks.filter((block) => block.type === "paragraph");
		expect(paragraphs[0]?.runs?.[0]?.text).toBe("Prepended");
	});

	test("insert --runs silently drops unsupported run types (round-trip safety)", async () => {
		// Simulates `docx read | jq | docx insert --runs '[...]'` where the
		// source paragraph contained an equation/footnoteRef/chart that we
		// surface in the AST but can't re-emit as fresh OOXML. Should not crash.
		const runsJson = JSON.stringify([
			{ type: "text", text: "Before " },
			{ type: "equation", text: "x_i", display: false },
			{ type: "text", text: " middle " },
			{ type: "footnoteRef", kind: "footnote", id: "fn1" },
			{ type: "chart", kind: "chart" },
			{ type: "text", text: " after" },
		]);
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--runs",
			runsJson,
		);
		expect(result.exitCode).toBe(0);
		const read = await runCli("read", docPath);
		const doc = read.parsed as {
			blocks: Array<{ type: string; runs?: Array<{ text?: string }> }>;
		};
		const lastParagraph = doc.blocks
			.filter((block) => block.type === "paragraph")
			.pop();
		const texts = (lastParagraph?.runs ?? [])
			.map((run) => run.text)
			.filter((text): text is string => text !== undefined);
		expect(texts.join("")).toBe("Before  middle  after");
	});

	test("insert --runs supports mixed-format paragraph", async () => {
		const runsJson = JSON.stringify([
			{ type: "text", text: "Mix: " },
			{ type: "text", text: "red", color: "CC0000" },
			{ type: "text", text: " / " },
			{ type: "text", text: "bold", bold: true },
		]);
		await runCli("insert", docPath, "--after", "p0", "--runs", runsJson);
		const read = await runCli("read", docPath);
		const doc = read.parsed as {
			blocks: Array<{
				type: string;
				runs?: Array<{ text: string; color?: string; bold?: boolean }>;
			}>;
		};
		const lastParagraph = doc.blocks
			.filter((block) => block.type === "paragraph")
			.pop();
		const runs = lastParagraph?.runs ?? [];
		expect(runs[0]?.text).toBe("Mix: ");
		expect(runs[1]).toMatchObject({ text: "red", color: "CC0000" });
		expect(runs[3]).toMatchObject({ text: "bold", bold: true });
	});

	test("edit replaces a paragraph at the locator", async () => {
		await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--text",
			"Replaced",
			"--style",
			"Heading1",
		);
		const read = await runCli("read", docPath);
		const doc = read.parsed as {
			blocks: Array<{
				type: string;
				style?: string;
				runs?: Array<{ text: string }>;
			}>;
		};
		const paragraph = doc.blocks.find((block) => block.type === "paragraph");
		expect(paragraph?.style).toBe("Heading1");
		expect(paragraph?.runs?.[0]?.text).toBe("Replaced");
	});

	test("delete removes the block at the locator", async () => {
		await runCli("insert", docPath, "--after", "p0", "--text", "Second");
		const beforeRead = await runCli("read", docPath);
		const before = beforeRead.parsed as {
			blocks: Array<{ type: string }>;
		};
		const beforeCount = before.blocks.filter(
			(block) => block.type === "paragraph",
		).length;

		await runCli("delete", docPath, "--at", "p0");
		const afterRead = await runCli("read", docPath);
		const after = afterRead.parsed as {
			blocks: Array<{ type: string; runs?: Array<{ text: string }> }>;
		};
		const afterCount = after.blocks.filter(
			(block) => block.type === "paragraph",
		).length;
		expect(afterCount).toBe(beforeCount - 1);
		const remaining = after.blocks.find((block) => block.type === "paragraph");
		expect(remaining?.runs?.[0]?.text).toBe("Second");
	});

	test("dry-run does not modify the file", async () => {
		const before = await Bun.file(docPath).arrayBuffer();
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Should not appear",
			"--dry-run",
		);
		const after = await Bun.file(docPath).arrayBuffer();
		expect(after.byteLength).toBe(before.byteLength);
	});

	test("--output writes to a parallel file and leaves FILE untouched", async () => {
		const beforeBytes = await Bun.file(docPath).arrayBuffer();
		const outPath = `${docPath}.copy.docx`;
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Out-of-band",
			"-o",
			outPath,
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { path: string }).path).toBe(outPath);

		const afterBytes = await Bun.file(docPath).arrayBuffer();
		expect(afterBytes.byteLength).toBe(beforeBytes.byteLength);
		expect(await Bun.file(outPath).exists()).toBe(true);

		const read = await runCli("read", outPath);
		const doc = read.parsed as {
			blocks: Array<{ type: string; runs?: Array<{ text: string }> }>;
		};
		const paragraphs = doc.blocks.filter((block) => block.type === "paragraph");
		expect(paragraphs[1]?.runs?.[0]?.text).toBe("Out-of-band");
	});

	test("--dry-run with --output writes nothing and echoes the intended output", async () => {
		const sourceBytes = await Bun.file(docPath).arrayBuffer();
		const outPath = `${docPath}.copy.docx`;
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Should not appear",
			"-o",
			outPath,
			"--dry-run",
		);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			ok: true,
			dryRun: true,
			path: docPath,
			output: outPath,
		});
		const afterBytes = await Bun.file(docPath).arrayBuffer();
		expect(afterBytes.byteLength).toBe(sourceBytes.byteLength);
		expect(await Bun.file(outPath).exists()).toBe(false);
	});

	test("invalid locator returns block-not-found", async () => {
		const result = await runCli("edit", docPath, "--at", "p99", "--text", "x");
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({
			ok: false,
			code: "BLOCK_NOT_FOUND",
		});
	});
});
