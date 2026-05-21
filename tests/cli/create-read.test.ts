import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

describe("docx create + read", () => {
	test("creates a minimal docx and reads it back", async () => {
		const workspace = tempWorkspace("create-read");
		const docPath = join(workspace, "out.docx");

		const create = await runCli(
			"create",
			docPath,
			"--title",
			"Test",
			"--author",
			"Tester",
			"--text",
			"Hello world",
		);
		expect(create.exitCode).toBe(0);
		expect(create.parsed).toMatchObject({
			ok: true,
			operation: "create",
			path: docPath,
		});

		const read = await runCli("read", docPath, "--ast");
		expect(read.exitCode).toBe(0);
		const doc = read.parsed as {
			properties: { title: string; author: string };
			blocks: Array<{ type: string; runs?: Array<{ text: string }> }>;
		};
		expect(doc.properties.title).toBe("Test");
		expect(doc.properties.author).toBe("Tester");
		const firstParagraph = doc.blocks.find(
			(block) => block.type === "paragraph",
		);
		expect(firstParagraph?.runs?.[0]?.text).toBe("Hello world");
	});

	test("escapes hostile XML in title/author/text round-trip", async () => {
		const workspace = tempWorkspace("escape");
		const docPath = join(workspace, "out.docx");

		await runCli(
			"create",
			docPath,
			"--title",
			'Has "quotes" & <stuff>',
			"--author",
			"<script>",
			"--text",
			"Body: <input> & 'data'",
		);
		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as {
			properties: { title: string; author: string };
			blocks: Array<{ type: string; runs?: Array<{ text: string }> }>;
		};
		expect(doc.properties.title).toBe('Has "quotes" & <stuff>');
		expect(doc.properties.author).toBe("<script>");
		const paragraph = doc.blocks.find((block) => block.type === "paragraph");
		expect(paragraph?.runs?.[0]?.text).toBe("Body: <input> & 'data'");
	});

	test("read on missing file returns not-found error", async () => {
		const result = await runCli("read", "/tmp/does-not-exist.docx");
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({
			ok: false,
			code: "FILE_NOT_FOUND",
		});
	});

	test("create rejects existing file without --force", async () => {
		const workspace = tempWorkspace("force");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "first");
		const second = await runCli("create", docPath, "--text", "second");
		expect(second.exitCode).toBe(2);
		expect(second.parsed).toMatchObject({ ok: false, code: "USAGE" });
	});
});
