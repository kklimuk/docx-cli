import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

describe("docx insert --page-break / --column-break", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("breaks");
		docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "First");
	});

	test("--page-break inserts a paragraph with a single page break run", async () => {
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--page-break",
		);
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath);
		const doc = read.parsed as {
			blocks: Array<{
				id: string;
				type: string;
				runs?: Array<{ type: string; kind?: string }>;
			}>;
		};
		const paragraphs = doc.blocks.filter((block) => block.type === "paragraph");
		expect(paragraphs[1]?.runs).toEqual([{ type: "break", kind: "page" }]);
	});

	test("--column-break inserts a paragraph with a single column break run", async () => {
		await runCli("insert", docPath, "--before", "p0", "--column-break");
		const read = await runCli("read", docPath);
		const doc = read.parsed as {
			blocks: Array<{
				id: string;
				type: string;
				runs?: Array<{ type: string; kind?: string }>;
			}>;
		};
		const paragraphs = doc.blocks.filter((block) => block.type === "paragraph");
		expect(paragraphs[0]?.runs).toEqual([{ type: "break", kind: "column" }]);
	});

	test("rejects --page-break alongside --text", async () => {
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--page-break",
			"--text",
			"x",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({
			ok: false,
			code: "USAGE",
		});
	});

	test("rejects --page-break alongside --column-break", async () => {
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--page-break",
			"--column-break",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({
			ok: false,
			code: "USAGE",
		});
	});

	test("requires content flag", async () => {
		const result = await runCli("insert", docPath, "--after", "p0");
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({
			ok: false,
			code: "USAGE",
		});
	});
});
