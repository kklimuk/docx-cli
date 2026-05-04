import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

describe("docx comments", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("comments");
		docPath = join(workspace, "out.docx");
		await runCli(
			"create",
			docPath,
			"--text",
			"The quick brown fox jumps over the lazy dog.",
		);
	});

	test("add anchors a comment to the whole paragraph", async () => {
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--range",
			"p0",
			"--text",
			"Reconsider this sentence",
			"--author",
			"Reviewer",
		);
		expect(result.parsed).toMatchObject({
			ok: true,
			operation: "comments.add",
			commentId: "c0",
		});

		const list = await runCli("comments", "list", docPath);
		const comments = list.parsed as Array<{
			id: string;
			author: string;
			text: string;
			anchor: { startOffset: number; endOffset: number };
		}>;
		expect(comments[0]).toMatchObject({
			id: "c0",
			author: "Reviewer",
			text: "Reconsider this sentence",
		});
		expect(comments[0]?.anchor.startOffset).toBe(0);
		expect(comments[0]?.anchor.endOffset).toBe(44);
	});

	test("add with span splits runs at offsets", async () => {
		await runCli(
			"comments",
			"add",
			docPath,
			"--range",
			"p0:16-19",
			"--text",
			"fox?",
			"--author",
			"Jane",
		);
		const read = await runCli("read", docPath);
		const doc = read.parsed as {
			blocks: Array<{
				type: string;
				runs?: Array<{ type: string; text: string; comments?: string[] }>;
			}>;
		};
		const paragraph = doc.blocks.find((block) => block.type === "paragraph");
		expect(paragraph?.runs).toEqual([
			{ type: "text", text: "The quick brown " },
			{ type: "text", text: "fox", comments: ["c0"] },
			{ type: "text", text: " jumps over the lazy dog." },
		]);
	});

	test("reply links to parent via parentId", async () => {
		await runCli(
			"comments",
			"add",
			docPath,
			"--range",
			"p0",
			"--text",
			"Q?",
			"--author",
			"A",
		);
		await runCli(
			"comments",
			"reply",
			docPath,
			"--to",
			"c0",
			"--text",
			"Answer",
			"--author",
			"B",
		);
		const list = await runCli("comments", "list", docPath, "--thread", "c0");
		const comments = list.parsed as Array<{ id: string; parentId?: string }>;
		expect(comments).toHaveLength(2);
		expect(comments[1]).toMatchObject({ id: "c1", parentId: "c0" });
	});

	test("resolve flips done flag and filter excludes by default", async () => {
		await runCli(
			"comments",
			"add",
			docPath,
			"--range",
			"p0",
			"--text",
			"x",
			"--author",
			"A",
		);
		await runCli("comments", "resolve", docPath, "--id", "c0");

		const defaultList = await runCli("comments", "list", docPath);
		expect((defaultList.parsed as unknown[]).length).toBe(0);

		const allList = await runCli(
			"comments",
			"list",
			docPath,
			"--include-resolved",
		);
		const all = allList.parsed as Array<{ resolved?: boolean }>;
		expect(all[0]?.resolved).toBe(true);
	});

	test("delete + restore round-trips", async () => {
		await runCli(
			"comments",
			"add",
			docPath,
			"--range",
			"p0",
			"--text",
			"hi",
			"--author",
			"A",
		);
		await runCli("comments", "delete", docPath, "--id", "c0");
		const afterDelete = await runCli(
			"comments",
			"list",
			docPath,
			"--include-resolved",
		);
		expect((afterDelete.parsed as unknown[]).length).toBe(0);

		await runCli("comments", "restore", docPath, "--id", "c0");
		const afterRestore = await runCli(
			"comments",
			"list",
			docPath,
			"--include-resolved",
		);
		const restored = afterRestore.parsed as Array<{
			id: string;
			text: string;
		}>;
		expect(restored[0]).toMatchObject({ id: "c0", text: "hi" });
	});

	test("add anchors a comment across two paragraphs", async () => {
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Second paragraph here.",
		);
		await runCli(
			"insert",
			docPath,
			"--after",
			"p1",
			"--text",
			"Third paragraph here.",
		);

		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--range",
			"p0:16-p2:6",
			"--text",
			"Spans three paragraphs",
			"--author",
			"Reviewer",
		);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			ok: true,
			operation: "comments.add",
			commentId: "c0",
			locator: "p0:16-p2:6",
		});

		const list = await runCli("comments", "list", docPath);
		const comments = list.parsed as Array<{
			id: string;
			text: string;
			anchor: {
				startBlockId: string;
				startOffset: number;
				endBlockId: string;
				endOffset: number;
			};
		}>;
		expect(comments[0]).toMatchObject({
			id: "c0",
			text: "Spans three paragraphs",
			anchor: {
				startBlockId: "p0",
				startOffset: 16,
				endBlockId: "p2",
				endOffset: 6,
			},
		});
	});

	test("delete + restore round-trips a cross-block comment", async () => {
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Second paragraph here.",
		);
		await runCli(
			"comments",
			"add",
			docPath,
			"--range",
			"p0:16-p1:6",
			"--text",
			"two paragraphs",
			"--author",
			"A",
		);
		await runCli("comments", "delete", docPath, "--id", "c0");
		const afterDelete = await runCli(
			"comments",
			"list",
			docPath,
			"--include-resolved",
		);
		expect((afterDelete.parsed as unknown[]).length).toBe(0);

		await runCli("comments", "restore", docPath, "--id", "c0");
		const afterRestore = await runCli(
			"comments",
			"list",
			docPath,
			"--include-resolved",
		);
		const restored = afterRestore.parsed as Array<{
			id: string;
			text: string;
			anchor: { startBlockId: string; endBlockId: string };
		}>;
		expect(restored[0]).toMatchObject({
			id: "c0",
			text: "two paragraphs",
			anchor: { startBlockId: "p0", endBlockId: "p1" },
		});
	});

	test("add rejects cross-block range with out-of-range offset", async () => {
		await runCli("insert", docPath, "--after", "p0", "--text", "Short.");
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--range",
			"p0:0-p1:999",
			"--text",
			"too far",
			"--author",
			"A",
		);
		expect(result.exitCode).not.toBe(0);
		expect(result.parsed).toMatchObject({
			ok: false,
			code: "INVALID_LOCATOR",
		});
	});

	test("resolve auto-injects paraId on legacy comments", async () => {
		// comments-simple.docx is a mammoth-authored fixture without paraIds
		const workspace = tempWorkspace("legacy");
		const docCopy = join(workspace, "legacy.docx");
		await Bun.write(docCopy, Bun.file("tests/fixtures/comments-simple.docx"));
		const result = await runCli("comments", "resolve", docCopy, "--id", "c0");
		expect(result.exitCode).toBe(0);
		const list = await runCli(
			"comments",
			"list",
			docCopy,
			"--include-resolved",
		);
		const found = (
			list.parsed as Array<{ id: string; resolved?: boolean }>
		).find((comment) => comment.id === "c0");
		expect(found?.resolved).toBe(true);
	});
});
