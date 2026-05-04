import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

const FIXTURE = "tests/fixtures/tracked-changes.docx";
// Layout (paragraph p0):
//   "This is a text with " (0..20, plain)
//   "two exciting "        (20..33, inside <w:ins>)
//   "insertions."          (33..44, plain)

async function freshCopy(label: string): Promise<string> {
	const workspace = tempWorkspace(label);
	const docPath = join(workspace, "doc.docx");
	await Bun.write(docPath, Bun.file(FIXTURE));
	return docPath;
}

async function flatText(docPath: string): Promise<string> {
	const read = await runCli("read", docPath);
	const blocks = (
		read.parsed as {
			blocks: Array<{ runs?: Array<{ type: string; text: string }> }>;
		}
	).blocks;
	return blocks
		.flatMap((block) => block.runs ?? [])
		.filter((run) => run.type === "text")
		.map((run) => run.text)
		.join("");
}

describe("docx find — tracked changes", () => {
	test("locates text inside a <w:ins> wrapper at its paragraph offset", async () => {
		const docPath = await freshCopy("find-tc");
		const result = await runCli("find", docPath, "exciting");
		const payload = result.parsed as {
			matches: Array<{ locator: string; start: number; end: number }>;
		};
		expect(payload.matches).toHaveLength(1);
		expect(payload.matches[0]?.locator).toBe("p0:24-32");
		expect(payload.matches[0]?.start).toBe(24);
		expect(payload.matches[0]?.end).toBe(32);
	});

	test("offset model is consistent across plain + tracked-change runs", async () => {
		const docPath = await freshCopy("find-tc-after");
		const result = await runCli("find", docPath, "insertions");
		const payload = result.parsed as {
			matches: Array<{ start: number; end: number }>;
		};
		// "insertions" begins at offset 33 (after the 13-char ins block).
		expect(payload.matches[0]?.start).toBe(33);
		expect(payload.matches[0]?.end).toBe(43);
	});
});

describe("docx replace — tracked changes", () => {
	test("replacement inside an <w:ins> stays inside the wrapper", async () => {
		const docPath = await freshCopy("replace-in-ins");
		const result = await runCli("replace", docPath, "exciting", "delightful");
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath);
		const blocks = (
			read.parsed as {
				blocks: Array<{
					runs?: Array<{
						type: string;
						text: string;
						trackedChange?: { kind: string };
					}>;
				}>;
			}
		).blocks;
		const insRun = blocks[0]?.runs?.find((run) => run.text === "delightful");
		expect(insRun?.trackedChange?.kind).toBe("ins");
	});

	test("rejects span that crosses a tracked-change boundary", async () => {
		const docPath = await freshCopy("replace-cross-tc");
		const result = await runCli("replace", docPath, "with two", "with three");
		expect(result.exitCode).toBe(1);
		expect(result.parsed).toMatchObject({
			ok: false,
			code: "TRACKED_CHANGE_CONFLICT",
		});
	});

	test("plain-text replacement before a tracked change still works", async () => {
		const docPath = await freshCopy("replace-before-tc");
		const result = await runCli("replace", docPath, "text", "string");
		expect(result.exitCode).toBe(0);

		const flat = await flatText(docPath);
		expect(flat).toBe("This is a string with two exciting insertions.");
	});

	test("plain-text replacement after a tracked change still works", async () => {
		const docPath = await freshCopy("replace-after-tc");
		const result = await runCli("replace", docPath, "insertions", "edits");
		expect(result.exitCode).toBe(0);

		const flat = await flatText(docPath);
		expect(flat).toBe("This is a text with two exciting edits.");
	});
});

describe("docx comments add — tracked changes", () => {
	test("anchors a comment fully inside a <w:ins> wrapper", async () => {
		const docPath = await freshCopy("comments-in-ins");
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--range",
			"p0:24-32",
			"--text",
			"Inside ins",
			"--author",
			"QA",
		);
		expect(result.exitCode).toBe(0);

		const list = await runCli("comments", "list", docPath);
		const comments = list.parsed as Array<{
			anchor: { startOffset: number; endOffset: number };
		}>;
		expect(comments[0]?.anchor.startOffset).toBe(24);
		expect(comments[0]?.anchor.endOffset).toBe(32);
	});

	test("anchors a comment that crosses a tracked-change boundary", async () => {
		const docPath = await freshCopy("comments-cross-tc");
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--range",
			"p0:15-35",
			"--text",
			"Spans the tracked change",
			"--author",
			"QA",
		);
		expect(result.exitCode).toBe(0);

		const list = await runCli("comments", "list", docPath);
		const comments = list.parsed as Array<{
			anchor: { startOffset: number; endOffset: number };
		}>;
		expect(comments[0]?.anchor.startOffset).toBe(15);
		expect(comments[0]?.anchor.endOffset).toBe(35);
	});
});
