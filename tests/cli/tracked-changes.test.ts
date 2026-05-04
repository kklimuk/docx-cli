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

describe("docx find — tracked-change overlap", () => {
	test("match inside a tracked insertion reports trackedChanges in the result", async () => {
		const docPath = await freshCopy("find-overlap-ins");
		const result = await runCli("find", docPath, "exciting");
		expect(result.exitCode).toBe(0);

		const payload = result.parsed as {
			matches: Array<{
				locator: string;
				trackedChanges?: Array<{ kind: string; author: string }>;
			}>;
		};
		expect(payload.matches[0]?.trackedChanges).toBeDefined();
		expect(payload.matches[0]?.trackedChanges?.[0]?.kind).toBe("ins");
		expect(payload.matches[0]?.trackedChanges?.[0]?.author).toBe("eng-dept");
	});

	test("match in plain text has no trackedChanges field", async () => {
		const docPath = await freshCopy("find-overlap-plain");
		const result = await runCli("find", docPath, "insertions");
		expect(result.exitCode).toBe(0);

		const payload = result.parsed as {
			matches: Array<{ trackedChanges?: unknown }>;
		};
		expect(payload.matches[0]?.trackedChanges).toBeUndefined();
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

	test("span crossing a tracked-change boundary replaces and splits the wrapper", async () => {
		const docPath = await freshCopy("replace-cross-tc");
		const result = await runCli("replace", docPath, "with two", "with three");
		expect(result.exitCode).toBe(0);

		const flat = await flatText(docPath);
		expect(flat).toBe("This is a text with three exciting insertions.");
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

	test("with track-changes flag on, plain-text replace emits del + ins", async () => {
		const docPath = await freshCopy("replace-tracked-plain");
		await runCli("track-changes", docPath, "on");
		const result = await runCli("replace", docPath, "text", "string");
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath);
		const runs = (
			read.parsed as {
				blocks: Array<{
					runs?: Array<{
						type: string;
						text: string;
						trackedChange?: { kind: string };
					}>;
				}>;
			}
		).blocks[0]?.runs;

		const insertedRun = runs?.find((run) => run.text === "string");
		expect(insertedRun?.trackedChange?.kind).toBe("ins");

		const deletedRun = runs?.find(
			(run) => run.text === "text" && run.trackedChange?.kind === "del",
		);
		expect(deletedRun).toBeDefined();
	});
});

describe("docx delete — tracked changes", () => {
	test("with track-changes flag on, paragraph delete wraps runs in <w:del>", async () => {
		const docPath = await freshCopy("delete-tracked");
		await runCli("track-changes", docPath, "on");
		const result = await runCli("delete", docPath, "--at", "p0");
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath);
		const blocks = (
			read.parsed as {
				blocks: Array<{
					type: string;
					runs?: Array<{
						type: string;
						text: string;
						trackedChange?: { kind: string };
					}>;
				}>;
			}
		).blocks;

		const paragraph = blocks.find((b) => b.type === "paragraph");
		expect(paragraph).toBeDefined();
		const deletedRun = paragraph?.runs?.find(
			(r) => r.trackedChange?.kind === "del",
		);
		expect(deletedRun).toBeDefined();
	});
});

describe("docx edit — tracked changes", () => {
	test("with track-changes flag on, edit preserves old runs as <w:del> + new as <w:ins>", async () => {
		const docPath = await freshCopy("edit-tracked");
		await runCli("track-changes", docPath, "on");
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--text",
			"Replaced",
		);
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath);
		const runs = (
			read.parsed as {
				blocks: Array<{
					runs?: Array<{
						type: string;
						text: string;
						trackedChange?: { kind: string };
					}>;
				}>;
			}
		).blocks[0]?.runs;

		const insertedRun = runs?.find((r) => r.text === "Replaced");
		expect(insertedRun?.trackedChange?.kind).toBe("ins");
		const deletedRun = runs?.find(
			(r) => r.trackedChange?.kind === "del" && r.text.includes("This is a"),
		);
		expect(deletedRun).toBeDefined();
	});
});

describe("docx insert — tracked changes", () => {
	test("with track-changes flag on, insert wraps new runs in <w:ins>", async () => {
		const docPath = await freshCopy("insert-tracked");
		await runCli("track-changes", docPath, "on");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Inserted",
		);
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath);
		const blocks = (
			read.parsed as {
				blocks: Array<{
					type: string;
					runs?: Array<{
						type: string;
						text: string;
						trackedChange?: { kind: string };
					}>;
				}>;
			}
		).blocks;

		const newParagraph = blocks.filter((b) => b.type === "paragraph")[1];
		const insertedRun = newParagraph?.runs?.find((r) => r.text === "Inserted");
		expect(insertedRun?.trackedChange?.kind).toBe("ins");
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
