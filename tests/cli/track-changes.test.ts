import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";
import { freshFixture as copyFixture } from "./helpers";

describe("docx track-changes", () => {
	test("on creates settings.xml and registers it", async () => {
		const workspace = tempWorkspace("track");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "hi");

		const result = await runCli("track-changes", docPath, "on");
		expect(result.parsed).toMatchObject({
			ok: true,
			operation: "track-changes",
			mode: "on",
			previouslyOn: false,
		});

		const xml = await (await Pkg.open(docPath)).readText("word/settings.xml");
		expect(xml).toContain("<w:trackChanges/>");
		expect(xml).toContain("xmlns:w=");
	});

	test("off removes the trackChanges element", async () => {
		const workspace = tempWorkspace("untrack");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "hi");
		await runCli("track-changes", docPath, "on");
		const result = await runCli("track-changes", docPath, "off");
		expect(result.parsed).toMatchObject({
			mode: "off",
			previouslyOn: true,
		});

		const xml = await (await Pkg.open(docPath)).readText("word/settings.xml");
		expect(xml).not.toContain("<w:trackChanges/>");
	});

	test("rejects invalid mode", async () => {
		const workspace = tempWorkspace("invalid");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "hi");
		const result = await runCli("track-changes", docPath, "maybe");
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("list returns every <w:ins>/<w:del> with metadata", async () => {
		const result = await runCli(
			"track-changes",
			"list",
			"tests/fixtures/tracked-changes.docx",
		);
		expect(result.exitCode).toBe(0);
		const changes = result.parsed as Array<{
			id: string;
			kind: string;
			author: string;
			date: string;
			blockId: string;
			text: string;
		}>;
		expect(changes).toHaveLength(1);
		expect(changes[0]).toMatchObject({
			id: "tc0",
			kind: "ins",
			author: "eng-dept",
			date: "2014-06-25T10:40:00Z",
			blockId: "p0",
			text: "two exciting ",
		});
	});

	test("list on a doc without tracked changes returns []", async () => {
		const result = await runCli(
			"track-changes",
			"list",
			"tests/fixtures/minimal.docx",
		);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toEqual([]);
	});
});

describe("docx track-changes accept / reject", () => {
	async function freshFixture(label: string): Promise<string> {
		const workspace = tempWorkspace(label);
		const docPath = join(workspace, "doc.docx");
		await Bun.write(docPath, Bun.file("tests/fixtures/tracked-changes.docx"));
		return docPath;
	}

	async function flatText(docPath: string): Promise<string> {
		const read = await runCli("read", docPath, "--ast");
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

	async function buildMixed(label: string): Promise<string> {
		const workspace = tempWorkspace(label);
		const docPath = join(workspace, "doc.docx");
		await runCli("create", docPath, "--text", "The quick brown fox jumps.");
		await runCli("track-changes", docPath, "on");
		// Yields a <w:del> for "quick brown " and a <w:ins> for "old slow ".
		await runCli("replace", docPath, "quick brown ", "old slow ");
		return docPath;
	}

	test("accept --at tcN unwraps an insertion (text becomes plain)", async () => {
		const docPath = await freshFixture("accept-at-ins");
		const result = await runCli(
			"track-changes",
			"accept",
			docPath,
			"--at",
			"tc0",
		);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			ok: true,
			operation: "track-changes.accept",
			applied: [{ id: "tc0", kind: "ins", action: "unwrap" }],
		});

		const list = await runCli("track-changes", "list", docPath);
		expect(list.parsed).toEqual([]);
		expect(await flatText(docPath)).toBe(
			"This is a text with two exciting insertions.",
		);
	});

	test("reject --at tcN on an insertion deletes the inserted text", async () => {
		const docPath = await freshFixture("reject-at-ins");
		const result = await runCli(
			"track-changes",
			"reject",
			docPath,
			"--at",
			"tc0",
		);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			ok: true,
			operation: "track-changes.reject",
			applied: [{ id: "tc0", kind: "ins", action: "delete" }],
		});

		const list = await runCli("track-changes", "list", docPath);
		expect(list.parsed).toEqual([]);
		expect(await flatText(docPath)).toBe("This is a text with insertions.");
	});

	test("accept --all on mixed ins+del yields the post-accept text", async () => {
		const docPath = await buildMixed("accept-all-mixed");
		const result = await runCli("track-changes", "accept", docPath, "--all");
		expect(result.exitCode).toBe(0);
		const applied = (
			result.parsed as {
				applied: Array<{ id: string; kind: string; action: string }>;
			}
		).applied;
		expect(applied).toHaveLength(2);
		// Reverse-order processing: tc1 (ins) before tc0 (del).
		expect(applied.map((a) => a.id).sort()).toEqual(["tc0", "tc1"]);
		expect(applied.find((a) => a.kind === "del")?.action).toBe("delete");
		expect(applied.find((a) => a.kind === "ins")?.action).toBe("unwrap");

		expect(await flatText(docPath)).toBe("The old slow fox jumps.");
		expect((await runCli("track-changes", "list", docPath)).parsed).toEqual([]);
	});

	test("reject --all on mixed ins+del reverts to the baseline text", async () => {
		const docPath = await buildMixed("reject-all-mixed");
		const result = await runCli("track-changes", "reject", docPath, "--all");
		expect(result.exitCode).toBe(0);
		const applied = (
			result.parsed as { applied: Array<{ kind: string; action: string }> }
		).applied;
		expect(applied.find((a) => a.kind === "ins")?.action).toBe("delete");
		expect(applied.find((a) => a.kind === "del")?.action).toBe("unwrap");

		expect(await flatText(docPath)).toBe("The quick brown fox jumps.");
		expect((await runCli("track-changes", "list", docPath)).parsed).toEqual([]);
	});

	test("--dry-run reports the plan without mutating the file", async () => {
		const docPath = await freshFixture("dry-run-accept");
		const result = await runCli(
			"track-changes",
			"accept",
			docPath,
			"--at",
			"tc0",
			"--dry-run",
		);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			dryRun: true,
			operation: "track-changes.accept",
			applied: [{ id: "tc0", kind: "ins", action: "unwrap" }],
		});

		const list = await runCli("track-changes", "list", docPath);
		expect(list.parsed).toMatchObject([{ id: "tc0" }]);
	});

	test("--at tcN with unknown id returns TRACKED_CHANGE_NOT_FOUND", async () => {
		const docPath = await freshFixture("missing-id");
		const result = await runCli(
			"track-changes",
			"accept",
			docPath,
			"--at",
			"tc99",
		);
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({
			code: "TRACKED_CHANGE_NOT_FOUND",
		});
	});

	test("--at and --all together are rejected", async () => {
		const docPath = await freshFixture("conflict");
		const result = await runCli(
			"track-changes",
			"accept",
			docPath,
			"--at",
			"tc0",
			"--all",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("missing --at and --all is rejected", async () => {
		const docPath = await freshFixture("missing-target");
		const result = await runCli("track-changes", "accept", docPath);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("--all on a doc with no tracked changes succeeds with empty applied", async () => {
		const workspace = tempWorkspace("accept-empty");
		const docPath = join(workspace, "doc.docx");
		await Bun.write(docPath, Bun.file("tests/fixtures/minimal.docx"));
		const result = await runCli("track-changes", "accept", docPath, "--all");
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			ok: true,
			operation: "track-changes.accept",
			applied: [],
		});
	});

	test("-o/--output writes to a parallel file, original untouched", async () => {
		const docPath = await freshFixture("accept-output");
		const outPath = `${docPath}.out.docx`;
		const result = await runCli(
			"track-changes",
			"accept",
			docPath,
			"--at",
			"tc0",
			"--output",
			outPath,
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { path: string }).path).toBe(outPath);

		// Original still has the change.
		expect(
			(
				(await runCli("track-changes", "list", docPath)).parsed as Array<{
					id: string;
				}>
			).length,
		).toBe(1);
		// Output has it accepted.
		expect((await runCli("track-changes", "list", outPath)).parsed).toEqual([]);
	});
});

const TRACKED_FIX = "tests/fixtures/tracked-changes.docx";
// Layout (paragraph p0):
//   "This is a text with " (0..20, plain)
//   "two exciting "        (20..33, inside <w:ins>)
//   "insertions."          (33..44, plain)

const freshTracked = (label: string) => copyFixture(label, TRACKED_FIX);

async function flatText(docPath: string): Promise<string> {
	const read = await runCli("read", docPath, "--ast");
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
		const docPath = await freshTracked("find-tc");
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
		const docPath = await freshTracked("find-tc-after");
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
		const docPath = await freshTracked("find-overlap-ins");
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
		const docPath = await freshTracked("find-overlap-plain");
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
		const docPath = await freshTracked("replace-in-ins");
		const result = await runCli("replace", docPath, "exciting", "delightful");
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath, "--ast");
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
		const docPath = await freshTracked("replace-cross-tc");
		const result = await runCli("replace", docPath, "with two", "with three");
		expect(result.exitCode).toBe(0);

		const flat = await flatText(docPath);
		expect(flat).toBe("This is a text with three exciting insertions.");
	});

	test("plain-text replacement before a tracked change still works", async () => {
		const docPath = await freshTracked("replace-before-tc");
		const result = await runCli("replace", docPath, "text", "string");
		expect(result.exitCode).toBe(0);

		const flat = await flatText(docPath);
		expect(flat).toBe("This is a string with two exciting insertions.");
	});

	test("plain-text replacement after a tracked change still works", async () => {
		const docPath = await freshTracked("replace-after-tc");
		const result = await runCli("replace", docPath, "insertions", "edits");
		expect(result.exitCode).toBe(0);

		const flat = await flatText(docPath);
		expect(flat).toBe("This is a text with two exciting edits.");
	});

	test("with track-changes flag on, plain-text replace emits del + ins", async () => {
		const docPath = await freshTracked("replace-tracked-plain");
		await runCli("track-changes", docPath, "on");
		const result = await runCli("replace", docPath, "text", "string");
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath, "--ast");
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
		const docPath = await freshTracked("delete-tracked");
		await runCli("track-changes", docPath, "on");
		const result = await runCli("delete", docPath, "--at", "p0");
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath, "--ast");
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
		const docPath = await freshTracked("edit-tracked");
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

		const read = await runCli("read", docPath, "--ast");
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
		const docPath = await freshTracked("insert-tracked");
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

		const read = await runCli("read", docPath, "--ast");
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

describe("--author flag — tracked-change attribution", () => {
	type RunWithChange = {
		type: string;
		text: string;
		trackedChange?: { kind: string; author: string };
	};

	async function readRuns(docPath: string): Promise<RunWithChange[]> {
		const read = await runCli("read", docPath, "--ast");
		const blocks = (
			read.parsed as { blocks: Array<{ runs?: RunWithChange[] }> }
		).blocks;
		return blocks.flatMap((block) => block.runs ?? []);
	}

	test("insert --author overrides $DOCX_AUTHOR on tracked insertion", async () => {
		const docPath = await freshTracked("insert-author");
		await runCli("track-changes", docPath, "on");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Inserted",
			"--author",
			"Reviewer-One",
		);
		expect(result.exitCode).toBe(0);

		const runs = await readRuns(docPath);
		const inserted = runs.find((run) => run.text === "Inserted");
		expect(inserted?.trackedChange?.kind).toBe("ins");
		expect(inserted?.trackedChange?.author).toBe("Reviewer-One");
	});

	test("edit --author attributes both <w:del> and <w:ins>", async () => {
		const docPath = await freshTracked("edit-author");
		await runCli("track-changes", docPath, "on");
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--text",
			"Replaced",
			"--author",
			"Reviewer-Two",
		);
		expect(result.exitCode).toBe(0);

		const runs = await readRuns(docPath);
		const inserted = runs.find((run) => run.text === "Replaced");
		const deleted = runs.find((run) => run.trackedChange?.kind === "del");
		expect(inserted?.trackedChange?.author).toBe("Reviewer-Two");
		expect(deleted?.trackedChange?.author).toBe("Reviewer-Two");
	});

	test("delete --author attributes the tracked deletion", async () => {
		const docPath = await freshTracked("delete-author");
		await runCli("track-changes", docPath, "on");
		const result = await runCli(
			"delete",
			docPath,
			"--at",
			"p0",
			"--author",
			"Reviewer-Three",
		);
		expect(result.exitCode).toBe(0);

		const runs = await readRuns(docPath);
		const deleted = runs.find((run) => run.trackedChange?.kind === "del");
		expect(deleted?.trackedChange?.author).toBe("Reviewer-Three");
	});

	test("replace --author attributes the inserted replacement", async () => {
		const docPath = await freshTracked("replace-author");
		await runCli("track-changes", docPath, "on");
		const result = await runCli(
			"replace",
			docPath,
			"text",
			"string",
			"--author",
			"Reviewer-Four",
		);
		expect(result.exitCode).toBe(0);

		const runs = await readRuns(docPath);
		const inserted = runs.find(
			(run) => run.text === "string" && run.trackedChange?.kind === "ins",
		);
		expect(inserted?.trackedChange?.author).toBe("Reviewer-Four");
	});
});

describe("docx comments add — tracked changes", () => {
	test("anchors a comment fully inside a <w:ins> wrapper", async () => {
		const docPath = await freshTracked("comments-in-ins");
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--at",
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
		const docPath = await freshTracked("comments-cross-tc");
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--at",
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

describe("audit-comments — hyperlinks/images under track-changes", () => {
	type CommentRow = {
		id: string;
		author: string;
		text: string;
		anchor: {
			startBlockId: string;
			endBlockId: string;
			startOffset: number;
			endOffset: number;
		};
	};

	async function listComments(docPath: string): Promise<CommentRow[]> {
		const result = await runCli("comments", "list", docPath);
		return result.parsed as CommentRow[];
	}

	async function listAuditComments(docPath: string): Promise<CommentRow[]> {
		const all = await listComments(docPath);
		return all.filter((comment) => comment.text.startsWith("[docx-cli]"));
	}

	test("hyperlinks add emits an audit comment anchored to the wrapped span", async () => {
		const workspace = tempWorkspace("hl-add-audit");
		const docPath = join(workspace, "doc.docx");
		await Bun.write(docPath, Bun.file("tests/fixtures/minimal.docx"));
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"the quick brown fox jumps",
		);
		await runCli("track-changes", docPath, "on");

		const result = await runCli(
			"hyperlinks",
			"add",
			docPath,
			"--at",
			"p1:10-15",
			"--url",
			"https://example.com/brown",
			"--author",
			"Auditor",
		);
		expect(result.exitCode).toBe(0);

		const comments = await listAuditComments(docPath);
		expect(comments).toHaveLength(1);
		expect(comments[0]?.author).toBe("Auditor");
		expect(comments[0]?.text).toBe(
			"[docx-cli] hyperlink added → https://example.com/brown",
		);
		expect(comments[0]?.anchor.startBlockId).toBe("p1");
		expect(comments[0]?.anchor.endBlockId).toBe("p1");
		expect(comments[0]?.anchor.startOffset).toBe(10);
		expect(comments[0]?.anchor.endOffset).toBe(15);
	});

	test("hyperlinks add stays silent when track-changes is off", async () => {
		const workspace = tempWorkspace("hl-add-silent");
		const docPath = join(workspace, "doc.docx");
		await Bun.write(docPath, Bun.file("tests/fixtures/minimal.docx"));
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"the quick brown fox jumps",
		);
		const result = await runCli(
			"hyperlinks",
			"add",
			docPath,
			"--at",
			"p1:10-15",
			"--url",
			"https://example.com/brown",
		);
		expect(result.exitCode).toBe(0);
		expect(await listAuditComments(docPath)).toHaveLength(0);
	});

	test("hyperlinks replace emits an audit comment with old → new URLs", async () => {
		const workspace = tempWorkspace("hl-replace-audit");
		const docPath = join(workspace, "doc.docx");
		await Bun.write(docPath, Bun.file("tests/fixtures/minimal.docx"));
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"see this site for details",
		);
		await runCli(
			"hyperlinks",
			"add",
			docPath,
			"--at",
			"p1:4-8",
			"--url",
			"https://old.example.com",
		);
		await runCli("track-changes", docPath, "on");

		const result = await runCli(
			"hyperlinks",
			"replace",
			docPath,
			"--at",
			"link0",
			"--with",
			"https://new.example.com",
			"--author",
			"Reviewer",
		);
		expect(result.exitCode).toBe(0);

		const comments = await listAuditComments(docPath);
		expect(comments).toHaveLength(1);
		expect(comments[0]?.author).toBe("Reviewer");
		expect(comments[0]?.text).toBe(
			"[docx-cli] hyperlink target changed: https://old.example.com → https://new.example.com",
		);
		expect(comments[0]?.anchor.startBlockId).toBe("p1");
		expect(comments[0]?.anchor.endBlockId).toBe("p1");
		expect(comments[0]?.anchor.startOffset).toBe(4);
		expect(comments[0]?.anchor.endOffset).toBe(8);
	});

	test("hyperlinks delete emits an audit comment over the surviving text", async () => {
		const workspace = tempWorkspace("hl-delete-audit");
		const docPath = join(workspace, "doc.docx");
		await Bun.write(docPath, Bun.file("tests/fixtures/minimal.docx"));
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"see this site for details",
		);
		await runCli(
			"hyperlinks",
			"add",
			docPath,
			"--at",
			"p1:4-8",
			"--url",
			"https://gone.example.com",
		);
		await runCli("track-changes", docPath, "on");

		const result = await runCli(
			"hyperlinks",
			"delete",
			docPath,
			"--at",
			"link0",
			"--author",
			"Cleaner",
		);
		expect(result.exitCode).toBe(0);

		const comments = await listAuditComments(docPath);
		expect(comments).toHaveLength(1);
		expect(comments[0]?.author).toBe("Cleaner");
		expect(comments[0]?.text).toBe(
			"[docx-cli] hyperlink removed (was: https://gone.example.com)",
		);
		expect(comments[0]?.anchor.startBlockId).toBe("p1");
		expect(comments[0]?.anchor.endBlockId).toBe("p1");
		expect(comments[0]?.anchor.startOffset).toBe(4);
		expect(comments[0]?.anchor.endOffset).toBe(8);
	});

	test("images replace emits an audit comment per drawing using the rId", async () => {
		const workspace = tempWorkspace("img-replace-audit");
		const docPath = join(workspace, "doc.docx");
		await Bun.write(docPath, Bun.file("tests/fixtures/large-mixed.docx"));

		const before = await runCli("images", "list", docPath);
		const beforeList = before.parsed as Array<{ id: string; hash: string }>;
		const target = beforeList.find((image) => image.id === "img0");
		const replacementSrc = beforeList.find(
			(image) => image.id !== "img0" && image.hash !== target?.hash,
		);
		expect(replacementSrc).toBeDefined();
		const extractDir = join(workspace, "extracted");
		require("node:fs").mkdirSync(extractDir, { recursive: true });
		await runCli(
			"images",
			"extract",
			docPath,
			"--to",
			extractDir,
			"--at",
			replacementSrc?.id ?? "img1",
		);
		const replacementPath = join(extractDir, `${replacementSrc?.hash}.jpeg`);

		await runCli("track-changes", docPath, "on");

		const result = await runCli(
			"images",
			"replace",
			docPath,
			"--at",
			"img0",
			"--with",
			replacementPath,
			"--author",
			"ImageBot",
		);
		expect(result.exitCode).toBe(0);

		const auditComments = await listAuditComments(docPath);
		expect(auditComments.length).toBeGreaterThanOrEqual(1);
		const replacementComment = auditComments.find((comment) =>
			comment.text.startsWith("[docx-cli] image replaced:"),
		);
		expect(replacementComment).toBeDefined();
		expect(replacementComment?.author).toBe("ImageBot");
	});
});

// B2 — atomic batch accept/reject. `--at` is a `multiple: true` flag, so
// `accept --at tc1 --at tc2 --at tc3` resolves all targets against the
// pre-mutation tree. Mid-batch renumbering doesn't shift the still-pending
// ids out from under the agent.

const MULTI_FIX = "tests/fixtures/multi-tracked.docx";
// Layout (built by tests/fixtures/setup/multi-tracked.ts):
//   p0: "Aleph is first."  — tc0 (del "Alpha") + tc1 (ins "Aleph")
//   p1: "Bet is second."   — tc2 (del "Beta")  + tc3 (ins "Bet")
//   p2: "Gimel is third."  — tc4 (del "Gamma") + tc5 (ins "Gimel")
// Six tracked changes, ids tc0..tc5 in document order. Picking
// non-adjacent ids (tc0, tc2, tc4) targets one wrapper per paragraph.

const freshMulti = (label: string) => copyFixture(label, MULTI_FIX);

async function listTracked(
	docPath: string,
): Promise<Array<{ id: string; kind: string }>> {
	const result = await runCli("track-changes", "list", docPath);
	return result.parsed as Array<{ id: string; kind: string }>;
}

async function paragraphTextDoc(
	docPath: string,
	blockId: string,
): Promise<string> {
	const read = await runCli("read", docPath, "--ast");
	const blocks = (
		read.parsed as {
			blocks: Array<{
				id: string;
				runs?: Array<{ type: string; text: string }>;
			}>;
		}
	).blocks;
	const block = blocks.find((candidate) => candidate.id === blockId);
	return (block?.runs ?? [])
		.filter((run) => run.type === "text")
		.map((run) => run.text)
		.join("");
}

describe("docx track-changes accept --at (batch)", () => {
	test("repeated --at accepts each id atomically against the pre-mutation tree", async () => {
		const docPath = await freshMulti("batch-accept");
		const before = await listTracked(docPath);
		// 3 replaces × 2 wrappers each = 6 tracked changes.
		expect(before).toHaveLength(6);
		const ids = before.map((change) => change.id);
		const targets = [ids[0], ids[2], ids[4]].filter(
			(id): id is string => id !== undefined,
		);
		expect(targets).toHaveLength(3);

		// Pick three non-adjacent ids that span all three paragraphs. In the
		// buggy world (one-at-a-time, no batch), accepting tc0 would shift
		// tc2/tc4's ids and the agent's pre-fetched list would be wrong.
		const result = await runCli(
			"track-changes",
			"accept",
			docPath,
			"--at",
			targets[0] as string,
			"--at",
			targets[1] as string,
			"--at",
			targets[2] as string,
			"--verbose",
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as {
			applied: Array<{ id: string }>;
		};
		expect(payload.applied.map((entry) => entry.id)).toEqual(targets);

		// 3 of the original 6 changes accepted → 3 remain.
		const after = await listTracked(docPath);
		expect(after).toHaveLength(3);
	});

	test("dedupes repeated ids", async () => {
		const docPath = await freshMulti("batch-dedupe");
		const before = await listTracked(docPath);
		const firstId = before[0]?.id;
		expect(firstId).toBeDefined();

		const result = await runCli(
			"track-changes",
			"accept",
			docPath,
			"--at",
			firstId as string,
			"--at",
			firstId as string,
			"--verbose",
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as {
			applied: Array<{ id: string }>;
		};
		expect(payload.applied).toHaveLength(1);
	});

	// NOTE: the --at + --all USAGE rejection is asserted once in the
	// accept/reject contract block above; not re-tested here (same code path).

	test("unknown id in the batch errors atomically (no writes)", async () => {
		const docPath = await freshMulti("batch-unknown");
		const beforeText = await paragraphTextDoc(docPath, "p0");

		const result = await runCli(
			"track-changes",
			"accept",
			docPath,
			"--at",
			"tc0",
			"--at",
			"tc99",
		);
		expect(result.exitCode).toBe(3); // NOT_FOUND
		expect(result.parsed).toMatchObject({
			code: "TRACKED_CHANGE_NOT_FOUND",
		});
		// p0 unchanged because the batch aborted before any apply.
		const afterText = await paragraphTextDoc(docPath, "p0");
		expect(afterText).toBe(beforeText);
	});

	test("reject --at also supports the multiple flag", async () => {
		const docPath = await freshMulti("batch-reject");
		const before = await listTracked(docPath);
		const idA = before[0]?.id;
		const idB = before[1]?.id;
		expect(idA).toBeDefined();
		expect(idB).toBeDefined();

		const result = await runCli(
			"track-changes",
			"reject",
			docPath,
			"--at",
			idA as string,
			"--at",
			idB as string,
			"--verbose",
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as { applied: unknown[] };
		expect(payload.applied).toHaveLength(2);
	});
});

// `--track` forces tracked emission for one command even when the document's
// global <w:trackChanges/> toggle is OFF (the Task-2 trap: tracked corrections
// that silently weren't tracked).

async function docFrom(label: string, md: string): Promise<string> {
	const workspace = tempWorkspace(label);
	const src = join(workspace, "src.md");
	await Bun.write(src, md);
	const docPath = join(workspace, "out.docx");
	expect((await runCli("create", docPath, "--from", src)).exitCode).toBe(0);
	return docPath;
}

async function changes(
	docPath: string,
): Promise<Array<{ kind: string; author: string }>> {
	const list = await runCli("track-changes", "list", docPath);
	return list.parsed as Array<{ kind: string; author: string }>;
}

describe("--track forces tracked emission with the global toggle off", () => {
	test("edit --track wraps the edit in del/ins (author defaults to Reviewer)", async () => {
		const docPath = await docFrom("track-edit", "The old plan is here.\n");
		const find = await runCli("find", docPath, "old");
		const locator = (find.parsed as { matches: Array<{ locator: string }> })
			.matches[0]?.locator;
		if (!locator) throw new Error("expected a match for 'old'");
		expect(
			(
				await runCli(
					"edit",
					docPath,
					"--at",
					locator,
					"--text",
					"new",
					"--track",
				)
			).exitCode,
		).toBe(0);
		const list = await changes(docPath);
		expect(list.map((change) => change.kind).sort()).toEqual(["del", "ins"]);
		expect(list.every((change) => change.author === "Reviewer")).toBe(true);
	});

	test("replace --track records substitutions as tracked changes", async () => {
		const docPath = await docFrom(
			"track-replace",
			"Acme Corp and Acme Corp.\n",
		);
		expect(
			(
				await runCli(
					"replace",
					docPath,
					"Acme Corp",
					"Acme Industries",
					"--track",
					"--all",
				)
			).exitCode,
		).toBe(0);
		expect((await changes(docPath)).length).toBeGreaterThan(0);
	});

	test("delete --track wraps the paragraph in a tracked deletion", async () => {
		const docPath = await docFrom(
			"track-delete",
			"First para.\n\nSecond para.\n",
		);
		expect(
			(await runCli("delete", docPath, "--at", "p0", "--track")).exitCode,
		).toBe(0);
		expect((await changes(docPath)).some((c) => c.kind === "del")).toBe(true);
	});

	test("tables delete-row --track emits a rowDel revision", async () => {
		const docPath = await docFrom(
			"track-table",
			"| A | B |\n| - | - |\n| 1 | 2 |\n| 3 | 4 |\n",
		);
		expect(
			(
				await runCli(
					"tables",
					"delete-row",
					docPath,
					"--at",
					"t0:r2",
					"--track",
				)
			).exitCode,
		).toBe(0);
		expect((await changes(docPath)).some((c) => c.kind === "rowDel")).toBe(
			true,
		);
	});

	test("without --track and global off, a mutation leaves no tracked changes", async () => {
		const docPath = await docFrom("track-none", "Acme Corp.\n");
		expect(
			(
				await runCli(
					"replace",
					docPath,
					"Acme Corp",
					"Acme Industries",
					"--all",
				)
			).exitCode,
		).toBe(0);
		expect((await changes(docPath)).length).toBe(0);
	});
});

/**
 * End-to-end coverage for tracked moves (<w:moveFrom>/<w:moveTo>) against
 * tests/fixtures/tracked-moves.docx. Layout:
 *   p0: "Origin paragraph: " + moveFrom("the moved sentence") + "."
 *   p1: "Destination paragraph: " + moveTo("the moved sentence") + "."
 */
const MOVES_FIX = "tests/fixtures/tracked-moves.docx";

type Body = {
	blocks: Array<{
		type: string;
		runs?: Array<{
			type: string;
			text?: string;
			trackedChange?: { id: string; kind: string; author: string };
		}>;
	}>;
};

const freshMoves = (label: string) => copyFixture(label, MOVES_FIX);

function paragraphTextAt(doc: Body, index: number): string {
	const block = doc.blocks[index];
	if (!block || block.type !== "paragraph") return "";
	return (block.runs ?? [])
		.filter((run) => run.type === "text")
		.map((run) => run.text ?? "")
		.join("");
}

describe("tracked moves — AST surface", () => {
	test("read exposes moveFrom and moveTo as TrackedChange entries", async () => {
		const result = await runCli("read", MOVES_FIX, "--ast");
		const doc = result.parsed as Body;

		const movedRunFrom = (doc.blocks[0]?.runs ?? []).find(
			(run) => run.type === "text" && run.text === "the moved sentence",
		);
		const movedRunTo = (doc.blocks[1]?.runs ?? []).find(
			(run) => run.type === "text" && run.text === "the moved sentence",
		);

		expect(movedRunFrom?.trackedChange?.kind).toBe("moveFrom");
		expect(movedRunFrom?.trackedChange?.author).toBe("Reviewer");
		expect(movedRunTo?.trackedChange?.kind).toBe("moveTo");
		expect(movedRunTo?.trackedChange?.author).toBe("Reviewer");
	});

	test("track-changes list reports both halves of the move", async () => {
		const result = await runCli("track-changes", "list", MOVES_FIX);
		const records = result.parsed as Array<{
			id: string;
			kind: string;
			text: string;
		}>;
		expect(records).toHaveLength(2);
		const kinds = records.map((record) => record.kind).sort();
		expect(kinds).toEqual(["moveFrom", "moveTo"]);
		// Both halves carry the same text.
		expect(
			records.every((record) => record.text === "the moved sentence"),
		).toBe(true);
	});
});

describe("tracked moves — wc views", () => {
	test("--accepted skips the moveFrom origin", async () => {
		const result = await runCli("wc", MOVES_FIX, "--accepted");
		// Accepted view: "Origin paragraph: ." + "Destination paragraph: the moved sentence."
		// Words: "Origin", "paragraph:", "." → "Origin paragraph: ." has 3 word-like tokens.
		// Word counter splits on whitespace and counts non-whitespace runs:
		//   p0 accepted: "Origin paragraph: ." = 3 tokens
		//   p1 accepted: "Destination paragraph: the moved sentence." = 5 tokens
		// total = 8
		expect((result.parsed as { words: number }).words).toBe(8);
	});

	test("--baseline skips the moveTo destination", async () => {
		const result = await runCli("wc", MOVES_FIX, "--baseline");
		// Baseline view: "Origin paragraph: the moved sentence." + "Destination paragraph: ."
		//   p0: "Origin", "paragraph:", "the", "moved", "sentence." = 5 tokens
		//   p1: "Destination", "paragraph:", "." = 3 tokens
		// total = 8
		expect((result.parsed as { words: number }).words).toBe(8);
	});

	test("default is the accepted view (skips moveFrom origin)", async () => {
		const result = await runCli("wc", MOVES_FIX);
		// Same as --accepted: 8 words (default flipped from "current"
		// for consistency with `read` / `find` / `replace`).
		expect((result.parsed as { words: number }).words).toBe(8);
	});

	test("--current counts both halves (legacy default)", async () => {
		const result = await runCli("wc", MOVES_FIX, "--current");
		// Both halves visible: 5 + 5 = 10
		expect((result.parsed as { words: number }).words).toBe(10);
	});
});

describe("tracked moves — accept", () => {
	test("accept --all unwraps moveTo (text stays) and deletes moveFrom (text gone)", async () => {
		const docPath = await freshMoves("moves-accept-all");
		const result = await runCli("track-changes", "accept", docPath, "--all");
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as Body;
		expect(paragraphTextAt(doc, 0)).toBe("Origin paragraph: .");
		expect(paragraphTextAt(doc, 1)).toBe(
			"Destination paragraph: the moved sentence.",
		);

		// And the underlying XML has no remaining move wrappers.
		const pkg = await Pkg.open(docPath);
		const xml = await pkg.readText("word/document.xml");
		expect(xml).not.toContain("<w:moveFrom");
		expect(xml).not.toContain("<w:moveTo");
	});

	test("accept --at one moveFrom alone leaves the moveTo intact", async () => {
		const docPath = await freshMoves("moves-accept-one");
		// tc0 is the moveFrom (it appears first in document order); tc1 is the moveTo.
		const result = await runCli(
			"track-changes",
			"accept",
			docPath,
			"--at",
			"tc0",
		);
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as Body;
		// p0: moveFrom accepted → text gone.
		expect(paragraphTextAt(doc, 0)).toBe("Origin paragraph: .");
		// p1: moveTo still wrapped.
		const movedRunTo = (doc.blocks[1]?.runs ?? []).find(
			(run) => run.type === "text" && run.text === "the moved sentence",
		);
		expect(movedRunTo?.trackedChange?.kind).toBe("moveTo");
	});
});

describe("tracked moves — reject", () => {
	test("reject --all unwraps moveFrom (text stays) and deletes moveTo (text gone)", async () => {
		const docPath = await freshMoves("moves-reject-all");
		const result = await runCli("track-changes", "reject", docPath, "--all");
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as Body;
		expect(paragraphTextAt(doc, 0)).toBe(
			"Origin paragraph: the moved sentence.",
		);
		expect(paragraphTextAt(doc, 1)).toBe("Destination paragraph: .");

		const pkg = await Pkg.open(docPath);
		const xml = await pkg.readText("word/document.xml");
		expect(xml).not.toContain("<w:moveFrom");
		expect(xml).not.toContain("<w:moveTo");
	});
});
