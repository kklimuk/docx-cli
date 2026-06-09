import { beforeEach, describe, expect, test } from "bun:test";
import { writeFileSync } from "node:fs";
import { join } from "node:path";
import { runCli, spawnCli, tempWorkspace } from "./harness";

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
			"--at",
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
			"--at",
			"p0:16-19",
			"--text",
			"fox?",
			"--author",
			"Jane",
		);
		const read = await runCli("read", docPath, "--ast");
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
			"--at",
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
			"--at",
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
			"--at",
			"p0",
			"--text",
			"x",
			"--author",
			"A",
		);
		await runCli("comments", "resolve", docPath, "--at", "c0");

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

	test("delete removes the comment from the listing", async () => {
		await runCli(
			"comments",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"hi",
			"--author",
			"A",
		);
		await runCli("comments", "delete", docPath, "--at", "c0");
		const afterDelete = await runCli(
			"comments",
			"list",
			docPath,
			"--include-resolved",
		);
		expect((afterDelete.parsed as unknown[]).length).toBe(0);
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
			"--at",
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

	test("add rejects cross-block range with out-of-range offset", async () => {
		await runCli("insert", docPath, "--after", "p0", "--text", "Short.");
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--at",
			"p0:0-p1:999",
			"--text",
			"too far",
			"--author",
			"A",
		);
		expect(result.exitCode).not.toBe(0);
		expect(result.parsed).toMatchObject({
			code: "INVALID_LOCATOR",
		});
	});

	test("add anchors a span that includes hyperlinked text", async () => {
		// Wrap "brown" (offsets 10..15) in a hyperlink, then anchor a comment
		// over "quick brown fox" (4..19). Pre-fix, paragraphTextLength would
		// have ignored hyperlinked runs and reported the paragraph as 39 chars
		// instead of 44, throwing SpanOutOfRangeError or mis-anchoring the end
		// marker. Post-fix, hyperlinked text counts toward offsets so the span
		// reads back exactly as anchored.
		await runCli(
			"hyperlinks",
			"add",
			docPath,
			"--at",
			"p0:10-15",
			"--url",
			"https://example.com/brown",
		);
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--at",
			"p0:4-19",
			"--text",
			"Spans the hyperlink",
			"--author",
			"Reviewer",
		);
		expect(result.exitCode).toBe(0);

		const list = await runCli("comments", "list", docPath);
		const comments = list.parsed as Array<{
			text: string;
			anchor: { startOffset: number; endOffset: number };
		}>;
		const spanning = comments.find(
			(comment) => comment.text === "Spans the hyperlink",
		);
		expect(spanning?.anchor.startOffset).toBe(4);
		expect(spanning?.anchor.endOffset).toBe(19);
	});

	test("resolve auto-injects paraId on legacy comments", async () => {
		// comments-simple.docx is a mammoth-authored fixture without paraIds
		const workspace = tempWorkspace("legacy");
		const docCopy = join(workspace, "legacy.docx");
		await Bun.write(docCopy, Bun.file("tests/fixtures/comments-simple.docx"));
		const result = await runCli("comments", "resolve", docCopy, "--at", "c0");
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

const FIXTURE = "tests/fixtures/comments-batch.docx";
// Layout (built by tests/fixtures/setup/comments-batch.ts):
//   p0: "Alpha is first."
//   p1: "Beta is second."
//   p2: "Gamma is third."
// Each paragraph has a unique opening word (Alpha/Beta/Gamma) for unique
// anchors, and shares "is" so multi-match / --occurrence is exercised.

async function freshCopy(label: string): Promise<{
	workspace: string;
	docPath: string;
}> {
	const workspace = tempWorkspace(label);
	const docPath = join(workspace, "out.docx");
	await Bun.write(docPath, Bun.file(FIXTURE));
	return { workspace, docPath };
}

async function listComments(docPath: string): Promise<unknown[]> {
	const result = await runCli("comments", "list", docPath);
	return result.parsed as unknown[];
}

describe("docx comments add --anchor", () => {
	test("anchors a comment via a unique phrase", async () => {
		const { docPath } = await freshCopy("anchor-unique");
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--anchor",
			"Beta",
			"--text",
			"check this",
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as { commentId: string; locator: string };
		expect(payload.commentId).toMatch(/^c\d+$/);
		expect(payload.locator).toBe("p1:0-4");
	});

	test("errors when the anchor matches multiple times without --occurrence", async () => {
		const { docPath } = await freshCopy("anchor-ambiguous");
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--anchor",
			"is",
			"--text",
			"check",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
		expect((result.parsed as { error: string }).error).toContain(
			"matches 3 times",
		);
	});

	test("--occurrence picks the Nth match", async () => {
		const { docPath } = await freshCopy("anchor-nth");
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--anchor",
			"is",
			"--occurrence",
			"2",
			"--text",
			"second match",
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as { locator: string };
		// "is" appears at offset 5 in p1 ("Beta is second."); second match
		// in document order.
		expect(payload.locator).toBe("p1:5-7");
	});

	test("anchor not found errors with MATCH_NOT_FOUND", async () => {
		const { docPath } = await freshCopy("anchor-missing");
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--anchor",
			"nonexistent phrase",
			"--text",
			"x",
		);
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({
			code: "MATCH_NOT_FOUND",
		});
	});
});

describe("docx comments add --batch", () => {
	test("applies a JSONL batch atomically and prints the minted ids", async () => {
		const { workspace, docPath } = await freshCopy("batch-add");
		const batchPath = join(workspace, "comments.jsonl");
		writeFileSync(
			batchPath,
			[
				'{"at": "p0", "text": "first comment"}',
				'{"anchor": "Beta", "text": "second comment", "author": "Reviewer"}',
				'{"at": "p2:0-5", "text": "third comment"}',
			].join("\n"),
		);

		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--batch",
			batchPath,
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as {
			batch: Array<{ commentId: string; locator: string }>;
		};
		expect(payload.batch).toHaveLength(3);
		expect(payload.batch[0]?.locator).toBe("p0");
		expect(payload.batch[1]?.locator).toBe("p1:0-4");
		expect(payload.batch[2]?.locator).toBe("p2:0-5");
		expect(new Set(payload.batch.map((entry) => entry.commentId)).size).toBe(3);

		const comments = await listComments(docPath);
		expect(comments).toHaveLength(3);
	});

	test("aborts atomically when one entry's anchor is missing", async () => {
		const { workspace, docPath } = await freshCopy("batch-abort");
		const batchPath = join(workspace, "bad.jsonl");
		writeFileSync(
			batchPath,
			[
				'{"at": "p0", "text": "first"}',
				'{"anchor": "nonexistent", "text": "missing"}',
				'{"at": "p1", "text": "third"}',
			].join("\n"),
		);

		const before = await listComments(docPath);
		expect(before).toHaveLength(0);

		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--batch",
			batchPath,
		);
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({
			code: "MATCH_NOT_FOUND",
		});

		// No comments were added because the batch aborted at validation.
		const after = await listComments(docPath);
		expect(after).toHaveLength(0);
	});

	test("rejects --batch combined with --at/--anchor/--text on the CLI", async () => {
		const { workspace, docPath } = await freshCopy("batch-mutex");
		const batchPath = join(workspace, "x.jsonl");
		writeFileSync(batchPath, '{"at": "p0", "text": "hi"}\n');
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--batch",
			batchPath,
			"--at",
			"p1",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});
});

describe("docx comments add — view-aware offset resolution (regression)", () => {
	// Bug: `find` returns offsets in the accepted view by default
	// (skipping <w:del>/<w:moveFrom>), but `addCommentMarkersToParagraph`
	// originally counted ALL run text via `sumRunBearingTextLength`. So
	// on a paragraph with a tracked deletion before the anchor, the
	// computed offset landed inside the deletion or at the wrong run.
	// Fix: thread a FindView through marker placement; default to
	// accepted-view (matching `find`).

	test("anchor in a paragraph with a prior tracked deletion wraps only the visible text", async () => {
		const workspace = tempWorkspace("anchor-after-del");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Hello old world today.");
		await runCli("track-changes", docPath, "on");
		// Replace "old " with "" — emits <w:del>old </w:del>. Accepted-
		// view text becomes "Hello world today." (18 chars).
		await runCli("replace", docPath, "old ", "");

		// Anchor on "world" — should resolve to accepted-view offset 6
		// (after "Hello ") and land the comment markers around "world"
		// without including the deleted "old ". Pre-fix, marker
		// placement counted <w:del> bytes too, so offset 6 landed
		// inside "old".
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--anchor",
			"world",
			"--text",
			"check",
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as { commentId: string; locator: string };
		expect(payload.locator).toBe("p0:6-11");

		// Render with comments inlined to verify the bracket lands
		// around "world" only (accepted view drops the deletion).
		const rendered = await runCli(
			"read",
			docPath,
			"--comments",
			"--from",
			"p0",
			"--to",
			"p0",
		);
		expect(rendered.stdout).toContain(`world[^${payload.commentId}]`);
		expect(rendered.stdout).toContain(`[^${payload.commentId}]: "world"`);
	});

	test("--current view places markers in the raw byte stream", async () => {
		const workspace = tempWorkspace("anchor-current");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Hello old world today.");
		await runCli("track-changes", docPath, "on");
		await runCli("replace", docPath, "old ", "");

		// In --current view, the <w:del> content "old " is part of the
		// haystack. Anchor on "old" should work there.
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--anchor",
			"old",
			"--text",
			"trace this deletion",
			"--current",
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as { commentId: string };

		const rendered = await runCli(
			"read",
			docPath,
			"--current",
			"--comments",
			"--from",
			"p0",
			"--to",
			"p0",
		);
		expect(rendered.stdout).toContain(`[^${payload.commentId}]: "old"`);
	});

	test("--current and --baseline are mutually exclusive", async () => {
		const { docPath } = await freshCopy("view-mutex");
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--anchor",
			"Beta",
			"--text",
			"x",
			"--current",
			"--baseline",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});
});

describe("docx comments delete --id (repeatable) and --batch", () => {
	async function withSeededComments(label: string): Promise<{
		workspace: string;
		docPath: string;
		ids: string[];
	}> {
		const { workspace, docPath } = await freshCopy(label);
		const ids: string[] = [];
		for (const para of ["p0", "p1", "p2"]) {
			const res = await runCli(
				"comments",
				"add",
				docPath,
				"--at",
				para,
				"--text",
				`note on ${para}`,
			);
			ids.push((res.parsed as { commentId: string }).commentId);
		}
		return { workspace, docPath, ids };
	}

	test("repeated --id removes multiple comments atomically", async () => {
		const { docPath, ids } = await withSeededComments("delete-multi");
		expect(ids).toHaveLength(3);
		const idA = ids[0];
		const idC = ids[2];
		expect(idA).toBeDefined();
		expect(idC).toBeDefined();

		const result = await runCli(
			"comments",
			"delete",
			docPath,
			"--at",
			idA as string,
			"--at",
			idC as string,
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as {
			batch: Array<{ commentId: string }>;
		};
		expect(payload.batch.map((entry) => entry.commentId)).toEqual([
			idA as string,
			idC as string,
		]);

		const remaining = await listComments(docPath);
		expect(remaining).toHaveLength(1);
	});

	test("unknown id aborts the batch (no writes)", async () => {
		const { docPath, ids } = await withSeededComments("delete-bad");
		const idA = ids[0];
		expect(idA).toBeDefined();
		const result = await runCli(
			"comments",
			"delete",
			docPath,
			"--at",
			idA as string,
			"--at",
			"c999",
		);
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({
			code: "COMMENT_NOT_FOUND",
		});
		const remaining = await listComments(docPath);
		expect(remaining).toHaveLength(3);
	});

	test("--batch JSONL also works", async () => {
		const { workspace, docPath, ids } =
			await withSeededComments("delete-batch");
		const batchPath = join(workspace, "rm.jsonl");
		writeFileSync(
			batchPath,
			[`{"id": "${ids[0]}"}`, `{"id": "${ids[1]}"}`].join("\n"),
		);
		const result = await runCli(
			"comments",
			"delete",
			docPath,
			"--batch",
			batchPath,
		);
		expect(result.exitCode).toBe(0);
		const remaining = await listComments(docPath);
		expect(remaining).toHaveLength(1);
	});
});

describe("docx comments resolve --id (repeatable) and --batch", () => {
	async function withSeededComments(label: string): Promise<{
		docPath: string;
		ids: string[];
	}> {
		const { docPath } = await freshCopy(label);
		const ids: string[] = [];
		for (const para of ["p0", "p1", "p2"]) {
			const res = await runCli(
				"comments",
				"add",
				docPath,
				"--at",
				para,
				"--text",
				`note ${para}`,
			);
			ids.push((res.parsed as { commentId: string }).commentId);
		}
		return { docPath, ids };
	}

	test("repeated --id resolves multiple comments at once", async () => {
		const { docPath, ids } = await withSeededComments("resolve-multi");
		const idA = ids[0];
		const idC = ids[2];
		expect(idA).toBeDefined();
		expect(idC).toBeDefined();

		const result = await runCli(
			"comments",
			"resolve",
			docPath,
			"--at",
			idA as string,
			"--at",
			idC as string,
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as {
			batch: Array<{ commentId: string; resolved: boolean }>;
		};
		expect(payload.batch).toEqual([
			{ commentId: idA as string, resolved: true },
			{ commentId: idC as string, resolved: true },
		]);

		// Read the comments-ext part to confirm the resolved flags landed.
		const proc = Bun.spawn(
			["unzip", "-p", docPath, "word/commentsExtended.xml"],
			{ stdout: "pipe" },
		);
		const xml = await new Response(proc.stdout).text();
		await proc.exited;
		const doneCount = (xml.match(/w15:done="1"/g) ?? []).length;
		expect(doneCount).toBe(2);
	});

	test("--unset applies to all ids in the batch", async () => {
		const { docPath, ids } = await withSeededComments("resolve-unset");
		const idA = ids[0];
		const idB = ids[1];
		expect(idA).toBeDefined();
		expect(idB).toBeDefined();

		// First mark them resolved.
		await runCli(
			"comments",
			"resolve",
			docPath,
			"--at",
			idA as string,
			"--at",
			idB as string,
		);
		// Now unset both.
		const result = await runCli(
			"comments",
			"resolve",
			docPath,
			"--at",
			idA as string,
			"--at",
			idB as string,
			"--unset",
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as {
			batch: Array<{ resolved: boolean }>;
		};
		expect(payload.batch.every((entry) => entry.resolved === false)).toBe(true);
	});
});

// -o is handled independently in each comments subcommand; verify the parallel
// write lands in the output and the SOURCE is left byte-for-byte untouched (a
// silent source overwrite would lose an agent's original).
describe("docx comments — -o parallel write", () => {
	test("comments add -o writes to the output and leaves the source byte-unchanged", async () => {
		const src = join(tempWorkspace("comments-o-src"), "doc.docx");
		await runCli("create", src, "--text", "Alpha beta gamma.");
		const before = await Bun.file(src).bytes();
		const out = join(tempWorkspace("comments-o-out"), "out.docx");

		const result = await runCli(
			"comments",
			"add",
			src,
			"--at",
			"p0",
			"--text",
			"note",
			"-o",
			out,
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { path: string }).path).toBe(out);
		expect(await Bun.file(src).bytes()).toEqual(before);

		expect(
			((await runCli("comments", "list", out)).parsed as unknown[]).length,
		).toBe(1);
		expect(
			((await runCli("comments", "list", src)).parsed as unknown[]).length,
		).toBe(0);
	});
});

// Only the batch path's dry-run was covered; the single-shot --dry-run for
// add/resolve/delete short-circuits before any mutation while still validating.
describe("docx comments — --dry-run previews without writing", () => {
	async function seedDoc(label: string): Promise<string> {
		const docPath = join(tempWorkspace(label), "doc.docx");
		await runCli("create", docPath, "--text", "alpha beta gamma");
		await runCli("comments", "add", docPath, "--at", "p0", "--text", "first");
		return docPath;
	}

	test("add --dry-run does not write the comment", async () => {
		const docPath = join(tempWorkspace("comments-dry-add"), "doc.docx");
		await runCli("create", docPath, "--text", "alpha beta");
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"note",
			"--dry-run",
		);
		expect((result.parsed as { dryRun: boolean }).dryRun).toBe(true);
		expect(await listComments(docPath)).toEqual([]);
	});

	test("resolve --dry-run leaves the comment unresolved", async () => {
		const docPath = await seedDoc("comments-dry-resolve");
		await runCli("comments", "resolve", docPath, "--at", "c0", "--dry-run");
		// Default list excludes resolved comments — it's still here, so unresolved.
		expect(await listComments(docPath)).toHaveLength(1);
	});

	test("delete --dry-run keeps the comment", async () => {
		const docPath = await seedDoc("comments-dry-delete");
		await runCli("comments", "delete", docPath, "--at", "c0", "--dry-run");
		expect(await listComments(docPath)).toHaveLength(1);
	});
});

describe("docx comments reply — output contract", () => {
	async function seedDoc(label: string): Promise<string> {
		const docPath = join(tempWorkspace(label), "doc.docx");
		await runCli("create", docPath, "--text", "alpha beta gamma");
		await runCli("comments", "add", docPath, "--at", "p0", "--text", "first");
		return docPath;
	}

	test("prints a bare cN by default (no --verbose)", async () => {
		const docPath = await seedDoc("reply-bare");
		// spawnCli bypasses the harness's --verbose injection, so we see the real
		// default (minted-locator) output.
		const result = await spawnCli(
			"comments",
			"reply",
			docPath,
			"--at",
			"c0",
			"--text",
			"follow-up",
		);
		expect(result.exitCode).toBe(0);
		expect(result.stdout.trim()).toMatch(/^c\d+$/);
	});

	test("--dry-run does not add the reply", async () => {
		const docPath = await seedDoc("reply-dry");
		const before = (await listComments(docPath)).length;
		await runCli(
			"comments",
			"reply",
			docPath,
			"--at",
			"c0",
			"--text",
			"skip",
			"--dry-run",
		);
		expect(await listComments(docPath)).toHaveLength(before);
	});

	test("-o writes the reply to the output and leaves the source byte-unchanged", async () => {
		const src = await seedDoc("reply-o-src");
		const before = await Bun.file(src).bytes();
		const out = join(tempWorkspace("reply-o-out"), "out.docx");

		const result = await runCli(
			"comments",
			"reply",
			src,
			"--at",
			"c0",
			"--text",
			"follow-up",
			"-o",
			out,
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { path: string }).path).toBe(out);
		expect(await Bun.file(src).bytes()).toEqual(before);
	});
});
