import { describe, expect, test } from "bun:test";
import { writeFileSync } from "node:fs";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

const FIXTURE = "tests/fixtures/comments-batch.docx";
// Layout (built by scripts/make-comments-batch-fixture.ts):
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
		expect(result.parsed).toMatchObject({ ok: false, code: "USAGE" });
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
			ok: false,
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
				'{"range": "p0", "text": "first comment"}',
				'{"anchor": "Beta", "text": "second comment", "author": "Reviewer"}',
				'{"range": "p2:0-5", "text": "third comment"}',
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
				'{"range": "p0", "text": "first"}',
				'{"anchor": "nonexistent", "text": "missing"}',
				'{"range": "p1", "text": "third"}',
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
			ok: false,
			code: "MATCH_NOT_FOUND",
		});

		// No comments were added because the batch aborted at validation.
		const after = await listComments(docPath);
		expect(after).toHaveLength(0);
	});

	test("rejects --batch combined with --range/--anchor/--text on the CLI", async () => {
		const { workspace, docPath } = await freshCopy("batch-mutex");
		const batchPath = join(workspace, "x.jsonl");
		writeFileSync(batchPath, '{"range": "p0", "text": "hi"}\n');
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--batch",
			batchPath,
			"--range",
			"p1",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ ok: false, code: "USAGE" });
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
		expect(result.parsed).toMatchObject({ ok: false, code: "USAGE" });
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
				"--range",
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
			"--id",
			idA as string,
			"--id",
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
			"--id",
			idA as string,
			"--id",
			"c999",
		);
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({
			ok: false,
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
				"--range",
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
			"--id",
			idA as string,
			"--id",
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
			"--id",
			idA as string,
			"--id",
			idB as string,
		);
		// Now unset both.
		const result = await runCli(
			"comments",
			"resolve",
			docPath,
			"--id",
			idA as string,
			"--id",
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
