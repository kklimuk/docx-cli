import { beforeEach, describe, expect, test } from "bun:test";
import { writeFileSync } from "node:fs";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import { runCli, spawnCli, tempWorkspace } from "./harness";
import { expectValidClean, readDocumentXml } from "./helpers";

async function readPart(docPath: string, partName: string): Promise<string> {
	const pkg = await Pkg.open(docPath);
	return await pkg.readText(partName);
}

describe("docx comments — schema hygiene of the comments part", () => {
	test("add on a comments part lacking mc:Ignorable registers w14 so validate stays clean", async () => {
		// Word-authored comments.xml declares xmlns:w14 but often omits
		// `mc:Ignorable`; our `w14:paraId`/`w14:textId` are then schema errors the
		// validator can't strip — every `comments add` made `validate` fail on a file
		// the agent hadn't broken. Simulate: strip the declaration our fresh root
		// carries, then add again.
		const workspace = tempWorkspace("comments-ignorable");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Clause one. Clause two.");
		await runCli("comments", "add", docPath, "--at", "p0", "--text", "first");
		const pkg = await Pkg.open(docPath);
		const stripped = (await pkg.readText("word/comments.xml")).replace(
			/ mc:Ignorable="[^"]*"/,
			"",
		);
		expect(stripped).not.toContain("mc:Ignorable");
		pkg.writeText("word/comments.xml", stripped);
		await pkg.save();
		// Sanity: the stripped part now fails validation on its own.
		expect((await runCli("validate", docPath)).exitCode).not.toBe(0);

		const added = await runCli(
			"comments",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"second",
		);
		expect(added.exitCode).toBe(0);
		const validation = await runCli("validate", docPath);
		expect(validation.exitCode).toBe(0);
		const root = (await readPart(docPath, "word/comments.xml")).match(
			/<w:comments[^>]*>/,
		)?.[0];
		expect(root).toContain("w14");
		expect(root).toMatch(/mc:Ignorable="[^"]*w14/);
	});
});

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

	test("reply anchors its own body markers in Word's thread nesting", async () => {
		// Word drops any comment whose <w:commentReference> is missing from the
		// body — a reply with only a comments.xml entry + sidecar link is
		// silently deleted on the next open/save (issue #1).
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

		// Word's thread layout: Start0, Start1, …, End0, ref0, End1, ref1 —
		// the reply's range closes after the parent's reference run.
		const documentXml = await readDocumentXml(docPath);
		const order = [
			'<w:commentRangeStart w:id="0"/>',
			'<w:commentRangeStart w:id="1"/>',
			'<w:commentRangeEnd w:id="0"/>',
			'<w:commentReference w:id="0"/>',
			'<w:commentRangeEnd w:id="1"/>',
			'<w:commentReference w:id="1"/>',
		].map((marker) => documentXml.indexOf(marker));
		expect(order).not.toContain(-1);
		expect(order).toEqual([...order].sort((left, right) => left - right));

		// The thread link lives ONLY in the extended sidecar — Word writes no
		// parent attribute in comments.xml itself.
		const extendedXml = await readPart(docPath, "word/commentsExtended.xml");
		expect(extendedXml).toContain("w15:paraIdParent=");
		const commentsXml = await readPart(docPath, "word/comments.xml");
		expect(commentsXml).not.toContain("paraIdParent");

		// The anchored reply survives the write-read loop: it now has a real
		// anchor and shows up in the default markdown view.
		const listed = await runCli("comments", "list", docPath);
		const reply = (
			listed.parsed as Array<{ id: string; anchor: { startBlockId: string } }>
		).find((comment) => comment.id === "c1");
		expect(reply?.anchor.startBlockId).toBe("p0");
		const rendered = await runCli("read", docPath, "--comments");
		expect(rendered.stdout).toContain("[^c1]");
	});

	test("sibling replies keep reply order; replying to a reply attaches to the thread root", async () => {
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
		for (const reply of ["first answer", "second answer"]) {
			await runCli(
				"comments",
				"reply",
				docPath,
				"--at",
				"c0",
				"--text",
				reply,
				"--author",
				"B",
			);
		}
		const nested = await runCli(
			"comments",
			"reply",
			docPath,
			"--at",
			"c1",
			"--text",
			"reply to the first answer",
			"--author",
			"C",
		);
		// Word threads are single-level: replying to c1 attaches to the thread
		// root c0 (Word rewrites deeper chains to the root on open anyway), and
		// the ack reports the effective parent.
		expect(nested.parsed).toMatchObject({ parentId: "c0" });
		const list = await runCli("comments", "list", docPath);
		const replies = list.parsed as Array<{ id: string; parentId?: string }>;
		expect(replies.find((comment) => comment.id === "c3")?.parentId).toBe("c0");

		// Thread order holds across document.xml and comments.xml alike.
		const documentXml = await readDocumentXml(docPath);
		const referenceOrder = [
			...documentXml.matchAll(/<w:commentReference w:id="(\d+)"\/>/g),
		].map(([, id]) => id);
		expect(referenceOrder).toEqual(["0", "1", "2", "3"]);
		const commentsXml = await readPart(docPath, "word/comments.xml");
		const commentOrder = [
			...commentsXml.matchAll(/<w:comment w:id="(\d+)"/g),
		].map(([, id]) => id);
		expect(commentOrder).toEqual(["0", "1", "2", "3"]);
	});

	test("the committed comments-threaded fixture reads back as one root-anchored thread", async () => {
		// tests/fixtures/setup/comments-threaded.ts authors this thread with the
		// CLI (Alice root; Bob, Carol, Dave replies — Dave addressed a reply and
		// attached to the root). Guards that the committed binary keeps the
		// Word-valid shape: every reply anchored, all parented to c0.
		const fixturePath = "tests/fixtures/comments-threaded.docx";
		const list = await runCli(
			"comments",
			"list",
			fixturePath,
			"--thread",
			"c0",
		);
		const thread = list.parsed as Array<{
			id: string;
			parentId?: string;
			anchor: { startBlockId: string };
		}>;
		expect(thread.map((comment) => comment.id)).toEqual([
			"c0",
			"c1",
			"c2",
			"c3",
		]);
		for (const comment of thread) {
			expect(comment.anchor.startBlockId).toBe("p0");
			if (comment.id !== "c0") expect(comment.parentId).toBe("c0");
		}
	});

	test("deleting a thread parent cascades to its replies", async () => {
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
		await runCli(
			"comments",
			"reply",
			docPath,
			"--at",
			"c1",
			"--text",
			"Nested",
			"--author",
			"C",
		);

		await runCli("comments", "delete", docPath, "--at", "c0");

		const list = await runCli("comments", "list", docPath);
		expect((list.parsed as unknown[]).length).toBe(0);

		// No orphaned reply markers or dangling thread links survive the cascade.
		const documentXml = await readDocumentXml(docPath);
		expect(documentXml).not.toContain("commentReference");
		expect(documentXml).not.toContain("commentRangeStart");
		const extendedXml = await readPart(docPath, "word/commentsExtended.xml");
		expect(extendedXml).not.toContain("w15:paraIdParent=");
	});

	test("resolve stays visible in the default list; --open filters it", async () => {
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

		// The resolve must survive the write→read loop: the comment stays
		// listed, flagged resolved — hiding it made resolve look like delete.
		const defaultList = await runCli("comments", "list", docPath);
		const all = defaultList.parsed as Array<{ resolved?: boolean }>;
		expect(all.length).toBe(1);
		expect(all[0]?.resolved).toBe(true);

		const openList = await runCli("comments", "list", docPath, "--open");
		expect((openList.parsed as unknown[]).length).toBe(0);
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
		const afterDelete = await runCli("comments", "list", docPath);
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
		const list = await runCli("comments", "list", docCopy);
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

	test("initials come from the name, not a trailing affiliation", async () => {
		const { docPath } = await freshCopy("anchor-initials");
		for (const author of ["Jane Doe (Acme Legal)", "Jane Doe, Acme Legal"]) {
			const added = await runCli(
				"comments",
				"add",
				docPath,
				"--anchor",
				"Beta",
				"--text",
				"check this",
				"--author",
				author,
			);
			expect(added.exitCode).toBe(0);
		}
		const listed = await runCli("comments", "list", docPath);
		const comments = listed.parsed as { initials: string }[];
		expect(comments.map((comment) => comment.initials)).toEqual(["JD", "JD"]);
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

describe("docx comments — inline escape decoding (--text)", () => {
	// `"a\\nb"` in this TS source IS the literal backslash-n the shell delivers.
	let docPath: string;

	beforeEach(async () => {
		docPath = join(tempWorkspace("comments-esc"), "out.docx");
		await runCli("create", docPath, "--text", "A paragraph to annotate.");
	});

	test("add --text decodes \\n / \\t into <w:br/> / <w:tab/>", async () => {
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"first\\nsecond\\tthird",
		);
		expect(result.exitCode).toBe(0);
		const xml = await readPart(docPath, "word/comments.xml");
		expect(xml).toContain("<w:br/>");
		expect(xml).toContain("<w:tab/>");
	});

	test("reply --text decodes \\n too", async () => {
		await runCli("comments", "add", docPath, "--at", "p0", "--text", "Parent.");
		const result = await runCli(
			"comments",
			"reply",
			docPath,
			"--at",
			"c0",
			"--text",
			"one\\ntwo",
		);
		expect(result.exitCode).toBe(0);
		const xml = await readPart(docPath, "word/comments.xml");
		expect(xml).toContain("<w:br/>");
	});

	test("--batch bodies are NOT re-decoded (JSON already decoded them)", async () => {
		// A JSONL entry with a literal backslash-n stays literal — only inline
		// argv is decoded. A real newline in JSON (\n) is handled by JSON.parse.
		const batch = join(tempWorkspace("comments-esc-batch"), "b.jsonl");
		// The \\n here is a JSON escape for a real newline; the \\\\t is a literal
		// backslash-t that must survive verbatim (batch is the trusted channel).
		writeFileSync(
			batch,
			`{"at":"p0","text":"real\\nbreak keep\\\\tliteral"}\n`,
		);
		const result = await runCli("comments", "add", docPath, "--batch", batch);
		expect(result.exitCode).toBe(0);
		const xml = await readPart(docPath, "word/comments.xml");
		expect(xml).toContain("<w:br/>"); // the real JSON newline became a break
		expect(xml).toContain("keep\\tliteral"); // the literal backslash-t survived
	});

	test("comment body round-trips: \\n reads back faithfully in the AST (no word-joining)", async () => {
		await runCli("comments", "add", docPath, "--at", "p0", "--text", "c1\\nc2");
		const list = await runCli("comments", "list", docPath);
		const comment = (list.parsed as Array<{ text: string }>)[0];
		expect(comment?.text).toBe("c1\nc2");
	});
});

describe("docx comments reply --batch", () => {
	async function seedTwoComments(label: string): Promise<string> {
		const docPath = join(tempWorkspace(label), "doc.docx");
		await runCli("create", docPath, "--text", "Alpha clause. Beta clause.");
		await runCli(
			"comments",
			"add",
			docPath,
			"--anchor",
			"Alpha",
			"--text",
			"Narrow this",
		);
		await runCli(
			"comments",
			"add",
			docPath,
			"--anchor",
			"Beta",
			"--text",
			"Why?",
		);
		return docPath;
	}

	test("answers several threads in one call, resolving the ones marked resolve", async () => {
		const docPath = await seedTwoComments("reply-batch");
		const batch = join(tempWorkspace("reply-batch-jsonl"), "replies.jsonl");
		writeFileSync(
			batch,
			[
				JSON.stringify({ at: "c0", text: "Narrowed.", resolve: true }),
				JSON.stringify({ at: "c1", text: "Keeping it." }),
			].join("\n"),
		);
		const result = await spawnCli(
			"comments",
			"reply",
			docPath,
			"--batch",
			batch,
		);
		expect(result.exitCode).toBe(0);
		expect(result.stdout.trim().split("\n")).toEqual(["c2", "c3"]);

		const comments = (await listComments(docPath)) as Array<{
			id: string;
			parentId: string | null;
			resolved: boolean;
			text: string;
		}>;
		const byId = new Map(comments.map((comment) => [comment.id, comment]));
		expect(byId.get("c2")).toMatchObject({ parentId: "c0", text: "Narrowed." });
		expect(byId.get("c3")).toMatchObject({
			parentId: "c1",
			text: "Keeping it.",
		});
		expect(byId.get("c0")?.resolved).toBe(true);
		expect(byId.get("c1")?.resolved).toBe(false);
	});

	test("an unknown parent aborts the whole batch before anything is written", async () => {
		const docPath = await seedTwoComments("reply-batch-bad");
		const before = await Bun.file(docPath).bytes();
		const batch = join(tempWorkspace("reply-batch-bad-jsonl"), "replies.jsonl");
		writeFileSync(
			batch,
			[
				JSON.stringify({ at: "c0", text: "ok" }),
				JSON.stringify({ at: "c9", text: "nope" }),
			].join("\n"),
		);
		const result = await runCli("comments", "reply", docPath, "--batch", batch);
		expect(result.exitCode).toBe(3);
		expect(result.stdout).toContain("entry 1");
		expect(await Bun.file(docPath).bytes()).toEqual(before);
	});

	test('"resolved" (the spelling comments list prints) closes the thread too', async () => {
		const docPath = await seedTwoComments("reply-batch-resolved");
		const batch = join(
			tempWorkspace("reply-batch-resolved-jsonl"),
			"replies.jsonl",
		);
		writeFileSync(
			batch,
			JSON.stringify({ at: "c1", text: "Done.", resolved: true }),
		);
		const result = await runCli("comments", "reply", docPath, "--batch", batch);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({ resolved: ["c1"] });
		const comments = (await listComments(docPath)) as Array<{
			id: string;
			resolved: boolean;
		}>;
		expect(comments.find((comment) => comment.id === "c1")?.resolved).toBe(
			true,
		);
	});

	test("an unknown field or a non-boolean resolve is refused, not ignored", async () => {
		const docPath = await seedTwoComments("reply-batch-unknown");
		const before = await Bun.file(docPath).bytes();
		const workspace = tempWorkspace("reply-batch-unknown-jsonl");
		for (const entry of [
			{ at: "c0", text: "ok", close: true },
			{ at: "c0", text: "ok", resolve: "true" },
		]) {
			const batch = join(workspace, "replies.jsonl");
			writeFileSync(batch, JSON.stringify(entry));
			const result = await runCli(
				"comments",
				"reply",
				docPath,
				"--batch",
				batch,
			);
			expect(result.exitCode).toBe(2);
			expect(result.stdout).toContain("entry 0");
		}
		expect(await Bun.file(docPath).bytes()).toEqual(before);
	});

	test("--dry-run previews the reply ids the real run mints", async () => {
		const docPath = await seedTwoComments("reply-batch-dry");
		const batch = join(tempWorkspace("reply-batch-dry-jsonl"), "replies.jsonl");
		writeFileSync(
			batch,
			[
				JSON.stringify({ at: "c0", text: "a" }),
				JSON.stringify({ at: "c1", text: "b", resolve: true }),
			].join("\n"),
		);
		const result = await runCli(
			"comments",
			"reply",
			docPath,
			"--batch",
			batch,
			"--dry-run",
		);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			batch: [
				{ commentId: "c2", parentId: "c0" },
				{ commentId: "c3", parentId: "c1" },
			],
			resolved: ["c1"],
		});
	});

	test("--batch refuses a CLI --at/--text alongside it", async () => {
		const docPath = await seedTwoComments("reply-batch-mixed");
		const result = await runCli(
			"comments",
			"reply",
			docPath,
			"--batch",
			"-",
			"--at",
			"c0",
		);
		expect(result.exitCode).toBe(2);
	});
});

describe("docx comments — Word's durable-id sidecars (issue #33)", () => {
	// Word 365 writes three sidecars beside comments.xml, each with ONE entry per
	// comment: commentsExtended (`w15:commentEx`, keyed by the comment's
	// `w14:paraId`), commentsIds (`w16cid:commentId`, paraId → durableId) and
	// commentsExtensible (`w16cex:commentExtensible`, keyed by durableId). A
	// comment missing from them makes Word offer to repair the file — and the
	// repair renumbers every comment.
	const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";

	async function wordShapedDocument(): Promise<string> {
		const workspace = tempWorkspace("comments-sidecars");
		const docPath = join(workspace, "word-shaped.docx");
		await runCli("create", docPath, "--text", "First clause. Second clause.");
		await runCli("comments", "add", docPath, "--at", "p0", "--text", "Seed");
		const pkg = await Pkg.open(docPath);
		const paraId = (await pkg.readText("word/comments.xml")).match(
			/w14:paraId="([0-9A-F]{8})"/,
		)?.[1];
		expect(paraId).toBeDefined();
		pkg.writeText(
			"word/commentsExtended.xml",
			`<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:mc="${NS_MC}" mc:Ignorable="w15"><w15:commentEx w15:paraId="${paraId}" w15:done="0"/></w15:commentsEx>`,
		);
		pkg.writeText(
			"word/commentsIds.xml",
			`<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:mc="${NS_MC}" mc:Ignorable="w16cid"><w16cid:commentId w16cid:paraId="${paraId}" w16cid:durableId="1A2B3C4D"/></w16cid:commentsIds>`,
		);
		pkg.writeText(
			"word/commentsExtensible.xml",
			`<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:mc="${NS_MC}" mc:Ignorable="w16cex"><w16cex:commentExtensible w16cex:durableId="1A2B3C4D" w16cex:dateUtc="2024-01-01T00:00:00Z"/></w16cex:commentsExtensible>`,
		);
		await pkg.save();
		return docPath;
	}

	async function sidecarState(docPath: string) {
		const all = (xml: string, pattern: RegExp) =>
			[...xml.matchAll(pattern)].map((match) => match[1]);
		const comments = await readPart(docPath, "word/comments.xml");
		const ids = await readPart(docPath, "word/commentsIds.xml");
		return {
			commentParaIds: all(
				comments,
				/<w:comment [^>]*>.*?w14:paraId="([0-9A-F]{8})"/gs,
			),
			extendedParaIds: all(
				await readPart(docPath, "word/commentsExtended.xml"),
				/<w15:commentEx [^>]*w15:paraId="([0-9A-F]{8})"/g,
			),
			idsParaIds: all(ids, /w16cid:paraId="([0-9A-F]{8})"/g),
			idsDurableIds: all(ids, /w16cid:durableId="([0-9A-F]{8})"/g),
			extensibleDurableIds: all(
				await readPart(docPath, "word/commentsExtensible.xml"),
				/w16cex:durableId="([0-9A-F]{8})"/g,
			),
		};
	}

	function expectConsistent(state: Awaited<ReturnType<typeof sidecarState>>) {
		const sorted = (values: (string | undefined)[]) => [...values].sort();
		expect(sorted(state.extendedParaIds)).toEqual(sorted(state.commentParaIds));
		expect(sorted(state.idsParaIds)).toEqual(sorted(state.commentParaIds));
		expect(sorted(state.extensibleDurableIds)).toEqual(
			sorted(state.idsDurableIds),
		);
		expect(new Set(state.idsDurableIds).size).toBe(state.idsDurableIds.length);
	}

	test("add, reply and delete keep every sidecar 1:1 with comments.xml", async () => {
		const docPath = await wordShapedDocument();

		await runCli(
			"comments",
			"add",
			docPath,
			"--anchor",
			"First",
			"--text",
			"A",
		);
		await runCli("comments", "add", docPath, "--at", "p0", "--text", "B");
		let state = await sidecarState(docPath);
		expect(state.commentParaIds).toHaveLength(3);
		expectConsistent(state);
		expect(state.idsDurableIds).toContain("1A2B3C4D");

		await runCli("comments", "reply", docPath, "--at", "c1", "--text", "R");
		state = await sidecarState(docPath);
		expect(state.commentParaIds).toHaveLength(4);
		expectConsistent(state);

		// Deleting the thread root cascades to its reply — both leave every part.
		await runCli("comments", "delete", docPath, "--at", "c1");
		state = await sidecarState(docPath);
		expect(state.commentParaIds).toHaveLength(2);
		expectConsistent(state);

		await runCli("comments", "delete", docPath, "--at", "c0");
		state = await sidecarState(docPath);
		expect(state.idsDurableIds).not.toContain("1A2B3C4D");
		expectConsistent(state);
		await expectValidClean(docPath);
	});

	test("a package without the sidecars stays without them", async () => {
		const workspace = tempWorkspace("comments-no-sidecars");
		const docPath = join(workspace, "plain.docx");
		await runCli("create", docPath, "--text", "Clause.");
		await runCli("comments", "add", docPath, "--at", "p0", "--text", "A");
		const pkg = await Pkg.open(docPath);
		for (const part of [
			"word/commentsExtended.xml",
			"word/commentsIds.xml",
			"word/commentsExtensible.xml",
		]) {
			expect(await pkg.readPart(part)).toBeUndefined();
		}
	});
});

describe("docx comments — a commentsExtended part we materialize covers every comment", () => {
	// A `create`d doc has no commentsExtended.xml until the first reply/resolve
	// mints it. It used to register only that thread, so `add` → `add` →
	// `reply c0` left c1 without an entry — the inconsistency `validate` reports
	// (Word writes one entry per comment once the part exists).
	async function commentedDoc(name: string): Promise<string> {
		const workspace = tempWorkspace(name);
		const docPath = join(workspace, "out.docx");
		const source = join(workspace, "source.md");
		writeFileSync(source, "Clause one.\n\nClause two.\n");
		await runCli("create", docPath, "--from", source);
		await runCli("comments", "add", docPath, "--at", "p0", "--text", "A");
		await runCli("comments", "add", docPath, "--at", "p1", "--text", "B");
		return docPath;
	}

	async function extendedEntries(docPath: string) {
		const xml = await readPart(docPath, "word/commentsExtended.xml");
		return [...xml.matchAll(/<w15:commentEx [^>]*\/>/g)].map((match) => ({
			paraId: match[0].match(/w15:paraId="([0-9A-F]{8})"/)?.[1],
			parent: match[0].match(/w15:paraIdParent="([0-9A-F]{8})"/)?.[1],
		}));
	}

	async function commentParaIds(docPath: string) {
		const xml = await readPart(docPath, "word/comments.xml");
		return [...xml.matchAll(/w14:paraId="([0-9A-F]{8})"/g)].map(
			(match) => match[1],
		);
	}

	test("the first reply registers the earlier unrelated comment too", async () => {
		const docPath = await commentedDoc("comments-ext-backfill-reply");
		expect(
			(await runCli("comments", "reply", docPath, "--at", "c0", "--text", "R"))
				.exitCode,
		).toBe(0);
		const entries = await extendedEntries(docPath);
		const paraIds = await commentParaIds(docPath);
		expect(entries.map((entry) => entry.paraId).sort()).toEqual(
			[...paraIds].sort(),
		);
		// Exactly one entry per comment — the reply isn't registered twice.
		expect(entries).toHaveLength(3);
		// comments.xml groups the reply under its thread: c0, reply, c1.
		expect(entries.filter((entry) => entry.parent)).toEqual([
			{ paraId: paraIds[1], parent: paraIds[0] },
		]);
		await expectValidClean(docPath);

		await runCli("comments", "add", docPath, "--at", "p1", "--text", "C");
		expect(await extendedEntries(docPath)).toHaveLength(4);
		await expectValidClean(docPath);
	});

	test("the first resolve registers every comment, resolving only its target", async () => {
		const docPath = await commentedDoc("comments-ext-backfill-resolve");
		await runCli("comments", "resolve", docPath, "--at", "c1");
		const xml = await readPart(docPath, "word/commentsExtended.xml");
		expect(xml.match(/<w15:commentEx /g)).toHaveLength(2);
		expect(xml.match(/w15:done="1"/g)).toHaveLength(1);
		await expectValidClean(docPath);
	});
});

describe("docx comments delete — markers in note parts", () => {
	test("deleting a comment anchored in a footnote strips its footnote markers", async () => {
		// Word lets a reviewer comment inside a footnote; the markers then live in
		// footnotes.xml. `delete` used to prune document.xml only, leaving a
		// `<w:commentReference>` that names no comment.
		const workspace = tempWorkspace("comments-delete-footnote");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Body clause.");
		await runCli("footnotes", "add", docPath, "--at", "p0", "--text", "Note.");
		await runCli("comments", "add", docPath, "--at", "p0", "--text", "A");
		// Move c0's markers from the body into the footnote, as Word writes them.
		const pkg = await Pkg.open(docPath);
		const body = (await pkg.readText("word/document.xml"))
			.replace(/<w:commentRange(Start|End) w:id="0"\/>/g, "")
			.replace(
				/<w:r>(?:(?!<w:r>).)*?<w:commentReference w:id="0"\/><\/w:r>/,
				"",
			);
		expect(body).not.toContain('w:id="0"/><w:r');
		const notes = (await pkg.readText("word/footnotes.xml")).replace(
			/(<w:r>(?:(?!<w:r>).)*?Note\.<\/w:t><\/w:r>)/,
			'<w:commentRangeStart w:id="0"/>$1<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>',
		);
		expect(notes).toContain('<w:commentReference w:id="0"/>');
		pkg.writeText("word/document.xml", body);
		pkg.writeText("word/footnotes.xml", notes);
		await pkg.save();
		await expectValidClean(docPath);

		expect(
			(await runCli("comments", "delete", docPath, "--at", "c0")).exitCode,
		).toBe(0);
		expect(await readPart(docPath, "word/footnotes.xml")).not.toContain(
			"comment",
		);
		await expectValidClean(docPath);
	});
});

describe("docx comments — removing one paragraph of a cross-paragraph comment", () => {
	// A comment spanning p0..p1 has its start in p0 and its end + reference in
	// p1. Removing either paragraph used to take its half along, leaving the
	// other half unbalanced — the inconsistency `validate` reports.
	async function spanningComment(name: string): Promise<string> {
		const workspace = tempWorkspace(name);
		const docPath = join(workspace, "out.docx");
		const source = join(workspace, "source.md");
		writeFileSync(source, "First clause.\n\nSecond clause.\n\nThird clause.\n");
		await runCli("create", docPath, "--from", source);
		await runCli(
			"comments",
			"add",
			docPath,
			"--at",
			"p0:6-p1:6",
			"--text",
			"A",
		);
		const comments = JSON.parse(
			(await runCli("comments", "list", docPath)).stdout,
		);
		expect(comments[0].anchor).toMatchObject({
			startBlockId: "p0",
			endBlockId: "p1",
		});
		return docPath;
	}

	async function markerCounts(docPath: string) {
		const xml = await readDocumentXml(docPath);
		const count = (tag: string) =>
			xml.match(new RegExp(`<w:${tag} w:id="0"/>`, "g"))?.length ?? 0;
		return [
			count("commentRangeStart"),
			count("commentRangeEnd"),
			count("commentReference"),
		];
	}

	for (const [label, args] of [
		["delete the start paragraph", ["delete", "--at", "p0"]],
		["delete the end paragraph", ["delete", "--at", "p1"]],
		[
			"edit --text '' on the start paragraph",
			["edit", "--at", "p0", "--text", ""],
		],
		[
			"edit --text '' on the end paragraph",
			["edit", "--at", "p1", "--text", ""],
		],
		[
			"edit --runs '[]' on the end paragraph",
			["edit", "--at", "p1", "--runs", "[]"],
		],
	] as const) {
		test(`${label} keeps the comment's markers balanced`, async () => {
			const docPath = await spanningComment(
				`comments-straddle-${label.replace(/\W+/g, "-")}`,
			);
			const [verb, ...rest] = args;
			expect((await runCli(verb, docPath, ...rest)).exitCode).toBe(0);
			expect(await markerCounts(docPath)).toEqual([1, 1, 1]);
			await expectValidClean(docPath);
		});
	}
});

describe("docx comments delete — malformed threading", () => {
	test("a paraIdParent cycle in commentsExtended.xml doesn't hang the reply cascade", async () => {
		const workspace = tempWorkspace("comments-thread-cycle");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Clause.");
		await runCli("comments", "add", docPath, "--at", "p0", "--text", "A");
		await runCli("comments", "reply", docPath, "--at", "c0", "--text", "R");
		// Point the root back at its own reply: root → reply → root → …
		const pkg = await Pkg.open(docPath);
		const extended = await pkg.readText("word/commentsExtended.xml");
		const [, replyParaId, rootParaId] =
			extended.match(
				/w15:paraId="([0-9A-F]+)" w15:paraIdParent="([0-9A-F]+)"/,
			) ?? [];
		expect(rootParaId).toBeDefined();
		pkg.writeText(
			"word/commentsExtended.xml",
			extended.replace(
				`w15:paraId="${rootParaId}" w15:done`,
				`w15:paraId="${rootParaId}" w15:paraIdParent="${replyParaId}" w15:done`,
			),
		);
		await pkg.save();

		const deleted = await runCli("comments", "delete", docPath, "--at", "c0");
		expect(deleted.exitCode).toBe(0);
		expect(
			JSON.parse((await runCli("comments", "list", docPath)).stdout),
		).toEqual([]);
	}, 10_000);
});

// An equation is its `read` token in --anchor/--at offsets, so a phrase copied
// from `read` anchors exactly, and a range can't open or close inside one.
describe("comments over an inline equation", () => {
	async function equationDoc(): Promise<string> {
		const workspace = tempWorkspace("comments-equation");
		const docPath = join(workspace, "eq.docx");
		const source = join(workspace, "eq.md");
		await Bun.write(source, "Energy is $E=mc^2$ in physics.\n");
		await runCli("create", docPath, "--from", source);
		return docPath;
	}
	const markerOrder = async (docPath: string) =>
		(await readDocumentXml(docPath)).match(
			/<w:commentRange(?:Start|End)|<m:oMath>|<\/m:oMath>|physics/g,
		);

	test("--anchor across the token brackets the whole equation", async () => {
		const docPath = await equationDoc();
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--anchor",
			"is $E=mc^2$ in",
			"--text",
			"Cite this.",
		);
		expect(result.exitCode).toBe(0);
		expect(await markerOrder(docPath)).toEqual([
			"<w:commentRangeStart",
			"<m:oMath>",
			"</m:oMath>",
			"<w:commentRangeEnd",
			"physics",
		]);
	});

	test("text after the equation anchors at its read offsets", async () => {
		const docPath = await equationDoc();
		await runCli(
			"comments",
			"add",
			docPath,
			"--anchor",
			"physics",
			"--text",
			"Which branch?",
		);
		const [comment] = (await runCli("comments", "list", docPath)).parsed as {
			anchor?: object;
		}[];
		// The same offsets `find physics` reports past the 8-char token.
		expect(comment?.anchor).toEqual({
			startBlockId: "p0",
			startOffset: 22,
			endBlockId: "p0",
			endOffset: 29,
		});
	});

	test("a span strictly inside the token snaps outward around the equation", async () => {
		const docPath = await equationDoc();
		await runCli(
			"comments",
			"add",
			docPath,
			"--at",
			"p0:13-17",
			"--text",
			"Squared?",
		);
		expect(await markerOrder(docPath)).toEqual([
			"<w:commentRangeStart",
			"<m:oMath>",
			"</m:oMath>",
			"<w:commentRangeEnd",
			"physics",
		]);
	});
});

// Release-merge cross-check: equation-anchored comments × the comment-integrity
// fixes. Deleting the paragraph that holds a comment's WHOLE range used to drop
// only its markers, leaving an anchorless `<w:comment>` (validate's "no
// anchor" warning; Word deletes it on open). It must go with its anchor, while
// a cross-paragraph comment that only straddles the line keeps both halves.
describe("comments wholly on a removed paragraph go with it", () => {
	async function commentedEquationDoc(name: string): Promise<string> {
		const workspace = tempWorkspace(name);
		const docPath = join(workspace, "eq.docx");
		const source = join(workspace, "eq.md");
		await Bun.write(
			source,
			"Intro paragraph.\n\nEnergy is $E=mc^2$ in physics.\n\nTail paragraph.\n",
		);
		await runCli("create", docPath, "--from", source);
		await runCli(
			"comments",
			"add",
			docPath,
			"--anchor",
			"$E=mc^2$",
			"--text",
			"On the equation",
		);
		await runCli("comments", "reply", docPath, "--at", "c0", "--text", "Ack");
		await runCli(
			"comments",
			"add",
			docPath,
			"--at",
			"p0:6-p1:12",
			"--text",
			"Straddles into the equation",
		);
		return docPath;
	}

	const remainingTexts = async (docPath: string) =>
		(
			(await runCli("comments", "list", docPath)).parsed as { text: string }[]
		).map((comment) => comment.text);

	test("delete --at pN removes the comment and its reply, keeps the straddler", async () => {
		const docPath = await commentedEquationDoc("comments-contained-delete");
		expect((await runCli("delete", docPath, "--at", "p1")).exitCode).toBe(0);
		expect(await remainingTexts(docPath)).toEqual([
			"Straddles into the equation",
		]);
		await expectValidClean(docPath);
	});

	test("edit --text '' on the line does the same", async () => {
		const docPath = await commentedEquationDoc("comments-contained-blank");
		expect(
			(await runCli("edit", docPath, "--at", "p1", "--text", "")).exitCode,
		).toBe(0);
		expect(await remainingTexts(docPath)).toEqual([
			"Straddles into the equation",
		]);
		await expectValidClean(docPath);
	});

	test("a cell's last paragraph is blanked, and its comment still goes", async () => {
		const workspace = tempWorkspace("comments-contained-cell");
		const docPath = join(workspace, "cell.docx");
		const source = join(workspace, "cell.md");
		await Bun.write(source, "| A | B |\n| --- | --- |\n| one | two |\n");
		await runCli("create", docPath, "--from", source);
		await runCli(
			"comments",
			"add",
			docPath,
			"--at",
			"t0:r1c0:p0",
			"--text",
			"Cell note",
		);
		expect(
			(await runCli("delete", docPath, "--at", "t0:r1c0:p0")).exitCode,
		).toBe(0);
		expect(await remainingTexts(docPath)).toEqual([]);
		await expectValidClean(docPath);
	});

	test("read --comments quotes an equation as its read token", async () => {
		const docPath = await commentedEquationDoc("comments-equation-quote");
		const read = await runCli("read", docPath, "--comments");
		expect(read.stdout).toContain('[^c0]: "$E=mc^2$" — ');
	});
});
