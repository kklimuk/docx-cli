import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

// Editing a paragraph that carries a comment must re-anchor the comment to the
// new content, not collapse it to a zero-length range (the Task-2 orphan bug).

type Anchor = {
	startOffset: number;
	endOffset: number;
	startBlockId: string;
};

async function doc(label: string, md: string): Promise<string> {
	const workspace = tempWorkspace(label);
	const src = join(workspace, "src.md");
	await Bun.write(src, md);
	const path = join(workspace, "out.docx");
	expect((await runCli("create", path, "--from", src)).exitCode).toBe(0);
	return path;
}

async function firstAnchor(path: string): Promise<Anchor> {
	const list = await runCli("comments", "list", path);
	const first = (list.parsed as Array<{ anchor: Anchor }>)[0];
	if (!first) throw new Error("expected at least one comment");
	return first.anchor;
}

describe("edit re-anchors comments instead of orphaning them", () => {
	test("untracked --text edit keeps the comment spanning the new text", async () => {
		const path = await doc("reanchor-text", "The tower opened in 1879.\n");
		expect(
			(
				await runCli(
					"comments",
					"add",
					path,
					"--at",
					"p0",
					"--text",
					"check year",
				)
			).exitCode,
		).toBe(0);
		expect(
			(
				await runCli(
					"edit",
					path,
					"--at",
					"p0",
					"--text",
					"The tower opened in 1889.",
				)
			).exitCode,
		).toBe(0);
		const anchor = await firstAnchor(path);
		expect(anchor.startOffset).toBe(0);
		// Spans the whole rewritten line (~25 chars), not a near-collapse.
		expect(anchor.endOffset - anchor.startOffset).toBeGreaterThanOrEqual(20);
	});

	test("tracked --text edit keeps the comment anchored (Task-2 regression)", async () => {
		const path = await doc("reanchor-tracked", "Completed in 1879.\n");
		expect(
			(
				await runCli(
					"comments",
					"add",
					path,
					"--at",
					"p0",
					"--text",
					"1889 not 1879",
				)
			).exitCode,
		).toBe(0);
		expect((await runCli("track-changes", path, "on")).exitCode).toBe(0);
		expect(
			(await runCli("edit", path, "--at", "p0", "--text", "Completed in 1889."))
				.exitCode,
		).toBe(0);
		const anchor = await firstAnchor(path);
		expect(anchor.startOffset).toBe(0);
		// Must cover the whole edited line, not collapse to a 1-char range or
		// drift to span only the inserted word (the Task-2 bug + off-by-one guard).
		expect(anchor.endOffset - anchor.startOffset).toBeGreaterThanOrEqual(15);
	});

	test("--markdown rewrite of a commented paragraph re-anchors the comment", async () => {
		const path = await doc("reanchor-md", "Old wording here.\n");
		expect(
			(await runCli("comments", "add", path, "--at", "p0", "--text", "reword"))
				.exitCode,
		).toBe(0);
		expect(
			(
				await runCli(
					"edit",
					path,
					"--at",
					"p0",
					"--markdown",
					"New wording entirely.",
				)
			).exitCode,
		).toBe(0);
		const anchor = await firstAnchor(path);
		expect(anchor.startOffset).toBe(0);
		// Spans the whole new paragraph ("New wording entirely." ~21 chars).
		expect(anchor.endOffset - anchor.startOffset).toBeGreaterThanOrEqual(15);
	});

	test("editing a different paragraph leaves an existing comment intact", async () => {
		const path = await doc(
			"reanchor-other",
			"First paragraph.\n\nSecond paragraph.\n",
		);
		expect(
			(await runCli("comments", "add", path, "--at", "p1", "--text", "note"))
				.exitCode,
		).toBe(0);
		expect(
			(await runCli("edit", path, "--at", "p0", "--text", "Edited first."))
				.exitCode,
		).toBe(0);
		const anchor = await firstAnchor(path);
		expect(anchor.startBlockId).toBe("p1");
		expect(anchor.endOffset).toBeGreaterThan(anchor.startOffset);
	});
});
