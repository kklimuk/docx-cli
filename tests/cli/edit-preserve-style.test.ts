import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

// Decision 2: a whole-paragraph edit with plain content (no explicit --style,
// no block markers) preserves the paragraph's existing style — re-titling a
// heading keeps it a heading. Markdown that carries its own block style wins.

async function docFrom(label: string, md: string): Promise<string> {
	const workspace = tempWorkspace(label);
	const src = join(workspace, "src.md");
	await Bun.write(src, md);
	const docPath = join(workspace, "out.docx");
	expect((await runCli("create", docPath, "--from", src)).exitCode).toBe(0);
	return docPath;
}

describe("edit preserves paragraph style on plain whole-paragraph edits", () => {
	test("--markdown plain text on a heading keeps the heading style", async () => {
		const docPath = await docFrom("md-heading", "# Quarterly Design Review\n");
		expect(
			(await runCli("edit", docPath, "--at", "p0", "--markdown", "Q3 Review"))
				.exitCode,
		).toBe(0);
		expect((await runCli("outline", docPath)).stdout).toContain("Q3 Review");
	});

	test("--text on a heading keeps the heading style", async () => {
		const docPath = await docFrom("text-heading", "# Original Title\n");
		expect(
			(await runCli("edit", docPath, "--at", "p0", "--text", "New Title"))
				.exitCode,
		).toBe(0);
		expect((await runCli("outline", docPath)).stdout).toContain("New Title");
	});

	test("--markdown that sets its own block style wins (## → Heading2)", async () => {
		const docPath = await docFrom("md-override", "# Big\n");
		expect(
			(await runCli("edit", docPath, "--at", "p0", "--markdown", "## Smaller"))
				.exitCode,
		).toBe(0);
		const ast = await runCli("read", docPath, "--ast");
		const block = (
			ast.parsed as { blocks: Array<{ id: string; style?: string }> }
		).blocks[0];
		expect(block?.style).toBe("Heading2");
	});

	test("rewording a plain paragraph does not invent a heading", async () => {
		const docPath = await docFrom("plain", "Just a plain paragraph.\n");
		expect(
			(
				await runCli(
					"edit",
					docPath,
					"--at",
					"p0",
					"--markdown",
					"Reworded plainly",
				)
			).exitCode,
		).toBe(0);
		// outline lists only styled headings; a still-plain paragraph never
		// appears there (it would if the edit had invented a heading style).
		expect((await runCli("outline", docPath)).stdout).not.toContain(
			"Reworded plainly",
		);
	});
});
