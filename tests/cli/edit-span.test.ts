import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

const FIXTURE = "tests/fixtures/word-formatted.docx";
// p1: "The " + "MESSENGER" [bold,#800080] + " is " + "fatally flawed" [italic] + "."

type Run = {
	type: string;
	text?: string;
	bold?: boolean;
	italic?: boolean;
	color?: string;
};

async function freshCopy(label: string): Promise<string> {
	const workspace = tempWorkspace(label);
	const docPath = join(workspace, "out.docx");
	await Bun.write(docPath, Bun.file(FIXTURE));
	return docPath;
}

async function readParagraph(docPath: string, blockId: string): Promise<Run[]> {
	const result = await runCli("read", docPath, "--ast");
	const blocks = (
		result.parsed as { blocks: Array<{ id: string; runs?: Run[] }> }
	).blocks;
	return blocks.find((candidate) => candidate.id === blockId)?.runs ?? [];
}

async function firstLocator(docPath: string, phrase: string): Promise<string> {
	// The harness injects --json into find, so read the structured payload.
	const result = await runCli("find", docPath, phrase);
	expect(result.exitCode).toBe(0);
	const locator = (result.parsed as { matches?: Array<{ locator: string }> })
		.matches?.[0]?.locator;
	if (!locator) throw new Error(`no match for ${phrase}`);
	return locator;
}

function flat(runs: Run[]): string {
	return runs
		.filter((run) => run.type === "text")
		.map((run) => run.text ?? "")
		.join("");
}

describe("docx edit --at pN:S-E — character-span edit", () => {
	test("find → edit --at <span> replaces just that span, inheriting its rPr", async () => {
		const docPath = await freshCopy("span-inherit");
		const locator = await firstLocator(docPath, "MESSENGER"); // p1:4-13
		expect(locator).toBe("p1:4-13");

		const result = await runCli(
			"edit",
			docPath,
			"--at",
			locator,
			"--text",
			"COURIER",
		);
		expect(result.exitCode).toBe(0);

		const runs = await readParagraph(docPath, "p1");
		expect(flat(runs)).toBe("The COURIER is fatally flawed.");
		// Replacement inherited bold + color from the old MESSENGER run.
		const replaced = runs.find((run) => run.text === "COURIER");
		expect(replaced?.bold).toBe(true);
		expect(replaced?.color).toBe("800080");
		// Untouched neighbor keeps its italic.
		expect(runs.find((run) => run.text === "fatally flawed")?.italic).toBe(
			true,
		);
	});

	test("sub-run span inherits italic and splits the run cleanly", async () => {
		const docPath = await freshCopy("span-subrun");
		const locator = await firstLocator(docPath, "fatally"); // within the italic run
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			locator,
			"--text",
			"deeply",
		);
		expect(result.exitCode).toBe(0);

		const runs = await readParagraph(docPath, "p1");
		expect(flat(runs)).toBe("The MESSENGER is deeply flawed.");
		// The "deeply" portion is still italic (inherited from the split run).
		const deeply = runs.find((run) => run.text === "deeply");
		expect(deeply?.italic).toBe(true);
	});

	test("span edit on a heading preserves its paragraph style", async () => {
		const workspace = tempWorkspace("span-heading");
		const md = join(workspace, "src.md");
		await Bun.write(md, "# Quarterly Design Review\n");
		const docPath = join(workspace, "h.docx");
		expect((await runCli("create", docPath, "--from", md)).exitCode).toBe(0);

		const locator = await firstLocator(docPath, "Quarterly"); // p0 heading
		expect(
			(await runCli("edit", docPath, "--at", locator, "--text", "Q3")).exitCode,
		).toBe(0);

		// Still a heading (outline only lists styled headings).
		const outline = await runCli("outline", docPath);
		expect(outline.stdout).toContain("Q3 Design Review");
	});

	test("cell-paragraph span (tN:rRcC:pK:S-E) edits in place", async () => {
		const workspace = tempWorkspace("span-cell");
		const md = join(workspace, "t.md");
		await Bun.write(
			md,
			"| Field | Value |\n| --- | --- |\n| State | fill in state |\n",
		);
		const docPath = join(workspace, "t.docx");
		expect((await runCli("create", docPath, "--from", md)).exitCode).toBe(0);

		const locator = await firstLocator(docPath, "fill in state");
		expect(locator).toContain(":"); // a cell-scoped span locator
		expect(
			(await runCli("edit", docPath, "--at", locator, "--text", "Delaware"))
				.exitCode,
		).toBe(0);

		const read = await runCli("read", docPath);
		expect(read.stdout).toContain("Delaware");
		expect(read.stdout).not.toContain("fill in state");
	});

	test("tracked span edit produces a del/ins pair attributed to the default author", async () => {
		const docPath = await freshCopy("span-tracked");
		expect((await runCli("track-changes", docPath, "on")).exitCode).toBe(0);
		const locator = await firstLocator(docPath, "MESSENGER");
		expect(
			(await runCli("edit", docPath, "--at", locator, "--text", "COURIER"))
				.exitCode,
		).toBe(0);

		const list = await runCli("track-changes", "list", docPath);
		const changes = list.parsed as Array<{
			kind: string;
			text: string;
			author: string;
		}>;
		expect(
			changes.some(
				(change) => change.kind === "del" && change.text === "MESSENGER",
			),
		).toBe(true);
		expect(
			changes.some(
				(change) => change.kind === "ins" && change.text === "COURIER",
			),
		).toBe(true);
		expect(changes.every((change) => change.author === "Reviewer")).toBe(true);
	});

	test("out-of-range span fails with INVALID_LOCATOR", async () => {
		const docPath = await freshCopy("span-oob");
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p1:5-999",
			"--text",
			"X",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("INVALID_LOCATOR");
	});

	test("--markdown on a span is rejected with a pointer to whole-paragraph edit", async () => {
		const docPath = await freshCopy("span-md");
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p1:0-3",
			"--markdown",
			"# H",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});
});
