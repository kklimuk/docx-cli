import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";
import {
	buildRawDoc,
	expectValidClean,
	readDocumentXml,
	readMarkdown,
	storyXml,
	wordTextBoxParagraphXml,
} from "./helpers";

/** `docx move` (issue #38): an untracked splice of the LIVE blocks, so
 *  everything inside them rides along — the read → insert-a-copy → delete
 *  workaround orphaned the comment and turned the footnote into literal
 *  `[^fn1]` text. */

const CLAUSES = `# Terms

1. Payment due in 30 days.
2. Late fees apply.

Clause B: Governing law is Delaware.

Clause C: Notices go to the address above.
`;

async function createDoc(label: string, markdown: string): Promise<string> {
	const workspace = tempWorkspace(label);
	const source = join(workspace, "source.md");
	const docPath = join(workspace, "doc.docx");
	await Bun.write(source, markdown);
	const created = await runCli("create", docPath, "--from", source);
	expect(created.exitCode).toBe(0);
	return docPath;
}

/** The text `read` shows, locator comments and blank lines stripped. */
async function readLines(docPath: string): Promise<string[]> {
	return (await readMarkdown(docPath))
		.split("\n")
		.filter((line) => line.trim() !== "" && !line.startsWith("<!--"))
		.map((line) => line.replace(/ <!-- [^>]*-->/g, ""));
}

function paragraphContaining(xml: string, text: string): string {
	const match = [...xml.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].find((entry) =>
		entry[0].includes(text),
	);
	if (!match) throw new Error(`no paragraph containing ${text}`);
	return match[0];
}

describe("move — the issue #38 repro", () => {
	test("comment, footnote and run formatting ride along; new locator printed", async () => {
		const docPath = await createDoc("move-repro", CLAUSES);
		await runCli(
			"edit",
			docPath,
			"--at",
			"p3:0-8",
			"--text",
			"Clause B",
			"--bold",
			"--color",
			"C00000",
		);
		await runCli(
			"comments",
			"add",
			docPath,
			"--anchor",
			"Delaware",
			"--text",
			"Confirm with legal",
		);
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--anchor",
			"Delaware",
			"--text",
			"See Del. Code.",
		);
		const before = paragraphContaining(
			await readDocumentXml(docPath),
			"Governing law",
		);

		const result = await runCli("move", docPath, "--at", "p3", "--after", "p4");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toBe("p4\n");

		// The paragraph moved byte-for-byte: markers, reference, formatting.
		const after = paragraphContaining(
			await readDocumentXml(docPath),
			"Governing law",
		);
		expect(after).toBe(before);
		const lines = await readLines(docPath);
		expect(lines.slice(-3)).toEqual([
			"Clause C: Notices go to the address above.",
			'<span style="color:#C00000">**Clause B**</span>: Governing law is Delaware[^fn1].',
			"[^fn1]: See Del. Code.",
		]);
		const comments = (await runCli("comments", "list", docPath)).parsed as {
			anchor: { startBlockId: string };
		}[];
		expect(comments[0]?.anchor.startBlockId).toBe("p4");
		await expectValidClean(docPath);
	});
});

describe("move — sources", () => {
	test("a range carries the blocks between its endpoints, tables included", async () => {
		const docPath = await createDoc(
			"move-range",
			"A.\n\nB.\n\n| x | y |\n|---|---|\n| 1 | 2 |\n\nC.\n\nD.\n",
		);
		const result = await runCli(
			"move",
			docPath,
			"--at",
			"p1-p2",
			"--after",
			"p3",
		);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toBe("p2\nt0\np3\n");
		expect(
			(await readLines(docPath)).filter((line) => !line.startsWith("|")),
		).toEqual(["A.", "D.", "B.", "C."]);
		await expectValidClean(docPath);
	});

	test("--subtree moves a heading with everything up to the next same-level heading", async () => {
		const docPath = await createDoc(
			"move-subtree",
			"# Intro\n\nIntro text.\n\n# Details\n\nDetail text.\n\n## Sub\n\nSub text.\n\n# Summary\n\nSummary text.\n",
		);
		const result = await runCli(
			"move",
			docPath,
			"--at",
			"p2",
			"--subtree",
			"--after",
			"p7",
		);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toBe("p4\np5\np6\np7\n");
		expect(await readLines(docPath)).toEqual([
			"# Intro",
			"Intro text.",
			"# Summary",
			"Summary text.",
			"# Details",
			"Detail text.",
			"## Sub",
			"Sub text.",
		]);
	});

	test("--subtree on a non-heading is refused", async () => {
		const docPath = await createDoc("move-subtree-plain", "A.\n\nB.\n");
		const result = await runCli(
			"move",
			docPath,
			"--at",
			"p0",
			"--subtree",
			"--after",
			"p1",
		);
		expect(result.exitCode).toBe(2);
		expect(result.stdout).toContain("--subtree needs a heading");
	});

	test("a table moves whole", async () => {
		const docPath = await createDoc(
			"move-table",
			"Lead.\n\n| x | y |\n|---|---|\n| 1 | 2 |\n\nTail.\n",
		);
		const result = await runCli("move", docPath, "--at", "t0", "--after", "p1");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toBe("t0\n");
		const lines = await readLines(docPath);
		expect(lines[0]).toBe("Lead.");
		expect(lines[1]).toBe("Tail.");
		expect(lines[2]).toStartWith("| x");
		await expectValidClean(docPath);
	});
});

describe("move — table cells", () => {
	const TABLE = "A one.\n\n| H1 | H2 |\n|---|---|\n| x |  |\n\nB two.\n";

	test("into a bare empty cell fills its mandatory paragraph", async () => {
		const docPath = await createDoc("move-into-cell", TABLE);
		const result = await runCli(
			"move",
			docPath,
			"--at",
			"p0",
			"--after",
			"t0:r1c1",
		);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toBe("t0:r1c1:p0\n");
		const markdown = await readMarkdown(docPath);
		expect(markdown).toContain("| A one. <!-- t0:r1c1:p0 --> |");
		expect(markdown).not.toContain("<!-- t0:r1c1:p1 -->");
		await expectValidClean(docPath);
	});

	test("out of a cell leaves Word's mandatory empty paragraph behind", async () => {
		const docPath = await createDoc("move-out-of-cell", TABLE);
		const result = await runCli(
			"move",
			docPath,
			"--at",
			"t0:r1c0:p0",
			"--before",
			"p0",
		);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toBe("p0\n");
		const markdown = await readMarkdown(docPath);
		expect(markdown).toContain("x <!-- p0 -->");
		expect(markdown).toContain("<!-- t0:r1c0 -->");
		await expectValidClean(docPath);
	});
});

describe("move — text boxes", () => {
	test("reorders paragraphs within one story; refuses leaving it", async () => {
		const docPath = await buildRawDoc(
			`${storyXml(["Body."])}${wordTextBoxParagraphXml(["First.", "Second."])}`,
			"move-textbox",
		);
		const result = await runCli(
			"move",
			docPath,
			"--at",
			"tbx0:p0",
			"--after",
			"tbx0:p1",
		);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toBe("tbx0:p1\n");
		const xml = await readDocumentXml(docPath);
		// Both branches of the AlternateContent carry the new order.
		const stories = [
			...xml.matchAll(/<w:txbxContent>([\s\S]*?)<\/w:txbxContent>/g),
		].map((match) => match[1] ?? "");
		expect(stories).toHaveLength(2);
		for (const story of stories) {
			expect(story.indexOf("Second.")).toBeLessThan(story.indexOf("First."));
		}

		const across = await runCli(
			"move",
			docPath,
			"--at",
			"tbx0:p0",
			"--after",
			"p0",
		);
		expect(across.exitCode).toBe(1);
		expect(across.stdout).toContain('"code":"UNSUPPORTED"');
	});
});

describe("move — refusals write nothing", () => {
	async function expectRefused(
		docPath: string,
		args: string[],
		exitCode: number,
		message: string,
	): Promise<void> {
		const before = await readDocumentXml(docPath);
		const result = await runCli("move", docPath, ...args);
		expect(result.exitCode).toBe(exitCode);
		expect(result.stdout).toContain(message);
		expect(await readDocumentXml(docPath)).toBe(before);
	}

	test("a move that lands where the blocks already are", async () => {
		const docPath = await createDoc("move-noop", "A.\n\nB.\n\nC.\n");
		await expectRefused(
			docPath,
			["--at", "p1", "--after", "p0"],
			2,
			"already after p0",
		);
		await expectRefused(
			docPath,
			["--at", "p1", "--before", "p2"],
			2,
			"already before p2",
		);
	});

	test("a destination inside the source", async () => {
		const docPath = await createDoc("move-inside", "A.\n\nB.\n\nC.\n");
		await expectRefused(
			docPath,
			["--at", "p0-p2", "--after", "p1"],
			2,
			"part of what's being moved",
		);
	});

	test("a comment with one end inside the source", async () => {
		const docPath = await createDoc(
			"move-straddle",
			"A one.\n\nB two.\n\nC three.\n",
		);
		await runCli(
			"comments",
			"add",
			docPath,
			"--at",
			"p0:0-p1:3",
			"--text",
			"hi",
		);
		await expectRefused(
			docPath,
			["--at", "p1", "--after", "p2"],
			1,
			"Comment c0 starts on one side",
		);
		// The whole range moves fine.
		const result = await runCli(
			"move",
			docPath,
			"--at",
			"p0-p1",
			"--after",
			"p2",
		);
		expect(result.exitCode).toBe(0);
		await expectValidClean(docPath);
	});

	test("a paragraph carrying a section break, and a section-break anchor", async () => {
		const docPath = await createDoc("move-sectpr", "A.\n\nB.\n\nC.\n");
		await runCli("sections", docPath, "--at", "p0-p0", "--columns", "2");
		const ids = (await runCli("read", docPath, "--ast")).parsed as {
			blocks: { id: string; type: string }[];
		};
		const sentinel = ids.blocks.findIndex(
			(block) => block.type === "sectionBreak",
		);
		const carrier = ids.blocks[sentinel - 1]?.id as string;
		await expectRefused(
			docPath,
			["--at", carrier, "--after", "p3"],
			1,
			"carries a section break",
		);
		await expectRefused(
			docPath,
			["--at", "p3", "--after", "s0"],
			2,
			"can't anchor a move",
		);
	});

	test("track changes on", async () => {
		const docPath = await createDoc("move-tracked", "A.\n\nB.\n");
		await runCli("track-changes", "on", docPath);
		await expectRefused(
			docPath,
			["--at", "p0", "--after", "p1"],
			1,
			"tracked moves aren't supported yet",
		);
	});

	test("usage errors", async () => {
		const docPath = await createDoc("move-usage", "A.\n\nB.\n");
		await expectRefused(docPath, ["--at", "p0"], 2, "Missing destination");
		await expectRefused(
			docPath,
			["--at", "p0", "--after", "p1", "--before", "p1"],
			2,
			"exactly one of --after or --before",
		);
		await expectRefused(
			docPath,
			["--at", "p0:0-1", "--after", "p1"],
			2,
			"whole blocks",
		);
		await expectRefused(
			docPath,
			["--at", "p9", "--after", "p1"],
			3,
			"Block not found",
		);
	});

	test("--dry-run validates but writes nothing", async () => {
		const docPath = await createDoc("move-dry", "A.\n\nB.\n");
		const before = await readDocumentXml(docPath);
		const result = await runCli(
			"move",
			docPath,
			"--at",
			"p0",
			"--after",
			"p1",
			"--dry-run",
		);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain('"dryRun":true');
		expect(await readDocumentXml(docPath)).toBe(before);
	});
});

describe("move --batch", () => {
	async function batchFile(lines: object[]): Promise<string> {
		const path = join(tempWorkspace("move-batch-jsonl"), "moves.jsonl");
		await Bun.write(path, lines.map((line) => JSON.stringify(line)).join("\n"));
		return path;
	}

	test("locators address the document as read; stacked moves keep entry order", async () => {
		const docPath = await createDoc(
			"move-batch",
			"A.\n\nB.\n\nC.\n\nD.\n\nE.\n",
		);
		const batch = await batchFile([
			{ at: "p0", after: "p4" },
			{ at: "p1", after: "p4" },
			{ at: "p3", before: "p2" },
		]);
		const result = await runCli("move", docPath, "--batch", batch);
		expect(result.exitCode).toBe(0);
		expect(await readLines(docPath)).toEqual(["D.", "C.", "E.", "A.", "B."]);
		expect(result.stdout).toBe("p0\np3\np4\n");
	});

	test("an empty cell's paragraph moved out by one entry isn't reused by another", async () => {
		const docPath = await createDoc(
			"move-batch-cell",
			"A one.\n\n| H1 | H2 |\n|---|---|\n| x |  |\n\nB two.\n",
		);
		const result = await runCli(
			"move",
			docPath,
			"--batch",
			await batchFile([
				{ at: "t0:r1c1:p0", after: "p1" },
				{ at: "p0", after: "t0:r1c1" },
			]),
		);
		expect(result.exitCode).toBe(0);
		expect(await readMarkdown(docPath)).toContain(
			"| A one. <!-- t0:r1c1:p0 --> |",
		);
		await expectValidClean(docPath);
	});

	test("overlapping sources and anchors on moved blocks are refused", async () => {
		const docPath = await createDoc("move-batch-bad", "A.\n\nB.\n\nC.\n");
		const before = await readDocumentXml(docPath);
		const overlap = await runCli(
			"move",
			docPath,
			"--batch",
			await batchFile([
				{ at: "p0-p1", after: "p2" },
				{ at: "p1", after: "p2" },
			]),
		);
		expect(overlap.exitCode).toBe(2);
		expect(overlap.stdout).toContain("entry 1: p1 overlaps p0-p1");

		const anchored = await runCli(
			"move",
			docPath,
			"--batch",
			await batchFile([
				{ at: "p0", after: "p2" },
				{ at: "p1", before: "p0" },
			]),
		);
		expect(anchored.exitCode).toBe(2);
		expect(anchored.stdout).toContain("entry 1: p0 is moved by entry 0");
		expect(await readDocumentXml(docPath)).toBe(before);
	});

	test("rejects --at beside --batch and malformed entries", async () => {
		const docPath = await createDoc("move-batch-usage", "A.\n\nB.\n");
		const batch = await batchFile([{ at: "p0", after: "p1" }]);
		const mixed = await runCli("move", docPath, "--batch", batch, "--at", "p0");
		expect(mixed.exitCode).toBe(2);
		const malformed = await runCli(
			"move",
			docPath,
			"--batch",
			await batchFile([{ at: "p0", to: "p1" }]),
		);
		expect(malformed.exitCode).toBe(2);
		expect(malformed.stdout).toContain(String.raw`unknown key \"to\"`);
	});
});
