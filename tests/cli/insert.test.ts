import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";
import { newTableDoc, trackedKinds } from "./helpers";

type AstParagraph = {
	id: string;
	type: string;
	style?: string;
	spacing?: Record<string, unknown>;
	indent?: Record<string, unknown>;
	runs?: Array<{ type: string; text?: string }>;
};

async function readParagraphs(docPath: string): Promise<AstParagraph[]> {
	const read = await runCli("read", docPath, "--ast");
	const doc = read.parsed as { blocks: AstParagraph[] };
	return doc.blocks.filter((block) => block.type === "paragraph");
}

function paragraphText(paragraph: AstParagraph | undefined): string {
	return (paragraph?.runs ?? []).map((run) => run.text ?? "").join("");
}

describe("docx insert / edit / delete", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("ied");
		docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Original body");
	});

	test("insert --after places paragraph after the locator", async () => {
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Appended",
			"--style",
			"Heading2",
			"--color",
			"CC0000",
			"--bold",
		);
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as {
			blocks: Array<{
				id: string;
				type: string;
				style?: string;
				runs?: Array<{ text: string; color?: string; bold?: boolean }>;
			}>;
		};
		const paragraphs = doc.blocks.filter((block) => block.type === "paragraph");
		expect(paragraphs[0]?.runs?.[0]?.text).toBe("Original body");
		expect(paragraphs[1]?.style).toBe("Heading2");
		expect(paragraphs[1]?.runs?.[0]).toMatchObject({
			text: "Appended",
			color: "CC0000",
			bold: true,
		});
	});

	test("insert --before places paragraph before the locator", async () => {
		await runCli("insert", docPath, "--before", "p0", "--text", "Prepended");
		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as {
			blocks: Array<{ type: string; runs?: Array<{ text: string }> }>;
		};
		const paragraphs = doc.blocks.filter((block) => block.type === "paragraph");
		expect(paragraphs[0]?.runs?.[0]?.text).toBe("Prepended");
	});

	test("insert --runs silently drops unsupported run types (round-trip safety)", async () => {
		// Simulates `docx read | jq | docx insert --runs '[...]'` where the
		// source paragraph contained an equation/footnoteRef/chart that we
		// surface in the AST but can't re-emit as fresh OOXML. Should not crash.
		const runsJson = JSON.stringify([
			{ type: "text", text: "Before " },
			{ type: "equation", text: "x_i", display: false },
			{ type: "text", text: " middle " },
			{ type: "noteRef", kind: "footnote", id: "fn1" },
			{ type: "chart", kind: "chart" },
			{ type: "text", text: " after" },
		]);
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--runs",
			runsJson,
		);
		expect(result.exitCode).toBe(0);
		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as {
			blocks: Array<{ type: string; runs?: Array<{ text?: string }> }>;
		};
		const lastParagraph = doc.blocks
			.filter((block) => block.type === "paragraph")
			.pop();
		const texts = (lastParagraph?.runs ?? [])
			.map((run) => run.text)
			.filter((text): text is string => text !== undefined);
		expect(texts.join("")).toBe("Before  middle  after");
	});

	test("insert --runs supports mixed-format paragraph", async () => {
		const runsJson = JSON.stringify([
			{ type: "text", text: "Mix: " },
			{ type: "text", text: "red", color: "CC0000" },
			{ type: "text", text: " / " },
			{ type: "text", text: "bold", bold: true },
		]);
		await runCli("insert", docPath, "--after", "p0", "--runs", runsJson);
		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as {
			blocks: Array<{
				type: string;
				runs?: Array<{ text: string; color?: string; bold?: boolean }>;
			}>;
		};
		const lastParagraph = doc.blocks
			.filter((block) => block.type === "paragraph")
			.pop();
		const runs = lastParagraph?.runs ?? [];
		expect(runs[0]?.text).toBe("Mix: ");
		expect(runs[1]).toMatchObject({ text: "red", color: "CC0000" });
		expect(runs[3]).toMatchObject({ text: "bold", bold: true });
	});

	test("edit replaces a paragraph at the locator", async () => {
		await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--text",
			"Replaced",
			"--style",
			"Heading1",
		);
		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as {
			blocks: Array<{
				type: string;
				style?: string;
				runs?: Array<{ text: string }>;
			}>;
		};
		const paragraph = doc.blocks.find((block) => block.type === "paragraph");
		expect(paragraph?.style).toBe("Heading1");
		expect(paragraph?.runs?.[0]?.text).toBe("Replaced");
	});

	test("delete removes the block at the locator", async () => {
		await runCli("insert", docPath, "--after", "p0", "--text", "Second");
		const beforeRead = await runCli("read", docPath, "--ast");
		const before = beforeRead.parsed as {
			blocks: Array<{ type: string }>;
		};
		const beforeCount = before.blocks.filter(
			(block) => block.type === "paragraph",
		).length;

		await runCli("delete", docPath, "--at", "p0");
		const afterRead = await runCli("read", docPath, "--ast");
		const after = afterRead.parsed as {
			blocks: Array<{ type: string; runs?: Array<{ text: string }> }>;
		};
		const afterCount = after.blocks.filter(
			(block) => block.type === "paragraph",
		).length;
		expect(afterCount).toBe(beforeCount - 1);
		const remaining = after.blocks.find((block) => block.type === "paragraph");
		expect(remaining?.runs?.[0]?.text).toBe("Second");
	});

	test("dry-run does not modify the file", async () => {
		const before = await Bun.file(docPath).arrayBuffer();
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Should not appear",
			"--dry-run",
		);
		const after = await Bun.file(docPath).arrayBuffer();
		expect(after.byteLength).toBe(before.byteLength);
	});

	test("--output writes to a parallel file and leaves FILE untouched", async () => {
		const beforeBytes = await Bun.file(docPath).arrayBuffer();
		const outPath = `${docPath}.copy.docx`;
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Out-of-band",
			"-o",
			outPath,
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { path: string }).path).toBe(outPath);

		const afterBytes = await Bun.file(docPath).arrayBuffer();
		expect(afterBytes.byteLength).toBe(beforeBytes.byteLength);
		expect(await Bun.file(outPath).exists()).toBe(true);

		const read = await runCli("read", outPath, "--ast");
		const doc = read.parsed as {
			blocks: Array<{ type: string; runs?: Array<{ text: string }> }>;
		};
		const paragraphs = doc.blocks.filter((block) => block.type === "paragraph");
		expect(paragraphs[1]?.runs?.[0]?.text).toBe("Out-of-band");
	});

	test("--dry-run with --output writes nothing and echoes the intended output", async () => {
		const sourceBytes = await Bun.file(docPath).arrayBuffer();
		const outPath = `${docPath}.copy.docx`;
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Should not appear",
			"-o",
			outPath,
			"--dry-run",
		);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			dryRun: true,
			path: docPath,
			output: outPath,
		});
		const afterBytes = await Bun.file(docPath).arrayBuffer();
		expect(afterBytes.byteLength).toBe(sourceBytes.byteLength);
		expect(await Bun.file(outPath).exists()).toBe(false);
	});

	test("invalid locator returns block-not-found", async () => {
		const result = await runCli("edit", docPath, "--at", "p99", "--text", "x");
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({
			code: "BLOCK_NOT_FOUND",
		});
	});
});

describe("docx insert --page-break / --column-break", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("breaks");
		docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "First");
	});

	test("--page-break inserts a paragraph with a single page break run", async () => {
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--page-break",
		);
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as {
			blocks: Array<{
				id: string;
				type: string;
				runs?: Array<{ type: string; kind?: string }>;
			}>;
		};
		const paragraphs = doc.blocks.filter((block) => block.type === "paragraph");
		expect(paragraphs[1]?.runs).toEqual([{ type: "break", kind: "page" }]);
	});

	test("--column-break inserts a paragraph with a single column break run", async () => {
		await runCli("insert", docPath, "--before", "p0", "--column-break");
		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as {
			blocks: Array<{
				id: string;
				type: string;
				runs?: Array<{ type: string; kind?: string }>;
			}>;
		};
		const paragraphs = doc.blocks.filter((block) => block.type === "paragraph");
		expect(paragraphs[0]?.runs).toEqual([{ type: "break", kind: "column" }]);
	});

	test("rejects --page-break alongside --text", async () => {
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--page-break",
			"--text",
			"x",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({
			code: "USAGE",
		});
	});

	test("rejects --page-break alongside --column-break", async () => {
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--page-break",
			"--column-break",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({
			code: "USAGE",
		});
	});

	test("requires content flag", async () => {
		const result = await runCli("insert", docPath, "--after", "p0");
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({
			code: "USAGE",
		});
	});
});

/**
 * Ergonomics fix surfaced by the weak-model adversarial run: `--text` with
 * embedded newlines/tabs → real `<w:br/>` / `<w:tab/>` (not a literal \n that
 * Word swallows). Verse/addresses stay line-per-line, and `read` round-trips
 * them. (Image captions moved to images.test.ts with the `docx images add` verb.)
 */

type Run = { type: string; text?: string; kind?: string };
type Block = { id: string; type: string; style?: string; runs?: Run[] };

async function blocks(path: string): Promise<Block[]> {
	const result = await runCli("read", path, "--ast");
	return (result.parsed as { blocks: Block[] }).blocks;
}

async function block(path: string, id: string): Promise<Block> {
	const found = (await blocks(path)).find((candidate) => candidate.id === id);
	if (!found) throw new Error(`block ${id} not found`);
	return found;
}

function newDoc(label: string): string {
	return join(tempWorkspace(label), "doc.docx");
}

describe("--text newlines and tabs", () => {
	test("insert --text with a newline becomes a <w:br/> line break", async () => {
		const path = newDoc("nl-insert");
		await runCli("create", path, "--text", "Intro.");
		await runCli(
			"insert",
			path,
			"--after",
			"p0",
			"--text",
			"line one\nline two",
		);
		const runs = (await block(path, "p1")).runs ?? [];
		expect(runs.map((run) => run.type)).toEqual(["text", "break", "text"]);
		expect(runs[0]?.text).toBe("line one");
		expect(runs[1]?.kind).toBe("line");
		expect(runs[2]?.text).toBe("line two");
	});

	test("insert --text with a tab becomes a <w:tab/>", async () => {
		const path = newDoc("tab-insert");
		await runCli("create", path, "--text", "Intro.");
		await runCli("insert", path, "--after", "p0", "--text", "a\tb");
		const runs = (await block(path, "p1")).runs ?? [];
		expect(runs.map((run) => run.type)).toEqual(["text", "tab", "text"]);
	});

	test("edit --text (whole paragraph) splits newlines too", async () => {
		const path = newDoc("nl-edit");
		await runCli("create", path, "--text", "placeholder");
		await runCli("edit", path, "--at", "p0", "--text", "first\nsecond\nthird");
		const runs = (await block(path, "p0")).runs ?? [];
		expect(runs.filter((run) => run.type === "break").length).toBe(2);
		expect(
			runs.filter((run) => run.type === "text").map((run) => run.text),
		).toEqual(["first", "second", "third"]);
	});

	test("a multi-line --text paragraph round-trips through read → markdown", async () => {
		const path = newDoc("nl-roundtrip");
		await runCli("create", path, "--text", "Roses are red\nViolets are blue");
		const md = (await runCli("read", path)).stdout;
		expect(md).toContain("Roses are red\nViolets are blue");
	});

	test("single-line --text is still one text run (no behavior change)", async () => {
		const path = newDoc("nl-single");
		await runCli("create", path, "--text", "just one line");
		const runs = (await block(path, "p0")).runs ?? [];
		expect(runs).toHaveLength(1);
		expect(runs[0]?.type).toBe("text");
	});
});

// Regression: spacing/indent flags used to be silently dropped on several insert
// content kinds (exit 0, no effect) — the weak-agent footgun. They must either
// take effect or be rejected up front (markdown: the source owns block layout).
// (Code, equations, and images moved to their own noun-verb commands; their
// spacing threading is covered in code.test.ts / equations.test.ts /
// images.test.ts.)
describe("insert — spacing/indent across content kinds (no silent drop)", () => {
	async function withAnchor(label: string): Promise<string> {
		const path = newDoc(label);
		await runCli("create", path, "--text", "Anchor.");
		return path;
	}

	test("--markdown rejects spacing/indent flags (the source owns block layout)", async () => {
		const path = await withAnchor("ins-md-reject");
		const result = await runCli(
			"insert",
			path,
			"--after",
			"p0",
			"--markdown",
			"A new paragraph.",
			"--space-after",
			"12",
		);
		expect(result.exitCode).not.toBe(0);
		expect((result.parsed as { error?: string }).error).toContain(
			"can't be combined with --markdown",
		);
	});
});

async function styleIds(docPath: string): Promise<string[]> {
	const pkg = await Pkg.open(docPath);
	if (!pkg.hasPart("word/styles.xml")) return [];
	const xml = await pkg.readText("word/styles.xml");
	return [...xml.matchAll(/w:styleId="([^"]+)"/g)].map((m) => m[1] ?? "");
}

describe("insert/edit --style provisioning", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("style-prov");
		docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Body");
	});

	test("insert --style Heading2 defines Heading2 (and Normal) in styles.xml", async () => {
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"A heading",
			"--style",
			"Heading2",
		);
		const ids = await styleIds(docPath);
		expect(ids).toContain("Heading2");
		expect(ids).toContain("Normal");
	});

	test("edit --style Quote defines Quote without dropping existing styles", async () => {
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"A heading",
			"--style",
			"Heading2",
		);
		await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--text",
			"Quoted",
			"--style",
			"Quote",
		);
		const ids = await styleIds(docPath);
		expect(ids).toContain("Quote");
		expect(ids).toContain("Heading2");
	});

	test("a custom (non-baseline) style is referenced but not defined", async () => {
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Custom",
			"--style",
			"MyCorporateStyle",
		);
		const ids = await styleIds(docPath);
		expect(ids).not.toContain("MyCorporateStyle");
		// The pStyle reference is still written even though the style is undefined.
		const pkg = await Pkg.open(docPath);
		const documentXml = await pkg.readText("word/document.xml");
		expect(documentXml).toContain('w:val="MyCorporateStyle"');
	});

	test("insert without --style adds no style definitions", async () => {
		await runCli("insert", docPath, "--after", "p0", "--text", "Plain");
		// `docx create` ships a styles.xml with only Normal; a plain insert
		// shouldn't add anything.
		expect(await styleIds(docPath)).toEqual(["Normal"]);
	});
});

// track-flag covers edit/replace/delete; insert's per-invocation --track (force
// one insertion tracked while the doc toggle is OFF) is its own code path.
describe("docx insert — --track forces tracking with the toggle off", () => {
	test("--track wraps the inserted runs in a tracked insertion", async () => {
		const path = newDoc("insert-track");
		await runCli("create", path, "--text", "alpha");
		await runCli("insert", path, "--after", "p0", "--text", "beta", "--track");
		expect(await trackedKinds(path)).toContain("ins");
	});

	test("no --track on an untracked doc records nothing", async () => {
		const path = newDoc("insert-track-control");
		await runCli("create", path, "--text", "alpha");
		await runCli("insert", path, "--after", "p0", "--text", "beta");
		expect(await trackedKinds(path)).toHaveLength(0);
	});
});

describe("insert --text-file (literal, parser-free)", () => {
	let docPath: string;

	beforeEach(async () => {
		docPath = newDoc("insert-literal");
		await runCli("create", docPath, "--text", "SEED");
	});

	test("inserts literal multi-paragraph text without GFM parsing", async () => {
		const notes = join(tempWorkspace("literal-src"), "notes.txt");
		// Every line is content GFM would corrupt — an ordered-list marker, emphasis
		// punctuation, a bare URL, CriticMarkup. Literal mode keeps them verbatim.
		await Bun.write(
			notes,
			"3. Reviewer 1 notes the issue\n*not italic* and _also not_\nSee https://example.com here\nCost {++5++} dollars",
		);

		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text-file",
			notes,
		);
		expect(result.exitCode).toBe(0);
		// 4 source lines → 4 new paragraphs, in order.
		expect((result.parsed as { locators: string[] }).locators).toEqual([
			"p1",
			"p2",
			"p3",
			"p4",
		]);

		const paragraphs = await readParagraphs(docPath);
		expect(paragraphText(paragraphs[1])).toBe("3. Reviewer 1 notes the issue");
		expect(paragraphText(paragraphs[2])).toBe("*not italic* and _also not_");
		expect(paragraphText(paragraphs[3])).toBe("See https://example.com here");
		expect(paragraphText(paragraphs[4])).toBe("Cost {++5++} dollars");
		// No run became a hyperlink or CriticMarkup ins/del — every run is plain text.
		for (const paragraph of paragraphs.slice(1)) {
			for (const run of paragraph.runs ?? []) expect(run.type).toBe("text");
		}
		// And the raw XML carries no list numbering or hyperlink that GFM would mint.
		const xml = await (await Pkg.open(docPath)).readText("word/document.xml");
		expect(xml).not.toContain("<w:hyperlink");
		expect(xml).not.toContain("<w:numPr");
	});

	test("each newline is a paragraph; blank lines become empty paragraphs", async () => {
		const notes = join(tempWorkspace("literal-blank"), "n.txt");
		// Interior blank line → an empty paragraph; the trailing newline must NOT
		// mint a stray trailing paragraph.
		await Bun.write(notes, "alpha\n\nbravo\n");

		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text-file",
			notes,
		);
		expect((result.parsed as { locators: string[] }).locators).toEqual([
			"p1",
			"p2",
			"p3",
		]);

		const paragraphs = await readParagraphs(docPath);
		expect(paragraphText(paragraphs[1])).toBe("alpha");
		expect(paragraphText(paragraphs[2])).toBe(""); // the blank line
		expect(paragraphText(paragraphs[3])).toBe("bravo");
	});
});

describe("insert --at-start / --at-end (boundary placement)", () => {
	let docPath: string;

	beforeEach(async () => {
		docPath = newDoc("insert-boundary");
		await runCli("create", docPath, "--text", "MIDDLE");
	});

	test("--at-start prepends; --at-end appends; both need no locator", async () => {
		const top = await runCli(
			"insert",
			docPath,
			"--at-start",
			"--text",
			"TOP",
			"--style",
			"Title",
		);
		expect(top.exitCode).toBe(0);
		expect((top.parsed as { locators: string[] }).locators).toEqual(["p0"]);

		const bottom = await runCli(
			"insert",
			docPath,
			"--at-end",
			"--text",
			"BOTTOM",
		);
		expect(bottom.exitCode).toBe(0);

		const paragraphs = await readParagraphs(docPath);
		expect(paragraphText(paragraphs[0])).toBe("TOP");
		expect(paragraphs[0]?.style).toBe("Title"); // boundary insert still styles
		expect(paragraphText(paragraphs[1])).toBe("MIDDLE");
		expect(paragraphText(paragraphs[2])).toBe("BOTTOM");
	});

	test("--at-start and --at-end are mutually exclusive", async () => {
		const result = await runCli(
			"insert",
			docPath,
			"--at-start",
			"--at-end",
			"--text",
			"x",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});

	test("--at-start is rejected inside --batch (no boundary anchors there)", async () => {
		const batch = join(tempWorkspace("boundary-batch"), "b.jsonl");
		await Bun.write(batch, '{"at-start":true,"text":"x"}\n');
		const result = await runCli("insert", docPath, "--batch", batch);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { error: string }).error).toContain(
			"aren't supported in --batch",
		);
	});

	// Regression for the table-cell anchoring bug: blockReferences holds cell
	// paragraphs (tag w:p) registered BEFORE their owning table, so a tag-only
	// boundary scan would pick the first CELL paragraph on a table-first doc and
	// splice INSIDE the cell. --at-start must anchor at the BODY top.
	test("--at-start on a table-first doc anchors at the body top, not inside the first cell", async () => {
		const docPath = newDoc("boundary-table-first");
		await runCli("create", docPath, "--text", "TAIL");
		// Push a table above p0 so the document now BEGINS with a table.
		await runCli(
			"tables",
			"create",
			docPath,
			"--before",
			"p0",
			"--rows",
			"2",
			"--cols",
			"2",
		);

		const result = await runCli(
			"insert",
			docPath,
			"--at-start",
			"--text",
			"TOP",
		);
		expect(result.exitCode).toBe(0);
		// A TOP-LEVEL paragraph locator (p0) — NOT a cell locator (t0:r0c0:pN).
		expect((result.parsed as { locators: string[] }).locators).toEqual(["p0"]);

		const paragraphs = await readParagraphs(docPath);
		expect(paragraphText(paragraphs[0])).toBe("TOP"); // first block in the body
	});

	test("--at-end on a table-last doc anchors at the body end, not inside the last cell", async () => {
		const docPath = newDoc("boundary-table-last");
		await runCli("create", docPath, "--text", "HEAD");
		await runCli(
			"tables",
			"create",
			docPath,
			"--after",
			"p0",
			"--rows",
			"2",
			"--cols",
			"2",
		);

		const result = await runCli(
			"insert",
			docPath,
			"--at-end",
			"--text",
			"BOTTOM",
		);
		expect(result.exitCode).toBe(0);
		const locators = (result.parsed as { locators: string[] }).locators;
		// Top-level paragraph after the table — no cell-scoped (":") locator.
		expect(locators).toEqual(["p1"]);
		expect(locators.some((l) => l.includes(":"))).toBe(false);
	});
});

describe("insert into a bare table-cell locator", () => {
	test("--at fills the mandatory empty paragraph and reports canonical p0", async () => {
		const docPath = await newTableDoc("cell-at-empty");
		const result = await runCli(
			"insert",
			docPath,
			"--at",
			"t0:r0c0",
			"--text",
			"Dana Okafor",
		);

		expect(result.exitCode).toBe(0);
		expect((result.parsed as { locators: string[] }).locators).toEqual([
			"t0:r0c0:p0",
		]);
		const read = await runCli("read", docPath);
		expect(read.stdout).toContain("Dana Okafor <!-- t0:r0c0:p0 -->");
		expect(read.stdout).not.toContain("<!-- t0:r0c0:p0 --><br>Dana Okafor");
	});

	test("inherits run formatting from an empty cell's paragraph mark", async () => {
		const docPath = await newTableDoc("cell-at-styled");
		expect(
			(
				await runCli(
					"raw",
					"replace",
					docPath,
					"--at",
					"t0:r0c0:p0",
					"--xml",
					'<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="18"/></w:rPr></w:pPr></w:p>',
				)
			).exitCode,
		).toBe(0);
		await runCli("insert", docPath, "--at", "t0:r0c0", "--text", "Styled");

		const read = await runCli("read", docPath, "--ast");
		const table = (
			read.parsed as {
				blocks: Array<{
					type: string;
					rows?: Array<{
						cells: Array<{
							blocks: Array<{
								runs?: Array<{
									text?: string;
									font?: string;
									sizeHalfPoints?: number;
								}>;
							}>;
						}>;
					}>;
				}>;
			}
		).blocks.find((candidate) => candidate.type === "table");
		const run = table?.rows?.[0]?.cells[0]?.blocks[0]?.runs?.[0];
		expect(run).toMatchObject({
			text: "Styled",
			font: "Arial",
			sizeHalfPoints: 18,
		});
	});

	test("--before, --at, and --after use stable cell start/end boundaries", async () => {
		const docPath = await newTableDoc("cell-boundaries");
		await runCli("insert", docPath, "--at", "t0:r0c0", "--text", "A");
		await runCli("insert", docPath, "--at", "t0:r0c0", "--text", "B");
		await runCli("insert", docPath, "--before", "t0:r0c0", "--text", "C");
		await runCli("insert", docPath, "--after", "t0:r0c0", "--text", "D");

		const read = await runCli("read", docPath);
		const cell = read.stdout.match(/\| C [^|]+\|/)?.[0] ?? "";
		expect(cell).toContain("C <!-- t0:r0c0:p0 -->");
		expect(cell).toContain("A <!-- t0:r0c0:p1 -->");
		expect(cell).toContain("B <!-- t0:r0c0:p2 -->");
		expect(cell).toContain("D <!-- t0:r0c0:p3 -->");
	});

	test("tracked empty-cell fill rejects back to one valid blank paragraph", async () => {
		const docPath = await newTableDoc("cell-track-reject");
		expect(
			(
				await runCli(
					"insert",
					docPath,
					"--at",
					"t0:r0c0",
					"--text",
					"Tracked",
					"--track",
				)
			).exitCode,
		).toBe(0);
		expect(await trackedKinds(docPath)).toEqual(["ins"]);

		expect(
			(await runCli("track-changes", "reject", docPath, "--all")).exitCode,
		).toBe(0);
		const read = await runCli("read", docPath);
		expect(read.stdout).toContain("<!-- t0:r0c0 -->");
		expect(read.stdout).not.toContain("Tracked");
	});

	test("--at appends after an ordinary block and rejects merged bare cells", async () => {
		const docPath = await newTableDoc("cell-rejections");
		const blockTarget = await runCli(
			"insert",
			docPath,
			"--at",
			"p0",
			"--text",
			"After block",
		);
		expect(blockTarget.exitCode).toBe(0);
		expect((blockTarget.parsed as { placement: string }).placement).toBe("at");
		expect((blockTarget.parsed as { locators: string[] }).locators).toEqual([
			"p1",
		]);
		expect(paragraphText((await readParagraphs(docPath))[1])).toBe(
			"After block",
		);

		expect(
			(await runCli("tables", "merge", docPath, "--at", "t0:r0c0-r0c1"))
				.exitCode,
		).toBe(0);
		const mergedTarget = await runCli(
			"insert",
			docPath,
			"--at",
			"t0:r0c0",
			"--text",
			"Nope",
		);
		expect(mergedTarget.exitCode).not.toBe(0);
		expect((mergedTarget.parsed as { code: string }).code).toBe(
			"TABLE_STRUCTURE",
		);
	});
});

describe("docx insert — paragraph spacing & indentation", () => {
	test("inserts a paragraph carrying spacing + indentation", async () => {
		const workspace = tempWorkspace("insert-spacing");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "First.");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Spaced and indented.",
			"--space-after",
			"6",
			"--line-spacing",
			"1.5",
			"--indent-left",
			"0.5in",
		);
		expect(result.exitCode).toBe(0);
		const p1 = (await readParagraphs(docPath)).find((p) => p.id === "p1");
		expect(p1?.spacing).toEqual({ after: 120, line: 360, lineRule: "auto" });
		expect(p1?.indent).toEqual({ left: 720 });
	});

	test("rejects --first-line together with --hanging", async () => {
		const workspace = tempWorkspace("insert-mutex");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "First.");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"x",
			"--first-line",
			"0.5in",
			"--hanging",
			"0.25in",
		);
		expect(result.exitCode).not.toBe(0);
		expect((result.parsed as { error?: string }).error).toContain(
			"mutually exclusive",
		);
	});
});

describe("docx insert — inline escape decoding (--text / --markdown)", () => {
	// `"a\\nb"` in this TS source IS the literal a,\,n,b the shell delivers when an
	// agent writes `--text "a\nb"`. The inline argv ingress decodes it so the break
	// lands as `<w:br/>`, matching the `--text-file` / batch channels.
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("insert-esc");
		docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Original body");
	});

	test("--text decodes \\n to a line break and \\t to a tab", async () => {
		const result = await runCli(
			"insert",
			docPath,
			"--at-end",
			"--text",
			"a\\nb\\tc",
		);
		expect(result.exitCode).toBe(0);
		const paragraphs = await readParagraphs(docPath);
		const last = paragraphs[paragraphs.length - 1];
		expect((last?.runs ?? []).map((run) => run.type)).toEqual([
			"text",
			"break",
			"text",
			"tab",
			"text",
		]);
		expect(paragraphText(last)).toBe("abc");
	});

	test("--markdown decodes \\n\\n into separate blocks", async () => {
		const result = await runCli(
			"insert",
			docPath,
			"--at-end",
			"--markdown",
			"## Heading\\n\\nA paragraph.",
		);
		expect(result.exitCode).toBe(0);
		const paragraphs = await readParagraphs(docPath);
		const heading = paragraphs.find((p) => p.style === "Heading2");
		expect(paragraphText(heading)).toBe("Heading");
		expect(paragraphs.some((p) => paragraphText(p) === "A paragraph.")).toBe(
			true,
		);
	});
});

// Contextual `--help`: `--text`/`--runs` route to a focused slice (same shim as
// edit's `pickHelp`), so a weak agent that types a content flag then `--help`
// lands on the guidance for that flag rather than the full default screen.
describe("docx insert — contextual --help", () => {
	test("--text --help focuses on formatting a new paragraph (mentions --bold and --markdown)", async () => {
		const def = (await runCli("insert", "--help")).stdout;
		const result = await runCli("insert", "--text", "--help");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("--bold");
		expect(result.stdout).toContain("--markdown");
		expect(result.stdout).not.toBe(def);
	});

	test("--runs --help exposes the Run[] JSON", async () => {
		const result = await runCli("insert", "--runs", "--help");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("Run[]");
		expect(result.stdout).toContain("vertAlign");
	});

	test("default --help leads with --markdown and defers the per-run JSON to --runs --help", async () => {
		const result = await runCli("insert", "--help");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("--markdown");
		// The default must NOT dump the full per-run JSON schema…
		expect(result.stdout).not.toContain("vertAlign");
		// …it points at the variant that does.
		expect(result.stdout).toContain("--runs --help");
	});
});
