import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";
import { readDocumentXml, trackedKinds } from "./helpers";

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
 * Two ergonomics fixes surfaced by the weak-model adversarial run:
 *   - `--text` with embedded newlines/tabs → real `<w:br/>` / `<w:tab/>` (not a
 *     literal \n that Word swallows). Verse/addresses stay line-per-line, and
 *     `read` round-trips them.
 *   - `insert --image --caption` / `images add --caption` → a native Word
 *     "Caption"-styled paragraph under the figure (Table-of-Figures-able).
 */

type Run = { type: string; text?: string; kind?: string };
type Block = { id: string; type: string; style?: string; runs?: Run[] };

const ASSETS = join(import.meta.dir, "..", "fixtures", "assets");
const PNG = join(ASSETS, "sample.png");

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

describe("image captions", () => {
	test("insert --image --caption adds a Caption-styled paragraph below the figure", async () => {
		const path = newDoc("cap-insert");
		await runCli("create", path, "--text", "Report.");
		const result = await runCli(
			"insert",
			path,
			"--after",
			"p0",
			"--image",
			PNG,
			"--caption",
			"Figure 1: Quarterly revenue",
		);
		expect(result.exitCode).toBe(0);
		// Two blocks minted: the figure paragraph and the caption paragraph.
		const all = await blocks(path);
		const caption = all.find((b) => b.style === "Caption");
		expect(caption).toBeDefined();
		expect((caption?.runs ?? []).map((run) => run.text).join("")).toBe(
			"Figure 1: Quarterly revenue",
		);
	});

	test("the Caption style is provisioned in styles.xml", async () => {
		const path = newDoc("cap-style");
		await runCli("create", path, "--text", "Report.");
		await runCli(
			"insert",
			path,
			"--after",
			"p0",
			"--image",
			PNG,
			"--caption",
			"Fig 1",
		);
		// styles.xml provisioning is exercised by the doc opening cleanly here.
		expect((await runCli("read", path, "--ast")).exitCode).toBe(0);
		// The caption paragraph carries the Caption pStyle (proves the style was
		// referenced).
		const caption = (await blocks(path)).find((b) => b.style === "Caption");
		expect(caption).toBeDefined();
	});

	test("images add --caption (alias) works the same way", async () => {
		const path = newDoc("cap-alias");
		await runCli("create", path, "--text", "Report.");
		const result = await runCli(
			"images",
			"add",
			path,
			"--image",
			PNG,
			"--after",
			"p0",
			"--caption",
			"Figure A",
		);
		expect(result.exitCode).toBe(0);
		const caption = (await blocks(path)).find((b) => b.style === "Caption");
		expect((caption?.runs ?? []).map((run) => run.text).join("")).toBe(
			"Figure A",
		);
	});

	test("no --caption → no Caption paragraph", async () => {
		const path = newDoc("cap-none");
		await runCli("create", path, "--text", "Report.");
		await runCli("insert", path, "--after", "p0", "--image", PNG);
		const caption = (await blocks(path)).find((b) => b.style === "Caption");
		expect(caption).toBeUndefined();
	});
});

// Regression: spacing/indent flags used to be silently dropped on several insert
// content kinds (exit 0, no effect) — the weak-agent footgun. They must either
// take effect (code/equation/image: meaningful, single/uniform paragraphs) or be
// rejected up front (markdown: the source owns block layout).
describe("insert — spacing/indent across content kinds (no silent drop)", () => {
	async function withAnchor(label: string): Promise<string> {
		const path = newDoc(label);
		await runCli("create", path, "--text", "Anchor.");
		return path;
	}

	test("--code threads spacing/indent onto every code paragraph", async () => {
		const path = await withAnchor("ins-code-spacing");
		expect(
			(
				await runCli(
					"insert",
					path,
					"--after",
					"p0",
					"--code",
					"a = 1\nb = 2",
					"--language",
					"python",
					"--space-after",
					"12",
					"--indent-left",
					"0.5",
				)
			).exitCode,
		).toBe(0);
		for (const id of ["p1", "p2"]) {
			const paragraph = (await readParagraphs(path)).find((b) => b.id === id);
			expect(paragraph?.spacing).toEqual({ after: 240 });
			expect(paragraph?.indent).toEqual({ left: 720 });
		}
	});

	test("--equation threads spacing/indent onto the equation paragraph", async () => {
		const path = await withAnchor("ins-eq-spacing");
		await runCli(
			"insert",
			path,
			"--after",
			"p0",
			"--equation",
			"x = y",
			"--space-before",
			"12",
			"--indent-left",
			"0.5",
		);
		const xml = await readDocumentXml(path);
		expect(xml).toContain('<w:spacing w:before="240"/>');
		expect(xml).toContain('<w:ind w:left="720"/>');
	});

	test("--image threads spacing onto the figure paragraph", async () => {
		const path = await withAnchor("ins-img-spacing");
		await runCli(
			"insert",
			path,
			"--after",
			"p0",
			"--image",
			PNG,
			"--width",
			"1",
			"--space-after",
			"12",
		);
		const figure = (await readParagraphs(path)).find((b) => b.id === "p1");
		expect(figure?.spacing).toEqual({ after: 240 });
	});

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

type CodeBlockAst = {
	id: string;
	type: string;
	style?: string;
	runs?: Array<{
		type: string;
		text?: string;
		runStyle?: string;
		color?: string;
		bold?: boolean;
	}>;
};

async function blocksOf(docPath: string): Promise<CodeBlockAst[]> {
	const result = await runCli("read", docPath, "--ast");
	return (result.parsed as { blocks: CodeBlockAst[] }).blocks;
}

describe("docx insert --code", () => {
	test("inline --code: each \\n becomes its own CodeBlock paragraph", async () => {
		const docPath = join(tempWorkspace("code-inline"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--code",
			"line one\nline two\nline three",
		);

		const blocks = await blocksOf(docPath);
		// 1 intro + 3 code lines + 1 sectionBreak.
		expect(blocks).toHaveLength(5);
		const codeBlocks = blocks.filter((b) => b.style === "CodeBlock");
		expect(codeBlocks).toHaveLength(3);
		const texts = codeBlocks.map(
			(b) => b.runs?.map((r) => r.text ?? "").join("") ?? "",
		);
		expect(texts).toEqual(["line one", "line two", "line three"]);
		// Every run carries the Code character style — defensive against Word
		// versions that don't cascade pStyle's font through to runs.
		for (const b of codeBlocks) {
			for (const r of b.runs ?? []) expect(r.runStyle).toBe("Code");
		}
	});

	test("--code-file PATH: reads content from the file", async () => {
		const workspace = tempWorkspace("code-file");
		const docPath = join(workspace, "out.docx");
		const snippetPath = join(workspace, "snippet.py");
		await Bun.write(snippetPath, "def hello():\n    return 42\n");

		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--code-file",
			snippetPath,
		);

		const blocks = await blocksOf(docPath);
		const codeBlocks = blocks.filter((b) => b.style === "CodeBlock");
		// Trailing newline becomes a (visible-but-empty) fourth paragraph; this
		// matches "what was in the file" semantically, and the markdown render
		// collapses it back to a clean fenced block.
		expect(codeBlocks.length).toBeGreaterThanOrEqual(2);
		const joined = codeBlocks
			.map((b) => b.runs?.map((r) => r.text ?? "").join("") ?? "")
			.join("\n");
		expect(joined).toContain("def hello():");
		expect(joined).toContain("    return 42");
	});

	// `--code-file -` (stdin) is exercised only by manual review and by an
	// agent running the real binary — the in-process harness shares the test
	// runner's stdin so attempting to consume it never returns EOF. The
	// `--code-file PATH` test above exercises the file-resolution flow; the
	// stdin branch is a one-line conditional that swaps `Bun.file(path)` for
	// `Bun.stdin.stream()`.

	test("--language typescript applies token colors", async () => {
		const docPath = join(tempWorkspace("code-ts"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--code",
			"function foo() { return 42; }",
			"--language",
			"typescript",
		);

		const blocks = await blocksOf(docPath);
		const codeBlock = blocks.find((b) => b.style?.startsWith("CodeBlock"));
		expect(codeBlock).toBeDefined();
		const runs = codeBlock?.runs ?? [];
		// At least the `function` and `return` keyword tokens should pick up
		// the keyword color (CF222E in our palette).
		const colored = runs.filter((r) => r.color);
		expect(colored.length).toBeGreaterThan(0);
		const keywordRun = runs.find((r) => r.text === "function");
		expect(keywordRun?.color).toBe("CF222E");
	});

	test("unknown --language degrades to uncolored (block still inserts)", async () => {
		const docPath = join(tempWorkspace("code-unknown-lang"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--code",
			"sub foo { 42 }",
			"--language",
			"klingon-script", // not a real lowlight language
		);
		expect(result.exitCode).toBe(0);
		const blocks = await blocksOf(docPath);
		const codeBlock = blocks.find((b) => b.style?.startsWith("CodeBlock"));
		expect(codeBlock).toBeDefined();
		// Single uncolored run when language is unrecognized.
		const runs = codeBlock?.runs ?? [];
		expect(runs.every((r) => r.color === undefined)).toBe(true);
	});

	test("provisions Code AND CodeBlock styles in styles.xml", async () => {
		const docPath = join(tempWorkspace("code-styles"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--code",
			"plain text\nsecond line",
		);

		const pkg = await Pkg.open(docPath);
		const stylesXml = await pkg.readText("word/styles.xml");
		expect(stylesXml).toContain('w:styleId="Code"');
		expect(stylesXml).toContain('w:styleId="CodeBlock"');
	});

	test("--language LANG provisions a CodeBlock-LANG style basedOn CodeBlock", async () => {
		const docPath = join(tempWorkspace("code-lang-style"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--code",
			"function foo() {}",
			"--language",
			"typescript",
		);

		// Each emitted paragraph carries the per-language pStyle, so the
		// language survives round-trip through Word/LibreOffice (basedOn
		// preserves the rendering).
		const blocks = await blocksOf(docPath);
		const codeBlock = blocks.find((b) => b.type === "paragraph" && b.style);
		expect(codeBlock?.style).toBe("CodeBlock-typescript");

		// And the derived style is provisioned in styles.xml, basedOn CodeBlock.
		const pkg = await Pkg.open(docPath);
		const stylesXml = await pkg.readText("word/styles.xml");
		expect(stylesXml).toContain('w:styleId="CodeBlock-typescript"');
		expect(stylesXml).toMatch(
			/w:styleId="CodeBlock-typescript"[\s\S]*?<w:basedOn w:val="CodeBlock"\/>/,
		);
	});

	test("an unknown language still gets a derived style (degraded, not lost)", async () => {
		// `klingon-script` isn't in lowlight's grammar list, so token colors
		// drop to plain. But the language NAME survives via the pStyle suffix
		// so `read --markdown` still tags the fenced block with the language.
		const docPath = join(tempWorkspace("code-lang-unknown"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--code",
			"sub foo { 42 }",
			"--language",
			"klingon-script",
		);
		const blocks = await blocksOf(docPath);
		const codeBlock = blocks.find((b) => b.type === "paragraph" && b.style);
		expect(codeBlock?.style).toBe("CodeBlock-klingon-script");
	});

	test("--language requires --code OR --code-file (orphan check)", async () => {
		const docPath = join(tempWorkspace("code-orphan"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		// --language without any --code/--code-file is an orphan sub-flag.
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Just text",
			"--language",
			"python",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});

	test("--code and --code-file are mutually exclusive", async () => {
		const docPath = join(tempWorkspace("code-conflict"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--code",
			"foo",
			"--code-file",
			"some.py",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});
});

describe("docx read --markdown for code blocks", () => {
	test("consecutive CodeBlock paragraphs collapse into one fenced block", async () => {
		const docPath = join(tempWorkspace("code-render"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--code",
			"function foo() {\n  return 42;\n}",
		);

		const result = await runCli("read", docPath);
		const md = result.stdout;
		// Three lines wrapped in one fenced block. Locator comments sit on their
		// OWN lines bracketing the fence — never glued to a fence line, which
		// would break a downstream markdown parser (CommonMark/markdown-it).
		expect(md).toContain("function foo() {");
		expect(md).toContain("  return 42;");
		expect(md).toContain("}");
		expect(md).toContain("<!-- p1 -->");
		expect(md).toContain("<!-- p3 -->");
		// The fence lines themselves are clean: no comment glued to an opening
		// info string, no content after a closing fence.
		expect(md).toMatch(/<!-- p1 -->\n```\n/); // open locator, then a bare fence
		expect(md).toMatch(/\n```\n<!-- p3 -->/); // bare close fence, then locator
		expect(md).not.toMatch(/```[^\n]*<!--/); // never a comment on a fence line
	});

	test("--language LANG round-trips to a tagged GFM fence", async () => {
		const docPath = join(tempWorkspace("code-render-lang"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--code",
			"function foo() {\n  return 42;\n}",
			"--language",
			"typescript",
		);
		const md = (await runCli("read", docPath)).stdout;
		// Fence opens with the language tag recovered from the pStyle suffix —
		// on a CLEAN info-string line (no locator glued to it, which would make a
		// parser read the language as `typescript<!--` and lose highlighting).
		expect(md).toContain("```typescript\n");
		expect(md).not.toContain("```typescript<!--");
		expect(md).toMatch(/<!-- p1 -->\n```typescript/);
		expect(md).toContain("function foo() {");
	});

	test("read-output code fence re-imports with its language intact (parse-validity)", async () => {
		// The regression that the string-shape assertions above missed: emitting
		// the locator ON the fence line corrupts the language id, so re-parsing
		// `read --markdown` output (what a downstream tool — or our own importer —
		// does) recovers the WRONG language. Round-trip read→import→read and
		// assert the fenced language survives.
		const docPath = join(tempWorkspace("code-fence-reparse"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--code",
			"const x = 1;\nconst y = 2;",
			"--language",
			"typescript",
		);

		// Feed read --markdown back through create --from, then read again.
		const ws = tempWorkspace("code-fence-reparse-rt");
		const mdPath = join(ws, "doc.md");
		await Bun.write(mdPath, (await runCli("read", docPath)).stdout);
		const dst = join(ws, "rt.docx");
		await runCli("create", dst, "--from", mdPath);

		const reread = (await runCli("read", dst)).stdout;
		// The language tag must survive the round-trip — a corrupted info string
		// would re-import as a bare/garbled CodeBlock and lose it.
		expect(reread).toContain("```typescript\n");
		expect(reread).toContain("const x = 1;");
		const blocks = (
			(await runCli("read", dst, "--ast")).parsed as {
				blocks: Array<{ style?: string }>;
			}
		).blocks;
		expect(blocks.some((b) => b.style === "CodeBlock-typescript")).toBe(true);
	});

	test("token colors are stripped from the fenced rendering (source survives)", async () => {
		const docPath = join(tempWorkspace("code-render-colors"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--code",
			"return 42;",
			"--language",
			"typescript",
		);

		const md = (await runCli("read", docPath)).stdout;
		// No HTML span / color leakage in the fenced rendering — just literal
		// source text between the fences.
		expect(md).not.toMatch(/<span style="color/);
		expect(md).toContain("return 42;");
	});

	test("inline runStyle: Code emits backticks in the rendered markdown", async () => {
		const docPath = join(tempWorkspace("code-inline-render"), "out.docx");
		await runCli("create", docPath, "--text", "Plain.");
		await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--runs",
			JSON.stringify([
				{ type: "text", text: "use " },
				{ type: "text", text: "foo()", runStyle: "Code" },
				{ type: "text", text: " here" },
			]),
		);

		const md = (await runCli("read", docPath)).stdout;
		expect(md).toContain("use `foo()` here");
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
			"insert",
			docPath,
			"--before",
			"p0",
			"--table",
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
			"insert",
			docPath,
			"--after",
			"p0",
			"--table",
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
