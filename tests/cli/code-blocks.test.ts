import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "../../src/core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";

type Block = {
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

async function blocksOf(docPath: string): Promise<Block[]> {
	const result = await runCli("read", docPath, "--ast");
	return (result.parsed as { blocks: Block[] }).blocks;
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
