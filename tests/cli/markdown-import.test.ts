import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "../../src/core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";

type Block = {
	id: string;
	type: string;
	style?: string;
	level?: number;
	list?: { level: number; numId: number };
	taskState?: "checked" | "unchecked";
	quoteDepth?: number;
	runs?: Array<{
		type: string;
		text?: string;
		latex?: string;
		display?: boolean;
		runStyle?: string;
		bold?: boolean;
		italic?: boolean;
		strike?: boolean;
		hyperlink?: { id: string; url?: string };
		trackedChange?: { id: string; kind: string };
	}>;
	rows?: Array<{
		cells: Array<{
			blocks: Block[];
		}>;
	}>;
};

async function readBlocks(docPath: string): Promise<Block[]> {
	const result = await runCli("read", docPath, "--ast");
	return (result.parsed as { blocks: Block[] }).blocks;
}

async function readMarkdown(docPath: string): Promise<string> {
	const result = await runCli("read", docPath);
	return result.stdout;
}

async function readPart(docPath: string, partPath: string): Promise<string> {
	const pkg = await Pkg.open(docPath);
	const bytes = await pkg.readBytes(partPath);
	return bytes ? new TextDecoder().decode(bytes) : "";
}

describe("docx insert/create --markdown — block features", () => {
	test("headings provision Heading1..6 baseline styles + apply pStyle", async () => {
		const docPath = join(tempWorkspace("md-headings"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"# H1\n\n## H2\n\n### H3\n\n#### H4\n\n##### H5\n\n###### H6",
		);

		const blocks = await readBlocks(docPath);
		const headings = blocks.filter((b) => b.style?.startsWith("Heading"));
		expect(headings.map((h) => h.style)).toEqual([
			"Heading1",
			"Heading2",
			"Heading3",
			"Heading4",
			"Heading5",
			"Heading6",
		]);

		// styles.xml should now define each baseline heading.
		const styles = await readPart(docPath, "word/styles.xml");
		for (let depth = 1; depth <= 6; depth++) {
			expect(styles).toContain(`w:styleId="Heading${depth}"`);
		}
	});

	test("paragraph + inline formatting: bold / italic / strike / inlineCode", async () => {
		const docPath = join(tempWorkspace("md-inline"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"Hello **bold**, *italic*, ~~strike~~, and `code`.",
		);

		const blocks = await readBlocks(docPath);
		const paragraph = blocks.find((b) => b.id === "p1");
		const runs = paragraph?.runs?.filter((r) => r.type === "text") ?? [];
		const bold = runs.find((r) => r.bold === true);
		const italic = runs.find((r) => r.italic === true);
		const strike = runs.find((r) => r.strike === true);
		const code = runs.find((r) => r.runStyle === "Code");
		expect(bold?.text).toBe("bold");
		expect(italic?.text).toBe("italic");
		expect(strike?.text).toBe("strike");
		expect(code?.text).toBe("code");
	});

	test("hyperlink: [text](url) mints a relationship + wraps the run", async () => {
		const docPath = join(tempWorkspace("md-link"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"Visit [Example](https://example.com) now.",
		);

		const blocks = await readBlocks(docPath);
		const paragraph = blocks.find((b) => b.id === "p1");
		const linked = paragraph?.runs?.find((r) => r.hyperlink);
		expect(linked?.hyperlink?.url).toBe("https://example.com");
		expect(linked?.text).toBe("Example");
	});

	test("bullet + ordered + nested lists allocate distinct numIds", async () => {
		const workspace = tempWorkspace("md-lists");
		const docPath = join(workspace, "out.docx");
		const mdPath = join(workspace, "lists.md");
		await runCli("create", docPath, "--text", "Intro.");
		await Bun.write(
			mdPath,
			"- a\n- b\n  - nested\n- c\n\n1. first\n2. second\n",
		);
		await runCli("insert", docPath, "--after", "p0", "--markdown-file", mdPath);

		const blocks = await readBlocks(docPath);
		const listItems = blocks.filter((b) => b.list !== undefined);
		expect(listItems.length).toBe(6); // 3 outer bullets + 1 nested + 2 ordered
		const bulletNumId = listItems[0]?.list?.numId;
		const orderedNumId = listItems[4]?.list?.numId;
		expect(bulletNumId).not.toBe(orderedNumId);
		// The nested bullet inherits the parent bullet's numId but at level 1.
		const nested = listItems[2];
		expect(nested?.list?.numId).toBe(bulletNumId);
		expect(nested?.list?.level).toBe(1);

		const numbering = await readPart(docPath, "word/numbering.xml");
		expect(numbering).toContain('w:numId="');
	});

	test("GFM task list: -[ ]/-[x] becomes a list paragraph with taskState", async () => {
		const workspace = tempWorkspace("md-task");
		const docPath = join(workspace, "out.docx");
		const mdPath = join(workspace, "task.md");
		await runCli("create", docPath, "--text", "Intro.");
		await Bun.write(mdPath, "- [ ] one\n- [x] two\n");
		await runCli("insert", docPath, "--after", "p0", "--markdown-file", mdPath);

		const blocks = await readBlocks(docPath);
		const tasks = blocks.filter((b) => b.taskState !== undefined);
		expect(tasks.map((t) => t.taskState)).toEqual(["unchecked", "checked"]);
	});

	test("fenced code block routes through CodeBlock emitter", async () => {
		const docPath = join(tempWorkspace("md-code"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"```python\nprint(1)\nprint(2)\n```",
		);

		const blocks = await readBlocks(docPath);
		// `--code` and `--markdown` with a fenced lang both emit
		// CodeBlock-<lang> as the pStyle (so the language survives round-trip);
		// match the prefix.
		const codeBlocks = blocks.filter((b) => b.style?.startsWith("CodeBlock"));
		expect(codeBlocks.length).toBe(2);
		expect(codeBlocks[0]?.runs?.map((r) => r.text).join("")).toBe("print(1)");
	});

	test("GFM table renders as <w:tbl> with one cell-paragraph per cell", async () => {
		const docPath = join(tempWorkspace("md-table"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"| A | B |\n| --- | --- |\n| 1 | 2 |\n| 3 | 4 |",
		);

		const blocks = await readBlocks(docPath);
		const table = blocks.find((b) => b.type === "table");
		expect(table).toBeDefined();
		expect(table?.rows?.length).toBe(3); // header + 2 body rows
		expect(table?.rows?.[0]?.cells.length).toBe(2);
		const firstCell = table?.rows?.[1]?.cells[0]?.blocks?.[0];
		expect(firstCell?.runs?.map((r) => r.text).join("")).toBe("1");
	});

	test("blockquote applies Quote style + quoteDepth=1 to each paragraph", async () => {
		const docPath = join(tempWorkspace("md-quote"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"> First line.\n>\n> Second line.",
		);

		const blocks = await readBlocks(docPath);
		const quotes = blocks.filter((b) => b.style === "Quote");
		expect(quotes.length).toBe(2);
		expect(quotes.every((b) => b.quoteDepth === 1)).toBe(true);
	});

	test("inline math + display math: $..$ and $$..$$ become EquationRun", async () => {
		const docPath = join(tempWorkspace("md-math"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"Inline: $x^2 + 1$ here.\n\n$$\n\\sum_{i=0}^n i\n$$",
		);

		const blocks = await readBlocks(docPath);
		const equations = blocks.flatMap((b) =>
			(b.runs ?? []).filter((r) => r.type === "equation"),
		);
		expect(equations.length).toBe(2);
		const inline = equations.find((e) => e.display === false);
		const display = equations.find((e) => e.display === true);
		// temml round-trip: `x^2` stays as `x^2`, `x^{long}` is brace-preserved.
		expect(inline?.latex?.replace(/\s/g, "")).toMatch(/x\^\{?2\}?/);
		expect(display?.latex?.replace(/\s/g, "")).toContain("\\sum");
	});

	test("footnoteReference + footnoteDefinition register in footnotes.xml", async () => {
		const docPath = join(tempWorkspace("md-footnote"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"Text with a ref[^1] inline.\n\n[^1]: This is the footnote body.",
		);

		const footnotes = await readPart(docPath, "word/footnotes.xml");
		expect(footnotes).toContain("This is the footnote body");

		const blocks = await readBlocks(docPath);
		const noteRef = blocks
			.flatMap((b) => b.runs ?? [])
			.find((r) => r.type === "noteRef");
		expect(noteRef).toBeDefined();
	});

	test("a footnote referenced multiple times mints one definition PER reference (OOXML 1:1)", async () => {
		// Markdown lets the same `[^1]` be cited repeatedly; Word treats N
		// references to a single footnote definition as corruption ("unreadable
		// content", repaired by cloning). So each reference must get its own
		// definition. Regression for that Word-validity bug.
		const docPath = join(tempWorkspace("md-footnote-reuse"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"First cite[^a] then second cite[^a] then a third[^a].\n\n[^a]: shared body.",
		);

		const footnotes = await readPart(docPath, "word/footnotes.xml");
		const definitions = [...footnotes.matchAll(/<w:footnote\s+w:id="(\d+)"/g)]
			.map((match) => match[1])
			.filter((id) => id !== "0"); // drop the id=0 continuationSeparator
		const document = await readPart(docPath, "word/document.xml");
		const references = [
			...document.matchAll(/<w:footnoteReference\s+w:id="(\d+)"/g),
		].map((match) => match[1]);

		// Three references → three distinct definitions, 1:1.
		expect(references).toHaveLength(3);
		expect(new Set(references).size).toBe(3);
		expect(definitions).toHaveLength(3);
		// Every reference points at a real definition.
		for (const id of references) expect(definitions).toContain(id);
	});

	test("a hyperlink inside a footnote body becomes a real <w:hyperlink> backed by the footnotes part's rels", async () => {
		const docPath = join(tempWorkspace("md-footnote-link"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"Cite[^s] here.\n\n[^s]: see [AP](https://ap.example.com/x) for detail.",
		);

		const footnotes = await readPart(docPath, "word/footnotes.xml");
		expect(footnotes).toContain("<w:hyperlink");

		// The hyperlink rel lives in the FOOTNOTES part's rels, not the document's,
		// so the `<w:hyperlink r:id>` inside footnotes.xml resolves correctly.
		const rels = await readPart(docPath, "word/_rels/footnotes.xml.rels");
		expect(rels).toContain("https://ap.example.com/x");
		expect(rels).toContain('TargetMode="External"');
		const rid = footnotes.match(/<w:hyperlink[^>]*r:id="([^"]+)"/)?.[1];
		expect(rid).toBeDefined();
		expect(rels).toContain(`Id="${rid}"`);
	});

	test("thematicBreak (---) becomes a HorizontalRule paragraph", async () => {
		const docPath = join(tempWorkspace("md-hr"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"Above.\n\n---\n\nBelow.",
		);

		const documentXml = await readPart(docPath, "word/document.xml");
		// The HR paragraph carries a `<w:pBdr><w:bottom/></w:pBdr>` — match that
		// shape rather than the markdown render (which uses `---` as a sentinel).
		expect(documentXml).toContain("<w:pBdr>");
		expect(documentXml).toContain("<w:bottom");
	});
});

describe("docx insert --markdown — blockquote depth + escape", () => {
	test("nested list inside blockquote uses QuoteListParagraph + quoteDepth=1", async () => {
		const workspace = tempWorkspace("md-quote-list");
		const docPath = join(workspace, "out.docx");
		const mdPath = join(workspace, "quote.md");
		await runCli("create", docPath, "--text", "Intro.");
		await Bun.write(mdPath, "> intro line.\n>\n> - one\n> - two\n");
		await runCli("insert", docPath, "--after", "p0", "--markdown-file", mdPath);

		const blocks = await readBlocks(docPath);
		const introQuote = blocks.find(
			(b) => b.style === "Quote" && b.runs?.[0]?.text?.includes("intro"),
		);
		expect(introQuote?.quoteDepth).toBe(1);

		const quotedListItems = blocks.filter(
			(b) => b.style === "QuoteListParagraph",
		);
		expect(quotedListItems.length).toBe(2);
		expect(quotedListItems.every((b) => b.quoteDepth === 1)).toBe(true);
		expect(quotedListItems.every((b) => b.list !== undefined)).toBe(true);
	});

	test("nested blockquote increments quoteDepth", async () => {
		const workspace = tempWorkspace("md-quote-nested");
		const docPath = join(workspace, "out.docx");
		const mdPath = join(workspace, "quote.md");
		await runCli("create", docPath, "--text", "Intro.");
		await Bun.write(mdPath, "> outer.\n>\n> > nested.\n> >\n> > > deep.\n");
		await runCli("insert", docPath, "--after", "p0", "--markdown-file", mdPath);

		const blocks = await readBlocks(docPath);
		const quotes = blocks.filter((b) => b.style === "Quote");
		const depths = quotes.map((b) => b.quoteDepth).sort();
		expect(depths).toEqual([1, 2, 3]);
	});

	test("code block inside blockquote escapes (no quoteDepth)", async () => {
		const workspace = tempWorkspace("md-quote-code");
		const docPath = join(workspace, "out.docx");
		const mdPath = join(workspace, "quote.md");
		await runCli("create", docPath, "--text", "Intro.");
		await Bun.write(
			mdPath,
			"> Intro.\n>\n> ```python\n> print(1)\n> ```\n>\n> Trailing.\n",
		);
		await runCli("insert", docPath, "--after", "p0", "--markdown-file", mdPath);

		const blocks = await readBlocks(docPath);
		// The two quote paragraphs flanking the code block are tagged; the
		// code block paragraphs themselves are not.
		const quotes = blocks.filter((b) => b.quoteDepth !== undefined);
		expect(quotes.length).toBe(2);
		const codeBlocks = blocks.filter((b) => b.style?.startsWith("CodeBlock"));
		expect(codeBlocks.length).toBe(1);
		expect(codeBlocks[0]?.quoteDepth).toBeUndefined();
	});

	test("read-back markdown reproduces `>` markers, including for nested lists", async () => {
		const workspace = tempWorkspace("md-quote-roundtrip");
		const docPath = join(workspace, "out.docx");
		const mdPath = join(workspace, "quote.md");
		await runCli("create", docPath, "--text", "Intro.");
		await Bun.write(
			mdPath,
			[
				"> outer paragraph.",
				">",
				"> - quoted bullet",
				"> - second bullet",
				">",
				"> > nested quote",
			].join("\n"),
		);
		await runCli("insert", docPath, "--after", "p0", "--markdown-file", mdPath);

		const rendered = (await runCli("read", docPath)).stdout;
		expect(rendered).toContain("> outer paragraph.");
		expect(rendered).toContain("> - quoted bullet");
		expect(rendered).toContain("> - second bullet");
		expect(rendered).toContain("> > nested quote");
	});
});

describe("docx insert --markdown — CriticMarkup", () => {
	test("tracking off: {++..++} flattens to plain text, {--..--} drops", async () => {
		const docPath = join(tempWorkspace("md-critic-off"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"Before {++kept++} and {--gone--} end.",
		);

		const blocks = await readBlocks(docPath);
		const paragraph = blocks.find((b) => b.id === "p1");
		const text = paragraph?.runs
			?.filter((r) => r.type === "text")
			.map((r) => r.text)
			.join("");
		// "Before " + "kept" + " and " (deletion dropped) + " end."
		expect(text).toContain("Before");
		expect(text).toContain("kept");
		expect(text).toContain("end.");
		expect(text).not.toContain("gone");
	});

	test("tracking on: {++..++} wraps runs in <w:ins>; {--..--} in <w:del>", async () => {
		const docPath = join(tempWorkspace("md-critic-on"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli("track-changes", docPath, "on");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"Before {++added++} and {--removed--} end.",
			"--author",
			"Test",
		);

		const documentXml = await readPart(docPath, "word/document.xml");
		expect(documentXml).toContain("<w:ins ");
		expect(documentXml).toContain("<w:del ");
		// The added text is inside a w:ins, removed inside w:del with delText.
		expect(documentXml).toMatch(/<w:ins[^>]*>[\s\S]*?added[\s\S]*?<\/w:ins>/);
		expect(documentXml).toMatch(/<w:delText[^>]*>removed</);
	});

	test("`edit --markdown` under tracking preserves CriticMarkup inner wrappers (no flatten)", async () => {
		// Regression for the code-review finding: `edit --markdown` routes
		// through `applyTrackedRangeReplace`, which previously flattened the
		// walker's inner CriticMarkup `<w:ins>` into the outer
		// `<w:ins>` wrapper carrying the editor's author — losing the
		// inner author/revisionId entirely. With `wrapContiguousTrackable`,
		// the inner wrapper passes through unchanged.
		const docPath = join(tempWorkspace("md-critic-edit-tracking"), "out.docx");
		await runCli("create", docPath, "--text", "Original paragraph.");
		await runCli("track-changes", docPath, "on");
		await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--markdown",
			"Rewritten paragraph with {++added++} inline.",
			"--author",
			"editor",
		);

		const documentXml = await readPart(docPath, "word/document.xml");
		// The outer Edit wrap carries the editor's author.
		const editorWraps = documentXml.match(/<w:ins[^>]*w:author="editor"/g);
		expect(editorWraps?.length ?? 0).toBeGreaterThanOrEqual(1);
		// The inner CriticMarkup `<w:ins>` from the markdown walker carries
		// its own (default) author — proving the wrapper survived rather
		// than being subsumed by the outer Edit wrap. Test against the
		// non-editor author to confirm the inner wrapper is distinct.
		const innerWraps = documentXml.match(
			/<w:ins[^>]*w:author="(?!editor")[^"]+"/g,
		);
		expect(innerWraps?.length ?? 0).toBeGreaterThanOrEqual(1);
		// And the inserted text itself must survive as content.
		expect(documentXml).toContain("added");
	});
});

describe("docx edit --markdown", () => {
	test("single-paragraph replace: pN ← parsed blocks (untracked)", async () => {
		const docPath = join(tempWorkspace("md-edit-single"), "out.docx");
		await runCli("create", docPath, "--text", "Old paragraph.");
		await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--markdown",
			"# Replacement\n\nNew body.",
		);

		const blocks = await readBlocks(docPath);
		// Body is now the replacement (heading + paragraph) plus the section break.
		expect(blocks[0]?.style).toBe("Heading1");
		expect(blocks[1]?.runs?.map((r) => r.text).join("")).toContain("New body");
	});

	test("range replace pN-pM: contiguous paragraphs collapse to parsed blocks", async () => {
		const workspace = tempWorkspace("md-edit-range");
		const docPath = join(workspace, "out.docx");
		const mdPath = join(workspace, "replacement.md");
		await runCli("create", docPath, "--text", "p0.");
		await runCli("insert", docPath, "--after", "p0", "--text", "p1.");
		await runCli("insert", docPath, "--after", "p1", "--text", "p2.");
		await runCli("insert", docPath, "--after", "p2", "--text", "p3.");
		// Replace p1-p2 with two list items. We write through --markdown-file
		// because a `--markdown` value starting with `-` confuses Node's
		// parseArgs ("ambiguous"); --markdown-file accepts any content.
		await Bun.write(mdPath, "- alpha\n- beta\n");
		await runCli("edit", docPath, "--at", "p1-p2", "--markdown-file", mdPath);

		const blocks = await readBlocks(docPath);
		const lists = blocks.filter((b) => b.list !== undefined);
		expect(lists.length).toBe(2);
		expect(lists.map((b) => b.runs?.map((r) => r.text).join(""))).toEqual([
			"alpha",
			"beta",
		]);
	});
});

describe("docx create --from PATH.md", () => {
	test("creates a doc whose body is the parsed markdown", async () => {
		const workspace = tempWorkspace("md-create-from");
		const docPath = join(workspace, "out.docx");
		const mdPath = join(workspace, "src.md");
		await Bun.write(mdPath, "# Title\n\nIntro paragraph.\n\n- one\n- two");

		await runCli("create", docPath, "--from", mdPath);

		const blocks = await readBlocks(docPath);
		expect(blocks[0]?.style).toBe("Heading1");
		expect(blocks[0]?.runs?.map((r) => r.text).join("")).toBe("Title");
		// Intro paragraph + 2 list items + section break.
		expect(blocks.length).toBeGreaterThanOrEqual(4);
	});

	test("--from with --text errors as mutex", async () => {
		const docPath = join(tempWorkspace("md-from-text"), "out.docx");
		const result = await runCli(
			"create",
			docPath,
			"--text",
			"x",
			"--from",
			"/dev/null",
		);
		expect(result.exitCode).toBeGreaterThan(0);
		expect(result.parsed).toMatchObject({
			code: "USAGE",
		});
	});
});

describe("docx insert/edit --markdown — flag conflicts", () => {
	test("insert --markdown rejects --style with a clear hint", async () => {
		const docPath = join(tempWorkspace("md-conflict-insert"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"# Heading",
			"--style",
			"Heading2",
		);
		expect(result.exitCode).toBeGreaterThan(0);
		expect(result.parsed).toMatchObject({
			code: "USAGE",
		});
		const parsed = result.parsed as { error?: string };
		expect(parsed.error ?? "").toContain("--markdown");
	});

	test("edit --markdown rejects --style with a clear hint", async () => {
		const docPath = join(tempWorkspace("md-conflict-edit"), "out.docx");
		await runCli("create", docPath, "--text", "Original.");
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--markdown",
			"# Replacement",
			"--style",
			"Heading2",
		);
		expect(result.exitCode).toBeGreaterThan(0);
		expect(result.parsed).toMatchObject({
			code: "USAGE",
		});
	});
});

describe("docx insert --markdown — error paths", () => {
	test("inline math with malformed LaTeX surfaces a usage error", async () => {
		const docPath = join(tempWorkspace("md-bad-math"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			// Mismatched brace — temml's parser rejects with a clear error.
			"Broken: $\\frac{a$ here.",
		);
		expect(result.exitCode).toBeGreaterThan(0);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("--markdown-file PATH not found surfaces FILE_NOT_FOUND", async () => {
		const docPath = join(tempWorkspace("md-no-file"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown-file",
			"/does/not/exist.md",
		);
		expect(result.exitCode).toBeGreaterThan(0);
		expect(result.parsed).toMatchObject({
			code: "FILE_NOT_FOUND",
		});
	});
});

describe("docx insert --markdown — image round-trip via SHA-256 hash", () => {
	test("read → edit --markdown round-trip preserves the existing media part", async () => {
		const workspace = tempWorkspace("md-img-roundtrip");
		const docPath = join(workspace, "a.docx");
		await runCli("create", docPath, "--text", "Doc A");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--image",
			"tests/fixtures/assets/sample.png",
		);

		// First read emits `![alt](<sha256>.<ext>)`. Capture the hash and
		// confirm the URL shape.
		const rendered = (await runCli("read", docPath)).stdout;
		const match = rendered.match(/\(([0-9a-f]{64})\.png\)/);
		expect(match).not.toBeNull();
		const hash = match?.[1] ?? "";
		expect(hash).toHaveLength(64);

		// Round-trip through `edit --markdown` — pass the hash back as the
		// image URL. The walker should reuse the existing relationship,
		// not mint a second media part.
		const mdPath = join(workspace, "edit.md");
		await Bun.write(mdPath, `Captioned: ![logo](${hash}.png)`);
		await runCli("edit", docPath, "--at", "p1", "--markdown-file", mdPath);

		// Only one image registered in the body — the same media part we
		// started with. A duplicate would surface as a second `imgN` entry.
		const { Document } = await import("../../src/core");
		const document = await Document.open(docPath);
		expect([...document.body.imageById.keys()]).toEqual(["img0"]);

		// The new alt text propagated; the hash is still recognized on a
		// subsequent read.
		const reRead = (await runCli("read", docPath)).stdout;
		expect(reRead).toContain(`![logo](${hash}.png)`);
	});

	test("hash-shaped URL that doesn't match any image in the target doc → clear error", async () => {
		const docPath = join(tempWorkspace("md-img-missing"), "a.docx");
		await runCli("create", docPath, "--text", "Intro.");
		const fakeHash = "0".repeat(64);
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			`Missing: ![alt](${fakeHash}.png)`,
		);
		expect(result.exitCode).toBeGreaterThan(0);
		expect(result.parsed).toMatchObject({
			code: "IMAGE_SOURCE",
		});
		const parsed = result.parsed as { error?: string };
		expect(parsed.error ?? "").toContain(fakeHash);
	});
});

describe("docx insert --markdown — round-trip via read --markdown", () => {
	test("comprehensive sample round-trips through read --markdown", async () => {
		const docPath = join(tempWorkspace("md-roundtrip"), "out.docx");
		const source = [
			"# Heading 1",
			"",
			"## Heading 2",
			"",
			"A paragraph with **bold**, *italic*, ~~strike~~, and `code`.",
			"",
			"Visit [Example](https://example.com).",
			"",
			"- bullet a",
			"- bullet b",
			"",
			"1. ordered one",
			"2. ordered two",
			"",
			"```python",
			"print(1)",
			"```",
			"",
			"| A | B |",
			"| --- | --- |",
			"| 1 | 2 |",
			"",
			"Inline: $x^2$.",
			"",
			"$$",
			"\\sum_{i=0}^n i",
			"$$",
			"",
			"Footnote ref[^a] here.",
			"",
			"[^a]: footnote body.",
		].join("\n");

		await runCli("create", docPath, "--text", "Seed.");
		await runCli("edit", docPath, "--at", "p0", "--markdown", source);

		const rendered = await readMarkdown(docPath);
		// The render side preserves the dialect we emit; assert the features
		// we put in came back out. We don't diff verbatim because read --markdown
		// adds locator comments and normalizes some whitespace.
		expect(rendered).toContain("# Heading 1");
		expect(rendered).toContain("## Heading 2");
		expect(rendered).toContain("**bold**");
		expect(rendered).toContain("*italic*");
		expect(rendered).toContain("~~strike~~");
		expect(rendered).toContain("`code`");
		expect(rendered).toContain("[Example](https://example.com)");
		expect(rendered).toContain("- bullet a");
		expect(rendered).toContain("```python");
		expect(rendered).toContain("| A");
		expect(rendered).toMatch(/\$x\^\{?2\}?\$/);
		expect(rendered).toContain("$$");
		expect(rendered).toContain("[^fn1]");
		expect(rendered).toContain("footnote body");
	});
});
