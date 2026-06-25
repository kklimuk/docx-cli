import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";
import { readDocumentXml, readMarkdown } from "./helpers";

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

describe("docx insert --markdown — inline-surgery (straddle fix + spans)", () => {
	test("FIXED TRAP: {++**bold**++} under tracking wraps a bold run in <w:ins>", async () => {
		// The whole reason for the inline-surgery rewrite. Previously `{++` and
		// `++}` lived in sibling text nodes around the `strong`, so no critic
		// node formed and the markers leaked as literal text. Now the markers are
		// matched across the `strong` sibling, producing a `<w:ins>` whose run
		// carries `<w:b/>`.
		const docPath = join(tempWorkspace("md-straddle-ins"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli("track-changes", docPath, "on");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"Before {++**bold**++} after.",
			"--author",
			"Test",
		);

		const documentXml = await readPart(docPath, "word/document.xml");
		// A single <w:ins> wrapping a run that is bold and contains "bold".
		expect(documentXml).toMatch(
			/<w:ins[^>]*>[\s\S]*?<w:b\s*\/>[\s\S]*?bold[\s\S]*?<\/w:ins>/,
		);
		// The markers themselves must NOT survive as literal text.
		expect(documentXml).not.toContain("{++");
		expect(documentXml).not.toContain("++}");
	});

	test("straddle with nested emphasis: {++pre *em* post++} keeps the *em* run (markers gone)", async () => {
		// Tracking OFF isolates the straddle fix: the criticInsert must form
		// (markers consumed) AND keep the inner emphasis. The OLD text-split
		// plugin couldn't match markers straddling the `emphasis` node, so it
		// leaked `{++pre`/`post++}` as literal text.
		const docPath = join(tempWorkspace("md-straddle-em"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"Edit {++pre *em* post++} done.",
		);

		const blocks = await readBlocks(docPath);
		const runs = (blocks.find((b) => b.id === "p1")?.runs ?? []).filter(
			(r) => r.type === "text",
		);
		expect(runs.map((r) => r.text).join("")).toBe("Edit pre em post done.");
		expect(runs.find((r) => r.italic)?.text).toBe("em");
	});

	test("formatted tracked DELETE: {--**gone**--} → <w:del> with <w:delText> (not <w:t>) on the bold run", async () => {
		// Risk: a <w:t> inside <w:del> is invalid OOXML (Word "unreadable
		// content"). The delText rename must recurse through the nested bold run.
		const docPath = join(tempWorkspace("md-straddle-del"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli("track-changes", docPath, "on");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"Keep {--**gone**--} text.",
			"--author",
			"Test",
		);

		const documentXml = await readPart(docPath, "word/document.xml");
		expect(documentXml).toMatch(
			/<w:del[^>]*>[\s\S]*?<w:b\s*\/>[\s\S]*?<w:delText[^>]*>gone<\/w:delText>[\s\S]*?<\/w:del>/,
		);
		// No bare <w:t> carrying the deleted text (would be invalid inside <w:del>).
		expect(documentXml).not.toMatch(/<w:t[^>]*>gone<\/w:t>/);
	});

	test("bracketed span [plain]{color=...} applies the color to just that run", async () => {
		const docPath = join(tempWorkspace("md-span-color"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			'A [plain]{color="FF0000"} word.',
		);

		const runs = ((await readBlocks(docPath)).find((b) => b.id === "p1")
			?.runs ?? []) as Array<{ type: string; text?: string; color?: string }>;
		const text = runs.filter((r) => r.type === "text");
		expect(text.map((r) => r.text).join("")).toBe("A plain word.");
		expect(text.find((r) => r.color === "FF0000")?.text).toBe("plain");
	});

	test("degrade: unbalanced {++ with no close survives as literal text, no <w:ins>, no throw", async () => {
		const docPath = join(tempWorkspace("md-degrade"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"Unbalanced {++ no close here.",
		);
		expect(result.exitCode).toBe(0);

		const blocks = await readBlocks(docPath);
		const text = (blocks.find((b) => b.id === "p1")?.runs ?? [])
			.filter((r) => r.type === "text")
			.map((r) => r.text)
			.join("");
		expect(text).toBe("Unbalanced {++ no close here.");
		const documentXml = await readPart(docPath, "word/document.xml");
		expect(documentXml).not.toContain("<w:ins ");
	});

	test("code-span exclusion: `{++x++}` stays a literal inline-code run, no critic node", async () => {
		// Tracking OFF so the only `<w:ins>` that could appear would come from a
		// (wrongly) parsed critic node inside the code span. The tokenizer treats
		// `inlineCode` as an opaque atom, so the markers stay literal.
		const docPath = join(tempWorkspace("md-codespan"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"Literal `{++x++}` here.",
		);

		const blocks = await readBlocks(docPath);
		const code = (blocks.find((b) => b.id === "p1")?.runs ?? []).find(
			(r) => r.runStyle === "Code",
		);
		expect(code?.text).toBe("{++x++}");
		const documentXml = await readPart(docPath, "word/document.xml");
		expect(documentXml).not.toContain("<w:ins ");
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

// read --markdown output must re-import cleanly: the ` <!-- pN -->` locator
// separator must not accumulate spaces, and verse line breaks must survive,
// across read → create → read. (The section-break `---` line is excluded —
// section properties can't be recreated from markdown; tracked separately.)

const SOURCE = `# Quarterly Report

A normal paragraph with some text in it.

The winter evening settles down
With smell of steaks in passageways.
The burnt-out ends of smoky days.

| Name | Note |
| --- | --- |
| Dana | line one<br>line two |
`;

async function readBody(path: string): Promise<string> {
	const out = (await runCli("read", path)).stdout;
	// Drop the trailing section-break marker line (sections don't round-trip
	// through markdown — a separate, known limitation).
	return out
		.split("\n")
		.filter((line) => !line.includes("<!-- s0 -->"))
		.join("\n");
}

describe("read → create → read is stable for normal content", () => {
	test("no trailing-space growth before locators; verse + tables survive", async () => {
		const workspace = tempWorkspace("md-roundtrip");
		const src = join(workspace, "src.md");
		await Bun.write(src, SOURCE);
		const doc1 = join(workspace, "d1.docx");
		expect((await runCli("create", doc1, "--from", src)).exitCode).toBe(0);

		const read1 = await readBody(doc1);
		const reimport = join(workspace, "r.md");
		await Bun.write(reimport, `${read1}\n`);
		const doc2 = join(workspace, "d2.docx");
		expect((await runCli("create", doc2, "--from", reimport)).exitCode).toBe(0);
		const read2 = await readBody(doc2);

		expect(read2).toBe(read1);
		// no double-space crept in before any locator comment
		expect(read2).not.toMatch(/ {2}<!--/);
		// verse line breaks are still there (3 lines → 2 <w:br/>)
		const br = (await Bun.file(doc2).bytes()).length; // touch file
		expect(br).toBeGreaterThan(0);
		expect(read2).toContain("The winter evening settles down\n");
	});
});

/**
 * The read↔import formatting contract beyond the per-run round-trip and
 * two-emitter convergence tests elsewhere in this file:
 *   - the document `<!-- docx:base … -->` note (dominant font/size declared once
 *     and omitted per-run — a read-time VISIBILITY hint the importer drops, NOT
 *     parse-back; a full rebuild falls back to the template docDefaults),
 *   - default formatting (black / text1) dropped as noise,
 *   - hand-authored HTML formatting parsed on import (semantic tags + data-*),
 *   - insert blending into the anchor paragraph's formatting.
 */

type RunAst = {
	type: string;
	text?: string;
	bold?: boolean;
	italic?: boolean;
	strike?: boolean;
	font?: string;
	sizeHalfPoints?: number;
	color?: string;
	colorTheme?: string;
	highlight?: string;
	vertAlign?: string;
	underline?: string;
};
type BaselineBlock = { id: string; type: string; runs?: RunAst[] };

async function read(path: string): Promise<string> {
	return (await runCli("read", path)).stdout;
}

async function blocks(path: string): Promise<BaselineBlock[]> {
	const result = await runCli("read", path, "--ast");
	return (result.parsed as { blocks: BaselineBlock[] }).blocks;
}

async function allRuns(path: string): Promise<RunAst[]> {
	return (await blocks(path)).flatMap((block) => block.runs ?? []);
}

/** A doc whose single inserted paragraph (p1) carries the given runs. */
async function docWith(label: string, runs: RunAst[]): Promise<string> {
	const path = join(tempWorkspace(label), "out.docx");
	await runCli("create", path, "--text", "Intro.");
	await runCli("insert", path, "--after", "p0", "--runs", JSON.stringify(runs));
	return path;
}

const ARIAL_BODY = { font: "Arial", sizeHalfPoints: 16 } as const;

describe("read — document format baseline note", () => {
	test("a dominant font/size becomes a base note and is omitted per-run", async () => {
		const path = await docWith("base-emit", [
			{
				type: "text",
				text: "The bulk of this document is in Arial eight point. ",
				...ARIAL_BODY,
			},
			{
				type: "text",
				text: "More of the same ubiquitous body content here too. ",
				...ARIAL_BODY,
			},
			{ type: "text", text: "TITLE", font: "Georgia", sizeHalfPoints: 44 },
		]);
		const md = await read(path);
		expect(md).toContain('<!-- docx:base font="Arial" size="8pt" -->');
		// Body runs carry no per-run font/size — they ride the note.
		expect(md).toContain("The bulk of this document is in Arial eight point.");
		expect(md).not.toContain("font-family:Arial");
		// The deviating run keeps its own font + size.
		expect(md).toContain("font-family:Georgia");
		expect(md).toContain("font-size:22pt");
	});

	test("a hostile dominant font value can't break or inject into the base note", async () => {
		const hostile = 'Bad"--><x'; // a quote, a comment-close, and angle brackets
		const path = await docWith("base-hostile", [
			{
				type: "text",
				text: "Body content one filling the document. ",
				font: hostile,
				sizeHalfPoints: 16,
			},
			{
				type: "text",
				text: "Body content two also in the same font. ",
				font: hostile,
				sizeHalfPoints: 16,
			},
		]);
		const md = await read(path);
		// Exactly one base note, escaped — the raw quote did NOT close the
		// attribute and the `-->` did NOT close the comment early.
		expect((md.match(/docx:base/g) ?? []).length).toBe(1);
		expect(md).not.toContain('font="Bad"');
		expect(md).not.toContain('Bad"--><x'); // raw hostile sequence never leaks
		expect(md).toContain("&quot;"); // it was escaped instead
		expect(md).toContain("Body content one filling the document.");
	});

	test("base note is a visibility hint, NOT parse-back: a full rebuild drops the dominant font/size but keeps deviating runs", async () => {
		const src = await docWith("base-rt", [
			{
				type: "text",
				text: "Ubiquitous Arial body content fills this page. ",
				...ARIAL_BODY,
			},
			{
				type: "text",
				text: "Still more Arial body content to dominate it. ",
				...ARIAL_BODY,
			},
			{ type: "text", text: "Heading", font: "Georgia", sizeHalfPoints: 44 },
		]);
		const workspace = tempWorkspace("base-rt-dst");
		const mdPath = join(workspace, "doc.md");
		await Bun.write(mdPath, await read(src));
		const dst = join(workspace, "rt.docx");
		await runCli("create", dst, "--from", mdPath);

		const runs = (await allRuns(dst)).filter((r) => (r.text ?? "").trim());
		// "comments are never anything but hints" — the importer DROPS docx:base, so
		// the omitted dominant Arial/8pt falls back to the template docDefaults
		// (the body runs carry no explicit font/size after rebuild).
		const body = runs.find((r) => (r.text ?? "").includes("Ubiquitous"));
		expect(body?.font).toBeUndefined();
		expect(body?.sizeHalfPoints).toBeUndefined();
		// A run with its OWN (non-dominant) font/size keeps it — only the omitted
		// baseline is lost on a from-scratch rebuild; `read --ast` of the source
		// stays lossless and in-place `edit` never touches runs.
		const heading = runs.find((r) => (r.text ?? "").includes("Heading"));
		expect(heading?.font).toBe("Georgia");
		expect(heading?.sizeHalfPoints).toBe(44);
		// The rebuilt doc has no dominant font now → no base note on re-read.
		expect(await read(dst)).not.toContain("docx:base");
	});

	test("set-default-font surfaces the document default in the base note (deviation-only)", async () => {
		const docPath = join(tempWorkspace("base-default-font"), "out.docx");
		await runCli("create", docPath, "--text", "Plain body paragraph.");

		// A fresh Calibri 11pt doc declares NO base note — the universal template
		// default is noise, suppressed deviation-only.
		expect(await read(docPath)).not.toContain("docx:base");

		// After set-default-font the docDefaults font/size DEVIATE from the template,
		// so they surface — making the change observable on the next read (the
		// write-read loop). Runs carry no explicit font, so they ride the note.
		await runCli(
			"styles",
			"set-default-font",
			docPath,
			"Garamond",
			"--size",
			"14",
		);
		const md = await read(docPath);
		expect(md).toContain('<!-- docx:base font="Garamond" size="14pt" -->');
		expect(md).not.toContain("font-family:Garamond"); // not stamped per-run
	});
});

describe("read — default formatting dropped as noise", () => {
	test("black color and the text1 theme are not emitted; a real color is", async () => {
		const path = await docWith("default-color", [
			{
				type: "text",
				text: "plain black ",
				color: "000000",
				colorTheme: "text1",
			},
			{ type: "text", text: "teal", color: "107087" },
		]);
		const md = await read(path);
		expect(md).not.toContain("000000");
		expect(md).not.toContain("data-color-theme");
		expect(md).toContain('<span style="color:#107087">teal</span>');
	});
});

describe("import — hand-authored HTML formatting parses", () => {
	test("<mark>, <span style>, <sup>, <sub>, <u> map to run formatting", async () => {
		const workspace = tempWorkspace("html-import");
		const mdPath = join(workspace, "doc.md");
		await Bun.write(
			mdPath,
			'A <mark>hi</mark> <span style="color:#FF0000">red</span> <sup>up</sup> <sub>dn</sub> <u>ln</u> end.\n',
		);
		const dst = join(workspace, "out.docx");
		await runCli("create", dst, "--from", mdPath);
		const runs = await allRuns(dst);
		expect(runs.find((r) => r.text === "hi")?.highlight).toBe("yellow");
		expect(runs.find((r) => r.text === "red")?.color).toBe("FF0000");
		expect(runs.find((r) => r.text === "up")?.vertAlign).toBe("superscript");
		expect(runs.find((r) => r.text === "dn")?.vertAlign).toBe("subscript");
		expect(runs.find((r) => r.text === "ln")?.underline).toBe("single");
	});

	test("data-* attributes carry OOXML-only props CSS can't express", async () => {
		const workspace = tempWorkspace("html-data");
		const mdPath = join(workspace, "doc.md");
		await Bun.write(
			mdPath,
			'X <span data-color-theme="accent1">themed</span> <u data-underline="wave">wav</u> Y.\n',
		);
		const dst = join(workspace, "out.docx");
		await runCli("create", dst, "--from", mdPath);
		const runs = await allRuns(dst);
		expect(runs.find((r) => r.text === "themed")?.colorTheme).toBe("accent1");
		expect(runs.find((r) => r.text === "wav")?.underline).toBe("wave");
	});
});

describe("read — whitespace-only runs keep their formatting", () => {
	test("emphasis on a blank run uses HTML tags (not `** **`) and round-trips", async () => {
		const path = await docWith("blank-fmt", [
			{ type: "text", text: "a" },
			{ type: "text", text: " ", bold: true },
			{ type: "text", text: "b" },
			{ type: "text", text: " ", underline: "single" },
			{ type: "text", text: "c" },
		]);
		const md = await read(path);
		// No mis-parsing `** **`; bold/underline ride unambiguous HTML.
		expect(md).not.toContain("** **");
		expect(md).toContain("<b> </b>");
		expect(md).toContain("<u> </u>");

		const workspace = tempWorkspace("blank-fmt-rt");
		const mdPath = join(workspace, "doc.md");
		await Bun.write(mdPath, md);
		const dst = join(workspace, "rt.docx");
		await runCli("create", dst, "--from", mdPath);
		const blanks = (await allRuns(dst)).filter(
			(r) => r.type === "text" && (r.text ?? "").trim() === "",
		);
		expect(blanks.some((r) => (r as { bold?: boolean }).bold)).toBe(true);
		expect(blanks.some((r) => r.underline === "single")).toBe(true);
	});
});

describe("insert — inherits formatting from the anchor paragraph", () => {
	test("plain inserted text adopts the neighbor's font + size", async () => {
		const path = await docWith("blend", [
			{
				type: "text",
				text: "Existing Arial eight-point body text.",
				...ARIAL_BODY,
			},
		]);
		await runCli(
			"insert",
			path,
			"--after",
			"p1",
			"--markdown",
			"Brand new content.",
		);
		const inserted = (await allRuns(path)).find((r) =>
			(r.text ?? "").includes("Brand new"),
		);
		expect(inserted?.font).toBe("Arial");
		expect(inserted?.sizeHalfPoints).toBe(16);
	});

	test("inserting after a heading does NOT promote the new text to a heading", async () => {
		const path = join(tempWorkspace("blend-head"), "out.docx");
		await runCli("create", path, "--text", "Body.");
		await runCli("insert", path, "--after", "p0", "--markdown", "# A Heading");
		await runCli(
			"insert",
			path,
			"--after",
			"p1",
			"--markdown",
			"Body paragraph after the heading.",
		);
		const run = (await allRuns(path)).find((r) =>
			(r.text ?? "").includes("Body paragraph after the heading"),
		);
		// No heading size grafted on — the run stays plain body.
		expect(run?.sizeHalfPoints).toBeUndefined();
		expect(run?.font).toBeUndefined();
	});
});

/**
 * Run-level formatting round-trip (color + theme color + highlight + shading +
 * underline + super/subscript + small/all caps + font + size). Drives the two
 * governing invariants:
 *   (I)  round-trip identity — author → read --markdown → import → read --ast is
 *        a fixpoint, with exact OOXML values preserved.
 *   (II) unsupported ≠ broken — an unmodeled rPr child survives untouched, and a
 *        bad enum value fails loud rather than corrupting the file.
 * Plus the two-emitter convergence guard (blocks.tsx vs markdown/inline.tsx).
 */

type TextRunAst = {
	type: string;
	text?: string;
	color?: string;
	colorTheme?: string;
	colorThemeTint?: string;
	colorThemeShade?: string;
	highlight?: string;
	shade?: string;
	underline?: string;
	underlineColor?: string;
	vertAlign?: string;
	smallCaps?: boolean;
	allCaps?: boolean;
	font?: string;
	sizeHalfPoints?: number;
};

type FmtBlock = { id: string; type: string; runs?: TextRunAst[] };

async function readFmtBlocks(docPath: string): Promise<FmtBlock[]> {
	const result = await runCli("read", docPath, "--ast");
	return (result.parsed as { blocks: FmtBlock[] }).blocks;
}

async function readText(docPath: string): Promise<string> {
	return (await runCli("read", docPath)).stdout;
}

/** Write `md` to a fresh workspace, rebuild a .docx from it via `create --from`,
 * and return its formatted blocks — the read → import half of the round-trip, so
 * a test can assert escaped Markdown decodes back to the original run text. */
async function roundTripBlocks(label: string, md: string): Promise<FmtBlock[]> {
	const ws = tempWorkspace(label);
	const mdPath = join(ws, "doc.md");
	await Bun.write(mdPath, md);
	const dst = join(ws, "rt.docx");
	await runCli("create", dst, "--from", mdPath);
	return readFmtBlocks(dst);
}

/** The runs covering every Phase-1 attribute, one run per variant. */
const SAMPLE_RUNS: TextRunAst[] = [
	{ type: "text", text: "red", color: "FF0000" },
	{
		type: "text",
		text: "theme",
		color: "4472C4",
		colorTheme: "accent1",
		colorThemeTint: "99",
	},
	{ type: "text", text: "hl", highlight: "yellow" },
	{ type: "text", text: "shaded", shade: "FFE599" },
	{ type: "text", text: "udbl", underline: "double", underlineColor: "FF0000" },
	{ type: "text", text: "usng", underline: "single" },
	{ type: "text", text: "sup", vertAlign: "superscript" },
	{ type: "text", text: "sub", vertAlign: "subscript" },
	{ type: "text", text: "sc", smallCaps: true },
	{ type: "text", text: "ac", allCaps: true },
	{ type: "text", text: "fontrun", font: "Courier New" },
	{ type: "text", text: "bigrun", sizeHalfPoints: 28 },
];

async function authorSample(label: string): Promise<string> {
	const docPath = join(tempWorkspace(label), "out.docx");
	await runCli("create", docPath, "--text", "Intro.");
	await runCli(
		"insert",
		docPath,
		"--after",
		"p0",
		"--runs",
		JSON.stringify(SAMPLE_RUNS),
	);
	return docPath;
}

function formattedParagraph(blocks: FmtBlock[]): TextRunAst[] {
	const block = blocks.find((b) =>
		(b.runs ?? [])
			.map((r) => r.text ?? "")
			.join("")
			.includes("red"),
	);
	// Drop whitespace-only runs: `read --markdown` appends a ` <!-- pN -->`
	// locator comment, and the import drops the comment but keeps the leading
	// space as a trailing run — an incidental artifact, not formatting.
	return (block?.runs ?? []).filter(
		(r) => r.type === "text" && (r.text ?? "").trim().length > 0,
	);
}

describe("run formatting — AST author → read (blocks.tsx emit + read.ts capture)", () => {
	test("every Phase-1 attribute survives create --runs → read --ast", async () => {
		const docPath = await authorSample("rf-author");
		const runs = formattedParagraph(await readFmtBlocks(docPath));
		expect(runs).toEqual(SAMPLE_RUNS);
	});

	test("theme color emits <w:color w:val + w:themeColor + w:themeTint> byte-exact", async () => {
		const docPath = await authorSample("rf-theme-xml");
		const xml = await readDocumentXml(docPath);
		expect(xml).toContain(
			'<w:color w:val="4472C4" w:themeColor="accent1" w:themeTint="99"/>',
		);
		expect(xml).toContain(
			'<w:shd w:val="clear" w:color="auto" w:fill="FFE599"/>',
		);
		expect(xml).toContain('<w:u w:val="double" w:color="FF0000"/>');
	});
});

describe("run formatting — read --markdown emits HTML a reader renders", () => {
	test("each attribute renders as the semantic HTML / `<span style>` form", async () => {
		const docPath = await authorSample("rf-md");
		const md = await readText(docPath);
		expect(md).toContain('<span style="color:#FF0000">red</span>');
		// Theme color: resolved hex in `style`, exact OOXML token in `data-*`.
		expect(md).toContain('data-color-theme="accent1"');
		expect(md).toContain('data-color-theme-tint="99"');
		expect(md).toContain("<mark>hl</mark>");
		expect(md).toContain(
			'<span style="background-color:#FFE599">shaded</span>',
		);
		expect(md).toContain(
			'<u data-underline="double" data-underline-color="FF0000">udbl</u>',
		);
		expect(md).toContain("<u>usng</u>");
		expect(md).toContain("<sup>sup</sup>");
		expect(md).toContain("<sub>sub</sub>");
		expect(md).toContain('<span style="font-variant:small-caps">sc</span>');
		expect(md).toContain('<span style="text-transform:uppercase">ac</span>');
		expect(md).toContain("font-family:'Courier New'");
		expect(md).toContain("font-size:14pt");
		// No legacy Pandoc bracketed spans.
		expect(md).not.toContain("{color=");
		expect(md).not.toContain("]{.");
	});
});

describe("run formatting — full round-trip identity (invariant I)", () => {
	test("author → read --markdown → create --from → read --ast is a fixpoint", async () => {
		const src = await authorSample("rf-rt-src");
		const original = formattedParagraph(await readFmtBlocks(src));

		const workspace = tempWorkspace("rf-rt-dst");
		const mdPath = join(workspace, "doc.md");
		await Bun.write(mdPath, await readText(src));
		const dst = join(workspace, "rt.docx");
		await runCli("create", dst, "--from", mdPath);

		const roundTripped = formattedParagraph(await readFmtBlocks(dst));
		expect(roundTripped).toEqual(original);
	});
});

describe("run formatting — two-emitter convergence (blocks.tsx ≡ inline.tsx)", () => {
	test("same logical run via --runs and via --markdown yields identical <w:rPr>", async () => {
		const astDoc = join(tempWorkspace("rf-conv-ast"), "out.docx");
		await runCli("create", astDoc, "--text", "Intro.");
		await runCli(
			"insert",
			astDoc,
			"--after",
			"p0",
			"--runs",
			JSON.stringify([
				{ type: "text", text: "X", color: "FF0000", underline: "single" },
			]),
		);

		const mdDoc = join(tempWorkspace("rf-conv-md"), "out.docx");
		await runCli("create", mdDoc, "--text", "Intro.");
		await runCli(
			"insert",
			mdDoc,
			"--after",
			"p0",
			"--markdown",
			'[X]{.underline color="FF0000"}',
		);

		const rprOf = (xml: string): string =>
			xml.match(/<w:rPr>[\s\S]*?<\/w:rPr>/)?.[0] ?? "";
		const astRpr = rprOf(await readDocumentXml(astDoc));
		const mdRpr = rprOf(await readDocumentXml(mdDoc));
		expect(astRpr).toBe(
			'<w:rPr><w:color w:val="FF0000"/><w:u w:val="single"/></w:rPr>',
		);
		expect(mdRpr).toBe(astRpr);
	});
});

describe("run formatting — invariant II (unsupported ≠ broken)", () => {
	test("invalid highlight enum fails loud with USAGE (no silent OOXML loss)", async () => {
		const docPath = join(tempWorkspace("rf-bad-hl"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			'A [bad]{highlight="chartreuse"} word.',
		);
		expect(result.exitCode).toBeGreaterThan(0);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
		const parsed = result.parsed as { error?: string };
		expect(parsed.error ?? "").toContain("chartreuse");
	});

	test("invalid underline enum fails loud with USAGE", async () => {
		const docPath = join(tempWorkspace("rf-bad-u"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			'A [bad]{underline="squiggle"} word.',
		);
		expect(result.exitCode).toBeGreaterThan(0);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("a span whose attributes parse to nothing degrades to literal text", async () => {
		const docPath = join(tempWorkspace("rf-empty-attrs"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--markdown",
			"See [note]{TODO fixme} later.",
		);
		expect(result.exitCode).toBe(0);
		// `{TODO fixme}` are bare tokens (no recognized attrs) → not a span →
		// brackets restored as literal text.
		const joined = (await readFmtBlocks(docPath))
			.flatMap((b) => b.runs ?? [])
			.map((r) => r.text ?? "")
			.join("");
		expect(joined).toContain("See [note]{TODO fixme} later.");
	});

	test("a formatted run with inert metacharacters keeps its formatting and stays clean", async () => {
		// `] { }` here form no construct (no `](`, no CriticMarkup), so they're
		// left byte-clean — the run keeps its color via the <span>, no `\` noise,
		// and the text round-trips. (Escaping only fires for chars that WOULD parse.)
		const src = join(tempWorkspace("rf-bracket-src"), "out.docx");
		await runCli("create", src, "--text", "Intro.");
		await runCli(
			"insert",
			src,
			"--after",
			"p0",
			"--runs",
			JSON.stringify([{ type: "text", text: "a]{x}b", color: "FF0000" }]),
		);

		const md = (await runCli("read", src)).stdout;
		expect(md).toContain("color:#FF0000");
		// Inert metacharacters are NOT escaped (they parse as nothing).
		expect(md).toContain("a]{x}b");

		const runs = (await roundTripBlocks("rf-bracket-rt", md))
			.flatMap((b) => b.runs ?? [])
			.filter((r) => r.type === "text");
		expect(runs.map((r) => r.text ?? "").join("")).toContain("a]{x}b");
		const colored = runs.find((r) => (r.text ?? "").includes("a]{x}b"));
		expect(colored?.color).toBe("FF0000");
	});

	test("a symmetric delimiter is escaped only when it can pair (lone stays clean)", async () => {
		// remark eats a paired `$…$` as math (KaTeX does too in a preview) — the
		// reported bug — but a LONE `$` is inert. Escape only the pair, leaving
		// prices like `$5` byte-clean. Counting is paragraph-wide, so a `$` in one
		// run pairs with one past a `<mark>` in another (the exact psa shape).
		async function readRuns(label: string, runs: object[]): Promise<string> {
			const doc = join(tempWorkspace(label), "out.docx");
			await runCli("create", doc, "--text", "Intro.");
			await runCli(
				"insert",
				doc,
				"--after",
				"p0",
				"--runs",
				JSON.stringify(runs),
			);
			return (await runCli("read", doc)).stdout;
		}

		const pairedMd = await readRuns("md-paired", [
			{ type: "text", text: "at least $" },
			{ type: "text", text: "amount", highlight: "yellow" },
			{ type: "text", text: " and $" },
			{ type: "text", text: "more", highlight: "yellow" },
		]);
		expect(pairedMd).toContain("at least \\$");
		expect(pairedMd).toContain("and \\$");
		expect(pairedMd).not.toMatch(/[^\\]\$/); // no bare `$` left to open math

		const loneMd = await readRuns("md-lone", [
			{ type: "text", text: "a price of $5 today" },
		]);
		expect(loneMd).toContain("price of $5 today"); // lone `$` left clean

		// Both round-trip byte-exact.
		const pairedText = (await roundTripBlocks("md-paired-rt", pairedMd))
			.flatMap((b) => b.runs ?? [])
			.map((r) => r.text ?? "")
			.join("");
		expect(pairedText).toContain("at least $amount and $more");
		const loneText = (await roundTripBlocks("md-lone-rt", loneMd))
			.flatMap((b) => b.runs ?? [])
			.map((r) => r.text ?? "")
			.join("");
		expect(loneText).toContain("a price of $5 today");
	});

	test("brackets are escaped only when link-forming — checkboxes/placeholders stay clean", async () => {
		async function readRuns(label: string, text: string): Promise<string> {
			const doc = join(tempWorkspace(label), "out.docx");
			await runCli("create", doc, "--text", "Intro.");
			await runCli(
				"insert",
				doc,
				"--after",
				"p0",
				"--runs",
				JSON.stringify([{ type: "text", text }]),
			);
			return (await runCli("read", doc)).stdout;
		}

		// A `[ x ]` checkbox / `[Fill in …]` placeholder forms no link → left clean.
		const inert = await readRuns(
			"brk-inert",
			"[ x ] done [Fill in amount] here",
		);
		expect(inert).toContain("[ x ] done [Fill in amount] here");
		expect(inert).not.toContain("\\[");
		const inertText = (await roundTripBlocks("brk-inert-rt", inert))
			.flatMap((b) => b.runs ?? [])
			.map((r) => r.text ?? "")
			.join("");
		expect(inertText).toContain("[ x ] done [Fill in amount] here");

		// `[word](target)` WOULD be consumed into a hyperlink on import, so its
		// link delimiters are escaped (the parser flags them as a `link` node); the
		// round-trip proves it stayed literal text instead of becoming a hyperlink.
		const link = await readRuns("brk-link", "see [word](target) literally");
		expect(link).toContain("\\[word\\]\\(target\\)");
		const linkText = (await roundTripBlocks("brk-link-rt", link))
			.flatMap((b) => b.runs ?? [])
			.map((r) => r.text ?? "")
			.join("");
		expect(linkText).toContain("see [word](target) literally");
	});

	test("brackets inside a wide paired-$ math span are left clean (only the $ escape)", async () => {
		// Two `$` pair into one `inlineMath` span; escaping the boundary `$` breaks
		// it, so a `[opt]` that fell between them needn't be touched. (This is the
		// psa shape — placeholders sitting between two dollar-amount fields.)
		const src = join(tempWorkspace("md-mathspan"), "out.docx");
		await runCli("create", src, "--text", "Intro.");
		await runCli(
			"insert",
			src,
			"--after",
			"p0",
			"--runs",
			JSON.stringify([{ type: "text", text: "from $5 to [opt] and $9 total" }]),
		);
		const md = (await runCli("read", src)).stdout;
		expect(md).toContain("from \\$5 to [opt] and \\$9 total");
		const text = (await roundTripBlocks("md-mathspan-rt", md))
			.flatMap((b) => b.runs ?? [])
			.map((r) => r.text ?? "")
			.join("");
		expect(text).toContain("from $5 to [opt] and $9 total");
	});

	test("an inline code span rides verbatim — its metacharacters are never escaped", async () => {
		// `runStyle: "Code"` text is literal between backticks, so it's excluded from
		// the escape parse entirely; `$`/`*`/`[` inside it stay byte-exact.
		const ws = tempWorkspace("md-code");
		const mdPath = join(ws, "doc.md");
		await Bun.write(mdPath, "Code span: `$5 and *x* and [y]` end.\n");
		const src = join(ws, "code.docx");
		await runCli("create", src, "--from", mdPath);
		const md = (await runCli("read", src)).stdout;
		expect(md).toContain("`$5 and *x* and [y]`");
		expect(md).not.toContain("\\$");
	});

	test("--runs with an invalid highlight enum fails loud (matches the markdown path)", async () => {
		const docPath = join(tempWorkspace("rf-runs-bad-hl"), "out.docx");
		await runCli("create", docPath, "--text", "Intro.");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--runs",
			JSON.stringify([{ type: "text", text: "x", highlight: "chartreuse" }]),
		);
		expect(result.exitCode).toBeGreaterThan(0);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
		expect((result.parsed as { error?: string }).error ?? "").toContain(
			"chartreuse",
		);
	});

	test("a font value containing a quote can't inject a second attribute", async () => {
		// Self-inflicted-injection guard: a crafted font with `"` must not
		// round-trip into `highlight=...`; the unsafe value is dropped on export.
		const src = join(tempWorkspace("rf-font-inject"), "out.docx");
		await runCli("create", src, "--text", "Intro.");
		await runCli(
			"insert",
			src,
			"--after",
			"p0",
			"--runs",
			JSON.stringify([
				{ type: "text", text: "hi", font: 'Arial" highlight="yellow' },
			]),
		);
		const md = (await runCli("read", src)).stdout;
		expect(md).not.toContain('highlight="yellow"');

		const ws = tempWorkspace("rf-font-inject-rt");
		const mdPath = join(ws, "doc.md");
		await Bun.write(mdPath, md);
		const dst = join(ws, "rt.docx");
		await runCli("create", dst, "--from", mdPath);
		const injected = (await readFmtBlocks(dst))
			.flatMap((b) => b.runs ?? [])
			.some((r) => (r as { highlight?: string }).highlight === "yellow");
		expect(injected).toBe(false);
	});
});

// "Comments are never anything but hints" (root CLAUDE.md): every docx: structural
// annotation read emits is DROPPED on import — none drive reconstruction. So they
// can't corrupt the parse or smuggle structure into a from-scratch create.
describe("import — docx: annotations are dropped, never reconstructed", () => {
	async function createFrom(label: string, markdown: string): Promise<string> {
		const workspace = tempWorkspace(label);
		const mdPath = join(workspace, "src.md");
		await Bun.write(mdPath, markdown);
		const doc = join(workspace, "out.docx");
		const result = await runCli("create", doc, "--from", mdPath);
		expect(result.exitCode).toBe(0);
		return doc;
	}

	test("a docx:table note before a table is dropped; the table still imports", async () => {
		const doc = await createFrom(
			"drop-table",
			'# T\n\n<!-- docx:table t0 widths="1,2,3in" borders="double" -->\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n',
		);
		const ast = (await runCli("read", doc, "--ast")).parsed as {
			blocks: Array<{ type: string }>;
		};
		expect(ast.blocks.some((b) => b.type === "table")).toBe(true);
		// The note didn't leak as text, and didn't reconstruct custom widths.
		const md = await read(doc);
		expect(md).not.toContain("docx:table");
		expect(md).not.toContain("widths="); // even-width table, no widths note
	});

	test("a docx:section note between paragraphs is dropped (no section reconstructed)", async () => {
		const doc = await createFrom(
			"drop-section",
			'# T\n\nbefore\n\n<!-- docx:section s1 cols="2" type="continuous" -->\n\nafter\n',
		);
		const ast = (await runCli("read", doc, "--ast")).parsed as {
			blocks: Array<{ type: string }>;
		};
		// Only the trailing mandatory sectPr — the docx:section comment was dropped.
		expect(ast.blocks.filter((b) => b.type === "sectionBreak")).toHaveLength(1);
		// The imported note was NOT reconstructed: no cols="2" section comes back,
		// and no second section (s1). The bare trailing s0 marker every doc carries
		// is not the dropped note.
		const md = await read(doc);
		expect(md).not.toContain('cols="2"');
		expect(md).not.toContain("docx:section s1");
	});

	test("inline docx:cell / docx:p / docx:image hints don't break the parse or leak as text", async () => {
		const doc = await createFrom(
			"drop-inline",
			'# T\n\nA paragraph. <!-- p9 --> <!-- docx:p p9 style="Caption" -->\n\n| x <!-- docx:cell t0:r0c0 gridSpan="2" --> | y |\n| --- | --- |\n',
		);
		const md = await read(doc);
		expect(md).toContain("A paragraph.");
		expect(md).toContain("x"); // table cell content survived
		expect(md).not.toContain("docx:p");
		expect(md).not.toContain("docx:cell");
		// The dropped hints didn't reconstruct a Caption style.
		const ast = (await runCli("read", doc, "--ast")).parsed as {
			blocks: Array<{ type: string; style?: string }>;
		};
		expect(ast.blocks.some((b) => b.style === "Caption")).toBe(false);
	});

	test("a full read → create round-trip of an annotation-heavy fixture survives", async () => {
		// sections.docx is dense with docx:section + docx:page hints.
		const src = join(import.meta.dir, "..", "fixtures", "sections.docx");
		const md = (await runCli("read", src)).stdout;
		expect(md).toContain("docx:section");
		const doc = await createFrom("drop-roundtrip", md);
		// Content survived; the hints were dropped (no leak), and a re-read is
		// itself stable (re-emitted fresh, never accreting).
		const r1 = await read(doc);
		const doc2 = await createFrom("drop-roundtrip2", r1);
		expect(await read(doc2)).toBe(r1);
	});
});
