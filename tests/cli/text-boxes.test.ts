import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";
import {
	buildRawDoc,
	readDocumentXml,
	readMarkdown,
	storyXml,
	wordTextBoxParagraphXml,
} from "./helpers";

/**
 * Text boxes (issue #4). Word writes every modern text box as an
 * `<mc:AlternateContent>` pair — a `wps` shape under `<mc:Choice>` holding the
 * story, and a VML `<w:pict>` twin under `<mc:Fallback>` holding the SAME
 * story. The reader used to drop the whole wrapper, so `replace --all` could
 * report two occurrences and exit 0 while a third sat in a text box on page 1.
 *
 * The fixture is authored through the CLI: `create` + `raw insert` of the exact
 * shape Word emits (both branches), so the tests exercise the gate pipeline
 * AND the reader. Every mutating verb addresses the story through `tbxN:pK`.
 */

/** The harness runs `find` with `--json`; pull the locator list back out. */
function findLocators(stdout: string): string[] {
	const parsed = JSON.parse(stdout) as { matches: { locator: string }[] };
	return parsed.matches.map((match) => match.locator);
}

const BOX_LINES = [
	"CONFIDENTIAL",
	"Confidential - Acme Corporation internal use only.",
];

describe("text boxes — the issue #4 document", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("text-boxes");
		docPath = join(workspace, "out.docx");
		await runCli(
			"create",
			docPath,
			"--text",
			"Acme Corporation shall deliver the goods by Friday.",
		);
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Payment is due to Acme within thirty days.",
		);
		const fragment = join(workspace, "box.xml");
		await Bun.write(fragment, wordTextBoxParagraphXml(BOX_LINES));
		const inserted = await runCli(
			"raw",
			"insert",
			docPath,
			"--after",
			"p0",
			"--xml-file",
			fragment,
		);
		expect(inserted.exitCode).toBe(0);
	});

	test("read renders the story after its anchor, bracketed by docx:textbox hints", async () => {
		const markdown = await readMarkdown(docPath);
		expect(markdown).toContain(
			'<!-- docx:textbox tbx0 anchor="p1" wrap="square" align="right" -->',
		);
		expect(markdown).toContain("CONFIDENTIAL <!-- tbx0:p0 -->");
		expect(markdown).toContain(
			"Confidential - Acme Corporation internal use only. <!-- tbx0:p1 -->",
		);
		expect(markdown).toContain("<!-- docx:textbox-end tbx0 -->");
		// Story comes between the anchor (p1, empty — so not printed) and p2.
		const storyAt = markdown.indexOf("docx:textbox tbx0");
		expect(storyAt).toBeGreaterThan(markdown.indexOf("<!-- p0 -->"));
		expect(storyAt).toBeLessThan(markdown.indexOf("<!-- p2 -->"));
	});

	test("read --ast carries the story as a textBox run with tbxN:pK blocks", async () => {
		const result = await runCli("read", docPath, "--ast");
		expect(result.exitCode).toBe(0);
		const ast = JSON.parse(result.stdout);
		const anchor = ast.blocks[1];
		expect(anchor.id).toBe("p1");
		expect(anchor.runs[0]).toMatchObject({
			type: "textBox",
			id: "tbx0",
			floating: true,
			wrap: "square",
			align: "right",
		});
		expect(
			anchor.runs[0].blocks.map((block: { id: string }) => block.id),
		).toEqual(["tbx0:p0", "tbx0:p1"]);
	});

	test("find sees the third occurrence a human reads first", async () => {
		const result = await runCli("find", docPath, "Acme");
		expect(result.exitCode).toBe(0);
		expect(findLocators(result.stdout)).toEqual([
			"p0:0-4",
			"tbx0:p1:15-19",
			"p2:18-22",
		]);
	});

	test("replace --all reaches the text box — no silent under-application", async () => {
		const result = await runCli(
			"replace",
			docPath,
			"Acme",
			"Globex",
			"--all",
			"--verbose",
		);
		expect(result.exitCode).toBe(0);
		expect(JSON.parse(result.stdout)).toMatchObject({
			totalMatches: 3,
			replaced: 3,
		});
		const markdown = await readMarkdown(docPath);
		expect(markdown).not.toContain("Acme");
		expect(markdown).toContain(
			"Confidential - Globex Corporation internal use only. <!-- tbx0:p1 -->",
		);
	});

	test("the Fallback twin is re-synced from the Choice story on save", async () => {
		await runCli("replace", docPath, "Acme", "Globex", "--all");
		const xml = await readDocumentXml(docPath);
		const fallback = xml.match(/<mc:Fallback>.*?<\/mc:Fallback>/s)?.[0] ?? "";
		const fallbackText = fallback.replace(/<[^>]+>/g, "");
		expect(fallbackText).toContain("Globex Corporation internal use only.");
		expect(fallbackText).not.toContain("Acme");
		// The whole wrapper survives — both branches, the shape geometry too.
		expect(xml).toContain('<mc:Choice Requires="wps">');
		expect(xml).toContain('<wps:cNvSpPr txBox="1"/>');
	});

	test("edit / comments add / insert / delete address the story like a cell", async () => {
		const edited = await runCli(
			"edit",
			docPath,
			"--at",
			"tbx0:p0",
			"--text",
			"TOP SECRET",
		);
		expect(edited.exitCode).toBe(0);
		expect(JSON.parse(edited.stdout)).toMatchObject({ locator: "tbx0:p0" });

		// Word discards a comment (or footnote) inside a text box on its next
		// save — verified against Word for Mac 365 — so both are refused up front
		// and pointed at the anchor paragraph.
		const commented = await runCli(
			"comments",
			"add",
			docPath,
			"--at",
			"tbx0:p1:15-19",
			"--text",
			"Confirm the party name",
		);
		expect(commented.exitCode).not.toBe(0);
		expect(commented.stdout).toContain("UNSUPPORTED");
		expect(commented.stdout).toContain("anchors the box");
		const byPhrase = await runCli(
			"comments",
			"add",
			docPath,
			"--anchor",
			"internal use only",
			"--text",
			"x",
		);
		expect(byPhrase.exitCode).not.toBe(0);
		expect(byPhrase.stdout).toContain("only inside a text box");
		const noted = await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"tbx0:p0",
			"--text",
			"src",
		);
		expect(noted.exitCode).not.toBe(0);
		expect(noted.stdout).toContain("UNSUPPORTED");
		expect((await runCli("comments", "list", docPath)).stdout.trim()).toBe(
			"[]",
		);

		const inserted = await runCli(
			"insert",
			docPath,
			"--after",
			"tbx0:p1",
			"--text",
			"Do not forward.",
		);
		expect(inserted.exitCode).toBe(0);
		expect(inserted.stdout).toContain("tbx0:p2");

		const deleted = await runCli("delete", docPath, "--at", "tbx0:p1");
		expect(deleted.exitCode).toBe(0);

		const markdown = await readMarkdown(docPath);
		expect(markdown).toContain("TOP SECRET <!-- tbx0:p0 -->");
		expect(markdown).toContain("Do not forward. <!-- tbx0:p1 -->");
		expect(markdown).not.toContain("internal use only");
		// Body untouched.
		expect(markdown).toContain(
			"Acme Corporation shall deliver the goods by Friday. <!-- p0 -->",
		);
	});

	test("deleting a box's last paragraph blanks it — a story keeps one <w:p>", async () => {
		await runCli("delete", docPath, "--at", "tbx0:p1");
		const result = await runCli("delete", docPath, "--at", "tbx0:p0");
		expect(result.exitCode).toBe(0);
		const ast = JSON.parse((await runCli("read", docPath, "--ast")).stdout);
		expect(ast.blocks[1].runs[0].blocks).toHaveLength(1);
		expect(ast.blocks[1].runs[0].blocks[0].runs).toEqual([]);
	});

	test("wc counts what read shows — text boxes included; tbxN counts one box", async () => {
		const whole = JSON.parse((await runCli("wc", docPath, "--json")).stdout);
		expect(whole.words).toBe(9 + 7 + 1 + 7);
		const range = JSON.parse(
			(await runCli("wc", docPath, "p0-p2", "--json")).stdout,
		);
		expect(range.words).toBe(whole.words);
		const box = JSON.parse(
			(await runCli("wc", docPath, "tbx0", "--json")).stdout,
		);
		expect(box).toMatchObject({ scope: "textBox", words: 1 + 7 });
		const line = JSON.parse(
			(await runCli("wc", docPath, "tbx0:p1", "--json")).stdout,
		);
		expect(line).toMatchObject({ scope: "paragraph", words: 7 });
	});

	test("a bad text-box locator fails BLOCK_NOT_FOUND, not silently", async () => {
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"tbx3:p0",
			"--text",
			"x",
		);
		expect(result.exitCode).not.toBe(0);
		expect(result.stdout).toContain("BLOCK_NOT_FOUND");
		const missingBox = await runCli("wc", docPath, "tbx3");
		expect(missingBox.exitCode).not.toBe(0);
	});

	test("a bare tbxN on edit/delete names the paragraph form to use instead", async () => {
		const result = await runCli("edit", docPath, "--at", "tbx0", "--text", "x");
		expect(result.exitCode).not.toBe(0);
		expect(result.stdout).toContain("whole text box");
		expect(result.stdout).toContain("tbx0:p0");
	});

	test("the ECMA schema gate accepts the mutated document", async () => {
		await runCli("replace", docPath, "Acme", "Globex", "--all");
		const result = await runCli("validate", docPath);
		expect(result.exitCode).toBe(0);
	});
});

describe("text boxes — a VML-only (Fallback-shaped) box reads the same way", () => {
	test("a bare <w:pict> text box surfaces as tbx0 with no floating geometry", async () => {
		const docPath = await buildRawDoc(
			`<w:p><w:r><w:t>Body.</w:t></w:r></w:p><w:p><w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:rect><v:textbox><w:txbxContent>${storyXml(["Legacy box"])}</w:txbxContent></v:textbox></v:rect></w:pict></w:r></w:p>`,
			"vml-box",
		);
		const markdown = await readMarkdown(docPath);
		expect(markdown).toContain('<!-- docx:textbox tbx0 anchor="p1" -->');
		expect(markdown).toContain("Legacy box <!-- tbx0:p0 -->");
		const found = await runCli("find", docPath, "Legacy");
		expect(findLocators(found.stdout)).toEqual(["tbx0:p0:0-6"]);
	});
});

describe("text boxes — the anchor paragraph and the box survive each other", () => {
	let docPath: string;
	beforeEach(async () => {
		docPath = await buildRawDoc(
			wordTextBoxParagraphXml(BOX_LINES, { leadingText: "Anchor text. " }) +
				`<w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p>`,
			"tb-anchor",
		);
	});

	test("replace --at tbxN:pK is a valid scope (find → replace pipeline)", async () => {
		const result = await runCli(
			"replace",
			docPath,
			"Acme",
			"Globex",
			"--at",
			"tbx0:p1",
		);
		expect(result.exitCode).toBe(0);
		expect(await readMarkdown(docPath)).toContain(
			"Confidential - Globex Corporation internal use only. <!-- tbx0:p1 -->",
		);
	});

	test("replacing the anchor run's own text keeps the box (trailing zero-width marker)", async () => {
		const result = await runCli("replace", docPath, "Anchor", "Lead");
		expect(result.exitCode).toBe(0);
		const markdown = await readMarkdown(docPath);
		expect(markdown).toContain("Lead text. <!-- p0 -->");
		expect(markdown).toContain("docx:textbox tbx0");
		const xml = await readDocumentXml(docPath);
		expect(xml.match(/<mc:AlternateContent/g)).toHaveLength(1);
		expect(xml.match(/<mc:Fallback/g)).toHaveLength(1);
	});

	test("a whole-paragraph edit of the anchor keeps the box — untracked and tracked", async () => {
		const edited = await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--text",
			"See the stamp.",
		);
		expect(edited.exitCode).toBe(0);
		let markdown = await readMarkdown(docPath);
		expect(markdown).toContain("See the stamp. <!-- p0 -->");
		expect(markdown).toContain("CONFIDENTIAL <!-- tbx0:p0 -->");

		const tracked = await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--text",
			"See the red stamp.",
			"--track",
		);
		expect(tracked.exitCode).toBe(0);
		markdown = await readMarkdown(docPath);
		expect(markdown).toContain("See the red stamp. <!-- p0 -->");
		expect(markdown).toContain("CONFIDENTIAL <!-- tbx0:p0 -->");
		const xml = await readDocumentXml(docPath);
		// The shape itself is never wrapped in a revision by a text edit.
		expect(xml).not.toMatch(
			/<w:(ins|del)[^>]*>\s*<w:r>(?:(?!<\/w:r>).)*<mc:AlternateContent/s,
		);
	});

	test("a tracked delete of the anchor hides the box in the accepted view and reject restores it", async () => {
		const deleted = await runCli("delete", docPath, "--at", "p0", "--track");
		expect(deleted.exitCode).toBe(0);
		const xml = await readDocumentXml(docPath);
		// The story keeps live <w:t>; the deletion is the <w:del> around the anchor run.
		expect(xml).toMatch(/<w:del[^>]*>.*<mc:AlternateContent/s);
		expect(xml).not.toMatch(
			/<w:txbxContent>(?:(?!<\/w:txbxContent>).)*<w:delText/s,
		);
		const accepted = await readMarkdown(docPath);
		expect(accepted).not.toContain("docx:textbox");
		const current = (await runCli("read", docPath, "--current")).stdout;
		expect(current).toContain("docx:textbox tbx0");
		const found = JSON.parse((await runCli("find", docPath, "Acme")).stdout);
		expect(found.matches).toHaveLength(0);
		expect((await runCli("validate", docPath)).exitCode).toBe(0);

		const rejected = await runCli("track-changes", "reject", docPath, "--all");
		expect(rejected.exitCode).toBe(0);
		expect(await readMarkdown(docPath)).toContain(
			"CONFIDENTIAL <!-- tbx0:p0 -->",
		);
	});

	test("raw insert INTO a box runs the schema gate against the story", async () => {
		const workspace = tempWorkspace("tb-raw-gate");
		const bad = join(workspace, "bad.xml");
		await Bun.write(
			bad,
			`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:jc w:val="nonsense-value"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
		);
		const before = await readDocumentXml(docPath);
		const result = await runCli(
			"raw",
			"insert",
			docPath,
			"--after",
			"tbx0:p0",
			"--xml-file",
			bad,
		);
		expect(result.exitCode).not.toBe(0);
		expect(result.stdout).toContain("VALIDATION_FAILED");
		expect(await readDocumentXml(docPath)).toBe(before);
	});

	test("a box anchored in a code-block paragraph still renders its story", async () => {
		await runCli("edit", docPath, "--at", "p0", "--style", "CodeBlock");
		const markdown = await readMarkdown(docPath);
		expect(markdown).toContain("```");
		expect(markdown).toContain("docx:textbox tbx0");
		expect(markdown).toContain("CONFIDENTIAL <!-- tbx0:p0 -->");
	});

	test("a cross-paragraph find enters the story but never joins it to the flow", async () => {
		const inside = JSON.parse(
			(await runCli("find", docPath, "CONFIDENTIAL\\nConfidential")).stdout,
		);
		expect(inside.matches[0]).toMatchObject({
			startBlockId: "tbx0:p0",
			endBlockId: "tbx0:p1",
		});
		const across = JSON.parse(
			(await runCli("find", docPath, "Anchor text. \\nCONFIDENTIAL")).stdout,
		);
		expect(across.matches).toHaveLength(0);
	});
});
