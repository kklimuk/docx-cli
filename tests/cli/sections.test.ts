import { beforeEach, describe, expect, test } from "bun:test";
import { copyFileSync } from "node:fs";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import JSZip from "jszip";
import { runCli, tempWorkspace } from "./harness";
import { readDocumentXml, readMarkdown, trackedKinds } from "./helpers";

const SECTIONS_FIXTURE = join(
	import.meta.dir,
	"..",
	"fixtures",
	"sections.docx",
);

type SectionBlock = {
	id: string;
	type: "sectionBreak";
	columns?: number;
	sectionType?: string;
	pageWidth?: number;
	pageHeight?: number;
	pageOrientation?: "portrait" | "landscape";
	marginTop?: number;
	marginRight?: number;
	marginBottom?: number;
	marginLeft?: number;
};

type ParagraphBlock = {
	id: string;
	type: "paragraph";
	runs?: Array<{ type: string; text?: string }>;
};

type Block = SectionBlock | ParagraphBlock | { id: string; type: string };

describe("sections via the sections verb / edit / delete", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("sections");
		docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "First");
	});

	test("sections --at p0 emits a sentinel paragraph + new sN block", async () => {
		// p0 is the doc's first block, so the wrap omits the leading break and
		// inserts only the trailing cols=2 break — the structure the removed
		// `insert --section` used to produce.
		await runCli(
			"sections",
			docPath,
			"--at",
			"p0",
			"--columns",
			"2",
			"--type",
			"continuous",
		);
		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as { blocks: Block[] };
		const sections = doc.blocks.filter(
			(block): block is SectionBlock => block.type === "sectionBreak",
		);
		expect(sections).toHaveLength(2);
		expect(sections[0]).toMatchObject({
			id: "s0",
			columns: 2,
			sectionType: "continuous",
		});
		expect(sections[1]).toMatchObject({ id: "s1" });
		expect(sections[1]?.columns).toBeUndefined();
	});

	test("insert no longer creates sections — it redirects to `docx sections`", async () => {
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--section",
			"--columns",
			"2",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
		expect((result.parsed as { error: string }).error).toContain(
			"docx sections",
		);
	});

	test("edit --at sN updates columns + type on the trailing section", async () => {
		await runCli(
			"edit",
			docPath,
			"--at",
			"s0",
			"--columns",
			"3",
			"--type",
			"continuous",
		);
		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as { blocks: Block[] };
		const section = doc.blocks.find(
			(block): block is SectionBlock => block.type === "sectionBreak",
		);
		expect(section).toMatchObject({
			id: "s0",
			columns: 3,
			sectionType: "continuous",
		});
	});

	test("edit --at sN supports updating only one property at a time", async () => {
		await runCli("edit", docPath, "--at", "s0", "--columns", "2");
		const firstRead = await runCli("read", docPath, "--ast");
		const after1 = (firstRead.parsed as { blocks: Block[] }).blocks.find(
			(block): block is SectionBlock => block.type === "sectionBreak",
		);
		expect(after1?.columns).toBe(2);
		expect(after1?.sectionType).toBeUndefined();

		await runCli("edit", docPath, "--at", "s0", "--type", "nextPage");
		const secondRead = await runCli("read", docPath, "--ast");
		const after2 = (secondRead.parsed as { blocks: Block[] }).blocks.find(
			(block): block is SectionBlock => block.type === "sectionBreak",
		);
		expect(after2?.columns).toBe(2);
		expect(after2?.sectionType).toBe("nextPage");
	});

	test("edit --at sN rejects --text/--runs", async () => {
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"s0",
			"--text",
			"nope",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("edit --at pN rejects --columns/--type", async () => {
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--text",
			"x",
			"--columns",
			"2",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("edit --at sN rejects invalid --type", async () => {
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"s0",
			"--type",
			"bogus",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("delete --at sN removes an inline sectPr; the sentinel paragraph remains", async () => {
		await runCli("sections", docPath, "--at", "p0", "--columns", "2");
		const beforeRead = await runCli("read", docPath, "--ast");
		const before = (beforeRead.parsed as { blocks: Block[] }).blocks;
		const beforeSections = before.filter((b) => b.type === "sectionBreak");
		expect(beforeSections).toHaveLength(2);

		const deleteResult = await runCli("delete", docPath, "--at", "s0");
		expect(deleteResult.exitCode).toBe(0);

		const afterRead = await runCli("read", docPath, "--ast");
		const after = (afterRead.parsed as { blocks: Block[] }).blocks;
		const afterSections = after.filter((b) => b.type === "sectionBreak");
		expect(afterSections).toHaveLength(1);
		// sentinel paragraph survives — only the sectPr was stripped.
		const paragraphs = after.filter((b) => b.type === "paragraph");
		expect(paragraphs).toHaveLength(2);
	});

	test("delete --at sN rejects the trailing section break", async () => {
		const result = await runCli("delete", docPath, "--at", "s0");
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("wc reports a sections count alongside words", async () => {
		await runCli("sections", docPath, "--at", "p0", "--columns", "2");
		const wc = await runCli("wc", docPath);
		expect(wc.parsed).toMatchObject({
			scope: "document",
			words: 1,
			sections: 2,
		});
	});
});

describe("sections under track-changes", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("sections-tc");
		docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Hello");
	});

	test("sections under tracking + reject removes the sentinel paragraph entirely", async () => {
		await runCli("track-changes", docPath, "on");
		await runCli(
			"sections",
			docPath,
			"--at",
			"p0",
			"--columns",
			"2",
			"--author",
			"Alice",
		);
		const beforeRead = await runCli("read", docPath, "--ast");
		const beforeBlocks = (beforeRead.parsed as { blocks: Block[] }).blocks;
		expect(beforeBlocks.filter((b) => b.type === "sectionBreak")).toHaveLength(
			2,
		);
		expect(beforeBlocks.filter((b) => b.type === "paragraph")).toHaveLength(2);

		await runCli("track-changes", "reject", docPath, "--all");

		const afterRead = await runCli("read", docPath, "--ast");
		const afterBlocks = (afterRead.parsed as { blocks: Block[] }).blocks;
		// Sentinel paragraph + its inline sectPr both gone; only the original
		// content paragraph and the trailing sectPr remain.
		expect(afterBlocks.filter((b) => b.type === "sectionBreak")).toHaveLength(
			1,
		);
		expect(afterBlocks.filter((b) => b.type === "paragraph")).toHaveLength(1);
	});

	test("delete --at pN under tracking + accept merges the deleted paragraph with the next", async () => {
		await runCli("insert", docPath, "--after", "p0", "--text", "Second");
		await runCli("track-changes", docPath, "on");
		await runCli("delete", docPath, "--at", "p0", "--author", "Alice");

		const list = await runCli("track-changes", "list", docPath);
		const changes = list.parsed as Array<{ id: string; kind: string }>;
		// Two trackings on p0: run-level <w:del> wrapping "Original body",
		// and paragraph-mark <w:del> on the paragraph mark itself.
		expect(changes).toHaveLength(2);
		expect(changes.map((c) => c.kind)).toEqual(["del", "del"]);

		await runCli("track-changes", "accept", docPath, "--all");

		const afterRead = await runCli("read", docPath, "--ast");
		const paragraphs = (afterRead.parsed as { blocks: Block[] }).blocks.filter(
			(b): b is ParagraphBlock => b.type === "paragraph",
		);
		// p0 is gone (its runs deleted, its paragraph mark merged with p1's
		// content). The single remaining paragraph carries "Second".
		expect(paragraphs).toHaveLength(1);
		const text = (paragraphs[0]?.runs ?? [])
			.map((run) => run.text ?? "")
			.join("");
		expect(text).toBe("Second");
	});

	test("edit --at sN under tracking emits a sectPrChange snapshot", async () => {
		await runCli("edit", docPath, "--at", "s0", "--columns", "1");
		await runCli("track-changes", docPath, "on");
		await runCli(
			"edit",
			docPath,
			"--at",
			"s0",
			"--columns",
			"3",
			"--type",
			"continuous",
			"--author",
			"Alice",
		);
		const list = await runCli("track-changes", "list", docPath);
		const changes = list.parsed as Array<{
			id: string;
			kind: string;
			author: string;
			blockId: string;
			prior?: { columns?: number; sectionType?: string };
			current?: { columns?: number; sectionType?: string };
		}>;
		expect(changes).toHaveLength(1);
		expect(changes[0]).toMatchObject({
			id: "tc0",
			kind: "sectPrChange",
			author: "Alice",
			blockId: "s0",
			prior: { columns: 1 },
			current: { columns: 3, sectionType: "continuous" },
		});
	});

	test("track-changes accept removes the sectPrChange but keeps the new properties", async () => {
		await runCli("edit", docPath, "--at", "s0", "--columns", "1");
		await runCli("track-changes", docPath, "on");
		await runCli(
			"edit",
			docPath,
			"--at",
			"s0",
			"--columns",
			"3",
			"--type",
			"continuous",
		);
		await runCli("track-changes", "accept", docPath, "--all");
		const read = await runCli("read", docPath, "--ast");
		const section = (read.parsed as { blocks: Block[] }).blocks.find(
			(b): b is SectionBlock => b.type === "sectionBreak",
		);
		expect(section).toMatchObject({
			id: "s0",
			columns: 3,
			sectionType: "continuous",
		});
		const list = await runCli("track-changes", "list", docPath);
		expect(list.parsed).toEqual([]);
	});

	test("track-changes reject restores prior section properties from snapshot", async () => {
		await runCli("edit", docPath, "--at", "s0", "--columns", "1");
		await runCli("track-changes", docPath, "on");
		await runCli(
			"edit",
			docPath,
			"--at",
			"s0",
			"--columns",
			"3",
			"--type",
			"continuous",
		);
		await runCli("track-changes", "reject", docPath, "--all");
		const read = await runCli("read", docPath, "--ast");
		const section = (read.parsed as { blocks: Block[] }).blocks.find(
			(b): b is SectionBlock => b.type === "sectionBreak",
		);
		expect(section).toMatchObject({ id: "s0", columns: 1 });
		expect(section?.sectionType).toBeUndefined();
	});

	// Per ECMA-376 §17.6.18, <w:sectPrChange> must be the LAST child of
	// <w:sectPr>. When tracking is on and we add a NEW property to a sectPr
	// that started empty, applyColumns / applySectionType must insert ahead
	// of the freshly-pushed sectPrChange, not after it.
	test("edit --at sN under tracking keeps sectPrChange as the last child", async () => {
		await runCli("track-changes", docPath, "on");
		await runCli(
			"edit",
			docPath,
			"--at",
			"s0",
			"--columns",
			"2",
			"--type",
			"continuous",
		);
		const pkg = await Pkg.open(docPath);
		const documentXml = await pkg.readText("word/document.xml");
		const sectPrMatch = documentXml.match(
			/<w:sectPr\b[^>]*>([\s\S]*?)<\/w:sectPr>/,
		);
		expect(sectPrMatch).not.toBeNull();
		const inner = sectPrMatch?.[1] ?? "";
		const childTags = Array.from(inner.matchAll(/<w:(\w+)\b/g)).map(
			(match) => match[1],
		);
		// Only top-level children of the live sectPr — exclude tags nested inside
		// the sectPrChange snapshot (those appear AFTER the sectPrChange opener).
		const changeIndex = childTags.indexOf("sectPrChange");
		expect(changeIndex).toBeGreaterThanOrEqual(0);
		const liveChildren = childTags.slice(0, changeIndex + 1);
		// The invariant per ECMA-376 §17.6.18: sectPrChange must be the LAST
		// child of the live sectPr. Don't assert on the full list of siblings —
		// the create template legitimately seeds pgSz/pgMar/docGrid which would
		// make a strict equality brittle.
		expect(liveChildren[liveChildren.length - 1]).toBe("sectPrChange");
		expect(liveChildren).toContain("type");
		expect(liveChildren).toContain("cols");
	});
});

describe("sections.docx fixture", () => {
	test("AST surfaces every section type with the expected properties", async () => {
		const read = await runCli("read", SECTIONS_FIXTURE, "--ast");
		const doc = read.parsed as { blocks: Block[] };
		const sections = doc.blocks.filter(
			(block): block is SectionBlock => block.type === "sectionBreak",
		);
		expect(sections).toEqual([
			{ id: "s0", type: "sectionBreak", columns: 1, sectionType: "continuous" },
			{ id: "s1", type: "sectionBreak", columns: 2, sectionType: "continuous" },
			{ id: "s2", type: "sectionBreak", columns: 2, sectionType: "nextColumn" },
			{ id: "s3", type: "sectionBreak", columns: 2, sectionType: "continuous" },
			{ id: "s4", type: "sectionBreak", columns: 1, sectionType: "nextPage" },
			{ id: "s5", type: "sectionBreak", columns: 1, sectionType: "evenPage" },
			{ id: "s6", type: "sectionBreak", columns: 1, sectionType: "oddPage" },
			// The trailing (mandatory) sectPr carries the document-wide page
			// geometry; the inline breaks above inherit it and surface none.
			{
				id: "s7",
				type: "sectionBreak",
				columns: 2,
				sectionType: "continuous",
				pageWidth: 12240,
				pageHeight: 15840,
				marginTop: 1440,
				marginRight: 1440,
				marginBottom: 1440,
				marginLeft: 1440,
			},
		]);
	});

	test("track-changes list surfaces the sectPrChange with prior+current diff", async () => {
		const list = await runCli("track-changes", "list", SECTIONS_FIXTURE);
		const changes = list.parsed as Array<{
			id: string;
			kind: string;
			author: string;
			date: string;
			blockId: string;
			prior?: { columns?: number; sectionType?: string };
			current?: { columns?: number; sectionType?: string };
		}>;
		expect(changes).toHaveLength(1);
		expect(changes[0]).toMatchObject({
			id: "tc0",
			kind: "sectPrChange",
			author: "Reviewer",
			date: "2026-05-06T10:00:00Z",
			blockId: "s7",
			prior: { columns: 1, sectionType: "continuous" },
			current: { columns: 2, sectionType: "continuous" },
		});
	});

	test("wc reports words and the section count from the fixture", async () => {
		const wc = await runCli("wc", SECTIONS_FIXTURE);
		expect(wc.parsed).toMatchObject({
			scope: "document",
			sections: 8,
		});
		const words = (wc.parsed as { words: number }).words;
		expect(words).toBeGreaterThan(800);
	});

	test("track-changes accept on the fixture removes the snapshot, keeps current props", async () => {
		const workspace = tempWorkspace("sections-fixture-accept");
		const docPath = join(workspace, "sections.docx");
		copyFileSync(SECTIONS_FIXTURE, docPath);

		await runCli("track-changes", "accept", docPath, "--all");

		const read = await runCli("read", docPath, "--ast");
		const sections = (read.parsed as { blocks: Block[] }).blocks.filter(
			(block): block is SectionBlock => block.type === "sectionBreak",
		);
		expect(sections.at(-1)).toMatchObject({
			id: "s7",
			columns: 2,
			sectionType: "continuous",
		});
		const list = await runCli("track-changes", "list", docPath);
		expect(list.parsed).toEqual([]);
	});

	test("track-changes reject on the fixture restores the prior 1-column state", async () => {
		const workspace = tempWorkspace("sections-fixture-reject");
		const docPath = join(workspace, "sections.docx");
		copyFileSync(SECTIONS_FIXTURE, docPath);

		await runCli("track-changes", "reject", docPath, "--all");

		const read = await runCli("read", docPath, "--ast");
		const sections = (read.parsed as { blocks: Block[] }).blocks.filter(
			(block): block is SectionBlock => block.type === "sectionBreak",
		);
		expect(sections.at(-1)).toMatchObject({
			id: "s7",
			columns: 1,
			sectionType: "continuous",
		});
		const list = await runCli("track-changes", "list", docPath);
		expect(list.parsed).toEqual([]);
	});

	test("wc sN counts words in the section's content range", async () => {
		const wc = await runCli("wc", SECTIONS_FIXTURE, "s4");
		expect(wc.parsed).toMatchObject({
			scope: "section",
		});
		const words = (wc.parsed as { words: number }).words;
		expect(words).toBeGreaterThan(0);
	});

	test("per-section wc sums to the doc-level total", async () => {
		const docResult = await runCli("wc", SECTIONS_FIXTURE);
		const docWords = (docResult.parsed as { words: number; sections: number })
			.words;
		const sectionCount = (
			docResult.parsed as { words: number; sections: number }
		).sections;

		let perSectionSum = 0;
		for (let index = 0; index < sectionCount; index += 1) {
			const result = await runCli("wc", SECTIONS_FIXTURE, `s${index}`);
			perSectionSum += (result.parsed as { words: number }).words;
		}
		expect(perSectionSum).toBe(docWords);
	});

	test("wc returns BLOCK_NOT_FOUND for an out-of-range section locator", async () => {
		const result = await runCli("wc", SECTIONS_FIXTURE, "s99");
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({
			code: "BLOCK_NOT_FOUND",
		});
	});

	test("each sN locator resolves to a sectPr — no overlap with paragraph/table ids", async () => {
		const workspace = tempWorkspace("sections-fixture-locators");
		const docPath = join(workspace, "sections.docx");
		copyFileSync(SECTIONS_FIXTURE, docPath);

		// edit --at sN succeeds for every enumerated section; if any sN
		// resolved to a paragraph instead, this would error with USAGE.
		for (let index = 0; index < 8; index += 1) {
			const result = await runCli(
				"edit",
				docPath,
				"--at",
				`s${index}`,
				"--columns",
				"1",
			);
			expect(result.exitCode).toBe(0);
		}
	});
});

// Read-side rendering of section breaks + page geometry. These are read-time
// VISIBILITY hints (per "comments are never anything but hints") — the bare `---`
// is gone (it round-tripped as a thematic break, corrupting layout), and page
// geometry deviations ride a leading docx:page note.
describe("section breaks render as docx:section, not bare ---", () => {
	test("a mid-doc section is an own-line docx:section hint with cols/type", async () => {
		const md = await readMarkdown(SECTIONS_FIXTURE);
		expect(md).not.toContain("---");
		// The note renders at the section's START and states the locator range it
		// governs — the content BELOW it — so the off-by-one "which side?" trap is explicit.
		expect(md).toContain(
			'<!-- docx:section s1 cols="2" type="continuous" applies-to="p3..p7 (below)" -->',
		);
		// cols=1 is the default → suppressed; type still shown (with scope).
		expect(md).toContain(
			'<!-- docx:section s0 type="continuous" applies-to="p0..p2 (below)" -->',
		);
		// The trailing mandatory section break is suppressed entirely.
		expect(md).not.toContain("docx:section s7");
	});
});

describe("docx:layout hazard — tab alignment inside a multi-column section", () => {
	test("flags tab paragraphs governed by a cols>1 section, and only those", async () => {
		const workspace = tempWorkspace("layout-hazard");
		const docPath = join(workspace, "doc.docx");
		// p0 intro, p1/p2 tab-aligned, p3 outro — all single column to start.
		await runCli("create", docPath, "--text", "Intro");
		await runCli("insert", docPath, "--after", "p0", "--text", "Name\tLoc");
		await runCli("insert", docPath, "--after", "p1", "--text", "Role\tDate");
		await runCli("insert", docPath, "--after", "p2", "--text", "Outro");
		// Single column → no hazard even though p1/p2 have tabs.
		expect(await readMarkdown(docPath)).not.toContain("docx:layout");
		// Wrap the two tab paragraphs in a 2-column section.
		await runCli("sections", docPath, "--at", "p1-p2", "--columns", "2");
		const md = await readMarkdown(docPath);
		// Both tab paragraphs warn; the single-column intro/outro do not.
		expect(md.match(/docx:layout/g)?.length).toBe(2);
		expect(md).toContain('cols="2"');
		expect(md).toContain("render to verify");
	});

	test("flags a fragile right-alignment via a near-margin LEFT tab (résumé pattern)", async () => {
		const workspace = tempWorkspace("layout-lefttab");
		const docPath = join(workspace, "doc.docx");
		await runCli("create", docPath, "--text", "seed");
		// Inject the résumé template's fragile shape into p0: a LEFT tab stop near
		// the right margin + a tabbed "Org⇥City" line (no authoring verb sets tab
		// stops, so build it directly).
		const pkg = await Pkg.open(docPath);
		const xml = (await pkg.readText("word/document.xml")).replace(
			'<w:p><w:r><w:t xml:space="preserve">seed</w:t></w:r></w:p>',
			'<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="10000"/></w:tabs></w:pPr><w:r><w:t xml:space="preserve">Org</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t xml:space="preserve">City, ST</w:t></w:r></w:p>',
		);
		pkg.writeText("word/document.xml", xml);
		await pkg.save(docPath);

		// AST exposes the tab stop; read flags the wrap risk.
		const ast = (await runCli("read", docPath, "--ast")).parsed as {
			blocks: Array<{
				id: string;
				tabStops?: { align: string; pos: number }[];
			}>;
		};
		expect(ast.blocks[0]?.tabStops).toEqual([{ align: "left", pos: 10000 }]);
		const md = await readMarkdown(docPath);
		expect(md).toContain("docx:layout");
		expect(md).toContain("LEFT tab");
		// Per-line note points at the cure; the consolidated top summary carries the
		// one-call fix-all command across every wrapping line.
		expect(md).toContain("--tabs right");
		expect(md).toContain('fix-all="edit FILE --at p0 --tabs right"');
	});

	test("consolidates MANY wrapping lines into one summary spanning their range", async () => {
		const workspace = tempWorkspace("layout-summary-multi");
		const docPath = join(workspace, "doc.docx");
		await runCli("create", docPath, "--text", "seed");
		// Two tabbed lines (p1, p3) with a near-margin LEFT tab + a plain line between.
		for (const after of ["p0", "p1", "p2"]) {
			await runCli("insert", docPath, "--after", after, "--text", "filler");
		}
		await runCli("edit", docPath, "--at", "p1", "--text", "Org\tCity, State");
		await runCli("edit", docPath, "--at", "p3", "--text", "Org\tCity, State");
		await runCli("edit", docPath, "--at", "p1", "--tabs", "left@5in");
		await runCli("edit", docPath, "--at", "p3", "--tabs", "left@5in");

		const md = await readMarkdown(docPath);
		// ONE consolidated summary, covering 2 lines, with a single range cure that
		// spans min..max (the range form skips the plain p2 in between).
		expect(md).toContain('wrap="2 lines"');
		expect(md).toContain('fix-all="edit FILE --at p1-p3 --tabs right"');
		expect(md.match(/fix-all=/g)).toHaveLength(1);
	});
});

describe("docx:page note (deviation-only page geometry)", () => {
	/** Rewrite a created doc's trailing sectPr geometry directly (faster than the
	 *  authoring verb, and decoupled from it) so we can exercise the read note. */
	async function withGeometry(
		path: string,
		pgSz: string,
		pgMar: string,
	): Promise<void> {
		const zip = await JSZip.loadAsync(await Bun.file(path).bytes());
		const entry = zip.file("word/document.xml");
		if (!entry) throw new Error("document.xml missing");
		const xml = (await entry.async("string"))
			.replace(/<w:pgSz[^/]*\/>/, pgSz)
			.replace(/<w:pgMar[^/]*\/>/, pgMar);
		zip.file("word/document.xml", xml);
		await Bun.write(path, await zip.generateAsync({ type: "uint8array" }));
	}

	async function defaultDoc(label: string): Promise<string> {
		const workspace = tempWorkspace(label);
		const md = join(workspace, "x.md");
		await Bun.write(md, "# T\n\nhello\n");
		const doc = join(workspace, "x.docx");
		await runCli("create", doc, "--from", md);
		return doc;
	}

	test("a plain default-Letter doc emits no page note", async () => {
		const md = await readMarkdown(await defaultDoc("page-default"));
		expect(md).not.toContain("docx:page");
	});

	test("landscape + narrow margins surface orientation, margins, text-width", async () => {
		const doc = await defaultDoc("page-landscape");
		await withGeometry(
			doc,
			'<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>',
			'<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>',
		);
		const md = await readMarkdown(doc);
		expect(md).toContain('orientation="landscape"');
		// Landscape Letter is still Letter — size is suppressed.
		expect(md).not.toContain("size=");
		expect(md).toContain('margins="0.5in"');
		expect(md).toContain('text-width="10in"'); // 15840 − 720 − 720 = 14400tw = 10in
		// The note leads with the trailing section's locator so an agent can
		// re-apply via `sections --at sN …` (matches docx:cell / docx:p).
		expect(md).toMatch(/<!-- docx:page s\d+ /);
	});

	// <w:pgSz> is schema-optional — a section can set non-default margins yet
	// inherit page size. Keying geometry detection on pageWidth alone treated this
	// as "no geometry": the margin deviation was hidden AND content width fell back
	// to the 6.5in default, mis-flagging image overflow.
	test("non-default margins with no <w:pgSz> still surface a page note + correct text-width", async () => {
		const doc = await defaultDoc("page-marginsonly");
		const zip = await JSZip.loadAsync(await Bun.file(doc).bytes());
		const entry = zip.file("word/document.xml");
		if (!entry) throw new Error("document.xml missing");
		const xml = (await entry.async("string"))
			.replace(/<w:pgSz[^/]*\/>/, "") // drop page size entirely
			.replace(
				/<w:pgMar[^/]*\/>/,
				'<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>',
			);
		zip.file("word/document.xml", xml);
		await Bun.write(doc, await zip.generateAsync({ type: "uint8array" }));

		const md = await readMarkdown(doc);
		expect(md).toContain('margins="0.5in"');
		// Default Letter width (12240) − 720 − 720 = 10800tw = 7.5in (the real
		// margins applied against the inherited default page width).
		expect(md).toContain('text-width="7.5in"');
	});

	test("a non-Letter size (A4) surfaces the size attribute, exact twips in --ast", async () => {
		const doc = await defaultDoc("page-a4");
		await withGeometry(
			doc,
			'<w:pgSz w:w="11906" w:h="16838"/>',
			'<w:pgMar w:top="1440" w:right="1134" w:bottom="1440" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/>',
		);
		// twipsToInches uses 3 decimals (margins need 1/8" precision), so A4's
		// non-round inches show in full.
		expect(await readMarkdown(doc)).toContain('size="8.268x11.693in"');
		const ast = (await runCli("read", doc, "--ast")).parsed as {
			blocks: SectionBlock[];
		};
		const trailing = ast.blocks
			.filter((block) => block.type === "sectionBreak")
			.pop();
		expect(trailing?.pageWidth).toBe(11906);
		expect(trailing?.marginLeft).toBe(1134);
	});

	// Regression: the note must describe page 1 and flag when later sections
	// differ — it used to report the TRAILING section, so a doc that was landscape
	// only on its last section read as "landscape" while page 1 was portrait.
	test("mixed-geometry doc flags varies=by-section (note describes page 1)", async () => {
		const doc = join(tempWorkspace("page-varies"), "x.docx");
		await runCli("create", doc, "--text", "Title.");
		await runCli("insert", doc, "--at-end", "--text", "Body one.");
		await runCli("insert", doc, "--at-end", "--text", "Body two.");
		// Wrap a middle range to mint sections, then set ONLY the body section
		// landscape — leaving page 1 (s0) portrait-default.
		await runCli("sections", doc, "--at", "p1-p2", "--columns", "2");
		await runCli("sections", doc, "--at", "s1", "--orientation", "landscape");
		const md = await readMarkdown(doc);
		expect(md).toContain("docx:page s0"); // note describes the FIRST section
		expect(md).toContain('varies="by-section"');
		// Page 1 is default portrait, so no orientation deviation is claimed.
		expect(md).not.toContain('orientation="landscape"');
	});

	test("a uniform default-Letter doc emits no varies flag (and no note)", async () => {
		const doc = join(tempWorkspace("page-uniform"), "x.docx");
		await runCli("create", doc, "--text", "Body.");
		const md = await readMarkdown(doc);
		expect(md).not.toContain("docx:page");
		expect(md).not.toContain("varies=");
	});
});

describe("page setup (docx sections --at sN: margins / orientation / size)", () => {
	async function plainDoc(label: string): Promise<string> {
		const doc = join(tempWorkspace(label), "x.docx");
		await runCli("create", doc, "--text", "Body.");
		return doc;
	}
	async function trailingSection(doc: string): Promise<SectionBlock> {
		const ast = (await runCli("read", doc, "--ast")).parsed as {
			blocks: Block[];
		};
		const trailing = ast.blocks
			.filter((block): block is SectionBlock => block.type === "sectionBreak")
			.pop();
		if (!trailing) throw new Error("no section break");
		return trailing;
	}

	test("orientation landscape swaps page dimensions + sets w:orient", async () => {
		const doc = await plainDoc("pg-orient");
		expect(
			(
				await runCli(
					"sections",
					doc,
					"--at",
					"s0",
					"--orientation",
					"landscape",
				)
			).exitCode,
		).toBe(0);
		const s = await trailingSection(doc);
		expect(s.pageWidth).toBe(15840);
		expect(s.pageHeight).toBe(12240);
		expect(s.pageOrientation).toBe("landscape");
	});

	test("named --size legal sets exact dimensions", async () => {
		const doc = await plainDoc("pg-legal");
		await runCli("sections", doc, "--at", "s0", "--size", "legal");
		const s = await trailingSection(doc);
		expect(s.pageWidth).toBe(12240);
		expect(s.pageHeight).toBe(20160);
	});

	test("WxH --size with W>H implies landscape", async () => {
		const doc = await plainDoc("pg-wxh");
		await runCli("sections", doc, "--at", "s0", "--size", "11x8.5in");
		const s = await trailingSection(doc);
		expect(s.pageWidth).toBe(15840);
		expect(s.pageHeight).toBe(12240);
		expect(s.pageOrientation).toBe("landscape");
	});

	test("--margins uniform and 4-tuple (CSS order top,right,bottom,left)", async () => {
		const uniform = await plainDoc("pg-mar-uniform");
		await runCli("sections", uniform, "--at", "s0", "--margins", "1in");
		const u = await trailingSection(uniform);
		expect([u.marginTop, u.marginRight, u.marginBottom, u.marginLeft]).toEqual([
			1440, 1440, 1440, 1440,
		]);

		const four = await plainDoc("pg-mar-four");
		await runCli(
			"sections",
			four,
			"--at",
			"s0",
			"--margins",
			"0.75,1,0.75,1.25",
		);
		const f = await trailingSection(four);
		expect([f.marginTop, f.marginRight, f.marginBottom, f.marginLeft]).toEqual([
			1080, 1440, 1080, 1800,
		]);
	});

	test("emits CT_SectPr in valid child order (pgSz → pgMar → cols → sectPrChange)", async () => {
		const doc = await plainDoc("pg-order");
		await runCli(
			"sections",
			doc,
			"--at",
			"s0",
			"--size",
			"a4",
			"--margins",
			"1in",
			"--columns",
			"2",
		);
		const xml = await readDocumentXml(doc);
		const sectPr = xml.match(/<w:sectPr>[\s\S]*?<\/w:sectPr>/)?.[0] ?? "";
		const order = ["w:pgSz", "w:pgMar", "w:cols"].map((tag) =>
			sectPr.indexOf(`<${tag}`),
		);
		expect(order.every((index) => index !== -1)).toBe(true);
		expect(order).toEqual([...order].sort((a, b) => a - b));
	});

	test("the read note round-trips: re-applying its margins/size reproduces them", async () => {
		const doc = await plainDoc("pg-roundtrip");
		await runCli(
			"sections",
			doc,
			"--at",
			"s0",
			"--size",
			"legal",
			"--margins",
			"1.5in",
		);
		const before = await trailingSection(doc);
		// Read the note, parse its size/margins, feed them back into a fresh doc.
		const md = await readMarkdown(doc);
		const size = md.match(/size="([^"]+)"/)?.[1];
		const margins = md.match(/margins="([^"]+)"/)?.[1];
		expect(size).toBeDefined();
		expect(margins).toBeDefined();
		const doc2 = await plainDoc("pg-roundtrip2");
		await runCli(
			"sections",
			doc2,
			"--at",
			"s0",
			"--size",
			size as string,
			"--margins",
			margins as string,
		);
		const after = await trailingSection(doc2);
		expect(after.pageWidth).toBe(before.pageWidth);
		expect(after.pageHeight).toBe(before.pageHeight);
		expect(after.marginTop).toBe(before.marginTop);
		expect(after.marginLeft).toBe(before.marginLeft);
	});

	test("page geometry is rejected on a column-wrap range (--at pN-pM)", async () => {
		const doc = join(tempWorkspace("pg-wrap"), "x.docx");
		await runCli("create", doc, "--text", "One.\nTwo.\nThree.");
		const result = await runCli(
			"sections",
			doc,
			"--at",
			"p0-p2",
			"--columns",
			"2",
			"--orientation",
			"landscape",
		);
		expect(result.exitCode).not.toBe(0);
		expect((result.parsed as { error?: string }).error).toContain(
			"EXISTING section",
		);
	});

	test("an sN edit with no flag at all is rejected", async () => {
		const doc = await plainDoc("pg-noflag");
		const result = await runCli("sections", doc, "--at", "s0");
		expect(result.exitCode).not.toBe(0);
		expect((result.parsed as { error?: string }).error).toContain(
			"--columns, --type, --orientation, --size, or --margins",
		);
	});

	test("invalid --size and --orientation are rejected", async () => {
		const doc = await plainDoc("pg-bad");
		expect(
			(await runCli("sections", doc, "--at", "s0", "--size", "huge")).exitCode,
		).not.toBe(0);
		expect(
			(await runCli("sections", doc, "--at", "s0", "--orientation", "sideways"))
				.exitCode,
		).not.toBe(0);
	});

	describe("under track-changes (one <w:sectPrChange> for the whole edit)", () => {
		async function trackedDoc(label: string): Promise<string> {
			const doc = await plainDoc(label);
			await runCli("track-changes", "on", doc);
			return doc;
		}

		test("records a sectPrChange with prior/current page geometry", async () => {
			const doc = await trackedDoc("pg-tc-list");
			await runCli(
				"sections",
				doc,
				"--at",
				"s0",
				"--orientation",
				"landscape",
				"--margins",
				"2in",
			);
			expect(await trackedKinds(doc)).toContain("sectPrChange");
			const changes = (await runCli("track-changes", "list", doc))
				.parsed as Array<{
				kind: string;
				prior?: { pageOrientation?: string; marginTop?: number };
				current?: { pageOrientation?: string; marginTop?: number };
			}>;
			const change = changes.find((c) => c.kind === "sectPrChange");
			expect(change?.prior?.pageOrientation).toBeUndefined(); // prior = portrait
			expect(change?.prior?.marginTop).toBe(1440);
			expect(change?.current?.pageOrientation).toBe("landscape");
			expect(change?.current?.marginTop).toBe(2880);
		});

		test("accept keeps the new geometry; reject restores the prior", async () => {
			const accept = await trackedDoc("pg-tc-accept");
			await runCli(
				"sections",
				accept,
				"--at",
				"s0",
				"--orientation",
				"landscape",
			);
			await runCli("track-changes", "accept", accept, "--all");
			const a = await trailingSection(accept);
			expect(a.pageOrientation).toBe("landscape");
			expect(await trackedKinds(accept)).not.toContain("sectPrChange");

			const reject = await trackedDoc("pg-tc-reject");
			await runCli(
				"sections",
				reject,
				"--at",
				"s0",
				"--orientation",
				"landscape",
				"--margins",
				"2in",
			);
			await runCli("track-changes", "reject", reject, "--all");
			const r = await trailingSection(reject);
			expect(r.pageWidth).toBe(12240); // back to portrait Letter
			expect(r.pageHeight).toBe(15840);
			expect(r.marginTop).toBe(1440); // back to 1in
		});
	});

	// Regression: a column wrap mints fresh sentinel sectPrs. Page geometry is
	// per-section, so without inheritance the new sections silently revert to
	// portrait-Letter — the "landscape vanishes after adding columns" blocker.
	test("column-wrap preserves the document's page geometry on every section", async () => {
		const doc = join(tempWorkspace("pg-wrap-inherit"), "x.docx");
		await runCli(
			"create",
			doc,
			"--text",
			"Title.",
			"--orientation",
			"landscape",
			"--margins",
			"0.5in",
		);
		// Wrap a middle range in 2 columns (needs ≥3 paragraphs to wrap p1).
		await runCli("insert", doc, "--at-end", "--text", "Body one.");
		await runCli("insert", doc, "--at-end", "--text", "Body two.");
		expect(
			(await runCli("sections", doc, "--at", "p1-p2", "--columns", "2"))
				.exitCode,
		).toBe(0);
		const ast = (await runCli("read", doc, "--ast")).parsed as {
			blocks: Block[];
		};
		const sections = ast.blocks.filter(
			(block): block is SectionBlock => block.type === "sectionBreak",
		);
		expect(sections.length).toBeGreaterThanOrEqual(2); // wrap added breaks
		// EVERY section keeps the landscape geometry + 0.5in margins.
		for (const s of sections) {
			expect(s.pageWidth).toBe(15840);
			expect(s.pageHeight).toBe(12240);
			expect(s.pageOrientation).toBe("landscape");
			expect(s.marginTop).toBe(720);
		}
		// The read note reports landscape with NO varies flag (uniform geometry).
		const md = await readMarkdown(doc);
		expect(md).toContain('orientation="landscape"');
		expect(md).not.toContain("varies=");
	});
});

describe("page setup, whole document (docx sections --margins/… with no --at)", () => {
	// A doc with THREE sections (two inline sectPrs from a column wrap + the
	// trailing body sectPr) — the multi-section shape where a per-section margin
	// sweep silently misses the trailing governing section (the resume s3 trap).
	async function multiSectionDoc(label: string): Promise<string> {
		const doc = join(tempWorkspace(label), "x.docx");
		await runCli("create", doc, "--text", "Alpha.");
		await runCli("insert", doc, "--at-end", "--text", "Beta.");
		await runCli("insert", doc, "--at-end", "--text", "Gamma.");
		await runCli("sections", doc, "--at", "p1", "--columns", "2");
		return doc;
	}
	async function allSections(doc: string): Promise<SectionBlock[]> {
		const ast = (await runCli("read", doc, "--ast")).parsed as {
			blocks: Block[];
		};
		return ast.blocks.filter(
			(block): block is SectionBlock => block.type === "sectionBreak",
		);
	}

	test("--margins with no --at sets EVERY section (the resume s3 trap)", async () => {
		const doc = await multiSectionDoc("pg-all-margins");
		expect((await allSections(doc)).length).toBeGreaterThanOrEqual(3);
		const result = await runCli("sections", doc, "--margins", "0.5");
		expect(result.exitCode).toBe(0);
		for (const s of await allSections(doc)) {
			expect([
				s.marginTop,
				s.marginRight,
				s.marginBottom,
				s.marginLeft,
			]).toEqual([720, 720, 720, 720]);
		}
	});

	test("--orientation with no --at sets every section landscape", async () => {
		const doc = await multiSectionDoc("pg-all-orient");
		expect(
			(await runCli("sections", doc, "--orientation", "landscape")).exitCode,
		).toBe(0);
		for (const s of await allSections(doc)) {
			expect(s.pageOrientation).toBe("landscape");
		}
	});

	test("no --at and no page geometry is still a USAGE error", async () => {
		const doc = await multiSectionDoc("pg-all-noflag");
		const result = await runCli("sections", doc);
		expect(result.exitCode).not.toBe(0);
		expect((result.parsed as { error?: string }).error).toContain("--at");
	});

	test("--columns with no --at is rejected (columns need a target section)", async () => {
		const doc = await multiSectionDoc("pg-all-cols");
		const result = await runCli("sections", doc, "--columns", "2");
		expect(result.exitCode).not.toBe(0);
		expect((result.parsed as { error?: string }).error).toContain("--at");
	});

	test("tracked doc-wide margins records a sectPrChange per section", async () => {
		const doc = await multiSectionDoc("pg-all-tracked");
		await runCli("track-changes", "on", doc);
		const sectionCount = (await allSections(doc)).length;
		await runCli("sections", doc, "--margins", "0.5");
		const sectPrChanges = (await trackedKinds(doc)).filter(
			(kind) => kind === "sectPrChange",
		).length;
		expect(sectPrChanges).toBe(sectionCount);
	});
});

describe("page setup auto-realigns wrapping tab columns", () => {
	// A résumé-style line: a bold org with a right-edge LEFT tab holding a long
	// location. The LEFT tab at 6.1in (~8784tw) is right-edge for the default 1in
	// margins (text width 9360tw) — so it's the fragile wrap hazard `read` flags.
	async function fragileTabDoc(label: string): Promise<string> {
		const doc = join(tempWorkspace(label), "x.docx");
		await runCli("create", doc, "--text", "Priya Raman");
		await runCli(
			"insert",
			doc,
			"--after",
			"p0",
			"--runs",
			'[{"type":"text","text":"Northwind Robotics","bold":true},{"type":"tab"},{"type":"text","text":"San Francisco, CA"}]',
		);
		await runCli("edit", doc, "--at", "p1", "--tabs", "left@6.1in");
		return doc;
	}
	async function tabStops(
		doc: string,
		id: string,
	): Promise<Array<{ align: string; pos: number }>> {
		const ast = (await runCli("read", doc, "--ast")).parsed as {
			blocks: Array<{
				id: string;
				tabStops?: Array<{ align: string; pos: number }>;
			}>;
		};
		return ast.blocks.find((block) => block.id === id)?.tabStops ?? [];
	}

	test("doc-wide margins cure the wrap — fragile LEFT tab → RIGHT tab at the new margin", async () => {
		const doc = await fragileTabDoc("pg-reflow");
		expect(await readMarkdown(doc)).toContain("docx:layout"); // flagged before
		const result = await runCli("sections", doc, "--margins", "0.5");
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { realignedTabs?: number }).realignedTabs).toBe(1);
		// 0.5in margins on letter → text width 12240 − 720 − 720 = 10800tw.
		expect(await tabStops(doc, "p1")).toEqual([{ align: "right", pos: 10800 }]);
		expect(await readMarkdown(doc)).not.toContain("docx:layout"); // cured
	});

	test("a non-fragile mid-line LEFT tab is left untouched", async () => {
		const doc = join(tempWorkspace("pg-reflow-midline"), "x.docx");
		await runCli("create", doc, "--text", "Name");
		await runCli(
			"insert",
			doc,
			"--after",
			"p0",
			"--runs",
			'[{"type":"text","text":"A"},{"type":"tab"},{"type":"text","text":"B"}]',
		);
		await runCli("edit", doc, "--at", "p1", "--tabs", "left@3in"); // mid-line
		const result = await runCli("sections", doc, "--margins", "0.5");
		expect(
			(result.parsed as { realignedTabs?: number }).realignedTabs,
		).toBeUndefined();
		expect(await tabStops(doc, "p1")).toEqual([{ align: "left", pos: 4320 }]);
	});

	test("an existing RIGHT tab is already robust — not touched", async () => {
		const doc = join(tempWorkspace("pg-reflow-right"), "x.docx");
		await runCli("create", doc, "--text", "Name");
		await runCli(
			"insert",
			doc,
			"--after",
			"p0",
			"--runs",
			'[{"type":"text","text":"Org"},{"type":"tab"},{"type":"text","text":"Date"}]',
		);
		await runCli("edit", doc, "--at", "p1", "--tabs", "right@6.1in");
		const result = await runCli("sections", doc, "--margins", "0.5");
		expect(
			(result.parsed as { realignedTabs?: number }).realignedTabs,
		).toBeUndefined();
	});

	test("single-section --at s0 margins also realigns fragile tabs", async () => {
		const doc = await fragileTabDoc("pg-reflow-at-s0");
		const result = await runCli(
			"sections",
			doc,
			"--at",
			"s0",
			"--margins",
			"0.5",
		);
		expect((result.parsed as { realignedTabs?: number }).realignedTabs).toBe(1);
		expect(await tabStops(doc, "p1")).toEqual([{ align: "right", pos: 10800 }]);
	});

	test("tracked margins record the tab reflow as a pPrChange; reject restores the LEFT tab", async () => {
		const doc = await fragileTabDoc("pg-reflow-tracked");
		await runCli("track-changes", "on", doc);
		await runCli("sections", doc, "--margins", "0.5");
		expect(await trackedKinds(doc)).toContain("pPrChange");
		await runCli("track-changes", "reject", doc, "--all");
		expect(
			(await tabStops(doc, "p1")).some((tab) => tab.align === "left"),
		).toBe(true);
	});

	// Regression: the cure must PRESERVE non-fragile stops, not wipe the whole
	// <w:tabs> — a line with a legit mid-line LEFT tab AND a fragile right-edge one
	// keeps the mid-line tab and only the right-edge one becomes a right tab.
	test("preserves a non-fragile mid-line tab while curing the fragile one", async () => {
		const doc = join(tempWorkspace("pg-reflow-preserve"), "x.docx");
		await runCli("create", doc, "--text", "Name");
		await runCli(
			"insert",
			doc,
			"--after",
			"p0",
			"--runs",
			'[{"type":"text","text":"A"},{"type":"tab"},{"type":"text","text":"B"},{"type":"tab"},{"type":"text","text":"C"}]',
		);
		await runCli("edit", doc, "--at", "p1", "--tabs", "left@3in,left@6.1in");
		const result = await runCli("sections", doc, "--margins", "0.5");
		expect((result.parsed as { realignedTabs?: number }).realignedTabs).toBe(1);
		// mid-line left@3in (4320tw) survives; the fragile left@6.1in becomes right@10800.
		expect(await tabStops(doc, "p1")).toEqual([
			{ align: "left", pos: 4320 },
			{ align: "right", pos: 10800 },
		]);
	});

	// Regression: a multi-column section's tab stops are column-relative — the
	// right-margin cure doesn't apply, so reflow must skip it.
	test("skips reflow in a multi-column section", async () => {
		const doc = join(tempWorkspace("pg-reflow-cols"), "x.docx");
		await runCli("create", doc, "--text", "Body alpha.");
		await runCli(
			"insert",
			doc,
			"--after",
			"p0",
			"--runs",
			'[{"type":"text","text":"Org"},{"type":"tab"},{"type":"text","text":"Date"}]',
		);
		await runCli("edit", doc, "--at", "p1", "--tabs", "left@6.1in");
		const result = await runCli(
			"sections",
			doc,
			"--at",
			"s0",
			"--columns",
			"2",
			"--margins",
			"0.5",
		);
		expect(
			(result.parsed as { realignedTabs?: number }).realignedTabs,
		).toBeUndefined();
		// The fragile tab is LEFT untouched (no full-width right tab in a column).
		expect(await tabStops(doc, "p1")).toEqual([{ align: "left", pos: 8784 }]);
	});
});

// Layout is pure-visibility: the importer drops docx:section, so a full
// read → create rebuild doesn't reconstruct sections (it doesn't corrupt them
// into border paragraphs either). A hand-authored `---` is a thematic break.
describe("section layout is pure-visibility on import", () => {
	test("read → create drops mid-doc section breaks (only trailing remains)", async () => {
		const workspace = tempWorkspace("section-vis");
		const md = await readMarkdown(SECTIONS_FIXTURE);
		expect(md).toContain("docx:section");
		const mdPath = join(workspace, "s.md");
		await Bun.write(mdPath, md);
		const rebuilt = join(workspace, "rebuilt.docx");
		expect((await runCli("create", rebuilt, "--from", mdPath)).exitCode).toBe(
			0,
		);
		const ast = (await runCli("read", rebuilt, "--ast")).parsed as {
			blocks: SectionBlock[];
		};
		const breaks = ast.blocks.filter((block) => block.type === "sectionBreak");
		expect(breaks).toHaveLength(1); // just the trailing mandatory sectPr
		expect(await readMarkdown(rebuilt)).not.toContain("---");
	});

	test("a hand-authored --- is a thematic break, not a section break", async () => {
		const workspace = tempWorkspace("thematic-break");
		const mdPath = join(workspace, "hr.md");
		await Bun.write(mdPath, "# Title\n\nbefore\n\n---\n\nafter\n");
		const doc = join(workspace, "hr.docx");
		expect((await runCli("create", doc, "--from", mdPath)).exitCode).toBe(0);
		const ast = (await runCli("read", doc, "--ast")).parsed as {
			blocks: SectionBlock[];
		};
		// Only the trailing mandatory sectPr — the `---` became a border paragraph.
		expect(ast.blocks.filter((b) => b.type === "sectionBreak")).toHaveLength(1);
		expect(await readMarkdown(doc)).not.toContain("docx:section");
	});
});
