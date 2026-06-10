import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

// `docx sections` is the intuitive verb for column layout: wrap a paragraph range
// in its own N-column section (--at pN-pM / pN), or recount an existing section
// (--at sN). The wrapping section breaks are written as real <w:sectPr>; they
// surface in `read` as docx:section hints but, per "comments are hints", do NOT
// round-trip through create (covered in sections.test.ts / markdown.test.ts).

type SectionBlock = {
	id: string;
	type: string;
	columns?: number;
	sectionType?: string;
};

async function makeDoc(label: string, paras: number): Promise<string> {
	const workspace = tempWorkspace(label);
	const body = Array.from({ length: paras }, (_, i) => `para${i}`).join("\n\n");
	const src = join(workspace, "src.md");
	await Bun.write(src, `# Doc\n\n${body}\n`);
	const doc = join(workspace, "doc.docx");
	await runCli("create", doc, "--from", src);
	return doc;
}

async function sections(path: string): Promise<SectionBlock[]> {
	const ast = (await runCli("read", path, "--ast")).parsed as {
		blocks: SectionBlock[];
	};
	return ast.blocks.filter((block) => block.type === "sectionBreak");
}

describe("docx sections — range wrap (--at pN-pM)", () => {
	test("wraps a paragraph range in its own N-column continuous section", async () => {
		const doc = await makeDoc("columns-range", 6); // p0=# Doc, p1..p6=para0..5
		expect(
			(await runCli("sections", doc, "--at", "p2-p4", "--columns", "2"))
				.exitCode,
		).toBe(0);

		const secs = await sections(doc);
		// One cols=2 break (ends the wrapped section) + a no-cols break bounding its
		// start + the trailing mandatory sectPr.
		expect(secs.filter((s) => s.columns === 2)).toHaveLength(1);

		// The 2-column break sits after the range; a bounding break sits before it.
		const md = (await runCli("read", doc)).stdout;
		const lines = md.split("\n").filter((line) => line.length > 0);
		const beforeIdx = lines.findIndex(
			(line) => line.includes("docx:section") && !line.includes('cols="2"'),
		);
		const afterIdx = lines.findIndex((line) => line.includes('cols="2"'));
		expect(beforeIdx).toBeGreaterThanOrEqual(0);
		expect(afterIdx).toBeGreaterThan(beforeIdx);
	});

	test("a single paragraph (--at pN) is a one-paragraph range", async () => {
		const doc = await makeDoc("columns-single", 4);
		expect(
			(await runCli("sections", doc, "--at", "p2", "--columns", "2")).exitCode,
		).toBe(0);
		expect((await sections(doc)).some((s) => s.columns === 2)).toBe(true);
	});

	test("a range at the document start omits the leading break", async () => {
		const doc = await makeDoc("columns-start", 4);
		await runCli("sections", doc, "--at", "p0-p2", "--columns", "2");
		const secs = await sections(doc);
		// Just the cols=2 break + the trailing sectPr — no spurious empty leading
		// section before p0.
		expect(secs).toHaveLength(2);
		expect(secs.filter((s) => s.columns === 2)).toHaveLength(1);
	});

	test("--type sets the wrapping break's section type", async () => {
		const doc = await makeDoc("columns-type", 5);
		await runCli(
			"sections",
			doc,
			"--at",
			"p2-p3",
			"--columns",
			"3",
			"--type",
			"continuous",
		);
		const wrapped = (await sections(doc)).find((s) => s.columns === 3);
		expect(wrapped?.sectionType).toBe("continuous");
	});

	test("rejects a range that already contains a section break", async () => {
		const doc = await makeDoc("columns-guard", 5);
		await runCli("sections", doc, "--at", "p2-p3", "--columns", "2");
		// A break now exists at the old p3 boundary; a spanning range must fail
		// rather than nest sections ambiguously.
		const result = await runCli(
			"sections",
			doc,
			"--at",
			"p1-p6",
			"--columns",
			"2",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});

	test("--dry-run previews without writing", async () => {
		const doc = await makeDoc("columns-dry", 4);
		const before = await Bun.file(doc).bytes();
		const result = await runCli(
			"sections",
			doc,
			"--at",
			"p1-p2",
			"--columns",
			"2",
			"--dry-run",
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { dryRun: boolean }).dryRun).toBe(true);
		expect(await Bun.file(doc).bytes()).toEqual(before);
	});
});

describe("docx sections — section edit (--at sN)", () => {
	test("recounts an existing section's columns in place", async () => {
		const doc = await makeDoc("columns-section", 5);
		await runCli("sections", doc, "--at", "p2-p3", "--columns", "2");
		const wrapped = (await sections(doc)).find((s) => s.columns === 2);
		expect(wrapped).toBeDefined();
		const id = wrapped?.id as string;

		expect(
			(await runCli("sections", doc, "--at", id, "--columns", "3")).exitCode,
		).toBe(0);
		const after = await sections(doc);
		expect(after.some((s) => s.columns === 3)).toBe(true);
		expect(after.some((s) => s.columns === 2)).toBe(false);
	});

	test("--columns 1 collapses a section back to a single column", async () => {
		const doc = await makeDoc("columns-collapse", 5);
		await runCli("sections", doc, "--at", "p2-p3", "--columns", "2");
		const id = (await sections(doc)).find((s) => s.columns === 2)?.id as string;
		await runCli("sections", doc, "--at", id, "--columns", "1");
		expect((await sections(doc)).some((s) => s.columns === 2)).toBe(false);
	});
});

describe("docx sections — validation", () => {
	test("rejects a missing --columns", async () => {
		const doc = await makeDoc("columns-nocount", 3);
		expect((await runCli("sections", doc, "--at", "p1-p2")).exitCode).toBe(2);
	});

	test("rejects a non-positive --columns", async () => {
		const doc = await makeDoc("columns-badcount", 3);
		expect(
			(await runCli("sections", doc, "--at", "p1-p2", "--columns", "0"))
				.exitCode,
		).toBe(2);
	});

	// `Number.parseInt` would silently TRUNCATE "2.5"→2 and "1e2"→1, both passing
	// the `>= 1` check — an agent that meant 100 columns would get 1, no error.
	test("rejects a non-integer --columns instead of truncating", async () => {
		const doc = await makeDoc("columns-floatcount", 3);
		for (const bad of ["2.5", "1e2", "0x4", "abc"]) {
			const result = await runCli(
				"sections",
				doc,
				"--at",
				"p1",
				"--columns",
				bad,
			);
			expect(result.exitCode).toBe(2);
			expect((result.parsed as { code: string }).code).toBe("USAGE");
		}
		// Nothing was written — no stray section break from a coerced count.
		expect((await sections(doc)).filter((s) => s.columns)).toHaveLength(0);
	});

	test("rejects a missing --at", async () => {
		const doc = await makeDoc("columns-noat", 3);
		expect((await runCli("sections", doc, "--columns", "2")).exitCode).toBe(2);
	});
});

// A <w:sectPr> may only live in <w:body> or a body-level paragraph's <w:pPr>
// (ECMA-376) — never inside <w:tc>. Wrapping a table-cell paragraph would write
// invalid OOXML ("unreadable content") AND be invisible on read-back (the reader
// doesn't enumerate in-cell sectPr), silently breaking the write-read loop.
describe("docx sections — body-level guard", () => {
	async function tableDoc(label: string): Promise<string> {
		const workspace = tempWorkspace(label);
		const src = join(workspace, "t.md");
		await Bun.write(src, "# Doc\n\n| A | B |\n| --- | --- |\n| c0 | c1 |\n");
		const doc = join(workspace, "doc.docx");
		await runCli("create", doc, "--from", src);
		return doc;
	}

	test("rejects a paragraph inside a table cell, leaving the doc untouched", async () => {
		const doc = await tableDoc("columns-cell");
		const before = await Bun.file(doc).bytes();
		const result = await runCli(
			"sections",
			doc,
			"--at",
			"t0:r0c0:p0",
			"--columns",
			"2",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
		expect((result.parsed as { error: string }).error).toContain("table cell");
		// No in-cell sectPr written: the file is byte-identical.
		expect(await Bun.file(doc).bytes()).toEqual(before);
		// And it still reads (no corruption) with no columns anywhere.
		const read = await runCli("read", doc, "--ast");
		expect(read.exitCode).toBe(0);
		expect((await sections(doc)).filter((s) => s.columns)).toHaveLength(0);
	});
});
