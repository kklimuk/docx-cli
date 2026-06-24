import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";
import { freshFixture, readMarkdown } from "./helpers";

const FIXTURES = join(import.meta.dir, "..", "fixtures");

async function newDoc(label: string, ...paragraphs: string[]): Promise<string> {
	const path = join(tempWorkspace(label), "doc.docx");
	await runCli("create", path, "--text", paragraphs[0] ?? "Body.", "--force");
	for (let index = 1; index < paragraphs.length; index++) {
		await runCli(
			"insert",
			path,
			"--after",
			`p${index - 1}`,
			"--text",
			paragraphs[index] as string,
		);
	}
	return path;
}

async function partText(path: string, part: string): Promise<string> {
	const pkg = await Pkg.open(path);
	return await pkg.readText(part);
}

type Marginal = {
	id: string;
	kind: string;
	type: string;
	sectionId: string;
	text: string;
};

async function list(
	noun: "headers" | "footers",
	path: string,
): Promise<Marginal[]> {
	const result = await runCli(noun, "list", path);
	return result.parsed as Marginal[];
}

describe("headers/footers — set basics", () => {
	test("headers set --text creates a header part + AST entry", async () => {
		const path = await newDoc("hf-text");
		const result = await runCli("headers", "set", path, "--text", "My Header");
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			ok: true,
			operation: "headers.set",
			kind: "header",
			type: "default",
		});
		const headers = await list("headers", path);
		expect(headers).toHaveLength(1);
		expect(headers[0]).toMatchObject({
			id: "hdr0",
			kind: "header",
			type: "default",
			text: "My Header",
		});
		const part = await partText(path, "word/header1.xml");
		expect(part).toContain("<w:hdr");
		expect(part).toContain("My Header");
	});

	test("footers set --page-number emits a PAGE field, centered by default", async () => {
		const path = await newDoc("hf-page");
		await runCli("footers", "set", path, "--page-number");
		const footers = await list("footers", path);
		expect(footers[0]?.text).toBe("{page}");
		const part = await partText(path, "word/footer1.xml");
		expect(part).toContain('<w:fldSimple w:instr=" PAGE "');
		expect(part).toContain('<w:jc w:val="center"/>');
	});

	test("footers set --page-number --of-pages emits PAGE + NUMPAGES", async () => {
		const path = await newDoc("hf-ofpages");
		await runCli("footers", "set", path, "--page-number", "--of-pages");
		const footers = await list("footers", path);
		expect(footers[0]?.text).toBe("Page {page} of {pages}");
		const part = await partText(path, "word/footer1.xml");
		expect(part).toContain('w:instr=" PAGE "');
		expect(part).toContain('w:instr=" NUMPAGES "');
	});

	test("--align right positions a lone field", async () => {
		const path = await newDoc("hf-align");
		await runCli("footers", "set", path, "--page-number", "--align", "right");
		const part = await partText(path, "word/footer1.xml");
		expect(part).toContain('<w:jc w:val="right"/>');
	});

	test("--text replaces existing content (reuses the same part)", async () => {
		const path = await newDoc("hf-replace");
		await runCli("headers", "set", path, "--text", "First");
		await runCli("headers", "set", path, "--text", "Second");
		const headers = await list("headers", path);
		expect(headers).toHaveLength(1);
		expect(headers[0]?.text).toBe("Second");
		// No header2.xml — the same part is reused.
		const pkg = await Pkg.open(path);
		expect(pkg.listParts().filter((p) => /header\d+\.xml$/.test(p))).toEqual([
			"word/header1.xml",
		]);
	});
});

describe("headers/footers — fields", () => {
	test("--date emits a DATE field, shown as {date}", async () => {
		const path = await newDoc("hf-date");
		await runCli("headers", "set", path, "--date");
		expect((await list("headers", path))[0]?.text).toBe("{date}");
		expect(await partText(path, "word/header1.xml")).toContain(
			'w:instr=" DATE "',
		);
	});

	test("--date --date-format emits the format switch", async () => {
		const path = await newDoc("hf-datefmt");
		await runCli(
			"headers",
			"set",
			path,
			"--date",
			"--date-format",
			"MMMM yyyy",
		);
		// Quotes in the instr attribute serialize as &quot; (Word reads them back).
		const part = await partText(path, "word/header1.xml");
		expect(part).toContain("DATE \\@");
		expect(part).toContain("MMMM yyyy");
	});

	test("--style-ref emits STYLEREF, shown as {styleref:NAME}", async () => {
		const path = await newDoc("hf-styleref");
		await runCli("headers", "set", path, "--style-ref", "Heading 1");
		expect((await list("headers", path))[0]?.text).toBe("{styleref:Heading 1}");
		const part = await partText(path, "word/header1.xml");
		expect(part).toContain("STYLEREF");
		expect(part).toContain("Heading 1");
	});

	test("--field filename emits FILENAME, shown as {filename}", async () => {
		const path = await newDoc("hf-filename");
		await runCli("footers", "set", path, "--field", "filename");
		expect((await list("footers", path))[0]?.text).toBe("{filename}");
		expect(await partText(path, "word/footer1.xml")).toContain("FILENAME \\p");
	});
});

describe("headers/footers — two-zone layout", () => {
	test("--text + --page-number puts text left, page right via a content-edge tab", async () => {
		const path = await newDoc("hf-twozone");
		await runCli(
			"headers",
			"set",
			path,
			"--text",
			"Confidential",
			"--page-number",
		);
		const part = await partText(path, "word/header1.xml");
		expect(part).toContain('<w:tab w:val="right" w:pos="9360"/>'); // 6.5in content
		expect(part).toContain("Confidential");
		expect(part).toContain('w:instr=" PAGE "');
		expect((await list("headers", path))[0]?.text).toBe("Confidential\t{page}");
	});
});

describe("headers/footers — placement types", () => {
	test("--first-page sets <w:titlePg/> + a first-type reference", async () => {
		const path = await newDoc("hf-first");
		await runCli("headers", "set", path, "--first-page", "--text", "Cover");
		const xml = await partText(path, "word/document.xml");
		expect(xml).toContain("<w:titlePg/>");
		expect(xml).toMatch(/<w:headerReference w:type="first"/);
		expect((await list("headers", path))[0]?.type).toBe("first");
	});

	test("--even toggles document-level <w:evenAndOddHeaders/>", async () => {
		const path = await newDoc("hf-even");
		await runCli("headers", "set", path, "--even", "--text", "Even");
		expect(await partText(path, "word/settings.xml")).toContain(
			"<w:evenAndOddHeaders/>",
		);
		expect((await list("headers", path))[0]?.type).toBe("even");
	});

	test("--odd is an alias for --type default", async () => {
		const path = await newDoc("hf-odd");
		await runCli("headers", "set", path, "--odd", "--text", "Odd");
		expect((await list("headers", path))[0]?.type).toBe("default");
	});

	test("references are spliced FIRST in CT_SectPr order (before pgSz)", async () => {
		const path = await newDoc("hf-order");
		await runCli("headers", "set", path, "--text", "H");
		await runCli("footers", "set", path, "--page-number");
		const xml = await partText(path, "word/document.xml");
		const headerAt = xml.indexOf("<w:headerReference");
		const footerAt = xml.indexOf("<w:footerReference");
		const pgSzAt = xml.indexOf("<w:pgSz");
		expect(headerAt).toBeGreaterThan(-1);
		expect(headerAt).toBeLessThan(pgSzAt);
		expect(footerAt).toBeLessThan(pgSzAt);
		expect(headerAt).toBeLessThan(footerAt);
	});
});

describe("headers/footers — document-wide vs one section", () => {
	test("no --at applies to every section, sharing one part", async () => {
		const path = await freshFixture("hf-all", join(FIXTURES, "sections.docx"));
		const sectionCount = (await runCli("read", path, "--ast")).parsed as {
			blocks: Array<{ type: string }>;
		};
		const sections = sectionCount.blocks.filter(
			(block) => block.type === "sectionBreak",
		).length;
		const result = await runCli("footers", "set", path, "--page-number");
		expect(result.parsed).toMatchObject({ sections });
		const footers = await list("footers", path);
		expect(footers).toHaveLength(sections);
		// One shared part, not one per section.
		const pkg = await Pkg.open(path);
		expect(
			pkg.listParts().filter((p) => /footer\d+\.xml$/.test(p)),
		).toHaveLength(1);
	});

	test("--at sN targets a single section", async () => {
		const path = await freshFixture("hf-one", join(FIXTURES, "sections.docx"));
		await runCli("footers", "set", path, "--at", "s0", "--page-number");
		const footers = await list("footers", path);
		expect(footers).toHaveLength(1);
		expect(footers[0]?.sectionId).toBe("s0");
	});

	test("--at a non-section locator is rejected", async () => {
		const path = await newDoc("hf-badat");
		const result = await runCli(
			"headers",
			"set",
			path,
			"--at",
			"p0",
			"--text",
			"X",
		);
		expect(result.exitCode).toBeGreaterThan(0);
		expect(result.parsed).toMatchObject({ code: "INVALID_LOCATOR" });
	});
});

describe("headers/footers — read hints + clear", () => {
	test("read surfaces docx:header / docx:footer hints (importer drops them)", async () => {
		const path = await newDoc("hf-hint");
		await runCli("headers", "set", path, "--text", "Title");
		await runCli("footers", "set", path, "--page-number", "--of-pages");
		const md = await readMarkdown(path);
		expect(md).toContain('<!-- docx:header text="Title" -->');
		expect(md).toContain('<!-- docx:footer text="Page {page} of {pages}" -->');
	});

	test("first-page hint carries the type attribute", async () => {
		const path = await newDoc("hf-hint-first");
		await runCli("headers", "set", path, "--first-page", "--text", "Cover");
		expect(await readMarkdown(path)).toContain(
			'<!-- docx:header type="first" text="Cover" -->',
		);
	});

	test("per-section headers co-locate at each section's START, not bunched together", async () => {
		const path = await newDoc(
			"hf-colocate",
			"First section body.",
			"Second section body.",
		);
		// Inline break after p0 → s0 (governs the first section); trailing → s1.
		await runCli("sections", path, "--at", "p0", "--columns", "1");
		await runCli(
			"headers",
			"set",
			path,
			"--at",
			"s0",
			"--text",
			"Alpha Header",
		);
		await runCli("headers", "set", path, "--at", "s1", "--text", "Beta Header");
		const lines = (await readMarkdown(path)).split("\n");
		const firstBody = lines.findIndex((line) =>
			line.includes("First section body."),
		);
		const secondBody = lines.findIndex((line) =>
			line.includes("Second section body."),
		);
		const alpha = lines.findIndex((line) => line.includes("Alpha Header"));
		const beta = lines.findIndex((line) => line.includes("Beta Header"));
		// Each header sits at its section's START: s0's header before the first
		// section's body, s1's header after it (before the second section's body) —
		// so the first section's content sits BETWEEN the two hints, proving they
		// are NOT bunched together.
		expect(alpha).toBeGreaterThanOrEqual(0);
		expect(alpha).toBeLessThan(firstBody);
		expect(beta).toBeGreaterThan(firstBody);
		expect(beta).toBeLessThan(secondBody);
	});

	test("clear removes the reference", async () => {
		const path = await newDoc("hf-clear");
		await runCli("headers", "set", path, "--text", "H");
		expect(await list("headers", path)).toHaveLength(1);
		await runCli("headers", "clear", path);
		expect(await list("headers", path)).toHaveLength(0);
		expect(await partText(path, "word/document.xml")).not.toContain(
			"<w:headerReference",
		);
	});
});

describe("headers/footers — tracked changes", () => {
	test("tracked set records a sectPrChange; reject restores (removes the reference)", async () => {
		const path = await newDoc("hf-track");
		await runCli("footers", "set", path, "--page-number", "--track");
		expect(await partText(path, "word/document.xml")).toContain(
			"<w:sectPrChange",
		);
		const listed = (await runCli("track-changes", "list", path))
			.parsed as Array<{
			id: string;
			kind: string;
		}>;
		expect(listed.some((c) => c.kind === "sectPrChange")).toBe(true);
		expect(await list("footers", path)).toHaveLength(1);
		await runCli("track-changes", "reject", path, "--at", "tc0");
		expect(await list("footers", path)).toHaveLength(0);
	});

	test("accept keeps the reference and drops the snapshot", async () => {
		const path = await newDoc("hf-accept");
		await runCli("footers", "set", path, "--page-number", "--track");
		await runCli("track-changes", "accept", path, "--at", "tc0");
		expect(await list("footers", path)).toHaveLength(1);
		expect(await partText(path, "word/document.xml")).not.toContain(
			"<w:sectPrChange",
		);
	});
});

describe("headers/footers — validation", () => {
	test("set with no content is a USAGE error", async () => {
		const path = await newDoc("hf-empty");
		const result = await runCli("headers", "set", path);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("two field sources are rejected", async () => {
		const path = await newDoc("hf-twofields");
		const result = await runCli(
			"footers",
			"set",
			path,
			"--page-number",
			"--date",
		);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("invalid --type / --field / --align are rejected", async () => {
		const path = await newDoc("hf-invalid");
		expect(
			(await runCli("headers", "set", path, "--type", "bogus", "--text", "X"))
				.parsed,
		).toMatchObject({ code: "USAGE" });
		expect(
			(await runCli("footers", "set", path, "--field", "bogus")).parsed,
		).toMatchObject({ code: "USAGE" });
		expect(
			(await runCli("headers", "set", path, "--text", "X", "--align", "up"))
				.parsed,
		).toMatchObject({ code: "USAGE" });
	});

	test("blank --text sets an (empty) marginal", async () => {
		const path = await newDoc("hf-blank");
		const result = await runCli(
			"headers",
			"set",
			path,
			"--first-page",
			"--text",
			"",
		);
		expect(result.exitCode).toBe(0);
		expect((await list("headers", path))[0]?.text).toBe("");
	});
});

describe("create — header/footer flags", () => {
	test("--header / --page-numbers land on the trailing section", async () => {
		const path = join(tempWorkspace("hf-create"), "doc.docx");
		await runCli(
			"create",
			path,
			"--text",
			"Body.",
			"--header",
			"Report",
			"--page-numbers",
			"--force",
		);
		expect((await list("headers", path))[0]?.text).toBe("Report");
		expect((await list("footers", path))[0]?.text).toBe("{page}");
	});

	test("--footer and --page-numbers are mutually exclusive", async () => {
		const path = join(tempWorkspace("hf-create-mutex"), "doc.docx");
		const result = await runCli(
			"create",
			path,
			"--footer",
			"X",
			"--page-numbers",
			"--force",
		);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});
});

describe("headers/footers — review regressions", () => {
	// M1: after a doc-wide set, all sections share ONE part. A scoped `--at sN`
	// set must NOT overwrite that shared body (which silently rewrote every other
	// section) — it copy-on-writes a fresh part for the targeted section.
	test("scoped set forks the shared part — other sections keep the doc-wide content", async () => {
		const path = await freshFixture("hf-cow", join(FIXTURES, "sections.docx"));
		await runCli("headers", "set", path, "--text", "SHARED");
		await runCli("headers", "set", path, "--at", "s0", "--text", "ONLY-S0");
		const headers = await list("headers", path);
		expect(headers.find((h) => h.sectionId === "s0")?.text).toBe("ONLY-S0");
		for (const header of headers.filter((h) => h.sectionId !== "s0")) {
			expect(header.text).toBe("SHARED");
		}
		// Two parts now (shared + the s0 fork), both referenced — no dangling rId.
		const pkg = await Pkg.open(path);
		expect(
			pkg.listParts().filter((p) => /header\d+\.xml$/.test(p)).length,
		).toBeGreaterThanOrEqual(2);
		// Re-read round-trips cleanly.
		expect((await runCli("read", path)).exitCode).toBe(0);
	});

	// clear removes the per-section <w:titlePg/> once no first-page marginal remains
	// (else it keeps suppressing the page-1 default header/footer).
	test("clear --first-page removes the orphan <w:titlePg/>", async () => {
		const path = await newDoc("hf-clear-titlepg");
		await runCli("headers", "set", path, "--first-page", "--text", "Cover");
		expect(await partText(path, "word/document.xml")).toContain("<w:titlePg");
		await runCli("headers", "clear", path, "--first-page");
		expect(await partText(path, "word/document.xml")).not.toContain(
			"<w:titlePg",
		);
	});

	test("clear --first-page keeps <w:titlePg/> while a first-page footer remains", async () => {
		const path = await newDoc("hf-titlepg-keep");
		await runCli("headers", "set", path, "--first-page", "--text", "H");
		await runCli("footers", "set", path, "--first-page", "--text", "F");
		await runCli("headers", "clear", path, "--first-page");
		expect(await partText(path, "word/document.xml")).toContain("<w:titlePg");
	});

	// clear removes the document-level <w:evenAndOddHeaders/> once no even marginal
	// remains (else even pages render blank instead of inheriting the default).
	test("clear --even removes the orphan <w:evenAndOddHeaders/>", async () => {
		const path = await newDoc("hf-clear-even");
		await runCli("headers", "set", path, "--even", "--text", "Even");
		expect(await partText(path, "word/settings.xml")).toContain(
			"evenAndOddHeaders",
		);
		await runCli("headers", "clear", path, "--even");
		expect(await partText(path, "word/settings.xml")).not.toContain(
			"evenAndOddHeaders",
		);
	});

	// A pure content replace under --track records no sectPrChange (body edits
	// aren't individually tracked in v1) — the ack warns rather than misleading.
	test("content-only replace under --track flags trackedRevision: false", async () => {
		const path = await newDoc("hf-track-content");
		await runCli("headers", "set", path, "--text", "V1");
		await runCli("track-changes", "on", path);
		const result = await runCli(
			"headers",
			"set",
			path,
			"--text",
			"V2",
			"--track",
		);
		expect(result.parsed).toMatchObject({ trackedRevision: false });
	});

	// A marginal on only SOME sections must be labeled per-section (sN), not as
	// document chrome (which implied it covered sections it doesn't).
	test("a header on a subset of sections is labeled per-section, not document-wide", async () => {
		const path = await freshFixture(
			"hf-subset",
			join(FIXTURES, "sections.docx"),
		);
		await runCli("headers", "set", path, "--at", "s0", "--text", "OnlyS0");
		const md = await readMarkdown(path);
		expect(md).toContain('docx:header s0 text="OnlyS0"');
		expect(md).not.toContain('<!-- docx:header text="OnlyS0" -->');
	});

	// The two-zone right tab tracks the section's content width — at 0.5in margins
	// on letter that's 12240 − 720 − 720 = 10800tw (not the default 9360).
	test("two-zone right tab tracks non-default margins", async () => {
		const path = await newDoc("hf-twozone-geom");
		await runCli("sections", path, "--margins", "0.5");
		await runCli("headers", "set", path, "--text", "Title", "--page-number");
		expect(await partText(path, "word/header1.xml")).toContain(
			'w:val="right" w:pos="10800"',
		);
	});
});
