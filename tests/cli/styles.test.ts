import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";
import { freshFixture } from "./helpers";

// `docx styles` is the one query whose data isn't in the document body — the
// style catalog lives in word/styles.xml. It tells an authoring agent which
// `--style NAME` values exist and what a style looks like. Text-first; `--json`
// for the structured form (the harness does NOT auto-inject --json here, so the
// text default is testable directly).

const FIXTURES = join(import.meta.dir, "..", "fixtures");
const fixture = (name: string): string => join(FIXTURES, name);

describe("docx styles", () => {
	test("lists the full catalog as a text table (id / type / name)", async () => {
		const result = await runCli("styles", fixture("resume-styling.docx"));
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toMatch(/^Normal\s+paragraph/m);
		// A character style is typed distinctly from a paragraph style.
		expect(result.stdout).toMatch(/Hyperlink\s+character/);
	});

	test("--used filters to styles actually applied in the body", async () => {
		const path = fixture("resume-styling.docx");
		const all = (await runCli("styles", path)).stdout;
		const used = (await runCli("styles", path, "--used")).stdout;
		// The resume applies a custom paragraph style — it must show under --used.
		expect(used).toContain("SD-BodyText9pt");
		// --used is a strict subset of the full catalog.
		expect(used.trim().split("\n").length).toBeLessThan(
			all.trim().split("\n").length,
		);
		// A defined-but-unapplied style (Normal is always defined) may be absent
		// from --used; the catalog always has it.
		expect(all).toContain("Normal");
	});

	test("--at describes one style (name, basedOn, key formatting)", async () => {
		const result = await runCli(
			"styles",
			fixture("academic-paper.docx"),
			"--at",
			"Heading1",
		);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toMatch(/^Heading1 \(paragraph\)/);
		expect(result.stdout).toContain("name:");
		expect(result.stdout).toContain("basedOn:");
	});

	test("--at on an unknown style id fails with NOT_FOUND", async () => {
		const result = await runCli(
			"styles",
			fixture("resume-styling.docx"),
			"--at",
			"NoSuchStyle",
		);
		expect(result.exitCode).toBe(3);
		expect((result.parsed as { code: string }).code).toBe("BLOCK_NOT_FOUND");
	});

	test("--json emits a structured array for the list", async () => {
		const result = await runCli(
			"styles",
			fixture("resume-styling.docx"),
			"--json",
		);
		const styles = result.parsed as Array<{
			id: string;
			type: string;
			name?: string;
			basedOn?: string;
		}>;
		expect(Array.isArray(styles)).toBe(true);
		const normal = styles.find((s) => s.id === "Normal");
		expect(normal?.type).toBe("paragraph");
		// basedOn is carried when present (a derived style references its parent).
		expect(styles.some((s) => typeof s.basedOn === "string")).toBe(true);
	});

	test("--json --at emits a single style-detail object", async () => {
		const result = await runCli(
			"styles",
			fixture("academic-paper.docx"),
			"--at",
			"Heading1",
			"--json",
		);
		const detail = result.parsed as { id: string; type: string };
		expect(detail.id).toBe("Heading1");
		expect(detail.type).toBe("paragraph");
	});

	// Bold/italic are OOXML *toggle* properties: `<w:b w:val="0"/>` means bold is
	// explicitly OFF (a style that un-bolds a bold parent). A presence-only check
	// would lie with bold:true. FigureNumber in large-mixed.docx has <w:b w:val="0"/>.
	test("--at resolves a toggle-off (w:b w:val=0) to bold:false, not bold:true", async () => {
		const result = await runCli(
			"styles",
			fixture("large-mixed.docx"),
			"--at",
			"FigureNumber",
			"--json",
		);
		const detail = result.parsed as { bold?: boolean };
		expect(detail.bold).toBe(false);
	});

	// Tables apply a style via <w:tblPr><w:tblStyle w:val>; a presence in the body
	// that the paragraph/run walk never sees. mnda.docx applies TableGrid.
	test("--used includes table styles applied via <w:tblStyle>", async () => {
		const path = fixture("mnda.docx");
		const used = (await runCli("styles", path, "--used")).stdout;
		expect(used).toMatch(/^TableGrid\s+table/m);
		// And the structured form carries it too.
		const usedJson = (await runCli("styles", path, "--used", "--json"))
			.parsed as Array<{ id: string; type: string }>;
		expect(
			usedJson.some((s) => s.id === "TableGrid" && s.type === "table"),
		).toBe(true);
	});

	test("--catalog lists the provisionable built-ins, no FILE needed", async () => {
		const result = await runCli("styles", "--catalog");
		expect(result.exitCode).toBe(0);
		// The styles an agent had to guess at before — Title and the deep headings —
		// are now discoverable even though no document defines them.
		expect(result.stdout).toMatch(/^Title\s+paragraph/m);
		expect(result.stdout).toMatch(/^Subtitle\s+paragraph/m);
		expect(result.stdout).toMatch(/^Heading9\s+paragraph/m);
	});

	test("--catalog --json emits the structured catalog", async () => {
		const result = await runCli("styles", "--catalog", "--json");
		expect(result.exitCode).toBe(0);
		const ids = (result.parsed as Array<{ id: string }>).map((s) => s.id);
		expect(ids).toEqual(
			expect.arrayContaining([
				"Title",
				"Subtitle",
				"Heading7",
				"Heading8",
				"Heading9",
			]),
		);
	});
});

describe("docx styles set-default-font", () => {
	async function docWithHeading(label: string): Promise<string> {
		const docPath = join(tempWorkspace(label), "out.docx");
		await runCli("create", docPath, "--text", "Body paragraph.");
		// A Heading1 pins its own explicit font (Calibri Light) — the override case.
		await runCli(
			"insert",
			docPath,
			"--at-start",
			"--text",
			"Heading",
			"--style",
			"Heading1",
		);
		return docPath;
	}

	const readPart = async (docPath: string, part: string): Promise<string> =>
		(await Pkg.open(docPath)).readText(part);

	test("sets docDefaults + theme fonts, preserving explicit-font styles", async () => {
		const docPath = await docWithHeading("font-default");
		const result = await runCli(
			"styles",
			"set-default-font",
			docPath,
			"Times New Roman",
		);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			ok: true,
			operation: "styles.set-default-font",
			font: "Times New Roman",
			themeUpdated: true,
		});
		// Heading1 pins its own font, so it's reported (not silently left different).
		expect(
			(result.parsed as { explicitStyles: string[] }).explicitStyles,
		).toContain("Heading1");

		const styles = await readPart(docPath, "word/styles.xml");
		expect(styles).toMatch(/<w:docDefaults>[\s\S]*w:ascii="Times New Roman"/);
		// Explicit (theme-attrs dropped) so it beats the theme.
		expect(styles).not.toMatch(/<w:docDefaults>[\s\S]*w:asciiTheme/);
		// Heading1's own font is untouched in the default scope.
		expect(styles).toMatch(/Heading1[\s\S]*?w:ascii="Calibri Light"/);

		const theme = await readPart(docPath, "word/theme/theme1.xml");
		// Both major (headings) and minor (body) latin typefaces became the font.
		expect([
			...theme.matchAll(/<a:latin typeface="Times New Roman"/g),
		]).toHaveLength(2);
	});

	test("--all repoints explicit style fonts too", async () => {
		const docPath = await docWithHeading("font-all");
		const result = await runCli(
			"styles",
			"set-default-font",
			docPath,
			"Georgia",
			"--all",
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { repointed: number }).repointed).toBeGreaterThan(
			0,
		);

		const styles = await readPart(docPath, "word/styles.xml");
		expect(styles).toMatch(/Heading1[\s\S]*?w:ascii="Georgia"/);
		expect(styles).not.toContain("Calibri Light"); // the only explicit font, now gone
	});

	test("--size sets the default size (points → half-points)", async () => {
		const docPath = await docWithHeading("font-size");
		await runCli(
			"styles",
			"set-default-font",
			docPath,
			"Calibri",
			"--size",
			"16",
		);
		const styles = await readPart(docPath, "word/styles.xml");
		expect(styles).toMatch(/<w:docDefaults>[\s\S]*<w:sz w:val="32"/);
	});

	test("--dry-run writes nothing", async () => {
		const docPath = await docWithHeading("font-dry");
		const before = (await Bun.file(docPath).arrayBuffer()).byteLength;
		const result = await runCli(
			"styles",
			"set-default-font",
			docPath,
			"Arial",
			"--dry-run",
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { dryRun?: boolean }).dryRun).toBe(true);
		expect((result.parsed as { ok?: boolean }).ok).toBeUndefined();
		// The preview must surface which styles stay off-font (same as the real run),
		// so an agent can decide on --all without mutating the file first.
		expect(
			(result.parsed as { explicitStyles: string[] }).explicitStyles,
		).toContain("Heading1");
		expect((await Bun.file(docPath).arrayBuffer()).byteLength).toBe(before);
	});

	test("requires a FONT name", async () => {
		const docPath = await docWithHeading("font-missing");
		const result = await runCli("styles", "set-default-font", docPath);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});
});

const readStyles = async (docPath: string): Promise<string> =>
	(await Pkg.open(docPath)).readText("word/styles.xml");

/** A doc whose body uses Heading1 (so the baseline def is materialized in
 *  styles.xml) plus a plain body paragraph. */
async function docWithHeading1(label: string): Promise<string> {
	const docPath = join(tempWorkspace(label), "out.docx");
	await runCli("create", docPath, "--text", "Body paragraph.");
	await runCli(
		"insert",
		docPath,
		"--at-start",
		"--text",
		"Heading",
		"--style",
		"Heading1",
	);
	return docPath;
}

describe("docx styles set", () => {
	test("edits an existing style's run formatting (read-back round-trips)", async () => {
		const docPath = await docWithHeading1("set-run");
		const result = await runCli(
			"styles",
			"set",
			docPath,
			"--at",
			"Heading1",
			"--color",
			"1F4E79",
			"--size",
			"18",
			"--bold",
		);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			ok: true,
			operation: "styles.set",
			id: "Heading1",
			type: "paragraph",
		});

		// The write-read loop: `styles --at` reflects the new properties.
		const detail = (
			await runCli("styles", docPath, "--at", "Heading1", "--json")
		).parsed as { color?: string; sizePt?: number; bold?: boolean };
		expect(detail.color).toBe("1F4E79");
		expect(detail.sizePt).toBe(18);
		expect(detail.bold).toBe(true);

		// Size sets both <w:sz> and <w:szCs> (18pt = 36 half-points).
		const styles = await readStyles(docPath);
		expect(styles).toMatch(/Heading1[\s\S]*?<w:sz w:val="36"\/>/);
		expect(styles).toMatch(/Heading1[\s\S]*?<w:szCs w:val="36"\/>/);
	});

	// Regression: a pretty-printed styles.xml — inter-element whitespace kept as
	// `#text` nodes, which is what Word / LibreOffice / third-party generators emit
	// — must not corrupt CT_RPr order. The order-splice helpers rank only real
	// element siblings; a leading whitespace node used to outrank the new child and
	// shove it to the front (<w:color> ahead of <w:rFonts>), which Word rejects as
	// "unreadable content." Our own output is single-line, so this only bit real
	// user inputs — hence the explicit pretty-print here.
	test("preserves CT_RPr order on a pretty-printed styles.xml", async () => {
		const docPath = await docWithHeading1("set-prettyprint");
		const pkg = await Pkg.open(docPath);
		const pretty = (await pkg.readText("word/styles.xml"))
			.replace(/<w:rPr>/g, "<w:rPr>\n      ")
			.replace(/\/><w:/g, "/>\n      <w:");
		pkg.writeText("word/styles.xml", pretty);
		await pkg.save();

		// Insert a property Heading1 lacks: <w:i> ranks between <w:b> and <w:color>.
		const result = await runCli(
			"styles",
			"set",
			docPath,
			"--at",
			"Heading1",
			"--italic",
		);
		expect(result.exitCode).toBe(0);

		const styles = await readStyles(docPath);
		const block = styles.match(
			/<w:style w:type="paragraph" w:styleId="Heading1">[\s\S]*?<\/w:style>/,
		)?.[0];
		const order = [
			...(block ?? "").matchAll(/<w:(rFonts|b|i|color|sz)\b/g),
		].map((match) => match[1]);
		// italic landed in its CT_RPr slot, not at the front of the rPr.
		expect(order).toEqual(["rFonts", "b", "i", "color", "sz"]);
		// And the change still round-trips (write-read loop intact).
		const detail = (
			await runCli("styles", docPath, "--at", "Heading1", "--json")
		).parsed as { italic?: boolean };
		expect(detail.italic).toBe(true);
	});

	// The describeStyle read-back was widened to surface every settable property so
	// the write-read loop holds for the WHOLE shared run+paragraph vocabulary, not
	// just bold/color/size. Pin the toggles and indent/spacing dims whose appliers
	// (applyRunFormatToRpr / applyParagraphPropsToPPr) are now shared with `edit` —
	// a regression in either, or in describeStyle, would otherwise pass silently.
	test("round-trips the full run + paragraph vocabulary", async () => {
		const docPath = await docWithHeading1("set-vocab");
		const result = await runCli(
			"styles",
			"set",
			docPath,
			"--at",
			"Heading1",
			"--underline",
			"--strike",
			"--caps",
			"--smallcaps",
			"--superscript",
			"--highlight",
			"yellow",
			"--line-spacing",
			"double",
			"--indent-right",
			"0.3",
			"--hanging",
			"0.25",
		);
		expect(result.exitCode).toBe(0);

		const detail = (
			await runCli("styles", docPath, "--at", "Heading1", "--json")
		).parsed as Record<string, unknown>;
		expect(detail).toMatchObject({
			underline: "single",
			strike: true,
			caps: true,
			smallCaps: true,
			vertAlign: "superscript",
			highlight: "yellow",
			lineSpacing: "2",
			indentRightIn: 0.3,
			hangingIn: 0.25,
		});

		// first-line is the firstLine/hanging alternative — cover it on its own doc.
		const flPath = await docWithHeading1("set-firstline");
		await runCli(
			"styles",
			"set",
			flPath,
			"--at",
			"Heading1",
			"--first-line",
			"0.4",
		);
		const fl = (await runCli("styles", flPath, "--at", "Heading1", "--json"))
			.parsed as { firstLineIn?: number };
		expect(fl.firstLineIn).toBe(0.4);
	});

	test("edits paragraph formatting on a paragraph style", async () => {
		const docPath = await docWithHeading1("set-para");
		await runCli(
			"styles",
			"set",
			docPath,
			"--at",
			"Heading1",
			"--alignment",
			"center",
			"--space-before",
			"24",
			"--indent-left",
			"0.5",
		);
		const detail = (
			await runCli("styles", docPath, "--at", "Heading1", "--json")
		).parsed as {
			alignment?: string;
			spaceBeforePt?: number;
			indentLeftIn?: number;
		};
		expect(detail.alignment).toBe("center");
		expect(detail.spaceBeforePt).toBe(24);
		expect(detail.indentLeftIn).toBe(0.5);
	});

	test("updates metadata (--name / --based-on)", async () => {
		const docPath = await docWithHeading1("set-meta");
		await runCli(
			"styles",
			"set",
			docPath,
			"--at",
			"Heading1",
			"--name",
			"Section Head",
			"--based-on",
			"Title",
		);
		const detail = (
			await runCli("styles", docPath, "--at", "Heading1", "--json")
		).parsed as { name?: string; basedOn?: string };
		expect(detail.name).toBe("Section Head");
		expect(detail.basedOn).toBe("Title");
	});

	test("auto-provisions an un-materialized baseline, then applies", async () => {
		const docPath = await docWithHeading1("set-prov");
		// Heading3 isn't used anywhere — set should provision then color it.
		const result = await runCli(
			"styles",
			"set",
			docPath,
			"--at",
			"Heading3",
			"--color",
			"00AA00",
		);
		expect(result.exitCode).toBe(0);
		const detail = (
			await runCli("styles", docPath, "--at", "Heading3", "--json")
		).parsed as { id: string; color?: string };
		expect(detail.id).toBe("Heading3");
		expect(detail.color).toBe("00AA00");
	});

	test("rejects paragraph flags on a character style", async () => {
		const docPath = await docWithHeading1("set-char-para");
		const result = await runCli(
			"styles",
			"set",
			docPath,
			"--at",
			"Hyperlink", // a baseline character style
			"--alignment",
			"center",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});

	test("unknown non-baseline id fails with NOT_FOUND", async () => {
		const docPath = await docWithHeading1("set-unknown");
		const result = await runCli(
			"styles",
			"set",
			docPath,
			"--at",
			"NoSuchStyle",
			"--bold",
		);
		expect(result.exitCode).toBe(3);
		expect((result.parsed as { code: string }).code).toBe("BLOCK_NOT_FOUND");
	});

	test("no formatting and no metadata is a USAGE error", async () => {
		const docPath = await docWithHeading1("set-empty");
		const result = await runCli("styles", "set", docPath, "--at", "Heading1");
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});

	// The empirical Word behavior (verified): editing a style definition is NOT a
	// tracked revision, even with track-changes ON. The edit lands directly in
	// styles.xml; the body gets no <w:ins>/<w:del>, styles.xml no rPrChange.
	test("style edits are untracked even with the doc toggle on", async () => {
		const docPath = await docWithHeading1("set-untracked");
		await runCli("track-changes", docPath, "on");
		// Edit BOTH run and paragraph formatting, so the rPrChange AND pPrChange
		// absence checks below are each actually exercised by a mutating input (an
		// --italic-only edit never touches pPr, so it can't prove the pPr path).
		await runCli(
			"styles",
			"set",
			docPath,
			"--at",
			"Heading1",
			"--italic",
			"--alignment",
			"center",
			"--space-before",
			"24",
		);

		const styles = await readStyles(docPath);
		expect(styles).toMatch(/Heading1[\s\S]*?<w:i\/>/); // run formatting applied directly
		expect(styles).toMatch(/Heading1[\s\S]*?<w:jc w:val="center"\/>/); // paragraph too
		expect(styles).not.toContain("rPrChange");
		expect(styles).not.toContain("pPrChange");

		const documentXml = await (await Pkg.open(docPath)).readText(
			"word/document.xml",
		);
		expect(documentXml).not.toMatch(/<w:ins[ >]/);
		expect(documentXml).not.toMatch(/<w:del[ >]/);

		// And the tracking engine sees zero changes from the style edit.
		const list = (await runCli("track-changes", "list", docPath))
			.parsed as unknown[];
		expect(list).toHaveLength(0);
	});

	test("--dry-run writes nothing", async () => {
		const docPath = await docWithHeading1("set-dry");
		const before = (await Bun.file(docPath).arrayBuffer()).byteLength;
		const result = await runCli(
			"styles",
			"set",
			docPath,
			"--at",
			"Heading1",
			"--bold",
			"--dry-run",
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { dryRun?: boolean }).dryRun).toBe(true);
		expect((result.parsed as { ok?: boolean }).ok).toBeUndefined();
		expect((await Bun.file(docPath).arrayBuffer()).byteLength).toBe(before);
	});
});

describe("docx styles create", () => {
	const createDoc = async (label: string): Promise<string> => {
		const docPath = join(tempWorkspace(label), "out.docx");
		await runCli("create", docPath, "--text", "Body paragraph.");
		return docPath;
	};

	test("defines a new paragraph style with metadata defaults + formatting", async () => {
		const docPath = await createDoc("create-para");
		const result = await runCli(
			"styles",
			"create",
			docPath,
			"Callout",
			"--name",
			"Callout",
			"--color",
			"C00000",
			"--bold",
			"--space-after",
			"6",
		);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			ok: true,
			operation: "styles.create",
			id: "Callout",
			type: "paragraph",
		});

		const detail = (
			await runCli("styles", docPath, "--at", "Callout", "--json")
		).parsed as {
			type: string;
			name?: string;
			basedOn?: string;
			color?: string;
			bold?: boolean;
			spaceAfterPt?: number;
		};
		expect(detail.type).toBe("paragraph");
		expect(detail.name).toBe("Callout");
		expect(detail.basedOn).toBe("Normal"); // paragraph default
		expect(detail.color).toBe("C00000");
		expect(detail.bold).toBe(true);
		expect(detail.spaceAfterPt).toBe(6);

		// CT_Style order: name → basedOn → next → qFormat → pPr → rPr.
		const styles = await readStyles(docPath);
		expect(styles).toMatch(
			/<w:style w:type="paragraph" w:styleId="Callout"><w:name[\s\S]*?<w:basedOn[\s\S]*?<w:next[\s\S]*?<w:qFormat\/><w:pPr>[\s\S]*?<\/w:pPr><w:rPr>[\s\S]*?<\/w:rPr><\/w:style>/,
		);
	});

	test("defines a character style (run formatting only, no pPr)", async () => {
		const docPath = await createDoc("create-char");
		await runCli(
			"styles",
			"create",
			docPath,
			"KbdKey",
			"--type",
			"character",
			"--font",
			"Consolas",
			"--shade",
			"EEEEEE",
		);
		const detail = (await runCli("styles", docPath, "--at", "KbdKey", "--json"))
			.parsed as { type: string; font?: string; shade?: string };
		expect(detail.type).toBe("character");
		expect(detail.font).toBe("Consolas");
		expect(detail.shade).toBe("EEEEEE");
		// A character style carries no <w:pPr>.
		const styles = await readStyles(docPath);
		expect(styles).toMatch(
			/<w:style w:type="character" w:styleId="KbdKey">[\s\S]*?<\/w:style>/,
		);
		const node = styles.match(
			/<w:style w:type="character" w:styleId="KbdKey">[\s\S]*?<\/w:style>/,
		)?.[0];
		expect(node).not.toContain("<w:pPr>");
	});

	test("the created style is usable via edit --style and round-trips in read", async () => {
		const docPath = await createDoc("create-usable");
		await runCli(
			"styles",
			"create",
			docPath,
			"Lead",
			"--italic",
			"--size",
			"13",
		);
		await runCli("edit", docPath, "--at", "p0", "--style", "Lead");
		const markdown = (await runCli("read", docPath)).stdout;
		expect(markdown).toContain('style="Lead"');
		const used = (await runCli("styles", docPath, "--used")).stdout;
		expect(used).toMatch(/^Lead\s+paragraph/m);
	});

	test("rejects an id that already exists (routes to set)", async () => {
		const docPath = await createDoc("create-dup");
		await runCli("styles", "create", docPath, "Callout", "--bold");
		const result = await runCli(
			"styles",
			"create",
			docPath,
			"Callout",
			"--italic",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
		expect((result.parsed as { hint?: string }).hint).toContain("styles set");
	});

	test("rejects a built-in id (routes to set)", async () => {
		const docPath = await createDoc("create-baseline");
		const result = await runCli(
			"styles",
			"create",
			docPath,
			"Heading1",
			"--bold",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
		expect((result.parsed as { hint?: string }).hint).toContain("styles set");
	});

	test("rejects a STYLEID containing whitespace", async () => {
		const docPath = await createDoc("create-ws");
		const result = await runCli(
			"styles",
			"create",
			docPath,
			"My Style",
			"--bold",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});

	test("rejects paragraph flags on a character style", async () => {
		const docPath = await createDoc("create-char-para");
		const result = await runCli(
			"styles",
			"create",
			docPath,
			"BadChar",
			"--type",
			"character",
			"--alignment",
			"center",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});

	test("rejects an invalid --type", async () => {
		const docPath = await createDoc("create-bad-type");
		const result = await runCli(
			"styles",
			"create",
			docPath,
			"Foo",
			"--type",
			"table",
			"--bold",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});

	test("--dry-run writes nothing", async () => {
		const docPath = await createDoc("create-dry");
		const before = (await Bun.file(docPath).arrayBuffer()).byteLength;
		const result = await runCli(
			"styles",
			"create",
			docPath,
			"Temp",
			"--bold",
			"--dry-run",
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { dryRun?: boolean }).dryRun).toBe(true);
		expect((await Bun.file(docPath).arrayBuffer()).byteLength).toBe(before);
	});
});

// Regressions caught by the tier-2.7 adversarial review.
describe("docx styles set/create — review hardening", () => {
	test("set --next round-trips through styles --at (write-read loop)", async () => {
		const docPath = await docWithHeading1("rh-next-set");
		await runCli(
			"styles",
			"set",
			docPath,
			"--at",
			"Heading1",
			"--next",
			"Title",
		);
		const detail = (
			await runCli("styles", docPath, "--at", "Heading1", "--json")
		).parsed as { next?: string };
		expect(detail.next).toBe("Title");
	});

	test("create --next round-trips through styles --at", async () => {
		const docPath = await docWithHeading1("rh-next-create");
		await runCli(
			"styles",
			"create",
			docPath,
			"Lead",
			"--next",
			"Normal",
			"--italic",
		);
		const detail = (await runCli("styles", docPath, "--at", "Lead", "--json"))
			.parsed as { next?: string };
		expect(detail.next).toBe("Normal");
	});

	// A table style is not a paragraph style: the ack must report its REAL type, and
	// paragraph flags must be rejected (not silently written into its pPr).
	test("set on a table style reports the real type and rejects paragraph flags", async () => {
		const docPath = await freshFixture("rh-table", fixture("mnda.docx"));
		const ok = await runCli(
			"styles",
			"set",
			docPath,
			"--at",
			"TableGrid",
			"--bold",
		);
		expect(ok.exitCode).toBe(0);
		expect((ok.parsed as { type: string }).type).toBe("table");

		const bad = await runCli(
			"styles",
			"set",
			docPath,
			"--at",
			"TableGrid",
			"--alignment",
			"center",
		);
		expect(bad.exitCode).toBe(2);
		expect((bad.parsed as { code: string }).code).toBe("USAGE");
		expect((bad.parsed as { error: string }).error).toContain("table");
	});

	test("set rejects --based-on equal to the style itself (cycle)", async () => {
		const docPath = await docWithHeading1("rh-self-set");
		const result = await runCli(
			"styles",
			"set",
			docPath,
			"--at",
			"Heading1",
			"--based-on",
			"Heading1",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});

	test("create rejects --based-on equal to the new style id (cycle)", async () => {
		const docPath = await docWithHeading1("rh-self-create");
		const result = await runCli(
			"styles",
			"create",
			docPath,
			"Foo",
			"--based-on",
			"Foo",
			"--bold",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});

	// An explicit empty --based-on must fall back to the Normal default, not write a
	// dangling <w:basedOn w:val=""/>.
	test("create --based-on '' falls back to the Normal default", async () => {
		const docPath = await docWithHeading1("rh-empty-basedon");
		await runCli(
			"styles",
			"create",
			docPath,
			"Bar",
			"--based-on",
			"",
			"--bold",
		);
		const detail = (await runCli("styles", docPath, "--at", "Bar", "--json"))
			.parsed as { basedOn?: string };
		expect(detail.basedOn).toBe("Normal");
		const styles = await (await Pkg.open(docPath)).readText("word/styles.xml");
		expect(styles).not.toContain('w:basedOn w:val=""');
	});
});
