import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";

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
