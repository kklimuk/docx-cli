import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli } from "./harness";

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
});
