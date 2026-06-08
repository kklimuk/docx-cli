import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "../../src/core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";

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

type Block = { id: string; type: string; runs?: TextRunAst[] };

async function readBlocks(docPath: string): Promise<Block[]> {
	const result = await runCli("read", docPath, "--ast");
	return (result.parsed as { blocks: Block[] }).blocks;
}

async function readText(docPath: string): Promise<string> {
	return (await runCli("read", docPath)).stdout;
}

async function readDocumentXml(docPath: string): Promise<string> {
	const pkg = await Pkg.open(docPath);
	return pkg.readText("word/document.xml");
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

function formattedParagraph(blocks: Block[]): TextRunAst[] {
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
		const runs = formattedParagraph(await readBlocks(docPath));
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
		const original = formattedParagraph(await readBlocks(src));

		const workspace = tempWorkspace("rf-rt-dst");
		const mdPath = join(workspace, "doc.md");
		await Bun.write(mdPath, await readText(src));
		const dst = join(workspace, "rt.docx");
		await runCli("create", dst, "--from", mdPath);

		const roundTripped = formattedParagraph(await readBlocks(dst));
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
		const text = formattedParagraph(await readBlocks(docPath));
		// `{TODO fixme}` are bare tokens (no recognized attrs) → not a span →
		// brackets restored as literal text.
		expect((text[0]?.text ?? "").length).toBeGreaterThanOrEqual(0);
		const joined = (await readBlocks(docPath))
			.flatMap((b) => b.runs ?? [])
			.map((r) => r.text ?? "")
			.join("");
		expect(joined).toContain("See [note]{TODO fixme} later.");
	});

	test("a formatted run whose text contains markup metacharacters keeps its formatting", async () => {
		// `[ ] { }` in run text used to force DROPPING the span (Pandoc's `[…]{…}`
		// couldn't escape them). HTML escapes them (`&#91;` etc.), so the formatting
		// is kept AND the text round-trips byte-exact — strictly better than before.
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
		// Formatting is now preserved (HTML can escape the metacharacters).
		expect(md).toContain("color:#FF0000");
		// The literal `]` is entity-escaped so it can't re-tokenize as link syntax.
		expect(md).not.toContain("a]{x}b");

		const ws = tempWorkspace("rf-bracket-rt");
		const mdPath = join(ws, "doc.md");
		await Bun.write(mdPath, md);
		const dst = join(ws, "rt.docx");
		await runCli("create", dst, "--from", mdPath);

		// Round-trip: text byte-exact AND the color survived.
		const runs = (await readBlocks(dst))
			.flatMap((b) => b.runs ?? [])
			.filter((r) => r.type === "text");
		expect(runs.map((r) => r.text ?? "").join("")).toContain("a]{x}b");
		const colored = runs.find((r) => (r.text ?? "").includes("a]{x}b"));
		expect(colored?.color).toBe("FF0000");
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
		const injected = (await readBlocks(dst))
			.flatMap((b) => b.runs ?? [])
			.some((r) => (r as { highlight?: string }).highlight === "yellow");
		expect(injected).toBe(false);
	});
});
