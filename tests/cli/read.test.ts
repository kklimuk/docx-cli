import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

const FIXTURES = join(import.meta.dir, "..", "fixtures");
const fixture = (name: string): string => join(FIXTURES, name);

async function render(...args: string[]): Promise<string> {
	const result = await runCli("read", ...args);
	expect(result.exitCode).toBe(0);
	return result.stdout;
}

describe("docx read (markdown)", () => {
	test("minimal: emits heading, bold, locator comments", async () => {
		const out = await render(fixture("minimal.docx"));
		expect(out).toContain("# Style Guide <!-- p0 -->");
		expect(out).toContain("**important**");
		expect(out).toMatch(/Use[^\n]*\*\*important\*\*[^\n]*terms/);
		expect(out).toContain("<!-- p2 -->");
		// The trailing mandatory section break is suppressed (it's implicit OOXML
		// structure that create re-adds, so emitting `---` would re-import as a
		// stray thematic-break paragraph that accretes each round-trip).
		expect(out).not.toContain("--- <!-- s0 -->");
	});

	test("academic-paper: heading levels 1 and 2 render with # and ##", async () => {
		const out = await render(fixture("academic-paper.docx"));
		expect(out).toMatch(/^# Guided Imagery and Progressive Muscle Relaxation/m);
		expect(out).toMatch(/^## Features of Guided Imagery/m);
		expect(out).toMatch(/^## Guided Imagery in Group Psychotherapy/m);
	});

	test("academic-paper: italics and hyperlinks render inline", async () => {
		const out = await render(fixture("academic-paper.docx"));
		expect(out).toContain("*Guided imagery*");
		expect(out).toContain("*Progressive muscle relaxation*");
	});

	test("tables-and-lists: pipe table with cell locators", async () => {
		const out = await render(fixture("tables-and-lists.docx"));
		expect(out).toMatch(/^\| \*\*Equipment\*\* <!-- t0:r0c0:p0 --> \|/m);
		expect(out).toMatch(/^\| --- \| --- \|$/m);
		expect(out).toContain("Agilent E3631A Triple Output DC Power Supply");
		expect(out).toContain("9.1 Ω Resistor");
	});

	test("lists fixture: bullet vs ordered hierarchy renders correctly", async () => {
		const out = await render(fixture("lists.docx"));
		// Bullets at nested levels
		expect(out).toMatch(/^- Apples/m);
		expect(out).toMatch(/^ {2}- Granny Smith/m);
		expect(out).toMatch(/^ {4}- From the orchard down the road/m);
		// Ordered items carry their real per-level ordinal; the first item at each
		// nesting level is 1 (deeper levels reset when a shallower one advances).
		expect(out).toMatch(/^1\. Preheat the oven/m);
		expect(out).toMatch(/^ {2}1\. Sift the flour/m);
		expect(out).toMatch(/^ {4}1\. Don't overmix/m);
		// The second bullet list (independent numId) still renders as bullets
		expect(out).toMatch(/^- Independent item one/m);
	});

	test("tables-and-lists: top-level ordered items render with real ordinals", async () => {
		// These are Word-numbered items (numFmt=decimal), so they render as an
		// ordered list with their real running ordinal (1. 2. 3.) — so the RAW
		// markdown reads correctly rather than as a wall of `1.`.
		const out = await render(fixture("tables-and-lists.docx"));
		expect(out).toMatch(/^1\. \*\*Introduction\*\*/m);
		expect(out).toMatch(/^2\. \*\*Background Information\*\*/m);
		expect(out).toMatch(/^3\. \*\*Methods and Materials\*\*/m);
	});

	test("resume-styling: list bullets and bold runs render", async () => {
		const out = await render(fixture("resume-styling.docx"));
		expect(out).toMatch(/^- GPA: 3\.5/m);
		expect(out).toContain("**JANE SMITH**");
		expect(out).toContain("[j1smith@business.rutgers.edu](mailto:");
	});

	test("multi-column: SIGCHI three-column table renders as pipe table", async () => {
		const out = await render(fixture("multi-column.docx"));
		expect(out).toMatch(/^\| Leave Authors Anonymous <!-- t0:r0c0:p0 -->/m);
		expect(out).toMatch(/^\| --- \| --- \| --- \|$/m);
		expect(out).toContain("<br>for Submission <!-- t0:r0c0:p1 -->");
	});

	test("large-mixed: images render as ![alt](<sha256>.<ext>) and locator survives", async () => {
		const out = await render(fixture("large-mixed.docx"));
		// Each image surfaces as a content-addressed URL: `<64-hex>.<ext>`.
		// The walker uses the hash on round-trip to reuse the existing
		// media part instead of re-fetching/duplicating.
		const hashMatches = out.match(/!\[[^\]]*\]\([0-9a-f]{64}\.[a-z0-9]+\)/g);
		expect(hashMatches?.length ?? 0).toBeGreaterThanOrEqual(2);
		expect(out).toContain("# Chinese Folding Fan Design Project");
	});

	test("equations: OMML walks to real LaTeX in `$…$` / `$$…$$` mid-paragraph", async () => {
		const out = await render(fixture("equations.docx"));
		// Inline atom equations on their own paragraphs ("Atom: superscript $x^2$"
		// shape — header text followed by the equation render).
		expect(out).toContain("$x^2$");
		expect(out).toContain("$E=mc^2$");
		// A famous formula round-tripped to its canonical LaTeX form.
		expect(out).toContain("$\\frac{a}{b}$");
		// Display: quadratic formula on its own block line.
		expect(out).toMatch(/\$\$x=\\frac\{-b\\pm \\sqrt\{b\^2-4ac\}\}\{2a\}\$\$/);
	});

	test("strict-profile: chart/smartart/drawing placeholders render", async () => {
		const out = await render(fixture("strict-profile.docx"));
		expect(out).toContain("`[chart]`");
		expect(out).toContain("`[smartart]`");
		expect(out).toContain("`[drawing]`");
	});

	test("locator pins use HTML comments, not <sup>", async () => {
		const out = await render(fixture("minimal.docx"));
		expect(out).not.toContain("<sup>");
		expect(out).toMatch(/<!-- p\d+ -->/);
	});
});

describe("docx read (markdown) --from / --to", () => {
	test("--from p1 starts at p1, drops p0", async () => {
		const out = await render(fixture("minimal.docx"), "--from", "p1");
		expect(out).not.toContain("# Style Guide");
		expect(out).toContain("**important**");
		expect(out).toMatch(/Use[^\n]*\*\*important\*\*[^\n]*terms/);
	});

	test("--to is inclusive", async () => {
		const out = await render(fixture("minimal.docx"), "--to", "p1");
		expect(out).toContain("# Style Guide");
		expect(out).toContain("**important**");
		expect(out).toMatch(/Use[^\n]*\*\*important\*\*[^\n]*terms/);
		expect(out).not.toContain("The quick brown fox");
	});

	test("--from p13 --to p15 slices the middle of academic-paper", async () => {
		const out = await render(
			fixture("academic-paper.docx"),
			"--from",
			"p13",
			"--to",
			"p15",
		);
		expect(out).toContain("Group psychotherapy effectively promotes");
		expect(out).toContain("# Guided Imagery <!-- p15 -->");
		expect(out).not.toContain("Hannah K. Greenbaum");
	});

	test("--from accepts a span locator (offsets ignored)", async () => {
		const out = await render(fixture("minimal.docx"), "--from", "p1:0-5");
		expect(out).toContain("**important**");
		expect(out).toMatch(/Use[^\n]*\*\*important\*\*[^\n]*terms/);
		expect(out).not.toContain("# Style Guide");
	});

	test("--from accepts a range locator (uses start paragraph)", async () => {
		const out = await render(fixture("minimal.docx"), "--from", "p1:0-p2:3");
		expect(out).toContain("**important**");
		expect(out).toContain("The quick brown fox");
		expect(out).not.toContain("# Style Guide");
	});

	test("--from cN rejected as invalid locator", async () => {
		const result = await runCli(
			"read",
			fixture("minimal.docx"),
			"--from",
			"c0",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({
			code: "INVALID_LOCATOR",
		});
	});

	test("--from references an unknown block id", async () => {
		const result = await runCli(
			"read",
			fixture("minimal.docx"),
			"--from",
			"p99",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({
			code: "INVALID_LOCATOR",
		});
	});

	test("--from with --ast is rejected (Markdown-only flag)", async () => {
		const result = await runCli(
			"read",
			fixture("minimal.docx"),
			"--from",
			"p1",
			"--ast",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});
});

describe("docx read (markdown) tracked changes", () => {
	test("default markdown render is the accepted view (ins as plain, del gone, no markers)", async () => {
		const out = await render(fixture("tracked-changes.docx"));
		expect(out).toContain("two exciting");
		expect(out).not.toContain("{++");
		expect(out).not.toContain("{--");
		expect(out).not.toMatch(/\[\^tc\d+\]/);
		expect(out).not.toContain("<ins>");
		expect(out).not.toContain("<del>");
	});

	test("--current renders insertions as CriticMarkup {++...++} with [^tcN] refs and an appendix", async () => {
		const out = await render(fixture("tracked-changes.docx"), "--current");
		expect(out).toContain("{++two exciting ++}[^tc0]insertions");
		expect(out).toContain(
			"[^tc0]: insertion by eng-dept (2014-06-25T10:40:00Z)",
		);
		expect(out).not.toContain("<ins>");
		expect(out).not.toContain("<del>");
	});

	test("--accepted shows post-accept view (ins as plain, del gone, no markers)", async () => {
		const out = await render(fixture("tracked-changes.docx"), "--accepted");
		expect(out).toContain("two exciting");
		expect(out).not.toContain("{++");
		expect(out).not.toContain("{--");
		expect(out).not.toMatch(/\[\^tc\d+\]/);
		expect(out).not.toContain("<ins>");
		expect(out).not.toContain("<del>");
	});

	test("--baseline shows pre-change view (ins gone, del as plain, no markers)", async () => {
		const out = await render(fixture("tracked-changes.docx"), "--baseline");
		// Fixture has only an insertion ("two exciting"), so baseline drops it.
		expect(out).toContain("This is a text with insertions.");
		expect(out).not.toContain("two exciting");
		expect(out).not.toContain("{++");
		expect(out).not.toContain("{--");
		expect(out).not.toMatch(/\[\^tc\d+\]/);
	});

	test("--accepted and --baseline together are rejected", async () => {
		const result = await runCli(
			"read",
			fixture("tracked-changes.docx"),
			"--accepted",
			"--baseline",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("AST exposes stable tcN id on tracked-change runs", async () => {
		const result = await runCli(
			"read",
			fixture("tracked-changes.docx"),
			"--ast",
		);
		expect(result.exitCode).toBe(0);
		const doc = result.parsed as {
			blocks: Array<{
				type: string;
				runs?: Array<{
					type: string;
					trackedChange?: { id: string; kind: string };
				}>;
			}>;
		};
		const changes = doc.blocks
			.flatMap((block) => block.runs ?? [])
			.filter((run) => run.trackedChange !== undefined)
			.map((run) => run.trackedChange);
		expect(changes.length).toBeGreaterThan(0);
		expect(changes[0]?.id).toBe("tc0");
		expect(changes[0]?.kind).toBe("ins");
	});
});

describe("docx read (markdown) --comments", () => {
	test("default view (no --comments) emits no [^cN] refs and no footnotes", async () => {
		const out = await render(fixture("comments-with-replies.docx"));
		expect(out).not.toMatch(/\[\^c\d+\]/);
		expect(out).not.toContain("[^c0]:");
	});

	test("--comments inlines [^cN] at span end and emits GFM footnotes", async () => {
		const out = await render(
			fixture("comments-with-replies.docx"),
			"--comments",
		);
		expect(out).toContain("[^c0]");
		expect(out).toContain("[^c1]");
		expect(out).toContain('[^c0]: "some text to have a comment"');
		expect(out).toContain("Jesse Rosenthal (2016-05-09T16:13:00Z)");
		expect(out).toContain(": I left a comment.");
	});

	test("--comments overlapping anchors stack: [^cP][^cQ]", async () => {
		const out = await render(
			fixture("comments-with-replies.docx"),
			"--comments",
		);
		expect(out).toMatch(/comment in a comment\[\^c3\]\[\^c4\]/);
	});

	test("--comments reply footnote shows ↳ parent", async () => {
		const out = await render(
			fixture("comments-with-replies.docx"),
			"--comments",
		);
		expect(out).toContain("↳ c3:");
	});

	test("--comments on doc without comments emits no footnote section", async () => {
		const out = await render(fixture("strict-profile.docx"), "--comments");
		expect(out).not.toContain("[^c");
	});

	test("--comments + --from drops footnotes whose spans are out of range", async () => {
		const out = await render(
			fixture("comments-with-replies.docx"),
			"--comments",
			"--from",
			"p2",
		);
		expect(out).not.toContain("[^c0]:");
		expect(out).not.toContain("[^c1]:");
		expect(out).toContain("[^c2]:");
	});

	test("--comments works on simple comments fixture", async () => {
		const out = await render(fixture("comments-simple.docx"), "--comments");
		expect(out).toContain("[^c0]:");
	});
});

describe("docx read (markdown) equations", () => {
	test("inline `<m:oMath>` renders as `$…$` mid-paragraph", async () => {
		const out = await render(fixture("equations.docx"));
		// `\hat{n}` — accent atom; matches the inline `$\hat{n}$` shape.
		expect(out).toContain("$\\hat{n}$");
	});

	test("display `<m:oMathPara>` renders on its own line with `$$…$$`", async () => {
		const out = await render(fixture("equations.docx"));
		// Quadratic formula is the first display equation that's pure math (no
		// surrounding prose), so it renders on its own block line.
		expect(out).toMatch(
			/^\$\$x=\\frac\{-b\\pm \\sqrt\{b\^2-4ac\}\}\{2a\}\$\$/m,
		);
	});

	test("equations marked display=true survive in JSON AST", async () => {
		const result = await runCli("read", fixture("equations.docx"), "--ast");
		const doc = result.parsed as {
			blocks: Array<{
				type: string;
				runs?: Array<{ type: string; display?: boolean; text?: string }>;
			}>;
		};
		const equations = doc.blocks
			.flatMap((b) => b.runs ?? [])
			.filter((r) => r.type === "equation");
		expect(equations.length).toBeGreaterThan(0);
		const display = equations.filter((e) => e.display === true);
		const inline = equations.filter((e) => e.display === false);
		expect(display.length).toBeGreaterThan(0);
		expect(inline.length).toBeGreaterThan(0);
	});
});

describe("docx read (markdown) footnotes / endnotes", () => {
	test("notes.docx: footnote ref + definition both render", async () => {
		const out = await render(fixture("notes.docx"));
		expect(out).toContain("[^fn1]");
		expect(out).toMatch(/\n\[\^fn1\]: My note\./);
	});

	test("notes.docx: endnote ref + definition both render", async () => {
		const out = await render(fixture("notes.docx"));
		expect(out).toContain("[^en1]");
		expect(out).toContain(
			"[^en1]: This is an endnote at the end of the document.",
		);
	});

	test("notes.docx: footnote and endnote refs both visible inline", async () => {
		const out = await render(fixture("notes.docx"));
		expect(out).toMatch(/Test footnote\.\[\^fn1\] Test endnote\.\[\^en1\]/);
	});

	test("doc without footnote refs in body emits no definitions", async () => {
		const out = await render(fixture("minimal.docx"));
		expect(out).not.toContain("[^fn");
		expect(out).not.toContain("[^en");
	});

	test("footnotes/endnotes arrays present in JSON AST", async () => {
		const result = await runCli("read", fixture("notes.docx"), "--ast");
		const doc = result.parsed as {
			footnotes: Array<{ id: string; text: string }>;
			endnotes: Array<{ id: string; text: string }>;
		};
		expect(doc.footnotes.length).toBe(1);
		expect(doc.endnotes.length).toBe(1);
		expect(doc.footnotes[0]?.id).toBe("fn1");
		expect(doc.endnotes[0]?.id).toBe("en1");
	});

	test("Word's reserved separator/continuation entries are filtered out", async () => {
		const result = await runCli("read", fixture("notes.docx"), "--ast");
		const doc = result.parsed as {
			footnotes: Array<{ id: string }>;
		};
		const hasNegativeId = doc.footnotes.some(
			(f) => f.id.startsWith("fn-") || f.id === "fn0",
		);
		expect(hasNegativeId).toBe(false);
	});
});

describe("docx read (markdown) color and highlight", () => {
	test("non-default run color emits a `<span style>` a markdown reader renders", async () => {
		const out = await render(fixture("minimal.docx"));
		expect(out).toContain('<span style="color:#800080">');
		// No legacy Pandoc bracketed spans.
		expect(out).not.toContain('{color="800080"}');
	});

	test("color span wraps the bold markdown (emphasis nests inside the span)", async () => {
		const out = await render(fixture("minimal.docx"));
		// "important" is bold + purple in minimal.docx
		expect(out).toContain('<span style="color:#800080">**important**</span>');
	});
});

describe("docx read (markdown) chart / drawing placeholders", () => {
	test("chart drawing renders as `[chart]`", async () => {
		const out = await render(fixture("strict-profile.docx"));
		expect(out).toContain("`[chart]`");
	});

	test("smartart and generic drawing placeholders render too", async () => {
		const out = await render(fixture("strict-profile.docx"));
		expect(out).toContain("`[smartart]`");
		expect(out).toContain("`[drawing]`");
	});

	test("chart placeholders survive in JSON AST", async () => {
		const result = await runCli(
			"read",
			fixture("strict-profile.docx"),
			"--ast",
		);
		const doc = result.parsed as {
			blocks: Array<{
				type: string;
				runs?: Array<{ type: string; kind?: string }>;
			}>;
		};
		const charts = doc.blocks
			.flatMap((b) => b.runs ?? [])
			.filter((r) => r.type === "chart");
		expect(charts.map((c) => c.kind).sort()).toEqual([
			"chart",
			"drawing",
			"smartart",
		]);
	});
});

describe("docx read --ast (JSON AST opt-in)", () => {
	test("--ast emits valid JSON AST", async () => {
		const result = await runCli("read", fixture("minimal.docx"), "--ast");
		expect(result.exitCode).toBe(0);
		const doc = result.parsed as { schemaVersion: number; blocks: unknown[] };
		expect(doc.schemaVersion).toBe(1);
		expect(Array.isArray(doc.blocks)).toBe(true);
	});

	test("read FILE (no flags) emits Markdown, not JSON", async () => {
		const result = await runCli("read", fixture("minimal.docx"));
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toBeUndefined();
		expect(result.stdout).toContain("# Style Guide <!-- p0 -->");
	});

	test("read FILE --help prints usage", async () => {
		const result = await runCli("read", "--help");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("docx read");
		expect(result.stdout).toContain("--ast");
		expect(result.stdout).toContain("--from");
		expect(result.stdout).toContain("--comments");
	});
});

describe("docx read (markdown) smoke-test all fixtures", () => {
	const FIXTURES_LIST = [
		"academic-paper.docx",
		"comments-simple.docx",
		"comments-with-replies.docx",
		"equations.docx",
		"large-mixed.docx",
		"minimal.docx",
		"multi-column.docx",
		"notes.docx",
		"resume-styling.docx",
		"strict-profile.docx",
		"tables-and-lists.docx",
		"tracked-changes.docx",
	];

	for (const name of FIXTURES_LIST) {
		test(`${name}: default markdown render completes successfully`, async () => {
			const result = await runCli("read", fixture(name));
			expect(result.exitCode).toBe(0);
			expect(result.stdout.length).toBeGreaterThan(0);
		});

		test(`${name}: --accepted --comments completes`, async () => {
			const result = await runCli(
				"read",
				fixture(name),
				"--accepted",
				"--comments",
			);
			expect(result.exitCode).toBe(0);
			expect(result.stdout.length).toBeGreaterThan(0);
		});
	}
});

type OutlineEntry = {
	id: string;
	locator: string;
	level: number;
	style: string;
	text: string;
	children: OutlineEntry[];
};

describe("docx outline", () => {
	test("builds a hierarchy from the academic-paper fixture", async () => {
		const result = await runCli(
			"outline",
			"tests/fixtures/academic-paper.docx",
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as OutlineEntry[];
		const topLevelTexts = payload.map((entry) => entry.text);
		expect(topLevelTexts).toContain("Guided Imagery");
		expect(topLevelTexts).toContain("Conclusion");
		expect(topLevelTexts).toContain("References");

		const guidedImagery = payload.find(
			(entry) => entry.text === "Guided Imagery",
		);
		expect(guidedImagery?.level).toBe(1);
		expect(guidedImagery?.children.map((child) => child.text)).toEqual([
			"Features of Guided Imagery",
			"Guided Imagery in Group Psychotherapy",
		]);
		expect(guidedImagery?.children[0]?.level).toBe(2);
		expect(guidedImagery?.children[0]?.locator).toMatch(/^p\d+$/);
	});

	test("doc with no headings returns an empty outline", async () => {
		const workspace = tempWorkspace("outline-empty");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Just a body paragraph.");
		const result = await runCli("outline", docPath);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toEqual([]);
	});

	test("--style-prefix targets a non-default style family", async () => {
		// academic-paper has a "Title" style on p3 — invisible under default Heading
		// prefix but should surface when we ask for it explicitly.
		const result = await runCli(
			"outline",
			"tests/fixtures/academic-paper.docx",
			"--style-prefix",
			"Title",
		);
		const payload = result.parsed as OutlineEntry[];
		expect(payload).toHaveLength(1);
		expect(payload[0]?.style).toBe("Title");
		expect(payload[0]?.level).toBe(1);
	});

	test("skipped levels nest directly under the nearest shallower level", async () => {
		const workspace = tempWorkspace("outline-skip");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "intro");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Top",
			"--style",
			"Heading1",
		);
		await runCli(
			"insert",
			docPath,
			"--after",
			"p1",
			"--text",
			"Skipped",
			"--style",
			"Heading3",
		);

		const result = await runCli("outline", docPath);
		const payload = result.parsed as OutlineEntry[];
		expect(payload).toHaveLength(1);
		expect(payload[0]?.text).toBe("Top");
		expect(payload[0]?.children).toHaveLength(1);
		expect(payload[0]?.children[0]).toMatchObject({
			text: "Skipped",
			level: 3,
		});
	});

	test("a Heading-styled paragraph inside a table cell is skipped", async () => {
		const workspace = tempWorkspace("outline-cell");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "intro");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Top Heading",
			"--style",
			"Heading1",
		);
		await runCli(
			"insert",
			docPath,
			"--after",
			"p1",
			"--table",
			"--rows",
			"1",
			"--cols",
			"1",
		);
		// Style the in-cell paragraph as a Heading2 — outline only walks
		// top-level blocks, so it must NOT surface.
		await runCli(
			"edit",
			docPath,
			"--at",
			"t0:r0c0:p0",
			"--markdown",
			"## In-Cell Heading",
		);

		const result = await runCli("outline", docPath);
		const payload = result.parsed as OutlineEntry[];
		expect(payload).toHaveLength(1);
		expect(payload[0]?.text).toBe("Top Heading");
		expect(JSON.stringify(payload)).not.toContain("In-Cell Heading");
	});
});

// Paragraph style + alignment that GFM can't show ride a docx:p hint (per the
// naming rule: bare = locator, docx: = metadata). Deviation-only: styles already
// conveyed by the construct (#, -, >, fences) and default-left alignment emit
// nothing. Read-time hints; the importer drops them.
describe("docx read — paragraph style/alignment hints (docx:p)", () => {
	async function styledDoc(): Promise<string> {
		const workspace = tempWorkspace("read-docx-p");
		const md = join(workspace, "s.md");
		await Bun.write(md, "# Heading\n\nnormal body\n");
		const doc = join(workspace, "d.docx");
		await runCli("create", doc, "--from", md);
		await runCli(
			"insert",
			doc,
			"--after",
			"p1",
			"--text",
			"A caption",
			"--style",
			"Caption",
		);
		await runCli(
			"insert",
			doc,
			"--after",
			"p1",
			"--text",
			"Centered",
			"--alignment",
			"center",
		);
		await runCli(
			"edit",
			doc,
			"--at",
			"p0",
			"--text",
			"Heading",
			"--alignment",
			"center",
		);
		return doc;
	}

	test("a non-construct style (Caption) surfaces as docx:p style", async () => {
		const out = await render(await styledDoc());
		expect(out).toMatch(/<!-- docx:p p3 style="Caption" -->/);
	});

	test("non-default alignment surfaces in a docx:p note that carries the locator (no bare dup)", async () => {
		const out = await render(await styledDoc());
		// The docx:p note carries `p2` as its leading token (like docx:cell), so the
		// bare `<!-- p2 -->` locator is NOT also emitted — that would duplicate p2.
		expect(out).toContain('Centered <!-- docx:p p2 align="center" -->');
		expect(out).not.toContain("<!-- p2 -->");
	});

	test("a heading construct gets no style= (but align= still shows, locator in the note)", async () => {
		const out = await render(await styledDoc());
		expect(out).toMatch(/# Heading <!-- docx:p p0 align="center" -->/);
		expect(out).not.toMatch(/docx:p p0[^>]*style=/);
		expect(out).not.toContain("<!-- p0 -->");
	});

	test("a plain Normal/left paragraph gets no docx:p note", async () => {
		const out = await render(await styledDoc());
		expect(out).toMatch(/normal body <!-- p1 -->(?! <!-- docx:p)/);
	});
});

describe("docx read — paragraph spacing & indentation (letter.docx fixture)", () => {
	// letter.docx is authored entirely via the insert/edit spacing+indent flags
	// (tests/fixtures/setup/letter.ts), so reading it back is the round-trip dogfood.
	type LetterPara = {
		id: string;
		type: string;
		spacing?: { before?: number; after?: number; line?: number };
		indent?: {
			left?: number;
			right?: number;
			firstLine?: number;
			hanging?: number;
		};
		runs?: Array<{ text?: string }>;
	};
	async function paras(): Promise<LetterPara[]> {
		const result = await runCli("read", fixture("letter.docx"), "--ast");
		expect(result.exitCode).toBe(0);
		return (result.parsed as { blocks: LetterPara[] }).blocks.filter(
			(b) => b.type === "paragraph",
		);
	}
	const byText = (list: LetterPara[], needle: string): LetterPara | undefined =>
		list.find((p) =>
			(p.runs ?? []).some((r) => (r.text ?? "").includes(needle)),
		);

	test("--ast carries every authored pPr (twips): tight blocks, body, quote, signature, enclosures", async () => {
		const list = await paras();
		// Tight address line: space-after 0.
		expect(byText(list, "Howard Street")?.spacing?.after).toBe(0);
		// Body: first-line indent (0.5in) + 1.5 line spacing (range-edited).
		const body = byText(list, "Thank you for sending");
		expect(body?.indent?.firstLine).toBe(720);
		expect(body?.spacing?.line).toBe(360);
		// Block quote: left + right indent.
		const quote = byText(list, "aggregate liability");
		expect(quote?.indent).toEqual({ left: 720, right: 720 });
		// Closing / signature: indented to ~center (3.5in = 5040 twips).
		expect(byText(list, "Sincerely")?.indent?.left).toBe(5040);
		// Enclosures: hanging indent.
		expect(byText(list, "Enclosures:")?.indent?.hanging).toBe(720);
	});

	test("read markdown annotates the layout deviation-only, in re-appliable units", async () => {
		const md = await render(fixture("letter.docx"));
		expect(md).toContain('first-line="0.5in"');
		expect(md).toContain('line-spacing="1.5"');
		expect(md).toContain('indent-left="3.5in"');
		expect(md).toContain('hanging="0.5in"');
		expect(md).toContain('space-after="0pt"');
	});
});

describe("docx read — direct indent on a Quote paragraph (write-read loop)", () => {
	// The Quote style's LEFT indent is structural (it drives quoteDepth), so it's
	// not surfaced as direct indent — but right/firstLine/hanging are direct
	// formatting and must round-trip, not vanish on read-back.
	type QuotePara = {
		id: string;
		quoteDepth?: number;
		indent?: { left?: number; right?: number; hanging?: number };
	};
	test("right/hanging surface; left stays structural (quoteDepth)", async () => {
		const docPath = join(tempWorkspace("quote-indent"), "doc.docx");
		await runCli("create", docPath, "--text", "A quote line.");
		await runCli("edit", docPath, "--at", "p0", "--style", "Quote");
		await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--indent-right",
			"0.5",
			"--hanging",
			"0.25",
		);
		const result = await runCli("read", docPath, "--ast");
		const p0 = (result.parsed as { blocks: QuotePara[] }).blocks.find(
			(b) => b.id === "p0",
		);
		expect(p0?.quoteDepth).toBe(1);
		expect(p0?.indent).toEqual({ right: 720, hanging: 360 });
		expect(p0?.indent?.left).toBeUndefined();
	});
});
