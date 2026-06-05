import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli } from "./harness";

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
		expect(out).toContain("--- <!-- s0 -->");
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
		// Ordered items emit "1. " regardless of depth (GFM auto-increments)
		expect(out).toMatch(/^1\. Preheat the oven/m);
		expect(out).toMatch(/^ {2}1\. Sift the flour/m);
		expect(out).toMatch(/^ {4}1\. Don't overmix/m);
		// The second bullet list (independent numId) still renders as bullets
		expect(out).toMatch(/^- Independent item one/m);
	});

	test("tables-and-lists: top-level ordered items render as `1.`", async () => {
		// These are Word-numbered items (numFmt=decimal), so they should
		// render as ordered list items rather than bullets. The renderer
		// uses `1. ` for every item since GFM auto-increments client-side.
		const out = await render(fixture("tables-and-lists.docx"));
		expect(out).toMatch(/^1\. \*\*Introduction\*\*/m);
		expect(out).toMatch(/^1\. \*\*Background Information\*\*/m);
		expect(out).toMatch(/^1\. \*\*Methods and Materials\*\*/m);
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
			ok: false,
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
			ok: false,
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
		expect(result.parsed).toMatchObject({ ok: false, code: "USAGE" });
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
		expect(result.parsed).toMatchObject({ ok: false, code: "USAGE" });
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

	test("--comments works on rich comments fixture", async () => {
		const out = await render(fixture("comments-rich.docx"), "--comments");
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
	test('non-default run color wraps in <span style="color:#hex">', async () => {
		const out = await render(fixture("minimal.docx"));
		expect(out).toContain('<span style="color:#800080">');
	});

	test("color span sits inside formatting markers (bold etc.)", async () => {
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
		"comments-rich.docx",
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
