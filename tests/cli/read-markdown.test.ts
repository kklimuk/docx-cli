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

describe("docx read --markdown", () => {
	test("minimal: emits heading, bold, locator comments", async () => {
		const out = await render(fixture("minimal.docx"), "--markdown");
		expect(out).toContain("# Style Guide <!-- p0 -->");
		expect(out).toContain("**important**");
		expect(out).toMatch(/Use[^\n]*\*\*important\*\*[^\n]*terms/);
		expect(out).toContain("<!-- p2 -->");
		expect(out).toContain("--- <!-- s0 -->");
	});

	test("academic-paper: heading levels 1 and 2 render with # and ##", async () => {
		const out = await render(fixture("academic-paper.docx"), "--markdown");
		expect(out).toMatch(/^# Guided Imagery and Progressive Muscle Relaxation/m);
		expect(out).toMatch(/^## Features of Guided Imagery/m);
		expect(out).toMatch(/^## Guided Imagery in Group Psychotherapy/m);
	});

	test("academic-paper: italics and hyperlinks render inline", async () => {
		const out = await render(fixture("academic-paper.docx"), "--markdown");
		expect(out).toContain("*Guided imagery*");
		expect(out).toContain("*Progressive muscle relaxation*");
	});

	test("tables-and-lists: pipe table with cell locators", async () => {
		const out = await render(fixture("tables-and-lists.docx"), "--markdown");
		expect(out).toMatch(/^\| \*\*Equipment\*\* <!-- t0:r0c0:p0 --> \|/m);
		expect(out).toMatch(/^\| --- \| --- \|$/m);
		expect(out).toContain("Agilent E3631A Triple Output DC Power Supply");
		expect(out).toContain("9.1 Ω Resistor");
	});

	test("tables-and-lists: top-level list bullets render", async () => {
		const out = await render(fixture("tables-and-lists.docx"), "--markdown");
		expect(out).toMatch(/^- \*\*Introduction\*\*/m);
		expect(out).toMatch(/^- \*\*Background Information\*\*/m);
		expect(out).toMatch(/^- \*\*Methods and Materials\*\*/m);
	});

	test("resume-styling: list bullets and bold runs render", async () => {
		const out = await render(fixture("resume-styling.docx"), "--markdown");
		expect(out).toMatch(/^- GPA: 3\.5/m);
		expect(out).toContain("**JANE SMITH**");
		expect(out).toContain("[j1smith@business.rutgers.edu](mailto:");
	});

	test("multi-column: SIGCHI three-column table renders as pipe table", async () => {
		const out = await render(fixture("multi-column.docx"), "--markdown");
		expect(out).toMatch(/^\| Leave Authors Anonymous <!-- t0:r0c0:p0 -->/m);
		expect(out).toMatch(/^\| --- \| --- \| --- \|$/m);
		expect(out).toContain("<br>for Submission <!-- t0:r0c0:p1 -->");
	});

	test("large-mixed: images render as ![alt](imgN) and locator survives", async () => {
		const out = await render(fixture("large-mixed.docx"), "--markdown");
		expect(out).toMatch(/!\[[^\]]*\]\(img0\)/);
		expect(out).toMatch(/!\[[^\]]*\]\(img1\)/);
		expect(out).toContain("# Chinese Folding Fan Design Project");
	});

	test("equations: hyperlinks survive; OOMath surfaces as `equation: ...` placeholder", async () => {
		const out = await render(fixture("equations.docx"), "--markdown");
		expect(out).toContain("[basis]");
		expect(out).toContain("[change of basis]");
		expect(out).toContain("Einstein summation convention");
		expect(out).toContain("`equation: ei`");
		expect(out).toContain("`equation: vi=R−1jivj,`");
	});

	test("strict-profile: chart/smartart/drawing placeholders render", async () => {
		const out = await render(fixture("strict-profile.docx"), "--markdown");
		expect(out).toContain("`[chart]`");
		expect(out).toContain("`[smartart]`");
		expect(out).toContain("`[drawing]`");
	});

	test("locator pins use HTML comments, not <sup>", async () => {
		const out = await render(fixture("minimal.docx"), "--markdown");
		expect(out).not.toContain("<sup>");
		expect(out).toMatch(/<!-- p\d+ -->/);
	});
});

describe("docx read --markdown --from / --to", () => {
	test("--from p1 starts at p1, drops p0", async () => {
		const out = await render(
			fixture("minimal.docx"),
			"--markdown",
			"--from",
			"p1",
		);
		expect(out).not.toContain("# Style Guide");
		expect(out).toContain("**important**");
		expect(out).toMatch(/Use[^\n]*\*\*important\*\*[^\n]*terms/);
	});

	test("--to is inclusive", async () => {
		const out = await render(
			fixture("minimal.docx"),
			"--markdown",
			"--to",
			"p1",
		);
		expect(out).toContain("# Style Guide");
		expect(out).toContain("**important**");
		expect(out).toMatch(/Use[^\n]*\*\*important\*\*[^\n]*terms/);
		expect(out).not.toContain("The quick brown fox");
	});

	test("--from p13 --to p15 slices the middle of academic-paper", async () => {
		const out = await render(
			fixture("academic-paper.docx"),
			"--markdown",
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
		const out = await render(
			fixture("minimal.docx"),
			"--markdown",
			"--from",
			"p1:0-5",
		);
		expect(out).toContain("**important**");
		expect(out).toMatch(/Use[^\n]*\*\*important\*\*[^\n]*terms/);
		expect(out).not.toContain("# Style Guide");
	});

	test("--from accepts a range locator (uses start paragraph)", async () => {
		const out = await render(
			fixture("minimal.docx"),
			"--markdown",
			"--from",
			"p1:0-p2:3",
		);
		expect(out).toContain("**important**");
		expect(out).toContain("The quick brown fox");
		expect(out).not.toContain("# Style Guide");
	});

	test("--from cN rejected as invalid locator", async () => {
		const result = await runCli(
			"read",
			fixture("minimal.docx"),
			"--markdown",
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
			"--markdown",
			"--from",
			"p99",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({
			ok: false,
			code: "INVALID_LOCATOR",
		});
	});

	test("--from without --markdown fails with USAGE", async () => {
		const result = await runCli(
			"read",
			fixture("minimal.docx"),
			"--from",
			"p1",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ ok: false, code: "USAGE" });
	});
});

describe("docx read --markdown --changes", () => {
	test("default view drops nothing (no del runs in fixture) but skips ins markup", async () => {
		const out = await render(fixture("tracked-changes.docx"), "--markdown");
		expect(out).toContain("two exciting");
		expect(out).not.toContain("<ins>");
		expect(out).not.toContain("<del>");
	});

	test("--changes wraps inserted text in <ins>", async () => {
		const out = await render(
			fixture("tracked-changes.docx"),
			"--markdown",
			"--changes",
		);
		expect(out).toContain("<ins>two exciting </ins>insertions");
	});
});

describe("docx read --markdown --comments", () => {
	test("default view (no --comments) emits no [^cN] refs and no footnotes", async () => {
		const out = await render(
			fixture("comments-with-replies.docx"),
			"--markdown",
		);
		expect(out).not.toMatch(/\[\^c\d+\]/);
		expect(out).not.toContain("[^c0]:");
	});

	test("--comments inlines [^cN] at span end and emits GFM footnotes", async () => {
		const out = await render(
			fixture("comments-with-replies.docx"),
			"--markdown",
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
			"--markdown",
			"--comments",
		);
		expect(out).toMatch(/comment in a comment\[\^c3\]\[\^c4\]/);
	});

	test("--comments reply footnote shows ↳ parent", async () => {
		const out = await render(
			fixture("comments-with-replies.docx"),
			"--markdown",
			"--comments",
		);
		expect(out).toContain("↳ c3:");
	});

	test("--comments on doc without comments emits no footnote section", async () => {
		const out = await render(
			fixture("strict-profile.docx"),
			"--markdown",
			"--comments",
		);
		expect(out).not.toContain("[^c");
	});

	test("--comments + --from drops footnotes whose spans are out of range", async () => {
		const out = await render(
			fixture("comments-with-replies.docx"),
			"--markdown",
			"--comments",
			"--from",
			"p2",
		);
		expect(out).not.toContain("[^c0]:");
		expect(out).not.toContain("[^c1]:");
		expect(out).toContain("[^c2]:");
	});

	test("--comments works on simple comments fixture", async () => {
		const out = await render(
			fixture("comments-simple.docx"),
			"--markdown",
			"--comments",
		);
		expect(out).toContain("[^c0]:");
	});

	test("--comments works on rich comments fixture", async () => {
		const out = await render(
			fixture("comments-rich.docx"),
			"--markdown",
			"--comments",
		);
		expect(out).toContain("[^c0]:");
	});
});

describe("docx read --markdown equations", () => {
	test("inline `<m:oMath>` renders as `equation: ...` mid-paragraph", async () => {
		const out = await render(fixture("equations.docx"), "--markdown");
		expect(out).toContain("`equation: ei`");
	});

	test("display `<m:oMathPara>` renders on its own line", async () => {
		const out = await render(fixture("equations.docx"), "--markdown");
		expect(out).toMatch(/^`equation: ei=j=1nejRij=ejRij\.` <!-- p1 -->/m);
	});

	test("equations marked display=true survive in JSON AST", async () => {
		const result = await runCli("read", fixture("equations.docx"));
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

describe("docx read --markdown footnotes / endnotes", () => {
	test("notes.docx: footnote ref + definition both render", async () => {
		const out = await render(fixture("notes.docx"), "--markdown");
		expect(out).toContain("[^fn1]");
		expect(out).toMatch(/\n\[\^fn1\]: My note\./);
	});

	test("notes.docx: endnote ref + definition both render", async () => {
		const out = await render(fixture("notes.docx"), "--markdown");
		expect(out).toContain("[^en1]");
		expect(out).toContain(
			"[^en1]: This is an endnote at the end of the document.",
		);
	});

	test("notes.docx: footnote and endnote refs both visible inline", async () => {
		const out = await render(fixture("notes.docx"), "--markdown");
		expect(out).toMatch(/Test footnote\.\[\^fn1\] Test endnote\.\[\^en1\]/);
	});

	test("equations.docx: substantive footnote with math symbols renders", async () => {
		const out = await render(fixture("equations.docx"), "--markdown");
		expect(out).toContain("[^fn26]");
		expect(out).toMatch(/\n\[\^fn26\]: The Einstein summation convention/);
	});

	test("footnote definitions only emitted for refs visible in slice", async () => {
		const out = await render(
			fixture("equations.docx"),
			"--markdown",
			"--from",
			"p3",
		);
		expect(out).not.toContain("[^fn26]");
	});

	test("doc without footnote refs in body emits no definitions", async () => {
		const out = await render(fixture("minimal.docx"), "--markdown");
		expect(out).not.toContain("[^fn");
		expect(out).not.toContain("[^en");
	});

	test("footnotes/endnotes arrays present in JSON AST", async () => {
		const result = await runCli("read", fixture("notes.docx"));
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
		const result = await runCli("read", fixture("notes.docx"));
		const doc = result.parsed as {
			footnotes: Array<{ id: string }>;
		};
		const hasNegativeId = doc.footnotes.some(
			(f) => f.id.startsWith("fn-") || f.id === "fn0",
		);
		expect(hasNegativeId).toBe(false);
	});
});

describe("docx read --markdown color and highlight", () => {
	test('non-default run color wraps in <span style="color:#hex">', async () => {
		const out = await render(fixture("minimal.docx"), "--markdown");
		expect(out).toContain('<span style="color:#800080">');
	});

	test("color span sits inside formatting markers (bold etc.)", async () => {
		const out = await render(fixture("minimal.docx"), "--markdown");
		// "important" is bold + purple in minimal.docx
		expect(out).toContain('<span style="color:#800080">**important**</span>');
	});
});

describe("docx read --markdown chart / drawing placeholders", () => {
	test("chart drawing renders as `[chart]`", async () => {
		const out = await render(fixture("strict-profile.docx"), "--markdown");
		expect(out).toContain("`[chart]`");
	});

	test("smartart and generic drawing placeholders render too", async () => {
		const out = await render(fixture("strict-profile.docx"), "--markdown");
		expect(out).toContain("`[smartart]`");
		expect(out).toContain("`[drawing]`");
	});

	test("chart placeholders survive in JSON AST", async () => {
		const result = await runCli("read", fixture("strict-profile.docx"));
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

describe("docx read JSON mode (default) still works", () => {
	test("read FILE emits valid JSON AST", async () => {
		const result = await runCli("read", fixture("minimal.docx"));
		expect(result.exitCode).toBe(0);
		const doc = result.parsed as { schemaVersion: number; blocks: unknown[] };
		expect(doc.schemaVersion).toBe(1);
		expect(Array.isArray(doc.blocks)).toBe(true);
	});

	test("read FILE --help prints usage", async () => {
		const result = await runCli("read", "--help");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("docx read");
		expect(result.stdout).toContain("--markdown");
		expect(result.stdout).toContain("--from");
		expect(result.stdout).toContain("--comments");
	});
});

describe("docx read --markdown smoke-test all fixtures", () => {
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
		test(`${name}: --markdown completes successfully`, async () => {
			const result = await runCli("read", fixture(name), "--markdown");
			expect(result.exitCode).toBe(0);
			expect(result.stdout.length).toBeGreaterThan(0);
		});

		test(`${name}: --markdown --changes --comments completes`, async () => {
			const result = await runCli(
				"read",
				fixture(name),
				"--markdown",
				"--changes",
				"--comments",
			);
			expect(result.exitCode).toBe(0);
			expect(result.stdout.length).toBeGreaterThan(0);
		});
	}
});
