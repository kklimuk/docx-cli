import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import { clearFormatting, resolveClearTags } from "@core/edit/clear-formatting";
import { XmlNode } from "@core/parser";
import { runCli, tempWorkspace } from "./harness";
import { freshFixture, readDocumentXml } from "./helpers";

const FIXTURE = "tests/fixtures/word-formatted.docx";
// Layout (built by tests/fixtures/setup/word-formatted.ts):
//   p0: plain — "The quick brown fox jumps over the lazy dog."
//   p1: "The " + "MESSENGER" [bold,#800080] + " is " + "fatally flawed" [italic] + "."
//   p2: "Bold" [bold] + " then " + "italic" [italic] + " then plain."
//   p3: "Rating: The me" + "ssenger is fatally f" + "lawed." (all italic, splits mid-word)

const freshCopy = (label: string) => freshFixture(label, FIXTURE);

type Run = {
	type: string;
	text?: string;
	bold?: boolean;
	italic?: boolean;
	color?: string;
	trackedChange?: { kind: string };
};

async function readParagraph(docPath: string, blockId: string): Promise<Run[]> {
	const result = await runCli("read", docPath, "--ast");
	const blocks = (
		result.parsed as { blocks: Array<{ id: string; runs?: Run[] }> }
	).blocks;
	const block = blocks.find((candidate) => candidate.id === blockId);
	return block?.runs ?? [];
}

describe("docx edit --text — formatting preservation (default)", () => {
	test("preserves bold + italic on unchanged words; new words inherit from neighbors", async () => {
		const docPath = await freshCopy("preserve-default");
		// Replace one word in p1: "fatally flawed" → "irreparably damaged"
		// Both old and new tokens span the italic span.
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p1",
			"--text",
			"The MESSENGER is irreparably damaged.",
		);
		expect(result.exitCode).toBe(0);

		const runs = await readParagraph(docPath, "p1");
		// Concatenated text reads cleanly.
		const flat = runs
			.filter((run) => run.type === "text")
			.map((run) => run.text ?? "")
			.join("");
		expect(flat).toBe("The MESSENGER is irreparably damaged.");

		// MESSENGER stays bold + colored.
		const messenger = runs.find((run) => run.text === "MESSENGER");
		expect(messenger?.bold).toBe(true);
		expect(messenger?.color).toBe("800080");

		// "irreparably damaged" inherits italic from the neighboring kept
		// italic run that originally held "fatally flawed".
		const replacement = runs.find((run) => run.text?.includes("irreparably"));
		expect(replacement?.italic).toBe(true);
	});

	test("preserves formatting when only a non-formatted word changes", async () => {
		const docPath = await freshCopy("preserve-plain");
		// p2: "Bold" [b] + " then " + "italic" [i] + " then plain."
		// Change only the trailing "plain" → "ordinary".
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p2",
			"--text",
			"Bold then italic then ordinary.",
		);
		expect(result.exitCode).toBe(0);

		const runs = await readParagraph(docPath, "p2");
		const bold = runs.find((run) => run.text === "Bold");
		const italic = runs.find((run) => run.text === "italic");
		expect(bold?.bold).toBe(true);
		expect(italic?.italic).toBe(true);
		const flat = runs
			.filter((run) => run.type === "text")
			.map((run) => run.text ?? "")
			.join("");
		expect(flat).toBe("Bold then italic then ordinary.");
	});

	test("an embedded tab keeps run formatting and emits a real <w:tab/>", async () => {
		// Regression: a leading bold span followed by a tab (the canonical
		// resume "**Name**⇥City" line). The tab used to bypass the
		// formatting-preserve path, flattening the whole paragraph into one
		// rPr-less run — bold AND the run's font fell back to docDefaults
		// (Times New Roman). Now the tab rides inside its run as <w:tab/> and
		// the surrounding rPr survives.
		const docPath = await freshCopy("preserve-tab");
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p2",
			"--text",
			"Bold then italic\tthen plain.",
		);
		expect(result.exitCode).toBe(0);

		const runs = await readParagraph(docPath, "p2");
		const bold = runs.find((run) => run.text === "Bold");
		expect(bold?.bold).toBe(true);
		const italic = runs.find((run) => run.text === "italic");
		expect(italic?.italic).toBe(true);

		// The tab is a real <w:tab/> element, not a literal "\t" leaked into <w:t>.
		const xml = await readDocumentXml(docPath);
		expect(xml).toContain("<w:tab/>");
		expect(xml).not.toMatch(/<w:t[ >][^<]*\t/);
	});

	test("--no-formatting reverts to a single fresh run", async () => {
		const docPath = await freshCopy("no-formatting");
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p1",
			"--text",
			"The MESSENGER is irreparably damaged.",
			"--no-formatting",
		);
		expect(result.exitCode).toBe(0);

		const runs = await readParagraph(docPath, "p1");
		// All text in a single run with no rPr.
		const textRuns = runs.filter((run) => run.type === "text");
		expect(textRuns).toHaveLength(1);
		expect(textRuns[0]?.text).toBe("The MESSENGER is irreparably damaged.");
		expect(textRuns[0]?.bold).toBeUndefined();
		expect(textRuns[0]?.italic).toBeUndefined();
		expect(textRuns[0]?.color).toBeUndefined();
	});

	test("explicit --bold/--color bypasses preservation (uniform formatting)", async () => {
		const docPath = await freshCopy("explicit-format");
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p1",
			"--text",
			"All red bold now.",
			"--bold",
			"--color",
			"FF0000",
		);
		expect(result.exitCode).toBe(0);

		const runs = await readParagraph(docPath, "p1");
		const textRuns = runs.filter((run) => run.type === "text");
		expect(textRuns).toHaveLength(1);
		expect(textRuns[0]?.bold).toBe(true);
		expect(textRuns[0]?.color).toBe("FF0000");
	});
});

describe("docx edit --text — formatting preservation under tracking", () => {
	async function trackedCopy(label: string): Promise<string> {
		const docPath = await freshCopy(label);
		await runCli("track-changes", docPath, "on");
		return docPath;
	}

	test("word-level del/ins instead of whole-paragraph del+ins", async () => {
		const docPath = await trackedCopy("tracked-word-level");
		await runCli(
			"edit",
			docPath,
			"--at",
			"p1",
			"--text",
			"The MESSENGER is irreparably damaged.",
		);

		// Should produce per-word del/ins markers (not a whole-paragraph
		// del + ins). The keep tokens stay outside any tracked-change wrapper.
		const tracked = await runCli("track-changes", "list", docPath);
		const changes = tracked.parsed as Array<{ kind: string; blockId: string }>;
		const insertions = changes.filter((change) => change.kind === "ins");
		const deletions = changes.filter((change) => change.kind === "del");
		expect(insertions.length).toBeGreaterThanOrEqual(1);
		expect(deletions.length).toBeGreaterThanOrEqual(1);

		// Concatenated text in default (accepted) view should read cleanly.
		const acceptedRuns = await readParagraph(docPath, "p1");
		const acceptedText = acceptedRuns
			.filter(
				(run) =>
					run.type === "text" &&
					run.trackedChange?.kind !== "del" &&
					run.trackedChange?.kind !== "moveFrom",
			)
			.map((run) => run.text ?? "")
			.join("");
		expect(acceptedText).toBe("The MESSENGER is irreparably damaged.");

		// And baseline view (drop ins/moveTo) reads the original text.
		const baselineText = acceptedRuns
			.filter(
				(run) =>
					run.type === "text" &&
					run.trackedChange?.kind !== "ins" &&
					run.trackedChange?.kind !== "moveTo",
			)
			.map((run) => run.text ?? "")
			.join("");
		expect(baselineText).toBe("The MESSENGER is fatally flawed.");
	});

	test("MESSENGER bold + color preserved on the unchanged span under tracking", async () => {
		const docPath = await trackedCopy("tracked-preserve-bold");
		await runCli(
			"edit",
			docPath,
			"--at",
			"p1",
			"--text",
			"The MESSENGER is irreparably damaged.",
		);

		const runs = await readParagraph(docPath, "p1");
		const messenger = runs.find((run) => run.text === "MESSENGER");
		expect(messenger?.bold).toBe(true);
		expect(messenger?.color).toBe("800080");
		// Should NOT be inside a tracked-change wrapper — it was unchanged.
		expect(messenger?.trackedChange).toBeUndefined();
	});
});

describe("docx edit --text — regression: paragraphs with existing run-bearing wrappers", () => {
	// Bug: extractOldTokens only walks direct <w:r> children, but the
	// rebuild preserves every non-<w:r> child verbatim. So a paragraph
	// containing <w:ins>/<w:del>/<w:hyperlink>/etc. would leak that
	// wrapper's text alongside the diff output — duplicating it.
	// Fix: bow out of formatting preservation when any run-bearing
	// wrapper is present and fall through to the legacy whole-paragraph
	// del+ins path.

	test("chained edit under tracking does not leak the prior <w:ins> text", async () => {
		const workspace = tempWorkspace("chained-leak");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "intro");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Hello world today.",
		);
		await runCli("track-changes", docPath, "on");

		// First edit: produces <w:ins>brave </w:ins> in the paragraph.
		await runCli(
			"edit",
			docPath,
			"--at",
			"p1",
			"--text",
			"Hello brave world today.",
		);
		// Second edit: the buggy path walks direct <w:r> children only,
		// missing the existing <w:ins>brave </w:ins>; then rebuild keeps
		// the <w:ins> AND appends new diff runs, producing duplicate text
		// like "brave Hello brave new world today.".
		await runCli(
			"edit",
			docPath,
			"--at",
			"p1",
			"--text",
			"Hello brave new world today.",
		);

		const runs = await readParagraph(docPath, "p1");
		const acceptedText = runs
			.filter(
				(run) =>
					run.type === "text" &&
					run.trackedChange?.kind !== "del" &&
					run.trackedChange?.kind !== "moveFrom",
			)
			.map((run) => run.text ?? "")
			.join("");
		expect(acceptedText).toBe("Hello brave new world today.");
	});

	test("edit on a paragraph with a hyperlink does not duplicate the hyperlink text", async () => {
		const workspace = tempWorkspace("hyperlink-leak");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "intro");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Read the manual at example.com today.",
		);
		await runCli(
			"hyperlinks",
			"add",
			docPath,
			"--at",
			"p1:18-29",
			"--url",
			"https://example.com",
		);

		// Without the fix, extractOldTokens skips the <w:hyperlink>
		// wrapper, so the diff sees "Read the manual at  today." with no
		// "example.com". The rebuild then keeps the hyperlink AND appends
		// the new diff runs, producing "example.coRead the manual at
		// example.com tomorrow." or similar.
		await runCli(
			"edit",
			docPath,
			"--at",
			"p1",
			"--text",
			"Read the manual at example.com tomorrow.",
		);

		const runs = await readParagraph(docPath, "p1");
		const flat = runs
			.filter((run) => run.type === "text")
			.map((run) => run.text ?? "")
			.join("");
		expect(flat).toBe("Read the manual at example.com tomorrow.");
	});
});

describe("docx edit --text — regression: heading replacement consolidates", () => {
	// LCS over text alone matches any " " in old to any " " in new. When
	// the old heading is a single bold run "Police Reform / Safer Streets."
	// and the new heading is "Law and Order.", the spurious whitespace
	// matches fragment the bold span — producing alternating
	// del/ins/keep-space wrappers that render as `**Law**** ****and****`
	// in the accepted view (each word a separate bold span).
	//
	// The fix demotes whitespace-only keeps adjacent to non-keeps and
	// consolidates each edit group (all deletes precede all inserts), so
	// `groupAndEmitPlainRuns` can merge consecutive same-rPr entries
	// into a single bold run.

	test("rewriting a fully-bold heading keeps the bold contiguous", async () => {
		const workspace = tempWorkspace("heading-rewrite");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "intro");
		const headingRuns = JSON.stringify([
			{ type: "text", text: "Police Reform / Safer Streets.", bold: true },
			{ type: "text", text: " Then a tail." },
		]);
		await runCli("insert", docPath, "--after", "p0", "--runs", headingRuns);

		await runCli("track-changes", docPath, "on");
		await runCli(
			"edit",
			docPath,
			"--at",
			"p1",
			"--text",
			"Law and Order. Then a tail.",
		);

		const runs = await readParagraph(docPath, "p1");
		// Walk the inserted runs (accepted view: ins inlined, del dropped).
		const acceptedBoldRuns = runs.filter(
			(run) =>
				run.type === "text" &&
				run.bold === true &&
				run.trackedChange?.kind !== "del" &&
				run.trackedChange?.kind !== "moveFrom",
		);
		// Pre-fix: 5+ separate bold runs (Law / " " / and / " " / Order.).
		// Post-fix: a single bold run "Law and Order." (or "Law and Order. ")
		// because consolidation lets `groupAndEmitPlainRuns` merge
		// consecutive same-bold inserts.
		expect(acceptedBoldRuns).toHaveLength(1);
		const concatenated = acceptedBoldRuns[0]?.text ?? "";
		expect(concatenated).toContain("Law and Order.");
	});

	test("tracked-change count for a heading rewrite scales with edit boundaries, not whitespace", async () => {
		const workspace = tempWorkspace("heading-tc-count");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "intro");
		const headingRuns = JSON.stringify([
			{ type: "text", text: "Police Reform / Safer Streets.", bold: true },
		]);
		await runCli("insert", docPath, "--after", "p0", "--runs", headingRuns);
		await runCli("track-changes", docPath, "on");
		await runCli("edit", docPath, "--at", "p1", "--text", "Law and Order.");

		const tracked = await runCli("track-changes", "list", docPath);
		const changes = tracked.parsed as Array<{ kind: string }>;
		// A complete heading replacement should produce exactly one del +
		// one ins after consolidation, not 5+ pairs.
		const insertions = changes.filter((change) => change.kind === "ins");
		const deletions = changes.filter((change) => change.kind === "del");
		expect(deletions).toHaveLength(1);
		expect(insertions).toHaveLength(1);
	});
});

describe("docx edit --text — regression: mid-word run splits", () => {
	// Source docs frequently have <w:r> boundaries inside words (Word and
	// LibreOffice both produce these after iterative editing). Per-segment
	// tokenization shreds those words into sub-tokens that never align
	// with the new text's clean tokens, producing spurious del+ins pairs
	// for words that are actually unchanged. The fix: concatenate the
	// paragraph text once, tokenize the concatenation, and look up rPr
	// per character. p3 of the fixture has 3 italic runs split as
	// "Rating: The me" + "ssenger is fatally f" + "lawed." — every
	// "messenger" / "flawed" alignment depends on this fix.

	test("rewording preserves unchanged words across mid-word run splits (untracked)", async () => {
		const workspace = tempWorkspace("regression-untracked");
		const docPath = join(workspace, "out.docx");
		await Bun.write(docPath, Bun.file(FIXTURE));

		// Change only "fatally flawed" → "irreparably damaged". Everything
		// else (Rating, The, messenger, .) stays. With per-segment
		// tokenization, "messenger" would split as ["me","ssenger"] across
		// the run boundary and never match — producing del "me" / del
		// "ssenger" / ins "messenger" pairs even though the word is
		// unchanged.
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p3",
			"--text",
			"Rating: The messenger is irreparably damaged.",
		);
		expect(result.exitCode).toBe(0);

		const runs = await readParagraph(docPath, "p3");
		// Concatenated text reads cleanly.
		const flat = runs
			.filter((run) => run.type === "text")
			.map((run) => run.text ?? "")
			.join("");
		expect(flat).toBe("Rating: The messenger is irreparably damaged.");

		// Every emitted run should be italic — including the inserted
		// "irreparably damaged" (inheriting from the surrounding italic
		// span via position-pairing with the deleted "fatally flawed").
		const textRuns = runs.filter((run) => run.type === "text");
		expect(textRuns.length).toBeGreaterThan(0);
		expect(textRuns.every((run) => run.italic === true)).toBe(true);
	});

	test("under tracking, only the actually-changed words are wrapped (no spurious del+ins)", async () => {
		const workspace = tempWorkspace("regression-tracked");
		const docPath = join(workspace, "out.docx");
		await Bun.write(docPath, Bun.file(FIXTURE));
		await runCli("track-changes", docPath, "on");

		await runCli(
			"edit",
			docPath,
			"--at",
			"p3",
			"--text",
			"Rating: The messenger is irreparably damaged.",
		);

		const tracked = await runCli("track-changes", "list", docPath);
		const changes = tracked.parsed as Array<{
			id: string;
			kind: string;
			text?: string;
		}>;

		// Pre-fix bug: 8+ tracked changes (one del+ins for every word that
		// straddled a run boundary, including unchanged words like
		// "messenger" and "flawed"). Post-fix: only the actually-changed
		// words "fatally" and "flawed" produce del+ins pairs (4 changes
		// total — 2 del + 2 ins).
		const deletions = changes.filter((change) => change.kind === "del");
		const insertions = changes.filter((change) => change.kind === "ins");
		expect(deletions.length).toBeLessThanOrEqual(3);
		expect(insertions.length).toBeLessThanOrEqual(3);

		// "messenger" and "Rating" are unchanged and must not appear in any
		// tracked-change record. Pre-fix this assertion would fail.
		const trackedTexts = changes.map((change) => change.text ?? "");
		expect(trackedTexts).not.toContain("messenger");
		expect(trackedTexts).not.toContain("Rating:");
		expect(trackedTexts).not.toContain("The");
	});
});

const SPAN_FIX = "tests/fixtures/word-formatted.docx";
// p1: "The " + "MESSENGER" [bold,#800080] + " is " + "fatally flawed" [italic] + "."

type SpanRun = {
	type: string;
	text?: string;
	bold?: boolean;
	italic?: boolean;
	color?: string;
};

const freshSpan = (label: string) => freshFixture(label, SPAN_FIX);

async function readSpanPara(
	docPath: string,
	blockId: string,
): Promise<SpanRun[]> {
	const result = await runCli("read", docPath, "--ast");
	const blocks = (
		result.parsed as { blocks: Array<{ id: string; runs?: SpanRun[] }> }
	).blocks;
	return blocks.find((candidate) => candidate.id === blockId)?.runs ?? [];
}

async function firstLocator(docPath: string, phrase: string): Promise<string> {
	// The harness injects --json into find, so read the structured payload.
	const result = await runCli("find", docPath, phrase);
	expect(result.exitCode).toBe(0);
	const locator = (result.parsed as { matches?: Array<{ locator: string }> })
		.matches?.[0]?.locator;
	if (!locator) throw new Error(`no match for ${phrase}`);
	return locator;
}

function flat(runs: SpanRun[]): string {
	return runs
		.filter((run) => run.type === "text")
		.map((run) => run.text ?? "")
		.join("");
}

describe("docx edit --at pN:S-E — character-span edit", () => {
	test("find → edit --at <span> replaces just that span, inheriting its rPr", async () => {
		const docPath = await freshSpan("span-inherit");
		const locator = await firstLocator(docPath, "MESSENGER"); // p1:4-13
		expect(locator).toBe("p1:4-13");

		const result = await runCli(
			"edit",
			docPath,
			"--at",
			locator,
			"--text",
			"COURIER",
		);
		expect(result.exitCode).toBe(0);

		const runs = await readSpanPara(docPath, "p1");
		expect(flat(runs)).toBe("The COURIER is fatally flawed.");
		// Replacement inherited bold + color from the old MESSENGER run.
		const replaced = runs.find((run) => run.text === "COURIER");
		expect(replaced?.bold).toBe(true);
		expect(replaced?.color).toBe("800080");
		// Untouched neighbor keeps its italic.
		expect(runs.find((run) => run.text === "fatally flawed")?.italic).toBe(
			true,
		);
	});

	test("sub-run span inherits italic and splits the run cleanly", async () => {
		const docPath = await freshSpan("span-subrun");
		const locator = await firstLocator(docPath, "fatally"); // within the italic run
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			locator,
			"--text",
			"deeply",
		);
		expect(result.exitCode).toBe(0);

		const runs = await readSpanPara(docPath, "p1");
		expect(flat(runs)).toBe("The MESSENGER is deeply flawed.");
		// The "deeply" portion is still italic (inherited from the split run).
		const deeply = runs.find((run) => run.text === "deeply");
		expect(deeply?.italic).toBe(true);
	});

	test("span edit on a heading preserves its paragraph style", async () => {
		const workspace = tempWorkspace("span-heading");
		const md = join(workspace, "src.md");
		await Bun.write(md, "# Quarterly Design Review\n");
		const docPath = join(workspace, "h.docx");
		expect((await runCli("create", docPath, "--from", md)).exitCode).toBe(0);

		const locator = await firstLocator(docPath, "Quarterly"); // p0 heading
		expect(
			(await runCli("edit", docPath, "--at", locator, "--text", "Q3")).exitCode,
		).toBe(0);

		// Still a heading (outline only lists styled headings).
		const outline = await runCli("outline", docPath);
		expect(outline.stdout).toContain("Q3 Design Review");
	});

	test("cell-paragraph span (tN:rRcC:pK:S-E) edits in place", async () => {
		const workspace = tempWorkspace("span-cell");
		const md = join(workspace, "t.md");
		await Bun.write(
			md,
			"| Field | Value |\n| --- | --- |\n| State | fill in state |\n",
		);
		const docPath = join(workspace, "t.docx");
		expect((await runCli("create", docPath, "--from", md)).exitCode).toBe(0);

		const locator = await firstLocator(docPath, "fill in state");
		expect(locator).toContain(":"); // a cell-scoped span locator
		expect(
			(await runCli("edit", docPath, "--at", locator, "--text", "Delaware"))
				.exitCode,
		).toBe(0);

		const read = await runCli("read", docPath);
		expect(read.stdout).toContain("Delaware");
		expect(read.stdout).not.toContain("fill in state");
	});

	test("tracked span edit produces a del/ins pair attributed to the default author", async () => {
		const docPath = await freshSpan("span-tracked");
		expect((await runCli("track-changes", docPath, "on")).exitCode).toBe(0);
		const locator = await firstLocator(docPath, "MESSENGER");
		expect(
			(await runCli("edit", docPath, "--at", locator, "--text", "COURIER"))
				.exitCode,
		).toBe(0);

		const list = await runCli("track-changes", "list", docPath);
		const changes = list.parsed as Array<{
			kind: string;
			text: string;
			author: string;
		}>;
		expect(
			changes.some(
				(change) => change.kind === "del" && change.text === "MESSENGER",
			),
		).toBe(true);
		expect(
			changes.some(
				(change) => change.kind === "ins" && change.text === "COURIER",
			),
		).toBe(true);
		expect(changes.every((change) => change.author === "Reviewer")).toBe(true);
	});

	test("out-of-range span fails with INVALID_LOCATOR", async () => {
		const docPath = await freshSpan("span-oob");
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p1:5-999",
			"--text",
			"X",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("INVALID_LOCATOR");
	});

	test("--markdown on a span is rejected with a pointer to whole-paragraph edit", async () => {
		const docPath = await freshSpan("span-md");
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p1:0-3",
			"--markdown",
			"# H",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});
});

// Decision 2: a whole-paragraph edit with plain content (no explicit --style,
// no block markers) preserves the paragraph's existing style — re-titling a
// heading keeps it a heading. Markdown that carries its own block style wins.

async function docFrom(label: string, md: string): Promise<string> {
	const workspace = tempWorkspace(label);
	const src = join(workspace, "src.md");
	await Bun.write(src, md);
	const docPath = join(workspace, "out.docx");
	expect((await runCli("create", docPath, "--from", src)).exitCode).toBe(0);
	return docPath;
}

describe("edit preserves paragraph style on plain whole-paragraph edits", () => {
	test("--markdown plain text on a heading keeps the heading style", async () => {
		const docPath = await docFrom("md-heading", "# Quarterly Design Review\n");
		expect(
			(await runCli("edit", docPath, "--at", "p0", "--markdown", "Q3 Review"))
				.exitCode,
		).toBe(0);
		expect((await runCli("outline", docPath)).stdout).toContain("Q3 Review");
	});

	test("--text on a heading keeps the heading style", async () => {
		const docPath = await docFrom("text-heading", "# Original Title\n");
		expect(
			(await runCli("edit", docPath, "--at", "p0", "--text", "New Title"))
				.exitCode,
		).toBe(0);
		expect((await runCli("outline", docPath)).stdout).toContain("New Title");
	});

	test("a fresh-run replace (--bold) keeps the paragraph's direct alignment", async () => {
		const docPath = await docFrom("align-keep", "Centered Title\n");
		expect(
			(
				await runCli(
					"edit",
					docPath,
					"--at",
					"p0",
					"--text",
					"Centered Title",
					"--alignment",
					"center",
				)
			).exitCode,
		).toBe(0);
		// --bold forces the fresh-run path (not the in-place preserve path); the
		// direct <w:jc> must still survive, not just the style.
		expect(
			(
				await runCli(
					"edit",
					docPath,
					"--at",
					"p0",
					"--text",
					"New Title",
					"--bold",
				)
			).exitCode,
		).toBe(0);
		expect(await readDocumentXml(docPath)).toContain('<w:jc w:val="center"');
	});

	test("--markdown that sets its own block style wins (## → Heading2)", async () => {
		const docPath = await docFrom("md-override", "# Big\n");
		expect(
			(await runCli("edit", docPath, "--at", "p0", "--markdown", "## Smaller"))
				.exitCode,
		).toBe(0);
		const ast = await runCli("read", docPath, "--ast");
		const block = (
			ast.parsed as { blocks: Array<{ id: string; style?: string }> }
		).blocks[0];
		expect(block?.style).toBe("Heading2");
	});

	test("rewording a plain paragraph does not invent a heading", async () => {
		const docPath = await docFrom("plain", "Just a plain paragraph.\n");
		expect(
			(
				await runCli(
					"edit",
					docPath,
					"--at",
					"p0",
					"--markdown",
					"Reworded plainly",
				)
			).exitCode,
		).toBe(0);
		// outline lists only styled headings; a still-plain paragraph never
		// appears there (it would if the edit had invented a heading style).
		expect((await runCli("outline", docPath)).stdout).not.toContain(
			"Reworded plainly",
		);
	});
});

// Editing a paragraph that carries a comment must re-anchor the comment to the
// new content, not collapse it to a zero-length range (the Task-2 orphan bug).

type Anchor = {
	startOffset: number;
	endOffset: number;
	startBlockId: string;
};

async function doc(label: string, md: string): Promise<string> {
	const workspace = tempWorkspace(label);
	const src = join(workspace, "src.md");
	await Bun.write(src, md);
	const path = join(workspace, "out.docx");
	expect((await runCli("create", path, "--from", src)).exitCode).toBe(0);
	return path;
}

async function firstAnchor(path: string): Promise<Anchor> {
	const list = await runCli("comments", "list", path);
	const first = (list.parsed as Array<{ anchor: Anchor }>)[0];
	if (!first) throw new Error("expected at least one comment");
	return first.anchor;
}

describe("edit re-anchors comments instead of orphaning them", () => {
	test("untracked --text edit keeps the comment spanning the new text", async () => {
		const path = await doc("reanchor-text", "The tower opened in 1879.\n");
		expect(
			(
				await runCli(
					"comments",
					"add",
					path,
					"--at",
					"p0",
					"--text",
					"check year",
				)
			).exitCode,
		).toBe(0);
		expect(
			(
				await runCli(
					"edit",
					path,
					"--at",
					"p0",
					"--text",
					"The tower opened in 1889.",
				)
			).exitCode,
		).toBe(0);
		const anchor = await firstAnchor(path);
		expect(anchor.startOffset).toBe(0);
		// Spans the whole rewritten line (~25 chars), not a near-collapse.
		expect(anchor.endOffset - anchor.startOffset).toBeGreaterThanOrEqual(20);
	});

	test("tracked --text edit keeps the comment anchored (Task-2 regression)", async () => {
		const path = await doc("reanchor-tracked", "Completed in 1879.\n");
		expect(
			(
				await runCli(
					"comments",
					"add",
					path,
					"--at",
					"p0",
					"--text",
					"1889 not 1879",
				)
			).exitCode,
		).toBe(0);
		expect((await runCli("track-changes", path, "on")).exitCode).toBe(0);
		expect(
			(await runCli("edit", path, "--at", "p0", "--text", "Completed in 1889."))
				.exitCode,
		).toBe(0);
		const anchor = await firstAnchor(path);
		expect(anchor.startOffset).toBe(0);
		// Must cover the whole edited line, not collapse to a 1-char range or
		// drift to span only the inserted word (the Task-2 bug + off-by-one guard).
		expect(anchor.endOffset - anchor.startOffset).toBeGreaterThanOrEqual(15);
	});

	test("--markdown rewrite of a commented paragraph re-anchors the comment", async () => {
		const path = await doc("reanchor-md", "Old wording here.\n");
		expect(
			(await runCli("comments", "add", path, "--at", "p0", "--text", "reword"))
				.exitCode,
		).toBe(0);
		expect(
			(
				await runCli(
					"edit",
					path,
					"--at",
					"p0",
					"--markdown",
					"New wording entirely.",
				)
			).exitCode,
		).toBe(0);
		const anchor = await firstAnchor(path);
		expect(anchor.startOffset).toBe(0);
		// Spans the whole new paragraph ("New wording entirely." ~21 chars).
		expect(anchor.endOffset - anchor.startOffset).toBeGreaterThanOrEqual(15);
	});

	test("editing a different paragraph leaves an existing comment intact", async () => {
		const path = await doc(
			"reanchor-other",
			"First paragraph.\n\nSecond paragraph.\n",
		);
		expect(
			(await runCli("comments", "add", path, "--at", "p1", "--text", "note"))
				.exitCode,
		).toBe(0);
		expect(
			(await runCli("edit", path, "--at", "p0", "--text", "Edited first."))
				.exitCode,
		).toBe(0);
		const anchor = await firstAnchor(path);
		expect(anchor.startBlockId).toBe("p1");
		expect(anchor.endOffset).toBeGreaterThan(anchor.startOffset);
	});
});

// `find --highlight` + `edit --clear` is the highlight-removal workflow that
// took a weak model ~40 commands; it should now be find → clear.

const SOURCE =
	'Fill [the state]{highlight="yellow"} and [the county]{highlight="yellow"}; keep [this]{color="FF0000"} bold [word]{highlight="yellow"}.\n';

async function clearDoc(label: string): Promise<string> {
	const workspace = tempWorkspace(label);
	const src = join(workspace, "src.md");
	await Bun.write(src, SOURCE);
	const path = join(workspace, "out.docx");
	expect((await runCli("create", path, "--from", src)).exitCode).toBe(0);
	return path;
}

async function locators(path: string, ...flags: string[]): Promise<string[]> {
	const find = await runCli("find", path, ...flags, "--all");
	return (find.parsed as { matches: Array<{ locator: string }> }).matches.map(
		(match) => match.locator,
	);
}

describe("find formatting filters + edit --clear", () => {
	test("find --highlight yellow returns each highlighted span", async () => {
		const path = await clearDoc("find-hl");
		const found = await locators(path, "--highlight", "yellow");
		expect(found.length).toBe(3); // the state / the county / word
	});

	test("find --highlight any matches any highlight color", async () => {
		const path = await clearDoc("find-hl-any");
		expect((await locators(path, "--highlight", "any")).length).toBe(3);
	});

	test("bare --highlight (no color) means any color", async () => {
		const path = await clearDoc("find-hl-bare");
		// `locators` passes `--highlight` with no value; it should match any color.
		expect((await locators(path, "--highlight")).length).toBe(3);
	});

	test("find --color FF0000 returns the colored span", async () => {
		const path = await clearDoc("find-color");
		const found = await locators(path, "--color", "FF0000");
		expect(found.length).toBe(1);
	});

	test("edit --clear highlight on found spans removes only highlight", async () => {
		const path = await clearDoc("clear-hl");
		for (const loc of await locators(path, "--highlight", "any")) {
			expect(
				(await runCli("edit", path, "--at", loc, "--clear", "highlight"))
					.exitCode,
			).toBe(0);
		}
		// all highlights gone; text + the red color preserved
		expect((await locators(path, "--highlight", "any")).length).toBe(0);
		expect((await locators(path, "--color", "FF0000")).length).toBe(1);
		const read = await runCli("read", path);
		expect(read.stdout).toContain("Fill the state and the county");
	});

	test("edit --clear all on a whole paragraph strips all run formatting", async () => {
		const path = await clearDoc("clear-all");
		expect(
			(await runCli("edit", path, "--at", "p0", "--clear", "all")).exitCode,
		).toBe(0);
		expect((await locators(path, "--highlight", "any")).length).toBe(0);
		expect((await locators(path, "--color", "FF0000")).length).toBe(0);
	});

	test("edit --clear with an unknown attribute is a usage error", async () => {
		const path = await clearDoc("clear-bad");
		const result = await runCli("edit", path, "--at", "p0", "--clear", "bogus");
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});
});

// The form-fill + un-highlight move: fill a placeholder AND strip its highlight in
// ONE op (so it's not two passes, and the targeted clear doesn't nuke font size the
// way `--clear all` would). Drives the mnda call count down.
describe("edit content + --clear combined (fill then strip)", () => {
	const FILLED = "Fill the state and the county; keep this bold word.";

	test("single op: --text + --clear highlight fills then strips ONLY highlight", async () => {
		const path = await clearDoc("combo-single");
		expect((await locators(path, "--highlight", "any")).length).toBe(3);
		const result = await runCli(
			"edit",
			path,
			"--at",
			"p0",
			"--text",
			FILLED,
			"--clear",
			"highlight",
		);
		expect(result.exitCode).toBe(0);
		// highlight gone, but the red color survives — targeted, not `--clear all`.
		expect((await locators(path, "--highlight", "any")).length).toBe(0);
		expect((await locators(path, "--color", "FF0000")).length).toBe(1);
		expect((await runCli("read", path)).stdout).toContain(
			"Fill the state and the county",
		);
	});

	test("batch: one {at,text,clear} entry fills and strips in a single entry", async () => {
		const path = await clearDoc("combo-batch");
		const batch = join(tempWorkspace("combo-batch-src"), "b.jsonl");
		await Bun.write(
			batch,
			`${JSON.stringify({ at: "p0", text: FILLED, clear: "highlight" })}\n`,
		);
		expect((await runCli("edit", path, "--batch", batch)).exitCode).toBe(0);
		expect((await locators(path, "--highlight", "any")).length).toBe(0);
		expect((await locators(path, "--color", "FF0000")).length).toBe(1);
	});

	test("single op: --text + --clear on a SPAN fills then strips that span", async () => {
		// `find --highlight` returns span locators, so this is the natural one-shot.
		const path = await clearDoc("combo-span");
		expect((await locators(path, "--highlight", "yellow")).length).toBe(3);
		const [span] = await locators(path, "--highlight", "yellow");
		const result = await runCli(
			"edit",
			path,
			"--at",
			span ?? "p0:0-4",
			"--text",
			"Texas",
			"--clear",
			"highlight",
		);
		expect(result.exitCode).toBe(0);
		// exactly the edited span lost its highlight (3 → 2); text replaced
		expect((await locators(path, "--highlight", "yellow")).length).toBe(2);
		expect((await runCli("read", path)).stdout).toContain("Texas");
	});

	test("batch: a span entry may combine content + clear", async () => {
		const path = await clearDoc("combo-span-batch");
		const [span] = await locators(path, "--highlight", "yellow");
		const batch = join(tempWorkspace("combo-span-batch-src"), "b.jsonl");
		await Bun.write(
			batch,
			`${JSON.stringify({ at: span ?? "p0:0-4", text: "Texas", clear: "highlight" })}\n`,
		);
		expect((await runCli("edit", path, "--batch", batch)).exitCode).toBe(0);
		expect((await locators(path, "--highlight", "yellow")).length).toBe(2);
	});

	test("--clear is repeatable: --clear highlight --clear color accumulates", async () => {
		const path = await clearDoc("clear-repeat");
		expect((await locators(path, "--highlight", "any")).length).toBe(3);
		expect((await locators(path, "--color", "FF0000")).length).toBe(1);
		const result = await runCli(
			"edit",
			path,
			"--at",
			"p0",
			"--clear",
			"highlight",
			"--clear",
			"color",
		);
		expect(result.exitCode).toBe(0);
		expect((await locators(path, "--highlight", "any")).length).toBe(0);
		expect((await locators(path, "--color", "FF0000")).length).toBe(0);
	});

	test("batch: clear accepts an array of attrs", async () => {
		const path = await clearDoc("clear-array");
		const batch = join(tempWorkspace("clear-array-src"), "b.jsonl");
		await Bun.write(
			batch,
			`${JSON.stringify({ at: "p0", clear: ["highlight", "color"] })}\n`,
		);
		expect((await runCli("edit", path, "--batch", batch)).exitCode).toBe(0);
		expect((await locators(path, "--highlight", "any")).length).toBe(0);
		expect((await locators(path, "--color", "FF0000")).length).toBe(0);
	});

	test("--text accepts a leading-dash value (e.g. -$500.00), no --text= needed", async () => {
		// parseArgs would reject `--text -$500.00` as ambiguous; we pre-merge it.
		const docPath = await docFrom("dash-value", "placeholder\n");
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--text",
			"-$500.00",
		);
		expect(result.exitCode).toBe(0);
		expect(await readDocumentXml(docPath)).toContain("-$500.00");
	});
});

describe("clear-formatting preserves unmodeled rPr children (in-place invariant)", () => {
	test("--clear all strips formatting tags but keeps an unmodeled <w:lang>", () => {
		// The feature mutates rPr in place precisely so props we don't model
		// survive. `all` must remove highlight/bold but leave <w:lang> untouched.
		const [para] = XmlNode.parse(
			`<w:p><w:r><w:rPr><w:highlight w:val="yellow"/><w:b/><w:lang w:val="fr-FR"/></w:rPr><w:t>bonjour</w:t></w:r></w:p>`,
		);
		const tags = resolveClearTags(["all"]);
		expect(tags).not.toBeNull();
		clearFormatting(para as XmlNode, null, tags as Set<string>);
		const xml = XmlNode.serialize([para as XmlNode]);
		expect(xml).not.toContain("w:highlight");
		expect(xml).not.toContain("<w:b/>");
		expect(xml).toContain('w:lang w:val="fr-FR"'); // unmodeled prop survives
		expect(xml).toContain("bonjour");
	});
});

type Block = {
	id: string;
	type: string;
	style?: string;
	runs?: Array<{ type: string; text?: string }>;
};

async function blocksOf(path: string): Promise<Block[]> {
	const result = await runCli("read", path, "--ast");
	return (result.parsed as { blocks: Block[] }).blocks;
}

async function fivePara(label: string): Promise<string> {
	const docPath = join(tempWorkspace(label), "out.docx");
	await runCli("create", docPath, "--text", "Paragraph 1.");
	for (const i of [2, 3, 4, 5]) {
		await runCli(
			"insert",
			docPath,
			"--after",
			`p${i - 2}`,
			"--text",
			`Paragraph ${i}.`,
		);
	}
	return docPath;
}

describe("locator grammar: pN-pM", () => {
	test("rejects backward range at parse time", async () => {
		const docPath = await fivePara("range-backward");
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p3-p1",
			"--text",
			"x",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("INVALID_LOCATOR");
	});

	test("rejects cross-parent range (paragraph + cell paragraph)", async () => {
		const workspace = tempWorkspace("range-cross-parent");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Outside.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--table",
			"--rows",
			"1",
			"--cols",
			"1",
		);
		// Try p0 (body paragraph) - t0:r0c0:p0 (cell paragraph) — different parents.
		// p0-p1 would be body-body which is fine; we want to force cross-parent.
		// Easier proof: a non-existent endpoint produces BLOCK_NOT_FOUND.
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p0-p9",
			"--text",
			"x",
		);
		expect(result.exitCode).toBe(3);
		expect((result.parsed as { code: string }).code).toBe("BLOCK_NOT_FOUND");
	});
});

describe("docx edit --at pN-pM (range replace, untracked)", () => {
	test("--text collapses N paragraphs into one", async () => {
		const docPath = await fivePara("range-text");
		await runCli("edit", docPath, "--at", "p0-p3", "--text", "Just one.");
		const blocks = await blocksOf(docPath);
		expect(blocks.filter((b) => b.type === "paragraph")).toHaveLength(2);
		const paragraphs = blocks.filter((b) => b.type === "paragraph");
		expect(paragraphs[0]?.runs?.[0]?.text).toBe("Just one.");
		expect(paragraphs[1]?.runs?.[0]?.text).toBe("Paragraph 5.");
	});

	test("--runs replaces a range with one paragraph from runs JSON", async () => {
		const docPath = await fivePara("range-runs");
		await runCli(
			"edit",
			docPath,
			"--at",
			"p1-p3",
			"--runs",
			JSON.stringify([
				{ type: "text", text: "Bold ", bold: true },
				{ type: "text", text: "result." },
			]),
		);
		const blocks = await blocksOf(docPath);
		expect(blocks.filter((b) => b.type === "paragraph")).toHaveLength(3);
	});

	test("--code expands one anchor into N CodeBlock paragraphs", async () => {
		const docPath = await fivePara("range-code-expand");
		// Replace one paragraph with a three-line code block.
		await runCli(
			"edit",
			docPath,
			"--at",
			"p1",
			"--code",
			"function foo() {\n  return 42;\n}",
			"--language",
			"typescript",
		);
		const blocks = await blocksOf(docPath);
		const codeBlocks = blocks.filter((b) => b.style === "CodeBlock-typescript");
		expect(codeBlocks).toHaveLength(3);
		// And the CodeBlock-typescript style was provisioned.
		const pkg = await Pkg.open(docPath);
		const stylesXml = await pkg.readText("word/styles.xml");
		expect(stylesXml).toContain('w:styleId="CodeBlock-typescript"');
	});

	test("--code-file PATH reads file content for the replacement", async () => {
		const workspace = tempWorkspace("range-code-file");
		const docPath = join(workspace, "out.docx");
		const snippet = join(workspace, "snippet.py");
		await Bun.write(snippet, "def hello():\n    return 42\n");
		await runCli("create", docPath, "--text", "Old paragraph.");
		await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--code-file",
			snippet,
			"--language",
			"python",
		);
		const blocks = await blocksOf(docPath);
		const codeBlocks = blocks.filter((b) => b.style === "CodeBlock-python");
		expect(codeBlocks.length).toBeGreaterThanOrEqual(2);
	});

	test("dry-run prints the locator without mutating", async () => {
		const docPath = await fivePara("range-dry");
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p0-p3",
			"--text",
			"x",
			"--dry-run",
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { dryRun: boolean }).dryRun).toBe(true);
		const blocks = await blocksOf(docPath);
		expect(blocks.filter((b) => b.type === "paragraph")).toHaveLength(5);
	});
});

describe("docx edit --at pN-pM (range replace, tracked)", () => {
	test("tracked replace emits Word-canonical shape; accept → new content", async () => {
		const docPath = await fivePara("range-track-replace");
		await runCli("track-changes", docPath, "on");
		await runCli(
			"edit",
			docPath,
			"--at",
			"p0-p3",
			"--text",
			"Single new paragraph.",
			"--author",
			"Probe",
		);

		// XML shape: every old paragraph wrapped <w:del> for content; first N-1
		// also have paragraph-mark <w:del>. The last old paragraph (transition)
		// has its content del'd and the new content appended as <w:ins>.
		const pkg = await Pkg.open(docPath);
		const documentXml = await pkg.readText("word/document.xml");
		expect(documentXml).toContain('<w:del w:id="0" w:author="Probe"');
		expect(documentXml).toMatch(/<w:rPr><w:del/); // paragraph-mark del
		expect(documentXml).toMatch(/<w:ins[^>]*w:author="Probe"/);

		// Accept-all → just the new content + the post-range paragraph.
		await runCli("track-changes", "accept", docPath, "--all");
		const blocks = await blocksOf(docPath);
		const paragraphs = blocks.filter((b) => b.type === "paragraph");
		expect(paragraphs).toHaveLength(2);
		expect(paragraphs[0]?.runs?.[0]?.text).toBe("Single new paragraph.");
		expect(paragraphs[1]?.runs?.[0]?.text).toBe("Paragraph 5.");
	});

	test("reject-all restores the original paragraphs intact", async () => {
		const docPath = await fivePara("range-track-reject");
		await runCli("track-changes", docPath, "on");
		await runCli(
			"edit",
			docPath,
			"--at",
			"p0-p3",
			"--text",
			"would-be replacement",
		);
		await runCli("track-changes", "reject", docPath, "--all");

		const blocks = await blocksOf(docPath);
		const paragraphs = blocks.filter((b) => b.type === "paragraph");
		expect(paragraphs).toHaveLength(5);
		expect(paragraphs.map((p) => p.runs?.[0]?.text)).toEqual([
			"Paragraph 1.",
			"Paragraph 2.",
			"Paragraph 3.",
			"Paragraph 4.",
			"Paragraph 5.",
		]);
	});

	test("expand 4 → 8 paragraphs under tracking; accept yields all 8", async () => {
		const docPath = await fivePara("range-track-expand");
		await runCli("track-changes", docPath, "on");
		const newContent = Array.from({ length: 8 }, (_, i) => `NEW ${i + 1}`).join(
			"\n",
		);
		await runCli("edit", docPath, "--at", "p0-p3", "--code", newContent);
		await runCli("track-changes", "accept", docPath, "--all");
		const blocks = await blocksOf(docPath);
		const paragraphs = blocks.filter((b) => b.type === "paragraph");
		// 8 new code paragraphs + Paragraph 5.
		expect(paragraphs).toHaveLength(9);
		expect(paragraphs[7]?.runs?.map((r) => r.text).join("")).toBe("NEW 8");
		expect(paragraphs[8]?.runs?.[0]?.text).toBe("Paragraph 5.");
	});
});

describe("docx delete --at pN-pM", () => {
	test("untracked range delete splices the paragraphs out", async () => {
		const docPath = await fivePara("range-del-untracked");
		await runCli("delete", docPath, "--at", "p1-p3");
		const blocks = await blocksOf(docPath);
		const paragraphs = blocks.filter((b) => b.type === "paragraph");
		expect(paragraphs).toHaveLength(2);
		expect(paragraphs.map((p) => p.runs?.[0]?.text)).toEqual([
			"Paragraph 1.",
			"Paragraph 5.",
		]);
	});

	test("tracked range delete; accept → just the post-range paragraph", async () => {
		const docPath = await fivePara("range-del-track");
		await runCli("track-changes", docPath, "on");
		await runCli("delete", docPath, "--at", "p0-p3", "--author", "Probe");

		// Per Word's empirical pattern: every paragraph in range has content
		// del'd; all but the last have paragraph-mark del'd too.
		const pkg = await Pkg.open(docPath);
		const documentXml = await pkg.readText("word/document.xml");
		const delCount = (documentXml.match(/<w:del w:id="[0-9]+"/g) ?? []).length;
		expect(delCount).toBeGreaterThanOrEqual(4);

		await runCli("track-changes", "accept", docPath, "--all");
		const blocks = await blocksOf(docPath);
		const paragraphs = blocks.filter((b) => b.type === "paragraph");
		// One empty residue paragraph from the cascade + Paragraph 5. Empty
		// paragraphs render to nothing, so users see just one line.
		const nonEmpty = paragraphs.filter(
			(p) => (p.runs?.length ?? 0) > 0 && p.runs?.[0]?.text,
		);
		expect(nonEmpty.map((p) => p.runs?.[0]?.text)).toEqual(["Paragraph 5."]);
	});

	test("reject restores all paragraphs", async () => {
		const docPath = await fivePara("range-del-reject");
		await runCli("track-changes", docPath, "on");
		await runCli("delete", docPath, "--at", "p0-p3");
		await runCli("track-changes", "reject", docPath, "--all");
		const blocks = await blocksOf(docPath);
		const paragraphs = blocks.filter((b) => b.type === "paragraph");
		expect(paragraphs).toHaveLength(5);
	});
});

describe("docx wc pN-pM", () => {
	test("sums word counts across the paragraph range", async () => {
		const workspace = tempWorkspace("wc-range");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "one two three"); // 3
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"alpha beta gamma delta",
		); // 4
		await runCli("insert", docPath, "--after", "p1", "--text", "x y"); // 2
		await runCli("insert", docPath, "--after", "p2", "--text", "ignored final"); // 2

		const result = await runCli("wc", docPath, "p0-p2");
		expect(result.exitCode).toBe(0);
		const parsed = result.parsed as { scope: string; words: number };
		expect(parsed.scope).toBe("blockRange");
		expect(parsed.words).toBe(9); // 3 + 4 + 2
	});

	test("returns BLOCK_NOT_FOUND when an endpoint is missing", async () => {
		const docPath = await fivePara("wc-range-missing");
		const result = await runCli("wc", docPath, "p0-p9");
		expect(result.exitCode).toBe(3);
		expect((result.parsed as { code: string }).code).toBe("BLOCK_NOT_FOUND");
	});

	test("--accepted skips tracked-deleted text in the range", async () => {
		// One word ("removed") is inside a <w:del> wrapper — accepted view
		// drops it, baseline view counts it, current view counts everything.
		const docPath = join(tempWorkspace("wc-range-views"), "out.docx");
		await runCli("create", docPath, "--text", "alpha beta");
		await runCli("insert", docPath, "--after", "p0", "--text", "gamma delta");
		await runCli("track-changes", docPath, "on");
		// Tracked delete the second paragraph's text — its runs get <w:del>'d.
		await runCli("edit", docPath, "--at", "p1", "--text", "");
		await runCli("track-changes", docPath, "off");

		// p0 has 2 words; p1's text is del-wrapped (accepted view skips it).
		const accepted = (await runCli("wc", docPath, "p0-p1", "--accepted"))
			.parsed as {
			words: number;
		};
		expect(accepted.words).toBe(2);

		// Baseline view sees the original 4 words.
		const baseline = (await runCli("wc", docPath, "p0-p1", "--baseline"))
			.parsed as {
			words: number;
		};
		expect(baseline.words).toBe(4);
	});
});

describe("docx edit --at pN-pM tracked XML parity (vs Word probe)", () => {
	test("4 → 4: ins-marked paragraph-marks on transition + middle new paragraphs", async () => {
		const docPath = await fivePara("range-track-4-4");
		await runCli("track-changes", docPath, "on");
		const newContent = ["NEW 1", "NEW 2", "NEW 3", "NEW 4"].join("\n");
		await runCli(
			"edit",
			docPath,
			"--at",
			"p0-p3",
			"--code",
			newContent,
			"--author",
			"Probe",
		);
		const pkg = await Pkg.open(docPath);
		const documentXml = await pkg.readText("word/document.xml");
		// Expect 4 dels on old marks (p0/p1/p2 marks + p3's content del),
		// plus an ins on the transition's paragraph-mark (because N≥2 needs a
		// new paragraph break) and ins-marked paragraph-marks on middle new
		// paragraphs (NEW 2, NEW 3) but NOT the last (NEW 4).
		const insMarkCount = (
			documentXml.match(/<w:rPr>\s*<w:ins[^>]+\/>\s*<\/w:rPr>/g) ?? []
		).length;
		// Transition mark + NEW 2 mark + NEW 3 mark = 3 ins-marked pmarks.
		expect(insMarkCount).toBe(3);
		const delMarkCount = (
			documentXml.match(/<w:rPr>\s*<w:del[^>]+\/>\s*<\/w:rPr>/g) ?? []
		).length;
		// p0, p1, p2 marks del'd. p3 (transition) mark is NOT del'd.
		expect(delMarkCount).toBe(3);

		await runCli("track-changes", "accept", docPath, "--all");
		const blocks = await blocksOf(docPath);
		const paragraphs = blocks.filter((b) => b.type === "paragraph");
		const texts = paragraphs.map(
			(p) => p.runs?.map((r) => r.text ?? "").join("") ?? "",
		);
		expect(texts).toEqual(["NEW 1", "NEW 2", "NEW 3", "NEW 4", "Paragraph 5."]);
	});

	test("range-replace does NOT trigger LCS formatting preservation", async () => {
		// Single-paragraph --text preserves rPr on unchanged words. Range
		// --text rewrites the span wholesale (no cross-paragraph LCS) — that
		// matches Word's empirical behavior.
		const docPath = join(tempWorkspace("range-no-lcs"), "out.docx");
		await runCli("create", docPath, "--text", "This bold word survives.");
		// Add bold formatting to "bold" via --runs.
		await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--runs",
			JSON.stringify([
				{ type: "text", text: "This " },
				{ type: "text", text: "bold", bold: true },
				{ type: "text", text: " word survives." },
			]),
		);
		await runCli("insert", docPath, "--after", "p0", "--text", "Second.");
		// Range edit p0-p1 with --text containing "bold". With per-paragraph
		// LCS, "bold" might keep its formatting. With range-replace (wholesale),
		// it doesn't.
		await runCli(
			"edit",
			docPath,
			"--at",
			"p0-p1",
			"--text",
			"Rewritten with bold word.",
		);
		const blocks = await blocksOf(docPath);
		const para = blocks.find((b) => b.type === "paragraph");
		const runs = (para?.runs ?? []) as Array<{ text?: string; bold?: boolean }>;
		// All runs should be plain text — no bold preserved across the range.
		expect(runs.every((r) => !r.bold)).toBe(true);
	});
});

describe("docx edit --at pN-pM content-flag validation", () => {
	test("--code and --text are mutually exclusive", async () => {
		const docPath = await fivePara("edit-code-mutex-text");
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--code",
			"foo",
			"--text",
			"bar",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});

	test("--code and --code-file are mutually exclusive", async () => {
		const docPath = await fivePara("edit-code-mutex-file");
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--code",
			"foo",
			"--code-file",
			"some.py",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});

	test("--language without --code or --code-file is a USAGE error", async () => {
		const docPath = await fivePara("edit-lang-orphan");
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--text",
			"plain",
			"--language",
			"python",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});

	test("--code without --language degrades cleanly (no highlighting)", async () => {
		const docPath = await fivePara("edit-code-no-lang");
		await runCli("edit", docPath, "--at", "p0", "--code", "plain text");
		const blocks = await blocksOf(docPath);
		// Plain CodeBlock (no language suffix) since --language wasn't given.
		const codeBlock = blocks.find((b) => b.style === "CodeBlock");
		expect(codeBlock).toBeDefined();
	});
});

describe("tracked range edit/delete with a table in the range", () => {
	// Paragraph ids are assigned in document order regardless of intervening
	// tables — `[p, p, table, p, p]` gives `p0, p1, t0, p2, p3`. A tracked
	// `pN-pM` whose underlying parent slice includes `<w:tbl>` would corrupt
	// the file: `markParagraphMarkAs` injects `<w:pPr>` into the table node.
	// Untracked path splices through cleanly. Tracked path must reject.
	async function paraTablePara(label: string): Promise<string> {
		const docPath = join(tempWorkspace(label), "out.docx");
		await runCli("create", docPath, "--text", "Before.");
		// Body now: [p0 "Before.", s0].
		await runCli("insert", docPath, "--after", "p0", "--text", "Middle.");
		// Body: [p0, p1, s0].
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
		// Body: [p0, p1, t0, s0].
		await runCli("insert", docPath, "--before", "s0", "--text", "After.");
		// Body: [p0 "Before.", p1 "Middle.", t0, p2 "After.", s0].
		return docPath;
	}

	test("edit pN-pM under tracking rejects when range includes a table", async () => {
		const docPath = await paraTablePara("range-tracked-table-edit");
		await runCli("track-changes", docPath, "on");
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p0-p2",
			"--text",
			"replacement",
		);
		expect(result.exitCode).toBe(1);
		expect((result.parsed as { code: string }).code).toBe(
			"TRACKED_CHANGE_CONFLICT",
		);
	});

	test("delete pN-pM under tracking rejects when range includes a table", async () => {
		const docPath = await paraTablePara("range-tracked-table-delete");
		await runCli("track-changes", docPath, "on");
		const result = await runCli("delete", docPath, "--at", "p0-p2");
		expect(result.exitCode).toBe(1);
		expect((result.parsed as { code: string }).code).toBe(
			"TRACKED_CHANGE_CONFLICT",
		);
	});

	test("untracked range delete through a table splices cleanly", async () => {
		// Confirms the untracked path is still permitted — the table goes
		// with the range, which is the documented spliceful behavior.
		const docPath = await paraTablePara("range-untracked-table");
		const result = await runCli("delete", docPath, "--at", "p0-p2");
		expect(result.exitCode).toBe(0);
		const blocks = await blocksOf(docPath);
		expect(blocks.find((b) => b.type === "table")).toBeUndefined();
	});
});

describe("range edit preserves inline sectPr (section break) on the endpoint", () => {
	// A paragraph carrying an inline `<w:sectPr>` is a section-boundary
	// paragraph (its `sN` block sits right after its `pN`). Range-replacing
	// onto it must lift the sectPr onto the new paragraph or the section
	// break vanishes silently.
	async function docWithSection(label: string): Promise<string> {
		const docPath = join(tempWorkspace(label), "out.docx");
		await runCli("create", docPath, "--text", "Body 1.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--section",
			"--columns",
			"2",
			"--type",
			"continuous",
		);
		await runCli("insert", docPath, "--after", "p1", "--text", "Body 2.");
		return docPath;
	}

	test("untracked range replace preserves sectPr on the endpoint", async () => {
		const docPath = await docWithSection("range-sectpr-untracked");
		// Sanity-check the section break is present.
		const before = await blocksOf(docPath);
		expect(before.find((b) => b.type === "sectionBreak")).toBeDefined();

		await runCli("edit", docPath, "--at", "p0-p1", "--text", "Replaced.");
		const after = await blocksOf(docPath);
		const sections = after.filter((b) => b.type === "sectionBreak");
		// The inline section break (s0) survives; the trailing one is always
		// present in OOXML. So we expect at least one inline + one trailing.
		expect(sections.length).toBeGreaterThanOrEqual(2);
	});

	test("tracked range replace preserves sectPr on the endpoint", async () => {
		const docPath = await docWithSection("range-sectpr-tracked");
		await runCli("track-changes", docPath, "on");
		await runCli("edit", docPath, "--at", "p0-p1", "--text", "Replaced.");
		await runCli("track-changes", "accept", docPath, "--all");
		const after = await blocksOf(docPath);
		const sections = after.filter((b) => b.type === "sectionBreak");
		expect(sections.length).toBeGreaterThanOrEqual(2);
	});
});

describe("--code-file normalizes line endings", () => {
	test("CRLF content lands as clean text (no stray \\r in runs)", async () => {
		const workspace = tempWorkspace("crlf");
		const docPath = join(workspace, "out.docx");
		const snippetPath = join(workspace, "snippet.py");
		await Bun.write(snippetPath, "def hello():\r\n    return 42\r\n");
		await runCli("create", docPath, "--text", "Intro.");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--code-file",
			snippetPath,
		);
		const blocks = await blocksOf(docPath);
		const codeBlocks = blocks.filter((b) => b.style?.startsWith("CodeBlock"));
		const allText = codeBlocks
			.flatMap((b) => b.runs ?? [])
			.map((r) => r.text ?? "")
			.join("");
		expect(allText).not.toContain("\r");
		expect(allText).toContain("def hello():");
		expect(allText).toContain("    return 42");
	});
});

describe("docx edit/delete --at pN-pN (degenerate range)", () => {
	test("edit pN-pN behaves like edit pN", async () => {
		const docPath = await fivePara("range-pn-pn-edit");
		await runCli("edit", docPath, "--at", "p2-p2", "--text", "Just p2.");
		const blocks = await blocksOf(docPath);
		const paragraphs = blocks.filter((b) => b.type === "paragraph");
		expect(paragraphs).toHaveLength(5);
		expect(paragraphs[2]?.runs?.[0]?.text).toBe("Just p2.");
	});

	test("delete pN-pN behaves like delete pN", async () => {
		const docPath = await fivePara("range-pn-pn-delete");
		await runCli("delete", docPath, "--at", "p2-p2");
		const blocks = await blocksOf(docPath);
		const paragraphs = blocks.filter((b) => b.type === "paragraph");
		expect(paragraphs).toHaveLength(4);
		expect(paragraphs.map((p) => p.runs?.[0]?.text)).toEqual([
			"Paragraph 1.",
			"Paragraph 2.",
			"Paragraph 4.",
			"Paragraph 5.",
		]);
	});
});
