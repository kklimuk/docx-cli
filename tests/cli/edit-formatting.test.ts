import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

const FIXTURE = "tests/fixtures/word-formatted.docx";
// Layout (built by scripts/make-word-formatted-fixture.ts):
//   p0: plain — "The quick brown fox jumps over the lazy dog."
//   p1: "The " + "MESSENGER" [bold,#800080] + " is " + "fatally flawed" [italic] + "."
//   p2: "Bold" [bold] + " then " + "italic" [italic] + " then plain."
//   p3: "Rating: The me" + "ssenger is fatally f" + "lawed." (all italic, splits mid-word)

async function freshCopy(label: string): Promise<string> {
	const workspace = tempWorkspace(label);
	const docPath = join(workspace, "out.docx");
	await Bun.write(docPath, Bun.file(FIXTURE));
	return docPath;
}

type Run = {
	type: string;
	text?: string;
	bold?: boolean;
	italic?: boolean;
	color?: string;
	trackedChange?: { kind: string };
};

async function readParagraph(docPath: string, blockId: string): Promise<Run[]> {
	const result = await runCli("read", docPath);
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
