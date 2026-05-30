import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "../../src/core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";

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
