import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

/**
 * The read↔import formatting contract beyond the per-run round-trip in
 * run-formatting-roundtrip.test.ts:
 *   - the document `<!-- docx:base … -->` note (dominant font/size declared once,
 *     omitted per-run, re-applied on import),
 *   - default formatting (black / text1) dropped as noise,
 *   - hand-authored HTML formatting parsed on import (semantic tags + data-*),
 *   - insert blending into the anchor paragraph's formatting.
 */

type RunAst = {
	type: string;
	text?: string;
	bold?: boolean;
	italic?: boolean;
	strike?: boolean;
	font?: string;
	sizeHalfPoints?: number;
	color?: string;
	colorTheme?: string;
	highlight?: string;
	vertAlign?: string;
	underline?: string;
};
type Block = { id: string; type: string; runs?: RunAst[] };

async function read(path: string): Promise<string> {
	return (await runCli("read", path)).stdout;
}

async function blocks(path: string): Promise<Block[]> {
	const result = await runCli("read", path, "--ast");
	return (result.parsed as { blocks: Block[] }).blocks;
}

async function allRuns(path: string): Promise<RunAst[]> {
	return (await blocks(path)).flatMap((block) => block.runs ?? []);
}

/** A doc whose single inserted paragraph (p1) carries the given runs. */
async function docWith(label: string, runs: RunAst[]): Promise<string> {
	const path = join(tempWorkspace(label), "out.docx");
	await runCli("create", path, "--text", "Intro.");
	await runCli("insert", path, "--after", "p0", "--runs", JSON.stringify(runs));
	return path;
}

const ARIAL_BODY = { font: "Arial", sizeHalfPoints: 16 } as const;

describe("read — document format baseline note", () => {
	test("a dominant font/size becomes a base note and is omitted per-run", async () => {
		const path = await docWith("base-emit", [
			{
				type: "text",
				text: "The bulk of this document is in Arial eight point. ",
				...ARIAL_BODY,
			},
			{
				type: "text",
				text: "More of the same ubiquitous body content here too. ",
				...ARIAL_BODY,
			},
			{ type: "text", text: "TITLE", font: "Georgia", sizeHalfPoints: 44 },
		]);
		const md = await read(path);
		expect(md).toContain('<!-- docx:base font="Arial" size="8pt" -->');
		// Body runs carry no per-run font/size — they ride the note.
		expect(md).toContain("The bulk of this document is in Arial eight point.");
		expect(md).not.toContain("font-family:Arial");
		// The deviating run keeps its own font + size.
		expect(md).toContain("font-family:Georgia");
		expect(md).toContain("font-size:22pt");
	});

	test("a hostile dominant font value can't break or inject into the base note", async () => {
		const hostile = 'Bad"--><x'; // a quote, a comment-close, and angle brackets
		const path = await docWith("base-hostile", [
			{
				type: "text",
				text: "Body content one filling the document. ",
				font: hostile,
				sizeHalfPoints: 16,
			},
			{
				type: "text",
				text: "Body content two also in the same font. ",
				font: hostile,
				sizeHalfPoints: 16,
			},
		]);
		const md = await read(path);
		// Exactly one base note, escaped — the raw quote did NOT close the
		// attribute and the `-->` did NOT close the comment early.
		expect((md.match(/docx:base/g) ?? []).length).toBe(1);
		expect(md).not.toContain('font="Bad"');
		expect(md).not.toContain('Bad"--><x'); // raw hostile sequence never leaks
		expect(md).toContain("&quot;"); // it was escaped instead
		expect(md).toContain("Body content one filling the document.");
	});

	test("base note round-trips: create --from restores the omitted font/size", async () => {
		const src = await docWith("base-rt", [
			{
				type: "text",
				text: "Ubiquitous Arial body content fills this page. ",
				...ARIAL_BODY,
			},
			{
				type: "text",
				text: "Still more Arial body content to dominate it. ",
				...ARIAL_BODY,
			},
			{ type: "text", text: "Heading", font: "Georgia", sizeHalfPoints: 44 },
		]);
		const workspace = tempWorkspace("base-rt-dst");
		const mdPath = join(workspace, "doc.md");
		await Bun.write(mdPath, await read(src));
		const dst = join(workspace, "rt.docx");
		await runCli("create", dst, "--from", mdPath);

		const runs = (await allRuns(dst)).filter((r) => (r.text ?? "").trim());
		const body = runs.find((r) => (r.text ?? "").includes("Ubiquitous"));
		expect(body?.font).toBe("Arial");
		expect(body?.sizeHalfPoints).toBe(16);
		const heading = runs.find((r) => (r.text ?? "").includes("Heading"));
		expect(heading?.font).toBe("Georgia");
		expect(heading?.sizeHalfPoints).toBe(44);
	});
});

describe("read — default formatting dropped as noise", () => {
	test("black color and the text1 theme are not emitted; a real color is", async () => {
		const path = await docWith("default-color", [
			{
				type: "text",
				text: "plain black ",
				color: "000000",
				colorTheme: "text1",
			},
			{ type: "text", text: "teal", color: "107087" },
		]);
		const md = await read(path);
		expect(md).not.toContain("000000");
		expect(md).not.toContain("data-color-theme");
		expect(md).toContain('<span style="color:#107087">teal</span>');
	});
});

describe("import — hand-authored HTML formatting parses", () => {
	test("<mark>, <span style>, <sup>, <sub>, <u> map to run formatting", async () => {
		const workspace = tempWorkspace("html-import");
		const mdPath = join(workspace, "doc.md");
		await Bun.write(
			mdPath,
			'A <mark>hi</mark> <span style="color:#FF0000">red</span> <sup>up</sup> <sub>dn</sub> <u>ln</u> end.\n',
		);
		const dst = join(workspace, "out.docx");
		await runCli("create", dst, "--from", mdPath);
		const runs = await allRuns(dst);
		expect(runs.find((r) => r.text === "hi")?.highlight).toBe("yellow");
		expect(runs.find((r) => r.text === "red")?.color).toBe("FF0000");
		expect(runs.find((r) => r.text === "up")?.vertAlign).toBe("superscript");
		expect(runs.find((r) => r.text === "dn")?.vertAlign).toBe("subscript");
		expect(runs.find((r) => r.text === "ln")?.underline).toBe("single");
	});

	test("data-* attributes carry OOXML-only props CSS can't express", async () => {
		const workspace = tempWorkspace("html-data");
		const mdPath = join(workspace, "doc.md");
		await Bun.write(
			mdPath,
			'X <span data-color-theme="accent1">themed</span> <u data-underline="wave">wav</u> Y.\n',
		);
		const dst = join(workspace, "out.docx");
		await runCli("create", dst, "--from", mdPath);
		const runs = await allRuns(dst);
		expect(runs.find((r) => r.text === "themed")?.colorTheme).toBe("accent1");
		expect(runs.find((r) => r.text === "wav")?.underline).toBe("wave");
	});
});

describe("read — whitespace-only runs keep their formatting", () => {
	test("emphasis on a blank run uses HTML tags (not `** **`) and round-trips", async () => {
		const path = await docWith("blank-fmt", [
			{ type: "text", text: "a" },
			{ type: "text", text: " ", bold: true },
			{ type: "text", text: "b" },
			{ type: "text", text: " ", underline: "single" },
			{ type: "text", text: "c" },
		]);
		const md = await read(path);
		// No mis-parsing `** **`; bold/underline ride unambiguous HTML.
		expect(md).not.toContain("** **");
		expect(md).toContain("<b> </b>");
		expect(md).toContain("<u> </u>");

		const workspace = tempWorkspace("blank-fmt-rt");
		const mdPath = join(workspace, "doc.md");
		await Bun.write(mdPath, md);
		const dst = join(workspace, "rt.docx");
		await runCli("create", dst, "--from", mdPath);
		const blanks = (await allRuns(dst)).filter(
			(r) => r.type === "text" && (r.text ?? "").trim() === "",
		);
		expect(blanks.some((r) => (r as { bold?: boolean }).bold)).toBe(true);
		expect(blanks.some((r) => r.underline === "single")).toBe(true);
	});
});

describe("insert — inherits formatting from the anchor paragraph", () => {
	test("plain inserted text adopts the neighbor's font + size", async () => {
		const path = await docWith("blend", [
			{
				type: "text",
				text: "Existing Arial eight-point body text.",
				...ARIAL_BODY,
			},
		]);
		await runCli(
			"insert",
			path,
			"--after",
			"p1",
			"--markdown",
			"Brand new content.",
		);
		const inserted = (await allRuns(path)).find((r) =>
			(r.text ?? "").includes("Brand new"),
		);
		expect(inserted?.font).toBe("Arial");
		expect(inserted?.sizeHalfPoints).toBe(16);
	});

	test("inserting after a heading does NOT promote the new text to a heading", async () => {
		const path = join(tempWorkspace("blend-head"), "out.docx");
		await runCli("create", path, "--text", "Body.");
		await runCli("insert", path, "--after", "p0", "--markdown", "# A Heading");
		await runCli(
			"insert",
			path,
			"--after",
			"p1",
			"--markdown",
			"Body paragraph after the heading.",
		);
		const run = (await allRuns(path)).find((r) =>
			(r.text ?? "").includes("Body paragraph after the heading"),
		);
		// No heading size grafted on — the run stays plain body.
		expect(run?.sizeHalfPoints).toBeUndefined();
		expect(run?.font).toBeUndefined();
	});
});
