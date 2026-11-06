import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";
import { freshFixture } from "./helpers";

/** Issue #42 step 1: a field is atomic structure. Every mutator either keeps
 *  it intact, removes it whole, or refuses with nothing written — never a
 *  silent flatten to frozen text, never half a field. */

const PAPER = "tests/fixtures/multi-column.docx";
// p34: "…referenced by numbers in brackets [1, 3, 4]." — three complex
// `REF _Ref… \r \h` fields whose cached results are "1", "3", "4]".
const NEW_TEXT =
	"Use a numbered list of references at the end of the article, ordered alphabetically by surname of first author, and referenced by numbers in brackets [1, 3, 4].";

async function paragraphXml(path: string, locator: string): Promise<string> {
	return (await runCli("raw", "get", path, "--at", locator)).stdout;
}

const count = (xml: string, needle: string): number =>
	xml.split(needle).length - 1;

/** begin/separate/end counts — a balanced field contributes one of each. */
function fieldChars(xml: string): Record<string, number> {
	return {
		begin: count(xml, 'w:fldCharType="begin"'),
		separate: count(xml, 'w:fldCharType="separate"'),
		end: count(xml, 'w:fldCharType="end"'),
	};
}

async function expectUnchanged(
	path: string,
	run: () => Promise<{ exitCode: number; stdout: string }>,
	code: string,
): Promise<void> {
	const before = Buffer.from(await Bun.file(path).arrayBuffer());
	const result = await run();
	expect(result.exitCode).not.toBe(0);
	expect(result.stdout).toContain(code);
	const after = Buffer.from(await Bun.file(path).arrayBuffer());
	expect(after.equals(before)).toBe(true);
}

describe("whole-paragraph rebuilds refuse a paragraph holding fields", () => {
	const cases: [string, string[]][] = [
		["markdown (the default channel)", [NEW_TEXT]],
		["--text", ["--text", NEW_TEXT]],
		["--track --text", ["--track", "--text", NEW_TEXT]],
		["--track markdown", ["--track", NEW_TEXT]],
		["--runs", ["--runs", JSON.stringify([{ text: NEW_TEXT }])]],
	];
	for (const [label, flags] of cases) {
		test(label, async () => {
			const path = await freshFixture(`fields-${label}`, PAPER);
			await expectUnchanged(
				path,
				() => runCli("edit", path, "--at", "p34", ...flags),
				"FIELD_IN_PARAGRAPH",
			);
		});
	}

	test("the refusal names the fields by kind and cached result", async () => {
		const path = await freshFixture("fields-message", PAPER);
		const result = await runCli("edit", path, "--at", "p34", NEW_TEXT);
		expect(result.stdout).toContain(
			'3 fields (REF ×3: \\"1\\", \\"3\\", \\"4]\\")',
		);
	});

	test("--batch refuses the whole batch, nothing written", async () => {
		const path = await freshFixture("fields-batch", PAPER);
		const batch = join(tempWorkspace("fields-batch"), "edits.jsonl");
		await Bun.write(
			batch,
			[
				JSON.stringify({ at: "p0", text: "Changed title" }),
				JSON.stringify({ at: "p34", text: NEW_TEXT }),
			].join("\n"),
		);
		await expectUnchanged(
			path,
			() => runCli("edit", path, "--batch", batch),
			"FIELD_IN_PARAGRAPH",
		);
	});

	test("a range edit over a field paragraph refuses", async () => {
		const path = await freshFixture("fields-range", PAPER);
		await expectUnchanged(
			path,
			() => runCli("edit", path, "--at", "p33-p34", "--text", "x"),
			"FIELD_IN_PARAGRAPH",
		);
	});
});

describe("span edits keep, cover, or refuse fields", () => {
	test("replace beside the fields keeps all three", async () => {
		const path = await freshFixture("fields-beside", PAPER);
		const result = await runCli("replace", path, "last name", "surname");
		expect(result.exitCode).toBe(0);
		expect(count(await paragraphXml(path, "p34"), " REF _Ref")).toBe(3);
	});

	test("a span inside a field's result edits the cached text, field kept", async () => {
		const path = await freshFixture("fields-inside", PAPER);
		const result = await runCli(
			"edit",
			path,
			"--at",
			"p34:159-161",
			"--text",
			"5]",
		);
		expect(result.exitCode).toBe(0);
		const xml = await paragraphXml(path, "p34");
		expect(count(xml, " REF _Ref")).toBe(3);
		expect(fieldChars(xml)).toEqual({ begin: 3, separate: 3, end: 3 });
		expect(xml).toContain(
			'>5]</w:t></w:r><w:r w:rsidR="00DE3B36"><w:fldChar w:fldCharType="end"/>',
		);
	});

	test("a span covering whole fields removes them whole — no orphan end", async () => {
		const path = await freshFixture("fields-covered", PAPER);
		const result = await runCli("replace", path, "[1, 3, 4]", "[2]");
		expect(result.exitCode).toBe(0);
		const xml = await paragraphXml(path, "p34");
		expect(fieldChars(xml)).toEqual({ begin: 0, separate: 0, end: 0 });
		expect(xml).not.toContain("instrText");
		const validate = await runCli("validate", path);
		expect(validate.stdout.startsWith("valid")).toBe(true);
	});

	test("tracked, the covered fields are deleted whole: accept removes, reject restores", async () => {
		const path = await freshFixture("fields-covered-tracked", PAPER);
		const result = await runCli("replace", path, "[1, 3, 4]", "[2]", "--track");
		expect(result.exitCode).toBe(0);
		const tracked = await paragraphXml(path, "p34");
		expect(count(tracked, "<w:delInstrText")).toBe(3);
		expect(count(tracked, "<w:instrText")).toBe(0);

		const accepted = join(tempWorkspace("fields-accept"), "a.docx");
		await runCli("track-changes", "accept", path, "--all", "-o", accepted);
		expect(fieldChars(await paragraphXml(accepted, "p34"))).toEqual({
			begin: 0,
			separate: 0,
			end: 0,
		});

		const rejected = join(tempWorkspace("fields-reject"), "r.docx");
		await runCli("track-changes", "reject", path, "--all", "-o", rejected);
		const restored = await paragraphXml(rejected, "p34");
		expect(fieldChars(restored)).toEqual({ begin: 3, separate: 3, end: 3 });
		expect(count(restored, "<w:instrText")).toBe(3);
		expect(restored).not.toContain("delInstrText");
	});

	test("a span cutting a field is refused, nothing written", async () => {
		const path = await freshFixture("fields-cut", PAPER);
		await expectUnchanged(
			path,
			() => runCli("replace", path, "3, 4", "9"),
			"cuts a field",
		);
		await expectUnchanged(
			path,
			() => runCli("edit", path, "--at", "p34:156-160", "--text", "9"),
			"cuts a field",
		);
	});
});

/** A fresh document holding one paragraph per entry of `lines`. */
async function createParagraphs(
	label: string,
	lines: string[],
): Promise<string> {
	const workspace = tempWorkspace(label);
	const source = join(workspace, "source.md");
	await Bun.write(source, lines.join("\n\n"));
	const path = join(workspace, "doc.docx");
	expect((await runCli("create", path, "--from", source)).exitCode).toBe(0);
	return path;
}

/** A TOC-shaped field: begin + instruction + separate in one paragraph, its
 *  `end` in the next. */
async function twoParagraphField(label: string): Promise<string> {
	const path = await createParagraphs(label, [
		"Contents",
		"placeholder",
		"placeholder",
		"Body text",
	]);
	const begin =
		'<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Introduction 1</w:t></w:r></w:p>';
	const end =
		'<w:p><w:r><w:t>Methods 2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
	expect(
		(await runCli("raw", "replace", path, "--at", "p1", "--xml", begin))
			.exitCode,
	).toBe(0);
	expect(
		(await runCli("raw", "replace", path, "--at", "p2", "--xml", end)).exitCode,
	).toBe(0);
	return path;
}

describe("paragraph removal never splits a multi-paragraph field", () => {
	test("delete --at one end refuses", async () => {
		const path = await twoParagraphField("fields-split-delete");
		await expectUnchanged(
			path,
			() => runCli("delete", path, "--at", "p2"),
			"split a field",
		);
		await expectUnchanged(
			path,
			() => runCli("delete", path, "--at", "p0-p1"),
			"split a field",
		);
	});

	test('edit --text "" (line removal) refuses, single and --batch', async () => {
		const path = await twoParagraphField("fields-split-edit");
		await expectUnchanged(
			path,
			() => runCli("edit", path, "--at", "p1", "--text", ""),
			"split a field",
		);
		const batch = join(tempWorkspace("fields-split-batch"), "b.jsonl");
		await Bun.write(batch, JSON.stringify({ at: "p2", delete: true }));
		await expectUnchanged(
			path,
			() => runCli("edit", path, "--batch", batch),
			"split a field",
		);
	});

	test("rebuilding one end refuses (it holds a field marker)", async () => {
		const path = await twoParagraphField("fields-split-rebuild");
		await expectUnchanged(
			path,
			() => runCli("edit", path, "--at", "p1", "Introduction 3"),
			"FIELD_IN_PARAGRAPH",
		);
	});

	test("removing both ends takes the whole field", async () => {
		const path = await twoParagraphField("fields-split-both");
		expect((await runCli("delete", path, "--at", "p1-p2")).exitCode).toBe(0);
		const batchPath = await twoParagraphField("fields-split-both-batch");
		const batch = join(tempWorkspace("fields-split-both-b"), "b.jsonl");
		await Bun.write(
			batch,
			`${JSON.stringify({ at: "p1" })}\n${JSON.stringify({ at: "p2" })}`,
		);
		expect((await runCli("delete", batchPath, "--batch", batch)).exitCode).toBe(
			0,
		);
		for (const done of [path, batchPath]) {
			const validate = await runCli("validate", done);
			expect(validate.stdout.startsWith("valid")).toBe(true);
		}
	});
});

describe("cross-paragraph replace keeps a multi-paragraph field whole", () => {
	test("a match inside the field's result is fine", async () => {
		const path = await twoParagraphField("fields-across-inside");
		const result = await runCli(
			"replace",
			path,
			"Introduction 1\nMethods",
			"Overview",
		);
		expect(result.exitCode).toBe(0);
		const validate = await runCli("validate", path);
		expect(validate.stdout.startsWith("valid")).toBe(true);
		expect(fieldChars(await paragraphXml(path, "p1"))).toEqual({
			begin: 1,
			separate: 1,
			end: 1,
		});
	});

	test("a match taking the field's begin but not its end is refused", async () => {
		const path = await twoParagraphField("fields-across-cut");
		await expectUnchanged(
			path,
			() => runCli("replace", path, "Contents\nIntroduction", "Contents"),
			"cuts a field",
		);
	});
});

describe("tracked paragraph deletion writes delInstrText", () => {
	test("delete --track; reject restores <w:instrText>", async () => {
		const path = await freshFixture("fields-tracked-delete", PAPER);
		expect(
			(await runCli("delete", path, "--at", "p34", "--track")).exitCode,
		).toBe(0);
		const deleted = await paragraphXml(path, "p34");
		expect(count(deleted, "<w:delInstrText")).toBe(3);
		expect(count(deleted, "<w:instrText")).toBe(0);
		const rejected = join(tempWorkspace("fields-tracked-delete-r"), "r.docx");
		await runCli("track-changes", "reject", path, "--all", "-o", rejected);
		const restored = await paragraphXml(rejected, "p34");
		expect(count(restored, "<w:instrText")).toBe(3);
		expect(restored).not.toContain("delInstrText");
	});
});

describe("validate reports an unbalanced field", () => {
	test("an orphan end is an error", async () => {
		const path = await createParagraphs("fields-orphan", ["one", "two"]);
		await runCli(
			"raw",
			"replace",
			path,
			"--at",
			"p1",
			"--xml",
			'<w:p><w:r><w:t>two</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
			"--no-validate",
		);
		const result = await runCli("validate", path);
		expect(result.exitCode).toBe(1);
		expect(result.stdout).toContain("unbalanced field");
	});
});
