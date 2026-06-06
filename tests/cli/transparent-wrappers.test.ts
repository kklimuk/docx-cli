import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";

/**
 * Regression coverage for the AST/XML offset alignment fix: when a paragraph
 * contains a transparent wrapper (<w:fldSimple>, <w:smartTag>) the AST
 * surfaces the inner text as part of `paragraph.runs`, so `find` reports
 * offsets that include the wrapper's content. Before the fix, the XML-side
 * traversal in `comments/helpers.tsx`, `replace/replace-span.tsx`, and
 * `hyperlinks/wrap.tsx` skipped these wrappers, producing offsets that
 * disagreed with the AST and breaking find→replace / find→comments-add
 * pipelines.
 *
 * Fixture (tests/fixtures/transparent-wrappers.docx):
 *   p0: "Today is " + <w:fldSimple>"2026-05-05"</w:fldSimple> + "."
 *   p1: "Hello " + <w:smartTag>"Alice"</w:smartTag> + "."
 */
const FIXTURE = "tests/fixtures/transparent-wrappers.docx";

async function freshCopy(label: string): Promise<string> {
	const workspace = tempWorkspace(label);
	const docPath = join(workspace, "out.docx");
	await Bun.write(docPath, Bun.file(FIXTURE));
	return docPath;
}

describe("transparent wrappers — find offsets include inner text", () => {
	test("find locates text inside <w:fldSimple>", async () => {
		const result = await runCli("find", FIXTURE, "2026-05-05");
		const payload = result.parsed as {
			totalMatches: number;
			matches: Array<{ locator: string; text: string }>;
		};
		expect(payload.totalMatches).toBe(1);
		expect(payload.matches[0]?.text).toBe("2026-05-05");
		// "Today is " is 9 chars; the field content starts at offset 9.
		expect(payload.matches[0]?.locator).toBe("p0:9-19");
	});

	test("find locates text inside <w:smartTag>", async () => {
		const result = await runCli("find", FIXTURE, "Alice");
		const payload = result.parsed as {
			totalMatches: number;
			matches: Array<{ locator: string; text: string }>;
		};
		expect(payload.totalMatches).toBe(1);
		// "Hello " is 6 chars; "Alice" starts at offset 6.
		expect(payload.matches[0]?.locator).toBe("p1:6-11");
	});

	test("find locates text spanning past the wrapper", async () => {
		// "is 2026" crosses the boundary into <w:fldSimple>.
		const result = await runCli("find", FIXTURE, "is 2026");
		const payload = result.parsed as {
			totalMatches: number;
			matches: Array<{ locator: string }>;
		};
		expect(payload.totalMatches).toBe(1);
		expect(payload.matches[0]?.locator).toBe("p0:6-13");
	});
});

describe("transparent wrappers — replace inside the wrapper", () => {
	test("replace text fully inside <w:fldSimple> swaps the cached field result", async () => {
		const docPath = await freshCopy("transparent-replace-fld");
		const result = await runCli("replace", docPath, "2026-05-05", "today");
		expect(result.exitCode).toBe(0);

		// Re-read and confirm the new text appears at the right paragraph.
		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as {
			blocks: Array<{
				type: string;
				runs?: Array<{ type: string; text?: string }>;
			}>;
		};
		const p0Text = (doc.blocks[0]?.runs ?? [])
			.filter((run) => run.type === "text")
			.map((run) => run.text ?? "")
			.join("");
		expect(p0Text).toBe("Today is today.");

		// The "." run on the right side of the fldSimple should still be there.
		expect(p0Text.endsWith(".")).toBe(true);
	});

	test("replace text fully inside <w:smartTag> swaps the inner content", async () => {
		const docPath = await freshCopy("transparent-replace-smart");
		const result = await runCli("replace", docPath, "Alice", "Bob");
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as {
			blocks: Array<{
				type: string;
				runs?: Array<{ type: string; text?: string }>;
			}>;
		};
		const p1Text = (doc.blocks[1]?.runs ?? [])
			.filter((run) => run.type === "text")
			.map((run) => run.text ?? "")
			.join("");
		expect(p1Text).toBe("Hello Bob.");
	});

	test("replace span crossing INTO a wrapper splits it cleanly", async () => {
		const docPath = await freshCopy("transparent-replace-cross");
		// "is 2026" crosses the fldSimple boundary; replacing it should leave
		// the rest of the date inside a (now-shorter) fldSimple half.
		const result = await runCli("replace", docPath, "is 2026", "was 2025");
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as {
			blocks: Array<{
				type: string;
				runs?: Array<{ type: string; text?: string }>;
			}>;
		};
		const p0Text = (doc.blocks[0]?.runs ?? [])
			.filter((run) => run.type === "text")
			.map((run) => run.text ?? "")
			.join("");
		expect(p0Text).toBe("Today was 2025-05-05.");
	});

	test("the fldSimple wrapper survives a fully-internal replace", async () => {
		const docPath = await freshCopy("transparent-replace-survives");
		await runCli("replace", docPath, "2026-05-05", "today");

		const pkg = await Pkg.open(docPath);
		const xml = await pkg.readText("word/document.xml");
		// fldSimple still wraps the (replaced) inner content.
		expect(xml).toContain("<w:fldSimple");
	});
});

describe("transparent wrappers — comments add against the wrapper", () => {
	test("comments add accepts a span that ends past <w:fldSimple>", async () => {
		const docPath = await freshCopy("transparent-comment-fld");
		// p0 length is 20: "Today is " (9) + "2026-05-05" (10) + "." (1).
		// Range covers "is 2026-05-05" (offsets 6-19).
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--at",
			"p0:6-19",
			"--text",
			"verify the date",
			"--author",
			"Tester",
		);
		expect(result.exitCode).toBe(0);
	});

	test("comments add accepts a span that ends past <w:smartTag>", async () => {
		const docPath = await freshCopy("transparent-comment-smart");
		// p1 length is 12: "Hello " (6) + "Alice" (5) + "." (1).
		// Range covers "Alice." (offsets 6-12).
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--at",
			"p1:6-12",
			"--text",
			"check the name",
			"--author",
			"Tester",
		);
		expect(result.exitCode).toBe(0);
	});

	test("comments add against a span fully past the wrapper succeeds with correct length", async () => {
		const docPath = await freshCopy("transparent-comment-past");
		// p0 length is 20; targeting the trailing "." at offset 19-20 should work.
		const result = await runCli(
			"comments",
			"add",
			docPath,
			"--at",
			"p0:19-20",
			"--text",
			"period",
			"--author",
			"Tester",
		);
		expect(result.exitCode).toBe(0);
	});
});
