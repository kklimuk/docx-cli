import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

const FIXTURE = "tests/fixtures/normalize-query.docx";
// Layout (built by scripts/make-normalize-query-fixture.ts):
//   p0: 'The plan: "hello" world—ready to ship. The figure: 5 * 3 = 15.'
//       (smart quotes around hello)
//   p1: 'plan: "hello" today.' (straight quotes)

async function paragraphText(
	docPath: string,
	blockId: string,
): Promise<string> {
	const read = await runCli("read", docPath);
	const blocks = (
		read.parsed as {
			blocks: Array<{
				id: string;
				runs?: Array<{ type: string; text: string }>;
			}>;
		}
	).blocks;
	const block = blocks.find((candidate) => candidate.id === blockId);
	return (block?.runs ?? [])
		.filter((run) => run.type === "text")
		.map((run) => run.text)
		.join("");
}

describe("docx replace — pattern normalization", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("replace-norm");
		docPath = join(workspace, "out.docx");
		await Bun.write(docPath, Bun.file(FIXTURE));
	});

	test("strips markdown emphasis from the pattern; replacement is literal", async () => {
		// Default: replace just the first match (p0's smart-quote "hello").
		const result = await runCli("replace", docPath, "**hello**", "goodbye");
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as {
			normalizedPattern?: string;
			normalizationApplied?: string[];
		};
		expect(payload.normalizedPattern).toBe("hello");
		expect(payload.normalizationApplied).toContain("strip-md-emphasis");

		// p0's surrounding smart quotes are preserved (they weren't part of
		// the matched span); "goodbye" replaces just "hello".
		expect(await paragraphText(docPath, "p0")).toContain("“goodbye”");
	});

	test("smart-quote pattern matches straight-quote document text via canonicalization", async () => {
		// Smart-quote pattern with --all hits both p0 (smart in doc) and
		// p1 (straight in doc) thanks to canonicalization.
		const result = await runCli(
			"replace",
			docPath,
			"“hello”", // smart quotes in the pattern.
			"goodbye",
			"--all",
		);
		expect(result.exitCode).toBe(0);

		// Replacement is LITERAL: the matched span (smart quote + hello +
		// smart quote in p0; straight quote + hello + straight quote in p1)
		// is replaced wholesale by the literal "goodbye". Surrounding
		// punctuation is preserved.
		expect(await paragraphText(docPath, "p0")).toBe(
			"The plan: goodbye world—ready to ship. The figure: 5 * 3 = 15.",
		);
		expect(await paragraphText(docPath, "p1")).toBe("plan: goodbye today.");
	});

	test("--exact disables pattern normalization", async () => {
		const result = await runCli(
			"replace",
			docPath,
			"**hello**",
			"goodbye",
			"--exact",
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as {
			totalMatches: number;
			replaced: number;
			normalizedPattern?: string;
		};
		expect(payload.totalMatches).toBe(0);
		expect(payload.replaced).toBe(0);
		expect(payload.normalizedPattern).toBeUndefined();
		// Both paragraphs unchanged.
		expect(await paragraphText(docPath, "p0")).toContain("“hello”");
		expect(await paragraphText(docPath, "p1")).toContain('"hello"');
	});
});
