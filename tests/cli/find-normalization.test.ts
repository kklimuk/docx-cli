import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

const FIXTURE = "tests/fixtures/normalize-query.docx";
// Layout (built by tests/fixtures/setup/normalize-query.ts):
//   p0: 'The plan: "hello" world—ready to ship. The figure: 5 * 3 = 15.'
//       (smart quotes around hello, em-dash, bare " * " between digits)
//   p1: 'plan: "hello" today.' (straight quotes)

describe("docx find — query normalization", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("find-norm");
		docPath = join(workspace, "out.docx");
		await Bun.write(docPath, Bun.file(FIXTURE));
	});

	test("strips balanced markdown emphasis around non-whitespace", async () => {
		const result = await runCli("find", docPath, "**hello**", "--all");
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as {
			totalMatches: number;
			matches: Array<{ text: string }>;
			normalizedQuery?: string;
			normalizationApplied?: string[];
		};
		// Both p0 (smart quotes around hello) and p1 (straight quotes) match
		// the stripped query "hello".
		expect(payload.totalMatches).toBe(2);
		expect(payload.matches[0]?.text).toBe("hello");
		expect(payload.normalizedQuery).toBe("hello");
		expect(payload.normalizationApplied).toContain("strip-md-emphasis");
	});

	test("does not strip an unmatched single asterisk used as multiplication", async () => {
		const result = await runCli("find", docPath, "5 * 3", "--all");
		const payload = result.parsed as {
			totalMatches: number;
			matches: Array<{ text: string }>;
			normalizedQuery?: string;
		};
		expect(payload.totalMatches).toBe(1);
		expect(payload.matches[0]?.text).toBe("5 * 3");
		// No emphasis stripping should happen for "5 * 3".
		expect(payload.normalizedQuery).toBeUndefined();
	});

	test("smart quotes in the query match the document's smart quotes", async () => {
		// The doc literally contains smart quotes; the query passes straight quotes.
		const result = await runCli("find", docPath, '"hello"');
		const payload = result.parsed as {
			matches: Array<{ text: string }>;
		};
		// Match preserves original-haystack characters (smart quotes).
		expect(payload.matches[0]?.text).toBe("“hello”");
	});

	test("em-dash in the document matches a hyphen in the query", async () => {
		// Body contains "world—ready"; the query passes a hyphen.
		const result = await runCli("find", docPath, "world-ready");
		const payload = result.parsed as {
			matches: Array<{ text: string }>;
		};
		expect(payload.matches).toHaveLength(1);
		// The match preserves the original document text (with em-dash).
		expect(payload.matches[0]?.text).toBe("world—ready");
	});

	test("--exact disables normalization entirely", async () => {
		const result = await runCli("find", docPath, "**hello**", "--exact");
		const payload = result.parsed as {
			totalMatches: number;
			normalizedQuery?: string;
		};
		expect(payload.totalMatches).toBe(0);
		expect(payload.normalizedQuery).toBeUndefined();
	});

	test("--regex bypasses normalization (raw regex)", async () => {
		const result = await runCli("find", docPath, "h[aeiou]llo", "--regex");
		const payload = result.parsed as {
			matches: Array<{ text: string }>;
			normalizedQuery?: string;
		};
		expect(payload.matches[0]?.text).toBe("hello");
		expect(payload.normalizedQuery).toBeUndefined();
	});
});
