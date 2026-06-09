import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, spawnCli, tempWorkspace } from "./harness";
import { freshFixture } from "./helpers";

describe("docx find", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("find");
		docPath = join(workspace, "out.docx");
		await runCli(
			"create",
			docPath,
			"--text",
			"The quick brown fox jumps over the lazy dog.",
		);
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Another fox appears, then a Fox.",
		);
	});

	test("returns ALL matches by default (totalMatches == returned)", async () => {
		const result = await runCli("find", docPath, "fox");
		const payload = result.parsed as {
			totalMatches: number;
			matches: Array<{ locator: string; text: string }>;
		};
		expect(payload.totalMatches).toBe(2);
		expect(payload.matches).toHaveLength(2);
		expect(payload.matches.map((match) => match.locator)).toEqual([
			"p0:16-19",
			"p1:8-11",
		]);
	});

	test("--nth N selects a single match", async () => {
		const result = await runCli("find", docPath, "fox", "--nth", "0");
		const payload = result.parsed as {
			matches: Array<{ locator: string; text: string }>;
		};
		expect(payload.matches).toHaveLength(1);
		expect(payload.matches[0]).toMatchObject({
			locator: "p0:16-19",
			text: "fox",
		});
	});

	test("--all still returns every match (redundant: all is now default)", async () => {
		const result = await runCli("find", docPath, "fox", "--all");
		const payload = result.parsed as {
			matches: Array<{ locator: string }>;
		};
		expect(payload.matches.map((match) => match.locator)).toEqual([
			"p0:16-19",
			"p1:8-11",
		]);
	});

	test("no matches: empty stdout, 'no matches' on stderr, exit 0", async () => {
		// Text default (no --json injection), so spawn the real binary.
		const result = await spawnCli("find", docPath, "zzz-not-present-zzz");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toBe("");
		expect(result.stderr.trim()).toBe("no matches");
	});

	test("--ignore-case finds upper- and lowercase", async () => {
		const result = await runCli(
			"find",
			docPath,
			"fox",
			"--ignore-case",
			"--all",
		);
		const payload = result.parsed as {
			matches: Array<{ locator: string; text: string }>;
		};
		expect(payload.matches).toHaveLength(3);
		expect(payload.matches[2]).toMatchObject({
			locator: "p1:28-31",
			text: "Fox",
		});
	});

	test("--regex supports JS regex syntax", async () => {
		const result = await runCli(
			"find",
			docPath,
			"(quick|lazy)",
			"--regex",
			"--all",
		);
		const payload = result.parsed as {
			matches: Array<{ text: string }>;
		};
		expect(payload.matches.map((match) => match.text)).toEqual([
			"quick",
			"lazy",
		]);
	});

	test("--nth picks a specific match", async () => {
		const result = await runCli("find", docPath, "fox", "--nth", "1");
		const payload = result.parsed as {
			matches: Array<{ locator: string }>;
		};
		expect(payload.matches[0]?.locator).toBe("p1:8-11");
	});

	test("--nth out of range is MATCH_NOT_FOUND with exit 3", async () => {
		const result = await runCli("find", docPath, "fox", "--nth", "5");
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({
			code: "MATCH_NOT_FOUND",
		});
	});

	test("zero matches returns empty array", async () => {
		const result = await runCli("find", docPath, "absent");
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			totalMatches: 0,
			matches: [],
		});
	});

	test("output composes with comments add", async () => {
		const find = await runCli("find", docPath, "fox");
		const locator = (
			find.parsed as {
				matches: Array<{ locator: string }>;
			}
		).matches[0]?.locator;
		expect(locator).toBe("p0:16-19");

		const add = await runCli(
			"comments",
			"add",
			docPath,
			"--at",
			locator ?? "p0",
			"--text",
			"Reconsider",
			"--author",
			"QA",
		);
		expect(add.exitCode).toBe(0);

		const list = await runCli("comments", "list", docPath);
		const comments = list.parsed as Array<{
			anchor: { startOffset: number; endOffset: number };
		}>;
		expect(comments[0]?.anchor.startOffset).toBe(16);
		expect(comments[0]?.anchor.endOffset).toBe(19);
	});

	test("invalid regex returns USAGE error", async () => {
		const result = await runCli("find", docPath, "(unclosed", "--regex");
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("matches inside table cells get tT:rRcC:pK locators", async () => {
		const fixture = "tests/fixtures/tables-and-lists.docx";
		const result = await runCli("find", fixture, "Equipment", "--all");
		const payload = result.parsed as {
			matches: Array<{ locator: string; blockId: string }>;
		};
		const cellMatch = payload.matches.find((match) =>
			match.blockId.startsWith("t0:r"),
		);
		expect(cellMatch).toBeDefined();
		expect(cellMatch?.blockId).toBe("t0:r0c0:p0");
		expect(cellMatch?.locator).toBe("t0:r0c0:p0:0-9");
	});

	test("table-cell locator composes with comments add", async () => {
		const workspace = tempWorkspace("find-cell-comments");
		const cellDocPath = join(workspace, "cells.docx");
		await Bun.write(
			cellDocPath,
			Bun.file("tests/fixtures/tables-and-lists.docx"),
		);

		const find = await runCli("find", cellDocPath, "Breadboard", "--all");
		const cellLocator = (
			find.parsed as { matches: Array<{ locator: string }> }
		).matches.find((match) => match.locator.startsWith("t0:r"))?.locator;
		expect(cellLocator).toBeDefined();

		const add = await runCli(
			"comments",
			"add",
			cellDocPath,
			"--at",
			cellLocator ?? "p0",
			"--text",
			"This row",
			"--author",
			"QA",
		);
		expect(add.exitCode).toBe(0);

		const list = await runCli("comments", "list", cellDocPath);
		const comments = list.parsed as Array<{
			anchor: { startBlockId: string; endBlockId: string };
		}>;
		expect(comments[0]?.anchor.startBlockId).toBe("t0:r3c0:p0");
		expect(comments[0]?.anchor.endBlockId).toBe("t0:r3c0:p0");
	});
});

const NORMALIZE_FIXTURE = "tests/fixtures/normalize-query.docx";
// Layout (built by tests/fixtures/setup/normalize-query.ts):
//   p0: 'The plan: "hello" world—ready to ship. The figure: 5 * 3 = 15.'
//       (smart quotes around hello, em-dash, bare " * " between digits)
//   p1: 'plan: "hello" today.' (straight quotes)

describe("docx find — query normalization", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("find-norm");
		docPath = join(workspace, "out.docx");
		await Bun.write(docPath, Bun.file(NORMALIZE_FIXTURE));
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

// tracked-changes.docx layout (paragraph p0):
//   "This is a text with " (0..20, plain)
//   "two exciting "        (inside <w:ins>)
//   "insertions."          (plain, after the ins)
//
// Visible text per view:
//   current   — "This is a text with two exciting insertions." (44)
//   accepted  — same as current (ins is kept)            (44)
//   baseline  — "This is a text with insertions."        (31)
const TRACKED_FIXTURE = "tests/fixtures/tracked-changes.docx";

const freshCopy = (label: string) => freshFixture(label, TRACKED_FIXTURE);

describe("docx find — views", () => {
	test("default (accepted) reports the view in the response", async () => {
		const docPath = await freshCopy("find-view-default");
		const result = await runCli("find", docPath, "insertions");
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as {
			view: string;
			matches: Array<{ start: number; end: number }>;
		};
		expect(payload.view).toBe("accepted");
		expect(payload.matches[0]?.start).toBe(33);
	});

	test("--current view sees the same text as accepted when only ins is present", async () => {
		const docPath = await freshCopy("find-view-current");
		const result = await runCli("find", docPath, "insertions", "--current");
		const payload = result.parsed as {
			view: string;
			matches: Array<{ start: number }>;
		};
		expect(payload.view).toBe("current");
		expect(payload.matches[0]?.start).toBe(33);
	});

	test("--baseline drops the inserted text and shifts offsets accordingly", async () => {
		const docPath = await freshCopy("find-view-baseline");
		const result = await runCli("find", docPath, "insertions", "--baseline");
		const payload = result.parsed as {
			view: string;
			matches: Array<{ start: number; end: number }>;
		};
		expect(payload.view).toBe("baseline");
		// Without the 13-char "two exciting " insertion, "insertions" starts at 20.
		expect(payload.matches[0]?.start).toBe(20);
		expect(payload.matches[0]?.end).toBe(30);
	});

	test("--baseline does not match text that lives only inside <w:ins>", async () => {
		const docPath = await freshCopy("find-view-baseline-miss");
		const result = await runCli("find", docPath, "exciting", "--baseline");
		const payload = result.parsed as { matches: unknown[] };
		expect(payload.matches).toEqual([]);
	});

	test("--current and --baseline together are rejected", async () => {
		const docPath = await freshCopy("find-view-mutex");
		const result = await runCli(
			"find",
			docPath,
			"insertions",
			"--current",
			"--baseline",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});
});

// clear-formatting covers the --highlight / --color filters; --bold, --italic,
// --underline (and their intersection) were untested. Spans of the seeded
// paragraph p1: "plain "(0-6) "bolded"(6-12,b) " mid "(12-17,i)
// "under"(17-22,u) " both"(22-27,b+i).
describe("docx find — bold / italic / underline formatting filters", () => {
	async function formattedDoc(label: string): Promise<string> {
		const docPath = join(tempWorkspace(label), "out.docx");
		await runCli("create", docPath, "--text", "seed");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--runs",
			JSON.stringify([
				{ type: "text", text: "plain " },
				{ type: "text", text: "bolded", bold: true },
				{ type: "text", text: " mid ", italic: true },
				{ type: "text", text: "under", underline: "single" },
				{ type: "text", text: " both", bold: true, italic: true },
			]),
		);
		return docPath;
	}

	async function locators(
		docPath: string,
		...flags: string[]
	): Promise<string[]> {
		const result = await runCli("find", docPath, ...flags); // harness injects --json
		return (
			result.parsed as { matches: Array<{ locator: string }> }
		).matches.map((match) => match.locator);
	}

	test("--bold --all returns every bold span", async () => {
		const docPath = await formattedDoc("find-bold");
		expect(await locators(docPath, "--bold", "--all")).toEqual([
			"p1:6-12",
			"p1:22-27",
		]);
	});

	test("--italic --all returns every italic span", async () => {
		const docPath = await formattedDoc("find-italic");
		expect(await locators(docPath, "--italic", "--all")).toEqual([
			"p1:12-17",
			"p1:22-27",
		]);
	});

	test("--underline --all returns the underlined span", async () => {
		const docPath = await formattedDoc("find-underline");
		expect(await locators(docPath, "--underline", "--all")).toEqual([
			"p1:17-22",
		]);
	});

	test("--bold --italic intersects (only the run carrying both)", async () => {
		const docPath = await formattedDoc("find-bold-italic");
		expect(await locators(docPath, "--bold", "--italic", "--all")).toEqual([
			"p1:22-27",
		]);
	});
});
