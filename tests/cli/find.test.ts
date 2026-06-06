import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

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

	test("returns first match by default", async () => {
		const result = await runCli("find", docPath, "fox");
		const payload = result.parsed as {
			totalMatches: number;
			matches: Array<{ locator: string; text: string }>;
		};
		expect(payload.totalMatches).toBe(2);
		expect(payload.matches).toHaveLength(1);
		expect(payload.matches[0]).toMatchObject({
			locator: "p0:16-19",
			text: "fox",
		});
	});

	test("--all returns every match in document order", async () => {
		const result = await runCli("find", docPath, "fox", "--all");
		const payload = result.parsed as {
			matches: Array<{ locator: string }>;
		};
		expect(payload.matches.map((match) => match.locator)).toEqual([
			"p0:16-19",
			"p1:8-11",
		]);
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
