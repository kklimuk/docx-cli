import { describe, expect, test } from "bun:test";
import { runCli } from "./harness";

describe("docx info schema", () => {
	test("default outputs JSON Schema", async () => {
		const result = await runCli("info", "schema");
		expect(result.exitCode).toBe(0);
		const schema = result.parsed as { $defs: Record<string, unknown> };
		expect(Object.keys(schema.$defs)).toEqual(
			expect.arrayContaining([
				"Block",
				"Paragraph",
				"Run",
				"TextRun",
				"ImageRun",
				"Comment",
			]),
		);
	});

	test("--ts outputs the live types.ts source", async () => {
		const result = await runCli("info", "schema", "--ts");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("export type DocProperties");
		expect(result.stdout).toContain("export type Paragraph");
		expect(result.stdout).toContain("CommentAnchor");
	});
});

describe("docx info locators", () => {
	test("default outputs the grammar reference", async () => {
		const result = await runCli("info", "locators");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("LOCATOR GRAMMAR");
		expect(result.stdout).toContain("p3:5-20");
	});

	test("--json outputs structured reference", async () => {
		const result = await runCli("info", "locators", "--json");
		expect(result.exitCode).toBe(0);
		const reference = result.parsed as {
			blockLocators: Record<string, unknown>;
			spanLocator: { syntax: string };
		};
		expect(reference.spanLocator.syntax).toBe("pN:S-E");
	});
});

describe("docx info dispatcher", () => {
	test("no topic prints help with usage exit", async () => {
		const result = await runCli("info");
		expect(result.exitCode).toBe(2);
		expect(result.stdout).toContain("docx info");
		expect(result.stdout).toContain("schema");
		expect(result.stdout).toContain("locators");
	});

	test("unknown topic returns USAGE", async () => {
		const result = await runCli("info", "nope");
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ ok: false, code: "USAGE" });
	});
});
