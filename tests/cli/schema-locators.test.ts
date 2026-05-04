import { describe, expect, test } from "bun:test";
import { runCli } from "./harness";

describe("docx schema", () => {
	test("default outputs JSON Schema", async () => {
		const result = await runCli("schema");
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
		const result = await runCli("schema", "--ts");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("export type Doc");
		expect(result.stdout).toContain("export type Paragraph");
		expect(result.stdout).toContain("CommentAnchor");
	});
});

describe("docx locators", () => {
	test("default outputs the grammar reference", async () => {
		const result = await runCli("locators");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("LOCATOR GRAMMAR");
		expect(result.stdout).toContain("p3:5-20");
	});

	test("--json outputs structured reference", async () => {
		const result = await runCli("locators", "--json");
		expect(result.exitCode).toBe(0);
		const reference = result.parsed as {
			blockLocators: Record<string, unknown>;
			spanLocator: { syntax: string };
		};
		expect(reference.spanLocator.syntax).toBe("pN:S-E");
	});
});
