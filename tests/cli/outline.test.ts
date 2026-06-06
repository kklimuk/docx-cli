import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

type OutlineEntry = {
	id: string;
	locator: string;
	level: number;
	style: string;
	text: string;
	children: OutlineEntry[];
};

describe("docx outline", () => {
	test("builds a hierarchy from the academic-paper fixture", async () => {
		const result = await runCli(
			"outline",
			"tests/fixtures/academic-paper.docx",
		);
		expect(result.exitCode).toBe(0);
		const payload = result.parsed as OutlineEntry[];
		const topLevelTexts = payload.map((entry) => entry.text);
		expect(topLevelTexts).toContain("Guided Imagery");
		expect(topLevelTexts).toContain("Conclusion");
		expect(topLevelTexts).toContain("References");

		const guidedImagery = payload.find(
			(entry) => entry.text === "Guided Imagery",
		);
		expect(guidedImagery?.level).toBe(1);
		expect(guidedImagery?.children.map((child) => child.text)).toEqual([
			"Features of Guided Imagery",
			"Guided Imagery in Group Psychotherapy",
		]);
		expect(guidedImagery?.children[0]?.level).toBe(2);
		expect(guidedImagery?.children[0]?.locator).toMatch(/^p\d+$/);
	});

	test("doc with no headings returns an empty outline", async () => {
		const workspace = tempWorkspace("outline-empty");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Just a body paragraph.");
		const result = await runCli("outline", docPath);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toEqual([]);
	});

	test("--style-prefix targets a non-default style family", async () => {
		// academic-paper has a "Title" style on p3 — invisible under default Heading
		// prefix but should surface when we ask for it explicitly.
		const result = await runCli(
			"outline",
			"tests/fixtures/academic-paper.docx",
			"--style-prefix",
			"Title",
		);
		const payload = result.parsed as OutlineEntry[];
		expect(payload).toHaveLength(1);
		expect(payload[0]?.style).toBe("Title");
		expect(payload[0]?.level).toBe(1);
	});

	test("skipped levels nest directly under the nearest shallower level", async () => {
		const workspace = tempWorkspace("outline-skip");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "intro");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Top",
			"--style",
			"Heading1",
		);
		await runCli(
			"insert",
			docPath,
			"--after",
			"p1",
			"--text",
			"Skipped",
			"--style",
			"Heading3",
		);

		const result = await runCli("outline", docPath);
		const payload = result.parsed as OutlineEntry[];
		expect(payload).toHaveLength(1);
		expect(payload[0]?.text).toBe("Top");
		expect(payload[0]?.children).toHaveLength(1);
		expect(payload[0]?.children[0]).toMatchObject({
			text: "Skipped",
			level: 3,
		});
	});
});
