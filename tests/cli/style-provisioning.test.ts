import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "@core/package";
import { runCli, tempWorkspace } from "./harness";

async function styleIds(docPath: string): Promise<string[]> {
	const pkg = await Pkg.open(docPath);
	if (!pkg.hasPart("word/styles.xml")) return [];
	const xml = await pkg.readText("word/styles.xml");
	return [...xml.matchAll(/w:styleId="([^"]+)"/g)].map((m) => m[1] ?? "");
}

describe("insert/edit --style provisioning", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("style-prov");
		docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Body");
	});

	test("insert --style Heading2 defines Heading2 (and Normal) in styles.xml", async () => {
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"A heading",
			"--style",
			"Heading2",
		);
		const ids = await styleIds(docPath);
		expect(ids).toContain("Heading2");
		expect(ids).toContain("Normal");
	});

	test("edit --style Quote defines Quote without dropping existing styles", async () => {
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"A heading",
			"--style",
			"Heading2",
		);
		await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--text",
			"Quoted",
			"--style",
			"Quote",
		);
		const ids = await styleIds(docPath);
		expect(ids).toContain("Quote");
		expect(ids).toContain("Heading2");
	});

	test("a custom (non-baseline) style is referenced but not defined", async () => {
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Custom",
			"--style",
			"MyCorporateStyle",
		);
		const ids = await styleIds(docPath);
		expect(ids).not.toContain("MyCorporateStyle");
		// The pStyle reference is still written even though the style is undefined.
		const pkg = await Pkg.open(docPath);
		const documentXml = await pkg.readText("word/document.xml");
		expect(documentXml).toContain('w:val="MyCorporateStyle"');
	});

	test("insert without --style adds no style definitions", async () => {
		await runCli("insert", docPath, "--after", "p0", "--text", "Plain");
		// `docx create` ships a styles.xml with only Normal; a plain insert
		// shouldn't add anything.
		expect(await styleIds(docPath)).toEqual(["Normal"]);
	});
});
