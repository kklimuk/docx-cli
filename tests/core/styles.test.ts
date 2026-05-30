import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Document } from "@core/ast/document";
import { XmlNode } from "@core/parser";
import type { BaselineStyleId } from "../../src/core/ast/document/styles";

let workspace: string;

beforeAll(() => {
	workspace = mkdtempSync(join(tmpdir(), "docx-cli-styles-"));
});

afterAll(() => {
	if (workspace) rmSync(workspace, { recursive: true, force: true });
});

describe("ensureStyle", () => {
	test("fixture ships canonical parts with Normal as the only defined style", async () => {
		const document = await Document.open(
			"tests/fixtures/styles-injection.docx",
		);
		expect(document.styles?.tree).toBeDefined();
		const tree = document.styles?.tree;
		if (!tree) throw new Error("expected stylesTree");
		expect(styleIds(tree)).toEqual(["Normal"]);
		// The styles part should already be registered in the canonical fixture.
		expect(relationshipTypes(document.relationships.tree)).toContain(
			"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
		);
		expect(overridePartNames(document.contentTypes.tree)).toContain(
			"/word/styles.xml",
		);
	});

	test("adds a missing style definition (Heading1) to existing styles.xml", async () => {
		const target = await stageFixture("add-missing.docx");
		const document = await Document.open(target);

		document.ensureStyles().ensureStyle("Heading1");

		const tree = document.styles?.tree;
		if (!tree) throw new Error("expected stylesTree after ensureStyle");
		const ids = styleIds(tree);
		expect(ids).toContain("Heading1");
		expect(ids).toContain("Normal");
	});

	test("seeds the styles part from scratch when the package lacks one", async () => {
		const target = await stageFixture("from-scratch.docx");
		const document = await Document.open(target);

		// Simulate an older / hand-rolled doc that omits styles.xml entirely.
		document.pkg.deletePart("word/styles.xml");
		document.styles = undefined;

		const stylesView = document.ensureStyles();
		stylesView.ensureStyle("Heading1");
		const ids = styleIds(stylesView.tree);
		expect(ids).toContain("Heading1");
		expect(ids).toContain("Normal");
		// The relationship + content-type override pre-existed in the canonical
		// fixture; registerPart is idempotent and should leave them as-is rather
		// than duplicating.
		const styleRelCount = relationshipTypes(document.relationships.tree).filter(
			(type) =>
				type ===
				"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
		).length;
		expect(styleRelCount).toBe(1);
	});

	test("registers the relationship + content-type override when the package lacks them entirely", async () => {
		const target = await stageFixture("rel-from-scratch.docx");
		const document = await Document.open(target);

		// Strip the part AND its rel/override entries — simulates a doc that
		// truly never had styles.xml registered (vs. the prior test which only
		// removed the part bytes). Exercises the mintRelationshipId + push path
		// inside registerPart.
		document.pkg.deletePart("word/styles.xml");
		document.styles = undefined;
		removeRelationshipByType(
			document.relationships.tree,
			"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
		);
		removeOverride(document.contentTypes.tree, "/word/styles.xml");

		const stylesView = document.ensureStyles();
		stylesView.ensureStyle("Heading1");

		const styleRels = relationshipTypes(document.relationships.tree).filter(
			(type) =>
				type ===
				"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
		);
		expect(styleRels).toHaveLength(1);
		expect(overridePartNames(document.contentTypes.tree)).toContain(
			"/word/styles.xml",
		);
		expect(styleIds(stylesView.tree)).toContain("Heading1");
	});

	test("IntenseQuote emits pPr children in OOXML schema order (pBdr before spacing/ind)", async () => {
		const target = await stageFixture("intense-quote-order.docx");
		const document = await Document.open(target);
		document.ensureStyles().ensureStyle("IntenseQuote");
		const tree = document.styles?.tree;
		if (!tree) throw new Error("expected stylesTree");
		const root = XmlNode.findRoot(tree, "w:styles");
		if (!root) throw new Error("expected <w:styles> root");
		const style = root
			.findChildren("w:style")
			.find((s) => s.getAttribute("w:styleId") === "IntenseQuote");
		if (!style) throw new Error("expected IntenseQuote style");
		const pPr = style.findChild("w:pPr");
		if (!pPr) throw new Error("expected pPr child");
		// Per ECMA-376 §17.3.1.26 (CT_PPrBase), pBdr precedes spacing and ind.
		const childTags = pPr.children.map((c) => c.tag);
		const pBdrIdx = childTags.indexOf("w:pBdr");
		const spacingIdx = childTags.indexOf("w:spacing");
		const indIdx = childTags.indexOf("w:ind");
		expect(pBdrIdx).toBeGreaterThanOrEqual(0);
		expect(pBdrIdx).toBeLessThan(spacingIdx);
		expect(pBdrIdx).toBeLessThan(indIdx);
	});

	test("is idempotent — calling twice does not duplicate definitions", async () => {
		const target = await stageFixture("idempotent.docx");
		const document = await Document.open(target);

		document.ensureStyles().ensureStyle("Heading2");
		document.ensureStyles().ensureStyle("Heading2");
		document.ensureStyles().ensureStyle("Heading2");

		const tree = document.styles?.tree;
		if (!tree) throw new Error("expected stylesTree after ensureStyle");
		const count = styleIds(tree).filter((id) => id === "Heading2").length;
		expect(count).toBe(1);
	});

	test("injected styles survive a save/reopen cycle", async () => {
		const target = await stageFixture("roundtrip.docx");
		const document = await Document.open(target);
		document.ensureStyles().ensureStyle("Heading1");
		document.ensureStyles().ensureStyle("Quote");
		document.ensureStyles().ensureStyle("Code");
		document.ensureStyles().ensureStyle("FootnoteReference");
		await document.save();

		const reopened = await Document.open(target);
		const tree = reopened.styles?.tree;
		if (!tree) throw new Error("expected reopened stylesTree");
		const ids = styleIds(tree);
		expect(ids).toEqual(
			expect.arrayContaining([
				"Normal",
				"Heading1",
				"Quote",
				"Code",
				"FootnoteReference",
			]),
		);
	});

	test("covers the full baseline catalog without error", async () => {
		const target = await stageFixture("baseline.docx");
		const document = await Document.open(target);
		const all: BaselineStyleId[] = [
			"Normal",
			"Heading1",
			"Heading2",
			"Heading3",
			"Heading4",
			"Heading5",
			"Heading6",
			"Quote",
			"IntenseQuote",
			"Code",
			"CodeBlock",
			"ListParagraph",
			"FootnoteReference",
			"FootnoteText",
		];
		for (const id of all) document.ensureStyles().ensureStyle(id);
		const tree = document.styles?.tree;
		if (!tree) throw new Error("expected stylesTree after ensureStyle");
		expect(styleIds(tree)).toEqual(expect.arrayContaining(all));
	});
});

async function stageFixture(label: string): Promise<string> {
	const target = join(workspace, label);
	await Bun.write(target, Bun.file("tests/fixtures/styles-injection.docx"));
	return target;
}

function styleIds(tree: XmlNode[]): string[] {
	const root = XmlNode.findRoot(tree, "w:styles");
	if (!root) return [];
	return root
		.findChildren("w:style")
		.map((node) => node.getAttribute("w:styleId") ?? "");
}

function relationshipTypes(tree: XmlNode[]): string[] {
	const root = XmlNode.findRoot(tree, "Relationships");
	if (!root) return [];
	return root
		.findChildren("Relationship")
		.map((node) => node.getAttribute("Type") ?? "");
}

function overridePartNames(tree: XmlNode[]): string[] {
	const root = XmlNode.findRoot(tree, "Types");
	if (!root) return [];
	return root
		.findChildren("Override")
		.map((node) => node.getAttribute("PartName") ?? "");
}

function removeRelationshipByType(tree: XmlNode[], type: string): void {
	const root = XmlNode.findRoot(tree, "Relationships");
	if (!root) return;
	root.children = root.children.filter(
		(child) =>
			!(child.tag === "Relationship" && child.getAttribute("Type") === type),
	);
}

function removeOverride(tree: XmlNode[], partName: string): void {
	const root = XmlNode.findRoot(tree, "Types");
	if (!root) return;
	root.children = root.children.filter(
		(child) =>
			!(
				child.tag === "Override" && child.getAttribute("PartName") === partName
			),
	);
}
