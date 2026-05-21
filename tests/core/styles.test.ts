import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { openDocView, saveDocView } from "@core/ast/doc-view";
import { XmlNode } from "@core/parser";
import { type BaselineStyleId, ensureStyle } from "../../src/core/styles";

let workspace: string;

beforeAll(() => {
	workspace = mkdtempSync(join(tmpdir(), "docx-cli-styles-"));
});

afterAll(() => {
	if (workspace) rmSync(workspace, { recursive: true, force: true });
});

describe("ensureStyle", () => {
	test("fixture ships canonical parts with Normal as the only defined style", async () => {
		const view = await openDocView("tests/fixtures/styles-injection.docx");
		expect(view.stylesTree).toBeDefined();
		const tree = view.stylesTree;
		if (!tree) throw new Error("expected stylesTree");
		expect(styleIds(tree)).toEqual(["Normal"]);
		// The styles part should already be registered in the canonical fixture.
		expect(relationshipTypes(view.relationshipsTree)).toContain(
			"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
		);
		expect(overridePartNames(view.contentTypesTree)).toContain(
			"/word/styles.xml",
		);
	});

	test("adds a missing style definition (Heading1) to existing styles.xml", async () => {
		const target = await stageFixture("add-missing.docx");
		const view = await openDocView(target);

		ensureStyle(view, "Heading1");

		const tree = view.stylesTree;
		if (!tree) throw new Error("expected stylesTree after ensureStyle");
		const ids = styleIds(tree);
		expect(ids).toContain("Heading1");
		expect(ids).toContain("Normal");
	});

	test("seeds the styles part from scratch when the package lacks one", async () => {
		const target = await stageFixture("from-scratch.docx");
		const view = await openDocView(target);

		// Simulate an older / hand-rolled doc that omits styles.xml entirely.
		view.pkg.deletePart("word/styles.xml");
		view.stylesTree = undefined;

		ensureStyle(view, "Heading1");

		const tree = view.stylesTree;
		if (!tree) throw new Error("expected stylesTree after ensureStyle");
		const ids = styleIds(tree);
		expect(ids).toContain("Heading1");
		expect(ids).toContain("Normal");
		// The relationship + content-type override pre-existed in the canonical
		// fixture; registerPart is idempotent and should leave them as-is rather
		// than duplicating.
		const styleRelCount = relationshipTypes(view.relationshipsTree).filter(
			(type) =>
				type ===
				"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
		).length;
		expect(styleRelCount).toBe(1);
	});

	test("registers the relationship + content-type override when the package lacks them entirely", async () => {
		const target = await stageFixture("rel-from-scratch.docx");
		const view = await openDocView(target);

		// Strip the part AND its rel/override entries — simulates a doc that
		// truly never had styles.xml registered (vs. the prior test which only
		// removed the part bytes). Exercises the mintRelationshipId + push path
		// inside registerPart.
		view.pkg.deletePart("word/styles.xml");
		view.stylesTree = undefined;
		removeRelationshipByType(
			view.relationshipsTree,
			"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
		);
		removeOverride(view.contentTypesTree, "/word/styles.xml");

		ensureStyle(view, "Heading1");

		const styleRels = relationshipTypes(view.relationshipsTree).filter(
			(type) =>
				type ===
				"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
		);
		expect(styleRels).toHaveLength(1);
		expect(overridePartNames(view.contentTypesTree)).toContain(
			"/word/styles.xml",
		);
		const tree = view.stylesTree;
		if (!tree) throw new Error("expected stylesTree after ensureStyle");
		expect(styleIds(tree)).toContain("Heading1");
	});

	test("IntenseQuote emits pPr children in OOXML schema order (pBdr before spacing/ind)", async () => {
		const target = await stageFixture("intense-quote-order.docx");
		const view = await openDocView(target);
		ensureStyle(view, "IntenseQuote");
		const tree = view.stylesTree;
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
		const view = await openDocView(target);

		ensureStyle(view, "Heading2");
		ensureStyle(view, "Heading2");
		ensureStyle(view, "Heading2");

		const tree = view.stylesTree;
		if (!tree) throw new Error("expected stylesTree after ensureStyle");
		const count = styleIds(tree).filter((id) => id === "Heading2").length;
		expect(count).toBe(1);
	});

	test("injected styles survive a save/reopen cycle", async () => {
		const target = await stageFixture("roundtrip.docx");
		const view = await openDocView(target);
		ensureStyle(view, "Heading1");
		ensureStyle(view, "Quote");
		ensureStyle(view, "Code");
		ensureStyle(view, "FootnoteReference");
		await saveDocView(view);

		const reopened = await openDocView(target);
		const tree = reopened.stylesTree;
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
		const view = await openDocView(target);
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
		for (const id of all) ensureStyle(view, id);
		const tree = view.stylesTree;
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
