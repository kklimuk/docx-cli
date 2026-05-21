import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { openDocView, saveDocView } from "@core/ast/doc-view";
import { XmlNode } from "@core/parser";
import { type AbstractNumKind, allocateNum } from "../../src/core/numbering";

let workspace: string;

beforeAll(() => {
	workspace = mkdtempSync(join(tmpdir(), "docx-cli-numbering-"));
});

afterAll(() => {
	if (workspace) rmSync(workspace, { recursive: true, force: true });
});

describe("allocateNum", () => {
	test("creates numbering.xml from scratch on first call, with bullet abstractNum and one num", async () => {
		const target = await stageFixture("from-scratch.docx");
		const view = await openDocView(target);
		expect(view.numberingTree).toBeUndefined();

		const numId = allocateNum(view, "bullet");

		expect(numId).toBe(1);
		const tree = view.numberingTree;
		if (!tree) throw new Error("expected numberingTree");
		const abstractNums = abstractNumIds(tree);
		const nums = numIdsFromTree(tree);
		expect(abstractNums).toHaveLength(1);
		expect(nums).toEqual([1]);
		expect(level0Format(tree, abstractNums[0] ?? -1)).toBe("bullet");
		expect(relationshipTypes(view.relationshipsTree)).toContain(
			"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
		);
		expect(overridePartNames(view.contentTypesTree)).toContain(
			"/word/numbering.xml",
		);
	});

	test("reuses the existing abstractNum for repeated calls of the same kind", async () => {
		const target = await stageFixture("reuse.docx");
		const view = await openDocView(target);

		const a = allocateNum(view, "bullet");
		const b = allocateNum(view, "bullet");
		const c = allocateNum(view, "bullet");

		expect([a, b, c]).toEqual([1, 2, 3]);
		const tree = view.numberingTree;
		if (!tree) throw new Error("expected numberingTree");
		expect(abstractNumIds(tree)).toHaveLength(1);
		expect(numIdsFromTree(tree)).toEqual([1, 2, 3]);
	});

	test("seeds a second abstractNum when a different kind is requested", async () => {
		const target = await stageFixture("two-kinds.docx");
		const view = await openDocView(target);

		const bullet = allocateNum(view, "bullet");
		const ordered = allocateNum(view, "ordered");

		expect(bullet).toBe(1);
		expect(ordered).toBe(2);
		const tree = view.numberingTree;
		if (!tree) throw new Error("expected numberingTree");
		const abstractNums = abstractNumIds(tree);
		expect(abstractNums).toHaveLength(2);
		const formats = abstractNums.map((id) => level0Format(tree, id));
		expect(formats).toEqual(expect.arrayContaining(["bullet", "decimal"]));
	});

	test("abstractNum elements precede num elements (per OOXML schema)", async () => {
		const target = await stageFixture("ordering.docx");
		const view = await openDocView(target);
		allocateNum(view, "ordered");
		allocateNum(view, "bullet");
		const tree = view.numberingTree;
		if (!tree) throw new Error("expected numberingTree");
		const root = XmlNode.findRoot(tree, "w:numbering");
		if (!root) throw new Error("expected <w:numbering> root");
		const tags = root.children.map((child) => child.tag);
		const lastAbstract = tags.lastIndexOf("w:abstractNum");
		const firstNum = tags.indexOf("w:num");
		expect(lastAbstract).toBeGreaterThanOrEqual(0);
		expect(firstNum).toBeGreaterThan(lastAbstract);
	});

	test("survives a save/reopen cycle", async () => {
		const target = await stageFixture("roundtrip.docx");
		const view = await openDocView(target);
		const bullet = allocateNum(view, "bullet");
		const ordered = allocateNum(view, "ordered");
		await saveDocView(view);

		const reopened = await openDocView(target);
		const tree = reopened.numberingTree;
		if (!tree) throw new Error("expected reopened numberingTree");
		expect(numIdsFromTree(tree)).toEqual([bullet, ordered]);
		expect(abstractNumIds(tree)).toHaveLength(2);
	});

	test("doesn't collide with existing num/abstractNum ids in a pre-existing numbering.xml", async () => {
		const target = await stageFixture("pre-existing.docx");
		const view = await openDocView(target);
		// Seed a synthetic numbering.xml with high ids — allocateNum should
		// pick ids that don't collide.
		view.numberingTree = XmlNode.parse(
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="42"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:numFmt w:val="upperLetter"/></w:lvl></w:abstractNum>
<w:num w:numId="99"><w:abstractNumId w:val="42"/></w:num>
</w:numbering>`,
		);

		const newId = allocateNum(view, "bullet");

		expect(newId).toBe(100);
		const tree = view.numberingTree;
		const abstractNums = abstractNumIds(tree);
		expect(abstractNums).toEqual(expect.arrayContaining([42, 43]));
		// Existing entries untouched.
		expect(numIdsFromTree(tree)).toEqual(expect.arrayContaining([99, 100]));
	});

	test("accepts both kinds via the AbstractNumKind union without surprises", async () => {
		const target = await stageFixture("type-cover.docx");
		const view = await openDocView(target);
		const kinds: AbstractNumKind[] = ["bullet", "ordered"];
		for (const kind of kinds) allocateNum(view, kind);
		const tree = view.numberingTree;
		if (!tree) throw new Error("expected numberingTree");
		expect(abstractNumIds(tree)).toHaveLength(2);
	});
});

async function stageFixture(label: string): Promise<string> {
	const target = join(workspace, label);
	await Bun.write(target, Bun.file("tests/fixtures/styles-injection.docx"));
	return target;
}

function abstractNumIds(tree: XmlNode[]): number[] {
	const root = XmlNode.findRoot(tree, "w:numbering");
	if (!root) return [];
	return root
		.findChildren("w:abstractNum")
		.map((child) => Number(child.getAttribute("w:abstractNumId") ?? "-1"));
}

function numIdsFromTree(tree: XmlNode[]): number[] {
	const root = XmlNode.findRoot(tree, "w:numbering");
	if (!root) return [];
	return root
		.findChildren("w:num")
		.map((child) => Number(child.getAttribute("w:numId") ?? "-1"));
}

function level0Format(tree: XmlNode[], abstractNumId: number): string | null {
	const root = XmlNode.findRoot(tree, "w:numbering");
	if (!root) return null;
	for (const child of root.findChildren("w:abstractNum")) {
		if (Number(child.getAttribute("w:abstractNumId") ?? "") !== abstractNumId) {
			continue;
		}
		const lvl0 = child.findChild("w:lvl");
		const numFmt = lvl0?.findChild("w:numFmt");
		return numFmt?.getAttribute("w:val") ?? null;
	}
	return null;
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
