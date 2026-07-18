import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Document } from "@core/ast/document";
import { XmlNode } from "@core/parser";
import type { AbstractNumKind } from "../../src/core/ast/document/numbering";

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
		const document = await Document.open(target);
		expect(document.numbering?.tree).toBeUndefined();

		const numId = document.ensureNumbering().allocate("bullet");

		expect(numId).toBe(1);
		const tree = document.numbering?.tree;
		if (!tree) throw new Error("expected numberingTree");
		const abstractNums = abstractNumIds(tree);
		const nums = numIdsFromTree(tree);
		expect(abstractNums).toHaveLength(1);
		expect(nums).toEqual([1]);
		expect(level0Format(tree, abstractNums[0] ?? -1)).toBe("bullet");
		expect(relationshipTypes(document.relationships.tree)).toContain(
			"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
		);
		expect(overridePartNames(document.contentTypes.tree)).toContain(
			"/word/numbering.xml",
		);
	});

	test("reuses the existing abstractNum for repeated calls of the same kind", async () => {
		const target = await stageFixture("reuse.docx");
		const document = await Document.open(target);

		const a = document.ensureNumbering().allocate("bullet");
		const b = document.ensureNumbering().allocate("bullet");
		const c = document.ensureNumbering().allocate("bullet");

		expect([a, b, c]).toEqual([1, 2, 3]);
		const tree = document.numbering?.tree;
		if (!tree) throw new Error("expected numberingTree");
		expect(abstractNumIds(tree)).toHaveLength(1);
		expect(numIdsFromTree(tree)).toEqual([1, 2, 3]);
	});

	test("seeds a second abstractNum when a different kind is requested", async () => {
		const target = await stageFixture("two-kinds.docx");
		const document = await Document.open(target);

		const bullet = document.ensureNumbering().allocate("bullet");
		const ordered = document.ensureNumbering().allocate("ordered");

		expect(bullet).toBe(1);
		expect(ordered).toBe(2);
		const tree = document.numbering?.tree;
		if (!tree) throw new Error("expected numberingTree");
		const abstractNums = abstractNumIds(tree);
		expect(abstractNums).toHaveLength(2);
		const formats = abstractNums.map((id) => level0Format(tree, id));
		expect(formats).toEqual(expect.arrayContaining(["bullet", "decimal"]));
	});

	test("abstractNum elements precede num elements (per OOXML schema)", async () => {
		const target = await stageFixture("ordering.docx");
		const document = await Document.open(target);
		document.ensureNumbering().allocate("ordered");
		document.ensureNumbering().allocate("bullet");
		const tree = document.numbering?.tree;
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
		const document = await Document.open(target);
		const bullet = document.ensureNumbering().allocate("bullet");
		const ordered = document.ensureNumbering().allocate("ordered");
		await document.save();

		const reopened = await Document.open(target);
		const tree = reopened.numbering?.tree;
		if (!tree) throw new Error("expected reopened numberingTree");
		expect(numIdsFromTree(tree)).toEqual([bullet, ordered]);
		expect(abstractNumIds(tree)).toHaveLength(2);
	});

	test("doesn't collide with existing num/abstractNum ids in a pre-existing numbering.xml", async () => {
		const target = await stageFixture("pre-existing.docx");
		const document = await Document.open(target);
		// Seed a synthetic numbering.xml with high ids — allocate should
		// pick ids that don't collide.
		const numberingView = document.ensureNumbering();
		numberingView.tree = XmlNode.parse(
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="42"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:numFmt w:val="upperLetter"/></w:lvl></w:abstractNum>
<w:num w:numId="99"><w:abstractNumId w:val="42"/></w:num>
</w:numbering>`,
		);

		const newId = numberingView.allocate("bullet");

		expect(newId).toBe(100);
		const abstractNums = abstractNumIds(numberingView.tree);
		expect(abstractNums).toEqual(expect.arrayContaining([42, 43]));
		// Existing entries untouched.
		expect(numIdsFromTree(numberingView.tree)).toEqual(
			expect.arrayContaining([99, 100]),
		);
	});

	test("accepts both kinds via the AbstractNumKind union without surprises", async () => {
		const target = await stageFixture("type-cover.docx");
		const document = await Document.open(target);
		const kinds: AbstractNumKind[] = ["bullet", "ordered"];
		for (const kind of kinds) document.ensureNumbering().allocate(kind);
		const tree = document.numbering?.tree;
		if (!tree) throw new Error("expected numberingTree");
		expect(abstractNumIds(tree)).toHaveLength(2);
	});
});

describe("numbering control (start / format / clone)", () => {
	test("setStart updates the level-0 startOverride; getStart reflects it", async () => {
		const target = await stageFixture("set-start.docx");
		const document = await Document.open(target);
		const view = document.ensureNumbering();
		const numId = String(view.allocate("ordered"));

		expect(view.getStart(numId, 0)).toBe(1);
		expect(view.setStart(numId, 0, 5)).toBe(true);
		expect(view.getStart(numId, 0)).toBe(5);
		// Idempotent update, not duplicated.
		view.setStart(numId, 0, 7);
		expect(view.getStart(numId, 0)).toBe(7);
		expect(startOverrideCount(view.tree, numId, 0)).toBe(1);
	});

	test("setStart returns false for an unknown numId", async () => {
		const target = await stageFixture("set-start-missing.docx");
		const document = await Document.open(target);
		const view = document.ensureNumbering();
		view.allocate("ordered");
		expect(view.setStart("999", 0, 3)).toBe(false);
	});

	test("setFormat overrides numFmt per-list; getFormat prefers the override", async () => {
		const target = await stageFixture("set-format.docx");
		const document = await Document.open(target);
		const view = document.ensureNumbering();
		const numId = String(view.allocate("ordered"));

		expect(view.getFormat(numId, 0)).toBe("decimal");
		expect(view.setFormat(numId, 0, "upperRoman")).toBe(true);
		expect(view.getFormat(numId, 0)).toBe("upperRoman");
		// The override is a well-formed CT_Lvl that keeps its lvlText AND its
		// <w:start> — Word starts a start-less override level at 0, and a
		// roman/alpha glyph of 0 renders as an empty marker.
		const lvl = overrideLvl(view.tree, numId, 0);
		expect(lvl?.findChild("w:lvlText")?.getAttribute("w:val")).toBe("%1.");
		expect(lvl?.findChild("w:start")?.getAttribute("w:val")).toBe("1");
	});

	test("setStart + setFormat agree: the inner <w:start> stays in lockstep with startOverride", async () => {
		const target = await stageFixture("start-and-format.docx");
		const document = await Document.open(target);
		const view = document.ensureNumbering();
		const numId = String(view.allocate("ordered"));

		view.setStart(numId, 0, 5);
		view.setFormat(numId, 0, "upperRoman");

		expect(view.getStart(numId, 0)).toBe(5);
		expect(view.getFormat(numId, 0)).toBe("upperRoman");
		// setFormat found the existing startOverride and synced the clone's start.
		expect(
			overrideLvl(view.tree, numId, 0)
				?.findChild("w:start")
				?.getAttribute("w:val"),
		).toBe("5");
		expect(overrideChildTags(view.tree, numId, 0)).toEqual([
			"w:startOverride",
			"w:lvl",
		]);

		// And the reverse order: a later setStart re-syncs the inner start too.
		view.setStart(numId, 0, 9);
		expect(
			overrideLvl(view.tree, numId, 0)
				?.findChild("w:start")
				?.getAttribute("w:val"),
		).toBe("9");
	});

	test("a format override on one level upgrades bare startOverride siblings (Word drops mixed overrides)", async () => {
		// Mirrors the contract fixture that surfaced the bug: num 2 has a
		// pre-existing startOverride-only override on ilvl 0 (the 15 section
		// numbers) and the agent reformats ilvl 1 to lower-roman. Word blanks
		// every startOverride-only level as soon as a sibling override carries a
		// full <w:lvl>, so setFormat must upgrade the bare sibling with a cloned
		// level of its own — keeping ITS format and start.
		const target = await stageFixture("mixed-overrides.docx");
		const document = await Document.open(target);
		const view = document.ensureNumbering();
		view.tree = XmlNode.parse(
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
<w:abstractNum w:abstractNumId="2">
<w:lvl w:ilvl="0" w15:tentative="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
<w:lvl w:ilvl="1" w15:tentative="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%2)"/></w:lvl>
</w:abstractNum>
<w:num w:numId="2"><w:abstractNumId w:val="2"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>
</w:numbering>`,
		);

		expect(view.setFormat("2", 1, "lowerRoman")).toBe(true);

		const numberingRoot = view.tree
			? XmlNode.findRoot(view.tree, "w:numbering")
			: undefined;
		const num = numberingRoot
			?.findChildren("w:num")
			.find((node) => node.getAttribute("w:numId") === "2");
		const overrides = num?.findChildren("w:lvlOverride") ?? [];
		// Ascending ilvl order — the new ilvl-1 override lands AFTER ilvl 0.
		expect(overrides.map((node) => node.getAttribute("w:ilvl"))).toEqual([
			"0",
			"1",
		]);
		// Homogeneous: every override now carries a full <w:lvl>.
		for (const override of overrides) {
			const lvl = override.findChild("w:lvl");
			if (!lvl) throw new Error("expected an override <w:lvl>");
			expect(lvl.findChild("w:start")?.getAttribute("w:val")).toBe("1");
			// The clone never carries w15:tentative — the level is in use.
			expect(lvl.getAttribute("w15:tentative")).toBeUndefined();
		}
		// The upgraded ilvl-0 sibling keeps ITS OWN format; ilvl 1 got the new one.
		expect(view.getFormat("2", 0)).toBe("decimal");
		expect(view.getFormat("2", 1)).toBe("lowerRoman");
		// The bare override kept its startOverride (CT_LvlOverride order intact).
		expect(overrideChildTags(view.tree, "2", 0)).toEqual([
			"w:startOverride",
			"w:lvl",
		]);
	});

	test("getStart falls back to 1 on a non-numeric startOverride (foreign doc)", async () => {
		const target = await stageFixture("bad-start.docx");
		const document = await Document.open(target);
		const view = document.ensureNumbering();
		view.tree = XmlNode.parse(
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="oops"/></w:lvlOverride></w:num>
</w:numbering>`,
		);
		// NaN must not propagate into the start; fall back to 1.
		expect(view.getStart("1", 0)).toBe(1);
		expect(view.getLevelInfo("1", 0).start).toBe(1);
	});

	test("a second list keeps the shared abstractNum default when one is overridden", async () => {
		const target = await stageFixture("format-isolation.docx");
		const document = await Document.open(target);
		const view = document.ensureNumbering();
		const a = String(view.allocate("ordered"));
		const b = String(view.allocate("ordered"));

		view.setFormat(a, 0, "upperRoman");

		expect(view.getFormat(a, 0)).toBe("upperRoman");
		expect(view.getFormat(b, 0)).toBe("decimal");
	});

	test("CT_LvlOverride keeps startOverride before lvl (both insertion orders)", async () => {
		const target = await stageFixture("override-order.docx");
		const document = await Document.open(target);
		const view = document.ensureNumbering();
		const numId = String(view.allocate("ordered"));

		// allocate seeds startOverride; setFormat must append lvl AFTER it.
		view.setFormat(numId, 0, "lowerLetter");
		expect(overrideChildTags(view.tree, numId, 0)).toEqual([
			"w:startOverride",
			"w:lvl",
		]);
	});

	test("cloneListDefinition mints a new numId, copies the format, and restarts", async () => {
		const target = await stageFixture("clone.docx");
		const document = await Document.open(target);
		const view = document.ensureNumbering();
		const src = String(view.allocate("ordered"));
		view.setFormat(src, 0, "upperRoman");

		const cloned = String(view.cloneListDefinition(src, 3));

		expect(cloned).not.toBe(src);
		// Same backing abstractNum (the shared ordered style).
		expect(abstractNumIdOf(view.tree, cloned)).toBe(
			abstractNumIdOf(view.tree, src),
		);
		// Format carried, start applied, and the override is well-ordered.
		expect(view.getFormat(cloned, 0)).toBe("upperRoman");
		expect(view.getStart(cloned, 0)).toBe(3);
		expect(overrideChildTags(view.tree, cloned, 0)).toEqual([
			"w:startOverride",
			"w:lvl",
		]);
	});
});

function lvlOverrideNode(
	tree: XmlNode[],
	numId: string,
	level: number,
): XmlNode | undefined {
	const root = XmlNode.findRoot(tree, "w:numbering");
	const num = root
		?.findChildren("w:num")
		.find((node) => node.getAttribute("w:numId") === numId);
	return num
		?.findChildren("w:lvlOverride")
		.find((node) => node.getAttribute("w:ilvl") === String(level));
}

function overrideChildTags(
	tree: XmlNode[],
	numId: string,
	level: number,
): string[] {
	return (
		lvlOverrideNode(tree, numId, level)
			?.children.filter((child) => !child.isText)
			.map((child) => child.tag) ?? []
	);
}

function overrideLvl(
	tree: XmlNode[],
	numId: string,
	level: number,
): XmlNode | undefined {
	return lvlOverrideNode(tree, numId, level)?.findChild("w:lvl");
}

function startOverrideCount(
	tree: XmlNode[],
	numId: string,
	level: number,
): number {
	return (
		lvlOverrideNode(tree, numId, level)?.findChildren("w:startOverride")
			.length ?? 0
	);
}

function abstractNumIdOf(tree: XmlNode[], numId: string): string | undefined {
	const root = XmlNode.findRoot(tree, "w:numbering");
	const num = root
		?.findChildren("w:num")
		.find((node) => node.getAttribute("w:numId") === numId);
	return num?.findChild("w:abstractNumId")?.getAttribute("w:val");
}

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
