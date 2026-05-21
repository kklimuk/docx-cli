import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { openDocView, saveDocView } from "@core/ast/doc-view";
import { XmlNode } from "@core/parser";

let workspace: string;

beforeAll(() => {
	workspace = mkdtempSync(join(tmpdir(), "docx-cli-docview-"));
});

afterAll(() => {
	if (workspace) rmSync(workspace, { recursive: true, force: true });
});

describe("saveDocView writeback", () => {
	test("footnotesTree mutations survive a save/reopen cycle", async () => {
		const target = join(workspace, "notes-footnotes.docx");
		await Bun.write(target, Bun.file("tests/fixtures/notes.docx"));

		const view = await openDocView(target);
		const tree = view.footnotesTree;
		if (!tree) throw new Error("expected footnotesTree to be loaded");
		setNoteText(tree, "w:footnote", "1", " Edited footnote body.");
		expect(noteText(tree, "w:footnote", "1")).toBe(" Edited footnote body.");

		await saveDocView(view);

		const reopened = await openDocView(target);
		if (!reopened.footnotesTree) {
			throw new Error("expected reopened footnotesTree");
		}
		expect(noteText(reopened.footnotesTree, "w:footnote", "1")).toBe(
			" Edited footnote body.",
		);
		expect(reopened.doc.footnotes?.[0]?.text).toContain(
			"Edited footnote body.",
		);
	});

	test("endnotesTree mutations survive a save/reopen cycle", async () => {
		const target = join(workspace, "notes-endnotes.docx");
		await Bun.write(target, Bun.file("tests/fixtures/notes.docx"));

		const view = await openDocView(target);
		const tree = view.endnotesTree;
		if (!tree) throw new Error("expected endnotesTree to be loaded");
		setNoteText(tree, "w:endnote", "1", " Edited endnote body.");

		await saveDocView(view);

		const reopened = await openDocView(target);
		if (!reopened.endnotesTree) {
			throw new Error("expected reopened endnotesTree");
		}
		expect(noteText(reopened.endnotesTree, "w:endnote", "1")).toBe(
			" Edited endnote body.",
		);
		expect(reopened.doc.endnotes?.[0]?.text).toContain("Edited endnote body.");
	});

	test("appending a new <w:footnote> persists through save", async () => {
		const target = join(workspace, "notes-append.docx");
		await Bun.write(target, Bun.file("tests/fixtures/notes.docx"));

		const view = await openDocView(target);
		const tree = view.footnotesTree;
		if (!tree) throw new Error("expected footnotesTree to be loaded");
		const root = XmlNode.findRoot(tree, "w:footnotes");
		if (!root) throw new Error("expected <w:footnotes> root");

		const fresh = XmlNode.element("w:footnote", { "w:id": "2" }, [
			XmlNode.element("w:p", {}, [
				XmlNode.element("w:r", {}, [
					XmlNode.element("w:t", { "xml:space": "preserve" }, [
						XmlNode.textNode("Brand new footnote."),
					]),
				]),
			]),
		]);
		root.children.push(fresh);

		await saveDocView(view);

		const reopened = await openDocView(target);
		if (!reopened.footnotesTree) {
			throw new Error("expected reopened footnotesTree");
		}
		expect(noteText(reopened.footnotesTree, "w:footnote", "2")).toBe(
			"Brand new footnote.",
		);
	});
});

function findNoteById(tree: XmlNode[], tag: string, id: string): XmlNode {
	const rootTag = tag === "w:footnote" ? "w:footnotes" : "w:endnotes";
	const root = XmlNode.findRoot(tree, rootTag);
	if (!root) throw new Error(`missing <${rootTag}> root`);
	for (const child of root.findChildren(tag)) {
		if (child.getAttribute("w:id") === id) return child;
	}
	throw new Error(`<${tag} w:id="${id}"> not found`);
}

function noteText(tree: XmlNode[], tag: string, id: string): string {
	const node = findNoteById(tree, tag, id);
	const text = node.findDescendant("w:t");
	return text ? text.collectText() : "";
}

function setNoteText(
	tree: XmlNode[],
	tag: string,
	id: string,
	value: string,
): void {
	const node = findNoteById(tree, tag, id);
	const text = node.findDescendant("w:t");
	if (!text) throw new Error(`<w:t> missing from ${tag} ${id}`);
	const first = text.children[0];
	if (!first) throw new Error(`<w:t> for ${tag} ${id} has no children`);
	first.text = value;
}
