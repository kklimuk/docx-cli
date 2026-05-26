import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";
import { openDocView, saveDocView } from "../../../src/core/ast/doc-view";
import { Paragraph } from "../../../src/core/blocks";
import { w } from "../../../src/core/jsx";
import { allocateNum } from "../../../src/core/numbering";
import { type NullableXmlNode, XmlNode } from "../../../src/core/parser";
import { ensureStyle } from "../../../src/core/styles";

/**
 * Build tests/fixtures/task-lists-web.docx — exercises the OTHER GFM task-list
 * shape Word for Web emits when authoring via the Home → Checklist button:
 * a regular bulleted list whose level-0 bullet character is Wingdings ☐
 * (U+F0A8), with `<w:strike>` on the paragraph-mark to mark items as done.
 * No SDT content controls; no `<w14:checkbox>`. Word for Web silently strips
 * our SDT shape on author, so we MUST be able to read this format too — see
 * the empirical probes in /tmp/checkbox-track-probe/ for the data this was
 * derived from.
 *
 * Implementation note: we leverage `allocateNum(view, "bullet")` to provision
 * the numbering.xml part + register the relationship, then mutate the
 * resulting abstractNum's level-0 lvlText to be Wingdings ☐. We don't expose
 * Wingdings-bullet as a public AbstractNumKind because we don't emit this
 * shape on the write side (SDT is canonical for our own task-list output).
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/task-lists-web.docx");
const cliEntry = resolve(root, "src/index.ts");

async function cli(...args: string[]): Promise<void> {
	await $`bun ${cliEntry} ${args}`.quiet();
}

mkdirSync(dirname(out), { recursive: true });

await cli("create", out, "--force", "--text", "Web-Checklist fixture");

const view = await openDocView(out);
ensureStyle(view, "ListParagraph");
ensureStyle(view, "Heading2");

const numId = allocateNum(view, "bullet");
rewriteLvl0AsWingdingsChecklist(view, numId);

const documentRoot = XmlNode.findRoot(view.documentTree, "w:document");
if (!documentRoot) throw new Error("expected <w:document> root");
const body = documentRoot.findChild("w:body");
if (!body) throw new Error("expected <w:body> child");
const lastChild = body.children[body.children.length - 1];
if (!lastChild || lastChild.tag !== "w:sectPr") {
	throw new Error("expected last child of <w:body> to be <w:sectPr>");
}
const sectPr = lastChild;

const newChildren: XmlNode[] = [
	Paragraph({ text: "Web Checklist", style: "Heading2" }),
	webChecklistItem({ numId, text: "buy groceries", checked: false }),
	webChecklistItem({ numId, text: "pay rent", checked: true }),
	webChecklistItem({ numId, text: "call dentist", checked: false }),
	webChecklistItem({ numId, text: "nested done item", checked: true }),
];

body.children = [...newChildren, sectPr];

await saveDocView(view);

const bytes = (await Bun.file(out).bytes()).length;
console.log(`Wrote ${out} (${bytes} bytes)`);
console.log(`  Wingdings ☐ numId=${numId}`);

/** Materialize a list paragraph whose bullet is Wingdings ☐, with optional
 * paragraph-mark strike to encode "checked" — the exact shape Word for Web's
 * Checklist feature emits. The bullet comes from numbering.xml; we only need
 * to wire numPr + the strike toggle here. */
function webChecklistItem({
	numId,
	text,
	checked,
}: {
	numId: number;
	text: string;
	checked: boolean;
}): XmlNode {
	const strikeRpr: NullableXmlNode = checked ? (
		<w.rPr>
			<w.strike />
		</w.rPr>
	) : null;
	return (
		<w.p>
			<w.pPr>
				<w.pStyle w-val="ListParagraph" />
				<w.numPr>
					<w.ilvl w-val="0" />
					<w.numId w-val={String(numId)} />
				</w.numPr>
				{strikeRpr}
			</w.pPr>
			<w.r>
				{checked && (
					<w.rPr>
						<w.strike />
					</w.rPr>
				)}
				<w.t {...{ "xml:space": "preserve" }}>{text}</w.t>
			</w.r>
		</w.p>
	);
}

/** Mutate the abstractNum referenced by `numId` so its level-0 bullet glyph
 * is Wingdings ☐ (U+F0A8) with the Wingdings font — the canonical Word-for-Web
 * Checklist bullet. Operates on the in-memory numberingTree; saveDocView will
 * persist. */
function rewriteLvl0AsWingdingsChecklist(
	v: Awaited<ReturnType<typeof openDocView>>,
	numId: number,
): void {
	const tree = v.numberingTree;
	if (!tree) throw new Error("numberingTree should exist after allocateNum");
	const numberingRoot = XmlNode.findRoot(tree, "w:numbering");
	if (!numberingRoot) throw new Error("expected <w:numbering> root");
	const num = numberingRoot
		.findChildren("w:num")
		.find((n) => n.getAttribute("w:numId") === String(numId));
	if (!num) throw new Error(`expected <w:num w:numId="${numId}">`);
	const abstractNumId = num.findChild("w:abstractNumId")?.getAttribute("w:val");
	if (!abstractNumId) throw new Error("missing abstractNumId");
	const abstractNum = numberingRoot
		.findChildren("w:abstractNum")
		.find((a) => a.getAttribute("w:abstractNumId") === abstractNumId);
	if (!abstractNum) throw new Error("missing abstractNum");
	const lvl0 = abstractNum
		.findChildren("w:lvl")
		.find((l) => l.getAttribute("w:ilvl") === "0");
	if (!lvl0) throw new Error("missing level 0");
	const lvlText = lvl0.findChild("w:lvlText");
	// Wingdings code point for ☐. The detection in core/ast/read.ts accepts both
	// U+F0A8 (Wingdings) and U+2610 (Unicode ☐). Use  escape so the PUA
	// character survives file-write round-trips (some editors strip raw PUA).
	if (lvlText) lvlText.setAttribute("w:val", "");
	const rPr = lvl0.findChild("w:rPr") ?? new XmlNode("w:rPr");
	if (!lvl0.findChild("w:rPr")) lvl0.children.push(rPr);
	rPr.children = rPr.children.filter((c) => c.tag !== "w:rFonts");
	rPr.children.unshift(
		new XmlNode("w:rFonts", {
			"w:ascii": "Wingdings",
			"w:hAnsi": "Wingdings",
		}),
	);
}
