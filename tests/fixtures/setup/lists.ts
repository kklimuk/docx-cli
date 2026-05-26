import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";
import { openDocView, saveDocView } from "../../../src/core/ast/doc-view";
import { ListParagraph, Paragraph } from "../../../src/core/blocks";
import { allocateNum } from "../../../src/core/numbering";
import { XmlNode } from "../../../src/core/parser";
import { ensureStyle } from "../../../src/core/styles";

/**
 * Build tests/fixtures/lists.docx — a docx exercising bulleted and ordered
 * lists, including 3 levels of nesting per CLAUDE.md S3 verification plan.
 *
 * Dogfoods the runtime API:
 *  1. `docx create` for the Word-canonical base
 *  2. allocateNum(view, kind) to mint a numId per top-level list
 *  3. ensureStyle(view, "ListParagraph") so list items inherit the right style
 *  4. <ListParagraph> JSX emitter to construct each item, spliced into <w:body>
 *
 * Once S8 lands, the markdown walker will produce equivalent output for
 * `- foo` / `1. bar` markdown input — this fixture remains useful for
 * exercising the read path and verifying renderer behavior in Word.
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/lists.docx");
const cliEntry = resolve(root, "src/index.ts");

async function cli(...args: string[]): Promise<void> {
	await $`bun ${cliEntry} ${args}`.quiet();
}

mkdirSync(dirname(out), { recursive: true });

// 1. Start from a Word-canonical base.
await cli("create", out, "--force", "--text", "Lists fixture");

// 2. Open the base, allocate a numId per top-level list, seed the style.
const view = await openDocView(out);
const bulletNumId = allocateNum(view, "bullet");
const orderedNumId = allocateNum(view, "ordered");
const nestedBulletNumId = allocateNum(view, "bullet");
ensureStyle(view, "ListParagraph");

// 3. Replace the body's single seed paragraph with a richer list demo.
const documentRoot = XmlNode.findRoot(view.documentTree, "w:document");
if (!documentRoot) throw new Error("expected <w:document> root");
const body = documentRoot.findChild("w:body");
if (!body) throw new Error("expected <w:body> child");

// Preserve the trailing <w:sectPr> that `docx create` seeded; everything else
// gets replaced with our list demo. Check the LAST child specifically — the
// invariant is "trailing", and using `find` would silently skip inline sectPrs
// (which docx create doesn't emit today but might in the future).
const lastChild = body.children[body.children.length - 1];
if (!lastChild || lastChild.tag !== "w:sectPr") {
	throw new Error("expected last child of <w:body> to be <w:sectPr>");
}
const sectPr = lastChild;

const newChildren: XmlNode[] = [];
const heading = (text: string) =>
	newChildren.push(Paragraph({ text, style: "Heading2" }));
const bullet = (level: number, text: string, numId = bulletNumId) =>
	newChildren.push(ListParagraph({ numId, level, text }));
const ordered = (level: number, text: string) =>
	newChildren.push(ListParagraph({ numId: orderedNumId, level, text }));

ensureStyle(view, "Heading2");

heading("Bulleted list");
bullet(0, "Apples");
bullet(1, "Granny Smith");
bullet(1, "Honeycrisp");
bullet(2, "From the orchard down the road");
bullet(0, "Bananas");
bullet(0, "Cherries");

heading("Ordered list");
ordered(0, "Preheat the oven");
ordered(0, "Combine the dry ingredients");
ordered(1, "Sift the flour");
ordered(1, "Whisk in baking powder");
ordered(2, "Don't overmix");
ordered(0, "Bake for forty minutes");

heading("Independent second bulleted list (restarts)");
bullet(0, "Independent item one", nestedBulletNumId);
bullet(0, "Independent item two", nestedBulletNumId);
bullet(1, "And a nested item beneath it", nestedBulletNumId);

// Splice the new children in front of the trailing sectPr.
body.children = [...newChildren, sectPr];

await saveDocView(view);

const bytes = (await Bun.file(out).bytes()).length;
console.log(`Wrote ${out} (${bytes} bytes)`);
console.log(`  bullet numId=${bulletNumId}`);
console.log(`  ordered numId=${orderedNumId}`);
console.log(`  second bullet numId=${nestedBulletNumId} (restarts numbering)`);
