import { XmlNode } from "../parser";
import type { TrackedMeta } from "../track-changes";
import { findCheckboxToggle } from "./detect";

/** Apply Word's accept-toggle behavior: keep the new glyph (unwrap `<w:ins>`),
 *  drop the deleted glyph (`<w:del>` and its contents go away). The
 *  `w14:checked` attribute is already in the "after" state (Word flipped it
 *  in place when the toggle happened), so no attribute fix is needed. */
export function acceptCheckboxToggle(sdt: XmlNode): void {
	const toggle = findCheckboxToggle(sdt);
	if (!toggle) return;
	const sdtContent = sdt.findChild("w:sdtContent");
	if (!sdtContent) return;
	unwrap(sdtContent.children, toggle.ins);
	remove(sdtContent.children, toggle.del);
}

/** Apply Word's reject-toggle behavior: drop the new glyph (`<w:ins>` and
 *  contents), restore the old glyph (unwrap `<w:del>` and rename its
 *  `<w:delText>` back to `<w:t>`), AND flip `w14:checked` back to match the
 *  restored glyph. We infer the prior attribute value from the deleted glyph
 *  (☐ → "0", ☒ → "1") because Word doesn't store it separately. */
export function rejectCheckboxToggle(sdt: XmlNode): void {
	const toggle = findCheckboxToggle(sdt);
	if (!toggle) return;
	const sdtContent = sdt.findChild("w:sdtContent");
	if (!sdtContent) return;
	remove(sdtContent.children, toggle.ins);
	unwrap(sdtContent.children, toggle.del);
	// Rename <w:delText> back to <w:t> so the restored glyph is a normal run.
	renameDelTextToText(toggle.del);
	const checkbox = sdt.findChild("w:sdtPr")?.findChild("w14:checkbox");
	const checkedNode = checkbox?.findChild("w14:checked");
	if (checkedNode) {
		// Infer prior state from the kept glyph: ☐ = unchecked (val=0), ☒ = checked (val=1).
		checkedNode.setAttribute(
			"w14:val",
			isCheckedGlyph(toggle.delGlyph) ? "1" : "0",
		);
	}
}

/** Flip a task-list paragraph's checkbox state in place (no tracking). Updates
 *  `w14:checked` and the visible glyph in `<w:sdtContent>`. Caller must verify
 *  the paragraph has a leading checkbox SDT via `detectTaskListState` or by
 *  inspecting `paragraph.taskState`. Returns true if the SDT was found and
 *  flipped, false otherwise (no-op). */
export function flipCheckboxUntracked(
	paragraph: XmlNode,
	checked: boolean,
): boolean {
	const sdt = paragraph.children.find((child) => child.tag === "w:sdt");
	if (!sdt) return false;
	const checkbox = sdt.findChild("w:sdtPr")?.findChild("w14:checkbox");
	const checkedNode = checkbox?.findChild("w14:checked");
	if (!checkedNode) return false;
	checkedNode.setAttribute("w14:val", checked ? "1" : "0");
	const sdtContent = sdt.findChild("w:sdtContent");
	if (sdtContent) rewriteGlyph(sdtContent, checked ? "☒" : "☐");
	return true;
}

/** Emit Word's canonical tracked-toggle XML inside the task-list paragraph's
 *  SDT: replace the existing glyph run with `<w:ins>new</w:ins><w:del>old</w:del>`
 *  and flip the `w14:checked` attribute. This mirrors what Microsoft Word for
 *  Mac emits when the user clicks a checkbox under tracking, validated
 *  empirically against the probe in /tmp/checkbox-track-probe/. The ins and
 *  del each consume a fresh revision id from `mintMeta` — Word emits distinct
 *  ids for the two halves of a toggle. Returns false if the paragraph has no
 *  leading checkbox SDT or the toggle is a no-op (already in the target state). */
export function flipCheckboxTracked(
	paragraph: XmlNode,
	checked: boolean,
	mintMeta: () => TrackedMeta,
): boolean {
	const sdt = paragraph.children.find((child) => child.tag === "w:sdt");
	if (!sdt) return false;
	const checkbox = sdt.findChild("w:sdtPr")?.findChild("w14:checkbox");
	const checkedNode = checkbox?.findChild("w14:checked");
	if (!checkedNode) return false;
	const currentlyChecked = checkedNode.getAttribute("w14:val") === "1";
	if (currentlyChecked === checked) return false;
	const sdtContent = sdt.findChild("w:sdtContent");
	if (!sdtContent) return false;
	const newGlyph = checked ? "☒" : "☐";
	const oldGlyph = checked ? "☐" : "☒";
	checkedNode.setAttribute("w14:val", checked ? "1" : "0");
	sdtContent.children = [
		buildGlyphWrapper("w:ins", newGlyph, mintMeta(), "w:t"),
		buildGlyphWrapper("w:del", oldGlyph, mintMeta(), "w:delText"),
	];
	return true;
}

function isCheckedGlyph(text: string): boolean {
	// Word's modern checklist canonically uses ☒ (U+2612) for checked in
	// `<w:sdtContent>`. The Wingdings PUA variants only appear in numbering
	// bullets (Word for Web), not SDT content.
	return text.includes("☒");
}

function unwrap(siblings: XmlNode[], wrapper: XmlNode): void {
	const index = siblings.indexOf(wrapper);
	if (index === -1) return;
	siblings.splice(index, 1, ...wrapper.children);
}

function remove(siblings: XmlNode[], target: XmlNode): void {
	const index = siblings.indexOf(target);
	if (index === -1) return;
	siblings.splice(index, 1);
}

function renameDelTextToText(wrapper: XmlNode): void {
	// `wrapper` has already been unwrapped, so its children now live in the
	// parent. Walk them and rename any <w:delText> → <w:t>. This works whether
	// they're still parented under `wrapper` or already moved — collectChildren
	// is referencing the same XmlNode objects.
	for (const run of wrapper.findChildren("w:r")) {
		for (const child of run.children) {
			if (child.tag === "w:delText") child.tag = "w:t";
		}
	}
}

function rewriteGlyph(sdtContent: XmlNode, glyph: string): void {
	// The canonical sdtContent shape is `<w:r/>` followed by `<w:r><w:t>☐|☒</w:t></w:r>`.
	// Find the run with text and update its text node; leave structural runs alone.
	for (const run of sdtContent.findChildren("w:r")) {
		const text = run.findChild("w:t") ?? run.findChild("w:delText");
		if (text) {
			text.children = [XmlNode.textNode(glyph)];
			return;
		}
	}
	// Fall back: append a new glyph run (defensive — shouldn't be needed).
	const r = new XmlNode("w:r");
	const t = new XmlNode("w:t");
	t.children = [XmlNode.textNode(glyph)];
	r.children.push(t);
	sdtContent.children.push(r);
}

function buildGlyphWrapper(
	wrapperTag: "w:ins" | "w:del",
	glyph: string,
	meta: TrackedMeta,
	textTag: "w:t" | "w:delText",
): XmlNode {
	const wrapper = new XmlNode(wrapperTag, {
		"w:id": String(meta.revisionId),
		"w:author": meta.author,
		"w:date": meta.date,
	});
	const r = new XmlNode("w:r");
	const t = new XmlNode(textTag);
	t.children = [XmlNode.textNode(glyph)];
	r.children.push(t);
	wrapper.children.push(r);
	return wrapper;
}
