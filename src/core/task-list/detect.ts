import type { DocView } from "../ast/doc-view";
import type { Paragraph } from "../ast/types";
import { getListBulletText } from "../numbering";
import type { XmlNode } from "../parser";

/** Recognize a GFM task-list paragraph. Two shapes in the wild:
 *
 *  1. **SDT content control** (Pandoc, LibreOffice, Word for Mac/Windows
 *     desktop, even after Word for Web round-trips it without editing) —
 *     `<w:sdt><w:sdtPr><w14:checkbox/></w:sdtPr><w:sdtContent>…☐|☒…</w:sdtContent></w:sdt>`
 *     followed by a whitespace `<w:r>`. The `w14:checked` attribute is the
 *     source of truth for state; the glyph in sdtContent is decoration.
 *
 *  2. **Wingdings bullet + strike** (Word for Web's "Checklist" feature) —
 *     a list paragraph whose bullet glyph is U+F0A8 (Wingdings ☐). Checked
 *     state is conveyed via `<w:strike w:val="1"/>` on the paragraph-mark
 *     `<w:rPr>` (and mirrored on every text run). No SDT involved.
 *
 *  Returns the set of children to skip during the run walk so the SDT glyph
 *  and trailing space don't leak into AST `runs`. The Wingdings-bullet path
 *  doesn't introduce skipped nodes — the text runs ARE the task content.
 *
 *  When called from `ast/read.ts`, also registers a `checkboxToggle` tracked-
 *  change reference for any toggle pattern (`<w:ins>`+`<w:del>` inside
 *  `<w:sdtContent>`); the apply walker in `cli/track-changes/apply.ts` does
 *  the same in matching order so `tcN` ids agree between `list` and
 *  `apply --at tcN`. */
export function detectTaskListState(
	view: DocView,
	paragraph: Paragraph,
	node: XmlNode,
	registerToggle: (sdt: XmlNode, parent: XmlNode[]) => void,
): Set<XmlNode> {
	const skip = new Set<XmlNode>();
	// Per GFM, task list items must live inside a list. Word permits a stray
	// checkbox content control on any paragraph but we don't surface those as
	// tasks — the SDT stays in the XmlNode tree and round-trips untouched.
	if (!paragraph.list) return skip;

	// Shape 1: leading SDT checkbox.
	const firstContent = node.children.find((child) => child.tag !== "w:pPr");
	if (firstContent?.tag === "w:sdt") {
		const checkbox = firstContent
			.findChild("w:sdtPr")
			?.findChild("w14:checkbox");
		if (checkbox) {
			const checked =
				checkbox.findChild("w14:checked")?.getAttribute("w14:val") === "1";
			paragraph.taskState = checked ? "checked" : "unchecked";
			skip.add(firstContent);
			// Skip the conventional space run between the SDT and the task text.
			// Under tracking the runs may be wrapped in `<w:ins>`/`<w:del>` — look
			// one level deep too so a tracked-insert task still strips the lead
			// space from `runs` (which would otherwise double-space the markdown).
			const next = nextNonPPrSibling(node, firstContent);
			if (next) {
				if (next.tag === "w:r" && isWhitespaceOnlyRun(next)) {
					skip.add(next);
				} else if (next.tag === "w:ins" || next.tag === "w:del") {
					const firstInner = next.children.find((c) => c.tag !== "w:rPr");
					if (firstInner?.tag === "w:r" && isWhitespaceOnlyRun(firstInner)) {
						skip.add(firstInner);
					}
				}
			}
			// Register a checkboxToggle reference if the SDT has an ins+del pair
			// (Word's canonical tracked-toggle shape). MUST match the order of
			// the apply walker, which sees the SDT as the first non-pPr child
			// and registers the toggle as the first tcN of this paragraph.
			if (findCheckboxToggle(firstContent)) {
				registerToggle(firstContent, node.children);
			}
			return skip;
		}
	}

	// Shape 2: Word-for-Web Checklist (Wingdings ☐ bullet + paragraph-mark
	// strike for "done"). The bullet character is U+F0A8 in Word for Web; we
	// also accept the Unicode ☐ U+2610 in case other tools emit it.
	const bulletText = getListBulletText(
		view,
		paragraph.list.numId,
		paragraph.list.level,
	);
	if (bulletText === "" || bulletText === "☐") {
		paragraph.taskState = isParagraphMarkStruck(node) ? "checked" : "unchecked";
	}
	return skip;
}

/** Empirically validated against Microsoft Word and LibreOffice: when the user
 *  toggles a `<w14:checkbox>` content control with track-changes on, Word
 *  emits an `<w:ins>` (new glyph ☒ or ☐) and `<w:del>` (old glyph ☐ or ☒) pair
 *  INSIDE `<w:sdtContent>` — and ALSO flips the `w14:checked` attribute in
 *  place (no separate `<w14:checkedChange>` element exists in the spec). The
 *  AST surfaces the pair as a single "checkboxToggle" tracked change so
 *  accept/reject keeps the glyph and the attribute consistent.
 *
 *  Returns `null` for any SDT that isn't a checkbox or doesn't show a
 *  canonical ins+del pair — e.g. a checkbox SDT with no tracking, a checkbox
 *  with only an `<w:ins>` (a partial state we don't recognize as a toggle),
 *  or other content-control kinds. */
export function findCheckboxToggle(sdt: XmlNode): {
	ins: XmlNode;
	del: XmlNode;
	insGlyph: string;
	delGlyph: string;
} | null {
	const sdtPr = sdt.findChild("w:sdtPr");
	if (!sdtPr?.findChild("w14:checkbox")) return null;
	const sdtContent = sdt.findChild("w:sdtContent");
	if (!sdtContent) return null;
	const ins = sdtContent.findChild("w:ins");
	const del = sdtContent.findChild("w:del");
	if (!ins || !del) return null;
	const insGlyph = collectGlyphText(ins);
	const delGlyph = collectGlyphText(del);
	if (!insGlyph || !delGlyph) return null;
	return { ins, del, insGlyph, delGlyph };
}

function isParagraphMarkStruck(node: XmlNode): boolean {
	const pPr = node.findChild("w:pPr");
	const rPr = pPr?.findChild("w:rPr");
	const strike = rPr?.findChild("w:strike");
	if (!strike) return false;
	// `<w:strike/>` (no val) defaults to true per ECMA-376; explicit "0" / "false" is off.
	const value = strike.getAttribute("w:val");
	if (value === undefined) return true;
	return value !== "0" && value !== "false";
}

function nextNonPPrSibling(
	parent: XmlNode,
	after: XmlNode,
): XmlNode | undefined {
	const index = parent.children.indexOf(after);
	if (index === -1) return undefined;
	for (let i = index + 1; i < parent.children.length; i++) {
		const child = parent.children[i];
		if (child && child.tag !== "w:pPr") return child;
	}
	return undefined;
}

function isWhitespaceOnlyRun(run: XmlNode): boolean {
	for (const child of run.children) {
		if (child.tag === "w:rPr") continue;
		if (child.tag !== "w:t") return false;
		if (child.collectText().trim().length > 0) return false;
	}
	return true;
}

function collectGlyphText(wrapper: XmlNode): string {
	let out = "";
	for (const run of wrapper.findChildren("w:r")) {
		for (const child of run.children) {
			if (child.tag === "w:t" || child.tag === "w:delText") {
				out += child.collectText();
			}
		}
	}
	return out;
}
