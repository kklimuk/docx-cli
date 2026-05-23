import type { DocView } from "../ast/doc-view";
import { applyParagraphOptionsInPlace, type ParagraphOptions } from "../blocks";
import { partitionParagraphRuns, XmlNode } from "../parser";
import { Del, Ins, markParagraphMarkAs } from "./emit";
import {
	convertTextToDelText,
	createRevisionAllocator,
	resolveAuthor,
	resolveDate,
	type TrackedMeta,
} from "./index";
import {
	buildTrackedRuns,
	buildUntrackedRuns,
	diffTokens,
	extractOldTokens,
	tokenize,
} from "./preserve-formatting";

/** Apply a word-level diff edit to a paragraph's content. Untracked mode emits
 *  only kept + inserted tokens with rPr inherited from the old runs; tracked
 *  mode emits the Word-shaped mix of plain runs / `<w:ins>` / `<w:del>` that
 *  Word produces when an author edits a few words mid-paragraph under tracking.
 *  Pure pPr properties (`--style`, `--alignment`) apply to the existing pPr in
 *  place via `applyParagraphOptionsInPlace`. */
export function applyFormattingPreservingEdit(
	view: DocView,
	paragraph: XmlNode,
	newText: string,
	paragraphOptions: ParagraphOptions,
	authorFlag: string | undefined,
	tracked: boolean,
): void {
	const oldTokens = extractOldTokens(paragraph);
	const newTokens = tokenize(newText);
	const ops = diffTokens(oldTokens, newTokens);

	const runChildren = tracked
		? buildTrackedRuns(ops, makeMetaMinter(view, authorFlag))
		: buildUntrackedRuns(ops);

	const { nonRuns } = partitionParagraphRuns(paragraph);
	applyParagraphOptionsInPlace(nonRuns, paragraphOptions);
	nonRuns.push(...runChildren);
	paragraph.children = nonRuns;
}

/** Replace a contiguous span of paragraphs with new paragraphs, tracking the
 *  change in the Word-canonical shape (empirically validated against
 *  Microsoft Word — see `scripts/word-redlines.sh` and the probe outputs in
 *  `/tmp/range-probe/`):
 *
 *  - For each old paragraph 1..M-1: content wrapped in `<w:del>`, the
 *    paragraph-mark marked as `<w:del>` so on accept it merges into the next.
 *  - For the LAST old paragraph (the "transition"): content wrapped in
 *    `<w:del>`; the FIRST new paragraph's content is appended inline wrapped
 *    in `<w:ins>`. The paragraph-mark gets an `<w:ins>` marker if N≥2 (so a
 *    new paragraph break opens between this and the next new paragraph),
 *    otherwise stays bare.
 *  - Each new paragraph 2..N-1 is inserted as a fresh `<w:p>` with content
 *    wrapped in `<w:ins>` and the paragraph-mark `<w:ins>`-marked.
 *  - The LAST new paragraph: content wrapped in `<w:ins>`, paragraph-mark
 *    bare (it becomes the trailing container — equivalent to where Word
 *    leaves the cursor after the operation).
 *
 *  Accept-all collapses the old M paragraphs and unwraps the new content;
 *  reject-all removes the new content and restores the old. Both end states
 *  match Word's own accept/reject behavior on equivalent input. */
export function applyTrackedRangeReplace(
	view: DocView,
	parent: XmlNode[],
	startIndex: number,
	endIndex: number,
	newParagraphs: XmlNode[],
	authorFlag: string | undefined,
): void {
	if (newParagraphs.length === 0) {
		throw new Error(
			"applyTrackedRangeReplace requires at least one new paragraph; use applyTrackedRangeDelete for pure deletion",
		);
	}
	const mintMeta = makeMetaMinter(view, authorFlag);
	const newN = newParagraphs.length;

	// Mark every old paragraph except the LAST: content del + paragraph-mark del.
	// On accept-all these all merge forward into the transition paragraph.
	for (let index = startIndex; index < endIndex; index++) {
		const para = parent[index];
		if (!para) continue;
		convertParagraphContentToDeleted(para, mintMeta);
		markParagraphMarkAs(para, "del", mintMeta());
	}

	// Transition paragraph (old endIndex): del its content, then append the
	// first new paragraph's content wrapped in <w:ins>. Adopt first new's
	// pPr if it has one (handles `--style` overrides on range edits).
	const transition = parent[endIndex];
	const firstNew = newParagraphs[0];
	if (!transition || !firstNew) return;
	const firstNewPPr = firstNew.findChild("w:pPr");
	convertParagraphContentToDeleted(transition, mintMeta);
	if (firstNewPPr) {
		replacePPr(transition, firstNewPPr);
	}
	const firstNewRuns: XmlNode[] = [];
	for (const child of firstNew.children) {
		if (child.tag === "w:pPr") continue;
		firstNewRuns.push(child);
	}
	if (firstNewRuns.length > 0) {
		transition.children.push(<Ins meta={mintMeta()}>{firstNewRuns}</Ins>);
	}

	// If N >= 2, the transition's paragraph-mark is itself "inserted" (Word's
	// model: typing Enter after writing the first new line opens a new
	// paragraph). Without it, accept-all would leave only one paragraph.
	if (newN >= 2) {
		markParagraphMarkAs(transition, "ins", mintMeta());
	}

	// Splice in new paragraphs 2..N. Each one's content gets <w:ins>-wrapped;
	// marks are <w:ins>-marked EXCEPT the trailing one (which is the final
	// container — its bare mark is what stops the merge cascade on accept).
	if (newN >= 2) {
		const toSplice: XmlNode[] = [];
		for (let index = 1; index < newN; index++) {
			const newPara = newParagraphs[index];
			if (!newPara) continue;
			wrapNewParagraphContentAsInserted(newPara, mintMeta);
			if (index < newN - 1) {
				markParagraphMarkAs(newPara, "ins", mintMeta());
			}
			toSplice.push(newPara);
		}
		parent.splice(endIndex + 1, 0, ...toSplice);
	}
}

/** Replace a contiguous span of paragraphs with new paragraphs without
 *  tracking — a plain splice. Companion of `applyTrackedRangeReplace`.
 *
 *  Before splicing, lifts an inline `<w:sectPr>` carried by the LAST old
 *  paragraph (a section-boundary paragraph) onto the LAST new paragraph's
 *  pPr. Otherwise the section break would vanish silently when the old
 *  paragraph is removed. */
export function applyUntrackedRangeReplace(
	parent: XmlNode[],
	startIndex: number,
	endIndex: number,
	newParagraphs: XmlNode[],
): void {
	const lastOld = parent[endIndex];
	const lastNew = newParagraphs[newParagraphs.length - 1];
	if (lastOld && lastNew) {
		const sectPr = lastOld.findChild("w:pPr")?.findChild("w:sectPr");
		if (sectPr) {
			let newPPr = lastNew.findChild("w:pPr");
			if (!newPPr) {
				newPPr = new XmlNode("w:pPr");
				lastNew.children.unshift(newPPr);
			}
			if (!newPPr.findChild("w:sectPr")) {
				newPPr.children.push(sectPr);
			}
		}
	}
	parent.splice(startIndex, endIndex - startIndex + 1, ...newParagraphs);
}

/** Delete a contiguous span of paragraphs, tracking the change in Word's
 *  canonical shape (empirically validated — see
 *  `/tmp/range-probe/delete-4.docx`):
 *
 *  - Every paragraph in `[startIndex, endIndex]` has its content wrapped in
 *    `<w:del>` (text → delText).
 *  - Paragraphs `startIndex..endIndex-1` have their paragraph-mark marked as
 *    `<w:del>`. The LAST paragraph's mark stays bare.
 *
 *  On accept-all the M paragraphs collapse into one empty residue paragraph
 *  (with the last paragraph's mark). On reject-all the original M paragraphs
 *  are restored intact. The residue matches Word's behavior — markdown render
 *  hides empty paragraphs, so an agent reading the doc sees a clean delete. */
export function applyTrackedRangeDelete(
	view: DocView,
	parent: XmlNode[],
	startIndex: number,
	endIndex: number,
	authorFlag: string | undefined,
): void {
	const mintMeta = makeMetaMinter(view, authorFlag);
	for (let index = startIndex; index <= endIndex; index++) {
		const para = parent[index];
		if (!para) continue;
		convertParagraphContentToDeleted(para, mintMeta);
		if (index < endIndex) {
			markParagraphMarkAs(para, "del", mintMeta());
		}
	}
}

/** Delete a contiguous span of paragraphs without tracking — a plain splice
 *  removing `[startIndex, endIndex]` inclusive. Companion of
 *  `applyTrackedRangeDelete`. */
export function applyUntrackedRangeDelete(
	parent: XmlNode[],
	startIndex: number,
	endIndex: number,
): void {
	parent.splice(startIndex, endIndex - startIndex + 1);
}

function convertParagraphContentToDeleted(
	paragraph: XmlNode,
	mintMeta: () => TrackedMeta,
): void {
	const { runs, nonRuns } = partitionParagraphRuns(paragraph);
	paragraph.children = nonRuns;
	if (runs.length === 0) return;
	const deletedRuns = runs.map((run) => convertTextToDelText(run));
	paragraph.children.push(<Del meta={mintMeta()}>{deletedRuns}</Del>);
}

function wrapNewParagraphContentAsInserted(
	paragraph: XmlNode,
	mintMeta: () => TrackedMeta,
): void {
	const { runs, nonRuns } = partitionParagraphRuns(paragraph);
	paragraph.children = nonRuns;
	if (runs.length === 0) return;
	paragraph.children.push(<Ins meta={mintMeta()}>{runs}</Ins>);
}

function replacePPr(paragraph: XmlNode, newPPr: XmlNode): void {
	const existing = paragraph.findChild("w:pPr");
	if (existing) {
		// Lift an inline `<w:sectPr>` (section break carried by this paragraph)
		// onto the new pPr — losing it would silently drop a section boundary.
		// Only copy if the new pPr doesn't already carry its own sectPr.
		const sectPr = existing.findChild("w:sectPr");
		if (sectPr && !newPPr.findChild("w:sectPr")) {
			newPPr.children.push(sectPr);
		}
		const idx = paragraph.children.indexOf(existing);
		paragraph.children.splice(idx, 1, newPPr);
	} else {
		paragraph.children.unshift(newPPr);
	}
}

function makeMetaMinter(
	view: DocView,
	authorFlag: string | undefined,
): () => TrackedMeta {
	const allocator = createRevisionAllocator(view);
	const baseMeta = { author: resolveAuthor(authorFlag), date: resolveDate() };
	return () => ({ ...baseMeta, revisionId: allocator.next() });
}
