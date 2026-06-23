import type { Document } from "../ast/document";
import type { BlockRangeReference } from "../ast/document/body";
import {
	applyParagraphOptionsInPlace,
	hasParagraphProperties,
	type ParagraphOptions,
	wrapPprChange,
} from "../blocks";
import { partitionParagraphRuns, XmlNode } from "../parser";
import { Del, Ins, markParagraphMarkAs } from "./emit";
import {
	convertTextToDelText,
	resolveAuthor,
	resolveDate,
	TrackChanges,
	type TrackedMeta,
	wrapContiguousTrackable,
} from "./index";
import {
	buildTrackedRuns,
	buildUntrackedRuns,
	diffTokens,
	extractOldTokens,
	paragraphMarkRunRpr,
	tokenize,
} from "./preserve-formatting";

/** Apply a word-level diff edit to a paragraph's content. Untracked mode emits
 *  only kept + inserted tokens with rPr inherited from the old runs; tracked
 *  mode emits the Word-shaped mix of plain runs / `<w:ins>` / `<w:del>` that
 *  Word produces when an author edits a few words mid-paragraph under tracking.
 *  Pure pPr properties (`--style`, `--alignment`) apply to the existing pPr in
 *  place via `applyParagraphOptionsInPlace`. */
export function applyFormattingPreservingEdit(
	document: Document,
	paragraph: XmlNode,
	newText: string,
	paragraphOptions: ParagraphOptions,
	authorFlag: string | undefined,
	tracked: boolean,
): void {
	const oldTokens = extractOldTokens(paragraph);
	const newTokens = tokenize(newText);
	const ops = diffTokens(oldTokens, newTokens);

	// When there's no run-level neighbor to inherit from (filling an empty
	// paragraph/cell), fall back to the paragraph-mark rPr so the new text picks
	// up the cell's declared font/size — matching how Word fills empty styled cells.
	const fallbackRpr = paragraphMarkRunRpr(paragraph);
	const runChildren = tracked
		? buildTrackedRuns(ops, makeMetaMinter(document, authorFlag), fallbackRpr)
		: buildUntrackedRuns(ops, fallbackRpr);

	const { nonRuns } = partitionParagraphRuns(paragraph);
	// Paragraph properties riding along with the text edit (`--style`/`--alignment`/
	// `--space-*`/`--line-spacing`/`--indent-*`/`--tabs`) are a real tracked
	// revision under tracking: snapshot the prior `<w:pPr>` into a `<w:pPrChange>`
	// BEFORE the in-place mutation so reject restores it (accept drops the marker).
	// Mirrors `Edit.paragraphProperties`; without this the change applied silently
	// and survived reject.
	if (tracked && hasParagraphProperties(paragraphOptions)) {
		let pPr = nonRuns.find((child) => child.tag === "w:pPr");
		if (!pPr) {
			pPr = new XmlNode("w:pPr");
			nonRuns.unshift(pPr);
		}
		wrapPprChange(pPr, makeMetaMinter(document, authorFlag)());
	}
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
	document: Document,
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
	const mintMeta = makeMetaMinter(document, authorFlag);
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
	document: Document,
	parent: XmlNode[],
	startIndex: number,
	endIndex: number,
	authorFlag: string | undefined,
): void {
	const mintMeta = makeMetaMinter(document, authorFlag);
	for (let index = startIndex; index <= endIndex; index++) {
		const para = parent[index];
		if (!para) continue;
		convertParagraphContentToDeleted(para, mintMeta);
		if (index < endIndex) {
			markParagraphMarkAs(para, "del", mintMeta());
		}
	}
}

/** Error thrown by `assertParagraphOnlyTrackedRange` when a tracked
 *  range-replace or range-delete would touch a non-paragraph block (most
 *  commonly a table). Both `Edit.range` and `delete --at pN-pM` rely on the
 *  same guard; this is the shared shape they map into their own error code. */
export class TrackedRangeConflictError extends Error {
	constructor(
		message: string,
		public hint: string,
	) {
		super(message);
		this.name = "TrackedRangeConflictError";
	}
}

/** Validate that every block in `[startIndex, endIndex]` is a `<w:p>` — the
 *  tracked-range walker (`applyTrackedRangeReplace` / `applyTrackedRangeDelete`)
 *  injects a `<w:pPr>` into every span block, which would corrupt a `<w:tbl>`
 *  or other non-paragraph node. Used by `Edit.range` and `delete --at pN-pM`
 *  before applying. The untracked paths splice cleanly across any block tag,
 *  so this guard only fires when tracking is on. */
export function assertParagraphOnlyTrackedRange(
	rangeRef: BlockRangeReference,
): void {
	for (let i = rangeRef.startIndex; i <= rangeRef.endIndex; i++) {
		const block = rangeRef.parent[i];
		if (block && block.tag !== "w:p") {
			const what = block.tag === "w:tbl" ? "a table" : `a ${block.tag} block`;
			throw new TrackedRangeConflictError(
				`Tracked range spans ${what} — not supported.`,
				"Toggle tracking off (`docx track-changes off`), or handle the table separately via `docx delete --at tN`.",
			);
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
	// Wrap each contiguous span of trackable run-level children
	// (`<w:r>`, `<m:oMath>`, `<m:oMathPara>`) in its own `<w:del>`, leaving
	// pre-existing wrappers (`<w:ins>`, `<w:del>`, `<w:hyperlink>`,
	// `<w:moveFrom>`, `<w:moveTo>`, `<w:fldSimple>`, `<w:smartTag>`) in
	// place. Mirrors the `applyDeletion` shape in `index.tsx` so a tracked
	// range-delete preserves walker-emitted CriticMarkup (or any other
	// inner-wrapper-carrying input) instead of nesting it under an outer
	// del that clobbers its metadata. See [CLAUDE.md "Tracked range edit /
	// delete — Word-canonical shapes"](./CLAUDE.md).
	paragraph.children = wrapContiguousTrackable(paragraph.children, (runs) => {
		const converted = runs.map((run) =>
			run.tag === "w:r" ? convertTextToDelText(run) : run,
		);
		return <Del meta={mintMeta()}>{converted}</Del>;
	});
}

function wrapNewParagraphContentAsInserted(
	paragraph: XmlNode,
	mintMeta: () => TrackedMeta,
): void {
	// See `convertParagraphContentToDeleted` for the rationale — same
	// contiguous-wrap pattern, just emitting `<w:ins>`. This is the path
	// that was previously losing CriticMarkup `<w:ins>`/`<w:del>` metadata
	// from `edit --markdown` source — the walker's inner wrapper would get
	// flattened into the outer one carrying the editor's author, dropping
	// the markdown source's author/revisionId entirely.
	paragraph.children = wrapContiguousTrackable(paragraph.children, (runs) => (
		<Ins meta={mintMeta()}>{runs}</Ins>
	));
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
	document: Document,
	authorFlag: string | undefined,
): () => TrackedMeta {
	const allocator = new TrackChanges(document).createAllocator();
	const baseMeta = { author: resolveAuthor(authorFlag), date: resolveDate() };
	return () => ({ ...baseMeta, revisionId: allocator.next() });
}
