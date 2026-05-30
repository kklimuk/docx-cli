import {
	isRunBearingWrapper,
	runTextLength,
	sliceRun,
	sumRunBearingTextLength,
	XmlNode,
} from "../parser";
import type { TrackedMeta } from "../track-changes";
import { Del } from "../track-changes/emit";
import { type NoteKind, noteConfig } from "./config";

/** Splice `noteRun` into `paragraph` at the given character offset. Mirrors
 *  the run-splitting machinery in `core/comments/markers.tsx`: walks the
 *  paragraph in document order, descends into run-bearing wrappers
 *  (`<w:ins>`, `<w:del>`, `<w:hyperlink>`, …), and when the cursor crosses the
 *  target offset slices the active `<w:r>` and drops `noteRun` between the
 *  halves. `pPr` is preserved at the head; offsets at zero land after the
 *  `<w:pPr>`; offsets at the paragraph end land after the final run. */
export function insertNoteReferenceAtOffset(
	paragraph: XmlNode,
	offset: number,
	noteRun: XmlNode,
): void {
	const total = sumRunBearingTextLength(paragraph.children);
	if (offset < 0 || offset > total) {
		throw new NoteOffsetOutOfRangeError(
			`Offset ${offset} out of paragraph length ${total}`,
		);
	}
	const state: PlacementState = { offset: 0, placed: false };
	paragraph.children = walkAndPlace(paragraph.children, offset, noteRun, state);
	if (!state.placed) {
		// Tail position — every visible child was passed through and we never
		// hit the offset. Append after the last child.
		paragraph.children.push(noteRun);
		state.placed = true;
	}
}

function walkAndPlace(
	children: XmlNode[],
	targetOffset: number,
	noteRun: XmlNode,
	state: PlacementState,
): XmlNode[] {
	const out: XmlNode[] = [];
	for (const child of children) {
		if (state.placed) {
			out.push(child);
			continue;
		}
		if (child.tag === "w:pPr") {
			out.push(child);
			if (state.offset === targetOffset) {
				out.push(noteRun);
				state.placed = true;
			}
			continue;
		}
		if (child.tag === "w:r") {
			const length = runTextLength(child);
			const runStart = state.offset;
			const runEnd = runStart + length;
			if (targetOffset > runEnd) {
				out.push(child);
				state.offset = runEnd;
				continue;
			}
			// Target lies inside (or at the boundary of) this run.
			const localOffset = targetOffset - runStart;
			if (localOffset === 0) {
				out.push(noteRun);
				state.placed = true;
				out.push(child);
				state.offset = runEnd;
				continue;
			}
			if (localOffset === length) {
				out.push(child);
				out.push(noteRun);
				state.placed = true;
				state.offset = runEnd;
				continue;
			}
			out.push(sliceRun(child, 0, localOffset));
			out.push(noteRun);
			state.placed = true;
			out.push(sliceRun(child, localOffset, length));
			state.offset = runEnd;
			continue;
		}
		if (isRunBearingWrapper(child.tag)) {
			const innerLength = sumRunBearingTextLength(child.children);
			const wrapperStart = state.offset;
			const wrapperEnd = wrapperStart + innerLength;
			if (targetOffset >= wrapperStart && targetOffset <= wrapperEnd) {
				const replacement = new XmlNode(child.tag, { ...child.attributes });
				replacement.children = walkAndPlace(
					child.children,
					targetOffset,
					noteRun,
					state,
				);
				out.push(replacement);
				continue;
			}
			state.offset = wrapperEnd;
			out.push(child);
			continue;
		}
		out.push(child);
	}
	return out;
}

export class NoteOffsetOutOfRangeError extends Error {
	constructor(message: string) {
		super(message);
		this.name = "NoteOffsetOutOfRangeError";
	}
}

/** Strip every `<w:r>` wrapping a reference to `numericId` from the document
 *  tree. Used by `delete` so the body no longer points at a footnote/endnote
 *  whose `<w:footnote>` / `<w:endnote>` body we've removed. Walks recursively
 *  through tracked-change wrappers and tables so deletion is honest no matter
 *  where the reference lives. */
export function removeNoteReferences(
	documentTree: XmlNode[],
	kind: NoteKind,
	numericId: string,
): void {
	const document = XmlNode.findRoot(documentTree, "w:document");
	if (!document) return;
	const config = noteConfig(kind);
	walkAndPrune(document, config.referenceTag, numericId);
}

function walkAndPrune(
	node: XmlNode,
	referenceTag: string,
	numericId: string,
): void {
	const filtered: XmlNode[] = [];
	for (const child of node.children) {
		if (
			child.tag === "w:r" &&
			containsReference(child, referenceTag, numericId)
		) {
			continue;
		}
		walkAndPrune(child, referenceTag, numericId);
		filtered.push(child);
	}
	node.children = filtered;
}

function containsReference(
	run: XmlNode,
	referenceTag: string,
	numericId: string,
): boolean {
	for (const child of run.children) {
		if (
			child.tag === referenceTag &&
			child.getAttribute("w:id") === numericId
		) {
			return true;
		}
	}
	return false;
}

/** The tracked counterpart of `removeNoteReferences`: find every `<w:r>` in
 *  `documentTree` that contains a reference to `(kind, numericId)` and wrap it
 *  in `<w:del meta>` (the run stays intact; the wrapper records the revision).
 *  The post-pass GC in `core/track-changes/apply.ts` removes both the
 *  reference and the body on accept. Skips runs already inside a `<w:del>`
 *  (idempotent under repeated `footnotes delete fnN`). */
export function wrapNoteReferencesAsDeleted(
	documentTree: XmlNode[],
	kind: NoteKind,
	numericId: string,
	meta: TrackedMeta,
): void {
	const document = XmlNode.findRoot(documentTree, "w:document");
	if (!document) return;
	const config = noteConfig(kind);
	walkAndWrap(document, config.referenceTag, numericId, meta);
}

function walkAndWrap(
	node: XmlNode,
	referenceTag: string,
	numericId: string,
	meta: TrackedMeta,
): void {
	if (node.tag === "w:del") return; // already inside a tracked deletion
	const replaced: XmlNode[] = [];
	for (const child of node.children) {
		if (
			child.tag === "w:r" &&
			containsReference(child, referenceTag, numericId)
		) {
			replaced.push(<Del meta={meta}>{child}</Del>);
			continue;
		}
		walkAndWrap(child, referenceTag, numericId, meta);
		replaced.push(child);
	}
	node.children = replaced;
}

type PlacementState = { offset: number; placed: boolean };
