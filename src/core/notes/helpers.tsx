import type { DocView } from "../ast/doc-view";
import { w } from "../jsx";
import { registerPart } from "../package";
import {
	isRunBearingWrapper,
	runTextLength,
	sliceRun,
	sumRunBearingTextLength,
	XmlNode,
} from "../parser";
import { ensureStyle } from "../styles";
import { convertTextToDelText, type TrackedMeta } from "../track-changes";
import { Del, Ins, markParagraphMarkAs } from "../track-changes/emit";

/** Provision footnotes.xml/endnotes.xml if absent. Word writes two reserved
 *  entries before any user notes — id=-1 separator (the rule above the note
 *  area) and id=0 continuationSeparator (used when notes wrap across pages).
 *  Both have `w:type` set; both are filtered out by core/ast/read.ts. We seed
 *  them so LibreOffice and Word render the note area correctly the first time
 *  an agent calls `footnotes add` against a doc that had none. */
export function ensureNotesPart(view: DocView, kind: NoteKind): XmlNode {
	const existing = getTree(view, kind);
	const config = noteConfig(kind);
	if (existing) {
		const root = XmlNode.findRoot(existing, config.rootTag);
		if (root) return root;
	}
	const root = buildEmptyNotesRoot(config);
	setTree(view, kind, [root]);
	registerPart(view.relationshipsTree, view.contentTypesTree, {
		partName: config.partName,
		contentType: config.contentType,
		relationshipType: config.relationshipType,
		target: config.target,
	});
	return root;
}

function buildEmptyNotesRoot(config: NoteConfig): XmlNode {
	const Root = config.kind === "footnote" ? w.footnotes : w.endnotes;
	return (
		<Root {...{ "xmlns:w": NS_W }}>
			<NoteBoilerplate config={config} type="separator" id="-1" />
			<NoteBoilerplate config={config} type="continuationSeparator" id="0" />
		</Root>
	);
}

function NoteBoilerplate({
	config,
	type,
	id,
}: {
	config: NoteConfig;
	type: "separator" | "continuationSeparator";
	id: string;
}): XmlNode {
	const Note = config.kind === "footnote" ? w.footnote : w.endnote;
	return (
		<Note w-id={id} w-type={type}>
			<w.p>
				<w.r>
					{type === "separator" ? <w.separator /> : <w.continuationSeparator />}
				</w.r>
			</w.p>
		</Note>
	);
}

function getTree(view: DocView, kind: NoteKind): XmlNode[] | undefined {
	return kind === "footnote" ? view.footnotesTree : view.endnotesTree;
}

function setTree(view: DocView, kind: NoteKind, tree: XmlNode[]): void {
	if (kind === "footnote") view.footnotesTree = tree;
	else view.endnotesTree = tree;
}

/** Lazily provision the two character/paragraph styles a note relies on. The
 *  baseline catalog defines both — these calls just register the style nodes
 *  in `styles.xml` if not already present, which avoids Word's fall-back to
 *  Normal (no superscript on the marker, default font on the body). */
export function ensureNoteStyles(view: DocView, kind: NoteKind): void {
	const config = noteConfig(kind);
	ensureStyle(view, config.referenceStyle);
	ensureStyle(view, config.textStyle);
}

/** Word reserves -1 (separator) and 0 (continuationSeparator) as the boilerplate
 *  ids; user notes start at 1. We allocate `max(existing user id) + 1` so we
 *  never collide with the reserved entries even if the document was authored
 *  by something that used different defaults. */
export function nextNoteId(view: DocView, kind: NoteKind): string {
	const tree = getTree(view, kind);
	const config = noteConfig(kind);
	if (!tree) return "1";
	const root = XmlNode.findRoot(tree, config.rootTag);
	if (!root) return "1";
	let highest = 0;
	for (const child of root.children) {
		if (child.tag !== config.itemTag) continue;
		const idAttribute = child.getAttribute("w:id");
		if (idAttribute == null) continue;
		const numeric = Number(idAttribute);
		if (Number.isFinite(numeric) && numeric > highest) highest = numeric;
	}
	return String(highest + 1);
}

export function findNoteByNumericId(
	view: DocView,
	kind: NoteKind,
	numericId: string,
): { node: XmlNode; parent: XmlNode[] } | undefined {
	const tree = getTree(view, kind);
	if (!tree) return undefined;
	const config = noteConfig(kind);
	const root = XmlNode.findRoot(tree, config.rootTag);
	if (!root) return undefined;
	for (const child of root.children) {
		if (child.tag !== config.itemTag) continue;
		if (child.getAttribute("w:id") !== numericId) continue;
		return { node: child, parent: root.children };
	}
	return undefined;
}

/** The run that lives in the body and points at the note. ECMA-376 §17.11.14
 *  binds the rStyle "FootnoteReference" / "EndnoteReference" character style
 *  (superscript number rendering) to the `<w:footnoteReference>` / `<w:endnoteReference>`
 *  marker — without it Word renders the marker without superscripting. */
export function NoteReferenceRun({
	config,
	id,
}: {
	config: NoteConfig;
	id: string;
}): XmlNode {
	const Reference =
		config.kind === "footnote" ? w.footnoteReference : w.endnoteReference;
	return (
		<w.r>
			<w.rPr>
				<w.rStyle w-val={config.referenceStyle} />
			</w.rPr>
			<Reference w-id={id} />
		</w.r>
	);
}

/** The note's own body paragraph — a `<w:footnote>` / `<w:endnote>` whose first
 *  run carries the back-reference marker (`<w:footnoteRef/>` / `<w:endnoteRef/>`)
 *  that Word renders as the same numeral as the in-body reference. The text
 *  follows as a separate run with a leading space, matching the shape Word
 *  emits. */
export function NoteBody({
	config,
	id,
	text,
}: {
	config: NoteConfig;
	id: string;
	text: string;
}): XmlNode {
	const Note = config.kind === "footnote" ? w.footnote : w.endnote;
	const BodyRef = config.kind === "footnote" ? w.footnoteRef : w.endnoteRef;
	return (
		<Note w-id={id}>
			<w.p>
				<w.pPr>
					<w.pStyle w-val={config.textStyle} />
				</w.pPr>
				<w.r>
					<w.rPr>
						<w.rStyle w-val={config.referenceStyle} />
					</w.rPr>
					<BodyRef />
				</w.r>
				<w.r>
					<w.t {...{ "xml:space": "preserve" }}>{` ${text}`}</w.t>
				</w.r>
			</w.p>
		</Note>
	);
}

/** Body of a footnote that's being added under tracking: the entire run
 *  content (back-reference marker + text) is wrapped in `<w:ins>` so accept
 *  unwraps the runs (footnote becomes normal) and reject lets the post-pass
 *  GC the now-orphan `<w:footnote>` wrapper. Matches Word's empirical shape
 *  from `/tmp/fn-probe/add.docx` — the `<w:footnote>` wrapper itself is NOT
 *  wrapped, only the inner runs. The paragraph mark stays bare too. */
export function TrackedNoteBody({
	config,
	id,
	text,
	meta,
}: {
	config: NoteConfig;
	id: string;
	text: string;
	meta: TrackedMeta;
}): XmlNode {
	const Note = config.kind === "footnote" ? w.footnote : w.endnote;
	const BodyRef = config.kind === "footnote" ? w.footnoteRef : w.endnoteRef;
	return (
		<Note w-id={id}>
			<w.p>
				<w.pPr>
					<w.pStyle w-val={config.textStyle} />
				</w.pPr>
				<Ins meta={meta}>
					<w.r>
						<w.rPr>
							<w.rStyle w-val={config.referenceStyle} />
						</w.rPr>
						<BodyRef />
					</w.r>
					<w.r>
						<w.t {...{ "xml:space": "preserve" }}>{` ${text}`}</w.t>
					</w.r>
				</Ins>
			</w.p>
		</Note>
	);
}

/** Splice `noteRun` into `paragraph` at the given character offset. Mirrors
 *  the run-splitting machinery in `cli/comments/helpers.tsx`: walks the
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

/** Wrap an existing note's body content in `<w:del>` (with `<w:t>` →
 *  `<w:delText>` rename) and mark the paragraph-mark for deletion. Mirrors
 *  Word's empirical shape from `/tmp/fn-probe/delete.docx`: three coupled
 *  revisions per delete — the reference run (caller emits), the body content
 *  (this fn), and the paragraph mark (this fn). Idempotent on already-wrapped
 *  bodies (the resulting OOXML still validates; the post-pass GC handles
 *  cleanup either way). */
export function wrapNoteBodyAsDeleted(
	noteBody: XmlNode,
	meta: TrackedMeta,
): void {
	const paragraph = noteBody.findChild("w:p");
	if (!paragraph) return;
	const head: XmlNode[] = [];
	const wrapped: XmlNode[] = [];
	for (const child of paragraph.children) {
		if (child.tag === "w:pPr") {
			head.push(child);
			continue;
		}
		// Convert any `<w:t>` inside this run to `<w:delText>` per ECMA-376
		// §17.16.5 — text inside `<w:del>` must use `<w:delText>`.
		wrapped.push(convertTextToDelText(child));
	}
	const del = (<Del meta={meta}>{wrapped}</Del>) as XmlNode;
	paragraph.children = [...head, del];
	markParagraphMarkAs(paragraph, "del", meta);
}

/** Replace the note body's text with a tracked replacement: the existing
 *  text run becomes a `<w:del>` (with delText conversion), preceded by a
 *  `<w:ins>` carrying the new text. The `<w:footnoteRef/>` run and the
 *  leading whitespace stay bare. Mirrors Word's order in
 *  `/tmp/fn-probe/edit.docx`: ins before del. */
export function wrapNoteBodyAsEdited(
	noteBody: XmlNode,
	newText: string,
	insMeta: TrackedMeta,
	delMeta: TrackedMeta,
): void {
	const paragraph = noteBody.findChild("w:p");
	if (!paragraph) return;
	const result: XmlNode[] = [];
	let inserted = false;
	for (const child of paragraph.children) {
		if (child.tag !== "w:r") {
			result.push(child);
			continue;
		}
		// Identify the user-text run by the absence of the back-reference
		// marker and the absence of an explicit whitespace-only `<w:t>`.
		if (isNoteBackRefRun(child) || isBareWhitespaceRun(child)) {
			result.push(child);
			continue;
		}
		if (!inserted) {
			const newRun = (
				<w.r>
					<w.t {...{ "xml:space": "preserve" }}>{newText}</w.t>
				</w.r>
			) as XmlNode;
			result.push((<Ins meta={insMeta}>{newRun}</Ins>) as XmlNode);
			inserted = true;
		}
		result.push(
			(<Del meta={delMeta}>{convertTextToDelText(child)}</Del>) as XmlNode,
		);
	}
	// No prior user-text run (e.g. body has only the back-ref + whitespace):
	// just append a tracked insertion of the new text.
	if (!inserted) {
		const newRun = (
			<w.r>
				<w.t {...{ "xml:space": "preserve" }}>{newText}</w.t>
			</w.r>
		) as XmlNode;
		result.push((<Ins meta={insMeta}>{newRun}</Ins>) as XmlNode);
	}
	paragraph.children = result;
}

function isNoteBackRefRun(run: XmlNode): boolean {
	for (const child of run.children) {
		if (child.tag === "w:footnoteRef" || child.tag === "w:endnoteRef") {
			return true;
		}
	}
	return false;
}

function isBareWhitespaceRun(run: XmlNode): boolean {
	let saw = false;
	for (const child of run.children) {
		if (child.tag === "w:rPr") continue;
		if (child.tag !== "w:t") return false;
		const text = child.collectText();
		if (text.length === 0 || /\S/.test(text)) return false;
		saw = true;
	}
	return saw;
}

/** Footnotes and endnotes share their entire mechanics — separate parts, but
 *  identical XML shape, same id-allocation rules (Word reserves -1/0 for the
 *  separator/continuationSeparator boilerplate), same reference-run pattern,
 *  same body-paragraph shape. We parameterize on this kind everywhere instead
 *  of duplicating two near-identical modules. */
export function noteConfig(kind: NoteKind): NoteConfig {
	return kind === "footnote" ? FOOTNOTE_CONFIG : ENDNOTE_CONFIG;
}

export type NoteKind = "footnote" | "endnote";

export type NoteConfig = {
	kind: NoteKind;
	idPrefix: "fn" | "en";
	rootTag: "w:footnotes" | "w:endnotes";
	itemTag: "w:footnote" | "w:endnote";
	referenceTag: "w:footnoteReference" | "w:endnoteReference";
	bodyRefTag: "w:footnoteRef" | "w:endnoteRef";
	textStyle: "FootnoteText" | "EndnoteText";
	referenceStyle: "FootnoteReference" | "EndnoteReference";
	partName: "word/footnotes.xml" | "word/endnotes.xml";
	target: "footnotes.xml" | "endnotes.xml";
	relationshipType: string;
	contentType: string;
};

type PlacementState = { offset: number; placed: boolean };

const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const FOOTNOTES_REL_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
const ENDNOTES_REL_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
const FOOTNOTES_CONTENT_TYPE =
	"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
const ENDNOTES_CONTENT_TYPE =
	"application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";

const FOOTNOTE_CONFIG: NoteConfig = {
	kind: "footnote",
	idPrefix: "fn",
	rootTag: "w:footnotes",
	itemTag: "w:footnote",
	referenceTag: "w:footnoteReference",
	bodyRefTag: "w:footnoteRef",
	textStyle: "FootnoteText",
	referenceStyle: "FootnoteReference",
	partName: "word/footnotes.xml",
	target: "footnotes.xml",
	relationshipType: FOOTNOTES_REL_TYPE,
	contentType: FOOTNOTES_CONTENT_TYPE,
};

const ENDNOTE_CONFIG: NoteConfig = {
	kind: "endnote",
	idPrefix: "en",
	rootTag: "w:endnotes",
	itemTag: "w:endnote",
	referenceTag: "w:endnoteReference",
	bodyRefTag: "w:endnoteRef",
	textStyle: "EndnoteText",
	referenceStyle: "EndnoteReference",
	partName: "word/endnotes.xml",
	target: "endnotes.xml",
	relationshipType: ENDNOTES_REL_TYPE,
	contentType: ENDNOTES_CONTENT_TYPE,
};
