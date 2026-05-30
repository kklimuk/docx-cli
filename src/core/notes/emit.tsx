import { w } from "../jsx";
import type { XmlNode } from "../parser";
import { convertTextToDelText, type TrackedMeta } from "../track-changes";
import { Del, Ins, markParagraphMarkAs } from "../track-changes/emit";
import type { NoteConfig } from "./config";

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
	const del = <Del meta={meta}>{wrapped}</Del>;
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
			);
			result.push(<Ins meta={insMeta}>{newRun}</Ins>);
			inserted = true;
		}
		result.push(<Del meta={delMeta}>{convertTextToDelText(child)}</Del>);
	}
	// No prior user-text run (e.g. body has only the back-ref + whitespace):
	// just append a tracked insertion of the new text.
	if (!inserted) {
		const newRun = (
			<w.r>
				<w.t {...{ "xml:space": "preserve" }}>{newText}</w.t>
			</w.r>
		);
		result.push(<Ins meta={insMeta}>{newRun}</Ins>);
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
