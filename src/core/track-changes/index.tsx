import type { Document } from "../ast/document";
import type { XmlNode } from "../parser";
import {
	type ApplyVerb,
	applyTrackedChanges,
	type ChangeRecord,
	collectTrackedChanges,
	previewTrackedChanges,
} from "./apply";
import { Del, Ins, markParagraphMarkAs } from "./emit";

export {
	type ChangeFound,
	type ChangeRecord,
	collectTrackedChanges,
	TrackedChangeNotFoundError,
} from "./apply";

/** Paragraph children that are trackable as inserted/deleted run-level
 * content. Text runs are obvious; OMML equations (`<m:oMath>` /
 * `<m:oMathPara>`) sit at the same nesting level (inside `<w:p>`, beside
 * `<w:r>` siblings) and Word's revision model wraps them the same way.
 * Hyperlinks and field codes already contain `<w:r>` themselves, so they get
 * tracked by their inner runs. */
const TRACKABLE_PARAGRAPH_CHILDREN = new Set(["w:r", "m:oMath", "m:oMathPara"]);

export type TrackedMeta = {
	author: string;
	date: string;
	revisionId: number;
};

export type RevisionAllocator = { next(): number };

/** Cross-cutting lens over the document's tracked-changes facilities. Owns
 * no state — constructed at call sites with `new TrackChanges(document)` and
 * holds only a back-reference. Surface: minting revision metadata
 * (`mintMeta`/`createAllocator`), toggling `<w:trackChanges/>` in
 * settings.xml (`setEnabled`), inventorying revisions (`list` — reads
 * `document.trackedChangeReferences`, the reader's single walk), previewing
 * accept/reject (`preview`), applying them (`accept`/`reject`), and the
 * paragraph-level emit helpers every mutating CLI verb threads through
 * (`applyInsertion`/`applyDeletion`). */
export class TrackChanges {
	constructor(private document: Document) {}

	/** Mint a fresh `TrackedMeta` (author + date + revisionId) for a single
	 * tracked change about to be emitted. Internally builds a one-shot
	 * allocator — for operations that emit several coupled revisions (e.g.
	 * a tracked delete with ref-run + body + paragraph-mark), hold an
	 * allocator via `createAllocator()` instead so the max-id scan runs once. */
	mintMeta(authorFlag?: string): TrackedMeta {
		return {
			author: resolveAuthor(authorFlag),
			date: resolveDate(),
			revisionId: this.createAllocator().next(),
		};
	}

	/** Build a revision-id allocator seeded from a fresh max-id scan over
	 * document.xml + footnotes.xml + endnotes.xml. Returns monotonically
	 * increasing ids; safe to call `.next()` repeatedly for multi-revision
	 * operations without re-scanning. */
	createAllocator(): RevisionAllocator {
		let nextId = computeMaxRevisionId(this.document) + 1;
		return {
			next(): number {
				const id = nextId;
				nextId += 1;
				return id;
			},
		};
	}

	/** Toggle `<w:trackChanges/>` in word/settings.xml. Turning ON
	 * materializes the settings part if absent; turning OFF on a doc with
	 * no settings part is a no-op (absence already implies tracking is off). */
	setEnabled(on: boolean): void {
		if (on) {
			this.document.ensureSettings().setTrackChangesEnabled(true);
		} else {
			this.document.settings?.setTrackChangesEnabled(false);
		}
	}

	/** The complete tracked-change inventory in document order, read from
	 * `document.trackedChangeReferences` — the reader's single walk over the body
	 * then the note parts. The single source of truth for tcN ids: `track-changes
	 * list` and `accept`/`reject --at tcN` read the same map, so they can't
	 * disagree. */
	list(): ReturnType<typeof collectTrackedChanges> {
		return collectTrackedChanges(this.document);
	}

	/** Records that an accept/reject of `target` (tcN ids or "all") WOULD
	 * produce, without mutating — powers `--dry-run`. Throws
	 * `TrackedChangeNotFoundError` for an unknown id. */
	preview(target: string[] | "all", verb: ApplyVerb): ChangeRecord[] {
		return previewTrackedChanges(this.document, target, verb);
	}

	/** Accept tracked changes (tcN ids or "all"), incorporating them into the
	 * document. Includes the body-side note pairing + table-grid resync
	 * post-passes. Returns the applied records; caller saves. */
	accept(target: string[] | "all"): ChangeRecord[] {
		return applyTrackedChanges(this.document, target, "accept");
	}

	/** Reject tracked changes (tcN ids or "all"), reverting them. Returns the
	 * applied records; caller saves. */
	reject(target: string[] | "all"): ChangeRecord[] {
		return applyTrackedChanges(this.document, target, "reject");
	}

	/** Wrap a freshly-built paragraph's trackable run-level children in
	 * `<w:ins>` and mark its paragraph break inserted — the tracked form of
	 * an `insert`. Accept keeps the content; reject removes the whole
	 * paragraph. */
	applyInsertion(paragraph: XmlNode, authorFlag?: string): void {
		const mintMeta = this.metaMinter(authorFlag);
		paragraph.children = wrapContiguousTrackable(paragraph.children, (runs) => (
			<Ins meta={mintMeta()}>{runs}</Ins>
		));
		markParagraphMarkAs(paragraph, "ins", mintMeta());
	}

	/** Wrap a paragraph's trackable run-level children in `<w:del>` (with
	 * `<w:t>`→`<w:delText>` on text runs) and mark its paragraph break
	 * deleted — the tracked form of a `delete`. Accept removes the paragraph
	 * (merging forward); reject restores it. */
	applyDeletion(paragraph: XmlNode, authorFlag?: string): void {
		const mintMeta = this.metaMinter(authorFlag);
		paragraph.children = wrapContiguousTrackable(paragraph.children, (runs) => {
			const converted = runs.map((child) =>
				child.tag === "w:r" ? convertTextToDelText(child) : child,
			);
			return <Del meta={mintMeta()}>{converted}</Del>;
		});
		markParagraphMarkAs(paragraph, "del", mintMeta());
	}

	/** A revision-meta minter backed by one allocator + a fixed author/date —
	 * for operations that emit several coupled revisions in one call. */
	private metaMinter(authorFlag?: string): () => TrackedMeta {
		const allocator = this.createAllocator();
		const base = { author: resolveAuthor(authorFlag), date: resolveDate() };
		return () => ({ ...base, revisionId: allocator.next() });
	}
}

/** Wrap each contiguous span of `TRACKABLE_PARAGRAPH_CHILDREN` in the wrapper
 * `build` produces, passing non-trackable siblings (e.g. `<w:pPr>`) through at
 * their existing positions. Shared by `applyInsertion` / `applyDeletion`. */
function wrapContiguousTrackable(
	children: XmlNode[],
	build: (runs: XmlNode[]) => XmlNode,
): XmlNode[] {
	const out: XmlNode[] = [];
	let buffer: XmlNode[] = [];
	const flush = (): void => {
		if (buffer.length === 0) return;
		out.push(build(buffer));
		buffer = [];
	};
	for (const child of children) {
		if (TRACKABLE_PARAGRAPH_CHILDREN.has(child.tag)) {
			buffer.push(child);
			continue;
		}
		flush();
		out.push(child);
	}
	flush();
	return out;
}

export function resolveAuthor(authorFlag?: string): string {
	if (authorFlag) return authorFlag;
	if (Bun.env.DOCX_AUTHOR) return Bun.env.DOCX_AUTHOR;
	return "docx-cli";
}

export function resolveDate(): string {
	return Bun.env.DOCX_CLI_NOW ?? new Date().toISOString();
}

export function convertTextToDelText(node: XmlNode): XmlNode {
	const cloned = node.clone();
	mutateTextToDelText([cloned]);
	return cloned;
}

function computeMaxRevisionId(document: Document): number {
	let max = -1;
	const visit = (node: XmlNode): void => {
		// All revision-tracking wrappers share the same `w:id` namespace —
		// scan moves alongside ins/del so newly minted ids don't collide.
		if (
			node.tag !== "w:ins" &&
			node.tag !== "w:del" &&
			node.tag !== "w:moveFrom" &&
			node.tag !== "w:moveTo"
		)
			return;
		const idAttr = node.getAttribute("w:id");
		if (!idAttr) return;
		const value = Number(idAttr);
		if (Number.isFinite(value) && value > max) max = value;
	};
	// Tracked changes can live in document.xml, footnotes.xml, AND endnotes.xml
	// — Word emits body-side <w:ins>/<w:del> inside <w:footnote>/<w:endnote>
	// when a footnote is added/deleted under tracking (see `TrackedNoteBody`
	// in `core/notes/helpers.tsx`). Allocate across all three parts so a
	// new revision never collides with one that's already in a note body.
	walkXml(document.documentTree, visit);
	if (document.footnotes?.tree) walkXml(document.footnotes?.tree, visit);
	if (document.endnotes?.tree) walkXml(document.endnotes?.tree, visit);
	return max;
}

function walkXml(nodes: XmlNode[], visit: (node: XmlNode) => void): void {
	for (const node of nodes) {
		visit(node);
		if (node.children.length > 0) walkXml(node.children, visit);
	}
}

function mutateTextToDelText(nodes: XmlNode[]): void {
	for (const node of nodes) {
		if (node.tag === "w:t") node.tag = "w:delText";
		mutateTextToDelText(node.children);
	}
}
